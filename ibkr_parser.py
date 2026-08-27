#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import bisect
import re
import datetime as dt
from collections import defaultdict

import pdfplumber

from tax_utils import boi_ils, _round_ils

# ─── LOOKUP TABLES ───────────────────────────────────────────────────────────

ISIN_COUNTRY = {
    "US": "United States",
    "GB": "United Kingdom",
    "BM": "Bermuda",
    "JE": "Jersey",
    "IE": "Ireland",
    "CA": "Canada",
    "DE": "Germany",
    "DK": "Denmark",
    "FR": "France",
    "NL": "Netherlands",
    "CH": "Switzerland",
}

_WHT_COUNTRY_CODE = {
    "US": "United States",
    "GB": "United Kingdom",
    "NL": "Netherlands",
    "DK": "Denmark",
    "DE": "Germany",
    "FR": "France",
    "CH": "Switzerland",
    "IE": "Ireland",
    "CA": "Canada",
    "AU": "Australia",
    "JP": "Japan",
    "SE": "Sweden",
    "NO": "Norway",
    "FI": "Finland",
    "IT": "Italy",
    "ES": "Spain",
    "BE": "Belgium",
    "AT": "Austria",
}

# ─── MODULE-LEVEL REGEX PATTERNS ─────────────────────────────────────────────

LOT_PATTERN = re.compile(
    r"^Closed Lot:\s+(\d{4}-\d{2}-\d{2})\s+"
    r"([\d.]+)\s+"
    r"([\d.]+)\s+"
    r"([\d,]+\.\d+)\s+"
    r"([-\d,]+\.\d+)",
    re.MULTILINE,
)

CHILD_FILL_PATTERN = re.compile(
    r"(\d{4}-\d{2}-\d{2}),\s*\n"
    r"([A-Z][A-Z0-9]*)\s+"
    r"(?!-\s)(\S+)\s+"
    r"(-[\d,]+(?:\.\d+)?)\s+"
    r"[\d.]+\s+[\d.]+\s+"
    r"([\d,]+\.\d+)\s+"
    r"([-\d.]+)\s+"
    r"[-\d,]+\.\d+\s+[-\d,]+\.\d+\s+[-\d,]+\.\d+\s+C",
    re.MULTILINE,
)

AGGREGATE_FILL_PATTERN = re.compile(
    r"(\d{4}-\d{2}-\d{2}),\s*\n"
    r"([A-Z][A-Z0-9]*)\s+-\s+"
    r"(-[\d,]+(?:\.\d+)?)\s+"
    r"[\d.]+\s+[\d.]+\s+"
    r"([\d,]+\.\d+)\s+"
    r"([-\d.]+)\s+"
    r"[-\d,]+\.\d+\s+[-\d,]+\.\d+\s+[-\d,]+\.\d+\s+C",
    re.MULTILINE,
)

TOTAL_PATTERN = re.compile(
    r"^Total ([A-Z]+)\s+(-?\d[\d,]*)\s+([\d,]+\.\d+)\s+([-\d.]+)",
    re.MULTILINE,
)

CURRENCY_HEADER_PATTERN = re.compile(r"^(USD|EUR|GBP|CAD|CHF|JPY|AUD)$", re.MULTILINE)

# ─── CORPORATE ACTION PATTERNS ───────────────────────────────────────────────

# Numbers line for a CA disposal: report_date eff_date,time qty proceeds value realized_pl
# Matches only negative qty (disposal) — zero-proceeds rows (splits/spinoffs) have positive qty
CA_DISPOSAL_NUMS_PATTERN = re.compile(
    r"\d{4}-\d{2}-\d{2}\s+"            # report date (discard)
    r"(\d{4}-\d{2}-\d{2}),\s*"         # effective date (capture)
    r"\d{2}:\d{2}:\d{2}\s+"            # time (discard)
    r"(-[\d,]+(?:\.\d+)?)\s+"          # negative quantity (disposal)
    r"([\d,]+\.\d+)\s+"                # proceeds (positive)
    r"[-\d,]+\.\d+\s+"                 # value (skip)
    r"([-\d,]+\.\d+)",                  # realized P/L
)

# TICKER(ISIN) — used to find the ticker from the description line preceding the numbers line
CA_TICKER_PATTERN = re.compile(r"([A-Z][A-Z0-9]*)\([^\)]+\)")

# Closed Lot line under a Corporate Action: acq_date, basis, lot_qty, lot_realized_pl, term
CA_LOT_PATTERN = re.compile(
    r"Closed Lot:\s+(\d{4}-\d{2}-\d{2})\s+"
    r"Basis:\s+([\d,]+\.\d+)\s+"
    r"([\d,]+(?:\.\d+)?)\s+"
    r"([-\d,]+\.\d+)\s+"
    r"(ST|LT)",
)

# Section boundary patterns for isolating the CA section and finding Total lines within it
CA_SECTION_START = re.compile(r"Corporate Actions")
CA_TOTAL_PATTERN = re.compile(r"^Total\b", re.MULTILINE)


# ─── TEXT EXTRACTION ─────────────────────────────────────────────────────────

def extract_full_text(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def extract_year_from_filename(path: str) -> str:
    m = re.search(r"_(\d{8})_(\d{8})_tax_statement\.pdf$", path)
    if m:
        return m.group(2)[:4]
    m = re.search(r"(\d{4})", path)
    return m.group(1) if m else "unknown"


# ─── CAPITAL-GAINS LOT PARSER (raw broker facts — no Moses, no ILS conversion) ──

def parse_ibkr_corporate_action_lots(text: str) -> list[dict]:
    """
    Parse taxable stock disposals from the IBKR Corporate Actions section.

    Handles cash corporate actions (mergers, acquisitions, tender offers) where the
    disposition is recorded as a negative-quantity row with positive proceeds and one or
    more associated Closed Lot records carrying per-lot Basis and realized P/L.

    Lot proceeds are reconstructed exactly as:
        lot_proceeds_fc = lot_basis_fc + lot_realized_pl_fc
    NOT allocated proportionally by quantity — this preserves per-lot Moses accuracy
    when multiple acquisition lots at different dates are closed in one corporate action.

    Zero-proceeds events (splits, spinoffs) are silently skipped.
    A taxable disposal (negative qty, positive proceeds) with no Closed Lot records
    raises RuntimeError — cost basis is required to compute capital gains.

    IDs:
        fill_id = IBKR/CA/{effective_date}/{ticker}/{ca_seq}
        lot_id  = {fill_id}/LOT/{lot_seq}
    """
    # Isolate the Corporate Actions section
    ca_match = CA_SECTION_START.search(text)
    if not ca_match:
        return []
    ca_section = text[ca_match.start():]

    # Build currency block lookup within the CA section
    ca_currency_headers = sorted(
        [(m.start(), m.group(1)) for m in CURRENCY_HEADER_PATTERN.finditer(ca_section)],
        key=lambda x: x[0],
    )
    ca_ccy_positions = [c[0] for c in ca_currency_headers]

    def _ca_block_currency(pos: int) -> str:
        idx = bisect.bisect_right(ca_ccy_positions, pos) - 1
        return ca_currency_headers[idx][1] if idx >= 0 else "USD"

    # Find all CA Total line positions (used as window boundaries)
    total_positions = sorted(m.start() for m in CA_TOTAL_PATTERN.finditer(ca_section))

    ca_seq_counter: dict = defaultdict(int)
    result_rows = []

    for m in CA_DISPOSAL_NUMS_PATTERN.finditer(ca_section):
        effective_date = m.group(1)
        parent_qty     = float(m.group(2).replace(",", ""))   # negative
        parent_proc    = float(m.group(3).replace(",", ""))   # positive
        parent_rpl     = float(m.group(4).replace(",", ""))

        # Skip zero-proceeds events (splits, spinoffs, non-cash)
        if parent_proc == 0.0:
            continue

        # Find the ticker: last TICKER(ISIN) match before this numbers line
        preceding = ca_section[:m.start()]
        ticker_matches = list(CA_TICKER_PATTERN.finditer(preceding))
        if not ticker_matches:
            raise RuntimeError(
                f"Corporate action disposal on {effective_date} (qty={parent_qty}, "
                f"proc={parent_proc}) has no preceding TICKER(ISIN) — cannot identify security"
            )
        sym = ticker_matches[-1].group(1)

        # Find Closed Lot lines in the window between this match and the next Total line
        window_start = m.end()
        idx_total = bisect.bisect_right(total_positions, window_start)
        window_end = total_positions[idx_total] if idx_total < len(total_positions) else len(ca_section)
        window = ca_section[window_start:window_end]

        lots = list(CA_LOT_PATTERN.finditer(window))
        if not lots:
            raise RuntimeError(
                f"Corporate action disposal for {sym} on {effective_date} has positive proceeds "
                f"({parent_proc}) but no Closed Lot records — cannot compute cost basis"
            )

        # Reconstruct per-lot proceeds from basis + realized P/L (not proportional allocation)
        lot_data = []
        for lot in lots:
            acq_date  = lot.group(1)
            basis_fc  = float(lot.group(2).replace(",", ""))
            lot_qty   = float(lot.group(3).replace(",", ""))
            lot_rpl   = float(lot.group(4).replace(",", ""))
            lot_proc  = basis_fc + lot_rpl
            lot_data.append((acq_date, basis_fc, lot_qty, lot_rpl, lot_proc))

        # Validate fill-level invariants
        total_qty  = sum(d[2] for d in lot_data)
        total_rpl  = sum(d[3] for d in lot_data)
        total_proc = sum(d[4] for d in lot_data)

        if abs(total_qty - abs(parent_qty)) > 0.01:
            raise RuntimeError(
                f"Quantity mismatch for {sym} CA on {effective_date}: "
                f"lots sum to {total_qty} but parent row shows {abs(parent_qty)}"
            )
        if abs(total_rpl - parent_rpl) > 0.02:
            raise RuntimeError(
                f"Realized P/L mismatch for {sym} CA on {effective_date}: "
                f"lots sum to {total_rpl:.2f} but parent row shows {parent_rpl:.2f}"
            )
        if abs(total_proc - parent_proc) > 0.02:
            raise RuntimeError(
                f"Proceeds mismatch for {sym} CA on {effective_date}: "
                f"reconstructed {total_proc:.2f} but parent row shows {parent_proc:.2f}"
            )

        ccy = _ca_block_currency(m.start())
        key = (effective_date, sym)
        ca_seq_counter[key] += 1
        ca_seq = ca_seq_counter[key]
        fill_id = f"IBKR/CA/{effective_date}/{sym}/{ca_seq}"

        print(f"  CA {effective_date} {sym} ({ccy}): proceeds={parent_proc:.2f} "
              f"rpl={parent_rpl:.2f} fill_id={fill_id}")

        for lot_seq, (acq_date, basis_fc, lot_qty, lot_rpl, lot_proc) in enumerate(lot_data, start=1):
            lot_id = f"{fill_id}/LOT/{lot_seq}"
            result_rows.append({
                "source":                 "IBKR",
                "ticker":                 sym,
                "isin":                   "",
                "currency":               ccy,
                "exchange":               "CA",
                "quantity":               lot_qty,
                "acquisition_date":       acq_date,
                "sale_date":              effective_date,
                "acquisition_basis_fc":   basis_fc,
                "cost_per_share_fc":      basis_fc / lot_qty if lot_qty else 0.0,
                "gross_sale_proceeds_fc": lot_proc,
                "sale_commission_fc":     0.0,
                "fill_quantity_fc":       total_qty,
                "fill_proceeds_fc":       total_proc,
                "fill_commission_fc":     0.0,
                "fill_id":                fill_id,
                "lot_id":                 lot_id,
                "fill_seq":               ca_seq,
                "lot_seq":                lot_seq,
            })

    return result_rows


def _parse_trade_lots(text: str) -> list[dict]:
    """
    Parse closed lot disposals from the IBKR Trades section (C-coded fills).
    Returns empty list if none found — combined check is in parse_ibkr_lots_detail().
    """
    lots_pos = [
        (m.start(), m.group(1), float(m.group(2)), float(m.group(3)),
         float(m.group(4).replace(",", "")))
        for m in LOT_PATTERN.finditer(text)
    ]  # (pos, acq_date, lot_qty, cost_per_share, lot_basis_fc)

    child_fills = [
        (m.start(), m.group(1), m.group(2), m.group(3),
         abs(float(m.group(4).replace(",", ""))),
         float(m.group(5).replace(",", "")),
         float(m.group(6)))
        for m in CHILD_FILL_PATTERN.finditer(text)
    ]  # (pos, date, symbol, exchange, disposal_qty, proceeds, comm_signed)

    aggregates_list = []
    for m in AGGREGATE_FILL_PATTERN.finditer(text):
        aggregates_list.append((
            m.start(), m.group(1), m.group(2),
            abs(float(m.group(3).replace(",", ""))),
            float(m.group(4).replace(",", "")),
            float(m.group(5)),
        ))  # (pos, date, sym, agg_qty, agg_proc, agg_comm)

    currency_headers = sorted(
        [(m.start(), m.group(1)) for m in CURRENCY_HEADER_PATTERN.finditer(text)],
        key=lambda x: x[0],
    )
    ccy_positions = [c[0] for c in currency_headers]

    if not lots_pos or not child_fills:
        return []

    def _block_currency(pos):
        idx = bisect.bisect_right(ccy_positions, pos) - 1
        return currency_headers[idx][1] if idx >= 0 else "USD"

    child_fills_sorted = sorted(child_fills, key=lambda x: x[0])
    lots_sorted = sorted(lots_pos, key=lambda x: x[0])
    lot_positions = [l[0] for l in lots_sorted]
    total_positions_all = sorted(m.start() for m in TOTAL_PATTERN.finditer(text))

    # Assign each Closed Lot to the child fill immediately preceding it
    fill_lots = []
    for i, (fill_pos, *_) in enumerate(child_fills_sorted):
        next_fill_pos = child_fills_sorted[i + 1][0] if i + 1 < len(child_fills_sorted) else len(text)
        idx_total = bisect.bisect_right(total_positions_all, fill_pos)
        next_total_pos = total_positions_all[idx_total] if idx_total < len(total_positions_all) else len(text)
        upper = min(next_fill_pos, next_total_pos)
        lo = bisect.bisect_right(lot_positions, fill_pos)
        hi = bisect.bisect_left(lot_positions, upper)
        fill_lots.append(lots_sorted[lo:hi])

    # Reconcile aggregate parent rows against child fills
    def _nearest_aggregate(fill_pos, sym, date):
        candidates = [a for a in aggregates_list if a[0] < fill_pos and a[1] == date and a[2] == sym]
        return candidates[-1] if candidates else None

    agg_child_qty      = defaultdict(float)
    agg_child_proceeds = defaultdict(float)
    agg_child_comm     = defaultdict(float)
    for fill_pos, date, sym, exch, disposal_qty, proceeds, comm in child_fills_sorted:
        agg = _nearest_aggregate(fill_pos, sym, date)
        if agg is not None:
            agg_pos = agg[0]
            agg_child_qty[agg_pos]      += disposal_qty
            agg_child_proceeds[agg_pos] += proceeds
            agg_child_comm[agg_pos]     += comm

    for agg_pos, agg_date, agg_sym, agg_qty, agg_proc, agg_comm in aggregates_list:
        if abs(agg_child_qty[agg_pos] - agg_qty) > 0.01:
            raise RuntimeError(
                f"Qty reconciliation failed for {agg_sym} {agg_date}: child fills sum to "
                f"{agg_child_qty[agg_pos]} but aggregate row shows {agg_qty}"
            )
        if abs(agg_child_proceeds[agg_pos] - agg_proc) > 0.02:
            raise RuntimeError(
                f"Proceeds reconciliation failed for {agg_sym} {agg_date}: child fills sum to "
                f"{agg_child_proceeds[agg_pos]:.2f} but aggregate row shows {agg_proc:.2f}"
            )
        if abs(agg_child_comm[agg_pos] - agg_comm) > 0.02:
            raise RuntimeError(
                f"Commission reconciliation failed for {agg_sym} {agg_date}: child fills sum to "
                f"{agg_child_comm[agg_pos]:.4f} but aggregate row shows {agg_comm:.4f}"
            )

    # Assign fill_seq: 1-based counter per (date, sym, exch) group, position-ordered
    fill_seq_counter: dict = defaultdict(int)
    fill_seqs = []
    for fill_pos, date, sym, exch, disposal_qty, proceeds, comm in child_fills_sorted:
        key = (date, sym, exch)
        fill_seq_counter[key] += 1
        fill_seqs.append(fill_seq_counter[key])

    result_rows = []

    for i, (fill_pos, date, sym, exch, disposal_qty, proceeds, comm_signed) in enumerate(child_fills_sorted):
        lots = fill_lots[i]
        fill_seq = fill_seqs[i]

        if not lots:
            raise RuntimeError(
                f"C-coded closing fill for {sym} on {date} (exchange {exch}) has no "
                f"associated Closed Lot rows — cannot compute cost basis"
            )

        lot_total_qty = sum(l[2] for l in lots)
        if abs(lot_total_qty - disposal_qty) > 0.01:
            raise RuntimeError(
                f"Lot qty mismatch for {sym} fill {date}: lots sum to "
                f"{lot_total_qty} but fill qty is {disposal_qty}"
            )

        ccy = _block_currency(fill_pos)
        fill_id = f"IBKR/{date}/{sym}/{exch}/{fill_seq}"

        print(f"  {date} {sym}/{exch} ({ccy}): proceeds={proceeds:.2f} "
              f"comm={comm_signed:.2f} fill_id={fill_id}")

        for lot_seq, (lot_pos, acq_date, lot_qty, cost_per_share, lot_basis_fc) in enumerate(lots, start=1):
            reconstructed = lot_qty * cost_per_share
            if abs(reconstructed - lot_basis_fc) > 0.01:
                print(f"    WARNING: basis mismatch for {sym} acq={acq_date}: "
                      f"qty*cps={reconstructed:.4f} vs lot_basis={lot_basis_fc:.4f}")

            lot_proceeds_fc = proceeds * (lot_qty / lot_total_qty)
            lot_comm_fc     = comm_signed * (lot_qty / lot_total_qty)
            lot_id = f"{fill_id}/LOT/{lot_seq}"

            result_rows.append({
                "source":               "IBKR",
                "ticker":               sym,
                "isin":                 "",
                "currency":             ccy,
                "exchange":             exch,
                "quantity":             lot_qty,
                "acquisition_date":     acq_date,
                "sale_date":            date,
                "acquisition_basis_fc": lot_basis_fc,
                "cost_per_share_fc":    cost_per_share,
                "gross_sale_proceeds_fc": lot_proceeds_fc,
                "sale_commission_fc":   lot_comm_fc,  # signed: negative=cost
                # fill-level totals — used by _verify_allocation to check lot sums
                "fill_quantity_fc":     disposal_qty,
                "fill_proceeds_fc":     proceeds,
                "fill_commission_fc":   comm_signed,
                "fill_id":              fill_id,
                "lot_id":               lot_id,
                "fill_seq":             fill_seq,
                "lot_seq":              lot_seq,
            })

    if not result_rows:
        return []

    return result_rows


def _has_unparsed_disposals(text: str) -> bool:
    """
    Return True if the text contains signals that stock disposals occurred but
    were not captured by the parsers.  Used to distinguish "no sales this year"
    from a genuine parse failure.

    Signals checked:
    1. C-coded fills (CHILD_FILL_PATTERN / AGGREGATE_FILL_PATTERN) in the text —
       closing trades that should have produced lots.
    2. Negative-quantity stock trade lines (excluding forex symbols like USD.ILS).
    3. Non-zero stock realized P/L in the Realized & Unrealized Performance Summary.
    """
    # 1. C-coded fill lines
    if CHILD_FILL_PATTERN.search(text) or AGGREGATE_FILL_PATTERN.search(text):
        return True

    # 2. Negative-quantity stock lines (not forex).
    # Stock trade lines look like: SYMBOL EXCHANGE -qty price ...
    # Forex lines contain a dot (e.g. USD.ILS) — exclude those.
    neg_stock_trade = re.compile(
        r"^([A-Z][A-Z0-9]{0,9})\s+\S+\s+(-[\d,]+(?:\.\d+)?)\s+[\d.]+",
        re.MULTILINE,
    )
    for m in neg_stock_trade.finditer(text):
        symbol = m.group(1)
        if "." not in symbol:   # forex symbols contain a dot
            return True

    # 3. Non-zero stock realized P/L in the Performance Summary.
    # Lines look like: SYMBOL  cost  s/t_profit  s/t_loss  l/t_profit  l/t_loss  total ...
    # A non-zero realized gain or loss means a position was closed.
    # Exclude currency/forex symbols — they appear in the same section but are not stock lots.
    _CURRENCY_SYMBOLS = frozenset({"USD", "EUR", "GBP", "CAD", "CHF", "JPY", "AUD",
                                   "ILS", "HKD", "SGD", "NZD", "SEK", "NOK", "DKK"})
    perf_line = re.compile(
        r"^([A-Z][A-Z0-9]{0,9})\s+"     # symbol (no dot → not forex pair)
        r"[\d,]+\.\d+\s+"               # cost adj
        r"([-\d,]+\.\d+)\s+"            # S/T profit
        r"([-\d,]+\.\d+)\s+"            # S/T loss
        r"([-\d,]+\.\d+)\s+"            # L/T profit
        r"([-\d,]+\.\d+)",              # L/T loss
        re.MULTILINE,
    )
    for m in perf_line.finditer(text):
        symbol = m.group(1)
        if "." in symbol or symbol in _CURRENCY_SYMBOLS:
            continue
        values = [float(m.group(i).replace(",", "")) for i in (2, 3, 4, 5)]
        if any(v != 0.0 for v in values):
            return True

    return False


def parse_ibkr_lots_detail(text: str) -> list[dict]:
    """
    Parse all taxable IBKR stock disposals: normal Trades (C-coded fills) and
    Corporate Action disposals (mergers, acquisitions, tender offers).

    Returns one dict per closed acquisition lot with raw broker facts only.
    No ILS conversion, no Moses calculation — those happen in capital_gains.py.

    Returns [] when the PDF is valid but contains no taxable stock disposals
    (e.g. a year where only opening trades occurred).

    Raises RuntimeError when disposal signals are present but no lots were parsed
    (indicating a parse failure rather than a genuinely empty year).
    """
    trade_lots = _parse_trade_lots(text)
    ca_lots    = parse_ibkr_corporate_action_lots(text)
    all_lots   = trade_lots + ca_lots
    if not all_lots and _has_unparsed_disposals(text):
        raise RuntimeError(
            "IBKR PDF contains signals of stock disposals (C-coded fills, "
            "negative-quantity trades, or non-zero realized P/L) but no lots "
            "were parsed. Check that the correct tax statement PDF was supplied."
        )
    return all_lots


# ─── INCOME PARSERS (return ILS amounts — not capital gains, not via capital_gains.py) ──

def _extract_dividends_text(pdf_path: str) -> str:
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "Dividends" not in text:
                continue
            right = page.crop((page.width * 0.5, 0, page.width, page.height))
            parts.append(right.extract_text() or "")
    return "\n".join(parts)


def parse_ticker_tax_country(pdf_path: str) -> dict:
    """Build ticker → tax-source country from IBKR withholding labels (takes priority over ISIN prefix)."""
    wht_pattern = re.compile(r"([A-Z0-9]+)\([A-Z]{2}[A-Z0-9]+\).*?-\s+([A-Z]{2})\s+Tax")
    ticker_country = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "Withholding Tax" not in text or "Dividends" not in text:
                continue
            left = page.crop((0, 0, page.width * 0.5, page.height))
            for line in (left.extract_text() or "").splitlines():
                m = wht_pattern.search(line)
                if m:
                    ticker, code = m.group(1), m.group(2)
                    ticker_country[ticker] = _WHT_COUNTRY_CODE.get(code, code)
    return ticker_country


def parse_dividends_by_country(pdf_path: str, ticker_tax_country: dict) -> dict:
    div_text = _extract_dividends_text(pdf_path)
    # Split-line format: TICKER (ISIN) description\nDATE amount
    # Allows space before '(' (e.g. "BEPC (CA...)") and numeric ISINs (e.g. BEP partnership "114646358")
    split_pattern = re.compile(
        r"([A-Z][A-Z0-9]*)\s*\(([A-Z0-9]+)\)\s+"
        r"(Cash Dividend|Payment in Lieu of Dividend)"
        r"(?:\s+(USD|GBP|EUR)\s+[\d.]+\s+per\s+Share)?"
        r"[^\n]*\n"
        r"(\d{4}-\d{2}-\d{2})\s+"
        r"(-?[\d,]+\.\d+)",
        re.MULTILINE,
    )
    # Inline format: DATE TICKER(ISIN) description amount (all on one line)
    inline_pattern = re.compile(
        r"^(\d{4}-\d{2}-\d{2})\s+"
        r"([A-Z][A-Z0-9]*)\s*\(([A-Z0-9]+)\)\s+"
        r"(Cash Dividend|Payment in Lieu of Dividend)"
        r"[^\n]*\s+(-?[\d,]+\.\d+)$",
        re.MULTILINE,
    )

    def _country(ticker: str, isin: str) -> str:
        if ticker in ticker_tax_country:
            return ticker_tax_country[ticker]
        prefix = isin[:2].upper()
        if prefix.isalpha():
            return ISIN_COUNTRY.get(prefix, prefix)
        return "United States"  # numeric ISIN = US partnership

    by_country: dict = defaultdict(float)
    matched_spans: set = set()

    for m in split_pattern.finditer(div_text):
        for pos in range(m.start(), m.end()):
            matched_spans.add(pos)
        ticker, isin, _ptype, currency, date_str, amount_str = m.groups()
        country = _country(ticker, isin)
        ccy = (currency or "USD").upper()
        amount = float(amount_str.replace(",", ""))
        rate = boi_ils(ccy, date_str)
        by_country[country] += amount * rate
        print(f"  {date_str} {ticker}/{isin} ({country}) {ccy} {amount:.2f} @ {rate:.4f} = {amount*rate:.2f} ILS")

    for m in inline_pattern.finditer(div_text):
        if m.start() in matched_spans:
            continue
        date_str, ticker, isin, _ptype, amount_str = m.groups()
        country = _country(ticker, isin)
        amount = float(amount_str.replace(",", ""))
        rate = boi_ils("USD", date_str)
        by_country[country] += amount * rate
        print(f"  {date_str} {ticker}/{isin} ({country}) USD {amount:.2f} @ {rate:.4f} = {amount*rate:.2f} ILS [inline]")

    if not by_country:
        raise RuntimeError("No dividend rows parsed from PDF")
    return {country: total for country, total in sorted(by_country.items())}


def parse_interest_transactions_ils(pdf_path: str) -> float:
    """Parse interest transactions from all pages; convert at per-transaction BoI rate."""
    row_pattern = re.compile(r"^(\d{4}-\d{2}-\d{2})\s+(USD|EUR|GBP|CAD|CHF|JPY|AUD)\s+.+\s+([\d,]+\.\d+)$")
    total = 0.0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text = page.extract_text() or ""
            if "Interest" not in full_text:
                continue
            in_interest = False
            for line in full_text.splitlines():
                s = line.strip()
                if s == "Interest":
                    in_interest = True
                    continue
                if not in_interest:
                    continue
                if s.startswith("Total in ILS") or s.startswith("Fees") or s.startswith("Total Interest in ILS"):
                    break
                m = row_pattern.match(s)
                if m:
                    date_str, ccy, amount_str = m.groups()
                    amount = float(amount_str.replace(",", ""))
                    rate = boi_ils(ccy, date_str)
                    ils = amount * rate
                    total += ils
                    print(f"  {date_str} {ccy} interest {amount:.2f} @ {rate:.4f} = {ils:.2f} ILS")
    return total


def parse_withholding_tax_ils(text: str) -> int:
    # Uses IBKR's pre-converted ILS total from Cash Report, not per-transaction BoI rates.
    # Methodology differs from dividends/interest (per-transaction BoI). Affects Form 1324.
    m = re.search(r"Withholding Tax\s+([-\d,]+\.\d+)", text)
    if not m:
        raise RuntimeError("Could not parse Withholding Tax from Cash Report")
    value = float(m.group(1).replace(",", ""))
    if value > 0:
        raise RuntimeError(f"Expected withholding tax to be negative, got {value}")
    return _round_ils(abs(value))
