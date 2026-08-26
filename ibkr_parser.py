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

def parse_ibkr_lots_detail(text: str) -> list[dict]:
    """
    Parse all closed lot disposals from IBKR Trades section text.

    Returns one dict per closed acquisition lot with raw broker facts only.
    No ILS conversion, no Moses calculation — those happen in capital_gains.py.

    Each dict contains:
        source, ticker, isin, currency, quantity, acquisition_date, sale_date,
        exchange, acquisition_basis_fc, cost_per_share_fc,
        gross_sale_proceeds_fc, sale_commission_fc (signed: negative=cost),
        fill_id, lot_id, fill_seq, lot_seq

    Fails loudly on:
    - C-coded fill with no Closed Lot lines
    - Lot qty mismatch for any fill
    - Aggregate qty/proceeds/commission mismatch for partial fills
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

    if not lots_pos:
        raise RuntimeError("No Closed Lot lines found in PDF")
    if not child_fills:
        raise RuntimeError("No child C-coded fill rows found in PDF")

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
        raise RuntimeError("No closed lot disposals found in PDF")

    return result_rows


# ─── INCOME PARSERS (return ILS amounts — not capital gains, not via capital_gains.py) ──

def _extract_dividends_text(pdf_path: str) -> str:
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if "Withholding Tax" in text and "Dividends" in text:
                right = page.crop((page.width * 0.5, 0, page.width, page.height))
                parts.append(right.extract_text() or "")
            elif "Dividends" in text and i == total - 1:
                parts.append(text)
            elif i == total - 1:
                parts.append(text)
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
    pattern = re.compile(
        r"([A-Z0-9]+)\(([A-Z]{2}[A-Z0-9]+)\)\s+"
        r"(Cash Dividend|Payment in Lieu of Dividend)"
        r"(?:\s+(USD|GBP|EUR)\s+[\d.]+\s+per\s+Share)?"
        r"[^\n]*\n"
        r"(\d{4}-\d{2}-\d{2})\s+"
        r"([\d,]+\.\d+)",
        re.MULTILINE,
    )
    by_country: dict = defaultdict(float)
    for m in pattern.finditer(div_text):
        ticker, isin, _ptype, currency, date_str, amount_str = m.groups()
        if ticker in ticker_tax_country:
            country = ticker_tax_country[ticker]
        else:
            country = ISIN_COUNTRY.get(isin[:2].upper(), isin[:2].upper())
        ccy = (currency or "USD").upper()
        amount = float(amount_str.replace(",", ""))
        rate = boi_ils(ccy, date_str)
        by_country[country] += amount * rate
        print(f"  {date_str} {ticker}/{isin} ({country}) {ccy} {amount:.2f} @ {rate:.4f} = {amount*rate:.2f} ILS")
    if not by_country:
        raise RuntimeError("No dividend rows parsed from PDF")
    return {country: total for country, total in sorted(by_country.items())}


def parse_interest_transactions_ils(pdf_path: str) -> float:
    """Parse individual interest transactions from last page; convert each at per-transaction BoI rate."""
    with pdfplumber.open(pdf_path) as pdf:
        last_text = pdf.pages[-1].extract_text() or ""

    lines = last_text.splitlines()
    in_interest = False
    total = 0.0
    row_pattern = re.compile(r"^(\d{4}-\d{2}-\d{2})\s+(\S+)\s+.+\s+([\d,]+\.\d+)$")
    ccy_header = re.compile(r"^(USD|EUR|GBP|CAD|CHF|JPY|AUD)$")
    current_ccy = "USD"
    for line in lines:
        if line.strip() == "Interest":
            in_interest = True
            continue
        if not in_interest:
            continue
        if line.startswith("Total"):
            break
        m_ccy = ccy_header.match(line.strip())
        if m_ccy:
            current_ccy = m_ccy.group(1)
            continue
        m = row_pattern.match(line.strip())
        if m:
            date_str, ccy_inline, amount_str = m.group(1), m.group(2), m.group(3)
            ccy = ccy_inline if ccy_inline in {"USD", "EUR", "GBP", "CAD", "CHF", "JPY", "AUD"} else current_ccy
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
