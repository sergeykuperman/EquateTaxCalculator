#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import bisect
import glob
import os
import re
import datetime as dt
from collections import defaultdict

import pdfplumber
import pandas as pd
import requests

ISIN_COUNTRY = {
    "US": "United States",
    "GB": "United Kingdom",
    "BM": "Bermuda",
    "JE": "Jersey",
    "IE": "Ireland",
    "CA": "Canada",
    "DE": "Germany",
    "FR": "France",
    "NL": "Netherlands",
    "CH": "Switzerland",
}

_fx_cache = {}


def boi_ils(currency: str, day: str | dt.date) -> float:
    if isinstance(day, dt.date):
        day = day.strftime("%Y-%m-%d")
    key = (currency.upper(), day)
    if key in _fx_cache:
        return _fx_cache[key]
    url = (
        f"https://edge.boi.gov.il/FusionEdgeServer/"
        f"sdmx/v2/data/dataflow/BOI.STATISTICS/EXR/1.0/RER_{currency.upper()}_ILS"
    )
    d = dt.date.fromisoformat(day)
    for delta in range(7):
        candidate = (d - dt.timedelta(days=delta)).strftime("%Y-%m-%d")
        params = {
            "c[DATA_TYPE]": "OF00",
            "startperiod": candidate,
            "endperiod": candidate,
            "format": "sdmx-json",
        }
        resp = requests.get(url, params=params, timeout=10)
        if resp.status_code != 200 or not resp.text.strip():
            continue
        try:
            js = resp.json()
            obs = js["data"]["dataSets"][0]["series"]["0:0:0:0:0:0"]["observations"]
            rate = float(list(obs.values())[0][0])
            _fx_cache[key] = rate
            return rate
        except (KeyError, IndexError, ValueError):
            continue
    raise RuntimeError(f"No BoI {currency}/ILS rate found within 7 days before {day}")


def extract_full_text(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def extract_dividends_text(pdf_path: str) -> str:
    # Pages 11-13 have a two-column layout (WHT left, Dividends right).
    # Crop to right half to get clean dividend text. Page 14 is dividends-only.
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if "Withholding Tax" in text and "Dividends" in text:
                # Two-column page: crop right half for dividends
                right = page.crop((page.width * 0.5, 0, page.width, page.height))
                parts.append(right.extract_text() or "")
            elif "Dividends" in text and i == total - 1:
                parts.append(text)
            elif i == total - 1:
                # Last page may contain dividends continuation + interest
                parts.append(text)
    return "\n".join(parts)


def parse_ibkr_totals_ils(pdf_path: str) -> tuple[float, float]:
    """
    Parse exact IBKR-computed ILS totals from the last page.
    Returns (dividend_total_ils, interest_total_ils) as floats (not yet rounded).
    """
    with pdfplumber.open(pdf_path) as pdf:
        last_page_text = pdf.pages[-1].extract_text() or ""

    m = re.search(r"Total Dividends in ILS\s+([\d,]+\.\d+)", last_page_text)
    if not m:
        raise RuntimeError("Could not parse 'Total Dividends in ILS' from PDF")
    dividend_total = float(m.group(1).replace(",", ""))

    m = re.search(
        r"\bInterest\b\s*\nDate Description Amount\b.*?Total in ILS\s+([\d,]+\.\d+)",
        last_page_text,
        re.DOTALL,
    )
    if not m:
        raise RuntimeError("Could not parse interest 'Total in ILS' from PDF")
    interest_total = float(m.group(1).replace(",", ""))

    return dividend_total, interest_total


def parse_dividends_by_country(pdf_path: str) -> dict:
    # Extract clean dividends-only text from right column of two-column pages
    div_text = extract_dividends_text(pdf_path)

    # Each entry spans 2-3 lines:
    #   Line 1: TICKER(ISIN) description [currency] [per share info]
    #   Line 2: DATE  amount
    #   Line 3: (Dividend type)  [optional, wrapped]
    # We anchor on the ISIN pattern and grab the date+amount from the next line.
    pattern = re.compile(
        r"([A-Z0-9]+)\(([A-Z]{2}[A-Z0-9]+)\)\s+"
        r"(Cash Dividend|Payment in Lieu of Dividend)"
        r"(?:\s+(USD|GBP|EUR)\s+[\d.]+\s+per\s+Share)?"
        r"[^\n]*\n"                        # rest of description line
        r"(\d{4}-\d{2}-\d{2})\s+"          # date on next line
        r"([\d,]+\.\d+)",                   # amount
        re.MULTILINE,
    )
    by_country = defaultdict(float)
    for m in pattern.finditer(div_text):
        _ticker, isin, _ptype, currency, date_str, amount_str = m.groups()
        isin_prefix = isin[:2].upper()
        country = ISIN_COUNTRY.get(isin_prefix, isin_prefix)
        ccy = (currency or "USD").upper()
        amount = float(amount_str.replace(",", ""))
        rate = boi_ils(ccy, date_str)
        by_country[country] += amount * rate
        print(f"  {date_str} {isin} ({country}) {ccy} {amount:.2f} @ {rate:.4f} = {amount*rate:.2f} ILS")
    if not by_country:
        raise RuntimeError("No dividend rows parsed from PDF")
    return {country: total for country, total in sorted(by_country.items())}


def parse_withholding_tax_ils(text: str) -> int:
    # Cash Report base-currency summary: "Withholding Tax  -N,NNN.NN"
    m = re.search(r"Withholding Tax\s+([-\d,]+\.\d+)", text)
    if not m:
        raise RuntimeError("Could not parse Withholding Tax from Cash Report")
    return round(float(m.group(1).replace(",", "")))


def parse_realized_stocks_ils(text: str) -> tuple[int, int]:
    # Performance Summary "Total Stocks" row: symbol + 12 space-separated numbers
    # Realized block columns (0-indexed from after symbol):
    #   0: Cost Adj, 1: S/T Profit, 2: S/T Loss, 3: L/T Profit, 4: L/T Loss, 5: Total
    m = re.search(
        r"Total Stocks\s+([-\d,]+\.\d+)\s+([-\d,]+\.\d+)\s+([-\d,]+\.\d+)"
        r"\s+([-\d,]+\.\d+)\s+([-\d,]+\.\d+)\s+([-\d,]+\.\d+)",
        text,
    )
    if not m:
        raise RuntimeError("Could not parse Total Stocks row from Performance Summary")
    # groups: cost_adj, st_profit, st_loss, lt_profit, lt_loss, total
    st_profit = float(m.group(2).replace(",", ""))
    st_loss   = float(m.group(3).replace(",", ""))
    lt_profit = float(m.group(4).replace(",", ""))
    lt_loss   = float(m.group(5).replace(",", ""))
    gains  = st_profit + lt_profit
    losses = st_loss + lt_loss
    return round(gains), round(losses)


def parse_closed_lots_ils(text: str) -> tuple[int, int, int]:
    """
    Reconstruct realized gains/losses and gross sale value from closed lot data
    using BoI FX rates on both sale date and acquisition date, consistent with
    Form 1325 methodology (same as tax_calculator.py for EquatePlus).

    Anchors on "Closed Lot:" lines as the primary signal of a disposal — this
    correctly handles symbols where net qty is 0 or positive (buy > sell) but
    actual disposals still occurred.

    Returns (gains_ils, losses_ils, gross_sales_ils) all as rounded ints.
    gains_ils  >= 0  (sum of profitable lots only)
    losses_ils <= 0  (sum of loss lots only)
    gross_sales_ils >= 0
    """
    # Parse all "Closed Lot:" lines with position
    lot_pattern = re.compile(
        r"^Closed Lot:\s+(\d{4}-\d{2}-\d{2})\s+"
        r"([\d.]+)\s+"           # qty
        r"([\d.]+)\s+"           # cost per share
        r"([\d,]+\.\d+)\s+"      # lot basis — not used, recalculated
        r"([-\d,]+\.\d+)",       # lot realized P/L — not used, recalculated
        re.MULTILINE,
    )
    # Parse all "Total SYMBOL qty proceeds comm ..." rows (any qty sign)
    total_pattern = re.compile(
        r"^Total ([A-Z]+)\s+(-?\d[\d,]*)\s+([\d,]+\.\d+)\s+([-\d.]+)",
        re.MULTILINE,
    )
    # Parse all sell fill rows (negative qty, C-coded) for sale date
    sell_fill_pattern = re.compile(
        r"(\d{4}-\d{2}-\d{2}),\s*\n"
        r"([A-Z]+)\s+\S+\s+-\d+\s+[\d.]+\s+[\d.]+\s+[\d,]+\.\d+\s+"
        r"[-\d.]+\s+[-\d,]+\.\d+\s+[-\d,]+\.\d+\s+[-\d,]+\.\d+\s+C",
        re.MULTILINE,
    )
    # Currency section headers: a line that is exactly a 3-letter currency code
    currency_header_pattern = re.compile(r"^(USD|EUR|GBP|CAD|CHF|JPY|AUD)$", re.MULTILINE)

    # Build position-sorted lists
    lots_pos = [(m.start(), m.group(1), float(m.group(2)),
                 float(m.group(3)), float(m.group(4).replace(",", "")))
                for m in lot_pattern.finditer(text)]
    totals_pos = [(m.start(), m.group(1), int(m.group(2).replace(",", "")),
                   float(m.group(3).replace(",", "")), abs(float(m.group(4))))
                  for m in total_pattern.finditer(text)]
    sell_fills = [(m.start(), m.group(1), m.group(2))  # (pos, date, symbol)
                  for m in sell_fill_pattern.finditer(text)]
    # currency headers sorted by position: [(pos, currency), ...]
    currency_headers = sorted(
        [(m.start(), m.group(1)) for m in currency_header_pattern.finditer(text)],
        key=lambda x: x[0],
    )
    ccy_positions = [c[0] for c in currency_headers]

    if not lots_pos:
        raise RuntimeError("No Closed Lot lines found in PDF")

    # For each closed lot, find its owning symbol block:
    # The lot belongs to the nearest "Total SYMBOL" that comes AFTER the lot's position.
    # Group lots by that Total row (identified by its position).
    totals_pos_sorted = sorted(totals_pos, key=lambda x: x[0])

    # Build: total_pos -> (symbol, qty, proceeds, comm, currency, [lots])
    def _block_currency(pos):
        idx = bisect.bisect_right(ccy_positions, pos) - 1
        return currency_headers[idx][1] if idx >= 0 else "USD"

    total_blocks = {pos: (sym, qty, proc, comm, _block_currency(pos), [])
                    for pos, sym, qty, proc, comm in totals_pos_sorted}
    total_positions = sorted(total_blocks.keys())

    for lot_pos, acq_date, lot_qty, cost_per_share, _ in lots_pos:
        # Find the first Total row position that is after this lot
        idx = bisect.bisect_right(total_positions, lot_pos)
        if idx >= len(total_positions):
            continue  # lot after last Total — shouldn't happen
        owner_pos = total_positions[idx]
        total_blocks[owner_pos][5].append((acq_date, lot_qty, cost_per_share))

    # sell_fills sorted by position for per-block date lookup
    sell_fills_sorted = sorted(sell_fills, key=lambda x: x[0])
    fill_positions = [f[0] for f in sell_fills_sorted]

    # Only process Total blocks that have closed lots (= had actual disposals)
    total_gains = 0.0
    total_losses = 0.0
    total_gross = 0.0

    for i, total_pos in enumerate(total_positions):
        sym, qty, proceeds_usd, comm_usd, ccy, lots = total_blocks[total_pos]
        if not lots:
            continue  # no disposals in this symbol block

        # Find the sale date from sell fill rows within this block's text range.
        # A block spans from the previous Total (exclusive) to this Total (exclusive).
        block_start = total_positions[i - 1] if i > 0 else 0
        # fills whose position falls in (block_start, total_pos)
        lo = bisect.bisect_right(fill_positions, block_start)
        hi = bisect.bisect_left(fill_positions, total_pos)
        block_fills = sell_fills_sorted[lo:hi]
        if not block_fills:
            raise RuntimeError(f"No sell fill date found for {sym} block at pos {total_pos}")
        # Last fill date in the block wins (handles partial fills on same day)
        sale_date = block_fills[-1][1]

        fx_sale = boi_ils(ccy, sale_date)
        gross_ils = proceeds_usd * fx_sale
        total_gross += gross_ils

        lot_total_qty = sum(l[1] for l in lots)
        print(f"  {sale_date} {sym} ({ccy}): proceeds={proceeds_usd:.2f} "
              f"comm={comm_usd:.2f} fx_sale={fx_sale:.4f}")

        for acq_date, lot_qty, cost_per_share in lots:
            fx_acq = boi_ils(ccy, acq_date)
            lot_proceeds_usd = proceeds_usd * (lot_qty / lot_total_qty)
            lot_gross_ils = lot_proceeds_usd * fx_sale
            lot_cost_ils = lot_qty * cost_per_share * fx_acq
            lot_comm_ils = comm_usd * (lot_qty / lot_total_qty) * fx_sale
            lot_gain_ils = lot_gross_ils - lot_cost_ils - lot_comm_ils
            print(f"    lot acq={acq_date} qty={lot_qty} cost/sh=${cost_per_share:.4f} "
                  f"fx_acq={fx_acq:.4f} → gain={lot_gain_ils:.2f} ILS")
            if lot_gain_ils >= 0:
                total_gains += lot_gain_ils
            else:
                total_losses += lot_gain_ils

    if total_gross == 0:
        raise RuntimeError("No closed lot disposals found in PDF")

    return round(total_gains), round(total_losses), round(total_gross)


def extract_year_from_filename(path: str) -> str:
    m = re.search(r"_(\d{8})_(\d{8})_tax_statement\.pdf$", path)
    if m:
        return m.group(2)[:4]  # end-date year
    m = re.search(r"(\d{4})", path)
    return m.group(1) if m else "unknown"


def build_summary_rows(
    interest_ils: int,
    dividends_by_country: dict,
    dividend_total: int,
    div_plus_income: int,
    wht_ils: int,
    gains_ils: int,
    losses_ils: int,
    gross_sales_ils: int,
) -> list[dict]:
    rows = [
        {"Label": "Interest income", "Value (ILS)": interest_ils},
        {"Label": "Dividend spread by country", "Value (ILS)": ""},
    ]
    for country, amount in dividends_by_country.items():
        rows.append({"Label": f"  {country}", "Value (ILS)": amount})
    rows += [
        {"Label": "Dividend total", "Value (ILS)": dividend_total},
        {"Label": "Dividend + external income total", "Value (ILS)": div_plus_income},
        {"Label": "Foreign withholding tax", "Value (ILS)": wht_ils},
        {"Label": "IBKR positive realized stock gains", "Value (ILS)": gains_ils},
        {"Label": "IBKR realized stock losses", "Value (ILS)": losses_ils},
        {"Label": "IBKR gross sale value (stock disposals)", "Value (ILS)": gross_sales_ils},
    ]
    return rows


def main():
    pdfs = glob.glob("*_tax_statement.pdf")
    if not pdfs:
        raise SystemExit("No *_tax_statement.pdf found in current directory")
    pdf_path = pdfs[0]
    if len(pdfs) > 1:
        print(f"Multiple tax statement PDFs found, using: {pdf_path}")

    year = extract_year_from_filename(pdf_path)
    print(f"Parsing {pdf_path} (year {year}) ...")

    text = extract_full_text(pdf_path)

    print("Parsing IBKR exact dividend and interest totals ...")
    dividend_total_exact, interest_exact = parse_ibkr_totals_ils(pdf_path)
    print(f"  Total Dividends in ILS: {dividend_total_exact}")
    print(f"  Total Interest in ILS:  {interest_exact}")

    print("Parsing dividends by country (for breakdown only) ...")
    dividends_by_country = parse_dividends_by_country(pdf_path)
    div_plus_income_exact = dividend_total_exact + interest_exact
    print(f"  Dividends by country: { {k: round(v) for k, v in dividends_by_country.items()} }")
    print(f"  Dividend + interest total: {div_plus_income_exact}")

    print("Parsing withholding tax ...")
    wht_ils = parse_withholding_tax_ils(text)
    print(f"  WHT in ILS: {wht_ils}")

    print("Parsing realized stock gains/losses and gross sales via closed lots (BoI FX rates) ...")
    gains_ils, losses_ils, gross_sales_ils = parse_closed_lots_ils(text)
    print(f"  Gains: {gains_ils}  Losses: {losses_ils}  Gross sales: {gross_sales_ils}")

    rows = build_summary_rows(
        interest_ils=round(interest_exact),
        dividends_by_country={k: round(v) for k, v in dividends_by_country.items()},
        dividend_total=round(dividend_total_exact),
        div_plus_income=round(div_plus_income_exact),
        wht_ils=wht_ils,
        gains_ils=gains_ils,
        losses_ils=losses_ils,
        gross_sales_ils=gross_sales_ils,
    )

    # Append EquatePlus section if tax_summary_{year}.xlsx exists
    equate_summary_path = f"tax_summary_{year}.xlsx"
    if os.path.exists(equate_summary_path):
        print(f"\nFound {equate_summary_path}, adding EquatePlus + combined totals ...")
        eq_df = pd.read_excel(equate_summary_path)
        total_row = eq_df[eq_df["Sale date"] == "TOTAL"].iloc[0]
        eq_gains  = round(float(total_row["Total_gain_shekel"]))
        eq_losses = round(float(total_row["Total_loss_shekel"]))
        eq_gross  = round(float(total_row["Total_gross_sale_shekel"]))

        combined_gains  = gains_ils + eq_gains
        combined_losses = losses_ils + eq_losses
        combined_gross  = gross_sales_ils + eq_gross

        rows += [
            {"Label": "", "Value (ILS)": ""},
            {"Label": "EquatePlus positive realized gains", "Value (ILS)": eq_gains},
            {"Label": "EquatePlus realized losses",         "Value (ILS)": eq_losses},
            {"Label": "EquatePlus gross sale value",        "Value (ILS)": eq_gross},
            {"Label": "", "Value (ILS)": ""},
            {"Label": "TOTAL positive realized gains (IBKR + EquatePlus)",  "Value (ILS)": combined_gains},
            {"Label": "TOTAL realized losses (IBKR + EquatePlus)",          "Value (ILS)": combined_losses},
            {"Label": "TOTAL gross sale value (IBKR + EquatePlus)",         "Value (ILS)": combined_gross},
        ]
        print(f"  EquatePlus gains: {eq_gains}  losses: {eq_losses}  gross: {eq_gross}")
        print(f"  Combined gains: {combined_gains}  losses: {combined_losses}  gross: {combined_gross}")

    out = f"tax_ibkr_summary_{year}.xlsx"
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="IBKR Summary")
        ws = writer.sheets["IBKR Summary"]
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 18

    print(f"\nWrote {out}")


if __name__ == "__main__":
    main()
