#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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


def parse_interest_ils(pdf_path: str) -> int:
    # Interest section is always on the last page, after the Dividends section.
    with pdfplumber.open(pdf_path) as pdf:
        last_page_text = pdf.pages[-1].extract_text() or ""
    # Find the Interest section header then the "Total in ILS" that follows it
    m = re.search(
        r"\bInterest\b\s*\nDate Description Amount\b.*?Total in ILS\s+([\d,]+\.\d+)",
        last_page_text,
        re.DOTALL,
    )
    if m:
        return round(float(m.group(1).replace(",", "")))
    raise RuntimeError("Could not parse interest ILS total from PDF")


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
    return {country: round(total) for country, total in sorted(by_country.items())}


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


def parse_gross_sales_ils(text: str) -> int:
    # Use "Total SYMBOL" rows: negative net qty = net sell position.
    # We need the sale date; find the last trade date for each sold symbol.
    # All sells in this statement are on 2024-01-31, but parse dynamically.

    # Step 1: find all trade dates per symbol (date appears before the first trade row)
    # Step 2: find Total rows with negative qty (net sells)
    # Step 3: use the date of the last close trade for that symbol

    # Build a map: symbol -> most recent sell date
    close_date_pattern = re.compile(
        r"(\d{4}-\d{2}-\d{2}),\s*\n"
        r"([A-Z]+)\s+\S+\s+-\d+\s+[\d.]+\s+[\d.]+\s+[\d,]+\.\d+\s+"
        r"[-\d.]+\s+[-\d,]+\.\d+\s+[-\d,]+\.\d+\s+[-\d,]+\.\d+\s+C",
        re.MULTILINE,
    )
    symbol_date = {}
    for m in close_date_pattern.finditer(text):
        symbol_date[m.group(2)] = m.group(1)  # last date wins

    # Find Total SYMBOL rows with negative net qty (net sell)
    total_pattern = re.compile(
        r"^Total ([A-Z]+)\s+(-\d[\d,]*)\s+([\d,]+\.\d+)",
        re.MULTILINE,
    )
    total_ils = 0.0
    for m in total_pattern.finditer(text):
        symbol, qty_str, proceeds_str = m.group(1), m.group(2), m.group(3)
        qty = int(qty_str.replace(",", ""))
        if qty >= 0:
            continue
        proceeds_usd = float(proceeds_str.replace(",", ""))
        date_str = symbol_date.get(symbol)
        if not date_str:
            raise RuntimeError(f"No sell date found for symbol {symbol}")
        rate = boi_ils("USD", date_str)
        total_ils += proceeds_usd * rate
        print(f"  {date_str} {symbol} sell qty={qty} proceeds=${proceeds_usd:.2f} @ {rate:.4f} = {proceeds_usd*rate:.2f} ILS")

    if total_ils == 0:
        raise RuntimeError("No stock sell trades found in PDF")
    return round(total_ils)


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

    print("Parsing interest ...")
    interest_ils = parse_interest_ils(pdf_path)
    print(f"  Interest in ILS: {interest_ils}")

    print("Parsing dividends by country ...")
    dividends_by_country = parse_dividends_by_country(pdf_path)
    dividend_total = sum(dividends_by_country.values())
    div_plus_income = dividend_total + interest_ils
    print(f"  Dividends by country: {dividends_by_country}")
    print(f"  Dividend total: {dividend_total}")

    print("Parsing withholding tax ...")
    wht_ils = parse_withholding_tax_ils(text)
    print(f"  WHT in ILS: {wht_ils}")

    print("Parsing realized stock gains/losses ...")
    gains_ils, losses_ils = parse_realized_stocks_ils(text)
    print(f"  Gains: {gains_ils}  Losses: {losses_ils}")

    print("Parsing gross sales ...")
    gross_sales_ils = parse_gross_sales_ils(text)
    print(f"  Gross sales in ILS: {gross_sales_ils}")

    rows = build_summary_rows(
        interest_ils=interest_ils,
        dividends_by_country=dividends_by_country,
        dividend_total=dividend_total,
        div_plus_income=div_plus_income,
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
