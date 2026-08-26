#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import glob
import os
import re
from datetime import datetime
import datetime as dt

import pdfplumber
import pandas as pd
import requests

# ─── CONSTANT ────────────────────────────────────────────────────────────────
TAX_RATE = 0.25  # 25%

# ─── PARSE SALE PARAMETERS FROM PDF ────────────────────────────────────────────
def parse_sale_pdf(path):
    """
    Extract SALE_PRICE (€), execution date, settlement date, EX_RATE (ILS/€), FEES_EURO from sale_*.pdf.
    """
    with pdfplumber.open(path) as pdf:
        text = pdf.pages[0].extract_text()

    # SALE_PRICE: first number + "EUR" after "Quantity - Shares"
    m = re.search(r"Quantity\s*-\s*Shares[\s\S]*?([\d.,]+)\s*(?:€|EUR)", text)
    sale_price = float(m.group(1).replace(",", "")) if m else None

    # Execution date (trade/sale date for FX purposes): e.g. "Execution date: 1 Jul 2024 09:00:00 CET"
    m = re.search(r"Execution date:\s*([\d]{1,2}\s+[A-Za-z]+\s+\d{4})", text)
    execution_date = datetime.strptime(m.group(1), "%d %b %Y") if m else None

    # Settlement date (kept for reference): e.g. "Settlement date: 3 Jul 2024"
    m = re.search(r"Settlement date:\s*([\d]{1,2}\s+[A-Za-z]+\s+\d{4})", text)
    settlement_date = datetime.strptime(m.group(1), "%d %b %Y") if m else None

    # EX_RATE: look for number with 5 decimals after "Foreign exchange"
    m = re.search(r"Foreign exchange[\s\S]*?(\d+\.\d{5})", text)
    if not m:
        # fallback to any 5-decimal number on the page
        m = re.search(r"\b(\d+\.\d{5})\b", text)
    ex_rate = float(m.group(1)) if m else None

    # FEES_EURO: look for "Total debits" line with a two-decimal euro amount
    m = re.search(r"Total debits[\s\S]*?(\d+\.\d{2})\s*(?:€|EUR)", text)
    if not m:
        # fallback to any two-decimal number under 100
        m = re.search(r"\b([0-9]{1,2}\.\d{2})\b", text)
    fees_euro = float(m.group(1)) if m else None

    if None in (sale_price, execution_date, ex_rate, fees_euro):
        raise RuntimeError(f"Failed to parse all sale params from {path}:\n"
                           f" sale_price={sale_price}, "
                           f"execution={execution_date}, "
                           f"FX={ex_rate}, "
                           f"fees={fees_euro}")
    return sale_price, execution_date, settlement_date, ex_rate, fees_euro

# ─── BANK OF ISRAEL EUR/ILS RATE ─────────────────────────────────────────────
def boi_eur_ils(day: str | dt.date, *, live: bool = False) -> float:
    """
    Bank-of-Israel EUR→ILS representative rate for a specific day,
    or the current screen rate when live=True.
    """
    if isinstance(day, dt.date):
        day = day.strftime("%Y-%m-%d")

    # ---- 1. live screen quote --------------------------------------------
    if live:
        js = requests.get("https://www.boi.org.il/PublicApi/GetExchangeRates",
                          timeout=10).json()
        items = js["exchangeRates"]
        rec   = next(i for i in items if i["key"] == "EUR")
        return float(rec["currentExchangeRate"]) / rec.get("unit", 1)

    # ---- 2. historical representative quote via SDMX ---------------------
    # Walk backwards up to 7 days to skip weekends and Israeli holidays
    url = ("https://edge.boi.gov.il/FusionEdgeServer/"
           "sdmx/v2/data/dataflow/BOI.STATISTICS/EXR/1.0/RER_EUR_ILS")
    d = dt.date.fromisoformat(day)
    for delta in range(7):
        candidate = (d - dt.timedelta(days=delta)).strftime("%Y-%m-%d")
        params = {
            "c[DATA_TYPE]": "OF00",
            "startperiod": candidate,
            "endperiod":   candidate,
            "format": "sdmx-json",
        }
        resp = requests.get(url, params=params, timeout=10)
        if resp.status_code != 200 or not resp.text.strip():
            continue
        try:
            js = resp.json()
            obs = js["data"]["dataSets"][0]["series"]["0:0:0:0:0:0"]["observations"]
            return float(list(obs.values())[0][0])
        except (KeyError, IndexError, ValueError):
            continue
    raise RuntimeError(f"No BoI EUR/ILS rate found within 7 days before {day}")

# ─── PROCESS ONE CSV + ITS MATCHING SALE.PDF ─────────────────────────────────
def process_pair(csv_path):
    # match the date key in the filename
    m = re.search(r'consumption_(\d+\.\d+\.\d{4})\.csv$', csv_path)
    if not m:
        print(f"Skipping unrecognized file: {csv_path}")
        return
    date_key = m.group(1)               # e.g. "8.7.2025"
    sale_pdf = f"sale_{date_key}.pdf"
    if not os.path.exists(sale_pdf):
        print(f"Warning: no {sale_pdf} for {csv_path}, skipping")
        return

    # parse that PDF, then fetch BoI FX rate on the execution (trade) date
    sale_price, execution_date, settlement_date, _pdf_ex_rate, fees_euro = parse_sale_pdf(sale_pdf)
    fx_set = boi_eur_ils(execution_date.date())
    print(f"[{date_key}] sale_price={sale_price}€, execution={execution_date.date()}, settle={settlement_date.date() if settlement_date else 'N/A'}, BoI FX={fx_set}, fees={fees_euro}€")

    # load the CSV
    df = pd.read_csv(
        csv_path,
        sep=";",
        decimal=",",
    )
    df["Acquisition date"] = pd.to_datetime(df["Acquisition date"], format="%d %b %Y")

    # fetch BoI EUR/ILS rate for each unique acquisition date
    # Under ITO Section 91(b), the exchange rate substitutes for CPI on foreign-currency assets:
    # the inflationary component = cost×(fx_set/fx_acq − 1) and is tax-exempt; only the real
    # gain above the FX movement is taxable. Both fx_set (execution date) and fx_acq use the
    # BoI representative rate, consistent with Form 1325 which asks for the sale date rate.
    unique_acq_dates = df["Acquisition date"].dt.date.unique()
    fx_cache = {}
    for d in unique_acq_dates:
        fx_cache[d] = boi_eur_ils(d)
        print(f"  Acq {d}: BoI FX = {fx_cache[d]}")

    df["FX_acq"] = df["Acquisition date"].dt.date.map(fx_cache)

    # compute gain: convert both legs to ILS using their respective period's BoI FX rate
    fees_shekels            = fees_euro * fx_set
    df["gross_sale_shekel"] = df["Consumption"] * sale_price * fx_set
    df["cost_shekel"]       = df["Consumption"] * df["Purchase price"] * df["FX_acq"]

    # allocate fees proportionally per lot by share count, deduct from gain before tax
    df["fees_shekel_lot"]  = fees_shekels * (df["Consumption"] / df["Consumption"].sum())
    df["real_gain_shekel"] = df["gross_sale_shekel"] - df["cost_shekel"] - df["fees_shekel_lot"]
    df["tax_to_pay"]       = df["real_gain_shekel"].clip(lower=0) * TAX_RATE

    # totals
    total_gross_sale_shekel = df["gross_sale_shekel"].sum()
    total_gain              = df.loc[df["real_gain_shekel"] > 0, "real_gain_shekel"].sum()
    total_loss              = df.loc[df["real_gain_shekel"] < 0, "real_gain_shekel"].sum()
    total_real_gain         = total_gain + total_loss
    total_tax_to_pay        = max(0.0, total_real_gain) * TAX_RATE

    # write Data + Summary into two sheets
    out = csv_path.replace(".csv", "_with_calc.xlsx")
    with pd.ExcelWriter(out) as writer:
        df.to_excel(writer, sheet_name="Data", index=False)

        summary = pd.DataFrame([{
            "Fees_shekels":             fees_shekels,
            "Total_gross_sale_shekel":  total_gross_sale_shekel,
            "Total_gain_shekel":        total_gain,
            "Total_loss_shekel":        total_loss,
            "Total_real_gain_shekel":   total_real_gain,
            "Total_tax_to_pay":         total_tax_to_pay,
        }])
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print(f"Wrote {out} (with summary sheet)")

    return {
        "Sale date":                date_key,
        "Fees_shekels":             fees_shekels,
        "Total_gross_sale_shekel":  total_gross_sale_shekel,
        "Total_gain_shekel":        total_gain,
        "Total_loss_shekel":        total_loss,
        "Total_real_gain_shekel":   total_real_gain,
        "Total_tax_to_pay":         total_tax_to_pay,
    }

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    rows = []
    for csv_file in sorted(glob.glob("consumption_*.csv")):
        row = process_pair(csv_file)
        if row:
            rows.append(row)

    if not rows:
        return

    summary_df = pd.DataFrame(rows).sort_values("Sale date")
    net_annual_gain = summary_df["Total_real_gain_shekel"].sum()
    total_row = {
        "Sale date":                "TOTAL",
        "Fees_shekels":             summary_df["Fees_shekels"].sum(),
        "Total_gross_sale_shekel":  summary_df["Total_gross_sale_shekel"].sum(),
        "Total_gain_shekel":        summary_df["Total_gain_shekel"].sum(),
        "Total_loss_shekel":        summary_df["Total_loss_shekel"].sum(),
        "Total_real_gain_shekel":   net_annual_gain,
        "Total_tax_to_pay":         max(0.0, net_annual_gain) * TAX_RATE,
    }
    summary_df = pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)

    years = sorted({r["Sale date"].split(".")[-1] for r in rows})
    out = "tax_summary_" + "_".join(years) + ".xlsx"
    summary_df.to_excel(out, index=False)
    print(f"\nWrote combined summary: {out}")

if __name__ == "__main__":
    main()
