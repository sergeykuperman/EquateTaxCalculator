#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import glob
import os
import re
from datetime import datetime

import pdfplumber
import pandas as pd
import requests
import xml.etree.ElementTree as ET
from io import StringIO

# ─── CONSTANT ────────────────────────────────────────────────────────────────
TAX_RATE = 0.25  # 25%

# ─── PARSE SALE PARAMETERS FROM PDF ────────────────────────────────────────────
def parse_sale_pdf(path):
    """
    Extract SALE_PRICE (€), Settlement date, EX_RATE (ILS/€), FEES_EURO from sale_*.pdf.
    """
    with pdfplumber.open(path) as pdf:
        text = pdf.pages[0].extract_text()

    # SALE_PRICE: first number + "EUR" after "Quantity - Shares"
    m = re.search(r"Quantity\s*-\s*Shares[\s\S]*?([\d.,]+)\s*(?:€|EUR)", text)
    sale_price = float(m.group(1).replace(",", "")) if m else None

    # Settlement date: e.g. "Settlement date: 10 Jul 2025"
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

    if None in (sale_price, settlement_date, ex_rate, fees_euro):
        raise RuntimeError(f"Failed to parse all sale params from {path}:\n"
                           f" sale_price={sale_price}, "
                           f"settle={settlement_date}, "
                           f"FX={ex_rate}, "
                           f"fees={fees_euro}")
    return sale_price, settlement_date, ex_rate, fees_euro

# ─── FETCH CPI VIA CBS SDMX API ───────────────────────────────────────────────
def fetch_cpi_series(start, end):
    """
    Fetch Israel monthly CPI (PCPI_IX) via CBS SDMX XML,
    period start → end inclusive, returning a pandas.Series
    indexed by the first of each month.
    """
    url = "https://apis.cbs.gov.il/sdmx/data/IMF/ECOFIN_CPI/1"
    params = {
        "startPeriod": start.strftime("%m-%Y"),
        "endPeriod":   end.strftime("%m-%Y"),
        "format":      "xml",
        "download":    "false",
        "addNull":     "false",
    }
    # pull raw XML
    resp = requests.get(url, params=params)
    resp.raise_for_status()
    xml_root = ET.fromstring(resp.content)

    # find all Obs nodes, regardless of namespace
    obs_nodes = xml_root.findall(".//{*}Obs")
    dates, values = [], []
    for obs in obs_nodes:
        tp = obs.attrib.get("TIME_PERIOD") or obs.attrib.get("TIME")  # some variants
        ov = obs.attrib.get("OBS_VALUE") or obs.attrib.get("OBS")     # some variants
        if not (tp and ov):
            continue

        # parse e.g. "2025-07" or "2025-07-01" to Timestamp("2025-07-01")
        ts = pd.to_datetime(tp if "-" in tp else tp.replace("M", "-"), format="%Y-%m", errors="coerce")
        if pd.isna(ts):
            continue

        dates.append(ts)
        values.append(float(ov))

    if not dates:
        raise RuntimeError("No CPI observations found in the XML response")

    # build and return series
    return pd.Series(data=values, index=dates).sort_index()

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

    # parse that PDF
    sale_price, settlement_date, ex_rate, fees_euro = parse_sale_pdf(sale_pdf)
    print(f"[{date_key}] sale_price={sale_price}€, settle={settlement_date.date()}, FX={ex_rate}, fees={fees_euro}€")

 # load the CSV
    df = pd.read_csv(
        csv_path,
        sep=";",
        decimal=",",
        parse_dates=["Acquisition date"],
        dayfirst=True,
    )

    # fetch CPI series
    start_date = df["Acquisition date"].min()
    cpi = fetch_cpi_series(start_date, settlement_date)

    # map acquisition → CPI
    df["CPI_acq"] = (
        df["Acquisition date"]
          .dt.to_period("M")
          .dt.to_timestamp()
          .map(cpi)
    )

    # ① Check for missing acquisition‐date CPI
    missing_acq = df.loc[df["CPI_acq"].isna(), "Acquisition date"]
    if not missing_acq.empty:
        dates = ", ".join(d.strftime("%Y-%m-%d") for d in missing_acq.unique())
        raise RuntimeError(f"Missing CPI value for acquisition date(s): {dates}")

    # lookup settlement CPI
    cpi_set = cpi.get(settlement_date.replace(day=1))

    # ② Check for missing settlement CPI
    if pd.isna(cpi_set):
        sd = settlement_date.strftime("%Y-%m-%d")
        raise RuntimeError(f"Missing CPI value for settlement date: {sd}")

    df["CPI_set"] = cpi_set

    # do the math
    df["CPI_MULTIPLIER"]  = df["CPI_set"] / df["CPI_acq"]
    fees_shekels     = fees_euro * ex_rate
    df["gross_sale_euro"]  = df["Consumption"] * sale_price
    df["cost_euro"]        = df["Consumption"] * df["Purchase price"] * df["CPI_MULTIPLIER"]
    df["real_gain_shekel"] = (df["gross_sale_euro"] - df["cost_euro"]) * ex_rate
    df["tax_to_pay"]       = df["real_gain_shekel"] * TAX_RATE

    # now compute totals
    total_real_gain        = df["real_gain_shekel"].sum()
    total_tax_before_fees  = df["tax_to_pay"].sum()
    total_tax_to_pay       = total_tax_before_fees - fees_shekels
    
    # write Data + Summary into two sheets
    out = csv_path.replace(".csv", "_with_calc.xlsx")
    with pd.ExcelWriter(out) as writer:
        df.to_excel(writer, sheet_name="Data", index=False)

        summary = pd.DataFrame([{
            "Fees_shekels":           fees_shekels,
            "Total_real_gain_shekel": total_real_gain,
            "Total_tax_before_fees":  total_tax_before_fees,
            "Total_tax_to_pay":       total_tax_to_pay,
        }])
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print(f"Wrote {out} (with summary sheet)")

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    for csv_file in glob.glob("consumption_*.csv"):
        process_pair(csv_file)

if __name__ == "__main__":
    main()
