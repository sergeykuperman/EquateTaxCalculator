#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import glob
import os
import re
import sys
from datetime import datetime
import datetime as dt

import pdfplumber
import pandas as pd
import requests

# ─── CONSTANT ────────────────────────────────────────────────────────────────
TAX_RATE = 0.25  # 25%


# ─── MOSES / FORM 1325 CAPITAL GAIN ALGORITHM ────────────────────────────────
def moses_gain_loss(
    original_cost_ils: float,
    net_sale_ils: float,
    fx_buy: float,
    fx_sell: float,
) -> tuple[float, float]:
    """
    Israeli capital-gain/loss calculation for foreign-currency securities per
    ITO Section 91(b), Form 1325 (2025), and Circular 10/2025 (Moses rule).

    Returns (taxable_gain, deductible_loss) both >= 0.

    The FX rate acts as an index (מדד). When FX rises, the inflationary component
    is exempt. When FX falls, the negative inflationary component is treated as
    zero — it cannot enlarge a gain or convert a loss into a deductible amount.
    """
    adjusted_cost_ils = original_cost_ils * (fx_sell / fx_buy)
    nominal_result = net_sale_ils - original_cost_ils

    if nominal_result > 0:
        if fx_sell >= fx_buy:
            inflationary = adjusted_cost_ils - original_cost_ils  # >= 0
            exempt = min(nominal_result, max(0.0, inflationary))
            taxable_gain = nominal_result - exempt
        else:
            # FX fell: negative inflationary component treated as zero — no exemption
            taxable_gain = nominal_result
        deductible_loss = 0.0

    elif nominal_result < 0:
        taxable_gain = 0.0
        nominal_loss = -nominal_result
        if fx_sell >= fx_buy:
            # FX rise cannot enlarge the loss; full nominal loss is deductible
            deductible_loss = nominal_loss
        else:
            # Remove the portion of the loss attributable to FX decline
            neg_inflationary = original_cost_ils - adjusted_cost_ils  # >= 0
            deductible_loss = max(0.0, nominal_loss - neg_inflationary)
            # equivalent: max(0.0, adjusted_cost_ils - net_sale_ils)

    else:
        taxable_gain = 0.0
        deductible_loss = 0.0

    return taxable_gain, deductible_loss


# ─── UNIT TESTS (Circular 10/2025 examples, no fees) ─────────────────────────
def _run_tests():
    def check(label, orig, net_sale, fx_buy, fx_sell, exp_gain, exp_loss):
        gain, loss = moses_gain_loss(orig, net_sale, fx_buy, fx_sell)
        ok = abs(gain - exp_gain) < 0.01 and abs(loss - exp_loss) < 0.01
        status = "PASS" if ok else "FAIL"
        print(f"  [{status}] {label}: gain={gain:.2f} loss={loss:.2f} "
              f"(expected gain={exp_gain} loss={exp_loss})")
        if not ok:
            raise AssertionError(f"Test failed: {label}")

    print("Running Moses algorithm unit tests ...")

    # 1. Price rises, FX rises → inflationary component exempt
    # buy $1000 @ 3.4 → cost=3400; sell $1200 @ 3.7 → net=4440
    # adjusted_cost=3700; nominal=1040; inflationary=300; taxable=740
    check("FX up, price up", 3400, 4440, 3.4, 3.7, exp_gain=740, exp_loss=0)

    # 2. Price falls, FX rises → full nominal loss deductible
    # buy $1000 @ 3.4 → cost=3400; sell $800 @ 3.7 → net=2960; nominal=-440
    check("FX up, price down (loss)", 3400, 2960, 3.4, 3.7, exp_gain=0, exp_loss=440)

    # 2b. Price falls, FX rises sharply, FX gain > price loss → zero result
    # buy $1000 @ 3.4 → cost=3400; sell $800 @ 5.0 → net=4000; nominal=+600
    # adjusted_cost=5000; inflationary=1600; exempt=min(600,1600)=600; taxable=0
    check("FX up big, price down (nominal gain, all inflationary)", 3400, 4000, 3.4, 5.0, exp_gain=0, exp_loss=0)

    # 3. Price rises, FX falls → full nominal gain taxable (no exemption)
    # buy $1000 @ 3.7 → cost=3700; sell $1200 @ 3.4 → net=4080; nominal=+380
    check("FX down, price up", 3700, 4080, 3.7, 3.4, exp_gain=380, exp_loss=0)

    # 4. Price falls, FX falls → partial loss deductible
    # buy $1000 @ 3.7 → cost=3700; sell $900 @ 3.4 → net=3060; nominal=-640
    # adjusted_cost=3400; neg_inflationary=300; deductible=max(0,640-300)=340
    check("FX down, price down", 3700, 3060, 3.7, 3.4, exp_gain=0, exp_loss=340)

    print("All tests passed.")


# ─── PARSE SALE PARAMETERS FROM PDF ──────────────────────────────────────────
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

    # EX_RATE: 5-decimal number after "Foreign exchange"
    m = re.search(r"Foreign exchange[\s\S]*?(\d+\.\d{5})", text)
    ex_rate = float(m.group(1)) if m else None

    # FEES_EURO: "Total debits" line with a two-decimal euro amount
    m = re.search(r"Total debits[\s\S]*?(\d+\.\d{2})\s*(?:€|EUR)", text)
    fees_euro = float(m.group(1)) if m else None

    if None in (sale_price, execution_date, ex_rate, fees_euro):
        raise RuntimeError(
            f"Failed to parse all sale params from {path}:\n"
            f" sale_price={sale_price}, execution={execution_date}, "
            f"FX={ex_rate}, fees={fees_euro}\n"
            f"--- page text ---\n{text}"
        )
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
    fx_sell = boi_eur_ils(execution_date.date())
    print(f"[{date_key}] sale_price={sale_price}€, execution={execution_date.date()}, "
          f"settle={settlement_date.date() if settlement_date else 'N/A'}, "
          f"BoI FX(sell)={fx_sell}, fees={fees_euro}€")

    # load the CSV
    df = pd.read_csv(csv_path, sep=";", decimal=",")
    df["Acquisition date"] = pd.to_datetime(df["Acquisition date"], format="%d %b %Y")

    # fetch BoI EUR/ILS rate for each unique acquisition date
    unique_acq_dates = df["Acquisition date"].dt.date.unique()
    fx_cache = {}
    for d in unique_acq_dates:
        fx_cache[d] = boi_eur_ils(d)
        print(f"  Acq {d}: BoI FX(buy) = {fx_cache[d]}")

    df["FX_acq"] = df["Acquisition date"].dt.date.map(fx_cache)

    # Per Form 1325:
    #   original_cost = purchase value (no acquisition commission for RSU grants)
    #   net_sale = gross_sale - allocated sale commission
    #   gross_sale is kept separately as turnover (מחזור מכירות) for Form 1322

    fees_shekels = fees_euro * fx_sell
    total_shares = df["Consumption"].sum()

    df["original_cost_shekel"] = df["Consumption"] * df["Purchase price"] * df["FX_acq"]
    df["gross_sale_shekel"]    = df["Consumption"] * sale_price * fx_sell
    df["sale_fee_shekel_lot"]  = fees_shekels * (df["Consumption"] / total_shares)
    df["net_sale_shekel"]      = df["gross_sale_shekel"] - df["sale_fee_shekel_lot"]
    df["adjusted_cost_shekel"] = df["original_cost_shekel"] * (fx_sell / df["FX_acq"])

    # Apply Moses 4-case algorithm per lot
    gains, losses = zip(*df.apply(
        lambda r: moses_gain_loss(
            r["original_cost_shekel"],
            r["net_sale_shekel"],
            r["FX_acq"],
            fx_sell,
        ),
        axis=1,
    ))
    df["taxable_gain_shekel"]    = list(gains)
    df["deductible_loss_shekel"] = list(losses)

    # Totals
    total_gross_sale_shekel = df["gross_sale_shekel"].sum()
    total_taxable_gain      = df["taxable_gain_shekel"].sum()
    total_deductible_loss   = df["deductible_loss_shekel"].sum()
    total_net_gain          = total_taxable_gain - total_deductible_loss
    total_tax_to_pay        = max(0.0, total_net_gain) * TAX_RATE

    # write Data + Summary into two sheets
    out = csv_path.replace(".csv", "_with_calc.xlsx")
    with pd.ExcelWriter(out) as writer:
        df.to_excel(writer, sheet_name="Data", index=False)

        summary = pd.DataFrame([{
            "Fees_shekels":             fees_shekels,
            "Total_gross_sale_shekel":  total_gross_sale_shekel,
            "Total_taxable_gain":       total_taxable_gain,
            "Total_deductible_loss":    total_deductible_loss,
            "Total_net_gain":           total_net_gain,
            "Total_tax_to_pay":         total_tax_to_pay,
        }])
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print(f"Wrote {out} (with summary sheet)")

    return {
        "Sale date":               date_key,
        "Fees_shekels":            fees_shekels,
        "Total_gross_sale_shekel": total_gross_sale_shekel,
        "Total_taxable_gain":      total_taxable_gain,
        "Total_deductible_loss":   total_deductible_loss,
        "Total_net_gain":          total_net_gain,
        "Total_tax_to_pay":        total_tax_to_pay,
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
    total_taxable_gain    = summary_df["Total_taxable_gain"].sum()
    total_deductible_loss = summary_df["Total_deductible_loss"].sum()
    net_annual_gain       = total_taxable_gain - total_deductible_loss
    total_row = {
        "Sale date":               "TOTAL",
        "Fees_shekels":            summary_df["Fees_shekels"].sum(),
        "Total_gross_sale_shekel": summary_df["Total_gross_sale_shekel"].sum(),
        "Total_taxable_gain":      total_taxable_gain,
        "Total_deductible_loss":   total_deductible_loss,
        "Total_net_gain":          net_annual_gain,
        "Total_tax_to_pay":        max(0.0, net_annual_gain) * TAX_RATE,
    }
    summary_df = pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)

    years = sorted({r["Sale date"].split(".")[-1] for r in rows})
    out = "tax_summary_" + "_".join(years) + ".xlsx"
    summary_df.to_excel(out, index=False)
    print(f"\nWrote combined summary: {out}")


if __name__ == "__main__":
    if "--test" in sys.argv:
        _run_tests()
    else:
        main()
