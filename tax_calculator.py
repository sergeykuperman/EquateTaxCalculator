#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import glob
import os
import re
import sys

import pandas as pd

from tax_utils import TAX_RATE, _round_ils, boi_ils
from equate_parser import parse_sale_pdf, parse_equateplus_pair
from capital_gains import build_equate_lots


# ─── UNIT TESTS (Circular 10/2025 examples, no fees) ─────────────────────────
def _run_tests():
    from tax_utils import moses_gain_loss

    def check(label, orig, net_sale, fx_buy, fx_sell, exp_gain, exp_loss):
        gain, loss = moses_gain_loss(orig, net_sale, fx_buy, fx_sell)
        ok = abs(gain - exp_gain) < 0.01 and abs(loss - exp_loss) < 0.01
        status = "PASS" if ok else "FAIL"
        print(f"  [{status}] {label}: gain={gain:.2f} loss={loss:.2f} "
              f"(expected gain={exp_gain} loss={exp_loss})")
        if not ok:
            raise AssertionError(f"Test failed: {label}")

    print("Running Moses algorithm unit tests ...")
    check("FX up, price up",        3400, 4440, 3.4, 3.7, exp_gain=740, exp_loss=0)
    check("FX up, price down (loss)", 3400, 2960, 3.4, 3.7, exp_gain=0, exp_loss=440)
    check("FX up big, price down (nominal gain, all inflationary)",
          3400, 4000, 3.4, 5.0, exp_gain=0, exp_loss=0)
    check("FX down, price up",      3700, 4080, 3.7, 3.4, exp_gain=380, exp_loss=0)
    check("FX down, price down",    3700, 3060, 3.7, 3.4, exp_gain=0,   exp_loss=340)
    print("All tests passed.")


# ─── PROCESS ONE CSV + ITS MATCHING SALE.PDF ─────────────────────────────────
def process_pair(csv_path):
    m = re.search(r'consumption_(\d+\.\d+\.\d{4})\.csv$', csv_path)
    if not m:
        print(f"Skipping unrecognized file: {csv_path}")
        return
    date_key = m.group(1)
    sale_pdf = f"sale_{date_key}.pdf"
    if not os.path.exists(sale_pdf):
        print(f"Warning: no {sale_pdf} for {csv_path}, skipping")
        return

    sale_price, execution_date, settlement_date, _pdf_ex_rate, fees_euro = parse_sale_pdf(sale_pdf)
    fx_sell = boi_ils("EUR", execution_date.date())

    print(f"[{date_key}] sale_price={sale_price}€, execution={execution_date.date()}, "
          f"settle={settlement_date.date() if settlement_date else 'N/A'}, "
          f"BoI FX(sell)={fx_sell}, fees={fees_euro}€")

    raw_lots = parse_equateplus_pair(csv_path, sale_pdf)

    for lot in raw_lots:
        acq_date = lot["acquisition_date"]
        print(f"  Acq {acq_date}: BoI FX(buy) = {boi_ils('EUR', acq_date)}")

    calc_lots = build_equate_lots(raw_lots)

    # Build per-lot DataFrame for the Data sheet
    data_rows = []
    for raw, calc in zip(raw_lots, calc_lots):
        data_rows.append({
            "Acquisition date":       raw["acquisition_date"],
            "Consumption":            raw["quantity"],
            "Purchase price":         raw["cost_per_share_fc"],
            "FX_acq":                 calc["fx_buy"],
            "original_cost_shekel":   calc["original_cost_ils_exact"],
            "gross_sale_shekel":      calc["gross_sale_proceeds_ils_exact"],
            "sale_fee_shekel_lot":    calc["sale_commission_ils_exact"],
            "net_sale_shekel":        calc["net_sale_consideration_ils_exact"],
            "adjusted_cost_shekel":   calc["adjusted_cost_ils_exact"],
            "taxable_gain_shekel":    calc["taxable_gain_ils_exact"],
            "deductible_loss_shekel": calc["deductible_loss_ils_exact"],
        })
    lot_df = pd.DataFrame(data_rows)

    fees_shekels            = fees_euro * fx_sell
    total_gross_sale_shekel = sum(r["gross_sale_proceeds_ils_exact"] for r in calc_lots)
    total_taxable_gain      = sum(r["taxable_gain_ils_exact"] for r in calc_lots)
    total_deductible_loss   = sum(r["deductible_loss_ils_exact"] for r in calc_lots)
    total_net_gain          = total_taxable_gain - total_deductible_loss
    total_tax_to_pay        = max(0.0, total_net_gain) * TAX_RATE

    out = csv_path.replace(".csv", "_with_calc.xlsx")
    with pd.ExcelWriter(out) as writer:
        lot_df.to_excel(writer, sheet_name="Data", index=False)
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

    years = sorted({r["Sale date"].split(".")[-1] for r in rows})
    if len(years) > 1:
        raise RuntimeError(
            f"Input files span multiple years: {years}. "
            f"Move files for each year to a separate directory and run tax_calculator.py there."
        )

    summary_df = pd.DataFrame(rows).sort_values("Sale date")
    total_taxable_gain    = summary_df["Total_taxable_gain"].sum()
    total_deductible_loss = summary_df["Total_deductible_loss"].sum()
    net_annual_gain       = total_taxable_gain - total_deductible_loss
    total_row = {
        "Sale date":               "TOTAL",
        "Fees_shekels":            _round_ils(summary_df["Fees_shekels"].sum()),
        "Total_gross_sale_shekel": _round_ils(summary_df["Total_gross_sale_shekel"].sum()),
        "Total_taxable_gain":      _round_ils(total_taxable_gain),
        "Total_deductible_loss":   _round_ils(total_deductible_loss),
        "Total_net_gain":          _round_ils(net_annual_gain),
        "Total_tax_to_pay":        _round_ils(max(0.0, net_annual_gain) * TAX_RATE),
    }
    summary_df = pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)

    out = "tax_summary_" + "_".join(years) + ".xlsx"
    summary_df.to_excel(out, index=False)
    print(f"\nWrote combined summary: {out}")


if __name__ == "__main__":
    if "--test" in sys.argv:
        _run_tests()
    else:
        main()
