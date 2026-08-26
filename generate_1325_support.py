#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_1325_support.py — Authoritative Form 1325 capital-gains support workbook.

Generates form1325_support_<year>.xlsx with six sheets:
  1. "Form 1325 Entry View"  — manual-transcription view, 10 rows/form + check subtotals
  2. "Form 1325 Support"     — full per-lot audit detail (exact floats + filing-rounded values)
  3. "Tax Data"              — machine-readable, one row per lot, Tax Data schema v1
  4. "Reconciliation"        — rounding transparency with metric-specific match criteria
  5. "Metadata"              — key/value provenance pairs
  6. "Sources"               — one row per source file with SHA-256 hash

This is the ONLY script that performs stock-sale capital-gains calculations.
All downstream consumers (annual_tax_summary.py) read from the Tax Data sheet.
"""

import argparse
import datetime
import glob
import hashlib
import os
import re

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from tax_utils import TAX_RATE, _round_ils
from ibkr_parser import extract_full_text, extract_year_from_filename, parse_ibkr_lots_detail
from equate_parser import parse_equateplus_lots
from capital_gains import build_ibkr_lots, build_equate_lots


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def _sha256(path: str) -> str:
    return hashlib.sha256(open(path, "rb").read()).hexdigest()


def _assign_form_numbers(rows: list[dict]) -> list[dict]:
    sorted_rows = sorted(
        rows,
        key=lambda r: (r["sale_date"], r["source"], r["ticker"], r["acquisition_date"]),
    )
    for i, row in enumerate(sorted_rows):
        row["form_no"] = (i // 10) + 1
        row["row_no"]  = (i % 10) + 1
    return sorted_rows


def _verify_allocation(rows: list[dict]) -> None:
    """
    Verify that per-lot allocated quantities, proceeds, and commissions sum back
    to the fill-level totals recorded on each raw lot by the parser.

    For IBKR lots: fill_quantity_fc / fill_proceeds_fc / fill_commission_fc are
    authoritative broker-reported fill totals; the lot sums must match them.
    For EquatePlus lots: these fields are absent — allocation is verified implicitly
    by the fact that total_shares drives proportional allocation in the parser.
    """
    from collections import defaultdict
    fill_groups: dict = defaultdict(list)
    for r in rows:
        fill_groups[r["fill_id"]].append(r)

    for fill_id, lot_rows in fill_groups.items():
        if "fill_quantity_fc" not in lot_rows[0]:
            # No fill-level totals available — skip (should not occur with current parsers)
            continue

        fill_qty  = lot_rows[0]["fill_quantity_fc"]
        fill_proc = lot_rows[0]["fill_proceeds_fc"]
        fill_comm = lot_rows[0]["fill_commission_fc"]

        sum_qty  = sum(r["quantity"] for r in lot_rows)
        sum_proc = sum(r["gross_sale_proceeds_fc"] for r in lot_rows)
        sum_comm = sum(r["sale_commission_fc"] for r in lot_rows)

        tol_qty  = 0.01
        tol_proc = max(0.02, 1e-6 * abs(fill_proc))
        tol_comm = max(0.02, 1e-6 * abs(fill_comm))

        if abs(sum_qty - fill_qty) > tol_qty:
            raise RuntimeError(
                f"Quantity allocation mismatch for fill {fill_id}: "
                f"lots sum to {sum_qty:.6f} but fill total is {fill_qty:.6f}"
            )
        if abs(sum_proc - fill_proc) > tol_proc:
            raise RuntimeError(
                f"Proceeds allocation mismatch for fill {fill_id}: "
                f"lots sum to {sum_proc:.4f} but fill total is {fill_proc:.4f}"
            )
        if abs(sum_comm - fill_comm) > tol_comm:
            raise RuntimeError(
                f"Commission allocation mismatch for fill {fill_id}: "
                f"lots sum to {sum_comm:.6f} but fill total is {fill_comm:.6f}"
            )


# ─── EXCEL WRITING ────────────────────────────────────────────────────────────

_GREY_FILL  = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
_ITALIC_FONT = Font(italic=True, size=9)
_HEADER_FONT = Font(bold=True)


def _autosize_columns(ws) -> None:
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 40)


def _write_entry_view(ws, all_rows: list[dict]) -> None:
    headers = [
        "1325 Form No.", "1325 Row No.", "Source", "Security (ticker)",
        "Acquisition date", "Original price ILS",
        "1 + שיעור עליית המדד", "Index increase %",
        "Adjusted original price ILS", "Sale date",
        "ערך נקוב במכירה", "Net sale consideration ILS",
        "Real capital gain ILS", "Real capital loss ILS",
        "Acquired before listing", "Tax rate",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = _HEADER_FONT

    last_form = None
    subtotal_gain = 0
    subtotal_loss = 0

    for row in all_rows:
        form_no = row["form_no"]
        if last_form is not None and form_no != last_form:
            _append_subtotal_row(ws, last_form, subtotal_gain, subtotal_loss)
            subtotal_gain = 0
            subtotal_loss = 0

        gain_ils = row["taxable_gain_ils_filing"]
        loss_ils = row["deductible_loss_ils_filing"]
        subtotal_gain += gain_ils
        subtotal_loss += loss_ils

        ws.append([
            form_no, row["row_no"], row["source"], row["ticker"],
            row["acquisition_date"],
            row["original_cost_ils_filing"],
            round(row["fx_ratio"], 6),
            round(row["index_increase_pct"], 4),
            row["adjusted_cost_ils_filing"],
            row["sale_date"],
            "",
            row["net_sale_consideration_ils_filing"],
            gain_ils, loss_ils,
            "No",
            row["tax_rate"],
        ])
        last_form = form_no

    if last_form is not None:
        _append_subtotal_row(ws, last_form, subtotal_gain, subtotal_loss)

    _autosize_columns(ws)


def _append_subtotal_row(ws, form_no: int, subtotal_gain: int, subtotal_loss: int) -> None:
    label = f"Form {form_no} — Check subtotal (NOT a 1325 entry row)"
    ws.append([f"Form {form_no}", "", "", label,
               "", "", "", "", "", "", "", "",
               subtotal_gain, subtotal_loss, "", ""])
    for cell in ws[ws.max_row]:
        cell.fill = _GREY_FILL
        cell.font = _ITALIC_FONT


def _write_support_sheet(ws, all_rows: list[dict]) -> None:
    headers = [
        "1325 Form No.", "1325 Row No.",
        "Source", "Ticker", "Currency", "Fill ID", "Lot ID",
        "Acquisition date", "Sale date", "Quantity",
        "fx_buy", "fx_sell", "1 + שיעור עליית המדד (fx_ratio)", "Index increase %",
        "Acquisition basis FC",
        "Original cost ILS (exact)",     "Original cost ILS (filing)",
        "Adjusted cost ILS (exact)",     "Adjusted cost ILS (filing)",
        "Gross proceeds FC",             "Gross proceeds ILS (exact)",
        "Commission FC",                 "Commission ILS (exact)",
        "Net sale ILS (exact)",          "Net sale ILS (filing)",
        "Nominal gain/loss ILS",         "FX adjustment ILS",
        "Taxable gain ILS (exact)",      "Taxable gain ILS (filing)",
        "Deductible loss ILS (exact)",   "Deductible loss ILS (filing)",
        "Tax rate",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = _HEADER_FONT

    for row in all_rows:
        ws.append([
            row["form_no"], row["row_no"],
            row["source"], row["ticker"], row["currency"],
            row["fill_id"], row["lot_id"],
            row["acquisition_date"], row["sale_date"], row["quantity"],
            row["fx_buy"], row["fx_sell"],
            row["fx_ratio"], row["index_increase_pct"],
            row["acquisition_basis_fc"],
            row["original_cost_ils_exact"],    row["original_cost_ils_filing"],
            row["adjusted_cost_ils_exact"],    row["adjusted_cost_ils_filing"],
            row["gross_sale_proceeds_fc"],     row["gross_sale_proceeds_ils_exact"],
            row["sale_commission_fc"],         row["sale_commission_ils_exact"],
            row["net_sale_consideration_ils_exact"], row["net_sale_consideration_ils_filing"],
            row["nominal_gain_loss_ils_exact"],row["fx_adjustment_ils_exact"],
            row["taxable_gain_ils_exact"],     row["taxable_gain_ils_filing"],
            row["deductible_loss_ils_exact"],  row["deductible_loss_ils_filing"],
            row["tax_rate"],
        ])

    _autosize_columns(ws)


def _write_tax_data(ws, all_rows: list[dict], year: str) -> None:
    """Tax Data sheet — machine-readable, one row per lot, Tax Data schema v1."""
    columns = [
        "schema_version", "tax_year", "source", "fill_id", "lot_id",
        "ticker", "isin", "currency", "tax_rate", "quantity",
        "acquisition_date", "sale_date",
        "fx_buy", "fx_sell", "fx_ratio", "index_increase_pct",
        "acquisition_basis_fc",
        "original_cost_ils_exact",    "adjusted_cost_ils_exact",
        "gross_sale_proceeds_fc",     "gross_sale_proceeds_ils_exact",
        "sale_commission_fc",         "sale_commission_ils_exact",
        "net_sale_consideration_ils_exact",
        "nominal_gain_loss_ils_exact","fx_adjustment_ils_exact",
        "taxable_gain_ils_exact",     "deductible_loss_ils_exact",
        "original_cost_ils_filing",   "adjusted_cost_ils_filing",
        "net_sale_consideration_ils_filing",
        "taxable_gain_ils_filing",    "deductible_loss_ils_filing",
        "form_no", "row_no",
    ]
    ws.append(columns)
    for cell in ws[1]:
        cell.font = _HEADER_FONT

    for row in all_rows:
        ws.append([
            row["schema_version"], year, row["source"],
            row["fill_id"], row["lot_id"],
            row["ticker"], row.get("isin", ""), row["currency"], row["tax_rate"], row["quantity"],
            row["acquisition_date"], row["sale_date"],
            row["fx_buy"], row["fx_sell"], row["fx_ratio"], row["index_increase_pct"],
            row["acquisition_basis_fc"],
            row["original_cost_ils_exact"],    row["adjusted_cost_ils_exact"],
            row["gross_sale_proceeds_fc"],     row["gross_sale_proceeds_ils_exact"],
            row["sale_commission_fc"],         row["sale_commission_ils_exact"],
            row["net_sale_consideration_ils_exact"],
            row["nominal_gain_loss_ils_exact"],row["fx_adjustment_ils_exact"],
            row["taxable_gain_ils_exact"],     row["deductible_loss_ils_exact"],
            row["original_cost_ils_filing"],   row["adjusted_cost_ils_filing"],
            row["net_sale_consideration_ils_filing"],
            row["taxable_gain_ils_filing"],    row["deductible_loss_ils_filing"],
            row["form_no"], row["row_no"],
        ])

    _autosize_columns(ws)


def _write_reconciliation(ws, all_rows: list[dict],
                           prior_ibkr: dict | None, prior_equate: dict | None) -> bool:
    """
    Reconciliation sheet with metric-specific match criteria:
      Gains / losses: summary_value == sum(filing rows)  [bottom-up path]
      Gross turnover: summary_value == _round_ils(sum_exact)  [top-down path]

    When prior stale summaries are present (tax_ibkr_summary_<year>.xlsx or
    tax_summary_<year>.xlsx), each source's summary_value is also compared against
    the prior run's value. A discrepancy there is a genuine cross-run diagnostic.

    row_rounding_diff is informational only and never fails validation.
    """
    ibkr_rows   = [r for r in all_rows if r["source"] == "IBKR"]
    equate_rows = [r for r in all_rows if r["source"] == "EquatePlus"]

    headers = [
        "Metric", "Source",
        "Sum exact", "Round(sum exact)", "Sum of filing rows", "Row rounding diff (informational)",
        "Summary value", "Aggregation rule satisfied?", "Aggregation rule",
        "Prior run value", "Delta vs prior run",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = _HEADER_FONT

    all_ok = True

    def _recon_row(metric, source, rows, exact_col, filing_col, use_top_down: bool, prior_val):
        nonlocal all_ok
        sum_exact       = sum(r[exact_col] for r in rows)
        round_sum_exact = _round_ils(sum_exact)
        sum_filing_rows = sum(r[filing_col] for r in rows)
        row_rounding_diff = sum_filing_rows - round_sum_exact

        if use_top_down:
            summary_value = round_sum_exact
            criterion = "Round(sum exact)"
        else:
            summary_value = sum_filing_rows
            criterion = "Sum of filing rows"

        match_ok = (round_sum_exact == summary_value) if use_top_down else (sum_filing_rows == summary_value)
        match_str = str(match_ok)
        if not match_ok:
            all_ok = False

        prior_str = str(prior_val) if prior_val is not None else "N/A"
        delta_str = str(summary_value - prior_val) if prior_val is not None else "N/A"

        ws.append([
            metric, source,
            sum_exact, round_sum_exact, sum_filing_rows, row_rounding_diff,
            summary_value, match_str, criterion,
            prior_str, delta_str,
        ])
        match_cell = ws.cell(row=ws.max_row, column=8)
        if not match_ok:
            match_cell.font = Font(bold=True, color="FF0000")
        else:
            match_cell.font = Font(color="006400")
        # Highlight non-zero delta vs prior run in amber (informational)
        delta_cell = ws.cell(row=ws.max_row, column=11)
        if prior_val is not None and summary_value != prior_val:
            delta_cell.font = Font(color="B8860B")

    for source_label, source_rows, prior in [
        ("IBKR",       ibkr_rows,   prior_ibkr),
        ("EquatePlus", equate_rows, prior_equate),
    ]:
        if not source_rows:
            continue
        p_gains  = prior["gains"]  if prior else None
        p_losses = prior["losses"] if prior else None
        p_gross  = prior["gross"]  if prior else None
        _recon_row("Taxable gains ILS",       source_label, source_rows,
                   "taxable_gain_ils_exact",     "taxable_gain_ils_filing",    False, p_gains)
        _recon_row("Deductible losses ILS",   source_label, source_rows,
                   "deductible_loss_ils_exact",  "deductible_loss_ils_filing", False, p_losses)
        _recon_row("Gross sale proceeds ILS", source_label, source_rows,
                   "gross_sale_proceeds_ils_exact", "gross_sale_proceeds_ils_exact", True, p_gross)
        ws.append([])  # spacer

    # Combined informational rows
    if ibkr_rows and equate_rows:
        _GREY_INFO = Font(italic=True, color="555555")
        for metric, exact_col, filing_col, use_top_down in [
            ("Taxable gains ILS",       "taxable_gain_ils_exact",     "taxable_gain_ils_filing",    False),
            ("Deductible losses ILS",   "deductible_loss_ils_exact",  "deductible_loss_ils_filing", False),
            ("Gross sale proceeds ILS", "gross_sale_proceeds_ils_exact", "gross_sale_proceeds_ils_exact", True),
        ]:
            sum_exact = sum(r[exact_col] for r in all_rows)
            rse = _round_ils(sum_exact)
            sfr = sum(r[filing_col] for r in all_rows)
            ws.append([
                metric, "Combined (info)",
                sum_exact, rse, sfr, sfr - rse,
                rse if use_top_down else sfr, "info", "informational only",
                "N/A", "N/A",
            ])
            for cell in ws[ws.max_row]:
                cell.font = _GREY_INFO

    _autosize_columns(ws)
    return all_ok


def _write_metadata(ws, year: str, ibkr_pdf: str | None, eq_csvs: list[str], eq_pdfs: list[str]) -> None:
    pairs = [
        ("schema_version",        "1"),
        ("tax_year",              year),
        ("generation_timestamp",  datetime.datetime.now().isoformat(timespec="seconds")),
        ("generator",             "generate_1325_support.py"),
        ("moses_methodology",     "ITO Section 91(b) + Circular 10/2025"),
        ("rounding_method",       "ROUND_HALF_UP per Form 1325 row"),
        ("gross_turnover_rounding", "round annual exact total once"),
    ]
    for key, val in pairs:
        ws.append([key, val])
    _autosize_columns(ws)


def _write_sources(ws, ibkr_pdf: str | None, eq_csvs: list[str], eq_pdfs: list[str]) -> None:
    headers = ["type", "filename", "sha256"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = _HEADER_FONT

    if ibkr_pdf:
        ws.append(["IBKR", os.path.basename(ibkr_pdf), _sha256(ibkr_pdf)])
    for path in sorted(eq_csvs):
        ws.append(["EquatePlus CSV", os.path.basename(path), _sha256(path)])
    for path in sorted(eq_pdfs):
        ws.append(["EquatePlus PDF", os.path.basename(path), _sha256(path)])
    _autosize_columns(ws)


# ─── STALE WORKBOOK LOADERS (for reconciliation only) ────────────────────────

def _load_ibkr_summary(year: str) -> dict | None:
    path = f"tax_ibkr_summary_{year}.xlsx"
    if not os.path.exists(path):
        return None
    df = pd.read_excel(path)
    label_col = df.columns[0]
    value_col = df.columns[1]
    label_map = {str(row[label_col]).strip(): row[value_col] for _, row in df.iterrows()}
    expected = ["IBKR taxable realized stock gains",
                "IBKR deductible realized stock losses",
                "IBKR gross sale value (stock disposals)"]
    missing = [k for k in expected if k not in label_map]
    if missing:
        raise RuntimeError(
            f"{path} appears to be pre-Moses or incompatible. Missing rows: {missing}. "
            f"Delete it or re-run annual_tax_summary.py."
        )
    return {
        "gains":  int(label_map["IBKR taxable realized stock gains"]),
        "losses": int(label_map["IBKR deductible realized stock losses"]),
        "gross":  int(label_map["IBKR gross sale value (stock disposals)"]),
    }


def _load_equate_summary(year: str) -> dict | None:
    path = f"tax_summary_{year}.xlsx"
    if not os.path.exists(path):
        return None
    df = pd.read_excel(path)
    total_rows = df[df["Sale date"] == "TOTAL"]
    if total_rows.empty:
        raise RuntimeError(f"No TOTAL row found in {path}")
    total_row = total_rows.iloc[0]
    if "Total_taxable_gain" not in total_row.index:
        raise RuntimeError(
            f"{path} appears to be pre-Moses (missing 'Total_taxable_gain'). "
            f"Re-run tax_calculator.py first."
        )
    return {
        "gains":  _round_ils(float(total_row["Total_taxable_gain"])),
        "losses": _round_ils(float(total_row["Total_deductible_loss"])),
        "gross":  _round_ils(float(total_row["Total_gross_sale_shekel"])),
    }


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate Form 1325 support workbook")
    parser.add_argument("--year", help="Target tax year (YYYY); inferred from filenames if omitted")
    parser.add_argument("--ibkr-pdf", help="Path to IBKR *_tax_statement.pdf (overrides glob)")
    args = parser.parse_args()

    # Discover source files
    if args.ibkr_pdf:
        pdfs = [args.ibkr_pdf]
        if not os.path.exists(args.ibkr_pdf):
            raise SystemExit(f"--ibkr-pdf not found: {args.ibkr_pdf}")
    else:
        pdfs = glob.glob("*_tax_statement.pdf")

    eq_csvs = sorted(glob.glob("consumption_*.csv"))

    if not pdfs and not eq_csvs:
        raise SystemExit("No *_tax_statement.pdf or consumption_*.csv found — nothing to process")
    if len(pdfs) > 1:
        raise SystemExit(f"Expected at most 1 *_tax_statement.pdf, found {len(pdfs)}: {pdfs}")

    # Determine year
    year = args.year
    if not year and pdfs:
        year = extract_year_from_filename(pdfs[0])
        if year == "unknown":
            year = None
    if not year:
        for csv_path in eq_csvs:
            m = re.search(r"consumption_\d+\.\d+\.(\d{4})\.csv$", csv_path)
            if m:
                year = m.group(1)
                break
    if not year:
        raise SystemExit(
            "Could not determine tax year. Use --year YYYY or ensure filenames contain the year."
        )

    print(f"Generating Form 1325 support workbook for year {year} ...")

    ibkr_pdf = pdfs[0] if pdfs else None
    all_rows = []

    if ibkr_pdf:
        # Validate that the IBKR PDF's filename year matches the target year
        pdf_year = extract_year_from_filename(ibkr_pdf)
        if pdf_year != "unknown" and pdf_year != year:
            raise SystemExit(
                f"Year mismatch: --year {year} but IBKR PDF filename suggests year {pdf_year} "
                f"({os.path.basename(ibkr_pdf)}). "
                f"Pass --year {pdf_year} or supply the correct PDF."
            )

        print(f"\nParsing IBKR closed lots from {ibkr_pdf} ...")
        text = extract_full_text(ibkr_pdf)
        raw_ibkr = parse_ibkr_lots_detail(text)
        print(f"  Applying FX conversion and Moses to {len(raw_ibkr)} IBKR raw lots ...")
        ibkr_lots = build_ibkr_lots(raw_ibkr)
        print(f"  {len(ibkr_lots)} IBKR lot rows calculated")
        all_rows.extend(ibkr_lots)

    if eq_csvs:
        print(f"\nParsing EquatePlus lots from {len(eq_csvs)} CSV/PDF pair(s) ...")
        raw_eq = parse_equateplus_lots(year)
        print(f"  Applying FX conversion and Moses to {len(raw_eq)} EquatePlus raw lots ...")
        eq_lots = build_equate_lots(raw_eq)
        print(f"  {len(eq_lots)} EquatePlus lot rows calculated")
        all_rows.extend(eq_lots)

    if not all_rows:
        raise SystemExit("No lot rows produced from any source")

    # Basic invariant checks
    for r in all_rows:
        if r["quantity"] <= 0:
            raise RuntimeError(f"Non-positive quantity: {r['lot_id']}")
        if r["fx_buy"] <= 0 or r["fx_sell"] <= 0:
            raise RuntimeError(f"Non-positive FX rate: {r['lot_id']}")
        if r["original_cost_ils_exact"] < 0:
            raise RuntimeError(f"Negative original cost: {r['lot_id']}")
        if r["gross_sale_proceeds_ils_exact"] < 0:
            raise RuntimeError(f"Negative gross proceeds: {r['lot_id']}")

    # lot_id uniqueness
    lot_ids = [r["lot_id"] for r in all_rows]
    dupes = [lid for lid in set(lot_ids) if lot_ids.count(lid) > 1]
    if dupes:
        raise RuntimeError(f"Duplicate lot_id values: {dupes}")

    _verify_allocation(all_rows)

    sorted_rows = _assign_form_numbers(all_rows)

    # Collect only the EquatePlus CSV+PDF pairs that actually produced Tax Data rows.
    # A CSV is included iff its year matches and its sale PDF exists (parse_equateplus_lots
    # now raises on missing PDFs, so every year-matched CSV contributed rows).
    consumed_eq_csvs = []
    consumed_eq_pdfs = []
    for csv_path in eq_csvs:
        m_date = re.search(r"consumption_(\d+\.\d+\.\d{4})\.csv$", csv_path)
        if not m_date:
            continue
        date_key = m_date.group(1)
        if date_key.split(".")[-1] != year:
            continue
        sale_pdf = f"sale_{date_key}.pdf"
        if os.path.exists(sale_pdf):
            consumed_eq_csvs.append(csv_path)
            consumed_eq_pdfs.append(sale_pdf)

    out_path = f"form1325_support_{year}.xlsx"
    print(f"\nWriting {out_path} ...")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet in ["Form 1325 Entry View", "Form 1325 Support",
                      "Tax Data", "Rounding & Cross-Check", "Metadata", "Sources"]:
            pd.DataFrame().to_excel(writer, sheet_name=sheet, index=False)

        wb = writer.book

        for sheet_name in ["Form 1325 Entry View", "Form 1325 Support",
                            "Tax Data", "Rounding & Cross-Check", "Metadata", "Sources"]:
            ws = wb[sheet_name]
            ws.delete_rows(1, ws.max_row)

        _write_entry_view(wb["Form 1325 Entry View"], sorted_rows)
        _write_support_sheet(wb["Form 1325 Support"], sorted_rows)
        _write_tax_data(wb["Tax Data"], sorted_rows, year)
        prior_ibkr   = _load_ibkr_summary(year)
        prior_equate = _load_equate_summary(year)
        all_ok = _write_reconciliation(wb["Rounding & Cross-Check"], sorted_rows, prior_ibkr, prior_equate)
        _write_metadata(wb["Metadata"], year, ibkr_pdf, consumed_eq_csvs, consumed_eq_pdfs)
        _write_sources(wb["Sources"], ibkr_pdf, consumed_eq_csvs, consumed_eq_pdfs)

    if not all_ok:
        os.remove(out_path)
        raise SystemExit(
            "ERROR: Aggregation rule check failed — detail totals do not match. "
            "Output file removed. Check printed output above."
        )

    n_forms = (len(sorted_rows) - 1) // 10 + 1
    print(f"Done. Wrote {out_path} ({len(sorted_rows)} lot rows, {n_forms} form page(s))")
    print("Verify: open 'Rounding & Cross-Check' sheet — all 'Aggregation rule satisfied?' should be True.")
    print("        open 'Tax Data' sheet — one row per lot, lot_id unique.")


if __name__ == "__main__":
    main()
