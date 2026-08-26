#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/integration/test_2022_pipeline.py

Integration tests for the 2022 tax year pipeline — IBKR Corporate Action disposal.

The 2022 statement contains no normal Trades disposals.  The only taxable
capital-gains event is a cash merger recorded in the Corporate Actions section:

    BRG (BLUEROCK RESIDENTIAL GROWTH) — Merged(Acquisition) for USD 24.25/share
    Effective date: 2022-10-10
    Quantity: 300 shares  |  Proceeds: $7,275.00  |  Acquisition basis: $1,432.04
    Acquisition date: 2020-04-09  |  Realized P/L (IBKR): $5,842.96

The pipeline reconstructs the Israeli taxable amount independently using the
2020-04-09 and 2022-10-10 BoI USD/ILS rates — NOT the IBKR ₪20,903.20 figure,
which was computed by IBKR using its own internal base-ILS rate.

Fixtures required (git-ignored — personal financial data):
  tests/integration/fixtures/2022/ibkr/U3484618_20220103_20221230_tax_statement.pdf

Run with:
  pytest tests/integration/test_2022_pipeline.py -v

Verified 2026-08-26 against real 2022 source file.
"""

import os
import shutil
import sys
import tempfile
from pathlib import Path

import pandas as pd
import pytest

# ─── VERIFIED 2022 REGRESSION FIXTURES ───────────────────────────────────────
# All values verified against real 2022 source files on 2026-08-26.
# Do not change without re-running the pipeline and manually reviewing the output.

EXPECTED_LOT_COUNT        = 1
EXPECTED_TICKER           = "BRG"
EXPECTED_QUANTITY         = 300.0
EXPECTED_BASIS_FC         = 1432.04
EXPECTED_PROCEEDS_FC      = 7275.00
EXPECTED_ACQ_DATE         = "2020-04-09"
EXPECTED_SALE_DATE        = "2022-10-10"
# ILS values use BoI rates for 2020-04-09 (acq) and 2022-10-10 (sale) — NOT IBKR's ₪20,903.20
EXPECTED_TAXABLE_GAIN_ILS   = 20483   # taxable_gain_ils_filing; exact = 20483.30
EXPECTED_GROSS_TURNOVER_ILS = 25644   # _round_ils(gross_sale_proceeds_ils_exact = 25644.375)

# ─── FIXTURE HELPERS ─────────────────────────────────────────────────────────

FIXTURE_DIR = Path(__file__).parent / "fixtures" / "2022"
IBKR_DIR    = FIXTURE_DIR / "ibkr"


def _fixtures_present() -> bool:
    return bool(list(IBKR_DIR.glob("*_tax_statement.pdf")))


def _copy_fixtures_to(tmp: str) -> str:
    """Copy 2022 IBKR PDF into tmp and return its path."""
    ibkr_pdf_src = next(IBKR_DIR.glob("*_tax_statement.pdf"))
    shutil.copy(ibkr_pdf_src, tmp)
    return os.path.join(tmp, ibkr_pdf_src.name)


def _run_generate(tmp: str, ibkr_pdf: str, year: str = "2022") -> str:
    """Run generate_1325_support.main() in tmp, return path to output workbook."""
    orig_dir  = os.getcwd()
    orig_argv = sys.argv[:]
    try:
        os.chdir(tmp)
        sys.argv = ["generate_1325_support.py", "--year", year, "--ibkr-pdf", ibkr_pdf]
        import generate_1325_support
        generate_1325_support.main()
    finally:
        os.chdir(orig_dir)
        sys.argv = orig_argv
    return os.path.join(tmp, f"form1325_support_{year}.xlsx")


# ─── SKIP GUARD ──────────────────────────────────────────────────────────────

_SKIP_REASON = (
    "2022 fixture files not found. "
    "Place IBKR PDF in tests/integration/fixtures/2022/ibkr/. "
)
_skip = pytest.mark.skipif(not _fixtures_present(), reason=_SKIP_REASON)


# ─── TESTS ───────────────────────────────────────────────────────────────────

@pytest.mark.integration
@_skip
def test_2022_pipeline_produces_workbook():
    """Pipeline runs without error and produces a workbook with the expected sheets."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        assert os.path.exists(wb_path)

        xl = pd.ExcelFile(wb_path)
        expected_sheets = {
            "Form 1325 Entry View", "Form 1325 Support",
            "Tax Data", "Rounding & Cross-Check", "Metadata", "Sources",
        }
        assert expected_sheets.issubset(set(xl.sheet_names)), (
            f"Missing sheets: {expected_sheets - set(xl.sheet_names)}"
        )


@pytest.mark.integration
@_skip
def test_2022_brg_corporate_action_lot():
    """BRG corporate-action disposal is parsed with correct raw broker facts."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")

        assert len(df) == EXPECTED_LOT_COUNT, f"Expected {EXPECTED_LOT_COUNT} lot row(s), got {len(df)}"
        row = df.iloc[0]

        assert row["ticker"]   == EXPECTED_TICKER,   f"ticker: {row['ticker']}"
        assert row["currency"] == "USD",              f"currency: {row['currency']}"
        assert abs(row["quantity"] - EXPECTED_QUANTITY) < 0.01,  f"quantity: {row['quantity']}"
        assert row["acquisition_date"] == EXPECTED_ACQ_DATE,     f"acq_date: {row['acquisition_date']}"
        assert row["sale_date"]        == EXPECTED_SALE_DATE,     f"sale_date: {row['sale_date']}"
        assert abs(row["gross_sale_proceeds_fc"] - EXPECTED_PROCEEDS_FC) < 0.02, (
            f"proceeds_fc: {row['gross_sale_proceeds_fc']}"
        )
        assert abs(row["acquisition_basis_fc"] - EXPECTED_BASIS_FC) < 0.02, (
            f"basis_fc: {row['acquisition_basis_fc']}"
        )
        assert abs(row["sale_commission_fc"]) < 0.001, f"commission should be 0: {row['sale_commission_fc']}"

        # fill_id must use CA namespace
        assert row["fill_id"] == "IBKR/CA/2022-10-10/BRG/1", f"fill_id: {row['fill_id']}"
        assert row["lot_id"]  == "IBKR/CA/2022-10-10/BRG/1/LOT/1", f"lot_id: {row['lot_id']}"


@pytest.mark.integration
@_skip
def test_2022_ils_totals():
    """Taxable gain and gross turnover in ILS match verified 2022 values."""
    from tax_utils import _round_ils
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")

        gain  = int(df["taxable_gain_ils_filing"].sum())
        loss  = int(df["deductible_loss_ils_filing"].sum())
        gross = _round_ils(df["gross_sale_proceeds_ils_exact"].sum())

        assert gain  == EXPECTED_TAXABLE_GAIN_ILS,   f"taxable gain ILS: got {gain}"
        assert loss  == 0,                            f"deductible loss ILS: got {loss}"
        assert gross == EXPECTED_GROSS_TURNOVER_ILS, f"gross turnover ILS: got {gross}"


@pytest.mark.integration
@_skip
def test_2022_aggregation_rule_all_satisfied():
    """All rows in 'Rounding & Cross-Check' sheet have Aggregation rule satisfied? = True."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Rounding & Cross-Check")

        col = "Aggregation rule satisfied?"
        assert col in df.columns, f"Column '{col}' not found"
        non_info = df[df[col].notna() & (df[col].astype(str) != "info")]
        failures = non_info[non_info[col].astype(str) == "False"]
        assert failures.empty, (
            f"Aggregation rule failures:\n{failures[['Metric', 'Source', col]].to_string()}"
        )


@pytest.mark.integration
@_skip
def test_2022_lot_id_unique():
    """Every lot_id in Tax Data is unique."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")

        dupes = df[df["lot_id"].duplicated(keep=False)]["lot_id"].unique().tolist()
        assert not dupes, f"Duplicate lot_ids: {dupes}"
