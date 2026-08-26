#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tests/integration/test_2024_pipeline.py

Integration tests for the 2024 tax year pipeline.

These tests require real 2024 source files placed in:
  tests/integration/fixtures/2024/ibkr/   — IBKR *_tax_statement.pdf
  tests/integration/fixtures/2024/equate/ — consumption_*.csv + sale_*.pdf

Run with:
  pytest tests/integration/test_2024_pipeline.py -v

Fixtures are copied into a temporary directory so the originals are never
modified and generated .xlsx files do not pollute the fixture tree.

Verified 2026-08-26 against real 2024 source files.
"""

import os
import shutil
import sys
import tempfile
from pathlib import Path

import pandas as pd
import pytest

# ─── VERIFIED 2024 REGRESSION FIXTURES ───────────────────────────────────────
# All values verified against real 2024 source files on 2026-08-26.
# Do not change without re-running the pipeline and manually reviewing the output.

# Lot counts
EXPECTED_IBKR_LOT_COUNT   = 7    # BEP×1 BEPC×1 FL×1 VZ×2 WBD×2
EXPECTED_EQUATE_LOT_COUNT = 70   # 14 EquatePlus sales

# IBKR filing totals @ 25% (from Tax Data filing columns)
EXPECTED_IBKR_GAIN  = 217     # sum(taxable_gain_ils_filing) for IBKR rows
EXPECTED_IBKR_LOSS  = 1861    # sum(deductible_loss_ils_filing) — note: exact sum ~1860.29,
                               # individually-rounded rows sum to 1861 (correct per rounding policy)
EXPECTED_IBKR_GROSS = 15459   # _round_ils(sum(gross_sale_proceeds_ils_exact)) for IBKR

# EquatePlus filing totals @ 25%
EXPECTED_EQUATE_GAIN  = 128993
EXPECTED_EQUATE_LOSS  = 0
EXPECTED_EQUATE_GROSS = 317855

# Combined filing totals
EXPECTED_COMBINED_GAIN  = 129210  # IBKR + EquatePlus gains
EXPECTED_COMBINED_LOSS  = 1861    # IBKR + EquatePlus losses
EXPECTED_COMBINED_GROSS = 333313  # _round_ils(sum of ALL exact proceeds)

# Structural facts
EXPECTED_EQUATE_SALE_COUNT = 14   # distinct sale_id values in EquatePlus rows
EXPECTED_IBKR_FILL_COUNT   = 6    # distinct fill_id values in IBKR rows
EXPECTED_TOTAL_LOT_COUNT   = 77   # IBKR + EquatePlus


# ─── FIXTURE HELPERS ─────────────────────────────────────────────────────────

FIXTURE_DIR = Path(__file__).parent / "fixtures" / "2024"
IBKR_DIR    = FIXTURE_DIR / "ibkr"
EQUATE_DIR  = FIXTURE_DIR / "equate"


def _fixtures_present() -> bool:
    ibkr_pdfs   = list(IBKR_DIR.glob("*_tax_statement.pdf"))
    equate_csvs = list(EQUATE_DIR.glob("consumption_*.csv"))
    return bool(ibkr_pdfs and equate_csvs)


def _copy_fixtures_to(tmp: str) -> str:
    """Copy all 2024 fixture files into tmp and return path to the IBKR PDF."""
    ibkr_pdf_src = next(IBKR_DIR.glob("*_tax_statement.pdf"))
    shutil.copy(ibkr_pdf_src, tmp)
    for path in EQUATE_DIR.glob("consumption_*.csv"):
        shutil.copy(path, tmp)
    for path in EQUATE_DIR.glob("sale_*.pdf"):
        shutil.copy(path, tmp)
    return os.path.join(tmp, ibkr_pdf_src.name)


def _run_generate(tmp: str, ibkr_pdf: str, year: str = "2024") -> str:
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
    "2024 fixture files not found. "
    "Place IBKR PDF in tests/integration/fixtures/2024/ibkr/ "
    "and EquatePlus CSV+PDF pairs in tests/integration/fixtures/2024/equate/. "
    "See tests/integration/fixtures/2024/README.md."
)
_skip = pytest.mark.skipif(not _fixtures_present(), reason=_SKIP_REASON)


# ─── TESTS ───────────────────────────────────────────────────────────────────

@pytest.mark.integration
@_skip
def test_2024_pipeline_produces_workbook():
    """Pipeline runs without error and produces a workbook with the expected sheets."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path = _run_generate(tmp, ibkr_pdf)
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
def test_2024_lot_counts():
    """IBKR and EquatePlus lot counts match verified 2024 values."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")

        assert len(df) == EXPECTED_TOTAL_LOT_COUNT, f"Total lots: got {len(df)}"
        assert len(df[df["source"] == "IBKR"])       == EXPECTED_IBKR_LOT_COUNT,   f"IBKR lots: got {len(df[df['source']=='IBKR'])}"
        assert len(df[df["source"] == "EquatePlus"]) == EXPECTED_EQUATE_LOT_COUNT, f"EquatePlus lots: got {len(df[df['source']=='EquatePlus'])}"


@pytest.mark.integration
@_skip
def test_2024_ibkr_totals():
    """IBKR gains, losses, gross turnover match verified 2024 values."""
    from tax_utils import _round_ils
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")
        ibkr = df[df["source"] == "IBKR"]

        gain  = int(ibkr["taxable_gain_ils_filing"].sum())
        loss  = int(ibkr["deductible_loss_ils_filing"].sum())
        gross = _round_ils(ibkr["gross_sale_proceeds_ils_exact"].sum())

        assert gain  == EXPECTED_IBKR_GAIN,  f"IBKR gains: got {gain}"
        assert loss  == EXPECTED_IBKR_LOSS,  f"IBKR losses: got {loss}"
        assert gross == EXPECTED_IBKR_GROSS, f"IBKR gross: got {gross}"


@pytest.mark.integration
@_skip
def test_2024_equate_totals():
    """EquatePlus gains, losses, gross turnover match verified 2024 values."""
    from tax_utils import _round_ils
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")
        eq = df[df["source"] == "EquatePlus"]

        gain  = int(eq["taxable_gain_ils_filing"].sum())
        loss  = int(eq["deductible_loss_ils_filing"].sum())
        gross = _round_ils(eq["gross_sale_proceeds_ils_exact"].sum())

        assert gain  == EXPECTED_EQUATE_GAIN,  f"EquatePlus gains: got {gain}"
        assert loss  == EXPECTED_EQUATE_LOSS,  f"EquatePlus losses: got {loss}"
        assert gross == EXPECTED_EQUATE_GROSS, f"EquatePlus gross: got {gross}"


@pytest.mark.integration
@_skip
def test_2024_combined_totals():
    """Combined gains, losses, gross turnover match verified 2024 values."""
    from tax_utils import _round_ils
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")

        gain  = int(df["taxable_gain_ils_filing"].sum())
        loss  = int(df["deductible_loss_ils_filing"].sum())
        gross = _round_ils(df["gross_sale_proceeds_ils_exact"].sum())

        assert gain  == EXPECTED_COMBINED_GAIN,  f"Combined gains: got {gain}"
        assert loss  == EXPECTED_COMBINED_LOSS,  f"Combined losses: got {loss}"
        assert gross == EXPECTED_COMBINED_GROSS, f"Combined gross: got {gross}"


@pytest.mark.integration
@_skip
def test_2024_structural_counts():
    """Number of distinct Equate sales and IBKR fills match verified 2024 values."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")

        ibkr = df[df["source"] == "IBKR"]
        eq   = df[df["source"] == "EquatePlus"]

        eq_sales   = eq["fill_id"].nunique()   # fill_id == sale_id for EquatePlus
        ibkr_fills = ibkr["fill_id"].nunique()

        assert eq_sales   == EXPECTED_EQUATE_SALE_COUNT, f"EquatePlus sales: got {eq_sales}"
        assert ibkr_fills == EXPECTED_IBKR_FILL_COUNT,   f"IBKR fills: got {ibkr_fills}"


@pytest.mark.integration
@_skip
def test_2024_wbd_multi_lot_fill():
    """WBD fill closes exactly 2 lots sharing one fill_id with distinct lot_ids."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")

        wbd = df[(df["source"] == "IBKR") & (df["ticker"] == "WBD")]
        assert len(wbd) == 2, f"WBD: expected 2 lots, got {len(wbd)}"
        assert wbd["fill_id"].nunique() == 1, "WBD: expected 1 fill_id shared by both lots"
        assert wbd["lot_id"].nunique()  == 2, "WBD: expected 2 distinct lot_ids"
        assert wbd["lot_id"].iloc[0].endswith("/LOT/1")
        assert wbd["lot_id"].iloc[1].endswith("/LOT/2")


@pytest.mark.integration
@_skip
def test_2024_aggregation_rule_all_satisfied():
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
def test_2024_lot_id_unique():
    """Every lot_id in Tax Data is unique."""
    with tempfile.TemporaryDirectory() as tmp:
        ibkr_pdf = _copy_fixtures_to(tmp)
        wb_path  = _run_generate(tmp, ibkr_pdf)
        df = pd.read_excel(wb_path, sheet_name="Tax Data")

        dupes = df[df["lot_id"].duplicated(keep=False)]["lot_id"].unique().tolist()
        assert not dupes, f"Duplicate lot_ids: {dupes}"
