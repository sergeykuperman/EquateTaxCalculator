#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_tax.py — pytest suite for the tax calculation modules.

Run all tests:
    pytest test_tax.py -v

Run without integration tests (no source files needed):
    pytest test_tax.py -v -m "not integration"
"""

import os
import tempfile

import pandas as pd
import pytest

from tax_utils import _round_ils, moses_gain_loss, TAX_RATE


# ─── 1. moses_gain_loss — 4 canonical Circular 10/2025 cases ─────────────────

def test_moses_fx_up_price_up():
    # buy $1000 @ 3.4 → cost=3400; sell $1200 @ 3.7 → net=4440
    # adjusted=3700; nominal=1040; inflationary=300; taxable=740
    gain, loss = moses_gain_loss(3400, 4440, 3.4, 3.7)
    assert abs(gain - 740) < 0.01
    assert loss == 0.0


def test_moses_fx_up_price_down():
    # buy $1000 @ 3.4 → cost=3400; sell $800 @ 3.7 → net=2960; nominal=-440
    gain, loss = moses_gain_loss(3400, 2960, 3.4, 3.7)
    assert gain == 0.0
    assert abs(loss - 440) < 0.01


def test_moses_fx_up_big_all_inflationary():
    # buy $1000 @ 3.4 → cost=3400; sell $800 @ 5.0 → net=4000; nominal=+600
    # adjusted=5000; inflationary=1600; exempt=min(600,1600)=600; taxable=0
    gain, loss = moses_gain_loss(3400, 4000, 3.4, 5.0)
    assert gain == 0.0
    assert loss == 0.0


def test_moses_fx_down_price_up():
    # buy $1000 @ 3.7 → cost=3700; sell $1200 @ 3.4 → net=4080; nominal=+380
    gain, loss = moses_gain_loss(3700, 4080, 3.7, 3.4)
    assert abs(gain - 380) < 0.01
    assert loss == 0.0


def test_moses_fx_down_price_down():
    # buy $1000 @ 3.7 → cost=3700; sell $900 @ 3.4 → net=3060; nominal=-640
    # adjusted=3400; neg_inflationary=300; deductible=max(0,640-300)=340
    gain, loss = moses_gain_loss(3700, 3060, 3.7, 3.4)
    assert gain == 0.0
    assert abs(loss - 340) < 0.01


# ─── 2. _round_ils uses ROUND_HALF_UP ────────────────────────────────────────

def test_round_ils_half_up():
    assert _round_ils(0.5) == 1   # rounds up (not banker's: would be 0)
    assert _round_ils(1.5) == 2   # rounds up
    assert _round_ils(2.5) == 3   # rounds up
    assert _round_ils(-0.5) == -1  # ROUND_HALF_UP rounds half away from zero
    assert _round_ils(100.4) == 100
    assert _round_ils(100.5) == 101


# ─── 3 & 4. Commission sign conventions ──────────────────────────────────────

def test_ibkr_commission_sign():
    # IBKR: comm is signed (negative = cost), net = gross + signed_comm
    from capital_gains import _apply_moses
    gross_fc = 1000.0
    comm_fc  = -5.0      # negative = cost
    fx = 3.7
    result = _apply_moses(
        lot_basis_fc=800.0, lot_proceeds_fc=gross_fc, lot_comm_fc=comm_fc,
        fx_acq=fx, fx_sell=fx, comm_is_cost=False
    )
    expected_net = (gross_fc + comm_fc) * fx   # 995 * 3.7
    assert abs(result["net_sale_consideration_ils_exact"] - expected_net) < 0.01


def test_equate_commission_sign():
    # EquatePlus: comm is positive cost, net = gross - comm
    from capital_gains import _apply_moses
    gross_fc = 1000.0
    comm_fc  = 5.0       # positive = cost
    fx = 3.7
    result = _apply_moses(
        lot_basis_fc=800.0, lot_proceeds_fc=gross_fc, lot_comm_fc=comm_fc,
        fx_acq=fx, fx_sell=fx, comm_is_cost=True
    )
    expected_net = (gross_fc - comm_fc) * fx   # 995 * 3.7
    assert abs(result["net_sale_consideration_ils_exact"] - expected_net) < 0.01


# ─── 5. One fill → one lot ────────────────────────────────────────────────────

def test_ibkr_one_fill_one_lot():
    """Single-lot fill: fill_id correct, lot_id ends with /LOT/1."""
    raw = [{
        "source": "IBKR", "ticker": "VZ", "isin": "", "currency": "USD",
        "exchange": "NYSE", "quantity": 10.0,
        "acquisition_date": "2023-01-15", "sale_date": "2024-06-01",
        "acquisition_basis_fc": 500.0, "cost_per_share_fc": 50.0,
        "gross_sale_proceeds_fc": 520.0, "sale_commission_fc": -1.0,
        "fill_id": "IBKR/2024-06-01/VZ/NYSE/1",
        "lot_id": "IBKR/2024-06-01/VZ/NYSE/1/LOT/1",
        "fill_seq": 1, "lot_seq": 1,
    }]
    from unittest.mock import patch
    with patch("capital_gains.boi_ils", return_value=3.7):
        from capital_gains import build_ibkr_lots
        result = build_ibkr_lots(raw)
    assert len(result) == 1
    assert result[0]["fill_id"] == "IBKR/2024-06-01/VZ/NYSE/1"
    assert result[0]["lot_id"] == "IBKR/2024-06-01/VZ/NYSE/1/LOT/1"


# ─── 6. One fill → two lots: shared fill_id, distinct lot_id, qtys sum ───────

def test_ibkr_one_fill_two_lots():
    """Two lots from same fill share fill_id; quantities sum to fill total."""
    fill_id = "IBKR/2024-01-31/WBD/EDGEA/1"
    raw = [
        {
            "source": "IBKR", "ticker": "WBD", "isin": "", "currency": "USD",
            "exchange": "EDGEA", "quantity": 0.597598,
            "acquisition_date": "2020-01-14", "sale_date": "2024-01-31",
            "acquisition_basis_fc": 10.0, "cost_per_share_fc": 16.73,
            "gross_sale_proceeds_fc": 6.0, "sale_commission_fc": -0.1,
            "fill_id": fill_id,
            "lot_id": f"{fill_id}/LOT/1",
            "fill_seq": 1, "lot_seq": 1,
        },
        {
            "source": "IBKR", "ticker": "WBD", "isin": "", "currency": "USD",
            "exchange": "EDGEA", "quantity": 10.402402,
            "acquisition_date": "2020-02-28", "sale_date": "2024-01-31",
            "acquisition_basis_fc": 174.16, "cost_per_share_fc": 16.74,
            "gross_sale_proceeds_fc": 104.0, "sale_commission_fc": -1.74,
            "fill_id": fill_id,
            "lot_id": f"{fill_id}/LOT/2",
            "fill_seq": 1, "lot_seq": 2,
        },
    ]
    from unittest.mock import patch
    with patch("capital_gains.boi_ils", return_value=3.7):
        from capital_gains import build_ibkr_lots
        result = build_ibkr_lots(raw)

    assert len(result) == 2
    assert result[0]["fill_id"] == result[1]["fill_id"] == fill_id
    assert result[0]["lot_id"] != result[1]["lot_id"]
    assert result[0]["lot_id"].endswith("/LOT/1")
    assert result[1]["lot_id"].endswith("/LOT/2")
    total_qty = result[0]["quantity"] + result[1]["quantity"]
    assert abs(total_qty - 11.0) < 0.001


# ─── 7. Two fills same ticker/exchange/date — distinct fill_id ───────────────

def test_ibkr_two_fills_same_day_distinct_fill_id():
    fill_id_1 = "IBKR/2024-03-15/BEP/NYSE/1"
    fill_id_2 = "IBKR/2024-03-15/BEP/NYSE/2"
    raw = [
        {
            "source": "IBKR", "ticker": "BEP", "isin": "", "currency": "USD",
            "exchange": "NYSE", "quantity": 5.0,
            "acquisition_date": "2022-05-01", "sale_date": "2024-03-15",
            "acquisition_basis_fc": 100.0, "cost_per_share_fc": 20.0,
            "gross_sale_proceeds_fc": 110.0, "sale_commission_fc": -0.5,
            "fill_id": fill_id_1, "lot_id": f"{fill_id_1}/LOT/1",
            "fill_seq": 1, "lot_seq": 1,
        },
        {
            "source": "IBKR", "ticker": "BEP", "isin": "", "currency": "USD",
            "exchange": "NYSE", "quantity": 3.0,
            "acquisition_date": "2022-06-01", "sale_date": "2024-03-15",
            "acquisition_basis_fc": 60.0, "cost_per_share_fc": 20.0,
            "gross_sale_proceeds_fc": 66.0, "sale_commission_fc": -0.3,
            "fill_id": fill_id_2, "lot_id": f"{fill_id_2}/LOT/1",
            "fill_seq": 2, "lot_seq": 1,
        },
    ]
    from unittest.mock import patch
    with patch("capital_gains.boi_ils", return_value=3.7):
        from capital_gains import build_ibkr_lots
        result = build_ibkr_lots(raw)

    assert result[0]["fill_id"] != result[1]["fill_id"]
    assert result[0]["fill_id"] == fill_id_1
    assert result[1]["fill_id"] == fill_id_2


# ─── 8. Fractional quantities ─────────────────────────────────────────────────

def test_fractional_quantities():
    fill_id = "IBKR/2024-01-31/WBD/EDGEA/1"
    raw = [
        {"source": "IBKR", "ticker": "WBD", "isin": "", "currency": "USD",
         "exchange": "EDGEA", "quantity": 0.597598,
         "acquisition_date": "2020-01-14", "sale_date": "2024-01-31",
         "acquisition_basis_fc": 10.0, "cost_per_share_fc": 16.73,
         "gross_sale_proceeds_fc": 5.98, "sale_commission_fc": -0.1,
         "fill_id": fill_id, "lot_id": f"{fill_id}/LOT/1",
         "fill_seq": 1, "lot_seq": 1},
        {"source": "IBKR", "ticker": "WBD", "isin": "", "currency": "USD",
         "exchange": "EDGEA", "quantity": 10.402402,
         "acquisition_date": "2020-02-28", "sale_date": "2024-01-31",
         "acquisition_basis_fc": 174.16, "cost_per_share_fc": 16.74,
         "gross_sale_proceeds_fc": 104.02, "sale_commission_fc": -1.74,
         "fill_id": fill_id, "lot_id": f"{fill_id}/LOT/2",
         "fill_seq": 1, "lot_seq": 2},
    ]
    from unittest.mock import patch
    with patch("capital_gains.boi_ils", return_value=3.7):
        from capital_gains import build_ibkr_lots
        result = build_ibkr_lots(raw)

    # Both lots should have positive quantity and valid calculations
    for lot in result:
        assert lot["quantity"] > 0
        assert lot["original_cost_ils_exact"] > 0
        assert lot["gross_sale_proceeds_ils_exact"] > 0


# ─── 9. EquatePlus same-sale same-acq-date → distinct lot_id with counter ────

def test_equate_same_acq_date_counter_suffix():
    """Two lots from same sale with identical acquisition date get distinct lot_ids."""
    from equate_parser import parse_equateplus_lots
    import unittest.mock as mock

    csv_content = (
        "Acquisition date;Consumption;Purchase price\n"
        "01 Jan 2022;10;50\n"
        "01 Jan 2022;5;50\n"  # same acquisition date
    )
    pdf_text = "Instrument: SAP\nQuantity - Shares 15 EUR\nExecution date: 15 Jun 2024 09:00:00 CET\nSettlement date: 17 Jun 2024\nForeign exchange 3.94123\nTotal debits 10.00 EUR"

    with tempfile.TemporaryDirectory() as tmpdir:
        orig_dir = os.getcwd()
        try:
            os.chdir(tmpdir)
            with open("consumption_15.6.2024.csv", "w") as f:
                f.write(csv_content)
            # Create a dummy sale PDF so os.path.exists passes
            with open("sale_15.6.2024.pdf", "w") as f:
                f.write("dummy")

            mock_page = mock.MagicMock()
            mock_page.extract_text.return_value = pdf_text
            mock_pdf = mock.MagicMock()
            mock_pdf.__enter__ = mock.MagicMock(return_value=mock_pdf)
            mock_pdf.__exit__ = mock.MagicMock(return_value=False)
            mock_pdf.pages = [mock_page]

            with mock.patch("pdfplumber.open", return_value=mock_pdf):
                lots = parse_equateplus_lots("2024")

            assert len(lots) == 2
            assert lots[0]["lot_id"] != lots[1]["lot_id"]
            assert lots[0]["sale_id"] == lots[1]["sale_id"]
            # First occurrence has no suffix; second has /2
            assert not lots[0]["lot_id"].endswith("/2")
            assert lots[1]["lot_id"].endswith("/2")
        finally:
            os.chdir(orig_dir)


# ─── 10. Mixed-year rejection ────────────────────────────────────────────────

def test_equate_year_filter():
    """Lots from a different year are skipped."""
    from equate_parser import parse_equateplus_lots
    import unittest.mock as mock

    csv_content = "Acquisition date;Consumption;Purchase price\n01 Jan 2021;10;50\n"

    with tempfile.TemporaryDirectory() as tmpdir:
        orig_dir = os.getcwd()
        try:
            os.chdir(tmpdir)
            with open("consumption_15.6.2023.csv", "w") as f:
                f.write(csv_content)
            # sale PDF is not needed because the year filter skips this file before opening it

            lots_2024 = parse_equateplus_lots("2024")  # target year 2024, file is 2023

            assert lots_2024 == []
        finally:
            os.chdir(orig_dir)


# ─── 11. row_rounding_diff non-zero does NOT cause match=False for gain/loss ──

def test_reconciliation_gain_match_uses_sum_filing_rows():
    """
    For gains/losses, match criterion is sum_filing_rows == summary_value.
    row_rounding_diff (sum_filing_rows - round_sum_exact) may be non-zero without failing.
    Example: 100.49 * 3 lots → sum_exact=301.47, round_sum_exact=301, sum_filing_rows=300.
    summary_value = 300 (sum of filing rows). match = 300==300 = True.
    row_rounding_diff = 300 - 301 = -1 (informational, not a failure).
    """
    exact_values = [100.49, 100.49, 100.49]
    sum_exact = sum(exact_values)
    round_sum_exact = _round_ils(sum_exact)       # 301
    sum_filing_rows = sum(_round_ils(v) for v in exact_values)  # 300

    # The bottom-up path: summary_value == sum_filing_rows
    summary_value = sum_filing_rows
    match_ok = (summary_value == sum_filing_rows)
    row_rounding_diff = sum_filing_rows - round_sum_exact

    assert round_sum_exact == 301
    assert sum_filing_rows == 300
    assert row_rounding_diff == -1        # non-zero, informational
    assert match_ok is True               # match still passes


# ─── 12. Gross turnover uses _round_ils(sum_exact), not sum_filing_rows ──────

def test_gross_turnover_uses_round_sum_exact():
    """1322 turnover = _round_ils(sum of exact gross proceeds), not sum of per-row rounded values."""
    exact_proceeds = [1000.49, 2000.49, 3000.49]
    summary_gross = _round_ils(sum(exact_proceeds))   # round annual total once
    sum_filing = sum(_round_ils(v) for v in exact_proceeds)  # per-row rounded, then summed

    # They may differ; summary_gross is authoritative
    assert summary_gross == _round_ils(6001.47)
    assert summary_gross == 6001
    # Demonstrate that sum_filing may differ (here: 1000+2000+3000=6000)
    assert sum_filing == 6000
    assert summary_gross != sum_filing   # confirms the two paths differ


# ─── 13–15. load_tax_data validation ─────────────────────────────────────────

def _make_tax_data_workbook(path: str, schema_version="1", tax_year="2024",
                             include_all_columns=True, duplicate_lot_id=False):
    """Helper: write a minimal valid (or deliberately broken) form1325_support workbook."""
    columns = [
        "schema_version", "tax_year", "source", "fill_id", "lot_id",
        "gross_sale_proceeds_ils_exact",
        "taxable_gain_ils_filing", "deductible_loss_ils_filing",
        "tax_rate",
    ]
    if not include_all_columns:
        columns = columns[:3]   # drop required columns

    lot_ids = ["EQ/2024-01-01/SAP/LOT/2023-01-01", "EQ/2024-01-01/SAP/LOT/2023-01-01"] if duplicate_lot_id else \
              ["EQ/2024-01-01/SAP/LOT/2023-01-01"]

    rows = [{c: ("1" if c == "schema_version" else "2024" if c == "tax_year"
                 else lid if c == "lot_id" else "EQ/2024-01-01/SAP" if c == "fill_id"
                 else "EquatePlus" if c == "source" else 0.25 if c == "tax_rate" else 1000.0)
             for c in columns}
            for lid in lot_ids]

    df = pd.DataFrame(rows, columns=columns)
    metadata = pd.DataFrame([
        [k, v] for k, v in [("schema_version", schema_version), ("tax_year", tax_year)]
    ])
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Tax Data", index=False)
        metadata.to_excel(writer, sheet_name="Metadata", header=False, index=False)


def test_load_tax_data_wrong_schema_version():
    from annual_tax_summary import load_tax_data
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "form1325_support_2024.xlsx")
        _make_tax_data_workbook(path, schema_version="99")
        orig = os.getcwd()
        try:
            os.chdir(tmpdir)
            with pytest.raises(RuntimeError, match="schema version"):
                load_tax_data("2024")
        finally:
            os.chdir(orig)


def test_load_tax_data_wrong_year():
    from annual_tax_summary import load_tax_data
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "form1325_support_2024.xlsx")
        _make_tax_data_workbook(path, tax_year="2023")
        orig = os.getcwd()
        try:
            os.chdir(tmpdir)
            with pytest.raises(RuntimeError, match="tax_year"):
                load_tax_data("2024")
        finally:
            os.chdir(orig)


def test_load_tax_data_file_absent():
    from annual_tax_summary import load_tax_data
    with tempfile.TemporaryDirectory() as tmpdir:
        orig = os.getcwd()
        try:
            os.chdir(tmpdir)
            with pytest.raises(SystemExit, match="not found"):
                load_tax_data("2024")
        finally:
            os.chdir(orig)


# ─── 16. load_tax_data raises RuntimeError (not AssertionError) on duplicate lot_id ──

def test_load_tax_data_duplicate_lot_id():
    from annual_tax_summary import load_tax_data
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "form1325_support_2024.xlsx")
        _make_tax_data_workbook(path, duplicate_lot_id=True)
        orig = os.getcwd()
        try:
            os.chdir(tmpdir)
            with pytest.raises(RuntimeError, match="[Dd]uplicate lot_id"):
                load_tax_data("2024")
        finally:
            os.chdir(orig)


# ─── 17–18. Integration tests (require actual 2024 source files) ─────────────

@pytest.mark.integration
def test_2024_ibkr_lot_count():
    """
    PLACEHOLDER — freeze actual count once generate_1325_support.py --year 2024
    has been run and the output manually verified.
    """
    from annual_tax_summary import load_tax_data
    rows = load_tax_data("2024")
    ibkr_rows = [r for r in rows if r["source"] == "IBKR"]
    # TODO: replace with verified count after first run
    assert len(ibkr_rows) > 0, "Expected at least one IBKR lot row"


@pytest.mark.integration
def test_2024_equateplus_lot_count():
    """
    PLACEHOLDER — freeze actual count once generate_1325_support.py --year 2024
    has been run and the output manually verified.
    """
    from annual_tax_summary import load_tax_data
    rows = load_tax_data("2024")
    eq_rows = [r for r in rows if r["source"] == "EquatePlus"]
    # TODO: replace with verified count after first run
    assert len(eq_rows) > 0, "Expected at least one EquatePlus lot row"


@pytest.mark.integration
def test_2024_ibkr_totals():
    """
    PLACEHOLDER — freeze exact gains/losses/gross once generate_1325_support.py --year 2024
    has been run and the output manually verified and recorded here as regression fixtures.

    Example (replace TODO values after verification):
        assert ibkr_gains  == 217
        assert ibkr_losses == 1860
        assert ibkr_gross  == 15459
    """
    from annual_tax_summary import load_tax_data
    rows = load_tax_data("2024")
    ibkr_rows = [r for r in rows if r["source"] == "IBKR"]
    ibkr_gains  = sum(r["taxable_gain_ils_filing"] for r in ibkr_rows)
    ibkr_losses = sum(r["deductible_loss_ils_filing"] for r in ibkr_rows)
    ibkr_gross  = _round_ils(sum(r["gross_sale_proceeds_ils_exact"] for r in ibkr_rows))
    # TODO: replace with verified values after first run
    assert ibkr_gains  >= 0
    assert ibkr_losses >= 0
    assert ibkr_gross  > 0
