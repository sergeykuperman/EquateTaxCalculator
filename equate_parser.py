#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import glob
import os
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
import pdfplumber


def parse_sale_pdf(path: str) -> tuple:
    """
    Extract sale parameters from an EquatePlus sale_*.pdf.
    Returns (sale_price_eur, execution_date, settlement_date, ex_rate_or_None, fees_euro).

    ex_rate is extracted for diagnostic/audit purposes only — the tax calculation
    uses the BoI representative rate, not this PDF-embedded FX rate.
    Missing ex_rate is tolerated (returns None) and logged as a warning.
    All other fields are required; RuntimeError is raised if any are absent.
    """
    with pdfplumber.open(path) as pdf:
        text = pdf.pages[0].extract_text()

    m = re.search(r"Quantity\s*-\s*Shares[\s\S]*?([\d.,]+)\s*(?:€|EUR)", text)
    sale_price = float(m.group(1).replace(",", "")) if m else None

    m = re.search(r"Execution date:\s*([\d]{1,2}\s+[A-Za-z]+\s+\d{4})", text)
    execution_date = datetime.strptime(m.group(1), "%d %b %Y") if m else None

    m = re.search(r"Settlement date:\s*([\d]{1,2}\s+[A-Za-z]+\s+\d{4})", text)
    settlement_date = datetime.strptime(m.group(1), "%d %b %Y") if m else None

    m = re.search(r"Foreign exchange[\s\S]*?(\d+\.\d{5})", text)
    ex_rate = float(m.group(1)) if m else None

    m = re.search(r"Total debits[\s\S]*?(\d+\.\d{2})\s*(?:€|EUR)", text)
    fees_euro = float(m.group(1)) if m else None

    if None in (sale_price, execution_date, fees_euro):
        raise RuntimeError(
            f"Failed to parse required sale params from {path}:\n"
            f" sale_price={sale_price}, execution={execution_date}, fees={fees_euro}\n"
            f"--- page text ---\n{text}"
        )
    if ex_rate is None:
        print(f"  WARNING: no Foreign exchange rate found in {path} "
              f"(diagnostic only — BoI rate is used for tax calculation)")
    return sale_price, execution_date, settlement_date, ex_rate, fees_euro


def parse_equateplus_ticker(pdf_path: str) -> str:
    """Parse the security ticker from the Instrument field of an EquatePlus sale PDF."""
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
    m = re.search(r"Instrument:\s+(\S+)", text)
    if not m:
        raise RuntimeError(f"Could not parse Instrument from {pdf_path}")
    return m.group(1)


def parse_equateplus_pair(csv_path: str, sale_pdf: str) -> list[dict]:
    """
    Build raw per-lot dicts for a single EquatePlus CSV + sale PDF pair.

    Returns one dict per acquisition lot with raw broker facts only.
    No ILS conversion, no Moses calculation — those happen in capital_gains.py.

    The sale_id includes a sale_seq disambiguator so that two sales of the
    same ticker on the same date (rare but possible) do not collide.
    sale_seq is 1-indexed and passed in by the caller; here it defaults to 1
    (single-call use). Use parse_equateplus_lots() for multi-sale year batches
    which assigns sale_seq correctly across all sales.
    """
    return _build_lots_for_pair(csv_path, sale_pdf, sale_seq=1)


def _build_lots_for_pair(csv_path: str, sale_pdf: str, sale_seq: int,
                         expected_date: "datetime.date | None" = None) -> list[dict]:
    """Internal: build raw lots for one CSV+PDF pair with an explicit sale_seq."""
    ticker = parse_equateplus_ticker(sale_pdf)
    sale_price_eur, execution_date, _settlement_date, _pdf_ex_rate, fees_euro = parse_sale_pdf(sale_pdf)

    if expected_date is not None:
        from datetime import timedelta
        delta = abs((execution_date.date() - expected_date).days)
        if delta > 1:
            raise RuntimeError(
                f"Execution date mismatch: {sale_pdf} contains execution date "
                f"{execution_date.date()} but filename implies {expected_date} "
                f"(delta={delta} days). "
                f"Verify that the correct sale PDF is paired with {csv_path}."
            )
        if delta == 1:
            print(f"  NOTE: {sale_pdf} execution date {execution_date.date()} "
                  f"differs from filename date {expected_date} by 1 day "
                  f"(order/execution day offset — normal for EquatePlus)")

    sale_date_str = execution_date.strftime("%Y-%m-%d")

    # sale_id encodes sale_seq so two sales of the same ticker on the same date are distinct
    sale_id = f"EQ/{sale_date_str}/{ticker}/{sale_seq}"

    df = pd.read_csv(csv_path, sep=";", decimal=",")
    df["Acquisition date"] = pd.to_datetime(df["Acquisition date"], format="%d %b %Y")

    total_shares = df["Consumption"].sum()
    acq_date_counter: dict = defaultdict(int)
    rows = []

    for _, lot in df.iterrows():
        acq_date = lot["Acquisition date"].date()
        lot_qty = float(lot["Consumption"])
        purchase_price_eur = float(lot["Purchase price"])

        lot_basis_fc = lot_qty * purchase_price_eur
        lot_proceeds_fc = lot_qty * sale_price_eur
        lot_comm_fc = fees_euro * (lot_qty / total_shares)  # positive = cost

        acq_date_str = acq_date.strftime("%Y-%m-%d")
        acq_date_counter[acq_date_str] += 1
        seq = acq_date_counter[acq_date_str]

        if seq == 1:
            lot_id = f"{sale_id}/LOT/{acq_date_str}"
        else:
            lot_id = f"{sale_id}/LOT/{acq_date_str}/{seq}"

        rows.append({
            "source":               "EquatePlus",
            "ticker":               ticker,
            "isin":                 "",
            "currency":             "EUR",
            "quantity":             lot_qty,
            "acquisition_date":     acq_date_str,
            "sale_date":            sale_date_str,
            "acquisition_basis_fc": lot_basis_fc,
            "cost_per_share_fc":    purchase_price_eur,
            "gross_sale_proceeds_fc": lot_proceeds_fc,
            "sale_commission_fc":   lot_comm_fc,  # positive = cost (EquatePlus convention)
            # sale-level totals — used by _verify_allocation to check lot sums
            "fill_quantity_fc":     total_shares,
            "fill_proceeds_fc":     total_shares * sale_price_eur,
            "fill_commission_fc":   fees_euro,
            "sale_id":              sale_id,
            "lot_id":               lot_id,
            "lot_seq":              seq,
        })
        print(f"  {sale_date_str} {ticker} acq={acq_date_str} qty={lot_qty:.4f} "
              f"lot_id={lot_id}")

    return rows


def parse_equateplus_lots(year: str) -> list[dict]:
    """
    Build raw per-lot dicts for all EquatePlus sales in the current directory
    whose filename year matches `year`.

    Raises RuntimeError if a year-matched CSV has no matching sale PDF.

    Returns one dict per acquisition lot with raw broker facts only.
    No ILS conversion, no Moses calculation — those happen in capital_gains.py.
    """
    csv_files = sorted(glob.glob("consumption_*.csv"))
    if not csv_files:
        return []

    # Group by (sale_date, ticker) to assign sale_seq for collision avoidance.
    # We first collect all (csv_path, sale_pdf) pairs for the target year.
    pairs = []
    for csv_path in csv_files:
        m = re.search(r"consumption_(\d+\.\d+\.\d{4})\.csv$", csv_path)
        if not m:
            continue
        date_key = m.group(1)
        csv_year = date_key.split(".")[-1]
        if csv_year != year:
            print(f"  Skipping {csv_path} (year {csv_year} != target {year})")
            continue
        sale_pdf = f"sale_{date_key}.pdf"
        if not os.path.exists(sale_pdf):
            raise RuntimeError(
                f"Missing sale PDF for year-matched CSV {csv_path}: "
                f"expected {sale_pdf}. "
                f"Provide the file or remove the CSV if this sale should not be processed."
            )
        expected_date = datetime.strptime(date_key, "%d.%m.%Y").date()
        pairs.append((csv_path, sale_pdf, expected_date))

    if not pairs:
        return []

    # Assign sale_seq per (ticker, sale_date) group to avoid sale_id collisions.
    # Parse just the execution_date+ticker from each PDF first (cheap: page 0 only).
    sale_key_counter: dict = defaultdict(int)
    pairs_with_seq = []
    for csv_path, sale_pdf, expected_date in pairs:
        ticker = parse_equateplus_ticker(sale_pdf)
        _, execution_date, *_ = parse_sale_pdf(sale_pdf)
        sale_date_str = execution_date.strftime("%Y-%m-%d")
        key = (sale_date_str, ticker)
        sale_key_counter[key] += 1
        pairs_with_seq.append((csv_path, sale_pdf, sale_key_counter[key], expected_date))

    rows = []
    for csv_path, sale_pdf, sale_seq, expected_date in pairs_with_seq:
        rows.extend(_build_lots_for_pair(csv_path, sale_pdf, sale_seq, expected_date))

    return rows
