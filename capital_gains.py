#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
capital_gains.py — Single capital-gains calculation layer.

Takes raw broker-fact dicts from ibkr_parser or equate_parser and produces
fully-calculated lot dicts matching the Tax Data schema.

Neither generate_1325_support.py nor tax_calculator.py may reproduce any of:
  - acquisition-basis ILS conversion
  - sale-proceeds ILS conversion
  - commission ILS conversion
  - adjusted-cost calculation
  - Moses gain/loss application

All callers must go through build_ibkr_lots() or build_equate_lots().
"""

from tax_utils import TAX_RATE, _round_ils, boi_ils, moses_gain_loss


def _apply_moses(lot_basis_fc: float, lot_proceeds_fc: float,
                 lot_comm_fc: float, fx_acq: float, fx_sell: float,
                 comm_is_cost: bool) -> dict:
    """
    Convert one lot's raw FC amounts to ILS and apply Moses.

    comm_is_cost=True  → EquatePlus convention: net = gross - commission
    comm_is_cost=False → IBKR convention: comm is signed (negative=cost), net = gross + signed_comm
    """
    original_cost_ils   = lot_basis_fc * fx_acq
    adjusted_cost_ils   = original_cost_ils * (fx_sell / fx_acq)
    gross_lot_ils       = lot_proceeds_fc * fx_sell

    if comm_is_cost:
        lot_comm_ils    = lot_comm_fc * fx_sell          # positive cost
        net_sale_ils    = gross_lot_ils - lot_comm_ils
    else:
        lot_comm_ils    = lot_comm_fc * fx_sell          # signed: negative=cost
        net_sale_ils    = gross_lot_ils + lot_comm_ils

    taxable_gain, deductible_loss = moses_gain_loss(
        original_cost_ils, net_sale_ils, fx_acq, fx_sell
    )

    nominal_gain_loss_ils = net_sale_ils - original_cost_ils
    fx_adjustment_ils     = adjusted_cost_ils - original_cost_ils

    return {
        "fx_buy":                         fx_acq,
        "fx_sell":                        fx_sell,
        "fx_ratio":                       fx_sell / fx_acq,
        "index_increase_pct":             (fx_sell / fx_acq - 1) * 100,
        "original_cost_ils_exact":        original_cost_ils,
        "adjusted_cost_ils_exact":        adjusted_cost_ils,
        "gross_sale_proceeds_ils_exact":  gross_lot_ils,
        "sale_commission_ils_exact":      lot_comm_ils,
        "net_sale_consideration_ils_exact": net_sale_ils,
        "nominal_gain_loss_ils_exact":    nominal_gain_loss_ils,
        "fx_adjustment_ils_exact":        fx_adjustment_ils,
        "taxable_gain_ils_exact":         taxable_gain,
        "deductible_loss_ils_exact":      deductible_loss,
        "original_cost_ils_filing":       _round_ils(original_cost_ils),
        "adjusted_cost_ils_filing":       _round_ils(adjusted_cost_ils),
        "net_sale_consideration_ils_filing": _round_ils(net_sale_ils),
        "taxable_gain_ils_filing":        _round_ils(taxable_gain),
        "deductible_loss_ils_filing":     _round_ils(deductible_loss),
    }


def build_ibkr_lots(raw_lots: list[dict]) -> list[dict]:
    """
    Apply FX conversion and Moses to IBKR raw lot dicts from ibkr_parser.

    Each raw lot has: currency, acquisition_date, sale_date, acquisition_basis_fc,
    gross_sale_proceeds_fc, sale_commission_fc (signed: negative=cost).

    Returns fully-calculated lot dicts matching Tax Data schema.
    """
    result = []
    for raw in raw_lots:
        ccy      = raw["currency"]
        fx_acq   = boi_ils(ccy, raw["acquisition_date"])
        fx_sell  = boi_ils(ccy, raw["sale_date"])

        calc = _apply_moses(
            lot_basis_fc   = raw["acquisition_basis_fc"],
            lot_proceeds_fc= raw["gross_sale_proceeds_fc"],
            lot_comm_fc    = raw["sale_commission_fc"],
            fx_acq         = fx_acq,
            fx_sell        = fx_sell,
            comm_is_cost   = False,  # IBKR: signed commission
        )

        print(f"    {raw['sale_date']} {raw['ticker']} acq={raw['acquisition_date']} "
              f"qty={raw['quantity']} fx_acq={fx_acq:.4f} fx_sell={fx_sell:.4f} "
              f"→ taxable={calc['taxable_gain_ils_exact']:.2f} "
              f"loss={calc['deductible_loss_ils_exact']:.2f} ILS")

        lot = {
            "schema_version":         "1",
            "source":                 raw["source"],
            "ticker":                 raw["ticker"],
            "isin":                   raw.get("isin", ""),
            "currency":               ccy,
            "quantity":               raw["quantity"],
            "acquisition_date":       raw["acquisition_date"],
            "sale_date":              raw["sale_date"],
            "fill_id":                raw["fill_id"],
            "lot_id":                 raw["lot_id"],
            "acquisition_basis_fc":   raw["acquisition_basis_fc"],
            "gross_sale_proceeds_fc": raw["gross_sale_proceeds_fc"],
            "sale_commission_fc":     raw["sale_commission_fc"],
            "tax_rate":               TAX_RATE,
        }
        lot.update(calc)
        result.append(lot)

    return result


def build_equate_lots(raw_lots: list[dict]) -> list[dict]:
    """
    Apply FX conversion and Moses to EquatePlus raw lot dicts from equate_parser.

    Each raw lot has: acquisition_date, sale_date, acquisition_basis_fc,
    gross_sale_proceeds_fc, sale_commission_fc (positive = cost).

    Returns fully-calculated lot dicts matching Tax Data schema.
    """
    result = []
    for raw in raw_lots:
        fx_acq  = boi_ils("EUR", raw["acquisition_date"])
        fx_sell = boi_ils("EUR", raw["sale_date"])

        calc = _apply_moses(
            lot_basis_fc   = raw["acquisition_basis_fc"],
            lot_proceeds_fc= raw["gross_sale_proceeds_fc"],
            lot_comm_fc    = raw["sale_commission_fc"],
            fx_acq         = fx_acq,
            fx_sell        = fx_sell,
            comm_is_cost   = True,   # EquatePlus: positive fee deducted
        )

        print(f"    {raw['sale_date']} {raw['ticker']} acq={raw['acquisition_date']} "
              f"qty={raw['quantity']} fx_acq={fx_acq:.4f} fx_sell={fx_sell:.4f} "
              f"→ taxable={calc['taxable_gain_ils_exact']:.2f} "
              f"loss={calc['deductible_loss_ils_exact']:.2f} ILS")

        lot = {
            "schema_version":         "1",
            "source":                 raw["source"],
            "ticker":                 raw["ticker"],
            "isin":                   raw.get("isin", ""),
            "currency":               "EUR",
            "quantity":               raw["quantity"],
            "acquisition_date":       raw["acquisition_date"],
            "sale_date":              raw["sale_date"],
            "fill_id":                raw["sale_id"],   # sale_id plays role of fill_id for EquatePlus
            "lot_id":                 raw["lot_id"],
            "acquisition_basis_fc":   raw["acquisition_basis_fc"],
            "gross_sale_proceeds_fc": raw["gross_sale_proceeds_fc"],
            "sale_commission_fc":     raw["sale_commission_fc"],
            "tax_rate":               TAX_RATE,
        }
        lot.update(calc)
        result.append(lot)

    return result
