#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime as dt
from decimal import Decimal, ROUND_HALF_UP

import requests

TAX_RATE = 0.25

_fx_cache: dict = {}


def _round_ils(value: float) -> int:
    """Arithmetic rounding to whole shekels (ROUND_HALF_UP). Single authoritative implementation."""
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def moses_gain_loss(
    original_cost_ils: float,
    net_sale_ils: float,
    fx_buy: float,
    fx_sell: float,
) -> tuple[float, float]:
    """
    Israeli capital-gain/loss per ITO Section 91(b), Form 1325 (2025), Circular 10/2025.
    FX rate acts as an index (מדד). When FX rises, inflationary component is exempt.
    When FX falls, negative inflationary component is treated as zero.
    Returns (taxable_gain, deductible_loss) both >= 0.
    """
    adjusted_cost_ils = original_cost_ils * (fx_sell / fx_buy)
    nominal_result = net_sale_ils - original_cost_ils

    if nominal_result > 0:
        if fx_sell >= fx_buy:
            inflationary = adjusted_cost_ils - original_cost_ils
            exempt = min(nominal_result, max(0.0, inflationary))
            taxable_gain = nominal_result - exempt
        else:
            taxable_gain = nominal_result
        deductible_loss = 0.0
    elif nominal_result < 0:
        taxable_gain = 0.0
        nominal_loss = -nominal_result
        if fx_sell >= fx_buy:
            deductible_loss = nominal_loss
        else:
            neg_inflationary = original_cost_ils - adjusted_cost_ils
            deductible_loss = max(0.0, nominal_loss - neg_inflationary)
    else:
        taxable_gain = 0.0
        deductible_loss = 0.0

    return taxable_gain, deductible_loss


def boi_ils(currency: str, day: str | dt.date) -> float:
    """
    Bank of Israel representative rate for currency/ILS on the given day.
    Walks back up to 7 days to skip weekends and Israeli holidays.
    Results are cached in _fx_cache.
    """
    if isinstance(day, dt.date):
        day = day.strftime("%Y-%m-%d")
    key = (currency.upper(), day)
    if key in _fx_cache:
        return _fx_cache[key]
    url = (
        f"https://edge.boi.gov.il/FusionEdgeServer/"
        f"sdmx/v2/data/dataflow/BOI.STATISTICS/EXR/1.0/RER_{currency.upper()}_ILS"
    )
    d = dt.date.fromisoformat(day)
    for delta in range(7):
        candidate = (d - dt.timedelta(days=delta)).strftime("%Y-%m-%d")
        params = {
            "c[DATA_TYPE]": "OF00",
            "startperiod": candidate,
            "endperiod": candidate,
            "format": "sdmx-json",
        }
        resp = requests.get(url, params=params, timeout=10)
        if resp.status_code != 200 or not resp.text.strip():
            continue
        try:
            js = resp.json()
            obs = js["data"]["dataSets"][0]["series"]["0:0:0:0:0:0"]["observations"]
            rate = float(list(obs.values())[0][0])
            _fx_cache[key] = rate
            return rate
        except (KeyError, IndexError, ValueError):
            continue
    raise RuntimeError(f"No BoI {currency}/ILS rate found within 7 days before {day}")
