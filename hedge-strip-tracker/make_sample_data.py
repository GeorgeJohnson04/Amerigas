"""
Generate SYNTHETIC inputs so hedge_tracker.py can be run without real market data.

Writes (paths taken from hedge_products.json):
    data/us_propane_priceindex.parquet   random-walk forward curves for the three
                                         hubs plus the blend, with winter seasonality
    data/demand_shapes.xlsx              one monthly demand-shape block per product

Nothing here is real pricing. The curves are a seeded random walk so results
are reproducible.

Usage:
    py make_sample_data.py
    py make_sample_data.py --start 2023-01-02 --end 2026-09-14 --seed 7
"""
from __future__ import annotations

import argparse
import os

import numpy as np
import pandas as pd

from build_price_index import BLEND_NAME, SCHEMA, build_blend
from hedge_tracker import DEFAULT_CONFIG, load_config

MONTHS_FULL = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]

# Illustrative demand shapes (relative weights by calendar month)
SHAPES = {
    "Annual Fixed":  [14, 13, 11, 7, 5, 4, 4, 4, 5, 8, 11, 14],
    "Annual Capped": [15, 13, 11, 7, 4, 3, 3, 3, 4, 8, 13, 16],
    "Winter":        [24, 20, 15, 0, 0, 0, 0, 0, 0, 8, 14, 19],
}

# Hub level ($/gal) and basis vs. Mont Belvieu Enterprise
HUBS = {
    "NGL fixed price - Financial Mont Belvieu Propane Enterprise": 0.000,
    "NGL fixed price - Financial Mont Belvieu Propane LST": -0.004,
    "NGL fixed price - Financial Conway Propane": -0.030,
}


def synthetic_prices(start: str, end: str, horizon_months: int, seed: int) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    dates = pd.bdate_range(start, end)
    # front-month level follows a mean-reverting random walk around $0.75
    level = np.empty(len(dates))
    level[0] = 0.75
    for i in range(1, len(dates)):
        level[i] = level[i - 1] + 3.0 * (0.75 - level[i - 1]) / 250 + rng.normal(0, 0.008)
    level = np.clip(level, 0.35, 1.40)

    rows = []
    for d, lvl in zip(dates, level):
        first = d.to_period("M").to_timestamp()
        for k in range(horizon_months):
            deliv = first + pd.DateOffset(months=k)
            # winter premium, mild contango-to-flat term structure
            season = 0.04 * np.cos(2 * np.pi * (deliv.month - 1) / 12)
            term = -0.0015 * k
            base = lvl + season + term
            for hub, basis in HUBS.items():
                noise = rng.normal(0, 0.002)
                rows.append((hub, d, deliv, round(base + basis + noise, 5)))
    df = pd.DataFrame(rows, columns=["priceindex", "pricedate", "delivdate", "price"])
    df["commodity"] = "LPG"
    df["basis"] = 0.0
    df["index"] = df["price"]
    return df[SCHEMA]


def write_shapes(path: str) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Shapes"
    r = 1
    for label, weights in SHAPES.items():
        ws.cell(row=r, column=1, value=label)
        for c, (m, w) in enumerate(zip(MONTHS_FULL, weights), start=1):
            ws.cell(row=r + 1, column=c, value=m)
            ws.cell(row=r + 2, column=c, value=w / sum(weights))
        r += 5
    wb.save(path)


def main() -> int:
    ap = argparse.ArgumentParser(description="Write synthetic price and shape inputs.")
    ap.add_argument("--config", default=DEFAULT_CONFIG)
    ap.add_argument("--start", default="2023-01-02")
    ap.add_argument("--end", default="2026-09-14")
    ap.add_argument("--horizon", type=int, default=60, help="forward months per price date")
    ap.add_argument("--seed", type=int, default=7)
    a = ap.parse_args()

    cfg = load_config(a.config)
    os.makedirs(os.path.dirname(cfg["price_parquet"]), exist_ok=True)
    os.makedirs(os.path.dirname(cfg["shape_workbook"]), exist_ok=True)

    comps = synthetic_prices(a.start, a.end, a.horizon, a.seed)
    weights = {o["index"]: o["weight"] for o in cfg["origin_blend"]}
    blend = build_blend(comps, name=cfg.get("blend_check_index", BLEND_NAME),
                        weights=weights, verbose=False)
    out = pd.concat([comps, blend], ignore_index=True)
    out.to_parquet(cfg["price_parquet"], index=False)
    print(f"wrote {cfg['price_parquet']} ({len(out):,} rows, "
          f"{out['pricedate'].min():%Y-%m-%d} .. {out['pricedate'].max():%Y-%m-%d})")

    write_shapes(cfg["shape_workbook"])
    print(f"wrote {cfg['shape_workbook']} ({', '.join(SHAPES)})")
    labels = {p["shape_label"] for p in cfg["products"]}
    missing = labels - set(SHAPES)
    if missing:
        print(f"WARNING: config uses shape labels not generated here: {sorted(missing)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
