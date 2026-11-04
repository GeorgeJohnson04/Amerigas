"""
Generate a fully synthetic dataset for the landed-cost variance demo.

Everything here is invented: the district, suppliers, locations, hubs, volumes and prices.
Nothing is derived from real company data.

Writes to ./data:
    loads.csv          one row per delivered load (mimics a warehouse landed-cost table)
    contracts.csv      supply contracts: index, pricing-day rule, seasonal premiums
    index_prices.csv   daily hub prices for two indices (TET and Non-TET)
    forecast.csv       monthly forecast of landed cost above index (fiscal year Apr-Mar)

Run:
    python generate_sample_data.py
"""
from pathlib import Path

import numpy as np
import pandas as pd

OUT = Path(__file__).resolve().parent / "data"
SEED = 42
DISTRICT = "5001 - SPRINGFIELD"
TET, NONTET = "Gulf Hub TET", "Gulf Hub Non-TET"
OVERHEAD = 0.065                         # flat supply fee, $/gal
TERMINAL_COST = {9: 0.905, 10: 0.915}    # flat terminal inventory cost by month, $/gal

MONTHS = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
FORECAST = [0.2440, 0.2430, 0.2360, 0.2380, 0.2410, 0.2400,
            0.3310, 0.3290, 0.3300, 0.3280, 0.3340, 0.3320]

CONTRACTS = pd.DataFrame([
    dict(contract_id="NOR26TP00001", supplier="Northstar Midstream", route="Delivered rail",
         index_name=TET, pricing_rule="same_day", summer_premium=0.150, winter_premium=0.250),
    dict(contract_id="PIN26TP00002", supplier="Pinecrest Supply", route="Remote rail",
         index_name=TET, pricing_rule="same_day", summer_premium=-0.040, winter_premium=0.040),
    dict(contract_id="BLU26TP00003", supplier="Bluewater Energy", route="Port spot trucks",
         index_name=NONTET, pricing_rule="prior_business_day", summer_premium=0.220, winter_premium=0.300),
    dict(contract_id="RED26TP00004", supplier="Redrock Terminals", route="Terminal trucks",
         index_name="None (terminal cost)", pricing_rule="none", summer_premium=np.nan, winter_premium=np.nan),
])


def make_prices(rng):
    """Two correlated daily hub series. Non-TET trades a little below TET."""
    days = pd.bdate_range("2026-08-24", "2026-10-09")
    tet = 0.72 + np.cumsum(rng.normal(0.002, 0.010, len(days)))
    nontet = tet - 0.018 + rng.normal(0, 0.003, len(days))
    wide = pd.DataFrame({"price_date": days, TET: tet, NONTET: nontet})
    long = wide.melt("price_date", var_name="index_name", value_name="price")
    long["price"] = long["price"].round(5)
    return long.sort_values(["index_name", "price_date"]).reset_index(drop=True)


def index_price(prices, index_name, when):
    s = prices[prices["index_name"] == index_name].set_index("price_date")["price"].sort_index()
    return float(s.asof(pd.Timestamp(when)))


def make_loads(rng, prices):
    by_route = CONTRACTS.set_index("route")
    sep_days = list(pd.bdate_range("2026-09-01", "2026-09-30"))
    rows = []

    def add(route, date, gallons, freight, origin, dest, move, storage=0.0, premium=None):
        k = by_route.loc[route]
        date = pd.Timestamp(date)
        if k["pricing_rule"] == "none":
            product = TERMINAL_COST[date.month]
        else:
            pricing_date = date - pd.offsets.BDay(1) if k["pricing_rule"] == "prior_business_day" else date
            if premium is None:
                premium = k["summer_premium"] if 4 <= date.month <= 9 else k["winter_premium"]
            product = index_price(prices, k["index_name"], pricing_date) + premium
        rows.append(dict(
            order_date=date, origin=origin, destination=dest, move_type=move,
            gallons=int(gallons), contract_id=k["contract_id"],
            product_cost=round(product, 5), overhead=OVERHEAD, storage=storage,
            freight=round(freight, 5)))

    # Delivered rail: freight is inside the delivered price
    for i in np.sort(rng.choice(len(sep_days), 9, replace=False)):
        add("Delivered rail", sep_days[i], rng.integers(29000, 31200), 0.0,
            "Lakeside Rail Hub - DELIVERED", "Springfield Plant", "Railcar")
    for d in ("2026-10-01", "2026-10-02"):
        add("Delivered rail", d, rng.integers(29000, 31200), 0.0,
            "Lakeside Rail Hub - DELIVERED", "Springfield Plant", "Railcar")

    # Remote rail: freight is a fixed cost per car, so $/gal depends on gallons shipped
    for i in np.sort(rng.choice(len(sep_days), 4, replace=False)):
        gallons = int(rng.integers(28000, 30000))
        car_freight = rng.normal(8200, 150)
        add("Remote rail", sep_days[i], gallons, car_freight / gallons,
            "Ridgeway Rail Yard", "Springfield Plant", "Railcar")

    # Terminal trucks: flat terminal cost plus storage
    for i in np.sort(rng.choice(len(sep_days), 12, replace=True)):
        add("Terminal trucks", sep_days[i], rng.integers(8000, 9600), 0.052 + rng.normal(0, 0.004),
            "Riverton Terminal", rng.choice(["Springfield Satellite A", "Springfield Satellite B"]),
            "Truck", storage=0.04)
    for d in ("2026-10-01", "2026-10-02", "2026-10-05"):
        add("Terminal trucks", d, rng.integers(8000, 9600), 0.052 + rng.normal(0, 0.004),
            "Riverton Terminal", rng.choice(["Springfield Satellite A", "Springfield Satellite B"]),
            "Truck", storage=0.04)

    # Port spot trucks (Non-TET, prior business day). The Sep 29 load is deliberately priced with the
    # winter premium so the season check has something to catch.
    add("Port spot trucks", "2026-09-15", 8650, 0.050, "Harbor Port", "Springfield Satellite A", "Truck")
    add("Port spot trucks", "2026-09-29", 8550, 0.052, "Harbor Port", "Springfield Satellite B", "Truck",
        premium=0.300)
    add("Port spot trucks", "2026-10-06", 8600, 0.051, "Harbor Port", "Springfield Satellite A", "Truck")

    loads = pd.DataFrame(rows).sort_values("order_date").reset_index(drop=True)
    loads.insert(0, "load_id", 9000001 + np.arange(len(loads)))
    loads.insert(2, "year_month", loads["order_date"].dt.to_period("M").dt.to_timestamp())
    loads.insert(3, "district", DISTRICT)
    loads["grand_total"] = (loads[["product_cost", "overhead", "storage", "freight"]].sum(axis=1)).round(5)
    loads["basis"] = "HUB"   # coarse warehouse label: says nothing about TET vs Non-TET
    # The warehouse stamps every row with the forecast for the snapshot's *current* month (Oct)
    loads["warehouse_lic"] = FORECAST[MONTHS.index("Oct")]
    return loads


def main():
    rng = np.random.default_rng(SEED)
    OUT.mkdir(exist_ok=True)
    prices = make_prices(rng)
    loads = make_loads(rng, prices)
    forecast = pd.DataFrame({"district": DISTRICT, "month": MONTHS, "forecast": FORECAST})

    loads.to_csv(OUT / "loads.csv", index=False)
    CONTRACTS.to_csv(OUT / "contracts.csv", index=False)
    prices.to_csv(OUT / "index_prices.csv", index=False)
    forecast.to_csv(OUT / "forecast.csv", index=False)
    print(f"Wrote {len(loads)} loads, {len(CONTRACTS)} contracts, {len(prices)} price rows to {OUT}")


if __name__ == "__main__":
    main()
