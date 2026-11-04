"""
Landed cost above index vs forecast, by supply route.

For every load:
    landed cost above index = (product cost - index price) + freight + overhead
Storage is shown separately. The result is compared with the forecast for the load's own month.

Things this handles that a naive pull gets wrong:
  * The warehouse "basis" label is coarse (HUB). The real index (TET vs Non-TET) comes from the
    contract table.
  * Pricing day follows the contract: same day, or prior business day.
  * The warehouse's own forecast column carries one month's value for every row, so it is ignored
    and each load is matched to the forecast for its own month.
  * Seasonal premiums switch on Oct 1. Each load's implied premium is checked against the contract.

Usage:
    python generate_sample_data.py          # build the synthetic data once
    python route_variance_report.py         # report from data/*.csv
    python route_variance_report.py --source snowflake   # same report from the SQL in ./sql
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
DATA_DIR = HERE / "data"
OUT_DIR = HERE / "Output"
SQL_DIR = HERE / "sql"

DEFAULT_INDEX = "Gulf Hub TET"          # used for contracts with no index (flat terminal cost)
ALT_INDEX = {"Gulf Hub TET": "Gulf Hub Non-TET", "Gulf Hub Non-TET": "Gulf Hub TET"}
PREMIUM_TOLERANCE = 0.0005

DARK = PatternFill("solid", fgColor="262626")
HFONT = Font(bold=True, color="FFFFFF")
GOOD = PatternFill("solid", fgColor="E2EFDA")
BAD = PatternFill("solid", fgColor="FCE4D6")
FLAG = PatternFill("solid", fgColor="FFF2CC")


# ----------------------------------------------------------------------------- inputs
def load_csv():
    names = ["loads", "contracts", "index_prices", "forecast"]
    missing = [n for n in names if not (DATA_DIR / f"{n}.csv").exists()]
    if missing:
        sys.exit(f"Missing data/{', data/'.join(m + '.csv' for m in missing)}. "
                 "Run: python generate_sample_data.py")
    return {n: pd.read_csv(DATA_DIR / f"{n}.csv") for n in names}


def load_snowflake():
    """Run the four queries in ./sql. Each must return the same columns as the matching CSV."""
    from snowflake_connect import get_connection

    conn = get_connection()
    cur = conn.cursor()
    try:
        tables = {}
        for name in ["loads", "contracts", "index_prices", "forecast"]:
            cur.execute((SQL_DIR / f"{name}.sql").read_text())
            df = cur.fetch_pandas_all()
            df.columns = [c.lower() for c in df.columns]
            tables[name] = df
        return tables
    finally:
        cur.close()
        conn.close()


def normalise(t):
    t["loads"]["order_date"] = pd.to_datetime(t["loads"]["order_date"])
    t["loads"]["year_month"] = pd.to_datetime(t["loads"]["year_month"])
    t["index_prices"]["price_date"] = pd.to_datetime(t["index_prices"]["price_date"])
    return t


# ----------------------------------------------------------------------------- calculation
def price_series(prices, name):
    return prices[prices["index_name"] == name].set_index("price_date")["price"].sort_index()


def compute(t):
    d = t["loads"].merge(t["contracts"], on="contract_id", how="left", validate="many_to_one")

    # Contracts with no index (flat terminal cost) are measured against the default hub, same day
    indexed = d["pricing_rule"].ne("none")
    d["index_used"] = d["index_name"].where(indexed, DEFAULT_INDEX)
    rule = d["pricing_rule"].where(indexed, "same_day")
    d["pricing_date"] = d["order_date"].where(rule.ne("prior_business_day"),
                                              d["order_date"] - pd.offsets.BDay(1))

    d["index_price"] = np.nan
    d["alt_index_gap"] = np.nan
    for name in d["index_used"].dropna().unique():
        m = d["index_used"].eq(name)
        when = pd.DatetimeIndex(d.loc[m, "pricing_date"])
        d.loc[m, "index_price"] = price_series(t["index_prices"], name).asof(when).to_numpy()
        if name in ALT_INDEX:
            alt = price_series(t["index_prices"], ALT_INDEX[name]).asof(when).to_numpy()
            d.loc[m, "alt_index_gap"] = alt - d.loc[m, "index_price"].to_numpy()

    d["differential"] = d["product_cost"] - d["index_price"]
    d["above_index"] = d["differential"] + d["freight"] + d["overhead"]

    # Forecast for each load's own month (not the warehouse column)
    d["month"] = d["year_month"].dt.strftime("%b")
    d["forecast"] = d["month"].map(t["forecast"].set_index("month")["forecast"])
    d["variance"] = d["above_index"] - d["forecast"]
    d["variance_usd"] = d["variance"] * d["gallons"]

    # Season check: indexed contracts should price at the premium for the load's season
    summer = d["year_month"].dt.month.between(4, 9)
    d["expected_premium"] = np.where(summer, d["summer_premium"], d["winter_premium"])
    bad = indexed & (d["differential"] - d["expected_premium"]).abs().gt(PREMIUM_TOLERANCE)
    d["premium_check"] = np.where(bad, "Premium differs from contract season", "")
    return d.sort_values(["order_date", "load_id"]).reset_index(drop=True)


def wavg(g, col):
    return float(np.average(g[col], weights=g["gallons"]))


def rollup(g):
    return dict(loads=len(g), gallons=g["gallons"].sum(), differential=wavg(g, "differential"),
                freight=wavg(g, "freight"), overhead=wavg(g, "overhead"),
                above_index=wavg(g, "above_index"), forecast=wavg(g, "forecast"),
                variance=wavg(g, "variance"), variance_usd=g["variance_usd"].sum())


# ----------------------------------------------------------------------------- excel
def clean(v):
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, np.generic):
        return v.item()
    return v


def put(ws, r, c, v, fmt=None, bold=False, fill=None):
    cell = ws.cell(row=r, column=c, value=clean(v))
    if fmt:
        cell.number_format = fmt
    if bold:
        cell.font = Font(bold=True)
    if fill:
        cell.fill = fill
    return cell


def header(ws, r, labels):
    for j, text in enumerate(labels, start=1):
        cell = ws.cell(row=r, column=j, value=text)
        cell.fill, cell.font = DARK, HFONT
        cell.alignment = Alignment(wrap_text=True, horizontal="center")


SUM_HEAD = ["Group", "Loads", "Gallons", "Differential", "Freight", "Overhead", "Above index",
            "Forecast", "Variance", "Variance $"]
SUM_FMT = [None, "0", "#,##0", "0.0000", "0.0000", "0.0000", "0.0000", "0.0000", "+0.0000;-0.0000",
           "$#,##0;[Red]-$#,##0"]


def summary_table(ws, r, title, groups):
    put(ws, r, 1, title, bold=True)
    r += 1
    header(ws, r, SUM_HEAD)
    first = r
    for label, g in groups + [("All loads", pd.concat([g for _, g in groups]))]:
        r += 1
        ru = rollup(g)
        vals = [label, ru["loads"], ru["gallons"], ru["differential"], ru["freight"], ru["overhead"],
                ru["above_index"], ru["forecast"], ru["variance"], ru["variance_usd"]]
        for j, (v, f) in enumerate(zip(vals, SUM_FMT), start=1):
            fill = (BAD if v > 0 else GOOD) if j in (9, 10) and not pd.isna(v) else None
            put(ws, r, j, v, fmt=f, fill=fill, bold=(label == "All loads"))
    return first, r


def write_workbook(d, t, district, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = f"{district}: landed cost above index vs forecast"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("$/gal, gallon-weighted. Variance = actual - forecast; positive means the district paid "
                "more than planned. Each load is compared with the forecast for its own month.")

    months = [(f"{g['year_month'].iloc[0]:%b %Y}", g) for _, g in d.groupby("year_month")]
    _, r = summary_table(ws, 4, "By month", months)

    routes = [(name, g) for name, g in d.groupby("route", sort=False)]
    first, last = summary_table(ws, r + 2, "By supply route (all months)", routes)

    chart = BarChart()
    chart.type, chart.title = "bar", "Landed cost above index vs forecast, by route ($/gal)"
    chart.height, chart.width = 8, 18
    chart.add_data(Reference(ws, min_col=7, max_col=8, min_row=first, max_row=last - 1),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first + 1, max_row=last - 1))
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    ws.add_chart(chart, "L4")

    r = last + 2
    put(ws, r, 1, "Checks", bold=True)
    flagged = d[d["premium_check"] != ""]
    if flagged.empty:
        put(ws, r + 1, 1, "All indexed loads match their contract premium for the season.")
    for _, row in flagged.iterrows():
        r += 1
        put(ws, r, 1, f"Load {row['load_id']} ({row['order_date']:%Y-%m-%d}, {row['route']}): "
                      f"priced at {row['differential']:+.3f} over {row['index_used']}, contract says "
                      f"{row['expected_premium']:+.3f} for the season.", fill=FLAG)
    ws.column_dimensions["A"].width = 30
    for j in range(2, 11):
        ws.column_dimensions[get_column_letter(j)].width = 13

    # ---- loads
    wl = wb.create_sheet("Loads")
    cols = [("Load", "load_id", "0"), ("Order date", "order_date", "yyyy-mm-dd"), ("Route", "route", None),
            ("Supplier", "supplier", None), ("Contract", "contract_id", None), ("Gallons", "gallons", "#,##0"),
            ("Index used", "index_used", None), ("Pricing date", "pricing_date", "yyyy-mm-dd"),
            ("Index price", "index_price", "0.00000"), ("Product cost", "product_cost", "0.00000"),
            ("Differential", "differential", "0.00000"), ("Freight", "freight", "0.00000"),
            ("Overhead", "overhead", "0.00000"), ("Above index", "above_index", "0.00000"),
            ("Forecast", "forecast", "0.00000"), ("Variance", "variance", "+0.00000;-0.00000"),
            ("Variance $", "variance_usd", "$#,##0.00;[Red]-$#,##0.00"), ("Storage", "storage", "0.00000"),
            ("Other index gap", "alt_index_gap", "+0.00000;-0.00000"), ("Premium check", "premium_check", None),
            ("Warehouse LIC col", "warehouse_lic", "0.00000")]
    header(wl, 1, [c[0] for c in cols])
    for i, (_, row) in enumerate(d.iterrows(), start=2):
        for j, (_, key, fmt) in enumerate(cols, start=1):
            put(wl, i, j, row[key], fmt=fmt, fill=FLAG if key == "premium_check" and row[key] else None)
    wl.freeze_panes = "C2"
    wl.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(d) + 1}"
    for j, (label, _, _) in enumerate(cols, start=1):
        wl.column_dimensions[get_column_letter(j)].width = 34 if label == "Premium check" else 13
    wl.column_dimensions["C"].width = 18
    wl.column_dimensions["D"].width = 22

    # ---- contracts
    wc = wb.create_sheet("Contracts")
    c = t["contracts"]
    header(wc, 1, list(c.columns))
    for i, (_, row) in enumerate(c.iterrows(), start=2):
        for j, v in enumerate(row, start=1):
            put(wc, i, j, v)
    for j in range(1, len(c.columns) + 1):
        wc.column_dimensions[get_column_letter(j)].width = 22

    # ---- method
    wm = wb.create_sheet("Method")
    notes = [
        "Landed cost above index = (product cost - index price) + freight + overhead. Storage is shown separately.",
        "The index comes from the contract table. The warehouse basis label only says HUB.",
        "Pricing day follows the contract rule: same day, or prior business day.",
        "Contracts with no index (flat terminal cost) are measured against the default hub on the order date.",
        "Each load is compared with the forecast for its own month. The warehouse LIC column carries one "
        "month's value on every row, so it is ignored.",
        "Seasonal premiums switch Oct 1. The premium check compares each load's implied premium with the "
        "contract premium for its season.",
        "'Other index gap' is how much the other hub (TET vs Non-TET) differs on the pricing date, i.e. the "
        "error if the wrong index were used.",
        "All data in this project is synthetic.",
    ]
    wm.column_dimensions["A"].width = 120
    for i, text in enumerate(notes, start=1):
        wm.cell(row=i, column=1, value=text).alignment = Alignment(wrap_text=True, vertical="top")

    try:
        wb.save(path)
    except PermissionError:
        path = path.with_name(f"{path.stem} {datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
        wb.save(path)
    return path


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Landed cost above index vs forecast, by route.")
    ap.add_argument("--source", choices=["csv", "snowflake"], default="csv")
    ap.add_argument("--district", help="only include districts containing this text")
    ap.add_argument("--out", help="output workbook path")
    a = ap.parse_args()

    tables = normalise(load_snowflake() if a.source == "snowflake" else load_csv())
    if a.district:
        tables["loads"] = tables["loads"][
            tables["loads"]["district"].str.contains(a.district, case=False, na=False)]
        if tables["loads"].empty:
            sys.exit(f"No loads match '{a.district}'.")

    d = compute(tables)
    district = d["district"].mode()[0]
    OUT_DIR.mkdir(exist_ok=True)
    slug = "".join(ch if ch.isalnum() else "_" for ch in district).strip("_")
    path = write_workbook(d, tables, district, Path(a.out) if a.out else OUT_DIR / f"landed_cost_variance_{slug}.xlsx")

    print(f"{district}: {len(d)} loads, {d['gallons'].sum():,.0f} gal")
    for label, g in [(f"{g['year_month'].iloc[0]:%b %Y}", g) for _, g in d.groupby("year_month")] + \
                    [(r, g) for r, g in d.groupby("route", sort=False)] + [("All loads", d)]:
        ru = rollup(g)
        print(f"  {label:18} loads={ru['loads']:3} above_index={ru['above_index']:.4f} "
              f"forecast={ru['forecast']:.4f} variance={ru['variance']:+.4f} (${ru['variance_usd']:,.0f})")
    flagged = d[d["premium_check"] != ""]
    print(f"Premium check flagged {len(flagged)} load(s): {flagged['load_id'].tolist()}")
    print(f"Report: {path}")


if __name__ == "__main__":
    main()
