"""
District - Outliers vs Non-Outliers
===================================
Pulls the zone-detail table from Snowflake, keeps every row for one district/location
(District, OriginLocation or DestinationLocation contains the --match text) and splits them
into "<Match> Outliers vs Non-Outliers.xlsx" with charts.

A row is an OUTLIER if any cost/volume column is unusual by either test:
  1. Local  - vs other loads on the same lane (Origin -> Destination, MoveType); lanes with
              fewer than 4 loads are compared with all of the district's loads of that MoveType.
              Outside Q1 - 1.5 x IQR .. Q3 + 1.5 x IQR AND at least 5% from the median.
  2. Wider  - the main report's test: outside 3 x IQR for both its State+MoveType and its
              MoveType across the whole table.
Data-rule problems (Qty x GrandTotal mismatch, duplicate PO, blank IDs...) are listed in
an "Other data issues" column but do not by themselves make a row an outlier.

Usage (from this folder):
    py district_report.py --match "SPRINGFIELD"                       # fresh pull from Snowflake
    py district_report.py --match "SPRINGFIELD" --from-file zone_detail_raw.parquet   # reuse the last pull
"""

import argparse
import datetime as dt
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font

import zone_detail_report as base

HERE = Path(__file__).resolve().parent
TEST_COLS = ["Qty", "Product", "Overhead", "Storage", "Transportation", "GrandTotal", "TotalZoneValue"]
MATCH_COLS = ["District", "OriginLocation", "DestinationLocation"]
LOCAL_MIN = 4       # smallest group used for the local test
LOCAL_K = 1.5       # local IQR multiplier
MIN_REL = 0.05      # local outlier must also be >= 5% from the group median
FRONT_COLS = ["District", "OrderDate", "OrderID", "TMWOrderNumber", "PONumber", "MoveType",
              "OriginLocation", "DestinationLocation", *TEST_COLS, "Basis", "Price", "LIC"]
FMT = {"Qty": "#,##0", "TotalZoneValue": "$#,##0.00", "Product": "0.00000", "Overhead": "0.00000",
       "Storage": "0.00000", "Transportation": "0.00000", "GrandTotal": "0.00000", "Price": "0.00000",
       "LIC": "0.0000"}
RED, LIGHT_RED, YELLOW, GREEN = "C00000", "FFC7CE", "FFEB9C", "C6EFCE"


def fmt_num(c, v):
    if pd.isna(v):
        return ""
    if c == "Qty":
        return f"{v:,.0f}"
    if c == "TotalZoneValue":
        return f"${v:,.2f}"
    return f"{v:.5f}"


def lane_of(d):
    return d["OriginLocation"].astype(str) + " -> " + d["DestinationLocation"].astype(str) + \
        " (" + d["MoveType"].astype(str) + ")"


def local_tests(sub):
    """Return (flags list, baseline rows) for the local lane / district tests."""
    lane = lane_of(sub)
    dmt = sub["District"].astype(str) + " - all " + sub["MoveType"].astype(str)
    lane_n, dmt_n = lane.map(lane.value_counts()), dmt.map(dmt.value_counts())
    group = np.where(lane_n >= LOCAL_MIN, lane, np.where(dmt_n >= LOCAL_MIN, dmt, ""))
    group = pd.Series(group, index=sub.index)
    flags, baseline = [], []
    for g in sorted(set(group) - {""}):
        members = group.index[group == g]
        # stats come from every row in the group definition (lane or district+MoveType)
        pool = sub.index[lane == g] if (lane == g).any() else sub.index[dmt == g]
        for c in TEST_COLS:
            v = pd.to_numeric(sub.loc[pool, c], errors="coerce").dropna()
            if v.empty:
                continue
            q1, q3, med = v.quantile(0.25), v.quantile(0.75), v.median()
            iqr = q3 - q1
            lo, hi = q1 - LOCAL_K * iqr, q3 + LOCAL_K * iqr
            baseline.append((g, c, len(v), med, lo, hi))
            for i in members:
                x = pd.to_numeric(sub.at[i, c], errors="coerce")
                if pd.isna(x) or lo <= x <= hi:
                    continue
                rel = (x - med) / abs(med) if med else np.inf
                if med and abs(rel) < MIN_REL:
                    continue
                pct = f"{rel:+.1%}" if np.isfinite(rel) else "vs 0"
                flags.append((i, c, "High" if x > hi else "Low",
                              f"{c} {fmt_num(c, x)} is {pct} vs {g} median {fmt_num(c, med)} "
                              f"(normal {fmt_num(c, lo)} to {fmt_num(c, hi)})"))
    skipped = sub.index[group == ""]
    return flags, baseline, skipped


def build(df, match, source, out_path):
    sub, cell_out, cell_other, baseline, skipped = find_outliers(df, match)
    data_cols = [c for c in FRONT_COLS if c in sub.columns] + \
                [c for c in df.columns if c not in FRONT_COLS]

    wb = Workbook()
    _summary(wb.active, sub, baseline, skipped, match, source)
    for name, status in (("Outliers", "Outlier"), ("Non-Outliers", "Non-Outlier")):
        _rows_sheet(wb.create_sheet(name), sub[sub["Status"] == status], data_cols, cell_out, cell_other,
                    with_reasons=(status == "Outlier"))
    _charts(wb.create_sheet("Charts"), sub, cell_out)

    try:
        wb.save(out_path)
    except PermissionError:
        out_path = out_path.with_name(f"{out_path.stem} {dt.datetime.now():%Y%m%d_%H%M%S}{out_path.suffix}")
        wb.save(out_path)
        print("Report was open in Excel - saved a copy instead.")
    n_out = (sub["Status"] == "Outlier").sum()
    print(f"Outliers: {n_out}   Non-outliers: {len(sub) - n_out}")
    for _, rec in sub[sub["Status"] == "Outlier"].iterrows():
        print(f"  Order {rec['OrderID']} {rec['MoveType']} {rec['OrderDate']:%m/%d}: {rec['Outlier reasons']}")
    print(f"Report: {out_path}")


def find_outliers(df, match):
    """Rows matching `match`, with Status / Outlier reasons / Other data issues / Lane columns added."""
    pat = re.escape(match)
    hit = pd.concat([df[c].astype(str).str.contains(pat, case=False, regex=True)
                     for c in MATCH_COLS if c in df.columns], axis=1).any(axis=1)
    sub = df[hit].copy()
    if sub.empty:
        sys.exit(f"No rows match '{match}'.")
    print(f"{len(sub)} rows match '{match}'")

    # whole-table tests (same as the main report), then keep this subset's flags
    work, kinds, notes = base.classify(df)
    all_flags, _ = base.find_issues(df, work, kinds, notes)
    mine = all_flags[all_flags["row"].isin(sub.index)]
    wide = mine[mine["issue"].str.startswith("Outlier") & mine["column"].isin(TEST_COLS)]
    other = mine[~mine.index.isin(wide.index)]

    local, baseline, skipped = local_tests(sub)

    reasons = {i: [] for i in sub.index}
    cell_out = set()
    for i, c, hl, txt in local:
        reasons[i].append(f"{hl} {txt}")
        cell_out.add((i, c))
    for r in wide.itertuples():
        reasons[r.row].append(f"{r.issue.replace('Outlier ', '').strip('()').title()} {r.column} "
                              f"{fmt_num(r.column, df.at[r.row, r.column])} vs wider table ({r.detail})")
        cell_out.add((r.row, r.column))
    other_txt = other.groupby("row").apply(
        lambda g: "; ".join(f"{a}: {b}" + (f" ({d})" if d else "")
                            for a, b, d in zip(g["issue"], g["column"], g["detail"])),
        include_groups=False) if len(other) else pd.Series(dtype=object)
    cell_other = {(r.row, r.column) for r in other.itertuples()}

    sub["Status"] = ["Outlier" if reasons[i] else "Non-Outlier" for i in sub.index]
    sub["Outlier reasons"] = [" | ".join(reasons[i]) for i in sub.index]
    sub["Other data issues"] = [other_txt.get(i, "") for i in sub.index]
    sub["Lane"] = lane_of(sub)
    sort = [c for c in ("MoveType", "Lane", "OrderDate", "OrderID") if c in sub.columns]
    sub = sub.sort_values(sort)
    return sub, cell_out, cell_other, baseline, skipped


def _summary(ws, sub, baseline, skipped, match, source):
    ws.title = "Summary"
    ws["A1"] = f"{sub['District'].mode()[0] if 'District' in sub else match} - Outliers vs Non-Outliers"
    ws["A1"].font = Font(bold=True, size=14)
    n_out = int((sub["Status"] == "Outlier").sum())
    facts = [("Source", source), ("Built", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
             ("Rows matched", f"{len(sub)}  (District / Origin / Destination contains '{match}')"),
             ("Outliers", n_out), ("Non-outliers", len(sub) - n_out),
             ("Order dates", f"{sub['OrderDate'].min():%Y-%m-%d} to {sub['OrderDate'].max():%Y-%m-%d}")]
    for i, (k, v) in enumerate(facts, start=3):
        base.put(ws, i, 1, k, bold=True)
        base.put(ws, i, 2, v)

    r = 10
    base.put(ws, r, 1, "By lane", bold=True)
    r += 1
    hdr = r
    base.header(ws, r, 1, ["Lane", "Loads", "Non-Outliers", "Outliers", "Avg Qty",
                           "Avg GrandTotal (non-outliers)", "Avg GrandTotal (outliers)", "Total value"])
    for lane, g in sub.groupby("Lane", sort=False):
        r += 1
        ok, bad = g[g["Status"] == "Non-Outlier"], g[g["Status"] == "Outlier"]
        vals = [lane, len(g), len(ok), len(bad), g["Qty"].mean(),
                ok["GrandTotal"].mean() if len(ok) else None,
                bad["GrandTotal"].mean() if len(bad) else None, g["TotalZoneValue"].sum()]
        for j, (v, f) in enumerate(zip(vals, [None, "0", "0", "0", "#,##0", "0.00000", "0.00000",
                                              "$#,##0"]), start=1):
            base.put(ws, r, j, v, fmt=f)
    last = r
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "bar", "stacked", 100
    ch.title, ch.height, ch.width = "Loads per lane: non-outliers vs outliers", 7.5, 20
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=hdr, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=hdr + 1, max_row=last))
    ch.series[0].graphicalProperties.solidFill = base.BLUE
    ch.series[1].graphicalProperties.solidFill = RED
    base._axes_on(ch)
    ch.x_axis.scaling.orientation = "maxMin"
    ch.legend.position = "b"
    ws.add_chart(ch, "J3")

    r += 2
    base.put(ws, r, 1, "How outliers are decided", bold=True)
    notes = [
        f"Local: each load vs other loads on the same lane (lanes with < {LOCAL_MIN} loads use the "
        f"district's loads of the same MoveType). Outside Q1 - {LOCAL_K:g} x IQR .. Q3 + {LOCAL_K:g} x IQR "
        f"and at least {MIN_REL:.0%} from the median.",
        "Wider: outside 3 x IQR for both its State+MoveType and its MoveType across the whole zone-detail table.",
        f"Columns tested: {', '.join(TEST_COLS)}. Price (market index, moves by date) and LIC are shown, not tested.",
        "Red cells = the value that made the row an outlier. Yellow cells = other data issues (not outliers).",
    ]
    if len(skipped):
        notes.append(f"{len(skipped)} load(s) had no local group of {LOCAL_MIN}+ and only got the wider test.")
    for t in notes:
        r += 1
        c = base.put(ws, r, 1, t)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30

    r += 2
    base.put(ws, r, 1, "Normal ranges used for the local test", bold=True)
    r += 1
    base.header(ws, r, 1, ["Group", "Column", "Loads", "Median", "Normal low", "Normal high"])
    for g, c, n, med, lo, hi in baseline:
        r += 1
        f = FMT.get(c)
        for j, (v, ff) in enumerate(zip([g, c, n, med, lo, hi], [None, None, "0", f, f, f]), start=1):
            base.put(ws, r, j, v, fmt=ff)
    base.widths(ws, {"A": 62, "B": 16, "C": 13, "D": 10, "E": 11, "F": 16, "G": 16, "H": 14})


def _rows_sheet(ws, rows, data_cols, cell_out, cell_other, with_reasons):
    lead = (["Outlier reasons"] if with_reasons else []) + ["Other data issues"]
    cols = lead + data_cols
    base.header(ws, 1, 1, cols)
    if rows.empty:
        base.put(ws, 2, 1, "None", bold=True)
    for r, (i, rec) in enumerate(rows.iterrows(), start=2):
        for j, c in enumerate(cols, start=1):
            color = LIGHT_RED if (i, c) in cell_out else YELLOW if (i, c) in cell_other else None
            if c == "Outlier reasons":
                color = LIGHT_RED
            cell = base.put(ws, r, j, rec[c] or None if c in lead else rec[c], color=color, fmt=FMT.get(c))
            if c in lead:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = ws.cell(row=2, column=len(lead) + 1)
    ws.auto_filter.ref = f"A1:{base.get_column_letter(len(cols))}{max(len(rows), 1) + 1}"
    for j, c in enumerate(cols, start=1):
        w = 70 if c == "Outlier reasons" else 40 if c == "Other data issues" else \
            34 if "Location" in c or c == "District" else 13
        ws.column_dimensions[base.get_column_letter(j)].width = w


def _charts(ws, sub, cell_out):
    base.put(ws, 1, 1, "Every load, in lane / date order. Red bar = that value made the load an outlier; "
                       "blue = normal.", bold=True)
    labels = [f"{str(m)[:4]} {d:%m/%d} #{o}" for m, d, o in zip(sub["MoveType"], sub["OrderDate"], sub["OrderID"])]
    base.header(ws, 3, 1, ["Load", "Lane", "Status", *TEST_COLS])
    for r, (i, lab) in enumerate(zip(sub.index, labels), start=4):
        base.put(ws, r, 1, lab)
        base.put(ws, r, 2, sub.at[i, "Lane"])
        base.put(ws, r, 3, sub.at[i, "Status"], color=LIGHT_RED if sub.at[i, "Status"] == "Outlier" else GREEN)
        for j, c in enumerate(TEST_COLS, start=4):
            base.put(ws, r, j, sub.at[i, c], fmt=FMT.get(c),
                     color=LIGHT_RED if (i, c) in cell_out else None)
    last = 3 + len(sub)
    anchor_row = 3
    for j, c in enumerate(TEST_COLS, start=4):
        red = {k: RED for k, i in enumerate(sub.index) if (i, c) in cell_out}
        ch = base.bar_chart(ws, f"{c} by load" + (f"  ({len(red)} outlier)" if red else ""), 1, j, 3, last,
                            colors=red)
        ch.width, ch.height = 30, 8
        ch.x_axis.tickLblSkip = 1
        ws.add_chart(ch, f"L{anchor_row}")
        anchor_row += 17
    base.widths(ws, {"A": 22, "B": 60, "C": 12})


def main():
    ap = argparse.ArgumentParser(description="Outliers vs non-outliers for one location/district.")
    ap.add_argument("--match", required=True,
                    help='text to find in District/Origin/Destination, e.g. "SPRINGFIELD"')
    ap.add_argument("--from-file", help="use a saved pull (.parquet/.csv/.xlsx) instead of Snowflake")
    ap.add_argument("--out", help="output workbook")
    a = ap.parse_args()

    if a.from_file:
        df = base.load_from_file(a.from_file)
        source = f"file: {Path(a.from_file).name}"
    else:
        df, table, total = base.load_from_snowflake(base.DEFAULT_TABLE, None, 10_000_000)
        source = f"Snowflake: {table.replace(chr(34), '')} ({total:,} rows), pulled {dt.datetime.now():%Y-%m-%d %H:%M}"
        base.archive_pull(df)
    df = df.reset_index(drop=True)
    if "SnapshotInstanceDateTime" in df.columns:
        source += f"; snapshot {pd.to_datetime(df['SnapshotInstanceDateTime']).max():%Y-%m-%d %H:%M}"

    name = a.match.title()
    out = Path(a.out) if a.out else HERE / f"{name} Outliers vs Non-Outliers.xlsx"
    build(df, a.match, source, out)


if __name__ == "__main__":
    main()
