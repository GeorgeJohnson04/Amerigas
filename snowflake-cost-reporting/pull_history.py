"""
Pull zone-detail raw data for as many months as possible and write it to Excel (no analysis).

The zone-detail table only holds the CURRENT snapshot (the whole table is replaced on each refresh)
and Snowflake Time Travel keeps 1 day. So this script:
  1. pulls the current table,
  2. walks back through the Time Travel window and pulls every other snapshot still reachable,
  3. archives each snapshot to snapshots\\zone_detail_snapshot_<id>_<date>.parquet,
  4. combines every archived snapshot, keeps YearMonth in the chosen year, and writes
     "Zone Detail Raw <year>.xlsx".

When the same OrderID appears in more than one snapshot, the row from the newest snapshot is kept
(the "SourceSnapshotID" column shows which snapshot each row came from).

Table and warehouse come from SNOWFLAKE_TABLE / SNOWFLAKE_WAREHOUSE (see zone_detail_report.py).

Usage (from this folder):
    py pull_history.py              # pull + build workbook for 2026
    py pull_history.py --no-pull    # rebuild the workbook from archived snapshots only
    py pull_history.py --year 2026
Run it daily to keep a copy of each snapshot before Snowflake replaces it.
"""

import argparse
import datetime as dt
from pathlib import Path

import pandas as pd

import zone_detail_report as base

HERE = Path(__file__).resolve().parent
TABLE = base.DEFAULT_TABLE  # from SNOWFLAKE_TABLE, e.g. ANALYTICS_DB.REPORTING.ZONE_DETAIL
TT_HOURS = list(range(0, 25, 2))  # Time Travel probes: now, 2h ago ... 24h ago


def pull_snowflake():
    import re

    from snowflake_connect import get_connection

    if not re.fullmatch(r'[A-Za-z0-9_$."]+', TABLE):
        raise SystemExit(f"Table name looks invalid: {TABLE!r}")
    print("Connecting to Snowflake (a browser window may open for SSO sign-in)...")
    conn = get_connection()
    cur = conn.cursor()
    saved = []
    try:
        cur.execute(f"USE WAREHOUSE {base._q(base.DEFAULT_WAREHOUSE)}")
        cur.execute(f'SELECT "SnapshotID", COUNT(*) FROM {TABLE} GROUP BY 1')
        current = {int(s): n for s, n in cur.fetchall()}
        print(f"Current table: snapshot(s) {current}")

        # which snapshot was in the table at each point in the Time Travel window
        found = {}
        for h in TT_HOURS:
            at = f"AT(OFFSET => -{h * 3600})" if h else ""
            try:
                cur.execute(f'SELECT DISTINCT "SnapshotID" FROM {TABLE} {at}')
                for (s,) in cur.fetchall():
                    found.setdefault(int(s), h)
            except Exception as e:
                print(f"  {h}h ago: not available ({str(e).splitlines()[0][:90]})")
                break
        print(f"Snapshots reachable now (snapshot -> hours ago): {found}")

        for sid, h in sorted(found.items()):
            existing = list(base.SNAP_DIR.glob(f"{base.SNAP_PREFIX}{sid}_*.parquet"))
            if existing:
                print(f"  snapshot {sid}: already archived ({existing[0].name})")
                continue
            at = f"AT(OFFSET => -{h * 3600})" if h else ""
            cur.execute(f'SELECT * FROM {TABLE} {at} WHERE "SnapshotID" = {sid}')
            df = cur.fetch_pandas_all()
            if h == 0:
                path = base.archive_pull(df)  # also refreshes zone_detail_raw.parquet
            else:
                base.SNAP_DIR.mkdir(exist_ok=True)
                when = pd.to_datetime(df["SnapshotInstanceDateTime"]).max()
                path = base.SNAP_DIR / f"{base.SNAP_PREFIX}{sid}_{when:%Y-%m-%d}.parquet"
                df.to_parquet(path, index=False)
                print(f"Archived snapshot {sid} ({h}h ago via Time Travel) -> snapshots\\{path.name}")
            saved.append(path)
    finally:
        cur.close()
        conn.close()
    return saved


def combine(year):
    files = sorted(base.SNAP_DIR.glob(f"{base.SNAP_PREFIX}*.parquet"))
    if not files:
        raise SystemExit("No archived snapshots in snapshots\\ - run without --no-pull first.")
    frames = []
    for f in files:
        d = pd.read_parquet(f)
        d["SourceSnapshotID"] = int(pd.to_numeric(d["SnapshotID"], errors="coerce").max())
        d["SourceSnapshotTime"] = pd.to_datetime(d["SnapshotInstanceDateTime"]).max()
        frames.append(d)
        print(f"  {f.name}: {len(d):,} rows, months "
              f"{sorted(pd.to_datetime(d['YearMonth']).dt.strftime('%Y-%m').unique())}")
    allrows = pd.concat(frames, ignore_index=True)
    allrows["YearMonth"] = pd.to_datetime(allrows["YearMonth"])

    # newest snapshot wins when an order appears in several snapshots
    allrows = allrows.sort_values(["SourceSnapshotID"]).drop_duplicates(
        subset=["OrderID", "Type", "YearMonth"], keep="last")
    other = allrows[allrows["YearMonth"].dt.year != year]
    data = allrows[allrows["YearMonth"].dt.year == year]
    data = data.sort_values(["YearMonth", "Type", "District", "OrderDate", "OrderID"]).reset_index(drop=True)
    if len(other):
        print(f"  left out {len(other)} row(s) outside {year}: "
              f"{other['YearMonth'].dt.strftime('%Y-%m').value_counts().to_dict()}")
    return data, files


def write(data, files, year):
    out = HERE / f"Zone Detail Raw {year}.xlsx"
    srcs = (data.groupby([data["YearMonth"].dt.strftime("%Y-%m").rename("YearMonth"), "Type",
                          "SourceSnapshotID", "SourceSnapshotTime"])
            .size().rename("Rows").reset_index())
    tz_cols = [c for c in data.columns if isinstance(data[c].dtype, pd.DatetimeTZDtype)]
    for c in tz_cols:
        data[c] = data[c].dt.tz_localize(None)
    try:
        writer = pd.ExcelWriter(out, engine="openpyxl")
    except PermissionError:
        out = out.with_name(f"{out.stem} {dt.datetime.now():%Y%m%d_%H%M%S}{out.suffix}")
        writer = pd.ExcelWriter(out, engine="openpyxl")
    with writer:
        data.to_excel(writer, sheet_name="Data", index=False)
        for m, g in data.groupby(data["YearMonth"].dt.strftime("%Y-%m")):
            g.to_excel(writer, sheet_name=m, index=False)
        srcs.to_excel(writer, sheet_name="Sources", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for col in ws.iter_cols(min_row=1, max_row=1):
                name = str(col[0].value)
                width = 22 if "Location" in name or name == "District" else \
                    20 if "Date" in name or "Time" in name else 13
                ws.column_dimensions[col[0].column_letter].width = width
                fmt = {"Qty": "#,##0", "TotalZoneValue": "#,##0.00"}.get(name)
                if name in ("YearMonth",):
                    fmt = "yyyy-mm"
                elif name in ("OrderDate",):
                    fmt = "yyyy-mm-dd"
                elif "Time" in name or name == "ProcessedDate":
                    fmt = "yyyy-mm-dd hh:mm"
                elif name in ("Product", "Overhead", "Storage", "Transportation", "GrandTotal", "Price", "LIC"):
                    fmt = "0.00000"
                if fmt:
                    for c in col[1:]:
                        c.number_format = fmt
            # make sure the header row stays readable after column formats are set
            for c in ws[1]:
                c.number_format = "General"
    return out, srcs


def main():
    ap = argparse.ArgumentParser(description="Raw zone-detail data for as many months as possible.")
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--no-pull", action="store_true", help="only rebuild from archived snapshots")
    a = ap.parse_args()

    if not a.no_pull:
        pull_snowflake()
    print("Combining archived snapshots:")
    data, files = combine(a.year)
    out, srcs = write(data, files, a.year)
    print("\nRows by month:")
    print(srcs.to_string(index=False))
    print(f"\nTotal rows: {len(data):,}")
    print(f"Workbook: {out}")


if __name__ == "__main__":
    main()
