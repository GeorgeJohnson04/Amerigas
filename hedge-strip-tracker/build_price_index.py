"""
Build a US-propane-only forward-curve price index parquet with full history.

Reads daily Argus forward-curve exports (ArgusFC_YYYYMMDD.csv) from LPG/ and
writes one long-format table:

    commodity, priceindex, pricedate, delivdate, basis, index, price

US propane indices (commodity = 'LPG'):

    NGL fixed price - Financial Conway Propane                    Conway, Kansas
    NGL fixed price - Financial Mont Belvieu Propane Enterprise   Mont Belvieu, Texas
    NGL fixed price - Financial Mont Belvieu Propane LST          Mont Belvieu, Texas

Plus a synthetic fourth index: a weighted blend of the three hubs
(weights are configurable with --blend-weights):

    US Propane Blend

Excluded: CIF ARA Propane EUR (European, EUR-denominated) and every non-propane
NGL (butane, isobutane, ethane, natural gasoline).

Parsing rules:
  - ValuationTypeName == 'Mid Market' only, which prevents the row triplication
  - ForwardPeriod 1-12 only, i.e. monthly contracts, not annual strips
  - delivdate built from Year + ForwardPeriod

Usage:
    python build_price_index.py
    python build_price_index.py --years 5
    python build_price_index.py --start 2021-09-01
    python build_price_index.py --csv
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import datetime

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
LPG_DIR = os.path.join(HERE, "LPG")
DEFAULT_OUT = os.path.join(HERE, "us_propane_priceindex.parquet")

SCHEMA = ["commodity", "priceindex", "pricedate", "delivdate", "basis", "index", "price"]

CONWAY = "NGL fixed price - Financial Conway Propane"
MB_ENTERPRISE = "NGL fixed price - Financial Mont Belvieu Propane Enterprise"
MB_LST = "NGL fixed price - Financial Mont Belvieu Propane LST"

US_PROPANE_INDICES = [CONWAY, MB_ENTERPRISE, MB_LST]

# --- Hub blend ------------------------------------------------------------
# A weighted composite of the three hub curves. The weights below are
# illustrative defaults; pass --blend-weights to use your own. The Mont Belvieu
# non-Enterprise (Non-TET) leg maps to the Argus "Mont Belvieu Propane LST"
# assessment.
BLEND_NAME = "US Propane Blend"
BLEND_WEIGHTS = {
    CONWAY: 0.30,
    MB_ENTERPRISE: 0.35,
    MB_LST: 0.35,
}

FNAME_RE = re.compile(r"ArgusFC_(\d{8})\.csv$", re.IGNORECASE)


def file_date(filename: str):
    m = FNAME_RE.search(filename)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d").date()
    except ValueError:
        return None


def parse_propane(filepath: str) -> pd.DataFrame:
    """Parse one ArgusFC file, keeping only the US propane monthly curve points."""
    df = pd.read_csv(filepath)

    required = {"ValuationTypeName", "ProductName", "MarketName",
                "ForwardPeriod", "Year", "TradeDate", "Value"}
    if not required.issubset(df.columns):
        return pd.DataFrame(columns=SCHEMA)

    # Mid Market only — the source carries Bid/Ask/Mid, which would triple rows
    df = df[df["ValuationTypeName"] == "Mid Market"]
    if df.empty:
        return pd.DataFrame(columns=SCHEMA)

    df = df[df["ProductName"].str.contains(
        "NGL fixed price - Financial", na=False, regex=False)]
    if df.empty:
        return pd.DataFrame(columns=SCHEMA)

    df = df.copy()
    df["priceindex"] = (
        df["ProductName"].str.strip() + " " + df["MarketName"].str.strip()
    )

    # Explicit allow-list: propane only, US hubs only
    df = df[df["priceindex"].isin(US_PROPANE_INDICES)]
    if df.empty:
        return pd.DataFrame(columns=SCHEMA)

    df["pricedate"] = pd.to_datetime(df["TradeDate"], format="mixed")

    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["ForwardPeriod"] = pd.to_numeric(df["ForwardPeriod"], errors="coerce")
    df = df.dropna(subset=["Year", "ForwardPeriod"])
    # Monthly contracts only; ForwardPeriod > 12 denotes annual/strip contracts
    df = df[(df["ForwardPeriod"] >= 1) & (df["ForwardPeriod"] <= 12)]
    if df.empty:
        return pd.DataFrame(columns=SCHEMA)

    df["delivdate"] = pd.to_datetime(
        dict(
            year=df["Year"].astype(int),
            month=df["ForwardPeriod"].astype(int),
            day=1,
        )
    )

    df["price"] = pd.to_numeric(df["Value"], errors="coerce")
    df = df.dropna(subset=["price"])
    if df.empty:
        return pd.DataFrame(columns=SCHEMA)

    df["commodity"] = "LPG"
    df["basis"] = 0.0
    df["index"] = df["price"]

    return df[SCHEMA]


def build_blend(
    components: pd.DataFrame,
    name: str = BLEND_NAME,
    weights: dict[str, float] | None = None,
    verbose: bool = True,
) -> pd.DataFrame:
    """Weighted composite of the component curves, as a synthetic priceindex.

    Weights are renormalised over whichever components have a quote on a given
    (pricedate, delivdate), so a partial day still produces a blend point rather
    than a gap.
    """
    import numpy as np

    weights = weights or BLEND_WEIGHTS

    wide = components.pivot_table(
        index=["pricedate", "delivdate"],
        columns="priceindex",
        values="price",
        aggfunc="last",
    )

    cols = [c for c in weights if c in wide.columns]
    missing = [c for c in weights if c not in wide.columns]
    if verbose and missing:
        print("  WARNING: blend component absent from source:")
        for m in missing:
            print(f"    {m!r} (weight {weights[m]})")
    if not cols:
        if verbose:
            print("  ERROR: no blend components available; skipping blend")
        return pd.DataFrame(columns=SCHEMA)

    vals = wide[cols].to_numpy(dtype=float)
    w = np.array([weights[c] for c in cols], dtype=float)

    have = ~np.isnan(vals)
    numerator = np.nansum(vals * w, axis=1)
    denominator = (have * w).sum(axis=1)

    with np.errstate(invalid="ignore", divide="ignore"):
        price = np.where(denominator > 0, numerator / denominator, np.nan)

    n_components = have.sum(axis=1)

    out = wide.reset_index()[["pricedate", "delivdate"]].copy()
    out["price"] = price
    out["_n_components"] = n_components
    out = out.dropna(subset=["price"])

    if verbose:
        full = int((out["_n_components"] == len(weights)).sum())
        partial = len(out) - full
        print(f"  blend points: {len(out):,}  "
              f"(all {len(weights)} components: {full:,}, partial: {partial:,})")
        if partial:
            breakdown = out.loc[out["_n_components"] < len(weights), "_n_components"]
            print(f"    partial rows use {sorted(breakdown.unique())} component(s); "
                  "weights renormalised")

        # Sanity: a weighted mean must sit within the component range
        comp_min = np.nanmin(vals, axis=1)
        comp_max = np.nanmax(vals, axis=1)
        inside = (price >= comp_min - 1e-9) & (price <= comp_max + 1e-9)
        n_ok = int(np.nansum(inside))
        n_tot = int(np.sum(~np.isnan(price)))
        print(f"    within component min/max: {n_ok:,}/{n_tot:,}"
              + ("" if n_ok == n_tot else "  <-- INVESTIGATE"))

    out = out.drop(columns="_n_components")
    out["priceindex"] = name
    out["commodity"] = "LPG"
    out["basis"] = 0.0
    out["index"] = out["price"]

    return out[SCHEMA]


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build a US propane price index parquet from the LPG/ Argus history."
    )
    parser.add_argument("--lpg-dir", default=LPG_DIR, help="Folder of ArgusFC_*.csv files")
    parser.add_argument("--out", default=DEFAULT_OUT, help="Output parquet path")
    parser.add_argument(
        "--years", type=float, default=5.0,
        help="Years of history to keep, counted back from the latest file (default: 5)",
    )
    parser.add_argument(
        "--start", default=None,
        help="Explicit start date YYYY-MM-DD; overrides --years",
    )
    parser.add_argument("--csv", action="store_true", help="Also write a .csv")
    parser.add_argument(
        "--no-blend", action="store_true",
        help="Omit the synthetic blend index; components only",
    )
    parser.add_argument(
        "--blend-name", default=BLEND_NAME,
        help=f"priceindex name for the blend (default: {BLEND_NAME!r})",
    )
    parser.add_argument(
        "--blend-weights", default=None,
        help="Override weights as conway,enterprise,lst e.g. '0.25,0.375,0.375'",
    )
    parser.add_argument(
        "--from-parquet", nargs="?", const=DEFAULT_OUT, default=None,
        help="Skip re-parsing LPG/ and read the component series from an existing "
             "parquet instead (defaults to the output path). Much faster when you "
             "only need to rebuild or retune the blend.",
    )
    args = parser.parse_args()

    weights = dict(BLEND_WEIGHTS)
    if args.blend_weights:
        parts = [p.strip() for p in args.blend_weights.split(",")]
        if len(parts) != 3:
            print("ERROR: --blend-weights needs 3 values: conway,enterprise,lst")
            return 2
        try:
            vals = [float(p) for p in parts]
        except ValueError:
            print(f"ERROR: --blend-weights must be numeric, got {args.blend_weights!r}")
            return 2
        if sum(vals) <= 0:
            print("ERROR: --blend-weights must sum to a positive number")
            return 2
        weights = dict(zip([CONWAY, MB_ENTERPRISE, MB_LST], vals))

    # ---- fast path: reuse already-parsed components ----------------------
    if args.from_parquet:
        if not os.path.exists(args.from_parquet):
            print(f"ERROR: --from-parquet file not found: {args.from_parquet}")
            return 2

        print("=" * 72)
        print("Build US propane price index (blend only, reusing components)")
        print("=" * 72)
        print(f"components from: {args.from_parquet}")

        src = pd.read_parquet(args.from_parquet)
        out = src[src["priceindex"].isin(US_PROPANE_INDICES)].copy()
        print(f"component rows: {len(out):,} "
              f"across {out['priceindex'].nunique()} indices")
        dropped = len(src) - len(out)
        if dropped:
            other = sorted(set(src["priceindex"]) - set(US_PROPANE_INDICES))
            print(f"ignored {dropped:,} non-component row(s): {other}")

        if out.empty:
            print("ERROR: no component rows found; nothing to blend.")
            return 1

        if not args.no_blend:
            print()
            print(f"Hub blend -> {args.blend_name!r}")
            for idx, w in weights.items():
                print(f"  {w:>6.3f}  {idx}")
            blend = build_blend(out, name=args.blend_name, weights=weights)
            if blend.empty:
                print("ERROR: blend produced no rows.")
                return 1
            out = pd.concat([out, blend], ignore_index=True)

        out = out.sort_values(
            ["priceindex", "pricedate", "delivdate"]
        ).reset_index(drop=True)

        print()
        print("per index:")
        for name, grp in out.groupby("priceindex"):
            print(f"  {name}")
            print(f"    rows        {len(grp):,}")
            print(f"    pricedates  {grp['pricedate'].nunique():,} "
                  f"({grp['pricedate'].min().date()} .. {grp['pricedate'].max().date()})")
            print(f"    price       {grp['price'].min():.4f} .. "
                  f"{grp['price'].max():.4f}  (nulls {int(grp['price'].isna().sum())})")
        print()

        out.to_parquet(args.out, index=False)
        print(f"Wrote {len(out):,} rows to {args.out} "
              f"({os.path.getsize(args.out) / 1e6:.2f} MB)")
        if args.csv:
            csv_path = os.path.splitext(args.out)[0] + ".csv"
            out.to_csv(csv_path, index=False)
            print(f"Wrote {len(out):,} rows to {csv_path}")

        check = pd.read_parquet(args.out)
        print()
        print(f"verify: rows {len(check):,}, columns {list(check.columns)}")
        print(f"verify: indices {sorted(check['priceindex'].unique())}")
        return 0

    if not os.path.isdir(args.lpg_dir):
        print(f"ERROR: LPG folder not found: {args.lpg_dir}")
        return 2

    dated = []
    for name in os.listdir(args.lpg_dir):
        d = file_date(name)
        if d:
            dated.append((d, os.path.join(args.lpg_dir, name)))
    dated.sort()

    if not dated:
        print(f"ERROR: no ArgusFC_YYYYMMDD.csv files in {args.lpg_dir}")
        return 2

    print("=" * 72)
    print("Build US propane price index")
    print("=" * 72)
    print(f"source folder: {args.lpg_dir}")
    print(f"files available: {len(dated):,}  ({dated[0][0]} .. {dated[-1][0]})")

    latest = dated[-1][0]
    if args.start:
        try:
            cutoff = datetime.strptime(args.start, "%Y-%m-%d").date()
        except ValueError:
            print(f"ERROR: --start must be YYYY-MM-DD, got '{args.start}'")
            return 2
    else:
        cutoff = latest - pd.Timedelta(days=round(args.years * 365.25))
        cutoff = cutoff.date() if hasattr(cutoff, "date") else cutoff

    selected = [(d, p) for d, p in dated if d >= cutoff]
    print(f"cutoff: {cutoff}  ->  {len(selected):,} file(s) to parse")
    print()

    frames = []
    failures = []
    for i, (d, path) in enumerate(selected, 1):
        try:
            part = parse_propane(path)
            if not part.empty:
                frames.append(part)
        except Exception as exc:
            failures.append((os.path.basename(path), f"{type(exc).__name__}: {exc}"))
        if i % 250 == 0 or i == len(selected):
            print(f"  parsed {i:,}/{len(selected):,} files, "
                  f"{sum(len(f) for f in frames):,} rows so far")

    if failures:
        print()
        print(f"WARNING: {len(failures)} file(s) failed to parse:")
        for name, err in failures[:10]:
            print(f"  {name}: {err}")

    if not frames:
        print("\nERROR: no US propane rows found; nothing written.")
        return 1

    out = pd.concat(frames, ignore_index=True)
    print()
    print(f"parsed rows: {len(out):,}")

    pre = len(out)
    out = out.drop_duplicates(
        subset=["commodity", "priceindex", "pricedate", "delivdate"], keep="last"
    )
    if pre != len(out):
        print(f"deduplicated: removed {pre - len(out):,} duplicate key(s)")

    if not args.no_blend:
        print()
        print(f"Hub blend -> {args.blend_name!r}")
        for idx, w in weights.items():
            print(f"  {w:>6.3f}  {idx}")
        blend = build_blend(out, name=args.blend_name, weights=weights)
        if not blend.empty:
            out = pd.concat([out, blend], ignore_index=True)

    out = out.sort_values(["priceindex", "pricedate", "delivdate"]).reset_index(drop=True)

    print()
    print("per index:")
    for name, grp in out.groupby("priceindex"):
        print(f"  {name}")
        print(f"    rows        {len(grp):,}")
        print(f"    pricedates  {grp['pricedate'].nunique():,} "
              f"({grp['pricedate'].min().date()} .. {grp['pricedate'].max().date()})")
        print(f"    delivdates  {grp['delivdate'].min().date()} .. "
              f"{grp['delivdate'].max().date()}")
        print(f"    price       {grp['price'].min():.4f} .. {grp['price'].max():.4f}  "
              f"(nulls {int(grp['price'].isna().sum())})")
    print()

    out.to_parquet(args.out, index=False)
    print(f"Wrote {len(out):,} rows to {args.out} "
          f"({os.path.getsize(args.out) / 1e6:.2f} MB)")

    if args.csv:
        csv_path = os.path.splitext(args.out)[0] + ".csv"
        out.to_csv(csv_path, index=False)
        print(f"Wrote {len(out):,} rows to {csv_path}")

    # Verify the round trip and the schema
    check = pd.read_parquet(args.out)
    print()
    print(f"verify: rows {len(check):,}, columns {list(check.columns)}")
    print(f"verify: indices {sorted(check['priceindex'].unique())}")
    if list(check.columns) != SCHEMA:
        print(f"  schema mismatch: expected {SCHEMA}")
        return 1

    span_days = (check["pricedate"].max() - check["pricedate"].min()).days
    print(f"verify: history span {span_days:,} days (~{span_days / 365.25:.1f} years)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
