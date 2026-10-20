"""SYNTHETIC zone-detail extract + Finance LIC workbook. Nothing here is real."""
import numpy as np
import openpyxl
import pandas as pd

rng = np.random.default_rng(7)
px = pd.read_parquet("data/us_propane_priceindex.parquet")
idx = {}
for b, name in (("MTB", "NGL fixed price - Financial Mont Belvieu Propane LST"),
                ("CON", "NGL fixed price - Financial Conway Propane")):
    s = px[px.priceindex == name].copy()
    s = s[s.delivdate == (s.pricedate.dt.to_period("M") + 1).dt.to_timestamp()]
    idx[b] = s.set_index("pricedate")["price"].sort_index()

DIST = [("1001", "SPRINGFIELD", "MT", "CON"), ("1002", "CEDAR FALLS", "MT", "CON"),
        ("1003", "ELK HOLLOW", "MT", "CON"), ("2001", "LAKESIDE", "WY", "MTB"),
        ("2002", "FAIRVIEW", "WY", "MTB"), ("3001", "BROOKHAVEN", "UT", "MTB"),
        ("3002", "JUNIPER FLATS", "UT", "MTB"), ("4001", "ASPEN GROVE", "CO", "CON"),
        ("4002", "CLEARWATER", "CO", "CON"), ("5001", "HARBOR VIEW", "MI", "CON"),
        ("5002", "MAPLE GROVE", "MI", "CON"), ("6001", "CORN VALLEY", "IL", "CON"),
        ("7001", "LIMESTONE", "IN", "MTB"), ("8001", "CANYON VIEW", "AZ", "MTB"),
        ("8002", "MESA VERDE", "AZ", "MTB")]
TERM = {"MT": ["NORTHFIELD TERMINAL", "BILLINGS RACK"], "WY": ["CASPER LINE", "GRANITE RAIL"],
        "UT": ["WASATCH TERMINAL"], "CO": ["FRONT RANGE LINE", "GRAND JUNCTION RACK"],
        "MI": ["LAKESHORE TERMINAL", "BAY LINE"], "IL": ["PRAIRIE HUB"], "IN": ["WABASH LINE"],
        "AZ": ["SAGUARO TERMINAL"]}

rows, oid = [], 400000
days = pd.bdate_range("2026-09-01", "2026-09-30")
for did, name, st, basis in DIST:
    lic = round(rng.uniform(0.11, 0.24), 5)
    diff = rng.uniform(0.03, 0.07)
    lanes = [(t, rng.uniform(0.05, 0.16)) for t in TERM[st]]
    for _ in range(int(rng.integers(60, 140))):
        oid += 1
        d = pd.Timestamp(rng.choice(days))
        t, frt = lanes[int(rng.integers(len(lanes)))]
        mv = str(rng.choice(["Delivered", "Transport", "Transfer"], p=[.25, .6, .15]))
        org = f"{t} - DELIVERED" if mv == "Delivered" else t
        i = float(idx[basis].asof(d))
        prod = round(i + diff + rng.normal(0, .004), 5)
        ovh = round(rng.uniform(.006, .011), 5)
        sto = round(rng.uniform(0, .012), 5) if mv == "Transfer" else 0.0
        trn = 0.0 if mv == "Delivered" else round(frt + rng.normal(0, .006), 5)
        qty = float(round(rng.normal(9600, 450)))
        gt = round(prod + ovh + sto + trn, 5)
        typ = "Actual" if d.day <= 24 else str(rng.choice(["Actual", "Forcast"]))
        rows.append(dict(
            SnapshotID=318, SnapshotInstanceDateTime=pd.Timestamp("2026-10-01 06:00"),
            YearMonth=pd.Timestamp("2026-09-01"), Type=typ, DeptID=did, District=name, State=st,
            OrderDate=d, OrderID=oid, TMWOrderNumber=f"T{oid + 881000}", PONumber=f"PO{oid + 52000}",
            MoveType=mv, OriginLocation=org, DestinationLocation=f"{name} PLANT", Qty=qty,
            Product=prod, Overhead=ovh, Storage=sto, Transportation=trn, GrandTotal=gt,
            TotalZoneValue=round(qty * gt, 2), Basis=basis,
            Price=round(i, 5) if rng.random() > .08 else np.nan, LIC=lic))
df = pd.DataFrame(rows)
df["MoveType"] = df["MoveType"].astype(object)
df["District"] = df["District"].astype(object)


def pick(n, where=None):
    pool = df.index if where is None else df.index[where]
    return rng.choice(pool, n, replace=False)


# --- seeded problems, so every check has something to find -------------------
sp = df.District == "SPRINGFIELD"
for k in pick(4, sp & (df.MoveType == "Transport")):
    df.loc[k, "Transportation"] *= 2.6                     # freight spikes in the focus district
for k in pick(2, sp):
    df.loc[k, "Product"] += 0.09                           # off-index purchases
df["GrandTotal"] = df[["Product", "Overhead", "Storage", "Transportation"]].sum(axis=1).round(5)
df["TotalZoneValue"] = (df.Qty * df.GrandTotal).round(2)
for k in pick(12):
    df.loc[k, "Transportation"] *= 3.4                     # outliers elsewhere
for k in pick(9):
    df.loc[k, "TotalZoneValue"] *= 1.11                    # value != qty x total
for k in pick(5):
    df.loc[k, "GrandTotal"] += 0.02                        # total != sum of parts
for k in pick(6):
    df.loc[k, "Overhead"] = 0.0                            # zero fees
a, b = pick(2)
df.loc[b, "PONumber"] = df.loc[a, "PONumber"]              # PO shared by two orders
for k in pick(4):
    df.loc[k, "LIC"] = np.nan                              # missing LIC
k = pick(1, df.District == "FAIRVIEW")[0]
df.loc[k, "LIC"] += 0.015                                  # LIC varies within district-month
for k in pick(3):
    df.loc[k, "YearMonth"] = pd.Timestamp("2026-08-01")   # outside report month
for k in pick(3):
    df.loc[k, "MoveType"] = "TRANSPORT "                   # case variant + trailing space
for k in pick(2):
    df.loc[k, "District"] = "N/A"                          # placeholder
for k in pick(2):
    df.loc[k, "OrderDate"] = pd.Timestamp("2027-09-14")   # future date
k = pick(1)[0]
df.loc[k, "Qty"] = -df.loc[k, "Qty"]                       # stray negative
df = pd.concat([df, df.loc[pick(3)]], ignore_index=True)  # fully duplicated rows
df.to_parquet("zone_detail_raw.parquet", index=False)
print(len(df), "rows", df.District.nunique(), "districts")

# --- Finance LIC workbook: planned landed cost above basis, $/gal ------------
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "LIC FY27"
ws.append(["Finance - planned landed cost above basis ($/gal), FY 2026-2027 (SYNTHETIC)"])
ws.append([])
M = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
ws.append(["DistrictID", "District", "Basis", *M])
for did, name, st, basis in DIST:
    g = df[df.DeptID == did]
    above = (g.Product - g.Price.fillna(g.Product - .05) + g.Overhead + g.Transportation).median()
    planned = float(above) + rng.normal(0, .006)
    ws.append([did, name.title(), basis, *[round(planned + rng.normal(0, .003), 5) for _ in M]])
wb.save("finance_costs 2026-2027.xlsx")
print("finance workbook written")
