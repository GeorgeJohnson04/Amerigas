"""
Build a SYNTHETIC locked base model for the web app, so it can run with no Excel
workbooks. Every name, price, volume and freight rate here is invented.
"""
import os
import sys

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "webapp"))

import model  # noqa: E402
import bases as B  # noqa: E402

PERIODS = ["Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26",
           "Oct-26", "Nov-26", "Dec-26", "Jan-27", "Feb-27", "Mar-27"]
SHAPE = np.array([5.0, 4.0, 3.4, 3.2, 3.6, 4.6, 7.6, 10.8, 13.6, 15.0, 12.8, 10.0])
SHAPE = SHAPE / SHAPE.sum()
# Forward hub curves, $/kgal (0.70-0.82 $/gal), winter premium
MTB = np.array([712, 708, 705, 709, 716, 726, 744, 768, 792, 806, 796, 760], float)
CON = MTB - np.array([22, 22, 24, 24, 24, 22, 18, 12, 8, 6, 10, 16], float)

AREAS = {
    # region, area, seed, supply points, plants, shortfall?
    ("Rockies", "MT-WY 26-27"): dict(seed=11, states=["MT", "WY"], shortfall=False,
        supply=[("Acme Energy", "Northfield Terminal, MT", "CON", 34, 0.90, 1.10, 1600),
                ("Summit NGL", "Northfield Terminal, MT", "CON", 41, 0.85, 1.15, 0),
                ("Prairie Midstream", "Granite Rail, WY", "MTB", 58, 0.90, 1.10, 900),
                ("Summit NGL", "Casper Line, WY", "MTB", 47, 0.80, 1.20, 0),
                ("Spot Market", "Billings Rack, MT", "CON", 96, 0.00, 2.50, 0)],
        plants=["1001-Springfield MT", "1002-Cedar Falls MT", "1003-Elk Hollow MT",
                "1004-Pine Ridge MT", "1005-Silver Creek MT", "2001-Lakeside WY",
                "2002-Fairview WY", "2003-Millbrook WY", "2004-Granite Pass WY",
                "2005-Red Butte WY"]),
    ("Rockies", "UT-CO 26-27"): dict(seed=23, states=["UT", "CO"], shortfall=False,
        supply=[("Acme Energy", "Wasatch Terminal, UT", "MTB", 52, 0.90, 1.10, 1200),
                ("Prairie Midstream", "Front Range Line, CO", "CON", 39, 0.85, 1.15, 0),
                ("Summit NGL", "Front Range Line, CO", "CON", 44, 0.85, 1.15, 0),
                ("Spot Market", "Grand Junction Rack, CO", "MTB", 104, 0.00, 2.50, 0)],
        plants=["3001-Brookhaven UT", "3002-Juniper Flats UT", "3003-Stony Point UT",
                "3004-Willow Bend UT", "4001-Aspen Grove CO", "4002-Clearwater CO",
                "4003-High Mesa CO", "4004-Kettle Ridge CO", "4005-Sunfield CO"]),
    ("Midwest", "MI 26-27"): dict(seed=37, states=["MI"], shortfall=False,
        supply=[("Acme Energy", "Lakeshore Terminal, MI", "CON", 29, 0.90, 1.10, 2000),
                ("Great Lakes Gas", "Bay Line, MI", "CON", 33, 0.85, 1.15, 0),
                ("Spot Market", "Thumb Rack, MI", "CON", 88, 0.00, 2.50, 0)],
        plants=["5001-Harbor View MI", "5002-Maple Grove MI", "5003-Copper Falls MI",
                "5004-North Bay MI", "5005-Birch Run MI", "5006-Iron Hill MI",
                "5007-Twin Lakes MI", "5008-Westwood MI"]),
    ("Midwest", "IL-IN 26-27"): dict(seed=41, states=["IL", "IN"], shortfall=True,
        supply=[("Great Lakes Gas", "Prairie Hub, IL", "CON", 27, 0.90, 1.10, 1500),
                ("Summit NGL", "Wabash Line, IN", "MTB", 36, 0.90, 1.10, 0),
                ("Acme Energy", "Wabash Line, IN", "MTB", 40, 0.85, 1.10, 0)],
        plants=["6001-Corn Valley IL", "6002-Oak Prairie IL", "6003-Riverbend IL",
                "6004-Hickory IL", "7001-Limestone IN", "7002-Cardinal IN",
                "7003-Walnut Creek IN"]),
    ("West", "AZ 26-27"): dict(seed=53, states=["AZ"], shortfall=False,
        supply=[("Desert Fuels", "Saguaro Terminal, AZ", "MTB", 88, 0.90, 1.10, 800),
                ("Spot Market", "Painted Rock Rack, AZ", "MTB", 131, 0.00, 2.50, 0)],
        plants=["8001-Canyon View AZ", "8002-Mesa Verde AZ", "8003-Sun Valley AZ",
                "8004-Copper Basin AZ", "8005-Red Rock AZ"]),
}


def build_area(spec):
    rng = np.random.default_rng(spec["seed"])
    P, N, M = 12, len(spec["supply"]), len(spec["plants"])
    annual = rng.uniform(900, 3200, M).round(0)          # kgal / plant / year
    demand = np.outer(SHAPE, annual) * rng.uniform(0.96, 1.04, (P, M))
    total = demand.sum(axis=1)

    # supply locations and plants on a 300 x 300 mile map
    s_xy = rng.uniform(40, 260, (N, 2))
    for i in range(1, N):                                 # same location -> same point
        if spec["supply"][i][1] == spec["supply"][i - 1][1]:
            s_xy[i] = s_xy[i - 1]
    d_xy = rng.uniform(0, 300, (M, 2))
    dist = np.linalg.norm(s_xy[:, None, :] - d_xy[None, :, :], axis=2)

    freight = 38 + 0.95 * dist + rng.normal(0, 6, (N, M))
    freight[dist > 210] = np.nan                          # too far: blank freight = lane not used
    for j in range(M):                                    # every plant keeps two lanes
        order = np.argsort(dist[:, j])[:2]
        freight[order, j] = 38 + 0.95 * dist[order, j]
    if spec["seed"] == 11:                                # one remote mountain plant: $1+/gal lanes
        j = 2
        freight[:, j] = np.where(np.isfinite(freight[:, j]), freight[:, j] + 965, np.nan)
        freight[0, j] = 1040.0
    trans = np.repeat(freight[None, :, :], P, axis=0)
    trans *= (1 + 0.04 * (SHAPE / SHAPE.max()))[:, None, None]  # small winter freight bump

    hubs = [s[2] for s in spec["supply"]]
    hub_price = np.stack([MTB if h == "MTB" else CON for h in hubs], axis=1)
    hub_price = hub_price + rng.normal(0, 1.5, (P, N))
    diff = np.array([s[3] for s in spec["supply"]], float)
    hub_premium = np.repeat(diff[None, :], P, axis=0)

    term = [i for i, s in enumerate(spec["supply"]) if s[0] != "Spot Market"]
    share = rng.dirichlet(np.ones(len(term)) * 4)
    contract = np.zeros((P, N))
    cover = 0.875 if spec["shortfall"] else 0.95
    for k, i in enumerate(term):
        contract[:, i] = total * cover * share[k]
    for i, s in enumerate(spec["supply"]):
        if s[0] == "Spot Market":
            contract[:, i] = total * 0.08
    if spec["shortfall"]:                                 # an over-committed peak winter
        contract[8:11] *= 0.97
    perc_min = np.repeat(np.array([s[4] for s in spec["supply"]])[None, :], P, axis=0)
    perc_max = np.repeat(np.array([s[5] for s in spec["supply"]])[None, :], P, axis=0)

    s_cap = np.repeat(np.array([s[6] for s in spec["supply"]], float)[None, :], P, axis=0)
    s_init = s_cap[0] * 0.30
    s_min = np.zeros((P, N))

    d_cap_1 = (annual * rng.uniform(0.05, 0.09, M)).round(0)
    d_cap = np.repeat(d_cap_1[None, :], P, axis=0)
    d_init = (d_cap_1 * 0.45).round(1)
    d_min = np.repeat((d_cap_1 * 0.15).round(1)[None, :], P, axis=0)

    arrays = {
        "contract": contract.round(1), "hub_price": hub_price.round(2),
        "hub_premium": hub_premium, "perc_min": perc_min, "perc_max": perc_max,
        "demand": demand.round(1), "trans_cost": np.round(trans, 2),
        "ship_min": np.zeros((P, N, M)), "ship_max": np.full((P, N, M), model.DEFAULT_MAX_SHIP),
        "supply_inv_init": s_init, "supply_inv_cap": s_cap, "supply_inv_min": s_min,
        "demand_inv_init": d_init, "demand_inv_cap": d_cap, "demand_inv_min": d_min,
    }
    return {
        "arrays": arrays, "cached_price": np.full((P, N), np.nan),
        "supply_names": [s[1] for s in spec["supply"]],
        "supply_groups": [s[0] for s in spec["supply"]],
        "supply_hubs": hubs, "demand_names": list(spec["plants"]),
        "period_labels": list(PERIODS),
    }, s_xy, d_xy


def all_raw():
    return {key: build_area(spec) for key, spec in AREAS.items()}


def save_bases(folder):
    bases = B.Bases(folder)
    if bases.list():
        return bases
    raws = all_raw()
    for name, official, bump in (("2026-27 Base", True, 0.0), ("2026-27 Working files", False, 0.0)):
        meta = bases._new_meta(name, "26-27", "setup", "import",
                               {"from": "synthetic sample data", "variant": "Baseline" if official else "Working"})
        meta["official"] = official
        pk = []
        for (region, area), (raw, _, _) in raws.items():
            states, pct = B.states_for(raw, None)
            pkg = B._pack(raw, region=region, area=area, source="synthetic",
                          demand_states=states, state_pct=pct)
            pk.append((region, area, pkg, {"source": "synthetic sample data"}))
        bases._save(meta, pk)
    return bases


if __name__ == "__main__":
    for (region, area), (raw, _, _) in all_raw().items():
        d = model.ModelBaseReader("synthetic", area, region).build_from_inputs(raw)
        sol = model.solve_area(d)
        print(f"{region:8s} {area:12s} {sol.status:9s} cost ${sol.objective if hasattr(sol,'objective') else 0:,.0f}"
              f"  unmet {sol.unmet_total:,.1f} kgal  demand {d.total_demand:,.0f} kgal")
