"""
legacy.py - a one-time head start: turn last year's Excel model workbooks into
a data set for this app.

This is the ONLY place the old workbooks are touched, and only to read them
(through the old reader in ../model.py, which opens them read-only). After the
conversion the app never looks at them again - every number lives in the data
set and is changed by uploads or by editing in the app.

What comes across, per area workbook:
  supply points   location, supplier, hub, contract, differential
                  (= delivered price - hub price), min / max lift %, storage
  plants          name, state (from the name, e.g. "... KS"), demand, storage
  lanes           every supply point -> plant pair with a freight rate in at
                  least one month. A blank rate stays blank (lane not used that
                  month); max shipments of 999,999+ become "no limit".
  hub prices      per area, from the hub codes on the supply points

Units are converted from the workbooks' thousands of gallons and $ per
thousand gallons to gallons and $ per gallon.

The workbook's demand comes across as the plants' budget demand with no
contract % applied, so the converted plan reproduces last year's model. Upload
the "Demand for Models" workbook afterwards to bring in the unadjusted budget
and the state contract %.
"""

from __future__ import annotations

import math
import os
import re
import sys
from typing import Callable, Dict, List, Optional, Tuple

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
OLD_DIR = os.path.dirname(HERE)            # parent folder, where model.py lives
NO_LIMIT = 999_999 * 1000.0                # gallons
_PLACEHOLDER = re.compile(r"^(supplier|supply|demand|plant)\s*\d+$", re.I)
_STATE = re.compile(r"(?:^|[\s,(\-])([A-Z]{2})\)?\s*$")


def _old():
    if OLD_DIR not in sys.path:
        sys.path.insert(0, OLD_DIR)
    import model  # noqa: E402  (the original command-line model)
    return model


def available(variant: str = "Working", root: Optional[str] = None) -> List[dict]:
    """The workbooks a conversion would read: [{region, area, path}]."""
    m = _old()
    root = root or m.default_root()
    out = []
    for region, files in m.discover(root, variant).items():
        for f in files:
            out.append({"region": region, "area": f.area, "path": f.path})
    return out


def _clean(v: float, nd: int = 6) -> Optional[float]:
    if v is None or not math.isfinite(v):
        return None
    return round(float(v), nd)


def _series(arr, scale: float = 1.0, nd: int = 6) -> List[Optional[float]]:
    return [_clean(float(x) * scale, nd) for x in np.asarray(arr, dtype=float)]


def area_id(code: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", code.lower()).strip("-")[:40] or "area"


def state_of(name: str, fallback: str = "") -> str:
    m = _STATE.search(name or "")
    if m:
        return m.group(1)
    return fallback


def convert_one(f: dict, names_mod=None) -> Tuple[dict, List[str]]:
    """One workbook -> (area dict, notes)."""
    m = _old()
    d = m.ModelBaseReader(f["path"], f["area"], f["region"]).read()
    notes: List[str] = []
    code = re.sub(r"\s*\d{2}-\d{2}\s*$", "", f["area"]).strip()
    disp = names_mod.area_display(f["area"]) if names_mod else {"title": code, "states": ""}
    area_states = [s.strip() for s in re.split(r"[\u00b7,/ ]+", disp.get("states") or "") if len(s.strip()) == 2]
    P, N, M = d.n_periods, d.n_supply, d.n_demand
    G = 1000.0
    contract = d.contract * G                        # (P, N) gal
    price = d.price / G                              # $/gal
    hub_price = d.hub_price / G
    demand = d.demand * G                            # (P, M)
    freight = d.freight_as_entered() / G             # (P, N, M), NaN = blank

    def blankish(x):
        return not np.isfinite(x).any() or np.nanmax(np.abs(x)) < 1e-12

    keep_s = []
    for i in range(N):
        nm = (d.supply_names[i] or "").strip()
        has_lane = np.isfinite(freight[:, i, :]).any()
        if (not nm or _PLACEHOLDER.match(nm)) and blankish(contract[:, i]):
            continue
        if not nm and not has_lane:
            continue
        keep_s.append(i)
    keep_p = []
    for j in range(M):
        nm = (d.demand_names[j] or "").strip()
        if (not nm or _PLACEHOLDER.match(nm)) and blankish(demand[:, j]):
            continue
        if not nm:
            continue
        keep_p.append(j)
    dropped = (N - len(keep_s), M - len(keep_p))
    if any(dropped):
        notes.append(f"Skipped {dropped[0]} empty supply row(s) and {dropped[1]} empty plant row(s).")

    # hubs: one per code, priced from the first supply point using it
    hubs: Dict[str, np.ndarray] = {}
    for i in keep_s:
        code_h = (d.supply_hubs[i] or "").strip()
        if code_h and code_h not in hubs and not blankish(hub_price[:, i]):
            hubs[code_h] = hub_price[:, i]
    supply = []
    sid = {}
    for n, i in enumerate(keep_s, start=1):
        code_h = (d.supply_hubs[i] or "").strip()
        base = hubs.get(code_h)
        if base is None:
            base = np.zeros(P)
            code_h = code_h if code_h in hubs else ""
        sid[i] = f"s{n}"
        supply.append({
            "id": f"s{n}", "location": (d.supply_names[i] or f"Supply {n}").strip(),
            "supplier": (d.supply_groups[i] or "").strip(), "hub": code_h,
            "contract": _series(contract[:, i], nd=3),
            "diff": _series(price[:, i] - base, nd=6),
            "min_pct": _series(d.perc_min[:, i], nd=6),
            "max_pct": _series(d.perc_max[:, i], nd=6),
            "open": _clean(float(d.supply_inv_init[i]) * G, 3) or 0.0,
            "cap": _series(d.supply_inv_cap[:, i], G, 3),
            "inv_min": _series(d.supply_inv_min[:, i], G, 3),
        })
    plants = []
    pid = {}
    for n, j in enumerate(keep_p, start=1):
        nm = (d.demand_names[j] or "").strip()
        st = state_of(nm, area_states[0] if len(area_states) == 1 else "")
        pid[j] = f"p{n}"
        plants.append({
            "id": f"p{n}", "name": nm, "state": st,
            "budget": _series(demand[:, j], nd=3),
            "open": _clean(float(d.demand_inv_init[j]) * G, 3) or 0.0,
            "cap": _series(d.demand_inv_cap[:, j], G, 3),
            "inv_min": _series(d.demand_inv_min[:, j], G, 3),
        })
    lanes = []
    for i in keep_s:
        for j in keep_p:
            r = freight[:, i, j]
            if not np.isfinite(r).any():
                continue
            mn = d.ship_min[:, i, j] * G
            mx = d.ship_max[:, i, j] * G
            mx_l = [None if (not math.isfinite(v) or v >= NO_LIMIT - 1) else round(float(v), 3) for v in mx]
            mn_l = [round(float(v), 3) if math.isfinite(v) else 0.0 for v in mn]
            lanes.append({"from": sid[i], "to": pid[j], "rate": _series(r, nd=6),
                          "min": None if not any(mn_l) else mn_l,
                          "max": None if all(v is None for v in mx_l) else mx_l})
    area = {
        "id": area_id(code), "name": code, "title": disp.get("title") or code, "region": f["region"],
        "months": list(d.period_labels),
        "hubs": [{"code": c, "name": "", "price": _series(v, nd=6)} for c, v in hubs.items()],
        "supply": supply, "plants": plants, "lanes": lanes,
        "source": os.path.basename(f["path"]),
    }
    notes += list(d.repairs or [])
    return area, notes


def convert(variant: str = "Working", root: Optional[str] = None,
            progress: Optional[Callable[[int, int, str], None]] = None, names_mod=None) -> Tuple[Dict[str, dict], List[dict]]:
    """Every workbook of a variant -> ({area id: area}, report rows)."""
    files = available(variant, root)
    areas: Dict[str, dict] = {}
    report = []
    for k, f in enumerate(files):
        if progress:
            progress(k, len(files), f["area"])
        try:
            a, notes = convert_one(f, names_mod)
            if a["id"] in areas:
                a["id"] = f"{a['id']}-{k}"
            areas[a["id"]] = a
            report.append({"area": a["name"], "region": f["region"], "ok": True, "file": os.path.basename(f["path"]),
                           "supply": len(a["supply"]), "plants": len(a["plants"]), "lanes": len(a["lanes"]), "notes": notes})
        except Exception as e:  # noqa: BLE001 - one bad workbook shouldn't stop the rest
            report.append({"area": f["area"], "region": f["region"], "ok": False, "file": os.path.basename(f["path"]),
                           "error": str(e)})
    return areas, report
