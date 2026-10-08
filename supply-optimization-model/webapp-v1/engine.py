"""
engine.py - the supply optimizer.

Works straight on the data people upload, in gallons and $ per gallon:

  supply points   where product is bought: a contract volume per month, a
                  price (hub price + differential), how much of the contract
                  may be lifted (min / max %), optional storage
  plants          where product is needed: planned demand per month,
                  optional storage
  lanes           the routes ("OD pairs") from a supply point to a plant, each
                  with a freight rate per month and optional min / max gallons.
                  A route that isn't listed - or has no rate in a month - is
                  not used. Every rate is a real price, however high; rates of
                  $1.00/gal or more are only flagged ("high freight").

For every area it finds the cheapest monthly plan: how much to buy from each
contract, what to ship on each lane, and what to put into or take out of
storage, so that every plant's demand is met.

The linear program, per month p:
  bought(i,p) + from storage(i,p) - into storage(i,p) = shipped out of i
  shipped into j + from storage(j,p) - into storage(j,p) = demand(j,p)
  storage(p) = storage(p-1) - from storage + into storage, kept between the
  minimum and the capacity
  min% x contract <= bought <= max% x contract,  lane min <= shipped <= lane max
minimising  sum(bought x price) + sum(shipped x freight rate).

When the data can't be satisfied (contracts too small, minimums too big) it
solves again allowing unmet demand / unplaceable volume at a heavy penalty, and
says exactly where and how much, so the plan is still useful.

Internally it solves in thousands of gallons for numerical stability; every
number going in and out is in gallons and $/gal.
"""

from __future__ import annotations

import hashlib
import json
import math
import time
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import numpy as np

K = 1000.0                 # solve in thousands of gallons
PENALTY = 25.0             # $/gal for unmet demand / unplaceable volume when relaxing
HIGH_FREIGHT = 1.0         # $/gal - freight at or above this gets a warning (never blocked)
REPORT_MIN = 100.0         # gal - smaller amounts are solver rounding, not worth a warning

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
WINTER = {10, 11, 12, 1, 2, 3}


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #

def month_num(label: str, p: int = 0) -> int:
    """Calendar month (1-12) of a label like 'Apr-26'; unreadable labels assume an Apr..Mar year."""
    t = str(label or "")[:3].title()
    return MONTHS.index(t) + 1 if t in MONTHS else (p + 3) % 12 + 1


def vec(x, P: int, default: float = 0.0) -> np.ndarray:
    """A monthly series (list, single number or None) as a float array; None/blank -> NaN."""
    if x is None:
        return np.full(P, np.nan if default is None else default, dtype=float)
    if isinstance(x, (int, float)):
        return np.full(P, float(x), dtype=float)
    out = np.full(P, np.nan, dtype=float)
    for k, v in enumerate(list(x)[:P]):
        if v is None or v == "":
            continue
        try:
            out[k] = float(v)
        except (TypeError, ValueError):
            pass
    if len(x) < P and len(x):                 # short series: repeat the last value
        out[len(x):] = out[len(x) - 1]
    if default is not None:
        out = np.where(np.isnan(out), default, out)
    return out


def fmt_gal(v: float) -> str:
    return f"{v:,.0f} gal"


def fmt_money(v: float) -> str:
    return f"${v:,.0f}"


def content_hash(*objs) -> str:
    h = hashlib.sha1()
    for o in objs:
        h.update(json.dumps(o, sort_keys=True, default=str).encode("utf-8"))
    return h.hexdigest()[:16]


# --------------------------------------------------------------------------- #
# Demand: budget -> contract % -> weather -> adjustments -> planned
# --------------------------------------------------------------------------- #

def rule_applies(rule: dict, area_id: str, plant: dict, mon: int) -> bool:
    if rule.get("area", "*") not in ("*", area_id):
        return False
    if mon not in set(rule.get("months") or []):
        return False
    scope = rule.get("scope") or {"type": "all"}
    kind, vals = scope.get("type", "all"), set(scope.get("values") or [])
    if kind == "states":
        return (plant.get("state") or "") in vals
    if kind == "plants":
        return plant.get("id") in vals
    return True


def demand_breakdown(area: dict, state_pct: Dict[str, Dict[str, float]],
                     weather: Optional[Callable[[dict, int], float]] = None,
                     rules: Sequence[dict] = ()) -> dict:
    """
    Planned demand for every plant and month, with each step kept so people
    can see where the number comes from:

      budget   the demand that was uploaded (gallons)
      pct      contract % of budget for the plant's state and month (1.0 if not set)
      weather  weather factor for the plan's weather scenario (1.0 = normal)
      adjust   the plan's own changes ("peak winter +10%"), multiplied together
      planned  budget x pct x weather x adjust - what the optimizer uses
    """
    months = area["months"]
    P = len(months)
    plants = area.get("plants", [])
    M = len(plants)
    budget = np.zeros((M, P))
    pct = np.ones((M, P))
    wf = np.ones((M, P))
    adj = np.ones((M, P))
    for j, pl in enumerate(plants):
        budget[j] = vec(pl.get("budget"), P, 0.0)
        st = (pl.get("state") or "").upper()
        table = state_pct.get(st) or {}
        for p, lab in enumerate(months):
            mon = month_num(lab, p)
            v = table.get(MONTHS[mon - 1])
            if v is not None:
                pct[j, p] = float(v)
            if weather is not None:
                w = weather(pl, mon)
                if w is not None and math.isfinite(w):
                    wf[j, p] = w
            for r in rules:
                if rule_applies(r, area["id"], pl, mon):
                    adj[j, p] *= 1.0 + float(r.get("pct", 0)) / 100.0
    planned = np.maximum(budget * pct * wf * adj, 0.0)
    return {"budget": budget, "pct": pct, "weather": wf, "adjust": adj, "planned": planned}


# --------------------------------------------------------------------------- #
# Turning an area into arrays
# --------------------------------------------------------------------------- #

class Arrays:
    """Everything the LP needs, as numpy arrays (gallons, $/gal)."""

    def __init__(self, area: dict, demand: np.ndarray, price_add: Optional[np.ndarray] = None,
                 freight_mult: float = 1.0):
        self.area = area
        self.months = list(area["months"])
        P = self.P = len(self.months)
        sup = self.supply = list(area.get("supply", []))
        pls = self.plants = list(area.get("plants", []))
        N, M = self.N, self.M = len(sup), len(pls)
        self.notes: List[str] = []
        sidx = {s["id"]: i for i, s in enumerate(sup)}
        pidx = {p["id"]: j for j, p in enumerate(pls)}
        hubs = {h["code"]: vec(h.get("price"), P, None) for h in area.get("hubs", [])}

        self.contract = np.zeros((N, P))
        self.price = np.zeros((N, P))
        self.hub_price = np.zeros((N, P))
        self.diff = np.zeros((N, P))
        self.min_pct = np.zeros((N, P))
        self.max_pct = np.ones((N, P))
        self.s_open = np.zeros(N)
        self.s_cap = np.zeros((N, P))
        self.s_min = np.zeros((N, P))
        missing_hub = set()
        for i, s in enumerate(sup):
            self.contract[i] = np.maximum(vec(s.get("contract"), P, 0.0), 0.0)
            self.diff[i] = vec(s.get("diff"), P, 0.0)
            hp = hubs.get(s.get("hub") or "")
            if hp is None:
                if s.get("hub") and self.contract[i].sum() > 0:
                    missing_hub.add(s["hub"])
                hp = np.zeros(P)
            hp = np.where(np.isnan(hp), 0.0, hp)
            self.hub_price[i] = hp
            self.price[i] = hp + self.diff[i]
            self.min_pct[i] = np.maximum(vec(s.get("min_pct"), P, 0.0), 0.0)
            self.max_pct[i] = np.maximum(vec(s.get("max_pct"), P, 1.0), 0.0)
            self.s_open[i] = max(float(s.get("open") or 0.0), 0.0)
            self.s_cap[i] = np.maximum(vec(s.get("cap"), P, 0.0), 0.0)
            self.s_min[i] = np.maximum(vec(s.get("inv_min"), P, 0.0), 0.0)
        if price_add is not None:
            self.price = self.price + price_add
        for h in sorted(missing_hub):
            self.notes.append(f"Hub {h} has no price uploaded, so supply points priced off it use their differential only.")

        self.demand = np.maximum(np.asarray(demand, dtype=float).reshape(M, P) if M else np.zeros((0, P)), 0.0)
        self.p_open = np.zeros(M)
        self.p_cap = np.zeros((M, P))
        self.p_min = np.zeros((M, P))
        for j, pl in enumerate(pls):
            self.p_open[j] = max(float(pl.get("open") or 0.0), 0.0)
            self.p_cap[j] = np.maximum(vec(pl.get("cap"), P, 0.0), 0.0)
            self.p_min[j] = np.maximum(vec(pl.get("inv_min"), P, 0.0), 0.0)

        lanes = []
        for ln in area.get("lanes", []):
            i, j = sidx.get(ln.get("from")), pidx.get(ln.get("to"))
            if i is None or j is None:
                continue
            lanes.append((i, j, ln))
        L = self.L = len(lanes)
        self.lane_i = np.array([x[0] for x in lanes], dtype=int)
        self.lane_j = np.array([x[1] for x in lanes], dtype=int)
        self.lanes = [x[2] for x in lanes]
        self.rate = np.full((L, P), np.nan)
        self.lmin = np.zeros((L, P))
        self.lmax = np.full((L, P), np.inf)
        for l, (_, _, ln) in enumerate(lanes):
            self.rate[l] = vec(ln.get("rate"), P, None) * freight_mult
            self.lmin[l] = np.maximum(vec(ln.get("min"), P, 0.0), 0.0)
            mx = vec(ln.get("max"), P, None)
            self.lmax[l] = np.where(np.isnan(mx), np.inf, np.maximum(mx, 0.0))
        self.closed = np.isnan(self.rate)
        dropped_min = self.closed & (self.lmin > 0)
        if dropped_min.any():
            self.notes.append(f"{int(dropped_min.sum())} lane/month(s) have a minimum but no freight rate, "
                              f"so the lane isn't used there and the minimum was ignored.")
        self.lmin = np.where(self.closed, 0.0, self.lmin)
        self.lmax = np.where(self.closed, 0.0, self.lmax)
        self.high = ~self.closed & (np.nan_to_num(self.rate) >= HIGH_FREIGHT - 1e-9)

        # crossed limits - fix the obvious typo and say so
        for name, lo, hi, label in (("lift", self.min_pct, self.max_pct, "minimum lift above maximum"),
                                    ("lane", self.lmin, self.lmax, "lane minimum above its maximum"),
                                    ("sstore", self.s_min, self.s_cap, "supply storage minimum above capacity"),
                                    ("pstore", self.p_min, self.p_cap, "plant storage minimum above capacity")):
            bad = lo > hi + 1e-9
            if bad.any():
                self.notes.append(f"{int(bad.sum())} cell(s) had a {label}; the minimum was lowered to match.")
                np.minimum(lo, hi, out=lo)


# --------------------------------------------------------------------------- #
# The LP
# --------------------------------------------------------------------------- #

class LP:
    def __init__(self, a: Arrays, relaxed: bool = False):
        self.a, self.relaxed = a, relaxed
        P, N, M, L = a.P, a.N, a.M, a.L
        self.o_ship = 0
        self.o_lift = L * P
        self.o_swd = self.o_lift + N * P
        self.o_sinj = self.o_swd + N * P
        self.o_send = self.o_sinj + N * P
        self.o_pwd = self.o_send + N * P
        self.o_pinj = self.o_pwd + M * P
        self.o_pend = self.o_pinj + M * P
        n = self.o_pend + M * P
        if relaxed:
            self.o_unmet, self.o_excess = n, n + M * P
            self.o_surplus, self.o_unsourced = n + 2 * M * P, n + 2 * M * P + N * P
            n = n + 2 * M * P + 2 * N * P
        self.n = n
        self.r_sbal, self.r_sinv = 0, N * P
        self.r_pbal, self.r_pinv = 2 * N * P, 2 * N * P + M * P
        self.rows = 2 * N * P + 2 * M * P

    def build(self):
        from scipy import sparse
        a, P, N, M, L = self.a, self.a.P, self.a.N, self.a.M, self.a.L
        r, c, v = [], [], []

        def add(rows, cols, vals):
            rows, cols = np.asarray(rows).ravel(), np.asarray(cols).ravel()
            r.append(rows); c.append(cols)
            v.append(np.broadcast_to(np.asarray(vals, dtype=float), rows.shape).ravel() if np.ndim(vals) == 0
                     else np.asarray(vals, dtype=float).ravel())

        ip = np.arange(N * P)                      # (i,p) flattened i*P+p
        jp = np.arange(M * P)
        ck = (a.contract / K).ravel()
        # supply balance
        add(self.r_sbal + ip, self.o_lift + ip, ck)
        add(self.r_sbal + ip, self.o_swd + ip, 1.0)
        add(self.r_sbal + ip, self.o_sinj + ip, -1.0)
        if L:
            lp = np.arange(L * P)
            li, pp = np.divmod(lp, P)
            add(self.r_sbal + a.lane_i[li] * P + pp, self.o_ship + lp, -1.0)
            add(self.r_pbal + a.lane_j[li] * P + pp, self.o_ship + lp, 1.0)
        # supply inventory chain
        add(self.r_sinv + ip, self.o_send + ip, 1.0)
        add(self.r_sinv + ip, self.o_swd + ip, 1.0)
        add(self.r_sinv + ip, self.o_sinj + ip, -1.0)
        prev = ip[(ip % P) > 0]
        add(self.r_sinv + prev, self.o_send + prev - 1, -1.0)
        # plant balance
        add(self.r_pbal + jp, self.o_pwd + jp, 1.0)
        add(self.r_pbal + jp, self.o_pinj + jp, -1.0)
        # plant inventory chain
        add(self.r_pinv + jp, self.o_pend + jp, 1.0)
        add(self.r_pinv + jp, self.o_pwd + jp, 1.0)
        add(self.r_pinv + jp, self.o_pinj + jp, -1.0)
        prevj = jp[(jp % P) > 0]
        add(self.r_pinv + prevj, self.o_pend + prevj - 1, -1.0)
        if self.relaxed:
            add(self.r_pbal + jp, self.o_unmet + jp, 1.0)
            add(self.r_pbal + jp, self.o_excess + jp, -1.0)
            add(self.r_sbal + ip, self.o_surplus + ip, -1.0)
            add(self.r_sbal + ip, self.o_unsourced + ip, 1.0)
        A = sparse.csr_matrix((np.concatenate(v), (np.concatenate(r), np.concatenate(c))),
                              shape=(self.rows, self.n))
        b = np.zeros(self.rows)
        b[self.r_sinv + np.arange(N) * P] = a.s_open / K
        b[self.r_pbal + jp] = (a.demand / K).ravel()
        b[self.r_pinv + np.arange(M) * P] = a.p_open / K

        cost = np.zeros(self.n)
        if L:
            cost[self.o_ship:self.o_ship + L * P] = np.nan_to_num(a.rate).ravel() * K
        cost[self.o_lift:self.o_lift + N * P] = (a.contract * a.price).ravel()
        if self.relaxed:
            cost[self.o_unmet:] = PENALTY * K

        lo, hi = np.zeros(self.n), np.full(self.n, np.inf)
        if L:
            lo[self.o_ship:self.o_ship + L * P] = (a.lmin / K).ravel()
            hi[self.o_ship:self.o_ship + L * P] = (a.lmax / K).ravel()
        live = a.contract > 0
        lo[self.o_lift:self.o_lift + N * P] = np.where(live, a.min_pct, 0.0).ravel()
        hi[self.o_lift:self.o_lift + N * P] = np.where(live, a.max_pct, 0.0).ravel()
        lo[self.o_send:self.o_send + N * P] = (a.s_min / K).ravel()
        hi[self.o_send:self.o_send + N * P] = (a.s_cap / K).ravel()
        lo[self.o_pend:self.o_pend + M * P] = (a.p_min / K).ravel()
        hi[self.o_pend:self.o_pend + M * P] = (a.p_cap / K).ravel()
        return cost, A, b, lo, hi


def solve(area: dict, demand: np.ndarray, price_add: Optional[np.ndarray] = None,
          freight_mult: float = 1.0, time_limit: float = 120.0) -> dict:
    """Solve one area. Returns a result dict (see _unpack)."""
    from scipy.optimize import linprog
    t0 = time.perf_counter()
    a = Arrays(area, demand, price_add=price_add, freight_mult=freight_mult)
    if a.N == 0 or a.M == 0:
        return {"status": "empty", "ok": False, "message": "This area needs at least one supply point and one plant.",
                "notes": a.notes, "arrays": a, "seconds": 0.0}
    res, lp = None, None
    for relaxed in (False, True):
        lp = LP(a, relaxed)
        cost, A, b, lo, hi = lp.build()
        res = linprog(cost, A_eq=A, b_eq=b, bounds=np.column_stack([lo, hi]), method="highs",
                      options={"presolve": True, "time_limit": time_limit})
        if res.status == 0:
            break
        if res.status != 2:                       # not "infeasible": relaxing won't help
            break
    secs = time.perf_counter() - t0
    if res is None or res.status != 0:
        msg = {1: "it ran out of time", 2: "the data contradicts itself", 3: "it is unbounded",
               4: "of a numerical problem"}.get(getattr(res, "status", -1), "of an unknown problem")
        return {"status": "failed", "ok": False, "message": f"The optimizer couldn't finish because {msg}.",
                "notes": a.notes, "arrays": a, "seconds": secs}
    return _unpack(a, lp, res, secs)


def _unpack(a: Arrays, lp: LP, res, secs: float) -> dict:
    P, N, M, L = a.P, a.N, a.M, a.L
    x = res.x

    def blk(off, n, shape):
        return np.maximum(x[off:off + n], 0.0).reshape(shape) * K
    ship = blk(lp.o_ship, L * P, (L, P)) if L else np.zeros((0, P))
    lift = np.clip(x[lp.o_lift:lp.o_lift + N * P].reshape(N, P), 0, None)
    bought = a.contract * lift
    out = {
        "ok": True, "status": "relaxed" if lp.relaxed else "optimal", "seconds": round(secs, 2),
        "arrays": a, "notes": list(a.notes),
        "ship": ship, "lift": lift, "bought": bought,
        "s_wd": blk(lp.o_swd, N * P, (N, P)), "s_inj": blk(lp.o_sinj, N * P, (N, P)),
        "s_end": blk(lp.o_send, N * P, (N, P)),
        "p_wd": blk(lp.o_pwd, M * P, (M, P)), "p_inj": blk(lp.o_pinj, M * P, (M, P)),
        "p_end": blk(lp.o_pend, M * P, (M, P)),
    }
    z = lambda shape: np.zeros(shape)
    if lp.relaxed:
        out.update(unmet=blk(lp.o_unmet, M * P, (M, P)), excess=blk(lp.o_excess, M * P, (M, P)),
                   surplus=blk(lp.o_surplus, N * P, (N, P)), unsourced=blk(lp.o_unsourced, N * P, (N, P)))
    else:
        out.update(unmet=z((M, P)), excess=z((M, P)), surplus=z((N, P)), unsourced=z((N, P)))
    for k in ("unmet", "excess", "surplus", "unsourced"):
        out[k] = np.where(out[k] < 1e-3, 0.0, out[k])
    freight = ship * np.nan_to_num(a.rate)
    commodity = bought * a.price
    out["freight_cost"] = freight
    out["commodity_cost"] = commodity
    delivered = np.zeros((M, P))
    if L:
        np.add.at(delivered, a.lane_j, ship)
    out["delivered"] = delivered
    try:
        out["duals"] = {"upper": np.asarray(res.upper.marginals, float), "eq": np.asarray(res.eqlin.marginals, float)}
    except AttributeError:
        out["duals"] = None
    out["lp"] = lp
    out["explain"] = explain(out)
    return out


# --------------------------------------------------------------------------- #
# Reading a result
# --------------------------------------------------------------------------- #

def summary(r: dict) -> dict:
    a: Arrays = r["arrays"]
    s = {"status": r["status"], "ok": r["ok"], "seconds": r.get("seconds"), "message": r.get("message", ""),
         "demand": float(a.demand.sum()) if a.M else 0.0}
    if not r["ok"]:
        return s
    comm, frt = float(r["commodity_cost"].sum()), float(r["freight_cost"].sum())
    bought = float(r["bought"].sum())
    hv = float(r["ship"][a.high].sum()) if a.L else 0.0
    s.update({
        "commodity": comm, "freight": frt, "total": comm + frt,
        "bought": bought, "shipped": float(r["ship"].sum()),
        "unmet": float(r["unmet"].sum()), "surplus": float(r["surplus"].sum() + r["excess"].sum()),
        "unsourced": float(r["unsourced"].sum()),
        "high_volume": hv, "high_cost": float(r["freight_cost"][a.high].sum()) if a.L else 0.0,
        "per_gal": (comm + frt) / bought if bought > 1e-6 else None,
        "high_lanes": int(a.high.any(axis=1).sum()) if a.L else 0,
    })
    return s


def monthly(r: dict) -> dict:
    a: Arrays = r["arrays"]
    if not r["ok"]:
        return {}
    return {"commodity": r["commodity_cost"].sum(axis=0).tolist(), "freight": r["freight_cost"].sum(axis=0).tolist(),
            "bought": r["bought"].sum(axis=0).tolist(), "shipped": r["ship"].sum(axis=0).tolist(),
            "demand": a.demand.sum(axis=0).tolist(), "unmet": r["unmet"].sum(axis=0).tolist()}


def explain(r: dict) -> List[dict]:
    """Plain-language notes about a result, worst first."""
    a: Arrays = r["arrays"]
    out = []
    m = a.months
    if r["status"] == "relaxed":
        out.append({"level": "error", "text": "The data can't be fully met as uploaded, so this plan allows the gaps "
                    "below at a heavy cost. Fix the inputs and run again for a clean plan."})
    for key, label, names in (("unmet", "of demand can't be supplied at", a.plants),
                              ("excess", "more than it can use or store is forced into", a.plants),
                              ("surplus", "must be bought (contract minimums) with nowhere to go at", a.supply),
                              ("unsourced", "would have to ship (lane minimums) without contract volume at", a.supply)):
        arr = r.get(key)
        if arr is None or arr.sum() < REPORT_MIN:
            continue
        for k in np.argsort(-arr.sum(axis=1))[:6]:
            tot = arr[k].sum()
            if tot < REPORT_MIN:
                break
            mons = [m[p] for p in range(a.P) if arr[k, p] >= REPORT_MIN]
            nm = names[k].get("name") or names[k].get("location")
            out.append({"level": "error", "text": f"{fmt_gal(tot)} {label} {nm} ({', '.join(mons)})."})
    if a.L and r.get("ok"):
        hv = float(r["ship"][a.high].sum())
        if hv >= REPORT_MIN:
            out.append({"level": "warn", "text": f"High freight: {fmt_gal(hv)} moves on lanes with freight of $1.00/gal "
                        f"or more ({fmt_money(float(r['freight_cost'][a.high].sum()))}). They were the cheapest routes "
                        f"the optimizer found - check those rates are right."})
    for n in a.notes:
        out.append({"level": "info", "text": n})
    return out


def insights(r: dict, top: int = 12) -> dict:
    """
    What one more gallon is worth, from the optimizer's shadow prices ($/gal),
    and how hard each contract is worked month by month.
    """
    out = {"contracts": [], "lanes": [], "storage": [], "demand": [], "use": [], "penalty": r["status"] == "relaxed"}
    a: Arrays = r["arrays"]
    if not r["ok"] or not r.get("duals"):
        return out
    lp: LP = r["lp"]
    up, eq = r["duals"]["upper"], r["duals"]["eq"]
    P, N, M, L = a.P, a.N, a.M, a.L
    tol = 1e-7

    def item(months, vals, **kw):
        return {"months": [int(p) for p in months], "value": float(np.mean(vals)), **kw}
    mu = up[lp.o_lift:lp.o_lift + N * P].reshape(N, P)
    with np.errstate(divide="ignore", invalid="ignore"):
        per = np.where(a.contract > 0, mu / np.where(a.contract > 0, a.contract, 1.0), 0.0)   # $/gal
    for i in range(N):
        ms = [p for p in range(P) if per[i, p] < -tol]
        if ms:
            out["contracts"].append(item(ms, -per[i, ms], supply=a.supply[i]["id"]))
    if L:
        sm = up[lp.o_ship:lp.o_ship + L * P].reshape(L, P) / K
        for l in range(L):
            ms = [p for p in range(P) if sm[l, p] < -tol]
            if ms:
                out["lanes"].append(item(ms, -sm[l, ms], frm=a.lanes[l]["from"], to=a.lanes[l]["to"]))
    for kind, off, n, ents in (("supply", lp.o_send, N, a.supply), ("plant", lp.o_pend, M, a.plants)):
        st = up[off:off + n * P].reshape(n, P) / K
        for k in range(n):
            ms = [p for p in range(P) if st[k, p] < -tol]
            if ms:
                out["storage"].append(item(ms, -st[k, ms], kind=kind, id=ents[k]["id"]))
    lam = eq[lp.r_pbal:lp.r_pbal + M * P].reshape(M, P) / K
    for j in range(M):
        ms = [p for p in range(P) if a.demand[j, p] > 0]
        if ms:
            e = item(ms, lam[j, ms], plant=a.plants[j]["id"])
            e["by_month"] = [float(v) for v in lam[j]]
            e["dearest"] = float(np.max(lam[j, ms]))
            out["demand"].append(e)
    for k in ("contracts", "lanes", "storage"):
        out[k] = sorted(out[k], key=lambda e: -e["value"])[:top]
    out["demand"] = sorted(out["demand"], key=lambda e: -e["value"])[:top]
    for i, s in enumerate(a.supply):
        live = a.contract[i] > 0
        if not live.any():
            continue
        at_max = live & (r["lift"][i] >= a.max_pct[i] - 1e-6)
        at_min = live & ~at_max & (r["lift"][i] <= a.min_pct[i] + 1e-6)
        out["use"].append({"supply": s["id"], "bought": float(r["bought"][i].sum()),
                           "contract": float(a.contract[i].sum()),
                           "months": ["" if not live[p] else "max" if at_max[p] else "min" if at_min[p] else "mid"
                                      for p in range(P)],
                           "at_max": int(at_max.sum()), "at_min": int(at_min.sum()), "live": int(live.sum())})
    return out


# --------------------------------------------------------------------------- #
# Data checks (before solving)
# --------------------------------------------------------------------------- #

def checks(area: dict, demand: np.ndarray) -> List[dict]:
    """Quick, plain-language checks on an area's data. level: ok | info | warn | error."""
    a = Arrays(area, demand)
    out: List[dict] = []
    P, m = a.P, a.months
    if a.N == 0:
        out.append({"level": "error", "text": "No supply points yet - upload supply points and contracts."})
    if a.M == 0:
        out.append({"level": "error", "text": "No plants yet - upload plants and demand."})
    if not a.N or not a.M:
        return out
    if a.L == 0:
        out.append({"level": "error", "text": "No lanes (routes from supply points to plants) yet - upload lanes and freight rates."})
        return out
    open_to = np.zeros((a.M, P), dtype=bool)
    open_from = np.zeros((a.N, P), dtype=bool)
    for l in range(a.L):
        open_to[a.lane_j[l]] |= ~a.closed[l]
        open_from[a.lane_i[l]] |= ~a.closed[l]
    stranded = [(j, [m[p] for p in range(P) if a.demand[j, p] > 0 and not open_to[j, p] and a.p_cap[j, p] <= 0])
                for j in range(a.M)]
    stranded = [(j, ms) for j, ms in stranded if ms]
    if stranded:
        out.append({"level": "error", "text": f"{len(stranded)} plant(s) need product in months with no lane to them: "
                    + "; ".join(f"{a.plants[j]['name']} ({', '.join(ms[:4])}{'...' if len(ms) > 4 else ''})" for j, ms in stranded[:5])})
    no_route = [i for i in range(a.N) if a.contract[i].sum() > 0 and not open_from[i].any()]
    if no_route:
        out.append({"level": "warn", "text": f"{len(no_route)} supply point(s) have a contract but no lane out: "
                    + ", ".join(a.supply[i]["location"] for i in no_route[:6])})
    can = (a.contract * a.max_pct).sum(axis=0) + a.s_open.sum() * (np.arange(P) == 0)
    need = a.demand.sum(axis=0) - (a.p_open.sum() * (np.arange(P) == 0))
    short = [(p, need[p] - can[p]) for p in range(P) if need[p] - can[p] > REPORT_MIN and a.s_cap[:, p].sum() + a.p_cap[:, p].sum() <= 0]
    if short:
        out.append({"level": "warn", "text": "Contracts at their maximum can't cover demand in "
                    + ", ".join(f"{m[p]} (short {fmt_gal(v)})" for p, v in short[:4])
                    + (" and more" if len(short) > 4 else "") + ". Unmet demand will show in the results."})
    forced = (a.contract * a.min_pct).sum(axis=0)
    over = [(p, forced[p] - a.demand[:, p].sum()) for p in range(P)
            if forced[p] - a.demand[:, p].sum() - a.s_cap[:, p].sum() - a.p_cap[:, p].sum() > REPORT_MIN]
    if over:
        out.append({"level": "warn", "text": "Contract minimums force more product than plants can take in "
                    + ", ".join(f"{m[p]} (+{fmt_gal(v)})" for p, v in over[:4]) + "."})
    hl = int(a.high.any(axis=1).sum())
    if hl:
        out.append({"level": "info", "text": f"{hl} lane(s) have freight of $1.00/gal or more in some month. They're used "
                    f"like any lane when cheapest, and flagged in the results."})
    no_state = [p["name"] for p in a.plants if not p.get("state")]
    if no_state:
        out.append({"level": "info", "text": f"{len(no_state)} plant(s) have no state, so state contract % and state "
                    f"weather don't apply to them: {', '.join(no_state[:5])}{'...' if len(no_state) > 5 else ''}"})
    for n in a.notes:
        out.append({"level": "info", "text": n})
    if not any(c["level"] in ("error", "warn") for c in out):
        out.insert(0, {"level": "ok", "text": "Data looks complete: every plant has a route, and contracts can cover demand."})
    return out
