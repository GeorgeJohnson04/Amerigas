"""
paste.py - turn a block copied from Excel into changes to a working model.

Data arrives in all sorts of Excel layouts, so instead of asking people to
reformat it, this works out the layout from what was pasted:

  wide     one row per plant / supply point, one column per month
           (header row with month names: Apr, Apr-26, 4/1/2026, April 2026, P1 ...)
  tall     the same, turned on its side: months down, names across
  long     one row per name and month: name | month | value
  list     one row per name with a single value (applies to the months you pick)
  matrix   (lanes) supply points down, plants across - one month at a time
  lanes    (lanes) one row per lane: from | to | value(s), with months across
           or a month column, or one value for the months you pick

Names are matched loosely: exact name, district number ("1001"), the name
without the code ("Springfield ST"), supplier + location, hub code / name, then a
fuzzy match on whatever's left. Every match is shown with how sure it is, and
people can change or skip any of them before anything is saved.

Values are converted from whatever unit they arrive in (gallons or thousand
gallons, $/gal or $/thousand gal, 90% or 0.9) - detected by comparing with
what the model already holds, and changeable.

Nothing here writes anything: analyse() returns the plan (matches, changes,
warnings) and app.py applies it as ordinary cell edits, so pasted changes show
up highlighted in the grids and in the change log, and can be reverted.
"""

from __future__ import annotations

import difflib
import math
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np

import model

GAL = model.GAL_PER_KGAL
MAX_ROWS = 20000
MAX_CHARS = 8_000_000

#: kind -> (model field, what each row names, unit kind, label)
KINDS: Dict[str, Tuple[str, str, str, str]] = {
    "demand": ("demand", "plant", "volume", "Demand"),
    "demand_inv_init": ("demand_inv_init", "plant", "volume", "Plant opening inventory"),
    "demand_inv_cap": ("demand_inv_cap", "plant", "volume", "Plant storage capacity"),
    "demand_inv_min": ("demand_inv_min", "plant", "volume", "Plant minimum inventory"),
    "contract": ("contract", "supply", "volume", "Contract volume"),
    "hub_premium": ("hub_premium", "supply", "price", "Differential to hub"),
    "hub_price": ("hub_price", "supply_hub", "price", "Hub price"),
    "perc_min": ("perc_min", "supply", "pct", "Minimum lift"),
    "perc_max": ("perc_max", "supply", "pct", "Maximum lift"),
    "supply_inv_init": ("supply_inv_init", "supply", "volume", "Supply opening inventory"),
    "supply_inv_cap": ("supply_inv_cap", "supply", "volume", "Supply storage capacity"),
    "supply_inv_min": ("supply_inv_min", "supply", "volume", "Supply minimum inventory"),
    "trans_cost": ("trans_cost", "lane", "price", "Freight rates"),
    "ship_min": ("ship_min", "lane", "volume", "Minimum shipments"),
    "ship_max": ("ship_max", "lane", "volume", "Maximum shipments"),
}

AUTO_OK = 0.86        # at or above: matched
AUTO_CHECK = 0.70     # at or above: matched, but flagged "check"

_MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
_MONTH_FULL = ["january", "february", "march", "april", "may", "june", "july", "august",
               "september", "october", "november", "december"]
_BLANK_TEXT = {"", "-", "--", "—", "n/a", "na", "#n/a", "x", "none", "null", "#value!", "#div/0!"}
_SKIP_LABEL = re.compile(r"^(grand\s*)?(sub\s*)?totals?\b|^summer$|^winter$|^ratio$|^average$|^avg$", re.I)
_PLACEHOLDER = re.compile(r"^(supplier|demand|supply)\s*\d+$", re.I)
_STATE = set("AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM "
             "NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY".split())


def fold(s: str) -> str:
    return model.fold_name(s or "")


def col_name(c: int) -> str:
    s = ""
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


# --------------------------------------------------------------------------- #
# Cells
# --------------------------------------------------------------------------- #

BLANK = "blank"
TEXT = "text"


def parse_number(s: str) -> Tuple[str, Optional[float]]:
    """('num', v) | ('pct', fraction) | ('blank', None) | ('text', None)."""
    t = (s or "").strip()
    if t.lower() in _BLANK_TEXT:
        return BLANK, None
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()").replace(",", "").replace("$", "").replace(" ", "").replace("\u00a0", "")
    pct = t.endswith("%")
    t = t.rstrip("%")
    if t.startswith("-") and t[1:2] == "$":
        t = "-" + t[2:]
    try:
        v = float(t)
    except ValueError:
        return TEXT, None
    if not math.isfinite(v):
        return TEXT, None
    v = -v if neg else v
    return ("pct", v / 100.0) if pct else ("num", v)


def parse_month(s: str):
    """A month header -> ('m', month 1-12, year or None) | ('p', period 1-based) | None."""
    t = (s or "").strip().lower().replace(".", "")
    if not t or len(t) > 32:
        return None
    m = re.fullmatch(r"(?:p|period|month)\s*(\d{1,2})", t)
    if m:
        return ("p", int(m.group(1)))
    m = re.fullmatch(r"([a-z]{3,9})(?:[\s\-/'_,]*(\d{2}|\d{4}))?", t)
    if m:
        w = m.group(1)
        k = next((i for i in range(12) if w in (_MONTHS[i], _MONTH_FULL[i]) or
                  (len(w) >= 3 and _MONTH_FULL[i].startswith(w))), None)
        if k is None:
            return None
        y = m.group(2)
        return ("m", k + 1, (2000 + int(y) if y and len(y) == 2 else int(y) if y else None))
    m = re.fullmatch(r"(\d{2}|\d{4})[\s\-/]([a-z]{3,9})", t)          # 26-apr, 2026-apr
    if m:
        r = parse_month(m.group(2))
        if r and r[0] == "m":
            y = m.group(1)
            return ("m", r[1], 2000 + int(y) if len(y) == 2 else int(y))
    m = re.fullmatch(r"(\d{4})-(\d{1,2})(?:-(\d{1,2}))?(?:[ t].*)?", t)
    if m and 1 <= int(m.group(2)) <= 12:
        return ("m", int(m.group(2)), int(m.group(1)))
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2}|\d{4})(?:\s.*)?", t)
    if m and 1 <= int(m.group(1)) <= 12:
        y = m.group(3)
        return ("m", int(m.group(1)), 2000 + int(y) if len(y) == 2 else int(y))
    m = re.fullmatch(r"(\d{1,2})/(\d{4})", t)
    if m and 1 <= int(m.group(1)) <= 12:
        return ("m", int(m.group(1)), int(m.group(2)))
    return None


def label_months(labels: Sequence[str]) -> List[Tuple[int, Optional[int]]]:
    """(month, year) of each model period; unreadable labels assume Apr..Mar order."""
    out = []
    for p, lab in enumerate(labels):
        r = parse_month(lab)
        out.append((r[1], r[2]) if r and r[0] == "m" else ((p + 3) % 12 + 1, None))
    return out


def period_of(tok, labels: Sequence[str], lm) -> Tuple[Optional[int], bool]:
    """Which model period a parsed month header is. Returns (index or None, year differs)."""
    if tok is None:
        return None, False
    if tok[0] == "p":
        k = tok[1] - 1
        return (k if 0 <= k < len(labels) else None), False
    _, month, year = tok
    cands = [(p, y) for p, (mm, y) in enumerate(lm) if mm == month]
    if year is not None:
        exact = [p for p, y in cands if y == year]
        if len(exact) == 1:
            return exact[0], False
    if len(cands) == 1:
        return cands[0][0], bool(year and cands[0][1] and cands[0][1] != year)
    return None, False


# --------------------------------------------------------------------------- #
# Names
# --------------------------------------------------------------------------- #

def _tokens(s: str) -> List[str]:
    return [t for t in re.findall(r"[a-z0-9]+", (s or "").lower()) if len(t) >= 2]


def _code(s: str) -> str:
    m = re.match(r"\s*#?(\d{3,4})\b", s or "")
    return m.group(1).zfill(4) if m else ""


@dataclass
class Entity:
    key: str
    region: str
    area: str
    kind: str                   # plant | supply | hub
    targets: List[int]          # positions in the model (a hub covers several supply rows)
    label: str
    aliases: List[str]
    code: str = ""
    tokens: set = field(default_factory=set)

    def info(self) -> dict:
        return {"key": self.key, "label": self.label, "area": self.area, "kind": self.kind}


def _aliases(full: str, others) -> List[str]:
    """Folded aliases, the full name first (it scores higher), no repeats."""
    out = [fold(full)]
    for x in others:
        f = fold(x)
        if f and f not in out:
            out.append(f)
    return [a for a in out if a]


def plant_entities(region: str, area: str, d) -> List[Entity]:
    out = []
    for j, n in enumerate(d.demand_names):
        if _PLACEHOLDER.match(n.strip()):
            continue
        bare = re.sub(r"^\s*\d{3,4}\s*-\s*", "", n)
        bare2 = re.sub(r"\(.*?\)", "", bare)
        no_state = re.sub(r"[\s,]+[A-Z]{2}\s*$", "", bare2.strip())
        al = _aliases(n, (bare, bare2, no_state))
        out.append(Entity(f"{region}|{area}|p|{j}", region, area, "plant", [j], n, al, _code(n),
                          set(_tokens(n))))
    return out


def supply_entities(region: str, area: str, d) -> List[Entity]:
    out = []
    for i, (loc, grp) in enumerate(zip(d.supply_names, d.supply_groups)):
        if _PLACEHOLDER.match(loc.strip()) or _PLACEHOLDER.match((grp or "").strip()):
            continue
        g = (grp or "").replace(model.BUILDER_TAG, "").strip()
        # "Supplier Location" is as exact as the location on its own.
        al = _aliases(loc, (f"{g} {loc}", f"{loc} {g}"))
        out.append(Entity(f"{region}|{area}|s|{i}", region, area, "supply", [i],
                          f"{loc} | {g}" if g else loc, al, "", set(_tokens(f"{loc} {g}"))))
    return out


def hub_entities(region: str, area: str, d) -> List[Entity]:
    out = []
    for h in sorted({h for h in d.supply_hubs if h}):
        rows = [i for i, x in enumerate(d.supply_hubs) if x == h]
        desc = model.HUB_DESCRIPTIONS.get(h, "")
        al = _aliases(h, (desc, f"{h} {desc}", f"{desc} hub", f"{h} hub"))
        out.append(Entity(f"{region}|{area}|h|{h}", region, area, "hub", rows,
                          f"{h} hub" + (f" ({desc})" if desc else ""), al, "", set(_tokens(f"{h} {desc}"))))
    return out


class Matcher:
    """Finds the entity a piece of text names, with a score 0..1."""

    def __init__(self, ents: List[Entity]):
        self.ents = ents
        self.by_alias: Dict[str, List[Entity]] = {}
        self.by_code: Dict[str, List[Entity]] = {}
        self.by_token: Dict[str, List[Entity]] = {}
        self.by_key = {e.key: e for e in ents}
        for e in ents:
            # The full name is a perfect match; shortened forms ("Springfield ST" for
            # "1001-Springfield ST") a near-perfect one, so "Riverton, ST" picks the plant
            # called exactly that over "1002-Riverton ST (T)".
            for n, a in enumerate(e.aliases):
                self.by_alias.setdefault(a, []).append((1.0 if n == 0 else 0.97, e))
            if e.code:
                self.by_code.setdefault(e.code, []).append(e)
            for t in e.tokens:
                if len(t) >= 3:
                    self.by_token.setdefault(t, []).append(e)
        self._cache: Dict[str, List[Tuple[float, Entity]]] = {}
        #: On a tie, an entity of this kind wins (hub prices: "Conway" = the CON hub).
        self.prefer = ""

    def scored(self, text: str) -> List[Tuple[float, Entity]]:
        """Candidates for one string, best first (cached)."""
        if text in self._cache:
            return self._cache[text]
        f = fold(text)
        res: Dict[str, Tuple[float, Entity]] = {}

        def put(score, e):
            if e.key not in res or res[e.key][0] < score:
                res[e.key] = (score, e)
        if f:
            for sc, e in self.by_alias.get(f, []):
                put(sc, e)
            code = _code(text)
            if code:
                for e in self.by_code.get(code, []):
                    put(0.96, e)
            if not res:
                toks = _tokens(text)
                short = {e.key: e for t in toks if len(t) >= 3 for e in self.by_token.get(t, [])}
                for e in short.values():
                    best = 0.0
                    for a in e.aliases:
                        if len(a) >= 5 and len(f) >= 5 and (a in f or f in a):
                            best = max(best, 0.88)
                        best = max(best, difflib.SequenceMatcher(None, f, a).ratio())
                    jac = len(e.tokens & set(toks)) / max(1, len(e.tokens | set(toks)))
                    best = max(best, 0.5 + 0.45 * jac if jac >= 0.5 else best)
                    put(best, e)
        out = sorted(res.values(), key=lambda x: -x[0])
        self._cache[text] = out
        return out

    def best(self, texts: Sequence[str], restrict=None) -> dict:
        """
        Best match for a row's label cells. Tries each cell, neighbouring pairs
        and all of them joined, and keeps the most certain, preferring a unique
        hit (two supply rows called "Northfield" need the supplier too).
        """
        texts = [t for t in texts if t and t.strip()]
        cands = list(texts)
        cands += [f"{a} {b}" for a, b in zip(texts, texts[1:])]
        if len(texts) > 2:
            cands.append(" ".join(texts))
        best = None
        for c in cands:
            sc = [(s, e) for s, e in self.scored(c) if restrict is None or restrict(e)]
            if not sc:
                continue
            top = sc[0][0]
            tied = [e for s, e in sc if s >= top - 0.015]
            if len(tied) > 1 and self.prefer:
                pref = [e for e in tied if e.kind == self.prefer]
                if len(pref) == 1:
                    tied = pref
            unique = len(tied) == 1
            # A sure, unique hit beats a perfect but ambiguous one ("SupplierX
            # Northfield" beats "Northfield" when two rows are called Northfield).
            rank = (unique and top >= AUTO_OK, top, unique, -len(tied))
            if best is None or rank > best["rank"]:
                best = {"rank": rank, "score": top, "entity": tied[0] if unique else None,
                        "options": [e.key for e in tied[:6]], "text": c}
        if best is None:
            return {"status": "unmatched", "score": 0.0, "entity": None, "options": []}
        if best["score"] < AUTO_CHECK:
            best["status"] = "unmatched"
            best["entity"] = None
        elif best["entity"] is None:
            best["status"] = "ambiguous"
        else:
            best["status"] = "matched" if best["score"] >= AUTO_OK else "check"
        best.pop("rank")
        return best


# --------------------------------------------------------------------------- #
# Tables
# --------------------------------------------------------------------------- #

def read_table(text: str) -> List[List[str]]:
    if len(text) > MAX_CHARS:
        raise ValueError("that's a very large paste - paste one area or one sheet at a time")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    rows = [r.split("\t") for r in text.split("\n")]
    if len(rows) > MAX_ROWS:
        raise ValueError(f"more than {MAX_ROWS:,} rows - paste a smaller block")
    if all(len(r) == 1 for r in rows) and any("," in r[0] for r in rows):
        import csv, io                                   # comma-separated (a CSV opened in Notepad)
        rows = list(csv.reader(io.StringIO(text)))
    rows = [[c.strip().strip('"') for c in r] for r in rows]
    while rows and not any(rows[-1]):
        rows.pop()
    while rows and not any(rows[0]):
        rows.pop(0)
    width = max((len(r) for r in rows), default=0)
    rows = [r + [""] * (width - len(r)) for r in rows]
    keep = [c for c in range(width) if any(r[c] for r in rows)]   # drop empty columns
    return [[r[c] for c in keep] for r in rows], keep


def _is_text(s: str) -> bool:
    return parse_number(s)[0] == TEXT


def _month_row(row, labels, lm):
    hits = {}
    for c, v in enumerate(row):
        p, _ = period_of(parse_month(v), labels, lm)
        if p is not None:
            hits.setdefault(c, p)
    return hits


def _blocks(cols: Dict[int, int]) -> List[List[int]]:
    """Month columns grouped into runs (a sheet can have Unadjusted Apr..Mar, then Adjusted Apr..Mar)."""
    out: List[List[int]] = []
    seen: set = set()
    for c in sorted(cols):
        if not out or cols[c] in seen or c - out[-1][-1] > 2:
            out.append([c])
            seen = {cols[c]}
        else:
            out[-1].append(c)
            seen.add(cols[c])
    return out


# --------------------------------------------------------------------------- #
# Analysis
# --------------------------------------------------------------------------- #

@dataclass
class Item:
    """One value to write: which row / lane names it, which months, the raw value."""
    row: int
    a: dict                    # match for the row label (plant / supply / from)
    b: Optional[dict]          # match for the lane's other end (to), or None
    values: Dict[int, Tuple[str, Optional[float]]]   # period -> parsed value; -1 = the single value
    a_text: str = ""
    b_text: str = ""


def input_array(d, fld: str) -> np.ndarray:
    """An input as entered (a closed lane's freight is NaN = blank)."""
    return np.asarray(d.freight_as_entered() if fld == "trans_cost" else getattr(d, fld), dtype=float)


class Context:
    """
    The models a paste can go into: one area, or every area of the base.
    *areas* is [(region, area, model now, model before this working model's
    cell edits)] - the second is what "back to the base value" means.
    """

    def __init__(self, areas: List[Tuple[str, str, object, object]], kind: str):
        self.kind = kind
        self.field, self.entity, self.unit, self.label = KINDS[kind]
        self.axes = model.CELL_EDIT_FIELDS[self.field]
        self.multi = len(areas) > 1
        self.current = {(r, a): input_array(d, self.field) for r, a, d, _ in areas}
        self.base = {(r, a): input_array(b, self.field) for r, a, _, b in areas}
        first = areas[0][2]
        self.labels = list(first.period_labels)
        self.lm = label_months(self.labels)
        plants, supplies, hubs = [], [], []
        for r, a, d, _ in areas:
            plants += plant_entities(r, a, d)
            supplies += supply_entities(r, a, d)
            hubs += hub_entities(r, a, d)
        self.plants, self.supplies, self.hubs = Matcher(plants), Matcher(supplies), Matcher(hubs + supplies)
        self.hubs.prefer = "hub"

    def row_matcher(self) -> Matcher:
        return {"plant": self.plants, "supply": self.supplies, "supply_hub": self.hubs,
                "lane": self.supplies}[self.entity]


def analyse(ctx: Context, text: str, opts: Optional[dict] = None, mapping: Optional[dict] = None) -> dict:
    opts = dict(opts or {})
    mapping = mapping or {}
    table, kept_cols = read_table(text)
    if not table:
        raise ValueError("nothing to read - copy the cells in Excel (including the header row) and paste here")
    notes: List[str] = []
    lane = ctx.entity == "lane"
    single = "p" not in ctx.axes
    layout, items, info = (_lanes(ctx, table, opts, notes) if lane else _rows(ctx, table, opts, notes))
    info["source_cols"] = [col_name(c) for c in kept_cols]

    # ---- name decisions (with the user's overrides) ------------------------ #
    # One record per distinct name; the same supply-point name in two areas
    # (every area has a "Conway") is two records when pasting into all areas.
    names: Dict[str, dict] = {}

    def decide(side: str, text_: str, m: dict, matcher: Matcher, area: str = "") -> Optional[Entity]:
        key = f"{area}::{text_}" if area else text_
        if (side, key) not in names:
            names[(side, key)] = {"side": side, "key": key, "text": text_, "area": area,
                                  "status": m["status"], "score": round(m["score"], 2),
                                  "suggested": m["entity"].key if m["entity"] else "",
                                  "options": m.get("options", []), "rows": 0, "choice": None}
        rec = names[(side, key)]
        rec["rows"] += 1
        over = (mapping.get(side) or {}).get(key)
        if over is not None:
            rec["choice"] = over
            return matcher.by_key.get(over) if over else None
        return m["entity"]

    # ---- units ------------------------------------------------------------- #
    factor, unit_used = _unit(ctx, items, opts, decide, notes)
    blank_mode = opts.get("blanks", "skip")
    months_opt = opts.get("months")
    months_all = list(range(len(ctx.labels))) if months_opt in (None, "", "all") else \
        [int(p) for p in months_opt if 0 <= int(p) < len(ctx.labels)]

    # ---- the changes ---------------------------------------------------------- #
    writes: Dict[Tuple[str, str, tuple], Optional[float]] = {}
    origin: Dict[Tuple[str, str, tuple], int] = {}
    dup = 0
    for it in items:
        eb = None
        if lane:
            eb = decide("b", it.b_text, it.b, ctx.plants)
            if eb is None:
                continue
            if ctx.multi and it.a.get("entity") is not None and it.a["entity"].area != eb.area:
                # the supply name was matched before we knew the plant's area: redo it there
                it.a = ctx.supplies.best(it.a_text.split(" \u00b7 "), restrict=lambda e, ar=eb.area: e.area == ar)
            ea = decide("a", it.a_text, it.a, ctx.supplies, eb.area if ctx.multi else "")
            if ea is None:
                continue
            if (eb.region, eb.area) != (ea.region, ea.area):
                names[("a", f"{eb.area}::{it.a_text}" if ctx.multi else it.a_text)]["status"] = "other area"
                continue
        else:
            if ctx.multi and it.a.get("status") == "ambiguous":
                # The same plant can be in two models (a district area and a rail
                # area). A "Model" column on the row says which one is meant.
                said = {fold(t) for t in it.a_text.split(" \u00b7 ")}
                pick = [k for k in it.a.get("options", [])
                        if fold(re.sub(r"\s*\d{2}-\d{2}\s*$", "", ctx.row_matcher().by_key[k].area)) in said]
                if len(pick) == 1:
                    it.a = {**it.a, "status": "matched", "entity": ctx.row_matcher().by_key[pick[0]]}
            ea = decide("a", it.a_text, it.a, ctx.row_matcher())
            if ea is None:
                continue
        for p, (kind, v) in it.values.items():
            periods = months_all if p == -1 and not single else [p]
            if single:
                periods = [None]
            if kind == BLANK:
                if blank_mode != "clear":
                    continue
                val = _cleared(ctx.field)
            elif kind == TEXT:
                continue
            else:
                val = _convert(ctx, kind, v, factor)
            for pp in periods:
                for t in ea.targets:
                    if single:
                        idx = (t,)
                    elif lane:
                        idx = (pp, t, eb.targets[0])
                    else:
                        idx = (pp, t)
                    k = (ea.region, ea.area, idx)
                    if k in writes and origin[k] != it.row:
                        dup += 1
                    writes[k] = val
                    origin[k] = it.row
    if dup:
        notes.append(f"{dup:,} cell(s) were given more than once (the same name on two rows) - the last one is used.")

    # ---- compare with what the model holds now ---------------------------- #
    changes, by_area, biggest = [], {}, []
    total_before = total_after = 0.0
    for (region, area, idx), val in writes.items():
        cur_arr, base_arr = ctx.current[(region, area)], ctx.base[(region, area)]
        if any(i >= s for i, s in zip(idx, cur_arr.shape)):
            continue
        cur, base = float(cur_arr[idx]), float(base_arr[idx])
        if ctx.unit == "volume":
            total_before += cur if np.isfinite(cur) else 0.0
            total_after += float(val) if val is not None else 0.0
        tol = TOLERANCE[ctx.unit]
        if _same(val, cur, tol):
            continue
        rec = {"region": region, "area": area, "index": list(idx), "value": val,
               "revert": _same(val, base, tol)}
        changes.append(rec)
        st = by_area.setdefault(area, {"region": region, "area": area, "cells": 0})
        st["cells"] += 1
        biggest.append((abs((val if val is not None else 0) - (cur if np.isfinite(cur) else 0)), region, area, idx,
                        float(cur) if np.isfinite(cur) else None, val))
    biggest.sort(key=lambda x: -x[0])

    name_list = sorted(names.values(), key=lambda r: ({"unmatched": 0, "ambiguous": 1, "other area": 2,
                                                         "check": 3, "matched": 4}[r["status"]], r["text"].lower()))
    covered = {(r, a, i[0] if single else i[1]) for r, a, i in writes}
    missing = []
    if not lane:
        for e in ctx.row_matcher().ents:
            if e.kind != "hub" and not any((e.region, e.area, t) in covered for t in e.targets):
                missing.append(e.label)
    return {
        "kind": ctx.kind, "field": ctx.field, "label": ctx.label, "layout": layout, "info": info,
        "unit": unit_used, "unit_kind": ctx.unit, "rows_read": len(table),
        "names": name_list,
        "changes": len(changes), "reverts": sum(1 for c in changes if c["revert"]),
        "cells_given": len(writes), "areas": sorted(by_area.values(), key=lambda x: x["area"]),
        "missing": missing[:200], "missing_count": len(missing),
        "total_before": total_before, "total_after": total_after,
        "biggest": [{"region": r, "area": a, "index": list(i), "old": o, "new": n} for _, r, a, i, o, n in biggest[:15]],
        "notes": notes, "months": ctx.labels, "single": single,
        "entities": {"a": [e.info() for e in (ctx.supplies if lane else ctx.row_matcher()).ents],
                     "b": [e.info() for e in ctx.plants.ents] if lane else []},
        "_changes": changes,
    }


def _cleared(fld: str) -> Optional[float]:
    if fld == "trans_cost":
        return None                     # blank freight = lane not used
    if fld == "ship_max":
        return model.DEFAULT_MAX_SHIP   # blank maximum = no limit
    return 0.0


def _convert(ctx: Context, kind: str, v: float, factor: float) -> float:
    if ctx.unit == "pct":
        return v if kind == "pct" else v * factor
    return v * factor


def _unit(ctx: Context, items: List[Item], opts: dict, decide, notes) -> Tuple[float, str]:
    """Factor from what was pasted to model units (kgal, $/kgal, fraction), and its name."""
    want = opts.get("unit") or "auto"
    nums = [v for it in items for k, v in it.values.values() if k == "num" and v]
    if ctx.unit == "pct":
        if want == "percent":
            return 0.01, "percent"
        if want == "fraction":
            return 1.0, "fraction"
        big = any(abs(v) > 3 for v in nums)
        return (0.01, "percent") if big else (1.0, "fraction")
    choices = {"volume": {"gal": 1 / GAL, "kgal": 1.0}, "price": {"$/gal": GAL, "$/kgal": 1.0}}[ctx.unit]
    if want in choices:
        return choices[want], want
    # auto: compare with what the matched cells hold now
    pairs = []
    lane = ctx.entity == "lane"
    for it in items[:4000]:
        ea = it.a.get("entity")
        eb = it.b.get("entity") if lane and it.b else None
        if ea is None or (lane and eb is None):
            continue
        cur = ctx.current.get((ea.region, ea.area))
        if cur is None:
            continue
        for p, (k, v) in it.values.items():
            if k != "num" or not v:
                continue
            pp = 0 if p == -1 else p
            try:
                if "p" not in ctx.axes:
                    c = cur[ea.targets[0]]
                elif lane:
                    c = cur[pp, ea.targets[0], eb.targets[0]]
                else:
                    c = cur[pp, ea.targets[0]]
            except IndexError:
                continue
            if np.isfinite(c) and c > 0 and not (ctx.field == "ship_max" and c >= model.DEFAULT_MAX_SHIP):
                pairs.append((abs(v), float(c)))
    if len(pairs) >= 3:
        best = min(choices.items(), key=lambda kv: float(np.median([abs(math.log10(v * kv[1] / c)) for v, c in pairs])))
        notes.append(f"Units: read as {best[0]} (compared with the values the model holds now).")
        return best[1], best[0]
    med = float(np.median([abs(v) for v in nums])) if nums else 0.0
    if ctx.unit == "volume":
        name = "gal" if med > 3000 else "kgal"
    else:
        name = "$/gal" if med < 5 else "$/kgal"
    notes.append(f"Units: read as {name} (from the size of the numbers) - change it below if that's wrong.")
    return choices[name], name


def _header_names(table, matcher: Matcher, max_rows: int = 8):
    """
    The first row whose cells name several different entities, and nothing
    else: (row, {col: match}). A header row holds no values, so a row with a
    number in it (a data row whose location and supplier both match) is not one.
    """
    for r in range(min(max_rows, len(table))):
        row = table[r]
        if any(v and not _nameish(v) for v in row):
            continue
        hits, seen = {}, set()
        for c, v in enumerate(row):
            if v:
                m = matcher.best([v])
                if m["status"] in ("matched", "check", "ambiguous"):
                    hits[c] = m
                    seen.add(m["entity"].key if m["entity"] else v)
        if len(seen) >= 2 and len(hits) >= 0.5 * sum(1 for v in row if v):
            return r, hits
    return None, {}


def _label_cells(row, cols) -> List[str]:
    # A district number with its leading zero ("0101") is a name, not a value.
    out = []
    for c in cols:
        v = row[c]
        if v and (_is_text(v) or re.fullmatch(r"0\d{3}", v)) and v not in out:
            out.append(v)
    return out


def _value_cols(data, cols, header) -> List[int]:
    """Columns that hold values: numbers (or blanks / n/a), hardly any text."""
    out = []
    for c in cols:
        kinds = [parse_number(row[c])[0] for row in data if row[c]]
        if not kinds:
            continue
        nums = sum(k in ("num", "pct") for k in kinds)
        texts = sum(k == TEXT for k in kinds)
        if texts <= 0.2 * len(kinds) and (nums or (header and header[c])):
            out.append(c)
    return out


#: Below these differences (model units) a pasted value counts as "the same" -
#: Excel sheets are rounded, and half a gallon is not a change anyone means.
TOLERANCE = {"volume": 0.0005, "price": 0.005, "pct": 1e-6}     # 0.5 gal, $0.000005/gal, 0.0001%


def _same(a, b, tol: float = 1e-9) -> bool:
    an = a is None or not math.isfinite(a)
    bn = b is None or not math.isfinite(b)
    if an or bn:
        return an and bn
    return abs(a - b) <= max(tol, 1e-9 * abs(b))


_CODE_CELL = re.compile(r"#?\d{3,4}")


def _nameish(v: str) -> bool:
    """Text, or a bare district number (a name in a header row)."""
    return bool(v) and (_is_text(v) or bool(_CODE_CELL.fullmatch(v.strip())))


def _rows(ctx: Context, table, opts, notes):
    """Plant / supply-point data: wide, tall, long or list."""
    matcher = ctx.row_matcher()
    labels, lm = ctx.labels, ctx.lm
    single = "p" not in ctx.axes
    info: dict = {}
    items: List[Item] = []

    # wide: a header row with months
    header, mcols = None, {}
    for r in range(min(10, len(table))):
        hits = _month_row(table[r], labels, lm)
        if len(hits) >= 2 or (len(hits) == 1 and len(labels) == 1):
            header, mcols = r, hits
            break
    if header is not None and not single:
        blocks = _blocks(mcols)
        bi = int(opts.get("block") or 0) if blocks else 0
        bi = min(max(bi, 0), len(blocks) - 1)
        cols = {c: mcols[c] for c in blocks[bi]}
        if len(blocks) > 1:
            info["blocks"] = [{"index": k, "label": f"columns {col_name(b[0])}\u2013{col_name(b[-1])}"
                               + (f" ({table[header - 1][b[0]]})" if header and table[header - 1][b[0]] else "")}
                              for k, b in enumerate(blocks)]
            notes.append(f"The paste has {len(blocks)} sets of month columns; using {info['blocks'][bi]['label']}. "
                         f"Pick another set below if that's the wrong one.")
        _year_note(table[header], cols, labels, lm, notes)
        label_cols = [c for c in range(len(table[0])) if c not in mcols]
        for r in range(header + 1, len(table)):
            row = table[r]
            texts = _label_cells(row, label_cols)
            if not texts or _SKIP_LABEL.match(texts[0]):
                continue
            vals = {p: parse_number(row[c]) for c, p in cols.items()}
            if all(k == TEXT for k, _ in vals.values()):
                continue
            items.append(Item(r + 1, matcher.best(texts), None, vals, " \u00b7 ".join(texts)))
        info["months_found"] = sorted({labels[p] for p in cols.values()}, key=labels.index)
        return "wide", items, info

    # tall: months down the first column, names across the top
    first = [r for r in range(len(table)) if period_of(parse_month(table[r][0]), labels, lm)[0] is not None]
    if len(first) >= 2 and not single:
        top = first[0]
        for c in range(1, len(table[0])):
            texts = [table[r][c] for r in range(top) if _nameish(table[r][c])]
            if not texts:
                continue
            vals = {}
            for r in first:
                p, _ = period_of(parse_month(table[r][0]), labels, lm)
                vals.setdefault(p, parse_number(table[r][c]))
            items.append(Item(c + 1, matcher.best(texts), None, vals, " \u00b7 ".join(texts)))
        info["months_found"] = sorted({labels[period_of(parse_month(table[r][0]), labels, lm)[0]] for r in first},
                                      key=labels.index)
        return "tall", items, info

    # long: a column of months
    hdr = 1 if table and all(not v or _is_text(v) for v in table[0]) and len(table) > 1 else 0
    data = table[hdr:]
    mcol = None
    for c in range(len(table[0])):
        n = sum(1 for row in data if period_of(parse_month(row[c]), labels, lm)[0] is not None)
        if n >= max(2, 0.6 * len(data)):
            mcol = c
            break
    num_cols = _value_cols(data, [c for c in range(len(table[0])) if c != mcol], table[0] if hdr else None)
    text_cols = [c for c in range(len(table[0])) if c != mcol and c not in num_cols]
    vc = _value_col(opts, num_cols, table[0] if hdr else None, info)
    if mcol is not None and vc is not None and not single:
        for r, row in enumerate(data, start=hdr + 1):
            texts = _label_cells(row, text_cols)
            p, _ = period_of(parse_month(row[mcol]), labels, lm)
            if not texts or p is None or _SKIP_LABEL.match(texts[0]):
                continue
            items.append(Item(r, matcher.best(texts), None, {p: parse_number(row[vc])}, " \u00b7 ".join(texts)))
        # one Item per (name, month): merge rows of the same name so names show once
        return "long", _merge(items), info

    # list: one value per name. A run of exactly P numeric columns = the months in order.
    if not single and len(num_cols) == len(labels) and not hdr:
        notes.append("No month headers found, but there are exactly as many number columns as months - "
                     f"read them as {labels[0]} to {labels[-1]} in order.")
        for r, row in enumerate(data, start=hdr + 1):
            texts = _label_cells(row, text_cols)
            if texts and not _SKIP_LABEL.match(texts[0]):
                items.append(Item(r, matcher.best(texts), None,
                                  {p: parse_number(row[c]) for p, c in enumerate(num_cols)}, " \u00b7 ".join(texts)))
        return "wide (by position)", items, info
    if vc is None:
        raise ValueError("couldn't find any numbers next to the names - copy the names and their values together")
    for r, row in enumerate(data, start=hdr + 1):
        texts = _label_cells(row, text_cols)
        if texts and not _SKIP_LABEL.match(texts[0]):
            items.append(Item(r, matcher.best(texts), None, {-1: parse_number(row[vc])}, " \u00b7 ".join(texts)))
    if not single:
        info["needs_months"] = True
    return "list", items, info


def _merge(items: List[Item]) -> List[Item]:
    out: Dict[str, Item] = {}
    for it in items:
        if it.a_text in out:
            out[it.a_text].values.update(it.values)
        else:
            out[it.a_text] = it
    return list(out.values())


def _value_col(opts, num_cols, header, info):
    if not num_cols:
        return None
    info["value_cols"] = [{"col": c, "label": (header[c] if header and header[c] else f"column {col_name(c)}")}
                          for c in num_cols]
    want = opts.get("value_col")
    if want not in (None, "") and int(want) in num_cols:
        return int(want)
    if header:
        for c in num_cols:
            if re.search(r"rate|freight|total|value|price|volume|demand|amount|cost|diff", header[c] or "", re.I):
                return c
    return num_cols[-1]


def _year_note(header_row, cols, labels, lm, notes):
    diff = [header_row[c] for c in cols if period_of(parse_month(header_row[c]), labels, lm)[1]]
    if diff:
        notes.append(f"The pasted months are for a different year ({diff[0]}, ...) than this model "
                     f"({labels[0]} to {labels[-1]}); matched by month.")


def _lanes(ctx: Context, table, opts, notes):
    """Lane data (freight, min/max shipment): a supply x plant matrix, or a list of lanes."""
    labels, lm = ctx.labels, ctx.lm
    info: dict = {}
    items: List[Item] = []
    fixed_from = opts.get("from") or ""
    fixed_to = opts.get("to") or ""

    # matrix: plants across a header row (or supply points across, plants down)
    hr, hits = _header_names(table, ctx.plants)
    if hr is not None:
        label_cols = [c for c in range(len(table[0])) if c not in hits]
        for r in range(hr + 1, len(table)):
            row = table[r]
            texts = _label_cells(row, label_cols)
            if not texts and fixed_from:
                texts = []
            elif not texts or _SKIP_LABEL.match(texts[0]):
                continue
            ma = _fixed(ctx.supplies, fixed_from) if fixed_from else ctx.supplies.best(texts)
            for c, mb in hits.items():
                v = parse_number(row[c])
                if v[0] == TEXT:
                    continue
                items.append(Item(r + 1, ma, mb, {-1: v}, " \u00b7 ".join(texts) if texts else "(chosen supply point)",
                                  table[hr][c]))
        info["needs_months"] = True
        return "matrix (supply points down, plants across)", items, info
    hr, hits = _header_names(table, ctx.supplies)
    if hr is not None:
        label_cols = [c for c in range(len(table[0])) if c not in hits]
        for r in range(hr + 1, len(table)):
            row = table[r]
            texts = _label_cells(row, label_cols)
            if not texts or _SKIP_LABEL.match(texts[0]):
                continue
            mb = ctx.plants.best(texts)
            for c, ma in hits.items():
                v = parse_number(row[c])
                if v[0] == TEXT:
                    continue
                items.append(Item(r + 1, ma, mb, {-1: v}, table[hr][c], " \u00b7 ".join(texts)))
        info["needs_months"] = True
        return "matrix (plants down, supply points across)", items, info

    # list of lanes
    header, mcols = None, {}
    for r in range(min(10, len(table))):
        h = _month_row(table[r], labels, lm)
        if len(h) >= 2:
            header, mcols = r, h
            break
    hdr = header + 1 if header is not None else \
        (1 if all(not v or _is_text(v) for v in table[0]) and len(table) > 1 else 0)
    data = table[hdr:]
    width = len(table[0])
    mcol = None
    if header is None:
        for c in range(width):
            n = sum(1 for row in data if period_of(parse_month(row[c]), labels, lm)[0] is not None)
            if n >= max(2, 0.6 * len(data)):
                mcol = c
                break
    num_cols = _value_cols(data, [c for c in range(width) if c not in mcols and c != mcol],
                           table[hdr - 1] if hdr else None)
    text_cols = [c for c in range(width) if c not in mcols and c != mcol and c not in num_cols]
    # numeric-looking district numbers ("0101") are names, not values
    for c in list(num_cols):
        vals = [row[c] for row in data if row[c]]
        if vals and all(re.fullmatch(r"0\d{3}|\d{4}", v) for v in vals) and \
                sum(1 for v in vals if ctx.plants.by_code.get(v.zfill(4))) >= 0.6 * len(vals):
            num_cols.remove(c)
            text_cols.append(c)
            text_cols.sort()
    if not text_cols and not (fixed_from and fixed_to):
        raise ValueError("couldn't find the supply point and plant names - each row needs where it's from and where it goes")
    sample = data[:300]

    def avg(c, m):
        sc = [m.best([row[c]])["score"] for row in sample if row[c]]
        return sum(sc) / len(sc) if sc else 0.0
    to_col = None
    if not fixed_to:
        to_col = max(text_cols, key=lambda c: avg(c, ctx.plants))
        if avg(to_col, ctx.plants) < 0.5:
            raise ValueError("couldn't tell which column is the plant - check the names, or paste a header row")
    from_cols = [c for c in text_cols if c != to_col]
    if not fixed_from and (not from_cols or max(avg(c, ctx.supplies) for c in from_cols) < 0.45):
        info["needs_from"] = True
        notes.append("No supply point column found - pick which supply point these rows are from.")
        if not from_cols:
            from_cols = []
    info["columns"] = {"from": [col_name(c) for c in from_cols], "to": col_name(to_col) if to_col is not None else ""}
    vc = None if mcols else _value_col(opts, num_cols, table[hdr - 1] if hdr else None, info)
    if mcols:
        blocks = _blocks(mcols)
        bi = min(max(int(opts.get("block") or 0), 0), len(blocks) - 1)
        cols = {c: mcols[c] for c in blocks[bi]}
        if len(blocks) > 1:
            info["blocks"] = [{"index": k, "label": f"columns {col_name(b[0])}\u2013{col_name(b[-1])}"} for k, b in enumerate(blocks)]
        _year_note(table[header], cols, labels, lm, notes)
    for r, row in enumerate(data, start=hdr + 1):
        ft = _label_cells(row, from_cols)
        tt = [row[to_col]] if to_col is not None and row[to_col] else []
        if (not ft and not fixed_from) or (not tt and not fixed_to):
            continue
        if ft and _SKIP_LABEL.match(ft[0]):
            continue
        ma = _fixed(ctx.supplies, fixed_from) if fixed_from else ctx.supplies.best(ft)
        mb = _fixed(ctx.plants, fixed_to) if fixed_to else ctx.plants.best(tt)
        if mcols:
            vals = {p: parse_number(row[c]) for c, p in cols.items()}
        elif mcol is not None:
            p, _ = period_of(parse_month(row[mcol]), labels, lm)
            if p is None or vc is None:
                continue
            vals = {p: parse_number(row[vc])}
        else:
            if vc is None:
                continue
            vals = {-1: parse_number(row[vc])}
            info["needs_months"] = True
        items.append(Item(r, ma, mb, vals, " \u00b7 ".join(ft) if ft else "(chosen supply point)",
                          " \u00b7 ".join(tt) if tt else "(chosen plant)"))
    if mcol is not None:
        merged: Dict[Tuple[str, str], Item] = {}
        for it in items:
            k = (it.a_text, it.b_text)
            if k in merged:
                merged[k].values.update(it.values)
            else:
                merged[k] = it
        items = list(merged.values())
    return "list of lanes", items, info


def _fixed(m: Matcher, key: str) -> dict:
    e = m.by_key.get(key)
    return {"status": "matched" if e else "unmatched", "score": 1.0 if e else 0.0, "entity": e,
            "options": [key] if e else []}


# --------------------------------------------------------------------------- #
# A template: what's in the model now, laid out the way analyse() reads it.
# --------------------------------------------------------------------------- #

def template(kind: str, d) -> str:
    fld, ent, unit, _ = KINDS[kind]
    axes = model.CELL_EDIT_FIELDS[fld]
    arr = np.asarray(d.freight_as_entered() if fld == "trans_cost" else getattr(d, fld), dtype=float)

    def fmt(v):
        if not np.isfinite(v):
            return ""
        if fld == "ship_max" and v >= model.DEFAULT_MAX_SHIP:
            return ""
        if unit == "volume":
            return f"{v * GAL:.12g}"
        if unit == "price":
            return f"{v / GAL:.12g}"
        return f"{v * 100:.12g}%"
    sup = [(i, n, (g or "").replace(model.BUILDER_TAG, "").strip()) for i, (n, g) in
           enumerate(zip(d.supply_names, d.supply_groups)) if not _PLACEHOLDER.match(n.strip())]
    pl = [(j, n) for j, n in enumerate(d.demand_names) if not _PLACEHOLDER.match(n.strip())]
    lines = []
    if ent == "lane":
        lines.append("\t".join(["Supply point", "Supplier", "Plant"] + list(d.period_labels)))
        for i, n, g in sup:
            for j, pn in pl:
                lines.append("\t".join([n, g, pn] + [fmt(arr[p, i, j]) for p in range(d.n_periods)]))
    elif ent == "plant":
        if "p" in axes:
            lines.append("\t".join(["Plant"] + list(d.period_labels)))
            lines += ["\t".join([n] + [fmt(arr[p, j]) for p in range(d.n_periods)]) for j, n in pl]
        else:
            lines.append("Plant\tValue")
            lines += [f"{n}\t{fmt(arr[j])}" for j, n in pl]
    else:
        if "p" in axes:
            lines.append("\t".join(["Supply point", "Supplier"] + list(d.period_labels)))
            lines += ["\t".join([n, g] + [fmt(arr[p, i]) for p in range(d.n_periods)]) for i, n, g in sup]
        else:
            lines.append("Supply point\tSupplier\tValue")
            lines += [f"{n}\t{g}\t{fmt(arr[i])}" for i, n, g in sup]
    return "\n".join(lines) + "\n"
