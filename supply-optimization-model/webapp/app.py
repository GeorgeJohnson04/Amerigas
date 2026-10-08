"""
app.py - local web app for the propane supply model.

WHAT IT DOES
------------
A browser front end over model.py.

  BASE MODELS  One locked set of area models per contract year (bases.py).
               A base is the raw model inputs - contracts, hub prices,
               differentials, lift limits, demand, freight, lane limits and
               storage - saved as JSON. Once made it can't be changed. The
               first bases are read from the Excel workbooks; after that,
               next year's base is made in the app from a base (or a working
               model) plus the new year's demand file - no Excel needed.

  WORKING MODELS ("scenarios")  Each manager's own copy on top of a base. It
               stores only that manager's changes (cells, model builder,
               demand % by state, an uploaded workbook) and is editable only
               by its owner, so nobody overwrites anybody else by accident.
               Others can look at it and make their own copy.

People pick an area, edit the inputs in spreadsheet-like grids, run the
optimisation, and read the results, insights and comparisons on screen.

HOW TO RUN
----------
    py -m pip install flask==3.1.3
    py webapp\\app.py                 (or double-click "Run Web App.bat")
    then open http://127.0.0.1:8050

SECURITY - READ BEFORE SHARING
------------------------------
This is built for local testing. It listens on 127.0.0.1 only, so nothing on
the network can reach it. There is NO login: the "your name" box decides whose
working models you see and may edit, which stops ACCIDENTAL overwrites but is
not security - anyone can type any name. Before hosting it for other people it
needs real sign-in (e.g. company SSO, which would replace the name box),
HTTPS, and a production WSGI server instead of Flask's development server.
Two cheap protections are in place for local use: requests must carry a Host
header of localhost/127.0.0.1 (blocks DNS rebinding), and every request that
changes data must carry an X-Requested-With header that a foreign web page
cannot set (blocks cross-site requests from other tabs).
"""

from __future__ import annotations

import dataclasses
import datetime as _dt
import glob
import json
import logging
import os
import re
import sys
import tempfile
import threading
import time
import uuid
from typing import Dict, List, Optional, Tuple

HERE = os.path.dirname(os.path.abspath(__file__))
TEST_DIR = os.path.dirname(HERE)
sys.path.insert(0, TEST_DIR)

import numpy as np  # noqa: E402

try:
    from flask import Flask, abort, jsonify, request, send_file  # noqa: E402
except ImportError:  # pragma: no cover
    raise SystemExit("Flask is required:  py -m pip install flask==3.1.3")

import model  # noqa: E402

sys.path.insert(0, HERE)
import names  # noqa: E402  (friendly display names - display only)
import bases as B  # noqa: E402
import paste as PASTE  # noqa: E402

LOG = logging.getLogger("webapp")

HOST = "127.0.0.1"
PORT = int(os.environ.get("MODEL_WEB_PORT", "8050"))
DATA_DIR = os.environ.get("MODEL_WEB_DATA") or os.path.join(HERE, "data")
SCEN_DIR = os.path.join(DATA_DIR, "scenarios")
DELETED_DIR = os.path.join(SCEN_DIR, "_deleted")
BASE_DIR = os.path.join(DATA_DIR, "bases")
DEMAND_DIR = os.path.join(DATA_DIR, "demand")
EXPORT_DIR = os.path.join(DATA_DIR, "exports")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
UPLOAD_EXTS = (".xlsm", ".xlsx")
UPLOAD_MAX_MB = 40
CSRF_HEADER = "X-Requested-With"
CSRF_VALUE = "propane-model"
BASE_PREFIX = "base:"
_ID_RE = re.compile(r"^[a-z0-9][a-z0-9-]{0,63}$")

app = Flask(__name__, static_folder=os.path.join(HERE, "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = UPLOAD_MAX_MB * 1024 * 1024


# --------------------------------------------------------------------------- #
# Small helpers.
# --------------------------------------------------------------------------- #

def now_iso() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")


def jarr(a, nd: int = 4):
    """numpy -> nested lists, rounded, with NaN/inf as null (JSON has no NaN)."""
    a = np.asarray(a, dtype=float)
    out = np.round(a, nd).astype(object)
    out[~np.isfinite(a)] = None
    return out.tolist()


def fnum(x) -> Optional[float]:
    x = float(x)
    return round(x, 4) if np.isfinite(x) else None


def fold(s: str) -> str:
    return model.fold_name(s or "")


def error(status: int, message: str):
    resp = jsonify({"error": message})
    resp.status_code = status
    return resp


def _atomic_write_json(path: str, obj) -> None:
    """Write JSON via a temp file + rename, retrying if a file-sync client holds a lock."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, indent=1)
    for attempt in range(10):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            time.sleep(0.1 * (attempt + 1))
    os.remove(tmp)
    raise RuntimeError(f"could not save {os.path.basename(path)} - is it open elsewhere?")


def area_key(region: str, area: str) -> str:
    return f"{region}|{area}"


def cell_key(field: str, index) -> str:
    return f"{field}|{','.join(str(int(k)) for k in index)}"


# --------------------------------------------------------------------------- #
# Request guards (local-only safety nets; see SECURITY in the module docstring).
# --------------------------------------------------------------------------- #

@app.before_request
def _guards():
    host = (request.host or "").split(":")[0].lower()
    if host not in ("127.0.0.1", "localhost"):
        abort(403)
    if request.method in ("POST", "PUT", "PATCH", "DELETE"):
        if request.headers.get(CSRF_HEADER) != CSRF_VALUE:
            abort(403)


@app.after_request
def _no_cache(resp):
    if request.path.startswith("/api/") or request.path.startswith("/static/") or request.path == "/":
        resp.headers["Cache-Control"] = "no-store"
    return resp


# --------------------------------------------------------------------------- #
# "Is this server running today's code?" (see open_browser.py)
# --------------------------------------------------------------------------- #

_CODE_FILES = (os.path.join(TEST_DIR, "model.py"), os.path.join(HERE, "app.py"),
               os.path.join(HERE, "names.py"), os.path.join(HERE, "bases.py"),
               os.path.join(HERE, "paste.py"))


def _code_fingerprint() -> str:
    parts = []
    for f in _CODE_FILES:
        try:
            st = os.stat(f)
            parts.append(f"{os.path.basename(f)}:{st.st_size}:{int(st.st_mtime)}")
        except OSError:
            parts.append(f"{os.path.basename(f)}:missing")
    return "|".join(parts)


STARTED_FINGERPRINT = _code_fingerprint()
STARTED_AT = now_iso()


@app.get("/api/health")
def api_health():
    stale = _code_fingerprint() != STARTED_FINGERPRINT
    return jsonify({"ok": True, "started": STARTED_AT, "stale": stale, "pid": os.getpid()})


@app.post("/api/shutdown")
def api_shutdown():
    """Let the launcher stop an out-of-date copy before starting a fresh one (local only)."""
    LOG.info("shutdown requested - exiting so a fresh copy can start")
    threading.Timer(0.3, lambda: os._exit(0)).start()
    return jsonify({"ok": True})


# --------------------------------------------------------------------------- #
# Workbooks: only for importing bases and reading uploaded copies.
# --------------------------------------------------------------------------- #

class Workbooks:
    """Finds the area workbooks in the Model folder and caches what's read from them."""

    def __init__(self, root: str):
        self.root = root
        self._lock = threading.Lock()
        self._raw: Dict[str, Tuple[int, dict]] = {}

    def raw(self, path: str, area: str, region: str) -> dict:
        """The raw inputs of a workbook file (cached until the file changes)."""
        mtime = os.stat(path).st_mtime_ns
        with self._lock:
            hit = self._raw.get(path)
            if hit and hit[0] == mtime:
                return hit[1]
        t0 = time.time()
        reader = model.ModelBaseReader(path, area, region)
        raw = reader.raw_inputs(*reader.read_parts())
        states, _ = B.states_for(raw, None)
        raw["demand_states"] = states
        with self._lock:
            self._raw[path] = (mtime, raw)
        LOG.info("read %s in %.1fs", os.path.basename(path), time.time() - t0)
        return raw


def guess_area(filename: str, candidates: List[Tuple[str, str]]) -> Optional[Tuple[str, str]]:
    """'MT-WY 26-27 Working (2).xlsm' -> ('Rockies', 'MT-WY 26-27') among *candidates*."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    stem = re.sub(r"\s*\(\d+\)\s*$", "", stem)
    stem = re.sub(r"(?i)^copy of\s+", "", stem)
    try:
        area = model._classify(stem + ".xlsm")[0]
    except Exception:  # noqa: BLE001
        area = stem
    hits = [c for c in candidates if fold(c[1]) == fold(area)]
    if len(hits) == 1:
        return hits[0]
    hits = [c for c in candidates if fold(c[1]) and fold(c[1]) in fold(stem)]
    if len(hits) == 1:
        return hits[0]
    code = fold(re.sub(r"\s*\d{2}-\d{2}\s*$", "", area))           # ignore the season
    hits = [c for c in candidates if fold(re.sub(r"\s*\d{2}-\d{2}\s*$", "", c[1])) == code]
    return hits[0] if len(hits) == 1 else None


def find_demand_files() -> List[str]:
    """Demand workbooks the app knows about, newest first."""
    pats = [os.path.join(d, "Demand for Models*.xls*") for d in
            (TEST_DIR, DEMAND_DIR, os.path.dirname(TEST_DIR))]
    files = {os.path.abspath(f) for p in pats for f in glob.glob(p)
             if not os.path.basename(f).startswith("~$")}
    return sorted(files, key=os.path.getmtime, reverse=True)


_DEMAND_CACHE: Dict[str, Tuple[float, B.DemandFile]] = {}


def demand_file(path: str) -> B.DemandFile:
    mt = os.path.getmtime(path)
    hit = _DEMAND_CACHE.get(path)
    if hit and hit[0] == mt:
        return hit[1]
    df = B.DemandFile(path)
    _DEMAND_CACHE[path] = (mt, df)
    return df


# --------------------------------------------------------------------------- #
# Working models (scenarios): one JSON file each, owned by one person.
# --------------------------------------------------------------------------- #

class Scenarios:
    """
    One JSON file per working model::

        {"id", "name", "base", "owner", "locked", "created", "updated", "version",
         "edits":     {"<region>|<area>": {"<field>|<i>,<j>": {"value", "user", "ts"}}},
         "structure": {"<region>|<area>": {model builder settings}},
         "sources":   {"<region>|<area>": {uploaded workbook}},
         "state_pct": {"<region>|<area>": {"<state>": [12 fractions, Apr..Mar]}}}

    "base:<base id>" is a read-only stand-in for the base model itself, so
    everything that reads a scenario also works on a base.
    """

    def __init__(self, folder: str, bases: B.Bases):
        self.folder = folder
        self.bases = bases
        self._lock = threading.RLock()
        os.makedirs(folder, exist_ok=True)

    def _path(self, sid: str) -> str:
        if not _ID_RE.match(sid or ""):
            abort(400, description="bad working model id")
        return os.path.join(self.folder, sid + ".json")

    def _legacy_base(self) -> str:
        """Base for working models made before bases existed (they were on the Working files)."""
        allb = self.bases.list()
        hit = next((b for b in allb if b.get("recipe", {}).get("variant") == "Working"), None)
        hit = hit or self.bases.default()
        return hit["id"] if hit else ""

    def _normalise(self, s: dict) -> dict:
        if not s.get("base"):
            s["base"] = self._legacy_base()
        if "owner" not in s:
            cb = s.get("created_by", "")
            s["owner"] = "" if cb in ("", "system") else cb
        for k in ("edits", "structure", "sources", "state_pct"):
            s.setdefault(k, {})
        s.setdefault("locked", False)
        s["is_base"] = False
        return s

    def get(self, sid: str) -> dict:
        if (sid or "").startswith(BASE_PREFIX):
            meta = self.bases.get(sid[len(BASE_PREFIX):])
            if not meta:
                abort(404, description="base model not found")
            return {"id": sid, "name": meta["name"], "base": meta["id"], "owner": "", "locked": True,
                    "is_base": True, "version": 0, "updated": meta.get("created"),
                    "edits": {}, "structure": {}, "sources": {}, "state_pct": {}}
        path = self._path(sid)
        with self._lock:
            if not os.path.isfile(path):
                abort(404, description="working model not found")
            with open(path, encoding="utf-8") as fh:
                return self._normalise(json.load(fh))

    def _write(self, s: dict) -> None:
        s = {k: v for k, v in s.items() if k != "is_base"}
        _atomic_write_json(self._path(s["id"]), s)

    def list(self) -> List[dict]:
        out = []
        with self._lock:
            for fn in sorted(os.listdir(self.folder)):
                if not fn.endswith(".json"):
                    continue
                try:
                    with open(os.path.join(self.folder, fn), encoding="utf-8") as fh:
                        s = self._normalise(json.load(fh))
                except (OSError, ValueError, KeyError):
                    continue
                out.append({
                    "id": s["id"], "name": s["name"], "base": s["base"], "owner": s["owner"],
                    "locked": bool(s.get("locked")), "updated": s.get("updated"),
                    "version": s.get("version", 0),
                    "edits": sum(len(v) for v in s["edits"].values()),
                    "areas": sorted({*s["edits"], *s["structure"], *s["sources"], *s["state_pct"]}),
                })
        return sorted(out, key=lambda s: (s["updated"] or ""), reverse=True)

    def create(self, name: str, user: str, base: str, copy_from: str = "") -> dict:
        name = (name or "").strip()[:80] or "Untitled working model"
        slug = model.slugify(name).lower().strip("_").replace("_", "-")[:40] or "model"
        sid = f"{slug}-{uuid.uuid4().hex[:6]}"
        s = {"id": sid, "name": name, "base": base, "owner": user, "locked": False,
             "created": now_iso(), "created_by": user, "updated": now_iso(), "version": 1,
             "edits": {}, "structure": {}, "sources": {}, "state_pct": {}}
        if copy_from:
            src = self.get(copy_from)
            s["base"] = src["base"]
            for k in ("edits", "structure", "sources", "state_pct"):
                s[k] = json.loads(json.dumps(src.get(k, {})))
            s["copied_from"] = src["name"]
        if not self.bases.get(s["base"]):
            abort(400, description="pick a base model first")
        with self._lock:
            self._write(s)
        return self._normalise(s)

    def update(self, sid: str, fn) -> dict:
        """Load, change with fn(s), bump the version, save. Returns the new scenario."""
        with self._lock:
            s = self.get(sid)
            fn(s)
            s["updated"] = now_iso()
            s["version"] = s.get("version", 0) + 1
            self._write(s)
            return s

    def delete(self, sid: str) -> None:
        with self._lock:
            path = self._path(sid)
            if not os.path.isfile(path):
                abort(404, description="working model not found")
            os.makedirs(DELETED_DIR, exist_ok=True)
            os.replace(path, os.path.join(DELETED_DIR, f"{sid}.{int(time.time())}.json"))

    @staticmethod
    def structure_for(s: dict, region: str, area: str) -> dict:
        return dict(s.get("structure", {}).get(area_key(region, area), {}))

    @staticmethod
    def edits_for(s: dict, region: str, area: str) -> List[dict]:
        cells = s.get("edits", {}).get(area_key(region, area), {})
        out = []
        for key, rec in cells.items():
            fld, idx = key.split("|", 1)
            out.append({"field": fld, "index": [int(k) for k in idx.split(",")],
                        "value": rec["value"], "user": rec.get("user", ""), "ts": rec.get("ts", ""),
                        "via": rec.get("via", "")})
        return out


def can_edit(s: dict, user: str) -> Tuple[bool, str]:
    """Whether *user* may change *s*, and if not, why (shown to them as is)."""
    if s.get("is_base"):
        return False, ("Base models are locked. Create a working model from this base "
                       "to make changes.")
    if s.get("locked"):
        return False, (f"'{s['name']}' is locked. "
                       + ("Unlock it from the working-model menu to change it." if fold(s["owner"]) == fold(user)
                          else "Make your own copy to change it."))
    if not s.get("owner"):
        return False, "This older shared working model is read-only. Make your own copy to change it."
    if not user.strip():
        return False, "Add your name at the top right first."
    if fold(s["owner"]) != fold(user):
        return False, (f"This working model belongs to {s['owner']}. Make your own copy to change it - "
                       f"theirs stays exactly as it is.")
    return True, ""


def require_edit(s: dict, user: str) -> None:
    ok, why = can_edit(s, user)
    if not ok:
        abort(403, description=why)


# --------------------------------------------------------------------------- #
# Building a model for an area of a working model.
# --------------------------------------------------------------------------- #

def area_raw(s: dict, region: str, area: str) -> Tuple[dict, str, str]:
    """(raw inputs, what it's based on, source path) for an area of *s*."""
    up = s.get("sources", {}).get(area_key(region, area))
    if up:
        path = os.path.join(UPLOAD_DIR, up["file"])
        if os.path.isfile(path):
            return WB.raw(path, area, region), f"{up['name']} (uploaded)", path
    raw = BASES.area_raw(s["base"], region, area)
    if raw is None:
        abort(404, description=f"{area} isn't in this working model's base")
    meta = BASES.get(s["base"]) or {}
    return raw, meta.get("name", "base"), f"{BASE_PREFIX}{s['base']}"


def base_pct_for(raw: dict, s: dict, region: str, area: str) -> Dict[str, List[float]]:
    """% of demand by state that the base demand was built at (an upload uses its base's)."""
    pct = dict(raw.get("state_pct") or {})
    if not pct and not s.get("is_base"):
        base = BASES.area_raw(s["base"], region, area) or {}
        pct = dict(base.get("state_pct") or {})
    return pct


def demand_scale_for(s: dict, region: str, area: str, raw: dict) -> Tuple[Optional[np.ndarray], str]:
    """(P, M) multipliers from the working model's "% of demand by state", or None."""
    want = s.get("state_pct", {}).get(area_key(region, area))
    if not want:
        return None, ""
    states = raw.get("demand_states") or []
    have = base_pct_for(raw, s, region, area)
    labels = raw["period_labels"]
    P, M = len(labels), len(raw["demand_names"])
    scale = np.ones((P, M))
    for j in range(min(M, len(states))):
        st = states[j]
        if st not in want:
            continue
        for p, lab in enumerate(labels):
            k = B.month_slot(lab, p)
            old = (have.get(st) or [1.0] * 12)[k] or 1.0
            scale[p, j] = float(want[st][k]) / float(old)
    return scale, "demand % by state"


def build(s: dict, region: str, area: str, edits: Optional[List[dict]] = None,
          structure: Optional[dict] = None, scale: bool = True) -> model.ModelData:
    raw, _, path = area_raw(s, region, area)
    edits = Scenarios.edits_for(s, region, area) if edits is None else edits
    structure = Scenarios.structure_for(s, region, area) if structure is None else structure
    ds, note = demand_scale_for(s, region, area, raw) if scale else (None, "")
    reader = model.ModelBaseReader(path, area, region, cell_edits=edits, structure=structure or None,
                                   demand_scale=ds, demand_scale_note=note)
    return reader.build_from_inputs(raw)


_BASELINE: Dict[Tuple[str, str, str], dict] = {}


def baseline_result(bid: str, region: str, area: str) -> Optional[dict]:
    """The base model's own result for an area (solved once, then remembered)."""
    key = (bid, region, area)
    if key not in _BASELINE:
        if BASES.area_raw(bid, region, area) is None:
            return None
        sol = model.solve_area(build(SC.get(BASE_PREFIX + bid), region, area))
        _BASELINE[key] = summary_of(sol)
    return _BASELINE[key]


def has_changes(s: dict, region: str, area: str) -> bool:
    ak = area_key(region, area)
    return any(s.get(k, {}).get(ak) for k in ("edits", "structure", "sources", "state_pct"))


# --------------------------------------------------------------------------- #
# Turning model objects into JSON.
# --------------------------------------------------------------------------- #

FIELD_INFO = {
    # field: (label, unit, group)
    "contract": ("Contract volume", "kgal", "supply"),
    "hub_price": ("Hub price", "$/kgal", "supply"),
    "hub_premium": ("Differential to hub", "$/kgal", "supply"),
    "price": ("Delivered price (hub + diff)", "$/kgal", "supply"),
    "perc_min": ("Minimum lift", "x contract", "supply"),
    "perc_max": ("Maximum lift", "x contract", "supply"),
    "supply_inv_init": ("Initial inventory", "kgal", "supply"),
    "supply_inv_cap": ("Storage capacity", "kgal", "supply"),
    "supply_inv_min": ("Minimum inventory", "kgal", "supply"),
    "demand": ("Demand", "kgal", "demand"),
    "demand_inv_init": ("Initial inventory", "kgal", "demand"),
    "demand_inv_cap": ("Storage capacity", "kgal", "demand"),
    "demand_inv_min": ("Minimum inventory", "kgal", "demand"),
    "trans_cost": ("Freight", "$/kgal", "lane"),
    "ship_min": ("Minimum shipment", "kgal", "lane"),
    "ship_max": ("Maximum shipment", "kgal", "lane"),
}


def input_values(d: model.ModelData, f: str) -> np.ndarray:
    """An input as the planner entered it: a closed lane's freight is blank (NaN), not 0."""
    return d.freight_as_entered() if f == "trans_cost" else np.asarray(getattr(d, f), float)


def scenario_info(s: dict, user: str) -> dict:
    ok, why = can_edit(s, user)
    meta = BASES.get(s["base"]) or {}
    return {"id": s["id"], "name": s["name"], "version": s.get("version", 0), "owner": s.get("owner", ""),
            "locked": bool(s.get("locked")), "is_base": bool(s.get("is_base")), "can_edit": ok,
            "why_not": why, "base": {"id": meta.get("id", ""), "name": meta.get("name", ""),
                                     "season": meta.get("season", "")}}


def demand_by_state(d: model.ModelData, states: List[str]) -> Dict[str, List[float]]:
    out: Dict[str, List[float]] = {}
    for j, st in enumerate(states[:d.n_demand]):
        if not st:
            continue
        row = out.setdefault(st, [0.0] * 12)
        for p, lab in enumerate(d.period_labels):
            row[B.month_slot(lab, p)] += float(d.demand[p, j])
    return out


def area_payload(s: dict, region: str, area: str, user: str) -> dict:
    raw, based_on, _ = area_raw(s, region, area)
    edits = Scenarios.edits_for(s, region, area)
    structure = Scenarios.structure_for(s, region, area)
    plain = build(s, region, area, edits=[], structure={}, scale=False)   # the base / upload as it is
    base = build(s, region, area, edits=[], structure=structure)           # before cell edits
    cur = build(s, region, area, edits=edits, structure=structure)
    base_dem = build(s, region, area, edits=[], structure=structure, scale=False)   # no demand changes at all
    up = s.get("sources", {}).get(area_key(region, area))
    upload = None
    if up:
        official = SC.get(BASE_PREFIX + s["base"])
        upload = {**up, "compare": compare_models(build(official, region, area, scale=False), plain)}
    edit_list = []
    base_arrays: Dict[str, np.ndarray] = {}         # once per field - a paste can make thousands of edits
    for e in edits:
        if e["field"] not in base_arrays:
            base_arrays[e["field"]] = input_values(base, e["field"])
        b = base_arrays[e["field"]][tuple(e["index"])]
        edit_list.append({**e, "base": fnum(b)})
    states = list(raw.get("demand_states") or [])
    base_pct = base_pct_for(raw, s, region, area)
    want = s.get("state_pct", {}).get(area_key(region, area)) or {}
    st_list = sorted({x for x in states if x})
    return {
        "region": region, "area": area,
        "source": based_on,
        "display": names.area_display(area), "region_display": names.region_display(region),
        "totals": {"demand": fnum(cur.total_demand), "contract": fnum(float(cur.contract.sum()))},
        "scenario": scenario_info(s, user),
        "periods": cur.period_labels,
        "supply": [{"name": n, "group": g, "hub": h}
                   for n, g, h in zip(cur.supply_names, cur.supply_groups, cur.supply_hubs)],
        "demand": cur.demand_names,
        "field_info": {k: {"label": v[0], "unit": v[1], "group": v[2],
                           "editable": k in model.CELL_EDIT_FIELDS}
                       for k, v in FIELD_INFO.items()},
        "fields": {f: jarr(input_values(cur, f)) for f in FIELD_INFO},
        "edits": edit_list,
        "high_rate": model.HIGH_FREIGHT_RATE,       # freight at/above this gets a warning
        "no_max": model.DEFAULT_MAX_SHIP,
        "workbook_size": {"periods": plain.n_periods, "supply": plain.n_supply, "demand": plain.n_demand,
                          "hubs": sorted({h for h in plain.supply_hubs if h})},
        "builder_hubs": model.builder_hub_choices(sorted({h for h in plain.supply_hubs if h})),
        "structure": {k: v for k, v in structure.items() if k not in ("user", "ts")},
        "structure_meta": {"user": structure.get("user", ""), "ts": structure.get("ts", "")},
        "builder_limits": {"periods": model.BUILDER_MAX_PERIODS, "added": model.BUILDER_MAX_ADDED},
        "demand_states": states,
        # Season demand by month: the base model's (before this working model's
        # demand % and typed-in changes) and now - for the Demand tab's chart.
        "demand_months": {"base": jarr(base_dem.demand.sum(axis=1)), "now": jarr(cur.demand.sum(axis=1))},
        "state_pct": {
            "months": B.month_names(), "states": st_list,
            "base": {st: base_pct.get(st) for st in st_list},
            "current": {st: want.get(st) or base_pct.get(st) for st in st_list},
            "changed": sorted(want),
            "base_gal": demand_by_state(plain, states),
            "current_gal": demand_by_state(cur, states),
        },
        "notes": [model.gal_text(t) for t in cur.notes],
        "repairs": [model.gal_text(t) for t in cur.repairs],
        "feasibility": [model.gal_text(t) for t in cur.feasibility_warnings()],
        "upload": upload,
    }


def compare_models(official: model.ModelData, other: model.ModelData) -> dict:
    """What an uploaded workbook changes relative to the base (gal / $ per gal, ready to show)."""
    out = {"size": [], "fields": []}
    for label, a, b in (("months", official.n_periods, other.n_periods),
                        ("supply points", official.n_supply, other.n_supply),
                        ("plants", official.n_demand, other.n_demand)):
        if a != b:
            out["size"].append(f"{label}: {a} in the base model, {b} in yours")
    for f in model.CELL_EDIT_FIELDS:
        a, b = input_values(official, f), input_values(other, f)
        if a.shape != b.shape:
            continue
        diff = ~np.isclose(a, b, rtol=1e-9, atol=1e-9, equal_nan=True)
        n = int(diff.sum())
        if not n:
            continue
        gap = np.where(np.isnan(a) ^ np.isnan(b), np.inf, np.abs(np.nan_to_num(b) - np.nan_to_num(a)))
        k = np.unravel_index(int(np.argmax(np.where(diff, gap, -1))), a.shape)
        scale = (model.GAL_PER_KGAL if f in model._VOLUME_FIELDS
                 else 1 / model.GAL_PER_KGAL if f in model._PRICE_FIELDS else 1.0)
        where = dict(zip(model.CELL_EDIT_FIELDS[f], (int(x) for x in k)))
        loc = []
        if "i" in where:
            loc.append(other.supply_names[where["i"]])
        if "j" in where:
            loc.append(other.demand_names[where["j"]])
        if "p" in where:
            loc.append(other.period_labels[where["p"]])
        out["fields"].append({"field": f, "label": FIELD_INFO[f][0], "cells": n,
                              "largest": {"where": " / ".join(loc), "official": fnum(a[k] * scale),
                                          "yours": fnum(b[k] * scale)}})
    if (official.supply_names != other.supply_names or official.demand_names != other.demand_names) \
            and not out["size"]:
        out["size"].append("some supply point or plant names differ")
    return out


def summary_of(sol: model.Solution) -> dict:
    d = sol.data
    out = {"status": sol.status, "ok": sol.ok, "relaxed": sol.relaxed,
           "message": model.gal_text(sol.message), "relax_note": model.gal_text(sol.relax_note),
           "solver": sol.solver, "seconds": round(sol.solve_seconds, 2),
           "demand": fnum(d.total_demand)}
    if sol.ok:
        comm = float(sol.commodity_cost().sum())
        frt = float(sol.freight_cost().sum())
        purch = float(sol.purchased().sum())
        out.update({
            "purchased": fnum(purch), "shipped": fnum(sol.ship.sum()),
            "unmet": fnum(sol.unmet_total), "surplus": fnum(sol.surplus_total),
            "other_slack": fnum(sol.other_slack_total),
            "blocked_volume": fnum(sol.blocked_lane_volume()),
            "blocked_cost": fnum(sol.blocked_lane_cost()),
            "commodity": fnum(comm), "freight": fnum(frt), "total": fnum(comm + frt),
            "avg": fnum((comm + frt) / purch) if purch > 1e-9 else None,
        })
    return out


def monthly_of(sol: model.Solution) -> Optional[dict]:
    if not sol.ok:
        return None
    return {"commodity": jarr(sol.commodity_cost().sum(axis=1), 2),
            "freight": jarr(sol.freight_cost().sum(axis=(1, 2)), 2),
            "purchased": jarr(sol.purchased().sum(axis=1)),
            "shipped": jarr(sol.ship.sum(axis=(1, 2))),
            "unmet": jarr(sol.shortfall.sum(axis=1) if sol.shortfall is not None
                          else np.zeros(sol.data.n_periods))}


def diagnostics_of(sol: model.Solution, based_on: str) -> List[dict]:
    d = sol.data
    rows = []

    def add(sev, text):
        rows.append({"severity": sev, "text": model.gal_text(text)})

    add("Info", f"Based on: {based_on}")
    add("Info", f"{d.n_supply} supply points, {d.n_demand} demand points, "
                f"{d.n_periods} periods, total demand {d.total_demand:,.1f} kgal.")
    if d.override_log:
        add("Info", f"{len(d.override_log)} change(s) from this working model applied - "
                    f"see the Changes tab.")
    closed = d.closed_lane_mask()
    if closed.any():
        add("Info", f"{int(closed.any(axis=0).sum())} lane(s) have a blank freight rate in at "
                    f"least one month and are not used in those months.")
    for r in d.repairs:
        add("Error", "DATA ERROR - " + r)
    for n in d.notes:
        add("Warning", n)
    for w in d.feasibility_warnings():
        add("Warning", w)
    if not sol.ok:
        add("Error", f"Did not solve: {sol.status}. {sol.message}")
    elif sol.relaxed:
        add("Warning", sol.relax_note)
        for v in sol.violations:
            add("Warning", v)
    if d.blocked_lane_pairs:
        add("Info", f"{d.blocked_lane_pairs} lane(s) have freight of $1/gal or more in at least one month. "
                    f"They are used like any other lane when they're the cheapest way; only a blank "
                    f"freight cell keeps a lane unused.")
    if sol.ok and sol.high_freight_volume() >= model.HIGH_FREIGHT_REPORT_MIN:
        add("Warning", f"High freight: {sol.high_freight_volume():,.1f} kgal moves on lanes with freight "
                       f"of $1/gal or more, costing ${sol.high_freight_cost():,.0f}. These are the cheapest "
                       f"routes the model found - check the rates are right (marked on Shipments and Lane volumes).")
    return rows


def insights_of(sol: model.Solution) -> dict:
    """'What is one more gallon worth' and how each contract is used."""
    mv = model.marginal_values(sol)
    for k in mv:
        for e in mv[k]:
            e["value"] = fnum(e["value"])
            e["best"] = fnum(e["best"])
            if "by_month" in e:
                e["by_month"] = [fnum(x) for x in e["by_month"]]
    use = model.contract_use(sol)
    for u in use:
        for k in ("bought", "contract", "share"):
            u[k] = fnum(u[k])
        u["lift"] = [fnum(x) for x in u["lift"]]
    return {"marginal": mv, "contract_use": use, "has_duals": bool(sol.duals)}


def result_payload(sol: model.Solution, baseline: Optional[dict], based_on: str) -> dict:
    d = sol.data
    out = {"area": d.area, "display": names.area_display(d.area),
           "summary": summary_of(sol), "baseline": baseline,
           "diagnostics": diagnostics_of(sol, based_on), "periods": d.period_labels,
           "supply": [{"name": n, "group": g, "hub": h}
                      for n, g, h in zip(d.supply_names, d.supply_groups, d.supply_hubs)],
           "demand": d.demand_names, "monthly": monthly_of(sol)}
    if not sol.ok:
        return out
    ship = sol.ship
    idx = np.argwhere(ship > 1e-6)
    blocked = d.blocked_lane_mask()
    # [month, supply, plant, kgal, $/kgal, freight $, high freight?] - volumes to
    # 1e-6 kgal (a thousandth of a gallon) so the lane tables can show exact gallons.
    out["shipments"] = [
        [int(p), int(i), int(j), round(float(ship[p, i, j]), 6), fnum(d.trans_cost[p, i, j]),
         round(float(ship[p, i, j] * d.trans_cost[p, i, j]), 2), bool(blocked[p, i, j])]
        for p, i, j in idx
    ]
    zeros_s = np.zeros((d.n_periods, d.n_supply))
    zeros_d = np.zeros((d.n_periods, d.n_demand))
    out["supply_plan"] = {
        "contract": jarr(d.contract), "perc": jarr(sol.perc),
        "purchased": jarr(sol.purchased()), "price": jarr(d.price),
        "cost": jarr(sol.commodity_cost()),
        "surplus": jarr(sol.surplus if sol.surplus is not None else zeros_s),
    }
    out["supply_inventory"] = {
        "init": jarr(d.supply_inv_init), "withdraw": jarr(sol.supply_wd),
        "inject": jarr(sol.supply_inj), "end": jarr(sol.supply_end),
        "min": jarr(d.supply_inv_min), "cap": jarr(d.supply_inv_cap),
    }
    out["demand_coverage"] = {
        "demand": jarr(d.demand), "delivered": jarr(sol.delivered()),
        "unmet": jarr(sol.shortfall if sol.shortfall is not None else zeros_d),
        "excess": jarr(sol.excess if sol.excess is not None else zeros_d),
        "withdraw": jarr(sol.demand_wd), "inject": jarr(sol.demand_inj),
        "end": jarr(sol.demand_end), "min": jarr(d.demand_inv_min),
        "cap": jarr(d.demand_inv_cap), "init": jarr(d.demand_inv_init),
    }
    out["insights"] = insights_of(sol)
    return out


# --------------------------------------------------------------------------- #
# Routes.
# --------------------------------------------------------------------------- #

WB: Workbooks
SC: Scenarios
BASES: B.Bases


def _body() -> dict:
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        abort(400, description="expected a JSON object")
    return data


def _user(b: Optional[dict] = None) -> str:
    v = (b or {}).get("user") if b is not None else None
    v = v if v is not None else request.args.get("user", "")
    return str(v or "").strip()[:60]


@app.errorhandler(400)
@app.errorhandler(403)
@app.errorhandler(404)
@app.errorhandler(409)
def _http_error(exc):
    return error(exc.code, getattr(exc, "description", str(exc)))


@app.errorhandler(model.ModelStructureError)
def _structure_error(exc):
    return error(422, f"This workbook can't be read by the model: {exc}")


@app.route("/")
def index():
    return send_file(os.path.join(HERE, "static", "index.html"))


# ---- base models ----------------------------------------------------------- #

def _official_path() -> str:
    return os.path.join(BASE_DIR, "official.json")


def official_base() -> Optional[dict]:
    try:
        with open(_official_path(), encoding="utf-8") as fh:
            bid = json.load(fh).get("id", "")
        meta = BASES.get(bid)
        if meta:
            return meta
    except (OSError, ValueError):
        pass
    return BASES.default()


def set_official(bid: str) -> None:
    _atomic_write_json(_official_path(), {"id": bid, "ts": now_iso()})


def base_summary(meta: dict) -> dict:
    off = official_base()
    return {k: meta.get(k) for k in ("id", "name", "season", "created", "created_by", "kind", "recipe")} | {
        "official": bool(off and off["id"] == meta["id"]), "locked": True,
        "areas": len(meta.get("areas", [])),
        "regions": sorted({a["region"] for a in meta.get("areas", [])}),
        "area_notes": [a for a in meta.get("areas", []) if a.get("demand_note") or a.get("demand_kept")],
    }


@app.get("/api/bases")
def api_bases():
    return jsonify({
        "bases": [base_summary(m) for m in BASES.list()],
        "official": (official_base() or {}).get("id", ""),
        "jobs": list(BASES.jobs.values()),
        "demand_files": [{"path": p, "name": os.path.basename(p),
                          "modified": _dt.datetime.fromtimestamp(os.path.getmtime(p)).isoformat(timespec="minutes")}
                         for p in find_demand_files()],
    })


@app.post("/api/bases")
def api_base_create():
    """
    Make a new (locked) base model set, in the background.

    {action: "import", variant: "Baseline"|"Working", name, user}
        read the area workbooks in the Model folder
    {action: "roll", source: "base:<id>" | "<working model id>", season: "2027-28",
     name, user, demand_file: path or "", freight_pct: number}
        next year's base from a base or a working model + a demand file
    """
    b = _body()
    user = _user(b)
    if not user:
        abort(400, description="add your name at the top right first")
    if BASES.running():
        abort(409, description="another base model is being made - wait for it to finish")
    action = b.get("action")
    df_path = str(b.get("demand_file") or "")
    if df_path and os.path.abspath(df_path) not in find_demand_files():
        abort(400, description="unknown demand file")
    name = str(b.get("name") or "").strip()[:80]
    if action == "import":
        variant = "Baseline" if str(b.get("variant")) == "Baseline" else "Working"

        def job(progress):
            df = demand_file(df_path) if df_path else None
            return BASES.import_workbooks(WB.root, variant, user, name=name, demand_file=df, progress=progress)
        jid = BASES.start_job(f"Import {variant} workbooks", job)
    elif action == "roll":
        src = SC.get(str(b.get("source") or ""))
        try:
            season = B.normalise_season(str(b.get("season") or ""))
            freight = float(b.get("freight_pct") or 0)
        except ValueError as exc:
            abort(400, description=str(exc))
        if not -50 <= freight <= 100:
            abort(400, description="freight change should be between -50% and +100%")
        meta = BASES.get(src["base"]) or {}
        areas = [(a["region"], a["area"]) for a in meta.get("areas", [])]
        label = f"{src['name']}" + ("" if src.get("is_base") else f" (working model of {src.get('owner') or 'shared'})")

        def job(progress):
            df = demand_file(df_path) if df_path else None
            sources = []
            for n, (region, area) in enumerate(areas):
                progress(n, len(areas) * 2, f"building {area}")
                sources.append((region, area, build(src, region, area)))
            return BASES.roll_forward(sources, season, user, name, label, demand_file=df,
                                      freight_pct=freight,
                                      progress=lambda d, t, s: progress(len(areas) + d, len(areas) * 2, s))
        jid = BASES.start_job(f"Roll forward to 20{season}", job)
    else:
        abort(400, description="unknown action")
    return jsonify({"job": jid})


@app.get("/api/bases/jobs/<jid>")
def api_base_job(jid):
    job = BASES.jobs.get(jid)
    if not job:
        abort(404, description="no such job")
    return jsonify(job)


@app.post("/api/bases/<bid>/official")
def api_base_official(bid):
    if not BASES.get(bid):
        abort(404, description="base model not found")
    set_official(bid)
    return jsonify({"ok": True})


@app.post("/api/demandfiles")
def api_demand_upload():
    f = request.files.get("file")
    if f is None or not f.filename.lower().endswith((".xlsx", ".xlsm")):
        abort(400, description="choose the demand workbook (.xlsx)")
    base = re.sub(r"[^\w\-. ]+", "_", os.path.basename(f.filename))
    if not base.lower().startswith("demand for models"):
        base = "Demand for Models - " + base
    os.makedirs(DEMAND_DIR, exist_ok=True)
    path = os.path.join(DEMAND_DIR, base)
    f.save(path)
    try:
        df = demand_file(path)
    except Exception as exc:  # noqa: BLE001
        os.remove(path)
        return error(422, f"couldn't read that demand workbook: {exc}")
    return jsonify({"path": os.path.abspath(path), "name": base, "plants": len(df.rows),
                    "states": len(df.state_pct)})


# ---- regions / working models ---------------------------------------------- #

@app.get("/api/regions")
def api_regions():
    bid = request.args.get("base") or (official_base() or {}).get("id", "")
    if request.args.get("refresh"):
        names.reload()
    meta = BASES.get(bid) if bid else None
    if not meta:
        return jsonify({"base": None, "regions": []})
    regions: Dict[str, list] = {}
    for a in meta.get("areas", []):
        regions.setdefault(a["region"], []).append(
            {"area": a["area"], "file": a.get("source", ""), "display": names.area_display(a["area"])})
    return jsonify({"base": base_summary(meta), "regions": [
        {"region": r, "display": names.region_display(r), "areas": regions[r]}
        for r in sorted(regions, key=lambda x: names.region_display(x)["title"])]})


@app.get("/api/scenarios")
def api_scenarios():
    user = _user()
    out = []
    for s in SC.list():
        full = {**s, "is_base": False}
        ok, _ = can_edit({**full, "name": s["name"]}, user)
        out.append({**s, "mine": bool(user) and fold(s["owner"]) == fold(user), "can_edit": ok})
    return jsonify(out)


@app.post("/api/scenarios")
def api_scenario_create():
    b = _body()
    user = _user(b)
    if not user:
        abort(400, description="add your name at the top right first")
    base = str(b.get("base") or (official_base() or {}).get("id", ""))
    s = SC.create(str(b.get("name", "")), user=user, base=base,
                  copy_from=str(b.get("copy_from", "") or ""))
    return jsonify({"id": s["id"], "name": s["name"], "base": s["base"]})


@app.patch("/api/scenarios/<sid>")
def api_scenario_patch(sid):
    """Rename and/or lock / unlock - owner only (a locked one can still be unlocked by its owner)."""
    b = _body()
    user = _user(b)
    s = SC.get(sid)
    if s.get("is_base"):
        abort(403, description="base models can't be changed")
    if not s.get("owner") or fold(s["owner"]) != fold(user):
        abort(403, description=f"only {s.get('owner') or 'its owner'} can change this working model")
    if "locked" not in b and s.get("locked"):
        abort(403, description="unlock it first")

    def change(x):
        if "name" in b:
            x["name"] = str(b.get("name") or "").strip()[:80] or x["name"]
        if "locked" in b:
            x["locked"] = bool(b["locked"])
    s = SC.update(sid, change)
    return jsonify({"id": s["id"], "name": s["name"], "locked": s["locked"]})


@app.delete("/api/scenarios/<sid>")
def api_scenario_delete(sid):
    s = SC.get(sid)
    require_edit(s, _user())
    SC.delete(sid)
    return jsonify({"ok": True})


@app.get("/api/scenarios/<sid>/version")
def api_scenario_version(sid):
    s = SC.get(sid)
    return jsonify({"version": s.get("version", 0), "updated": s.get("updated")})


@app.get("/api/scenarios/<sid>/changes")
def api_scenario_changes(sid):
    """Every change in the working model, across all areas, for the Changes tab."""
    s = SC.get(sid)
    rows = []
    for akey, cells in s.get("edits", {}).items():
        region, area = akey.split("|", 1)
        for key, rec in cells.items():
            fld, idx = key.split("|", 1)
            rows.append({"region": region, "area": area, "field": fld,
                         "label": FIELD_INFO.get(fld, (fld,))[0],
                         "index": [int(k) for k in idx.split(",")],
                         "value": rec["value"], "user": rec.get("user", ""), "ts": rec.get("ts", "")})
    for akey, pct in s.get("state_pct", {}).items():
        region, area = akey.split("|", 1)
        rows.append({"region": region, "area": area, "field": "state_pct", "label": "Demand % by state",
                     "states": sorted(pct), "user": s.get("owner", ""), "ts": s.get("state_pct_ts", {}).get(akey, "")})
    rows.sort(key=lambda r: r["ts"] or "", reverse=True)
    return jsonify(rows)


@app.get("/api/area")
def api_area():
    a = request.args
    s = SC.get(a.get("scenario", ""))
    return jsonify(area_payload(s, a.get("region", ""), a.get("area", ""), _user()))


@app.post("/api/edits")
def api_edits():
    b = _body()
    sid, region, area = str(b.get("scenario", "")), str(b.get("region", "")), str(b.get("area", ""))
    user = _user(b)
    s = SC.get(sid)
    require_edit(s, user)
    changes = b.get("changes")
    if not isinstance(changes, list) or not changes:
        abort(400, description="no changes sent")
    for ch in changes:
        if not isinstance(ch, dict) or ch.get("field") not in model.CELL_EDIT_FIELDS \
                or not isinstance(ch.get("index"), list):
            abort(400, description=f"bad change: {ch!r}")
        if not ch.get("revert"):
            if ch.get("value") is None and ch["field"] in model.BLANKABLE_EDIT_FIELDS:
                continue                       # blank freight = lane not used
            try:
                v = float(ch.get("value"))
            except (TypeError, ValueError):
                abort(400, description=f"{ch.get('value')!r} is not a number")
            if not np.isfinite(v):
                abort(400, description="values must be finite numbers")
    ak = area_key(region, area)

    def change(x):
        cells = dict(x["edits"].get(ak, {}))
        for ch in changes:
            key = cell_key(ch["field"], ch["index"])
            if ch.get("revert"):
                cells.pop(key, None)
            else:
                v = ch.get("value")
                cells[key] = {"value": None if v is None else float(v), "user": user, "ts": now_iso()}
        trial = {**x, "edits": {**x["edits"], ak: cells}}
        build(trial, region, area)                  # raises OverrideError if anything is off
        if cells:
            x["edits"][ak] = cells
        else:
            x["edits"].pop(ak, None)
    try:
        s = SC.update(sid, change)
    except model.OverrideError as exc:
        return error(400, str(exc))
    return jsonify(area_payload(s, region, area, user))


@app.post("/api/statepct")
def api_state_pct():
    """
    Set the % of budget demand to contract for, by state and month, for one
    area of a working model ({state: [12 fractions Apr..Mar]}); {} = back to base.
    Every plant's demand is scaled by new % / base %.
    """
    b = _body()
    sid, region, area = str(b.get("scenario", "")), str(b.get("region", "")), str(b.get("area", ""))
    user = _user(b)
    s = SC.get(sid)
    require_edit(s, user)
    pct = b.get("pct") or {}
    if not isinstance(pct, dict):
        abort(400, description="pct must be an object")
    raw, _, _ = area_raw(s, region, area)
    base_pct = base_pct_for(raw, s, region, area)
    clean = {}
    for st, vals in pct.items():
        if not isinstance(vals, list) or len(vals) != 12:
            abort(400, description=f"{st}: give 12 monthly values")
        try:
            v = [float(x) for x in vals]
        except (TypeError, ValueError):
            abort(400, description=f"{st}: values must be numbers")
        if any(not 0 <= x <= 3 for x in v):
            abort(400, description=f"{st}: % must be between 0% and 300%")
        if not np.allclose(v, base_pct.get(st) or [1.0] * 12):
            clean[str(st)[:12]] = v
    ak = area_key(region, area)

    def change(x):
        if clean:
            x["state_pct"][ak] = clean
            x.setdefault("state_pct_ts", {})[ak] = now_iso()
        else:
            x["state_pct"].pop(ak, None)
            x.setdefault("state_pct_ts", {}).pop(ak, None)
    s = SC.update(sid, change)
    return jsonify(area_payload(s, region, area, user))


@app.post("/api/structure")
def api_structure():
    """Save the model-builder settings (the sliders) for one area; {} = base size."""
    b = _body()
    sid, region, area = str(b.get("scenario", "")), str(b.get("region", "")), str(b.get("area", ""))
    user = _user(b)
    s = SC.get(sid)
    require_edit(s, user)
    st = b.get("structure") or {}
    if not isinstance(st, dict):
        abort(400, description="structure must be an object")
    plain = build(s, region, area, edits=[], structure={}, scale=False)
    clean = {}
    periods = st.get("periods")
    if periods not in (None, "", plain.n_periods):
        try:
            clean["periods"] = int(periods)
        except (TypeError, ValueError):
            abort(400, description="time periods must be a whole number")
    for key, fields in (("hubs", ("name", "price")), ("supply", ("location", "supplier", "hub")),
                        ("demand", ("name",))):
        items = st.get(key) or []
        if not isinstance(items, list):
            abort(400, description=f"{key} must be a list")
        rows = []
        for it in items:
            if not isinstance(it, dict):
                abort(400, description=f"bad {key} entry")
            row = {}
            for f in fields:
                v = it.get(f)
                if f == "price":
                    row[f] = v if isinstance(v, list) else ([v] if v not in (None, "") else [0])
                else:
                    row[f] = str(v or "").strip()[:80]
            rows.append(row)
        if rows:
            clean[key] = rows
    ak = area_key(region, area)
    info = {"dropped": 0}

    def change(x):
        d = build(x, region, area, edits=[], structure=clean)
        P, N, M = d.n_periods, d.n_supply, d.n_demand
        cells = dict(x["edits"].get(ak, {}))
        kept = {}
        for key, rec in cells.items():
            fld, idx = key.split("|", 1)
            pos = [int(k) for k in idx.split(",")]
            limit = {"p": P, "i": N, "j": M}
            if all(pos[n] < limit[a] for n, a in enumerate(model.CELL_EDIT_FIELDS.get(fld, ()))):
                kept[key] = rec
        build({**x, "edits": {**x["edits"], ak: kept}}, region, area, structure=clean)
        if clean:
            x["structure"][ak] = {**clean, "user": user, "ts": now_iso()}
        else:
            x["structure"].pop(ak, None)
        if kept:
            x["edits"][ak] = kept
        else:
            x["edits"].pop(ak, None)
        info["dropped"] = len(cells) - len(kept)
    try:
        s = SC.update(sid, change)
    except model.OverrideError as exc:
        return error(400, str(exc))
    out = area_payload(s, region, area, user)
    out["dropped_edits"] = info["dropped"]
    return jsonify(out)


# ---- uploads --------------------------------------------------------------- #

_CELL_XML_RE = re.compile(r"<c\b[^>]*?(?:/>|>(.*?)</c>)", re.S)
_VALUE_XML_RE = re.compile(r"<v(?:\s[^>]*)?>[^<]+</v>")


def uncalculated_formulas(path: str, tolerance: float = 0.02) -> int:
    """
    How many ModelBase formulas have no saved result (0 = the file is fine).

    The model reads the values Excel calculated and saved with the file. A
    workbook saved by something that doesn't calculate (a script, some
    converters) keeps the formulas but not their results, and every price
    built from a formula would silently read as 0. A few blanks are normal,
    so only more than *tolerance* of all formulas counts. Reads the sheet's
    XML straight out of the zip: a formula cell is ``<f>`` and its saved
    result a non-empty ``<v>``. Raises zipfile.BadZipFile for a non-workbook.
    """
    import zipfile
    with zipfile.ZipFile(path) as z:
        wb_xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
        m = re.search(r'<sheet\b[^>]*\bname="%s"[^>]*\br:id="([^"]+)"' % re.escape(model.MODEL_SHEET), wb_xml)
        if not m:
            return 0
        t = re.search(r'<Relationship\b[^>]*\bId="%s"[^>]*\bTarget="([^"]+)"' % re.escape(m.group(1)), rels) \
            or re.search(r'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="%s"' % re.escape(m.group(1)), rels)
        if not t:
            return 0
        target = t.group(1)
        part = target.lstrip("/") if target.startswith("/") else "xl/" + target
        sheet = z.read(part).decode("utf-8", "replace")
    formulas = missing = 0
    for cm in _CELL_XML_RE.finditer(sheet):
        body = cm.group(1) or ""
        if "<f" not in body:
            continue
        formulas += 1
        if not _VALUE_XML_RE.search(body):
            missing += 1
    if not formulas:
        return 0
    return missing if missing / formulas > tolerance else 0


@app.post("/api/upload")
def api_upload():
    """
    Upload an area workbook someone worked on offline. It goes into a NEW
    working model owned by the uploader, on top of the chosen base.
    """
    f = request.files.get("file")
    user = str(request.form.get("user", "")).strip()[:60]
    if not user:
        abort(400, description="add your name at the top right first")
    if f is None or not f.filename:
        abort(400, description="choose a workbook to upload")
    orig = os.path.basename(f.filename)
    ext = os.path.splitext(orig)[1].lower()
    if ext not in UPLOAD_EXTS:
        abort(400, description=f"'{orig}' isn't an Excel workbook (.xlsm or .xlsx)")
    bid = request.form.get("base") or (official_base() or {}).get("id", "")
    meta = BASES.get(bid)
    if not meta:
        abort(400, description="pick a base model first")
    cands = [(a["region"], a["area"]) for a in meta.get("areas", [])]
    region, area = request.form.get("region", ""), request.form.get("area", "")
    if not (region and area):
        guess = guess_area(orig, cands)
        if not guess:
            abort(400, description=f"couldn't tell which area '{orig}' is for - pick the area and upload again")
        region, area = guess
    if (region, area) not in cands:
        abort(400, description=f"{area} isn't in the base model {meta['name']}")

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    stored = f"{uuid.uuid4().hex}{ext}"
    path = os.path.join(UPLOAD_DIR, stored)
    f.save(path)
    try:
        try:
            stale = uncalculated_formulas(path)
        except Exception as exc:  # noqa: BLE001
            hint = (" It may be password-protected or still open and locked in Excel - "
                    "close it in Excel and try again." if "zip" in str(exc).lower() else "")
            raise ValueError(f"this file couldn't be opened as a model workbook ({exc}).{hint}")
        if stale:
            raise ValueError(
                f"{stale:,} formulas in this workbook have no calculated value saved, so its "
                f"prices and freight can't be trusted. That happens when a file is saved by "
                f"something other than Excel. Open it in Excel, save it, and upload it again.")
        try:
            data = model.ModelBaseReader(path, area, region).build_from_inputs(WB.raw(path, area, region))
        except model.ModelStructureError as exc:
            raise ValueError(f"the model can't read this workbook: {exc}")
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"this file couldn't be opened as a model workbook ({exc})")
    except ValueError as exc:
        os.remove(path)
        return error(422, str(exc))

    title = names.area_display(area)["title"]
    stamp = _dt.datetime.now().strftime("%b %d %H:%M")
    s = SC.create(f"{title} - {orig} ({stamp})", user=user, base=bid)
    ak = area_key(region, area)
    SC.update(s["id"], lambda x: x["sources"].__setitem__(ak, {
        "file": stored, "name": orig, "user": user, "ts": now_iso(), "size": os.path.getsize(path)}))
    LOG.info("upload %s -> %s / %s (%d supply x %d demand) as %s",
             orig, region, area, data.n_supply, data.n_demand, s["id"])
    return jsonify({"scenario": s["id"], "scenario_name": s["name"], "region": region, "area": area})


@app.post("/api/upload/remove")
def api_upload_remove():
    """Put an area back on the base model in this working model."""
    b = _body()
    sid, region, area = str(b.get("scenario", "")), str(b.get("region", "")), str(b.get("area", ""))
    user = _user(b)
    s = SC.get(sid)
    require_edit(s, user)
    ak = area_key(region, area)

    def change(x):
        for k in ("sources", "edits", "structure"):
            x[k].pop(ak, None)
    s = SC.update(sid, change)
    return jsonify(area_payload(s, region, area, user))


# ---- paste from Excel ------------------------------------------------------ #

def _paste_context(s: dict, region: str, area: str, kind: str, scope: str) -> PASTE.Context:
    if kind not in PASTE.KINDS:
        abort(400, description="pick what you're pasting")
    if scope == "all":
        meta = BASES.get(s["base"]) or {}
        targets = [(a["region"], a["area"]) for a in meta.get("areas", [])]
    else:
        targets = [(region, area)]
    areas = []
    for r, a in targets:
        try:
            areas.append((r, a, build(s, r, a), build(s, r, a, edits=[])))
        except model.OverrideError as exc:
            abort(400, description=f"{a}: {exc}")
    if not areas:
        abort(404, description="no areas to paste into")
    return PASTE.Context(areas, kind)


def _paste_run(b: dict) -> Tuple[dict, dict, str]:
    s = SC.get(str(b.get("scenario", "")))
    region, area = str(b.get("region", "")), str(b.get("area", ""))
    kind, scope = str(b.get("kind", "")), "all" if b.get("scope") == "all" else "area"
    if scope == "all" and PASTE.KINDS.get(kind, ("", "supply"))[1] not in ("plant", "lane"):
        abort(400, description="supply-point data can only be pasted into one area at a time "
                               "(the same supply point name appears in many areas)")
    opts = b.get("options") or {}
    if scope == "all" and (opts.get("from") or opts.get("to")):
        abort(400, description="pick one area to use a fixed supply point or plant")
    ctx = _paste_context(s, region, area, kind, scope)
    try:
        res = PASTE.analyse(ctx, str(b.get("text") or ""), opts, b.get("mapping") or {})
    except ValueError as exc:
        abort(400, description=str(exc))
    return s, res, scope


@app.post("/api/paste/preview")
def api_paste_preview():
    """What a paste would change - nothing is saved."""
    b = _body()
    s, res, scope = _paste_run(b)
    res.pop("_changes")
    res["can_edit"], res["why_not"] = can_edit(s, _user(b))
    res["scope"] = scope
    return jsonify(res)


@app.post("/api/paste/apply")
def api_paste_apply():
    """Apply a paste as ordinary cell edits (revertable in the change log)."""
    b = _body()
    user = _user(b)
    s, res, _ = _paste_run(b)
    require_edit(s, user)
    changes = res["_changes"]
    if not changes:
        abort(400, description="nothing to change - everything pasted matches the model already")
    fld = res["field"]
    by_area: Dict[Tuple[str, str], List[dict]] = {}
    for c in changes:
        by_area.setdefault((c["region"], c["area"]), []).append(c)
    stamp = now_iso()

    def change(x):
        for (r, a), lst in by_area.items():
            ak = area_key(r, a)
            cells = dict(x["edits"].get(ak, {}))
            for c in lst:
                key = cell_key(fld, c["index"])
                if c["revert"]:
                    cells.pop(key, None)
                else:
                    cells[key] = {"value": c["value"], "user": user, "ts": stamp, "via": f"paste: {res['label']}"}
            build({**x, "edits": {**x["edits"], ak: cells}}, r, a)      # raises OverrideError if off
            if cells:
                x["edits"][ak] = cells
            else:
                x["edits"].pop(ak, None)
    try:
        s = SC.update(s["id"], change)
    except model.OverrideError as exc:
        return error(400, str(exc))
    LOG.info("paste %s by %s: %d cell(s) in %d area(s) of %s", res["kind"], user, len(changes), len(by_area), s["id"])
    region, area = str(b.get("region", "")), str(b.get("area", ""))
    out = {"applied": len(changes), "areas": [{"region": r, "area": a, "cells": len(l)} for (r, a), l in by_area.items()]}
    if region and area:
        out["data"] = area_payload(s, region, area, user)
    return jsonify(out)


@app.get("/api/paste/template")
def api_paste_template():
    """What the model holds now, as a tab-separated table in the layout the paste reader expects."""
    a = request.args
    s = SC.get(a.get("scenario", ""))
    kind = a.get("kind", "")
    if kind not in PASTE.KINDS:
        abort(400, description="unknown kind")
    d = build(s, a.get("region", ""), a.get("area", ""))
    return app.response_class(PASTE.template(kind, d), mimetype="text/tab-separated-values")


@app.errorhandler(413)
def _too_big(exc):
    return error(413, f"that file is larger than {UPLOAD_MAX_MB} MB - is it the right workbook?")


# ---- solving and analysis -------------------------------------------------- #

def _areas_of(s: dict, region: str) -> List[str]:
    meta = BASES.get(s["base"]) or {}
    return [a["area"] for a in meta.get("areas", []) if a["region"] == region]


def solve_one(s: dict, region: str, area: str) -> dict:
    _, based_on, _ = area_raw(s, region, area)
    try:
        data = build(s, region, area)
    except model.OverrideError as exc:
        return {"area": area, "display": names.area_display(area),
                "summary": {"status": "error", "ok": False, "message": str(exc)},
                "diagnostics": [{"severity": "Error", "text": str(exc)}]}
    sol = model.solve_area(data)
    baseline = baseline_result(s["base"], region, area) \
        if (not s.get("is_base") and has_changes(s, region, area)) else None
    return result_payload(sol, baseline, based_on)


@app.post("/api/solve")
def api_solve():
    b = _body()
    s = SC.get(str(b.get("scenario", "")))
    region, area = str(b.get("region", "")), b.get("area")
    areas = [area] if area else _areas_of(s, region)
    if not areas:
        abort(404, description=f"no region {region!r}")
    return jsonify({"scenario": s["name"], "region": region,
                    "region_display": names.region_display(region),
                    "results": [solve_one(s, region, a) for a in areas]})


def _what_if(d: model.ModelData, **changes) -> dict:
    sol = model.solve_area(dataclasses.replace(d, **changes))
    return {"summary": summary_of(sol), "monthly": monthly_of(sol)}


@app.post("/api/analysis")
def api_analysis():
    """
    What-if runs on one area, never saved:

      {kind: "sensitivity", cents: 5, freight_pct: 5}
          each hub price up / down by `cents` per gallon, and freight up / down by %
      {kind: "stress", months: [8, 9, 10], pct: 15}
          demand up by `pct` % in those months (positions in the model's months)
    """
    b = _body()
    s = SC.get(str(b.get("scenario", "")))
    region, area = str(b.get("region", "")), str(b.get("area", ""))
    d = build(s, region, area)
    now = model.solve_area(d)
    base = {"summary": summary_of(now), "monthly": monthly_of(now)}
    kind = b.get("kind")
    if kind == "sensitivity":
        try:
            cents = float(b.get("cents", 5))
            fpct = float(b.get("freight_pct", 5))
        except (TypeError, ValueError):
            abort(400, description="numbers please")
        step = cents / 100.0 * model.GAL_PER_KGAL          # $/gal -> $/kgal
        rows = []
        hubs = sorted({h for h in d.supply_hubs if h})
        groups = [(f"{h} hub price", np.array([x == h for x in d.supply_hubs])) for h in hubs]
        if len(hubs) > 1:
            groups.append(("All hub prices", np.array([bool(x) for x in d.supply_hubs])))
        for label, mask in groups:
            up = _what_if(d, price=d.price + step * mask[None, :])
            dn = _what_if(d, price=d.price - step * mask[None, :])
            rows.append({"label": label, "change": f"\u00b1{cents:g}\u00a2/gal", "up": up["summary"],
                         "down": dn["summary"]})
        tc = d.trans_cost                 # every rate moves; closed (blank) lanes stay closed
        rows.append({"label": "Freight", "change": f"\u00b1{fpct:g}%",
                     "up": _what_if(d, trans_cost=tc * (1 + fpct / 100))["summary"],
                     "down": _what_if(d, trans_cost=tc * (1 - fpct / 100))["summary"]})
        return jsonify({"kind": kind, "now": base, "rows": rows})
    if kind == "stress":
        try:
            pct = float(b.get("pct", 15))
            months = [int(m) for m in (b.get("months") or [])]
        except (TypeError, ValueError):
            abort(400, description="numbers please")
        if not months or any(not 0 <= m < d.n_periods for m in months):
            abort(400, description="pick at least one month")
        if not -90 <= pct <= 200:
            abort(400, description="demand change should be between -90% and +200%")
        dem = d.demand.copy()
        dem[months] *= 1 + pct / 100.0
        sol = model.solve_area(dataclasses.replace(d, demand=dem))
        unmet = []
        if sol.ok and sol.shortfall is not None:
            for j in np.argsort(-sol.shortfall.sum(axis=0))[:10]:
                v = float(sol.shortfall[:, j].sum())
                if v > 1e-6:
                    unmet.append({"plant": d.demand_names[j], "unmet": fnum(v)})
        return jsonify({"kind": kind, "now": base, "pct": pct, "months": months,
                        "stressed": {"summary": summary_of(sol), "monthly": monthly_of(sol)},
                        "unmet_plants": unmet,
                        "extra_demand": fnum(float(dem.sum() - d.demand.sum()))})
    abort(400, description="unknown analysis")


@app.get("/api/compare")
def api_compare():
    """Two working models (or a base) side by side for one region: ?a=&b=&region="""
    a, b = SC.get(request.args.get("a", "")), SC.get(request.args.get("b", ""))
    region = request.args.get("region", "")
    areas_b = {fold(re.sub(r"\s*\d{2}-\d{2}\s*$", "", x)): x for x in _areas_of(b, region)}
    rows = []
    for area in _areas_of(a, region):
        other = areas_b.get(fold(re.sub(r"\s*\d{2}-\d{2}\s*$", "", area)))
        row = {"area": area, "display": names.area_display(area)}
        for key, s, ar in (("a", a, area), ("b", b, other)):
            if not ar:
                row[key] = None
                continue
            try:
                sol = model.solve_area(build(s, region, ar))
                row[key] = {"summary": summary_of(sol), "monthly": monthly_of(sol),
                            "periods": sol.data.period_labels}
            except model.OverrideError as exc:
                row[key] = {"summary": {"ok": False, "status": "error", "message": str(exc)}}
        rows.append(row)
    return jsonify({"a": {"id": a["id"], "name": a["name"], "owner": a.get("owner", ""), "is_base": a.get("is_base")},
                    "b": {"id": b["id"], "name": b["name"], "owner": b.get("owner", ""), "is_base": b.get("is_base")},
                    "region": region, "region_display": names.region_display(region), "areas": rows})


@app.get("/api/export")
def api_export():
    """The region results as the familiar region workbook (.xlsx), for download."""
    s = SC.get(request.args.get("scenario", ""))
    region = request.args.get("region", "")
    areas = _areas_of(s, region)
    if not areas:
        abort(404, description=f"no region {region!r}")
    sols = [model.solve_area(build(s, region, a)) for a in areas]
    os.makedirs(EXPORT_DIR, exist_ok=True)
    for fn in os.listdir(EXPORT_DIR):
        p = os.path.join(EXPORT_DIR, fn)
        if time.time() - os.path.getmtime(p) > 3600:
            try:
                os.remove(p)
            except OSError:
                pass
    today = _dt.date.today()
    name = model.region_filename(region, today, scenario=s["name"])
    path = os.path.join(EXPORT_DIR, f"{uuid.uuid4().hex[:8]}_{name}")
    note = f"Web app: '{s['name']}' on base '{(BASES.get(s['base']) or {}).get('name', '')}'"
    model.RegionWorkbookWriter(region, sols, today, source_note=note, scenario=s["name"]).write(path)
    return send_file(path, as_attachment=True, download_name=name)


# --------------------------------------------------------------------------- #

def first_run_setup() -> None:
    """No base models yet: read them from the Excel workbooks, once, in the background."""
    if BASES.list() or BASES.running():
        return
    dfs = find_demand_files()

    def job(progress):
        df = demand_file(dfs[0]) if dfs else None
        progress(0, 2, "Working workbooks")
        working = BASES.import_workbooks(
            WB.root, "Working", "setup", name="2026-27 Working files (from Excel)", demand_file=df,
            progress=lambda d, t, s: progress(d, t * 2, f"Working files: {s}"))
        baseline = BASES.import_workbooks(
            WB.root, "Baseline", "setup", name="2026-27 Base (from Excel Baseline files)", demand_file=df,
            progress=lambda d, t, s: progress(t + d, t * 2, f"Baseline files: {s}"))
        set_official(baseline["id"])
        return working
    BASES.start_job("First-time setup: reading the Excel models", job)


def init(root: Optional[str] = None) -> None:
    global WB, SC, BASES
    root = os.path.abspath(root or os.environ.get("MODEL_ROOT") or model.default_root())
    WB = Workbooks(root)
    BASES = B.Bases(BASE_DIR)
    SC = Scenarios(SCEN_DIR, BASES)
    LOG.info("model root: %s", root)
    LOG.info("data      : %s", DATA_DIR)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(message)s")
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleTitleW(f"Propane Model Web App - http://{HOST}:{PORT}")
        except Exception:  # noqa: BLE001 - cosmetic only
            pass
    init()
    first_run_setup()
    LOG.info("open http://%s:%d in your browser", HOST, PORT)
    app.run(host=HOST, port=PORT, threaded=True, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
