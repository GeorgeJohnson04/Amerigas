"""
app.py - the Supply Planner web app (local only).

    py app.py            then open http://127.0.0.1:8070/
    (earlier prototype; the current app is in ..\\webapp)

Everything runs on this PC. Data lives in .\\data (see store.py). The original
Excel workbooks are only read once, if someone chooses "convert last year's
models" on the welcome screen.

SECURITY: there is no login. The app only listens on 127.0.0.1 so nobody else
on the network can reach it; requests must carry a Host header of
localhost/127.0.0.1 (blocks DNS rebinding) and every request that changes data
must carry an X-Requested-With header a foreign web page can't set (blocks
cross-site requests). The name people type in the top bar is a signature for
the change log and plan ownership, not a password.
"""

from __future__ import annotations

import copy
import datetime as _dt
import io
import logging
import os
import re
import sys
import tempfile
import threading
import time
import urllib.parse
import uuid
from collections import OrderedDict
from typing import Dict, List, Optional

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import numpy as np  # noqa: E402

try:
    from flask import Flask, abort, jsonify, request, send_file  # noqa: E402
except ImportError:  # pragma: no cover
    raise SystemExit("Flask is required:  py -m pip install flask==3.1.3")

import engine as E  # noqa: E402
import importer as I  # noqa: E402
import names  # noqa: E402
import store as S  # noqa: E402
import weather as W  # noqa: E402

LOG = logging.getLogger("planner")
HOST = "127.0.0.1"
PORT = int(os.environ.get("PLANNER_PORT", "8070"))
DATA_DIR = os.environ.get("PLANNER_DATA") or os.path.join(HERE, "data")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
UPLOAD_MAX_MB = 40
CSRF_HEADER, CSRF_VALUE = "X-Requested-With", "supply-planner"

app = Flask(__name__, static_folder=os.path.join(HERE, "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = UPLOAD_MAX_MB * 1024 * 1024
STORE = S.Store(DATA_DIR)
os.makedirs(UPLOAD_DIR, exist_ok=True)


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #

def jarr(a, nd: int = 4):
    """numpy -> lists, rounded, NaN/inf -> null."""
    a = np.asarray(a, dtype=float)
    out = np.round(a, nd).astype(object)
    out[~np.isfinite(a)] = None
    return out.tolist()


def fnum(x, nd: int = 4):
    if x is None:
        return None
    x = float(x)
    return round(x, nd) if np.isfinite(x) else None


def error(status: int, message: str):
    r = jsonify({"error": message})
    r.status_code = status
    return r


def user() -> str:
    u = request.headers.get("X-User", "")
    return urllib.parse.unquote(u).strip()[:60]


def body() -> dict:
    return request.get_json(silent=True) or {}


def view_arg(b: Optional[dict] = None) -> str:
    return ((b or {}).get("view") or request.args.get("view") or "").strip()


@app.errorhandler(S.Denied)
def _denied(e):
    return error(403, str(e))


@app.errorhandler(S.NotFound)
def _nf(e):
    return error(404, str(e))


@app.errorhandler(ValueError)
def _bad(e):
    return error(400, str(e))


@app.errorhandler(413)
def _big(e):
    return error(413, f"that file is bigger than {UPLOAD_MAX_MB} MB")


@app.before_request
def _guards():
    host = (request.host or "").split(":")[0].lower()
    if host not in ("127.0.0.1", "localhost"):
        abort(403)
    if request.method in ("POST", "PUT", "PATCH", "DELETE") and request.headers.get(CSRF_HEADER) != CSRF_VALUE:
        abort(403)


@app.after_request
def _no_cache(resp):
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    return resp


# --------------------------------------------------------------------------- #
# is this server running today's code? (open_browser.py restarts stale ones)
# --------------------------------------------------------------------------- #

_CODE = [os.path.join(HERE, f) for f in ("app.py", "engine.py", "importer.py", "store.py", "weather.py", "legacy.py", "names.py")]


def _fingerprint() -> str:
    parts = []
    for f in _CODE:
        try:
            st = os.stat(f)
            parts.append(f"{os.path.basename(f)}:{st.st_size}:{int(st.st_mtime)}")
        except OSError:
            parts.append("missing")
    return "|".join(parts)


STARTED = _fingerprint()


@app.get("/api/health")
def health():
    return jsonify({"ok": True, "app": "supply-planner", "stale": _fingerprint() != STARTED})


@app.post("/api/shutdown")
def shutdown():
    threading.Timer(0.3, lambda: os._exit(0)).start()
    return jsonify({"ok": True})


@app.get("/")
def index():
    return send_file(os.path.join(HERE, "static", "index.html"))


# --------------------------------------------------------------------------- #
# demand + solving (with a cache, so moving around the app is instant)
# --------------------------------------------------------------------------- #

_HIST: Dict[str, W.History] = {}
_SOLVES: "OrderedDict[str, dict]" = OrderedDict()
_LAST: Dict[tuple, str] = {}               # (view, area) -> cache key of its last run
_SOLVE_LOCK = threading.Lock()
SOLVE_CACHE = 80


def history() -> W.History:
    v = STORE.weather_version()
    if v not in _HIST:
        _HIST.clear()
        _HIST[v] = W.History(STORE.weather())
    return _HIST[v]


def breakdown(view: str, area: dict, settings: Optional[dict] = None, state_pct: Optional[dict] = None) -> dict:
    settings = settings or STORE.demand_settings(view)
    sp = state_pct if state_pct is not None else STORE.state_pct(view)
    wf = history().weather_fn(area, settings.get("weather"))
    return E.demand_breakdown(area, sp, wf, settings.get("adjust") or [])


def solve_key(area: dict, planned: np.ndarray, extra=None) -> str:
    return E.content_hash(area, np.round(planned, 3).tolist(), extra)


def solve_cached(area: dict, planned: np.ndarray, price_add=None, freight_mult: float = 1.0, extra=None) -> dict:
    key = solve_key(area, planned, [None if price_add is None else np.round(price_add, 6).tolist(), freight_mult, extra])
    with _SOLVE_LOCK:
        hit = _SOLVES.get(key)
        if hit is not None:
            _SOLVES.move_to_end(key)
            return hit
    r = E.solve(area, planned, price_add=price_add, freight_mult=freight_mult)
    r["key"] = key
    with _SOLVE_LOCK:
        _SOLVES[key] = r
        while len(_SOLVES) > SOLVE_CACHE:
            _SOLVES.popitem(last=False)
    return r


def run_view_area(view: str, aid: str) -> dict:
    area = STORE.area(view, aid)
    if area is None:
        raise S.NotFound("unknown area")
    bd = breakdown(view, area)
    r = solve_cached(area, bd["planned"])
    _LAST[(view, aid)] = r["key"]
    return r


def last_result(view: str, aid: str):
    """(result or None, fresh?)"""
    key = _LAST.get((view, aid))
    r = _SOLVES.get(key) if key else None
    if r is None:
        return None, False
    area = STORE.area(view, aid)
    bd = breakdown(view, area)
    fresh = solve_key(area, bd["planned"], [None, 1.0, None]) == key
    return r, fresh


# --------------------------------------------------------------------------- #
# what people see about areas
# --------------------------------------------------------------------------- #

def area_label(a: dict) -> dict:
    d = names.area_display(a.get("name", ""))
    return {"id": a["id"], "name": a.get("name", ""), "title": a.get("title") or d["title"], "region": a.get("region", "Other"),
            "states": d.get("states") or " \u00b7 ".join(sorted({p.get("state", "") for p in a.get("plants", []) if p.get("state")}))}


def region_info(r: str) -> dict:
    return names.region_display(r)


def view_info(view: str) -> dict:
    ds, p = STORE.parse_view(view)
    meta = STORE.meta(ds)
    u = user()
    out = {"view": view, "dataset": meta, "plan": None, "can_edit": STORE.can_edit(view, u),
           "months": STORE.months(view), "demand": STORE.demand_settings(view)}
    out["demand_text"] = W.describe(out["demand"]["weather"])
    out["state_pct_states"] = len([s for s, v in STORE.state_pct(view).items() if v])
    if p:
        out["plan"] = {k: p.get(k) for k in ("id", "name", "owner", "dataset", "created", "updated", "note", "origin")}
        out["plan"]["areas_changed"] = sorted(p.get("areas", {}))
        out["plan"]["state_pct_changed"] = sum(len(v) for v in (p.get("state_pct") or {}).values())
    return out


@app.get("/api/boot")
def boot():
    try:
        import legacy
        n_legacy = len(legacy.available("Working"))
    except BaseException:  # noqa: BLE001 - no old models on this PC is fine (discover may SystemExit)
        n_legacy = 0
    return jsonify({
        "datasets": STORE.list_datasets(), "plans": STORE.list_plans(), "legacy": n_legacy,
        "kinds": {k: {x: v[x] for x in ("label", "group", "help", "unit")} for k, v in I.KINDS.items()},
        "order": I.ORDER, "weather_modes": W.MODES, "high_freight": E.HIGH_FREIGHT,
        "regions": {r: region_info(r) for r in sorted({a.get("region", "") for d in STORE.list_datasets()
                                                        for a in STORE.ds_areas(d["id"]).values()})},
    })


@app.get("/api/view")
def get_view():
    view = view_arg()
    info = view_info(view)
    areas = STORE.areas(view)
    changed = set((info.get("plan") or {}).get("areas_changed") or [])
    rows = []
    for aid, a in sorted(areas.items(), key=lambda kv: (kv[1].get("region", ""), kv[1].get("name", ""))):
        bd = breakdown(view, a)
        r, fresh = last_result(view, aid)
        row = area_label(a) | {
            "supply": len(a.get("supply", [])), "plants": len(a.get("plants", [])), "lanes": len(a.get("lanes", [])),
            "budget": fnum(bd["budget"].sum(), 0), "planned": fnum(bd["planned"].sum(), 0),
            "changed": aid in changed, "run": None,
        }
        if r is not None:
            s = E.summary(r)
            row["run"] = {"status": s["status"], "fresh": fresh, "total": fnum(s.get("total"), 0),
                          "per_gal": fnum(s.get("per_gal")), "unmet": fnum(s.get("unmet"), 0)}
        rows.append(row)
    info["areas"] = rows
    info["regions"] = {r: region_info(r) for r in sorted({x["region"] for x in rows})}
    w = history()
    info["weather"] = {"loaded": not w.empty, "seasons": len(w.seasons), "places": len(w.places), "unit": w.unit}
    return jsonify(info)


# --------------------------------------------------------------------------- #
# setup, data sets, plans
# --------------------------------------------------------------------------- #

@app.post("/api/setup/convert")
def setup_convert():
    import legacy
    b = body()
    u = user()
    if not u:
        raise S.Denied("enter your name first (top right)")
    variant = b.get("variant") if b.get("variant") in ("Working", "Baseline") else "Working"
    areas, report = legacy.convert(variant, names_mod=names)
    if not areas:
        raise ValueError("no workbooks could be read")
    year = b.get("year") or "2026-27"
    ds = STORE.create_dataset(b.get("name") or f"{year} (from last year\u2019s {variant} models)", year, u, areas, {},
                              source=f"last year\u2019s {variant} Excel models", note="Converted from Excel")
    return jsonify({"dataset": ds, "report": report})


@app.post("/api/datasets")
def create_dataset():
    b = body()
    u = user()
    if not u:
        raise S.Denied("enter your name first (top right)")
    if b.get("copy_from"):
        ds = STORE.copy_dataset(b["copy_from"], b.get("name") or "", b.get("year") or "", u)
    else:
        ds = STORE.create_dataset(b.get("name") or b.get("year") or "New data set", b.get("year") or "", u)
    return jsonify({"dataset": ds})


@app.patch("/api/datasets/<ds>")
def rename_dataset(ds):
    STORE.rename_dataset(ds, body().get("name") or ds, user() or "?")
    return jsonify({"ok": True})


@app.delete("/api/datasets/<ds>")
def delete_dataset(ds):
    if not user():
        raise S.Denied("enter your name first")
    STORE.delete_dataset(ds, user())
    return jsonify({"ok": True})


@app.post("/api/plans")
def create_plan():
    b = body()
    p = STORE.create_plan(b.get("dataset") or "", b.get("name") or "My plan", user(), copy_from=b.get("copy_from"),
                          note=b.get("note") or "")
    return jsonify({"plan": p["id"], "view": f"plan:{p['id']}"})


@app.patch("/api/plans/<pid>")
def rename_plan(pid):
    b = body()
    STORE.rename_plan(pid, user(), b.get("name") or "", b.get("note"))
    return jsonify({"ok": True})


@app.delete("/api/plans/<pid>")
def delete_plan(pid):
    STORE.delete_plan(pid, user())
    return jsonify({"ok": True})


# --------------------------------------------------------------------------- #
# one area
# --------------------------------------------------------------------------- #

@app.get("/api/area")
def get_area():
    view, aid = view_arg(), request.args.get("area", "")
    a = STORE.area(view, aid)
    if a is None:
        raise S.NotFound("unknown area")
    bd = breakdown(view, a)
    h = history()
    places = {}
    for pl in a.get("plants", []):
        pk = h.place_for(pl, a)
        places[pl["id"]] = h.label.get(pk, "") if pk else ""
    r, fresh = last_result(view, aid)
    ship = None
    if r is not None and r.get("ok"):
        ar = r["arrays"]
        ship = {f"{ln['from']}|{ln['to']}": jarr(r["ship"][l], 0) for l, ln in enumerate(ar.lanes)}
    ds, p = STORE.parse_view(view)
    return jsonify({
        "area": a, "label": area_label(a),
        "breakdown": {k: jarr(v, 6 if k in ("pct", "weather", "adjust") else 1) for k, v in bd.items()},
        "checks": E.checks(a, bd["planned"]),
        "weather": {"coverage": h.coverage(a) if not h.empty else None, "places": places},
        "changed_in_plan": bool(p and aid in p.get("areas", {})),
        "ship": ship, "ship_fresh": fresh,
        "state_pct": STORE.state_pct(view),
    })


@app.post("/api/edit")
def edit():
    b = body()
    e = STORE.edit(view_arg(b), b.get("area", ""), b.get("items") or [], user(), b.get("label") or "")
    return jsonify(e)


@app.post("/api/add")
def add():
    b = body()
    e = STORE.add(view_arg(b), b.get("area", ""), b.get("kind", ""), b.get("obj") or {}, user())
    return jsonify(e)


@app.post("/api/remove")
def remove():
    b = body()
    e = STORE.remove(view_arg(b), b.get("area", ""), b.get("kind", ""), b.get("ident"), user())
    return jsonify(e)


@app.post("/api/areas")
def add_area():
    b = body()
    a = STORE.add_area(view_arg(b), b.get("name") or "", b.get("region") or "", user(), b.get("title") or "")
    return jsonify({"area": a["id"]})


@app.patch("/api/areas")
def update_area():
    b = body()
    return jsonify(STORE.update_area_info(view_arg(b), b.get("area", ""), user(),
                                          name=b.get("name"), title=b.get("title"), region=b.get("region")))


@app.delete("/api/areas")
def delete_area():
    b = body()
    return jsonify(STORE.delete_area(view_arg(b), b.get("area", ""), user()))


@app.post("/api/reset_area")
def reset_area():
    b = body()
    return jsonify(STORE.reset_area(view_arg(b), b.get("area", ""), user()))


@app.post("/api/promote_area")
def promote_area():
    b = body()
    return jsonify(STORE.promote_area(view_arg(b), b.get("area", ""), user()))


# --------------------------------------------------------------------------- #
# demand: contract % by state, weather + adjustment rules
# --------------------------------------------------------------------------- #

@app.post("/api/state_pct")
def set_state_pct():
    b = body()
    return jsonify(STORE.set_state_pct(view_arg(b), b.get("items") or [], user()))


@app.post("/api/demand")
def set_demand():
    b = body()
    return jsonify(STORE.set_demand(view_arg(b), b.get("settings") or {}, user(), b.get("label") or ""))


@app.post("/api/demand/preview")
def demand_preview():
    """Planned demand by month now vs with unsaved settings - for one area or all of them."""
    b = body()
    view = view_arg(b)
    new = S.normal_demand(b.get("settings") or {})
    aids = [b["area"]] if b.get("area") else STORE.area_ids(view)
    months = STORE.months(view)
    now_tot, new_tot, budget = np.zeros(len(months)), np.zeros(len(months)), np.zeros(len(months))
    by_area = []
    sp = STORE.state_pct(view)
    for aid in aids:
        a = STORE.area(view, aid)
        if not a:
            continue
        b0 = breakdown(view, a, state_pct=sp)
        b1 = breakdown(view, a, settings=new, state_pct=sp)
        P = min(len(months), b0["planned"].shape[1])
        now_tot[:P] += b0["planned"].sum(axis=0)[:P]
        new_tot[:P] += b1["planned"].sum(axis=0)[:P]
        budget[:P] += b0["budget"].sum(axis=0)[:P]
        d = float(b1["planned"].sum() - b0["planned"].sum())
        if abs(d) > 0.5:
            by_area.append({"area": aid, "name": a.get("name", aid), "now": fnum(b0["planned"].sum(), 0),
                            "new": fnum(b1["planned"].sum(), 0), "change": fnum(d, 0)})
    by_area.sort(key=lambda x: -abs(x["change"]))
    return jsonify({"months": months, "budget": jarr(budget, 0), "now": jarr(now_tot, 0), "new": jarr(new_tot, 0),
                    "by_area": by_area[:40], "text": W.describe(new["weather"])})


# --------------------------------------------------------------------------- #
# weather history
# --------------------------------------------------------------------------- #

@app.get("/api/weather")
def weather_summary():
    h = history()
    view = view_arg()
    cov = []
    if view:
        for aid, a in STORE.areas(view).items():
            if not h.empty:
                cov.append({"area": aid, "name": a.get("name", aid), **h.coverage(a)})
    return jsonify({"loaded": not h.empty, **h.summary(), "coverage": cov, "modes": W.MODES})


@app.get("/api/weather/table")
def weather_table():
    h = history()
    place = W.fold(request.args.get("place", ""))
    if place not in h.places:
        raise S.NotFound("no history for that place")
    return jsonify(h.table(place))


@app.post("/api/weather/clear")
def weather_clear():
    if not user():
        raise S.Denied("enter your name first")
    b = body()
    ops = [{"op": "weather", "rows": [], "unit": "gal", "replace": True}]
    return jsonify(STORE.apply_import(view_arg(b), ops, user(), "Cleared the weather history"))


# --------------------------------------------------------------------------- #
# uploads: paste or Excel -> preview -> apply
# --------------------------------------------------------------------------- #

_TOKEN = re.compile(r"^[a-f0-9]{16}$")


def _clean_uploads():
    cut = time.time() - 86400
    for f in os.listdir(UPLOAD_DIR):
        p = os.path.join(UPLOAD_DIR, f)
        try:
            if os.path.getmtime(p) < cut:
                os.remove(p)
        except OSError:
            pass


@app.post("/api/import/upload")
def import_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        raise ValueError("choose an Excel file")
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("upload an .xlsx or .xlsm workbook (save older .xls files as .xlsx first)")
    _clean_uploads()
    tok = uuid.uuid4().hex[:16]
    path = os.path.join(UPLOAD_DIR, f"{tok}.xlsx")
    f.save(path)
    try:
        sheets = I.read_workbook(path)
    except Exception as e:  # noqa: BLE001
        os.remove(path)
        raise ValueError(f"couldn't read that workbook ({e})")
    meta = [{"sheet": s["sheet"], "kind": s["kind"], "maybe": s.get("maybe"), "rows": s["rows"],
             "head": [r[:14] for r in s["table"][:6]]} for s in sheets]
    _UPLOADS[tok] = {"name": f.filename, "sheets": {s["sheet"]: s["table"] for s in sheets}, "ts": time.time()}
    return jsonify({"upload": tok, "name": f.filename, "sheets": meta})


_UPLOADS: Dict[str, dict] = {}


def _table_for(b: dict):
    tok = b.get("upload")
    if not tok:
        return None, b.get("text") or ""
    if not _TOKEN.match(tok):
        raise ValueError("bad upload")
    up = _UPLOADS.get(tok)
    if up is None:
        path = os.path.join(UPLOAD_DIR, f"{tok}.xlsx")
        if not os.path.exists(path):
            raise ValueError("that upload has expired - choose the file again")
        up = _UPLOADS[tok] = {"name": "", "sheets": {s["sheet"]: s["table"] for s in I.read_workbook(path)}}
    t = up["sheets"].get(b.get("sheet") or "")
    if t is None:
        raise ValueError("that sheet isn't in the workbook")
    return t, ""


def _parse(b: dict):
    view = view_arg(b)
    ctx = {"areas": STORE.areas(view), "state_pct": STORE.state_pct(view), "months": STORE.months(view),
           "weather": STORE.weather()}
    table, text = _table_for(b)
    return I.parse(b.get("kind") or "", text, ctx, scope=b.get("scope") or None, opts=b.get("opts") or {},
                   mapping=b.get("mapping") or {}, table=table)


@app.post("/api/import/preview")
def import_preview():
    pv = _parse(body())
    pv.pop("_ops", None)
    return jsonify(pv)


@app.post("/api/import/apply")
def import_apply():
    b = body()
    pv = _parse(b)
    if not pv["_ops"]:
        return jsonify({"count": 0, "message": "Nothing to change - the data already matches."})
    src = b.get("source") or ("pasted" if not b.get("upload") else f"from {b.get('sheet')}")
    e = STORE.apply_import(view_arg(b), pv["_ops"], user(), f"{pv['label']} ({src})")
    return jsonify({**e, "changes": pv["changes"], "creates": len(pv["creates"])})


@app.post("/api/import/apply_workbook")
def import_apply_workbook():
    """Several sheets of one upload, one after another (each can be undone on its own)."""
    b = body()
    done = []
    for s in b.get("sheets") or []:
        if not s.get("kind"):
            continue
        sb = {**b, "sheet": s["sheet"], "kind": s["kind"], "opts": s.get("opts") or {}, "mapping": s.get("mapping") or {}}
        pv = _parse(sb)
        if not pv["_ops"]:
            done.append({"sheet": s["sheet"], "count": 0})
            continue
        e = STORE.apply_import(view_arg(b), pv["_ops"], user(), f"{pv['label']} (sheet \u201c{s['sheet']}\u201d)")
        done.append({"sheet": s["sheet"], "count": e.get("count", 0), "id": e.get("id")})
    return jsonify({"done": done})


@app.get("/api/template")
def template():
    view = view_arg()
    kinds = [k for k in (request.args.get("kinds") or "").split(",") if k in I.KINDS] or None
    blank = request.args.get("blank") == "1"
    aid = request.args.get("area")
    if blank:
        areas, sp, w, title = {}, {}, {}, "Supply planning data - blank template"
    else:
        areas = STORE.areas(view)
        if aid:
            areas = {aid: areas[aid]} if aid in areas else {}
        sp, w = STORE.state_pct(view), STORE.weather()
        title = f"Supply planning data - {view_info(view)['dataset']['name']}"
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    I.export_workbook(path, areas, sp, w, kinds, title)
    with open(path, "rb") as fh:
        data = io.BytesIO(fh.read())
    os.remove(path)
    nm = "Supply data template.xlsx" if blank else f"Supply data {_dt.date.today().isoformat()}.xlsx"
    return send_file(data, as_attachment=True, download_name=nm,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --------------------------------------------------------------------------- #
# changes
# --------------------------------------------------------------------------- #

@app.get("/api/changes")
def changes():
    view = view_arg()
    aid = request.args.get("area")
    out = []
    for e in STORE.changes(view):
        if aid and e.get("area") != aid and aid not in (e.get("areas") or []):
            continue
        out.append({k: e.get(k) for k in ("id", "ts", "user", "kind", "label", "area", "areas", "count", "undone", "undo")}
                   | {"can_undo": e.get("kind") in ("edit", "state_pct", "demand") or bool(e.get("undo")),
                      "sample": (e.get("items") or [])[:6]})
        if len(out) >= 400:
            break
    return jsonify({"changes": out})


@app.post("/api/undo")
def undo():
    b = body()
    return jsonify(STORE.undo(view_arg(b), b.get("id", ""), user()))


# --------------------------------------------------------------------------- #
# results
# --------------------------------------------------------------------------- #

def results_payload(view: str, aid: str, r: dict, fresh: bool = True, baseline: bool = True) -> dict:
    a: E.Arrays = r["arrays"]
    s = E.summary(r)
    out = {"area": aid, "status": r["status"], "ok": r["ok"], "fresh": fresh, "summary": {k: (fnum(v) if isinstance(v, float) else v)
                                                                                       for k, v in s.items()},
           "months": a.months,
           "explain": r["explain"] if "explain" in r else ([{"level": "error", "text": r["message"]}] if r.get("message") else [])}
    if not r["ok"]:
        return out
    P = a.P
    out["monthly"] = {k: jarr(v, 0) for k, v in E.monthly(r).items()}
    out["supply"] = [{
        "id": s_["id"], "location": s_["location"], "supplier": s_.get("supplier", ""), "hub": s_.get("hub", ""),
        "contract": jarr(a.contract[i], 0), "bought": jarr(r["bought"][i], 0), "price": jarr(a.price[i], 5),
        "min": jarr(a.contract[i] * a.min_pct[i], 0), "max": jarr(a.contract[i] * a.max_pct[i], 0),
        "cost": fnum(r["commodity_cost"][i].sum(), 0), "surplus": jarr(r["surplus"][i], 0),
        "store_end": jarr(r["s_end"][i], 0), "store_cap": jarr(a.s_cap[i], 0),
    } for i, s_ in enumerate(a.supply)]
    out["plants"] = [{
        "id": p["id"], "name": p["name"], "state": p.get("state", ""),
        "demand": jarr(a.demand[j], 0), "delivered": jarr(r["delivered"][j], 0), "unmet": jarr(r["unmet"][j], 0),
        "store_end": jarr(r["p_end"][j], 0), "store_cap": jarr(a.p_cap[j], 0),
        "from_store": jarr(r["p_wd"][j], 0), "into_store": jarr(r["p_inj"][j], 0),
        "freight": fnum(sum(float(r["freight_cost"][l].sum()) for l in range(a.L) if a.lane_j[l] == j), 0),
    } for j, p in enumerate(a.plants)]
    lanes = []
    for l, ln in enumerate(a.lanes):
        lanes.append({"from": ln["from"], "to": ln["to"], "rate": jarr(a.rate[l], 5), "ship": jarr(r["ship"][l], 0),
                      "cost": fnum(r["freight_cost"][l].sum(), 0), "high": [bool(x) for x in a.high[l]],
                      "total": fnum(r["ship"][l].sum(), 0)})
    out["lanes"] = lanes
    ins = E.insights(r)
    out["insights"] = ins
    if baseline:
        ds, p = STORE.parse_view(view)
        if p is not None:
            base_area = STORE.ds_area(ds, aid)
            if base_area is not None:
                bv = f"ds:{ds}"
                bb = breakdown(bv, base_area)
                br = solve_cached(base_area, bb["planned"])
                out["baseline"] = {k: (fnum(v) if isinstance(v, float) else v) for k, v in E.summary(br).items()}
                if br["ok"]:
                    out["baseline_monthly"] = {k: jarr(v, 0) for k, v in E.monthly(br).items()}
    return out


@app.post("/api/solve")
def solve():
    b = body()
    view, aid = view_arg(b), b.get("area", "")
    r = run_view_area(view, aid)
    return jsonify(results_payload(view, aid, r, True))


@app.get("/api/results")
def results():
    view, aid = view_arg(), request.args.get("area", "")
    r, fresh = last_result(view, aid)
    if r is None:
        return jsonify({"none": True})
    return jsonify(results_payload(view, aid, r, fresh))


@app.get("/api/results/export")
def results_export():
    """Results of every area that has been run in this view (or one area) as Excel."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    view = view_arg()
    only = request.args.get("area")
    info = view_info(view)
    areas = STORE.areas(view)
    aids = [only] if only else sorted(areas, key=lambda x: (areas[x].get("region", ""), areas[x].get("name", "")))
    wb = openpyxl.Workbook()
    head = PatternFill("solid", fgColor="1F3A5F")
    wsum = wb.active
    wsum.title = "Summary"
    title = info["dataset"]["name"] + (f" \u2014 {info['plan']['owner']}: {info['plan']['name']}" if info["plan"] else " \u2014 data set")
    wsum.append([title])
    wsum["A1"].font = Font(bold=True, size=14)
    wsum.append([f"Demand: {info['demand_text']}", f"Exported {S.now_iso()}"])
    wsum.append([])
    hdr = ["Region", "Area", "Status", "Planned demand (gal)", "Bought (gal)", "Propane $", "Freight $", "Total $",
           "$/gal", "Unmet demand (gal)", "High-freight gal"]
    wsum.append(hdr)
    sheets = {n: wb.create_sheet(n) for n in ("Supply plan", "Lane volumes", "Plants")}
    months = info["months"]
    sheets["Supply plan"].append(["Area", "Location", "Supplier", "Hub", "What"] + months + ["Total"])
    sheets["Lane volumes"].append(["Area", "From", "Supplier", "To", "What"] + months + ["Total"])
    sheets["Plants"].append(["Area", "Plant", "State", "What"] + months + ["Total"])
    for aid in aids:
        if aid not in areas:
            continue
        r = run_view_area(view, aid)
        s = E.summary(r)
        a = areas[aid]
        wsum.append([a.get("region", ""), a.get("name", ""), s["status"], s.get("demand"), s.get("bought"), s.get("commodity"),
                     s.get("freight"), s.get("total"), s.get("per_gal"), s.get("unmet"), s.get("high_volume")])
        if not r["ok"]:
            continue
        ar = r["arrays"]
        for i, sp in enumerate(ar.supply):
            for what, v in (("Contract gal", ar.contract[i]), ("Bought gal", r["bought"][i]), ("Price $/gal", ar.price[i])):
                sheets["Supply plan"].append([a["name"], sp["location"], sp.get("supplier", ""), sp.get("hub", ""), what]
                                             + [float(x) for x in v] + ([float(v.sum())] if "gal" in what and "$" not in what else [None]))
        Sx = {x["id"]: x for x in ar.supply}
        Px = {x["id"]: x for x in ar.plants}
        for l, ln in enumerate(ar.lanes):
            if r["ship"][l].sum() < 1:
                continue
            sp, pl = Sx[ln["from"]], Px[ln["to"]]
            sheets["Lane volumes"].append([a["name"], sp["location"], sp.get("supplier", ""), pl["name"], "Shipped gal"]
                                          + [float(x) for x in r["ship"][l]] + [float(r["ship"][l].sum())])
            sheets["Lane volumes"].append([a["name"], sp["location"], sp.get("supplier", ""), pl["name"], "Freight $/gal"]
                                          + [None if not np.isfinite(x) else float(x) for x in ar.rate[l]] + [None])
        for j, pl in enumerate(ar.plants):
            for what, v in (("Planned demand gal", ar.demand[j]), ("Delivered gal", r["delivered"][j]),
                            ("Unmet gal", r["unmet"][j]), ("Storage at month end gal", r["p_end"][j])):
                sheets["Plants"].append([a["name"], pl["name"], pl.get("state", ""), what] + [float(x) for x in v]
                                        + [float(v.sum()) if "Storage" not in what else None])
    for ws in wb.worksheets:
        hr = 4 if ws.title == "Summary" else 1
        for c in ws[hr]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head
            c.alignment = Alignment(wrap_text=True, horizontal="center")
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 14
        ws.column_dimensions["B"].width = 30
        ws.freeze_panes = ws.cell(row=hr + 1, column=2)
        for row in ws.iter_rows(min_row=hr + 1):
            for c in row:
                if isinstance(c.value, float):
                    c.number_format = "0.0000" if abs(c.value) < 10 and c.value != 0 else "#,##0"
    data = io.BytesIO()
    wb.save(data)
    data.seek(0)
    return send_file(data, as_attachment=True, download_name=f"Supply plan results {_dt.date.today().isoformat()}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --------------------------------------------------------------------------- #
# what-if: try changes without saving them
# --------------------------------------------------------------------------- #

def _case(view: str, area: dict, base_bd: dict, c: dict) -> dict:
    """One what-if case -> solved summary."""
    t = c.get("type")
    P = len(area["months"])
    months = set(c.get("months") or range(1, 13))
    mask = np.array([E.month_num(m, p) in months for p, m in enumerate(area["months"])])
    planned = base_bd["planned"].copy()
    price_add, fm, a2 = None, 1.0, area
    if t == "hub":
        add = float(c.get("add", 0))
        price_add = np.zeros((len(area["supply"]), P))
        for i, s in enumerate(area["supply"]):
            if c.get("hub") in (None, "", "*") or s.get("hub") == c.get("hub"):
                price_add[i, mask] = add
    elif t == "freight":
        fm = 1.0 + float(c.get("pct", 0)) / 100.0
    elif t == "demand":
        planned[:, mask] *= 1.0 + float(c.get("pct", 0)) / 100.0
    elif t == "weather":
        settings = copy.deepcopy(STORE.demand_settings(view))
        settings["weather"] = {"mode": c.get("mode", "cold"), "months": sorted(months) if c.get("months") else W.WINTER,
                               "season": c.get("season", "")}
        planned = breakdown(view, area, settings=settings)["planned"]
    elif t == "contract":
        a2 = copy.deepcopy(area)
        for s in a2["supply"]:
            if c.get("supply") in (None, "", "*") or s["id"] == c.get("supply"):
                v = E.vec(s.get("contract"), P, 0.0)
                v[mask] *= 1.0 + float(c.get("pct", 0)) / 100.0
                s["contract"] = v.tolist()
    elif t == "lane_off":
        a2 = copy.deepcopy(area)
        a2["lanes"] = [l for l in a2["lanes"] if not (l["from"] == c.get("from") and l["to"] == c.get("to"))]
    else:
        raise ValueError("unknown what-if")
    r = solve_cached(a2, planned, price_add=price_add, freight_mult=fm, extra=["whatif", c])
    s = E.summary(r)
    return {k: (fnum(v) if isinstance(v, float) else v) for k, v in s.items()} | {
        "monthly": {k: jarr(v, 0) for k, v in E.monthly(r).items()} if r["ok"] else None}


@app.post("/api/whatif")
def whatif():
    b = body()
    view, aid = view_arg(b), b.get("area", "")
    area = STORE.area(view, aid)
    if area is None:
        raise S.NotFound("unknown area")
    bd = breakdown(view, area)
    base = run_view_area(view, aid)
    bs = E.summary(base)
    out = []
    for c in (b.get("cases") or [])[:12]:
        try:
            s = _case(view, area, bd, c)
        except ValueError as e:
            out.append({"case": c, "error": str(e)})
            continue
        d = (s.get("total") or 0) - (bs.get("total") or 0) if s.get("ok") and bs.get("ok") else None
        out.append({"case": c, "result": s, "change": fnum(d, 0),
                    "unmet_change": fnum((s.get("unmet") or 0) - (bs.get("unmet") or 0), 0)})
    return jsonify({"base": {k: (fnum(v) if isinstance(v, float) else v) for k, v in bs.items()},
                    "base_monthly": {k: jarr(v, 0) for k, v in E.monthly(base).items()} if base["ok"] else None,
                    "cases": out, "months": area["months"]})


# --------------------------------------------------------------------------- #
# compare two views (data set vs plan, plan vs plan)
# --------------------------------------------------------------------------- #

@app.post("/api/compare")
def compare():
    b = body()
    va, vb = b.get("a", ""), b.get("b", "")
    aid = b.get("area", "")
    A, B = STORE.area(va, aid), STORE.area(vb, aid)
    out = {"area": aid, "a": None, "b": None}
    ra = run_view_area(va, aid) if A else None
    rb = run_view_area(vb, aid) if B else None
    for side, r in (("a", ra), ("b", rb)):
        if r is None:
            continue
        s = E.summary(r)
        out[side] = {k: (fnum(v) if isinstance(v, float) else v) for k, v in s.items()} | {
            "monthly": {k: jarr(v, 0) for k, v in E.monthly(r).items()} if r["ok"] else None}
    # what's different in the inputs
    diffs = []
    if A and B:
        da, db = breakdown(va, A), breakdown(vb, B)
        diffs.append({"what": "Planned demand", "a": fnum(da["planned"].sum(), 0), "b": fnum(db["planned"].sum(), 0)})
        for coll, key, lab in (("supply", "contract", "Contract volume"),):
            ta = sum(float(np.nansum(E.vec(x.get(key), len(A["months"]), 0.0))) for x in A.get(coll, []))
            tb = sum(float(np.nansum(E.vec(x.get(key), len(B["months"]), 0.0))) for x in B.get(coll, []))
            diffs.append({"what": lab, "a": fnum(ta, 0), "b": fnum(tb, 0)})
        diffs.append({"what": "Lanes", "a": len(A.get("lanes", [])), "b": len(B.get("lanes", []))})
        if ra is not None and rb is not None and ra["ok"] and rb["ok"]:
            sa = {s["id"]: float(ra["bought"][i].sum()) for i, s in enumerate(ra["arrays"].supply)}
            sb = {s["id"]: float(rb["bought"][i].sum()) for i, s in enumerate(rb["arrays"].supply)}
            names_ = {s["id"]: s["location"] + (f" | {s['supplier']}" if s.get("supplier") else "") for s in A["supply"] + B["supply"]}
            moves = sorted(({"supply": names_.get(k, k), "a": fnum(sa.get(k, 0), 0), "b": fnum(sb.get(k, 0), 0),
                             "change": fnum(sb.get(k, 0) - sa.get(k, 0), 0)} for k in set(sa) | set(sb)),
                           key=lambda x: -abs(x["change"] or 0))
            out["supply_moves"] = [m for m in moves if abs(m["change"] or 0) >= 1][:15]
    out["inputs"] = diffs
    out["months"] = (A or B or {}).get("months", [])
    return jsonify(out)


# --------------------------------------------------------------------------- #

def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    logging.getLogger("werkzeug").setLevel(logging.WARNING)
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleTitleW("Supply Planner")
        except Exception:  # noqa: BLE001
            pass
    print(f"Supply Planner running at http://{HOST}:{PORT}/  (close this window to stop it)")
    print(f"Data folder: {DATA_DIR}")
    app.run(host=HOST, port=PORT, debug=False, threaded=True, use_reloader=False)


if __name__ == "__main__":
    main()
