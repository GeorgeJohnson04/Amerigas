"""
store.py - where the app keeps its data, and who may change what.

  data/
    weather.json                       the weather history (shared by everything)
    datasets/<id>/meta.json            a data set = one planning year's inputs
    datasets/<id>/areas/<area>.json    one file per area (supply, plants, lanes, hubs)
    datasets/<id>/tables.json          contract % by state + the demand settings
    datasets/<id>/log.json             every change: who, when, what (newest last)
    datasets/<id>/undo/<change>.json   what things looked like before an upload
    plans/<id>.json                    one person's plan on top of a data set
    plans/_undo/<plan>/<change>.json
    trash/                             deleted plans and data sets (recoverable)

DATA SETS are shared and anyone can change them - there are no locks. Every
change is logged with the person's name and can be undone.

PLANS belong to one person. Only the owner changes a plan; everyone else can
look at it or copy it. A plan starts as "the data set as it is" and only keeps
its own copy of an area once its owner changes something there, so a plan
shows exactly which areas it changed, and areas it didn't touch pick up later
corrections to the data set automatically.

A "view" is what the screen is showing: "ds:<id>" (the data set itself) or
"plan:<id>".
"""

from __future__ import annotations

import copy
import datetime as _dt
import json
import os
import re
import shutil
import tempfile
import threading
import time
import uuid
from typing import Dict, List, Optional, Tuple

import importer as I

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
LOG_KEEP = 1500                 # change-log entries kept per data set / plan
UNDO_KEEP = 60                  # upload snapshots kept per data set / plan
_ID = re.compile(r"^[a-z0-9][a-z0-9\-]{0,63}$")


class Denied(Exception):
    """The person can't change this (someone else's plan)."""


class NotFound(Exception):
    pass


def now_iso() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")


def slug(text: str, fallback: str = "x") -> str:
    s = re.sub(r"[^a-z0-9]+", "-", (text or "").lower()).strip("-")[:48]
    return s or fallback


def season_months(year_label: str) -> List[str]:
    """'2026-27' -> ['Apr-26', ..., 'Mar-27']."""
    m = re.search(r"(\d{4})", year_label or "")
    y = int(m.group(1)) if m else _dt.date.today().year
    out = []
    for k in range(12):
        mon = (3 + k) % 12 + 1
        yy = y if mon >= 4 else y + 1
        out.append(f"{MONTHS[mon - 1]}-{yy % 100:02d}")
    return out


def _write_json(path: str, obj) -> None:
    """Write via a temp file + rename, retrying while a file-sync client holds a lock."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, separators=(",", ":"))
    for attempt in range(12):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            time.sleep(0.1 * (attempt + 1))
    os.remove(tmp)
    raise RuntimeError(f"couldn't save {os.path.basename(path)} - is it open somewhere else?")


def _read_json(path: str, default=None):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except FileNotFoundError:
        return copy.deepcopy(default)


class Store:
    def __init__(self, root: str):
        self.root = root
        self.ds_dir = os.path.join(root, "datasets")
        self.plan_dir = os.path.join(root, "plans")
        self.trash = os.path.join(root, "trash")
        for d in (self.ds_dir, self.plan_dir, self.trash):
            os.makedirs(d, exist_ok=True)
        self.lock = threading.RLock()
        self._cache: Dict[str, Tuple[float, object]] = {}

    # ------------------------------------------------------------------ #
    # cached reads
    # ------------------------------------------------------------------ #

    def _load(self, path: str, default=None):
        try:
            mt = os.stat(path).st_mtime_ns
        except FileNotFoundError:
            return copy.deepcopy(default)
        hit = self._cache.get(path)
        if hit and hit[0] == mt:
            return hit[1]
        obj = _read_json(path, default)
        self._cache[path] = (mt, obj)
        return obj

    def _save(self, path: str, obj) -> None:
        _write_json(path, obj)
        self._cache.pop(path, None)

    # ------------------------------------------------------------------ #
    # weather (shared)
    # ------------------------------------------------------------------ #

    @property
    def weather_path(self) -> str:
        return os.path.join(self.root, "weather.json")

    def weather(self) -> dict:
        return self._load(self.weather_path, {}) or {}

    def weather_version(self) -> str:
        try:
            return str(os.stat(self.weather_path).st_mtime_ns)
        except FileNotFoundError:
            return "0"

    # ------------------------------------------------------------------ #
    # data sets
    # ------------------------------------------------------------------ #

    def _ds(self, ds: str, *parts) -> str:
        if not _ID.match(ds or ""):
            raise NotFound("unknown data set")
        return os.path.join(self.ds_dir, ds, *parts)

    def list_datasets(self) -> List[dict]:
        out = []
        for d in sorted(os.listdir(self.ds_dir)):
            meta = self._load(os.path.join(self.ds_dir, d, "meta.json"))
            if not meta:
                continue
            m = dict(meta)
            m["areas"] = len([f for f in os.listdir(os.path.join(self.ds_dir, d, "areas")) if f.endswith(".json")]) \
                if os.path.isdir(os.path.join(self.ds_dir, d, "areas")) else 0
            out.append(m)
        return sorted(out, key=lambda m: (m.get("year", ""), m.get("created", "")), reverse=True)

    def meta(self, ds: str) -> dict:
        m = self._load(self._ds(ds, "meta.json"))
        if not m:
            raise NotFound("unknown data set")
        return m

    def create_dataset(self, name: str, year: str, user: str, areas: Optional[Dict[str, dict]] = None,
                       state_pct: Optional[dict] = None, source: str = "", note: str = "") -> str:
        with self.lock:
            base = slug(year or name, "data")
            ds, n = base, 2
            while os.path.exists(os.path.join(self.ds_dir, ds)):
                ds, n = f"{base}-{n}", n + 1
            months = None
            for a in (areas or {}).values():
                months = a.get("months")
                break
            months = months or season_months(year or name)
            meta = {"id": ds, "name": name or year, "year": year, "months": months, "created": now_iso(),
                    "by": user, "source": source, "note": note}
            self._save(self._ds(ds, "meta.json"), meta)
            os.makedirs(self._ds(ds, "areas"), exist_ok=True)
            for a in (areas or {}).values():
                self._save(self._ds(ds, "areas", f"{a['id']}.json"), a)
            self._save(self._ds(ds, "tables.json"), {"state_pct": state_pct or {}, "demand": default_demand()})
            self._log_ds(ds, {"kind": "create", "user": user, "label": f"Created data set \u201c{meta['name']}\u201d"
                              + (f" from {source}" if source else "")})
            return ds

    def copy_dataset(self, src: str, name: str, year: str, user: str) -> str:
        areas = self.ds_areas(src)
        t = self.ds_tables(src)
        return self.create_dataset(name, year or self.meta(src).get("year", ""), user, copy.deepcopy(areas),
                                   copy.deepcopy(t.get("state_pct") or {}), source=f"a copy of \u201c{self.meta(src)['name']}\u201d")

    def rename_dataset(self, ds: str, name: str, user: str) -> None:
        with self.lock:
            m = dict(self.meta(ds))
            old, m["name"] = m["name"], name
            self._save(self._ds(ds, "meta.json"), m)
            self._log_ds(ds, {"kind": "rename", "user": user, "label": f"Renamed data set \u201c{old}\u201d to \u201c{name}\u201d"})

    def delete_dataset(self, ds: str, user: str) -> None:
        with self.lock:
            src = self._ds(ds)
            if not os.path.isdir(src):
                raise NotFound("unknown data set")
            dst = os.path.join(self.trash, f"dataset-{ds}-{int(time.time())}")
            shutil.move(src, dst)

    def ds_area_ids(self, ds: str) -> List[str]:
        d = self._ds(ds, "areas")
        if not os.path.isdir(d):
            return []
        return sorted(f[:-5] for f in os.listdir(d) if f.endswith(".json"))

    def ds_area(self, ds: str, aid: str) -> Optional[dict]:
        if not _ID.match(aid or ""):
            return None
        return self._load(self._ds(ds, "areas", f"{aid}.json"))

    def ds_areas(self, ds: str) -> Dict[str, dict]:
        out = {}
        for aid in self.ds_area_ids(ds):
            a = self.ds_area(ds, aid)
            if a:
                out[aid] = a
        return out

    def ds_tables(self, ds: str) -> dict:
        t = self._load(self._ds(ds, "tables.json"), {}) or {}
        t.setdefault("state_pct", {})
        t.setdefault("demand", default_demand())
        return t

    def _save_ds_area(self, ds: str, area: dict) -> None:
        self._save(self._ds(ds, "areas", f"{area['id']}.json"), area)

    def _log_ds(self, ds: str, entry: dict) -> dict:
        path = self._ds(ds, "log.json")
        log = list(self._load(path, []) or [])
        entry = {"id": uuid.uuid4().hex[:10], "ts": now_iso(), **entry}
        log.append(entry)
        if len(log) > LOG_KEEP:
            log = log[-LOG_KEEP:]
        self._save(path, log)
        return entry

    # ------------------------------------------------------------------ #
    # plans
    # ------------------------------------------------------------------ #

    def _plan_path(self, pid: str) -> str:
        if not _ID.match(pid or ""):
            raise NotFound("unknown plan")
        return os.path.join(self.plan_dir, f"{pid}.json")

    def plan(self, pid: str) -> dict:
        p = self._load(self._plan_path(pid))
        if not p:
            raise NotFound("unknown plan")
        return p

    def list_plans(self, ds: Optional[str] = None) -> List[dict]:
        out = []
        for f in sorted(os.listdir(self.plan_dir)):
            if not f.endswith(".json"):
                continue
            p = self._load(os.path.join(self.plan_dir, f))
            if not p or (ds and p.get("dataset") != ds):
                continue
            out.append({k: p.get(k) for k in ("id", "name", "owner", "dataset", "created", "updated", "note")}
                       | {"areas_changed": sorted(p.get("areas", {})), "changes": len(p.get("changes", [])),
                          "weather": (p.get("demand") or {}).get("weather", {}).get("mode", "normal"),
                          "rules": len((p.get("demand") or {}).get("adjust", []))})
        return sorted(out, key=lambda p: p.get("updated") or "", reverse=True)

    def create_plan(self, ds: str, name: str, owner: str, copy_from: Optional[str] = None, note: str = "") -> dict:
        if not owner:
            raise Denied("enter your name first (top right) so the plan is yours")
        self.meta(ds)
        with self.lock:
            pid = f"{slug(owner, 'me')[:20]}-{uuid.uuid4().hex[:6]}"
            if copy_from:
                src = self.plan(copy_from)
                body = {"areas": copy.deepcopy(src.get("areas", {})), "state_pct": copy.deepcopy(src.get("state_pct", {})),
                        "demand": copy.deepcopy(src.get("demand") or self.ds_tables(ds)["demand"])}
                origin = f"a copy of {src['owner']}\u2019s \u201c{src['name']}\u201d"
                ds = src["dataset"]
            else:
                body = {"areas": {}, "state_pct": {}, "demand": copy.deepcopy(self.ds_tables(ds)["demand"])}
                origin = f"the data set \u201c{self.meta(ds)['name']}\u201d"
            p = {"id": pid, "name": name or "My plan", "owner": owner, "dataset": ds, "created": now_iso(),
                 "updated": now_iso(), "note": note, "origin": origin, **body, "changes": []}
            self._plan_log(p, {"kind": "create", "user": owner, "label": f"Started from {origin}"})
            self._save(self._plan_path(pid), p)
            return p

    def save_plan(self, p: dict) -> None:
        p["updated"] = now_iso()
        self._save(self._plan_path(p["id"]), p)

    def delete_plan(self, pid: str, user: str) -> None:
        with self.lock:
            p = self.plan(pid)
            if p["owner"] != user:
                raise Denied(f"only {p['owner']} can delete this plan")
            shutil.move(self._plan_path(pid), os.path.join(self.trash, f"plan-{pid}-{int(time.time())}.json"))
            self._cache.pop(self._plan_path(pid), None)

    def rename_plan(self, pid: str, user: str, name: str, note: Optional[str] = None) -> None:
        with self.lock:
            p = copy.deepcopy(self.plan(pid))
            if p["owner"] != user:
                raise Denied(f"only {p['owner']} can rename this plan")
            p["name"] = name or p["name"]
            if note is not None:
                p["note"] = note
            self.save_plan(p)

    def _plan_log(self, p: dict, entry: dict) -> dict:
        entry = {"id": uuid.uuid4().hex[:10], "ts": now_iso(), **entry}
        p.setdefault("changes", []).append(entry)
        if len(p["changes"]) > LOG_KEEP:
            p["changes"] = p["changes"][-LOG_KEEP:]
        return entry

    # ------------------------------------------------------------------ #
    # views
    # ------------------------------------------------------------------ #

    def parse_view(self, view: str) -> Tuple[str, Optional[dict]]:
        """-> (data set id, plan or None)."""
        kind, _, ident = (view or "").partition(":")
        if kind == "ds":
            self.meta(ident)
            return ident, None
        if kind == "plan":
            p = self.plan(ident)
            return p["dataset"], p
        raise NotFound("pick a data set or a plan")

    def can_edit(self, view: str, user: str) -> bool:
        _, p = self.parse_view(view)
        return bool(user) and (p is None or p["owner"] == user)

    def _need_edit(self, view: str, user: str) -> Tuple[str, Optional[dict]]:
        ds, p = self.parse_view(view)
        if not user:
            raise Denied("enter your name first (top right) - every change is signed")
        if p is not None and p["owner"] != user:
            raise Denied(f"this plan belongs to {p['owner']}. Make a copy to change it")
        return ds, (copy.deepcopy(p) if p is not None else None)

    def area_ids(self, view: str) -> List[str]:
        ds, p = self.parse_view(view)
        ids = set(self.ds_area_ids(ds))
        if p:
            ids |= set(p.get("areas", {}))
        return sorted(ids)

    def area(self, view: str, aid: str) -> Optional[dict]:
        ds, p = self.parse_view(view)
        if p and aid in p.get("areas", {}):
            return p["areas"][aid]
        return self.ds_area(ds, aid)

    def areas(self, view: str) -> Dict[str, dict]:
        ds, p = self.parse_view(view)
        out = dict(self.ds_areas(ds))
        if p:
            out.update(p.get("areas", {}))
        return out

    def state_pct(self, view: str) -> dict:
        ds, p = self.parse_view(view)
        base = copy.deepcopy(self.ds_tables(ds)["state_pct"])
        if p:
            for st, by in (p.get("state_pct") or {}).items():
                base.setdefault(st, {}).update(by)
        return base

    def demand_settings(self, view: str) -> dict:
        ds, p = self.parse_view(view)
        d = (p.get("demand") if p else None) or self.ds_tables(ds)["demand"]
        return normal_demand(d)

    def months(self, view: str) -> List[str]:
        ds, _ = self.parse_view(view)
        return list(self.meta(ds).get("months") or season_months(self.meta(ds).get("year", "")))

    def changes(self, view: str) -> List[dict]:
        ds, p = self.parse_view(view)
        log = p.get("changes", []) if p else (self._load(self._ds(ds, "log.json"), []) or [])
        return list(reversed(log))

    def plan_diff_areas(self, view: str) -> List[str]:
        _, p = self.parse_view(view)
        return sorted((p or {}).get("areas", {}))

    # ------------------------------------------------------------------ #
    # changing things
    # ------------------------------------------------------------------ #

    def _undo_dir(self, ds: str, p: Optional[dict]) -> str:
        return os.path.join(self.plan_dir, "_undo", p["id"]) if p else self._ds(ds, "undo")

    def _snapshot(self, ds: str, p: Optional[dict], cid: str, aids, tables: bool = False, weather: bool = False) -> None:
        """Before-picture of areas (None = didn't exist / plan had no copy), state %, weather."""
        snap: dict = {"areas": {}}
        for aid in aids:
            if p is not None:
                snap["areas"][aid] = copy.deepcopy(p.get("areas", {}).get(aid))
            else:
                snap["areas"][aid] = copy.deepcopy(self.ds_area(ds, aid))
        if tables:
            snap["state_pct"] = copy.deepcopy(p.get("state_pct", {}) if p is not None else self.ds_tables(ds)["state_pct"])
        if weather:
            snap["weather"] = copy.deepcopy(self.weather())
        d = self._undo_dir(ds, p)
        _write_json(os.path.join(d, f"{cid}.json"), snap)
        files = sorted((os.path.join(d, f) for f in os.listdir(d)), key=os.path.getmtime)
        for f in files[:-UNDO_KEEP]:
            try:
                os.remove(f)
            except OSError:
                pass

    def _put_area(self, ds: str, p: Optional[dict], area: dict) -> None:
        if p is not None:
            p.setdefault("areas", {})[area["id"]] = area
        else:
            self._save_ds_area(ds, area)

    def _drop_area(self, ds: str, p: Optional[dict], aid: str) -> None:
        if p is not None:
            p.get("areas", {}).pop(aid, None)
        else:
            try:
                os.remove(self._ds(ds, "areas", f"{aid}.json"))
            except FileNotFoundError:
                pass

    def _finish(self, ds: str, p: Optional[dict], entry: dict) -> dict:
        if p is not None:
            e = self._plan_log(p, entry)
            self.save_plan(p)
            return e
        return self._log_ds(ds, entry)

    def _working_area(self, ds: str, p: Optional[dict], aid: str) -> dict:
        a = (p.get("areas", {}).get(aid) if p is not None else None) or self.ds_area(ds, aid)
        if a is None:
            raise NotFound("unknown area")
        return copy.deepcopy(a)

    def edit(self, view: str, aid: str, items: List[dict], user: str, label: str = "") -> dict:
        """Change cells: items = [{path, value}] (paths as importer.get_path)."""
        with self.lock:
            ds, p = self._need_edit(view, user)
            area = self._working_area(ds, p, aid)
            done = []
            for it in items[:20000]:
                path = list(it["path"])
                old = I.get_path(area, path)
                new = it.get("value")
                if old == new:
                    continue
                I.set_path(area, path, new)
                done.append({"path": path, "old": old, "new": new})
            if not done:
                return {"id": None, "count": 0}
            self._put_area(ds, p, area)
            e = self._finish(ds, p, {"kind": "edit", "user": user, "area": aid, "label": label or describe_edit(area, done),
                                     "items": done[:4000], "count": len(done)})
            return {**e, "count": len(done)}

    def add(self, view: str, aid: str, kind: str, obj: dict, user: str) -> dict:
        """Add a supply point, plant, lane or hub to an area."""
        with self.lock:
            ds, p = self._need_edit(view, user)
            area = self._working_area(ds, p, aid)
            P = len(area["months"])
            cid = uuid.uuid4().hex[:10]
            if kind == "supply":
                o = {"id": I.new_id("s"), "location": obj.get("location", "").strip() or "New supply point",
                     "supplier": obj.get("supplier", "").strip(), "hub": obj.get("hub", "").strip(),
                     "contract": [0.0] * P, "diff": [0.0] * P, "min_pct": [0.0] * P, "max_pct": [1.0] * P,
                     "open": 0.0, "cap": [0.0] * P, "inv_min": [0.0] * P}
                area["supply"].append(o)
                label = f"Added supply point {o['location']}"
            elif kind == "plant":
                o = {"id": I.new_id("p"), "name": obj.get("name", "").strip() or "New plant",
                     "state": (obj.get("state") or "").strip().upper()[:4],
                     "budget": [0.0] * P, "open": 0.0, "cap": [0.0] * P, "inv_min": [0.0] * P}
                area["plants"].append(o)
                label = f"Added plant {o['name']}"
            elif kind == "lane":
                frm, to = obj.get("from"), obj.get("to")
                S = {s["id"]: s for s in area["supply"]}
                Pl = {x["id"]: x for x in area["plants"]}
                if frm not in S or to not in Pl:
                    raise NotFound("pick a supply point and a plant")
                if I.find(area, "lanes", (frm, to)):
                    raise ValueError("that lane already exists")
                rate = obj.get("rate")
                rate = None if rate in (None, "") else float(rate)
                o = {"from": frm, "to": to, "rate": [rate] * P, "min": None, "max": None}
                area.setdefault("lanes", []).append(o)
                label = f"Added lane {S[frm]['location']} \u2192 {Pl[to]['name']}"
            elif kind == "hub":
                code = (obj.get("code") or "").strip()
                if not code:
                    raise ValueError("give the hub a code, e.g. CON")
                if any(h["code"].lower() == code.lower() for h in area.get("hubs", [])):
                    raise ValueError("that hub already exists")
                pr = obj.get("price")
                o = {"code": code, "name": obj.get("name", ""), "price": [None if pr in (None, "") else float(pr)] * P}
                area.setdefault("hubs", []).append(o)
                label = f"Added hub {code}"
            else:
                raise ValueError("unknown kind")
            self._snapshot(ds, p, cid, [aid])
            self._put_area(ds, p, area)
            e = self._finish(ds, p, {"id": cid, "kind": "add", "user": user, "area": aid, "label": label, "undo": True})
            return {**e, "obj": o}

    def remove(self, view: str, aid: str, kind: str, ident, user: str) -> dict:
        with self.lock:
            ds, p = self._need_edit(view, user)
            area = self._working_area(ds, p, aid)
            cid = uuid.uuid4().hex[:10]
            if kind == "supply":
                s = I.find(area, "supply", ident)
                if not s:
                    raise NotFound("not found")
                area["supply"] = [x for x in area["supply"] if x["id"] != ident]
                area["lanes"] = [l for l in area.get("lanes", []) if l["from"] != ident]
                label = f"Removed supply point {s['location']} and its lanes"
            elif kind == "plant":
                x = I.find(area, "plants", ident)
                if not x:
                    raise NotFound("not found")
                area["plants"] = [y for y in area["plants"] if y["id"] != ident]
                area["lanes"] = [l for l in area.get("lanes", []) if l["to"] != ident]
                label = f"Removed plant {x['name']} and its lanes"
            elif kind == "lane":
                frm, to = ident
                before = len(area.get("lanes", []))
                area["lanes"] = [l for l in area.get("lanes", []) if not (l["from"] == frm and l["to"] == to)]
                if len(area["lanes"]) == before:
                    raise NotFound("not found")
                label = "Removed a lane"
            elif kind == "hub":
                area["hubs"] = [h for h in area.get("hubs", []) if h["code"] != ident]
                label = f"Removed hub {ident}"
            else:
                raise ValueError("unknown kind")
            self._snapshot(ds, p, cid, [aid])
            self._put_area(ds, p, area)
            return self._finish(ds, p, {"id": cid, "kind": "remove", "user": user, "area": aid, "label": label, "undo": True})

    def add_area(self, view: str, name: str, region: str, user: str, title: str = "") -> dict:
        with self.lock:
            ds, p = self._need_edit(view, user)
            aid, n = slug(name, "area"), 2
            existing = set(self.area_ids(view))
            base = aid
            while aid in existing:
                aid, n = f"{base}-{n}", n + 1
            area = {"id": aid, "name": name.strip() or "New area", "title": title or name, "region": region.strip() or "Other",
                    "months": self.months(view), "hubs": [], "supply": [], "plants": [], "lanes": []}
            cid = uuid.uuid4().hex[:10]
            self._snapshot(ds, p, cid, [aid])
            self._put_area(ds, p, area)
            self._finish(ds, p, {"id": cid, "kind": "add", "user": user, "area": aid, "label": f"Added area {area['name']}", "undo": True})
            return area

    def update_area_info(self, view: str, aid: str, user: str, **fields) -> dict:
        with self.lock:
            ds, p = self._need_edit(view, user)
            area = self._working_area(ds, p, aid)
            done = []
            for k in ("name", "title", "region"):
                if fields.get(k) is not None and fields[k] != area.get(k):
                    done.append({"path": [k], "old": area.get(k), "new": fields[k]})
                    area[k] = fields[k]
            if not done:
                return {"count": 0}
            self._put_area(ds, p, area)
            return self._finish(ds, p, {"kind": "edit", "user": user, "area": aid, "label": "Renamed / moved the area",
                                        "items": done, "count": len(done), "area_info": True})

    def delete_area(self, view: str, aid: str, user: str) -> dict:
        with self.lock:
            ds, p = self._need_edit(view, user)
            if p is not None:
                raise Denied("areas are removed from the data set, not from a plan")
            a = self.ds_area(ds, aid)
            if not a:
                raise NotFound("unknown area")
            cid = uuid.uuid4().hex[:10]
            self._snapshot(ds, p, cid, [aid])
            self._drop_area(ds, p, aid)
            return self._finish(ds, p, {"id": cid, "kind": "remove", "user": user, "area": aid,
                                        "label": f"Removed area {a['name']}", "undo": True})

    def reset_area(self, view: str, aid: str, user: str) -> dict:
        """Plan only: throw away the plan's changes to one area (back to the data set)."""
        with self.lock:
            ds, p = self._need_edit(view, user)
            if p is None:
                raise ValueError("only a plan can be reset to the data set")
            if aid not in p.get("areas", {}):
                return {"count": 0}
            cid = uuid.uuid4().hex[:10]
            self._snapshot(ds, p, cid, [aid])
            name = p["areas"][aid].get("name", aid)
            self._drop_area(ds, p, aid)
            return self._finish(ds, p, {"id": cid, "kind": "reset", "user": user, "area": aid,
                                        "label": f"Put {name} back to the data set", "undo": True})

    def set_state_pct(self, view: str, items: List[dict], user: str, label: str = "") -> dict:
        """items = [{state, month ('Jan'..), value (fraction or None to clear)}]."""
        with self.lock:
            ds, p = self._need_edit(view, user)
            cur_all = self.state_pct(view)
            done = []
            target = p.setdefault("state_pct", {}) if p is not None else copy.deepcopy(self.ds_tables(ds)["state_pct"])
            for it in items:
                st, mon, v = str(it["state"]).upper(), it["month"], it.get("value")
                if mon not in MONTHS:
                    continue
                old = (cur_all.get(st) or {}).get(mon)
                if old == v:
                    continue
                if v is None:
                    (target.get(st) or {}).pop(mon, None)
                else:
                    target.setdefault(st, {})[mon] = float(v)
                done.append({"state": st, "month": mon, "old": old, "new": v})
            if not done:
                return {"count": 0}
            if p is None:
                t = copy.deepcopy(self.ds_tables(ds))
                t["state_pct"] = target
                self._save(self._ds(ds, "tables.json"), t)
            return {**self._finish(ds, p, {"kind": "state_pct", "user": user, "label": label or
                                           f"Contract % by state: {len(done)} cell(s) changed", "pct_items": done,
                                           "count": len(done)}), "count": len(done)}

    def set_demand(self, view: str, settings: dict, user: str, label: str = "") -> dict:
        """Weather scenario + demand adjustment rules."""
        with self.lock:
            ds, p = self._need_edit(view, user)
            new = normal_demand(settings)
            old = self.demand_settings(view)
            if new == old:
                return {"count": 0}
            if p is not None:
                p["demand"] = new
            else:
                t = copy.deepcopy(self.ds_tables(ds))
                t["demand"] = new
                self._save(self._ds(ds, "tables.json"), t)
            return self._finish(ds, p, {"kind": "demand", "user": user, "label": label or "Changed demand settings",
                                        "old_demand": old, "new_demand": new, "count": 1})

    def apply_import(self, view: str, ops: List[dict], user: str, label: str) -> dict:
        """Carry out an upload's operations; snapshot first so it can be undone."""
        with self.lock:
            ds, p = self._need_edit(view, user)
            ctx = {"areas": copy.deepcopy(self.areas(view)), "state_pct": self.state_pct(view),
                   "months": self.months(view), "weather": copy.deepcopy(self.weather())}
            before_pct = copy.deepcopy(ctx["state_pct"])
            touched = {op["area"] for op in ops if op.get("area")}
            has_pct = any(op["op"] == "state_pct" for op in ops)
            has_w = any(op["op"] == "weather" for op in ops)
            cid = uuid.uuid4().hex[:10]
            self._snapshot(ds, p, cid, sorted(touched), tables=has_pct, weather=has_w)
            counts = I.apply_ops(ctx, ops, user, now_iso())
            for aid in sorted(touched):
                if aid in ctx["areas"]:
                    self._put_area(ds, p, ctx["areas"][aid])
            if has_pct:
                if p is not None:
                    base = self.ds_tables(ds)["state_pct"]
                    over = {}
                    for st, by in ctx["state_pct"].items():
                        for mon, v in by.items():
                            if (base.get(st) or {}).get(mon) != v:
                                over.setdefault(st, {})[mon] = v
                    p["state_pct"] = over
                else:
                    t = copy.deepcopy(self.ds_tables(ds))
                    t["state_pct"] = ctx["state_pct"]
                    self._save(self._ds(ds, "tables.json"), t)
            if has_w:
                w = ctx["weather"]
                w["source"] = label
                self._save(self.weather_path, w)
            total = sum(counts.values()) + sum(1 for op in ops if op["op"] in ("create", "create_area"))
            e = self._finish(ds, p, {"id": cid, "kind": "upload", "user": user, "label": label, "undo": True,
                                     "areas": sorted(touched), "count": total, "weather": has_w, "state_pct": has_pct})
            if has_w and p is not None:
                # weather is shared - note it on the data set's log too
                self._log_ds(ds, {"kind": "upload", "user": user, "label": f"{label} (weather history, shared)", "weather": True})
            _ = before_pct
            return e

    def undo(self, view: str, cid: str, user: str) -> dict:
        """Put things back the way they were before change `cid`."""
        with self.lock:
            ds, p = self._need_edit(view, user)
            log = p.get("changes", []) if p is not None else (self._load(self._ds(ds, "log.json"), []) or [])
            entry = next((e for e in log if e.get("id") == cid), None)
            if entry is None:
                raise NotFound("that change isn't in the log any more")
            if entry.get("undone"):
                raise ValueError("that change was already undone")
            kind = entry.get("kind")
            if kind == "edit":
                aid = entry["area"]
                area = self._working_area(ds, p, aid)
                for it in reversed(entry.get("items", [])):
                    if entry.get("area_info"):
                        area[it["path"][0]] = it["old"]
                    else:
                        try:
                            I.set_path(area, it["path"], it["old"])
                        except KeyError:
                            pass
                self._put_area(ds, p, area)
            elif kind == "state_pct":
                self.set_state_pct(view, [{"state": i["state"], "month": i["month"], "value": i["old"]}
                                          for i in entry.get("pct_items", [])], user, label="(undo)")
                if p is not None:
                    p = copy.deepcopy(self.plan(p["id"]))
                    log = p.get("changes", [])
            elif kind == "demand":
                old = entry.get("old_demand") or default_demand()
                if p is not None:
                    p["demand"] = old
                else:
                    t = copy.deepcopy(self.ds_tables(ds))
                    t["demand"] = old
                    self._save(self._ds(ds, "tables.json"), t)
            elif entry.get("undo"):
                path = os.path.join(self._undo_dir(ds, p), f"{cid}.json")
                snap = _read_json(path)
                if snap is None:
                    raise ValueError("this change is too old to undo automatically")
                for aid, a in snap.get("areas", {}).items():
                    if a is None:
                        self._drop_area(ds, p, aid)
                    else:
                        self._put_area(ds, p, a)
                if "state_pct" in snap:
                    if p is not None:
                        p["state_pct"] = snap["state_pct"]
                    else:
                        t = copy.deepcopy(self.ds_tables(ds))
                        t["state_pct"] = snap["state_pct"]
                        self._save(self._ds(ds, "tables.json"), t)
                if "weather" in snap:
                    self._save(self.weather_path, snap["weather"] or {})
            else:
                raise ValueError("this kind of change can't be undone")
            # mark it
            if p is not None:
                for e in p.get("changes", []):
                    if e.get("id") == cid:
                        e["undone"] = True
                return self._finish(ds, p, {"kind": "undo", "user": user, "label": f"Undid: {entry.get('label', '')}",
                                            "area": entry.get("area")})
            log = list(self._load(self._ds(ds, "log.json"), []) or [])
            for e in log:
                if e.get("id") == cid:
                    e["undone"] = True
            self._save(self._ds(ds, "log.json"), log)
            return self._log_ds(ds, {"kind": "undo", "user": user, "label": f"Undid: {entry.get('label', '')}",
                                     "area": entry.get("area")})

    def promote_area(self, view: str, aid: str, user: str) -> dict:
        """Copy a plan's version of an area into the data set (anyone may change the data set)."""
        with self.lock:
            ds, p = self.parse_view(view)
            if p is None or aid not in p.get("areas", {}):
                raise ValueError("this plan hasn't changed that area")
            if not user:
                raise Denied("enter your name first")
            cid = uuid.uuid4().hex[:10]
            self._snapshot(ds, None, cid, [aid])
            self._save_ds_area(ds, copy.deepcopy(p["areas"][aid]))
            return self._log_ds(ds, {"id": cid, "kind": "upload", "user": user, "area": aid, "undo": True,
                                     "label": f"Copied {p['areas'][aid].get('name', aid)} from {p['owner']}\u2019s plan \u201c{p['name']}\u201d into the data set"})


# --------------------------------------------------------------------------- #

def default_demand() -> dict:
    return {"weather": {"mode": "normal", "months": [10, 11, 12, 1, 2, 3], "season": ""}, "adjust": []}


def normal_demand(d: Optional[dict]) -> dict:
    d = d or {}
    w = dict(d.get("weather") or {})
    mode = w.get("mode", "normal")
    out_w = {"mode": mode if mode in ("normal", "cold", "warm", "coldest", "warmest", "season") else "normal",
             "months": sorted({int(m) for m in (w.get("months") or [10, 11, 12, 1, 2, 3]) if 1 <= int(m) <= 12}),
             "season": str(w.get("season") or "")}
    rules = []
    for r in d.get("adjust") or []:
        try:
            pct = float(r.get("pct", 0))
        except (TypeError, ValueError):
            continue
        months = sorted({int(m) for m in (r.get("months") or []) if 1 <= int(m) <= 12})
        if not months or not pct:
            continue
        sc = r.get("scope") or {"type": "all"}
        rules.append({"id": r.get("id") or uuid.uuid4().hex[:6], "months": months, "pct": pct,
                      "scope": {"type": sc.get("type", "all") if sc.get("type") in ("all", "states", "plants") else "all",
                                "values": list(sc.get("values") or [])},
                      "area": r.get("area") or "*", "label": str(r.get("label") or "")[:120]})
    return {"weather": out_w, "adjust": rules}


def describe_edit(area: dict, done: List[dict]) -> str:
    fields = {"budget": "budget demand", "contract": "contract", "diff": "differential", "min_pct": "min lift %",
              "max_pct": "max lift %", "cap": "storage capacity", "inv_min": "minimum inventory", "open": "opening inventory",
              "rate": "freight", "min": "lane minimum", "max": "lane maximum", "price": "hub price", "name": "name",
              "state": "state", "location": "location", "supplier": "supplier", "hub": "hub"}
    kinds = sorted({fields.get(d["path"][3] if d["path"][0] == "lanes" else d["path"][2], "value") for d in done
                    if len(d["path"]) >= 3})
    return f"{area.get('name', '')}: {len(done)} cell(s) of {', '.join(kinds) or 'data'} changed"
