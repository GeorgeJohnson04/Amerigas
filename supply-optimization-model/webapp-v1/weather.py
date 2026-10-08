"""
weather.py - weather-driven demand from a 20-year history.

The history is uploaded like any other data: one row per place and season,
twelve monthly values (gallons delivered, or heating degree days). "Place" can
be a plant, a state, an area, a region, or "All" for the whole company - the
most specific one that exists is used for each plant.

From it we work out a WEATHER FACTOR per place, season and month: how that
month compared with normal. 1.10 means 10% more demand than normal.

  * Degree days: factor = this month's degree days / the 20-year average.
  * Gallons: gallons also grow (or shrink) with the business, so the growth
    trend is taken out first - factor = actual / trend line for that month.

A plan picks a weather scenario, and only for the months it chooses
(default: winter, Oct-Mar):

  normal     factor 1.0
  cold       each month as cold as a 1-in-10 cold month (90th percentile)
  warm       each month as mild as a 1-in-10 mild month (10th percentile)
  coldest    replay the coldest winter in the history
  warmest    replay the mildest winter in the history
  season     replay a season you pick (e.g. "2013-14")
"""

from __future__ import annotations

import math
import re
from typing import Callable, Dict, List, Optional

import numpy as np

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
SEASON_ORDER = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
WINTER = [10, 11, 12, 1, 2, 3]
MODES = {
    "normal": "Normal weather",
    "cold": "Cold (1-in-10 cold months)",
    "warm": "Mild (1-in-10 mild months)",
    "coldest": "Coldest winter on record",
    "warmest": "Mildest winter on record",
    "season": "Repeat a past season",
}
ALL_KEYS = {"all", "total", "company", "companytotal", "national", "us"}


def fold(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def season_label(text) -> str:
    """'2006-07', '2006/2007', 2006, '06-07', '2006-2007' -> '2006-07'."""
    nums = re.findall(r"\d+", str(text or ""))
    if not nums:
        return ""
    a = int(nums[0])
    if a < 100:
        a += 2000 if a < 70 else 1900
    return f"{a}-{(a + 1) % 100:02d}"


def _detrend(values: np.ndarray) -> np.ndarray:
    """Trend line through one month's values across seasons (NaN-safe)."""
    x = np.arange(len(values), dtype=float)
    ok = np.isfinite(values) & (values > 0)
    if ok.sum() < 5:
        mean = np.nanmean(np.where(ok, values, np.nan)) if ok.any() else np.nan
        return np.full(len(values), mean)
    slope, icept = np.polyfit(x[ok], values[ok], 1)
    fit = icept + slope * x
    mean = values[ok].mean()
    return np.where(fit > 0.2 * mean, fit, mean)      # never let a trend line go silly


class History:
    """A parsed weather history with factors per place."""

    def __init__(self, data: Optional[dict]):
        data = data or {}
        self.data = data
        self.unit = data.get("unit", "gal")
        self.detrend = bool(data.get("detrend", self.unit == "gal"))
        rows = data.get("rows") or []
        self.places: Dict[str, Dict[str, Dict[str, float]]] = {}   # place -> season -> Mon -> value
        self.label: Dict[str, str] = {}
        for r in rows:
            key, season = str(r.get("key") or "").strip(), season_label(r.get("season"))
            if not key or not season:
                continue
            f = fold(key)
            self.label.setdefault(f, key)
            vals = self.places.setdefault(f, {}).setdefault(season, {})
            for mon, v in (r.get("values") or {}).items():
                try:
                    if v is not None and v != "" and math.isfinite(float(v)):
                        vals[mon] = float(v)
                except (TypeError, ValueError):
                    pass
        self.seasons = sorted({s for p in self.places.values() for s in p})
        self._factors: Dict[str, Dict[str, Dict[str, float]]] = {}

    @property
    def empty(self) -> bool:
        return not self.places

    def factors(self, place: str) -> Dict[str, Dict[str, float]]:
        """season -> Mon -> factor for one place (folded key)."""
        if place in self._factors:
            return self._factors[place]
        hist = self.places.get(place, {})
        seasons = sorted(hist)
        out: Dict[str, Dict[str, float]] = {s: {} for s in seasons}
        for mon in MONTHS:
            v = np.array([hist[s].get(mon, np.nan) for s in seasons], dtype=float)
            if not np.isfinite(v).any():
                continue
            base = _detrend(v) if self.detrend else np.full(len(v), np.nanmean(np.where(v > 0, v, np.nan)))
            with np.errstate(divide="ignore", invalid="ignore"):
                f = v / base
            for s, x in zip(seasons, f):
                if np.isfinite(x):
                    out[s][mon] = float(min(max(x, 0.2), 4.0))
        self._factors[place] = out
        return out

    def winter_index(self, place: str) -> Dict[str, float]:
        """Season -> average winter factor (above 1 = colder than normal)."""
        out = {}
        for s, fm in self.factors(place).items():
            vals = [fm[MONTHS[m - 1]] for m in WINTER if MONTHS[m - 1] in fm]
            if vals:
                out[s] = float(np.mean(vals))
        return out

    def scenario_factor(self, place: str, mode: str, mon: int, season: str = "") -> Optional[float]:
        fac = self.factors(place)
        if not fac:
            return None
        name = MONTHS[mon - 1]
        if mode == "normal":
            return 1.0
        if mode in ("cold", "warm"):
            vals = [fm[name] for fm in fac.values() if name in fm]
            if not vals:
                return None
            return float(np.percentile(vals, 90 if mode == "cold" else 10))
        if mode in ("coldest", "warmest"):
            wi = self.winter_index(place)
            if not wi:
                return None
            season = (max if mode == "coldest" else min)(wi, key=wi.get)
        return fac.get(season_label(season), {}).get(name)

    def summary(self) -> dict:
        places = []
        for f, by in self.places.items():
            wi = self.winter_index(f)
            ranked = sorted(wi.items(), key=lambda kv: -kv[1])
            places.append({"key": f, "label": self.label[f], "seasons": len(by),
                           "coldest": ranked[:3], "mildest": ranked[-3:][::-1]})
        return {"unit": self.unit, "detrend": self.detrend, "seasons": self.seasons,
                "places": sorted(places, key=lambda p: p["label"].lower()),
                "source": self.data.get("source", ""), "by": self.data.get("by", ""), "ts": self.data.get("ts", "")}

    def table(self, place: str) -> dict:
        """Factor heat map for one place: seasons down, months (Apr..Mar) across."""
        fac = self.factors(place)
        rows = [{"season": s, "factors": [fac[s].get(m) for m in SEASON_ORDER],
                 "values": [self.places[place][s].get(m) for m in SEASON_ORDER]} for s in sorted(fac)]
        return {"place": self.label.get(place, place), "months": SEASON_ORDER, "rows": rows,
                "winter": self.winter_index(place)}

    # ---- matching plants to the most specific place ----------------------- #

    def place_for(self, plant: dict, area: dict) -> Optional[str]:
        name = plant.get("name") or ""
        cands = [fold(name)]
        m = re.match(r"\s*(\d{3,4})\b", name)
        if m:
            cands.append(fold(m.group(1).zfill(4)))
            cands.append(fold(m.group(1)))
        bare = re.sub(r"^\s*\d{3,4}\s*-\s*", "", name)
        cands.append(fold(bare))
        if plant.get("state"):
            cands.append(fold(plant["state"]))
        cands += [fold(area.get("name", "")), fold(area.get("title", "")), fold(area.get("id", "")),
                  fold(area.get("region", ""))]
        for c in cands:
            if c and c in self.places:
                return c
        for k in self.places:
            if k in ALL_KEYS:
                return k
        return None

    def weather_fn(self, area: dict, setting: Optional[dict]) -> Optional[Callable[[dict, int], float]]:
        """A function (plant, calendar month) -> factor for the engine, or None for normal weather."""
        setting = setting or {}
        mode = setting.get("mode", "normal")
        if mode == "normal" or self.empty:
            return None
        months = set(setting.get("months") or WINTER)
        season = setting.get("season", "")
        cache: Dict[str, Optional[str]] = {}

        def fn(plant: dict, mon: int) -> float:
            if mon not in months:
                return 1.0
            pid = plant.get("id", "")
            if pid not in cache:
                cache[pid] = self.place_for(plant, area)
            place = cache[pid]
            if place is None:
                return 1.0
            v = self.scenario_factor(place, mode, mon, season)
            return 1.0 if v is None else v
        return fn

    def coverage(self, area: dict) -> dict:
        """How many of an area's plants have history, and from what level."""
        lv = {"plant": 0, "state": 0, "area": 0, "all": 0, "none": 0}
        for pl in area.get("plants", []):
            p = self.place_for(pl, area)
            if p is None:
                lv["none"] += 1
            elif p in ALL_KEYS:
                lv["all"] += 1
            elif plant_level(p, pl):
                lv["plant"] += 1
            elif pl.get("state") and p == fold(pl["state"]):
                lv["state"] += 1
            else:
                lv["area"] += 1
        return lv


def plant_level(place: str, plant: dict) -> bool:
    name = plant.get("name") or ""
    m = re.match(r"\s*(\d{3,4})\b", name)
    keys = {fold(name), fold(re.sub(r"^\s*\d{3,4}\s*-\s*", "", name))}
    if m:
        keys |= {fold(m.group(1)), fold(m.group(1).zfill(4))}
    return place in keys


def describe(setting: Optional[dict]) -> str:
    setting = setting or {}
    mode = setting.get("mode", "normal")
    if mode == "normal":
        return "Normal weather"
    months = setting.get("months") or WINTER
    when = "winter (Oct\u2013Mar)" if sorted(months) == sorted(WINTER) else ", ".join(MONTHS[m - 1] for m in months)
    if mode == "season":
        return f"Repeat {season_label(setting.get('season'))} in {when}"
    return f"{MODES.get(mode, mode)} in {when}"
