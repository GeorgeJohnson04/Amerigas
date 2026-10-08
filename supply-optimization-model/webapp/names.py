"""
names.py - friendly display names for regions and areas.

The workbooks are named with short codes ("MT-WY 26-27", "NM-W TX 26-27").
Those codes stay the internal keys everywhere (scenarios, the override deck,
the CLI) - this module only decides what a PERSON sees:

    "MT-WY 26-27"   ->  title "Montana & Wyoming", season "2026-27", states "MT · WY"
    "S TX 26-27"    ->  title "South Texas", ...

Names come from display_names.json first (hand-written, editable without
touching code), then from an automatic spelling-out of state codes, so a new
workbook still gets a readable name the day it appears.
"""

from __future__ import annotations

import json
import os
import re
from typing import Dict

HERE = os.path.dirname(os.path.abspath(__file__))
NAMES_FILE = os.path.join(HERE, "display_names.json")

STATES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas", "CA": "California",
    "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware", "FL": "Florida", "GA": "Georgia",
    "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
    "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada", "NH": "New Hampshire",
    "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York", "NC": "North Carolina",
    "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania",
    "RI": "Rhode Island", "SC": "South Carolina", "SD": "South Dakota", "TN": "Tennessee",
    "TX": "Texas", "UT": "Utah", "VT": "Vermont", "VA": "Virginia", "WA": "Washington",
    "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming",
}
DIRECTIONS = {"N": "North", "S": "South", "E": "East", "W": "West"}

_SEASON_RE = re.compile(r"\s*(\d{2})-(\d{2})\s*$")


def _load() -> Dict[str, Dict[str, dict]]:
    try:
        with open(NAMES_FILE, encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return {"regions": {}, "areas": {}}
    return {"regions": data.get("regions", {}) or {}, "areas": data.get("areas", {}) or {}}


_NAMES = _load()


def reload() -> None:
    global _NAMES
    _NAMES = _load()


def _join(parts):
    if len(parts) <= 1:
        return "".join(parts)
    return ", ".join(parts[:-1]) + " & " + parts[-1]


def split_season(area: str):
    """'MT-WY 26-27' -> ('MT-WY', '2026-27'); no season -> (area, '')."""
    m = _SEASON_RE.search(area)
    if not m:
        return area.strip(), ""
    return area[:m.start()].strip(), f"20{m.group(1)}\u2013{m.group(2)}"


def _auto(code: str):
    """Spell out state codes: 'NM-W TX' -> ('New Mexico & West Texas', 'NM · TX')."""
    names, abbrs = [], []
    for tok in code.split("-"):
        words = tok.split()
        if len(words) == 2 and words[0].upper() in DIRECTIONS and words[1].upper() in STATES:
            names.append(f"{DIRECTIONS[words[0].upper()]} {STATES[words[1].upper()]}")
            abbrs.append(words[1].upper())
        elif len(words) == 1 and words[0].upper() in STATES:
            names.append(STATES[words[0].upper()])
            abbrs.append(words[0].upper())
        else:
            return code, ""                     # not a state code - leave as written
    return _join(names), " \u00b7 ".join(abbrs)


def area_display(area: str) -> dict:
    code, season = split_season(area)
    over = _NAMES["areas"].get(code, {})
    auto_title, auto_states = _auto(code)
    return {
        "code": code,
        "title": over.get("title") or auto_title,
        "states": over.get("states", auto_states),
        "note": over.get("note", ""),
        "season": season,
    }


def region_display(region: str) -> dict:
    over = _NAMES["regions"].get(region, {})
    return {"title": over.get("title") or region, "tagline": over.get("tagline", "")}
