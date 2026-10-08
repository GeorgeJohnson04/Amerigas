# Propane Supply Model — Python rebuild

`model.py` is a from-scratch Python rebuild of the propane supply/transport optimisation
that currently lives inside the `*.xlsm` workbooks and gets solved by OpenSolver from
inside Excel.

It reads **the same input cells** out of **the same workbooks**, rebuilds **the same
linear program**, hands it to HiGHS, and writes results out **one workbook per region**
instead of one per area.

The source workbooks are opened **read-only**. Nothing is ever written back to them.

This document is for whoever has to change the model next. It explains what the code is
doing, why it is structured the way it is, and — importantly — which parts you can move
freely and which parts will silently produce wrong numbers if you get them wrong.

---

## 1. Quick start

```bat
py -m pip install openpyxl scipy numpy

py model.py                       :: every region
py model.py --region Rockies      :: one region
py model.py --list-regions        :: what's available
```

Or just double-click **`Run Model.bat`** for a menu. That is what the managers use;
it finds Python, installs the packages if they're missing, and explains the exit code
in plain English at the end.

Results land in `Output\<Region>\<Region>_Model_MM_DD_YY.xlsx`.

---

## 2. What problem it solves

Each **area** (CA-NV, AZ, MI, IL-IN, Alaska, …) is an independent supply plan: buy
propane from N supply points under contract, ship it to M demand points over 12 monthly
periods, using storage at both ends to shift volume between months. Minimise total cost.

A **region** is just the folder the areas sit in. Regions are a reporting grouping, not
a modelling one — see §5.

| Excel concept | Python equivalent |
|---|---|
| One `.xlsm` workbook | One `ModelData` → one LP → one `Solution` |
| `ModelBase` sheet | Everything the LP reads |
| OpenSolver + CBC | `scipy.optimize.linprog` → HiGHS |
| `Output Propane` sheet | Only read, and only under `--compare-saved` |
| One output file per area | One output file per region, areas as rows |

---

## 3. The linear program

Transcribed from `ModelBase`. The objective is Excel's `ModelBase!D4`, minimised.

**Indices** — `p = 1..12` periods, `i = 1..N` supply points, `j = 1..M` demand points.

**Decision variables**

| Variable | Meaning |
|---|---|
| `ship[p,i,j]` | volume moved supply *i* → demand *j* in period *p* |
| `perc[p,i]` | fraction of supply *i*'s period-*p* contract actually lifted |
| `s_wd`, `s_inj`, `s_end` | supply-point storage withdraw / inject / ending inventory |
| `d_wd`, `d_inj`, `d_end` | demand-point storage withdraw / inject / ending inventory |

**Objective**

```
min  Σ_p [ Σ_ij trans_cost[p,i,j] · ship[p,i,j]
         + Σ_i  contract[p,i] · price[p,i] · perc[p,i] ]

price[p,i] = hub_price[p,i] + hub_premium[p,i]      (Excel col D = col R + col AF)
```

`contract` and `price` are **data**, not variables, so `perc · contract · price` is
linear even though it reads like a product of two unknowns. This is the single most
important thing to preserve: if you ever make price a decision variable the problem
stops being an LP.

**Constraints**

```
(A) supply balance,  = 0        contract·perc + s_wd − (Σ_j ship + s_inj) = 0
(B) supply inventory            s_end[p] = s_begin[p] − s_wd + s_inj
                                s_begin[1] = initial inv;  s_begin[p] = s_end[p−1]
                                supply_inv_min ≤ s_end ≤ supply storage capacity
(C) demand balance,  = 0        Σ_i ship + d_wd − (demand + d_inj) = 0
(D) demand inventory            same shape as (B)
                                demand_inv_min ≤ d_end ≤ demand storage capacity
```

**Bounds** — `min_perc ≤ perc ≤ max_perc`, `min_ship ≤ ship ≤ max_ship`, all
withdraw/inject `≥ 0`.

Every one of these mirrors a specific `SolverAdd` call in `MixedFormulator7.8.xlsm`'s
VBA. That workbook is the generator that builds the `.xlsm` files, and it is the
authority on layout — if you ever need to settle an argument about where a block lives
or what a constraint is, decompile its VBA rather than guessing from a sample workbook.

### Units — get this wrong and everything is off by 1000

| Quantity | Unit |
|---|---|
| Volumes | **kgal** (thousands of gallons). 300,000 gal is entered as `300`. |
| Prices, differentials, freight | **$/kgal**. MTB + $0.045 is entered as `45.0`. |

Volume × price therefore lands in plain **dollars**, which is how cost gets reported.
Every output sheet says this in its header line. Don't "helpfully" convert.

---

## 4. Architecture

One pass, five stages. Line numbers are approximate — search by name.

```
discover()            ~3214   walk Model/, group workbooks into {region: [AreaFile]}
  ↓
ModelBaseReader       ~1417   one workbook  → ModelData   (read-only, no solving)
  ↓ apply_overrides() ~1275   optional hand-entered changes, before price is derived
  ↓
LPBuilder             ~2328   ModelData → sparse (c, A_eq, b_eq, lo, hi)
solve_area()          ~2871   the escalating solve ladder → Solution
  ↓
RegionWorkbookWriter  ~3319   [Solution] → one .xlsx per region
```

Orchestration is `run_region()` ~4002 and `main()` ~4449. Supporting types:
`SheetGrid` ~496, `ModelData` ~613, `Solution` ~2228, `OverrideDeck` ~834.

### 4.1 `SheetGrid` — why the whole sheet gets slurped

openpyxl's read-only mode streams rows top to bottom and won't let you jump to an
arbitrary cell. Resolving this model means hopping all over the sheet constantly, so the
whole worksheet goes into one sparse `dict` keyed by `(row, col)` up front. Simpler and
faster than re-scanning. Blank cells, text, Excel error strings and NaN all coerce to
`0.0`, which is not arbitrary — it is exactly what `SUMPRODUCT` does with text inside a
numeric block.

### 4.2 `ModelBaseReader` — label-and-defined-name anchoring

**This is the part most likely to bite you, so it's worth understanding before you
touch it.**

Every workbook comes out of the same generator, which makes it very tempting to compute
block positions arithmetically from N and M. **Don't.** The row offsets are not a clean
function of N and M. Between the demand table and the demand-inventory table, CA-NV has
3 spacer rows and Alaska has 5 — close enough to tempt you, just different enough to
break silently.

So anchors come from two places, and both get verified on load:

1. **The workbook's own defined names** — `Period<p>_p_Contract`, `_Price`, `_HubPrice`,
   `_HubPremium`, `_Perc`, `_Trans_Cost`, `_Ship`. Only **workbook-global** names are
   used. The sheet-local copies point at *other workbooks entirely* via stale external
   links left over from the Scenarios sheet, and following them gives you another area's
   data with no error.
2. **Caption scanning in columns A–C** — the `CAPTION_*` constants near the top of the
   file, matched case-insensitively with whitespace collapsed (the workbooks are full of
   trailing spaces).

A sheet that doesn't match raises `ModelStructureError` immediately with a specific
message. Fail loudly over guessing.

Two layout asymmetries worth knowing:

- **Supply-side** inventory columns (`initial inv`, `end inv`, `min inv`, `max inv`) are
  **single columns** beside each Ship block. Each period's Ship block has its own rows,
  so one column covers all 12 periods.
- **Demand-side** ones are **12-wide blocks** on the `Propane Demand Balance` row, with
  data starting one column right of the caption.

### 4.3 `LPBuilder`

Flattens everything into one variable vector and emits `A_eq` as COO triplets, then
solves `min cᵀx s.t. A_eq x = b_eq, lo ≤ x ≤ hi`. Constraint rows come out in the order
given by `FAM_SUPPLY_BALANCE`, `FAM_SUPPLY_INVENTORY`, `FAM_DEMAND_BALANCE`,
`FAM_DEMAND_INVENTORY`, which is what lets a bare row number from the solver be turned
back into "supply balance, period 4, Northfield ST".

scipy/HiGHS is the solver. There's a PuLP + CBC fallback if scipy is missing — the same
CBC the Excel version used, just driven directly instead of through OpenSolver's COM
layer.

### 4.4 `solve_area()` — the escalating ladder

The workbooks are frequently **over-committed**: their own saved "Feasible Check" cells
say `Max Too Low` or `Min Too High`, and Excel's solver fails on them too, which is why
several saved workbooks contain a stale plan. Returning nothing is the least useful
answer, so:

| Stage | What it allows | Result |
|---|---|---|
| 1 | Nothing. Completely faithful to the sheet. | `Optimal` — an answer Excel could have produced |
| 2 | Four penalised slacks: unmet demand, surplus, excess, unsourced | `Relaxed` — plus the exact volume and location of each gap |
| 3 | Full elastic phase-1 (balance equations themselves bend) | `Relaxed` + named violated constraints, indicative only |

Inventory chains stay **rigid at every stage** — they're accounting identities, and
bending them yields nonsense rather than insight.

Penalties are `SHORTFALL_PENALTY = SURPLUS_PENALTY = 1e7`, `ELASTIC_PENALTY = 1e9`. Real
rates are hundreds to low thousands per kgal, so 1e7 is about four orders of magnitude
clear of anything legitimate, and the 1e9 ordering makes the solver exhaust the
meaningful escape valves before bending an equation — which keeps the diagnosis pointed
at the real culprit.

`--no-auto-relax` stops after stage 1 for strict fidelity. `--allow-shortfall` starts at
stage 2 deliberately, for asking "how short am I?" directly.

**Relaxation is never silent.** Status goes to `Relaxed`, the cell is highlighted with a
hover comment, Diagnostics spells it out, and the process exits **4**.

### 4.5 Input repairs

Real data errors found in the workbooks get repaired so the run can proceed, and every
repair is reported as an **Error** on Diagnostics so somebody fixes the source:

- **min > max** (e.g. a minimum lift % above the maximum) — the *minimum is clamped down
  to the maximum*. Trusting the ceiling is the more conservative purchase.
- **Negative demand** — clamped to 0. For example, a plant such as `3003-Lakeside ST`
  entered with −5 kgal in a month would be read as 0, with an Error logged.

---

## 5. Region vs area — why areas are solved separately

Each area is solved as its **own LP**, and a region workbook is just an aggregation of
those results.

This is deliberate. Merging a region's areas into one combined LP would let volume flow
between areas and would change the recommendations — and it isn't even expressible from
the inputs, because the transport matrices contain no cross-area lanes. Region grouping
is a reporting convenience for managers who own several areas, nothing more.

Practical consequence: **adding, removing or regrouping regions cannot change any
numbers.** If it does, something is wrong.

---

## 6. Output workbook

One `.xlsx` per region, sheets in this order:

| Sheet | Contents |
|---|---|
| **Summary** | One row per area: status, volumes, unmet/surplus/slack, high-freight (≥ $1/gal) volume and cost, commodity/freight/total $, avg $/gal, solver, seconds, source workbook. Region totals below. |
| **Supply Plan** | Per supply point per period: contract, lift %, purchased volume, price, cost |
| **Shipments** | Every non-zero lane-period with its exact volume, with a `CHECK: >= $1/gal` flag on high-freight lanes |
| **Demand Coverage** | Demand vs delivered vs storage movement per demand point per period |
| **Supply Inventory** | Opening / withdraw / inject / ending, against min and max |
| **Demand Inventory** | Same, demand side |
| **Diagnostics** | Notes, Warnings and Errors per area — read this whenever status isn't `Optimal` |
| **Overrides** | *Only present when a deck was in play.* One row per resolved change, before → after, with deck line number |

Non-`Optimal` statuses are highlighted and carry the explanation as a cell comment.

### High freight ($1.00/gal and up) is a warning, not "do not use"

Every freight rate is an ordinary price, whatever its size. Planners used to type `1000`
($1.00/gal) to mean "do not use this lane", but real freight on some lanes is now above
$1.00/gal, so the model no longer treats any rate as special. A lane at $1.00/gal or more
is used whenever it is the cheapest way to supply a plant, exactly as any other lane
(and exactly as Excel did; there was never sentinel logic in the generator).

What it does get is a **warning**: "High freight: x gal moves on lanes with freight of
$1/gal or more, costing $y - check the rates are right", the busiest lanes listed on the
Diagnostics sheet, a `CHECK: >= $1/gal` flag on the Shipments sheet, an amber mark on the
lane in the web app's tables, and a "High-freight lanes" card on Results
(`HIGH_FREIGHT_RATE = 1000` $/kgal, `high_freight_mask()`). Volumes below
`HIGH_FREIGHT_REPORT_MIN = 0.1` kgal aren't reported, because warning about "0.0 kgal"
of simplex rounding dust trains people to ignore the warning that matters. Freight
escalation (next year's base, What-if) moves these rates like every other one.

### Blank freight = lane not used

The **only** way to keep a lane unused is to leave its freight cell **blank** (or put
text such as `n/a` in it). The LP then holds that lane at zero in that month. A minimum
shipment on a blank-freight lane is dropped with a note. A freight of `0` still means
free freight. In the web app, clearing a freight cell closes the lane (shown grey,
"not used"), and lanes added by the model builder start blank. A workbook with no
blank freight cells solves exactly as it did before this rule existed.

---

## 7. Manual entry — the override deck

A plain-text file of hand-entered changes, applied on top of the workbooks without
touching them. Default filename `overrides.txt` beside the script, picked up
automatically; `--overrides PATH` names one explicitly (and a *named* file that doesn't
exist is an error, never a silent skip).

```
[scenario]
name = Wide open first pass          # goes in the filename and on every sheet

[all]
max_pct = 5                          # blanket sweep, every area in the run

[area: MT-WY 26-27]
supply "Northfield, ST" : store_min = 250
supply "Terminal A, ST" | Acme : min_pct = 0.9
demand "2002-Riverton ST" : demand = 0
lane "Terminal A, ST" -> "1001-Springfield ST" : freight = 1000
hub CON : price = 650
```

| Target | Settable fields |
|---|---|
| `supply` | `contract`, `diff`, `hub_price`, `min_pct`, `max_pct`, `store_max`, `store_min`, `store_init` |
| `demand` | `demand`, `store_max`, `store_min`, `store_init` |
| `lane` | `freight`, `min_ship`, `max_ship` |
| `hub` | `price` |
| bare key under `[all]` / `[area:]` | `min_pct`, `max_pct` |

A value is either **one number** (applies to all 12 months) or **exactly 12**.
`min_pct`/`max_pct` are multipliers, not percents — `1.0` = 100%, `5` = 500%.
Precedence is workbook → `[all]` → `[area: X]`, last wins. Name matching ignores case
and spacing.

**Anything that matches nothing is a hard error** naming the line and the nearest
spellings. Never a silent no-op — a typo that quietly does nothing is the worst possible
failure mode for a planning input.

Two design decisions to be aware of:

- An **ambiguous name is an error for `supply`** (it changes commercial terms, and the
  `| <supplier>` qualifier exists precisely to disambiguate) but **allowed for `lane`**,
  where it applies to every matching route and logs each one. If "Terminal A, ST" is
  three supply rows (one per supplier), a planner blocking that lane means all of them.
- Overrides are applied **after the workbook read, before price is derived**, so
  `hub_price`/`diff` changes propagate into `price` properly. All validations and repairs
  re-run afterwards.

Helpers:

```bat
py model.py --template-overrides        :: annotated blank deck listing every real
                                        :: location, correctly spelled, all commented out
py model.py --check-overrides FILE      :: validate + echo every resolved change,
                                        :: before → after, without solving
```

The audit sheet records **no-ops too** — the deck asserted a value, so that assertion is
part of the record.

---

## 8. Exit codes

Shared by `model.py` and `Run Model.bat`.

| Code | Meaning |
|---|---|
| 0 | every area solved cleanly |
| 1 | no workbooks found / bad setup / a named override file is missing |
| 2 | unrecognised region or group name |
| 3 | at least one area could not be solved at all (includes a refused unsupported build) |
| 4 | solved, **but** some areas needed relaxations — inputs are over-committed |
| 5 | the deck wouldn't parse, or one of its lines matched nothing |

4 is a success for file-writing purposes and a failure for trust-the-numbers purposes.
Any automation wrapping this should treat it as "read Diagnostics first".

---

## 9. Fidelity — what's guaranteed, and what isn't

**Verified against the workbooks:**

- Every clean-optimal area reproduces its cost exactly; every relaxed area reproduces
  Excel's own feasibility gap volumes exactly.
- All areas load, including the rail builds.
- `--compare-saved` reads the plan cost already stored on `Output Propane`
  (`p_TopAnnualCommCost`) and compares it to the fresh solve. A current workbook
  matches **to the dollar** at component level (e.g. $1,000,000 saved vs $1,000,000
  re-solved).
- No source workbook is modified (asserted on mtime).

`--compare-saved` is the regression oracle. **Use it after any change to the reader or
the LP.** A difference is reported, never fatal — a stale saved plan is a fact about the
workbook, not a bug in the code. A workbook saved before its last input change can
legitimately be several percent stale.

**Deliberately not modelled — detected and refused, exit 3:**

- **Alt routes** — defined names matching `Period\d+_p_Alt(Ship|Trans_Cost)`
- **Mixed fuel / blending** — an `output mix` sheet, or `m_`-prefixed names

These are optional generator features that add a whole second block of shipment
variables. Refusing is the point: a reader that ignored them would produce a
clean-looking, confidently wrong answer. If you implement either one, the refusal in
`ModelBaseReader._refuse_unsupported_features` is what you delete — and not before.

---

## 10. Extending it

### Add a field to the override deck
1. Add it to the `allowed` dict for that target keyword in `parse_override_deck`.
2. Handle it in `apply_overrides` — mutate the right `ModelData` array. If it feeds
   `price`, set the price-touched flag so `price` gets re-derived.
3. Add a commented example line to `write_override_template`.
4. `--check-overrides` picks it up for free.

### Add a constraint
Add it in `LPBuilder` as a new family, append the family name to the `FAM_*` constants in
emission order, and keep `equality_triplets()` and the row-labelling in step — otherwise
violation messages point at the wrong thing. Then run `--compare-saved` across every
region: a genuine new constraint should change *no* cost that was `Optimal` before,
unless you intend it to.

### Add an output sheet
Add a `_my_sheet(self, wb)` method on `RegionWorkbookWriter` and call it from the same
place the others are called. Follow `_shipments` for the `_header()` helper, which also
sets the frozen pane. Keep the kgal / $-per-kgal units line.

### Read a new block from the workbook
Add a `CAPTION_*` constant, locate it by caption or defined name, and **verify what you
found** rather than trusting an offset. Re-read §4.2 first.

### Traps that have already cost time
- **`Run Model.bat` must be CRLF.** `cmd.exe` needs CRLF for reliable `goto`; LF-only
  gives `The system cannot find the batch label specified`. This was masked for a long
  time because most paths fall through to `:done` instead of jumping to it. If you edit
  the batch file with a Unix-y tool, check the line endings afterwards.
- **Region names contain spaces.** Separate multiple regions with **commas**, not spaces.
- **Only `py` is reliably on PATH here**; plain `python` is the Microsoft Store stub.
- **Never capture `%ERRORLEVEL%` inside an `if (...)` block** in the batch file — it's
  expanded at parse time, so you get the *previous* exit code. That's why the
  information-only menu choices jump to labels.
- **`--out-dir` defaults to `Output` next to the script**, and `NON_REGION_DIRS` keeps
  `test`/`output`/`archive` from being mistaken for regions. Renaming the output folder
  to something not in that set will make it discover itself.

### Verifying a change
```bat
py -m py_compile model.py
py model.py --region Rockies --compare-saved
py model.py                                    :: expect 0, or 4 with Diagnostics
py model.py --check-overrides overrides.example.txt
"Run Model.bat" CHECK
```

---

## 11. Files

| File | Role |
|---|---|
| `model.py` | Everything above. Single module, no package. |
| `Run Model.bat` | Menu launcher for non-technical users. **CRLF.** |
| `groups.json` | `{group name: [region, ...]}` — manager groupings for `--group` |
| `overrides.example.txt` | Short fictional example deck showing the syntax. Generate a real annotated blank deck from your workbooks with `--template-overrides` |
| `MixedFormulator7.8.xlsm` | The VBA generator that builds the source workbooks. Reference only — the authority on layout and on what each constraint is. |
| `webapp\bases.py` | Locked base models: import from the workbooks, roll forward to next year, read the demand workbook (§12) |
| `webapp\paste.py` | Paste from Excel: layout, name and unit detection, preview of the changes (§12) |
| `webapp\data\bases\` | The base models (read-only JSON, one folder per base) |
| `Output\<Region>\` | Generated region workbooks |
| `Run Web App.bat`, `webapp\` | The local web app (§12) |

Dependencies: `openpyxl`, `numpy`, `scipy` (HiGHS). Optional fallback: `pulp` (CBC).

---

## 12. Web app (local, for testing)

A browser front end over the same model. Double-click **`Run Web App.bat`** (or run
`py webapp\app.py`) and open <http://127.0.0.1:8050>.

It lets you:

- **Pick a region and area**, then edit any input in a spreadsheet-style grid: contract
  volumes, hub prices, differentials, min/max lift, storage, demand, and lane freight and
  min/max shipments (one month at a time). Type, then Enter or Tab to move on. Paste a
  block straight from Excel. Clearing a cell puts the workbook value back. `150%` is
  accepted for lift limits.
- **Run the model** for the area or the whole region. You get cost cards with the change
  versus the untouched workbook, then the supply plan, shipments, demand coverage, supply
  inventory and diagnostics on screen.
- **Download the region Excel.** It's the same region workbook the CLI writes, and edits
  appear on its Overrides sheet.

### Base models and working models

**How the yearly models were made in Excel.** From the model documentation and the
workbooks: a year's **Baseline** is last year's Working model with the new year's demand,
but last year's contract volumes, differentials and freight. The **Working** model is that
baseline plus each manager's updates (new supply points, agreed basis curves, new freight,
wholesale volumes). Everything the LP needs is typed into the `ModelBase` sheet; the rest
of the sheet only copies and adds those numbers up (the hub price per supply point is an
`INDEX/MATCH` on the hub table, the delivered price is hub + differential). So a model is
fully described by its raw inputs. `ModelBaseReader.raw_inputs()` reads exactly those, and
`build_from_inputs()` makes the LP from them with no workbook.

**Base models** (`webapp\bases.py`, stored in `webapp\data\bases\<id>\`, one JSON per
area) are those raw inputs for every area, one set per contract year. They are **locked**:
the app has no way to change one and the files are written read-only.

- **First start:** the app reads the Excel models once (about 20 seconds) and makes two
  bases, "2026-27 Base (from Excel Baseline files)" (official) and "2026-27 Working files
  (from Excel)". Every area of the Working base reproduces the workbook's result to the
  dollar (checked for every area).
- **Next year, no Excel:** **Base models → Make next year's base** starts from a base or a
  finished working model, keeps contracts, differentials, hub prices, limits and storage,
  takes the new demand from a demand workbook (matched by plant name, then
  district number; plants it doesn't list, and the Rail areas, keep last year's demand),
  optionally changes every freight rate by x% (blank lanes stay blank), moves the
  months on a year and renames the areas (`MT-WY 27-28`). Demand files are picked up from
  `Test\`, `Model\` or `webapp\data\demand\`, or uploaded on that page.
- **Re-reading Excel:** **Read a base from the Excel models** makes a new base from the
  Baseline or Working workbooks in the Model folder.
- **Official:** one base is marked official (★); new working models start from it.

**Working models** (`webapp\data\scenarios\*.json`) sit on top of a base and store only a
manager's changes: cell edits, model-builder settings, demand % by state and uploaded
workbooks. Each belongs to the person whose name is in the top-right box. **Only the owner
can change it**; everyone else sees it read-only and can **make their own copy**. The
owner can **lock** a finished one so it can't change by accident. Older scenarios made
before this were attached to the Working-files base; those without a recorded owner are
read-only for everyone (copy them to keep working). Deleted ones go to `_deleted\`.

The name box stops accidental overwrites; it is not security, since anyone can type any
name. Real sign-in would replace it.

**Colour code.** Slate = locked base, green = your working model, purple = yours but
locked, amber = somebody else's. The same colours are used on the top bar, the banner on
every area and the home page. Each area tab has its own colour (Inputs blue, Demand %
teal, Builder grey, Results green, Insights purple, What-if orange, Changes amber), and
input tables are grouped Supply (blue), Demand (teal) and Lanes (orange).

**Paste from Excel** (`webapp\paste.py`). For data that arrives as Excel sheets (demand,
freight, contracts, differentials, hub prices, lift %, storage, lane limits). Pick what
you're pasting, copy the cells in Excel including the header row, paste, check, apply.

- **Layouts it reads:** names down and months across (`Apr`, `Apr-26`, `4/1/2026`,
  `April 2026`, `P1`; extra columns like Total/Ratio are ignored, and when a sheet has two
  sets of month columns you pick which); months down and names across; name | month |
  value rows; one value per name for the months you tick; freight as a supply × plant
  grid (one month at a time) or as a list of lanes (from | supplier | to | rate, or with
  months across). Tab- or comma-separated.
- **Names:** exact name, district number (`1001`), the name without its code
  (`Springfield, ST`), supplier + location (needed where two supply rows share a location),
  hub code or name (`CON`, `Mont Belvieu`), then close spellings. Every name shows as
  Matched / Check / Which one? / Not found, and any of them can be changed or skipped.
  Unknown names are skipped, never guessed silently.
- **Units:** gallons or thousand gallons, $/gal or $/thousand gal, `90%` / `90` / `0.9`,
  worked out by comparing with what the model holds, and changeable.
- **Every area at once** (demand and lanes): each row goes to the area its plant is in.
  A "Model" column, like the one in the demand workbook, settles plants that are
  in two areas. A full demand sheet reads in well under a second; rows for plants no
  model has are reported as Not found.
- **Blank cells** leave the model alone by default, or can mean lane not used / no
  limit / zero.
- **Preview first:** cells that would change (and how many go back to the base value),
  biggest changes and season totals. **Apply** saves them as ordinary changes in your
  working model: yellow in the grids, marked "pasted" in the change log, revertable.
- **Copy current values as a starting table** puts the model's numbers in the paste
  layout, to round-trip through Excel.

To lock pasted data in as a base for the same year, use **Base models → Make next year's
base** with the working model as the source and the same year (e.g. 2026-27). The months
and area names then stay as they are.

**Quick demand change** (top of the **Demand changes** tab). Three steps, no maths:
1. **When:** season buttons (All year, Winter Oct–Mar, Peak winter Dec–Feb, Shoulder
   Oct/Nov/Mar, Summer Apr–Sep), or tick single months. Winter months are blue and summer
   gold everywhere months appear.
2. **Where:** every plant in the area, some states (with their gallons), or some plants
   (with a search box).
3. **How much:** −20 / −10 / −5 / +5 / +10 / +20 % buttons or an exact %, or "contract at
   a % of budget" (like the State Adjustments sheet).

Presets fill all three in one click: Cold winter (Dec–Feb +10%), Very cold winter
(Oct–Mar +20%), Mild winter (Dec–Feb −10%), Slow summer (Apr–Sep −10%). Before saving, a
sentence says what will happen ("Demand for every plant in Montana & Wyoming in winter
(Oct–Mar) goes up 20%: 1,000,000 → 1,200,000 gal"), with before → after per month. Plants
with a state are changed through that state's %, and the rest are saved as typed-in
demand. "Put all demand back to the base" undoes it all, and a chart compares demand by
month (base vs now) with winter shaded. The detailed % table below has Summer/Winter
headers and "set summer / winter" links per state. The What-if stress test uses the same
season buttons, plus Cold / Very cold / Polar vortex / Mild winter presets that run in one
click. The Inputs demand table links to the quick tool, and every monthly input table
shades winter month headers.

**Demand % by state.** The base demand is budget demand × the "State Adjustments" % (for
example, a state's summer months at a hypothetical 70%). The **Demand % by state** tab shows each state's 12 monthly
percentages. Change one and every plant in that state is scaled by new % / base %, in
this working model only. Plants are assigned a state from the demand file's district list.

**Analysis.**

- **Insights** (after a run): read from the LP's dual values. Lists contracts at their
  maximum and what one more gallon there would save, lanes at their limit (including
  closed ones), full storage, the delivered cost of one more gallon at each plant, and a
  month-by-month **contract use** strip (at max / in between / at min). These hold for
  small changes (`model.marginal_values`, `model.contract_use`).
- **What-if:** **price sensitivity** (each hub, all hubs and freight up and down; the
  optimizer re-plans each time) and a **cold-winter stress test** (demand +x% in the
  months you pick, default Dec–Feb, with unmet demand by plant). Nothing is saved.
- **Compare:** any two models (bases or anybody's working models) for a region, area by
  area and month by month.
- **Chart:** Results' cost chart switches between **Both / Propane only / Freight only**
  and **Total $ / Per gallon**.

**Edits never touch the source workbooks or the base.** Cell edits are applied through
`model.apply_cell_edits`, at the same point in the pipeline as the override deck, so every
validation and repair still runs over the final numbers. Each change records who made it
and when, and can be reverted cell by cell from the **Changes** tab.

Edits address cells by **position** (month, supply row, demand row), not by name,
because several areas have two supply rows with the same location name.

**Units: gallons and $ per gallon.** Everything people see or type is in gallons and
$ per gallon: the web app and the region Excel workbook (from the CLI and from the
app's download). That covers volumes, prices, freight, cost per gallon and diagnostic
messages. The model itself still calculates in the workbooks' own units, kgal and $/kgal,
and figures are converted only when they're shown or typed in: `model.to_gal`,
`to_per_gal`, `gal_text`, and in the browser `toShow`/`toModel`. Results are therefore
identical; 300 kgal simply shows as 300,000 gal and $650.00/kgal as $0.6500/gal. Freight
of $1.00/gal or more is used like any rate and marked as high freight; only a blank
freight cell means "do not use" (section 6).

**Lane volumes.** After a run, every lane table on Inputs (freight, min and max shipment)
shows under each rate the exact gallons that lane carries in the month picked, with
totals per supply point and per plant ("Show shipped gallons" turns it off; a note says
when inputs changed since the run). Results has a **Lane volumes** view: supply points
down, plants across, for one month or the whole season; hovering a cell gives the
month-by-month gallons to the hundredth, the rate and the freight cost. The "Cost by month" chart on Results has a
**Total $ / Per gallon** switch; per gallon is commodity $ per gallon bought that month
plus freight $ per gallon shipped. Two things still use workbook units:

- the source workbooks themselves
- the text override deck (`overrides.txt`)

**Uploading a workbook worked on offline.** **Upload**, on the home page or any area,
takes an `.xlsm`/`.xlsx`.

- **Area:** worked out from the file name ("MT-WY 26-27 Working (2).xlsm" becomes
  Montana & Wyoming), or picked from a list.
- **Reading:** the file is read with the same reader as the official workbooks, and a
  file it can't read is refused with the reason.
- **New scenario:** it lands in a new scenario, so nobody else's work changes.
- **What you see:** a banner lists every input that differs from the official file
  (cell counts plus the biggest change), and "vs workbook" on Results compares against
  the official file.
- **Going back:** "Use the official workbook instead" switches the area back.
- **Storage:** uploads are kept in `webapp\data\uploads` under generated names, up to
  40 MB each. The official workbook is never changed.
- **Unsaved formula results:** a file whose formula results weren't saved (anything
  saved by a tool that doesn't recalculate) is refused with "open it in Excel, save it,
  and upload again". Without those results every formula-driven price would silently
  read as zero.

**Model builder (adding contracts).** Every area has a **Model builder** tab with the four
MixedFormulator sliders: Time Periods, Supply Points, Demand Points and Price Hubs. It
has no Mixed Fuel or Alternative Transportation options. The sliders start at the
workbook's own size.

- **Supply Points:** sliding up adds new supply points. Each one is a new contract; you
  name the location, supplier and price hub. The hub list is **CON** (Conway),
  **MTB-Tet** and **MTB-Non-T** (Mont Belvieu TET / non-TET), limited to the hubs the
  area's workbook uses (`model.HUB_VARIANTS`, `builder_hub_choices`). Both MTB choices
  start on the workbook's MTB price, and the contract's differential sets the rest. A
  new price hub named `MTB-Non-T` with its own prices takes over that choice.
- **Demand Points:** adds plants.
- **Price Hubs:** adds hubs, each with its own monthly price.
- **Time Periods:** shortens the horizon, or lengthens it up to 36 months. Added months
  start with no volume or demand, and repeat the last month's prices, freight and limits.
- **Build:** press **Build Optimization Model**, then fill the new rows in on the Inputs
  tab like any other cell.

New rows do nothing until they're filled in: volume starts at 0, and every new lane
starts with blank freight (not used). This is `model.extend_structure`, applied before
the cell edits. Rows are only ever added after the workbook's own, so saved edits keep
pointing at the same cells. Shrinking the model drops any edit that no longer fits, and
the page says how many were dropped. Workbook rows can't be removed, and the workbook is
never rebuilt or changed. The settings are saved in the scenario.

**Friendly names.** The app shows "Montana & Wyoming · 2026–27" instead of
`MT-WY 26-27`. Names come from `webapp\display_names.json`, which is editable and
display-only. Restart the app after changing it. Codes that aren't listed there are
spelled out automatically from their state codes (`KS-NE 27-28` → "Kansas & Nebraska ·
2027–28"). The workbook codes remain the real keys everywhere, so scenarios, the
override deck and the CLI don't change. Plant codes are also tidied for display only:
`1001-Springfield ST` shows as "Springfield, ST" with `#1001` underneath. Unused placeholder
rows ("Supplier 24", "Demand 25") are hidden by default, and a toggle shows them.

**Security: this is local-only for now.** It binds to 127.0.0.1, so nothing else on the
network can reach it. There is **no login**: the name box decides whose working models
you can change, but anyone can type any name, so it only prevents accidents. Anyone can
also make a base model or mark one official. Before other people can use it, it needs:

- real sign-in (company SSO)
- HTTPS
- a production WSGI server in place of Flask's dev server
- file storage that tolerates several writers (a database, or a server-local folder
  rather than a synced cloud-drive folder)

Mutating requests already require an `X-Requested-With` header, and the Host header must
be localhost. Those block other web pages and DNS rebinding from driving it.

Files: `webapp\app.py` (Flask API), `webapp\static\` (`index.html`, `app.js`,
`styles.css`; no build step), `webapp\data\` (scenarios, temporary exports). Dependency:
`flask==3.1.3`.
