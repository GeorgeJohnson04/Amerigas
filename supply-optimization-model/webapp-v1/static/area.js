/* area.js - one area: summary, supply, demand, lanes, change log.
   (Results, insights and what-if are in results.js.) */
"use strict";

S.ui = Object.assign({ supplyField: "contract", plantField: "budget", laneMode: "grid", laneMonth: 0, laneFilter: "",
  resView: "lanes", resCost: "both", resPer: "total", resMonth: -1, hideEmpty: true }, JSON.parse(localStorage.getItem("planner.ui") || "{}"));
function saveUi() { localStorage.setItem("planner.ui", JSON.stringify(S.ui)); }

const TABS = [
  { id: "summary", label: "Summary" },
  { id: "supply", label: "Supply", cls: "c-supply" },
  { id: "demand", label: "Demand", cls: "c-demand" },
  { id: "lanes", label: "Lanes & freight", cls: "c-lanes" },
  { id: "results", label: "Results", cls: "c-results" },
  { id: "insights", label: "Insights", cls: "c-results" },
  { id: "whatif", label: "What-if", cls: "c-compare" },
  { id: "changes", label: "Changes" },
];

async function renderArea(aid, tab, sub) {
  const main = $("#main");
  S.A = await API.get("/api/area", { view: S.view, area: aid });
  S.A.id = aid;
  const a = S.A.area, L = S.A.label;
  const row = areaRow(aid) || {};
  const st = statusOf(row);
  const can = S.info.can_edit;
  const runBtn = h("button", { class: "btn btn-primary", type: "button" }, row.run && row.run.fresh ? "Run again" : "Run the plan");
  runBtn.addEventListener("click", () => runArea(runBtn));
  const more = [];
  more.push(h("button", { class: "btn btn-quiet", type: "button", onclick: () => API.download("/api/template", { view: S.view, area: aid }) }, "Download data (Excel)"));
  if (S.info.plan && S.A.changed_in_plan && can) more.push(h("button", { class: "btn btn-quiet", type: "button", onclick: resetArea }, "Undo all my changes here"));
  if (S.info.plan && S.A.changed_in_plan) more.push(h("button", { class: "btn btn-quiet", type: "button", onclick: promoteArea }, "Copy into the data set"));
  if (!S.info.plan && can) more.push(h("button", { class: "btn btn-quiet", type: "button", onclick: editAreaInfo }, "Rename / move"));
  main.append(
    h("div", { class: "area-head" },
      h("div", null, h("div", { class: "crumbs" }, h("a", { href: "#/overview" }, "Overview"), " \u203a ", ((S.info.regions || {})[a.region] || {}).title || a.region),
        h("h1", null, a.name, h("span", { class: "area-title" }, L.title && L.title !== a.name ? L.title : "")),
        h("div", { class: "area-meta" }, h("span", { class: "status" }, h("i", { class: "dot " + st.cls }), st.text),
          L.states ? h("span", null, L.states) : null, h("span", null, `${a.supply.length} supply points \u00b7 ${a.plants.length} plants \u00b7 ${a.lanes.length} lanes`),
          S.A.changed_in_plan ? h("span", { class: "tag changed" }, "changed in this plan") : null)),
      h("div", { class: "page-actions" }, more, runBtn)),
    h("nav", { class: "tabs", "aria-label": "Area sections" }, TABS.map(t => h("a", { class: "tab " + (t.cls || "") + (t.id === tab ? " on" : ""), href: `#/area/${aid}/${t.id}`, "aria-current": t.id === tab ? "page" : null }, t.label))),
    h("div", { id: "tabBody", class: "tab-body" }));
  await drawTab(tab, sub);
}
async function drawTab(tab, sub) {
  const body = clear($("#tabBody"));
  const fn = { summary: tabSummary, supply: tabSupply, demand: tabDemand, lanes: tabLanes, results: tabResults, insights: tabInsights, whatif: tabWhatif, changes: tabChanges }[tab] || tabSummary;
  try { await fn(body, sub); } catch (e) { body.appendChild(h("div", { class: "card error" }, e.message)); }
}
async function reloadArea() {
  const y = window.scrollY;
  S.A = Object.assign(await API.get("/api/area", { view: S.view, area: S.A.id }), { id: S.A.id });
  await refreshNav();
  await drawTab(S.route.tab, S.route.sub);
  window.scrollTo(0, y);
}
async function runArea(btn) {
  busy(btn, true, "Planning\u2026");
  try {
    const r = await API.post("/api/solve", { view: S.view, area: S.A.id });
    toast(r.status === "optimal" ? `Planned: ${F.money(r.summary.total)} (${F.rate(r.summary.per_gal)}/gal)` : "Planned, but the inputs leave gaps - see the notes.", r.status === "optimal" ? "ok" : "error");
    await refreshNav();
    if (S.route.tab !== "results") location.hash = `#/area/${S.A.id}/results`; else route();
  } catch (e) { toast(e.message, "error"); busy(btn, false); }
}
async function saveEdits(items, label) {
  if (!(await needName())) throw new Error("Enter your name to save");
  const r = await API.post("/api/edit", { view: S.view, area: S.A.id, items, label });
  toast(r.count ? `Saved ${r.count} change(s).` : "Nothing changed.");
  await reloadArea();
}
async function resetArea() {
  if (!(await confirmBox("Undo all your changes in this area?", "This area goes back to how it is in the shared data set. (You can undo this from the change log.)", "Put it back"))) return;
  await API.post("/api/reset_area", { view: S.view, area: S.A.id }); toast("Back to the data set."); await refreshNav(); route();
}
async function promoteArea() {
  if (!(await needName())) return;
  if (!(await confirmBox("Copy this area into the data set?", "Everyone will see this plan\u2019s version of the area in the shared data set. The change is logged and can be undone.", "Copy into data set"))) return;
  await API.post("/api/promote_area", { view: S.view, area: S.A.id }); toast("Copied into the data set."); await refreshNav(); route();
}
async function editAreaInfo() {
  const a = S.A.area;
  const v = await promptBox("Rename / move area", [{ key: "name", label: "Short name", value: a.name }, { key: "title", label: "Full name", value: a.title || "" }, { key: "region", label: "Region", value: a.region || "" }], "Save");
  if (!v) return;
  await API.patch("/api/areas", { view: S.view, area: a.id, ...v });
  toast("Saved."); await refreshNav(); route();
}

const supName = s => s ? s.location + (s.supplier ? ` | ${s.supplier}` : "") : "?";
const supById = id => S.A.area.supply.find(s => s.id === id);
const plantById = id => S.A.area.plants.find(p => p.id === id);
const shortPlant = n => String(n || "").replace(/^\s*\d{3,4}\s*-\s*/, "");
const series = (v, P) => { const out = []; for (let p = 0; p < P; p++) out.push(Array.isArray(v) ? (v[p] ?? (v.length ? v[v.length - 1] : null)) : (v ?? null)); return out; };
const sum = arr => (arr || []).reduce((s, v) => s + (+v || 0), 0);

/* ---------------- summary ---------------- */
async function tabSummary(body) {
  const a = S.A.area, bd = S.A.breakdown, P = a.months.length;
  const contract = a.supply.reduce((s, x) => s + sum(series(x.contract, P)), 0);
  const maxLift = a.supply.reduce((s, x) => { const c = series(x.contract, P), m = series(x.max_pct, P); return s + c.reduce((t, v, p) => t + (+v || 0) * (m[p] ?? 1), 0); }, 0);
  const planned = bd.planned.reduce((s, r) => s + sum(r), 0), budget = bd.budget.reduce((s, r) => s + sum(r), 0);
  const high = a.lanes.filter(l => series(l.rate, P).some(v => v !== null && v >= S.boot.high_freight)).length;
  const unusedCells = a.lanes.reduce((s, l) => s + series(l.rate, P).filter(v => v === null).length, 0);
  body.append(
    h("div", { class: "flow" },
      h("a", { class: "flow-step c-supply", href: `#/area/${a.id}/supply` }, h("div", { class: "fs-title" }, "Supply"), h("div", { class: "fs-big" }, F.galShort(contract)),
        h("div", { class: "fs-sub" }, `contracted at ${a.supply.length} supply points \u00b7 up to ${F.galShort(maxLift)} at max lift`)),
      h("div", { class: "flow-arrow", "aria-hidden": "true" }, "\u2192"),
      h("a", { class: "flow-step c-lanes", href: `#/area/${a.id}/lanes` }, h("div", { class: "fs-title" }, "Lanes"), h("div", { class: "fs-big" }, a.lanes.length),
        h("div", { class: "fs-sub" }, `routes${high ? ` \u00b7 ${high} with $1.00+/gal freight` : ""}${unusedCells ? ` \u00b7 ${unusedCells} lane-months not used (blank)` : ""}`)),
      h("div", { class: "flow-arrow", "aria-hidden": "true" }, "\u2192"),
      h("a", { class: "flow-step c-demand", href: `#/area/${a.id}/demand` }, h("div", { class: "fs-title" }, "Demand"), h("div", { class: "fs-big" }, F.galShort(planned)),
        h("div", { class: "fs-sub" }, `planned at ${a.plants.length} plants \u00b7 budget ${F.galShort(budget)}`))),
    card("Data checks", notesList(S.A.checks), "", h("span", { class: "muted small" }, "Before running")),
    card("Planned demand vs what contracts can supply", chart({ labels: a.months, yFmt: F.galShort, series: [
      { name: "Contract volume", values: a.months.map((_, p) => a.supply.reduce((s, x) => s + (+series(x.contract, P)[p] || 0), 0)), color: "#c7c7f1" },
      { name: "Planned demand", values: a.months.map((_, p) => bd.planned.reduce((s, r) => s + (+r[p] || 0), 0)), type: "line", color: COLORS.demand },
      { name: "Most contracts allow", values: a.months.map((_, p) => a.supply.reduce((s, x) => s + (+series(x.contract, P)[p] || 0) * (series(x.max_pct, P)[p] ?? 1), 0)), type: "line", color: COLORS.supply, dash: "5 4" }] })));
  const r = await API.get("/api/results", { view: S.view, area: a.id });
  if (!r.none && r.ok) {
    const s = r.summary;
    body.appendChild(card("Last run", [
      r.fresh ? null : h("p", { class: "warn-text" }, "The inputs have changed since this run - run again to update it."),
      h("div", { class: "kpis" }, kpi("Delivered cost", F.money(s.total), r.baseline ? `${F.signedMoney(s.total - r.baseline.total)} vs data set` : "", "c-results"),
        kpi("$/gal", F.rate(s.per_gal), "propane + freight"), kpi("Unmet demand", F.galShort(s.unmet), "", s.unmet > 100 ? "c-bad" : "")),
      notesList(r.explain.slice(0, 4)), h("a", { class: "btn", href: `#/area/${a.id}/results` }, "See the results")], "c-results"));
  }
}

/* ---------------- supply ---------------- */
const SUPPLY_FIELDS = [
  { value: "contract", label: "Contract", unit: "gal", fmt: F.num },
  { value: "diff", label: "Differential", unit: "$/gal", fmt: F.rate, price: true },
  { value: "price", label: "Price", unit: "$/gal", fmt: F.rate, ro: true, price: true },
  { value: "min_pct", label: "Min lift %", unit: "%", pct: true },
  { value: "max_pct", label: "Max lift %", unit: "%", pct: true },
  { value: "cap", label: "Storage capacity", unit: "gal", fmt: F.num },
  { value: "inv_min", label: "Min inventory", unit: "gal", fmt: F.num },
];
const pctParse = t => { const s = String(t).trim().replace(/[%,\s]/g, ""); if (s === "") return { ok: true, v: null }; const v = Number(s); return isFinite(v) ? { ok: true, v: v / 100 } : { ok: false }; };
const priceParse = t => { const s = String(t).trim().replace(/[$,\s]/g, ""); if (s === "") return { ok: true, v: null }; const v = Number(s); return isFinite(v) ? { ok: true, v } : { ok: false }; };

function tabSupply(body) {
  const a = S.A.area, P = a.months.length, can = S.info.can_edit;
  const hubs = Object.fromEntries((a.hubs || []).map(x => [x.code, series(x.price, P)]));
  const f = SUPPLY_FIELDS.find(x => x.value === S.ui.supplyField) || SUPPLY_FIELDS[0];
  const rows = a.supply.map(s => {
    const vals = f.value === "price" ? series(s.diff, P).map((d, p) => (hubs[s.hub] ? (hubs[s.hub][p] || 0) : 0) + (+d || 0)) : series(s[f.value], P);
    return { key: s.id, label: s.location, sub: [s.supplier, s.hub ? `hub ${s.hub}` : "no hub", s.open ? `opening ${F.gal(s.open)}` : ""].filter(Boolean).join(" \u00b7 "), obj: s,
      cells: vals.map((v, p) => ({ v, path: f.ro ? null : ["supply", s.id, f.value, p], text: f.pct ? (v === null ? "" : F.pct(v)) : undefined })),
      total: f.value === "contract" || f.value === "cap" ? F.num(sum(vals)) : f.pct ? "" : f.price ? F.rate(vals.reduce((x, y) => x + (+y || 0), 0) / P) : "" };
  });
  const g = grid({ cols: a.months, rows, editable: can && !f.ro, fmt: f.fmt || F.pct, parse: f.pct ? pctParse : f.price ? priceParse : undefined,
    editText: f.pct ? v => String(Math.round(v * 10000) / 100) : undefined, rowTitle: "Supply point", totalLabel: f.price ? "Average" : "Season",
    unitHint: f.pct ? "Percent: type 90 for 90%." : f.price ? "$ per gallon." : "Gallons.", onSave: items => saveEdits(items),
    rowActions: can ? row => h("span", { class: "row-act" },
      h("button", { class: "icon-btn", type: "button", title: "Edit name, supplier, hub, opening inventory", "aria-label": "Edit " + row.label, onclick: () => editSupply(row.obj) }, "\u270e"),
      h("button", { class: "icon-btn danger", type: "button", title: "Remove", "aria-label": "Remove " + row.label, onclick: () => removeThing("supply", row.obj.id, `Remove ${row.label}? Its lanes go too.`) }, "\u00d7")) : null });
  body.append(card("Supply points", [
    h("div", { class: "toolbar" }, seg(SUPPLY_FIELDS.map(x => ({ value: x.value, label: x.label })), f.value, v => { S.ui.supplyField = v; saveUi(); drawTab("supply"); }),
      h("span", { class: "spacer" }), can ? h("button", { class: "btn", type: "button", onclick: addSupply }, "+ Add supply point") : null,
      h("a", { class: "btn btn-quiet", href: "#/uploads/contract" }, "Paste / upload")),
    h("p", { class: "muted small" }, f.value === "price" ? "Price = hub price + differential. Change the hub price below or the differential above." :
      f.value === "contract" ? "Gallons contracted each month. Click a cell to type, or paste a block straight from Excel." :
      f.pct ? "Share of the contract that must (min) or may (max) be taken. 100% = the contract volume." : ""),
    a.supply.length ? g : h("p", { class: "empty" }, "No supply points yet. Add one, or upload them on Data uploads.")], "c-supply"));
  const hubRows = (a.hubs || []).map(x => ({ key: x.code, label: x.code, sub: x.name || "", obj: x, cells: series(x.price, P).map((v, p) => ({ v, path: ["hubs", x.code, "price", p] })),
    total: F.rate(series(x.price, P).reduce((s, v) => s + (+v || 0), 0) / P) }));
  body.append(card("Hub prices", [
    h("div", { class: "toolbar" }, h("p", { class: "muted small" }, "Monthly price per hub in $/gal. Every supply point priced off a hub uses it."), h("span", { class: "spacer" }),
      can ? h("button", { class: "btn", type: "button", onclick: addHub }, "+ Add hub") : null),
    hubRows.length ? grid({ cols: a.months, rows: hubRows, editable: can, fmt: F.rate, parse: priceParse, rowTitle: "Hub", totalLabel: "Average", onSave: items => saveEdits(items, "Hub prices changed"),
      rowActions: can ? row => h("button", { class: "icon-btn danger", type: "button", "aria-label": "Remove hub " + row.label, onclick: () => removeThing("hub", row.key, `Remove hub ${row.label}?`) }, "\u00d7") : null })
      : h("p", { class: "empty" }, "No hubs yet.")], "c-supply"));
}
async function addSupply() {
  if (!(await needName())) return;
  const hubs = (S.A.area.hubs || []).map(x => ({ value: x.code, label: x.code }));
  const v = await promptBox("Add a supply point", [{ key: "location", label: "Location", placeholder: "e.g. Northfield, ST" }, { key: "supplier", label: "Supplier", placeholder: "e.g. SupplierX" },
    { key: "hub", label: "Hub", options: [{ value: "", label: "(none)" }].concat(hubs), value: hubs[0] ? hubs[0].value : "" }], "Add");
  if (!v || !v.location) return;
  await API.post("/api/add", { view: S.view, area: S.A.id, kind: "supply", obj: v });
  toast("Added. Now give it a contract, and lanes to the plants it can serve."); await reloadArea();
}
async function editSupply(s) {
  const hubs = (S.A.area.hubs || []).map(x => ({ value: x.code, label: x.code }));
  const v = await promptBox("Supply point", [{ key: "location", label: "Location", value: s.location }, { key: "supplier", label: "Supplier", value: s.supplier || "" },
    { key: "hub", label: "Hub", options: [{ value: "", label: "(none)" }].concat(hubs), value: s.hub || "" },
    { key: "open", label: "Opening inventory (gal)", value: s.open || 0, type: "number" }], "Save");
  if (!v) return;
  await saveEdits([["location", v.location], ["supplier", v.supplier], ["hub", v.hub], ["open", Number(v.open) || 0]].map(([k, val]) => ({ path: ["supply", s.id, k], value: val })), `Supply point ${v.location} details`);
}
async function addHub() {
  if (!(await needName())) return;
  const v = await promptBox("Add a hub", [{ key: "code", label: "Code", placeholder: "e.g. CON, MTB-Tet, MTB-Non-T" }, { key: "name", label: "Name", placeholder: "e.g. Conway" },
    { key: "price", label: "Price $/gal (all months, optional)", type: "number", step: "0.0001" }], "Add");
  if (!v || !v.code) return;
  await API.post("/api/add", { view: S.view, area: S.A.id, kind: "hub", obj: v }); toast("Hub added."); await reloadArea();
}
async function removeThing(kind, ident, text) {
  if (!(await needName())) return;
  if (!(await confirmBox("Remove?", text + " You can undo this from the change log.", "Remove", true))) return;
  await API.post("/api/remove", { view: S.view, area: S.A.id, kind, ident }); toast("Removed."); await reloadArea();
}

/* ---------------- demand ---------------- */
async function tabDemand(body) {
  const a = S.A.area, bd = S.A.breakdown, P = a.months.length, can = S.info.can_edit;
  const T = k => bd[k === "planned" || k === "budget" ? k : "planned"].reduce((s, r) => s + sum(r), 0);
  // step totals: budget -> after % -> after weather -> after changes
  const M = a.plants.length;
  let afterPct = 0, afterW = 0, budget = 0, planned = 0;
  for (let j = 0; j < M; j++) for (let p = 0; p < P; p++) {
    const b = +bd.budget[j][p] || 0; budget += b;
    afterPct += b * bd.pct[j][p]; afterW += b * bd.pct[j][p] * bd.weather[j][p]; planned += +bd.planned[j][p] || 0;
  }
  const step = (label, val, delta, cls, link) => h("div", { class: "wf-step " + cls }, h("div", { class: "wf-label" }, label), h("div", { class: "wf-val" }, F.galShort(val)),
    delta !== null ? h("div", { class: "wf-delta " + (delta > 0 ? "up" : delta < 0 ? "down" : "") }, delta === 0 ? "no change" : F.signedGal(delta)) : h("div", { class: "wf-delta" }, "as uploaded"),
    link ? h("div", { class: "wf-link" }, link) : null);
  const dset = S.info.demand;
  body.append(card("Where planned demand comes from", [
    h("div", { class: "waterfall" },
      step("Budget demand", budget, null, "s-budget", h("a", { href: "#plantsCard", onclick: e => { e.preventDefault(); $("#plantsCard").scrollIntoView({ behavior: "smooth" }); } }, "edit below")),
      h("span", { class: "wf-op" }, "\u00d7"), step("Contract % by state", afterPct, afterPct - budget, "s-pct", h("a", { href: "#pctCard", onclick: e => { e.preventDefault(); $("#pctCard").scrollIntoView({ behavior: "smooth" }); } }, "edit below")),
      h("span", { class: "wf-op" }, "\u00d7"), step("Weather", afterW, afterW - afterPct, "s-weather", S.info.demand_text),
      h("span", { class: "wf-op" }, "\u00d7"), step("Your demand changes", planned, planned - afterW, "s-adjust", `${dset.adjust.length} change(s)`),
      h("span", { class: "wf-op" }, "="), step("Planned demand", planned, planned - budget, "s-planned", "what the plan covers")),
    chart({ labels: a.months, yFmt: F.galShort, height: 200, series: [
      { name: "Budget", values: a.months.map((_, p) => bd.budget.reduce((s, r) => s + (+r[p] || 0), 0)), color: "#cfe9e6" },
      { name: "Planned", values: a.months.map((_, p) => bd.planned.reduce((s, r) => s + (+r[p] || 0), 0)), type: "line", color: COLORS.demand }] })], "c-demand"));

  body.appendChild(quickChangeCard());
  const w = await API.get("/api/weather", {});
  if (w.loaded) body.appendChild(weatherSettingsCard(w, () => reloadArea()));
  else body.appendChild(card("Weather", h("p", null, "No weather history uploaded yet, so demand uses normal weather. ", h("a", { href: "#/weather" }, "Upload a weather history"), " to plan for a cold or mild winter."), "c-weather"));

  // contract % by state
  const states = [...new Set(a.plants.map(p => (p.state || "").toUpperCase()).filter(Boolean))].sort();
  const sp = S.A.state_pct || {};
  const pctRows = states.map(st => ({ key: st, label: st, sub: `${a.plants.filter(p => (p.state || "").toUpperCase() === st).length} plant(s)`,
    cells: a.months.map((m, p) => { const mon = MONTHS[monthNum(m, p) - 1]; const v = (sp[st] || {})[mon]; return { v: v ?? null, path: { state: st, month: mon }, text: v === undefined || v === null ? "100%" : F.pct(v), cls: v === undefined || v === null ? "unset" : "" }; }) }));
  body.append(h("div", { id: "pctCard" }, card("Contract % by state", [
    h("p", { class: "muted small" }, "Share of budget demand to cover with contracts. Blank means 100%. This table is shared by every area with plants in the state", S.info.plan ? " (in this plan only)." : "."),
    pctRows.length ? grid({ cols: a.months, rows: pctRows, editable: can, totals: false, fmt: F.pct, parse: pctParse, editText: v => String(Math.round(v * 10000) / 100), rowTitle: "State",
      unitHint: "Percent: type 85 for 85%. Clear a cell for 100%.",
      onSave: async items => { if (!(await needName())) return; const r = await API.post("/api/state_pct", { view: S.view, items: items.map(i => ({ ...i.path, value: i.value })) }); toast(`Saved ${r.count || 0} change(s).`); await reloadArea(); } })
      : h("p", { class: "empty" }, "The plants here have no state, so contract % doesn\u2019t apply. Give plants a state (pencil button below).")], "c-demand")));

  // plants
  const PF = [{ value: "budget", label: "Budget (gal)" }, { value: "planned", label: "Planned (gal)" }, { value: "cap", label: "Storage capacity" }, { value: "inv_min", label: "Min inventory" }];
  const f = S.ui.plantField;
  const places = (S.A.weather || {}).places || {};
  const rows = a.plants.map((pl, j) => {
    const vals = f === "planned" ? bd.planned[j] : series(pl[f], P);
    return { key: pl.id, obj: pl, label: pl.name, sub: [pl.state || "no state", places[pl.id] ? `weather: ${places[pl.id]}` : "", pl.open ? `opening ${F.gal(pl.open)}` : ""].filter(Boolean).join(" \u00b7 "),
      cells: vals.map((v, p) => ({ v, path: f === "planned" ? null : ["plants", pl.id, f, p],
        note: f === "planned" ? `Budget ${F.gal(bd.budget[j][p])} \u00d7 ${F.pct(bd.pct[j][p])} contract \u00d7 ${F.factor(bd.weather[j][p])} weather \u00d7 ${F.factor(bd.adjust[j][p])} changes = ${F.gal(bd.planned[j][p])}` : null,
        cls: f === "planned" && Math.abs((bd.planned[j][p] || 0) - (bd.budget[j][p] || 0)) > 0.5 ? ((bd.planned[j][p] > bd.budget[j][p]) ? "up" : "down") : "" })) };
  });
  body.append(h("div", { id: "plantsCard" }, card("Plants", [
    h("div", { class: "toolbar" }, seg(PF, f, v => { S.ui.plantField = v; saveUi(); drawTab("demand"); }), h("span", { class: "spacer" }),
      can ? h("button", { class: "btn", type: "button", onclick: addPlant }, "+ Add plant") : null, h("a", { class: "btn btn-quiet", href: "#/uploads/budget" }, "Paste / upload")),
    h("p", { class: "muted small" }, f === "planned" ? "Read-only: hover a cell to see how it was worked out. Green = above budget, red = below." : f === "budget" ? "Budget demand as uploaded, in gallons. The contract %, weather and your changes are applied on top." : "Gallons."),
    a.plants.length ? grid({ cols: a.months, rows, editable: can && f !== "planned", rowTitle: "Plant", onSave: items => saveEdits(items), unitHint: "Gallons.",
      rowActions: can ? row => h("span", { class: "row-act" },
        h("button", { class: "icon-btn", type: "button", "aria-label": "Edit " + row.label, title: "Edit name, state, opening inventory", onclick: () => editPlant(row.obj) }, "\u270e"),
        h("button", { class: "icon-btn danger", type: "button", "aria-label": "Remove " + row.label, onclick: () => removeThing("plant", row.obj.id, `Remove ${row.label}? Its lanes go too.`) }, "\u00d7")) : null })
      : h("p", { class: "empty" }, "No plants yet.")], "c-demand")));
}
async function addPlant() {
  if (!(await needName())) return;
  const v = await promptBox("Add a plant", [{ key: "name", label: "Plant", placeholder: "e.g. 1001 - Springfield KS" }, { key: "state", label: "State", placeholder: "e.g. KS" }], "Add");
  if (!v || !v.name) return;
  await API.post("/api/add", { view: S.view, area: S.A.id, kind: "plant", obj: v });
  toast("Added. Give it budget demand and at least one lane."); await reloadArea();
}
async function editPlant(pl) {
  const v = await promptBox("Plant", [{ key: "name", label: "Name", value: pl.name }, { key: "state", label: "State", value: pl.state || "" }, { key: "open", label: "Opening inventory (gal)", value: pl.open || 0, type: "number" }], "Save");
  if (!v) return;
  await saveEdits([["name", v.name], ["state", (v.state || "").toUpperCase()], ["open", Number(v.open) || 0]].map(([k, val]) => ({ path: ["plants", pl.id, k], value: val })), `Plant ${v.name} details`);
}

/* quick demand change: when, where, how much -> a rule */
function quickChangeCard() {
  const a = S.A.area, can = S.info.can_edit;
  const st = { months: [12, 1, 2], scope: "area", values: [], pct: 10 };
  const states = [...new Set(a.plants.map(p => (p.state || "").toUpperCase()).filter(Boolean))].sort();
  const whereBox = h("div", { class: "where-values" });
  const drawWhere = () => {
    clear(whereBox);
    if (st.scope === "states") whereBox.append(h("div", { class: "chips" }, states.map(s => h("button", { type: "button", class: "chip" + (st.values.includes(s) ? " on" : ""), onclick: e => { st.values = st.values.includes(s) ? st.values.filter(x => x !== s) : st.values.concat([s]); e.target.classList.toggle("on"); preview(); } }, s))));
    if (st.scope === "plants") whereBox.append(h("div", { class: "check-list" }, a.plants.map(p => h("label", null, h("input", { type: "checkbox", checked: st.values.includes(p.id), onchange: e => { st.values = e.target.checked ? st.values.concat([p.id]) : st.values.filter(x => x !== p.id); preview(); } }), p.name))));
  };
  const pctIn = h("input", { type: "number", value: st.pct, step: "1", "aria-label": "Change in percent", class: "pct-input" });
  pctIn.addEventListener("input", () => { st.pct = Number(pctIn.value) || 0; preview(); });
  const presets = h("div", { class: "seg" }, [-10, -5, 5, 10, 20].map(v => h("button", { type: "button", onclick: () => { st.pct = v; pctIn.value = v; preview(); } }, (v > 0 ? "+" : "") + v + "%")));
  const out = h("div", { class: "preview" });
  const rule = () => {
    const scope = st.scope === "states" ? { type: "states", values: st.values } : st.scope === "plants" ? { type: "plants", values: st.values } : { type: "all" };
    const where = st.scope === "states" ? st.values.join(", ") : st.scope === "plants" ? `${st.values.length} plant(s)` : st.scope === "every" ? "every area" : a.name;
    return { months: st.months, pct: st.pct, scope, area: st.scope === "every" ? "*" : a.id, label: `${st.pct > 0 ? "+" : ""}${st.pct}% ${monthsText(st.months)}, ${where}` };
  };
  let timer;
  function preview() {
    clearTimeout(timer);
    timer = setTimeout(async () => {
      if (!st.months.length || !st.pct || ((st.scope === "states" || st.scope === "plants") && !st.values.length)) { clear(out).append(h("p", { class: "muted" }, "Pick months, where, and how much to see the effect.")); return; }
      const settings = { weather: S.info.demand.weather, adjust: S.info.demand.adjust.concat([rule()]) };
      const pv = await API.post("/api/demand/preview", { view: S.view, settings, area: st.scope === "every" ? null : a.id });
      const d = sum(pv.new) - sum(pv.now);
      clear(out).append(h("p", null, h("b", null, rule().label), ` \u2192 ${F.signedGal(d)} planned demand${st.scope === "every" ? " across all areas" : " in " + a.name}.`),
        chart({ labels: pv.months, height: 170, yFmt: F.galShort, series: [{ name: "Planned now", values: pv.now, color: "#cfe9e6" }, { name: "With this change", values: pv.new, type: "line", color: COLORS.demand }] }));
    }, 200);
  }
  const addBtn = h("button", { class: "btn btn-primary", type: "button", disabled: !can }, "Add this change");
  addBtn.addEventListener("click", async () => {
    if (!(await needName())) return;
    if (!st.months.length || !st.pct) return toast("Pick months and a % first.", "error");
    const r = rule();
    await API.post("/api/demand", { view: S.view, settings: { weather: S.info.demand.weather, adjust: S.info.demand.adjust.concat([r]) }, label: "Demand change: " + r.label });
    toast("Demand change added. Run the plan to see the cost."); await reloadArea();
  });
  const rules = S.info.demand.adjust;
  const list = rules.length ? h("ul", { class: "rules" }, rules.map((r, k) => h("li", null, h("span", { class: "rule-pct " + (r.pct > 0 ? "up" : "down") }, (r.pct > 0 ? "+" : "") + r.pct + "%"),
    h("span", null, r.label || `${monthsText(r.months)}`), r.area !== "*" && r.area !== a.id ? h("span", { class: "muted small" }, ` (other area: ${r.area})`) : null,
    can ? h("button", { class: "icon-btn danger", type: "button", "aria-label": "Remove change", onclick: async () => {
      if (!(await needName())) return;
      await API.post("/api/demand", { view: S.view, settings: { weather: S.info.demand.weather, adjust: rules.filter((_, i) => i !== k) }, label: "Removed demand change: " + (r.label || "") });
      toast("Removed."); await reloadArea();
    } }, "\u00d7") : null))) : h("p", { class: "muted" }, "No demand changes yet.");
  drawWhere(); preview();
  return card("Quick demand change", [
    h("div", { class: "qc-grid" },
      h("div", { class: "qc-col" }, h("h3", null, "1. When"), monthChips(st.months, ms => { st.months = ms; preview(); })),
      h("div", { class: "qc-col" }, h("h3", null, "2. Where"), seg([{ value: "area", label: "This area" }, { value: "states", label: "Some states" }, { value: "plants", label: "Some plants" }, { value: "every", label: "Every area" }], st.scope,
        v => { st.scope = v; st.values = []; drawWhere(); preview(); }), whereBox),
      h("div", { class: "qc-col" }, h("h3", null, "3. How much"), h("div", { class: "row" }, pctIn, h("span", null, "%")), presets)),
    out, h("div", { class: "row end" }, can ? null : h("span", { class: "muted" }, "View only"), addBtn),
    h("h3", null, "Changes in this " + (S.info.plan ? "plan" : "data set")), list], "c-demand");
}

/* ---------------- lanes ---------------- */
function tabLanes(body) {
  const a = S.A.area, P = a.months.length, can = S.info.can_edit;
  const ship = S.A.ship || {};
  const hf = S.boot.high_freight;
  const shipNote = !S.A.ship ? "Run the plan to see gallons shipped under each rate." : S.A.ship_fresh ? "Small number under a rate = gallons shipped in the last run." : "Gallons shipped are from a run before the latest changes.";
  const legend = h("div", { class: "lane-legend" }, h("span", null, h("i", { class: "sw sw-rate" }), "freight $/gal"), h("span", null, h("i", { class: "sw sw-high" }), `$${hf.toFixed(2)}+/gal: used if cheapest, flagged`),
    h("span", null, h("i", { class: "sw sw-blank" }), "blank: lane not used that month"), h("span", null, h("i", { class: "sw sw-none" }), "no lane yet (type a rate to add one)"), h("span", { class: "muted" }, shipNote));
  const tools = h("div", { class: "toolbar" }, seg([{ value: "grid", label: "Grid (one month)" }, { value: "list", label: "List (all months)" }], S.ui.laneMode, v => { S.ui.laneMode = v; saveUi(); drawTab("lanes"); }));
  const content = h("div");
  if (S.ui.laneMode === "grid") {
    const p = Math.min(S.ui.laneMonth || 0, P - 1);
    const msel = h("select", { "aria-label": "Month" }, a.months.map((m, k) => h("option", { value: k, selected: k === p }, m)));
    msel.addEventListener("change", () => { S.ui.laneMonth = +msel.value; saveUi(); drawTab("lanes"); });
    tools.append(h("label", { class: "inline-label" }, "Month ", msel));
    const lanes = new Map(a.lanes.map(l => [l.from + "|" + l.to, l]));
    const rows = a.supply.map(s => ({ key: s.id, label: s.location, sub: [s.supplier, s.hub].filter(Boolean).join(" \u00b7 "),
      cells: a.plants.map(pl => {
        const l = lanes.get(s.id + "|" + pl.id);
        if (!l) return { v: null, path: ["lanes", s.id, pl.id, "rate", p], text: "", cls: "nolane", note: `No lane from ${s.location} to ${pl.name}. Type a rate to add one.` };
        const v = series(l.rate, P)[p];
        const sh = ship[s.id + "|" + pl.id];
        const g = sh ? sh[p] : null;
        return { v, path: ["lanes", s.id, pl.id, "rate", p], cls: (v === null ? "closed" : v >= hf ? "high" : "") + (g > 0.5 ? " used" : ""),
          extra: g > 0.5 ? F.num(g) : "", note: `${s.location} \u2192 ${pl.name}, ${a.months[p]}: ${v === null ? "not used (blank)" : F.rate(v) + "/gal"}${g > 0.5 ? ` \u00b7 ${F.gal(g)} shipped` : ""}` };
      }) }));
    content.append(a.supply.length && a.plants.length ? grid({ cols: a.plants.map(pl => shortPlant(pl.name)), rows, editable: can, totals: false, compact: true, fmt: F.rate, parse: priceParse, rowTitle: "From \\ To",
      unitHint: "$ per gallon. Clear a cell to stop using that lane in this month.", onSave: items => saveEdits(items, `Freight for ${a.months[p]} changed`) })
      : h("p", { class: "empty" }, "Add supply points and plants first."));
  } else {
    const filt = h("input", { type: "search", placeholder: "Filter lanes by name", value: S.ui.laneFilter || "", "aria-label": "Filter lanes" });
    tools.append(filt);
    const draw = () => {
      clear(content);
      const q = (filt.value || "").toLowerCase();
      const list = a.lanes.filter(l => !q || (supName(supById(l.from)) + " " + (plantById(l.to) || {}).name).toLowerCase().includes(q));
      content.append(pager(list, 60, page => grid({ cols: a.months, editable: can, fmt: F.rate, parse: priceParse, rowTitle: "Lane", totals: true, totalLabel: "Shipped",
        onSave: items => saveEdits(items), unitHint: "$ per gallon. Blank = not used that month.",
        rows: page.map(l => { const s = supById(l.from), pl = plantById(l.to), sh = ship[l.from + "|" + l.to];
          const lim = [l.min && sum(series(l.min, P)) ? `min ${F.num(series(l.min, P)[0])}` : "", l.max && series(l.max, P).some(v => v !== null) ? `max ${F.num(series(l.max, P)[0])}` : ""].filter(Boolean).join(" \u00b7 ");
          return { key: l.from + "|" + l.to, obj: l, label: `${s ? s.location : "?"} \u2192 ${pl ? shortPlant(pl.name) : "?"}`, sub: [s && s.supplier, lim].filter(Boolean).join(" \u00b7 "),
            total: sh ? F.num(sum(sh)) : "",
            cells: series(l.rate, P).map((v, p) => ({ v, path: ["lanes", l.from, l.to, "rate", p], cls: v === null ? "closed" : v >= hf ? "high" : "", extra: sh && sh[p] > 0.5 ? F.num(sh[p]) : "" })) }; }),
        rowActions: can ? row => h("span", { class: "row-act" },
          h("button", { class: "icon-btn", type: "button", title: "Min / max gallons", "aria-label": "Limits for " + row.label, onclick: () => laneLimits(row.obj) }, "\u2195"),
          h("button", { class: "icon-btn danger", type: "button", "aria-label": "Remove lane " + row.label, onclick: () => removeThing("lane", [row.obj.from, row.obj.to], `Remove the lane ${row.label}?`) }, "\u00d7")) : null })));
    };
    filt.addEventListener("input", () => { S.ui.laneFilter = filt.value; saveUi(); draw(); });
    draw();
  }
  tools.append(h("span", { class: "spacer" }), can ? h("button", { class: "btn", type: "button", onclick: addLane }, "+ Add lane") : null, h("a", { class: "btn btn-quiet", href: "#/uploads/lanes" }, "Paste / upload"));
  body.append(card("Lanes and freight rates", [tools, legend, content], "c-lanes"));
}
async function addLane() {
  if (!(await needName())) return;
  const a = S.A.area;
  const v = await promptBox("Add a lane", [
    { key: "from", label: "From (supply point)", options: a.supply.map(s => ({ value: s.id, label: supName(s) })) },
    { key: "to", label: "To (plant)", options: a.plants.map(p => ({ value: p.id, label: p.name })) },
    { key: "rate", label: "Freight $/gal (all months)", type: "number", step: "0.0001", help: "Leave blank and fill months in later if rates differ." }], "Add lane");
  if (!v) return;
  await API.post("/api/add", { view: S.view, area: S.A.id, kind: "lane", obj: v }); toast("Lane added."); await reloadArea();
}
async function laneLimits(l) {
  const P = S.A.area.months.length;
  const mn = series(l.min, P)[0], mx = series(l.max, P)[0];
  const v = await promptBox("Lane limits (every month)", [{ key: "min", label: "Minimum gallons per month", value: mn || "", type: "number", help: "Must ship at least this much when the lane is used. Blank = none." },
    { key: "max", label: "Maximum gallons per month", value: mx ?? "", type: "number", help: "Blank = no limit." }], "Save");
  if (!v) return;
  const items = [];
  for (let p = 0; p < P; p++) {
    items.push({ path: ["lanes", l.from, l.to, "min", p], value: v.min === "" ? 0 : Number(v.min) });
    items.push({ path: ["lanes", l.from, l.to, "max", p], value: v.max === "" ? null : Number(v.max) });
  }
  await saveEdits(items, "Lane limits changed");
}

/* ---------------- changes ---------------- */
function tabChanges(body) { body.append(card("Changes in this area", changeList(S.A.id))); }
