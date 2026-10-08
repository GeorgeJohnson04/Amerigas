/* app.js - the shell: who you are, which year and plan you're looking at,
   the left-hand navigation, and the pages that aren't about one area
   (welcome, overview, weather, compare, change log, help). */
"use strict";

const S = { boot: null, ds: null, view: null, info: null, route: { page: "overview" } };

/* ---------------- who am I ---------------- */
function drawUser() {
  const u = API.user();
  $("#userChip .user-name").textContent = u || "Set your name";
  $("#userChip .avatar").textContent = u ? u.trim().split(/\s+/).map(w => w[0]).join("").slice(0, 2).toUpperCase() : "?";
  $("#userChip").classList.toggle("unset", !u);
}
async function askName(why) {
  const v = await promptBox("Your name", [{ key: "name", label: "Name", value: API.user(), placeholder: "e.g. Jordan Smith",
    help: why || "Your name signs every change in the log and makes your plans yours. It isn't a password." }], "Save");
  if (v && v.name.trim()) { localStorage.setItem("planner.user", v.name.trim()); drawUser(); await reload(); return true; }
  return false;
}
async function needName() {
  if (API.user()) return true;
  return await askName("Enter your name first - every change is signed with it.");
}

/* ---------------- loading ---------------- */
async function reload(keepRoute) {
  S.boot = await API.get("/api/boot");
  if (!S.boot.datasets.length) { S.ds = null; S.view = null; S.info = null; drawPickers(); drawNav(); renderWelcome(); return; }
  const ids = S.boot.datasets.map(d => d.id);
  S.ds = ids.includes(S.ds) ? S.ds : (ids.includes(localStorage.getItem("planner.ds")) ? localStorage.getItem("planner.ds") : ids[0]);
  const plans = S.boot.plans.filter(p => p.dataset === S.ds).map(p => "plan:" + p.id);
  const saved = localStorage.getItem("planner.view");
  if (!(S.view === "ds:" + S.ds || plans.includes(S.view))) S.view = (saved === "ds:" + S.ds || plans.includes(saved)) ? saved : "ds:" + S.ds;
  localStorage.setItem("planner.ds", S.ds);
  localStorage.setItem("planner.view", S.view);
  await loadView();
  if (!keepRoute) route();
}
async function loadView() {
  S.info = await API.get("/api/view", { view: S.view });
  drawPickers(); drawBanner(); drawNav();
}
async function refreshNav() { try { S.info = await API.get("/api/view", { view: S.view }); drawNav(); drawBanner(); } catch (e) { /* ignore */ } }

function areaRow(aid) { return (S.info && S.info.areas.find(a => a.id === aid)) || null; }
function mode() { if (!S.info) return "none"; if (!S.info.plan) return "ds"; return S.info.can_edit ? "mine" : "theirs"; }

/* ---------------- top bar ---------------- */
function drawPickers() {
  const dp = clear($("#dsPick"));
  for (const d of (S.boot ? S.boot.datasets : [])) dp.appendChild(h("option", { value: d.id, selected: d.id === S.ds }, `${d.name}`));
  dp.appendChild(h("option", { value: "__new" }, "+ New planning year\u2026"));
  const vp = clear($("#viewPick"));
  if (!S.ds) return;
  const me = API.user();
  const plans = S.boot.plans.filter(p => p.dataset === S.ds);
  vp.appendChild(h("option", { value: "ds:" + S.ds, selected: S.view === "ds:" + S.ds }, "The data set (shared)"));
  const mine = plans.filter(p => p.owner === me), theirs = plans.filter(p => p.owner !== me);
  if (mine.length) vp.appendChild(h("optgroup", { label: "My plans" }, mine.map(p => h("option", { value: "plan:" + p.id, selected: S.view === "plan:" + p.id }, p.name))));
  if (theirs.length) vp.appendChild(h("optgroup", { label: "Other people\u2019s plans (view only)" }, theirs.map(p => h("option", { value: "plan:" + p.id, selected: S.view === "plan:" + p.id }, `${p.owner}: ${p.name}`))));
  document.body.dataset.mode = mode();
}
$("#dsPick").addEventListener("change", async e => {
  if (e.target.value === "__new") { e.target.value = S.ds || ""; return newDataset(); }
  S.ds = e.target.value; S.view = "ds:" + S.ds; await reload();
});
$("#viewPick").addEventListener("change", async e => { S.view = e.target.value; localStorage.setItem("planner.view", S.view); await loadView(); route(); });
$("#userChip").addEventListener("click", () => askName());
$("#newPlanBtn").addEventListener("click", () => newPlan());

async function newPlan(copyFrom) {
  if (!S.ds || !(await needName())) return;
  const src = copyFrom ? S.boot.plans.find(p => p.id === copyFrom) : null;
  const v = await promptBox(copyFrom ? "Copy this plan" : "New plan", [
    { key: "name", label: "Plan name", value: src ? `${src.name} (my copy)` : "", placeholder: "e.g. Cold winter, extra Conway supply" },
    { key: "note", label: "What are you trying? (optional)", value: "" }],
    copyFrom ? "Make my copy" : "Create plan");
  if (!v) return;
  const r = await API.post("/api/plans", { dataset: S.ds, name: v.name || "My plan", note: v.note, copy_from: copyFrom || null });
  S.view = r.view;
  localStorage.setItem("planner.view", S.view);
  toast("Plan created. Only you can change it; the shared data set stays as it is.");
  await reload(true); route();
}
async function newDataset() {
  if (!(await needName())) return;
  const opts = [{ value: "", label: "Start empty (upload everything)" }].concat((S.boot.datasets || []).map(d => ({ value: d.id, label: `Copy \u201c${d.name}\u201d` })));
  const v = await promptBox("New planning year", [
    { key: "year", label: "Planning year", value: "", placeholder: "e.g. 2027-28", help: "Months run April to March." },
    { key: "name", label: "Name", value: "", placeholder: "e.g. 2027-28 budget" },
    { key: "copy", label: "Start from", options: opts, value: "" }], "Create");
  if (!v) return;
  const r = await API.post("/api/datasets", { year: v.year, name: v.name || v.year, copy_from: v.copy || null });
  S.ds = r.dataset; S.view = "ds:" + r.dataset;
  toast("Planning year created.");
  await reload();
}

/* ---------------- banner: what am I looking at, can I change it ---------------- */
function drawBanner() {
  const b = clear($("#banner"));
  const m = mode();
  document.body.dataset.mode = m;
  if (!S.info) { b.hidden = true; return; }
  b.hidden = false;
  const ds = S.info.dataset, p = S.info.plan;
  if (m === "ds") {
    b.append(h("span", { class: "banner-tag" }, "Shared data set"),
      h("span", null, h("b", null, ds.name), " \u2014 the inputs everyone plans from. Anyone can change them; every change is signed and can be undone."),
      h("span", { class: "banner-actions" }, h("button", { class: "btn btn-small", type: "button", onclick: () => newPlan() }, "Try ideas in a plan instead")));
  } else if (m === "mine") {
    const n = (p.areas_changed || []).length;
    b.append(h("span", { class: "banner-tag" }, "My plan"),
      h("span", null, h("b", null, p.name), ` \u2014 only you can change this. It starts from the data set and keeps your changes separate`,
        n ? ` (${n} area${n > 1 ? "s" : ""} changed).` : " (nothing changed yet)."),
      h("span", { class: "banner-actions" },
        h("button", { class: "btn btn-small", type: "button", onclick: renamePlan }, "Rename"),
        h("button", { class: "btn btn-small", type: "button", onclick: deletePlan }, "Delete")));
  } else {
    b.append(h("span", { class: "banner-tag" }, "View only"),
      h("span", null, h("b", null, `${p.owner}\u2019s plan \u201c${p.name}\u201d`), " \u2014 you can look and run it, but only ", p.owner, " can change it."),
      h("span", { class: "banner-actions" }, h("button", { class: "btn btn-small", type: "button", onclick: () => newPlan(p.id) }, "Make my own copy")));
  }
}
async function renamePlan() {
  const p = S.info.plan;
  const v = await promptBox("Rename plan", [{ key: "name", label: "Plan name", value: p.name }, { key: "note", label: "Note", value: p.note || "" }], "Save");
  if (!v) return;
  await API.patch("/api/plans/" + p.id, v);
  await reload(true); route();
}
async function deletePlan() {
  const p = S.info.plan;
  if (!(await confirmBox("Delete plan?", `\u201c${p.name}\u201d will be moved to the app's trash folder. The data set isn't affected.`, "Delete plan", true))) return;
  await API.del("/api/plans/" + p.id);
  S.view = "ds:" + S.ds;
  await reload();
}

/* ---------------- left navigation ---------------- */
function statusOf(a) {
  if (!a.run) return { cls: "none", text: "Not run yet" };
  if (a.run.status === "relaxed") return { cls: "bad", text: "Has gaps (unmet demand or extra product)" + (a.run.fresh ? "" : " \u00b7 inputs changed since") };
  if (!a.run.fresh) return { cls: "warn", text: "Inputs changed since the last run" };
  if (a.run.status === "optimal") return { cls: "ok", text: "Planned" };
  return { cls: "bad", text: a.run.status };
}
function drawNav() {
  const nav = clear($("#areaNav"));
  $$(".nav-link").forEach(l => l.classList.toggle("on", l.dataset.page === S.route.page));
  if (!S.info) return;
  const q = ($("#areaFilter").value || "").toLowerCase();
  const byRegion = {};
  for (const a of S.info.areas) {
    if (q && !(a.name + " " + a.title + " " + a.states + " " + a.region).toLowerCase().includes(q)) continue;
    (byRegion[a.region] = byRegion[a.region] || []).push(a);
  }
  for (const [r, list] of Object.entries(byRegion)) {
    const info = (S.info.regions || {})[r] || {};
    const det = h("details", { class: "region", open: true }, h("summary", null, info.title || r, h("span", { class: "count" }, list.length)));
    for (const a of list) {
      const st = statusOf(a);
      det.appendChild(h("a", { class: "area-link" + (S.route.area === a.id ? " on" : ""), href: `#/area/${a.id}/${S.route.area === a.id ? (S.route.tab || "summary") : "summary"}`, title: `${a.title} \u2014 ${st.text}` },
        h("i", { class: "dot " + st.cls, "aria-hidden": "true" }), h("span", { class: "al-name" }, a.name),
        a.changed ? h("span", { class: "al-changed", title: "Changed in this plan" }, "\u270e") : null,
        h("span", { class: "sr-only" }, st.text)));
    }
    nav.appendChild(det);
  }
  if (mode() !== "theirs" && S.info) nav.appendChild(h("button", { class: "btn btn-quiet btn-block", type: "button", onclick: addArea }, "+ Add an area"));
}
$("#areaFilter").addEventListener("input", drawNav);
async function addArea() {
  if (!(await needName())) return;
  const regions = [...new Set(S.info.areas.map(a => a.region))];
  const v = await promptBox("Add an area", [
    { key: "name", label: "Short name", placeholder: "e.g. KS-NE" },
    { key: "title", label: "Full name (optional)", placeholder: "e.g. Kansas & Nebraska" },
    { key: "region", label: "Region", value: regions[0] || "", help: regions.length ? "Existing: " + regions.join(", ") : "" }], "Add area");
  if (!v || !v.name) return;
  const r = await API.post("/api/areas", { view: S.view, ...v });
  await refreshNav();
  location.hash = `#/area/${r.area}/summary`;
}

/* ---------------- routing ---------------- */
window.addEventListener("hashchange", route);
function route() {
  const parts = (location.hash || "#/overview").replace(/^#\/?/, "").split("/");
  const [page, a, b, c] = parts;
  if (!S.info) { if (S.boot && !S.boot.datasets.length) renderWelcome(); return; }
  S.route = page === "area" ? { page, area: a, tab: b || "summary", sub: c } : { page: page || "overview", arg: a };
  drawNav();
  const main = clear($("#main"));
  main.dataset.page = S.route.page;
  const pages = { overview: renderOverview, uploads: () => renderUploads(S.route.arg), weather: renderWeather, compare: renderCompare,
    changes: renderChanges, help: renderHelp, area: () => renderArea(S.route.area, S.route.tab, S.route.sub) };
  Promise.resolve().then(() => (pages[S.route.page] || renderOverview)())
    .catch(e => { main.appendChild(h("div", { class: "card error" }, e.message)); });
  window.scrollTo(0, 0);
}
function pageHead(title, sub, actions, cls) {
  return h("div", { class: "page-head " + (cls || "") }, h("div", null, h("h1", null, title), sub ? h("p", { class: "lede" }, sub) : null),
    actions ? h("div", { class: "page-actions" }, actions) : null);
}
function card(title, kids, cls, actions) {
  return h("section", { class: "card " + (cls || "") }, title ? h("div", { class: "card-head" }, h("h2", null, title), actions ? h("div", { class: "card-actions" }, actions) : null) : null, kids);
}

/* ---------------- welcome (no data yet) ---------------- */
function renderWelcome() {
  document.body.dataset.mode = "none";
  $("#banner").hidden = true;
  const main = clear($("#main"));
  const conv = h("button", { class: "btn btn-primary", type: "button", disabled: !S.boot.legacy }, `Convert ${S.boot.legacy} model workbooks`);
  conv.addEventListener("click", async () => {
    if (!(await needName())) return;
    busy(conv, true, "Reading the workbooks\u2026 (about a minute)");
    try {
      const r = await API.post("/api/setup/convert", { variant: "Working", year: $("#wYear").value || "2026-27" });
      const bad = r.report.filter(x => !x.ok);
      toast(`Converted ${r.report.length - bad.length} areas.` + (bad.length ? ` ${bad.length} couldn't be read.` : ""), bad.length ? "error" : "ok");
      S.ds = r.dataset; S.view = "ds:" + r.dataset; location.hash = "#/overview"; await reload();
    } catch (e) { toast(e.message, "error"); busy(conv, false); }
  });
  const empty = h("button", { class: "btn", type: "button" }, "Start empty");
  empty.addEventListener("click", async () => {
    if (!(await needName())) return;
    const r = await API.post("/api/datasets", { year: $("#wYear").value || "2026-27", name: ($("#wYear").value || "2026-27") + " plan" });
    S.ds = r.dataset; S.view = "ds:" + r.dataset; location.hash = "#/uploads"; await reload();
  });
  main.append(
    h("div", { class: "welcome" },
      h("h1", null, "Welcome to Supply Planner"),
      h("p", { class: "lede" }, "Plan where each plant\u2019s propane comes from, month by month, at the lowest delivered cost. Everything comes from data you upload or paste from Excel: plants and demand, supply contracts, hub prices, freight lanes, contract % by state and a weather history."),
      h("div", { class: "field inline" }, h("label", { for: "wYear" }, "Planning year"), h("input", { id: "wYear", value: "2026-27" })),
      h("div", { class: "welcome-grid" },
        h("div", { class: "welcome-card c-supply" }, h("h2", null, "Get a head start"),
          h("p", null, "Read last year\u2019s Excel model workbooks once and turn them into this year\u2019s data. The workbooks aren\u2019t changed, and the app won\u2019t need them again."),
          S.boot.legacy ? null : h("p", { class: "muted" }, "No model workbooks were found on this PC."), conv),
        h("div", { class: "welcome-card c-upload" }, h("h2", null, "Upload everything"),
          h("p", null, "Start with an empty year and bring data in on the Data uploads page: paste from Excel or upload a workbook. Download the template to see the layout."),
          h("div", { class: "row" }, empty, h("button", { class: "btn btn-quiet", type: "button", onclick: () => API.download("/api/template", { blank: 1 }) }, "Download template")))),
      h("p", { class: "muted" }, "Tip: set your name (top right) first - it signs every change, and plans you make are yours.")));
}

/* ---------------- overview ---------------- */
async function renderOverview() {
  const main = $("#main");
  const info = S.info;
  const areas = info.areas;
  const tot = k => areas.reduce((s, a) => s + (+a[k] || 0), 0);
  const runs = areas.filter(a => a.run);
  const fresh = runs.filter(a => a.run.fresh);
  const cost = runs.reduce((s, a) => s + (a.run.total || 0), 0);
  const plannedRun = runs.reduce((s, a) => s + (a.planned || 0), 0);
  const unmet = runs.reduce((s, a) => s + (a.run.unmet || 0), 0);
  const runBtn = h("button", { class: "btn btn-primary", type: "button" }, runs.length && fresh.length === areas.length ? "Run all areas again" : "Run all areas");
  const prog = h("div", { class: "progress", hidden: true }, h("div", { class: "bar" }), h("span", { class: "progress-text" }));
  runBtn.addEventListener("click", () => runAll(runBtn, prog));
  const title = info.plan ? `${info.plan.owner === API.user() ? "My plan" : info.plan.owner + "\u2019s plan"}: ${info.plan.name}` : info.dataset.name;
  main.append(pageHead(title, info.plan ? `Built on the ${info.dataset.name} data set. Demand: ${info.demand_text}${info.demand.adjust.length ? ` \u00b7 ${info.demand.adjust.length} demand change(s)` : ""}.`
      : `The shared inputs for ${info.dataset.year || info.dataset.name}. Demand: ${info.demand_text}.`,
    [h("button", { class: "btn", type: "button", onclick: () => API.download("/api/results/export", { view: S.view }) }, "Download results (Excel)"), runBtn]));
  main.appendChild(prog);
  main.appendChild(h("div", { class: "kpis" },
    kpi("Areas", areas.length, `${fresh.length} planned and up to date`),
    kpi("Planned demand", F.galShort(tot("planned")), `budget ${F.galShort(tot("budget"))}`, "c-demand"),
    kpi("Delivered cost", runs.length ? F.moneyShort(cost) : "\u2013", runs.length ? `${runs.length} of ${areas.length} areas run` : "run the areas to see it", "c-results"),
    kpi("Average $/gal", runs.length && plannedRun ? F.rate(cost / plannedRun) : "\u2013", "propane + freight", "c-results"),
    kpi("Unmet demand", runs.length ? F.galShort(unmet) : "\u2013", unmet > 100 ? "some areas can\u2019t be fully supplied" : "every plant covered", unmet > 100 ? "c-bad" : "")));

  // what's loaded
  const sum = k => areas.reduce((s, a) => s + (a[k] || 0), 0);
  const nStates = info.state_pct_states || 0;
  const w = info.weather;
  const item = (ok, label, value, link, cls) => h("a", { class: "load-item " + cls + (ok ? " ok" : " missing"), href: link },
    h("span", { class: "li-mark", "aria-hidden": "true" }, ok ? "\u2714" : "!"), h("span", { class: "li-label" }, label), h("span", { class: "li-value" }, value));
  main.appendChild(card("What\u2019s loaded", h("div", { class: "load-grid" },
    item(sum("plants") > 0, "Plants & budget demand", `${F.num(sum("plants"))} plants`, "#/uploads/budget", "c-demand"),
    item(nStates > 0, "Contract % by state", nStates ? `${nStates} states` : "not uploaded (100% used)", "#/uploads/state_pct", "c-demand"),
    item(sum("supply") > 0, "Supply points & contracts", `${F.num(sum("supply"))} supply points`, "#/uploads/contract", "c-supply"),
    item(sum("lanes") > 0, "Lanes & freight rates", `${F.num(sum("lanes"))} lanes`, "#/uploads/lanes", "c-lanes"),
    item(w.loaded, "Weather history", w.loaded ? `${w.seasons} seasons, ${w.places} places` : "not uploaded yet", "#/weather", "c-weather")),
    "", h("a", { class: "btn btn-small", href: "#/uploads" }, "Go to data uploads")));

  // areas by region
  const byRegion = {};
  areas.forEach(a => (byRegion[a.region] = byRegion[a.region] || []).push(a));
  const tbl = h("table", { class: "table" }, h("thead", null, h("tr", null, ["Area", "States", "Plants", "Supply", "Lanes", "Planned demand", "Delivered cost", "$/gal", "Status"].map((t, i) => h("th", { class: i >= 2 && i <= 7 ? "num" : "" }, t)))));
  const tb = h("tbody");
  for (const [r, list] of Object.entries(byRegion)) {
    tb.appendChild(h("tr", { class: "group" }, h("th", { colspan: 9 }, ((info.regions || {})[r] || {}).title || r)));
    for (const a of list) {
      const st = statusOf(a);
      tb.appendChild(h("tr", { class: "click", onclick: () => location.hash = `#/area/${a.id}/summary` },
        h("td", null, h("a", { href: `#/area/${a.id}/summary` }, a.name), h("div", { class: "muted small" }, a.title), a.changed ? h("span", { class: "tag changed" }, "changed in plan") : null),
        h("td", null, a.states), h("td", { class: "num" }, a.plants), h("td", { class: "num" }, a.supply), h("td", { class: "num" }, a.lanes),
        h("td", { class: "num" }, F.gal(a.planned)), h("td", { class: "num" }, a.run ? F.money(a.run.total) : ""), h("td", { class: "num" }, a.run ? F.rate(a.run.per_gal) : ""),
        h("td", null, h("span", { class: "status" }, h("i", { class: "dot " + st.cls }), st.text))));
    }
  }
  tbl.appendChild(tb);
  main.appendChild(card("Areas", h("div", { class: "table-scroll" }, tbl)));
}
async function runAll(btn, prog) {
  const list = S.info.areas;
  busy(btn, true, "Running\u2026");
  prog.hidden = false;
  let k = 0, bad = 0;
  for (const a of list) {
    $(".progress-text", prog).textContent = `Planning ${a.name} (${k + 1} of ${list.length})\u2026`;
    $(".bar", prog).style.width = (100 * k / list.length) + "%";
    try { const r = await API.post("/api/solve", { view: S.view, area: a.id }); if (r.status !== "optimal") bad++; }
    catch (e) { bad++; }
    k++;
  }
  $(".bar", prog).style.width = "100%";
  toast(bad ? `Done. ${bad} area(s) have gaps or problems - they're marked red.` : "All areas planned.", bad ? "error" : "ok");
  await refreshNav();
  route();
}

/* ---------------- weather ---------------- */
async function renderWeather() {
  const main = $("#main");
  const w = await API.get("/api/weather", { view: S.view });
  main.appendChild(pageHead("Weather", "Weather-driven demand from your history: how much colder or milder each month was than normal, and what that does to the plan.",
    [h("a", { class: "btn", href: "#/uploads/weather" }, w.loaded ? "Upload a new history" : "Upload weather history"),
     h("button", { class: "btn btn-quiet", type: "button", onclick: () => API.download("/api/template", { blank: 1, kinds: "weather" }) }, "Template")], "c-weather"));
  if (!w.loaded) {
    main.appendChild(card("No weather history yet", [
      h("p", null, "Upload up to 20+ years of monthly history: one row per place and season, with the twelve months across. A place can be a plant, a state, an area, a region, or \u201cAll\u201d for the whole company. Values can be gallons delivered or heating degree days."),
      h("pre", { class: "example" }, "Place\tSeason\tApr\tMay\tJun\tJul\tAug\tSep\tOct\tNov\tDec\tJan\tFeb\tMar\nKS\t2006-07\t310\t120\t20\t5\t8\t60\t300\t620\t910\t1040\t820\t600\nAll\t2006-07\t\u2026"),
      h("p", null, "Each plant uses the most specific history there is (its own, then its state, area, region, then All). Gallon histories have the growth trend taken out first so a bigger business doesn\u2019t look like colder weather.")]));
    return;
  }
  main.appendChild(weatherSettingsCard(w));
  // places
  const placeSel = h("select", { "aria-label": "Place" }, w.places.map(p => h("option", { value: p.label }, `${p.label} (${p.seasons} seasons)`)));
  const heat = h("div");
  const drawHeat = async () => {
    clear(heat);
    const t = await API.get("/api/weather/table", { place: placeSel.value });
    const tbl = h("table", { class: "table heat" }, h("thead", null, h("tr", null, h("th", null, "Season"), t.months.map(m => h("th", { class: "num" }, m)), h("th", { class: "num" }, "Winter"))));
    const tb = h("tbody");
    for (const r of t.rows.slice().reverse()) {
      tb.appendChild(h("tr", null, h("th", null, r.season), r.factors.map((f, i) => h("td", { class: "num", style: heatStyle(f), title: `${t.months[i]} ${r.season}: ${f ? Math.round((f - 1) * 100) + "% vs normal" : "no data"} (${r.values[i] ?? ""} ${w.unit === "hdd" ? "degree days" : "gal"})` }, f ? f.toFixed(2) : "")),
        h("td", { class: "num", style: heatStyle(t.winter[r.season]) }, t.winter[r.season] ? t.winter[r.season].toFixed(2) : "")));
    }
    tbl.appendChild(tb);
    heat.append(h("p", { class: "muted" }, "1.00 = normal. Blue = colder than normal (more demand), orange = milder."), h("div", { class: "table-scroll" }, tbl));
  };
  placeSel.addEventListener("change", drawHeat);
  main.appendChild(card("History by place", [h("div", { class: "row" }, h("label", null, "Place "), placeSel), heat], "", null));
  drawHeat();
  const list = h("table", { class: "table" }, h("thead", null, h("tr", null, ["Place", "Seasons", "Coldest winters", "Mildest winters"].map(t => h("th", null, t)))),
    h("tbody", null, w.places.map(p => h("tr", null, h("td", null, p.label), h("td", null, p.seasons),
      h("td", null, p.coldest.map(([s, f]) => `${s} (+${Math.round((f - 1) * 100)}%)`).join(", ")),
      h("td", null, p.mildest.map(([s, f]) => `${s} (${Math.round((f - 1) * 100)}%)`).join(", "))))));
  main.appendChild(card("Coldest and mildest winters", h("div", { class: "table-scroll" }, list)));
  const cov = h("table", { class: "table" }, h("thead", null, h("tr", null, ["Area", "Own history", "State", "Area / region", "Company (All)", "None"].map((t, i) => h("th", { class: i ? "num" : "" }, t)))),
    h("tbody", null, w.coverage.map(c => h("tr", null, h("td", null, c.name), ["plant", "state", "area", "all", "none"].map(k => h("td", { class: "num" + (k === "none" && c[k] ? " warn-text" : "") }, c[k] || ""))))));
  main.appendChild(card("Which history each area\u2019s plants use", [h("p", { class: "muted" }, "Number of plants using each level. Plants with no history keep normal weather."), h("div", { class: "table-scroll" }, cov)]));
  main.appendChild(h("p", null, h("button", { class: "btn btn-quiet", type: "button", onclick: async () => {
    if (!(await needName())) return;
    if (!(await confirmBox("Clear the weather history?", "Every plan goes back to normal weather. You can undo this from the change log.", "Clear history", true))) return;
    await API.post("/api/weather/clear", { view: S.view }); toast("Weather history cleared."); route();
  } }, "Clear the weather history")));
}
function heatStyle(f) {
  if (!f) return {};
  const d = Math.max(-0.5, Math.min(0.5, f - 1));
  const a = Math.min(Math.abs(d) * 2.2, 0.85);
  return { background: d >= 0 ? `rgba(2,132,199,${a})` : `rgba(234,88,12,${a})`, color: a > 0.5 ? "#fff" : "" };
}

/* the plan-wide weather choice - used on the Weather page and the Demand tab */
function weatherSettingsCard(w, onSaved) {
  const cur = JSON.parse(JSON.stringify(S.info.demand));
  const set = cur.weather;
  const can = S.info.can_edit;
  const box = h("div", { class: "weather-pick" });
  const modes = Object.entries(w.modes || S.boot.weather_modes);
  const help = { normal: "Demand as budgeted.", cold: "Every chosen month as cold as a 1-in-10 cold month.", warm: "Every chosen month as mild as a 1-in-10 mild month.",
    coldest: "Replay the coldest winter in the history.", warmest: "Replay the mildest winter in the history.", season: "Replay one past season you pick." };
  const seasonSel = h("select", { "aria-label": "Season to repeat", disabled: !can }, (w.seasons || []).slice().reverse().map(s => h("option", { value: s, selected: s === set.season }, s)));
  const preview = h("div", { class: "preview" });
  const months = monthChips(set.months, ms => { set.months = ms; doPreview(); });
  const cards = h("div", { class: "mode-cards" }, modes.map(([k, label]) => h("button", { type: "button", class: "mode-card" + (k === set.mode ? " on" : ""), disabled: !can && k !== set.mode, "data-k": k,
    onclick: e => { set.mode = k; $$(".mode-card", cards).forEach(c => c.classList.toggle("on", c.dataset.k === k)); seasonRow.hidden = k !== "season"; doPreview(); } },
    h("b", null, label), h("small", null, help[k] || ""))));
  const seasonRow = h("div", { class: "row", hidden: set.mode !== "season" }, h("label", null, "Season "), seasonSel);
  seasonSel.addEventListener("change", () => { set.season = seasonSel.value; doPreview(); });
  if (set.mode === "season" && !set.season && w.seasons && w.seasons.length) set.season = w.seasons[w.seasons.length - 1];
  let timer = null;
  async function doPreview() {
    clearTimeout(timer);
    timer = setTimeout(async () => {
      try {
        const pv = await API.post("/api/demand/preview", { view: S.view, settings: cur });
        clear(preview).append(h("p", null, h("b", null, pv.text), ": planned demand ", h("b", null, F.galShort(pv.new.reduce((a, b) => a + b, 0))),
          " vs ", F.galShort(pv.now.reduce((a, b) => a + b, 0)), " now (", F.signedGal(pv.new.reduce((a, b) => a + b, 0) - pv.now.reduce((a, b) => a + b, 0)), ")."),
          chart({ labels: pv.months, height: 180, yFmt: F.galShort, series: [{ name: "Now", values: pv.now, color: "#cbd5e1" }, { name: "With this weather", values: pv.new, type: "line", color: COLORS.weather }] }));
      } catch (e) { clear(preview).append(h("p", { class: "error-text" }, e.message)); }
    }, 250);
  }
  const save = h("button", { class: "btn btn-primary", type: "button", disabled: !can }, "Use this weather");
  save.addEventListener("click", async () => {
    if (!(await needName())) return;
    busy(save, true, "Saving\u2026");
    try { await API.post("/api/demand", { view: S.view, settings: cur, label: "Weather: " + describeWeather(set) }); toast("Weather saved. Re-run areas to see the cost."); await refreshNav(); if (onSaved) onSaved(); else route(); }
    catch (e) { toast(e.message, "error"); } finally { busy(save, false); }
  });
  box.append(cards, seasonRow, h("div", { class: "field" }, h("label", null, "Which months it applies to"), months), preview,
    h("div", { class: "row end" }, can ? null : h("span", { class: "muted" }, "View only"), save));
  doPreview();
  return card(`Weather for this ${S.info.plan ? "plan" : "data set"}: ${S.info.demand_text}`, box, "c-weather",
    h("span", { class: "muted small" }, "Applies to every area"));
}
function describeWeather(w) {
  const label = (S.boot.weather_modes || {})[w.mode] || w.mode;
  if (w.mode === "normal") return "Normal weather";
  if (w.mode === "season") return `Repeat ${w.season} in ${monthsText(w.months)}`;
  return `${label} in ${monthsText(w.months)}`;
}

/* ---------------- compare ---------------- */
async function renderCompare() {
  const main = $("#main");
  const views = [{ value: "ds:" + S.ds, label: "The data set (shared)" }].concat(S.boot.plans.filter(p => p.dataset === S.ds).map(p => ({ value: "plan:" + p.id, label: `${p.owner}: ${p.name}` })));
  const mk = (id, val) => h("select", { id }, views.map(v => h("option", { value: v.value, selected: v.value === val }, v.label)));
  const a = mk("cmpA", "ds:" + S.ds), b = mk("cmpB", S.view !== "ds:" + S.ds ? S.view : (views[1] || views[0]).value);
  const areaSel = h("select", { id: "cmpArea" }, h("option", { value: "" }, "All areas"), S.info.areas.map(x => h("option", { value: x.id, selected: S.route.arg === x.id }, x.name)));
  const go = h("button", { class: "btn btn-primary", type: "button" }, "Compare");
  const out = h("div");
  main.append(pageHead("Compare", "Put the data set and plans side by side: what each costs, and where the product comes from.", null, "c-compare"),
    card(null, h("div", { class: "row wrap" }, h("div", { class: "field" }, h("label", { for: "cmpA" }, "A"), a), h("div", { class: "field" }, h("label", { for: "cmpB" }, "B"), b),
      h("div", { class: "field" }, h("label", { for: "cmpArea" }, "Area"), areaSel), go)), out);
  go.addEventListener("click", async () => {
    clear(out);
    busy(go, true, "Comparing\u2026");
    try {
      const ids = areaSel.value ? [areaSel.value] : S.info.areas.map(x => x.id);
      const rows = [];
      const prog = h("p", { class: "muted" });
      out.appendChild(prog);
      for (const [k, id] of ids.entries()) {
        prog.textContent = `Working on ${id} (${k + 1} of ${ids.length})\u2026`;
        rows.push(await API.post("/api/compare", { a: a.value, b: b.value, area: id }));
      }
      prog.remove();
      const name = id => (S.info.areas.find(x => x.id === id) || {}).name || id;
      const T = (x, k) => x && x[k] !== undefined && x[k] !== null ? x[k] : null;
      const ta = rows.reduce((s, r) => s + (T(r.a, "total") || 0), 0), tb = rows.reduce((s, r) => s + (T(r.b, "total") || 0), 0);
      out.appendChild(h("div", { class: "kpis" }, kpi("A", F.moneyShort(ta), $("option:checked", a).textContent), kpi("B", F.moneyShort(tb), $("option:checked", b).textContent),
        kpi("B vs A", F.signedMoney(tb - ta), tb - ta > 0 ? "B costs more" : tb - ta < 0 ? "B saves money" : "no difference", tb - ta > 0 ? "c-bad" : "c-results")));
      const tbl = h("table", { class: "table" }, h("thead", null, h("tr", null, ["Area", "A cost", "B cost", "Change", "A $/gal", "B $/gal", "A unmet", "B unmet"].map((t, i) => h("th", { class: i ? "num" : "" }, t)))),
        h("tbody", null, rows.map(r => { const d = (T(r.b, "total") || 0) - (T(r.a, "total") || 0); return h("tr", null, h("td", null, name(r.area)), h("td", { class: "num" }, F.money(T(r.a, "total"))), h("td", { class: "num" }, F.money(T(r.b, "total"))),
          h("td", { class: "num " + (d > 0.5 ? "up" : d < -0.5 ? "down" : "") }, F.signedMoney(d)), h("td", { class: "num" }, F.rate(T(r.a, "per_gal"))), h("td", { class: "num" }, F.rate(T(r.b, "per_gal"))),
          h("td", { class: "num" }, F.gal(T(r.a, "unmet"))), h("td", { class: "num" }, F.gal(T(r.b, "unmet")))); })));
      out.appendChild(card("By area", h("div", { class: "table-scroll" }, tbl)));
      if (rows.length === 1 && rows[0].a && rows[0].b && rows[0].a.monthly && rows[0].b.monthly) {
        const r = rows[0];
        const add = m => m.commodity.map((v, i) => v + m.freight[i]);
        out.appendChild(card("Cost by month", chart({ labels: r.months, yFmt: F.moneyShort, series: [{ name: "A", values: add(r.a.monthly), color: "#94a3b8" }, { name: "B", values: add(r.b.monthly), color: COLORS.results }] })));
        out.appendChild(card("What\u2019s different in the inputs", h("table", { class: "table" }, h("thead", null, h("tr", null, h("th", null, ""), h("th", { class: "num" }, "A"), h("th", { class: "num" }, "B"), h("th", { class: "num" }, "Change"))),
          h("tbody", null, r.inputs.map(x => h("tr", null, h("td", null, x.what), h("td", { class: "num" }, F.num(x.a)), h("td", { class: "num" }, F.num(x.b)), h("td", { class: "num" }, F.num(x.b - x.a))))))));
        if (r.supply_moves && r.supply_moves.length) out.appendChild(card("Where the product comes from - biggest moves", h("table", { class: "table" },
          h("thead", null, h("tr", null, h("th", null, "Supply point"), h("th", { class: "num" }, "A bought"), h("th", { class: "num" }, "B bought"), h("th", { class: "num" }, "Change"))),
          h("tbody", null, r.supply_moves.map(m => h("tr", null, h("td", null, m.supply), h("td", { class: "num" }, F.gal(m.a)), h("td", { class: "num" }, F.gal(m.b)), h("td", { class: "num" }, F.signedGal(m.change))))))));
      }
      await refreshNav();
    } catch (e) { out.appendChild(h("div", { class: "card error" }, e.message)); }
    finally { busy(go, false); }
  });
}

/* ---------------- change log ---------------- */
async function renderChanges() {
  const main = $("#main");
  main.append(pageHead("Change log", S.info.plan ? "Every change made to this plan, newest first." : "Every change made to the shared data set, newest first. Anyone can undo a change."));
  main.appendChild(changeList(null));
}
function changeList(area) {
  const box = h("div");
  (async () => {
    const r = await API.get("/api/changes", { view: S.view, area });
    const filter = h("input", { type: "search", placeholder: "Filter by name or text", "aria-label": "Filter changes" });
    const list = h("div");
    const draw = () => {
      clear(list);
      const q = filter.value.toLowerCase();
      const items = r.changes.filter(c => !q || (c.label + " " + c.user).toLowerCase().includes(q));
      if (!items.length) { list.appendChild(h("p", { class: "muted" }, "No changes yet.")); return; }
      list.appendChild(pager(items, 50, page => h("ul", { class: "changes" }, page.map(c => h("li", { class: "change k-" + c.kind + (c.undone ? " undone" : "") },
        h("div", { class: "ch-main" }, h("span", { class: "ch-label" }, c.label), c.undone ? h("span", { class: "tag" }, "undone") : null),
        h("div", { class: "ch-meta" }, h("b", null, c.user || "?"), " \u00b7 ", F.when(c.ts), c.count ? ` \u00b7 ${F.num(c.count)} change(s)` : ""),
        c.can_undo && !c.undone && S.info.can_edit && c.kind !== "undo" && c.kind !== "create" ? h("button", { class: "btn btn-small btn-quiet", type: "button", onclick: async ev => {
          if (!(await needName())) return;
          if (!(await confirmBox("Undo this change?", `\u201c${c.label}\u201d - the things it changed go back to how they were just before it.`, "Undo"))) return;
          busy(ev.target, true);
          try { await API.post("/api/undo", { view: S.view, id: c.id }); toast("Undone."); await refreshNav(); route(); } catch (e) { toast(e.message, "error"); busy(ev.target, false); }
        } }, "Undo") : null)))));
    };
    filter.addEventListener("input", draw);
    box.append(h("div", { class: "row" }, filter), list);
    draw();
  })().catch(e => box.appendChild(h("p", { class: "error-text" }, e.message)));
  return box;
}

/* ---------------- help ---------------- */
function renderHelp() {
  const main = $("#main");
  const term = (t, d) => h("div", { class: "term" }, h("dt", null, t), h("dd", null, d));
  main.append(pageHead("How it works", "What the planner does, in plain words."),
    card("The idea", [h("ol", { class: "steps" },
      h("li", null, h("b", null, "Bring in the data. "), "Plants with their budget demand, supply points with their contracts, hub prices, the lanes (routes) between them with freight rates, the contract % for each state, and a weather history. Paste from Excel or upload a workbook on ", h("a", { href: "#/uploads" }, "Data uploads"), "."),
      h("li", null, h("b", null, "Work out demand. "), "Planned demand = budget \u00d7 contract % for the state \u00d7 weather \u00d7 your own changes. The Demand tab shows each step."),
      h("li", null, h("b", null, "Plan. "), "For each area the optimizer picks how much to buy from each contract, how much to ship on each lane and what to store, month by month, so every plant is covered at the lowest propane + freight cost."),
      h("li", null, h("b", null, "Try ideas. "), "Make a plan (top bar) to try a cold winter, a new contract or different freight without touching the shared data. Use What-if for quick tests and Compare to put plans side by side."))]),
    card("Words used in the app", h("dl", { class: "glossary" },
      term("Data set", "One planning year\u2019s shared inputs. Anyone can change it; every change is in the change log with a name and can be undone."),
      term("Plan", "Your own version on top of a data set. Only you can change it; others can look or make a copy. A plan only keeps its own copy of an area once you change something there."),
      term("Supply point", "Where product is bought: a terminal or pipeline location with a supplier and a hub."),
      term("Contract", "Gallons agreed with a supplier for a month. Minimum lift % is the least that must be taken, maximum lift % the most that may be taken."),
      term("Hub price / differential", "Price = the hub\u2019s price (CON, MTB-Tet, MTB-Non-T\u2026) + the supply point\u2019s differential, in $ per gallon."),
      term("Plant", "Where product is needed, with its budget demand by month."),
      term("Lane", "A route from a supply point to a plant (an OD pair) with a freight rate per month. A blank rate means the lane isn\u2019t used that month - that\u2019s the only way a lane is switched off. Rates of $1.00/gal or more are used like any other but flagged."),
      term("Contract % by state", "The share of budget demand to cover with contracts, by state and month."),
      term("Weather factor", "How a month compared with normal in the history: 1.10 = 10% more demand. Gallon histories have the growth trend removed first."),
      term("Unmet demand", "Gallons a plant needs that the inputs can\u2019t supply (contracts too small, no lane). The plan still shows everything else so you can see where the gap is."),
      term("Shadow price", "What one more gallon of something (a contract, a lane, storage) would save. Shown on the Insights tab."))),
    card("Colours", h("div", { class: "swatches" },
      [["Supply", "c-supply"], ["Demand", "c-demand"], ["Lanes & freight", "c-lanes"], ["Results", "c-results"], ["Weather", "c-weather"]].map(([t, c]) => h("span", { class: "swatch " + c }, t)),
      h("span", { class: "swatch m-ds" }, "Shared data set"), h("span", { class: "swatch m-mine" }, "My plan"), h("span", { class: "swatch m-theirs" }, "Someone else\u2019s plan"))));
}

/* ---------------- start ---------------- */
drawUser();
reload().catch(e => { clear($("#main")).appendChild(h("div", { class: "card error" }, e.message)); });
