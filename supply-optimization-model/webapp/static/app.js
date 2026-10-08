/* Propane Supply Model - browser front end.
 *
 * Talks to app.py's JSON API. Nothing here writes to a workbook: edits are
 * posted to /api/edits and stored in the selected scenario on the server.
 * Unsaved edits live in state.pending until "Save changes".
 *
 * Area codes like "MT-WY 26-27" are the keys the API uses; what people SEE is
 * the friendly display name the server sends alongside ("Montana & Wyoming").
 */
"use strict";

const $ = (sel, el = document) => el.querySelector(sel);
const $$ = (sel, el = document) => [...el.querySelectorAll(sel)];
const HDR = { "Content-Type": "application/json", "X-Requested-With": "propane-model" };

// Which index positions each input uses: p = month, i = supply point, j = demand point.
const AXES = {
  contract: ["p", "i"], hub_price: ["p", "i"], hub_premium: ["p", "i"], price: ["p", "i"],
  perc_min: ["p", "i"], perc_max: ["p", "i"],
  supply_inv_init: ["i"], supply_inv_cap: ["p", "i"], supply_inv_min: ["p", "i"],
  demand: ["p", "j"], demand_inv_init: ["j"], demand_inv_cap: ["p", "j"], demand_inv_min: ["p", "j"],
  trans_cost: ["p", "i", "j"], ship_min: ["p", "i", "j"], ship_max: ["p", "i", "j"],
};
const TABLES = [
  { group: "Supply", items: ["contract", "hub_price", "hub_premium", "price", "perc_min",
                             "perc_max", "supply_inv_init", "supply_inv_cap", "supply_inv_min"] },
  { group: "Demand", items: ["demand", "demand_inv_init", "demand_inv_cap", "demand_inv_min"] },
  { group: "Lanes", items: ["trans_cost", "ship_min", "ship_max"] },
];
const NAV_LABEL = {
  contract: "Contract volume", hub_price: "Hub price", hub_premium: "Differential", price: "Delivered price",
  perc_min: "Minimum lift", perc_max: "Maximum lift", supply_inv_init: "Opening inventory",
  supply_inv_cap: "Storage capacity", supply_inv_min: "Minimum inventory",
  demand: "Demand", demand_inv_init: "Opening inventory", demand_inv_cap: "Storage capacity",
  demand_inv_min: "Minimum inventory", trans_cost: "Freight rates", ship_min: "Minimum shipments",
  ship_max: "Maximum shipments",
};
const DESC = {
  contract: "Contracted volume available from each supply point, by month (gallons).",
  hub_price: "The market hub price each supply point is priced off ($ per gallon).",
  hub_premium: "Differential over or under the hub, $ per gallon. Delivered price = hub price + differential.",
  price: "Hub price plus differential ($ per gallon). Calculated for you, so it isn't editable here.",
  perc_min: "The least share of the contract that must be lifted. Type 90% or 0.9.",
  perc_max: "The most that can be lifted. 150% (or 1.5) means half as much again as the contract.",
  supply_inv_init: "Product in storage at the supply point when the season starts (gallons).",
  supply_inv_cap: "The most the supply point can hold at month end (gallons).",
  supply_inv_min: "The least that has to stay in storage at month end (gallons).",
  demand: "What each plant needs, by month (gallons).",
  demand_inv_init: "Product in the plant's storage when the season starts (gallons).",
  demand_inv_cap: "The most the plant can hold at month end (gallons).",
  demand_inv_min: "The least that has to stay in the plant's storage at month end (gallons).",
  trans_cost: "Freight per gallon from each supply point (rows) to each plant (columns). Every rate is used as a real price, $1.00/gal and up included (those get a warning). Only a blank cell keeps a lane unused.",
  ship_min: "Gallons that must move on a lane in the month.",
  ship_max: "The most gallons that can move on a lane in the month. Leave blank for no limit.",
};
const TOTALLED = new Set(["contract", "demand"]);
const CHANGE_ROWS = 400;          // a paste can make thousands of changes; the log shows the newest

// ---- units -------------------------------------------------------------
// The model and the server work in the workbooks' units: thousands of gallons
// and $ per thousand gallons. People see GALLONS and $ PER GALLON. Conversion
// happens only here, at display and entry, so nothing is ever stored twice.
const GAL = 1000;
const VOLUME_FIELDS = new Set(["contract", "demand", "ship_min", "ship_max", "supply_inv_init", "supply_inv_cap",
                               "supply_inv_min", "demand_inv_init", "demand_inv_cap", "demand_inv_min"]);
const PRICE_FIELDS = new Set(["hub_price", "hub_premium", "price", "trans_cost"]);
const scaleOf = f => VOLUME_FIELDS.has(f) ? GAL : PRICE_FIELDS.has(f) ? 1 / GAL : 1;
const UNIT = f => VOLUME_FIELDS.has(f) ? "gal" : PRICE_FIELDS.has(f) ? "$/gal" : "x contract";
const toShow = (f, v) => v == null ? v : v * scaleOf(f);        // model units -> display
const toModel = (f, v) => v == null ? v : v / scaleOf(f);       // display -> model units
const ACCENTS = ["#2457d6", "#f97316", "#0e9f6e", "#7c3aed", "#0891b2", "#db2777", "#ca8a04", "#4f46e5", "#059669", "#dc2626"];
const STATE_CODES = new Set("AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY".split(" "));

const state = {
  regions: [], scenarios: [], scenario: "", region: "", area: "",
  bases: [], base: "", official: "", baseInfo: null, page: "home",
  data: null, pending: new Map(), table: "contract", month: 0, tab: "inputs",
  version: 0, results: null, resultArea: "", resultView: "plan", saving: false,
  hideUnused: true, search: "",
  chartMode: localStorage.getItem("model.chartMode") === "gal" ? "gal" : "total",   // cost chart: total $ or per gallon
  chartParts: localStorage.getItem("model.chartParts") || "both",                   // both | propane | freight
  showShip: localStorage.getItem("model.showShip") !== "0",                         // shipped gallons on lane tables
  laneMonth: 0, laneEmpty: false,
};
const TABS = ["inputs", "paste", "demand", "builder", "results", "insights", "whatif", "changes"];
const BASE_PREFIX = "base:";
const isBaseId = id => String(id || "").startsWith(BASE_PREFIX);
const userName = () => $("#user").value.trim();
const enc = encodeURIComponent;

/* ------------------------------------------------------------------ utils */

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
const nf = (d) => new Intl.NumberFormat("en-US", { maximumFractionDigits: d, minimumFractionDigits: 0 });
const NF0 = nf(0), NF1 = nf(1), NF2 = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const money = v => v == null ? "" : (v < 0 ? "-$" : "$") + NF0.format(Math.abs(v));
function moneyShort(v) {
  if (v == null) return "";
  const a = Math.abs(v), s = v < 0 ? "-$" : "$";
  if (a >= 1e6) return s + (a / 1e6).toFixed(a >= 1e7 ? 1 : 2) + "M";
  if (a >= 1e3) return s + (a / 1e3).toFixed(0) + "K";
  return s + NF0.format(a);
}
// Volumes arrive in kgal; show gallons. (Name kept short because it is used everywhere.)
const kgal = v => v == null ? "" : NF0.format(v * GAL);
const NF4 = new Intl.NumberFormat("en-US", { minimumFractionDigits: 4, maximumFractionDigits: 4 });
const perGal = v => v == null ? "" : "$" + NF4.format(v / GAL);          // $/kgal in, $/gal out
// Full value for editing (already in display units). 12 significant digits
// round-trips every workbook value exactly, so focusing a cell never looks
// like a change.
function inputText(v) {
  if (v == null || !isFinite(v)) return "";
  return String(Number(Number(v).toPrecision(12)));
}
// Gallons to the hundredth, for "exactly how much" tooltips (kgal in).
const NF2g = new Intl.NumberFormat("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
const galExact = v => v == null ? "" : NF2g.format(v * GAL);

// What the last run of THIS area moved on each lane in month p:
// {map: "i,j" -> [kgal, rate, freight $], at, stale}, or null with no usable run.
function laneShipments(p) {
  const r = currentResult(), d = state.data;
  if (!r || !r.shipments || !d || r.supply.length !== d.supply.length || r.demand.length !== d.demand.length
      || r.periods.length !== d.periods.length) return null;
  const map = new Map();
  for (const s of r.shipments) if (s[0] === p) map.set(`${s[1]},${s[2]}`, [s[3], s[4], s[5], s[6]]);
  return { map, at: state.results.at ? state.results.at.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }) : "the last run",
           stale: state.results.version != null && state.results.version !== state.version || state.pending.size > 0 };
}

// A saved or workbook value in words, for tooltips and the change log (model units in).
function editText(field, v) {
  if (v == null && field === "trans_cost") return "blank (lane not used)";
  return cellText(toShow(field, v), field);
}
// What a cell shows when you're not in it, in display units: whole gallons
// with commas, or prices to 4 decimals ($0.7500). Clicking in shows the full value.
function cellText(v, field) {
  if (v == null || !isFinite(v)) return "";
  if (VOLUME_FIELDS.has(field)) return NF0.format(v);
  if (PRICE_FIELDS.has(field)) return NF4.format(v);
  return String(Math.round(v * 100) / 100);
}
function keyOf(field, index) { return field + "|" + index.join(","); }
function initials(name) {
  const w = String(name || "").trim().split(/\s+/).filter(Boolean);
  return w.length ? (w[0][0] + (w.length > 1 ? w[w.length - 1][0] : "")).toUpperCase() : "?";
}
const isPlaceholder = n => /^(supplier|demand)\s*\d+$/i.test(String(n).trim());

// "1001-Springfield MT" -> { main: "Springfield, MT", meta: "#1001" };  "1002-NORTHFIELD MT" -> "Northfield, MT"
function prettyLoc(raw) {
  let s = String(raw || "").trim(), code = "";
  const m = s.match(/^(\d{3,4})\s*-\s*(.+)$/);
  if (m) { code = m[1]; s = m[2].trim(); }
  if (/[A-Z]/.test(s) && s === s.toUpperCase()) {
    s = s.toLowerCase().replace(/\b([a-z])/g, c => c.toUpperCase())
         .replace(/\b([A-Za-z]{2})$/, t => STATE_CODES.has(t.toUpperCase()) ? t.toUpperCase() : t);
  }
  const st = s.match(/^(.*?[a-z.)])\s+([A-Z]{2})$/);
  if (st && STATE_CODES.has(st[2]) && !st[1].endsWith(",")) s = `${st[1]}, ${st[2]}`;
  return { main: s, meta: code ? "#" + code : "" };
}
function areaTitle(d) { return d ? d.title : ""; }
function areaSub(d) { return [d.states, d.season ? d.season + " season" : ""].filter(Boolean).join("  ·  "); }

async function api(method, url, body) {
  const opt = { method, headers: method === "GET" ? {} : HDR };
  if (body !== undefined) opt.body = JSON.stringify(body);
  const r = await fetch(url, opt);
  const isJson = (r.headers.get("content-type") || "").includes("json");
  const data = isJson ? await r.json() : null;
  if (!r.ok) throw new Error((data && data.error) || `${r.status} ${r.statusText}`);
  return data;
}

// Kept as a top-level const so outside scripts (tests) can see busy state.
const busyEl = $("#busy");
async function busy(msg, fn) {
  $("#busy-text").textContent = msg; busyEl.hidden = false;
  try { return await fn(); } finally { busyEl.hidden = true; }
}

function toast(msg, kind = "", action = null, sticky = false) {
  const t = document.createElement("div");
  t.className = "toast " + kind;
  t.setAttribute("role", kind === "error" ? "alert" : "status");
  t.innerHTML = `<div class="txt">${esc(msg)}</div>`;
  if (action) {
    const b = document.createElement("button");
    b.type = "button"; b.textContent = action.label;
    b.onclick = () => { t.remove(); action.run(); };
    t.appendChild(b);
  }
  const x = document.createElement("button");
  x.type = "button"; x.textContent = "✕"; x.setAttribute("aria-label", "Dismiss"); x.onclick = () => t.remove();
  t.appendChild(x);
  $("#toasts").appendChild(t);
  if (!sticky) setTimeout(() => t.remove(), kind === "error" ? 9000 : 4500);
  return t;
}
const fail = err => toast(err.message || String(err), "error");

const RESTART_MSG = "The web app was updated while it was running, so this copy is out of date. " +
  "Close the 'Propane Model Web App' window, then double-click Run Web App.bat to start the new version.";

// The page is always fresh from disk, but the server's code is loaded once at
// start-up - after an update, a copy left running is out of date. Say so up
// front instead of letting new buttons fail one by one.
async function checkServerFresh() {
  try {
    const r = await fetch("/api/health", { cache: "no-store" });
    const stale = r.status === 404 || (r.ok && (await r.json()).stale);
    if (stale) toast(RESTART_MSG, "error", null, true);
  } catch (_) { /* the server is down; other calls will say so */ }
}

function confirmDiscard() {
  if (!state.pending.size) return true;
  return confirm(`You have ${state.pending.size} unsaved change(s). Discard them?`);
}

/* --------------------------------------------------------------- startup */

async function init() {
  const user = $("#user");
  user.value = localStorage.getItem("model.user") || "";
  const syncAvatar = () => { $("#avatar").textContent = initials(user.value); };
  syncAvatar();
  user.addEventListener("input", syncAvatar);
  // A different name = a different set of "my" working models.
  user.addEventListener("change", async () => {
    localStorage.setItem("model.user", user.value.trim());
    await loadScenarios(state.scenario);
    if (state.area) await openArea(state.region, state.area); else showPage(state.page);
  });

  $$(".side-link").forEach(b => b.addEventListener("click", () => goPage(b.dataset.page)));
  $("#base-pick").addEventListener("change", onBasePick);
  initBases(); initCompare(); initDemandPct(); initWhatIf(); initPaste();
  $("#name-form").addEventListener("submit", e => {
    e.preventDefault();
    const v = $("#name-input").value.trim();
    if (!v) return;
    user.value = v; syncAvatar(); localStorage.setItem("model.user", v);
    $("#name-dialog").close();
    loadScenarios(state.scenario).then(() => showPage(state.page));
  });

  $$('[role="tab"]').forEach(t => t.addEventListener("click", () => showTab(t.id.replace("tabbtn-", ""))));
  $("#month-pick").addEventListener("change", e => { state.month = +e.target.value; renderGrid(); });
  $("#hide-unused").addEventListener("change", e => { state.hideUnused = e.target.checked; renderGrid(); });
  $("#show-ship").checked = state.showShip;
  $("#show-ship").addEventListener("change", e => {
    state.showShip = e.target.checked; localStorage.setItem("model.showShip", state.showShip ? "1" : "0"); renderGrid(); });
  $("#save").addEventListener("click", save);
  $("#discard").addEventListener("click", () => { state.pending.clear(); renderGrid(); renderTableNav(); });
  $("#run-area").addEventListener("click", () => run(true));
  $("#run-region").addEventListener("click", () => run(false));
  $("#changes-all").addEventListener("change", renderChanges);
  $("#builder-build").addEventListener("click", () => {
    const b = state.builder, wb = state.data.workbook_size;
    const st = {};
    if (b.periods !== wb.periods) st.periods = b.periods;
    if (b.hubs.length) st.hubs = b.hubs;
    const names = builderHubChoices(state.data, b).map(c => c.value);
    if (b.supply.length) st.supply = b.supply.map(s => ({ ...s, hub: hubPick(s.hub, names) }));
    if (b.demand.length) st.demand = b.demand;
    buildModel(st);
  });
  $("#builder-reset").addEventListener("click", () => {
    if (confirm("Put this area back to the base model's size? New supply points, plants and hubs - and any changes typed into them - are removed.")) buildModel({});
  });
  $("#scenario").addEventListener("change", onScenarioPick);
  $("#brand").addEventListener("click", goHome);
  $("#area-search").addEventListener("input", e => { state.search = e.target.value.trim().toLowerCase(); renderSidebar(); });

  // scenario menu
  const menuBtn = $("#scen-menu-btn"), menu = $("#scen-menu");
  const closeMenu = () => { menu.hidden = true; menuBtn.setAttribute("aria-expanded", "false"); };
  menuBtn.addEventListener("click", e => {
    e.stopPropagation();
    menu.hidden = !menu.hidden; menuBtn.setAttribute("aria-expanded", String(!menu.hidden));
    if (!menu.hidden) menu.querySelector("button").focus();
  });
  document.addEventListener("click", e => { if (!menu.contains(e.target)) closeMenu(); });
  document.addEventListener("keydown", e => { if (e.key === "Escape" && !menu.hidden) { closeMenu(); menuBtn.focus(); } });
  const act = (id, fn) => $(id).addEventListener("click", () => { closeMenu(); fn(); });
  act("#scen-new", () => scenarioCreate(false));
  act("#scen-copy", () => scenarioCreate(true));
  act("#scen-rename", scenarioRename);
  act("#scen-lock", scenarioLock);
  act("#scen-delete", scenarioDelete);

  window.addEventListener("beforeunload", e => { if (state.pending.size) { e.preventDefault(); e.returnValue = ""; } });
  initUpload();
  checkServerFresh();

  try {
    if (!(await waitForSetup())) return;
    await loadBases();
    const pref = localStorage.getItem("model.scenario") || "";
    state.base = localStorage.getItem("model.base") || state.official;
    await loadScenarios(pref);
    await loadRegions();
    if (!userName()) { $("#name-dialog").showModal(); $("#name-input").focus(); }
    // Deep link: #region=Rockies&area=MT-WY%2026-27 reopens that area on load/refresh.
    const h = new URLSearchParams(location.hash.slice(1));
    if (h.get("region") && h.get("area") && areaInfo(h.get("region"), h.get("area"))) {
      if (TABS.includes(h.get("tab"))) state.tab = h.get("tab");
      await openArea(h.get("region"), h.get("area"));
    } else showPage(["bases", "compare"].includes(h.get("page")) ? h.get("page") : "home");
  } catch (err) { fail(err); }
  setInterval(pollVersion, 15000);
}

// First start: the server reads the Excel models into base models once.
async function waitForSetup() {
  let first = true;
  for (;;) {
    const b = await api("GET", "/api/bases");
    const job = b.jobs.find(j => j.state === "running");
    const setupRunning = job && /^First-time/.test(job.label);
    if (b.bases.length && !setupRunning) {
      if (!first) { location.reload(); return false; }       // setup just finished: start clean
      $("#setup-view").hidden = true; return true;
    }
    const err = b.jobs.find(j => j.state === "error");
    if (!job && err) { toast("Setting up the base models failed: " + err.error, "error", null, true); return false; }
    if (first) { $$("#main > section").forEach(s => s.hidden = s.id !== "setup-view"); first = false; }
    if (job) {
      $("#setup-text").textContent = `${job.step || "starting"} (${job.done} of ${job.total || "?"})`;
      $("#setup-bar").style.width = (job.total ? Math.round(100 * job.done / job.total) : 3) + "%";
    }
    await new Promise(r => setTimeout(r, 1500));
  }
}

async function loadBases() {
  const b = await api("GET", "/api/bases");
  state.bases = b.bases; state.official = b.official; state.jobs = b.jobs; state.demandFiles = b.demand_files;
  if (!state.bases.some(x => x.id === state.base)) state.base = state.official || (state.bases[0] || {}).id || "";
  $("#base-pick").innerHTML = state.bases.map(x =>
    `<option value="${esc(x.id)}"${x.id === state.base ? " selected" : ""}>${esc(x.name)}${x.official ? "  \u2605" : ""}</option>`).join("");
  return b;
}

async function loadRegions() {
  const r = await api("GET", `/api/regions?base=${enc(state.base)}`);
  state.regions = r.regions; state.baseInfo = r.base;
  renderSidebar(); renderCrumbs();
  if (!state.area) renderHome();
}

async function onBasePick(e) {
  if (!confirmDiscard()) { e.target.value = state.base; return; }
  state.pending.clear();
  state.base = e.target.value; localStorage.setItem("model.base", state.base);
  state.scenario = BASE_PREFIX + state.base;
  state.results = null;
  await loadScenarios(state.scenario);
  await loadRegions();
  if (state.area && areaInfo(state.region, state.area)) await openArea(state.region, state.area);
  else { state.area = ""; showPage("home"); }
}

function regionAccent(region) {
  const k = state.regions.findIndex(r => r.region === region);
  return ACCENTS[(k < 0 ? 0 : k) % ACCENTS.length];
}
function regionInfo(region) { return state.regions.find(r => r.region === region); }
function areaInfo(region, area) { return (regionInfo(region)?.areas || []).find(a => a.area === area); }

/* ------------------------------------------------------------- scenarios */

// The working-model picker lists, for the current base: the locked base itself,
// then MY working models, then everybody else's (view only).
async function loadScenarios(prefer) {
  state.scenarios = await api("GET", `/api/scenarios?user=${enc(userName())}`);
  const want = state.scenarios.find(s => s.id === prefer);
  if (want && want.base !== state.base) {                 // a saved working model moves the base with it
    state.base = want.base; $("#base-pick").value = state.base; await loadRegions();
  }
  const here = state.scenarios.filter(s => s.base === state.base);
  const mine = here.filter(s => s.mine), others = here.filter(s => !s.mine);
  const ids = [BASE_PREFIX + state.base, ...here.map(s => s.id)];
  state.scenario = ids.includes(prefer) ? prefer : ids.includes(state.scenario) ? state.scenario
                 : (mine[0] ? mine[0].id : BASE_PREFIX + state.base);
  const opt = s => `<option value="${esc(s.id)}"${s.id === state.scenario ? " selected" : ""}>${s.locked ? "\uD83D\uDD12 " : ""}${esc(s.name)}${s.edits ? `  \u00b7  ${s.edits} change${s.edits === 1 ? "" : "s"}` : ""}${s.mine ? "" : `  \u00b7  ${esc(s.owner || "shared")}`}</option>`;
  const base = state.bases.find(b => b.id === state.base);
  $("#scenario").innerHTML =
    `<optgroup label="Locked base (view only)"><option value="${esc(BASE_PREFIX + state.base)}"${state.scenario === BASE_PREFIX + state.base ? " selected" : ""}>\uD83D\uDD12 ${esc(base ? base.name : "Base model")}</option></optgroup>` +
    (mine.length ? `<optgroup label="My working models">${mine.map(opt).join("")}</optgroup>` : "") +
    (others.length ? `<optgroup label="Other people's (view only)">${others.map(opt).join("")}</optgroup>` : "");
  localStorage.setItem("model.scenario", state.scenario);
  localStorage.setItem("model.base", state.base);
  paintContext();
  updateExportLink();
  if (!state.area && state.page === "home") renderHome();
}

// Colour the working-model chip: slate = locked base, green = mine, purple = mine
// but locked, amber = somebody else's.
function modeOf(id = state.scenario) {
  if (isBaseId(id)) return "base";
  const s = state.scenarios.find(x => x.id === id);
  if (!s) return "base";
  if (!s.mine) return "other";
  return s.locked ? "locked" : "mine";
}
function paintContext() {
  const m = modeOf();
  $("#ctx-work").className = "ctx ctx-work mode-" + m;
  const s = state.scenarios.find(x => x.id === state.scenario);
  const mine = m === "mine" || m === "locked";
  $("#scen-copy").hidden = m === "base";
  $("#scen-rename").hidden = !mine; $("#scen-delete").hidden = !(m === "mine");
  $("#scen-lock").hidden = !mine;
  $("#scen-lock").textContent = s && s.locked ? "Unlock (allow changes again)" : "Lock (finished - stop changes)";
}

async function onScenarioPick(e) {
  if (!confirmDiscard()) { e.target.value = state.scenario; return; }
  state.pending.clear();
  state.scenario = e.target.value;
  localStorage.setItem("model.scenario", state.scenario);
  state.results = null;
  paintContext(); updateExportLink();
  if (state.area) await openArea(state.region, state.area); else showPage(state.page);
}

function needName() {
  if (userName()) return false;
  $("#name-dialog").showModal(); $("#name-input").focus();
  return true;
}

async function scenarioCreate(copy, presetName) {
  if (needName()) return;
  const cur = state.scenarios.find(s => s.id === state.scenario);
  const base = state.bases.find(b => b.id === state.base);
  const name = presetName || prompt(copy ? "Name for your copy:" : `Name for your new working model (on ${base ? base.name : "the base"}):`,
                      copy ? `${userName()}'s copy of ${cur ? cur.name : (base ? base.name : "base")}` : `${userName()} - ${base ? "20" + base.season : ""} working`);
  if (!name) return;
  try {
    const s = await api("POST", "/api/scenarios", { name, user: userName(), base: state.base,
                                                    copy_from: copy && !isBaseId(state.scenario) ? state.scenario : "" });
    state.pending.clear();
    await loadScenarios(s.id);
    if (state.area) await openArea(state.region, state.area); else showPage(state.page);
    toast(`Working model "${s.name}" created. It's yours: only you can change it.`, "ok");
  } catch (err) { fail(err); }
}

async function scenarioRename() {
  const cur = state.scenarios.find(s => s.id === state.scenario);
  const name = prompt("New name:", cur ? cur.name : "");
  if (!name) return;
  try { await api("PATCH", `/api/scenarios/${state.scenario}`, { name, user: userName() }); await loadScenarios(state.scenario); toast("Renamed.", "ok"); }
  catch (err) { fail(err); }
}

async function scenarioLock() {
  const cur = state.scenarios.find(s => s.id === state.scenario);
  if (!cur) return;
  try {
    await api("PATCH", `/api/scenarios/${state.scenario}`, { locked: !cur.locked, user: userName() });
    await loadScenarios(state.scenario);
    if (state.area) await openArea(state.region, state.area);
    toast(cur.locked ? "Unlocked - you can change it again." : "Locked. Nobody, including you, can change it until you unlock it.", "ok");
  } catch (err) { fail(err); }
}

async function scenarioDelete() {
  const cur = state.scenarios.find(s => s.id === state.scenario);
  if (!cur || !confirm(`Delete your working model "${cur.name}" and its ${cur.edits} change(s)?\n\n` +
                       `It is moved to webapp\\data\\scenarios\\_deleted and can be restored by hand.`)) return;
  try {
    await api("DELETE", `/api/scenarios/${state.scenario}?user=${enc(userName())}`);
    state.pending.clear();
    await loadScenarios(BASE_PREFIX + state.base);
    if (state.area) await openArea(state.region, state.area); else showPage(state.page);
    toast("Working model deleted.", "ok");
  } catch (err) { fail(err); }
}

let versionToast = null;
async function pollVersion() {
  if (!state.scenario || state.saving || !state.data) return;
  try {
    const v = await api("GET", `/api/scenarios/${state.scenario}/version`);
    if (v.version > state.version) {
      state.version = v.version;
      if (versionToast) versionToast.remove();
      versionToast = toast(`This working model was changed in another window.`, "", {
        label: "Reload", run: async () => { if (confirmDiscard()) { state.pending.clear(); await openArea(state.region, state.area); } },
      }, true);
    }
  } catch (_) { /* server may be restarting; try again next tick */ }
}

function updateExportLink() {
  const a = $("#export");
  if (!state.region) { a.removeAttribute("href"); return; }
  a.href = `/api/export?scenario=${encodeURIComponent(state.scenario)}&region=${encodeURIComponent(state.region)}`;
}

/* ------------------------------------------------------- navigation / home */

function renderSidebar() {
  const q = state.search;
  const html = state.regions.map(r => {
    const areas = r.areas.filter(a => !q || [a.display.title, a.display.states, a.area, r.display.title]
                                           .join(" ").toLowerCase().includes(q));
    if (!areas.length) return "";
    return `<div class="nav-region">
      <div class="nav-region-head"><span class="swatch" style="background:${regionAccent(r.region)}"></span>${esc(r.display.title)}</div>
      ${areas.map(a => `<button type="button" class="nav-item" data-region="${esc(r.region)}" data-area="${esc(a.area)}"
          title="${esc(a.file)}"${r.region === state.region && a.area === state.area ? ' aria-current="true"' : ""}>
          <span class="t">${esc(a.display.title)}</span>
          <span class="s">${esc([a.display.states || a.display.note, a.display.season].filter(Boolean).join("  ·  "))}</span>
        </button>`).join("")}
    </div>`;
  }).join("");
  $("#areas").innerHTML = html || `<p class="nav-empty">No area matches "${esc(state.search)}".</p>`;
  $$("#areas .nav-item").forEach(b => b.addEventListener("click", () => pickArea(b.dataset.region, b.dataset.area)));
}

function pickArea(region, area) {
  if (!confirmDiscard()) return;
  state.pending.clear();
  openArea(region, area);
}

function goHome() { goPage("home"); }

// Pages: home, bases, compare (an open area is its own view).
function goPage(page) {
  if (!confirmDiscard()) return;
  state.pending.clear();
  state.area = ""; state.region = ""; state.data = null; state.results = null;
  showPage(page);
}
function showPage(page) {
  state.page = page;
  const views = { home: "home", bases: "bases-view", compare: "compare-view" };
  $$("#main > section").forEach(s => s.hidden = s.id !== (state.area ? "area-view" : views[page]));
  $$(".side-link").forEach(b => b.toggleAttribute("aria-current", !state.area && b.dataset.page === page));
  if (!state.area) {
    history.replaceState(null, "", page === "home" ? location.pathname : `#page=${page}`);
    document.title = { home: "Propane Supply Model", bases: "Base models · Propane Supply Model", compare: "Compare · Propane Supply Model" }[page];
    if (page === "home") renderHome();
    if (page === "bases") renderBases();
    if (page === "compare") renderCompareForm();
  }
  renderSidebar(); renderCrumbs(); updateExportLink();
}

function renderHome() {
  const nAreas = state.regions.reduce((a, r) => a + r.areas.length, 0);
  const base = state.bases.find(b => b.id === state.base);
  const mine = state.scenarios.filter(s => s.mine);
  $("#hero-title").textContent = base ? `${base.season ? "20" + base.season.replace("-", "\u2013") + " season" : base.name}` : "Plan the season, one area at a time";
  $("#hero-stats").innerHTML = `
    <div class="hero-stat"><div class="v">${nAreas}</div><div class="k">Areas</div></div>
    <div class="hero-stat"><div class="v">${state.bases.length}</div><div class="k">Base models</div></div>
    <div class="hero-stat"><div class="v">${mine.length}</div><div class="k">My working models</div></div>`;
  const cur = modeOf();
  const step = (n, done, title, body, btn) => `<li class="step${done ? " done" : ""}"><span class="n">${done ? "\u2713" : n}</span>
    <div><strong>${title}</strong><p class="muted">${body}</p>${btn || ""}</div></li>`;
  $("#steps").innerHTML =
    step(1, !!base, "Pick the base", base ? `You're on <strong>${esc(base.name)}</strong>${base.official ? " (the official one)" : ""}. Locked: nobody can change it.` : "No base model yet.",
         `<button type="button" class="btn sm" data-go="bases">See base models</button>`) +
    step(2, cur === "mine", "Make your working model", cur === "mine" ? `Working in <strong>${esc(state.scenarios.find(s => s.id === state.scenario).name)}</strong>.`
         : "Your own copy of the base. Only you can change it.",
         cur === "mine" ? "" : `<button type="button" class="btn sm primary" data-new>New working model</button>`) +
    step(3, false, "Pick an area and run it", "Change inputs, run the model, then look at Results, Insights and What-if.", "");
  $$("#steps [data-go]").forEach(b => b.onclick = () => goPage(b.dataset.go));
  $$("#steps [data-new]").forEach(b => b.onclick = () => scenarioCreate(false));
  const mineHere = mine.filter(s => s.base === state.base);
  $("#my-models").innerHTML = !mineHere.length ? "" : `<h2 class="section-title">My working models</h2><div class="model-cards">` +
    mineHere.map(s => `<button type="button" class="model-card mode-${s.locked ? "locked" : "mine"}${s.id === state.scenario ? " current" : ""}" data-sid="${esc(s.id)}">
      <span class="t">${s.locked ? "\uD83D\uDD12 " : ""}${esc(s.name)}</span>
      <span class="s">${s.edits} change${s.edits === 1 ? "" : "s"}${s.areas.length ? ` in ${s.areas.length} area${s.areas.length === 1 ? "" : "s"}` : ""} \u00b7 ${esc(whenText(s.updated))}</span></button>`).join("") + `</div>`;
  $$("#my-models [data-sid]").forEach(b => b.onclick = async () => {
    state.scenario = b.dataset.sid; await loadScenarios(state.scenario); renderHome();
  });
  $("#region-cards").innerHTML = state.regions.map(r => `
    <article class="region-card" style="--accent:${regionAccent(r.region)}">
      <h2>${esc(r.display.title)}</h2>
      <p class="tag">${esc(r.display.tagline || `${r.areas.length} area${r.areas.length === 1 ? "" : "s"}`)}</p>
      <div class="area-chips">
        ${r.areas.map(a => `<button type="button" class="area-chip" data-region="${esc(r.region)}" data-area="${esc(a.area)}">
            <span><span class="t">${esc(a.display.title)}</span><br>
              <span class="s">${esc([a.display.states || a.display.note, a.display.season].filter(Boolean).join("  ·  "))}</span></span>
            <svg class="go" viewBox="0 0 24 24" aria-hidden="true"><path d="M9 6l6 6-6 6" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          </button>`).join("")}
      </div>
    </article>`).join("");
  $$("#region-cards .area-chip").forEach(b => b.addEventListener("click", () => pickArea(b.dataset.region, b.dataset.area)));
}

// The topbar has no breadcrumb any more - the sidebar and the mode banner say
// where you are. Kept as a no-op so older call sites stay simple.
function renderCrumbs() {}

async function openArea(region, area) {
  const info = areaInfo(region, area);
  const label = info ? info.display.title : area;
  try {
    const data = await busy(`Opening ${label}...`,
      () => api("GET", `/api/area?scenario=${enc(state.scenario)}&region=${enc(region)}&area=${enc(area)}&user=${enc(userName())}`));
    if (state.area !== area || state.region !== region) { state.results = null; state.whatif = null; }
    state.region = region; state.area = area; state.data = data;
    state.builder = null;                      // re-seeded from the saved structure
    state.pctDraft = null;
    state.version = data.scenario.version;
    state.month = Math.min(state.month, data.periods.length - 1);
    $$("#main > section").forEach(s => s.hidden = s.id !== "area-view");
    $$(".side-link").forEach(b => b.removeAttribute("aria-current"));
    history.replaceState(null, "", `#region=${enc(region)}&area=${enc(area)}`);
    document.title = `${data.display.title} · Propane Supply Model`;
    renderAreaHeader();
    updateExportLink();
    renderSidebar();
    renderInputs(); renderResults(); renderChanges();
    showTab(state.tab);
  } catch (err) { fail(err); }
}

// One coloured strip at the top of every area: what you're looking at, and
// whether you can change it.
function renderModeBanner() {
  const d = state.data, s = d.scenario, m = s.is_base ? "base" : !s.can_edit && s.locked && modeOf() !== "other" ? "locked"
    : s.can_edit ? "mine" : "other";
  const icon = { base: "\uD83D\uDD12", locked: "\uD83D\uDD12", mine: "\u270E", other: "\uD83D\uDC41" }[m];
  const title = {
    base: `Locked base model: ${s.base.name}`,
    locked: `Your working model "${s.name}" is locked`,
    mine: `Your working model: ${s.name}`,
    other: `${s.owner || "A shared"} working model: ${s.name}`,
  }[m];
  const text = {
    base: "View only. Make a working model to try changes - the base stays as it is.",
    locked: "Finished and locked, so nothing changes by accident. Unlock it from the \u2026 menu to edit again.",
    mine: `On base ${s.base.name}. Only you can change it. Changes are highlighted in yellow.`,
    other: "View only - it's somebody else's. Make your own copy to try changes.",
  }[m];
  const btn = m === "base" ? `<button type="button" class="btn sm primary" id="mode-new">New working model</button>`
    : m === "other" ? `<button type="button" class="btn sm primary" id="mode-copy">Make my own copy</button>` : "";
  $("#mode-banner").innerHTML = `<div class="mode-banner mode-${m}"><span class="mode-icon" aria-hidden="true">${icon}</span>
    <div><strong>${esc(title)}</strong><p>${esc(text)}</p></div>${btn}</div>`;
  if ($("#mode-new")) $("#mode-new").onclick = () => scenarioCreate(false);
  if ($("#mode-copy")) $("#mode-copy").onclick = () => scenarioCreate(true);
}

/* ----------------------------------------------------------------- upload */
// A manager who worked on their copy of an area workbook offline uploads it;
// the server reads it with the same reader as the official files and puts it
// in a NEW scenario, which this then opens.

function initUpload() {
  const dlg = $("#upload-dialog"), fileIn = $("#upload-file"), sel = $("#upload-area");
  const drop = $("#upload-drop"), guessEl = $("#upload-guess");
  const describe = () => {
    const f = fileIn.files[0];
    $("#upload-file-text").innerHTML = f
      ? `<strong>${esc(f.name)}</strong><br><span class="muted">${NF1.format(f.size / 1048576)} MB - click to choose a different file</span>`
      : `<strong>Choose a file</strong> or drop it here<br><span class="muted">.xlsm or .xlsx, up to 40 MB</span>`;
    const g = f && guessArea(f.name);
    guessEl.textContent = !f ? "" : g ? `Looks like ${g.title} (${g.region}).`
      : "Couldn't tell the area from the file name - pick it above.";
    if (g && !sel.value) sel.value = `${g.region}|${g.area}`;
  };
  $$("[data-upload]").forEach(b => b.addEventListener("click", () => {
    sel.innerHTML = `<option value="">Work it out from the file name</option>` + state.regions.map(r =>
      `<optgroup label="${esc(r.display.title)}">${r.areas.map(a =>
        `<option value="${esc(r.region + "|" + a.area)}">${esc(a.display.title)} - ${esc(a.area)}</option>`).join("")}</optgroup>`).join("");
    if (state.area) sel.value = `${state.region}|${state.area}`;
    fileIn.value = ""; describe();
    dlg.showModal();
  }));
  fileIn.addEventListener("change", describe);
  ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, () => drop.classList.remove("over")));
  drop.addEventListener("drop", e => {
    e.preventDefault();
    if (e.dataTransfer.files.length) { fileIn.files = e.dataTransfer.files; describe(); }
  });
  $("#upload-cancel").addEventListener("click", () => dlg.close());
  $("#upload-form").addEventListener("submit", async e => {
    e.preventDefault();
    const f = fileIn.files[0];
    if (!f) { toast("Choose a workbook first.", "error"); return; }
    if (!/\.(xlsm|xlsx)$/i.test(f.name)) { toast(`"${f.name}" isn't an Excel workbook (.xlsm or .xlsx).`, "error"); return; }
    const user = $("#user").value.trim();
    if (!user) { toast("Add your name at the top right first.", "error"); $("#user").focus(); return; }
    if (!confirmDiscard()) return;
    const fd = new FormData();
    fd.append("file", f); fd.append("user", user); fd.append("base", state.base);
    if (sel.value) { const [r, a] = sel.value.split("|"); fd.append("region", r); fd.append("area", a); }
    try {
      const res = await busy(`Reading ${f.name}... (big areas take a few seconds)`, async () => {
        let r;
        try {
          r = await fetch("/api/upload", { method: "POST", body: fd, headers: { "X-Requested-With": "propane-model" } });
        } catch (_) {
          throw new Error("Couldn't reach the web app. Is the 'Propane Model Web App' window still open?");
        }
        const j = await r.json().catch(() => ({}));
        if (r.status === 404 || r.status === 405) throw new Error(RESTART_MSG);
        if (!r.ok) throw new Error(j.error || `Upload failed (${r.status} ${r.statusText})`);
        return j;
      });
      dlg.close();
      state.pending.clear();
      await loadScenarios(res.scenario);
      state.tab = "inputs";
      await openArea(res.region, res.area);
      toast(`Uploaded. You're now in your new working model "${res.scenario_name}".`, "ok");
    } catch (err) { fail(err); }
  });
}

// Mirror of the server's guess: area code from a workbook file name.
function guessArea(filename) {
  const fold = s => String(s).toLowerCase().replace(/[^a-z0-9]+/g, "");
  let stem = filename.replace(/\.[^.]+$/, "").replace(/\s*\(\d+\)\s*$/, "").replace(/^copy of\s+/i, "");
  stem = stem.replace(/\s*(v\d+\s*)?(working|baseline)\s*\d*$/i, "");
  const all = state.regions.flatMap(r => r.areas.map(a => ({ region: r.region, area: a.area, title: a.display.title })));
  const hits = all.filter(a => fold(a.area) === fold(stem));
  return hits.length === 1 ? hits[0] : null;
}

function renderUploadBanner() {
  const d = state.data, up = d.upload, box = $("#upload-banner");
  if (!up) { box.innerHTML = ""; return; }
  const c = up.compare || { size: [], fields: [] };
  const fmtField = (f, v) => v == null ? "" : VOLUME_FIELDS.has(f) ? NF0.format(v) + " gal"
    : PRICE_FIELDS.has(f) ? "$" + NF4.format(v) + "/gal" : NF0.format(v * 100) + "%";
  const items = [
    ...c.size.map(t => `<li>${esc(t)}</li>`),
    ...c.fields.map(f => `<li><strong>${esc(NAV_LABEL[f.field] || f.label)}</strong>: ${NF0.format(f.cells)} cell${f.cells === 1 ? "" : "s"} differ.
      Biggest: ${esc(f.largest.where)}, base ${esc(fmtField(f.field, f.largest.official))} &rarr; yours ${esc(fmtField(f.field, f.largest.yours))}</li>`),
  ];
  box.innerHTML = `<div class="upload-banner">
    <div><h3>Running on your uploaded file: ${esc(up.name)}</h3>
      <p class="muted small">Uploaded by ${esc(up.user || "someone")} on ${esc(whenText(up.ts))}. "vs base" on the Results tab compares against the base model.</p>
      ${items.length ? `<ul>${items.join("")}</ul>` : `<p class="small" style="margin-top:6px">Every input matches the base model.</p>`}</div>
    ${d.scenario.can_edit ? `<button type="button" class="btn sm" id="upload-remove">Use the base model instead</button>` : ""}</div>`;
  if (!$("#upload-remove")) return;
  $("#upload-remove").onclick = async () => {
    if (!confirm("Switch this area back to the base model? Your uploaded file stays on the server, but this working model stops using it, and any changes typed in on top of it are cleared.")) return;
    try {
      const data = await busy("Switching back...", () => api("POST", "/api/upload/remove",
        { scenario: state.scenario, region: state.region, area: state.area, user: userName() }));
      state.data = data; state.version = data.scenario.version; state.pending.clear(); state.results = null; state.builder = null;
      renderAreaHeader(); renderInputs(); renderResults(); renderChanges();
      toast("Back on the base model.", "ok");
    } catch (err) { fail(err); }
  };
}

function renderAreaHeader() {
  const d = state.data, disp = d.display;
  renderModeBanner();
  renderUploadBanner();
  $("#save-bar").hidden = !d.scenario.can_edit;
  $("#builder-build").disabled = !d.scenario.can_edit; $("#builder-reset").disabled = !d.scenario.can_edit;
  $("#pct-save").disabled = !d.scenario.can_edit; $("#pct-reset").disabled = !d.scenario.can_edit;
  $("#demand-count").textContent = (d.state_pct.changed || []).length || "";
  $("#area-eyebrow").textContent = d.region_display.title;
  $("#area-title").textContent = disp.title;
  const chips = [];
  if (disp.states) chips.push(`<span class="chip">${esc(disp.states)}</span>`);
  if (disp.season) chips.push(`<span class="chip blue">${esc(disp.season)} season</span>`);
  if (disp.note) chips.push(`<span class="chip flame">${esc(disp.note)}</span>`);
  chips.push(`<span class="muted" title="What this area is built from">Based on ${esc(d.source)}</span>`);
  $("#area-sub").innerHTML = chips.join("");
  const nSup = d.supply.filter(s => !isPlaceholder(s.name)).length;
  const nDem = d.demand.filter(n => !isPlaceholder(n)).length;
  const nEd = d.edits.length;
  $("#area-stats").innerHTML = `
    <div class="stat"><div class="k">Supply points</div><div class="v">${nSup}</div></div>
    <div class="stat"><div class="k">Plants served</div><div class="v">${nDem}</div></div>
    <div class="stat"><div class="k">Season demand</div><div class="v">${kgal(d.totals.demand)}</div><div class="d muted">gallons</div></div>
    <div class="stat"><div class="k">Contracted</div><div class="v">${kgal(d.totals.contract)}</div><div class="d muted">gallons</div></div>
    <div class="stat"><div class="k">Changes from the base</div><div class="v">${nEd}</div><div class="d muted">${esc(d.scenario.is_base ? "this is the base" : d.scenario.name)}</div></div>`;
  $("#changes-count").textContent = nEd || "";
  const st = d.structure || {};
  $("#builder-count").textContent = (st.supply || []).length + (st.demand || []).length + (st.hubs || []).length || "";
}

function showTab(name) {
  if (!TABS.includes(name)) name = "inputs";
  state.tab = name;
  TABS.forEach(t => {
    $("#tab-" + t).hidden = t !== name;
    $("#tabbtn-" + t).setAttribute("aria-selected", String(t === name));
  });
  if (name === "changes") renderChanges();
  if (name === "builder") renderBuilder();
  if (name === "demand") renderDemandPct();
  if (name === "insights") renderInsights();
  if (name === "whatif") renderWhatIf();
  if (name === "paste") renderPasteKinds();
  if (state.area) history.replaceState(null, "", `#region=${enc(state.region)}&area=${enc(state.area)}&tab=${name}`);
}

/* ---------------------------------------------------------- model builder */
// The Excel workbook's "Build Optimization Model" panel, as sliders. The
// sliders start at the workbook's own size; sliding up adds rows that are
// named here and then filled in on the Inputs tab like any other cell.

function builderFromData() {
  const d = state.data, st = d.structure || {}, wb = d.workbook_size;
  return {
    periods: st.periods || wb.periods,
    hubs: (st.hubs || []).map(h => ({ name: h.name || "", price: [...(h.price || [0])] })),
    supply: (st.supply || []).map(s => ({ location: s.location || "", supplier: s.supplier || "", hub: s.hub || "" })),
    demand: (st.demand || []).map(x => ({ name: x.name || "" })),
  };
}

// Price hubs a new supply point can buy at: the server's list (CON, MTB-Tet,
// MTB-Non-T, ...) followed by any new hubs named in the builder. A new hub with
// the same name as a listed one replaces it, since it has its own prices.
const foldName = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
function builderHubChoices(d, b) {
  const wb = d.workbook_size;
  const added = b.hubs.map((h, k) => h.name || `Hub ${wb.hubs.length + k + 1}`);
  return [
    ...(d.builder_hubs || wb.hubs.map(h => ({ value: h, label: h })))
      .filter(c => !added.some(a => foldName(a) === foldName(c.value))),
    ...added.map(n => ({ value: n, label: `${n} (new hub)` })),
  ];
}
// The choice a saved hub corresponds to. Older saves used the workbook code
// "MTB", which now means MTB-Tet; anything unknown falls back to the first hub.
function hubPick(saved, names) {
  const hit = names.find(n => foldName(n) === foldName(saved));
  if (hit) return hit;
  const pref = names.find(n => foldName(n).startsWith(foldName(saved)) && foldName(saved));
  return pref || names[0] || "";
}

function builderDirty() {
  return JSON.stringify(state.builder) !== JSON.stringify(builderFromData());
}

function resizeList(list, n, blank) {
  while (list.length < n) list.push(blank(list.length));
  list.length = n;
}

function renderBuilder() {
  if (!state.data) return;
  const d = state.data, wb = d.workbook_size, lim = d.builder_limits;
  if (!state.builder) state.builder = builderFromData();
  const b = state.builder;
  const rows = [
    ["periods", "Time Periods", b.periods, 1, lim.periods, wb.periods],
    ["supply", "Supply Points", wb.supply + b.supply.length, wb.supply, wb.supply + lim.added, wb.supply],
    ["demand", "Demand Points", wb.demand + b.demand.length, wb.demand, wb.demand + lim.added, wb.demand],
    ["hubs", "Price Hubs", wb.hubs.length + b.hubs.length, wb.hubs.length, wb.hubs.length + Math.min(lim.added, 10), wb.hubs.length],
  ];
  $("#builder-sliders").innerHTML = rows.map(([k, label, v, min, max, base]) => `
    <div class="slider-row">
      <label for="bnum-${k}">${label}<span class="base">Base: ${base}</span></label>
      <input type="number" id="bnum-${k}" data-k="${k}" min="${min}" max="${max}" value="${v}" class="${v !== base ? "changed" : ""}">
      <input type="range" id="brange-${k}" data-k="${k}" min="${min}" max="${max}" value="${v}" aria-label="${label}">
    </div>`).join("");

  const setCount = (k, raw) => {
    const row = rows.find(r => r[0] === k);
    const v = Math.max(row[3], Math.min(row[4], Math.round(+raw || row[3])));
    if (k === "periods") b.periods = v;
    else if (k === "supply") resizeList(b.supply, v - wb.supply, () => ({ location: "", supplier: "", hub: "" }));
    else if (k === "demand") resizeList(b.demand, v - wb.demand, () => ({ name: "" }));
    else resizeList(b.hubs, v - wb.hubs.length, () => ({ name: "", price: [0] }));
    renderBuilder();
    $(`#brange-${k}`)?.focus();
  };
  $$("#builder-sliders input[type=range]").forEach(r => r.addEventListener("input", () => {
    $(`#bnum-${r.dataset.k}`).value = r.value;
  }));
  $$("#builder-sliders input[type=range]").forEach(r => r.addEventListener("change", () => setCount(r.dataset.k, r.value)));
  $$("#builder-sliders input[type=number]").forEach(n => n.addEventListener("change", () => setCount(n.dataset.k, n.value)));

  // names for whatever was added
  const hubChoices = builderHubChoices(d, b), hubNames = hubChoices.map(c => c.value);
  const P = b.periods, labels = periodLabels(P);
  let h = "";
  if (b.supply.length) {
    h += `<div class="builder-section"><h3>New supply points (contracts)</h3>
      <p class="hint">Name each one. After building, fill in contract volume, differential and lift limits on the Inputs tab,
      and freight under Lanes. New lanes start with blank freight (not used), so nothing moves until you price them.</p>
      <table class="builder-table"><thead><tr><th></th><th>Supply location</th><th>Supplier</th><th>Price hub</th></tr></thead><tbody>` +
      b.supply.map((s, k) => `<tr><td class="n">${k + 1}</td>
        <td><input type="text" data-list="supply" data-k="${k}" data-f="location" value="${esc(s.location)}" placeholder="e.g. Northfield, ST" aria-label="Supply location ${k + 1}"></td>
        <td><input type="text" data-list="supply" data-k="${k}" data-f="supplier" value="${esc(s.supplier)}" placeholder="e.g. SupplierX" aria-label="Supplier ${k + 1}"></td>
        <td><select data-list="supply" data-k="${k}" data-f="hub" aria-label="Price hub ${k + 1}">${hubChoices.map(c =>
          `<option value="${esc(c.value)}"${hubPick(s.hub, hubNames) === c.value ? " selected" : ""}>${esc(c.label)}</option>`).join("")}</select></td></tr>`).join("") +
      `</tbody></table></div>`;
  }
  if (b.demand.length) {
    h += `<div class="builder-section"><h3>New plants (demand points)</h3>
      <p class="hint">After building, enter their demand on the Inputs tab and freight to them under Lanes.</p>
      <table class="builder-table"><thead><tr><th></th><th>Plant name</th></tr></thead><tbody>` +
      b.demand.map((x, k) => `<tr><td class="n">${k + 1}</td>
        <td><input type="text" data-list="demand" data-k="${k}" data-f="name" value="${esc(x.name)}" placeholder="e.g. 1001-Springfield ST" aria-label="Plant name ${k + 1}"></td></tr>`).join("") +
      `</tbody></table></div>`;
  }
  if (b.hubs.length) {
    h += `<div class="builder-section"><h3>New price hubs</h3>
      <p class="hint">Monthly hub price, $ per gallon (e.g. 0.75). Type one value in the first month to use it for every month.</p>
      <div style="overflow-x:auto"><table class="builder-table"><thead><tr><th></th><th>Hub name</th>${labels.map(l => `<th>${esc(l)}</th>`).join("")}</tr></thead><tbody>` +
      b.hubs.map((x, k) => `<tr><td class="n">${k + 1}</td>
        <td><input type="text" data-list="hubs" data-k="${k}" data-f="name" value="${esc(x.name)}" placeholder="e.g. Conway" aria-label="Hub name ${k + 1}" style="width:140px"></td>` +
        labels.map((l, p) => `<td><input class="price" data-list="hubs" data-k="${k}" data-f="price" data-p="${p}" inputmode="decimal"
          value="${x.price.length === 1 ? (p === 0 ? esc(inputText(toShow("hub_price", x.price[0]))) : "") : esc(inputText(toShow("hub_price", x.price[p] ?? null)))}"
          placeholder="${x.price.length === 1 && p > 0 ? esc(inputText(toShow("hub_price", x.price[0]))) : ""}" aria-label="${esc(x.name || "Hub " + (k + 1))} ${esc(l)}, $ per gallon"></td>`).join("") +
        `</tr>`).join("") + `</tbody></table></div></div>`;
  }
  $("#builder-lists").innerHTML = h;
  $$("#builder-lists [data-list]").forEach(el => el.addEventListener("change", () => {
    const item = b[el.dataset.list][+el.dataset.k];
    if (el.dataset.f === "price") {
      const p = +el.dataset.p, shown = el.value.trim() === "" ? null : Number(el.value.replace(/[$,]/g, ""));
      if (shown != null && !isFinite(shown)) { toast(`"${el.value}" is not a price`, "error"); return; }
      const v = toModel("hub_price", shown);       // typed in $/gal, stored in $/kgal
      if (p === 0 && item.price.length === 1) item.price = [v ?? 0];
      else {
        const full = item.price.length === 1 ? new Array(P).fill(item.price[0]) : [...item.price];
        full[p] = v ?? full[0]; item.price = full;
      }
    } else item[el.dataset.f] = el.value;
    if (el.dataset.list === "hubs" && el.dataset.f === "name") renderBuilder();
    else updateBuilderStatus();
  }));
  updateBuilderStatus();
}

function periodLabels(P) {
  const have = state.data.periods;
  const out = have.slice(0, P);
  const last = have[have.length - 1];
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const m = /^([A-Z][a-z]{2})-(\d{2})$/.exec(last || "");
  for (let k = out.length; k < P; k++) {
    if (m) {
      const idx = MON.indexOf(m[1]) + (k - have.length + 1), y = +m[2] + Math.floor(idx / 12);
      out.push(`${MON[idx % 12]}-${String(y % 100).padStart(2, "0")}`);
    } else out.push(`P${k + 1}`);
  }
  return out;
}

function updateBuilderStatus() {
  const d = state.data, st = d.structure || {}, meta = d.structure_meta || {};
  const dirty = builderDirty();
  const added = (st.supply || []).length + (st.demand || []).length + (st.hubs || []).length + (st.periods ? 1 : 0);
  $("#builder-build").disabled = !dirty;
  $("#builder-reset").disabled = !added && !dirty;
  $("#builder-status").textContent = dirty ? "Not built yet - press Build to apply."
    : added ? `Built${meta.user ? " by " + meta.user : ""}${meta.ts ? " on " + whenText(meta.ts) : ""}.` : "Same size as the base.";
  $("#builder-count").textContent = (st.supply || []).length + (st.demand || []).length + (st.hubs || []).length || "";
}

async function buildModel(structure) {
  const user = $("#user").value.trim();
  if (!user) { toast("Add your name at the top right first.", "error"); $("#user").focus(); return; }
  if (state.pending.size && !confirm("You have unsaved grid changes. Building will discard them. Continue?")) return;
  try {
    const data = await busy("Building model...", () => api("POST", "/api/structure", {
      scenario: state.scenario, region: state.region, area: state.area, user, structure,
    }));
    state.pending.clear();
    state.data = data; state.version = data.scenario.version; state.results = null;
    state.builder = builderFromData();
    renderAreaHeader(); renderInputs(); renderResults(); renderChanges(); renderBuilder();
    await loadScenarios(state.scenario);
    const w = data.workbook_size;
    toast(`Model built: ${data.periods.length} periods, ${data.supply.length} supply points ` +
          `(${data.supply.length - w.supply} new), ${data.demand.length} plants.` +
          (data.dropped_edits ? ` ${data.dropped_edits} cell change(s) outside the new size were removed.` : ""), "ok");
  } catch (err) { fail(err); }
}

/* ---------------------------------------------------------------- inputs */

function renderInputs() {
  const d = state.data;
  $("#month-pick").innerHTML = d.periods.map((p, k) => `<option value="${k}"${k === state.month ? " selected" : ""}>${esc(p)}</option>`).join("");
  $("#hide-unused").checked = state.hideUnused;
  const msgs = [
    ...d.repairs.map(t => ["Error", "Data error in the base data (the model works around it): " + t]),
    ...d.feasibility.map(t => ["Warning", t]),
    ...d.notes.map(t => ["Warning", t]),
  ];
  $("#warnings").innerHTML = !msgs.length ? "" :
    `<details${msgs.some(m => m[0] === "Error") ? " open" : ""}><summary>${msgs.length} thing${msgs.length === 1 ? "" : "s"} to check in this area's data</summary>` +
    msgs.map(([s, t]) => `<div class="msg ${s}">${esc(t)}</div>`).join("") + "</details>";
  renderTableNav();
  renderGrid();
}

function renderTableNav() {
  const d = state.data, counts = {};
  for (const e of d.edits) counts[e.field] = (counts[e.field] || 0) + 1;
  for (const p of state.pending.values()) counts[p.field] = (counts[p.field] || 0) + (p.revert ? 0 : 0);
  // Colour per group, used again as the grid's header stripe: supply blue, demand teal, lanes orange.
  const gcls = { Supply: "g-supply", Demand: "g-demand", Lanes: "g-lanes" };
  $("#table-nav").innerHTML = TABLES.map(g => `<h3 class="${gcls[g.group]}"><i></i>${esc(g.group)}</h3>` + g.items.map(f => {
    const fi = d.field_info[f];
    return `<button type="button" class="${gcls[g.group]}" data-table="${f}"${f === state.table ? ' aria-current="true"' : ""}>
      <span>${esc(NAV_LABEL[f] || fi.label)}</span>
      ${counts[f] ? `<span class="n" title="${counts[f]} saved change(s)">${counts[f]}</span>` : (fi.editable ? "" : '<span class="ro">calc</span>')}
    </button>`;
  }).join("")).join("");
  $$("#table-nav button").forEach(b => b.addEventListener("click", () => {
    state.table = b.dataset.table; renderTableNav(); renderGrid();
  }));
  const grp = TABLES.find(g => g.items.includes(state.table));
  $(".grid-card").className = "card grid-card " + (grp ? gcls[grp.group] : "");
}

function supplyLabel(i, d = state.data) {
  const s = d.supply[i];
  return s.group ? `${s.name} | ${s.group}` : s.name;
}

function gridModel(field) {
  const d = state.data, axes = AXES[field], vals = d.fields[field];
  const keep = (name) => !(state.hideUnused && isPlaceholder(name));
  const supRows = d.supply.map((s, i) => ({ i, main: s.name, meta: [s.group, s.hub].filter(Boolean).join(" · "), full: supplyLabel(i), name: s.name }))
                   .filter(r => keep(r.name));
  const demRows = d.demand.map((n, j) => { const p = prettyLoc(n); return { i: j, main: p.main, meta: p.meta, full: n, name: n }; })
                   .filter(r => keep(r.name));
  if (axes.length === 3) {
    const p = state.month;
    return { rows: supRows, cols: demRows.map(r => ({ i: r.i, label: r.main, full: r.full })), lane: true,
             get: (r, c) => vals[p][r][c], index: (r, c) => [p, r, c] };
  }
  const rows = axes.includes("i") ? supRows : demRows;
  if (axes.length === 1) {
    return { rows, cols: [{ i: 0, label: "Value" }], get: (r) => vals[r], index: (r) => [r] };
  }
  return { rows, cols: d.periods.map((p, k) => ({ i: k, label: p })), get: (r, c) => vals[c][r], index: (r, c) => [c, r] };
}

function editsMap() {
  const m = new Map();
  for (const e of state.data.edits) m.set(keyOf(e.field, e.index), e);
  return m;
}

// What a cell should show right now: unsaved value > saved edit > model value.
function cellValue(field, index, current, edits) {
  const k = keyOf(field, index), p = state.pending.get(k), e = edits.get(k);
  if (p) return p.revert ? (e ? e.base : current) : p.value;
  if (e) return e.value;
  return current;
}

function renderGrid() {
  const d = state.data, field = state.table;
  // Not yours (or a locked base): every cell is shown read-only.
  const fi = { ...d.field_info[field], editable: d.field_info[field].editable && d.scenario.can_edit };
  const gm = gridModel(field), edits = editsMap();
  $("#table-title").textContent = `${NAV_LABEL[field] || fi.label}`;
  $("#table-desc").textContent = DESC[field] || "";
  $("#month-wrap").hidden = !gm.lane;
  const freight = d.fields.trans_cost, noMax = d.no_max, highRate = d.high_rate ?? 1000;
  const totals = TOTALLED.has(field) && !gm.lane;
  const unit = UNIT(field);
  const show = v => toShow(field, v);          // every number below is shown in gal / $/gal

  // Lanes: the gallons the last run moved on each lane in this month, under every cell.
  const sv = gm.lane ? laneShipments(state.month) : null;
  const showShip = !!sv && state.showShip;
  $("#ship-toggle").hidden = !gm.lane || !sv;
  $("#lane-note").innerHTML = !gm.lane ? (field === "demand" && fi.editable
      ? `<div class="msg Info qd-hint">Changing demand for a season, a state or a few plants?
          <button type="button" class="btn sm" id="go-quick-demand">Use Quick demand change</button>
          <span class="muted small">Pick Winter / Summer, pick where, pick +10% - no maths.</span></div>` : "") : !sv
    ? `<div class="msg Info">Run the model to see the exact gallons moved on every lane, month by month, right in this table.</div>`
    : `<div class="msg ${sv.stale ? "Warning" : "Info"}"><strong>Shipped gallons</strong> under each rate are from the run at
        ${esc(sv.at)} for <strong>${esc(d.periods[state.month])}</strong>${sv.stale ? " - you've changed inputs since then, run the model again for current volumes" : ""}.
        Lanes with freight of $1.00/gal or more are <span class="hf-key">marked</span>: they're used like any other lane, so check the rate.</div>`;

  let h = `<table class="grid${showShip ? " with-ship" : ""}" aria-label="${esc(fi.label)}"><thead><tr><th class="rowhead" scope="col">${gm.lane ? "From supply  \u2192  to plant" : (AXES[field].includes("i") ? "Supply point" : "Plant")}</th>`;
  // Month columns: winter months get a blue header so the season is visible at a glance.
  const monthly = !gm.lane && AXES[field].includes("p");
  h += gm.cols.map((c, k) => `<th scope="col"${gm.lane ? ` class="lane" title="${esc(c.full)}"` : monthly ? ` class="${isWinter(c.label, k) ? "mh-w" : "mh-s"}"` : ""}>${esc(c.label)}</th>`).join("");
  if (totals) h += `<th scope="col">Season</th>`;
  if (showShip) h += `<th scope="col" class="shipcol">Shipped from<br>this supply, gal</th>`;
  h += `</tr></thead><tbody>`;
  const colTotals = new Array(gm.cols.length).fill(0), colShip = new Array(gm.cols.length).fill(0);
  gm.rows.forEach((row, r) => {
    h += `<tr><th class="rowhead" scope="row" title="${esc(row.full)}"><span class="loc">${esc(row.main)}</span>${row.meta ? `<span class="meta">${esc(row.meta)}</span>` : ""}</th>`;
    let rowTotal = 0, rowShip = 0;
    gm.cols.forEach((col, c) => {
      const idx = gm.index(row.i, col.i), k = keyOf(field, idx);
      const v = cellValue(field, idx, gm.get(row.i, col.i), edits);      // model units
      const dv = show(v);                                                  // display units
      rowTotal += +dv || 0; colTotals[c] += +dv || 0;
      const cls = [];
      if (state.pending.has(k)) cls.push("pending"); else if (edits.has(k)) cls.push("edited");
      // Blank freight = lane not used (shown greyed on every lane table).
      const rate = gm.lane ? (field === "trans_cost" ? v : freight[idx[0]][idx[1]][idx[2]]) : null;
      const closed = gm.lane && rate == null;
      if (closed) cls.push("closed");
      // High freight ($1.00/gal or more): a warning only - the lane is used like any other.
      const high = gm.lane && !closed && rate >= highRate - 1e-6;
      if (high) cls.push("high");
      if (!v) cls.push("zero");
      const e = edits.get(k);
      let tip = e ? `Base value ${editText(field, e.base)}. Changed to ${editText(field, e.value)} by ${e.user || "someone"} on ${String(e.ts).replace("T", " at ")}` : "";
      if (high) tip = `High freight: ${perGal(rate)}/gal. Used like any other lane when it's the cheapest way - check the rate.` + (tip ? "\n" + tip : "");
      let shipHtml = "";
      if (showShip) {
        const s = sv.map.get(`${row.i},${col.i}`);
        if (s) { rowShip += s[0]; colShip[c] += s[0]; }
        shipHtml = closed ? `<span class="shipv none">not used</span>`
          : s ? `<span class="shipv${high ? " hf" : ""}" title="${esc(`${d.periods[state.month]}: ${galExact(s[0])} gal from ${row.full} to ${col.full}, freight ${money(s[2])}`)}">${kgal(s[0])} gal</span>`
          : `<span class="shipv none">0 gal</span>`;
        cls.push("has-ship");
      }
      if (!fi.editable) {
        h += `<td class="ro ${cls.join(" ")}"${tip ? ` title="${esc(tip)}"` : ""}><span class="rov">${esc(cellText(dv ?? 0, field)) || (closed ? "not used" : "")}</span>${shipHtml}</td>`;
      } else {
        const unbounded = field === "ship_max" && v >= noMax;
        const ph = unbounded ? "no limit" : field === "trans_cost" && v == null ? "not used" : "";
        h += `<td class="${cls.join(" ")}"${tip ? ` title="${esc(tip)}"` : ""}><input data-r="${r}" data-c="${c}" inputmode="decimal"
               value="${unbounded ? "" : esc(cellText(dv, field))}" data-full="${unbounded ? "" : esc(inputText(dv))}"${ph ? ` placeholder="${ph}"` : ""}
               aria-label="${esc(row.full)}, ${esc(col.full || col.label)}, ${esc(unit)}">${shipHtml}</td>`;
      }
    });
    if (totals) h += `<td class="num totalcol">${NF0.format(rowTotal)}</td>`;
    if (showShip) h += `<td class="num totalcol shipcol">${kgal(rowShip)}</td>`;
    h += `</tr>`;
  });
  if (totals) {
    h += `<tr class="total"><th class="rowhead" scope="row">Total</th>${colTotals.map(t => `<td class="num">${NF0.format(t)}</td>`).join("")}<td class="num">${NF0.format(colTotals.reduce((a, b) => a + b, 0))}</td></tr>`;
  }
  if (showShip) {
    h += `<tr class="total"><th class="rowhead" scope="row">Delivered to this plant, gal</th>${colShip.map(t => `<td class="num">${kgal(t)}</td>`).join("")}<td class="num shipcol">${kgal(colShip.reduce((a, b) => a + b, 0))}</td></tr>`;
  }
  h += `</tbody></table>`;
  if (!gm.rows.length) h = `<div class="empty-state"><div class="big">Nothing to show</div>Every row here is an unused placeholder. Untick "Hide unused rows" to see them.</div>`;
  const wrap = $("#grid-wrap");
  wrap.innerHTML = h;
  if ($("#go-quick-demand")) $("#go-quick-demand").onclick = () => showTab("demand");
  wrap.onchange = e => { if (e.target.matches("input")) commit(e.target); };
  // Full precision while editing, tidy display otherwise. Swapping the value on
  // focus happens before the user types, so it never counts as a change.
  // (addEventListener, bound once: browsers have no onfocusin property.)
  if (!wrap.dataset.focusBound) {
    wrap.dataset.focusBound = "1";
    wrap.addEventListener("focusin", e => {
      const t = e.target;
      if (t.matches("input") && t.dataset.full && t.value === cellText(+t.dataset.full, state.table)) {
        t.value = t.dataset.full; t.select();
      }
    });
    wrap.addEventListener("focusout", e => {
      const t = e.target;
      if (t.matches("input") && t.dataset.full && t.value === t.dataset.full) t.value = cellText(+t.dataset.full, state.table);
    });
  }
  wrap.onkeydown = onGridKey;
  wrap.onpaste = onGridPaste;
  state.gm = gm;
  updatePending();
}

function parseEntry(raw, field) {
  let s = String(raw).trim().replace(/[$,\s]/g, "");
  if (s === "") return { empty: true };
  let pct = false;
  if (s.endsWith("%")) { pct = true; s = s.slice(0, -1); }
  if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return { error: `"${raw}" is not a number` };
  let v = Number(s);
  if (pct) v = v / 100;
  if (!isFinite(v)) return { error: `"${raw}" is not a number` };
  if (v < 0 && field !== "hub_premium") return { error: `${NAV_LABEL[field] || field} can't be negative` };
  return { value: v };
}

function commit(input) {
  const field = state.table, gm = state.gm;
  const row = gm.rows[+input.dataset.r], col = gm.cols[+input.dataset.c];
  const idx = gm.index(row.i, col.i), k = keyOf(field, idx);
  const td = input.parentElement, edits = editsMap(), e = edits.get(k);
  const cur = gm.get(row.i, col.i);                       // model units
  const parsed = parseEntry(input.value, field);          // what was typed: gal or $/gal
  if (parsed.value != null) parsed.value = toModel(field, parsed.value);
  td.classList.remove("bad"); input.removeAttribute("aria-invalid");
  if (parsed.error) {
    td.classList.add("bad"); input.setAttribute("aria-invalid", "true");
    toast(parsed.error, "error");
    return;
  }
  const isBlank = v => v == null;
  if (field === "trans_cost") {
    // Freight: blank is a value of its own - "lane not used". So clearing a
    // freight cell closes the lane rather than restoring the workbook value
    // (use Revert on the Changes tab for that).
    const want = parsed.empty ? null : parsed.value;
    const now = e ? e.value : cur;
    const same = isBlank(want) || isBlank(now) ? isBlank(want) && isBlank(now)
      : Math.abs(want - now) <= 1e-9 * Math.max(1, Math.abs(now));
    if (same) state.pending.delete(k);
    else if (e && (isBlank(want) ? isBlank(e.base) : !isBlank(e.base) && Math.abs(want - e.base) <= 1e-9 * Math.max(1, Math.abs(e.base))))
      state.pending.set(k, { field, index: idx, revert: true });       // back to the workbook value
    else state.pending.set(k, { field, index: idx, value: want });
  } else if (parsed.empty) {
    if (field === "ship_max" && !e) {          // blank max = no limit
      if (cur < state.data.no_max) state.pending.set(k, { field, index: idx, value: state.data.no_max });
      else state.pending.delete(k);
    } else if (e) state.pending.set(k, { field, index: idx, revert: true });
    else state.pending.delete(k);
  } else {
    const same = (a, b) => Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b));
    if (e ? same(parsed.value, e.value) : same(parsed.value, cur)) state.pending.delete(k);
    else state.pending.set(k, { field, index: idx, value: parsed.value });
  }
  td.classList.toggle("pending", state.pending.has(k));
  td.classList.toggle("edited", !state.pending.has(k) && !!e);
  if (field === "trans_cost") {
    const shown = cellValue(field, idx, cur, edits);
    td.classList.toggle("closed", shown == null);
    input.placeholder = shown == null ? "not used" : "";
  }
  if (state.pending.has(k) && state.pending.get(k).revert) {
    input.value = cellText(toShow(field, e.base), field);
    input.dataset.full = inputText(toShow(field, e.base));
  }
  updatePending();
}

function updatePending() {
  const n = state.pending.size;
  $("#pending").textContent = n ? `${n} unsaved change${n === 1 ? "" : "s"}` : "";
  $("#save").disabled = !n; $("#discard").disabled = !n;
}

function cellInput(r, c) { return $(`#grid-wrap input[data-r="${r}"][data-c="${c}"]`); }

function onGridKey(e) {
  const t = e.target;
  if (!t.matches("input")) return;
  const r = +t.dataset.r, c = +t.dataset.c;
  let next = null;
  if (e.key === "Enter") next = cellInput(e.shiftKey ? r - 1 : r + 1, c);
  else if (e.key === "ArrowDown") next = cellInput(r + 1, c);
  else if (e.key === "ArrowUp") next = cellInput(r - 1, c);
  else if (e.key === "Escape") { renderGrid(); cellInput(r, c)?.focus(); return; }
  else return;
  e.preventDefault();
  commit(t);
  if (next) { next.focus(); next.select(); }
}

// Paste a block copied from Excel: tab-separated columns, one line per row.
function onGridPaste(e) {
  const t = e.target;
  if (!t.matches("input")) return;
  const text = (e.clipboardData || window.clipboardData).getData("text");
  if (!/[\t\n]/.test(text.trim())) return;          // single value: let the browser paste it
  e.preventDefault();
  const lines = text.replace(/\r/g, "").split("\n");
  while (lines.length && lines[lines.length - 1] === "") lines.pop();
  const r0 = +t.dataset.r, c0 = +t.dataset.c;
  let n = 0;
  lines.forEach((line, dr) => line.split("\t").forEach((val, dc) => {
    const inp = cellInput(r0 + dr, c0 + dc);
    if (inp) { inp.value = val; commit(inp); n++; }
  }));
  toast(`Pasted ${n} value${n === 1 ? "" : "s"}. Press Save changes to keep them.`, "ok");
}

async function save() {
  const user = $("#user").value.trim();
  if (!user) {
    toast("Add your name at the top right first, so the change log shows who made the change.", "error");
    $("#user").focus();
    return;
  }
  localStorage.setItem("model.user", user);
  const changes = [...state.pending.values()];
  state.saving = true;
  try {
    const data = await busy("Saving...", () => api("POST", "/api/edits", {
      scenario: state.scenario, region: state.region, area: state.area, user, changes,
    }));
    state.data = data; state.version = data.scenario.version; state.pending.clear();
    renderAreaHeader(); renderInputs(); renderChanges();
    await loadScenarios(state.scenario);
    toast(`Saved ${changes.length} change${changes.length === 1 ? "" : "s"} to "${data.scenario.name}".`, "ok");
  } catch (err) { fail(err); }
  finally { state.saving = false; }
}

/* --------------------------------------------------------------- results */

async function run(oneArea) {
  if (state.pending.size && !confirm("You have unsaved changes. They will NOT be included in this run. Continue?")) return;
  try {
    const body = { scenario: state.scenario, region: state.region };
    if (oneArea) body.area = state.area;
    const label = oneArea ? state.data.display.title : state.data.region_display.title;
    state.results = await busy(oneArea ? `Optimizing ${label}...` : `Optimizing every area in ${label}...`,
                               () => api("POST", "/api/solve", body));
    state.results.at = new Date();
    state.results.version = state.version;           // to tell when lane volumes are out of date
    if (state.tab === "inputs") renderGrid();
    state.resultArea = oneArea ? state.area : (state.results.results.find(r => r.area === state.area) || state.results.results[0]).area;
    renderResults();
    if (state.tab !== "insights") showTab("results"); else renderInsights();
  } catch (err) { fail(err); }
}

function statusChip(s) {
  if (!s.ok) return `<span class="chip red">${s.status === "error" ? "Error" : "Did not solve"}</span>`;
  return s.relaxed ? `<span class="chip amber">Solved with gaps</span>` : `<span class="chip green">Optimal</span>`;
}

// Change in cost per gallon vs the workbook, e.g. "▼ $0.0123/gal vs workbook".
function deltaPerGal(cur, base) {
  if (base == null || cur == null) return "";
  const d = (cur - base) / GAL;
  if (Math.abs(d) < 0.00005) return `<div class="d muted">Same as base</div>`;
  return `<div class="d ${d < 0 ? "down" : "up"}">${d < 0 ? "\u25BC" : "\u25B2"} $${NF4.format(Math.abs(d))}/gal vs base</div>`;
}

function deltaText(cur, base) {
  if (base == null || cur == null) return "";
  const d = cur - base, pct = base ? d / base * 100 : 0;
  if (Math.abs(d) < 0.5) return `<div class="d muted">Same as base</div>`;
  return `<div class="d ${d < 0 ? "down" : "up"}">${d < 0 ? "\u25BC" : "\u25B2"} ${moneyShort(Math.abs(d))} (${pct >= 0 ? "+" : ""}${NF1.format(pct)}%) vs base</div>`;
}

function renderResults() {
  const box = $("#results");
  if (!state.results) {
    box.innerHTML = `<div class="card empty-state">
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>
      <div class="big">No results yet</div>
      <p>Press <strong>Run model</strong> to optimize ${esc(state.data ? state.data.display.title : "this area")} with the "${esc(state.data ? state.data.scenario.name : "")}".</p></div>`;
    return;
  }
  const res = state.results.results;
  let h = `<div class="results-head"><div><h2>Results</h2><p class="muted">"${esc(state.results.scenario)}"${state.results.at ? " · run at " + state.results.at.toLocaleTimeString() : ""}</p></div></div>`;
  if (res.length > 1) {
    let tot = 0, base = 0, anyBase = false;
    for (const r of res) { tot += r.summary.total || 0; base += (r.baseline ? r.baseline.total : r.summary.total) || 0; anyBase = anyBase || !!r.baseline; }
    h += `<div class="card" style="margin-bottom:16px"><div class="card-head"><div><h2>${esc(state.results.region_display.title)}: every area</h2>
          <p class="muted">Region total ${money(tot)}${anyBase ? " · " + deltaText(tot, base).replace(/<[^>]+>/g, "") : ""}</p></div></div>
      <div class="scroll"><table class="list"><thead><tr><th>Area</th><th>Status</th><th class="num">Total cost</th><th class="num">Commodity</th>
      <th class="num">Freight</th><th class="num">Avg $/gal</th><th class="num">Unmet gal</th><th class="num">High freight (&ge; $1/gal), gal</th><th>vs base</th></tr></thead><tbody>`;
    for (const r of res) {
      const s = r.summary, sel = r.area === state.resultArea;
      h += `<tr class="clickable${sel ? " selected" : ""}" data-view="${esc(r.area)}" tabindex="0">
        <td><strong>${esc(r.display.title)}</strong></td><td>${statusChip(s)}</td>
        <td class="num">${money(s.total)}</td><td class="num">${money(s.commodity)}</td><td class="num">${money(s.freight)}</td>
        <td class="num">${perGal(s.avg)}</td><td class="num">${kgal(s.unmet)}</td><td class="num">${kgal(s.blocked_volume)}</td>
        <td>${r.baseline ? deltaText(s.total, r.baseline.total) : '<span class="muted">Same as base</span>'}</td></tr>`;
    }
    h += `</tbody></table></div></div>`;
  }
  const r = res.find(x => x.area === state.resultArea) || res[0];
  h += renderOneResult(r);
  box.innerHTML = h;
  $$("#results tr[data-view]").forEach(tr => {
    const go = () => { state.resultArea = tr.dataset.view; renderResults(); };
    tr.onclick = go; tr.onkeydown = e => { if (e.key === "Enter") go(); };
  });
  $$("#results button[data-rview]").forEach(b => b.onclick = () => { state.resultView = b.dataset.rview; renderResults(); });
  $$("#results button[data-chart]").forEach(b => b.onclick = () => {
    state.chartMode = b.dataset.chart;
    localStorage.setItem("model.chartMode", state.chartMode);
    renderResults();
  });
  $$("#results button[data-parts]").forEach(b => b.onclick = () => {
    state.chartParts = b.dataset.parts;
    localStorage.setItem("model.chartParts", state.chartParts);
    renderResults();
  });
  $$("#results select[data-metric]").forEach(s => s.onchange = () => { state["metric_" + s.dataset.metric] = s.value; renderResults(); });
  const lm = $("#lane-month"), le = $("#lane-empty");
  if (lm) lm.onchange = () => { state.laneMonth = lm.value === "all" ? "all" : +lm.value; renderResults(); };
  if (le) le.onchange = () => { state.laneEmpty = le.checked; renderResults(); };
  const sm = $("#ship-month"), sf = $("#ship-filter");
  if (sm) sm.onchange = () => { state.shipMonth = sm.value; renderResults(); };
  if (sf) sf.onchange = () => { state.shipFilter = sf.value; renderResults(); };
}

function kpi(label, value, extra = "", cls = "") {
  return `<div class="kpi ${cls}"><div class="k">${label}</div><div class="v">${value}</div>${extra}</div>`;
}

function renderOneResult(r) {
  const s = r.summary, b = r.baseline;
  let h = `<div class="results-head"><div><p class="eyebrow">${esc(r.display.states || "")}</p><h2 style="font-size:20px">${esc(r.display.title)}</h2></div>
           <div>${statusChip(s)} <span class="muted">${esc(s.solver || "")}${s.seconds != null ? " · " + s.seconds + "s" : ""}</span></div></div>`;
  if (s.ok) {
    h += `<div class="kpis">
      ${kpi("Total cost", money(s.total), b ? deltaText(s.total, b.total) : '<div class="d" style="color:#aab7d6">The base model as is</div>', "hero-kpi")}
      ${kpi("Commodity", money(s.commodity), b ? deltaText(s.commodity, b.commodity) : "")}
      ${kpi("Freight", money(s.freight), b ? deltaText(s.freight, b.freight) : "")}
      ${kpi("Cost per gallon", perGal(s.avg), b && b.avg != null ? deltaPerGal(s.avg, b.avg) : '<div class="d muted">commodity + freight, all in</div>')}
      ${kpi("Purchased", kgal(s.purchased), `<div class="d muted">gal, of ${kgal(s.demand)} gal demand</div>`)}
      ${kpi("Unmet demand", kgal(s.unmet), '<div class="d muted">gal</div>', s.unmet > 1e-6 ? "bad" : "")}
      ${kpi("Surplus", kgal(s.surplus), '<div class="d muted">gal with nowhere to go</div>', s.surplus > 1e-6 ? "warn" : "")}
      ${kpi("High-freight lanes", kgal(s.blocked_volume), `<div class="d muted">gal on lanes at $1.00/gal or more${s.blocked_volume >= 0.1 ? " - check those rates" : ""}</div>`, s.blocked_volume >= 0.1 ? "warn" : "")}
    </div>`;
  }
  if (s.relax_note) h += `<div class="msg Warning">${esc(s.relax_note)}</div>`;
  if (!s.ok && s.message) h += `<div class="msg Error">${esc(s.message)}</div>`;
  if (s.ok) h += costChart(r);

  const views = s.ok
    ? [["plan", "Supply plan"], ["lanes", "Lane volumes"], ["ship", "Shipments"], ["demand", "Demand coverage"], ["sinv", "Supply inventory"], ["diag", "Diagnostics"]]
    : [["diag", "Diagnostics"]];
  if (!views.some(v => v[0] === state.resultView)) state.resultView = views[0][0];
  h += `<div class="pills" role="group" aria-label="Result views">` +
       views.map(([k, l]) => `<button type="button" data-rview="${k}" aria-pressed="${k === state.resultView}">${l}</button>`).join("") + `</div>`;

  const supRows = r.supply.map((x, i) => ({ i, main: x.name, meta: [x.group, x.hub].filter(Boolean).join(" · "), full: x.group ? `${x.name} | ${x.group}` : x.name, ph: isPlaceholder(x.name) }));
  const demRows = r.demand.map((n, i) => { const p = prettyLoc(n); return { i, main: p.main, meta: p.meta, full: n, ph: isPlaceholder(n) }; });
  h += `<div class="card">`;
  if (state.resultView === "plan") {
    h += matrixView(r, "plan", r.supply_plan, supRows, [
      ["purchased", "Purchased (gal)", kgal, true], ["perc", "Lift (share of contract)", v => NF0.format(v * 100) + "%", false],
      ["contract", "Contract (gal)", kgal, true], ["price", "Delivered price ($/gal)", perGal, false],
      ["cost", "Commodity cost", money, true], ["surplus", "Surplus, nowhere to go (gal)", kgal, true]]);
  } else if (state.resultView === "demand") {
    h += matrixView(r, "demand", r.demand_coverage, demRows, [
      ["delivered", "Delivered (gal)", kgal, true], ["demand", "Demand (gal)", kgal, true],
      ["unmet", "Unmet demand (gal)", kgal, true], ["excess", "Forced in beyond use (gal)", kgal, true],
      ["end", "Ending inventory (gal)", kgal, false], ["withdraw", "Storage withdrawal (gal)", kgal, true],
      ["inject", "Storage injection (gal)", kgal, true]]);
  } else if (state.resultView === "sinv") {
    h += matrixView(r, "sinv", r.supply_inventory, supRows, [
      ["end", "Ending inventory (gal)", kgal, false], ["withdraw", "Withdrawal (gal)", kgal, true],
      ["inject", "Injection (gal)", kgal, true], ["cap", "Storage capacity (gal)", kgal, false],
      ["min", "Minimum inventory (gal)", kgal, false]]);
  } else if (state.resultView === "ship") {
    h += shipmentsView(r, supRows, demRows);
  } else if (state.resultView === "lanes") {
    h += laneVolumesView(r, supRows, demRows);
  } else {
    h += `<div class="card-body" style="padding-top:16px">` +
         r.diagnostics.map(m => `<div class="msg ${m.severity}"><strong>${m.severity}</strong>${esc(m.text)}</div>`).join("") + `</div>`;
  }
  return h + `</div>`;
}

// Monthly cost, stacked: commodity (blue) + freight (orange). The toggle in the
// header switches between total dollars and cost per gallon: commodity $ per
// gallon bought that month plus freight $ per gallon shipped, so the two parts
// add up to the delivered cost per gallon.
function costChart(r) {
  const P = r.periods.length, perGalMode = state.chartMode === "gal";
  const sum = a => a.reduce((x, v) => x + (v || 0), 0);
  const comm$ = r.periods.map((_, p) => sum(r.supply_plan.cost[p]));
  const bought = r.periods.map((_, p) => sum(r.supply_plan.purchased[p]));    // kgal
  const frt$ = new Array(P).fill(0), shipped = new Array(P).fill(0);
  for (const s of r.shipments) { frt$[s[0]] += s[5] || 0; shipped[s[0]] += s[3] || 0; }
  // Per gallon: $/kgal internally (perGal() shows it as $/gal); 0 when nothing moved.
  const parts = state.chartParts;                       // both | propane | freight
  const showC = parts !== "freight", showF = parts !== "propane";
  const commAll = perGalMode ? comm$.map((c, p) => bought[p] > 1e-9 ? c / bought[p] : 0) : comm$;
  const frtAll = perGalMode ? frt$.map((f, p) => shipped[p] > 1e-9 ? f / shipped[p] : 0) : frt$;
  const comm = showC ? commAll : commAll.map(() => 0), frt = showF ? frtAll : frtAll.map(() => 0);
  const tot = comm.map((c, p) => c + frt[p]);
  const fmtAxis = perGalMode ? v => "$" + (v / GAL).toFixed(2) : moneyShort;
  const fmtVal = perGalMode ? v => perGal(v) + "/gal" : money;
  const max = Math.max(perGalMode ? 1e-6 : 1, ...tot);
  const W = 760, H = 190, L = 56, B = 24, T = 10, bw = (W - L - 8) / P;
  let g = "";
  for (let k = 0; k <= 4; k++) {
    const y = T + (H - T - B) * (1 - k / 4);
    g += `<line class="gridline" x1="${L}" x2="${W - 4}" y1="${y}" y2="${y}"/><text x="${L - 8}" y="${y + 4}" text-anchor="end">${fmtAxis(max * k / 4)}</text>`;
  }
  r.periods.forEach((lab, p) => {
    const x = L + p * bw + bw * 0.18, w = bw * 0.64, sc = (H - T - B) / max;
    const hc = comm[p] * sc, hf = frt[p] * sc, y0 = H - B;
    const tip = perGalMode
      ? `${lab}: ${fmtVal(tot[p])} delivered (commodity ${fmtVal(comm[p])}, freight ${fmtVal(frt[p])}; ${kgal(bought[p])} gal bought)`
      : `${lab}: ${money(tot[p])} (commodity ${money(comm[p])}, freight ${money(frt[p])})`;
    g += `<g><title>${esc(tip)}</title>
      <rect class="bar-c" x="${x}" y="${y0 - hc}" width="${w}" height="${Math.max(0, hc)}" rx="2"/>
      <rect class="bar-f" x="${x}" y="${y0 - hc - hf}" width="${w}" height="${Math.max(0, hf)}" rx="2"/></g>
      <text x="${x + w / 2}" y="${H - 6}" text-anchor="middle">${esc(lab)}</text>`;
  });
  const peak = tot.indexOf(Math.max(...tot));
  const what = { both: "Cost", propane: "Propane cost", freight: "Freight cost" }[parts];
  const toggleParts = `<div class="seg" role="group" aria-label="What to show">
      <button type="button" data-parts="both" aria-pressed="${parts === "both"}">Both</button>
      <button type="button" data-parts="propane" aria-pressed="${parts === "propane"}"><i class="key c"></i>Propane only</button>
      <button type="button" data-parts="freight" aria-pressed="${parts === "freight"}"><i class="key f"></i>Freight only</button></div>`;
  const toggle = `<div class="seg" role="group" aria-label="Chart units">
      <button type="button" data-chart="total" aria-pressed="${!perGalMode}">Total $</button>
      <button type="button" data-chart="gal" aria-pressed="${perGalMode}">Per gallon</button></div>`;
  return `<div class="card chart-card chart">
    <div class="chart-head"><h3>${perGalMode ? `${what} per gallon by month` : `${what} by month`}</h3>
      <div class="chart-legend">${showC ? `<span><i style="background:var(--blue)"></i>Propane (commodity)</span>` : ""}${showF ? `<span><i style="background:#f59e42"></i>Freight</span>` : ""}
      <span class="muted">Peak: ${esc(r.periods[peak])}, ${perGalMode ? fmtVal(tot[peak]) : moneyShort(tot[peak])}</span>${toggleParts}${toggle}</div></div>
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${perGalMode ? `Monthly ${what.toLowerCase()} per gallon` : `Monthly ${what.toLowerCase()}`}">${g}</svg></div>`;
}

function matrixView(r, key, data, rows, metrics) {
  const pick = state["metric_" + key] || metrics[0][0];
  const m = metrics.find(x => x[0] === pick) || metrics[0];
  const [f, label, fmt, summable] = m;
  const vals = data[f];
  const shown = rows.filter(row => !(row.ph && r.periods.every((_, p) => !vals[p][row.i])));
  let h = `<div class="view-tools"><label class="field" for="metric-${key}">Show
           <select id="metric-${key}" data-metric="${key}">` +
          metrics.map(x => `<option value="${x[0]}"${x[0] === m[0] ? " selected" : ""}>${esc(x[1])}</option>`).join("") + `</select></label></div>`;
  h += `<div class="grid-wrap" style="max-height:calc(100vh - 420px)"><table class="grid" aria-label="${esc(label)}"><thead><tr><th class="rowhead" scope="col">Location</th>` +
       r.periods.map(p => `<th scope="col">${esc(p)}</th>`).join("") + (summable ? `<th scope="col">Season</th>` : "") + `</tr></thead><tbody>`;
  const colT = new Array(r.periods.length).fill(0);
  shown.forEach(row => {
    let t = 0;
    h += `<tr><th class="rowhead" scope="row" title="${esc(row.full)}"><span class="loc">${esc(row.main)}</span>${row.meta ? `<span class="meta">${esc(row.meta)}</span>` : ""}</th>`;
    r.periods.forEach((_, p) => { const v = vals[p][row.i] || 0; t += v; colT[p] += v; h += `<td class="num${v ? "" : " zero"}">${fmt(v)}</td>`; });
    if (summable) h += `<td class="num totalcol">${fmt(t)}</td>`;
    h += `</tr>`;
  });
  if (summable) h += `<tr class="total"><th class="rowhead" scope="row">Total</th>${colT.map(v => `<td class="num">${fmt(v)}</td>`).join("")}<td class="num">${fmt(colT.reduce((a, b) => a + b, 0))}</td></tr>`;
  return h + `</tbody></table></div>`;
}

// Supply points down, plants across: the exact gallons each lane carries in the
// month picked (or the season). Row = everything that supply point ships; column =
// everything that plant receives. High-freight lanes are marked, not hidden.
function laneVolumesView(r, supRows, demRows) {
  const m = state.laneMonth;                         // month index, or "all"
  const vol = new Map(), rates = new Map(), hf = new Set();
  for (const s of r.shipments) {
    if (m !== "all" && s[0] !== m) continue;
    const k = `${s[1]},${s[2]}`;
    vol.set(k, (vol.get(k) || 0) + s[3]);
    rates.set(k, [...(rates.get(k) || []), [s[0], s[3], s[4], s[5]]]);
    if (s[6]) hf.add(k);
  }
  const used = new Set([...vol.keys()]);
  const rows = supRows.filter(x => state.laneEmpty ? !x.ph : [...used].some(k => k.startsWith(x.i + ",")));
  const cols = demRows.filter(x => state.laneEmpty ? !x.ph : [...used].some(k => k.endsWith("," + x.i)));
  let h = `<div class="view-tools"><label class="field" for="lane-month">Month <select id="lane-month">
      <option value="all"${m === "all" ? " selected" : ""}>Whole season</option>` +
      r.periods.map((p, k) => `<option value="${k}"${k === m ? " selected" : ""}>${esc(p)}</option>`).join("") + `</select></label>
      <label class="toggle"><input type="checkbox" id="lane-empty"${state.laneEmpty ? " checked" : ""}> Show lanes with nothing on them</label>
      <span class="chip">${NF0.format(used.size)} lanes used</span>
      <span class="chip">${kgal([...vol.values()].reduce((a, b) => a + b, 0))} gal</span>
      ${hf.size ? `<span class="chip amber">${hf.size} high-freight lane${hf.size === 1 ? "" : "s"} (&ge; $1.00/gal) - check the rates</span>` : ""}</div>`;
  if (!rows.length) return h + `<div class="empty-state"><div class="big">Nothing moves${m === "all" ? "" : " in " + esc(r.periods[m])}</div></div>`;
  h += `<div class="grid-wrap" style="max-height:calc(100vh - 420px)"><table class="grid lane-vol"><thead><tr>
      <th class="rowhead" scope="col">From supply &rarr; to plant</th>` +
      cols.map(c => `<th scope="col" class="lane" title="${esc(c.full)}">${esc(c.main)}</th>`).join("") +
      `<th scope="col" class="shipcol">Shipped, gal</th></tr></thead><tbody>`;
  const colT = new Array(cols.length).fill(0);
  for (const row of rows) {
    let t = 0;
    h += `<tr><th class="rowhead" scope="row" title="${esc(row.full)}"><span class="loc">${esc(row.main)}</span>${row.meta ? `<span class="meta">${esc(row.meta)}</span>` : ""}</th>`;
    cols.forEach((c, k) => {
      const key = `${row.i},${c.i}`, v = vol.get(key) || 0;
      t += v; colT[k] += v;
      const tip = v ? (rates.get(key) || []).map(([p, q, rt, cost]) =>
        `${r.periods[p]}: ${galExact(q)} gal at ${perGal(rt)}/gal = ${money(cost)}`).join("\n") : "";
      h += `<td class="num${v ? "" : " zero"}${hf.has(key) ? " high" : ""}"${tip ? ` title="${esc(`${row.full} \u2192 ${c.full}\n${tip}`)}"` : ""}>${v ? kgal(v) : "\u2013"}</td>`;
    });
    h += `<td class="num totalcol shipcol">${kgal(t)}</td></tr>`;
  }
  h += `<tr class="total"><th class="rowhead" scope="row">Received, gal</th>${colT.map(v => `<td class="num">${kgal(v)}</td>`).join("")}
      <td class="num shipcol">${kgal(colT.reduce((a, b) => a + b, 0))}</td></tr></tbody></table></div>
    <p class="muted small" style="padding:6px 18px 16px">Exact gallons, rounded to the gallon; hover a cell for the month-by-month detail
      to the hundredth of a gallon, the rate and the freight cost.</p>`;
  return h;
}

function shipmentsView(r, supRows, demRows) {
  const month = state.shipMonth ?? "all", q = (state.shipFilter || "").toLowerCase();
  let rows = r.shipments.filter(s => month === "all" || String(s[0]) === month);
  if (q) rows = rows.filter(s => (supRows[s[1]].full + " " + demRows[s[2]].main + " " + demRows[s[2]].full).toLowerCase().includes(q));
  rows.sort((a, b) => b[5] - a[5]);
  const vol = rows.reduce((a, s) => a + s[3], 0), cost = rows.reduce((a, s) => a + s[5], 0);
  const LIMIT = 1500;
  let h = `<div class="view-tools"><label class="field" for="ship-month">Month <select id="ship-month"><option value="all">All months</option>` +
          r.periods.map((p, k) => `<option value="${k}"${String(k) === month ? " selected" : ""}>${esc(p)}</option>`).join("") + `</select></label>
          <label class="field" for="ship-filter">Filter <input id="ship-filter" type="search" value="${esc(state.shipFilter || "")}" placeholder="Location name"></label>
          <span class="chip">${rows.length} shipments</span><span class="chip">${kgal(vol)} gal</span><span class="chip">${money(cost)} freight</span></div>`;
  h += `<div class="scroll" style="margin:0 18px 18px;border:1px solid var(--line);border-radius:10px"><table class="list"><thead><tr><th>Month</th><th>From</th><th>To</th>
        <th class="num">Volume gal</th><th class="num">Rate $/gal</th><th class="num">Freight</th><th></th></tr></thead><tbody>`;
  for (const s of rows.slice(0, LIMIT)) {
    const sup = supRows[s[1]], dem = demRows[s[2]];
    h += `<tr><td>${esc(r.periods[s[0]])}</td><td title="${esc(sup.full)}">${esc(sup.main)} <span class="muted">${esc(sup.meta)}</span></td>
          <td title="${esc(dem.full)}">${esc(dem.main)}</td>
          <td class="num">${kgal(s[3])}</td><td class="num">${perGal(s[4])}</td><td class="num">${money(s[5])}</td>
          <td>${s[6] ? '<span class="chip amber" title="Freight of $1.00/gal or more - used because it was the cheapest route; check the rate">high freight</span>' : ""}</td></tr>`;
  }
  h += `</tbody></table></div>`;
  if (rows.length > LIMIT) h += `<p class="muted" style="padding:0 18px 16px">Showing the ${LIMIT} largest by freight. Filter, or download the Excel for the rest.</p>`;
  return h;
}

/* --------------------------------------------------------------- changes */

function targetText(d, field, index) {
  const axes = AXES[field] || [], w = {};
  axes.forEach((a, k) => w[a] = index[k]);
  const parts = [];
  if ("i" in w && "j" in w) parts.push(`${supplyLabel(w.i, d)} \u2192 ${prettyLoc(d.demand[w.j]).main}`);
  else if ("i" in w) parts.push(supplyLabel(w.i, d));
  else if ("j" in w) parts.push(prettyLoc(d.demand[w.j]).main);
  if ("p" in w) parts.push(d.periods[w.p]);
  return parts.join(" · ");
}
function whenText(ts) {
  if (!ts) return "";
  const d = new Date(ts);
  return isNaN(d) ? ts : d.toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
}
function whoCell(user) { return `<span class="who"><span class="avatar">${esc(initials(user))}</span>${esc(user || "Unknown")}</span>`; }

async function renderChanges() {
  const box = $("#changes");
  if (!state.data) return;
  const d = state.data;
  if ($("#changes-all").checked) {
    try {
      const rows = await api("GET", `/api/scenarios/${state.scenario}/changes`);
      box.innerHTML = !rows.length ? `<div class="empty-state"><div class="big">No changes yet</div>Everything in this working model still matches the base.</div>` :
        `<div class="scroll"><table class="list"><thead><tr><th>When</th><th>Who</th><th>Area</th><th>Input</th><th>Where</th><th class="num">New value</th></tr></thead><tbody>` +
        rows.slice(0, CHANGE_ROWS).map(x => {
          const a = areaInfo(x.region, x.area);
          const here = x.region === state.region && x.area === state.area;
          if (x.field === "state_pct") return `<tr><td>${esc(whenText(x.ts))}</td><td>${whoCell(x.user)}</td><td>${esc(a ? a.display.title : x.area)}</td>
            <td>Demand % by state</td><td>${esc(x.states.join(", "))}</td><td class="num"><span class="muted">see Demand % tab</span></td></tr>`;
          return `<tr><td>${esc(whenText(x.ts))}</td><td>${whoCell(x.user)}</td><td>${esc(a ? a.display.title : x.area)}</td>
            <td>${esc(NAV_LABEL[x.field] || x.label)}</td><td>${here ? esc(targetText(d, x.field, x.index)) : '<span class="muted">open the area to see</span>'}</td>
            <td class="num">${esc(editText(x.field, x.value))} <span class="muted">${esc(UNIT(x.field))}</span></td></tr>`;
        }).join("") + `</tbody></table></div>`;
    } catch (err) { fail(err); }
    return;
  }
  const rows = [...d.edits].sort((a, b) => (b.ts || "").localeCompare(a.ts || ""));
  const pctNote = d.state_pct.changed.length ? `<div class="msg Info" style="margin:12px 18px 0">Demand % changed for
    ${esc(d.state_pct.changed.join(", "))} - see the <strong>Demand % by state</strong> tab.</div>` : "";
  if (!rows.length) {
    box.innerHTML = pctNote + `<div class="empty-state"><div class="big">No cell changes to ${esc(d.display.title)}</div>Every input here still matches the base.</div>`;
    return;
  }
  box.innerHTML = pctNote + `<div class="view-tools">${d.scenario.can_edit ? `<button type="button" id="revert-all" class="btn danger sm">Revert all ${rows.length} in this area</button>` : ""}</div>
    <div class="scroll"><table class="list"><thead><tr><th>When</th><th>Who</th><th>Input</th><th>Where</th>
    <th class="num">Change</th><th></th></tr></thead><tbody>` +
    rows.slice(0, CHANGE_ROWS).map((e, k) => `<tr><td>${esc(whenText(e.ts))}${e.via ? `<br><span class="chip cyan">pasted</span>` : ""}</td><td>${whoCell(e.user)}</td><td>${esc(NAV_LABEL[e.field] || d.field_info[e.field].label)}</td>
      <td>${esc(targetText(d, e.field, e.index))}</td>
      <td class="num"><span class="muted">${esc(editText(e.field, e.base))}</span><span class="arrow">\u2192</span><strong>${esc(editText(e.field, e.value))}</strong> <span class="muted">${esc(UNIT(e.field))}</span></td>
      <td>${d.scenario.can_edit ? `<button type="button" class="btn sm" data-revert="${k}">Revert</button>` : ""}</td></tr>`).join("") + `</tbody></table></div>` +
    (rows.length > CHANGE_ROWS ? `<p class="muted small" style="padding:8px 18px 16px">Showing the newest ${CHANGE_ROWS} of ${NF0.format(rows.length)}. "Revert all" covers every one.</p>` : "");
  const doRevert = async (list) => {
    const user = $("#user").value.trim();
    if (!user) { toast("Add your name at the top right first.", "error"); $("#user").focus(); return; }
    try {
      const data = await busy("Reverting...", () => api("POST", "/api/edits", {
        scenario: state.scenario, region: state.region, area: state.area, user,
        changes: list.map(e => ({ field: e.field, index: e.index, revert: true })),
      }));
      state.data = data; state.version = data.scenario.version;
      renderAreaHeader(); renderInputs(); renderChanges(); await loadScenarios(state.scenario);
      toast(`Reverted ${list.length} change${list.length === 1 ? "" : "s"} to the base value.`, "ok");
    } catch (err) { fail(err); }
  };
  $$("#changes button[data-revert]").forEach(b => b.onclick = () => doRevert([rows[+b.dataset.revert]]));
  if ($("#revert-all")) $("#revert-all").onclick = () => { if (confirm(`Put all ${rows.length} value(s) in ${d.display.title} back to the base?`)) doRevert(rows); };
}

/* ------------------------------------------------------ demand % by state */
// The base demand is budget demand x the "State Adjustments" %. Changing a
// state's % scales every plant in that state by new % / base %.

const pctText = v => v == null ? "" : String(Math.round(v * 1000) / 10);

/* ----------------------------------------------------- seasons + month picker */
// One set of season buttons used everywhere months are picked, so nobody has
// to remember which months "winter" is. Winter = Oct-Mar and summer = Apr-Sep,
// the same split the Demand for Models workbook uses.

const MON3 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const SEASONS = [
  { key: "all", label: "All year", months: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] },
  { key: "winter", label: "Winter", sub: "Oct\u2013Mar", months: [10, 11, 12, 1, 2, 3] },
  { key: "peak", label: "Peak winter", sub: "Dec\u2013Feb", months: [12, 1, 2] },
  { key: "shoulder", label: "Shoulder", sub: "Oct, Nov, Mar", months: [10, 11, 3] },
  { key: "summer", label: "Summer", sub: "Apr\u2013Sep", months: [4, 5, 6, 7, 8, 9] },
];
// Calendar month (1-12) of a period label like "Apr-26"; unreadable labels assume Apr..Mar order.
function monthOf(label, p) {
  const k = MON3.indexOf(String(label).slice(0, 3));
  return k >= 0 ? k + 1 : (p + 3) % 12 + 1;
}
const isWinter = (label, p) => SEASONS[1].months.includes(monthOf(label, p));
function seasonPeriods(key, periods) {
  const s = SEASONS.find(x => x.key === key);
  return periods.map((l, p) => s.months.includes(monthOf(l, p)) ? p : -1).filter(p => p >= 0);
}
// Which season button matches the current pick exactly (to show it pressed).
function seasonOf(sel, periods) {
  for (const s of SEASONS) {
    const ps = seasonPeriods(s.key, periods);
    if (ps.length === sel.size && ps.every(p => sel.has(p))) return s.key;
  }
  return "";
}

// Season buttons + one chip per month (winter months blue, summer months gold).
function monthPickerHtml(id, periods, sel) {
  const on = seasonOf(sel, periods);
  return `<div class="mp" id="${id}">
    <div class="mp-seasons" role="group" aria-label="Pick a season">${SEASONS.map(s =>
      `<button type="button" class="mp-season s-${s.key}" data-season="${s.key}" aria-pressed="${on === s.key}">${s.label}${s.sub ? `<span>${s.sub}</span>` : ""}</button>`).join("")}
      <button type="button" class="mp-season s-none" data-season="none" aria-pressed="${sel.size === 0}">Clear</button></div>
    <div class="mp-months" role="group" aria-label="Or pick months">${periods.map((l, p) =>
      `<label class="mp-m ${isWinter(l, p) ? "w" : "s"}"><input type="checkbox" data-p="${p}"${sel.has(p) ? " checked" : ""}>${esc(l)}</label>`).join("")}</div></div>`;
}
function bindMonthPicker(id, periods, sel, onChange) {
  const root = $("#" + id);
  if (!root) return;
  root.querySelectorAll("[data-season]").forEach(b => b.onclick = () => {
    sel.clear();
    if (b.dataset.season !== "none") seasonPeriods(b.dataset.season, periods).forEach(p => sel.add(p));
    onChange();
  });
  root.querySelectorAll("[data-p]").forEach(c => c.onchange = () => {
    c.checked ? sel.add(+c.dataset.p) : sel.delete(+c.dataset.p);
    onChange();
  });
}
// "Dec-26, Jan-27, Feb-27" or "Peak winter (Dec-Feb)".
function monthsText(sel, periods) {
  const key = seasonOf(sel, periods);
  if (key) { const s = SEASONS.find(x => x.key === key); return s.sub ? `${s.label.toLowerCase()} (${s.sub})` : "all year"; }
  return [...sel].sort((a, b) => a - b).map(p => periods[p]).join(", ");
}

/* ------------------------------------------------------- quick demand change */

const QD_PRESETS = [
  { label: "Cold winter", sub: "Dec\u2013Feb +10%", season: "peak", pct: 10 },
  { label: "Very cold winter", sub: "Oct\u2013Mar +20%", season: "winter", pct: 20 },
  { label: "Mild winter", sub: "Dec\u2013Feb \u221210%", season: "peak", pct: -10 },
  { label: "Slow summer", sub: "Apr\u2013Sep \u221210%", season: "summer", pct: -10 },
];

function qdState() {
  const d = state.data;
  if (!state.qd || state.qd.area !== state.area) {
    state.qd = { area: state.area, months: new Set(seasonPeriods("peak", d.periods)), where: "all",
                 states: new Set(), plants: new Set(), mode: "change", pct: 10, setPct: 95, filter: "" };
  }
  return state.qd;
}

// Every plant the change reaches, and for each picked month its demand now and after (kgal).
function qdPlan() {
  const d = state.data, q = qdState(), sp = d.state_pct, st = d.demand_states || [];
  const plants = d.demand.map((n, j) => j).filter(j => !isPlaceholder(d.demand[j]) &&
    (q.where === "all" || (q.where === "states" ? q.states.has(st[j]) : q.plants.has(j))));
  const months = [...q.months].sort((a, b) => a - b);
  const cur = d.fields.demand;
  const pctNow = (s, p) => (state.pctDraft && state.pctDraft[s] ? state.pctDraft[s] : sp.current[s] || [])[(monthOf(d.periods[p], p) + 8) % 12];
  const cells = [];
  let before = 0, after = 0, blocked = "";
  for (const j of plants) for (const p of months) {
    const v = cur[p][j] || 0;
    let nv;
    if (q.mode === "change") nv = v * (1 + q.pct / 100);
    else {
      const now = pctNow(st[j], p);
      if (!st[j] || !now) { blocked = "Setting a % of budget works by state, and some of these plants have no state - use Raise or lower instead."; nv = v; }
      else nv = v * (q.setPct / 100) / now;
    }
    before += v; after += nv;
    cells.push({ p, j, v, nv });
  }
  // Each cell goes the right way by itself: through its state's % when the plant has
  // a state and wasn't typed in by hand (whole area / states only), otherwise as a
  // typed-in value (picked plants, plants with no state, cells already typed in).
  const editedDemand = new Set(d.edits.filter(e => e.field === "demand").map(e => e.index.join(",")));
  const viaPct = c => q.where !== "plants" && st[c.j] && sp.states.includes(st[c.j]) && !editedDemand.has(`${c.p},${c.j}`);
  const pctCells = cells.filter(viaPct), editCells = cells.filter(c => !viaPct(c));
  const statesInScope = [...new Set(pctCells.map(c => st[c.j]))];
  const byState = pctCells.length > 0;
  return { plants, months, cells, pctCells, editCells, before, after, byState, statesInScope, blocked,
           editPlants: new Set(editCells.map(c => c.j)).size };
}

function renderQuickDemand() {
  const d = state.data, q = qdState(), sp = d.state_pct, box = $("#quick-demand");
  const ro = !d.scenario.can_edit;
  const plan = qdPlan();
  const stGal = s => (sp.current_gal[s] || []).reduce((a, v) => a + v, 0);
  const plantsList = d.demand.map((n, j) => [j, n]).filter(([, n]) => !isPlaceholder(n));
  const flt = q.filter.toLowerCase();
  let h = `<div class="qd-presets"><span class="qd-h">Common changes</span>${QD_PRESETS.map((x, k) =>
      `<button type="button" class="qd-preset" data-preset="${k}">${x.label}<span>${x.sub}</span></button>`).join("")}</div>
    <div class="qd-grid">
      <div class="qd-step"><h3><span class="sn">1</span>When?</h3>${monthPickerHtml("qd-months", d.periods, q.months)}</div>
      <div class="qd-step"><h3><span class="sn">2</span>Where?</h3>
        <div class="qd-where" role="radiogroup">
          <label><input type="radio" name="qd-where" value="all"${q.where === "all" ? " checked" : ""}> Every plant in ${esc(d.display.title)}</label>
          ${sp.states.length ? `<label><input type="radio" name="qd-where" value="states"${q.where === "states" ? " checked" : ""}> Only some states</label>` : ""}
          <label><input type="radio" name="qd-where" value="plants"${q.where === "plants" ? " checked" : ""}> Only some plants</label></div>
        ${q.where === "states" ? `<div class="qd-chips">${sp.states.map(s => `<label class="qd-chip"><input type="checkbox" data-state="${esc(s)}"${q.states.has(s) ? " checked" : ""}>
            <strong>${esc(s)}</strong> <span>${kgal(stGal(s))} gal</span></label>`).join("")}</div>` : ""}
        ${q.where === "plants" ? `<input type="search" id="qd-filter" class="qd-filter" placeholder="Find a plant" value="${esc(q.filter)}">
          <div class="qd-chips qd-plants">${plantsList.filter(([, n]) => !flt || n.toLowerCase().includes(flt) || prettyLoc(n).main.toLowerCase().includes(flt)).map(([j, n]) =>
            `<label class="qd-chip"><input type="checkbox" data-plant="${j}"${q.plants.has(j) ? " checked" : ""}>${esc(prettyLoc(n).main)}</label>`).join("")}</div>
          <div class="qd-mini"><button type="button" class="linkish" id="qd-all-plants">Select all shown</button>
            <button type="button" class="linkish" id="qd-no-plants">Clear</button></div>` : ""}
      </div>
      <div class="qd-step"><h3><span class="sn">3</span>How much?</h3>
        <div class="seg qd-mode" role="group">
          <button type="button" data-mode="change" aria-pressed="${q.mode === "change"}">Raise or lower demand</button>
          ${sp.states.length ? `<button type="button" data-mode="set" aria-pressed="${q.mode === "set"}">Contract at a % of budget</button>` : ""}</div>
        ${q.mode === "change" ? `<div class="qd-steps">${[-20, -10, -5, 5, 10, 20].map(v =>
            `<button type="button" class="qd-step-btn ${v < 0 ? "down" : "up"}" data-pct="${v}" aria-pressed="${q.pct === v}">${v > 0 ? "+" : "\u2212"}${Math.abs(v)}%</button>`).join("")}</div>
          <label class="field">or exactly <input type="number" id="qd-pct" value="${q.pct}" step="1" min="-90" max="200"> %</label>`
        : `<label class="field">Contract for <input type="number" id="qd-set" value="${q.setPct}" step="1" min="0" max="300"> % of budget demand</label>
           <p class="hint">Like the State Adjustments sheet: 95 means plan for 95% of the budget.</p>`}
      </div>
    </div>`;
  // the plain-language preview
  const diff = plan.after - plan.before;
  const where = q.where === "all" ? `every plant in ${d.display.title}` : q.where === "states"
    ? (q.states.size ? `plants in ${[...q.states].join(", ")}` : "") : (q.plants.size ? `${q.plants.size} plant${q.plants.size === 1 ? "" : "s"}` : "");
  const how = q.mode === "change" ? `${q.pct >= 0 ? "up" : "down"} ${Math.abs(q.pct)}%` : `contracted at ${q.setPct}% of budget`;
  const ready = plan.cells.length && where && !plan.blocked && Math.abs(diff) > 0.0005;
  h += `<div class="qd-preview ${ready ? (diff > 0 ? "up" : "down") : "idle"}">`;
  if (!q.months.size) h += `<p>Pick the months first (step 1).</p>`;
  else if (!where) h += `<p>Pick ${q.where === "states" ? "at least one state" : "at least one plant"} (step 2).</p>`;
  else if (plan.blocked) h += `<p>${esc(plan.blocked)}</p>`;
  else {
    h += `<p class="qd-say">Demand for <strong>${esc(where)}</strong> in <strong>${esc(monthsText(q.months, d.periods))}</strong> goes <strong>${esc(how)}</strong>:</p>
      <div class="qd-nums"><span>${kgal(plan.before)} gal</span><span class="arrow">\u2192</span><strong>${kgal(plan.after)} gal</strong>
        <span class="${diff >= 0 ? "up" : "down"}">(${diff >= 0 ? "+" : "\u2212"}${kgal(Math.abs(diff))} gal)</span></div>
      <div class="qd-bymonth">${plan.months.map(p => {
        const b = plan.cells.filter(c => c.p === p).reduce((a, c) => a + c.v, 0), n = plan.cells.filter(c => c.p === p).reduce((a, c) => a + c.nv, 0);
        return `<span class="${isWinter(d.periods[p], p) ? "w" : "s"}"><b>${esc(d.periods[p])}</b>${kgal(b)} \u2192 ${kgal(n)}</span>`; }).join("")}</div>
      <p class="muted small">${[
          plan.byState ? `Saved as the state percentages for ${esc(plan.statesInScope.join(", "))} (see the table below)` : "",
          plan.editCells.length ? `${plan.byState ? "plus " : "Saved as "}typed-in demand for ${plan.editPlants} plant${plan.editPlants === 1 ? "" : "s"}${q.where !== "plants" ? " that have no state or were typed in already" : ""} (yellow in the Demand table)` : "",
        ].filter(Boolean).join(", ")}.
        It's added to what's there now, so pressing it twice applies it twice.</p>`;
  }
  h += `<div class="qd-actions">${ro ? `<span class="muted small">${esc(d.scenario.why_not)}</span>`
    : `<button type="button" class="btn primary" id="qd-apply"${ready ? "" : " disabled"}>Apply this change</button>`}</div></div>`;
  h += demandMonthChart();
  box.innerHTML = h;

  const again = () => renderQuickDemand();
  bindMonthPicker("qd-months", d.periods, q.months, again);
  $$("#quick-demand [data-preset]").forEach(b => b.onclick = () => {
    const x = QD_PRESETS[+b.dataset.preset];
    q.months.clear(); seasonPeriods(x.season, d.periods).forEach(p => q.months.add(p));
    q.mode = "change"; q.pct = x.pct; again();
  });
  $$('#quick-demand input[name="qd-where"]').forEach(r => r.onchange = () => { q.where = r.value; again(); });
  $$("#quick-demand [data-state]").forEach(c => c.onchange = () => { c.checked ? q.states.add(c.dataset.state) : q.states.delete(c.dataset.state); again(); });
  $$("#quick-demand [data-plant]").forEach(c => c.onchange = () => { c.checked ? q.plants.add(+c.dataset.plant) : q.plants.delete(+c.dataset.plant); again(); });
  if ($("#qd-filter")) $("#qd-filter").oninput = e => { q.filter = e.target.value; again(); $("#qd-filter").focus(); $("#qd-filter").setSelectionRange(q.filter.length, q.filter.length); };
  if ($("#qd-all-plants")) $("#qd-all-plants").onclick = () => { $$("#quick-demand [data-plant]").forEach(c => q.plants.add(+c.dataset.plant)); again(); };
  if ($("#qd-no-plants")) $("#qd-no-plants").onclick = () => { q.plants.clear(); again(); };
  $$("#quick-demand [data-mode]").forEach(b => b.onclick = () => { q.mode = b.dataset.mode; again(); });
  $$("#quick-demand [data-pct]").forEach(b => b.onclick = () => { q.pct = +b.dataset.pct; again(); });
  if ($("#qd-pct")) $("#qd-pct").onchange = e => { const v = +e.target.value; if (isFinite(v) && v > -100) q.pct = v; again(); };
  if ($("#qd-set")) $("#qd-set").onchange = e => { const v = +e.target.value; if (isFinite(v) && v >= 0 && v <= 300) q.setPct = v; again(); };
  if ($("#qd-apply")) $("#qd-apply").onclick = applyQuickDemand;
}

// Demand by month: the base model's (grey) and now (teal); winter months shaded.
function demandMonthChart() {
  const d = state.data, dm = d.demand_months;
  if (!dm) return "";
  const P = d.periods.length, base = dm.base.slice(0, P), now = dm.now.slice(0, P);
  const max = Math.max(1, ...base, ...now);
  const W = 760, H = 170, L = 60, B = 24, T = 10, bw = (W - L - 8) / P;
  let g = "";
  d.periods.forEach((l, p) => { if (isWinter(l, p)) g += `<rect class="winter-band" x="${L + p * bw}" y="${T}" width="${bw}" height="${H - T - B}"/>`; });
  for (let k = 0; k <= 4; k++) { const y = T + (H - T - B) * (1 - k / 4);
    g += `<line class="gridline" x1="${L}" x2="${W - 4}" y1="${y}" y2="${y}"/><text x="${L - 8}" y="${y + 4}" text-anchor="end">${NF0.format(max * k / 4 * GAL / 1000)}k</text>`; }
  d.periods.forEach((l, p) => {
    const x = L + p * bw + bw * 0.14, w = bw * 0.34, sc = (H - T - B) / max;
    g += `<g><title>${esc(l)}: base ${kgal(base[p])} gal, now ${kgal(now[p])} gal</title>
      <rect class="bar-base" x="${x}" y="${H - B - base[p] * sc}" width="${w}" height="${base[p] * sc}" rx="2"/>
      <rect class="bar-now" x="${x + w + 2}" y="${H - B - now[p] * sc}" width="${w}" height="${now[p] * sc}" rx="2"/></g>
      <text x="${x + w}" y="${H - 6}" text-anchor="middle">${esc(l)}</text>`;
  });
  const tb = base.reduce((a, v) => a + v, 0), tn = now.reduce((a, v) => a + v, 0);
  return `<div class="chart qd-chart"><div class="chart-head"><h3>Demand by month</h3>
      <div class="chart-legend"><span><i style="background:#cbd5e1"></i>Base model</span><span><i style="background:var(--t-demand)"></i>This working model</span>
      <span><i class="winter-key"></i>Winter</span>
      <span class="muted">Season: ${kgal(tb)} \u2192 <strong>${kgal(tn)}</strong> gal${Math.abs(tn - tb) > 0.5 ? ` <span class="${tn > tb ? "up" : "down"}">(${tn > tb ? "+" : "\u2212"}${NF1.format(Math.abs(tn / tb - 1) * 100)}%)</span>` : ""}</span></div></div>
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Demand by month, base against this working model">${g}</svg></div>`;
}

async function applyQuickDemand() {
  if (needName()) return;
  const d = state.data, q = qdState(), plan = qdPlan(), sp = d.state_pct, st = d.demand_states;
  try {
    // 1. typed-in values first (plants picked one by one, plants with no state, cells already typed in)
    if (plan.editCells.length) {
      const edits = new Map(d.edits.filter(e => e.field === "demand").map(e => [e.index.join(","), e]));
      const changes = plan.editCells.map(c => {
        const e = edits.get(`${c.p},${c.j}`);
        if (e && Math.abs(c.nv - e.base) <= 1e-9 * Math.max(1, Math.abs(e.base))) return { field: "demand", index: [c.p, c.j], revert: true };
        return { field: "demand", index: [c.p, c.j], value: c.nv };
      }).filter(c => c.revert || Math.abs(c.value - (d.fields.demand[c.index[0]][c.index[1]] || 0)) > 1e-9);
      if (changes.length) {
        const data = await busy("Saving demand...", () => api("POST", "/api/edits", {
          scenario: state.scenario, region: state.region, area: state.area, user: userName(), changes }));
        state.data = data; state.version = data.scenario.version;
      }
    }
    // 2. then the state percentages: only the picked months of the states in scope
    if (plan.byState) {
      const sp2 = state.data.state_pct;
      if (!state.pctDraft) state.pctDraft = Object.fromEntries(sp2.states.map(s => [s, [...(sp2.current[s] || new Array(12).fill(1))]]));
      const slots = [...new Set(plan.months.map(p => (monthOf(d.periods[p], p) + 8) % 12))];
      for (const s of plan.statesInScope) for (const k of slots) {
        const now = state.pctDraft[s][k];
        state.pctDraft[s][k] = q.mode === "change" ? now * (1 + q.pct / 100) : q.setPct / 100;
      }
      await saveDemandPct(false);                       // re-renders and toasts
    } else {
      state.results = null;
      renderAreaHeader(); renderInputs(); renderDemandPct(); renderResults(); renderChanges();
      await loadScenarios(state.scenario);
      toast(`Saved. Season demand is now ${kgal(state.data.totals.demand)} gal - run the model to see the cost.`, "ok");
    }
  } catch (err) { fail(err); }
}

async function undoAllDemand() {
  const d = state.data;
  const demandEdits = d.edits.filter(e => e.field === "demand");
  if (!d.state_pct.changed.length && !demandEdits.length) { toast("Demand already matches the base.", ""); return; }
  if (!confirm(`Put all demand in ${d.display.title} back to the base model? This clears the state percentages`
               + (demandEdits.length ? ` and ${demandEdits.length} typed-in demand value(s)` : "") + ".")) return;
  if (needName()) return;
  try {
    if (demandEdits.length) {
      const data = await api("POST", "/api/edits", { scenario: state.scenario, region: state.region, area: state.area, user: userName(),
        changes: demandEdits.map(e => ({ field: "demand", index: e.index, revert: true })) });
      state.data = data; state.version = data.scenario.version;
    }
    if (state.data.state_pct.changed.length) await saveDemandPct(true);
    else { state.results = null; renderAreaHeader(); renderInputs(); renderDemandPct(); renderResults(); renderChanges(); await loadScenarios(state.scenario); }
    toast("Demand is back to the base model.", "ok");
  } catch (err) { fail(err); }
}

function initDemandPct() {
  $("#pct-save").addEventListener("click", () => saveDemandPct(false));
  $("#pct-reset").addEventListener("click", () => {
    if (confirm("Put every state back to the base model's percentages?")) saveDemandPct(true);
  });
  $("#qd-undo").addEventListener("click", undoAllDemand);
}

function renderDemandPct() {
  const d = state.data, sp = d.state_pct, box = $("#pct-table");
  renderQuickDemand();
  if (!sp.states.length) {
    box.innerHTML = `<div class="empty-state"><div class="big">No states found for this area</div>
      Plants here couldn't be matched to a state, so there is nothing to scale. Use the Demand table on the Inputs tab instead.</div>`;
    return;
  }
  if (!state.pctDraft) state.pctDraft = Object.fromEntries(sp.states.map(st => [st, [...(sp.current[st] || new Array(12).fill(1))]]));
  const draft = state.pctDraft, ro = !d.scenario.can_edit;
  const noBase = sp.states.filter(st => !sp.base[st]);
  const total = arr => (arr || []).reduce((a, v) => a + (v || 0), 0);
  let h = `<p class="muted small" style="margin-bottom:10px">Percent of budget demand, by month. Highlighted cells differ from the base.
      ${noBase.length ? `No base % on file for ${esc(noBase.join(", "))}: 100% there means "base demand as is".` : ""}</p>
    <div class="grid-wrap" style="margin:0"><table class="grid pct-grid"><thead><tr><th class="rowhead">State</th>` +
    sp.months.map((m, k) => `<th class="${k >= 6 ? "mh-w" : "mh-s"}">${m}</th>`).join("") + `<th>Season demand, gal</th></tr>
    <tr class="season-row"><th class="rowhead"></th><th colspan="6" class="mh-s">Summer</th><th colspan="6" class="mh-w">Winter</th><th></th></tr></thead><tbody>`;
  for (const st of sp.states) {
    const base = sp.base[st] || new Array(12).fill(1);
    const g0 = total(sp.base_gal[st]), g1 = total(sp.current_gal[st]);
    // Preview of the new season total from the draft (exact for the base months).
    const preview = (sp.base_gal[st] || []).reduce((a, v, k) => a + v * (draft[st][k] / (base[k] || 1)), 0);
    h += `<tr><th class="rowhead" scope="row"><span class="loc">${esc(st)}</span><span class="meta">${sp.changed.includes(st) ? "changed" : "base"}${ro ? "" :
        ` &middot; set <button type="button" class="linkish" data-fill="${esc(st)}|summer">summer</button> / <button type="button" class="linkish" data-fill="${esc(st)}|winter">winter</button>`}</span></th>` +
      draft[st].map((v, k) => {
        const diff = Math.abs(v - base[k]) > 1e-9;
        return `<td class="${diff ? "edited" : ""}">${ro ? `<span class="ro-pct">${pctText(v)}%</span>`
          : `<input data-st="${esc(st)}" data-k="${k}" inputmode="decimal" value="${pctText(v)}" aria-label="${esc(st)} ${sp.months[k]} %" title="Base ${pctText(base[k])}%">`}</td>`;
      }).join("") +
      `<td class="num">${kgal(preview)}${Math.abs(preview - g0) > 0.5 ? ` <span class="${preview > g0 ? "up" : "down"} small">(${preview > g0 ? "+" : "-"}${kgal(Math.abs(preview - g0))} vs base)</span>` : ""}
        ${Math.abs(g1 - preview) > 0.5 ? `<br><span class="muted small">saved: ${kgal(g1)}</span>` : ""}</td></tr>`;
  }
  box.innerHTML = h + `</tbody></table></div>`;
  $$("#pct-table input").forEach(inp => inp.addEventListener("change", () => {
    const v = Number(inp.value.replace(/[%\s]/g, ""));
    if (!isFinite(v) || v < 0 || v > 300) { toast("Enter a percent between 0 and 300.", "error"); inp.value = pctText(draft[inp.dataset.st][+inp.dataset.k]); return; }
    draft[inp.dataset.st][+inp.dataset.k] = v / 100;
    renderDemandPct();
    $(`#pct-table input[data-st="${inp.dataset.st}"][data-k="${+inp.dataset.k}"]`)?.focus();
  }));
  // Fill a whole row quickly: same % for a season.
  $$("#pct-table [data-fill]").forEach(b => b.onclick = () => {
    const [st, season] = b.dataset.fill.split("|");
    const v = prompt(`${st}, ${season === "winter" ? "winter (Oct\u2013Mar)" : "summer (Apr\u2013Sep)"}: what % of budget demand?`,
                     pctText(draft[st][season === "winter" ? 8 : 0]));
    if (v == null) return;
    const x = Number(String(v).replace(/[%\s]/g, ""));
    if (!isFinite(x) || x < 0 || x > 300) { toast("Enter a percent between 0 and 300.", "error"); return; }
    for (let k = season === "winter" ? 6 : 0; k < (season === "winter" ? 12 : 6); k++) draft[st][k] = x / 100;
    renderDemandPct();
  });
}

async function saveDemandPct(reset) {
  if (needName()) return;
  try {
    const data = await busy("Saving demand percentages...", () => api("POST", "/api/statepct", {
      scenario: state.scenario, region: state.region, area: state.area, user: userName(),
      pct: reset ? {} : state.pctDraft || {},
    }));
    state.data = data; state.version = data.scenario.version; state.pctDraft = null; state.results = null;
    renderAreaHeader(); renderInputs(); renderDemandPct(); renderResults();
    await loadScenarios(state.scenario);
    toast(reset ? "Back to the base percentages." : `Saved. Season demand is now ${kgal(data.totals.demand)} gal - run the model to see the cost.`, "ok");
  } catch (err) { fail(err); }
}

/* --------------------------------------------------------------- insights */
// Read from the solved LP: what one more gallon is worth, and how hard each
// contract is worked. Needs a run first.

function currentResult() {
  if (!state.results) return null;
  return state.results.results.find(x => x.area === state.area) || null;
}

function renderInsights() {
  const box = $("#insights"), r = currentResult();
  if (!r || !r.insights) {
    box.innerHTML = `<div class="card empty-state"><div class="big">Run the model first</div>
      Insights come from the solved model. Press <strong>Run model</strong>, then come back here.
      <div style="margin-top:12px"><button type="button" class="btn run" id="ins-run">Run model</button></div></div>`;
    $("#ins-run").onclick = () => run(true).then(() => showTab("insights"));
    return;
  }
  const ins = r.insights, mv = ins.marginal;
  const supRows = r.supply.map((x, i) => ({ main: x.name, meta: [x.group, x.hub].filter(Boolean).join(" · ") }));
  const dem = j => prettyLoc(r.demand[j]).main;
  const months = e => e.months.length === r.periods.length ? "every month" : e.months.map(p => r.periods[p]).join(", ");
  const save = v => `<span class="down"><strong>${perGal(Math.abs(v))}</strong>/gal</span>`;
  const card = (title, sub, rows, empty) => `<div class="card ins-card"><div class="card-head"><div><h2>${title}</h2><p class="muted">${sub}</p></div></div>
    <div class="card-body">${rows.length ? `<table class="list"><tbody>${rows.join("")}</tbody></table>` : `<p class="muted">${empty}</p>`}</div></div>`;

  const contracts = mv.contracts.slice(0, 10).map(e => `<tr><td><strong>${esc(supRows[e.i].main)}</strong><br><span class="muted small">${esc(supRows[e.i].meta)}</span></td>
    <td>${save(e.value)} saved per extra gallon<br><span class="muted small">at its maximum in ${esc(months(e))}</span></td></tr>`);
  const lanes = mv.lanes.slice(0, 10).map(e => `<tr><td><strong>${esc(supRows[e.i].main)}</strong> \u2192 ${esc(dem(e.j))}
    ${e.closed ? '<span class="chip">closed lane</span>' : ""}</td><td>${save(e.value)} per extra gallon<br><span class="muted small">${esc(months(e))}</span></td></tr>`);
  const storage = mv.storage.slice(0, 10).map(e => `<tr><td><strong>${esc(e.kind === "supply storage" ? supRows[e.i].main : dem(e.j))}</strong>
    <span class="chip">${e.kind === "supply storage" ? "supply storage" : "plant storage"}</span></td><td>${save(e.value)} per extra gallon of tank<br><span class="muted small">full in ${esc(months(e))}</span></td></tr>`);
  const demand = mv.demand.slice(0, 10).map(e => `<tr><td><strong>${esc(dem(e.j))}</strong></td>
    <td>${perGal(e.value)}/gal average<br><span class="muted small">dearest ${perGal(e.best)}/gal${e.penalty ? " - includes unmet-demand penalty" : ""}</span></td></tr>`);

  const use = ins.contract_use.map(u => {
    const s = supRows[u.i];
    const strip = u.months.map((m, p) => `<i class="u-${m || "none"}" title="${esc(r.periods[p])}: ${m === "max" ? "at maximum" : m === "min" ? "at minimum" : m === "mid" ? "in between" : "no contract"}"></i>`).join("");
    const tag = u.at_max >= Math.max(3, u.months.filter(Boolean).length / 2) ? `<span class="chip green">consider more volume</span>`
      : u.at_min === u.months.filter(Boolean).length && u.at_min ? `<span class="chip amber">only taking the minimum</span>` : "";
    return `<tr><td><strong>${esc(s.main)}</strong><br><span class="muted small">${esc(s.meta)}</span></td>
      <td class="num">${kgal(u.bought)} gal<br><span class="muted small">${NF0.format(u.share * 100)}% of contract</span></td>
      <td><span class="use-strip">${strip}</span></td><td>${tag}</td></tr>`;
  });
  box.innerHTML = `<div class="ins-intro card"><div class="card-body" style="padding-top:16px">
      <h2>What the solved model says about ${esc(r.display.title)}</h2>
      <p class="muted">These come from the optimizer's shadow prices: how much the season's cost would drop with one more gallon of
        something that's holding it back. They hold for small changes - use What-if or a working model to test bigger ones.</p></div></div>
    <div class="ins-grid">
      ${card("Contracts worth more volume", "Contracts at their maximum lift, and what one more gallon there would save.", contracts, "No contract is holding the plan back.")}
      ${card("Lanes worth opening up", "Lanes at their shipment limit (or closed), and what one more gallon on them would save.", lanes, "No lane limit is holding the plan back.")}
      ${card("Storage worth adding", "Tanks that are full at month end, and what one more gallon of space would save.", storage, "No storage limit is holding the plan back.")}
      ${card("What one more gallon of demand costs", "The delivered cost of the next gallon at each plant - dearest first.", demand, "")}
    </div>
    <div class="card"><div class="card-head"><div><h2>Contract use</h2>
      <p class="muted">Each contract's season, month by month: <i class="u-max key"></i> at maximum, <i class="u-mid key"></i> in between,
        <i class="u-min key"></i> at minimum.</p></div></div>
      <div class="scroll" style="margin:0 18px 18px"><table class="list"><thead><tr><th>Contract</th><th class="num">Bought</th>
        <th>${r.periods.map(p => esc(p.slice(0, 3))).join(" ")}</th><th></th></tr></thead><tbody>${use.join("")}</tbody></table></div></div>`;
}

/* ---------------------------------------------------------------- what-if */

function initWhatIf() {
  $("#sens-form").addEventListener("submit", async e => {
    e.preventDefault();
    try {
      const r = await busy("Running price sensitivity (several solves)...", () => api("POST", "/api/analysis", {
        kind: "sensitivity", scenario: state.scenario, region: state.region, area: state.area,
        cents: +$("#sens-cents").value, freight_pct: +$("#sens-freight").value }));
      state.whatif = { ...(state.whatif || {}), sens: r }; renderWhatIf();
    } catch (err) { fail(err); }
  });
  $("#stress-form").addEventListener("submit", async e => {
    e.preventDefault();
    const months = [...stressMonths()].sort((a, b) => a - b);
    if (!months.length) { toast("Pick at least one month.", "error"); return; }
    try {
      const r = await busy("Running the stress test...", () => api("POST", "/api/analysis", {
        kind: "stress", scenario: state.scenario, region: state.region, area: state.area, months, pct: +$("#stress-pct").value }));
      state.whatif = { ...(state.whatif || {}), stress: r }; renderWhatIf();
    } catch (err) { fail(err); }
  });
}

const STRESS_PRESETS = [
  { label: "Cold winter", sub: "Dec\u2013Feb +15%", season: "peak", pct: 15 },
  { label: "Very cold winter", sub: "Oct\u2013Mar +25%", season: "winter", pct: 25 },
  { label: "Polar vortex", sub: "Dec\u2013Feb +40%", season: "peak", pct: 40 },
  { label: "Mild winter", sub: "Dec\u2013Feb \u221215%", season: "peak", pct: -15 },
];
function stressMonths() {
  const d = state.data;
  if (!state.stressSel || state.stressSel.area !== state.area) {
    state.stressSel = { area: state.area, set: new Set(seasonPeriods("peak", d.periods)) };
  }
  return state.stressSel.set;
}

function renderWhatIf() {
  const d = state.data;
  if (!d) return;
  // Ready-made scenarios first, then the same season buttons as everywhere else.
  const sel = stressMonths();
  $("#stress-presets").innerHTML = `<span class="qd-h">Try one</span>` + STRESS_PRESETS.map((x, k) =>
    `<button type="button" class="qd-preset" data-stress="${k}">${x.label}<span>${x.sub}</span></button>`).join("");
  $$("#stress-presets [data-stress]").forEach(b => b.onclick = () => {
    const x = STRESS_PRESETS[+b.dataset.stress];
    sel.clear(); seasonPeriods(x.season, d.periods).forEach(p => sel.add(p));
    $("#stress-pct").value = x.pct; renderWhatIf();
    $("#stress-form").requestSubmit();
  });
  $("#stress-months").innerHTML = monthPickerHtml("stress-mp", d.periods, sel) +
    `<p class="small muted" style="margin:6px 0 10px">${sel.size ? `Demand in ${esc(monthsText(sel, d.periods))} goes up by the % below.` : "Pick the months to test."}</p>`;
  bindMonthPicker("stress-mp", d.periods, sel, renderWhatIf);
  const w = state.whatif || {};
  const dlt = (a, b) => b == null || a == null ? "" : `<span class="${a > b ? "up" : "down"}">${a > b ? "+" : "-"}${moneyShort(Math.abs(a - b))}</span>`;
  if (w.sens) {
    const now = w.sens.now.summary;
    $("#sens-out").innerHTML = `<p class="muted small">Now: ${money(now.total)} (${perGal(now.avg)}/gal)</p>
      <table class="list"><thead><tr><th>What moves</th><th>By</th><th class="num">If it goes up</th><th class="num">If it goes down</th></tr></thead><tbody>` +
      w.sens.rows.map(x => `<tr><td><strong>${esc(x.label)}</strong></td><td>${esc(x.change)}</td>
        <td class="num">${x.up.ok ? dlt(x.up.total, now.total) : statusChip(x.up)}</td><td class="num">${x.down.ok ? dlt(x.down.total, now.total) : statusChip(x.down)}</td></tr>`).join("") +
      `</tbody></table><p class="muted small" style="margin-top:8px">The optimizer re-plans each time, so a drop can be smaller than the straight
        price change: it shifts volume to whatever got relatively cheaper.</p>`;
  } else $("#sens-out").innerHTML = "";
  if (w.stress) {
    const a = w.stress.now.summary, b = w.stress.stressed.summary;
    $("#stress-out").innerHTML = `<div class="kpis small-kpis">
        ${kpi("Extra demand", kgal(w.stress.extra_demand), '<div class="d muted">gal</div>')}
        ${kpi("Season cost", b.ok ? money(b.total) : statusChip(b), b.ok ? `<div class="d">${dlt(b.total, a.total)} vs now</div>` : "")}
        ${kpi("Cost per gallon", b.ok ? perGal(b.avg) : "", b.ok && a.avg ? deltaPerGal(b.avg, a.avg).replace("vs base", "vs now") : "")}
        ${kpi("Unmet demand", kgal(b.unmet), '<div class="d muted">gal the supply can\'t cover</div>', b.unmet > 1e-6 ? "bad" : "")}
      </div>
      ${b.unmet > 1e-6 ? `<div class="msg Warning"><strong>Short:</strong> supply can't keep up. Plants hit hardest:
        ${w.stress.unmet_plants.map(u => `${esc(prettyLoc(u.plant).main)} (${kgal(u.unmet)} gal)`).join(", ")}.</div>`
        : `<div class="msg Info">Supply keeps up with the extra demand.</div>`}`;
  } else $("#stress-out").innerHTML = "";
}

/* --------------------------------------------------------- paste from Excel */
// People receive demand, freight, contracts... as Excel sheets in all sorts of
// layouts. Paste the block as it came; the server works out the layout, names
// and units (paste.py) and returns a preview; Apply saves it as ordinary cell
// edits in this working model.

const PASTE_KINDS = [
  { group: "Demand", cls: "g-demand", items: [["demand", "Demand"], ["demand_inv_cap", "Plant storage capacity"],
      ["demand_inv_min", "Plant minimum inventory"], ["demand_inv_init", "Plant opening inventory"]] },
  { group: "Lanes", cls: "g-lanes", items: [["trans_cost", "Freight rates"], ["ship_max", "Maximum shipments"], ["ship_min", "Minimum shipments"]] },
  { group: "Supply", cls: "g-supply", items: [["contract", "Contract volumes"], ["hub_premium", "Differentials"], ["hub_price", "Hub prices"],
      ["perc_min", "Minimum lift %"], ["perc_max", "Maximum lift %"], ["supply_inv_cap", "Supply storage capacity"],
      ["supply_inv_min", "Supply minimum inventory"], ["supply_inv_init", "Supply opening inventory"]] },
];
const PASTE_FIELD = { demand: "demand", demand_inv_cap: "demand_inv_cap", demand_inv_min: "demand_inv_min", demand_inv_init: "demand_inv_init",
  trans_cost: "trans_cost", ship_max: "ship_max", ship_min: "ship_min", contract: "contract", hub_premium: "hub_premium",
  hub_price: "hub_price", perc_min: "perc_min", perc_max: "perc_max", supply_inv_cap: "supply_inv_cap",
  supply_inv_min: "supply_inv_min", supply_inv_init: "supply_inv_init" };
const PLANT_OR_LANE = new Set(["demand", "demand_inv_cap", "demand_inv_min", "demand_inv_init", "trans_cost", "ship_max", "ship_min"]);

function pasteState() {
  if (!state.paste) state.paste = { kind: "demand", scope: "area", options: {}, mapping: { a: {}, b: {} }, res: null, seq: 0 };
  return state.paste;
}

function initPaste() {
  const ta = $("#paste-text");
  let timer = null;
  ta.addEventListener("input", () => {
    const ps = pasteState();
    ps.mapping = { a: {}, b: {} };                    // new data, new name decisions
    $("#paste-size").textContent = ta.value ? `${NF0.format(ta.value.split("\n").filter(Boolean).length)} rows pasted` : "";
    clearTimeout(timer); timer = setTimeout(pastePreview, 350);
  });
  $("#paste-clear").addEventListener("click", () => { ta.value = ""; $("#paste-size").textContent = ""; pasteState().res = null; renderPasteOut(); ta.focus(); });
  $$('input[name="paste-scope"]').forEach(r => r.addEventListener("change", () => { pasteState().scope = r.value; pastePreview(); }));
  $("#paste-template").addEventListener("click", async () => {
    const ps = pasteState();
    try {
      const r = await fetch(`/api/paste/template?scenario=${enc(state.scenario)}&region=${enc(state.region)}&area=${enc(state.area)}&kind=${ps.kind}`);
      if (!r.ok) throw new Error((await r.json().catch(() => ({}))).error || "couldn't make the table");
      const t = await r.text();
      try { await navigator.clipboard.writeText(t); toast("Copied. Paste it into Excel, update the numbers, then copy it back and paste here.", "ok"); }
      catch (_) { toast("Put in the box below - select it all and copy it into Excel.", "ok"); }
      ta.value = t; ta.dispatchEvent(new Event("input"));
    } catch (err) { fail(err); }
  });
}

function renderPasteKinds() {
  const ps = pasteState();
  $("#paste-kinds").innerHTML = PASTE_KINDS.map(g => `<div class="pk-group ${g.cls}"><span class="pk-h"><i></i>${g.group}</span>` +
    g.items.map(([k, l]) => `<button type="button" class="pk ${g.cls}" data-kind="${k}" aria-pressed="${k === ps.kind}">${esc(l)}</button>`).join("") + `</div>`).join("");
  $$("#paste-kinds .pk").forEach(b => b.onclick = () => {
    ps.kind = b.dataset.kind; ps.options = {}; ps.mapping = { a: {}, b: {} };
    if (!PLANT_OR_LANE.has(ps.kind)) { ps.scope = "area"; $('input[name="paste-scope"][value="area"]').checked = true; }
    renderPasteKinds(); pastePreview();
  });
  $('input[name="paste-scope"][value="all"]').disabled = !PLANT_OR_LANE.has(ps.kind);
  renderPasteOut();
}

async function pastePreview() {
  const ps = pasteState(), text = $("#paste-text").value;
  if (!text.trim()) { ps.res = null; renderPasteOut(); return; }
  const seq = ++ps.seq;
  $("#paste-sub").textContent = "Reading...";
  try {
    const r = await api("POST", "/api/paste/preview", { scenario: state.scenario, region: state.region, area: state.area,
      user: userName(), kind: ps.kind, scope: ps.scope, text, options: ps.options, mapping: ps.mapping });
    if (seq !== ps.seq) return;                       // a newer preview is on its way
    ps.res = r; ps.error = "";
  } catch (err) {
    if (seq !== ps.seq) return;
    ps.res = null; ps.error = err.message;
  }
  renderPasteOut();
}

const PASTE_STATUS = {
  matched: ["green", "Matched"], check: ["amber", "Check"], ambiguous: ["purple", "Which one?"],
  unmatched: ["red", "Not found"], "other area": ["", "Other area"],
};

function renderPasteOut() {
  const ps = pasteState(), r = ps.res, box = $("#paste-out");
  if (ps.error) { $("#paste-sub").textContent = "Couldn't read that"; box.innerHTML = `<div class="msg Error">${esc(ps.error)}</div>`; return; }
  if (!r) { $("#paste-sub").textContent = "Paste something to see what it would change."; box.innerHTML = pasteHelp(); return; }
  const fld = PASTE_FIELD[r.kind], d = state.data;
  const counts = {};
  r.names.forEach(n => { const s = n.choice !== null ? (n.choice ? "matched" : "skipped") : n.status; counts[s] = (counts[s] || 0) + 1; });
  $("#paste-sub").textContent = `${r.label}: read ${NF0.format(r.rows_read)} rows as a ${r.layout} table.`;
  const unitOpts = { volume: [["auto", "Work it out"], ["gal", "Gallons"], ["kgal", "Thousand gallons"]],
                     price: [["auto", "Work it out"], ["$/gal", "$ per gallon"], ["$/kgal", "$ per thousand gallons"]],
                     pct: [["auto", "Work it out"], ["percent", "Percent (90 = 90%)"], ["fraction", "Fraction (0.9 = 90%)"]] }[r.unit_kind];
  const opt = ps.options;
  let h = `<div class="paste-summary">
      <div class="ps-big ${r.changes ? "" : "none"}"><span class="v">${NF0.format(r.changes)}</span><span class="k">cells will change${r.reverts ? ` (${NF0.format(r.reverts)} back to the base value)` : ""}</span></div>
      <div class="ps-chips">
        <span class="chip green">${counts.matched || 0} names matched</span>
        ${counts.check ? `<span class="chip amber">${counts.check} to check</span>` : ""}
        ${counts.ambiguous ? `<span class="chip purple">${counts.ambiguous} need a choice</span>` : ""}
        ${counts.unmatched ? `<span class="chip red">${counts.unmatched} not found (skipped)</span>` : ""}
        ${counts.skipped ? `<span class="chip">${counts.skipped} skipped by you</span>` : ""}
        <span class="chip">${NF0.format(r.cells_given)} cells read</span>
      </div></div>`;
  if (["demand", "contract"].includes(r.kind) && r.total_before)
    h += `<p class="small">Season total for the cells pasted: ${kgal(r.total_before)} gal &rarr; <strong>${kgal(r.total_after)} gal</strong>
      <span class="${r.total_after > r.total_before ? "up" : "down"}">(${r.total_after >= r.total_before ? "+" : ""}${NF1.format((r.total_after / r.total_before - 1) * 100)}%)</span></p>`;
  r.notes.forEach(n => { h += `<div class="msg Info">${esc(n)}</div>`; });
  if (r.areas.length > 1 || r.scope === "all") h += `<p class="small muted">Areas: ${r.areas.map(a => `${esc(names_title(a.area))} (${NF0.format(a.cells)})`).join(", ") || "none"}</p>`;

  // settings
  h += `<div class="paste-opts">
    <label class="field">Numbers are in <select data-opt="unit">${unitOpts.map(([v, l]) => `<option value="${v}"${(opt.unit || "auto") === v ? " selected" : ""}>${l}${v === "auto" && (opt.unit || "auto") === "auto" ? ` (${esc(r.unit)})` : ""}</option>`).join("")}</select></label>
    <label class="field">Blank cells <select data-opt="blanks"><option value="skip">leave the model as is</option>
      <option value="clear"${opt.blanks === "clear" ? " selected" : ""}>${fld === "trans_cost" ? "mean lane not used" : fld === "ship_max" ? "mean no limit" : "mean zero"}</option></select></label>`;
  if (r.info.blocks) h += `<label class="field">Use <select data-opt="block">${r.info.blocks.map(b => `<option value="${b.index}"${+(opt.block || 0) === b.index ? " selected" : ""}>${esc(b.label)}</option>`).join("")}</select></label>`;
  if (r.info.value_cols && r.info.value_cols.length > 1) h += `<label class="field">Value column <select data-opt="value_col">${r.info.value_cols.map(v => `<option value="${v.col}"${String(opt.value_col) === String(v.col) ? " selected" : ""}>${esc(v.label)}</option>`).join("")}</select></label>`;
  if (r.info.needs_from && r.scope !== "all") h += `<label class="field">All rows are from <select data-opt="from"><option value="">pick a supply point</option>${r.entities.a.map(e => `<option value="${esc(e.key)}"${opt.from === e.key ? " selected" : ""}>${esc(e.label)}</option>`).join("")}</select></label>`;
  h += `</div>`;
  if (r.info.needs_months && !r.single) {
    ps.monthSel = opt.months && opt.months !== "all" ? new Set(opt.months.map(Number)) : new Set(r.months.map((_, k) => k));
    h += `<div class="paste-months"><p class="small"><strong>Which months?</strong> The paste has one value per ${r.layout.startsWith("matrix") ? "cell" : "row"}, with no month -
      it goes into <strong>${esc(monthsText(ps.monthSel, r.months) || "no months")}</strong>.</p>${monthPickerHtml("paste-mp", r.months, ps.monthSel)}</div>`;
  }

  // names
  const ents = side => (side === "b" ? r.entities.b : r.entities.a);
  const label = (side, key) => (ents(side).find(e => e.key === key) || {}).label || key;
  const problem = r.names.filter(n => n.status !== "matched" || n.choice !== null);
  const fine = r.names.filter(n => n.status === "matched" && n.choice === null);
  const nameRow = n => {
    const [c, t] = PASTE_STATUS[n.status] || ["", n.status];
    const side = n.side, cur = n.choice !== null ? n.choice : n.suggested;
    const list = ents(side).filter(e => !n.area || e.area === n.area);
    const optsHtml = `<option value="">- skip this name -</option>` + (n.options.length ? `<optgroup label="Best guesses">${n.options.map(k =>
      `<option value="${esc(k)}"${k === cur ? " selected" : ""}>${esc(label(side, k))}</option>`).join("")}</optgroup>` : "") +
      `<optgroup label="All">${list.map(e => `<option value="${esc(e.key)}"${e.key === cur && !n.options.includes(e.key) ? " selected" : ""}>${esc(e.label)}${r.scope === "all" ? " - " + esc(names_title(e.area)) : ""}</option>`).join("")}</optgroup>`;
    return `<tr class="st-${c}"><td>${esc(n.text)}${n.area && r.scope === "all" ? `<br><span class="muted small">${esc(names_title(n.area))}</span>` : ""}
      ${r.kind === "trans_cost" || r.kind.startsWith("ship_") ? `<span class="muted small"> (${side === "a" ? "from" : "to"})</span>` : ""}</td>
      <td><span class="chip ${n.choice !== null ? (n.choice ? "green" : "") : c}">${n.choice !== null ? (n.choice ? "Your pick" : "Skipped") : t}</span></td>
      <td><select data-map="${side}" data-key="${esc(n.key)}" aria-label="Match for ${esc(n.text)}">${optsHtml}</select></td></tr>`;
  };
  if (problem.length) h += `<h3 class="ps-h">Names to look at</h3><div class="scroll ps-names"><table class="list"><thead><tr><th>In your paste</th><th></th><th>Goes to</th></tr></thead><tbody>${problem.map(nameRow).join("")}</tbody></table></div>`;
  if (fine.length) h += `<details class="ps-ok"><summary>${fine.length} name${fine.length === 1 ? "" : "s"} matched exactly - show</summary>
      <div class="scroll ps-names"><table class="list"><tbody>${fine.slice(0, 600).map(nameRow).join("")}</tbody></table></div></details>`;
  if (r.missing_count && r.scope !== "all") h += `<details class="ps-ok"><summary>${r.missing_count} in the model but not in your paste (left as they are)</summary>
      <p class="small muted">${r.missing.map(esc).join(" &middot; ")}</p></details>`;

  // biggest changes
  if (r.biggest.length) h += `<h3 class="ps-h">Biggest changes</h3><div class="scroll"><table class="list"><thead><tr><th>Where</th><th class="num">Now</th><th class="num">Pasted</th></tr></thead><tbody>` +
    r.biggest.map(b => `<tr><td>${esc(b.area === state.area ? targetText(d, fld, b.index) : names_title(b.area) + " · " + b.index.join(","))}</td>
      <td class="num">${esc(editText(fld, b.old))}</td><td class="num"><strong>${esc(editText(fld, b.new))}</strong></td></tr>`).join("") + `</tbody></table></div>`;

  h += `<div class="paste-apply">${r.can_edit ? `<button type="button" class="btn primary" id="paste-apply"${r.changes ? "" : " disabled"}>Apply ${NF0.format(r.changes)} change${r.changes === 1 ? "" : "s"} to this working model</button>
      <span class="muted small">They're saved like typed changes: highlighted in the grids, listed under Changes, and revertable.</span>`
    : `<div class="msg Warning">${esc(r.why_not)}</div>`}</div>`;
  box.innerHTML = h;

  $$("#paste-out [data-opt]").forEach(s => s.onchange = () => { ps.options[s.dataset.opt] = s.value; pastePreview(); });
  $$("#paste-out [data-map]").forEach(s => s.onchange = () => { ps.mapping[s.dataset.map][s.dataset.key] = s.value; pastePreview(); });
  if ($("#paste-mp")) bindMonthPicker("paste-mp", r.months, ps.monthSel, () => {
    ps.options.months = ps.monthSel.size === r.months.length ? "all" : [...ps.monthSel].sort((a, b) => a - b);
    pastePreview();
  });
  if ($("#paste-apply")) $("#paste-apply").onclick = pasteApply;
}

function names_title(area) {
  for (const reg of state.regions) { const a = reg.areas.find(x => x.area === area); if (a) return a.display.title; }
  return area;
}

function pasteHelp() {
  return `<div class="paste-help">
    <p><strong>Any of these layouts works</strong> - paste it the way it arrived:</p>
    <ul>
      <li><strong>One row per plant or supply point, months across</strong> (Apr, Apr-26, 4/1/2026, April 2026 or P1..P12). Extra columns like Total are ignored.</li>
      <li><strong>Months down, names across</strong>, or <strong>name | month | value</strong> rows.</li>
      <li><strong>Freight as a grid</strong>: supply points down, plants across (one month at a time - you pick the months).</li>
      <li><strong>Freight as a list</strong>: from | (supplier) | to | rate, or with months across.</li>
      <li><strong>Every area at once</strong> for demand or freight: each row goes to the area its plant is in.</li>
    </ul>
    <p class="muted small">Names are matched by district number, full name, or a close spelling ("Springfield, ST" finds "1001-Springfield ST").
      Gallons or thousand gallons, $/gal or $/thousand gal, 90% or 0.9 are all understood. Nothing is saved until you press Apply.</p></div>`;
}

async function pasteApply() {
  const ps = pasteState(), r = ps.res;
  if (needName()) return;
  if (r.scope === "all" && !confirm(`Apply ${r.changes} changes across ${r.areas.length} areas of this working model?`)) return;
  try {
    const out = await busy(`Saving ${NF0.format(r.changes)} changes...`, () => api("POST", "/api/paste/apply", {
      scenario: state.scenario, region: state.region, area: state.area, user: userName(),
      kind: ps.kind, scope: ps.scope, text: $("#paste-text").value, options: ps.options, mapping: ps.mapping }));
    if (out.data) { state.data = out.data; state.version = out.data.scenario.version; }
    state.results = null; state.pctDraft = null;
    renderAreaHeader(); renderInputs(); renderResults(); renderChanges();
    await loadScenarios(state.scenario);
    const fld = PASTE_FIELD[ps.kind];
    toast(`Applied ${NF0.format(out.applied)} changes${out.areas.length > 1 ? ` in ${out.areas.length} areas` : ""}.`, "ok",
          { label: "See them", run: () => { state.table = fld; showTab("inputs"); renderTableNav(); renderGrid(); } });
    pastePreview();                                   // now shows 0 changes - a handy check
  } catch (err) { fail(err); }
}

/* ------------------------------------------------------------ base models */

function initBases() {
  $("#roll-form").addEventListener("submit", async e => {
    e.preventDefault();
    if (needName()) return;
    try {
      const r = await api("POST", "/api/bases", { action: "roll", user: userName(), source: $("#roll-source").value,
        season: $("#roll-season").value, name: $("#roll-name").value, demand_file: $("#roll-demand").value,
        freight_pct: +$("#roll-freight").value || 0 });
      toast("Making the new base model. This takes about a minute.", "ok");
      watchJob(r.job);
    } catch (err) { fail(err); }
  });
  $("#import-form").addEventListener("submit", async e => {
    e.preventDefault();
    if (needName()) return;
    try {
      const r = await api("POST", "/api/bases", { action: "import", user: userName(), variant: $("#import-variant").value,
        name: $("#import-name").value, demand_file: (state.demandFiles[0] || {}).path || "" });
      toast("Reading the workbooks. This takes about a minute.", "ok");
      watchJob(r.job);
    } catch (err) { fail(err); }
  });
  $("#demand-upload-btn").addEventListener("click", () => $("#demand-file").click());
  $("#demand-file").addEventListener("change", async e => {
    const f = e.target.files[0];
    if (!f) return;
    const fd = new FormData(); fd.append("file", f);
    try {
      const r = await busy(`Reading ${f.name}...`, async () => {
        const res = await fetch("/api/demandfiles", { method: "POST", body: fd, headers: { "X-Requested-With": "propane-model" } });
        const j = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(j.error || `Upload failed (${res.status})`);
        return j;
      });
      await loadBases(); renderBases(); $("#roll-demand").value = r.path;
      toast(`Demand file added: ${r.plants} plants, ${r.states} states.`, "ok");
    } catch (err) { fail(err); }
    e.target.value = "";
  });
}

async function watchJob(jid) {
  for (;;) {
    let j;
    try { j = await api("GET", `/api/bases/jobs/${jid}`); } catch (err) { fail(err); return; }
    const box = $("#jobs");
    if (box) box.innerHTML = `<div class="job ${j.state}"><strong>${esc(j.label)}</strong>
      <div class="bar"><div class="bar-fill" style="width:${j.total ? Math.round(100 * j.done / j.total) : 3}%"></div></div>
      <span class="muted small">${esc(j.state === "error" ? j.error : j.state === "done" ? "Done." : `${j.step} (${j.done} of ${j.total || "?"})`)}</span></div>`;
    if (j.state !== "running") {
      await loadBases();
      if (state.page === "bases" && !state.area) renderBases();
      if (j.state === "done") toast("New base model ready. It's locked - pick it at the top to use it.", "ok");
      else toast("Making the base model failed: " + j.error, "error");
      return;
    }
    await new Promise(r => setTimeout(r, 1500));
  }
}

function renderBases() {
  const list = $("#bases-list");
  list.innerHTML = `<div class="base-cards">` + state.bases.map(b => `
    <article class="base-card${b.official ? " official" : ""}${b.id === state.base ? " current" : ""}">
      <div class="base-top"><span class="lock" aria-hidden="true">\uD83D\uDD12</span>
        <div><h3>${esc(b.name)}</h3><p class="muted small">${b.season ? "20" + esc(b.season) + " \u00b7 " : ""}${b.areas} areas \u00b7 made ${esc(whenText(b.created))}${b.created_by && b.created_by !== "setup" ? " by " + esc(b.created_by) : ""}</p></div>
        ${b.official ? `<span class="chip flame">\u2605 Official</span>` : ""}</div>
      <p class="small base-recipe">${esc(b.recipe && b.recipe.from ? "From " + b.recipe.from : "")}${b.recipe && b.recipe.demand_file ? ` \u00b7 demand: ${esc(b.recipe.demand_file)}` : ""}${b.recipe && b.recipe.freight_pct ? ` \u00b7 freight ${b.recipe.freight_pct > 0 ? "+" : ""}${b.recipe.freight_pct}%` : ""}</p>
      ${b.area_notes.length ? `<details class="small"><summary>${b.area_notes.length} area note(s)</summary><ul>${b.area_notes.map(a =>
        `<li>${esc(a.area)}: ${esc(a.demand_note || `kept last year's demand for ${a.demand_kept.join(", ")}`)}</li>`).join("")}</ul></details>` : ""}
      <div class="base-actions">
        <button type="button" class="btn sm" data-use="${esc(b.id)}">${b.id === state.base ? "In use" : "Use this base"}</button>
        ${b.official ? "" : `<button type="button" class="btn sm ghost" data-official="${esc(b.id)}">Make official</button>`}
      </div>
    </article>`).join("") + `</div>`;
  $$("#bases-list [data-use]").forEach(b => b.onclick = async () => {
    $("#base-pick").value = b.dataset.use; await onBasePick({ target: $("#base-pick") }); goPage("home");
  });
  $$("#bases-list [data-official]").forEach(b => b.onclick = async () => {
    if (!confirm("Make this the official base? New working models start from it by default.")) return;
    try { await api("POST", `/api/bases/${b.dataset.official}/official`, {}); await loadBases(); renderBases(); toast("Official base changed.", "ok"); }
    catch (err) { fail(err); }
  });
  const src = $("#roll-source");
  const cur = src.value;
  src.innerHTML = `<optgroup label="Base models">${state.bases.map(b => `<option value="${esc(BASE_PREFIX + b.id)}">${esc(b.name)}</option>`).join("")}</optgroup>` +
    `<optgroup label="Working models">${state.scenarios.map(s => `<option value="${esc(s.id)}">${esc(s.name)} (${esc(s.owner || "shared")})</option>`).join("")}</optgroup>`;
  src.value = cur || BASE_PREFIX + (state.official || state.base);
  const sb = state.bases.find(b => BASE_PREFIX + b.id === src.value) || state.bases.find(b => b.id === state.base);
  if (!$("#roll-season").value && sb && /^\d{2}-\d{2}$/.test(sb.season || "")) {
    const a = (+sb.season.slice(0, 2) + 1) % 100;
    $("#roll-season").value = `20${String(a).padStart(2, "0")}-${String((a + 1) % 100).padStart(2, "0")}`;
  }
  $("#roll-demand").innerHTML = `<option value="">Keep last year's demand</option>` +
    (state.demandFiles || []).map(f => `<option value="${esc(f.path)}">${esc(f.name)}</option>`).join("");
  if ((state.demandFiles || []).length && !$("#roll-demand").dataset.touched) $("#roll-demand").value = state.demandFiles[0].path;
  $("#roll-demand").onchange = () => { $("#roll-demand").dataset.touched = "1"; };
  const running = (state.jobs || []).find(j => j.state === "running");
  if (running) watchJob(running.id);
}

/* ---------------------------------------------------------------- compare */

function initCompare() {
  $("#compare-form").addEventListener("submit", async e => {
    e.preventDefault();
    const a = $("#cmp-a").value, b = $("#cmp-b").value, region = $("#cmp-region").value;
    if (a === b) { toast("Pick two different models.", "error"); return; }
    try {
      const r = await busy("Running both models for every area in the region...", () =>
        api("GET", `/api/compare?a=${enc(a)}&b=${enc(b)}&region=${enc(region)}`));
      renderCompare(r);
    } catch (err) { fail(err); }
  });
}

function renderCompareForm() {
  const opts = `<optgroup label="Base models">${state.bases.map(b => `<option value="${esc(BASE_PREFIX + b.id)}">\uD83D\uDD12 ${esc(b.name)}</option>`).join("")}</optgroup>` +
    `<optgroup label="My working models">${state.scenarios.filter(s => s.mine).map(s => `<option value="${esc(s.id)}">${esc(s.name)}</option>`).join("")}</optgroup>` +
    `<optgroup label="Other people's">${state.scenarios.filter(s => !s.mine).map(s => `<option value="${esc(s.id)}">${esc(s.name)} (${esc(s.owner || "shared")})</option>`).join("")}</optgroup>`;
  for (const id of ["#cmp-a", "#cmp-b"]) { const v = $(id).value; $(id).innerHTML = opts; if (v) $(id).value = v; }
  if (!$("#cmp-a").value || $("#cmp-a").selectedIndex < 0) $("#cmp-a").value = BASE_PREFIX + state.base;
  if ($("#cmp-b").selectedIndex < 0 || $("#cmp-b").value === $("#cmp-a").value) {
    const mine = state.scenarios.find(s => s.mine && s.base === state.base);
    if (mine) $("#cmp-b").value = mine.id; else if (!isBaseId(state.scenario)) $("#cmp-b").value = state.scenario;
  }
  const rv = $("#cmp-region").value;
  $("#cmp-region").innerHTML = state.regions.map(r => `<option value="${esc(r.region)}">${esc(r.display.title)}</option>`).join("");
  if (rv) $("#cmp-region").value = rv;
}

function renderCompare(r) {
  const out = $("#compare-out");
  let ta = 0, tb = 0;
  const rows = r.areas.map(x => {
    const a = x.a && x.a.summary, b = x.b && x.b.summary;
    if (a && a.ok) ta += a.total; if (b && b.ok) tb += b.total;
    const cell = s => !s ? '<span class="muted">not in this model</span>' : s.ok ? `${money(s.total)}<br><span class="muted small">${perGal(s.avg)}/gal${s.unmet > 1e-6 ? ` \u00b7 <span class="up">${kgal(s.unmet)} gal unmet</span>` : ""}</span>` : statusChip(s);
    const diff = a && b && a.ok && b.ok ? b.total - a.total : null;
    return `<tr><td><strong>${esc(x.display.title)}</strong><br><span class="muted small">${esc(x.area)}</span></td>
      <td class="num">${cell(a)}</td><td class="num">${cell(b)}</td>
      <td class="num">${diff == null ? "" : Math.abs(diff) < 0.5 ? '<span class="muted">same</span>' : `<span class="${diff > 0 ? "up" : "down"}"><strong>${diff > 0 ? "+" : "-"}${money(Math.abs(diff))}</strong></span>`}</td></tr>`;
  });
  // Monthly bars: A vs B, summed over the region (aligned by month position).
  const P = Math.max(0, ...r.areas.map(x => (x.a && x.a.periods || x.b && x.b.periods || []).length));
  const labels = (r.areas.find(x => x.a && x.a.periods) || {}).a?.periods || [];
  const sumM = key => Array.from({ length: P }, (_, p) => r.areas.reduce((t, x) => {
    const m = x[key] && x[key].monthly; return t + (m ? (m.commodity[p] || 0) + (m.freight[p] || 0) : 0); }, 0));
  const A = sumM("a"), Bv = sumM("b"), max = Math.max(1, ...A, ...Bv);
  const W = 760, H = 200, L = 56, Bt = 24, T = 10, bw = (W - L - 8) / Math.max(1, P);
  let g = "";
  for (let k = 0; k <= 4; k++) { const y = T + (H - T - Bt) * (1 - k / 4);
    g += `<line class="gridline" x1="${L}" x2="${W - 4}" y1="${y}" y2="${y}"/><text x="${L - 8}" y="${y + 4}" text-anchor="end">${moneyShort(max * k / 4)}</text>`; }
  for (let p = 0; p < P; p++) {
    const x = L + p * bw + bw * 0.14, w = bw * 0.34, sc = (H - T - Bt) / max;
    g += `<g><title>${esc(labels[p] || "")}: A ${money(A[p])}, B ${money(Bv[p])}</title>
      <rect class="bar-a" x="${x}" y="${H - Bt - A[p] * sc}" width="${w}" height="${A[p] * sc}" rx="2"/>
      <rect class="bar-b" x="${x + w + 2}" y="${H - Bt - Bv[p] * sc}" width="${w}" height="${Bv[p] * sc}" rx="2"/></g>
      <text x="${x + w}" y="${H - 6}" text-anchor="middle">${esc(labels[p] || "P" + (p + 1))}</text>`;
  }
  const diff = tb - ta;
  out.innerHTML = `<div class="kpis">
      ${kpi(`<i class="sw-a"></i>A: ${esc(r.a.name)}`, money(ta), `<div class="d muted">${r.a.is_base ? "base model" : esc(r.a.owner || "")}</div>`)}
      ${kpi(`<i class="sw-b"></i>B: ${esc(r.b.name)}`, money(tb), `<div class="d muted">${r.b.is_base ? "base model" : esc(r.b.owner || "")}</div>`)}
      ${kpi("B minus A", `<span class="${diff > 0 ? "up" : "down"}">${diff > 0 ? "+" : "-"}${money(Math.abs(diff))}</span>`,
            `<div class="d muted">${ta ? NF1.format(diff / ta * 100) + "%" : ""} for ${esc(r.region_display.title)}</div>`, "hero-kpi")}
    </div>
    <div class="card chart-card chart"><div class="chart-head"><h3>Cost by month, whole region</h3>
      <div class="chart-legend"><span><i class="sw-a"></i>A</span><span><i class="sw-b"></i>B</span></div></div>
      <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Monthly cost, model A against model B">${g}</svg></div>
    <div class="card"><div class="scroll"><table class="list"><thead><tr><th>Area</th><th class="num"><i class="sw-a"></i>A</th><th class="num"><i class="sw-b"></i>B</th><th class="num">B minus A</th></tr></thead>
      <tbody>${rows.join("")}</tbody></table></div></div>`;
}

init();
