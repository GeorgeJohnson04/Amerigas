/* uploads.js - bring data in (paste from Excel or upload a workbook),
   check what will change, then save. Also the Excel downloads. */
"use strict";

const U = { kind: null, text: "", scope: "", opts: {}, mapping: {}, upload: null, uploadName: "", sheet: null, sheets: [], last: null };
const GROUPS = [
  { id: "demand", label: "Demand", cls: "c-demand" }, { id: "supply", label: "Supply", cls: "c-supply" },
  { id: "lanes", label: "Lanes & freight", cls: "c-lanes" }, { id: "weather", label: "Weather", cls: "c-weather" }];
const UNIT_CHOICES = { volume: [["auto", "Work it out"], ["gal", "Gallons"], ["kgal", "Thousand gallons"]],
  price: [["auto", "Work it out"], ["$/gal", "$ per gallon"], ["$/kgal", "$ per thousand gallons"]],
  pct: [["auto", "Work it out"], ["percent", "Percent (85 = 85%)"], ["fraction", "Fraction (0.85 = 85%)"]] };

async function renderUploads(kindArg) {
  const main = $("#main");
  if (kindArg && S.boot.kinds[kindArg] && U.kind !== kindArg && !U.upload) { U.kind = kindArg; U.opts = {}; U.mapping = {}; U.last = null; }
  const can = S.info.can_edit;
  main.append(pageHead("Data uploads", `Paste from Excel or upload a workbook. Nothing is saved until you\u2019ve seen what will change. Going into: ${S.info.plan ? `the plan \u201c${S.info.plan.name}\u201d` : "the shared data set"}.`,
    [h("button", { class: "btn", type: "button", onclick: () => API.download("/api/template", { view: S.view }) }, "Download current data (Excel)"),
     h("button", { class: "btn btn-quiet", type: "button", onclick: () => API.download("/api/template", { blank: 1 }) }, "Blank template")], "c-upload"));
  if (!can) main.append(h("div", { class: "alert warn" }, "This is someone else\u2019s plan, so you can preview uploads but not save them. Make your own copy from the banner above."));

  // step 1
  const kinds = S.boot.kinds;
  const pickers = h("div", { class: "kind-groups" }, GROUPS.map(g => h("div", { class: "kind-group " + g.cls }, h("h3", null, g.label),
    S.boot.order.filter(k => kinds[k].group === g.id).map(k => h("button", { type: "button", class: "kind-btn" + (U.kind === k && !U.upload ? " on" : ""), title: kinds[k].help,
      onclick: () => { U.kind = k; U.upload = null; U.sheet = null; U.opts = {}; U.mapping = {}; U.last = null; location.hash = "#/uploads/" + k; route(); } }, kinds[k].label)))));
  const file = h("input", { type: "file", accept: ".xlsx,.xlsm", id: "wbFile", class: "sr-only" });
  const fileBtn = h("label", { for: "wbFile", class: "btn btn-primary" }, "Choose an Excel workbook\u2026");
  file.addEventListener("change", async () => {
    if (!file.files.length) return;
    const fd = new FormData(); fd.append("file", file.files[0]);
    fileBtn.textContent = "Reading\u2026";
    try {
      const r = await API.call("POST", "/api/import/upload", fd);
      U.upload = r.upload; U.uploadName = r.name; U.sheets = r.sheets.map(s => ({ ...s, use: !!s.kind, kind: s.kind || "", opts: {}, mapping: {} }));
      U.kind = null; U.sheet = null; U.last = null; route();
    } catch (e) { toast(e.message, "error"); fileBtn.textContent = "Choose an Excel workbook\u2026"; }
  });
  main.append(card("1. What are you bringing in?", [pickers,
    h("div", { class: "or" }, h("span", null, "or")),
    h("div", { class: "upload-box" }, h("div", null, h("b", null, "Upload a whole workbook"), h("p", { class: "muted small" }, "Every sheet is read and matched to a kind of data (e.g. \u201cDemand Calculations\u201d, \u201cState Adjustments\u201d, or the sheets in the template). You choose which sheets to load.")), file, fileBtn)]));

  if (U.upload) return renderWorkbook(main);
  if (!U.kind) { main.append(h("p", { class: "muted" }, "Pick a kind of data above to start.")); return; }
  renderPaste(main);
}

function optionsBox(kind, opts, onChange, info) {
  const k = S.boot.kinds[kind];
  const box = h("div", { class: "opts" });
  if (k.unit && UNIT_CHOICES[k.unit]) {
    const sel = h("select", { "aria-label": "Units" }, UNIT_CHOICES[k.unit].map(([v, l]) => h("option", { value: v, selected: (opts.unit || "auto") === v }, l)));
    sel.addEventListener("change", () => { opts.unit = sel.value === "auto" ? undefined : sel.value; onChange(); });
    box.append(h("label", { class: "inline-label" }, "Units ", sel));
  }
  if (kind === "weather") {
    const sel = h("select", { "aria-label": "Weather units" }, [["", "Work it out"], ["gal", "Gallons delivered"], ["hdd", "Heating degree days"]].map(([v, l]) => h("option", { value: v, selected: (opts.weather_unit || "") === v }, l)));
    sel.addEventListener("change", () => { opts.weather_unit = sel.value || undefined; onChange(); });
    const rep = h("select", { "aria-label": "Replace or add" }, [["1", "Replace the whole history"], ["0", "Add these rows to it"]].map(([v, l]) => h("option", { value: v, selected: (opts.replace === false ? "0" : "1") === v }, l)));
    rep.addEventListener("change", () => { opts.replace = rep.value === "1"; onChange(); });
    box.append(h("label", { class: "inline-label" }, "Values are ", sel), h("label", { class: "inline-label" }, rep));
  }
  if (["plants", "supply", "lanes"].includes(kind)) {
    const sel = h("select", { "aria-label": "Volume units" }, [["gal", "gallons"], ["kgal", "thousand gallons"]].map(([v, l]) => h("option", { value: v, selected: (opts.vol_unit || "gal") === v }, l)));
    sel.addEventListener("change", () => { opts.vol_unit = sel.value; onChange(); });
    box.append(h("label", { class: "inline-label" }, kind === "lanes" ? "Min / max are in " : "Opening inventory is in ", sel));
  }
  if (["budget", "contract"].includes(kind)) {
    const cb = h("input", { type: "checkbox", checked: !!opts.add_new });
    cb.addEventListener("change", () => { opts.add_new = cb.checked; onChange(); });
    box.append(h("label", { class: "inline-label" }, cb, ` Add ${kind === "budget" ? "plants" : "supply points"} that aren\u2019t there yet`));
  }
  if (k.field || kind === "lanes") {
    const sel = h("select", { "aria-label": "Blank cells" }, [["skip", "leave what\u2019s there"], ["clear", kind === "lanes" ? "mean \u201cnot used\u201d" : "mean zero"]].map(([v, l]) => h("option", { value: v, selected: (opts.blanks || "skip") === v }, l)));
    sel.addEventListener("change", () => { opts.blanks = sel.value; onChange(); });
    box.append(h("label", { class: "inline-label" }, "Blank cells ", sel));
  }
  if (info && info.blocks && info.blocks.length > 1) {
    const sel = h("select", { "aria-label": "Which month columns" }, info.blocks.map(b => h("option", { value: b.index, selected: (+opts.block || 0) === b.index }, b.label)));
    sel.addEventListener("change", () => { opts.block = +sel.value; onChange(); });
    box.append(h("label", { class: "inline-label" }, "Two sets of month columns - use ", sel));
  }
  if (info && (info.single || info.matrix) && kind !== "weather") {
    const mp = monthChips(opts.months || SEASON, ms => { opts.months = ms.length === 12 ? undefined : ms; onChange(); });
    box.append(h("div", { class: "field" }, h("label", null, "One value per row - which months does it apply to?"), mp));
  }
  return box;
}

function renderPaste(main) {
  const kind = U.kind, k = S.boot.kinds[kind];
  const ta = h("textarea", { class: "paste", rows: 10, placeholder: "Copy the cells in Excel - including the header row - and paste them here (Ctrl+V).", "aria-label": "Pasted data" }, U.text);
  const scope = h("select", { "aria-label": "Which area" }, h("option", { value: "" }, "Rows say which area (or match across all areas)"), S.info.areas.map(a => h("option", { value: a.id, selected: U.scope === a.id }, `Only ${a.name}`)));
  const prevBtn = h("button", { class: "btn btn-primary", type: "button" }, "Check it");
  const out = h("div", { id: "previewOut" });
  const doPreview = async () => {
    U.text = ta.value; U.scope = scope.value;
    if (!U.text.trim()) return toast("Paste something first.", "error");
    busy(prevBtn, true, "Reading\u2026");
    try { U.last = await API.post("/api/import/preview", payload()); drawPreview(out, U.last, doPreview, doApply); }
    catch (e) { clear(out).append(h("div", { class: "alert bad" }, e.message)); }
    finally { busy(prevBtn, false); }
  };
  const doApply = async btn => {
    if (!(await needName())) return;
    busy(btn, true, "Saving\u2026");
    try {
      const r = await API.post("/api/import/apply", payload());
      toast(r.count ? `Saved: ${F.num(r.changes || r.count)} change(s)${r.creates ? `, ${r.creates} new item(s)` : ""}. Undo it from the change log if needed.` : (r.message || "Nothing changed."));
      U.last = null; U.text = ""; await refreshNav(); route();
    } catch (e) { toast(e.message, "error"); busy(btn, false); }
  };
  const payload = () => ({ view: S.view, kind, text: U.text, scope: U.scope || null, opts: U.opts, mapping: U.mapping });
  main.append(card(`2. Paste ${k.label.toLowerCase()}`, [
    h("p", null, k.help), exampleFor(kind),
    kind === "state_pct" || kind === "weather" ? null : h("div", { class: "row" }, scope),
    ta, optionsBox(kind, U.opts, () => { if (U.last) doPreview(); }, U.last && U.last.info),
    h("div", { class: "row end" }, prevBtn)], S.boot.kinds[kind] ? GROUPS.find(g => g.id === k.group).cls : ""), out);
  prevBtn.addEventListener("click", doPreview);
  if (U.last) drawPreview(out, U.last, doPreview, doApply);
}

function exampleFor(kind) {
  const ex = {
    plants: "Area\tPlant\tState\tOpening inventory\nMT-WY\t1001 - Springfield MT\tMT\t10000",
    budget: "Area\tPlant\tApr-26\tMay-26\t\u2026\tMar-27\nMT-WY\t1001 - Springfield MT\t50000\t30000\t\u2026\t100000",
    state_pct: "State\tApr\tMay\t\u2026\tMar\nMT\t100\t100\t\u2026\t80",
    supply: "Area\tLocation\tSupplier\tHub\nMT-WY\tNorthfield, ST\tSupplierX\tCON",
    contract: "Area\tLocation\tSupplier\tApr-26\t\u2026\tMar-27\nMT-WY\tNorthfield, ST\tSupplierX\t200000\t\u2026\t400000",
    hubs: "Hub\tApr-26\tMay-26\t\u2026\nCON\t0.7000\t0.7100\t\u2026",
    lanes: "From\tSupplier\tTo\tRate\tMax gal\nNorthfield, ST\tSupplierX\t1001 - Springfield MT\t0.1500\t\n   or a grid: supply points down, plants across, one rate per cell",
    weather: "Place\tSeason\tApr\tMay\t\u2026\tMar\nAll\t2006-07\t300\t100\t\u2026\t600",
  }[kind] || "Name\tApr-26\tMay-26\t\u2026\nNorthfield, ST\t1.0\t1.0\t\u2026";
  return h("details", { class: "example-box" }, h("summary", null, "Example layout"), h("pre", { class: "example" }, ex),
    h("p", { class: "muted small" }, "Headers can be in any order and spelled loosely; extra columns like Total are ignored. Months can be Apr, Apr-26, 4/1/2026 or P1."));
}

const STATUS = { matched: ["ok", "Matched"], check: ["warn", "Check"], ambiguous: ["warn", "Which one?"], unmatched: ["bad", "Not found"], new: ["info", "Will be added"] };
function drawPreview(out, pv, redo, apply) {
  clear(out);
  const can = S.info.can_edit;
  const sideMap = side => (U.mapping[side] = U.mapping[side] || {});
  const ents = pv.entities || {};
  const problems = pv.names.filter(n => n.status !== "matched");
  const matched = pv.names.filter(n => n.status === "matched");
  const nameRow = n => {
    const list = ents[n.side] || ents.name || [];
    const opts = [h("option", { value: "__auto", selected: n.choice === null }, n.status === "new" ? "Add it as new" : n.suggested ? "Use the suggestion" : "Skip these rows")]
      .concat(n.options.length ? [h("optgroup", { label: "Closest" }, n.options.map(key => { const e = list.find(x => x.key === key); return e ? h("option", { value: key, selected: n.choice === key }, `${e.label} (${e.area_label})`) : null; }))] : [])
      .concat([h("option", { value: "", selected: n.choice === "" }, "Skip these rows")])
      .concat(["plants", "budget", "supply", "contract", "hubs"].includes(pv.kind) && n.side === "name" ? [h("option", { value: "new", selected: n.choice === "new" }, "Add it as new")] : [])
      .concat(list.length <= 800 ? [h("optgroup", { label: "Everything" }, list.map(e => h("option", { value: e.key, selected: n.choice === e.key }, `${e.label} (${e.area_label})`)))] : []);
    const sel = h("select", { "aria-label": "Match for " + n.text, disabled: n.side === "area" }, opts);
    sel.addEventListener("change", () => { const m = sideMap(n.side); if (sel.value === "__auto") delete m[n.key]; else m[n.key] = sel.value; redo(); });
    const sug = n.suggested ? (list.find(e => e.key === n.suggested) || {}).label : "";
    const [cls, txt] = STATUS[n.status] || ["", n.status];
    return h("tr", null, h("td", null, h("b", null, n.text), n.area ? h("div", { class: "muted small" }, n.side === "from" ? "supply point" : n.side === "to" ? "plant" : n.side) : null),
      h("td", null, statusPill(cls, txt), n.score && n.status !== "new" ? h("span", { class: "muted small" }, ` ${Math.round(n.score * 100)}%`) : null),
      h("td", null, sug || ""), h("td", null, n.side === "area" ? h("span", { class: "muted small" }, "Area not found - add it first or fix the name") : sel), h("td", { class: "num" }, n.rows));
  };
  const nameTable = rows => h("table", { class: "table" }, h("thead", null, h("tr", null, ["In your data", "Match", "Suggested", "Use", "Rows"].map((t, i) => h("th", { class: i === 4 ? "num" : "" }, t)))), h("tbody", null, rows.map(nameRow)));
  const unitText = pv.unit ? { gal: "gallons", kgal: "thousand gallons (\u00d71,000)", "$/gal": "$ per gallon", "$/kgal": "$ per thousand gallons (\u00f71,000)", percent: "percent", fraction: "fractions", hdd: "degree days" }[pv.unit] || pv.unit : "";
  const applyBtn = h("button", { class: "btn btn-primary", type: "button", disabled: !can || !pv.changes && !pv.creates.length }, pv.changes || pv.creates.length ? `Save ${F.num(pv.changes)} change(s)${pv.creates.length ? ` + ${pv.creates.length} new` : ""}` : "Nothing to save");
  applyBtn.addEventListener("click", () => apply(applyBtn));
  out.append(card("3. Check what will change", [
    h("div", { class: "pv-summary" },
      h("div", null, h("span", { class: "muted" }, "Read as"), h("b", null, pv.layout || "")), h("div", null, h("span", { class: "muted" }, "Rows"), h("b", null, F.num(pv.rows_read))),
      unitText ? h("div", null, h("span", { class: "muted" }, "Units"), h("b", null, unitText)) : null,
      h("div", null, h("span", { class: "muted" }, "Cells that change"), h("b", null, F.num(pv.changes))),
      pv.creates.length ? h("div", null, h("span", { class: "muted" }, "New"), h("b", null, pv.creates.length)) : null,
      problems.length ? h("div", { class: "warn-text" }, h("span", { class: "muted" }, "Need a look"), h("b", null, problems.length)) : null),
    pv.weather ? h("p", null, `${pv.weather.places} place(s), ${pv.weather.n_seasons} season(s): ${pv.weather.seasons.join(" to ")}.`) : null,
    notesList((pv.notes || []).map(t => ({ level: "info", text: t }))),
    pv.new_areas && pv.new_areas.length ? h("div", { class: "alert info" }, `New area(s) will be added: ${pv.new_areas.join(", ")}`) : null,
    pv.creates.length ? h("details", null, h("summary", null, `${pv.creates.length} new item(s) will be added`), h("ul", { class: "small" }, pv.creates.slice(0, 200).map(c => h("li", null, `${c.type}: ${c.label}`)))) : null,
    problems.length ? [h("h3", null, "Names that need a look"), h("p", { class: "muted small" }, "Point each one at the right plant or supply point, add it as new, or skip it. Rows you skip aren\u2019t loaded."), h("div", { class: "table-scroll" }, nameTable(problems))] : null,
    matched.length ? h("details", null, h("summary", null, `${matched.length} name(s) matched`), h("div", { class: "table-scroll" }, nameTable(matched))) : null,
    pv.areas && pv.areas.length ? h("p", { class: "small" }, "Changes by area: ", pv.areas.map(a => `${a.label} ${F.num(a.cells)}`).join(" \u00b7 ")) : null,
    pv.biggest && pv.biggest.length ? h("details", { open: true }, h("summary", null, "Biggest changes"), h("table", { class: "table" }, h("thead", null, h("tr", null, h("th", null, "Where"), h("th", { class: "num" }, "Now"), h("th", { class: "num" }, "New"))),
      h("tbody", null, pv.biggest.map(b => h("tr", null, h("td", null, b.where), h("td", { class: "num" }, fmtVal(b.old, pv.unit_kind)), h("td", { class: "num" }, fmtVal(b.new, pv.unit_kind))))))) : null,
    h("div", { class: "row end" }, can ? null : h("span", { class: "muted" }, "View only"), applyBtn)], "preview"));
  out.scrollIntoView({ behavior: "smooth", block: "start" });
}
function fmtVal(v, unit) {
  if (v === null || v === undefined) return "blank";
  if (unit === "price") return F.rate(v);
  if (unit === "pct") return F.pct(v);
  return F.num(v);
}

function renderWorkbook(main) {
  const kinds = S.boot.kinds;
  const out = h("div");
  const rows = U.sheets.map((s, i) => {
    const sel = h("select", { "aria-label": "Kind for sheet " + s.sheet }, h("option", { value: "" }, "(skip this sheet)"), S.boot.order.map(k => h("option", { value: k, selected: s.kind === k }, kinds[k].label)));
    const use = h("input", { type: "checkbox", checked: s.use && !!s.kind, "aria-label": "Load sheet " + s.sheet });
    sel.addEventListener("change", () => { s.kind = sel.value; s.use = !!sel.value; use.checked = s.use; });
    use.addEventListener("change", () => { s.use = use.checked; if (s.use && !s.kind) { toast("Pick what this sheet holds first.", "error"); use.checked = s.use = false; } });
    const pv = h("button", { class: "btn btn-small", type: "button" }, "Check");
    pv.addEventListener("click", async () => {
      if (!s.kind) return toast("Pick what this sheet holds first.", "error");
      U.sheet = s.sheet;
      busy(pv, true, "Reading\u2026");
      try { await previewSheet(s, out); } finally { busy(pv, false); }
    });
    return h("tr", { class: U.sheet === s.sheet ? "on" : "" }, h("td", null, use), h("td", null, h("b", null, s.sheet), s.maybe && !s.kind ? h("div", { class: "muted small" }, `looks like ${kinds[s.maybe].label} too - only one sheet of each kind is ticked`) : null),
      h("td", { class: "num" }, F.num(s.rows)), h("td", null, sel), h("td", null, pv));
  });
  const applyAll = h("button", { class: "btn btn-primary", type: "button", disabled: !S.info.can_edit }, "Save the ticked sheets");
  applyAll.addEventListener("click", async () => {
    const list = U.sheets.filter(s => s.use && s.kind);
    if (!list.length) return toast("Tick at least one sheet.", "error");
    if (!(await needName())) return;
    if (!(await confirmBox("Save these sheets?", `${list.map(s => `\u201c${s.sheet}\u201d as ${kinds[s.kind].label}`).join(", ")}. Each sheet is its own entry in the change log, so any one can be undone.`, "Save"))) return;
    busy(applyAll, true, "Saving\u2026");
    try {
      const r = await API.post("/api/import/apply_workbook", { view: S.view, upload: U.upload, sheets: list.map(s => ({ sheet: s.sheet, kind: s.kind, opts: s.opts, mapping: s.mapping })) });
      toast("Saved: " + r.done.map(d => `${d.sheet} (${F.num(d.count)})`).join(", "));
      U.upload = null; U.sheets = []; await refreshNav(); route();
    } catch (e) { toast(e.message, "error"); busy(applyAll, false); }
  });
  main.append(card(`2. Sheets in \u201c${U.uploadName}\u201d`, [
    h("p", { class: "muted" }, "Each sheet was matched to a kind of data by its name and headers. Check a sheet to see exactly what it would change; unticked sheets are ignored."),
    h("table", { class: "table" }, h("thead", null, h("tr", null, h("th", null, "Load"), h("th", null, "Sheet"), h("th", { class: "num" }, "Rows"), h("th", null, "Holds"), h("th", null, ""))), h("tbody", null, rows)),
    h("div", { class: "row end" }, h("button", { class: "btn btn-quiet", type: "button", onclick: () => { U.upload = null; U.sheets = []; route(); } }, "Choose a different file"), applyAll)]), out);
  const cur = U.sheets.find(s => s.sheet === U.sheet);
  if (cur && cur.kind) previewSheet(cur, out);
}
async function previewSheet(s, out) {
  const payload = () => ({ view: S.view, kind: s.kind, upload: U.upload, sheet: s.sheet, opts: s.opts, mapping: s.mapping, scope: s.scope || null });
  const redo = async () => {
    U.mapping = s.mapping;
    try { const pv = await API.post("/api/import/preview", payload()); show(pv); } catch (e) { clear(out).append(h("div", { class: "alert bad" }, e.message)); }
  };
  const apply = async btn => {
    if (!(await needName())) return;
    busy(btn, true, "Saving\u2026");
    try { const r = await API.post("/api/import/apply", { ...payload(), source: `sheet \u201c${s.sheet}\u201d` }); toast(`Saved \u201c${s.sheet}\u201d: ${F.num(r.changes || r.count || 0)} change(s).`); s.use = false; await refreshNav(); route(); }
    catch (e) { toast(e.message, "error"); busy(btn, false); }
  };
  const show = pv => {
    U.mapping = s.mapping;
    drawPreview(out, pv, redo, apply);
    const opt = optionsBox(s.kind, s.opts, redo, pv.info);
    const scope = h("select", { "aria-label": "Which area" }, h("option", { value: "" }, "Rows say which area"), S.info.areas.map(a => h("option", { value: a.id, selected: s.scope === a.id }, `Only ${a.name}`)));
    scope.addEventListener("change", () => { s.scope = scope.value; redo(); });
    const head = $(".card", out);
    head.insertBefore(h("div", { class: "sheet-opts" }, h("b", null, `Sheet \u201c${s.sheet}\u201d as ${S.boot.kinds[s.kind].label}`),
      s.kind === "state_pct" || s.kind === "weather" ? null : scope, opt), head.children[1]);
  };
  U.mapping = s.mapping;
  try { show(await API.post("/api/import/preview", payload())); } catch (e) { clear(out).append(h("div", { class: "alert bad" }, e.message)); }
}
