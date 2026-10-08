/* ui.js - small shared helpers: API calls, formatting, DOM building, toasts,
   dialogs, an SVG chart, and the editable grid used for every monthly table. */
"use strict";

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const SEASON = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];
const SEASONS = {
  winter: { label: "Winter", sub: "Oct\u2013Mar", months: [10, 11, 12, 1, 2, 3] },
  peak: { label: "Peak winter", sub: "Dec\u2013Feb", months: [12, 1, 2] },
  shoulder: { label: "Shoulder", sub: "Oct, Nov, Mar", months: [10, 11, 3] },
  summer: { label: "Summer", sub: "Apr\u2013Sep", months: [4, 5, 6, 7, 8, 9] },
  year: { label: "Whole year", sub: "Apr\u2013Mar", months: [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3] },
};
const COLORS = { supply: "#5b5bd6", demand: "#0e9f94", lanes: "#ef7d1a", results: "#15803d", weather: "#0284c7",
                 propane: "#5b5bd6", freight: "#ef7d1a", muted: "#94a3b8", bad: "#dc2626", base: "#64748b" };

function monthNum(label, p) {
  const t = String(label || "").slice(0, 3);
  const k = MONTHS.indexOf(t.charAt(0).toUpperCase() + t.slice(1).toLowerCase());
  return k >= 0 ? k + 1 : ((p + 3) % 12) + 1;
}
function monthsText(ms) {
  const key = [...ms].sort((a, b) => a - b).join(",");
  for (const s of Object.values(SEASONS)) if ([...s.months].sort((a, b) => a - b).join(",") === key) return `${s.label} (${s.sub})`;
  return SEASON.filter(m => ms.includes(m)).map(m => MONTHS[m - 1]).join(", ");
}

/* ---------------- API ---------------- */
const API = {
  user: () => localStorage.getItem("planner.user") || "",
  headers(json) {
    const h = { "X-Requested-With": "supply-planner", "X-User": encodeURIComponent(API.user()) };
    if (json) h["Content-Type"] = "application/json";
    return h;
  },
  async call(method, url, data, params) {
    if (params) url += "?" + new URLSearchParams(Object.entries(params).filter(([, v]) => v !== undefined && v !== null && v !== "")).toString();
    const opts = { method, headers: API.headers(data !== undefined && !(data instanceof FormData)) };
    if (data instanceof FormData) opts.body = data;
    else if (data !== undefined) opts.body = JSON.stringify(data);
    let r;
    try { r = await fetch(url, opts); }
    catch (e) { throw new Error("Can't reach the app. Is the Supply Planner window still open?"); }
    let js = null;
    try { js = await r.json(); } catch (e) { /* not json */ }
    if (!r.ok) throw new Error((js && js.error) || `Something went wrong (${r.status})`);
    return js;
  },
  get: (u, p) => API.call("GET", u, undefined, p),
  post: (u, d) => API.call("POST", u, d || {}),
  patch: (u, d) => API.call("PATCH", u, d || {}),
  del: (u, d) => API.call("DELETE", u, d || {}),
  download(url, params) {
    const q = new URLSearchParams(Object.entries(params || {}).filter(([, v]) => v !== undefined && v !== null && v !== "")).toString();
    const a = document.createElement("a");
    a.href = url + (q ? "?" + q : "");
    a.download = "";
    document.body.appendChild(a); a.click(); a.remove();
  },
};

/* ---------------- formatting ---------------- */
const nf0 = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const F = {
  num: v => (v === null || v === undefined || !isFinite(v)) ? "" : nf0.format(Math.round(v)),
  gal: v => (v === null || v === undefined || !isFinite(v)) ? "\u2013" : nf0.format(Math.round(v)) + " gal",
  galShort(v) {
    if (v === null || v === undefined || !isFinite(v)) return "\u2013";
    const a = Math.abs(v);
    if (a >= 1e6) return (v / 1e6).toFixed(a >= 1e8 ? 0 : a >= 1e7 ? 1 : 2) + "M gal";
    if (a >= 1e3) return (v / 1e3).toFixed(a >= 1e5 ? 0 : 1) + "k gal";
    return nf0.format(v) + " gal";
  },
  money: v => (v === null || v === undefined || !isFinite(v)) ? "\u2013" : (v < 0 ? "-$" : "$") + nf0.format(Math.abs(Math.round(v))),
  moneyShort(v) {
    if (v === null || v === undefined || !isFinite(v)) return "\u2013";
    const a = Math.abs(v), s = v < 0 ? "-$" : "$";
    if (a >= 1e6) return s + (a / 1e6).toFixed(a >= 1e8 ? 0 : 2) + "M";
    if (a >= 1e3) return s + (a / 1e3).toFixed(0) + "k";
    return s + a.toFixed(0);
  },
  signedMoney: v => (v === null || v === undefined || !isFinite(v)) ? "\u2013" : (v > 0 ? "+" : v < 0 ? "\u2212" : "\u00b1") + "$" + nf0.format(Math.abs(Math.round(v))),
  signedGal: v => (v === null || v === undefined || !isFinite(v)) ? "\u2013" : (v > 0 ? "+" : v < 0 ? "\u2212" : "\u00b1") + nf0.format(Math.abs(Math.round(v))) + " gal",
  rate: v => (v === null || v === undefined || !isFinite(v)) ? "" : "$" + Number(v).toFixed(4),
  pct: v => (v === null || v === undefined || !isFinite(v)) ? "" : (Math.round(v * 1000) / 10) + "%",
  factor: v => (v === null || v === undefined || !isFinite(v)) ? "" : "\u00d7" + Number(v).toFixed(2),
  when(ts) {
    if (!ts) return "";
    const d = new Date(ts), now = new Date();
    const same = d.toDateString() === now.toDateString();
    return same ? "today " + d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
                : d.toLocaleDateString([], { month: "short", day: "numeric" }) + " " + d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  },
};

/* ---------------- DOM ---------------- */
function h(tag, attrs, ...kids) {
  const e = document.createElement(tag);
  if (attrs) for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === "class") e.className = v;
    else if (k === "style" && typeof v === "object") Object.assign(e.style, v);
    else if (k.startsWith("on") && typeof v === "function") e.addEventListener(k.slice(2), v);
    else if (k === "html") e.innerHTML = v;
    else if (v === true) e.setAttribute(k, "");
    else e.setAttribute(k, v);
  }
  for (const k of kids.flat(Infinity)) {
    if (k === null || k === undefined || k === false) continue;
    e.appendChild(k instanceof Node ? k : document.createTextNode(String(k)));
  }
  return e;
}
const $ = (s, root) => (root || document).querySelector(s);
const $$ = (s, root) => [...(root || document).querySelectorAll(s)];
function clear(e) { while (e.firstChild) e.removeChild(e.firstChild); return e; }
function esc(s) { return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

function toast(msg, kind) {
  const t = h("div", { class: "toast " + (kind || "ok"), role: kind === "error" ? "alert" : "status" }, msg);
  $("#toasts").appendChild(t);
  setTimeout(() => t.classList.add("out"), kind === "error" ? 7000 : 3500);
  setTimeout(() => t.remove(), kind === "error" ? 7600 : 4100);
}

/* modal: returns a promise of the clicked action's value (or null) */
function modal(title, bodyNodes, actions) {
  return new Promise(resolve => {
    const m = $("#modal");
    $("#modalTitle").textContent = title;
    clear($("#modalBody")).append(...[].concat(bodyNodes));
    const acts = clear($("#modalActions"));
    const prev = document.activeElement;
    const close = v => { m.hidden = true; document.removeEventListener("keydown", onKey); if (prev && prev.focus) prev.focus(); resolve(v); };
    const onKey = e => { if (e.key === "Escape") close(null); };
    document.addEventListener("keydown", onKey);
    for (const a of actions || [{ label: "OK", value: true, primary: true }]) {
      acts.appendChild(h("button", { class: "btn " + (a.primary ? "btn-primary" : a.danger ? "btn-danger" : "btn-quiet"), type: "button",
        onclick: () => { if (a.check && !a.check()) return; close(typeof a.value === "function" ? a.value() : a.value); } }, a.label));
    }
    m.hidden = false;
    const first = $("#modalBody input, #modalBody select, #modalBody textarea") || acts.querySelector(".btn-primary") || acts.firstChild;
    if (first) setTimeout(() => first.focus(), 30);
  });
}
async function confirmBox(title, text, okLabel, danger) {
  return await modal(title, [h("p", null, text)], [{ label: "Cancel", value: false }, { label: okLabel || "OK", value: true, primary: !danger, danger }]);
}
async function promptBox(title, fields, okLabel) {
  /* fields: [{key, label, value, type, options, help, placeholder}] */
  const inputs = {};
  const nodes = fields.map(f => {
    let inp;
    if (f.options) {
      inp = h("select", { id: "pf_" + f.key }, f.options.map(o => h("option", { value: o.value, selected: o.value === f.value }, o.label)));
    } else {
      inp = h("input", { id: "pf_" + f.key, type: f.type || "text", value: f.value ?? "", placeholder: f.placeholder || "", step: f.step });
    }
    inputs[f.key] = inp;
    return h("div", { class: "field" }, h("label", { for: "pf_" + f.key }, f.label), inp, f.help ? h("div", { class: "help" }, f.help) : null);
  });
  const v = await modal(title, nodes, [{ label: "Cancel", value: null },
    { label: okLabel || "OK", primary: true, value: () => Object.fromEntries(Object.entries(inputs).map(([k, i]) => [k, i.value])) }]);
  return v;
}

function busy(btn, on, label) {
  if (!btn) return;
  if (on) { btn.dataset.label = btn.textContent; btn.disabled = true; btn.classList.add("is-busy"); if (label) btn.textContent = label; }
  else { btn.disabled = false; btn.classList.remove("is-busy"); if (btn.dataset.label) btn.textContent = btn.dataset.label; }
}

function seg(options, value, onChange, cls) {
  /* a segmented toggle; options: [{value,label}] */
  const wrap = h("div", { class: "seg " + (cls || ""), role: "group" });
  for (const o of options) {
    const b = h("button", { type: "button", class: o.value === value ? "on" : "", "aria-pressed": o.value === value ? "true" : "false",
      title: o.title || null, onclick: () => {
        $$("button", wrap).forEach(x => { x.classList.remove("on"); x.setAttribute("aria-pressed", "false"); });
        b.classList.add("on"); b.setAttribute("aria-pressed", "true"); onChange(o.value);
      } }, o.label);
    wrap.appendChild(b);
  }
  return wrap;
}

function monthChips(selected, onChange) {
  /* 12 toggle chips Apr..Mar plus season shortcuts */
  const sel = new Set(selected);
  const wrap = h("div", { class: "month-picker" });
  const chips = h("div", { class: "chips" });
  const draw = () => {
    $$(".chip", chips).forEach(c => { const m = +c.dataset.m; c.classList.toggle("on", sel.has(m)); c.setAttribute("aria-pressed", sel.has(m) ? "true" : "false"); });
    $$(".season-btn", wrap).forEach(b => {
      const ms = SEASONS[b.dataset.s].months;
      b.classList.toggle("on", ms.length === sel.size && ms.every(m => sel.has(m)));
    });
  };
  const seasons = h("div", { class: "seasons" }, Object.entries(SEASONS).map(([k, s]) =>
    h("button", { type: "button", class: "season-btn", "data-s": k, onclick: () => { sel.clear(); s.months.forEach(m => sel.add(m)); draw(); onChange([...sel]); } },
      h("b", null, s.label), h("small", null, s.sub))));
  for (const m of SEASON) {
    chips.appendChild(h("button", { type: "button", class: "chip", "data-m": m, onclick: () => { sel.has(m) ? sel.delete(m) : sel.add(m); draw(); onChange([...sel]); } }, MONTHS[m - 1]));
  }
  wrap.append(seasons, chips);
  draw();
  wrap.set = ms => { sel.clear(); ms.forEach(m => sel.add(m)); draw(); };
  return wrap;
}

function statusPill(level, text) { return h("span", { class: "pill " + level }, text); }

function kpi(label, value, sub, cls) {
  return h("div", { class: "kpi " + (cls || "") }, h("div", { class: "kpi-label" }, label), h("div", { class: "kpi-value" }, value), sub ? h("div", { class: "kpi-sub" }, sub) : null);
}

function notesList(items) {
  if (!items || !items.length) return null;
  const icon = { error: "\u26d4", warn: "\u26a0", info: "\u2139", ok: "\u2714" };
  return h("ul", { class: "notes" }, items.map(n => h("li", { class: "note " + n.level }, h("span", { class: "note-ico", "aria-hidden": "true" }, icon[n.level] || "\u2022"),
    h("span", null, n.text))));
}

/* ---------------- chart ---------------- */
/* chart({labels, series:[{name, values, color, type:'bar'|'line', dash, stack}], yFmt, height}) */
function chart(opts) {
  /* drawn at roughly the size it's shown, so axis text stays a normal size */
  const W = opts.width || 1100, H = Math.round((opts.height || 240) * 1.1), padL = 70, padR = 12, padT = 12, padB = 28;
  const labels = opts.labels || [];
  const n = labels.length;
  const series = (opts.series || []).filter(s => s && s.values);
  const bars = series.filter(s => (s.type || "bar") === "bar");
  const lines = series.filter(s => s.type === "line");
  const stacks = {};
  bars.forEach(s => { const k = s.stack || s.name; (stacks[k] = stacks[k] || []).push(s); });
  const groups = Object.values(stacks);
  let max = 0, min = 0;
  for (let i = 0; i < n; i++) {
    for (const g of groups) { let pos = 0, neg = 0; for (const s of g) { const v = +s.values[i] || 0; v >= 0 ? pos += v : neg += v; } max = Math.max(max, pos); min = Math.min(min, neg); }
    for (const s of lines) { const v = s.values[i]; if (v !== null && isFinite(v)) { max = Math.max(max, v); min = Math.min(min, v); } }
  }
  if (max === min) max = min + 1;
  const nice = x => { const p = Math.pow(10, Math.floor(Math.log10(Math.abs(x) || 1))); const m = Math.abs(x) / p; const s = m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10; return Math.sign(x) * s * p; };
  const top = max > 0 ? nice(max * 1.05) : 0, bot = min < 0 ? nice(min * 1.05) : 0;
  const y = v => padT + (H - padT - padB) * (top - v) / (top - bot);
  const cw = (W - padL - padR) / Math.max(n, 1);
  const NS = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(NS, "svg");
  svg.setAttribute("viewBox", `0 0 ${W} ${H}`);
  svg.setAttribute("class", "chart");
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", opts.title || series.map(s => s.name).join(", "));
  const add = (tag, at, text) => { const e = document.createElementNS(NS, tag); for (const [k, v] of Object.entries(at)) e.setAttribute(k, v); if (text !== undefined) e.textContent = text; svg.appendChild(e); return e; };
  const fmt = opts.yFmt || F.num;
  for (let k = 0; k <= 4; k++) {
    const v = bot + (top - bot) * k / 4;
    add("line", { x1: padL, x2: W - padR, y1: y(v), y2: y(v), class: "gridline" });
    add("text", { x: padL - 6, y: y(v) + 4, "text-anchor": "end", class: "axis" }, fmt(v));
  }
  labels.forEach((l, i) => add("text", { x: padL + cw * (i + 0.5), y: H - 8, "text-anchor": "middle", class: "axis" }, l));
  const gw = cw * 0.72 / Math.max(groups.length, 1);
  groups.forEach((g, gi) => {
    for (let i = 0; i < n; i++) {
      let pos = 0, neg = 0;
      for (const s of g) {
        const v = +s.values[i] || 0;
        if (!v) continue;
        const a = v >= 0 ? pos : neg, b = a + v;
        v >= 0 ? pos = b : neg = b;
        const r = add("rect", { x: padL + cw * i + cw * 0.14 + gi * gw, y: Math.min(y(a), y(b)), width: Math.max(gw - 2, 1), height: Math.max(Math.abs(y(a) - y(b)), 0.5),
          fill: s.color || COLORS.muted, rx: 2, opacity: s.opacity || 1 });
        const t = document.createElementNS(NS, "title"); t.textContent = `${s.name} \u00b7 ${labels[i]}: ${fmt(v)}`; r.appendChild(t);
      }
    }
  });
  for (const s of lines) {
    const pts = [];
    s.values.forEach((v, i) => { if (v !== null && isFinite(v)) pts.push([padL + cw * (i + 0.5), y(v), v, i]); });
    if (pts.length) add("polyline", { points: pts.map(p => p[0] + "," + p[1]).join(" "), fill: "none", stroke: s.color || "#111", "stroke-width": 2.2, "stroke-dasharray": s.dash || "" });
    for (const p of pts) { const c = add("circle", { cx: p[0], cy: p[1], r: 3, fill: s.color || "#111" }); const t = document.createElementNS(NS, "title"); t.textContent = `${s.name} \u00b7 ${labels[p[3]]}: ${fmt(p[2])}`; c.appendChild(t); }
  }
  if (bot < 0) add("line", { x1: padL, x2: W - padR, y1: y(0), y2: y(0), class: "zero" });
  const legend = h("div", { class: "legend" }, series.map(s => h("span", null, h("i", { style: { background: s.color || COLORS.muted, height: s.type === "line" ? "3px" : "10px" } }), s.name)));
  return h("figure", { class: "chart-wrap" }, svg, legend);
}

/* ---------------- editable grid ---------------- */
/*
  grid({
    cols: ["Apr-26", ...], rows: [{key, label, sub, cells:[{v, path, cls, note, extra}], total}],
    fmt(v), parse(text) -> {ok, v}, editable, onSave(items) -> Promise, unitHint, rowActions(row)->Node
  })
  Click a cell (or Tab / arrows) to type; paste a block copied from Excel into
  any cell to fill the cells to its right and below. Changes wait in yellow
  until "Save changes".
*/
function grid(o) {
  const pending = new Map();          // "r,c" -> {path, v, text}
  const wrap = h("div", { class: "grid-wrap" });
  const bar = h("div", { class: "grid-bar", hidden: true });
  const table = h("table", { class: "grid" + (o.compact ? " compact" : "") });
  const thead = h("thead", null, h("tr", null, h("th", { class: "row-head", scope: "col" }, o.rowTitle || ""), o.cols.map(c => h("th", { scope: "col" }, c)),
    o.totals !== false ? h("th", { class: "total", scope: "col" }, o.totalLabel || "Total") : null, o.rowActions ? h("th", { class: "act" }, "") : null));
  const tbody = h("tbody");
  table.append(thead, tbody);
  const fmt = o.fmt || F.num;
  const parse = o.parse || (t => { const s = String(t).trim().replace(/[$,\s]/g, ""); if (s === "") return { ok: true, v: null }; const v = Number(s.replace(/%$/, "")); return isFinite(v) ? { ok: true, v } : { ok: false }; });
  const cellsEl = [];
  o.rows.forEach((row, r) => {
    const tr = h("tr", { class: row.cls || "" });
    tr.appendChild(h("th", { class: "row-head", scope: "row", title: row.title || null }, h("div", { class: "rh-main" }, row.label), row.sub ? h("div", { class: "rh-sub" }, row.sub) : null));
    cellsEl[r] = [];
    row.cells.forEach((c, k) => {
      const can = o.editable && c.path && c.edit !== false;
      const td = h("td", { class: (c.cls || "") + (can ? " can" : "") + (c.v === null || c.v === undefined ? " blank" : ""), tabindex: can ? "0" : null,
        title: c.note || null, "data-r": r, "data-c": k });
      td.appendChild(h("span", { class: "val" }, c.text !== undefined ? c.text : fmt(c.v)));
      if (c.extra) td.appendChild(h("span", { class: "extra" }, c.extra));
      tr.appendChild(td);
      cellsEl[r][k] = td;
    });
    if (o.totals !== false) tr.appendChild(h("td", { class: "total" }, row.total !== undefined ? row.total : fmt(row.cells.reduce((a, c) => a + (+c.v || 0), 0))));
    if (o.rowActions) tr.appendChild(h("td", { class: "act" }, o.rowActions(row)));
    tbody.appendChild(tr);
  });

  let editing = null;
  function setPending(r, k, text) {
    const c = o.rows[r].cells[k];
    const p = parse(text);
    const td = cellsEl[r][k];
    if (!p.ok) { td.classList.add("bad"); toast(`\u201c${text}\u201d isn't a number`, "error"); return; }
    td.classList.remove("bad");
    const same = (p.v === null && (c.v === null || c.v === undefined)) || (p.v !== null && c.v !== null && c.v !== undefined && Math.abs(p.v - c.v) < 1e-9);
    if (same) { pending.delete(r + "," + k); td.classList.remove("dirty"); $(".val", td).textContent = c.text !== undefined ? c.text : fmt(c.v); }
    else { pending.set(r + "," + k, { path: c.path, v: p.v }); td.classList.add("dirty"); $(".val", td).textContent = p.v === null ? "" : fmt(p.v); }
    drawBar();
  }
  function drawBar() {
    clear(bar);
    bar.hidden = pending.size === 0;
    if (!pending.size) return;
    const save = h("button", { class: "btn btn-primary", type: "button", onclick: async () => {
      busy(save, true, "Saving\u2026");
      try { await o.onSave([...pending.values()].map(x => ({ path: x.path, value: x.v }))); pending.clear(); }
      catch (e) { toast(e.message, "error"); busy(save, false); }
    } }, `Save ${pending.size} change${pending.size > 1 ? "s" : ""}`);
    bar.append(h("span", null, h("b", null, pending.size), " cell(s) changed, not saved yet."), o.unitHint ? h("span", { class: "muted" }, o.unitHint) : null,
      h("button", { class: "btn btn-quiet", type: "button", onclick: () => { for (const key of [...pending.keys()]) { const [r, k] = key.split(",").map(Number); const c = o.rows[r].cells[k]; const td = cellsEl[r][k]; td.classList.remove("dirty"); $(".val", td).textContent = c.text !== undefined ? c.text : fmt(c.v); } pending.clear(); drawBar(); } }, "Discard"),
      save);
  }
  function startEdit(td) {
    if (editing) finishEdit(true);
    const r = +td.dataset.r, k = +td.dataset.c;
    const c = o.rows[r].cells[k];
    const pend = pending.get(r + "," + k);
    const cur = pend ? pend.v : c.v;
    const inp = h("input", { class: "cell-input", value: cur === null || cur === undefined ? "" : (o.editText ? o.editText(cur) : String(cur)), "aria-label": `${o.rows[r].label} ${o.cols[k]}` });
    td.classList.add("editing");
    td.appendChild(inp);
    inp.focus(); inp.select();
    editing = { td, r, k, inp };
    inp.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === "Tab" || e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        finishEdit(true);
        const dr = e.key === "ArrowUp" ? -1 : (e.key === "Enter" || e.key === "ArrowDown") ? 1 : 0, dk = e.key === "Tab" ? (e.shiftKey ? -1 : 1) : 0;
        const nx = cellsEl[r + dr] && cellsEl[r + dr][k + dk];
        if (nx && nx.classList.contains("can")) startEdit(nx); else td.focus();
      } else if (e.key === "Escape") { finishEdit(false); td.focus(); }
    });
    inp.addEventListener("blur", () => setTimeout(() => { if (editing && editing.inp === inp) finishEdit(true); }, 80));
    inp.addEventListener("paste", e => {
      const text = (e.clipboardData || window.clipboardData).getData("text");
      if (!/[\t\n]/.test(text.trim())) return;
      e.preventDefault();
      finishEdit(false);
      const rows = text.replace(/\r/g, "").replace(/\n$/, "").split("\n").map(l => l.split("\t"));
      let n = 0;
      rows.forEach((vals, dr) => vals.forEach((t, dk) => {
        const cell = cellsEl[r + dr] && cellsEl[r + dr][k + dk];
        if (cell && cell.classList.contains("can")) { setPending(r + dr, k + dk, o.pasteText ? o.pasteText(t) : t); n++; }
      }));
      toast(`Pasted ${n} cell(s) - check them, then Save.`);
    });
  }
  function finishEdit(keep) {
    if (!editing) return;
    const { td, r, k, inp } = editing;
    editing = null;
    const text = inp.value;
    inp.remove();
    td.classList.remove("editing");
    if (keep) setPending(r, k, o.pasteText ? o.pasteText(text) : text);
  }
  tbody.addEventListener("click", e => { const td = e.target.closest("td.can"); if (td && !e.target.classList.contains("cell-input")) startEdit(td); });
  tbody.addEventListener("keydown", e => {
    const td = e.target.closest && e.target.closest("td.can");
    if (!td || e.target.classList.contains("cell-input")) return;
    if (e.key === "Enter" || e.key === "F2" || (e.key.length === 1 && !e.ctrlKey && !e.metaKey)) { e.preventDefault(); startEdit(td); if (e.key.length === 1) { editing.inp.value = e.key; } }
    const r = +td.dataset.r, k = +td.dataset.c;
    const mv = { ArrowRight: [0, 1], ArrowLeft: [0, -1], ArrowDown: [1, 0], ArrowUp: [-1, 0] }[e.key];
    if (mv) { e.preventDefault(); const nx = cellsEl[r + mv[0]] && cellsEl[r + mv[0]][k + mv[1]]; if (nx && nx.tabIndex >= 0) nx.focus(); }
  });
  wrap.append(h("div", { class: "grid-scroll" }, table), bar);
  wrap.hasPending = () => pending.size > 0;
  return wrap;
}

/* rows of a long list in pages, so very big areas stay quick */
function pager(items, size, render) {
  let page = 0;
  const box = h("div");
  const draw = () => {
    clear(box);
    const pages = Math.ceil(items.length / size);
    box.appendChild(render(items.slice(page * size, page * size + size)));
    if (pages > 1) {
      box.appendChild(h("div", { class: "pager" },
        h("button", { class: "btn btn-quiet", type: "button", disabled: page === 0, onclick: () => { page--; draw(); } }, "\u2039 Previous"),
        h("span", null, `${page * size + 1}\u2013${Math.min(items.length, (page + 1) * size)} of ${items.length}`),
        h("button", { class: "btn btn-quiet", type: "button", disabled: page >= pages - 1, onclick: () => { page++; draw(); } }, "Next \u203a")));
    }
  };
  draw();
  return box;
}
