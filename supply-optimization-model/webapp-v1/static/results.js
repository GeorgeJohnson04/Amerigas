/* results.js - an area's plan: results, insights, what-if. */
"use strict";

async function getResults(body) {
  const r = await API.get("/api/results", { view: S.view, area: S.A.id });
  if (r.none) {
    const b = h("button", { class: "btn btn-primary", type: "button" }, "Run the plan");
    b.addEventListener("click", () => runArea(b));
    body.append(card(null, h("div", { class: "empty-state" }, h("h2", null, "Not run yet"), h("p", null, "Run the plan to see what to buy, what to ship on each lane and what it costs."), b)));
    return null;
  }
  if (!r.fresh) {
    const b = h("button", { class: "btn btn-small", type: "button" }, "Run again");
    b.addEventListener("click", () => runArea(b));
    body.append(h("div", { class: "alert warn" }, h("span", null, "The inputs have changed since this run, so these numbers are out of date."), b));
  }
  return r;
}

async function tabResults(body) {
  const r = await getResults(body);
  if (!r) return;
  const a = S.A.area, s = r.summary, P = r.months.length;
  if (!r.ok) { body.append(card("Couldn\u2019t plan this area", notesList(r.explain))); return; }
  const base = r.baseline;
  const vs = (k, money) => base && base[k] !== undefined && base[k] !== null && s[k] !== null ? `${money ? F.signedMoney(s[k] - base[k]) : F.signedGal(s[k] - base[k])} vs data set` : "";
  body.append(h("div", { class: "kpis" },
    kpi("Delivered cost", F.money(s.total), vs("total", true) || "propane + freight", "c-results big"),
    kpi("Propane", F.money(s.commodity), vs("commodity", true), "c-supply"),
    kpi("Freight", F.money(s.freight), vs("freight", true), "c-lanes"),
    kpi("$ per gallon", F.rate(s.per_gal), base && base.per_gal ? `${F.rate(base.per_gal)} in data set` : "delivered"),
    kpi("Gallons bought", F.galShort(s.bought), `for ${F.galShort(s.demand)} planned demand`),
    kpi("Unmet demand", F.galShort(s.unmet), s.unmet > 100 ? "see notes" : "every plant covered", s.unmet > 100 ? "c-bad" : ""),
    kpi("High-freight gallons", F.galShort(s.high_volume), s.high_volume > 100 ? `${F.money(s.high_cost)} on $1.00+/gal lanes` : "none", s.high_volume > 100 ? "c-warn" : "")),
    r.explain.length ? card("Notes", notesList(r.explain)) : null);

  // cost chart with toggles
  const chartBox = h("div");
  const drawChart = () => {
    clear(chartBox);
    const m = r.monthly, bm = r.baseline_monthly;
    const per = S.ui.resPer === "per";
    const div = (v, i, mm) => per ? (mm.bought[i] > 1 ? v / mm.bought[i] : null) : v;
    const ser = [];
    const showP = S.ui.resCost !== "freight", showF = S.ui.resCost !== "propane";
    if (showP) ser.push({ name: per ? "Propane $/gal" : "Propane $", values: m.commodity.map((v, i) => div(v, i, m)), color: COLORS.propane, stack: "c" });
    if (showF) ser.push({ name: per ? "Freight $/gal" : "Freight $", values: m.freight.map((v, i) => div(v, i, m)), color: COLORS.freight, stack: "c" });
    if (bm) {
      const tot = bm.commodity.map((v, i) => (showP ? v : 0) + (showF ? bm.freight[i] : 0));
      ser.push({ name: "Data set", values: tot.map((v, i) => div(v, i, bm)), type: "line", color: COLORS.base, dash: "5 4" });
    }
    chartBox.append(chart({ labels: r.months, series: ser, yFmt: per ? v => "$" + Number(v).toFixed(2) : F.moneyShort }));
  };
  drawChart();
  body.append(card("Cost by month", [h("div", { class: "toolbar" },
    seg([{ value: "both", label: "Propane + freight" }, { value: "propane", label: "Propane only" }, { value: "freight", label: "Freight only" }], S.ui.resCost, v => { S.ui.resCost = v; saveUi(); drawChart(); }),
    seg([{ value: "total", label: "Total $" }, { value: "per", label: "Per gallon" }], S.ui.resPer, v => { S.ui.resPer = v; saveUi(); drawChart(); }),
    h("span", { class: "spacer" }), h("button", { class: "btn btn-quiet", type: "button", onclick: () => API.download("/api/results/export", { view: S.view, area: S.A.id }) }, "Download (Excel)")), chartBox], "c-results"));

  // detail views
  const VIEWS = [{ value: "lanes", label: "Lane volumes" }, { value: "supply", label: "Supply plan" }, { value: "ship", label: "Shipments" }, { value: "plants", label: "Demand coverage" }, { value: "storage", label: "Storage" }];
  const detail = h("div");
  const drawDetail = () => { clear(detail); ({ lanes: resLanes, supply: resSupply, ship: resShip, plants: resPlants, storage: resStorage }[S.ui.resView] || resLanes)(detail, r); };
  body.append(card("Details", [h("div", { class: "toolbar" }, seg(VIEWS, S.ui.resView, v => { S.ui.resView = v; saveUi(); drawDetail(); })), detail], "c-results"));
  drawDetail();
}

function resLanes(box, r) {
  /* exact gallons from each supply point to each plant, for a month or the season */
  const a = S.A.area, P = r.months.length, hf = S.boot.high_freight;
  const p = S.ui.resMonth ?? -1;
  const msel = h("select", { "aria-label": "Month" }, h("option", { value: -1, selected: p === -1 }, "Whole season"), r.months.map((m, k) => h("option", { value: k, selected: k === p }, m)));
  msel.addEventListener("change", () => { S.ui.resMonth = +msel.value; saveUi(); clear(box); resLanes(box, r); });
  const hide = h("input", { type: "checkbox", checked: S.ui.hideEmpty, id: "hideEmpty" });
  hide.addEventListener("change", () => { S.ui.hideEmpty = hide.checked; saveUi(); clear(box); resLanes(box, r); });
  const val = l => p === -1 ? sum(l.ship) : (l.ship[p] || 0);
  const by = new Map(r.lanes.map(l => [l.from + "|" + l.to, l]));
  let sup = a.supply, pls = a.plants;
  if (S.ui.hideEmpty) {
    sup = sup.filter(s => r.lanes.some(l => l.from === s.id && val(l) > 0.5));
    pls = pls.filter(pl => r.lanes.some(l => l.to === pl.id && val(l) > 0.5));
  }
  const tbl = h("table", { class: "table matrix" }, h("thead", null, h("tr", null, h("th", { class: "row-head" }, "From \\ To"), pls.map(pl => h("th", { class: "num", title: pl.name }, shortPlant(pl.name))), h("th", { class: "num total" }, "Total"))));
  const tb = h("tbody");
  for (const s of sup) {
    let tot = 0;
    tb.appendChild(h("tr", null, h("th", { class: "row-head", title: supName(s) }, h("div", { class: "rh-main" }, s.location), h("div", { class: "rh-sub" }, s.supplier || "")),
      pls.map(pl => { const l = by.get(s.id + "|" + pl.id); if (!l) return h("td", { class: "nolane" }, ""); const g = val(l); tot += g;
        const rate = p === -1 ? null : l.rate[p];
        const isHigh = p === -1 ? l.high.some((x, k) => x && l.ship[k] > 0.5) : l.high[p];
        return h("td", { class: "num" + (g > 0.5 ? " used" : "") + (isHigh && g > 0.5 ? " high" : "") + (p !== -1 && rate === null ? " closed" : ""),
          title: `${s.location} \u2192 ${pl.name}: ${F.gal(g)}${rate !== null && rate !== undefined ? ` at ${F.rate(rate)}/gal` : ""}` }, g > 0.5 ? F.num(g) : ""); }),
      h("td", { class: "num total" }, F.num(tot))));
  }
  const totRow = h("tr", { class: "total-row" }, h("th", { class: "row-head" }, "Total delivered"), pls.map(pl => h("td", { class: "num" }, F.num(r.lanes.filter(l => l.to === pl.id).reduce((t, l) => t + val(l), 0)))),
    h("td", { class: "num total" }, F.num(r.lanes.reduce((t, l) => t + val(l), 0))));
  tb.appendChild(totRow);
  tbl.appendChild(tb);
  box.append(h("div", { class: "toolbar" }, h("label", { class: "inline-label" }, "Month ", msel), h("label", { class: "inline-label" }, hide, " Hide supply points and plants with nothing shipped"),
    h("span", { class: "muted small" }, `Gallons shipped from each supply point to each plant. Amber = freight of $${hf.toFixed(2)}/gal or more.`)),
    h("div", { class: "table-scroll" }, tbl));
}

function resSupply(box, r) {
  const rows = r.supply.map(s => ({ key: s.id, label: s.location, sub: `${s.supplier || ""} \u00b7 contract ${F.galShort(sum(s.contract))} \u00b7 ${F.money(s.cost)}`,
    cells: s.bought.map((v, p) => { const c = s.contract[p] || 0; const atMax = c > 0 && v >= s.max[p] - 1, atMin = c > 0 && !atMax && v <= s.min[p] + 1;
      return { v, cls: c <= 0 ? "blank" : atMax ? "at-max" : atMin ? "at-min" : "", extra: c > 0 ? F.pct(v / c) + " of contract" : "",
        note: `Bought ${F.gal(v)} of ${F.gal(c)} contract (min ${F.gal(s.min[p])}, max ${F.gal(s.max[p])}) at ${F.rate(s.price[p])}/gal` }; }) }));
  box.append(h("p", { class: "muted small" }, "Gallons bought from each supply point. ", h("span", { class: "key at-max" }, "at maximum"), " ", h("span", { class: "key at-min" }, "at minimum"), " \u2013 hover for details."),
    grid({ cols: r.months, rows, editable: false, rowTitle: "Supply point", totalLabel: "Season" }));
}

function resShip(box, r) {
  const a = S.A.area;
  const list = r.lanes.filter(l => l.total > 0.5).sort((x, y) => y.total - x.total);
  const tbl = items => h("table", { class: "table" }, h("thead", null, h("tr", null, ["From", "To", "Gallons", "Avg freight", "Freight $", "Months used"].map((t, i) => h("th", { class: i >= 2 && i <= 4 ? "num" : "" }, t)))),
    h("tbody", null, items.map(l => { const s = supById(l.from), pl = plantById(l.to); const used = l.ship.map((g, p) => g > 0.5 ? r.months[p] : null).filter(Boolean);
      return h("tr", { class: l.high.some((x, p) => x && l.ship[p] > 0.5) ? "high" : "" }, h("td", null, supName(s)), h("td", null, pl ? pl.name : "?"), h("td", { class: "num" }, F.num(l.total)),
        h("td", { class: "num" }, F.rate(l.cost / l.total)), h("td", { class: "num" }, F.money(l.cost)), h("td", { class: "small" }, used.length === r.months.length ? "every month" : used.join(", "))); })));
  box.append(h("p", { class: "muted small" }, `${list.length} lanes carry product. Rows in amber use freight of $1.00/gal or more.`), pager(list, 80, tbl));
}

function resPlants(box, r) {
  const rows = r.plants.map(pl => ({ key: pl.id, label: pl.name, sub: `${pl.state || ""} \u00b7 freight ${F.money(pl.freight)}`,
    cells: pl.demand.map((d, p) => ({ v: pl.delivered[p], cls: pl.unmet[p] > 0.5 ? "short" : "", extra: pl.unmet[p] > 0.5 ? `short ${F.num(pl.unmet[p])}` : (pl.from_store[p] > 0.5 ? `+${F.num(pl.from_store[p])} from storage` : ""),
      note: `Planned ${F.gal(d)} \u00b7 delivered ${F.gal(pl.delivered[p])} \u00b7 from storage ${F.gal(pl.from_store[p])} \u00b7 into storage ${F.gal(pl.into_store[p])} \u00b7 unmet ${F.gal(pl.unmet[p])}` })) }));
  box.append(h("p", { class: "muted small" }, "Gallons delivered to each plant. Red = demand that can\u2019t be covered."), grid({ cols: r.months, rows, editable: false, rowTitle: "Plant", totalLabel: "Season" }));
}

function resStorage(box, r) {
  const rows = [];
  r.supply.filter(s => s.store_cap.some(v => v > 0)).forEach(s => rows.push({ label: s.location, sub: "supply storage", cells: s.store_end.map((v, p) => ({ v, extra: s.store_cap[p] ? `of ${F.num(s.store_cap[p])}` : "" })) }));
  r.plants.filter(p => p.store_cap.some(v => v > 0)).forEach(pl => rows.push({ label: pl.name, sub: "plant storage", cells: pl.store_end.map((v, p) => ({ v, extra: pl.store_cap[p] ? `of ${F.num(pl.store_cap[p])}` : "" })) }));
  box.append(rows.length ? [h("p", { class: "muted small" }, "Gallons in storage at the end of each month."), grid({ cols: r.months, rows, editable: false, totals: false, rowTitle: "Storage" })]
    : h("p", { class: "empty" }, "No storage in this area."));
}

/* ---------------- insights ---------------- */
async function tabInsights(body) {
  const r = await getResults(body);
  if (!r || !r.ok) return;
  const ins = r.insights, months = r.months;
  const ms = list => list.length === months.length ? "every month" : list.map(p => months[p]).join(", ");
  const items = [];
  if (ins.penalty) body.append(h("div", { class: "alert warn" }, "This plan has gaps, so some values below reflect the cost of those gaps rather than real savings. Fix the gaps first for reliable numbers."));
  for (const c of ins.contracts.slice(0, 6)) { const s = supById(c.supply); items.push({ kind: "c-supply", title: `More contract at ${supName(s)}`, text: `Every extra gallon here would save about ${F.rate(c.value)} in ${ms(c.months)}. It\u2019s taken at its maximum because it\u2019s the cheapest option.` }); }
  for (const l of ins.lanes.slice(0, 6)) { const s = supById(l.frm), pl = plantById(l.to); items.push({ kind: "c-lanes", title: `Lane ${s ? s.location : "?"} \u2192 ${pl ? shortPlant(pl.name) : "?"} is full`, text: `It\u2019s at its maximum; each extra gallon allowed would save about ${F.rate(l.value)} in ${ms(l.months)}.` }); }
  for (const st of ins.storage.slice(0, 4)) { const nm = st.kind === "supply" ? supName(supById(st.id)) : (plantById(st.id) || {}).name; items.push({ kind: "c-weather", title: `Storage at ${nm} is full`, text: `More room would be worth about ${F.rate(st.value)}/gal in ${ms(st.months)}.` }); }
  body.append(card("What would make this plan cheaper", items.length ? h("div", { class: "insights" }, items.map(i => h("div", { class: "insight " + i.kind }, h("h3", null, i.title), h("p", null, i.text))))
    : h("p", null, "Nothing is holding the plan back: no contract, lane or storage limit is costing money right now."), "c-results"));
  const dear = ins.demand.slice().sort((x, y) => y.value - x.value);
  body.append(card("What one more gallon costs at each plant", [h("p", { class: "muted small" }, "The delivered cost of the next gallon (propane + freight) - useful for pricing and for spotting expensive plants."),
    h("table", { class: "table" }, h("thead", null, h("tr", null, h("th", null, "Plant"), h("th", { class: "num" }, "Average"), h("th", { class: "num" }, "Dearest month"))),
      h("tbody", null, dear.map(d => { const pl = plantById(d.plant); const k = d.by_month.indexOf(d.dearest); return h("tr", null, h("td", null, pl ? pl.name : d.plant), h("td", { class: "num" }, F.rate(d.value)), h("td", { class: "num" }, `${F.rate(d.dearest)} (${months[k] || ""})`)); })))]));
  body.append(card("How hard each contract is worked", [h("p", { class: "muted small" }, h("span", { class: "key at-max" }, "max"), " taken at its maximum \u00b7 ", h("span", { class: "key at-mid" }, "mid"), " in between \u00b7 ", h("span", { class: "key at-min" }, "min"), " only the minimum"),
    h("table", { class: "table use" }, h("thead", null, h("tr", null, h("th", null, "Supply point"), months.map(m => h("th", null, m)), h("th", { class: "num" }, "Bought / contract"))),
      h("tbody", null, ins.use.map(u => h("tr", null, h("td", null, supName(supById(u.supply))), u.months.map((m, p) => h("td", { class: "use-" + (m || "none"), title: m ? `${months[p]}: ${m}` : "no contract" }, m ? m : "")),
        h("td", { class: "num" }, u.contract ? F.pct(u.bought / u.contract) : "")))))]));
}

/* ---------------- what-if ---------------- */
function tabWhatif(body) {
  const a = S.A.area;
  const hubs = [...new Set(a.supply.map(s => s.hub).filter(Boolean))];
  const presets = [
    { group: "Prices", cases: [{ type: "hub", hub: "*", add: 0.05, label: "All hub prices +$0.05/gal" }, { type: "hub", hub: "*", add: -0.05, label: "All hub prices \u2212$0.05/gal" }]
      .concat(hubs.map(hb => ({ type: "hub", hub: hb, add: 0.05, label: `${hb} +$0.05/gal` }))) },
    { group: "Freight", cases: [{ type: "freight", pct: 10, label: "Freight +10%" }, { type: "freight", pct: -10, label: "Freight \u221210%" }] },
    { group: "Demand", cases: [{ type: "demand", pct: 10, months: [12, 1, 2], label: "Peak winter demand +10%" }, { type: "demand", pct: 5, label: "All demand +5%" }, { type: "demand", pct: -10, months: [10, 11, 12, 1, 2, 3], label: "Winter demand \u221210%" }] },
    { group: "Weather (needs a history)", cases: [{ type: "weather", mode: "cold", label: "1-in-10 cold winter" }, { type: "weather", mode: "coldest", label: "Coldest winter on record" }, { type: "weather", mode: "warm", label: "1-in-10 mild winter" }] },
    { group: "Contracts", cases: [{ type: "contract", supply: "*", pct: -20, label: "Every contract \u221220%" }, { type: "contract", supply: "*", pct: 20, label: "Every contract +20%" }] },
  ];
  const chosen = new Set(["All hub prices +$0.05/gal", "Freight +10%", "Peak winter demand +10%", "Coldest winter on record"]);
  const list = h("div", { class: "whatif-presets" }, presets.map(g => h("fieldset", null, h("legend", null, g.group), g.cases.map(c => h("label", { class: "check" },
    h("input", { type: "checkbox", checked: chosen.has(c.label), onchange: e => e.target.checked ? chosen.add(c.label) : chosen.delete(c.label) }), c.label)))));
  // custom case
  const type = h("select", { "aria-label": "Kind of change" }, [["contract", "One contract changes by %"], ["hub", "One hub price changes by $/gal"], ["lane_off", "A lane is closed"], ["demand", "Demand changes by % in months"]].map(([v, l]) => h("option", { value: v }, l)));
  const pick = h("div", { class: "row wrap" });
  const custom = [];
  const drawPick = () => {
    clear(pick);
    const t = type.value;
    if (t === "contract") pick.append(h("select", { id: "wiSup", "aria-label": "Supply point" }, a.supply.map(s => h("option", { value: s.id }, supName(s)))), h("input", { id: "wiPct", type: "number", value: -20, "aria-label": "Percent" }), "%");
    if (t === "hub") pick.append(h("select", { id: "wiHub", "aria-label": "Hub" }, hubs.map(x => h("option", { value: x }, x))), h("input", { id: "wiAdd", type: "number", step: "0.01", value: 0.05, "aria-label": "$/gal" }), "$/gal");
    if (t === "lane_off") pick.append(h("select", { id: "wiLane", "aria-label": "Lane" }, a.lanes.slice(0, 2000).map(l => h("option", { value: l.from + "|" + l.to }, `${supName(supById(l.from))} \u2192 ${(plantById(l.to) || {}).name}`))));
    if (t === "demand") pick.append(h("input", { id: "wiDp", type: "number", value: 10, "aria-label": "Percent" }), "% in", h("select", { id: "wiDm", "aria-label": "Months" }, Object.entries(SEASONS).map(([k, s]) => h("option", { value: k }, `${s.label} (${s.sub})`))));
  };
  type.addEventListener("change", drawPick); drawPick();
  const customList = h("ul", { class: "rules" });
  const addCustom = () => {
    const t = type.value; let c;
    if (t === "contract") { const s = supById($("#wiSup").value); c = { type: t, supply: s.id, pct: +$("#wiPct").value, label: `${supName(s)} contract ${(+$("#wiPct").value > 0 ? "+" : "")}${$("#wiPct").value}%` }; }
    if (t === "hub") c = { type: t, hub: $("#wiHub").value, add: +$("#wiAdd").value, label: `${$("#wiHub").value} ${+$("#wiAdd").value >= 0 ? "+" : ""}$${$("#wiAdd").value}/gal` };
    if (t === "lane_off") { const [f, to] = $("#wiLane").value.split("|"); c = { type: t, from: f, to, label: `Close ${supName(supById(f))} \u2192 ${(plantById(to) || {}).name}` }; }
    if (t === "demand") { const s = SEASONS[$("#wiDm").value]; c = { type: t, pct: +$("#wiDp").value, months: s.months, label: `Demand ${+$("#wiDp").value > 0 ? "+" : ""}${$("#wiDp").value}% ${s.label}` }; }
    if (!c) return;
    custom.push(c);
    customList.appendChild(h("li", null, c.label));
  };
  const out = h("div");
  const run = h("button", { class: "btn btn-primary", type: "button" }, "Try these");
  run.addEventListener("click", async () => {
    const cases = presets.flatMap(g => g.cases).filter(c => chosen.has(c.label)).concat(custom);
    if (!cases.length) return toast("Tick at least one change to try.", "error");
    busy(run, true, `Trying ${cases.length} change(s)\u2026`);
    clear(out);
    try {
      const r = await API.post("/api/whatif", { view: S.view, area: a.id, cases });
      const b = r.base;
      out.append(h("p", null, `Compared with the plan as it is: ${F.money(b.total)} (${F.rate(b.per_gal)}/gal). Nothing is saved - these are tests only.`),
        h("table", { class: "table" }, h("thead", null, h("tr", null, ["What if\u2026", "Delivered cost", "Change", "$/gal", "Unmet demand", ""].map((t, i) => h("th", { class: i >= 1 && i <= 4 ? "num" : "" }, t)))),
          h("tbody", null, r.cases.map(c => c.error ? h("tr", null, h("td", null, c.case.label), h("td", { colspan: 5, class: "error-text" }, c.error)) :
            h("tr", null, h("td", null, c.case.label), h("td", { class: "num" }, F.money(c.result.total)), h("td", { class: "num " + (c.change > 0.5 ? "up" : c.change < -0.5 ? "down" : "") }, F.signedMoney(c.change)),
              h("td", { class: "num" }, F.rate(c.result.per_gal)), h("td", { class: "num" + (c.result.unmet > 100 ? " warn-text" : "") }, F.gal(c.result.unmet)),
              h("td", null, h("span", { class: "bar-inline " + (c.change > 0 ? "up" : "down"), style: { width: Math.min(100, Math.abs(c.change) / Math.max(...r.cases.map(x => Math.abs(x.change || 0)), 1) * 100) + "%" } })))))));
    } catch (e) { out.append(h("p", { class: "error-text" }, e.message)); }
    finally { busy(run, false); }
  });
  body.append(card("What if\u2026?", [h("p", { class: "muted" }, "Test changes on this area without saving anything. Each one is planned from scratch and compared with the plan as it is now."),
    list, h("h3", null, "Add your own"), h("div", { class: "row wrap" }, type, pick, h("button", { class: "btn", type: "button", onclick: addCustom }, "Add")), customList,
    h("div", { class: "row end" }, run), out], "c-compare"));
}
