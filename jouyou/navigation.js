/* 詳細日報の入口・現場の絞り込み・記入者別表示（2026-11-05）
   日報や共有データの形式は変えず、選択状態だけを端末に保存する。 */
"use strict";
function detailMeta(s) {
  const m = MS.現場.find(x => s.コード && String(x.コード) === String(s.コード)) ||
    MS.現場.find(x => !s.コード && siteMatch(s.名前, x)) || {};
  return { period: String(m.期 ?? s.期 ?? "").replace(/^第|期$/g, ""), area: String(m.地域 || s.地域 || "") };
}
function detailSites() {
  const f = UI.siteFilters || {};
  return DB.現場.filter(s => {
    const m = detailMeta(s);
    return (!f.query || [s.名前, s.略称, s.コード].some(v => nk(v).includes(nk(f.query)))) &&
      (!f.period || (m.period || "未設定") === f.period) &&
      (!f.area || (m.area || "未設定") === f.area) && (!f.gen || s.元請id === f.gen);
  }).sort((a, b) => (!!a.完了 - !!b.完了) || siteLabel(a).localeCompare(siteLabel(b), "ja", { numeric: true }));
}
function detailOptions(id, items, value, all) {
  const values = [...new Set(items.filter(Boolean))].sort((a, b) => a.localeCompare(b, "ja", { numeric: true }));
  if (value && !values.includes(value)) values.push(value);
  $(id).innerHTML = `<option value="">${all}</option>` + values.map(v => `<option value="${esc(v)}"${v === value ? " selected" : ""}>${esc(v)}</option>`).join("");
}
function detailRenderPickers() {
  const f = UI.siteFilters || {}, chosen = site();
  if (chosen) UI.gen = chosen.元請id;
  detailOptions("sitePeriod", DB.現場.map(s => detailMeta(s).period || "未設定"), f.period, "すべて");
  detailOptions("siteArea", DB.現場.map(s => detailMeta(s).area || "未設定"), f.area, "すべて");
  const gs = DB.元請.filter(g => DB.現場.some(s => s.元請id === g.id)).sort((a, b) => a.名前.localeCompare(b.名前, "ja"));
  $("inGen").innerHTML = '<option value="">すべて</option>' + gs.map(g => `<option value="${esc(g.id)}"${g.id === f.gen ? " selected" : ""}>${esc(g.名前)}</option>`).join("");
  if (document.activeElement !== $("siteQuery")) $("siteQuery").value = f.query || "";
  const ss = detailSites(), outside = chosen && !ss.some(s => s.id === chosen.id);
  $("inSite").innerHTML = '<option value="">現場名を選んでください</option>' +
    (outside ? `<option value="${esc(chosen.id)}" selected>選択中（絞り込み対象外）：${esc(siteLabel(chosen))}</option>` : "") +
    ss.map(s => `<option value="${esc(s.id)}"${s.id === UI.site ? " selected" : ""}>${esc(siteLabel(s))}${s.完了 ? "（完了）" : ""}</option>`).join("");
  if (!chosen) $("inSite").value = "";
  $("siteCount").textContent = ss.length + "件" + (!ss.length ? "（条件に合う現場はありません）" : "");
  $("siteResults").innerHTML = ss.length ? ss.map(s => {
    const m = detailMeta(s), g = DB.元請.find(x => x.id === s.元請id);
    return `<button type="button" class="site-choice${s.id === UI.site ? " selected" : ""}" data-select-site="${esc(s.id)}"><b>${esc(s.略称 || s.名前)}</b><span>${esc(s.名前)}</span><small>${esc([s.コード, m.period ? m.period + "期" : "", m.area, g && g.名前].filter(Boolean).join(" ／ "))}</small></button>`;
  }).join("") : '<p>条件に合う現場はありません。絞り込み条件を変えてください。</p>';
  saveUI(); detailRenderBrowse();
}
function detailFilterChange(key, value) {
  UI.siteFilters = { ...(UI.siteFilters || {}), [key]: value };
  saveUI(); detailRenderPickers();
}
function detailCommit() {
  for (const k of ["up", "dn", "gai", "car"]) if (grids[k] && grids[k].editing()) grids[k].commit();
}
function detailSelectSite(id, date) {
  detailCommit(); askEmptyPages();
  const s = DB.現場.find(x => x.id === id);
  UI.site = s ? s.id : null; UI.gen = s ? s.元請id : null;
  if (date) UI.day = date;
  UI.pg = 0; UNDO = []; REDO = [];
  saveUI(); renderAll(); priceRemind(); askForm();
}
function detailReports() {
  const out = [];
  for (const s of DB.現場) for (const [date, d] of Object.entries(s.日)) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !hasData(d)) continue;
    out.push({ s, date, d, writer: String(d.担当者 || "").trim() || "（記入者未記入）" });
  }
  return out.sort((a, b) => a.date.localeCompare(b.date) || siteLabel(a.s).localeCompare(siteLabel(b.s), "ja", { numeric: true }));
}
function detailSequence() {
  return detailReports().filter(r => UI.detailView === "writer" ? !!UI.detailWriter && nk(r.writer) === nk(UI.detailWriter) : r.s.id === UI.site);
}
function detailAdjacent(n) {
  const rows = detailSequence(), i = rows.findIndex(r => r.s.id === UI.site && r.date === UI.day);
  if (i >= 0) return rows[i + n] || null;
  return n > 0 ? rows.find(r => r.date >= UI.day) || null : rows.slice().reverse().find(r => r.date <= UI.day) || null;
}
function detailStep(n) {
  detailCommit();
  const r = detailAdjacent(n);
  if (!r) { toast(n > 0 ? "次の日報はありません" : "前の日報はありません"); return; }
  detailSelectSite(r.s.id, r.date);
}
function detailWriterNames() {
  const names = new Map();
  detailReports().forEach(r => { if (!names.has(nk(r.writer))) names.set(nk(r.writer), r.writer); });
  return [...names.values()];
}
function detailSetMode(mode) {
  detailCommit(); UI.detailView = mode;
  if (mode === "writer") {
    const names = detailWriterNames(), current = day() && String(day().担当者 || "").trim();
    if (!UI.detailWriter) UI.detailWriter = names.find(n => nk(n) === nk(current || ME.名前)) || names[0] || "";
    const rows = detailSequence(), r = rows.find(x => x.s.id === UI.site && x.date === UI.day) || rows.find(x => x.date >= UI.day) || rows[rows.length - 1];
    if (r) { detailSelectSite(r.s.id, r.date); return; }
  } else if (!site() && UI.lastSite) { detailSelectSite(UI.lastSite); return; }
  saveUI(); refreshSheet();
}
function detailWriterDate(date) {
  const r = detailSequence().find(x => x.date === date && x.s.id === UI.site) || detailSequence().find(x => x.date === date);
  detailSelectSite(r ? r.s.id : "", date);
}
function detailRenderBrowse() {
  if (!$("paperNav")) return;
  const byWriter = UI.detailView === "writer";
  $("bySite").setAttribute("aria-pressed", String(!byWriter));
  $("byWriter").setAttribute("aria-pressed", String(byWriter));
  $("writerControl").hidden = !byWriter;
  detailOptions("detailWriter", detailWriterNames(), UI.detailWriter, "記入者を選ぶ");
  const rows = detailSequence(), i = rows.findIndex(r => r.s.id === UI.site && r.date === UI.day);
  const mismatch = byWriter && i < 0;
  $("sheet").hidden = mismatch;
  document.querySelector(".sheet-actions").hidden = mismatch;
  const fxbar = document.querySelector("#pSheet .fbar"); if (fxbar) fxbar.hidden = mismatch;
  $("detailEmpty").hidden = !mismatch;
  $("detailEmptyText").textContent = UI.detailWriter ? UI.day + " の「" + UI.detailWriter + "」の日報はありません。矢印で前後の日報を開くか、現場名から選んでください。" : "記入者を選んでください。日報の記入者欄が空のものは「記入者未記入」で確認できます。";
  $("detailContext").textContent = (byWriter ? "記入者の日報" : "この現場の日報") + (rows.length ? ` ${i >= 0 ? i + 1 : "—"} / ${rows.length}件` : " 0件");
  for (const [id, n] of [["bPrevD", -1], ["bNextD", 1]]) {
    const r = detailAdjacent(n), el = $(id);
    el.disabled = !r;
    el.title = r ? `${n < 0 ? "前" : "次"}の日報：${r.date} ${siteLabel(r.s)}（${r.writer}）` : `${n < 0 ? "前" : "次"}の日報はありません`;
  }
}
function detailOpenSites() {
  detailCommit(); detailRenderPickers(); openDlg($("dSite"));
  setTimeout(() => $("siteQuery").focus(), 0);
}
function detailInit() {
  UI.site = UI.lastSite || UI.site;
  $("siteQuery").oninput = () => detailFilterChange("query", $("siteQuery").value);
  $("sitePeriod").onchange = () => detailFilterChange("period", $("sitePeriod").value);
  $("siteArea").onchange = () => detailFilterChange("area", $("siteArea").value);
  $("clearSiteFilters").onclick = () => { UI.siteFilters = {}; saveUI(); renderPickers(); };
  $("bySite").onclick = () => detailSetMode("site");
  $("byWriter").onclick = () => detailSetMode("writer");
  $("detailWriter").onchange = () => { UI.detailWriter = $("detailWriter").value; detailSetMode("writer"); };
  for (const id of ["hGen", "hSite"]) for (const el of [$(id), $(id).previousElementSibling]) {
    el.classList.add("site-trigger"); el.tabIndex = 0; el.setAttribute("role", "button"); el.title = "現場名を選ぶ";
    el.onclick = detailOpenSites;
    el.onkeydown = e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); detailOpenSites(); } };
  }
  $("detailChooseSite").onclick = detailOpenSites;
  $("hDate").tabIndex = 0; $("hDate").setAttribute("role", "button");
  $("hDate").onkeydown = e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openCal(); } };
  $("bPrevD").onclick = () => detailStep(-1); $("bNextD").onclick = () => detailStep(1);
  $("siteResults").onclick = e => {
    const b = e.target.closest("[data-select-site]"); if (!b) return;
    UI.detailView = "site"; closeDlg($("dSite")); detailSelectSite(b.dataset.selectSite);
  };
  if (!site()) setTimeout(detailOpenSites, 0);
}
