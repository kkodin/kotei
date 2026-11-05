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
  const month = UI.detailMonth || UI.day.slice(0, 7), out = [];
  for (const s of detailSites()) for (const [date, d] of Object.entries(s.日)) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !date.startsWith(month + "-") || !hasData(d)) continue;
    // 現場の既定担当者を代用しない。記入者のない過去の日は「未記入」に分ける。
    out.push({ s, date, d, writer: String(d.担当者 || "").trim() || "（記入者未記入）" });
  }
  return out.sort((a, b) => b.date.localeCompare(a.date) || siteLabel(a.s).localeCompare(siteLabel(b.s), "ja"));
}
function detailRenderBrowse() {
  if (!$("writerReports")) return;
  const byWriter = UI.detailView === "writer";
  $("bySite").setAttribute("aria-pressed", String(!byWriter));
  $("byWriter").setAttribute("aria-pressed", String(byWriter));
  $("writerControl").hidden = $("monthControl").hidden = $("writerReports").hidden = !byWriter;
  $("detailContext").textContent = site() ? "選択中：" + siteLabel(site()) : "現場名を選ぶと、この日の日報を開きます";
  if (!byWriter) return;
  $("detailMonth").value = UI.detailMonth || UI.day.slice(0, 7);
  const reports = detailReports(), names = new Map();
  reports.forEach(r => { if (!names.has(nk(r.writer))) names.set(nk(r.writer), r.writer); });
  detailOptions("detailWriter", [...names.values()], UI.detailWriter, "全員");
  const shown = reports.filter(r => !UI.detailWriter || nk(r.writer) === nk(UI.detailWriter));
  $("writerReports").innerHTML = `<p class="note">${shown.length}件。日報を開くと、下の入力欄で内容を確認・編集できます。</p>` +
    (shown.length ? `<table class="detail-list"><thead><tr><th>日付</th><th>記入者</th><th>現場名</th><th>作業予定・完了報告</th><th></th></tr></thead><tbody>${shown.map(r => {
      const work = (r.d.上 || []).map(row => row && [row[1], row[4]].filter(Boolean).join(" → ")).filter(Boolean).join(" ／ ");
      return `<tr${r.s.id === UI.site && r.date === UI.day ? ' class="selected"' : ""}><td>${esc(r.date)}</td><td>${esc(r.writer)}</td><td>${esc(siteLabel(r.s))}</td><td>${esc(work)}</td><td><button type="button" class="sub" data-report-site="${esc(r.s.id)}" data-report-day="${r.date}">日報を開く</button></td></tr>`;
    }).join("")}</tbody></table>` : '<p>この月・絞り込み条件に合う日報はありません。</p>');
}
function detailInit() {
  UI.site = UI.lastSite || UI.site;
  $("siteQuery").oninput = () => detailFilterChange("query", $("siteQuery").value);
  $("sitePeriod").onchange = () => detailFilterChange("period", $("sitePeriod").value);
  $("siteArea").onchange = () => detailFilterChange("area", $("siteArea").value);
  $("clearSiteFilters").onclick = () => { UI.siteFilters = {}; saveUI(); renderPickers(); };
  $("bySite").onclick = () => { UI.detailView = "site"; saveUI(); detailRenderBrowse(); };
  $("byWriter").onclick = () => { UI.detailView = "writer"; saveUI(); detailRenderBrowse(); };
  $("detailWriter").onchange = () => { UI.detailWriter = $("detailWriter").value; saveUI(); detailRenderBrowse(); };
  $("detailMonth").onchange = () => { if (!/^\d{4}-\d{2}$/.test($("detailMonth").value)) return; UI.detailMonth = $("detailMonth").value; saveUI(); detailRenderBrowse(); };
  $("writerReports").onclick = e => {
    const b = e.target.closest("[data-report-site]");
    if (b) detailSelectSite(b.dataset.reportSite, b.dataset.reportDay);
  };
}
