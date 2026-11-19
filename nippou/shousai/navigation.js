/* 詳細日報の入口・現場の絞り込み・記入者別表示（2026-11-05）
   日報や共有データの形式は変えず、選択状態だけを端末に保存する。 */
"use strict";
const detailPeopleState = { purpose: "view", group: "", query: "", site: "", onlySite: true };
// 退職した人：退職者.json の最後の月の次の月から出さない（開いている日の月で見る。過去の月の日報はそのまま）
const detailGone = (name, ym) => { const g = (MS.退職 || {})[nk(name)]; return !!g && (ym || String(UI.day || "").slice(0, 7)) > g; };
function detailMasterPeople() {
  const found = new Map();
  for (const [key, group] of [["社員", "社員"], ["従業員", "作業員"]]) for (const value of (MS[key] || [])) {
    const name = String(typeof value === "string" ? value : value && (value.名前 || value.氏名) || "").trim();
    if (name && !found.has(nk(name)) && !detailGone(name)) found.set(nk(name), { name, group });
  }
  return [...found.values()].sort((a, b) => a.name.localeCompare(b.name, "ja"));
}
function detailFilterChoices(id, key, items, chosen) {
  const unique = new Map(items.map(x => [x.value, x]));
  const rows = [{ value: "", label: "すべて" }, ...unique.values()];
  $(id).innerHTML = rows.map(x => `<button type="button" class="sub" data-filter-key="${key}" data-filter-value="${esc(x.value)}" aria-pressed="${String((chosen || "") === x.value)}">${esc(x.label)}</button>`).join("");
}
function detailMeta(s) {
  const m = MS.現場.find(x => s.コード && String(x.コード) === String(s.コード)) ||
    MS.現場.find(x => !s.コード && siteMatch(s.名前, x)) || {};
  return { period: String(m.期 ?? s.期 ?? "").replace(/^第|期$/g, ""), area: String(m.地域 || s.地域 || "") };
}
function detailSites() {
  const f = UI.siteFilters || {};
  if (f.gen) { const g = DB.元請.find(x => x.id === f.gen); if (g) f.gen = genName(g); }   // 前の版はボタンに元請の id を覚えていた → 見せる名前に
  return DB.現場.filter(s => {
    const m = detailMeta(s);
    return (!f.query || [s.名前, s.略称, s.コード].some(v => nk(v).includes(nk(f.query)))) &&
      (!f.period || (m.period || "未設定") === f.period) &&
      (!f.area || areaLabel(m.area || "未設定") === f.area) && (!f.gen || genLabelOf(s) === f.gen);   // 2026-11-17：ボタンは見せる名前で絞る（名前を直した・足した候補も）
  }).sort((a, b) => (isHonshaSite(b) - isHonshaSite(a)) || (!!a.完了 - !!b.完了) || siteLabel(a).localeCompare(siteLabel(b), "ja", { numeric: true }));   // 本社をいちばん上に
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
  detailOptions("siteArea", DB.現場.map(s => areaLabel(detailMeta(s).area || "未設定")), f.area, "すべて");
  const gs = DB.元請.filter(g => DB.現場.some(s => s.元請id === g.id)).sort((a, b) => genName(a).localeCompare(genName(b), "ja"));
  $("inGen").innerHTML = '<option value="">すべて</option>' + chipList("元請", gs.map(g => genName(g))).map(v => `<option value="${esc(v)}"${v === f.gen ? " selected" : ""}>${esc(v)}</option>`).join("");
  const sortValues = xs => [...new Set(xs)].sort((a, b) => a.localeCompare(b, "ja", { numeric: true })).map(x => ({ value: x, label: x }));
  detailFilterChoices("periodChoices", "period", sortValues(DB.現場.map(s => detailMeta(s).period || "未設定")), f.period);
  detailFilterChoices("areaChoices", "area", chipList("地域", DB.現場.map(s => areaLabel(detailMeta(s).area || "未設定"))).map(x => ({ value: x, label: x })), f.area);
  detailFilterChoices("genChoices", "gen", chipList("元請", gs.map(g => genName(g))).map(x => ({ value: x, label: x })), f.gen);
  if (document.activeElement !== $("siteQuery")) $("siteQuery").value = f.query || "";
  const ss = detailSites(), outside = chosen && !ss.some(s => s.id === chosen.id);
  $("inSite").innerHTML = '<option value="">現場名を選んでください</option>' +
    (outside ? `<option value="${esc(chosen.id)}" selected>選択中（絞り込み対象外）：${esc(siteLabel(chosen))}</option>` : "") +
    ss.map(s => `<option value="${esc(s.id)}"${s.id === UI.site ? " selected" : ""}>${esc(siteLabel(s))}${s.完了 ? "（完了）" : ""}</option>`).join("");
  if (!chosen) $("inSite").value = "";
  $("siteCount").textContent = ss.length + "件" + (!ss.length ? "（条件に合う現場はありません）" : "");
  $("siteResults").innerHTML = ss.length ? ss.map(s => {
    const m = detailMeta(s), g = DB.元請.find(x => x.id === s.元請id);
    return `<button type="button" class="site-choice${s.id === UI.site ? " selected" : ""}" data-select-site="${esc(s.id)}"><b>${esc(s.略称 || s.名前)}</b>${s.略称 && s.略称 !== s.名前 ? `<span>${esc(s.名前)}</span>` : ""}<small>${esc([s.コード, m.period ? m.period + "期" : "", m.area, genName(g)].filter(Boolean).join(" ／ "))}</small></button>`;
  }).join("") : '<p>条件に合う現場はありません。絞り込み条件を変えてください。</p>';
  saveUI(); detailRenderBrowse();
}
/* ---------- 元請名・地域のボタン：右クリック（iPad は長押し）で 消す・候補を足す・名前を直す（2026-11-17 ユーザー） ----------
   みんなで同じ（共通.json の 絞り込み）。消す＝ボタンを出さないだけ（現場・日報のデータは変えない）。
   名前を直す＝ボタンと用紙の元請名の見せ方だけ（元の名前はデータに残る）。候補を足す＝現場がまだ無くてもボタンを出す。 */
const chipCfg = kind => (DB.絞り込み ||= {})[kind] ||= { 隠す: [], 足す: [], 名前: {} };
const areaLabel = a => chipCfg("地域").名前[a] || a;
const genLabelOf = s => genName(DB.元請.find(g => g.id === s.元請id));
function chipList(kind, labels) {
  const c = chipCfg(kind), hide = new Set(c.隠す);
  return [...new Set([...labels, ...c.足す])].filter(v => v && !hide.has(v)).sort((a, b) => a.localeCompare(b, "ja", { numeric: true }));
}
function chipSave(msg) { DB.共通更新 = now(); save(); detailRenderPickers(); refreshSheet(); if (msg) toast(msg); }
function chipMenu(kind, value, x, y) {
  const dlg = $("dSite"); closeChipMenu();
  const c = chipCfg(kind), word = kind === "元請" ? "元請名" : "地域";
  const items = [];
  if (value) {
    items.push([`「${value}」を一覧から消す`, () => { if (!confirm(`${word}のボタン「${value}」を出さないようにしますか？\n（現場や日報のデータは変わりません。みんなの画面から消えます。あとで「消したものを戻す」で戻せます）`)) return; c.隠す.push(value); c.足す = c.足す.filter(v => v !== value); const f = UI.siteFilters || {}; if (f[kind === "元請" ? "gen" : "area"] === value) f[kind === "元請" ? "gen" : "area"] = ""; chipSave(`「${value}」を一覧から消しました`); }]);
    items.push([`「${value}」の名前を直す（編集する）`, () => {
      const v = (prompt(`${word}「${value}」の見せる名前（空にすると元に戻す）`, value) || "").trim(); if (v === value) return;
      const orig = Object.keys(c.名前).find(k => c.名前[k] === value) || (kind === "元請" ? (DB.元請.find(g => genName(g) === value) || {}).名前 : null) || value;
      if (c.足す.includes(value)) { c.足す = c.足す.map(x => x === value ? v : x).filter(Boolean); }
      else if (v) c.名前[orig] = v; else delete c.名前[orig];
      chipSave(v ? `「${value}」を「${v}」と見せるようにしました` : `「${value}」の名前を元に戻しました`);
    }]);
  }
  items.push([`${word}の候補を足す`, () => { const v = (prompt(`足す${word}（現場がまだ無くてもボタンを出します）`) || "").trim(); if (!v) return; c.隠す = c.隠す.filter(x => x !== v); if (!c.足す.includes(v)) c.足す.push(v); chipSave(`${word}「${v}」を足しました`); }]);
  if (c.隠す.length) items.push([`消したものを戻す（${c.隠す.length}件）`, () => { const v = (prompt(`戻す${word}を1つ打ってください：\n${c.隠す.join("、")}`, c.隠す[0]) || "").trim(); if (!v || !c.隠す.includes(v)) return; c.隠す = c.隠す.filter(x => x !== v); chipSave(`「${v}」を戻しました`); }]);
  const m = document.createElement("div"); m.className = "ctx"; m.id = "chipCtx";
  m.innerHTML = items.map((it, i) => `<div data-i="${i}"><span>${esc(it[0])}</span></div>`).join("");
  dlg.appendChild(m);   // 開いている窓の中に出す（窓の外は押せないため）
  const r = dlg.getBoundingClientRect();
  m.style.position = "fixed"; m.style.left = Math.min(x, innerWidth - m.offsetWidth - 4) + "px"; m.style.top = Math.min(y, innerHeight - m.offsetHeight - 4) + "px";
  m.addEventListener("click", e => { const d = e.target.closest("[data-i]"); if (!d) return; closeChipMenu(); items[+d.dataset.i][1](); });
}
function closeChipMenu() { const m = $("chipCtx"); if (m) m.remove(); }
document.addEventListener("pointerdown", e => { if (!e.target.closest("#chipCtx")) closeChipMenu(); });
function chipBind() {
  for (const [id, kind] of [["areaChoices", "地域"], ["genChoices", "元請"]]) {
    const el = $(id); if (!el || el.dataset.menu) continue; el.dataset.menu = "1";
    const target = e => { const b = e.target.closest("[data-filter-value]"); return b ? b.dataset.filterValue : ""; };
    el.addEventListener("contextmenu", e => { e.preventDefault(); chipMenu(kind, target(e), e.clientX, e.clientY); });
    let t = null;   // iPad：長押し
    el.addEventListener("pointerdown", e => { if (e.pointerType !== "touch") return; const v = target(e), x = e.clientX, y = e.clientY; t = setTimeout(() => { t = "done"; chipMenu(kind, v, x, y); }, 550); });
    const stop = () => { if (t && t !== "done") clearTimeout(t); };
    el.addEventListener("pointerup", e => { if (t === "done") { e.preventDefault(); e.stopPropagation(); } stop(); t = null; }, true);
    el.addEventListener("pointercancel", stop); el.addEventListener("pointermove", e => { if (e.pointerType === "touch") stop(); });
    el.addEventListener("click", e => { if (t === "done") { e.stopPropagation(); e.preventDefault(); } }, true);
  }
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
  if (s && UI.newForm) {   // 新規の日報で先に選んだ区分を、この現場のこの日に（まだ書いていない日は覚えるだけ。書き始めたときに入る）
    const d0 = s.日[date || UI.day];
    if (d0 && hasData(d0)) { d0.形態 = UI.newForm; d0.更新 = now(); }
    s.前回形態 = UI.newForm; s.設定更新 = now(); if (!s.形態) { s.形態 = UI.newForm; formShare(s); }
    UI.newForm = "";
  }
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
  return detailReports().filter(r => UI.detailView === "writer" ? detailWriterNames().some(n => nk(n) === nk(UI.detailWriter)) && nk(r.writer) === nk(UI.detailWriter) : r.s.id === UI.site);
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
function detailWriterNames() { return detailMasterPeople().map(p => p.name); }
function detailOpenPeople(purpose) {
  if (purpose === "entry" && (!site() || !canEdit())) { toast("編集できる現場の日報で選んでください"); return; }
  detailCommit(); detailPeopleState.purpose = purpose || "view";
  detailPeopleState.group = ""; detailPeopleState.query = ""; $("peopleQuery").value = "";
  detailPeopleState.site = UI.site || ""; detailPeopleState.onlySite = !!UI.site;   // 開いている現場に行った人から（2026-11-17 ユーザー）
  detailRenderPeople(); openDlg($("dPeople"));
}
// その現場に行った人：その現場の日報の 記入者・作業員（下の表）・外注（名前）と、現場の名簿
function detailSitePeople(id) {
  const s = DB.現場.find(x => x.id === id), out = new Set(); if (!s) return out;
  Object.values(s.日 || {}).forEach(d => { if (!d) return; if (d.担当者) out.add(nk(d.担当者)); [...(d.下 || []), ...(d.外注 || [])].forEach(r => { if (r && r[0]) out.add(nk(r[0])); }); });
  (s.作業員 || []).forEach(w => { if (w && w.名前) out.add(nk(w.名前)); });
  return out;
}
function detailRenderPeople() {
  if (!$("peopleResults")) return;
  const state = detailPeopleState, all = detailMasterPeople();
  const s = state.site && DB.現場.find(x => x.id === state.site), went = s ? detailSitePeople(s.id) : null, bySite = !!(s && state.onlySite);
  $("peopleGroups").innerHTML = [["", "すべて"], ["社員", "社員"], ["作業員", "作業員"]].map(([key, label]) => `<button type="button" class="sub" data-people-group="${key}" aria-pressed="${String(state.group === key)}">${label}</button>`).join("") +
    (s ? `<button type="button" class="sub" data-people-site="1" aria-pressed="${String(bySite)}" title="押すと切り替わります">${bySite ? "この現場に行った人：" + esc(s.略称 || s.名前) : "全員（現場で絞らない）"}</button>` : "");
  const selected = state.purpose === "entry" ? (day() && day().担当者) : UI.detailWriter;
  const visible = all.filter(p => (!state.group || p.group === state.group) && (!state.query || nk(p.name).includes(nk(state.query))) && (!bySite || went.has(nk(p.name))));
  const reports = detailReports();
  $("peopleResults").innerHTML = ["社員", "作業員"].filter(group => !state.group || group === state.group).map(group => {
    const people = visible.filter(p => p.group === group);
    return `<section><h4>${group} <small>${people.length}名</small></h4><div class="people-grid">${people.map(p => {
      const count = reports.filter(r => nk(r.writer) === nk(p.name)).length;
      return `<button type="button" class="person-choice" data-person-name="${esc(p.name)}" aria-pressed="${String(nk(selected) === nk(p.name))}"><b>${esc(p.name)}</b><small>${count ? "日報 " + count + "件" : "日報はまだありません"}</small></button>`;
    }).join("") || '<p class="note">該当する人はいません</p>'}</div></section>`;
  }).join("");
  $("peopleMasterStatus").textContent = all.length ? `社員・作業員マスター ${all.length}名${MS.読んだ ? " ／ 読み込み " + MS.読んだ : ""}` : "マスターが未読、または登録がありません。サインインしてマスターを読み直してください。";
}
function detailPickPerson(name) {
  const person = detailMasterPeople().find(p => nk(p.name) === nk(name));
  if (!person) { toast("マスターに登録されている人を選んでください"); detailRenderPeople(); return; }
  if (detailPeopleState.purpose === "entry") {
    if (!site() || !canEdit()) return;
    $("hTan").value = person.name; $("hTan").dispatchEvent(new Event("change"));
    closeDlg($("dPeople"));
  } else {
    UI.detailWriter = person.name; closeDlg($("dPeople")); detailSetMode("writer");
  }
}
function detailSetMode(mode) {
  detailCommit(); UI.detailView = mode;
  if (mode === "writer") {
    const names = detailWriterNames(), current = day() && String(day().担当者 || "").trim();
    if (!names.some(n => nk(n) === nk(UI.detailWriter))) UI.detailWriter = names.find(n => nk(n) === nk(current || ME.名前)) || names[0] || "";
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
  $("detailWriter").textContent = UI.detailWriter && detailWriterNames().some(n => nk(n) === nk(UI.detailWriter)) ? "記入者：" + UI.detailWriter : "記入者を選ぶ";
  if ($("dPeople").open) detailRenderPeople();
  const rows = detailSequence(), i = rows.findIndex(r => r.s.id === UI.site && r.date === UI.day);
  const mismatch = byWriter && i < 0;
  $("sheet").hidden = mismatch;
  document.querySelector(".sheet-actions").hidden = mismatch;
  const fxbar = document.querySelector("#pSheet .fbar"); if (fxbar) fxbar.hidden = mismatch;
  $("detailEmpty").hidden = !mismatch;
  $("detailEmptyText").textContent = UI.detailWriter ? UI.day + " の「" + UI.detailWriter + "」の日報はありません。矢印で前後の日報を開くか、現場名から選んでください。" : "社員・作業員マスターの一覧から記入者を選んでください。";
  $("detailContext").textContent = (byWriter ? "記入者の日報" : "この現場の日報") + (rows.length ? ` ${i >= 0 ? i + 1 : "—"} / ${rows.length}件` : " 0件");
  for (const [id, n] of [["bPrevD", -1], ["bNextD", 1]]) {
    const r = detailAdjacent(n), el = $(id);
    el.disabled = !r;
    el.title = r ? `${n < 0 ? "前" : "次"}の日報：${r.date} ${siteLabel(r.s)}（${r.writer}）` : `${n < 0 ? "前" : "次"}の日報はありません`;
  }
}
/* 現場名を選んだ日にまだ日報が無ければ、記入者（自分）だけの日報を作る（2026-11-18 ユーザー：現場名と記入者が入っていれば、ほかが空でも保存する）。
   ほかの人の日報を上書きしないよう、この現場の「使用中」の札が自分のものになり、共有を読み直してもまだ無いときだけ作る */
async function detailNewDay(sid, k) {
  const same = () => UI.site === sid && UI.day === k && !day() && !RO;
  for (let i = 0; i < 40 && signedIn() && !DEMO; i++) {   // 札を待つ（20秒まで）
    if (!same()) return;
    if (LOCK.site === sid && LOCK.mine) break;
    await new Promise(r => setTimeout(r, 500));
  }
  if (!same() || (signedIn() && !DEMO && !(LOCK.site === sid && LOCK.mine))) return;
  if (signedIn() && !DEMO) { try { await sync(); } catch (e) { return; } if (!same()) return; }
  if (!myName() || !canEdit()) return;
  const d = ensureDay(); if (!d || !d.担当者) return;
  d.更新 = now(); save(); renderTabs(); refreshSheet();
  toast(`${siteLabel(site())}　${jpDate(k)} の日報を作りました（記入者：${d.担当者}）`, 4000);
}
function detailOpenSites() {
  detailCommit(); detailRenderPickers(); chipBind(); openDlg($("dSite"));
  setTimeout(() => $("siteQuery").focus(), 0);
}
function detailInit() {
  dayToolsInit();
  UI.site = UI.lastSite || UI.site;
  $("siteQuery").oninput = () => detailFilterChange("query", $("siteQuery").value);
  $("sitePeriod").onchange = () => detailFilterChange("period", $("sitePeriod").value);
  $("siteArea").onchange = () => detailFilterChange("area", $("siteArea").value);
  $("clearSiteFilters").onclick = () => { UI.siteFilters = {}; saveUI(); renderPickers(); };
  $("bySite").onclick = () => detailSetMode("site");
  $("byWriter").onclick = () => { detailSetMode("writer"); detailOpenPeople("view"); };
  $("detailWriter").onclick = () => detailOpenPeople("view");
  $("siteFilterButtons").onclick = e => {
    const b = e.target.closest("[data-filter-key]"); if (b) detailFilterChange(b.dataset.filterKey, b.dataset.filterValue);
  };
  $("peopleGroups").onclick = e => {
    if (e.target.closest("[data-people-site]")) { detailPeopleState.onlySite = !detailPeopleState.onlySite; detailRenderPeople(); return; }
    const b = e.target.closest("[data-people-group]"); if (b) { detailPeopleState.group = b.dataset.peopleGroup; detailRenderPeople(); }
  };
  $("peopleQuery").oninput = () => { detailPeopleState.query = $("peopleQuery").value; detailRenderPeople(); };
  $("peopleClear").onclick = () => { detailPeopleState.group = ""; detailPeopleState.query = ""; $("peopleQuery").value = ""; detailRenderPeople(); };
  $("peopleResults").onclick = e => { const b = e.target.closest("[data-person-name]"); if (b) detailPickPerson(b.dataset.personName); };
  $("peopleReload").onclick = async () => {
    if (DEMO) { detailRenderPeople(); toast("試し画面では用意したマスターを参照しています"); return; }
    if (!signedIn()) { toast("サインインしてからマスターを読み直してください"); return; }
    const b = $("peopleReload"); b.disabled = true;
    try { await loadMasters(false); detailRenderPeople(); } finally { b.disabled = false; }
  };
  // 元請業者は押しても選ばない（2026-11-18 ユーザー：現場名を選べば出る）。現場名だけが「現場名を選ぶ」の入口。（本社）✓ を押したときは窓を出さない
  for (const el of [$("hSite"), $("hSite").previousElementSibling]) {
    el.classList.add("site-trigger"); el.tabIndex = 0; el.setAttribute("role", "button"); el.title = "現場名を選ぶ";
    el.onclick = e => { if (e.target.closest(".honChk")) return; detailOpenSites(); };
    el.onkeydown = e => { if (e.target.closest(".honChk")) return; if (e.key === "Enter" || e.key === " ") { e.preventDefault(); detailOpenSites(); } };
  }
  $("hGen").title = "現場名を選ぶと出ます";
  $("detailChooseSite").onclick = detailOpenSites;
  $("hDate").tabIndex = 0; $("hDate").setAttribute("role", "button");
  $("hDate").onkeydown = e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openCal(); } };
  $("bPrevD").onclick = () => detailStep(-1); $("bNextD").onclick = () => detailStep(1);
  $("siteResults").onclick = e => {
    const b = e.target.closest("[data-select-site]"); if (!b) return;
    UI.detailView = "site"; closeDlg($("dSite")); const k = UI.day; detailSelectSite(b.dataset.selectSite);
    if (!day()) detailNewDay(b.dataset.selectSite, k);
  };
  // 開いたときに「現場名を選ぶ」を自動で出すのはやめた（2026-11-15 ユーザー：入口の工事日報から現場・日を決めて開く）。現場名の欄を押せば今までどおり選べる
}
