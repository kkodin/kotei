/* 日付フォームの削除・コピー。コピー元と先を明示し、元の日報は変更しない。 */
"use strict";
let DAY_COPY = null;
function dayCopyOpen() {
  detailCommit(); const s = site(); if (!s || !canEdit()) return;
  const dates = Object.keys(s.日).filter(k => hasData(s.日[k])).sort();
  if (!dates.length) { toast("コピー元の日報がありません"); return; }
  const from = hasData(s.日[UI.day]) ? UI.day : dates.filter(k => k < UI.day).pop() || dates[0];
  DAY_COPY = { site: s.id, from, to: UI.day, fromMonth: from.slice(0, 7), toMonth: UI.day.slice(0, 7) };
  if ($("dCal").open) closeDlg($("dCal"));
  dayCopyRender(); openDlg($("dDayCopy"));
}
function dayCopyCalendar(side) {
  const c = DAY_COPY, s = site(), month = c[side + "Month"], [year, m] = month.split("-").map(Number);
  const first = new Date(year, m - 1, 1), n = new Date(year, m, 0).getDate();
  let html = `<div class="calnav"><button type="button" class="sub" data-copy-month="${side}|-1" aria-label="${side === "from" ? "コピー元" : "コピー先"}の前の月">◀</button><b>${year}年${m}月</b><button type="button" class="sub" data-copy-month="${side}|1" aria-label="${side === "from" ? "コピー元" : "コピー先"}の次の月">▶</button></div><div class="calgrid">`;
  html += [...WD].map((w, i) => `<b class="${i === 0 ? "sun" : i === 6 ? "sat" : ""}">${w}</b>`).join("");
  html += "<span></span>".repeat(first.getDay());
  for (let d = 1; d <= n; d++) {
    const date = `${month}-${pad(d)}`, has = !!s && hasData(s.日[date]);
    html += `<button type="button" data-copy-date="${side}|${date}" class="${has ? "has " : ""}${date === c[side] ? "on" : ""}"${side === "from" && !has ? " disabled" : ""}><span class="dn">${d}</span>${has ? '<span class="mk">●</span>' : ""}</button>`;
  }
  return html + "</div>";
}
function dayCopyRender() {
  const c = DAY_COPY, s = site(); if (!c) return;
  $("dayCopySite").textContent = s ? "現場：" + siteLabel(s) : "現場を選んでください";
  $("copyFromCal").innerHTML = dayCopyCalendar("from"); $("copyToCal").innerHTML = dayCopyCalendar("to");
  $("dayCopySummary").textContent = `${c.from} の内容 → ${c.to} へコピー`;
  const invalid = !s || s.id !== c.site || !hasData(s.日[c.from]) || c.from === c.to;
  $("dayCopyWarning").textContent = c.from === c.to ? "コピー元とコピー先に違う日を選んでください。" : s && hasData(s.日[c.to]) ? "コピー先には日報があります。作業予定・人員・機材・材料を置き換えます。" : "";
  $("dayCopyRun").disabled = invalid;
}
function dayCopyRun() {
  const c = DAY_COPY, s = site(); detailCommit();
  if (!c || !s || s.id !== c.site || !canEdit() || c.from === c.to || !hasData(s.日[c.from])) return;
  if (!confirm(`「${siteLabel(s)}」\n${c.from} から ${c.to} へコピーします。${hasData(s.日[c.to]) ? "\nコピー先の作業予定・人員・機材・材料を置き換えます。" : ""}\nコピー先の完了報告は空欄になります。実行しますか？`)) return;
  if (site() !== s || !canEdit() || !hasData(s.日[c.from])) return;
  const original = JSON.parse(JSON.stringify(s.日[c.from]));
  UI.detailView = "site"; detailSelectSite(s.id, c.to); pushUndo();
  const d = ensureDay();
  d.上 = (original.上 || []).map(r => { if (r) { while (r.length < 5) r.push(""); r[4] = ""; } return r; });
  d.下 = original.下 || []; d.外注 = original.外注 || []; d.頁 = pageCount(original); d.人員手 = original.人員手;
  d.更新 = now();
  if (s.消した日[c.to] && s.消した日[c.to] >= d.更新) {
    const t = new Date(s.消した日[c.to].replace(" ", "T")); t.setSeconds(t.getSeconds() + 1);
    d.更新 = `${ymd(t)} ${pad(t.getHours())}:${pad(t.getMinutes())}:${pad(t.getSeconds())}`;
  }
  tidy(d); UI.pg = 0; save(); saveUI(); closeDlg($("dDayCopy")); refreshSheet();
  toast(`${c.from} → ${c.to} にコピーしました（↶で戻せます）`);
}
function dayToolsInit() {
  $("calDeleteMode").onclick = () => { if (site() && canEdit()) openCal("del"); };
  $("calCopyMode").onclick = dayCopyOpen;
  $("calNormalMode").onclick = () => openCal();
  $("dDayCopy").addEventListener("click", e => {
    if (!DAY_COPY) return;
    const date = e.target.closest("[data-copy-date]"), month = e.target.closest("[data-copy-month]");
    if (date && !date.disabled) { const [side, value] = date.dataset.copyDate.split("|"); DAY_COPY[side] = value; dayCopyRender(); }
    if (month) { const [side, n] = month.dataset.copyMonth.split("|"), [y, m] = DAY_COPY[side + "Month"].split("-").map(Number); DAY_COPY[side + "Month"] = ymd(new Date(y, m - 1 + +n, 1)).slice(0, 7); dayCopyRender(); }
  });
  $("dayCopyRun").onclick = dayCopyRun;
}
