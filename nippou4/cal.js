/* 工事日報 Vr004（2026-11-02 ユーザー：クロジカで日々の報告をしているので、みんなが慣れたクロジカに似た形に）
   ・カレンダー：週（人×日）・日（人の列）・月。全体工程表のその日の予定（現場・作業）を上の段に重ねて出す。
   ・予定・報告（1行）を押すと「スケジュール詳細」、詳細の「編集する」・空いたます目・下の「＋」で「スケジュール編集」の窓。
     記入者＝サインインして書いている人（代理で書くときは「自分も参加」を外して参加者を選ぶ）。現場＝報告する現場。
   ・自社請けの現場は発注者から選べる（現場一覧の 元請・発注者）。
   ・設備（自社機械）は自社機械稼働表の予約と連動：その日・その現場の予約を予定の設備に入れ、報告で使った機械は稼働表に予約を足す
     （続いている予約は延ばす）。「使わなかった」は稼働表の 未使用 にする。稼働表を読めない人（ゲスト）は日報だけ。
   ・現場ごとの請負形態（常用／請負）：日報のフォルダの 現場設定.json。常用の現場は常用日報の形で集計する（2回目で作る）。
   ・集計：人別（人×日の行先・日数・時間）・現場別（現場×日の人数・人日）。
   データは Vr003 と同じ 個人_<名前>_<月>.json の 行。足したもの：行.終日・行.本人（false＝記入者は参加していない＝代理）・行.未使用機械・行.記入者・行.更新 */
"use strict";
const CAL = { mode: "week", hide: {}, home: "" };
try { Object.assign(CAL, JSON.parse(localStorage.getItem("nippou_cal_v4")) || {}); } catch (e) { }
const calSave = () => { try { localStorage.setItem("nippou_cal_v4", JSON.stringify({ mode: CAL.mode, hide: CAL.hide, home: CAL.home })); } catch (e) { } };
// はじめに開く画面：いつもカレンダー（PC もスマホも。クロジカと同じ）。カードに切り替えても覚えない
// （2026-11-02 ユーザー：開くと入力画面が出る。クロジカのカレンダーの画面からにしたい）
function calHome() { return "cal"; }
const MODES = [["week", "週"], ["day", "日"], ["month", "月"], ["person", "人別"], ["site", "現場別"]];
const PAL = ["#3f51b5", "#e53935", "#795548", "#827717", "#7986cb", "#00acc1", "#5e35b1", "#00695c", "#e57373", "#fb8c00", "#558b2f", "#1565c0", "#ad1457", "#6d4c41"];
function colorOf(n) {
  const i = people().findIndex(x => nk(x) === nk(n));
  return PAL[(i >= 0 ? i : [...nk(n)].reduce((a, c) => a + c.charCodeAt(0), 0)) % PAL.length];
}
const weekStart = k => dayAdd(k, -((dow(k) + 6) % 7));   // 月曜はじまり（クロジカの週）
function calDays() {
  if (CAL.mode === "week") { const a = weekStart(S.day); return Array.from({ length: 7 }, (_, i) => dayAdd(a, i)); }
  if (CAL.mode === "day") return [S.day];
  if (CAL.mode === "month") { const a = dayAdd(monthOf(S.day) + "-01", -dow(monthOf(S.day) + "-01")), out = []; for (let k = a; out.length < 42; k = dayAdd(k, 1)) { if (out.length >= 35 && monthOf(k) !== monthOf(S.day) && dow(k) === 0) break; out.push(k); } return out; }
  return monthDays(monthOf(S.day));
}
function calStep(n) {
  if (CAL.mode === "week") return dayAdd(S.day, 7 * n);
  if (CAL.mode === "day") return dayAdd(S.day, n);
  const d = dateOf(monthOf(S.day) + "-01"); d.setMonth(d.getMonth() + n); return keyOf(d);
}

/* ---------- 全員の日報を集める ---------- */
async function docsFor(days) {
  const ms = [...new Set(days.map(monthOf))];
  await Promise.all(ms.map(m => loadMonth(m)));
  return (await Promise.all(ms.map(m => monthDocs(m).catch(() => [])))).flat();
}
// その日の行（だれが書いたか・報告したか）
function evsOn(docs, k) {
  const out = [];
  docs.forEach(doc => { if (monthOf(k) !== doc.月 && doc.月) return; const d = (doc.日 || {})[k]; (d && d.行 || []).forEach(r => out.push({ k, who: doc.名前, r, rep: !!d.報告, mine: nk(doc.名前) === nk(S.me.name) })); });
  return out;
}
// 行に出てくる人（記入者が参加していれば記入者・報告の参加者。報告が無ければ予定の参加者）
function membersOf(x) {
  const m = new Map(), add = n => { if (n && !m.has(nk(n))) m.set(nk(n), n); };
  if (x.r.本人 !== false) add(x.who);
  const rep = listIn(x.r, "r", "人"); (rep.length ? rep : listIn(x.r, "p", "人")).forEach(add);
  return [...m.values()];
}
const titleOf = r => (r.行先 || "（行先なし）") + (r.関連現場 && r.関連現場.名前 ? `（${r.関連現場.名前}）` : "");
const timed = r => !r.終日 && r.始め && r.終わり && hoursOf(r) < 6;   // 短い用事は時刻つきで（クロジカの時間の予定）
function evTip(x) {
  const r = x.r, L = workLines(r).filter(l => l.予定 || l.報告);
  return [titleOf(r) + (r.始め || r.終わり ? `  ${r.始め || ""}〜${r.終わり || ""}` : ""), "記入者：" + x.who + (x.rep ? "" : "（書きかけ）"),
    membersOf(x).length ? "参加者：" + membersOf(x).join("、") : "", gaiTxt(listIn(r, "r", "外注")) ? "外注：" + gaiTxt(listIn(r, "r", "外注")) : "",
    listIn(r, "r", "機械").length ? "設備：" + listIn(r, "r", "機械").join("、") : listIn(r, "p", "機械").length ? "設備（予定）：" + listIn(r, "p", "機械").join("、") : "",
    ...L.map(l => "・" + (l.報告 || "（予定）" + l.予定))].filter(Boolean).join("\n");
}
// 全体工程表のその日の予定（現場ごと）
function plansOn(k) {
  const out = [];
  S.K.sites.forEach(s => { const L = planLines(s.名前, s.コード, k); if (L.length) out.push({ name: s.名前, code: s.コード || "", lines: L }); });
  return out;
}
// from：その線を出している人の行（週・日）。自分が書いた線は動かせる（ドラッグで移す・Ctrl で写す・右端で続きの日を増やす／減らす）
const evHTML = (x, color, from) => {
  const r = x.r, t = timed(r), mv = x.mine && CAL.mode !== "person" && CAL.mode !== "site";
  return `<div class="ev${t ? " tm" : ""}${x.rep ? "" : " draft"}${mv ? " mv" : ""}" data-ev="${esc(x.who)}|${x.k}|${r.id}" data-from="${esc(from || membersOf(x)[0] || x.who)}"${mv ? ` draggable="true"` : ""} style="--c:${color}" title="${esc(evTip(x))}">` +
    (mv && from ? `<span class="rz" title="右へ引くと次の日も同じ予定（減らすときは左へ）"></span>` : "") +
    (t ? `<small>${esc(r.始め)} - ${esc(r.終わり)}</small><br>` : "") + (kubunOf(r) === "夜" ? "🌙" : "") + esc(titleOf(r)) + `</div>`;
};

/* ---------- カレンダーを描く ---------- */
async function renderCal() {
  $("cards").hidden = true; $("empty").hidden = true; $("all").hidden = false;
  const days = calDays(), d0 = dateOf(days[0]), d1 = dateOf(days[days.length - 1]);
  const ttl = CAL.mode === "week" ? `${d0.getFullYear()}年${d0.getMonth() + 1}月${d0.getDate()}日〜${d1.getMonth() + 1}月${d1.getDate()}日`
    : CAL.mode === "day" ? `${d0.getFullYear()}年${d0.getMonth() + 1}月${d0.getDate()}日(${WD[d0.getDay()]})` : `${dateOf(S.day).getFullYear()}年${+monthOf(S.day).slice(5)}月`;
  $("dayHead").innerHTML = `<div class="calbar"><div class="tabs">${MODES.map(([v, t]) => `<button type="button" data-cm="${v}" class="${CAL.mode === v ? "on" : ""}">${t}</button>`).join("")}</div>
    <b class="cttl">${ttl}</b><span class="sp"></span><button type="button" id="cKuro" title="クロジカの CSV から過去の予定・報告を取り込む">クロジカから取り込む</button><button type="button" id="cWho">表示する人</button><button type="button" id="cCards" title="スマホと同じカードで書く">私の日報（カード）</button></div>`;
  $("dayHead").querySelectorAll("[data-cm]").forEach(b => b.onclick = () => { CAL.mode = b.dataset.cm; calSave(); render(); });
  $("cWho").onclick = chooseWho;
  $("cKuro").onclick = kuroOpen;
  $("cCards").onclick = () => { S.view = "me"; render(); };
  if (!$("all").querySelector(".cal")) $("all").innerHTML = `<p>読んでいます…</p>`;
  let docs;
  try { docs = await docsFor(days); } catch (e) { $("all").innerHTML = `<p class="ng">日報を読めませんでした：${esc(e.message)}</p>`; return; }
  if (S.view !== "cal" || calDays().join() !== days.join()) return;
  const evs = new Map(days.map(k => [k, evsOn(docs, k)]));
  const narrow = window.innerWidth < 760;   // スマホ：週・日は縦の一覧（クロジカのスマホと同じく、押すと詳細）
  const html = narrow && (CAL.mode === "week" || CAL.mode === "day") ? listHTML(days, evs) : CAL.mode === "week" ? weekHTML(days, evs) : CAL.mode === "day" ? dayHTML(days[0], evs.get(days[0])) : CAL.mode === "month" ? monthHTML(days, evs)
    : CAL.mode === "person" ? personHTML(days, evs) : siteHTML(days, evs);
  $("all").innerHTML = `<div class="cal ${CAL.mode}">${html}</div>`;
  bindCal(evs);
}
// 表に出す人：社員と、この期間に日報に出てくる人（「表示する人」で外した人は出さない）
function calNames(evs) {
  const m = new Map(), add = n => { if (n && !m.has(nk(n))) m.set(nk(n), n); };
  S.M.社員.forEach(add); add(S.me.name);
  [...evs.values()].flat().forEach(x => membersOf(x).forEach(add));
  return [...m.values()].filter(n => !CAL.hide[nk(n)]);
}
const dayTh = k => `<th class="${dayCls(k)}${k === todayKey() ? " today" : ""}" data-day="${k}"><b>${+k.slice(8)}</b> <small>${WD[dow(k)]}</small>${holidayOf(k) ? `<br><small>${esc(holidayOf(k))}</small>` : ""}</th>`;
const planCell = k => plansOn(k).map(p => `<div class="pv" draggable="true" data-plan="${k}|${esc(p.name)}|${esc(p.code)}" title="${esc(p.name + "\n" + p.lines.join("\n"))}">${esc(p.name)}</div>`).join("");
function weekHTML(days, evs) {
  const names = calNames(evs);
  return `<div class="cwrap"><table class="cg"><thead><tr><th class="nm"></th>${days.map(dayTh).join("")}</tr></thead><tbody>` +
    (S.K.sites.length ? `<tr class="plan"><th class="nm">全体工程<br><small>（予定）</small></th>${days.map(k => `<td class="${dayCls(k)}">${planCell(k)}</td>`).join("")}</tr>` : "") +
    names.map(n => `<tr><th class="nm"><span class="dot" style="background:${colorOf(n)}"></span>${esc(n)}</th>` +
      days.map(k => `<td class="${dayCls(k)}" data-new="${k}|${esc(n)}">${evs.get(k).filter(x => membersOf(x).some(v => nk(v) === nk(n))).map(x => evHTML(x, colorOf(n), n)).join("")}</td>`).join("") + `</tr>`).join("") +
    `</tbody></table></div><p class="hint">空いたます目を押すと、その人・その日の日報を書けます。上の「全体工程」は押すとその現場で書き始め、人のます目へ引くとその人の予定に入ります（全体工程の欄は残る）。
    自分が書いた線は、引くと移し、Ctrl を押しながら引くと写します。線の右端を右へ引くと次の日も同じ予定（左へ引くと減らす）。入れたあと1日ずつ押して直せます。</p>`;
}
// スマホの週・日：日ごとに縦に並べる（時刻順。色は記入者、参加者を下に）。押すと詳細、＋でその日に書く、全体工程の現場を押すとその現場で書く
function listHTML(days, evs) {
  return `<div class="mlist2">` + days.map(k => {
    const d = dateOf(k), list = evs.get(k).filter(x => membersOf(x).some(n => !CAL.hide[nk(n)])).sort((a, b) => (a.r.終日 || !a.r.始め) - (b.r.終日 || !b.r.始め) || (tmin(a.r.始め) ?? 0) - (tmin(b.r.始め) ?? 0));
    const pl = S.K.sites.length ? plansOn(k) : [];
    return `<section class="mday ${dayCls(k)}${k === todayKey() ? " today" : ""}"><div class="mdh"><a data-day="${k}"><b>${d.getMonth() + 1}/${d.getDate()}</b>（${WD[d.getDay()]}）${holidayOf(k) ? `<small>${esc(holidayOf(k))}</small>` : ""}</a><span class="sp"></span><button type="button" class="madd" data-new="${k}|${esc(S.me.name)}" aria-label="この日に書く">＋</button></div>` +
      (pl.length ? `<div class="mpl"><small>全体工程</small>${pl.map(p => `<span class="pv" data-plan="${k}|${esc(p.name)}|${esc(p.code)}">${esc(p.name)}</span>`).join("")}</div>` : "") +
      (list.length ? list.map(x => { const r = x.r, mem = membersOf(x);
        return `<div class="mev${x.rep ? "" : " draft"}" data-ev="${esc(x.who)}|${k}|${r.id}" style="--c:${colorOf(x.r.本人 === false ? mem[0] || x.who : x.who)}"><span class="mt">${r.終日 || !r.始め ? "終日" : esc(`${r.始め}-${r.終わり || ""}`)}${kubunOf(r) === "夜" ? " 🌙" : ""}</span>` +
          `<span class="mb"><b>${esc(titleOf(r))}</b><small>${esc(mem.join("、"))}${x.rep ? "" : "（書きかけ）"}</small></span></div>`; }).join("") : `<p class="none">予定・報告はありません</p>`) + `</section>`;
  }).join("") + `</div>`;
}
function dayHTML(k, list) {
  const names = calNames(new Map([[k, list]]));
  return `<div class="cwrap"><table class="cg dayv"><thead><tr>${S.K.sites.length ? `<th>全体工程（予定）</th>` : ""}${names.map(n => `<th><span class="dot" style="background:${colorOf(n)}"></span>${esc(n)}</th>`).join("")}</tr></thead><tbody><tr>` +
    (S.K.sites.length ? `<td class="plan">${plansOn(k).map(p => `<div class="pv" draggable="true" data-plan="${k}|${esc(p.name)}|${esc(p.code)}"><b>${esc(p.name)}</b>${p.lines.map(l => `<br><small>${esc(l)}</small>`).join("")}</div>`).join("") || `<small class="none">なし</small>`}</td>` : "") +
    names.map(n => `<td data-new="${k}|${esc(n)}">${list.filter(x => membersOf(x).some(v => nk(v) === nk(n))).sort((a, b) => (tmin(a.r.始め) ?? 0) - (tmin(b.r.始め) ?? 0)).map(x => evHTML({ ...x, r: { ...x.r, 終日: false } }, colorOf(n), n).replace("<div class=\"ev", "<div class=\"ev big")).join("")}</td>`).join("") + `</tr></tbody></table></div>`;
}
function monthHTML(days, evs) {
  const m = monthOf(S.day);
  return `<table class="mg"><thead><tr>${[..."日月火水木金土"].map((w, i) => `<th class="${i === 0 ? "sun" : i === 6 ? "sat" : ""}">${w}</th>`).join("")}</tr></thead><tbody>` +
    Array.from({ length: days.length / 7 }, (_, w) => `<tr>${days.slice(w * 7, w * 7 + 7).map(k => {
      const list = evs.get(k).filter(x => membersOf(x).some(n => !CAL.hide[nk(n)]));
      return `<td class="${dayCls(k)}${monthOf(k) !== m ? " out" : ""}${k === todayKey() ? " today" : ""}" data-new="${k}|"><a class="dn" data-day="${k}">${+k.slice(8)}</a>${holidayOf(k) ? `<small class="hol">${esc(holidayOf(k))}</small>` : ""}` +
        list.map(x => evHTML(x, colorOf(x.r.本人 === false ? membersOf(x)[0] || x.who : x.who)).replace(/>([^<]*)<\/div>$/, `>${esc(x.r.本人 === false ? (membersOf(x)[0] || x.who) : x.who)}：$1</div>`)).join("") + `</td>`;
    }).join("")}</tr>`).join("") + `</tbody></table>`;
}
// 人別：人×日のます目に行先（短く）。日数・時間
function personHTML(days, evs) {
  const names = calNames(evs);
  const rows = names.map(n => {
    const on = k => evs.get(k).filter(x => membersOf(x).some(v => nk(v) === nk(n)));
    let nd = 0, h = 0;
    days.forEach(k => { const l = on(k); if (l.length) nd++; l.forEach(x => h += hoursOf(x.r)); });
    return { n, on, nd, h };
  }).filter(o => o.nd);
  if (!rows.length) return `<p class="none">この月はまだだれも書いていません。</p>`;
  return `<h3>${+monthOf(S.day).slice(5)}月の人別<small>ます目は行った現場（押すとその日の日報）。時間は記入者が書いた始め〜終わり</small></h3><div class="gwrap"><table class="grid pg"><thead><tr><th></th>${days.map(k => `<th class="${dayCls(k)}">${+k.slice(8)}<br><small>${WD[dow(k)]}</small></th>`).join("")}<th>日数</th><th>時間</th></tr></thead><tbody>` +
    rows.map(o => `<tr><th><span class="dot" style="background:${colorOf(o.n)}"></span>${esc(o.n)}</th>${days.map(k => { const l = o.on(k), t = [...new Set(l.map(x => titleOf(x.r)))];
      return `<td class="${dayCls(k)}"${l.length ? ` data-pd="${k}" title="${esc(t.join("\n"))}"` : ""}>${esc(t.length ? t[0].slice(0, 4) + (t.length > 1 ? "他" : "") : "")}</td>`; }).join("")}<td class="tot">${o.nd}</td><td class="tot">${fmtH(o.h)}</td></tr>`).join("") + `</tbody></table></div>`;
}
// 現場別：現場×日の人数（参加者＋外注の人数）・人日。押すとその現場の現場別日報（別ウィンドウ）
function siteHTML(days, evs) {
  const by = new Map();
  days.forEach(k => evs.get(k).forEach(x => {
    if (!x.r.行先 || isOffice(x.r.行先)) return;
    const st = siteOf(x.r.行先, x.r.コード), key = st ? "c" + st.コード : "n" + nk(x.r.行先);
    if (!by.has(key)) by.set(key, { place: st ? siteName(st) : x.r.行先, code: st ? st.コード : x.r.コード || "", d: {} });
    const o = by.get(key).d, mem = new Set((o[k] || { m: new Set() }).m); membersOf(x).forEach(n => mem.add(nk(n)));
    o[k] = { m: mem, g: ((o[k] || {}).g || 0) + listIn(x.r, "r", "外注").reduce((a, g) => a + (parseFloat(g.人数) || 0), 0) };
  }));
  if (!by.size) return `<p class="none">この月はまだ現場の日報がありません。</p>`;
  const list = [...by.values()].sort((a, b) => a.place.localeCompare(b.place, "ja"));
  return `<h3>${+monthOf(S.day).slice(5)}月の現場別<small>ます目＝その日の人数（社員・作業員＋外注）。現場名を押すと現場別日報（別ウィンドウ）</small></h3><div class="gwrap"><table class="grid sg"><thead><tr><th></th><th>形態</th>${days.map(k => `<th class="${dayCls(k)}">${+k.slice(8)}<br><small>${WD[dow(k)]}</small></th>`).join("")}<th>人日</th></tr></thead><tbody>` +
    list.map(s => { let tot = 0; const cells = days.map(k => { const c = s.d[k]; const n = c ? c.m.size + c.g : 0; tot += n; return `<td class="${dayCls(k)}">${n || ""}</td>`; }).join("");
      return `<tr><th><a data-site="${esc(s.place)}|${esc(s.code)}">${esc(s.place)}</a></th><td>${esc(formOf(s.place, s.code) || "―")}</td>${cells}<td class="tot">${tot}</td></tr>`; }).join("") + `</tbody></table></div>`;
}
function bindCal(evs) {
  const box = $("all");
  box.onclick = e => {
    const ev = e.target.closest("[data-ev]");
    if (ev) { const [who, k, id] = ev.dataset.ev.split("|"), x = evs.get(k).find(y => nk(y.who) === nk(who) && y.r.id === id); if (x) showEv(x); return; }
    const pv = e.target.closest("[data-plan]");
    if (pv) { const [k, place, code] = pv.dataset.plan.split("|"); return editEv(k, null, { place, code }); }
    const dd = e.target.closest("[data-day]");
    if (dd && CAL.mode !== "day") { S.day = dd.dataset.day; CAL.mode = "day"; calSave(); return render(); }
    const pd = e.target.closest("[data-pd]");
    if (pd) { CAL.mode = "day"; calSave(); return setDay(pd.dataset.pd); }
    const st = e.target.closest("[data-site]");
    if (st) { const [p, c] = st.dataset.site.split("|"); return openSite(p, c); }
    const nw = e.target.closest("[data-new]");
    if (nw) { const [k, n] = nw.dataset.new.split("|"); editEv(k, null, { person: n }); }
  };
  bindDrag(box);
}

/* ---------- ドラッグ（PC のマウス）：全体工程の予定を人のます目へ（予定の欄は残る）・自分の線を移す／Ctrl で写す・右端で続きの日を増やす／減らす ---------- */
let DRAG = null, RZ = null;
const cellOf = el => { const td = el && el.closest && el.closest("td[data-new]"); if (!td) return null; const [k, n] = td.dataset.new.split("|"); return { td, k, n }; };
const whatOf = r => ({ 人: membersOf({ who: S.me.name, r }), 機械: listIn(r, "r", "機械").length ? listIn(r, "r", "機械") : listIn(r, "p", "機械") });
async function putRow(k, r) { await loadMonth(monthOf(k)); mutate(k, d => { const i = d.行.findIndex(y => y.id === r.id); if (i < 0) d.行.push(r); else d.行[i] = r; }, true); }
async function delRow(k, id) { await loadMonth(monthOf(k)); mutate(k, d => { d.行 = (d.行 || []).filter(y => y.id !== id); if (!d.行.length) d.報告 = null; }, true); }
// 人を入れ替える（from の行から to の行へ）。記入者（自分）は 本人、ほかの人は参加者（報告が無ければ予定）
function swapMember(y, from, to) {
  const me = n => nk(n) === nk(S.me.name), rep = listOf(y, "r", "人"), pl = listOf(y, "p", "人");
  if (from && nk(from) === nk(to)) return;
  if (from) { if (me(from)) y.本人 = false; [rep, pl].forEach(a => { const i = a.findIndex(v => nk(v) === nk(from)); if (i >= 0) a.splice(i, 1); }); }
  if (me(to)) delete y.本人;
  else { const a = rep.length || !pl.length ? rep : pl; if (!a.some(v => nk(v) === nk(to))) a.push(to); }
}
// 写した線は予定にする（報告の人・外注・機械・作業を予定へ。報告は空に）。その日の全体工程があれば作業・機械はその日の予定に
function asPlan(y, k) {
  ["人", "機械"].forEach(x => { const p = listOf(y, "p", x); listIn(y, "r", x).forEach(v => { if (!p.some(w => nk(w) === nk(v))) p.push(v); }); y[x] = []; });
  const pg = listOf(y, "p", "外注"); listIn(y, "r", "外注").forEach(g => { if (!pg.some(v => nk(v.業者) === nk(g.業者))) pg.push({ ...g }); }); y.外注 = [];
  y.作業行 = workLines(y).map(l => ({ 予定: l.予定 || l.報告, 報告: "" }));
  const pl = planLines(y.行先, y.コード, k); if (pl.length) y.作業行 = pl.map(t => ({ 予定: t, 報告: "" }));
  const pk = planKikai(y.行先, y.コード, k); if (pk.length) y.予定.機械 = pk;
  delete y.未使用機械; delete y.更新; syncWork(y);
  return y;
}
const nightPlan = (place, code, k) => planLines(place, code, k).some(t => t.includes("（夜）"));
async function dropOn(D, k, n, copy) {
  if (D.t === "pv") {   // 全体工程の予定 → その人・その日の予定
    const st = siteOf(D.place, D.code), y = { id: uid(), 行先: st ? siteName(st) : D.place, コード: D.code || (st && st.コード) || "", 作業: "", 作業行: [{ 予定: "", 報告: "" }], 予定: { 人: [], 外注: [], 機械: [] }, 人: [], 外注: [], 機械: [] };
    const night = nightPlan(y.行先, y.コード, k); y.区分 = night ? "夜" : "昼"; y.始め = night ? "21:00" : "8:00"; y.終わり = night ? "6:00" : "17:00";
    fillPlan(y, k); swapMember(y, "", n || S.me.name); y.記入者 = S.me.name;
    if (!await okClash(k, y, whatOf(y))) return;
    await putRow(k, y); render(); toast(`${n || S.me.name} の ${+k.slice(5, 7)}/${+k.slice(8)} に「${y.行先}」の予定を入れました`); return;
  }
  await loadMonth(monthOf(D.k));
  const src = rowsOf(D.k).find(y => y.id === D.id); if (!src) return;
  const to = n || D.from;
  if (!copy && D.k === k && nk(D.from) === nk(to)) return;
  const y = JSON.parse(JSON.stringify(src));
  if (copy) { asPlan(y, k); y.id = uid(); delete y.連; } else if (D.k !== k) delete y.連;
  swapMember(y, D.from, to); y.記入者 = S.me.name; y.更新 = nowStr();
  if (!await okClash(k, y, whatOf(y))) return;
  if (!copy) await delRow(D.k, src.id);
  await putRow(k, y); render(); toast(copy ? "写しました（あとで1日ずつ直せます）" : "移しました");
}
// 右端を引いた：src の日から t の日まで同じ予定を足す（日曜・祝日は飛ばす）／t より後のつながった予定を消す
async function stretch(D, t) {
  if (t === D.k) return;
  await loadMonth(monthOf(D.k));
  const src = rowsOf(D.k).find(y => y.id === D.id); if (!src) return;
  const chain = src.連 || src.id, inChain = y => y.id === src.id || (y.連 && y.連 === chain);
  if (t > D.k) {
    const days = []; for (let d = dayAdd(D.k, 1); d <= t; d = dayAdd(d, 1)) if (dow(d) !== 0 && !holidayOf(d)) days.push(d);
    await Promise.all([...new Set(days.map(monthOf))].map(m => loadMonth(m)));
    const add = days.filter(d => !rowsOf(d).some(inChain)).map(d => { const y = asPlan(JSON.parse(JSON.stringify(src)), d); y.id = uid(); y.連 = chain; y.記入者 = S.me.name; return [d, y]; });
    if (!add.length) return;
    const msg = []; for (const [d, y] of add) (await clashes(d, y, whatOf(y))).forEach(m => msg.push(`${+d.slice(5, 7)}/${+d.slice(8)} ${m.slice(1)}`));
    if (msg.length && !confirm(`同じ時間帯に、ほかでも入っています。\n\n${msg.join("\n")}\n\nこのままでよいですか？`)) return;
    if (!src.連) mutate(D.k, d => { const y = findRow(d, src.id); if (y) y.連 = chain; }, true);
    for (const [d, y] of add) await putRow(d, y);
    render(); toast(`${add.length}日分、同じ予定を入れました（日曜・祝日は飛ばす。あとで1日ずつ直せます）`);
  } else {
    const days = []; for (let d = dayAdd(t, 1); d <= D.k; d = dayAdd(d, 1)) days.push(d);
    await Promise.all([...new Set(days.map(monthOf))].map(m => loadMonth(m)));
    const del = days.flatMap(d => rowsOf(d).filter(inChain).map(y => [d, y]));
    if (!del.length) return;
    if (!confirm(`${del.map(([d, y]) => `${+d.slice(5, 7)}/${+d.slice(8)}「${titleOf(y)}」`).join("、")} を消しますか？`)) return;
    for (const [d, y] of del) await delRow(d, y.id);
    render(); toast(`${del.length}日分を消しました`);
  }
}
function bindDrag(box) {
  const clear = () => box.querySelectorAll(".dropok,.rzsel").forEach(e => e.classList.remove("dropok", "rzsel"));
  box.ondragstart = e => {
    if (RZ) { e.preventDefault(); return; }
    const pv = e.target.closest("[data-plan]"), ev = e.target.closest(".ev.mv");
    if (pv) { const [k, place, code] = pv.dataset.plan.split("|"); DRAG = { t: "pv", k, place, code }; e.dataTransfer.effectAllowed = "copy"; }
    else if (ev) { const [, k, id] = ev.dataset.ev.split("|"); DRAG = { t: "ev", k, id, from: ev.dataset.from }; e.dataTransfer.effectAllowed = "copyMove"; }
    else return;
    e.dataTransfer.setData("text/plain", "nippou");
  };
  box.ondragover = e => { const c = DRAG && cellOf(e.target); if (!c) return; e.preventDefault(); e.dataTransfer.dropEffect = DRAG.t === "pv" || e.ctrlKey ? "copy" : "move"; clear(); c.td.classList.add("dropok"); };
  box.ondrop = e => { const c = DRAG && cellOf(e.target); clear(); if (!c) return; e.preventDefault(); const D = DRAG; DRAG = null; dropOn(D, c.k, c.n, e.ctrlKey).catch(err => toast("できませんでした：" + err.message)); };
  box.ondragend = () => { DRAG = null; clear(); };
  // 右端（.rz）を引く：同じ人の行で、引いた先の日までを塗って見せる
  box.onpointerdown = e => {
    const rz = e.target.closest(".rz"); if (!rz) return;
    e.preventDefault(); e.stopPropagation();
    const ev = rz.closest(".ev"), [, k, id] = ev.dataset.ev.split("|"), row = ev.closest("tr");
    RZ = { k, id, from: ev.dataset.from, t: k };
    const move = m => { const c = cellOf(document.elementFromPoint(m.clientX, m.clientY)); if (!c || c.td.closest("tr") !== row) return; RZ.t = c.k; clear();
      row.querySelectorAll("td[data-new]").forEach(td => { const d = td.dataset.new.split("|")[0]; if ((d > k && d <= c.k) || (d <= k && d > c.k)) td.classList.add("rzsel"); }); };
    const up = () => { document.removeEventListener("pointermove", move); document.removeEventListener("pointerup", up); clear(); const R = RZ; setTimeout(() => { RZ = null; }, 0); stretch(R, R.t).catch(err => toast("できませんでした：" + err.message)); };
    document.addEventListener("pointermove", move); document.addEventListener("pointerup", up);
  };
  box.addEventListener("click", e => { if (e.target.closest(".rz")) e.stopImmediatePropagation(); }, true);
}
function chooseWho() {
  const all = [...new Set([...S.M.社員, ...S.M.従業員])];
  dlg(`<h2>表示する人</h2><p style="font-size:14px">社員と、その週・月に日報に出てくる人を出します。外した人は出しません（この端末だけ）。</p>
    <div class="btns" style="justify-content:flex-start;margin:0 0 6px"><button type="button" id="wAll">全部</button><button type="button" id="wNone">全部外す</button></div>
    <div class="checks">${all.map(n => `<label><input type="checkbox" value="${esc(n)}"${CAL.hide[nk(n)] ? "" : " checked"}><span class="dot" style="background:${colorOf(n)}"></span>${esc(n)}</label>`).join("")}</div>
    <div class="btns"><button data-close>やめる</button><button class="pri" id="wOk">決める</button></div>`);
  const cbs = () => $("dlgBody").querySelectorAll(".checks input");
  $("wAll").onclick = () => cbs().forEach(c => c.checked = true);
  $("wNone").onclick = () => cbs().forEach(c => c.checked = false);
  $("wOk").onclick = () => { CAL.hide = {}; cbs().forEach(c => { if (!c.checked) CAL.hide[nk(c.value)] = 1; }); calSave(); closeDlg(); render(); };
}

/* ---------- 請負形態（常用／請負）：日報のフォルダの 現場設定.json ---------- */
const SET_FILE = "現場設定.json";
let SETS = { doc: { 形式: "工事日報の現場設定", 版: 1, 現場: {} }, eTag: null };
const setKey = (place, code) => code ? "c:" + code : "n:" + nk(place);
async function loadSets() { try { const r = await nippouStore.get(SET_FILE); if (r) SETS = { doc: r.data, eTag: r.eTag }; } catch (e) { } }
function formOf(place, code) {
  const st = siteOf(place, code), c = code || (st && st.コード) || "", M = SETS.doc.現場 || {};
  return ((c && M[setKey("", c)]) || M[setKey(place, "")] || {}).形態 || "";
}
async function setForm(place, code, v) {
  const st = siteOf(place, code), c = code || (st && st.コード) || "", key = setKey(place, c);
  const fn = doc => { doc.現場 ||= {}; if (v) doc.現場[key] = { 形態: v, 名前: st ? siteName(st) : place, だれ: S.me.name, 更新: nowStr() }; else delete doc.現場[key]; };
  for (let i = 0; i < 4; i++) {
    const cur = await nippouStore.get(SET_FILE).catch(() => null), doc = cur ? cur.data : { 形式: "工事日報の現場設定", 版: 1, 現場: {} };
    fn(doc);
    try { const r = await nippouStore.put(SET_FILE, doc, cur ? cur.eTag : null); SETS = { doc, eTag: r.eTag }; toast(`請負形態を「${v || "未設定"}」にしました`); return; }
    catch (e) { if (e.status !== 412 && e.status !== 409) { toast("請負形態を保存できませんでした：" + e.message); return; } }
  }
}
const formSel = (place, code, id) => `<select id="${id}">${["", "常用", "請負"].map(v => `<option value="${v}"${formOf(place, code) === v ? " selected" : ""}>${v || "未設定"}</option>`).join("")}</select>`;

/* ---------- スケジュール詳細 ---------- */
const fmtDay = k => { const d = dateOf(k); return `${d.getFullYear()} 年 ${d.getMonth() + 1} 月 ${d.getDate()} 日 (${WD[d.getDay()]})`; };
function siteInfo(place, code) {
  const st = place && !isOffice(place) ? siteOf(place, code) : null;
  if (!st) return place && !isOffice(place) ? "現場一覧にない現場" : "";
  return `工事コード ${esc(st.コード || "―")}　` + (st.元請 === "自社請け" ? `自社請け　発注者 ${esc(st.発注者 || "―")}` : `元請 ${esc(st.元請 || "―")}${st.発注者 ? `　発注者 ${esc(st.発注者)}` : ""}`);
}
function showEv(x) {
  const r = x.r, L = workLines(r).filter(l => l.予定 || l.報告), two = (k, f) => {
    const p = f(listIn(r, "p", k)), q = f(listIn(r, "r", k));
    return p || q ? `<div class="pr2"><div><small>予定</small>${p || "―"}</div><div><small>報告</small>${q || "―"}</div></div>` : "―";
  };
  const chipsT = a => a.map(n => `<span class="chip ro">${esc(n)}</span>`).join("");
  const mem = [...(r.本人 !== false ? [x.who] : []), ...listIn(r, "r", "人")];
  dlg(`<div class="dh"><h2>スケジュール詳細</h2><button type="button" data-close aria-label="閉じる">×</button></div><table class="ft">
    <tr><th>日時</th><td>${fmtDay(x.k)}　${kubunOf(r)}勤　${r.終日 || !(r.始め || r.終わり) ? "終日" : esc(`${r.始め || ""} - ${r.終わり || ""}`)}</td></tr>
    <tr><th>現場</th><td><b>${esc(titleOf(r))}</b><div class="pinfo">${siteInfo(r.行先, r.コード)}</div></td></tr>
    ${r.行先 && !isOffice(r.行先) ? `<tr><th>請負形態</th><td>${formSel(r.行先, r.コード, "sForm")}</td></tr>` : ""}
    <tr><th>内容</th><td>${L.length ? `<table class="wt"><tr><th>予定</th><th>報告</th></tr>${L.map(l => `<tr><td>${esc(l.予定 || "")}</td><td>${esc(l.報告 || "")}</td></tr>`).join("")}</table>` : "―"}</td></tr>
    <tr><th>参加者</th><td><div class="pr2"><div><small>予定</small>${chipsT(listIn(r, "p", "人")) || "―"}</div><div><small>報告</small>${chipsT(mem) || "―"}</div></div></td></tr>
    <tr><th>外注</th><td>${two("外注", a => esc(gaiTxt(a)))}</td></tr>
    <tr><th>設備</th><td>${two("機械", chipsT)}${(r.未使用機械 || []).length ? `<div class="pinfo">使わなかった：${esc(r.未使用機械.join("、"))}</div>` : ""}</td></tr>
    <tr><th>記入者</th><td>${esc(x.who)}${r.更新 ? `（${esc(r.更新)}）` : ""}${x.rep ? "" : "　<span class=\"badge draft\">書きかけ</span>"}</td></tr>
  </table><div class="btns"><button data-close>閉じる</button>${x.mine ? `<button id="sEdit">編集する</button><button id="sDel">削除する</button>` : ""}<button id="sCopy">コピーして登録する</button></div>`, "wide");
  if ($("sForm")) $("sForm").onchange = () => setForm(r.行先, r.コード, $("sForm").value);
  if ($("sEdit")) $("sEdit").onclick = () => editEv(x.k, r);
  if ($("sDel")) $("sDel").onclick = async () => {
    if (!confirm(`「${titleOf(r)}」（${fmtDay(x.k)}）を削除しますか？`)) return;
    mutate(x.k, d => { d.行 = (d.行 || []).filter(y => y.id !== r.id); if (!d.行.length) d.報告 = null; }, true);
    closeDlg(); render();
  };
  $("sCopy").onclick = () => { const c = JSON.parse(JSON.stringify(r)); c.id = uid(); delete c.更新; if (!x.mine) c.本人 = true; editEv(x.k, null, { copy: c }); };
}

/* ---------- スケジュール編集（登録） ----------
   o.person：その人のます目から（自分でなければ代理）。o.place/code：全体工程の予定から。o.copy：コピーして登録 */
function editEv(k, r0, o = {}) {
  const isNew = !r0;
  const E = r0 ? JSON.parse(JSON.stringify(r0)) : o.copy || { id: uid(), 行先: "", コード: "", 始め: "", 終わり: "", 終日: false, 作業: "", 作業行: [{ 予定: "", 報告: "" }], 予定: { 人: [], 外注: [], 機械: [] }, 人: [], 外注: [], 機械: [] };
  let day = k;
  if (isNew && !o.copy && o.person && nk(o.person) !== nk(S.me.name)) { E.本人 = false; E.人 = [o.person]; }
  if (isNew && !o.copy && o.place) { const st = siteOf(o.place, o.code); E.行先 = st ? siteName(st) : o.place; E.コード = o.code || (st && st.コード) || ""; fillPlan(E, day); }
  const kikaiOn = () => !!S.kikai;
  // 現場は名前で絞る（打つと候補。候補の右に工事コード・元請／自社請けは発注者。発注者名で打っても出る）
  const items = placeItems;
  // 設備：稼働表のこの日の予約（ほかの現場に予約がある機械は ⚠）
  const elsewhere = n => { if (!kikaiOn()) return ""; const names = kikaiNames(E.行先, E.コード), mc = (S.kikai.機械 || []).find(m => nk(machName(m)) === nk(n)); if (!mc) return "";
    const y = (S.kikai.予約 || []).find(y => y.機械 === mc.id && y.状態 !== "仮" && kikaiWorks(y, day) && !names.includes(nk(y.現場)));
    return y ? `⚠ ${n} はこの日「${y.現場}」に予約あり` : ""; };
  const draw = () => {
    const st = E.行先 && !isOffice(E.行先) ? siteOf(E.行先, E.コード) : null, L = workLines(E);
    const planM = listIn(E, "p", "機械").filter(n => !listIn(E, "r", "機械").some(v => nk(v) === nk(n)));
    const warn = listIn(E, "r", "機械").map(elsewhere).filter(Boolean);
    dlg(`<div class="dh"><h2>${isNew ? "スケジュール登録" : "スケジュール編集"}</h2><button type="button" data-close aria-label="閉じる">×</button></div><table class="ft">
      <tr><th>日時</th><td><div class="dt"><input type="date" id="eDay" value="${day}"><label class="ck"><input type="checkbox" id="eAll"${E.終日 ? " checked" : ""}> 終日</label>
        ${kubunBox(E)}<span class="time"${E.終日 ? " hidden" : ""}><input class="tm" id="eT0" value="${esc(E.始め || "")}" placeholder="始め" inputmode="decimal"><span>〜</span><input class="tm" id="eT1" value="${esc(E.終わり || "")}" placeholder="終わり" inputmode="decimal"></span></div></td></tr>
      <tr><th>現場<span class="req">*</span></th><td><div class="pick"><input id="ePlace" value="${esc(E.行先 || "")}" placeholder="報告する現場（現場名・発注者・コードを打つと候補。本社も）" autocomplete="off"></div><div class="pinfo">${siteInfo(E.行先, E.コード)}</div>${honshaBox(E)}</td></tr>
      ${st ? `<tr><th>請負形態</th><td>${formSel(E.行先, E.コード, "eForm")} <small class="pinfo">現場ごと（常用の現場は常用日報の形で集計）</small></td></tr>` : ""}
      <tr><th>内容</th><td><div class="wl"><div class="wh"><span>予定</span><span>報告（したこと）</span><span></span></div>${L.map((l, j) => `<div class="wr"><textarea data-wj="${j}" data-wk="予定" rows="1" placeholder="予定">${esc(l.予定 || "")}</textarea><textarea data-wj="${j}" data-wk="報告" rows="1" placeholder="したこと">${esc(l.報告 || "")}</textarea>${L.length > 1 ? `<button type="button" data-delw="${j}" aria-label="消す">×</button>` : "<span></span>"}</div>`).join("")}</div>
        <button type="button" class="addline" data-act="addw">＋ 作業内容を追加</button></td></tr>
      <tr><th>参加者</th><td><div class="pr2"><div><small>予定</small>${chipsBox(E, "p", "人")}</div><div><small>報告</small><label class="ck me"><input type="checkbox" id="eMe"${E.本人 !== false ? " checked" : ""}> 自分（${esc(S.me.name)}）も参加</label>${chipsBox(E, "r", "人")}</div></div></td></tr>
      <tr><th>外注</th><td><div class="pr2"><div><small>予定</small>${gaiBox(E, "p")}</div><div><small>報告</small>${gaiBox(E, "r")}</div></div></td></tr>
      <tr><th>設備</th><td><div class="pr2"><div><small>予定${kikaiOn() ? "（自社機械稼働表の予約）" : ""}</small>${chipsBox(E, "p", "機械")}</div><div><small>報告（使った機械）</small>${chipsBox(E, "r", "機械")}
        ${planM.length ? `<div class="unused">${planM.map(n => `<label class="ck"><input type="checkbox" data-unused="${esc(n)}"${(E.未使用機械 || []).some(v => nk(v) === nk(n)) ? " checked" : ""}> ${esc(n)} は使わなかった</label>`).join("")}</div>` : ""}
        ${warn.map(w => `<div class="warn">${esc(w)}</div>`).join("")}</div></div>
        ${kikaiOn() ? `<small class="pinfo">報告の自社機械は自社機械稼働表の予約に入ります（続いている予約は延ばす）。「使わなかった」は稼働表に申告します。</small>` : ""}</td></tr>
      <tr><th>記入者</th><td>${esc(S.me.name)}${E.本人 === false ? "（代理で書く）" : ""}</td></tr>
    </table><div class="btns"><button data-close>閉じる</button><button class="pri" id="eSave">${isNew ? "登録する" : "更新する"}</button></div>`, "wide keep");
    bindEdit();
  };
  const redraw = () => { const sc = $("dlgBody").scrollTop; draw(); $("dlgBody").scrollTop = sc; };
  const setPlaceE = (v, code) => { E.行先 = v; E.コード = code || ""; fillPlan(E, day); redraw(); };
  function bindEdit() {
    const B = $("dlgBody");
    $("eDay").onchange = () => { if ($("eDay").value) { day = $("eDay").value; redraw(); } };
    $("eAll").onchange = () => { E.終日 = $("eAll").checked; redraw(); };
    bindTime($("eT0"), () => "8:00", v => { E.始め = v; });
    bindTime($("eT1"), () => stepTime(E.始め, "h", 9) || "17:00", v => { E.終わり = v; });
    bindExtras(B, () => E, (fn, re) => { fn(E); if (re) redraw(); }, sel => setTimeout(() => { const p = $("dlgBody").querySelector(sel); if (p) p.focus(); }, 0));
    const pl = $("ePlace");
    bindPick(pl, items, x => { pl.value = x.v; setPlaceE(x.v, x.コード); });
    pl.addEventListener("change", () => { const v = pl.value.trim(); if (v !== E.行先) { const s = siteOf(v), kk = koteiSite(v); setPlaceE(v, s ? s.コード : kk ? kk.コード : ""); } });
    if ($("eForm")) $("eForm").onchange = () => setForm(E.行先, E.コード, $("eForm").value);
    $("eMe").onchange = () => { E.本人 = $("eMe").checked ? undefined : false; };
    B.querySelectorAll("[data-wj]").forEach(el => { grow(el); el.addEventListener("input", () => { grow(el); const j = +el.dataset.wj; E.作業行 = workLines(E).map(l => ({ ...l })); E.作業行[j][el.dataset.wk] = el.value; }); });
    B.querySelectorAll("[data-ci]").forEach(el => { const [s, kind] = el.dataset.ci.split("|");
      bindPick(el, () => kind === "人" ? personItems(E.行先, listIn(E, s, kind)) : machItems(listIn(E, s, kind)), x => {
        const n = String(x.v || "").trim(); if (!n || listIn(E, s, kind).some(v => nk(v) === nk(n))) { el.value = ""; return; }
        listOf(E, s, kind).push(n); redraw(); setTimeout(() => { const p = $("dlgBody").querySelector(`[data-ci="${s}|${kind}"]`); if (p) p.focus(); }, 0);
      }); });
    const gOf = sj => { const [s, j] = sj.split("|"); return listOf(E, s, "外注")[+j]; };
    B.querySelectorAll("[data-gi]").forEach(el => { bindPick(el, gaiItems, x => { el.value = x.v; gOf(el.dataset.gi).業者 = x.v; }); el.addEventListener("change", () => { gOf(el.dataset.gi).業者 = el.value.trim(); }); });
    B.querySelectorAll("[data-gn]").forEach(el => el.addEventListener("change", () => { const v = el.value.normalize("NFKC").replace(/[^\d.]/g, ""); el.value = v; gOf(el.dataset.gn).人数 = v; }));
    B.querySelectorAll("[data-unused]").forEach(el => el.onchange = () => { const n = el.dataset.unused; E.未使用機械 = (E.未使用機械 || []).filter(v => nk(v) !== nk(n)); if (el.checked) E.未使用機械.push(n); });
    B.onclick = e => {
      const b = e.target.closest("button"); if (!b) return;
      if (b.dataset.dc) { const [s, kind, j] = b.dataset.dc.split("|"); listOf(E, s, kind).splice(+j, 1); redraw(); }
      if (b.dataset.delg) { const [s, j] = b.dataset.delg.split("|"); listOf(E, s, "外注").splice(+j, 1); redraw(); }
      if (b.dataset.addg) { listOf(E, b.dataset.addg, "外注").push({ 業者: "", 人数: "" }); redraw(); }
      if (b.dataset.act === "addw") { E.作業行 = [...workLines(E).map(l => ({ ...l })), { 予定: "", 報告: "" }]; redraw(); }
      if (b.dataset.delw) { E.作業行 = workLines(E).filter((l, i) => i !== +b.dataset.delw); redraw(); }
    };
    $("eSave").onclick = save;
  }
  async function save() {
    const pv = $("ePlace").value.trim();
    if (pv !== E.行先) { const s = siteOf(pv); E.行先 = pv; E.コード = s ? s.コード : E.コード; }
    if (!E.行先) { toast("現場を入れてください"); $("ePlace").focus(); return; }
    E.始め = normTime($("eT0").value); E.終わり = normTime($("eT1").value);
    if (E.終日) { E.始め = ""; E.終わり = ""; }
    E.作業行 = workLines(E).filter((l, i, a) => l.予定 || l.報告 || a.length === 1);
    E.未使用機械 = (E.未使用機械 || []).filter(n => listIn(E, "p", "機械").some(v => nk(v) === nk(n)) && !listIn(E, "r", "機械").some(v => nk(v) === nk(n)));
    ["人", "機械"].forEach(x => { listOf(E, "r", x); listOf(E, "p", x); });
    ["p", "r"].forEach(s => { const a = listOf(E, s, "外注"); a.splice(0, a.length, ...a.filter(g => g.業者)); });
    syncWork(E); E.記入者 = S.me.name; E.更新 = nowStr();
    $("eSave").disabled = true;
    // 同じ時間帯にほかで入っている人・機械（報告が無ければ予定）を尋ねる
    const pick = x => listIn(E, "r", x).length ? listIn(E, "r", x) : listIn(E, "p", x);
    if (!await okClash(day, E, { 人: [...(E.本人 !== false ? [S.me.name] : []), ...pick("人")], 機械: pick("機械") })) { $("eSave").disabled = false; return; }
    await loadMonth(monthOf(day));
    if (!isNew && day !== k) mutate(k, d => { d.行 = (d.行 || []).filter(y => y.id !== E.id); if (!d.行.length) d.報告 = null; }, true);
    mutate(day, d => { const i = d.行.findIndex(y => y.id === E.id); if (i < 0) d.行.push(E); else d.行[i] = E; }, true);
    closeDlg(); S.day = day; render();
    toast(isNew ? "登録しました" : "更新しました");
    syncKikai(day, E).catch(e => toast("自社機械稼働表に書けませんでした：" + e.message, 6000));
  }
  draw();
}

/* ---------- 自社機械稼働表の予約と連動（稼働データ.json。稼働表と同じく eTag で止めて読み直し、同じ変更をかけ直す） ---------- */
const workDayOf = (y, k) => { if ((y.休み || []).includes(k)) return false; if ((y.出勤 || []).includes(k)) return true; if (holidayOf(k) || dow(k) === 0) return !!y.日祝; if (dow(k) === 6) return !!y.土; return true; };
async function syncKikai(k, r) {
  if (!S.kikai || !r.行先 || isOffice(r.行先)) return;
  const ms = S.kikai.機械 || [], find = n => ms.find(m => nk(machName(m)) === nk(n)) || ms.find(m => nk(m.名前) === nk(n));
  const used = listIn(r, "r", "機械").map(find).filter(Boolean), unused = (r.未使用機械 || []).map(find).filter(Boolean);
  if (!used.length && !unused.length) return;
  const names = kikaiNames(r.行先, r.コード), kubun = kubunOf(r);
  const ids = used.map(() => "n" + uid());   // かけ直しても同じ id
  const fn = doc => {
    const R = doc.予約 ||= [], gen = (doc.現場 || []).find(s => names.includes(nk(s.名前))), place = gen ? gen.名前 : (koteiSite(r.行先, r.コード) || {}).名前 || r.行先;
    const here = m => R.filter(y => y.機械 === m.id && (names.includes(nk(y.現場)) || nk(y.現場) === nk(place)) && (y.区分 || "昼") === kubun);
    let n = 0;
    used.forEach((m, i) => {
      const hs = here(m), cov = hs.find(y => y.開始 <= k && y.終了 >= k);
      if (cov) {
        if (kikaiWorks(cov, k)) return;
        cov.未使用 = (cov.未使用 || []).filter(x => x !== k); cov.休み = (cov.休み || []).filter(x => x !== k);
        if (!kikaiWorks(cov, k)) (cov.出勤 ||= []).push(k);
        cov.更新 = nowStr(); n++; return;
      }
      // 前の予約が続いている（あいだの日がその予約で使わない日だけ・4日まで）→ 延ばす
      const prev = hs.filter(y => y.終了 < k && y.終了 >= dayAdd(k, -4)).find(y => { for (let d = dayAdd(y.終了, 1); d < k; d = dayAdd(d, 1)) if (workDayOf(y, d)) return false; return true; });
      if (prev) { prev.終了 = k; if (!kikaiWorks(prev, k)) (prev.出勤 ||= []).push(k); prev.更新 = nowStr(); n++; return; }
      R.push({ id: ids[i], 機械: m.id, 区分: kubun, 現場: place, 開始: k, 終了: k, 土: dow(k) === 6, 日祝: dow(k) === 0 || !!holidayOf(k), 休み: [], 出勤: [], 状態: "確定", メモ: `工事日報から（${S.me.name}）`, 更新: nowStr() });
      n++;
    });
    unused.forEach(m => { const cov = here(m).find(y => y.開始 <= k && y.終了 >= k); if (cov && !(cov.未使用 || []).includes(k)) { (cov.未使用 ||= []).push(k); cov.更新 = nowStr(); n++; } });
    return n;
  };
  for (let i = 0; i < 5; i++) {
    const cur = await kikaiStore.get("稼働データ.json"); if (!cur) return;
    if (!fn(cur.data)) { S.kikai = cur.data; return; }
    try { await kikaiStore.put("稼働データ.json", cur.data, cur.eTag); S.kikai = cur.data; toast("自社機械稼働表の予約も直しました"); return; }
    catch (e) { if (e.status !== 412 && e.status !== 409) throw e; }
  }
  throw new Error("ほかの人と同時に書いていて保存できませんでした");
}

start();
