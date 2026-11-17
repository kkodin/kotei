/* 詳細日報とつなぐ（2026-11-03 ユーザー：入り口は工事日報の1本。常用工は詳細日報を最初から開く。どちらに書いても反映。同じ窓で移る）
   ・入り口：今日の自分の現場（日報の行、無ければ配置表）が常用の現場なら、はじめに詳細日報を開く。詳細日報の「工事日報」ボタンから来たとき（?stay）・同じタブで2回目からは開かない。
   ・詳細日報 → 工事日報：カレンダーを描くときに、詳細日報の人を行として重ねる（行.詳細。書き込まない）。押すと詳細日報のその日を開く。
     その人がその日その現場の行をもう持っていれば重ねない。LINK_FROM より前の日は重ねない（前の日報は変えない）。
   ・工事日報 → 詳細日報：詳細日報を開いたときに、詳細日報の側（link.js）が入れる。 */
"use strict";
const LINK_FROM = "2026-11-01";
const LINKC = {};   // 月 → { t, docs }
const detailUrl = (place, code, k) => HX.appUrl("shousai", `?open=entry&code=${encodeURIComponent(code || "")}&name=${encodeURIComponent(place || "")}&d=${k}`);
async function jouyouDocs() {
  if (DEMO) { let a = {}; try { a = JSON.parse(localStorage.getItem("jouyou_demo_store")) || {}; } catch (e) { } return Object.entries(a).filter(([n]) => /^現場_.*\.json$/.test(n)).map(([, x]) => { try { return JSON.parse(x.data); } catch (e) { return null; } }).filter(Boolean); }
  const files = (await jouyouStore.list()).filter(f => /^現場_.*\.json$/.test(f.name));
  return (await Promise.all(files.map(f => jouyouStore.get(f.name).then(x => x && x.data).catch(() => null)))).filter(Boolean);
}
// 詳細日報の人を、工事日報の行の形にする（文書：{ 名前, 月, 日, 仮: true }）
async function linkDocs(ms, real) {
  const out = [];
  for (const m of ms) {
    if (m + "-31" < LINK_FROM) continue;
    const c = LINKC[m];
    if (!c || Date.now() - c.t > 30000) LINKC[m] = { t: Date.now(), js: await jouyouDocs().catch(() => []) };
    const by = new Map();
    LINKC[m].js.forEach(x => { const s = x.現場 || {};
      Object.entries(s.日 || {}).forEach(([k, d]) => { if (monthOf(k) !== m || k < LINK_FROM || !d) return;
        const pl = plOf(s.略称 || s.名前, s.コード), L = (d.上 || []).filter(r => r && (r[1] || r[4])).map(r => ({ 予定: String(r[1] || ""), 報告: String(r[4] || "") }));
        [...(d.下 || []), ...(d.外注 || [])].forEach(r => { if (!r || !r[0]) return; const n = findPerson(r[0]) || haiPerson(r[0]); if (!n || !findPerson(n)) return;
          const has = real.some(doc => ((doc.日 || {})[k] && doc.日[k].行 || []).some(y => y.行先 && plOf(y.行先, y.コード).key === pl.key && membersOf({ who: doc.名前, r: y }).some(v => nk(v) === nk(n)) || (y.外注 || []).some(g => (g.名前 || []).some(w => nk(w.名前) === nk(n)))));   // 外注の名前に出ている人も重ねない（人数を二重にしない）
          if (has) return;
          if (!by.has(n)) by.set(n, { 名前: n, 月: m, 日: {}, 仮: true });
          const row = { id: "jl" + nk(s.コード || s.名前) + k + (r[14] || ""), 行先: pl.label, コード: pl.code, 始め: r[1] || "", 終わり: r[2] || "", 区分: r[14] === "夜" ? "夜" : "昼", 作業: "", 作業行: L.length ? L : [{ 予定: "", 報告: "" }], 予定: { 人: [], 外注: [], 機械: [] }, 人: [], 外注: [], 機械: [], 記入者: d.担当者 || "", 詳細: true };
          syncWork(row);
          const D = by.get(n).日[k] ||= { 行: [], 報告: { だれ: d.担当者 || "", いつ: "詳細日報" } }; D.行.push(row);
        }); }); });
    out.push(...by.values());
  }
  return out;
}
const openDetail = x => { location.href = detailUrl(x.r.行先, x.r.コード, x.k); };

// カードの上に、詳細日報に書いてあることを知らせる
async function linkNotice() {
  const k = S.day; if (k < LINK_FROM) return;
  const real = await monthDocs(monthOf(k)).catch(() => []), v = (await linkDocs([monthOf(k)], real)).find(x => nk(x.名前) === nk(S.me.name)), L = v && v.日[k] && v.日[k].行;
  if (S.view !== "me" || S.day !== k || !L || !L.length || !$("dayHead") || $("linkMsg")) return;
  const el = document.createElement("div"); el.id = "linkMsg"; el.className = "linkmsg";
  el.innerHTML = `詳細日報に ${esc(L[0].行先)} の記録があります <button type="button">詳細日報を開く</button>`;
  el.querySelector("button").onclick = () => { location.href = detailUrl(L[0].行先, L[0].コード, k); };
  $("dayHead").appendChild(el);
}
const linkR0 = render;
render = function () { const v = linkR0.apply(this, arguments); if (S.view === "me") linkNotice().catch(() => { }); return v; };

// 入り口：常用の現場の人は、はじめに詳細日報を開く。移ったら true
async function entryJump() {
  const q = new URLSearchParams(location.search);
  if (q.has("stay") || q.has("site") || q.has("setup")) { try { sessionStorage.setItem("nippou_entry", "1"); } catch (e) { } return false; }
  try { if (sessionStorage.getItem("nippou_entry")) return false; sessionStorage.setItem("nippou_entry", "1"); } catch (e) { return false; }
  const k = todayKey(); let pl = null;
  const f = S.F[monthOf(k)], rows = f && f.doc && f.doc.日[k] ? f.doc.日[k].行 : [], r = (rows || []).find(y => y.行先 && !isOffice(y.行先));
  if (r) pl = plOf(r.行先, r.コード);
  else if (!(rows || []).length) {
    try { const H = DIFF.H; H.from = k; H.to = k; await haiLoad(H); const x = haiDay(H, k);
      const g = x && (x.d.現場 || []).find(g => (g.作業員 || []).some(w => nk(w.氏名) === nk(S.me.name)) || (g.担当が行った && nk(tantoName(g.担当)) === nk(S.me.name)));
      if (g) pl = plOf(g.現場名, ""); } catch (e) { }
  }
  if (!pl || formOf(pl.label, pl.code) !== "常用") return false;
  location.replace(detailUrl(pl.label, pl.code, k));
  return true;
}
