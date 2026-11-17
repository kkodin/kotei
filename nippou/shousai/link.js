/* 工事日報とつなぐ（2026-11-03 ユーザー：入り口は工事日報の1本。常用工は詳細日報を最初から開く。どちらに書いても反映。同じ窓で移る）
   ・ほかのアプリへ：上の帯の 工事日報・全体工程・機械予定（同じ窓）。
   ・工事日報 → 詳細日報：開いている日に、工事日報に手で書いた行（配置表から取り込んだ行は入れない）の人・外注・作業・機械を入れる。
     一度入れたものは 日.連携 に覚える（詳細日報で消したものを、また入れない）。LINK_FROM より前の日はつながない（前の日報は変えない）。
   ・詳細日報 → 工事日報：工事日報の側が読むときに重ねる（工事日報の link.js）。
   ・機械の予定：自社機械稼働表のこの日・この現場の予約を見せ、記入者が「使った」「使わなかった」を確定する。
   ・自社機械稼働表へ：所有者が自社の使用機材は予約に入れる（続いている予約は延ばす）。「使わなかった」は 未使用 に入れる。
   ・確かめ：配置表・前の日・機械の予定とくらべ、前の日のコピーのまま・人や機械のちがいを上に出す。「確認した」で消える（中身が変わるとまた出す）。 */
"use strict";
const LINK_FROM = "2026-11-01";
const LK = { key: "", busy: false, ndocs: {}, kk: null, chk: 0, timer: null };
const linkUrl = (id, q) => HX.appUrl(id, q);   // アプリの一覧とアドレスは共通の部品（common/apps.js）
// 上の帯のボタン：書いた分を共有してから移る
async function linkGo(id, q) {
  for (const k of ["up", "dn", "gai", "car"]) if (grids[k] && grids[k].editing()) grids[k].commit();
  try { await Promise.race([sync(), new Promise(r => setTimeout(r, 2500))]); } catch (e) { }
  location.href = linkUrl(id, q);
}
$("goNippou").onclick = () => linkGo("nippou", "?stay=1");
$("goKotei").onclick = () => linkGo("kotei");
$("goKikai").onclick = () => linkGo("kikai");
$("calTop").onclick = () => linkGo("nippou", "?stay=1");   // 「日を選ぶ」から入口（1週間の表）へ（2026-11-15 ユーザー）
// 用紙の上の「現場ごと／記入者ごと」と左右の ◀▶ はやめた（2026-11-15 ユーザー：記入者ごとで戻れなくなる）。前に記入者ごとを選んでいた端末も現場ごとに戻す
if (UI.detailView === "writer") { UI.detailView = "site"; saveUI(); }

/* ---------- 読む・書く ---------- */
const demoNippou = () => { try { return JSON.parse(localStorage.getItem("nippou_demo_v4")) || {}; } catch (e) { return {}; } };
async function nippouDocs(m, force) {
  const c = LK.ndocs[m]; if (c && !force && Date.now() - c.t < 30000) return c.docs;
  let docs = [];
  if (DEMO) docs = Object.entries(demoNippou()).filter(([n]) => n.startsWith("日報/個人_") && n.endsWith(`_${m}.json`)).map(([, x]) => x.data);
  else {
    const F = CFG.kotei, auth = { headers: { Authorization: "Bearer " + await token() } };
    const res = await fetch(`${GRAPH}/drives/${F.driveId}/items/${F.itemId}:/${enc("日報")}:/children?$top=999&$select=name,@microsoft.graph.downloadUrl`, auth);
    if (!res.ok) throw httpErr(res, "工事日報");
    const L = (await res.json()).value.filter(x => x.name.startsWith("個人_") && x.name.endsWith(`_${m}.json`));
    docs = (await Promise.all(L.map(x => fetch(x["@microsoft.graph.downloadUrl"], { cache: "no-store" }).then(r => r.json()).catch(() => null)))).filter(Boolean);
  }
  LK.ndocs[m] = { t: Date.now(), docs };
  return docs;
}
const KFILE = "稼働データ.json";
async function kikaiRead() {
  if (DEMO) { const x = demoNippou()["機械/" + KFILE]; return x ? { doc: x.data, eTag: x.eTag } : null; }
  const F = CFG.kikai, res = await fetch(`${GRAPH}/drives/${F.driveId}/items/${F.itemId}:/${enc(KFILE)}`, { headers: { Authorization: "Bearer " + await token() } });
  if (res.status === 404) return null; if (!res.ok) throw httpErr(res, KFILE);
  const m = await res.json(), b = await fetch(m["@microsoft.graph.downloadUrl"], { cache: "no-store" });
  return { doc: await b.json(), eTag: m.eTag };
}
async function kikaiWrite(doc, eTag) {
  if (DEMO) { const a = demoNippou(); a["機械/" + KFILE] = { data: doc, eTag: "d" + Date.now() }; localStorage.setItem("nippou_demo_v4", JSON.stringify(a)); return; }
  const F = CFG.kikai, res = await fetch(`${GRAPH}/drives/${F.driveId}/items/${F.itemId}:/${enc(KFILE)}:/content`, { method: "PUT", headers: { Authorization: "Bearer " + await token(), "Content-Type": "application/json", "If-Match": eTag }, body: JSON.stringify(doc) });
  if (!res.ok) throw httpErr(res, KFILE);
}
async function kikaiDoc(force) {
  if (LK.kk && !force && Date.now() - LK.kk.t < 60000) return LK.kk.doc;
  const r = await kikaiRead(); LK.kk = { t: Date.now(), doc: r ? r.doc : null };
  return LK.kk.doc;
}
const JP_HOL = {}, pad2 = pad, dayAdd = addD, dow = wd;   // 祝日（工事日報と同じ）
function jpHolidays(y) {
  if (JP_HOL[y]) return JP_HOL[y];
  const h = {}, k = (m, d) => y + "-" + pad2(m) + "-" + pad2(d);
  const nthMon = (m, n) => { const w = new Date(y, m - 1, 1).getDay(); return 1 + ((8 - w) % 7) + (n - 1) * 7; };
  const add = (m, d, nm) => { h[k(m, d)] = nm; };
  add(1, 1, "元日"); add(1, nthMon(1, 2), "成人の日"); add(2, 11, "建国記念の日");
  if (y >= 2020) add(2, 23, "天皇誕生日");
  add(3, Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4)), "春分の日");
  add(4, 29, "昭和の日"); add(5, 3, "憲法記念日"); add(5, 4, "みどりの日"); add(5, 5, "こどもの日");
  add(7, nthMon(7, 3), "海の日"); add(8, 11, "山の日"); add(10, nthMon(10, 2), "スポーツの日");
  add(9, nthMon(9, 3), "敬老の日");
  add(9, Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4)), "秋分の日");
  add(11, 3, "文化の日"); add(11, 23, "勤労感謝の日");
  Object.keys(h).forEach(x => { const n = dayAdd(x, 2), mid = dayAdd(x, 1);
    if (h[n] && !h[mid] && dow(mid) !== 0 && mid.slice(0, 4) === String(y)) h[mid] = "国民の休日"; });
  Object.keys(h).sort().forEach(x => { if (dow(x) !== 0) return; let n = dayAdd(x, 1); while (h[n]) n = dayAdd(n, 1); if (n.slice(0, 4) === String(y)) h[n] = "振替休日"; });
  return JP_HOL[y] = h;
}
const kikaiName = m => m.名前 + (m.規格 ? " " + m.規格 : "");
function kikaiWorks(y, k) {
  if (k < y.開始 || k > y.終了) return false;
  if ((y.未使用 || []).includes(k) || (y.休み || []).includes(k)) return (y.出勤 || []).includes(k);
  if ((y.出勤 || []).includes(k)) return true;
  const w = wd(k);
  if (jpHolidays(+k.slice(0, 4))[k] || w === 0) return !!y.日祝;
  return w === 6 ? !!y.土 : true;
}
// 稼働表の現場の名前：コードが同じ現場・略称・正式名
function kikaiSiteNames(s, doc) {
  const g = (doc.現場 || []).filter(x => s.コード && x.コード && String(x.コード) === String(s.コード)).map(x => x.名前);
  return [...g, s.略称, s.名前].filter(Boolean).map(nk);
}
// この日・この現場の予約（使わなかった日も見せる）
function kikaiPlans(s, k, doc) {
  const names = kikaiSiteNames(s, doc), ms = doc.機械 || [], out = [];
  (doc.予約 || []).forEach(y => {
    if (y.状態 === "仮" || !names.includes(nk(y.現場)) || !(kikaiWorks(y, k) || (y.未使用 || []).includes(k))) return;
    const m = ms.find(x => x.id === y.機械); if (m && !out.some(o => o.m.id === m.id)) out.push({ y, m });
  });
  return out;
}
const rowOfMach = (d, m) => ((d && d.下) || []).find(r => r && r[5] && nk(r[5]) === nk(m.名前) && nk(r[6] || "") === nk(m.規格 || ""));
const slotOf = (arr, i) => { let r = 0; while (arr[r] && arr[r][i]) r++; return r; };
// 足りなければページを足して、その行を入れる
function linkPut(key, r, cells) {
  const d = ensureDay(), n = Math.ceil((r + 1) / PG[key]); if (n > pageCount(d)) d.頁 = n;
  applyCells(key, cells.map(([c, v]) => ({ r, c, v })), true);
}

/* ---------- 工事日報 → 詳細日報 ---------- */
const sameSite = (s, r) => (s.コード && r.コード ? String(s.コード) === String(r.コード) : siteScore(s, r.行先) === 3);
async function linkPull(s, k) {
  const docs = await nippouDocs(k.slice(0, 7));
  if (site() !== s || UI.day !== k || !canEdit() || busyEditing()) return;
  const got = [];
  docs.forEach(doc => (((doc.日 || {})[k] || {}).行 || []).forEach(r => {
    if (!r.行先 || !sameSite(s, r) || /^(jouyou|haichi):/.test(String(r.取込 || "")) || r.詳細 || r.書式) return;   // 書式のある行は、この用紙で書いた日報そのもの
    const h = parseInt(String(r.始め || "").split(":")[0], 10), sh = r.区分 === "夜" || (!r.区分 && (h >= 17 || h < 5)) ? "夜" : "昼";
    const rep = r.人 || [], who = [...(r.本人 !== false && doc.名前 ? [doc.名前] : []), ...(rep.length ? rep : ((r.予定 || {}).人 || []))];
    who.forEach(n => got.push({ t: "人", key: `人|${nk(n)}|${sh}`, n, sh }));
    (r.外注 || []).forEach(g => (g.名前 || []).forEach(x => x && x.名前 && got.push({ t: "外", key: `外|${nk(x.名前)}|${sh}`, n: x.名前, sh, co: g.業者 })));
    (r.作業行 || []).forEach(l => { if (l.予定) got.push({ t: "予", key: "予|" + nk(l.予定), v: l.予定 }); if (l.報告) got.push({ t: "報", key: "報|" + nk(l.報告), v: l.報告, p: l.予定 }); });
    (r.機械 || []).forEach(v => got.push({ t: "機", key: "機|" + nk(v), v }));
  }));
  const seen = (day() || {}).連携 || {}, news = got.filter((x, i) => !seen[x.key] && got.findIndex(y => y.key === x.key) === i);
  if (!news.length) return;
  const mach = news.some(x => x.t === "機") ? await kikaiDoc().catch(() => null) : null;
  if (site() !== s || UI.day !== k) return;
  let n = 0;
  news.forEach(x => {
    const d = ensureDay();
    if (x.t === "人" || x.t === "外") {
      const dup = rowsOf(d).some(r => r && r[0] && nk(r[0]) === nk(x.n) && shiftOf(r, s) === x.sh);
      if (!dup) { const key = x.t === "人" ? "dn" : "gai"; linkPut(key, slotOf(d[ARR[key]] ||= [], 0), [[0, masterName(x.n)], [14, x.sh], ...(x.co ? [[12, x.co]] : [])]); n++; }
    } else if (x.t === "予") {
      if (!d.上.some(r => r && String(r[1] || "").trim() === x.v.trim())) { linkPut("up", slotOf(d.上, 1), [[1, x.v]]); n++; }
    } else if (x.t === "報") {
      const i = x.p ? d.上.findIndex(r => r && String(r[1] || "").trim() === x.p.trim()) : -1;
      if (!d.上.some(r => r && String(r[4] || "").trim() === x.v.trim())) { linkPut("up", i >= 0 && !(d.上[i] || [])[4] ? i : slotOf(d.上, 4), [[4, x.v]]); n++; }
    } else if (x.t === "機") {
      const m = mach && (mach.機械 || []).find(v => nk(kikaiName(v)) === nk(x.v));
      if (!(m ? rowOfMach(d, m) : d.下.some(r => r && nk(r[5] || "") === nk(x.v)))) { linkPut("dn", slotOf(d.下, 5), m ? [[5, m.名前], [6, m.規格 || ""], [13, SELF], [7, "1"]] : [[5, x.v]]); n++; }
    }
    const dd = ensureDay(); (dd.連携 ||= {})[x.key] = 1;
  });
  const dd = day(); if (dd) { dd.更新 = now(); save(); }
  if (n) toast(`工事日報に書いてある分を ${n}件 入れました`, 4000);
}

/* ---------- 自社機械稼働表へ（所有者が自社の使用機材・使わなかった機械） ---------- */
async function kikaiPush(s, k) {
  const d = s.日[k]; if (!d) return;
  const first = await kikaiDoc(); if (!first) return;
  const ms = (first.機械 || []).filter(m => !m.使わない), find = r => { const L = ms.filter(m => nk(m.名前) === nk(r[5])), e = L.find(m => nk(m.規格 || "") === nk(r[6] || "")); return e || (L.length === 1 ? L[0] : null); };
  const used = [];
  (d.下 || []).forEach(r => { if (!r || !r[5] || r[13] !== SELF) return; const m = find(r); if (m && !used.some(u => u.m.id === m.id)) used.push({ m, sh: shiftOf(r, s) }); });
  const unused = d.未使用機材 || [], sig = JSON.stringify([used.map(u => u.m.id + u.sh).sort(), [...unused].sort()]);
  if (d.稼働 === sig || (!used.length && !unused.length && !d.稼働)) return;
  const ids = used.map(() => "n" + uid());
  const fn = doc => {
    const R = doc.予約 ||= [], names = kikaiSiteNames(s, doc), gen = (doc.現場 || []).find(x => names.includes(nk(x.名前))), place = gen ? gen.名前 : s.略称 || s.名前;
    const here = (id, sh) => R.filter(y => y.機械 === id && names.includes(nk(y.現場)) && (!sh || (y.区分 || "昼") === sh));
    let n = 0;
    used.forEach((u, i) => {
      const hs = here(u.m.id, u.sh), cov = hs.find(y => y.開始 <= k && y.終了 >= k);
      if (cov) {
        if (kikaiWorks(cov, k)) return;
        cov.未使用 = (cov.未使用 || []).filter(x => x !== k); cov.休み = (cov.休み || []).filter(x => x !== k);
        if (!kikaiWorks(cov, k)) (cov.出勤 ||= []).push(k);
        cov.更新 = now(); n++; return;
      }
      const prev = hs.filter(y => y.終了 < k && y.終了 >= addD(k, -4)).find(y => { for (let t = addD(y.終了, 1); t < k; t = addD(t, 1)) if (kikaiWorks({ ...y, 終了: "9999-12-31" }, t)) return false; return true; });
      if (prev) { prev.終了 = k; if (!kikaiWorks(prev, k)) (prev.出勤 ||= []).push(k); prev.更新 = now(); n++; return; }
      const w = wd(k);
      R.push({ id: ids[i], 機械: u.m.id, 区分: u.sh, 現場: place, 開始: k, 終了: k, 土: w === 6, 日祝: w === 0 || !!jpHolidays(+k.slice(0, 4))[k], 休み: [], 出勤: [], 未使用: [], 状態: "確定", メモ: `詳細日報から（${ME.名前 || ""}）`, 更新: now() });
      n++;
    });
    unused.forEach(id => { const cov = here(id).find(y => y.開始 <= k && y.終了 >= k); if (cov && !(cov.未使用 || []).includes(k)) { (cov.未使用 ||= []).push(k); cov.更新 = now(); n++; } });
    return n;
  };
  for (let i = 0; i < 5; i++) {
    const cur = await kikaiRead(); if (!cur) return;
    const n = fn(cur.doc);
    try {
      if (n) await kikaiWrite(cur.doc, cur.eTag);
      LK.kk = { t: Date.now(), doc: cur.doc }; d.稼働 = sig; save();
      if (n) toast("自社機械稼働表の予約も直しました"); linkPaintKikai(); return;
    } catch (e) { if (e.status !== 412 && e.status !== 409) throw e; }
  }
}

/* ---------- 機械の予定（記入者が確定する） ---------- */
async function linkPaintKikai() {
  const el = $("kikaiPlan"), s = site(), k = UI.day; if (!el) return;
  if (!s || !signedIn()) { el.innerHTML = ""; return; }
  const doc = await kikaiDoc().catch(() => null);
  if (site() !== s || UI.day !== k) return;
  const P = doc ? kikaiPlans(s, k, doc) : [], d = day();
  if (!P.length) { el.innerHTML = ""; LK.left = []; return; }
  const st = p => rowOfMach(d, p.m) ? "used" : (((d && d.未使用機材) || []).includes(p.m.id) || (p.y.未使用 || []).includes(k)) ? "no" : "";
  const left = P.filter(p => !st(p)); LK.left = left.map(p => kikaiName(p.m));
  el.innerHTML = `<div class="secT"><b>機械の予定（自社機械稼働表）</b>${left.length ? `<span class="warn">${left.length}台 まだ確定していません。使ったかどうかを押してください</span>` : `<span class="cnt">全部確定しました</span>`}</div>
    <div class="kpL">${P.map(p => { const t = st(p); return `<div class="kpI ${t}"><span class="kpN">${t === "used" ? "✓ " : t === "no" ? "× " : ""}${esc(kikaiName(p.m))}<small>${esc(p.y.区分 || "昼")}</small></span><button type="button" class="sub small${t === "used" ? " on" : ""}" data-kuse="${esc(p.m.id)}">使った（確定）</button><button type="button" class="sub small${t === "no" ? " on" : ""}" data-kno="${esc(p.m.id)}">使わなかった</button></div>`; }).join("")}</div>`;
  el.querySelectorAll("[data-kuse]").forEach(b => b.onclick = () => { if (!canEdit()) return; const p = P.find(x => x.m.id === b.dataset.kuse), dd = ensureDay();
    dd.未使用機材 = (dd.未使用機材 || []).filter(x => x !== p.m.id);
    if (!rowOfMach(dd, p.m)) linkPut("dn", slotOf(dd.下, 5), [[5, p.m.名前], [6, p.m.規格 || ""], [13, SELF], [7, "1"]]); else { dd.更新 = now(); save(); refreshSheet(); }
    kikaiPush(s, k).catch(e => toast("自社機械稼働表に書けませんでした：" + e.message, 6000)); });
  el.querySelectorAll("[data-kno]").forEach(b => b.onclick = () => { if (!canEdit()) return; const p = P.find(x => x.m.id === b.dataset.kno), dd = ensureDay();
    (dd.未使用機材 ||= []).includes(p.m.id) || dd.未使用機材.push(p.m.id);
    const row = rowOfMach(dd, p.m); if (row) { const i = dd.下.indexOf(row); applyCells("dn", [5, 6, 7, 13].map(c => ({ r: i, c, v: "" })), true); }
    dd.更新 = now(); save(); refreshSheet();
    kikaiPush(s, k).catch(e => toast("自社機械稼働表に書けませんでした：" + e.message, 6000)); });
}

/* ---------- 確かめ（前の日のコピーのまま・配置表とのちがい） ---------- */
async function linkCheck(s, k, tk) {
  const el = $("chkBar"), d = s.日[k]; if (!el) return;
  if (!hasData(d)) { el.innerHTML = ""; return; }
  const L = [], names = rows => rows.filter(r => r && r[0]).map(r => r[0]);
  const prevK = Object.keys(s.日).filter(x => x < k && hasData(s.日[x])).sort().pop(), prev = prevK ? s.日[prevK] : null, ml = x => `${+x.slice(5, 7)}/${+x.slice(8)}`;
  const work = x => (x.上 || []).map(r => r && String(r[1] || "").trim()).filter(Boolean).sort().join("\n"), ppl = x => names(rowsOf(x)).map(nk).sort().join();
  const mach = x => ((x.下 || []).filter(r => r && r[5] && !/通勤/.test(r[5])).map(r => nk(r[5] + (r[6] || "")))).sort().join();
  let hai = null; try { hai = await haichiDay(k); } catch (e) { }
  if (tk !== LK.chk) return;
  const dn = new Set(names(rowsOf(d)).map(nk));
  if (hai) {
    const D = hai.data, rows = bestOf((D.現場 || []).filter(r => r && r.現場名), r => siteScore(s, r.現場名, r.現場コード)), hp = [];
    rows.forEach(r => (r.作業員 || []).forEach(w => { const n = masterName(String((w && w.氏名) || w || "").trim()); if (n && !hp.includes(n)) hp.push(n); }));
    // 担当が現場に行った人（配置表は名字だけ）は、日報にいれば配置表にもいるとみる
    const tanto = rows.filter(r => r.担当が行った && r.担当).map(r => nk(r.担当));
    const lack = hp.filter(n => !dn.has(nk(n)));
    if (lack.length) L.push(`配置表（${hai.版}）にいるのに、日報にいない人：${lack.join("、")}`);
    if (rows.length || hp.length) {
      const hs = new Set(hp.map(nk)), other = names(rowsOf(d)).filter(n => !hs.has(nk(n)) && !tanto.some(t => nk(n).startsWith(t)));
      const where = n => { const x = (D.現場 || []).find(r => (r.作業員 || []).some(w => nk((w && w.氏名) || w) === nk(n))); return x ? `${n}（配置表は ${x.現場名}）` : (D.本社 || []).some(h => nk(h.氏名) === nk(n)) ? `${n}（配置表は 本社など）` : n; };
      if (other.length) L.push(`日報にいるのに、配置表（${hai.版}）のこの現場にいない人：${other.map(where).join("、")}`);
    }
    const hw = [...new Set(rows.map(r => String(r.作業内容 || "").trim()).filter(Boolean))].join(" ／ ");
    if (hw && prev && work(d) && work(d) === work(prev)) L.push(`作業予定が前の日（${ml(prevK)}）と同じです。配置表の作業内容：${hw}`);
  }
  if (prev) {
    if (!hai && work(d) && work(d) === work(prev)) L.push(`作業予定が前の日（${ml(prevK)}）と同じです。今日の作業か確かめてください`);
    if (!hai && ppl(d) && ppl(d) === ppl(prev)) L.push(`作業員が前の日（${ml(prevK)}）と全員同じです。今日の人か確かめてください`);
    if (mach(d) && mach(d) === mach(prev)) L.push(`使用機材が前の日（${ml(prevK)}）と全部同じです。今日使った機械か確かめてください`);
  }
  const doc = await kikaiDoc().catch(() => null);
  if (tk !== LK.chk) return;
  if (doc) {
    const P = kikaiPlans(s, k, doc), names2 = kikaiSiteNames(s, doc), ms = doc.機械 || [];
    const left = P.filter(p => !rowOfMach(d, p.m) && !(d.未使用機材 || []).includes(p.m.id) && !(p.y.未使用 || []).includes(k)).map(p => kikaiName(p.m));
    if (left.length) L.push(`稼働表に予定があるのに、確定していない機械：${left.join("、")}`);
    (d.下 || []).forEach(r => { if (!r || !r[5] || r[13] !== SELF) return; const m = ms.find(v => nk(v.名前) === nk(r[5]) && nk(v.規格 || "") === nk(r[6] || "")); if (!m) return;
      const y = (doc.予約 || []).find(y => y.機械 === m.id && y.状態 !== "仮" && (y.区分 || "昼") === shiftOf(r, s) && kikaiWorks(y, k) && !names2.includes(nk(y.現場)));
      if (y) L.push(`${kikaiName(m)} は稼働表ではこの日「${y.現場}」に予約があります`); });
  }
  const fp = JSON.stringify(L);
  if (!L.length) { el.innerHTML = ""; return; }
  if (d.確認済 === fp) { el.innerHTML = `<div class="chkOk">確認済みです（${L.length}件）<button type="button" class="sub small" id="chkShow">見る</button></div>`; $("chkShow").onclick = () => { delete d.確認済; linkPaint(true); }; return; }
  el.innerHTML = `<div class="chkW"><b>⚠ 確かめてください（${L.length}件）</b><ul>${L.map(x => `<li>${esc(x)}</li>`).join("")}</ul><button type="button" class="sub small" id="chkDone">確認した（このままでよい）</button></div>`;
  $("chkDone").onclick = () => { if (!canEdit()) return; d.確認済 = fp; d.更新 = now(); save(); linkPaint(true); };
}

/* ---------- 画面を描くたびに ---------- */
function linkPaint(force) {
  const s = site(), k = UI.day; if (!$("kikaiPlan")) return;
  linkPaintKikai();
  const tk = ++LK.chk; if (s) linkCheck(s, k, tk).catch(e => console.warn(e)); else $("chkBar").innerHTML = "";
  clearTimeout(LK.timer);
  if (!s || !signedIn() || k < LINK_FROM) return;
  const key = s.id + "|" + k, fresh = key !== LK.key;
  LK.timer = setTimeout(async () => {
    if (site() !== s || UI.day !== k || !canEdit() || busyEditing() || LK.busy) return;
    LK.busy = true;
    try { if (fresh || force) { LK.key = key; await linkPull(s, k); } await kikaiPush(s, k); }
    catch (e) { console.warn("工事日報・稼働表とつなげませんでした", e); }
    LK.busy = false;
  }, fresh ? 300 : 2000);
}
const linkRS0 = refreshSheet;
refreshSheet = function () { linkRS0(); linkPaint(); };
setInterval(() => { if (!document.hidden && UI.tab !== "master" && !busyEditing()) { LK.key = ""; linkPaint(); } }, 60000);

/* ---------- 工事日報の入り口から（?open=entry&code=&name=&d=） ---------- */
function linkEntry(n) {
  const q = new URLSearchParams(location.search); if (q.get("open") !== "entry") return;
  if (!q.get("code") && !q.get("name")) {   // 現場を決めずに開く（今日の日報でまだ現場が無いとき）：用紙の上の現場名の欄に入れる
    history.replaceState(null, "", location.pathname + (DEMO ? "?demo" : ""));
    UI.site = null; UI.gen = null; UI.newForm = ""; UI.day = q.get("d") || ymd(new Date()); UI.pg = 0; saveUI(); showTab("sheet"); renderAll();   // 新規の日報：元請・現場は空欄、区分は選んでから
    setTimeout(blinkForm, 300); return;
  }
  const code = q.get("code"), name = q.get("name"), x = DB.現場.find(v => code && String(v.コード) === code) || (name && bestOf(DB.現場, v => siteScore(v, name))[0]);
  if (!x) {   // 現場一覧を読む前かもしれないので、しばらく待ってから探し直す。見つからなければ知らせる
    if ((n || 0) < 6) setTimeout(() => linkEntry((n || 0) + 1), 2000);
    else toast(`「${name || code}」は詳細日報の現場にまだありません。「現場名を選ぶ」か「現場・名簿」で現場を足してから開いてください`, 8000);
    return;
  }
  history.replaceState(null, "", location.pathname + (DEMO ? "?demo" : ""));
  UI.gen = x.元請id; UI.site = x.id; if (q.get("d")) UI.day = q.get("d"); UI.pg = 0; saveUI(); showTab("sheet"); renderAll(); askForm();
}
linkEntry();
linkPaint();
