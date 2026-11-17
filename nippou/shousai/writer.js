/* 記入者ごとの日報（工事日報 Vr005 段階2。2026-11-15 ユーザー：データの持ち方は A＝記入者ごとのファイル）
   ・用紙（この詳細日報の書式）はそのまま。日報1つ（現場・日）の保存先を、記入者のファイル 工程データ\日報\個人_<記入者>_<年月>.json の行にする。
     行.書式 ＝ その日の用紙の中身（上・下・外注・車・天気・気温・担当者…）＋ 現場id。行の 行先・人・外注・機械・作業行 は書式から作る（工事日報の表・一覧にそのまま出る）。
   ・同じ現場・同じ日の日報は1つ。最初に書いた人（担当者）が記入者。現場ごとのファイル（現場_<id>.json）には、移した日は書かない（名簿・単価・決まりは残る）。
   ・他の人の日報も直せる（2026-11-17 ユーザー：見るだけをやめた。同時に直すのは「使用中」の札と交代で防ぐ）。
   ・記入者を押して人を選ぶと、その人の日報を見る。現場名を押すと「現場名を選ぶ」窓（2026-11-17 に入れる欄から窓へ戻した）。
   ・MOVE_FROM より前の日は今までどおり現場ごとのファイル（前の日報は動かさない）。 */
"use strict";
const MOVE_FROM = "2026-11-01";
const WS = { files: {}, cache: {}, seen: {}, sent: {}, where: {}, done: {}, allow: {}, busy: false, readOk: false, dups: [], lost: [], unknown: [], filled: [] };   // files: ファイル名 → eTag／sent・where・done: "現場id|日" ごと
const wKey = (sid, k) => sid + "|" + k;
const wFile = (who, k) => `個人_${nk(who)}_${k.slice(0, 7)}.json`;
const myName = () => (typeof ME !== "undefined" && ME && ME.名前) || "";

/* ---------- 日報のフォルダを読む・書く ---------- */
const WDEMO = "nippou_demo_v4";
const wDemo = () => { try { return JSON.parse(localStorage.getItem(WDEMO)) || {}; } catch (e) { return {}; } };
const wUrl = name => `${GRAPH}/drives/${CFG.kotei.driveId}/items/${CFG.kotei.itemId}:/${enc("日報" + (name ? "/" + name : ""))}`;
const wStore = {
  async list() {
    if (DEMO) return Object.entries(wDemo()).filter(([n]) => n.startsWith("日報/")).map(([n, x]) => ({ name: n.slice(3), eTag: x.eTag }));
    const out = []; let url = wUrl() + ":/children?$top=999&$select=name,eTag,file";
    while (url) {
      const res = await fetch(url, { headers: { Authorization: "Bearer " + await token() } });
      if (res.status === 404) return out; if (!res.ok) throw httpErr(res, "日報の一覧");
      const j = await res.json(); j.value.forEach(x => { if (x.file) out.push({ name: x.name, eTag: x.eTag }); }); url = j["@odata.nextLink"] || null;
    }
    return out;
  },
  async get(name) {
    if (DEMO) { const x = wDemo()["日報/" + name]; return x ? { data: JSON.parse(JSON.stringify(x.data)), eTag: x.eTag } : null; }
    const res = await fetch(wUrl(name), { headers: { Authorization: "Bearer " + await token() } });
    if (res.status === 404) return null; if (!res.ok) throw httpErr(res, name);
    const m = await res.json(), b = await fetch(m["@microsoft.graph.downloadUrl"], { cache: "no-store" });
    return { data: await b.json(), eTag: m.eTag };
  },
  async put(name, data, eTag) {
    if (DEMO) { const a = wDemo(), cur = a["日報/" + name]; if (eTag ? (!cur || cur.eTag !== eTag) : cur) throw Object.assign(new Error("保存 412"), { status: 412 });
      const t = "w" + Date.now() + Math.random().toString(36).slice(2, 5); a["日報/" + name] = { data, eTag: t }; localStorage.setItem(WDEMO, JSON.stringify(a)); return { eTag: t }; }
    const res = await fetch(wUrl(name) + ":/content" + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail"), { method: "PUT",
      headers: { Authorization: "Bearer " + await token(), "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: JSON.stringify(data) });
    if (!res.ok) throw httpErr(res, name + " の保存");
    return { eTag: (await res.json()).eTag };
  }
};

/* ---------- 用紙の中身 → 工事日報の行 ---------- */
function rowOfDay(s, k, d, old) {
  const w = d.担当者, self = rowsOf(d).find(r => r && r[0] && nk(r[0]) === nk(w));
  const ppl = (d.下 || []).filter(r => r && r[0] && nk(r[0]) !== nk(w)).map(r => masterName(r[0]));
  const gai = new Map();
  (d.外注 || []).forEach(r => { if (!r || !r[0]) return; const co = r[12] || "（会社不明）"; if (!gai.has(co)) gai.set(co, { 業者: co, 人数: "0", 名前: [] }); const g = gai.get(co); g.名前.push({ 名前: r[0], 役: /職長/.test(r[3] || "") ? "職長" : "作業員" }); g.人数 = String(g.名前.length); });
  const mach = [...new Set((d.下 || []).filter(r => r && r[5] && !/通勤/.test(r[5])).map(r => r[5] + (r[6] ? " " + r[6] : "")))];
  const L = (d.上 || []).filter(r => r && (r[1] || r[4])).map(r => ({ 予定: String(r[1] || ""), 報告: String(r[4] || "") }));
  const sh = self ? shiftOf(self, s) : s.勤務 === "夜" ? "夜" : "昼";
  const row = { ...(old || {}), id: (old && old.id) || uid(), 行先: s.略称 || s.名前, コード: s.コード || "", 始め: self ? self[1] || "" : "", 終わり: self ? self[2] || "" : "", 終日: false, 区分: sh,
    作業行: L.length ? L : [{ 予定: "", 報告: "" }], 作業: L.map(l => l.報告 || l.予定).join("\n"), 予定: (old && old.予定) || { 人: [], 外注: [], 機械: [] },
    人: ppl, 外注: [...gai.values()], 機械: mach, 記入者: w, 書式: { ...JSON.parse(JSON.stringify(d)), 現場id: s.id }, 更新: now() };
  if (self) delete row.本人; else row.本人 = false;   // 記入者が作業員に入っていなければ代理で書いた日報
  return row;
}
const isMine = (r, sid) => r && r.書式 && r.書式.現場id === sid;

/* ---------- 合わせる（読む → 書く）。sync（現場ごとのファイル）のあとに流す ---------- */
async function writerSync() {
  if (!signedIn() || WS.busy || !DB.現場.length) return;
  WS.busy = true;
  try {
    // 1) 読む：記入者ごとのファイル（MOVE_FROM の月から）。変わったファイルだけ取り直し（WS.cache に中身を持つ）、
    //    どの日報がどのファイルにあるか（WS.seen）は毎回、全部のファイルの中身から作り直す（2026-11-16：読み直さなかったファイルの分を「消された」と取り違えたので）
    const from = MOVE_FROM.slice(0, 7);
    WS.readOk = false;
    const listed = (await wStore.list()).filter(it => { const m = /^個人_.+_(\d{4}-\d{2})\.json$/.exec(it.name); return m && m[1] >= from; });
    for (const it of listed) {
      if (WS.files[it.name] === it.eTag && WS.cache[it.name]) continue;
      const r = await wStore.get(it.name); if (!r) continue;
      WS.files[it.name] = r.eTag; WS.cache[it.name] = r.data;
    }
    WS.seen = {}; WS.unknown = [];
    for (const it of listed) {
      const data = WS.cache[it.name]; if (!data) continue;
      Object.entries(data.日 || {}).forEach(([k, dd]) => (dd.行 || []).forEach(row => {
        if (!row.書式 || k < MOVE_FROM) return;
        const s = DB.現場.find(x => x.id === row.書式.現場id) || DB.現場.find(x => row.コード && String(x.コード) === String(row.コード));
        if (!s) { WS.unknown.includes(it.name + " " + k) || WS.unknown.push(it.name + " " + k); return; }
        const key = wKey(s.id, k); (WS.seen[key] ||= []).includes(it.name) || WS.seen[key].push(it.name);
        const b = { ...row.書式 }; delete b.現場id;
        WS.where[key] = it.name; WS.done[key] = true;
        if (s.消した日[k] && s.消した日[k] >= (b.更新 || "")) return;   // この端末で消した日は戻さない（書くときに記入者のファイルからも外す）
        const a = s.日[k], text = JSON.stringify(b);
        if (!a || (b.更新 || "") >= (a.更新 || "")) { if (!a || JSON.stringify(a) !== text) s.日[k] = b; WS.sent[key] = text; }
      }));
    }
    // 同じ現場・同じ日に、2人（2つのファイル）の日報がある → 片方しか集計されないので知らせる
    WS.dups = Object.entries(WS.seen).filter(([, a]) => a.length > 1).map(([key, a]) => { const [sid, k] = key.split("|"), s = DB.現場.find(x => x.id === sid); return `${s ? siteLabel(s) : sid} ${k}（${a.map(n => n.replace(/^個人_|_\d{4}-\d{2}\.json$/g, "")).join("・")}）`; });
    // 記入者のファイルへ移した日が、どのファイルにも無い ＝ ほかの端末で消された？ → その記入者のファイルを取り直して、本当に無いときだけ控えから外す
    WS.lost = [];
    for (const s of DB.現場) for (const [k, d] of Object.entries(s.日)) {
      const key = wKey(s.id, k);
      if (k < MOVE_FROM || d.保存先 !== "記入者" || (WS.seen[key] || []).length || !d.担当者) continue;
      if (!WS.sent[key] || WS.sent[key] !== JSON.stringify(d)) continue;   // この端末で直して、まだ書いていない（または書いたことが分からない）ものは外さない
      const name = wFile(d.担当者, k), r = await wStore.get(name);
      if (r) { WS.files[name] = r.eTag; WS.cache[name] = r.data; }
      if (!r || !listed.some(it => it.name === name)) continue;   // ファイルが読めない・一覧に無い：判断しない
      if (((r.data.日 || {})[k] || { 行: [] }).行.some(row => isMine(row, s.id))) { (WS.seen[key] ||= []).push(name); continue; }
      delete s.日[k]; WS.lost.push(`${siteLabel(s)} ${k}`); delete WS.where[key]; delete WS.done[key]; delete WS.sent[key];
    }
    if (WS.lost.length) save();
    WS.readOk = true;
    // 2) 書く：MOVE_FROM からの日で、記入者が決まっていて、変わったもの
    const jobs = new Map();   // ファイル名 → [{ s, k, d|null（消す） }]
    const add = (name, j) => { if (!jobs.has(name)) jobs.set(name, []); jobs.get(name).push(j); };
    for (const s of DB.現場) {
      for (const [k, d] of Object.entries(s.日)) {
        if (k < MOVE_FROM || !hasData(d)) continue;
        if (!d.担当者 && !(await fillFromHaichi(s, k, d))) continue;   // 記入者が空 → 配置表のこの現場の「担当」。決まらなければ現場ごとのファイルのまま
        d.保存先 = "記入者";
        const key = wKey(s.id, k), text = JSON.stringify(d); if (WS.sent[key] === text) continue;
        const name = wFile(d.担当者, k);
        add(name, { s, k, d, text });
        if (WS.where[key] && WS.where[key] !== name) add(WS.where[key], { s, k, d: null });   // 記入者が替わった → 前の人のファイルから外す
      }
      for (const k of Object.keys(s.消した日)) { const key = wKey(s.id, k); if (s.日[k]) continue; (WS.seen[key] || (WS.where[key] ? [WS.where[key]] : [])).forEach(name => add(name, { s, k, d: null })); }   // 消した日：入っている記入者のファイル全部から外す
    }
    for (const [name, list] of jobs) {
      for (let n = 0; n < 4; n++) {
        const cur = await wStore.get(name), m = name.match(/_(\d{4}-\d{2})\.json$/)[1];
        const doc = cur ? cur.data : { 形式: "私の日報", 版: 1, 名前: list.find(j => j.d) ? masterName(list.find(j => j.d).d.担当者) : "", アカウント: "", 月: m, 日: {} };
        list.forEach(({ s, k, d }) => {
          const dd = doc.日[k] ||= { 行: [], 報告: null }, i = dd.行.findIndex(r => isMine(r, s.id));
          if (!d) { if (i >= 0) dd.行.splice(i, 1); if (!dd.行.length) dd.報告 = null; return; }
          const row = rowOfDay(s, k, d, i >= 0 ? dd.行[i] : null);
          if (i >= 0) dd.行[i] = row; else dd.行.push(row);
          dd.報告 ||= { だれ: masterName(d.担当者), いつ: "詳細日報" }; dd.更新 = now();
        });
        doc.更新 = now();
        try {
          const r = await wStore.put(name, doc, cur ? cur.eTag : null); WS.files[name] = r.eTag; WS.cache[name] = JSON.parse(JSON.stringify(doc));   // 書いた中身を覚える（次の回の一覧に入る）
          list.forEach(({ s, k, d, text }) => { const key = wKey(s.id, k); if (d) { WS.sent[key] = text; WS.where[key] = name; WS.done[key] = true; } else if (WS.where[key] === name) { delete WS.where[key]; delete WS.sent[key]; } });
          break;
        } catch (e) { if ((e.status !== 412 && e.status !== 409) || n === 3) throw e; }
      }
    }
  } catch (e) { console.warn("記入者ごとの日報を合わせられませんでした", e); syncMark("err", "日報を共有できません"); $("sync").title = e.message; }
  WS.busy = false;
}
// WS.seen："現場id|日" → その日報が入っている記入者のファイル名（2つ以上なら重複）。WS.cache：ファイル名 → 中身
// 記入者が空の日：配置表のその日・この現場の「担当」（名字）から、名簿の人を1人に決める（2026-11-16 ユーザー：配置表の担当の欄を参照）
async function fillFromHaichi(s, k, d) {
  let H = null; try { H = await haichiDay(k); } catch (e) { return false; }
  if (!H || !H.data) return false;
  const rows = bestOf((H.data.現場 || []).filter(r => r && r.現場名), r => siteScore(s, r.現場名, r.現場コード));
  const t = [...new Set(rows.map(r => String(r.担当 || "").replace(/\s/g, "")).filter(Boolean))];
  if (t.length !== 1) return false;
  const all = [...new Set([...MS.社員, ...MS.従業員, ...s.作業員.map(w => w.名前)].map(masterName))];
  const hit = all.filter(n => nk(n).startsWith(nk(t[0])));
  if (hit.length !== 1) return false;
  d.担当者 = hit[0]; d.担当者の元 = `配置表の担当（${H.版}）`; d.更新 = now();
  WS.filled.push(`${siteLabel(s)} ${k} → ${hit[0]}`);
  return true;
}
// 集計（画面・PDF・Excel）の前の確かめ：記入者ごとのファイルを読み終わったか・重複・消された日・現場の分からない日報
function sumWarnings() {
  const w = [];
  if (signedIn() && !WS.readOk) w.push("記入者ごとの日報をまだ読み終わっていません（共有 ✓ が出るまで待ってください）。この集計は足りない日があるかもしれません。");
  WS.dups.forEach(x => w.push("同じ現場・同じ日に2人の日報があります：" + x + "。片方しか集計されていません。どちらかを消してください。"));
  WS.unknown.forEach(x => w.push("現場が分からない日報があります：" + x + "（現場・名簿に無い現場）。集計に入っていません。"));
  WS.lost.forEach(x => w.push("ほかの端末で消された日報を、この端末からも外しました：" + x));
  return w;
}
const wRenderSum0 = renderSum;
renderSum = function (all) {
  wRenderSum0(all);
  const w = sumWarnings(); if (!w.length) return;
  const box = document.createElement("div"); box.className = "sumwarn"; box.innerHTML = `<b>⚠ 集計を使う前に確かめてください</b><ul>${w.map(x => `<li>${esc(x)}</li>`).join("")}</ul>`;
  $("sumBody").prepend(box);
};
for (const f of ["exportXlsx", "printDays"]) {
  const f0 = window[f];
  window[f] = function (...a) { const w = sumWarnings(); if (w.length && !confirm("集計を使う前に確かめてください：\n・" + w.join("\n・") + "\n\nこのまま出しますか？")) return; return f0.apply(this, a); };
}
// 現場ごとのファイルには、記入者のファイルへ移した日を書かない（名簿・単価・決まりは残る）
siteForShare = s => { const days = {}; Object.entries(s.日).forEach(([k, d]) => { if (!(k >= MOVE_FROM && WS.done[wKey(s.id, k)] && d.担当者)) days[k] = d; }); return { ...s, 日: days }; };
const wSync0 = sync;
sync = async function () { await wSync0(); await writerSync(); if (UI.tab === "sheet") refreshSheet(); if (UI.tab === "sum") renderSum(); };

/* ---------- 他の人の日報も直せる（2026-11-17 ユーザー：見るだけをやめる。同時に直すのは「使用中」の札と交代で防ぐ） ----------
   前は「他の人の日報は見るだけ・同じ現場に行った人は『この日報を直す』で直せる」だった。いまは だれの日報でも直せる。
   上に「○○さんが書いた日報です」と出すだけ。同時に直すのは、現場ごとの「使用中」の札（app.js）で止める。 */
const otherWriter = () => { const d = day(), s = site(); return !!(s && d && d.担当者 && myName() && nk(d.担当者) !== nk(myName())); };
// 新しく書く日は、自分が記入者
const wEnsure0 = ensureDay;
ensureDay = function () { const had = day(), d = wEnsure0(); if (d && !had && !d.担当者) d.担当者 = myName(); return d; };

/* ---------- 記入者を押す → その人の日報を見る（2026-11-17 ユーザー：人を選んでもその人の日報に行かなかった） ----------
   1) この日に、その人が記入者の日報 → 2) この日に、その人が作業員・外注に入っている日報 → 3) この日より前で、いちばん近い日のその人の日報。
   現場を選んであれば（人の窓の「この現場に行った人」）その現場を先に探す。 */
$("hTan").onclick = () => detailOpenPeople("view");
$("hTan").onkeydown = e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); detailOpenPeople("view"); } };
function wPersonDays(name) {
  const n = nk(name), out = [];
  for (const s of DB.現場) for (const [k, d] of Object.entries(s.日 || {})) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(k) || !d || !hasData(d)) continue;
    const w = nk(d.担当者) === n, inRow = [...(d.下 || []), ...(d.外注 || [])].some(r => r && r[0] && nk(r[0]) === n);
    if (w || inRow) out.push({ s, k, w });
  }
  return out;
}
detailPickPerson = function (name) {
  const person = detailMasterPeople().find(p => nk(p.name) === nk(name));
  if (!person) { toast("マスターに登録されている人を選んでください"); detailRenderPeople(); return; }
  if (detailPeopleState.purpose === "entry") {   // ▾：この日報の記入者を決める
    if (!site() || !canEdit()) return;
    $("hTan").value = person.name; $("hTan").dispatchEvent(new Event("change")); closeDlg($("dPeople")); return;
  }
  closeDlg($("dPeople"));
  const k = UI.day, pref = detailPeopleState.onlySite ? detailPeopleState.site : "", all = wPersonDays(person.name);
  const rank = x => (x.s.id === pref ? 0 : 2) + (x.w ? 0 : 1);
  const today = all.filter(x => x.k === k).sort((a, b) => rank(a) - rank(b));
  let hit = today[0], moved = false;
  if (!hit) { const before = all.filter(x => x.k < k && (!pref || x.s.id === pref)).sort((a, b) => b.k.localeCompare(a.k) || rank(a) - rank(b)); hit = before[0] || all.filter(x => x.k < k).sort((a, b) => b.k.localeCompare(a.k))[0]; moved = !!hit; }
  if (!hit) { toast(`${person.name} さんの日報は、${jpDate(k)} とそれより前にはありません`, 5000); return; }
  UI.detailWriter = person.name; UI.detailView = "site";
  detailSelectSite(hit.s.id, hit.k);
  if (moved) toast(`${person.name} さんの ${jpDate(k)} の日報はありません。いちばん近い ${jpDate(hit.k)}（${siteLabel(hit.s)}）を開きました`, 6000);
  else if (today.length > 1) toast(`${person.name} さんはこの日 ${today.length} つの日報に入っています。ほかは現場名から開けます`, 5000);
  else if (!hit.w) toast(`${person.name} さんが作業員に入っている日報です（記入者：${hit.s.日[hit.k].担当者 || "未記入"}）`, 5000);
};

/* ---------- 現場名を押すと「現場名を選ぶ」窓（2026-11-17 ユーザー：入れる欄をやめて窓に戻す。navigation.js の detailOpenSites） ---------- */
function paintWriterBar() {
  let bar = $("wrBar");
  if (!bar) { bar = document.createElement("div"); bar.id = "wrBar"; bar.className = "wrbar"; $("roBar").after(bar); }
  const d = day();
  if (!otherWriter()) { bar.hidden = true; return; }
  bar.hidden = false;
  bar.innerHTML = `✎ <b>${esc(d.担当者)}</b> さんが書いた日報です。直したときは、${esc(d.担当者)} さんにも知らせてください。`;
}
const wRS0 = refreshSheet;
refreshSheet = function () { wRS0(); paintWriterBar(); };
paintWriterBar();
