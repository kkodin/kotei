/* クロジカの CSV・常用日報（.json）を取り込む（2026-11-02 ユーザー：クロジカの CSV をダウンロードのフォルダに置くので、過去のものを日報に。
   取り込むボタン。現場名が間違っていることがあるので確かめられるように。前に作った常用日報（入力済みの塩浜）も取り込む）
   ・ファイルはいくつでも一度に選べる（クロジカの月ごとの CSV・同じ月を何度も落とした CSV・常用日報の .json）。重なりは1つにまとめる。
   ・クロジカの CSV（「⋮ → ダウンロード」。Shift-JIS か UTF-8・セル内改行あり）。列：開始日・開始時刻・終了日・終了時刻・場所・予定（＝現場名）・カテゴリー・内容・名前・設備。
     - 終日の予定は 0:00〜0:00 で出てくる → 終日。同じ予定（日・時刻・予定・内容が同じ）は1つにまとめ、名前の列の人を参加者に。
       はじめの人の日報（個人_<名前>_<月>.json）に入れる（記入者＝その人）。ほかの人は参加者。
     - 設備の列には人（作業員）・外注（「マエダ建設　５名」）・機械（㎥ が「?」に化ける）が混ざる → 名簿の人＝参加者、「○○　n名」＝外注、ほかは使用機械
       （自社機械稼働表の機械と ? ㎥ 空白を無視して同じなら、その名前にそろえる）。
     - 予定（現場名）の書き方がまちまち（大瀬その2／その２、浄明寺(夜勤)／（夜間）／(夜)、館町下水（本社）、会社（平沢雨水）、本社　西谷浄水場その2、夜勤　鶴間4号雨水、閉じていない「（夜間」）
       → 夜勤・夜間・昼・本社の印を外して寄せる。本社(○○)・会社(○○)・○○(本社)・本社　○○＝本社＋現場名。会社だけ＝本社。
   ・常用日報（kotei/jouyou の「保存」.json）：現場ごと・日ごとに1行。記入者＝その日の担当者。作業員は会社で分け、自社（ハシックス・名簿の人）は参加者、
     ほかの会社は外注（業者＝会社・人数・名前と役＝職長は職長、ほかは作業員）。使用機材（通勤車は除く）＝使用機械。上の表（作業予定｜完了報告）＝内容。
     常用日報そのもの（現場ごと）は日報のフォルダの 常用日報_<コード>.json に写して残す（常用日報の形の編集・集計で使う）。その現場の請負形態は「常用」にする。
   ・予定（現場名）ごとに、日報に入れる現場名を一覧で確かめる：一致／候補（確かめる）／一覧にない／現場ではない（有給・休み・明けなど。はじめは外す）。直せる・外せる。
   ・取り込んだ行は 行.取込（"kuro:…"／"jouyou:コード|日"）。もう一度取り込んでも重ならない。同じ人・同じ日・同じ現場にクロジカと常用日報の両方があれば1行にまとめる。
     その日に報告が無ければ「報告済み（○○から取り込み）」にする。
   ・ファイルを選ぶか、フォルダの Kurojica_calendar_….csv を全部読むか。月ごとの取り込み方（アプリの予定を優先 など）と書き込みは imp.js（配置表と共通）。 */
"use strict";
function kuroCsv(text) {
  const rows = []; let row = [], f = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true;
    else if (c === ",") { row.push(f); f = ""; }
    else if (c === "\n" || c === "\r") { if (c === "\r" && text[i + 1] === "\n") i++; row.push(f); rows.push(row); row = []; f = ""; }
    else f += c;
  }
  if (f !== "" || row.length) { row.push(f); rows.push(row); }
  return rows;
}
const kuroDay = v => { const m = /^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/.exec(String(v || "").trim()); return m ? m[1] + "-" + pad2(+m[2]) + "-" + pad2(+m[3]) : ""; };
const kuroTime = v => { const m = /(\d{1,2}):(\d{2})/.exec(String(v || "")); return m ? `${+m[1]}:${m[2]}` : ""; };
async function fileText(f) {
  const u8 = new Uint8Array(await f.arrayBuffer());
  if (u8[0] === 0xEF && u8[1] === 0xBB && u8[2] === 0xBF) return new TextDecoder("utf-8").decode(u8.subarray(3));
  try { return new TextDecoder("utf-8", { fatal: true }).decode(u8); } catch (e) { return new TextDecoder("shift_jis").decode(u8); }
}
function readKuroText(t) {
  const rows = kuroCsv(t), hd = (rows[0] || []).map(v => String(v).trim()), ix = k => hd.indexOf(k);
  const I = { d: ix("開始日"), t0: ix("開始時刻"), d1: ix("終了日"), t1: ix("終了時刻"), pl: ix("場所"), s: ix("予定"), c: ix("内容"), n: ix("名前"), eq: ix("設備") };
  if (I.d < 0 || I.s < 0) throw new Error("クロジカの CSV ではないようです（「開始日」「予定」の列がありません）");
  const at = (r, i) => i >= 0 ? String(r[i] || "").replace(/\r\n?/g, "\n").trim() : "";
  return rows.slice(1).map(r => {
    let t0 = kuroTime(at(r, I.t0)), t1 = kuroTime(at(r, I.t1));
    if (t0 === "0:00" && (t1 === "0:00" || !t1)) { t0 = ""; t1 = ""; }   // 終日は 0:00〜0:00 で出てくる
    return { 日: kuroDay(at(r, I.d)), 終日まで: kuroDay(at(r, I.d1)), 始め: t0, 終わり: t1, 場所: at(r, I.pl), 予定: at(r, I.s), 内容: at(r, I.c), 名前: at(r, I.n), 設備: at(r, I.eq) };
  }).filter(e => e.日 && e.予定);
}
// 予定の名前 → { base（現場名）, honsha（本社＋現場名・本社だけ）, night }
function kuroTitle(raw) {
  let t = String(raw || "").normalize("NFKC").replace(/[\s　]+/g, " ").trim();
  t = t.replace(/\s*残\s*[0-9.]+\s*[hH時間]*\s*$/, "").trim();
  let night = /夜勤|夜間|\(夜\)/.test(t), honsha = false;
  t = t.replace(/^(夜勤|夜間)\s+/, "").replace(/\((夜勤|夜間|昼勤|昼間|夜|昼)\)?\s*$/, "").replace(/\((夜勤|夜間|昼勤|昼間|夜|昼)\)/g, "").replace(/\s*(夜勤|夜間)$/, "").trim();
  let m;
  if ((m = /^(?:本社|会社)\s*\((.+?)\)?$/.exec(t))) { t = m[1]; honsha = true; }
  else if ((m = /^(?:本社|会社)\s+(.+)$/.exec(t))) { t = m[1]; honsha = true; }
  else if ((m = /^(.+?)\s*\(本社\)?$/.exec(t))) { t = m[1]; honsha = true; }
  else if (/^(本社|会社)$/.test(t)) { t = ""; honsha = true; }
  t = t.replace(/\((夜勤|夜間|昼勤|昼間|夜|昼)\)?\s*$/, "").replace(/\($/, "").trim();
  if (honsha && /^(昼から|午後から|午前|午後|午前中|AM|PM)$/i.test(t)) t = "";   // 本社(昼から) は本社だけ
  if ((m = /^(.+?)\((夜勤|夜間)/.exec(t))) { t = m[1].trim(); night = true; }
  return { base: t, honsha, night };
}
const RE_KURO_OFF = /安全協議会|工事部会|安全大会|健康診断|パトロール|朝礼|明け休み|^明け$|有給|有休|^休み|忌引|代休|振替|休暇|^残\s*[0-9.]+\s*[hH時間]*$/;
const kKey = v => nk(v).replace(/[ヶケヵが]/g, "ガ").replace(/([一二三四五六七八九])(丁目|工区|号)/g, (a, k, u) => ("一二三四五六七八九".indexOf(k) + 1) + u).replace(/[-‐－−―()（）]/g, "");
function kuroScore(names, t) {
  const a = kKey(t); if (!a) return 0;
  const ks = names.filter(Boolean).map(kKey);
  if (ks.includes(a)) return 3;
  return ks.some(k => k.length >= 2 && a.length >= 2 && (k.includes(a) || a.includes(k))) ? 2 : 0;
}
function kuroGuess(base, code) {
  if (code) { const st = siteOf("", code), ks = koteiSite("", code); if (st || ks) return { v: st ? siteName(st) : ks.名前, code, how: "一致" }; }
  const cand = [...S.sites.map(s => ({ v: siteName(s), code: s.コード || "", names: [s.名前, s.正式名] })),
    ...S.K.sites.map(s => ({ v: s.名前, code: s.コード || "", names: [s.名前, s.略称] }))];
  const sc = cand.map(c => [c, kuroScore(c.names, base)]).filter(x => x[1]);
  const best = Math.max(0, ...sc.map(x => x[1])), top = [...new Map(sc.filter(x => x[1] === best).map(([c]) => [nk(c.v), c])).values()];
  if (best === 3) return { v: top[0].v, code: top[0].code, how: "一致" };
  if (top.length) return { v: top[0].v, code: top[0].code, how: "候補", alt: top.slice(1).map(c => c.v) };
  return { v: base, code: "", how: "一覧にない" };
}
const findPerson = n => people().find(p => nk(p) === nk(n));
const personName = n => findPerson(n) || String(n).trim().replace(/\s+/g, "　");
const splitList = v => String(v || "").split(/[\n、,，]+/).map(x => x.replace(/^[\s・\-－]+|\s+$/g, "")).filter(Boolean);
// 機械の名前をそろえる（? ㎥ m3 空白・並び順を無視して、自社機械稼働表の機械と同じならその名前）
const machKey = v => nk(String(v).replace(/[?？]|㎥|m3/g, " ")).replace(/[ｔt]/g, "t").split("").sort().join("");
function machName2(v) {
  const k = machKey(v), hit = (S.M.機械 || []).find(m => machKey(m) === k);
  return hit || String(v).replace(/(\d)\s*[?？]/g, "$1㎥").trim();
}
// 設備の列の1つ → 人／外注／機械
function eqKind(v) {
  const t = String(v).normalize("NFKC").trim(), m = /^(.+?)\s*([0-9]+)\s*名$/.exec(t);
  if (m) return { 外注: { 業者: m[1].trim(), 人数: m[2] } };
  if (findPerson(t)) return { 人: findPerson(t) };
  return { 機械: machName2(v) };
}

// クロジカ：同じ予定をまとめる。何日にもわたる終日の予定は1日ずつに（31日まで）
function kuroGroup(rows) {
  const m = new Map();
  rows.forEach(e => {
    const days = [e.日];
    if (!e.始め && e.終日まで && e.終日まで > e.日) for (let d = dayAdd(e.日, 1); d <= e.終日まで && days.length < 31; d = dayAdd(d, 1)) days.push(d);
    days.forEach(k => {
      const key = [k, e.始め, e.終わり, e.予定, e.内容].join("\u0001");
      if (!m.has(key)) m.set(key, { ...e, 日: k, 名前: [], 設備: [], key, src: "kuro" });
      const g = m.get(key);
      splitList(e.名前).forEach(n => { if (!g.名前.some(v => nk(v) === nk(n))) g.名前.push(n); });
      splitList(e.設備).forEach(n => { if (!g.設備.some(v => nk(v) === nk(n))) g.設備.push(n); });
    });
  });
  return [...m.values()];
}
const kuroId = g => "kuro:" + [g.日, g.始め, g.終わり, nk(g.予定), nk(g.内容).slice(0, 60)].join("|");
function kuroRow(g, x) {
  const names = g.名前.map(personName), owner = names[0] || S.me.name, ppl = names.slice(1), gai = [], mach = [];
  g.設備.forEach(v => { const e = eqKind(v); if (e.人) ppl.push(e.人); else if (e.外注) gai.push(e.外注); else if (e.機械 && !mach.includes(e.機械)) mach.push(e.機械); });
  const r = { id: uid(), 行先: x.honsha ? "本社" : x.v, コード: x.honsha ? "" : x.code || "", 始め: g.始め, 終わり: g.終わり, 終日: !g.始め,
    作業: "", 作業行: [{ 予定: "", 報告: [g.内容, g.場所 ? "場所：" + g.場所 : ""].filter(Boolean).join("\n") }], 予定: { 人: [], 外注: [], 機械: [] },
    人: ppl.filter((n, i, a) => a.findIndex(v => nk(v) === nk(n)) === i && nk(n) !== nk(owner)), 外注: gai, 機械: mach,
    記入者: owner, 取込: kuroId(g), 更新: nowStr() };
  if (x.honsha && x.v) r.関連現場 = { 名前: x.v, コード: x.code || "" };
  if (g.night || (tmin(g.始め) != null && (tmin(g.始め) >= 17 * 60 || tmin(g.始め) < 5 * 60))) r.区分 = "夜";
  syncWork(r);
  return { owner, k: g.日, r };
}

// 常用日報（.json）：現場ごと・日ごと
const SELF_CO = /ハシックス|自社/;
function readJouyou(doc, file) {
  if (!doc || doc.形式 !== "常用日報" || !Array.isArray(doc.現場)) throw new Error(file + " は詳細日報の保存ファイルではないようです");
  return doc.現場.map(s => {
    const co = n => { const w = (s.作業員 || []).find(w => nk(w.名前) === nk(n)); return w && w.会社 || ""; };
    const days = Object.entries(s.日 || {}).filter(([k, d]) => d && ([...(d.下 || []), ...(d.外注 || [])].some(r => r && r[0]) || (d.上 || []).some(r => r && r.some(Boolean))));
    return { src: "jouyou", site: s, file, code: s.コード || "", name: s.名前, days: days.map(([k, d]) => ({ k, d, co })), from: days.length ? days.map(x => x[0]).sort()[0] : "", to: days.length ? days.map(x => x[0]).sort().pop() : "" };
  }).filter(j => j.days.length);
}
function jouyouRow(j, k, d, x) {
  const owner = personName(d.担当者 || j.site.担当者 || S.me.name), ppl = [], gai = new Map(), mach = [];
  let t0 = "", t1 = "";
  (d.外注 || []).forEach(row => { const n = String(row && row[0] || "").trim(); if (!n || nk(n) === nk(owner)) return;   // 外注の欄の人は外注（名簿にいる人でも）
    const key = row[12] || "（会社不明）"; if (!gai.has(key)) gai.set(key, { 業者: key, 人数: "0", 名前: [] }); const g = gai.get(key); g.名前.push({ 名前: n, 役: /職長/.test(row[3] || "") ? "職長" : "作業員" }); g.人数 = String(g.名前.length); });
  (d.下 || []).forEach(row => {
    const n = String(row && row[0] || "").trim(); if (n) {
      const c = (row.length >= 14 && row[13]) || co(n), self = SELF_CO.test(c) || (!c && findPerson(n));
      if (nk(n) === nk(owner)) { t0 = row[1] || t0; t1 = row[2] || t1; }
      else if (self) { if (!ppl.some(v => nk(v) === nk(personName(n)))) ppl.push(personName(n)); }
      else { const key = c || "（会社不明）"; if (!gai.has(key)) gai.set(key, { 業者: key, 人数: "0", 名前: [] }); const g = gai.get(key); g.名前.push({ 名前: n, 役: /職長/.test(row[3] || "") ? "職長" : "作業員" }); g.人数 = String(g.名前.length); }
    }
    const mc = String(row && row[5] || "").trim();
    if (mc && !/通勤/.test(mc)) { const v = machName2(mc + (row[6] ? " " + row[6] : "")); if (!mach.includes(v)) mach.push(v); }
  });
  function co(n) { return (j.days[0] && j.days[0].co(n)) || ""; }
  const L = (d.上 || []).filter(r => r && r.some(Boolean)).map(r => ({ 予定: [r[0], r[1], r[2]].filter(Boolean).join(" "), 報告: String(r[4] || "") }));
  const r = { id: uid(), 行先: x.v, コード: x.code || j.code || "", 始め: normTime(t0) || j.site.始業 || "8:00", 終わり: normTime(t1) || j.site.終業 || "17:00",
    作業: "", 作業行: L.length ? L : [{ 予定: "", 報告: "" }], 予定: { 人: [], 外注: [], 機械: [] }, 人: ppl, 外注: [...gai.values()], 機械: mach,
    記入者: owner, 取込: `jouyou:${j.code || nk(j.name)}|${k}`, 更新: nowStr() };
  if (d.備考) r.作業行.push({ 予定: "", 報告: "備考：" + d.備考 });
  if (tmin(r.始め) != null && (tmin(r.始め) >= 17 * 60 || tmin(r.始め) < 5 * 60)) r.区分 = "夜";
  syncWork(r);
  return { owner, k, r };
}

/* ---------- 取り込みの窓 ---------- */
let KURO = null;   // { files, groups, jous, map: Map(予定の名前 → { raw, base, honsha, n, v, code, how, on }), from, to }
function kuroOpen() { KURO = null; kuroDraw(); }
// フォルダから：その中の Kurojica_calendar_<期間>( (n)).csv を全部（2026-11-02 ユーザー：ファイルか、フォルダにあるものを全部か選ぶ）
const RE_KURO_FILE = /^Kurojica_calendar_.*\.csv$/i;
async function kuroFolder() {
  if (window.showDirectoryPicker) {
    try {
      const dir = await showDirectoryPicker({ id: "nippou_kuro_dir", startIn: "downloads" }), out = [];
      for await (const h of dir.values()) if (h.kind === "file" && RE_KURO_FILE.test(h.name)) out.push(await h.getFile());
      return out;
    } catch (e) { if (e.name === "AbortError") return null; }
  }
  return new Promise(res => { const i = document.createElement("input"); i.type = "file"; i.webkitdirectory = true; i.onchange = () => res([...i.files].filter(f => RE_KURO_FILE.test(f.name))); i.click(); });
}
async function kuroPick(folder) {
  let fs = [];
  if (folder === true) {
    fs = await kuroFolder(); if (!fs) return;
    if (!fs.length) { toast("そのフォルダに Kurojica_calendar_….csv がありません", 6000); return; }
  } else {
    try {
      if (window.showOpenFilePicker) fs = await Promise.all((await showOpenFilePicker({ id: "nippou_kuro", startIn: "downloads", multiple: true,
        types: [{ description: "クロジカの CSV・詳細日報の .json", accept: { "text/csv": [".csv"], "application/json": [".json"] } }] })).map(h => h.getFile()));
    } catch (e) { if (e.name === "AbortError") return; }
    if (!fs.length) {
      fs = await new Promise(res => { const i = document.createElement("input"); i.type = "file"; i.multiple = true; i.accept = ".csv,.json,text/csv,application/json"; i.onchange = () => res([...i.files]); i.click(); });
      if (!fs.length) return;
    }
  }
  try {
    const rows = [], jous = [], names = [];
    for (const f of fs) {
      const t = await fileText(f); names.push(f.name);
      if (/\.json$/i.test(f.name) || /^\s*\{/.test(t)) {
        const js = readJouyou(JSON.parse(t), f.name);
        js.forEach(j => { const i = jous.findIndex(o => (o.code || nk(o.name)) === (j.code || nk(j.name))); if (i < 0) jous.push(j); else if ((j.site.設定更新 || "") + j.days.length > (jous[i].site.設定更新 || "") + jous[i].days.length) jous[i] = j; });
      } else rows.push(...readKuroText(t));
    }
    const groups = kuroGroup(rows), map = new Map();
    groups.forEach(g => {
      const t = kuroTitle(g.予定), key = (t.honsha ? "本社:" : "") + kKey(t.base);
      g.mk = key; g.night = t.night;
      if (!map.has(key)) {
        const off = RE_KURO_OFF.test(t.base) || RE_KURO_OFF.test(g.予定);
        const gs = t.honsha && !t.base ? { v: "", how: "本社" } : off ? { v: t.base, how: "現場ではない" } : kuroGuess(t.base);
        map.set(key, { raw: g.予定, base: t.base, honsha: t.honsha, n: 0, ...gs, on: !off });
      }
      map.get(key).n++;
    });
    jous.forEach(j => { const key = "jouyou:" + (j.code || nk(j.name)); j.mk = key; map.set(key, { raw: "詳細日報：" + j.name, base: j.name, honsha: false, n: j.days.length, ...kuroGuess(j.name, j.code), on: true, jouyou: true }); });
    const ks = [...groups.map(g => g.日), ...jous.flatMap(j => [j.from, j.to])].filter(Boolean).sort();
    if (!ks.length) throw new Error("取り込むものがありません");
    KURO = { files: names, groups, jous, map, from: ks[0], to: ks[ks.length - 1], all: [ks[0], ks[ks.length - 1]], pol: KURO ? KURO.pol : {}, def: KURO ? KURO.def : "" };
  } catch (e) { toast("読めませんでした：" + e.message, 8000); return; }
  kuroDraw();
}
function kuroItems() {
  const K = KURO, out = [];
  K.groups.forEach(g => { if (g.日 < K.from || g.日 > K.to) return; const x = K.map.get(g.mk); if (x && x.on && (x.v || x.honsha)) out.push(kuroRow(g, x)); });
  K.jous.forEach(j => { const x = K.map.get(j.mk); if (!x || !x.on || !x.v) return; j.days.forEach(({ k, d }) => { if (k >= K.from && k <= K.to) out.push(jouyouRow(j, k, d, x)); }); });
  return out;
}
function kuroDraw() {
  const K = KURO, cls = { 一致: "ok", 候補: "maybe", 一覧にない: "ng", 現場ではない: "off", 本社: "ok" };
  if (!K) {
    dlg(`<div class="dh"><h2>クロジカ・詳細日報から取り込む</h2><button type="button" data-close aria-label="閉じる">×</button></div>
      <ol class="kuro-steps"><li><a href="https://app.aipo.com/" target="_blank" rel="noopener">クロジカ</a>で取り込みたい月を出し、「⋮ → ダウンロード」で CSV を落とす（ダウンロードのフォルダに入ります。何か月分でも）</li>
      <li>どちらかを押す<br>・<b>フォルダから全部</b>：選んだフォルダ（ダウンロードなど）の <code>Kurojica_calendar_….csv</code> を全部読む（同じ月を何度も落とした「(1)」などのファイルも。重なりは1つにまとめます）<br>
        ・<b>ファイルを選ぶ</b>：CSV を選ぶ（Ctrl か Shift を押しながら、いくつでも。詳細日報の保存ファイル .json も一緒に選べます）</li>
      <li>現場名と、月ごとの取り込み方（その日にアプリの予定があるとき）を確かめてから「取り込む」</li></ol>
      <div class="btns"><button data-close>閉じる</button><button id="kFile">ファイルを選ぶ</button><button class="pri" id="kDir">フォルダから全部</button></div>`, "wide keep");
    $("kFile").onclick = () => kuroPick(false); $("kDir").onclick = () => kuroPick(true); return;
  }
  const items = kuroItems(), ppl = new Set(items.map(i => nk(i.owner)));
  const list = [...K.map.entries()].sort((a, b) => (b[1].jouyou ? 1 : 0) - (a[1].jouyou ? 1 : 0) || (a[1].how === "一致") - (b[1].how === "一致") || b[1].n - a[1].n);
  dlg(`<div class="dh"><h2>クロジカ・詳細日報から取り込む</h2><button type="button" data-close aria-label="閉じる">×</button></div>
    <p class="kuro-file">${K.files.length}ファイル（${esc(K.all[0])}〜${esc(K.all[1])}）：クロジカ ${K.groups.length}件${K.jous.length ? `・詳細日報 ${K.jous.map(j => `${esc(j.name)} ${j.days.length}日`).join("・")}` : ""}</p>
    <div class="dt">取り込む期間 <input type="date" id="kFrom" value="${K.from}"> 〜 <input type="date" id="kTo" value="${K.to}"></div>
    <p class="pinfo">予定（現場名）ごとに、日報に入れる現場名を確かめてください。<b class="maybe">候補</b>・<b class="ng">一覧にない</b>は打ち直せます（打つと候補）。チェックを外すと取り込みません。</p>
    <div class="gwrap kuro-wrap"><table class="list kuro"><tr><th>取込</th><th>クロジカの予定・詳細日報</th><th>件数</th><th>日報の現場名</th><th></th></tr>
    ${list.map(([key, x], i) => `<tr${x.jouyou ? ` class="jy"` : ""}><td><input type="checkbox" data-kon="${i}"${x.on ? " checked" : ""}></td><td>${esc(x.raw)}</td><td class="num">${x.n}</td>
      <td>${x.honsha ? `<span class="pinfo">本社${x.v || x.how !== "本社" ? "＋" : "だけ"}</span>` : ""}${x.honsha && x.how === "本社" && !x.v ? "" : `<div class="pick"><input data-kv="${i}" value="${esc(x.v)}" autocomplete="off"></div>`}${x.alt && x.alt.length ? `<small class="pinfo">ほかの候補：${esc(x.alt.join("、"))}</small>` : ""}</td>
      <td><span class="kh ${cls[x.how] || ""}">${esc(x.how)}</span></td></tr>`).join("")}</table></div>
    ${polHTML(K, items)}
    <div class="btns"><button type="button" id="kFile">別のファイルを選ぶ</button><button type="button" id="kDir">フォルダから全部</button><span style="flex:1"></span><button data-close>閉じる</button><button class="pri" id="kGo"${items.length ? "" : " disabled"}>${items.length}件を取り込む（${ppl.size}人）</button></div>`, "wide keep");
  const B = $("dlgBody");
  $("kFile").onclick = () => kuroPick(false); $("kDir").onclick = () => kuroPick(true);
  bindPol(K, kuroDraw);
  $("kFrom").onchange = () => { K.from = $("kFrom").value || K.from; kuroDraw(); };
  $("kTo").onchange = () => { K.to = $("kTo").value || K.to; kuroDraw(); };
  B.querySelectorAll("[data-kon]").forEach(el => el.onchange = () => { list[+el.dataset.kon][1].on = el.checked; kuroDraw(); });
  B.querySelectorAll("[data-kv]").forEach(el => {
    const x = list[+el.dataset.kv][1];
    const set = (v, code) => { const st = siteOf(v, code), ks = koteiSite(v, code); x.v = v; x.code = code || (st && st.コード) || (ks && ks.コード) || ""; x.how = st || ks ? "一致" : v ? "一覧にない" : x.how; const sc = B.querySelector(".kuro-wrap").scrollTop; kuroDraw(); $("dlgBody").querySelector(".kuro-wrap").scrollTop = sc; };
    bindPick(el, () => placeItems().filter(y => y.grp !== "本社など"), y => { el.value = y.v; set(y.v, y.コード); });
    el.addEventListener("change", () => { const v = el.value.trim(); if (v !== x.v) set(v, ""); });
  });
  $("kGo").onclick = () => kuroImport(items).catch(e => { toast("取り込めませんでした：" + e.message, 8000); kuroDraw(); });
}
async function kuroImport(items) {
  const K = KURO;
  $("kGo").disabled = true;
  const st = await importItems(items, { srcs: ["kuro", "jouyou"], pol: K.pol, label: "クロジカ", how: r => srcOf(r) === "jouyou" ? "詳細日報" : "クロジカ", prog: impProg });
  if (!st) { kuroDraw(); return; }
  // 常用日報そのものを日報のフォルダに写して残す・請負形態を「常用」に
  for (const j of K.jous) {
    const x = K.map.get(j.mk); if (!x || !x.on || !x.v) continue;
    const name = `常用日報_${j.code || nk(j.name)}.json`, cur = await nippouStore.get(name).catch(() => null);
    await nippouStore.put(name, { 形式: "常用日報（現場）", 版: 1, 元: j.file, 取り込み: nowStr(), だれ: S.me.name, 現場: j.site }, cur ? cur.eTag : null).catch(e => toast("常用日報の写しを保存できませんでした：" + e.message, 6000));
    if (formOf(x.v, x.code) !== "常用") await setForm(x.v, x.code, "常用");
  }
  KURO = null; render();
  importDone(st);
}
