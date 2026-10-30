/* 工事日報 Vr002（2026-10-30）
   全体工程表の工程を「予定」として日報に入れ、セルは Excel と同じ動きで直す。
   Vr002：入力欄を6種類（施工箇所〜作業内容／社員／作業員／外注／材料／使用機械）に分けた。上の ◀▶ は前の日・次の日。
   ・画面は GitHub Pages（kotei/nippou/）、データは会社の OneDrive。工程は 工程データ（読むだけ）、日報は 工程データ\日報\ に書く。
   ・自社機械稼働表の予約を「使用機材」の予定に、前の日の使用機材（リース品）を続きとして入れる。
   ・変えた作業・人数・機械は予定と比べて色で出す。田村さんは「配置まとめ」で全現場の予定・報告を見て配置表に写す。
   ?demo を付けて開くと、サインインせずこのブラウザの中だけで試せる（OneDrive には書かない）。 */
"use strict";

const CFG = window.NIPPOU_CONFIG;
const SCOPES = ["User.Read", "Files.ReadWrite.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const DEMO = new URLSearchParams(location.search).has("demo");
const POLL_MS = 30000;
const NDIR = "日報";                 // 工程データ の中の日報のフォルダ
const TIME = { 昼: ["8:00", "17:00"], 夜: ["21:00", "6:00"] };
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------- 日付（"YYYY-MM-DD" の文字列で持つ） ---------- */
const pad2 = n => String(n).padStart(2, "0");
const keyOf = d => d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
const dateOf = k => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
const dayAdd = (k, n) => { const d = dateOf(k); d.setDate(d.getDate() + n); return keyOf(d); };
const dayDow = k => dateOf(k).getDay();
const todayKey = () => keyOf(new Date());
const WD = "日月火水木金土";
const monthOf = k => k.slice(0, 7);
const monthAdd = (m, n) => { const d = new Date(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1); return d.getFullYear() + "-" + pad2(d.getMonth() + 1); };
const daysIn = m => { const n = new Date(+m.slice(0, 4), +m.slice(5, 7), 0).getDate(); return Array.from({ length: n }, (_, i) => m + "-" + pad2(i + 1)); };
// 令和8年9月2日(水)
const wareki = k => { const d = dateOf(k), y = d.getFullYear(); return (y >= 2019 ? "令和" + (y - 2018 === 1 ? "元" : y - 2018) : "平成" + (y - 1988)) + "年" + (d.getMonth() + 1) + "月" + d.getDate() + "日(" + WD[d.getDay()] + ")"; };
const tabName = k => (+k.slice(8)) + "日(" + WD[dayDow(k)] + ")";
const md = k => (+k.slice(5, 7)) + "/" + (+k.slice(8));

const JP_HOL = {};
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
    if (h[n] && !h[mid] && dayDow(mid) !== 0 && mid.slice(0, 4) === String(y)) h[mid] = "国民の休日"; });
  Object.keys(h).sort().forEach(x => { if (dayDow(x) !== 0) return; let n = dayAdd(x, 1); while (h[n]) n = dayAdd(n, 1); if (n.slice(0, 4) === String(y)) h[n] = "振替休日"; });
  return JP_HOL[y] = h;
}
const holidayOf = k => jpHolidays(+k.slice(0, 4))[k] || "";

/* 全角の数字・記号を半角に（数字の欄はIMEが入ったままでも困らないように。入力欄のIMEルール） */
const hankaku = s => String(s ?? "").replace(/[０-９．，：－＋]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));

/* ---------- 保存先：OneDrive（Graph）／試し ---------- */
let msalApp = null, me = { name: "" };
async function token() {
  try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken; }
  catch (e) { if (e instanceof msal.InteractionRequiredAuthError) await msalApp.acquireTokenRedirect({ scopes: SCOPES }); throw e; }
}
async function g(path, opt = {}) {
  return fetch(GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + await token(), ...(opt.headers || {}) } });
}
function httpErr(res, what) { const e = new Error(what + " " + res.status); e.status = res.status; return e; }
const enc = p => p.split("/").map(encodeURIComponent).join("/");
// where：kotei（工程データ）／kikai（稼働データ）。sub：中のフォルダ（日報）
function GraphFolder(where, sub) {
  const F = CFG[where];
  const base = () => `/drives/${F.driveId}/items/${F.itemId}`;
  const p = name => base() + ":/" + enc((sub ? sub + "/" : "") + name);
  return {
    async list() {
      const out = [];
      let url = sub ? base() + ":/" + enc(sub) + ":/children?$top=999" : base() + "/children?$top=999";
      while (url) {
        const res = await g(url);
        if (res.status === 404 && sub) return [];
        if (!res.ok) throw httpErr(res, "一覧");
        const j = await res.json();
        j.value.forEach(x => { if (x.file) out.push({ name: x.name, eTag: x.eTag }); });
        url = j["@odata.nextLink"] ? j["@odata.nextLink"].replace(GRAPH, "") : null;
      }
      return out;
    },
    async get(name) {
      const res = await g(p(name));
      if (res.status === 404) return null;
      if (!res.ok) throw httpErr(res, "読み込み");
      const meta = await res.json();
      const dl = meta["@microsoft.graph.downloadUrl"];
      const body = dl ? await fetch(dl, { cache: "no-store" }) : await g(p(name) + ":/content");
      if (!body.ok) throw httpErr(body, "中身の読み込み");
      return { data: await body.json(), eTag: meta.eTag };
    },
    async put(name, data, eTag) {
      const url = p(name) + ":/content" + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
      const send = () => g(url, { method: "PUT", headers: { "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: JSON.stringify(data) });
      let res = await send();
      if (res.status === 404 && sub && !eTag) {
        // 日報 フォルダがまだ無い → 作ってからもう一度
        await g(base() + "/children", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: sub, folder: {}, "@microsoft.graph.conflictBehavior": "replace" }) });
        res = await send();
      }
      if (!res.ok) throw httpErr(res, "保存");
      return { eTag: (await res.json()).eTag };
    }
  };
}
function DemoFolder(prefix) {
  const key = DEMO_KEY;
  const all = () => { try { return JSON.parse(localStorage.getItem(key)) || {}; } catch (e) { return {}; } };
  const save = a => { try { localStorage.setItem(key, JSON.stringify(a)); } catch (e) { } };
  return {
    async list() { const a = all(); return Object.keys(a).filter(n => n.startsWith(prefix)).map(n => ({ name: n.slice(prefix.length), eTag: a[n].eTag })); },
    async get(name) { const a = all(), x = a[prefix + name]; return x ? { data: JSON.parse(JSON.stringify(x.data)), eTag: x.eTag } : null; },
    async put(name, data, eTag) {
      const a = all(), cur = a[prefix + name];
      if (eTag ? (!cur || cur.eTag !== eTag) : cur) { const e = new Error("保存 " + (eTag ? 412 : 409)); e.status = eTag ? 412 : 409; throw e; }
      const t = "d" + Date.now() + Math.random().toString(36).slice(2, 6);
      a[prefix + name] = { data: JSON.parse(JSON.stringify(data)), eTag: t }; save(a);
      return { eTag: t };
    }
  };
}
const DEMO_KEY = "nippou_demo_v2";
const koteiStore = DEMO ? DemoFolder("工程/") : GraphFolder("kotei");
const nippouStore = DEMO ? DemoFolder("日報/") : GraphFolder("kotei", NDIR);
const kikaiStore = DEMO ? DemoFolder("機械/") : GraphFolder("kikai");

/* ---------- 状態 ---------- */
const S = {
  K: {},          // 工程データ：name → { doc, eTag }
  kikai: null,    // 稼働データ.json の中身
  N: {},          // 日報：name → { doc, eTag, pending:[fn], saving, timer, err }
  site: "", day: "", tab: "sheet",
  weather: {},    // "YYYY-MM-DD" → "晴"
  undo: {}, redo: {}   // 日報ファイル名＋日 → [{before, after}]
};
const stamp = d => { d.更新 = { だれ: me.name, いつ: new Date().toLocaleString("ja-JP") }; };

/* ---------- 全体工程表のデータの読み方（全体工程表 app.js と同じ決まり） ---------- */
const nm = s => s.略称 || s.名前;
const kdoc = n => S.K[n] && S.K[n].doc;
const sites = () => Object.keys(S.K).filter(n => n.startsWith("現場_")).map(kdoc).filter(s => s && !s.消した)
  .sort((a, b) => (a.並び ?? 0) - (b.並び ?? 0) || String(nm(a)).localeCompare(nm(b), "ja"));
const siteById = id => sites().find(s => s.id === id);
const M = () => kdoc("マスタ.json") || { 社員: [], 従業員: [], 現場: [] };
const isRest = k => { const f = kdoc("休日.json"); return !!(f && (f.休み || []).includes(k)); };
function ruleWorks(ln, k) {
  if (holidayOf(k)) return ln.祝 === true;
  const w = dayDow(k);
  if (w === 0) return ln.日 === true;
  if (w === 6) return ln.土 !== false;
  return true;
}
function barWorks(ln, b, k) {
  if (isRest(k)) return false;
  if (b && (b.休み || []).includes(k)) return false;
  if (b && (b.出勤 || []).includes(k)) return true;
  return !!ln && ruleWorks(ln, k);
}
function laneKubun(ln) { return !ln ? "昼" : ln.区分 || (/夜/.test(ln.名前 || "") ? "夜" : "昼"); }
function laneWork(ln) { return !ln ? "" : ln.区分 ? (ln.名前 || "") : String(ln.名前 || "").replace(/昼間|夜間/g, "").trim(); }
function parseHead(t) {
  const s = hankaku(t).replace(/[＊*]/g, "+").replace(/[名人\s　]/g, "");
  if (!s) return null;
  const parts = s.split("+").filter(x => x !== "");
  if (!parts.length || parts.some(x => !/^\d+$/.test(x))) return null;
  let n = parts.map(Number);
  if (n.length > 1 && n.some(x => x)) n = n.filter(x => x);
  return { parts: n, total: n.reduce((a, x) => a + x, 0) };
}
const headOwn = b => b.人数内訳 || (b.人数 !== "" && b.人数 != null ? String(b.人数) : "");
const headTxt = (b, ln) => headOwn(b) || (ln && ln.人数) || "";
const headTxtAt = (b, k, ln) => (b.人数日 && b.人数日[k] != null && b.人数日[k] !== "") ? b.人数日[k] : headTxt(b, ln);
const fmtHead = t => { const p = parseHead(t); return p ? p.parts.map(x => x + "名").join("＋") : ""; };
const isGai = ln => !!ln && ln.種別 === "外注";
const gaiName = ln => isGai(ln) ? (ln.業者 || "外注") : "";

/* 自社機械稼働表の予約がその日に使うか（自社機械稼働表 app.js の works と同じ） */
function kikaiWorks(r, k) {
  if (k < r.開始 || k > r.終了) return false;
  if ((r.未使用 || []).includes(k)) return false;
  if ((r.出勤 || []).includes(k)) return true;
  if ((r.休み || []).includes(k)) return false;
  if (holidayOf(k) || dayDow(k) === 0) return !!r.日祝;
  if (dayDow(k) === 6) return !!r.土;
  return true;
}

/* ---------- 日報ファイル ---------- */
// 日報_<現場id>_<YYYY-MM>.json：{ 形式:"工事日報", 版:2, 現場, 月, 日:{ "YYYY-MM-DD": 日 } }
// 日：{ 天候, 表:{ 基本, 社員, 作業員, 外注, 材料, 機械 }, 予定表:（同じ形）, 予定日時, 報告:{だれ,いつ}|null, 確認:{だれ,いつ}|null }
//   表の中は SECS の列の順の配列の配列。後ろの空の行は保存しない。
// Vr001 の 行／予定（Excel の A〜M の13列）は読むときに 表 に直す（fromV1）。行・予定は消さずに残す（データの形は足すだけ）
const fileName = (siteId, m) => `日報_${siteId}_${m}.json`;
const nfile = (siteId, m) => S.N[fileName(siteId, m)];
const storedDay = (siteId, k) => { const f = nfile(siteId, monthOf(k)); return f && f.doc && f.doc.日 && f.doc.日[k] || null; };

/* 入力欄の6種類（2026-10-30 ユーザー指定：施工箇所〜作業内容／社員1行1名／作業員1行1名／外注1行1業者／材料／使用機械1行1つ） */
const SECS = [
  { key: "基本", label: "施工", title: "施工箇所・工種・作業時間・作業内容", min: 3, cols: [
    { t: "施工箇所", w: 130 }, { t: "工種", w: 130 }, { t: "開始", w: 58, align: "c", kind: "time" }, { t: "終了", w: 58, align: "c", kind: "time" }, { t: "作業内容", w: 384 }] },
  { key: "社員", label: "社員", title: "社員（1行に1名）", min: 2, unit: "名", cols: [{ t: "氏名", w: 160, kind: "name" }, { t: "作業内容", w: 600 }] },
  { key: "作業員", label: "作業員", title: "作業員（1行に1名）", min: 2, unit: "名", cols: [{ t: "氏名", w: 160, kind: "name" }, { t: "作業内容", w: 600 }] },
  { key: "外注", label: "外注", title: "外注業者（1行に1業者）", min: 2, unit: "名", cols: [{ t: "業者名", w: 220 }, { t: "人数", w: 60, align: "r", kind: "num" }, { t: "作業内容", w: 480 }] },
  { key: "材料", label: "材料", title: "材料", min: 2, cols: [{ t: "材料名", w: 180 }, { t: "詳細", w: 200 }, { t: "数量", w: 80, align: "r" }, { t: "業者名", w: 150 }, { t: "備考", w: 150 }] },
  { key: "機械", label: "機械", title: "使用機械（1行に1つ）", min: 2, unit: "台", cols: [{ t: "機械名", w: 180 }, { t: "能力等", w: 180 }, { t: "数量", w: 60, align: "r", kind: "num" }, { t: "自社／リース会社名", w: 340 }] }
];
SECS.forEach(s => s.cols.forEach(c => { c.cls = (c.align === "r" ? "r" : c.align === "c" ? "c" : "") + (c.kind ? " " + c.kind : ""); c.letter = c.t; }));
const SK = SECS.map(s => s.key);
const secOf = k => SECS.find(s => s.key === k);
const emptyT = () => Object.fromEntries(SK.map(k => [k, []]));
const blankRow = k => Array(secOf(k).cols.length).fill("");
const cleanRow = (k, r) => { const x = blankRow(k); (r || []).forEach((v, i) => { if (i < x.length) x[i] = v ?? ""; }); return x; };
const trimRows = rows => { const out = rows.map(r => r.slice()); while (out.length && out[out.length - 1].every(v => v === "")) out.pop(); return out; };
const normT = t => Object.fromEntries(SK.map(k => [k, trimRows(((t || {})[k] || []).map(r => cleanRow(k, r)))]));
const cloneT = t => JSON.parse(JSON.stringify(t));
const hasData = t => !!t && SK.some(k => (t[k] || []).some(r => r.some(v => v)));
const dayT = d => d.表 ? normT(d.表) : fromV1(d.行);
const dayP = d => d.予定表 ? normT(d.予定表) : d.予定 ? fromV1(d.予定) : null;
const firstLine = t => String(t || "").split(/\r?\n/)[0].trim();
const worksOf = s => String(s || "").split(/[・、,，]/).map(x => x.trim()).filter(Boolean);

/* Vr001 の Excel の形（A 施工箇所・B 工種・C〜E 時間・F〜H 作業内容・I〜M 使用機材）を6種類に直す */
function fromV1(rows) {
  const t = emptyT();
  let work = "";
  for (const r0 of rows || []) {
    const r = Array.from({ length: 13 }, (_, i) => (r0 || [])[i] ?? "");
    const f = String(r[5]).trim();
    if (r[0] || r[1] || r[2] || r[4]) { if (r[1]) work = r[1]; t.基本.push([r[0], r[1], r[2], r[4], [r[5], r[6], r[7]].filter(Boolean).join(" ")]); }
    else if (f) {
      let m;
      if ((m = f.match(/^社員[\s　:：]*(.+)$/))) m[1].split(/[・、,，]+/).map(x => x.trim()).filter(Boolean).forEach(n => t.社員.push([n, work]));
      else if ((m = hankaku(f).match(/^(.*?)[\s　]*((?:\d+\s*[名人]\s*[＋+]?\s*)+)$/))) {
        const label = m[1].trim(), n = (m[2].match(/\d+/g) || []).reduce((a, x) => a + Number(x), 0);
        if (!label || label === "作業員") for (let i = 0; i < n; i++) t.作業員.push(["", work]);
        else t.外注.push([label, String(n), work]);
      } else t.基本.push(["", "", "", "", f]);
    }
    if (r[8]) t.機械.push([r[8], r[9], r[10], [r[12], r[11]].filter(Boolean).join(" ")]);
  }
  return t;
}

function mutate(siteId, m, fn, opt = {}) {
  const name = fileName(siteId, m);
  if (!S.N[name]) S.N[name] = { doc: { 形式: "工事日報", 版: 2, 現場: siteId, 月: m, 日: {} }, eTag: null, pending: [], saving: false, timer: null, err: null };
  const f = S.N[name];
  fn(f.doc); f.doc.版 = 2; stamp(f.doc);
  f.pending.push(fn);
  scheduleSave(name);
  if (!opt.quiet) render();
}
function scheduleSave(name, ms = 600) {
  const f = S.N[name];
  clearTimeout(f.timer);
  f.timer = setTimeout(() => saveFile(name), ms);
  showSync();
}
async function saveFile(name) {
  const f = S.N[name];
  if (f.saving) { scheduleSave(name, 500); return; }
  if (!f.pending.length) return;
  f.saving = true; showSync();
  try {
    for (let tries = 0; tries < 5; tries++) {
      const n = f.pending.length;
      try {
        const r = await nippouStore.put(name, f.doc, f.eTag);
        f.eTag = r.eTag; f.pending.splice(0, n); f.err = null;
        break;
      } catch (e) {
        if (e.status !== 412 && e.status !== 409) throw e;
        // ほかの人（田村さんの確認など）が先に保存した → 最新を読んで、自分の変更を重ね直す
        const fresh = await nippouStore.get(name);
        const d = fresh ? fresh.data : f.doc;
        f.pending.forEach(fn => fn(d)); d.版 = 2; stamp(d);
        f.doc = d; f.eTag = fresh ? fresh.eTag : null;
        toast("ほかの人の変更を読み込んで、重ねて保存しました");
        render();
      }
    }
  } catch (e) {
    f.err = e.message; console.error(e);
    toast("保存できませんでした：" + e.message + "（自動でやり直します）");
    setTimeout(() => scheduleSave(name, 0), 10000);
  } finally {
    f.saving = false; showSync();
    if (f.pending.length && !f.err) scheduleSave(name, 300);
  }
}
function showSync() {
  const fs = Object.values(S.N);
  const busy = fs.some(f => f.pending.length || f.saving), err = fs.some(f => f.err);
  const el = $("sync");
  el.className = "sync " + (err ? "err" : busy ? "busy" : "ok");
  el.textContent = err ? "保存できていません" : busy ? "保存中…" : "保存済み";
}

/* ---------- 工程から予定を作る ----------
   基本：工種ごとに 施工箇所（前の日に同じ工種で書いた場所）・工種・作業時間（昼 8:00〜17:00／夜 21:00〜6:00）・作業内容
   社員：バーの社員を1行1名（作業内容＝工種）。作業員：段の人数の分だけ1行1名（氏名は前の日に同じ工種をした作業員から）
   外注：外注の段は 業者名・人数。機械：自社機械稼働表の予約＋前の日の機械のうち自社以外（リース品の続き）。材料は予定なし */
function genPlan(siteId, k) {
  const s = siteById(siteId), t = emptyT();
  if (!s) return t;
  const prev = prevT(siteId, k) || emptyT();
  const placeOf = w => { const r = prev.基本.find(x => x[1] === w && x[0]); return r ? r[0] : ""; };
  const used = new Set();
  for (const ln of s.段 || []) {
    const bars = (s.バー || []).filter(b => b.段 === ln.id && b.開始 <= k && b.終了 >= k && barWorks(ln, b, k)).sort((a, b) => a.開始.localeCompare(b.開始));
    for (const b of bars) {
      const kb = laneKubun(ln), work = laneWork(ln) || b.作業 || "";
      const detail = work === (b.作業 || "") ? firstLine(b.メモ) : (b.作業 || "");
      const tm = TIME[kb] || TIME.昼;
      t.基本.push([placeOf(work), work, tm[0], tm[1], detail]);
      for (const n of b.社員 || []) {
        const hit = t.社員.find(x => x[0] === n);
        if (!hit) t.社員.push([n, work]); else if (!worksOf(hit[1]).includes(work)) hit[1] += "・" + work;
      }
      const h = parseHead(headTxtAt(b, k, ln)), n = h ? h.total : 0;
      if (isGai(ln)) t.外注.push([gaiName(ln), n ? String(n) : "", work]);
      else for (let i = 0; i < n; i++) {
        const j = prev.作業員.findIndex((x, xi) => !used.has(xi) && x[0] && worksOf(x[1]).includes(work));
        if (j >= 0) used.add(j);
        t.作業員.push([j >= 0 ? prev.作業員[j][0] : "", work]);
      }
    }
  }
  // 機械：自社機械稼働表の予約
  const kk = S.kikai, mine = [];
  if (kk && kk.予約) {
    for (const r of kk.予約) {
      if (r.現場 !== nm(s) && r.現場 !== s.名前) continue;
      if (r.状態 === "仮" || !kikaiWorks(r, k)) continue;
      const m = (kk.機械 || []).find(x => x.id === r.機械);
      if (!m) continue;
      const hit = mine.find(x => x.name === m.名前 && x.spec === (m.規格 || "") && x.kb === (r.区分 || "昼"));
      if (hit) hit.n++; else mine.push({ name: m.名前, spec: m.規格 || "", n: 1, kb: r.区分 || "昼" });
    }
  }
  mine.forEach(x => t.機械.push([x.name, x.spec, String(x.n), "自社" + (x.kb === "夜" ? "（夜）" : "")]));
  prev.機械.forEach(r => { if (r[0] && !/^自社/.test(r[3])) t.機械.push(r.slice()); });
  return t;
}
// 前の日（7日前まで）に書いた日報（無ければ null）
function prevT(siteId, k) {
  for (let i = 1; i <= 7; i++) {
    const d = storedDay(siteId, dayAdd(k, -i));
    if (d) { const t = dayT(d); if (hasData(t)) return t; }
  }
  return null;
}
const prevDayKey = () => { for (let i = 1; i <= 7; i++) { const k = dayAdd(S.day, -i), d = storedDay(S.site, k); if (d && hasData(dayT(d))) return k; } return ""; };
// 日を開くときの中身：保存してあればそれ、無ければ工程から作った予定（保存はまだしない）
function viewDay(siteId, k) {
  const d = storedDay(siteId, k);
  if (d) return { stored: true, 天候: d.天候 || "", 表: dayT(d), 予定: dayP(d), 報告: d.報告 || null, 確認: d.確認 || null, 予定日時: d.予定日時 || "" };
  const p = genPlan(siteId, k);
  return { stored: false, 天候: "", 表: cloneT(p), 予定: p, 報告: null, 確認: null, 予定日時: "" };
}
// その日をファイルに作る（はじめて直したとき）
function ensureDay(doc, siteId, k) {
  if (!doc.日[k]) {
    const p = normT(genPlan(siteId, k));
    doc.日[k] = { 天候: S.weather[k] || "", 表: p, 予定表: cloneT(p), 予定日時: new Date().toLocaleString("ja-JP"), 報告: null, 確認: null };
  }
  return doc.日[k];
}
/* 工程が後から変わったとき：まだ報告していない日は、自分で直していないセルだけ新しい予定に入れ替える */
function refreshPlan(siteId, k) {
  const d = storedDay(siteId, k);
  if (!d || d.報告) return false;
  const op = dayP(d);
  if (!op) return false;
  const np = normT(genPlan(siteId, k));
  if (JSON.stringify(np) === JSON.stringify(op)) return false;
  mutate(siteId, monthOf(k), doc => {
    const x = doc.日[k]; if (!x) return;
    const o = dayP(x) || emptyT(), c = dayT(x), out = emptyT();
    for (const key of SK) {
      const n = Math.max(np[key].length, c[key].length, o[key].length), w = secOf(key).cols.length;
      for (let r = 0; r < n; r++) out[key].push(Array.from({ length: w }, (_, ci) => {
        const cv = (c[key][r] || [])[ci] ?? "", ov = (o[key][r] || [])[ci] ?? "";
        return cv === ov ? ((np[key][r] || [])[ci] ?? "") : cv;
      }));
    }
    x.表 = normT(out); x.予定表 = np; x.予定日時 = new Date().toLocaleString("ja-JP");
  }, { quiet: true });
  return true;
}

/* ---------- 予定と実績を比べる ----------
   作業（基本の工種）・社員（氏名）・人数（作業員の行の数・外注の業者ごとの人数）・機械（機械名＋能力等ごとの数量） */
const numOf = v => { const x = parseFloat(hankaku(v)); return isNaN(x) ? 0 : x; };
function extract(t) {
  const work = new Map(), ppl = new Map(), mach = new Map(), staff = new Set(), crew = [], mat = new Map();
  const add = (m, k, n) => m.set(k, (m.get(k) || 0) + n);
  if (t) {
    t.基本.forEach(r => { if (r[1]) add(work, r[1], 1); });
    t.社員.forEach(r => { if (r[0]) staff.add(r[0]); });
    t.作業員.forEach(r => { if (r[0] || r[1]) { add(ppl, "作業員", 1); if (r[0]) crew.push(r[0]); } });
    t.外注.forEach(r => { if (r[0]) add(ppl, r[0], numOf(r[1])); });
    t.機械.forEach(r => { if (r[0]) add(mach, r[0] + (r[1] ? " " + r[1] : ""), numOf(r[2]) || 1); });
    t.材料.forEach(r => {
      if (!r[0]) return;
      const key = r[0] + (r[1] ? " " + r[1] : ""), x = mat.get(key) || { n: 0, unit: "", who: new Set() };
      x.n += numOf(r[2]); x.unit = x.unit || hankaku(r[2]).replace(/^[\d.\s]+/, "").trim(); if (r[3]) x.who.add(r[3]);
      mat.set(key, x);
    });
  }
  return { work, ppl, mach, staff, crew, mat };
}
function diffDay(v) {
  const a = extract(v.表), p = extract(v.予定);
  const cmp = (pm, am) => [...new Set([...pm.keys(), ...am.keys()])].map(k => ({ k, p: pm.get(k) || 0, a: am.get(k) || 0 }));
  const st = [...new Set([...p.staff, ...a.staff])].map(k => ({ k, p: p.staff.has(k) ? 1 : 0, a: a.staff.has(k) ? 1 : 0 }));
  const out = { 作業: cmp(p.work, a.work), 社員: st, 人数: cmp(p.ppl, a.ppl), 機械: cmp(p.mach, a.mach) };
  out.changed = Object.values(out).some(list => Array.isArray(list) && list.some(x => x.p !== x.a));
  out.cells = 0;
  if (v.予定) for (const key of SK) {
    const P = v.予定[key], A = v.表[key];
    for (let r = 0; r < Math.max(P.length, A.length); r++) for (let c = 0; c < secOf(key).cols.length; c++) if (((P[r] || [])[c] ?? "") !== ((A[r] || [])[c] ?? "")) out.cells++;
  }
  out.a = a; out.p = p;
  return out;
}

/* ---------- 表（6つ。どれも Excel と同じセルの動き） ---------- */
const grids = {};
let cur = "基本", V = null;   // V＝いま開いている日の中身（viewDay）。cur＝いま選んでいる表
const anyEditing = () => Object.values(grids).some(g => g.editing());
const commitAll = () => Object.values(grids).forEach(g => { if (g.editing()) g.commit(); });
const curGrid = () => grids[cur];

function normalize(kind, v) {
  v = String(v ?? "");
  if (kind === "num") return hankaku(v).trim();
  if (kind === "time") {
    let t = hankaku(v).trim().replace(/[時;]/g, ":").replace(/分/g, "");
    if (/^\d{1,2}$/.test(t)) t = (+t) + ":00";
    else if (/^\d{3,4}$/.test(t)) t = (+t.slice(0, -2)) + ":" + t.slice(-2);
    else if (/^\d{1,2}:\d{1,2}$/.test(t)) { const [h, m] = t.split(":"); t = (+h) + ":" + pad2(+m); }
    return t;
  }
  return v.replace(/\s+$/, "");
}
function setupGrids() {
  $("sheet").innerHTML = `<div id="dayHead" class="dayhead"></div>` + SECS.map((s, i) =>
    `<section class="sec${i ? "" : " cur"}" data-sec="${s.key}"><div class="secT"><b>${i + 1}. ${esc(s.title)}</b><span class="cnt" id="cnt_${s.key}"></span></div><div class="gh" id="g_${s.key}"></div></section>`).join("");
  SECS.forEach((s, si) => {
    grids[s.key] = createGrid({
      host: $("g_" + s.key), scroller: $("sheetWrap"), cols: s.cols, headRows: 1, rowBase: 1, rowHeadW: 30, noLetters: true,
      stickyTop: () => $("dayHead").offsetHeight,
      formula: $("fx"), isCur: () => cur === s.key, onFocus: () => setCur(s.key),
      rows: () => V ? Math.max(s.min, V.表[s.key].length + 1) : s.min,    // いつも下に空の行を1つ
      head: () => `<tr class="xlhr sh"><th class="xlrh"></th>${s.cols.map((c, i) => `<th data-col="${i}" class="${c.cls}">${esc(c.t)}</th>`).join("")}</tr>`,
      get: (r, c) => V ? ((V.表[s.key][r] || [])[c] ?? "") : "",
      deco: (r, c) => {
        if (!V || !V.予定) return null;
        const a = (V.表[s.key][r] || [])[c] ?? "", p = (V.予定[s.key][r] || [])[c] ?? "";
        if (a === p) return null;
        return { cls: p ? "chg" : "add", title: p ? "予定：" + p + (a ? "" : "（消した）") : "予定に無かったもの" };
      },
      normalize: (c, v) => normalize(s.cols[c].kind, v),
      complete: (c, v) => { const lv = v.toLowerCase(); return colWords(s.key, c).find(x => x.toLowerCase().startsWith(lv) && x.length > v.length) || null; },
      list: c => colWords(s.key, c),
      apply: ch => applyCells(s.key, ch),
      undo, redo,
      onSheet: dir => moveDay(dir),
      onEdge: (dir, c) => {
        const t = SECS[si + dir];
        if (!t) return false;
        const j = t.cols.findIndex(x => x.t === s.cols[c].t);
        grids[t.key].enter(dir, j >= 0 ? j : Math.min(c, t.cols.length - 1));
        return true;
      },
      rowOp: (op, r, n) => rowOp(s.key, op, r, n),
      menu: (x, y, sel) => openMenu(s.key, x, y, sel),
      print: () => openPrint(),
      find: () => openFind(),
      toast,
      cellName: (r, c) => s.label + (r + 1) + "・" + s.cols[c].t,
      onSelect: (name, val, stat) => {
        if (cur !== s.key) return;
        $("nameBox").value = name;
        if (document.activeElement !== $("fx")) $("fx").value = val ?? grids[s.key]?.editText() ?? "";
        $("stat").textContent = stat || "";
      },
      onEditText: t => { if (document.activeElement !== $("fx")) $("fx").value = t; },
      onMode: m => { $("mode").textContent = m === "edit" ? "編集" : m ? "入力" : "準備完了"; }
    });
  });
  $("sheet").addEventListener("click", e => { if (e.target.closest("#wxCell")) editWeather(); });
}
function setCur(key) {
  if (cur === key) return;
  if (grids[cur] && grids[cur].editing()) grids[cur].commit();
  cur = key;
  document.querySelectorAll("#sheet .sec").forEach(x => x.classList.toggle("cur", x.dataset.sec === key));
  grids[key].place();
}
// その列で今月・先月に使った言葉（オートコンプリートと Alt+↓ の一覧）。氏名の欄はマスタの社員・従業員も
const wordCache = {};
const nameOf = x => typeof x === "string" ? x : (x && (x.名前 || x.氏名)) || "";
function colWords(key, c) {
  const ck = S.site + "|" + monthOf(S.day) + "|" + key + "|" + c;
  if (wordCache[ck]) return wordCache[ck];
  const cnt = new Map();
  const addRows = rows => (rows || []).forEach(r => { const v = r[c]; if (v) cnt.set(v, (cnt.get(v) || 0) + 1); });
  for (const m of [monthOf(S.day), monthAdd(monthOf(S.day), -1)]) {
    for (const f of Object.values(S.N)) {
      if (!f.doc || f.doc.月 !== m) continue;
      Object.values(f.doc.日 || {}).forEach(d => { addRows(dayT(d)[key]); const p = dayP(d); if (p) addRows(p[key]); });
    }
  }
  if (V) { addRows(V.表[key]); if (V.予定) addRows(V.予定[key]); }
  const col = secOf(key).cols[c];
  if (col.kind === "name") (key === "社員" ? M().社員 : M().従業員 || []).map(nameOf).filter(Boolean).forEach(n => cnt.set(n, cnt.get(n) || 0));
  if (key === "機械" && c === 3) ["自社", "東リース", "レント", "セフテック", "第一産商"].forEach(x => cnt.set(x, cnt.get(x) || 0));
  return wordCache[ck] = [...cnt.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "ja")).map(x => x[0]);
}
const clearWords = () => Object.keys(wordCache).forEach(k => delete wordCache[k]);

/* ---------- 直す（元に戻せるように、その日の中身を前後で覚える） ---------- */
const ukey = () => S.site + "|" + S.day;
function snapshot() { return { 表: cloneT(V.表), 予定: V.予定 ? cloneT(V.予定) : null, 天候: V.天候 }; }
function commitDay(before, after) {
  (S.undo[ukey()] = S.undo[ukey()] || []).push({ before, after });
  S.redo[ukey()] = [];
  writeDay(S.site, S.day, after);
}
function writeDay(siteId, k, snap) {
  mutate(siteId, monthOf(k), doc => {
    const d = ensureDay(doc, siteId, k);
    d.表 = normT(snap.表);
    if (snap.予定) d.予定表 = normT(snap.予定);
    if (snap.天候 !== undefined) d.天候 = snap.天候;
  });
}
function applyCells(key, changes) {
  const before = snapshot(), rows = V.表[key];
  for (const ch of changes) { while (rows.length <= ch.r) rows.push(blankRow(key)); rows[ch.r][ch.c] = ch.v; }
  if (!V.stored && !V.天候 && S.weather[S.day]) V.天候 = S.weather[S.day];
  commitDay(before, snapshot());
}
function undo() {
  const st = S.undo[ukey()]; if (!st || !st.length) { toast("元に戻すものはありません"); return; }
  const x = st.pop(); (S.redo[ukey()] = S.redo[ukey()] || []).push(x);
  writeDay(S.site, S.day, x.before); toast("元に戻しました");
}
function redo() {
  const st = S.redo[ukey()]; if (!st || !st.length) return;
  const x = st.pop(); S.undo[ukey()].push(x);
  writeDay(S.site, S.day, x.after);
}
// 行の挿入・削除（予定の行も一緒にずらして、比べる位置がずれないように）
function rowOp(key, op, r, n) {
  const before = snapshot(), A = V.表[key], P = V.予定 && V.予定[key];
  if (op === "insert") { for (let i = 0; i < n; i++) { A.splice(r, 0, blankRow(key)); if (P && r <= P.length) P.splice(r, 0, blankRow(key)); } }
  else { A.splice(r, n); if (P) P.splice(r, n); }
  commitDay(before, snapshot());
}

/* ---------- 画面 ---------- */
function render() {
  clearWords();
  renderTop();
  if (S.tab === "sheet") renderSheet();
  else if (S.tab === "haichi") renderHaichi();
  else if (S.tab === "month") renderMonth();
}
function renderTop() {
  const list = sites().filter(s => !s.完了 || s.id === S.site);
  const sel = $("inSite");
  const html = list.map(s => `<option value="${esc(s.id)}">${esc(nm(s))}${s.担当 ? "（" + esc(s.担当) + "）" : ""}</option>`).join("");
  if (sel.dataset.h !== html) { sel.innerHTML = html; sel.dataset.h = html; }
  sel.value = S.site;
  $("inDay").value = S.day;
  showSync();
}
function renderSheet() {
  $("noSite").hidden = !!S.site; $("sheet").hidden = !S.site;
  if (!S.site) return;
  V = viewDay(S.site, S.day);
  renderDayHead();
  SK.forEach(k => grids[k].render());
  renderCounts();
  renderTabs();
  renderSide();
  renderStatus();
}
function renderDayHead() {
  const k = S.day, w = V.天候, sug = !w && S.weather[k] ? S.weather[k] : "";
  const rep = V.報告 ? `<span class="rep">報告済み ${esc(V.報告.だれ || "")} ${esc(V.報告.いつ || "")}</span>` : V.stored ? `<span class="edit">書きかけ</span>` : `<span class="plan">予定（工程から）</span>`;
  $("dayHead").innerHTML = `<span class="date ${dayDow(k) === 0 || holidayOf(k) ? "sun" : dayDow(k) === 6 ? "sat" : ""}">${esc(wareki(k))}${holidayOf(k) ? " " + esc(holidayOf(k)) : ""}</span>` +
    `<span class="lbl">天候</span><span class="wx" id="wxCell" title="押して天候を入れる（晴・曇・雨…）">${w ? esc(w) : sug ? `<span class="sug">${esc(sug)}</span>` : "　　"}</span>${rep}`;
}
// 表の見出しの右に人数・台数
function renderCounts() {
  const e = extract(V.表);
  const tot = m => [...m.values()].reduce((a, x) => a + x, 0);
  const n = { 基本: e.work.size ? e.work.size + "工種" : "", 社員: e.staff.size ? e.staff.size + "名" : "", 作業員: e.ppl.get("作業員") ? e.ppl.get("作業員") + "名" : "",
    外注: (() => { const x = tot(e.ppl) - (e.ppl.get("作業員") || 0); return x ? x + "名" : ""; })(), 材料: e.mat.size ? e.mat.size + "品目" : "", 機械: tot(e.mach) ? tot(e.mach) + "台" : "" };
  SK.forEach(k => { $("cnt_" + k).textContent = n[k]; });
}
// 下の日の見出し（1日(火) 2日(水) …）
function renderTabs() {
  const m = monthOf(S.day);
  const h = daysIn(m).map(k => {
    const d = storedDay(S.site, k), dw = dayDow(k);
    const cls = [k === S.day ? "on" : "", dw === 0 || holidayOf(k) ? "sun" : dw === 6 ? "sat" : "", k === todayKey() ? "today" : "",
      d && d.報告 ? "rep" : d ? "edited" : ""].join(" ");
    const mark = d && d.報告 ? (diffDay(viewDay(S.site, k)).changed ? "◆" : "✓") : "";
    return `<button type="button" class="${cls}" data-day="${k}" title="${esc(wareki(k) + (holidayOf(k) ? " " + holidayOf(k) : "") + (d && d.報告 ? "\n報告済み" : d ? "\n書きかけ" : "\n予定のまま（まだ開いていない）"))}">${tabName(k)}${mark ? `<i>${mark}</i>` : ""}</button>`;
  }).join("");
  $("tabs").innerHTML = h;
  const on = $("tabs").querySelector(".on");
  if (on) on.scrollIntoView({ block: "nearest", inline: "nearest" });
}
// 右の「予定と比べる」
function renderSide() {
  const d = diffDay(V), s = siteById(S.site);
  const sec = (ttl, list, unit, fmt) => {
    if (!list.length) return `<h4>${ttl}</h4><p class="none">なし</p>`;
    return `<h4>${ttl}</h4><table class="cmp"><tr><th></th><th>予定</th><th>実績</th></tr>` + list.map(x => {
      const ch = x.p !== x.a;
      return `<tr class="${ch ? (x.a > x.p ? "up" : "down") : ""}"><td>${esc(x.k)}</td><td>${fmt ? fmt(x.p) : x.p ? x.p + unit : "—"}</td><td>${fmt ? fmt(x.a) : x.a ? x.a + unit : "—"}</td></tr>`;
    }).join("") + `</table>`;
  };
  const yes = n => n ? "○" : "—";
  const pplTot = m => [...m.values()].reduce((a, x) => a + x, 0);
  $("side").innerHTML =
    `<div class="sideTtl">予定と比べる</div>` +
    `<div class="sum ${d.changed ? "ng" : "ok"}">${d.changed ? "予定から変わっています" : "予定どおり"}${d.cells ? `<span>（直したセル ${d.cells}）</span>` : ""}</div>` +
    `<div class="tot">人数　予定 <b>${pplTot(d.p.ppl)}</b> 名 → 実績 <b>${pplTot(d.a.ppl)}</b> 名　／　社員 ${d.p.staff.size} → ${d.a.staff.size}</div>` +
    sec("作業（工種）", d.作業, "", yes) + sec("社員", d.社員, "", yes) + sec("人数（作業員・外注）", d.人数, "名") + sec("使用機械", d.機械, "台") +
    `<div class="legend"><span class="chgbox"></span>予定から直したセル　<span class="addbox"></span>予定に無かったセル</div>` +
    (s ? `<div class="note">予定は全体工程表「${esc(nm(s))}」の工程${S.kikai ? "・自社機械稼働表の予約" : ""}・前の日の機械と作業員から作っています。${V.予定日時 ? "（予定を作った時刻 " + esc(V.予定日時) + "）" : ""}</div>` : "") +
    (V.確認 ? `<div class="kakunin">✔ ${esc(V.確認.だれ)} さんが確認（${esc(V.確認.いつ)}）</div>` : "");
}
function renderStatus() {
  const d = storedDay(S.site, S.day);
  $("status").textContent = d && d.報告 ? "報告済み" : d ? "書きかけ（まだ報告していません）" : "予定のまま（工程から作った予定です。直すか「報告する」で保存します）";
  $("bReport").textContent = d && d.報告 ? "報告を取り消す" : diffDay(V).changed ? "この内容で報告する" : "予定どおりで報告する";
  $("bReport").classList.toggle("sub", !!(d && d.報告));
}

function setDay(k) {
  commitAll();
  const m = monthOf(k);
  S.day = k;
  try { localStorage.setItem("nippou_view", JSON.stringify({ site: S.site, day: k })); } catch (e) { }
  ensureMonth(S.site, m).then(() => { if (S.day === k) { refreshPlan(S.site, k); render(); } });
  render();
  if (S.tab === "sheet" && curGrid()) curGrid().focus();
}
function moveDay(dir) { setDay(dayAdd(S.day, dir)); }
// その月（と前の月）の日報ファイルを読んでおく
const loadedMonths = new Set();
async function ensureMonth(siteId, m) {
  const want = [fileName(siteId, m), fileName(siteId, monthAdd(m, -1))].filter(n => !loadedMonths.has(n));
  for (const n of want) {
    loadedMonths.add(n);
    if (S.N[n]) continue;
    try { const r = await nippouStore.get(n); if (r && !S.N[n]) S.N[n] = { doc: r.data, eTag: r.eTag, pending: [], saving: false, timer: null, err: null }; }
    catch (e) { console.warn(e); loadedMonths.delete(n); }
  }
}

/* ---------- 天候 ---------- */
const WX = ["晴", "曇", "雨", "雪", "晴後曇", "曇後晴", "曇後雨", "雨後曇", "晴時々曇", "曇時々雨", "小雨"];
function editWeather() {
  const cw = V.天候 || S.weather[S.day] || "";
  const box = $("wxPick");
  box.innerHTML = WX.map(w => `<button type="button" class="${w === cw ? "on" : ""}">${w}</button>`).join("") + `<input id="wxFree" placeholder="ほか" value="">`;
  const td = $("wxCell").getBoundingClientRect();
  Object.assign(box.style, { left: td.left + "px", top: td.bottom + "px" });
  box.hidden = false;
  box.onclick = e => { const b = e.target.closest("button"); if (b) setWeather(b.textContent); };
  $("wxFree").onkeydown = e => { if (e.key === "Enter" && !e.isComposing) setWeather(e.target.value.trim()); if (e.key === "Escape") box.hidden = true; };
}
function setWeather(w) {
  $("wxPick").hidden = true;
  const before = snapshot(); V.天候 = w;
  commitDay(before, snapshot());
  curGrid().focus();
}
document.addEventListener("mousedown", e => { if (!$("wxPick").hidden && !$("wxPick").contains(e.target) && !e.target.closest("#wxCell")) $("wxPick").hidden = true; });
// open-meteo の日ごとの天気（過去31日〜7日先）。緯度経度だけを送る
const WXCODE = c => c <= 1 ? "晴" : c <= 3 ? "曇" : c <= 48 ? "曇" : c <= 67 ? "雨" : c <= 77 ? "雪" : c <= 82 ? "雨" : c <= 86 ? "雪" : "雨";
async function loadWeather() {
  try {
    const w = CFG.weather;
    const r = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${w.lat}&longitude=${w.lon}&daily=weather_code&timezone=Asia%2FTokyo&past_days=31&forecast_days=7`);
    if (!r.ok) return;
    const j = await r.json();
    j.daily.time.forEach((k, i) => { S.weather[k] = WXCODE(j.daily.weather_code[i]); });
    if (S.tab === "sheet") renderSheet();
  } catch (e) { console.warn("天気が取れません", e); }
}

/* ---------- ボタン ---------- */
function report() {
  commitAll();
  const d = storedDay(S.site, S.day);
  if (d && d.報告) {
    if (!confirm("報告を取り消して、書きかけに戻しますか？")) return;
    mutate(S.site, monthOf(S.day), doc => { const x = doc.日[S.day]; if (x) x.報告 = null; });
    return;
  }
  const site = S.site, k = S.day, snap = snapshot();
  mutate(site, monthOf(k), doc => {
    const x = ensureDay(doc, site, k);
    if (!d) { x.表 = normT(snap.表); x.予定表 = normT(snap.予定 || snap.表); }
    if (!x.天候) x.天候 = snap.天候 || S.weather[k] || "";
    x.報告 = { だれ: me.name, いつ: new Date().toLocaleString("ja-JP") };
  });
  toast(md(k) + " を報告しました" + (diffDay(viewDay(site, k)).changed ? "（予定からの変更あり）" : "（予定どおり）"));
}
function copyPrev() {
  const p = prevT(S.site, S.day);
  if (!p) { toast("7日前までに書いた日報がありません"); return; }
  if (!confirm("前に書いた日（" + md(prevDayKey()) + "）の内容で、この日を置き換えますか？（予定は残るので、比べられます）")) return;
  const before = snapshot(); V.表 = cloneT(p);
  commitDay(before, snapshot());
}
function resetPlan() {
  if (!confirm("工程から予定を作り直して、この日の内容を予定に戻しますか？（いま書いてある内容は消えます。Ctrl+Z で戻せます）")) return;
  const before = snapshot(), p = genPlan(S.site, S.day);
  V.表 = cloneT(p); V.予定 = p;
  commitDay(before, snapshot());
}
function backToPlan(key, sel) {
  if (!V.予定) return;
  const before = snapshot(), A = V.表[key], P = V.予定[key];
  for (let r = sel.r0; r <= sel.r1; r++) {
    while (A.length <= r) A.push(blankRow(key));
    for (let c = sel.c0; c <= sel.c1; c++) A[r][c] = (P[r] || [])[c] ?? "";
  }
  commitDay(before, snapshot());
}

/* 右クリックのメニュー */
function openMenu(key, x, y, sel) {
  const m = $("menu");
  const items = [
    ["切り取り", "Ctrl+X", () => document.execCommand("cut")],
    ["コピー", "Ctrl+C", () => document.execCommand("copy")],
    ["貼り付け", "Ctrl+V", async () => { try { grids[key].paste(await navigator.clipboard.readText()); } catch (e) { toast("貼り付けは Ctrl+V を使ってください"); } }],
    null,
    ["行を挿入", "Ctrl+Shift++", () => rowOp(key, "insert", sel.r0, sel.r1 - sel.r0 + 1)],
    ["行を削除", "Ctrl+-", () => rowOp(key, "delete", sel.r0, sel.r1 - sel.r0 + 1)],
    ["数式と値のクリア", "Delete", () => { const ch = []; for (let r = sel.r0; r <= sel.r1; r++) for (let c = sel.c0; c <= sel.c1; c++) ch.push({ r, c, v: "" }); applyCells(key, ch); }],
    null,
    ["予定に戻す（選んだセル）", "", () => backToPlan(key, sel)],
    ["ドロップダウン リストから選択", "Alt+↓", () => grids[key].showList()]
  ];
  m.innerHTML = items.map((it, i) => it ? `<div data-i="${i}"><span>${it[0]}</span><kbd>${it[1]}</kbd></div>` : `<hr>`).join("");
  m.hidden = false;
  const r = m.getBoundingClientRect();
  m.style.left = Math.min(x, innerWidth - r.width - 4) + "px"; m.style.top = Math.min(y, innerHeight - r.height - 4) + "px";
  m.onmousedown = e => { e.preventDefault(); const d = e.target.closest("[data-i]"); if (!d) return; m.hidden = true; items[+d.dataset.i][2](); grids[key].focus(); };
}
document.addEventListener("mousedown", e => { if (!$("menu").hidden && !$("menu").contains(e.target)) $("menu").hidden = true; });

/* 検索（Ctrl+F）：この月の日報から言葉を探す（Excel の MY_FIND の代わり） */
function openFind() {
  const q = prompt("この月の日報から探す言葉");
  if (!q) return;
  const hits = [];
  for (const k of daysIn(monthOf(S.day))) {
    const v = viewDay(S.site, k);
    SK.forEach((key, si) => v.表[key].forEach((r, ri) => r.forEach((x, ci) => { if (String(x).includes(q)) hits.push({ k, si, key, ri, ci }); })));
  }
  if (!hits.length) { toast("「" + q + "」は見つかりませんでした"); return; }
  const s = curGrid().selection(), ci = SK.indexOf(cur);
  const after = h => h.k > S.day || (h.k === S.day && (h.si > ci || (h.si === ci && (h.ri > s.r || (h.ri === s.r && h.ci > s.c)))));
  const next = hits.find(after) || hits[0];
  if (next.k !== S.day) setDay(next.k);
  setCur(next.key); grids[next.key].select(next.ri, next.ci);
  toast(`「${q}」${hits.length}か所（${md(next.k)} ${secOf(next.key).label}${next.ri + 1}）`);
}

/* カレンダー（Excel の S_S フォームの代わり：日のボタンを押すとその日へ） */
function openCal() {
  const m = monthOf(S.day), days = daysIn(m), lead = (dayDow(days[0]) + 6) % 7;
  const cells = Array(lead).fill("").concat(days);
  $("calTtl").textContent = wareki(days[0]).replace(/\d+日\(.\)$/, "") + "分 工事日報";
  $("calGrid").innerHTML = ["月", "火", "水", "木", "金", "土", "日"].map(w => `<b>${w}</b>`).join("") + cells.map(k => {
    if (!k) return `<span></span>`;
    const d = storedDay(S.site, k), dw = dayDow(k);
    return `<button type="button" data-day="${k}" class="${k === S.day ? "on" : ""} ${k === todayKey() ? "today" : ""} ${dw === 0 || holidayOf(k) ? "sun" : dw === 6 ? "sat" : ""} ${d && d.報告 ? "rep" : d ? "edited" : ""}">${+k.slice(8)}</button>`;
  }).join("");
  $("calToday").textContent = "今日は　" + wareki(todayKey());
  $("dCal").showModal();
}

/* ---------- 配置まとめ（田村さん：全現場の予定・報告を見て配置表へ） ---------- */
let haichiDay = "";
function siteActive(s, k) { return (s.バー || []).some(b => b.開始 <= k && b.終了 >= k && barWorks((s.段 || []).find(l => l.id === b.段), b, k)); }
const haichiSites = k => sites().filter(s => !s.完了 && (siteActive(s, k) || storedDay(s.id, k)));
async function renderHaichi() {
  const k = haichiDay || (haichiDay = dayAdd(todayKey(), 1));
  $("hDay").value = k;
  $("hTtl").textContent = wareki(k) + (holidayOf(k) ? "　" + holidayOf(k) : "");
  const list = haichiSites(k);
  const rows = [];
  const totStaff = new Set();
  let crewP = 0, crewA = 0, gaiP = 0, gaiA = 0;
  for (const s of list) {
    const v = viewDay(s.id, k), d = diffDay(v), sd = storedDay(s.id, k);
    d.a.staff.forEach(n => totStaff.add(n));
    const crew = d.人数.filter(x => x.k === "作業員"), gai = d.人数.filter(x => x.k !== "作業員");
    crew.forEach(x => { crewP += x.p; crewA += x.a; }); gai.forEach(x => { gaiP += x.p; gaiA += x.a; });
    const listOf = (arr, unit) => arr.filter(x => x.a || x.p).map(x => {
      const t = esc(x.k) + (unit && x.a ? " " + x.a + unit : "");
      return x.p === x.a ? t : x.a === 0 ? `<s class="del" title="予定 ${x.p}${unit}">${esc(x.k)}${unit ? " " + x.p + unit : ""}</s>` : `<span class="chg" title="予定 ${x.p ? x.p + unit : "なし"}">${t}</span>`;
    }).join("<br>");
    const state = sd && sd.報告 ? (d.changed ? `<span class="b chg">変更あり</span>` : `<span class="b ok">予定どおり</span>`) + `<div class="when">${esc(sd.報告.だれ)} ${esc(sd.報告.いつ)}</div>`
      : sd ? `<span class="b edit">書きかけ</span>` : k <= todayKey() ? `<span class="b none">未報告</span>` : `<span class="b plan">工程の予定</span>`;
    rows.push(`<tr data-site="${esc(s.id)}" class="${d.changed ? "changed" : ""}"><td class="nm"><a href="#" data-open="${esc(s.id)}">${esc(nm(s))}</a><div class="tanto">${esc(s.担当 || "")}</div></td>` +
      `<td>${state}</td><td>${listOf(d.作業, "")}</td><td>${listOf(d.社員, "")}</td>` +
      `<td>${listOf(crew, "名")}${d.a.crew.length ? `<div class="names">${esc(d.a.crew.join("・"))}</div>` : ""}</td><td>${listOf(gai, "名")}</td><td>${listOf(d.機械, "台")}</td>` +
      `<td class="wx">${esc(v.天候 || S.weather[k] || "")}</td>` +
      `<td class="ck"><label><input type="checkbox" data-ck="${esc(s.id)}" ${v.確認 ? "checked" : ""}> 配置表に入れた</label>${v.確認 ? `<div class="when">${esc(v.確認.だれ)} ${esc(v.確認.いつ)}</div>` : ""}</td></tr>`);
  }
  $("haichi").innerHTML = list.length ? `<table class="hz"><thead><tr><th>現場</th><th>状態</th><th>作業</th><th>社員</th><th>作業員</th><th>外注</th><th>機械</th><th>天候</th><th>田村さん</th></tr></thead><tbody>${rows.join("")}</tbody>` +
    `<tfoot><tr><th>計 ${list.length}現場</th><td></td><td></td><td>${totStaff.size}人</td><td>予定 ${crewP}名 → ${crewA}名</td><td>予定 ${gaiP}名 → ${gaiA}名</td><td colspan="3"></td></tr></tfoot></table>`
    : `<p class="empty">この日に工程のある現場はありません。</p>`;
}
function haichiTSV() {
  const k = haichiDay, out = [["現場", "担当", "状態", "作業", "社員", "作業員", "外注", "機械", "変更"].join("\t")];
  for (const s of haichiSites(k)) {
    const v = viewDay(s.id, k), d = diffDay(v), sd = storedDay(s.id, k);
    const j = (arr, u) => arr.filter(x => x.a).map(x => x.k + (u ? x.a + u : "")).join("・");
    const crew = d.人数.find(x => x.k === "作業員");
    const chg = [...d.作業, ...d.社員, ...d.人数, ...d.機械].filter(x => x.p !== x.a).map(x => x.k + " " + x.p + "→" + x.a).join("、");
    out.push([nm(s), s.担当 || "", sd && sd.報告 ? "報告済み" : sd ? "書きかけ" : "予定", j(d.作業), j(d.社員),
      (crew && crew.a ? crew.a + "名" : "") + (d.a.crew.length ? "（" + d.a.crew.join("・") + "）" : ""), j(d.人数.filter(x => x.k !== "作業員"), "名"), j(d.機械, "台"), chg].join("\t"));
  }
  return out.join("\r\n");
}

/* ---------- 月の集計（Excel の 出面・誘導員 のマクロの代わり） ----------
   上の「計」は報告した日・書きかけの日だけ。予定のままの日（灰色）も入れた数は「予定も入れた計」に出す */
function renderMonth() {
  const m = monthOf(S.day), days = daysIn(m);
  const per = days.map(k => ({ k, v: viewDay(S.site, k), sd: storedDay(S.site, k) }));
  const ppl = new Set(), mach = new Set(), staff = new Set();
  per.forEach(x => { x.e = extract(x.v.表); x.e.ppl.forEach((_, k) => ppl.add(k)); x.e.mach.forEach((_, k) => mach.add(k)); x.e.staff.forEach(n => staff.add(n)); });
  const byJa = (a, b) => a.localeCompare(b, "ja");
  const P = [...ppl].sort((a, b) => (b === "作業員") - (a === "作業員") || byJa(a, b)), Mc = [...mach].sort(byJa), St = [...staff].sort(byJa);
  const sum = (f, all) => per.reduce((a, x) => a + (x.sd || all ? f(x) : 0), 0);
  const foot = (ttl, all) => `<tr class="${all ? "all" : ""}"><th>${ttl}</th><td></td><td></td><td></td>` + St.map(n => `<td class="c">${sum(x => x.e.staff.has(n) ? 1 : 0, all)}日</td>`).join("") +
    P.map(n => `<td class="r">${sum(x => x.e.ppl.get(n) || 0, all)}</td>`).join("") + Mc.map(n => `<td class="r">${sum(x => x.e.mach.get(n) || 0, all)}</td>`).join("") + `</tr>`;
  const h = [`<p class="note">「計」は報告した日・書きかけの日（一度でも直した日）だけを数えます。予定のままの日（灰色）は数えず、下の「予定も入れた計」に入ります。表は選んでコピーすると Excel に貼れます。</p>`,
    `<table class="mz"><thead><tr><th>日</th><th>状態</th><th>天候</th><th>作業</th>${St.map(n => `<th class="st">${esc(n)}</th>`).join("")}${P.map(n => `<th class="pp">${esc(n)}</th>`).join("")}${Mc.map(n => `<th class="mc">${esc(n)}</th>`).join("")}</tr></thead><tbody>`];
  for (const x of per) {
    const dw = dayDow(x.k);
    h.push(`<tr class="${x.sd ? "" : "plan"} ${dw === 0 || holidayOf(x.k) ? "sun" : dw === 6 ? "sat" : ""}"><td><a href="#" data-goto="${x.k}">${tabName(x.k)}</a></td><td>${x.sd && x.sd.報告 ? "報告" : x.sd ? "書きかけ" : "予定"}</td><td>${esc(x.v.天候)}</td><td>${esc([...x.e.work.keys()].join("・"))}</td>` +
      St.map(n => `<td class="c">${x.e.staff.has(n) ? "○" : ""}</td>`).join("") + P.map(n => `<td class="r">${x.e.ppl.get(n) || ""}</td>`).join("") + Mc.map(n => `<td class="r">${x.e.mach.get(n) || ""}</td>`).join("") + `</tr>`);
  }
  h.push(`</tbody><tfoot>${foot("計", false)}${foot("予定も入れた計", true)}</tfoot></table>`);
  // 材料（報告した日・書きかけの日）
  const mat = new Map();
  per.filter(x => x.sd).forEach(x => x.e.mat.forEach((v, k) => { const y = mat.get(k) || { n: 0, unit: "", who: new Set(), days: 0 }; y.n += v.n; y.unit = y.unit || v.unit; v.who.forEach(w => y.who.add(w)); y.days++; mat.set(k, y); }));
  h.push(`<h4 class="mh">材料（報告した日・書きかけの日の合計）</h4>` + (mat.size ? `<table class="mz"><thead><tr><th>材料名・詳細</th><th>数量</th><th>使った日数</th><th>業者名</th></tr></thead><tbody>` +
    [...mat.entries()].sort((a, b) => byJa(a[0], b[0])).map(([k, y]) => `<tr><td>${esc(k)}</td><td class="r">${Math.round(y.n * 100) / 100}${esc(y.unit)}</td><td class="r">${y.days}日</td><td>${esc([...y.who].join("・"))}</td></tr>`).join("") + `</tbody></table>` : `<p class="note">まだありません。</p>`));
  $("month").innerHTML = h.join("");
  $("mTtl").textContent = (siteById(S.site) ? nm(siteById(S.site)) : "") + "　" + wareki(days[0]).replace(/\d+日\(.\)$/, "") + "分";
}

/* ---------- 印刷（B5 縦。1日＝1ページ、6つの欄を上から順に） ---------- */
function openPrint() { $("dPrint").showModal(); }
function printPages(which) {
  commitAll();
  const days = which === "month" ? daysIn(monthOf(S.day)).filter(k => storedDay(S.site, k)) : [S.day];
  if (!days.length) { toast("この月に書いた日報がありません"); return; }
  const s = siteById(S.site);
  const out = days.map(k => {
    const v = viewDay(S.site, k);
    const tables = SECS.map(sec => {
      const rows = v.表[sec.key].filter(r => r.some(x => x));
      while (rows.length < sec.min) rows.push(blankRow(sec.key));
      const tw = sec.cols.reduce((a, c) => a + c.w, 0);
      return `<h4>${esc(sec.title)}</h4><table><colgroup>${sec.cols.map(c => `<col style="width:${(c.w / tw * 100).toFixed(1)}%">`).join("")}</colgroup>` +
        `<tr class="h">${sec.cols.map(c => `<th>${esc(c.t)}</th>`).join("")}</tr>` +
        rows.map(r => `<tr>${r.map((x, i) => `<td class="${sec.cols[i].cls}">${esc(x)}</td>`).join("")}</tr>`).join("") + `</table>`;
    }).join("");
    return `<div class="pp"><div class="ph"><b>工事日報</b><span>${esc(s ? s.名前 : "")}</span><span>${esc(wareki(k))}</span><span>天候　${esc(v.天候)}</span>` +
      `<span class="rp">${v.報告 ? "報告 " + esc(v.報告.だれ) : ""}</span></div>${tables}</div>`;
  });
  $("printArea").innerHTML = out.join("");
  $("dPrint").close();
  setTimeout(() => window.print(), 50);
}

/* ---------- サインイン・読み込み ---------- */
let ready = false, pulling = false;
async function pull() {
  if (pulling || !ready) return;
  pulling = true;
  let changed = false;
  try {
    for (const it of await koteiStore.list()) {
      if (!it.name.endsWith(".json") || !(it.name.startsWith("現場_") || it.name === "マスタ.json" || it.name === "休日.json")) continue;
      const f = S.K[it.name];
      if (f && f.eTag === it.eTag) continue;
      const r = await koteiStore.get(it.name);
      if (r) { S.K[it.name] = { doc: r.data, eTag: r.eTag }; changed = true; }
    }
    // 日報：いま開いている月と前後の月だけ
    const m = monthOf(S.day || todayKey()), want = new Set([monthAdd(m, -1), m, monthAdd(m, 1)]);
    for (const it of await nippouStore.list()) {
      const mm = (it.name.match(/_(\d{4}-\d{2})\.json$/) || [])[1];
      if (!mm || !want.has(mm)) continue;
      const f = S.N[it.name];
      if (f && (f.eTag === it.eTag || f.pending.length || f.saving)) continue;
      const r = await nippouStore.get(it.name);
      if (r) { S.N[it.name] = { doc: r.data, eTag: r.eTag, pending: [], saving: false, timer: null, err: null }; loadedMonths.add(it.name); changed = true; }
    }
  } catch (e) {
    console.error(e);
    if (e.status === 404 || e.status === 403) toast("工程データ フォルダを開けません。北澤さんに共有（編集可）を頼んでください");
    else toast("最新の読み込みに失敗：" + e.message);
  } finally { pulling = false; }
  if (changed && !anyEditing()) { if (S.day && S.site) refreshPlan(S.site, S.day); render(); }
}
async function pullKikai() {
  try { const r = await kikaiStore.get("稼働データ.json"); if (r) { S.kikai = r.data; if (ready && !anyEditing()) render(); } }
  catch (e) { console.warn("自社機械稼働表は読めませんでした（機械の予定なしで動きます）", e); }
}

let toastTimer = null;
function toast(msg, ms = 3500) {
  const el = $("toast"); el.textContent = msg; el.classList.add("on");
  clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove("on"), ms);
}

function bind() {
  $("inSite").onchange = e => { S.site = e.target.value; S.undo = {}; S.redo = {}; setDay(S.day); };
  // 上の ◀▶ は前の日・次の日（2026-10-30 ユーザー指定）
  $("inDay").onchange = e => { if (e.target.value) setDay(e.target.value); };
  $("bPrevD").onclick = () => moveDay(-1);
  $("bNextD").onclick = () => moveDay(1);
  $("bToday").onclick = () => setDay(todayKey());
  $("bTomorrow").onclick = () => setDay(dayAdd(todayKey(), 1));
  $("tabs").onclick = e => { const b = e.target.closest("[data-day]"); if (b) setDay(b.dataset.day); };
  $("bTabL").onclick = () => moveDay(-1);
  $("bTabR").onclick = () => moveDay(1);
  $("bCal").onclick = openCal;
  $("calGrid").onclick = e => { const b = e.target.closest("[data-day]"); if (b) { $("dCal").close(); setDay(b.dataset.day); } };
  $("calToday").onclick = () => { $("dCal").close(); setDay(todayKey()); };
  $("bCalClose").onclick = () => $("dCal").close();
  $("bReport").onclick = report;
  $("bCopyPrev").onclick = copyPrev;
  $("bResetPlan").onclick = resetPlan;
  $("bInsRow").onclick = () => { const s = curGrid().selection(); rowOp(cur, "insert", s.r0, s.r1 - s.r0 + 1); curGrid().focus(); };
  $("bDelRow").onclick = () => { const s = curGrid().selection(); rowOp(cur, "delete", s.r0, s.r1 - s.r0 + 1); curGrid().focus(); };
  $("bUndo").onclick = () => { undo(); curGrid().focus(); };
  $("bRedo").onclick = () => { redo(); curGrid().focus(); };
  $("bPrint").onclick = openPrint;
  $("bPrDay").onclick = () => printPages("day");
  $("bPrMonth").onclick = () => printPages("month");
  $("bPrCancel").onclick = () => $("dPrint").close();
  document.querySelectorAll("#views button").forEach(b => b.onclick = () => {
    S.tab = b.dataset.tab;
    document.querySelectorAll("#views button").forEach(x => x.classList.toggle("on", x === b));
    $("pSheet").hidden = S.tab !== "sheet"; $("pHaichi").hidden = S.tab !== "haichi"; $("pMonth").hidden = S.tab !== "month";
    render();
    if (S.tab === "sheet") curGrid().focus();
  });
  $("hDay").onchange = e => { haichiDay = e.target.value; renderHaichi(); };
  $("hPrev").onclick = () => { haichiDay = dayAdd(haichiDay, -1); renderHaichi(); };
  $("hNext").onclick = () => { haichiDay = dayAdd(haichiDay, 1); renderHaichi(); };
  $("hToday").onclick = () => { haichiDay = todayKey(); renderHaichi(); };
  $("hTomorrow").onclick = () => { haichiDay = dayAdd(todayKey(), 1); renderHaichi(); };
  $("hCopy").onclick = async () => { try { await navigator.clipboard.writeText(haichiTSV()); toast("コピーしました。Excel に貼り付けられます"); } catch (e) { toast("コピーできませんでした"); } };
  $("haichi").addEventListener("click", e => {
    const a = e.target.closest("[data-open]");
    if (a) { e.preventDefault(); S.site = a.dataset.open; document.querySelector('#views [data-tab="sheet"]').click(); setDay(haichiDay); }
  });
  $("haichi").addEventListener("change", e => {
    const c = e.target.closest("[data-ck]"); if (!c) return;
    const id = c.dataset.ck, k = haichiDay, on = c.checked;
    mutate(id, monthOf(k), doc => { const x = ensureDay(doc, id, k); x.確認 = on ? { だれ: me.name, いつ: new Date().toLocaleString("ja-JP") } : null; });
  });
  $("month").addEventListener("click", e => { const a = e.target.closest("[data-goto]"); if (a) { e.preventDefault(); document.querySelector('#views [data-tab="sheet"]').click(); setDay(a.dataset.goto); } });
  $("bMonthCopy").onclick = async () => { const t = $("month").querySelector("table"); if (!t) return; const txt = [...t.rows].map(r => [...r.cells].map(c => c.textContent).join("\t")).join("\r\n"); try { await navigator.clipboard.writeText(txt); toast("コピーしました。Excel に貼り付けられます"); } catch (e) { toast("コピーできませんでした"); } };
  $("bReload").onclick = () => { pull(); pullKikai(); toast("最新を読み込みました"); };
  window.addEventListener("beforeunload", e => { if (Object.values(S.N).some(f => f.pending.length || f.saving)) { e.preventDefault(); e.returnValue = ""; } });
  document.addEventListener("visibilitychange", () => { if (!document.hidden) pull(); });
  // シートの外を押しても、表にキーが届くように（Excel と同じく、いつも表が受ける）
  $("sheetWrap").addEventListener("mousedown", e => { if (e.target === $("sheetWrap") || e.target === $("sheet") || e.target.classList.contains("sec")) { e.preventDefault(); curGrid().focus(); } });
}

async function start() {
  let v = {};
  try { v = JSON.parse(localStorage.getItem("nippou_view")) || {}; } catch (e) { }
  S.day = v.day || todayKey();
  if (DEMO) {
    seedDemo();
    me.name = "試し";
    $("who").textContent = "試し（このブラウザの中だけ。OneDrive には書きません）";
  } else {
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: new URL("./", location.href).href },
      cache: { cacheLocation: "localStorage" }
    });
    await msalApp.initialize();
    try { const r = await msalApp.handleRedirectPromise(); if (r) msalApp.setActiveAccount(r.account); }
    catch (e) { $("signinMsg").textContent = "サインインできませんでした：" + e.message; }
    const acc = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
    if (!acc) { $("signin").hidden = false; $("app").hidden = true; $("bSignIn").onclick = () => msalApp.loginRedirect({ scopes: SCOPES }); return; }
    msalApp.setActiveAccount(acc);
    me.name = acc.name || acc.username;
    $("who").textContent = me.name;
  }
  bind();
  setupGrids();
  ready = true;
  await Promise.all([pull(), pullKikai()]);
  const all = sites();
  const mineSite = all.find(s => !s.完了 && s.担当 && String(s.担当).replace(/[\s　]/g, "") === String(me.name).replace(/[\s　]/g, ""));
  S.site = (v.site && siteById(v.site)) ? v.site : (mineSite || all.find(s => !s.完了) || all[0] || {}).id || "";
  setDay(S.day);
  loadWeather();
  setInterval(pull, POLL_MS);
  setInterval(pullKikai, 5 * 60 * 1000);
}

/* ---------- 試し（?demo）のデータ：今日の前後の工程・機械の予約（架空の名前） ---------- */
function seedDemo() {
  const key = DEMO_KEY;
  try { if (localStorage.getItem(key)) return; } catch (e) { return; }
  const t = todayKey(), a = n => dayAdd(t, n);
  const put = (name, data) => { const all = JSON.parse(localStorage.getItem(key) || "{}"); all[name] = { data, eTag: "e" + Math.random().toString(36).slice(2) }; localStorage.setItem(key, JSON.stringify(all)); };
  put("工程/マスタ.json", { 社員: ["見本　一郎", "見本　二郎", "試験　花子"], 従業員: ["作業　三郎", "作業　四郎", "作業　五郎", "作業　六郎"], 現場: [] });
  put("工程/休日.json", { 休み: [] });
  put("工程/現場_s1.json", { id: "s1", 名前: "A現場 配水管", 担当: "見本　一郎", コード: "510001", 並び: 1,
    段: [{ id: "l1", 区分: "夜", 名前: "管布設工", 土: false }, { id: "l2", 区分: "夜", 名前: "路面復旧工", 種別: "外注", 業者: "第一カッター興業", 人数: "1" }],
    バー: [{ id: "b1", 段: "l1", 作業: "配水管布設", 開始: a(-10), 終了: a(10), 社員: ["見本　一郎"], 人数: "1+3", 色: "青" },
      { id: "b2", 段: "l2", 作業: "舗装切断工", 開始: a(-2), 終了: a(3), 社員: [], 色: "橙" }] });
  put("工程/現場_s2.json", { id: "s2", 名前: "B現場 道路改良", 担当: "見本　二郎", コード: "510002", 並び: 2,
    段: [{ id: "m1", 区分: "昼", 名前: "", 土: true }],
    バー: [{ id: "c1", 段: "m1", 作業: "伐採工", 開始: a(-5), 終了: a(2), 社員: ["見本　二郎"], 人数: "2", 色: "緑" },
      { id: "c2", 段: "m1", 作業: "掘削工", 開始: a(3), 終了: a(15), 社員: ["見本　二郎", "試験　花子"], 人数: "1+2", 色: "茶" }] });
  put("機械/稼働データ.json", { 機械: [{ id: "k1", 名前: "バックホウ", 規格: "0.25㎥" }, { id: "k2", 名前: "規制車", 規格: "2t" }],
    予約: [{ id: "y1", 機械: "k1", 区分: "昼", 現場: "B現場 道路改良", 開始: a(-5), 終了: a(15), 土: true, 日祝: false, 休み: [], 出勤: [], 状態: "確定" },
      { id: "y2", 機械: "k2", 区分: "夜", 現場: "A現場 配水管", 開始: a(-10), 終了: a(10), 土: false, 日祝: false, 休み: [], 出勤: [], 状態: "確定" }] });
  // A現場の前の日の日報（作業員の名前・施工箇所・リース品が次の日の予定に引き継がれるのを見せる）
  const y = a(-1), 表 = {
    基本: [["○○町2丁目", "管布設工", "21:00", "6:00", "配水管布設 φ150 L=12m"], ["○○町2丁目", "路面復旧工", "21:00", "6:00", "舗装切断工"]],
    社員: [["見本　一郎", "管布設工"]], 作業員: [["作業　三郎", "管布設工"], ["作業　四郎", "管布設工"], ["作業　五郎", "管布設工"]],
    外注: [["第一カッター興業", "1", "路面復旧工"]], 材料: [["ダクタイル鋳鉄管", "GX形 φ150", "3本", "○○管材", ""], ["砕石", "RC-40", "4㎥", "○○建材", ""]],
    機械: [["規制車", "2t", "1", "自社（夜）"], ["バックホウ", "0.1㎥", "1", "東リース"]] };
  put("日報/" + `日報_s1_${monthOf(y)}.json`, { 形式: "工事日報", 版: 2, 現場: "s1", 月: monthOf(y),
    日: { [y]: { 天候: "晴", 表, 予定表: 表, 予定日時: "", 報告: { だれ: "見本　一郎", いつ: "" }, 確認: null } } });
}

start();
