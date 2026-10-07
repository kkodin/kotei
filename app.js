/* 全体工程表（2026-10-07）
   画面は GitHub Pages、データは会社の OneDrive「工程データ」フォルダ（Microsoft Graph）。
   現場ごとに1ファイル（現場_<id>.json）。同時に書いたときは eTag で止め、読み直して自分の変更を重ねて保存する。
   ?demo を付けて開くと、サインインせずこのブラウザの中だけで試せる（OneDrive には書かない）。 */
"use strict";

const CFG = window.KOTEI_CONFIG;
const SCOPES = ["User.Read", "Files.ReadWrite.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const DEMO = new URLSearchParams(location.search).has("demo");
const POLL_MS = 30000;
const SHOW_EMP = false;   // 従業員の名前を入れる（将来の仕様。いまは作業員は人数だけ）
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------- 日付（"YYYY-MM-DD" の文字列で持つ。時差の狂いを避ける） ---------- */
const pad2 = n => String(n).padStart(2, "0");
const keyOf = d => d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
const dateOf = k => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
const dayAdd = (k, n) => { const d = dateOf(k); d.setDate(d.getDate() + n); return keyOf(d); };
const dayDow = k => dateOf(k).getDay();
const dayDiff = (a, b) => Math.round((dateOf(b) - dateOf(a)) / 86400000);
const todayKey = () => keyOf(new Date());
const md = k => { const d = dateOf(k); return (d.getMonth() + 1) + "/" + d.getDate(); };
const WD = "日月火水木金土";

/* 日本の祝日（原価管理 Vr013 の jpHolidays と同じ決まり） */
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

/* ---------- 色 ---------- */
const COLORS = [
  ["青", "#3d7cc9"], ["緑", "#3f9a45"], ["黄", "#f2c230"], ["橙", "#ef7d1a"],
  ["赤", "#d93a36"], ["紫", "#8a4fb0"], ["茶", "#8d6e63"], ["灰", "#8a96a3"], ["黒", "#37414b"]
];
const colorOf = n => (COLORS.find(c => c[0] === n) || COLORS[0])[1];
const isLight = n => n === "黄";

/* ---------- 保存先：OneDrive（Graph）／試し（このブラウザの中） ---------- */
let msalApp = null, me = { name: "" };

async function token() {
  try {
    return (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) await msalApp.acquireTokenRedirect({ scopes: SCOPES });
    throw e;
  }
}
async function g(path, opt = {}) {
  return fetch(GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + await token(), ...(opt.headers || {}) } });
}
function httpErr(res, what) { const e = new Error(what + " " + res.status); e.status = res.status; return e; }

const GraphStore = {
  base: () => `/drives/${CFG.folder.driveId}/items/${CFG.folder.itemId}`,
  async list() {
    const out = [];
    let url = this.base() + "/children?$top=999";
    while (url) {
      const res = await g(url);
      if (!res.ok) throw httpErr(res, "一覧");
      const j = await res.json();
      j.value.forEach(x => out.push({ name: x.name, eTag: x.eTag, url: x["@microsoft.graph.downloadUrl"] }));
      url = j["@odata.nextLink"] ? j["@odata.nextLink"].replace(GRAPH, "") : null;
    }
    return out;
  },
  async get(name) {
    const p = this.base() + ":/" + encodeURIComponent(name);
    const res = await g(p);
    if (res.status === 404) return null;
    if (!res.ok) throw httpErr(res, "読み込み");
    const meta = await res.json();
    const dl = meta["@microsoft.graph.downloadUrl"];
    const body = dl ? await fetch(dl, { cache: "no-store" }) : await g(p + ":/content");
    if (!body.ok) throw httpErr(body, "中身の読み込み");
    return { data: await body.json(), eTag: meta.eTag };
  },
  // eTag があれば If-Match（違えば 412）。無ければ新規（あれば 409）
  async put(name, data, eTag) {
    const p = this.base() + ":/" + encodeURIComponent(name) + ":/content" + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
    const res = await g(p, { method: "PUT", headers: { "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: JSON.stringify(data, null, 1) });
    if (!res.ok) throw httpErr(res, "保存");
    return { eTag: (await res.json()).eTag };
  }
};

const DemoStore = {
  key: "kotei_demo_v1",
  all() { try { return JSON.parse(localStorage.getItem(this.key)) || {}; } catch (e) { return {}; } },
  save(a) { try { localStorage.setItem(this.key, JSON.stringify(a)); } catch (e) { } },
  async list() { const a = this.all(); return Object.keys(a).map(n => ({ name: n, eTag: a[n].eTag })); },
  async get(name) { const a = this.all(); return a[name] ? { data: JSON.parse(JSON.stringify(a[name].data)), eTag: a[name].eTag } : null; },
  async put(name, data, eTag) {
    const a = this.all(), cur = a[name];
    if (eTag ? (!cur || cur.eTag !== eTag) : cur) { const e = new Error("保存 " + (eTag ? 412 : 409)); e.status = eTag ? 412 : 409; throw e; }
    const t = "d" + Date.now() + Math.random().toString(36).slice(2, 6);
    a[name] = { data: JSON.parse(JSON.stringify(data)), eTag: t }; this.save(a);
    return { eTag: t };
  }
};
let store = DEMO ? DemoStore : GraphStore;

/* ---------- 状態 ---------- */
// S.files[name] = { doc, eTag, pending:[fn], saving, timer, err }
const S = { files: {}, view: { from: "", months: 2 }, tab: "chart" };
const newId = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const siteFiles = () => Object.keys(S.files).filter(n => n.startsWith("現場_") && S.files[n].doc && !S.files[n].doc.消した);
const sites = () => siteFiles().map(n => S.files[n].doc).sort((a, b) => (a.並び ?? 0) - (b.並び ?? 0) || String(a.名前).localeCompare(b.名前, "ja"));
// 終わった現場（完了）は、上の「完了も出す」にチェックしたときだけ出す（その人のブラウザに覚える）
let showDone = false;
try { showDone = localStorage.getItem("kotei_showDone") === "1"; } catch (e) { }
const shownSites = () => sites().filter(s => showDone || !s.完了);
const fileOfSite = id => "現場_" + id + ".json";
const stamp = d => { d.更新 = { だれ: me.name, いつ: new Date().toLocaleString("ja-JP") }; };

// 変更は「関数」で持つ。保存がぶつかったら、読み直した最新の中身に同じ関数をもう一度かける
function mutate(name, fn) {
  const f = S.files[name];
  fn(f.doc); stamp(f.doc);
  f.pending.push(fn);
  scheduleSave(name);
  render();
}
function scheduleSave(name, ms = 700) {
  const f = S.files[name];
  clearTimeout(f.timer);
  f.timer = setTimeout(() => saveFile(name), ms);
  showSync();
}
async function saveFile(name) {
  const f = S.files[name];
  if (f.saving) { scheduleSave(name, 500); return; }
  if (!f.pending.length) return;
  f.saving = true; showSync();
  try {
    for (let tries = 0; tries < 5; tries++) {
      const n = f.pending.length;
      try {
        const r = await store.put(name, f.doc, f.eTag);
        f.eTag = r.eTag; f.pending.splice(0, n); f.err = null;
        break;
      } catch (e) {
        if (e.status !== 412 && e.status !== 409) throw e;
        // ほかの人が先に保存した → 最新を読んで、自分の変更を重ね直す
        const fresh = await store.get(name);
        const d = fresh ? fresh.data : f.doc;
        f.pending.forEach(fn => fn(d)); stamp(d);
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
  const files = Object.values(S.files);
  const busy = files.some(f => f.pending.length || f.saving), err = files.some(f => f.err);
  const el = $("sync");
  el.className = "sync " + (err ? "err" : busy ? "busy" : "ok");
  el.textContent = err ? "保存できていません" : busy ? "保存中…" : "保存済み";
}

// ほかの人の書き込みを取り込む（自分の未保存の変更があるファイルは、保存のときに重ねるので飛ばす）
let pulling = false, ready = false;   // ready：サインインが済んで読み込みを始めてよい
async function pull() {
  if (pulling || !ready) return;
  pulling = true;
  try {
    const items = await store.list();
    let changed = false;
    for (const it of items) {
      if (!it.name.endsWith(".json") || it.name === "接続テスト.json") continue;
      const f = S.files[it.name];
      if (f && (f.eTag === it.eTag || f.pending.length || f.saving)) continue;
      const r = await store.get(it.name);
      if (!r) continue;
      S.files[it.name] = { doc: r.data, eTag: r.eTag, pending: [], saving: false, timer: null, err: null };
      changed = true;
    }
    if (changed) render();
    showSync();
  } catch (e) {
    console.error(e);
    if (e.status === 404 || e.status === 403) toast("工程データ フォルダを開けません。北澤さんに共有（編集可）を頼んでください");
    else toast("最新の読み込みに失敗：" + e.message);
  } finally { pulling = false; }
}

/* ---------- マスタ（工程データ\マスタ.json。配置表入力の名簿から tools\マスタを書き出す.py で作る） ---------- */
const M = () => (S.files["マスタ.json"] && S.files["マスタ.json"].doc) || { 社員: [], 従業員: [], 現場: [] };
const nk = n => String(n).replace(/[\s　]+/g, "");
const surname = n => String(n).split(/[\s　]+/)[0];
// 打った名前をマスタの書き方に寄せる（「吉成」→「吉成　直也」。名字が同じ人が2人いるときは寄せない）
function resolveName(n, list) {
  if (!n) return n;
  const hit = list.find(x => x === n) || list.find(x => nk(x) === nk(n));
  if (hit) return hit;
  const c = list.filter(x => nk(x).startsWith(nk(n)));
  return c.length === 1 ? c[0] : n;
}
// 前の作りで「吉成　直也」が「吉成」「直也」に分かれて保存されたものを、マスタの氏名に戻す
function joinSplit(names, list) {
  const out = [];
  for (let i = 0; i < names.length; i++) {
    const two = i + 1 < names.length && list.find(x => nk(x) === nk(names[i] + names[i + 1]));
    if (two) { out.push(two); i++; } else out.push(names[i]);
  }
  return out;
}
// バーに入っている名前（マスタに無い人も出す）
function usedNames(field) {
  const set = new Set();
  sites().forEach(s => { if (field === "社員" && s.担当) set.add(s.担当); s.バー.forEach(b => (b[field] || []).forEach(n => set.add(n))); });
  return set;
}
function namesFor(field) {
  const base = M()[field] || [];
  const extra = [...usedNames(field)].filter(n => !base.includes(n)).sort((a, b) => a.localeCompare(b, "ja"));
  return [...base, ...extra];
}

/* ---------- 表示する期間 ---------- */
function range() {
  const from = S.view.from + "-01";
  const end = dayAdd(keyOf(new Date(+S.view.from.slice(0, 4), +S.view.from.slice(5, 7) - 1 + S.view.months, 1)), -1);
  const days = [];
  for (let k = from; k <= end; k = dayAdd(k, 1)) days.push(k);
  return { from, end, days };
}
/* 会社の休み（工程データ\休日.json。日付の見出しを押して切り替える。全員で共有） */
const REST_FILE = "休日.json";
const isRest = k => { const f = S.files[REST_FILE]; return !!(f && f.doc && (f.doc.休み || []).includes(k)); };
function toggleRest(k) {
  if (!S.files[REST_FILE]) S.files[REST_FILE] = { doc: { 休み: [] }, eTag: null, pending: [], saving: false, timer: null, err: null };
  const on = !isRest(k);
  mutate(REST_FILE, d => {
    const set = new Set(d.休み || []);
    if (on) set.add(k); else set.delete(k);
    d.休み = [...set].sort();
  });
  toast(`${md(k)}（${WD[dayDow(k)]}）を${on ? "会社の休みにしました" : "休みから外しました"}`);
}

/* 段ごとの作業日（2026-10-07 ユーザー：現場ごとに土曜・日曜・祝日に作業するかが違う。昼夜で違うこともある）
   段.土／日／祝 … その曜日も作業するか（決めていなければ 土＝する・日＝しない・祝＝しない）
   段.休み／出勤 … その日だけの例外（段の上の日付を右クリックで切り替え） */
function ruleWorks(ln, k) {
  if (holidayOf(k)) return ln.祝 === true;
  const w = dayDow(k);
  if (w === 0) return ln.日 === true;
  if (w === 6) return ln.土 !== false;
  return true;
}
function laneWorks(ln, k) {
  if (!ln || isRest(k)) return false;
  if ((ln.休み || []).includes(k)) return false;
  if ((ln.出勤 || []).includes(k)) return true;
  return ruleWorks(ln, k);
}
// 工程（バー）ごとの休工日（バーの小窓の「休工日を設定」）。段の決まりの上にかける
function barWorks(ln, b, k) {
  if (isRest(k)) return false;
  if (b && (b.休み || []).includes(k)) return false;
  if (b && (b.出勤 || []).includes(k)) return true;
  return laneWorks(ln, k);
}
function toggleLaneDay(siteId, laneId, k) {
  const s = S.files[fileOfSite(siteId)].doc, ln = s.段.find(l => l.id === laneId);
  if (isRest(k)) { toast("この日は会社の休みです（日付の見出しを押すと戻せます）"); return; }
  const on = !laneWorks(ln, k);   // 押したあと作業する日になるか
  mutate(fileOfSite(siteId), d => {
    const l = d.段.find(x => x.id === laneId);
    if (!l) return;
    l.休み = (l.休み || []).filter(x => x !== k);
    l.出勤 = (l.出勤 || []).filter(x => x !== k);
    if (on && !ruleWorks(l, k)) l.出勤.push(k);
    if (!on && ruleWorks(l, k)) l.休み.push(k);
  });
  toast(`${s.名前}（${ln.名前}）${md(k)}（${WD[dayDow(k)]}）を${on ? "作業する日" : "休み"}にしました`);
}

function dayClass(k) {
  if (isRest(k)) return "rest";
  if (k === todayKey()) return "today";
  if (holidayOf(k)) return "hol";
  const w = dayDow(k);
  return w === 0 ? "sun" : w === 6 ? "sat" : "";
}

/* ---------- 工程表を描く ---------- */
const DW = () => parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--dw"));

function barLabel(b) {
  const parts = [];
  if (b.作業) parts.push(esc(b.作業));
  if (b.人数) parts.push(esc(b.人数) + "名");
  if (b.社員 && b.社員.length) parts.push(`<span class="st">${esc(b.社員.join("・"))}</span>`);
  return parts.join(" ");
}
function barTitle(s, b) {
  const ln = s.段.find(l => l.id === b.段);
  let work = 0;
  for (let k = b.開始; k <= b.終了; k = dayAdd(k, 1)) if (barWorks(ln, b, k)) work++;
  return `${s.名前}／${b.作業 || ""}\n${md(b.開始)}〜${md(b.終了)}（${dayDiff(b.開始, b.終了) + 1}日・うち作業 ${work}日）` +
    (b.社員 && b.社員.length ? "\n社員：" + b.社員.join("・") : "") + (b.人数 ? "\n作業員：" + b.人数 + "名" : "") + (b.メモ ? "\n" + b.メモ : "");
}

/* 同じ段で日が重なるバー（複数の工種）は、下に行を足して並べる。開始の早い順に、空いている一番上の行へ */
function packLane(bars) {
  const ends = [], pos = {};
  [...bars].sort((a, b) => a.開始.localeCompare(b.開始) || b.終了.localeCompare(a.終了)).forEach(b => {
    let r = ends.findIndex(e => e < b.開始);
    if (r < 0) { r = ends.length; ends.push(""); }
    ends[r] = b.終了; pos[b.id] = r;
  });
  return { pos, rows: Math.max(1, ends.length) };
}
/* 現場のたたみ（その人のブラウザだけに覚える） */
const FOLD_KEY = "kotei_fold";
let folded = new Set();
try { folded = new Set(JSON.parse(localStorage.getItem(FOLD_KEY)) || []); } catch (e) { }
function setFold(ids, on) {
  ids.forEach(id => on ? folded.add(id) : folded.delete(id));
  try { localStorage.setItem(FOLD_KEY, JSON.stringify([...folded])); } catch (e) { }
  render();
}
// たたんだ現場の帯：工事のある日をつないだ期間
function spansOf(s, from, end) {
  const ks = new Set();
  s.バー.forEach(b => { for (let k = b.開始 < from ? from : b.開始; k <= b.終了 && k <= end; k = dayAdd(k, 1)) ks.add(k); });
  const out = [];
  [...ks].sort().forEach(k => { const last = out[out.length - 1]; if (last && dayAdd(last.z, 1) === k) last.z = k; else out.push({ a: k, z: k }); });
  return out;
}

function renderChart() {
  const { from, end, days } = range();
  const dw = DW(), rh = parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--rh"));
  const width = days.length * dw;
  const h = [];
  // 見出し：月・日・曜
  h.push(`<div class="hrow m"><div class="corner">${esc(from.slice(0, 4))}年 <button type="button" class="sub small fold" data-fold="all">▸ 全部たたむ</button><button type="button" class="sub small fold" data-fold="none">▾ 全部ひらく</button></div>` +
    days.map(k => `<div class="hd ${k.endsWith("-01") ? "mon" : ""}">${k.endsWith("-01") ? (+k.slice(5, 7)) + "月" : ""}</div>`).join("") + "</div>");
  h.push(`<div class="hrow d"><div class="corner"></div>` +
    days.map(k => `<div class="hd ${dayClass(k)}" data-day="${k}" title="${esc((holidayOf(k) ? holidayOf(k) + "\n" : "") + "押すと会社の休みにする／戻す")}">${+k.slice(8)}</div>`).join("") + "</div>");
  h.push(`<div class="hrow w"><div class="corner">現場／段</div>` +
    days.map(k => `<div class="hd ${dayClass(k)}" data-day="${k}" title="押すと会社の休みにする／戻す">${isRest(k) ? "休" : WD[dayDow(k)]}</div>`).join("") + "</div>");
  // 本体
  h.push(`<div id="body">`);
  const left = `calc(var(--sw) + var(--lw))`;
  days.forEach((k, i) => { const c = dayClass(k); if (c) h.push(`<div class="col ${c}" style="left:calc(${left} + ${i * dw}px)"></div>`); });
  for (const s of shownSites()) {
    const fold = folded.has(s.id);
    h.push(`<div class="site ${fold ? "folded" : ""} ${s.完了 ? "done" : ""}" data-site="${esc(s.id)}"><div class="sname" data-act="site"><span class="tg" title="${fold ? "ひらく" : "たたむ"}">${fold ? "▸" : "▾"}</span>${s.完了 ? '<span class="donemark">完了</span>' : ""}${esc(s.名前)}` +
      (fold ? "" : (s.コード ? `<span class="code">${esc(s.コード)}</span>` : "") + (s.担当 ? `<span class="tanto">担当 ${esc(s.担当)}</span>` : "")) + `</div><div class="lanes">`);
    if (fold) {
      // たたんだ現場：1行に、工事のある期間だけを帯で出す
      const inR = s.バー.filter(b => b.終了 >= from && b.開始 <= end);
      const works = [...new Set(inR.map(b => b.作業).filter(Boolean))];
      h.push(`<div class="lane"><div class="llab" data-act="fold" title="押すとひらく">${inR.length ? works.length + "工種" : "―"}</div><div class="track sum" style="width:${width}px">`);
      for (const p of spansOf(s, from, end)) {
        const x = dayDiff(from, p.a) * dw, w = (dayDiff(p.a, p.z) + 1) * dw;
        h.push(`<div class="bar sumbar" data-act="fold" style="left:${x}px;width:${w}px" title="${esc(md(p.a) + "〜" + md(p.z) + "\n" + works.join("・") + "\n押すとひらく")}">${esc(works.join("・"))}</div>`);
      }
      h.push(`</div></div>`);
    } else for (const ln of s.段) {
      const bars = s.バー.filter(b => b.段 === ln.id && b.終了 >= from && b.開始 <= end);
      const { pos, rows } = packLane(bars);
      h.push(`<div class="lane" style="height:${rows * rh}px"><div class="llab" data-act="site" title="${esc(ln.名前 + "\n" + laneRuleText(ln))}">${esc(ln.名前)}</div><div class="track" data-site="${esc(s.id)}" data-lane="${esc(ln.id)}" style="width:${width}px">`);
      // 段の休みの日に薄い斜線。日曜・祝日は赤い色で分かるので斜線は付けない（2026-10-07 ユーザー指定）。右クリックで休みにした日は付ける
      days.forEach((k, i) => {
        if (isRest(k) || laneWorks(ln, k)) return;
        if ((dayDow(k) === 0 || holidayOf(k)) && !(ln.休み || []).includes(k)) return;
        h.push(`<div class="off" style="left:${i * dw}px"></div>`);
      });
      for (const b of bars) {
        const a = b.開始 < from ? from : b.開始, z = b.終了 > end ? end : b.終了;
        const x = dayDiff(from, a) * dw, w = (dayDiff(a, z) + 1) * dw;
        // 1本のバーの中で、その段が休む日は白く抜いて細い線でつなぐ（「8〜23日、この日とこの日は休み」を1本で書けるように）
        const gaps = [];
        for (let k = a, i = 0; k <= z; k = dayAdd(k, 1), i++) {
          if (barWorks(ln, b, k)) continue;
          const g = gaps[gaps.length - 1];
          if (g && g.i + g.n === i) g.n++; else gaps.push({ i, n: 1 });
        }
        h.push(`<div class="bar ${isLight(b.色) ? "dark" : ""} ${b.開始 < from ? "cut-l" : ""} ${b.終了 > end ? "cut-r" : ""}" data-bar="${esc(b.id)}" ` +
          `style="left:${x}px;top:${pos[b.id] * rh + 3}px;width:${w}px;--bc:${colorOf(b.色)}" title="${esc(barTitle(s, b))}">` +
          gaps.map(g => `<span class="gap" style="left:${g.i * dw}px;width:${g.n * dw}px"></span>`).join("") +
          `<span class="h l"></span><span class="lbl">${barLabel(b)}</span><span class="h r"></span></div>`);
      }
      h.push(`</div></div>`);
    }
    h.push(`</div></div>`);
  }
  h.push(`</div><div class="addsite"><button id="bAddSite" class="sub">＋ 現場を追加</button></div>`);
  const hidden = sites().length - shownSites().length;
  if (hidden) h.push(`<div class="addsite note">完了した現場 ${hidden} 件を隠しています（上の「完了も出す」で表示）</div>`);
  $("chart").innerHTML = h.join("");
  $("bAddSite").onclick = () => openSite(null);
}

/* ---------- 一覧 ---------- */
function renderList() {
  const { from, end } = range();
  const only = $("inListRange").checked, q = $("inListQ").value.trim();
  const rows = [];
  for (const s of shownSites()) for (const b of s.バー) {
    if (only && (b.終了 < from || b.開始 > end)) continue;
    const ln = s.段.find(l => l.id === b.段);
    const text = [s.名前, ln && ln.名前, b.作業, (b.社員 || []).join(" "), (b.従業員 || []).join(" "), b.メモ].join(" ");
    if (q && !q.split(/\s+/).every(w => text.includes(w))) continue;
    rows.push({ s, b, ln });
  }
  rows.sort((x, y) => x.b.開始.localeCompare(y.b.開始) || (x.s.並び ?? 0) - (y.s.並び ?? 0));
  const tot = rows.reduce((a, r) => a + (dayDiff(r.b.開始, r.b.終了) + 1), 0);
  $("list").innerHTML = `<table><tr><th>現場</th><th>段</th><th>作業</th><th>開始</th><th>終了</th><th>日数</th><th>社員</th>${SHOW_EMP ? "<th>従業員</th>" : ""}<th>作業員</th><th>メモ</th><th>更新</th></tr>` +
    rows.map(({ s, b, ln }) => `<tr class="r" data-site="${esc(s.id)}" data-bar="${esc(b.id)}"><td>${esc(s.名前)}</td><td>${esc(ln ? ln.名前 : "")}</td>` +
      `<td><span class="chip" style="background:${colorOf(b.色)}"></span>${esc(b.作業)}</td><td>${md(b.開始)}（${WD[dayDow(b.開始)]}）</td><td>${md(b.終了)}（${WD[dayDow(b.終了)]}）</td>` +
      `<td class="n">${dayDiff(b.開始, b.終了) + 1}</td><td>${esc((b.社員 || []).join("・"))}</td>${SHOW_EMP ? `<td>${esc((b.従業員 || []).join("・"))}</td>` : ""}<td class="n">${b.人数 ? esc(b.人数) + "名" : ""}</td>` +
      `<td>${esc(b.メモ)}</td><td class="note">${esc(b.更新 ? b.更新.だれ : "")}</td></tr>`).join("") +
    `</table><p class="note">${rows.length} 件（のべ ${tot} 日）</p>`;
}

/* ---------- 人の流れ（工程表のバーから、社員の行き先と作業員の日ごとの人数を出す） ---------- */
const shortOf = s => s.略称 || String(s.名前).slice(0, 4);
const isNight = ln => /夜/.test(ln ? ln.名前 : "");

function renderFlow() {
  const { from, end, days } = range();
  // 日ごとに、その日にかかっているバーを集める
  const at = {};   // day -> [{s, b, ln}]
  for (const s of sites()) for (const b of s.バー) {
    if (b.終了 < from || b.開始 > end) continue;
    const ln = s.段.find(l => l.id === b.段);
    for (let k = b.開始 < from ? from : b.開始; k <= b.終了 && k <= end; k = dayAdd(k, 1)) if (barWorks(ln, b, k)) (at[k] = at[k] || []).push({ s, b, ln });
  }
  const staff = namesFor("社員"), emps = namesFor("従業員");
  const head = `<thead><tr><th class="nm"></th>` +
    days.map(k => `<th class="${dayClass(k)}" title="${esc(holidayOf(k))}">${k.endsWith("-01") || k === from ? `<b>${+k.slice(5, 7)}/</b>` : ""}${+k.slice(8)}</th>`).join("") +
    `</tr><tr><th class="nm"></th>` + days.map(k => `<th class="${dayClass(k)}">${isRest(k) ? "休" : WD[dayDow(k)]}</th>`).join("") + `</tr></thead>`;
  const cols = days.length + 1;
  const h = [`<table class="fl">`, head, `<tbody>`];

  // 1. 社員・従業員の行き先（1人1行。マスタの全員を出すので、空いている人も分かる）
  const personRows = (field, list) => { for (const n of list) {
    h.push(`<tr><th class="nm">${esc(n)}</th>`);
    for (const k of days) {
      const hits = (at[k] || []).filter(x => (x.b[field] || []).includes(n));
      const siteIds = [...new Set(hits.map(x => x.s.id))];
      if (!hits.length) h.push(`<td class="${dayClass(k)}"></td>`);
      else if (siteIds.length > 1) h.push(`<td class="dup" data-site="${esc(hits[0].s.id)}" data-bar="${esc(hits[0].b.id)}" title="${esc(md(k) + " 重複：" + hits.map(x => x.s.名前 + "／" + (x.b.作業 || "")).join("、"))}">${siteIds.length}現場</td>`);
      else { const x = hits[0];
        h.push(`<td class="as ${isLight(x.b.色) ? "dark" : ""}" style="background:${colorOf(x.b.色)}" data-site="${esc(x.s.id)}" data-bar="${esc(x.b.id)}" title="${esc(md(k) + " " + x.s.名前 + "／" + (x.b.作業 || "") + (x.ln ? "（" + x.ln.名前 + "）" : ""))}">${esc(shortOf(x.s))}</td>`); }
    }
    h.push(`</tr>`);
  } };
  h.push(`<tr class="sec"><th colspan="${cols}">社員の行き先（${staff.length}人）</th></tr>`);
  if (!staff.length) h.push(`<tr><th class="nm note">（社員がいません）</th></tr>`);
  personRows("社員", staff);
  // 社員の入っていない工程
  h.push(`<tr><th class="nm" style="color:#c62828">社員 未定</th>`);
  for (const k of days) {
    const un = (at[k] || []).filter(x => !(x.b.社員 || []).length);
    const ss = [...new Set(un.map(x => shortOf(x.s)))];
    h.push(un.length ? `<td class="un ${dayClass(k)}" data-site="${esc(un[0].s.id)}" data-bar="${esc(un[0].b.id)}" title="${esc(md(k) + " 社員未定：" + un.map(x => x.s.名前 + "／" + (x.b.作業 || "")).join("、"))}">${ss.length > 1 ? ss.length + "件" : esc(ss[0])}</td>` : `<td class="${dayClass(k)}"></td>`);
  }
  h.push(`</tr><tr class="tot"><th class="nm">社員 計</th>`);
  for (const k of days) { const c = new Set((at[k] || []).flatMap(x => x.b.社員 || [])).size; h.push(`<td class="${c ? "" : "n0"} ${dayClass(k)}">${c || ""}</td>`); }
  h.push(`</tr>`);

  if (SHOW_EMP) {
    h.push(`<tr class="sec"><th colspan="${cols}">従業員の行き先（${emps.length}人）</th></tr>`);
    personRows("従業員", emps);
    h.push(`<tr class="tot"><th class="nm">従業員 計</th>`);
    for (const k of days) { const c = new Set((at[k] || []).flatMap(x => x.b.従業員 || [])).size; h.push(`<td class="${c ? "" : "n0"} ${dayClass(k)}">${c || ""}</td>`); }
    h.push(`</tr>`);
  }

  // 2. 作業員の人数（現場ごと・昼夜）
  h.push(`<tr class="sec"><th colspan="${cols}">作業員の人数（工程表のバーの「作業員」の合計。段の名前に「夜」があれば夜間）</th></tr>`);
  const sum = (k, f) => (at[k] || []).filter(f).reduce((a, x) => a + (+x.b.人数 || 0), 0);
  for (const s of shownSites()) {
    if (!s.バー.some(b => b.人数 && b.終了 >= from && b.開始 <= end)) continue;
    h.push(`<tr><th class="nm">${esc(s.名前)}</th>` + days.map(k => { const v = sum(k, x => x.s === s); return `<td class="${v ? "" : "n0"} ${dayClass(k)}">${v || ""}</td>`; }).join("") + `</tr>`);
  }
  [["昼間 計", x => !isNight(x.ln)], ["夜間 計", x => isNight(x.ln)], ["作業員 計", () => true]].forEach(([nm, f]) => {
    h.push(`<tr class="tot"><th class="nm">${nm}</th>` + days.map(k => { const v = sum(k, f); return `<td class="${v ? "" : "n0"} ${dayClass(k)}">${v || ""}</td>`; }).join("") + `</tr>`);
  });
  h.push(`</tbody></table>`);
  $("flow").innerHTML = h.join("");
}

function render() {
  if (S.tab === "chart") renderChart(); else if (S.tab === "flow") renderFlow(); else renderList();
  // 候補（マスタ＋入力済み）
  $("staffList").innerHTML = namesFor("社員").map(n => `<option value="${esc(n)}">`).join("");
  $("siteList").innerHTML = M().現場.map(x => `<option value="${esc(x.名前)}">${esc([x.コード, x.略称 && x.略称 !== x.名前 ? "配置表：" + x.略称 : "", x.担当 ? "担当 " + x.担当 : ""].filter(Boolean).join("　"))}</option>`).join("");
}

/* ---------- ドラッグ（動かす・延ばす・作る） ---------- */
let drag = null;
function dayAt(track, clientX) {
  const r = track.getBoundingClientRect();
  return Math.floor((clientX - r.left) / DW());
}
$("wrap").addEventListener("pointerdown", e => {
  if (e.button !== 0) return;
  const barEl = e.target.closest(".bar"), track = e.target.closest(".track");
  if (!track || track.classList.contains("sum")) return;
  const site = S.files[fileOfSite(track.dataset.site)].doc;
  const { from } = range();
  const d0 = dayAt(track, e.clientX);
  if (barEl) {
    const b = site.バー.find(x => x.id === barEl.dataset.bar);
    const mode = e.target.classList.contains("l") ? "l" : e.target.classList.contains("r") ? "r" : "move";
    drag = { kind: "bar", mode, el: barEl, track, site, b, d0, x0: e.clientX, moved: false, a: b.開始, z: b.終了, from };
    barEl.setPointerCapture(e.pointerId);
  } else {
    // 指で触ったときは、なぞると画面が流れるので「押した日に1日のバー」を作る小窓だけ出す
    if (e.pointerType === "touch") { const k = dayAdd(from, d0); openBar(site, null, { 段: track.dataset.lane, 開始: k, 終了: k }); return; }
    const ghost = document.createElement("div");
    ghost.className = "bar ghost";
    track.appendChild(ghost);
    drag = { kind: "new", track, site, lane: track.dataset.lane, d0, d1: d0, el: ghost, from };
    track.setPointerCapture(e.pointerId);
    placeGhost();
  }
  e.preventDefault();
});
function placeGhost() {
  const dw = DW(), a = Math.min(drag.d0, drag.d1), z = Math.max(drag.d0, drag.d1);
  drag.el.style.left = a * dw + "px"; drag.el.style.width = (z - a + 1) * dw + "px";
  drag.el.textContent = `${md(dayAdd(drag.from, a))}〜${md(dayAdd(drag.from, z))}（${z - a + 1}日）`;
}
window.addEventListener("pointermove", e => {
  if (!drag) return;
  if (drag.kind === "new") { drag.d1 = dayAt(drag.track, e.clientX); placeGhost(); return; }
  const dd = Math.round((e.clientX - drag.x0) / DW());
  if (Math.abs(e.clientX - drag.x0) > 4) drag.moved = true;
  if (!drag.moved) return;
  const b = drag.b;
  let a = b.開始, z = b.終了;
  if (drag.mode === "move") { a = dayAdd(b.開始, dd); z = dayAdd(b.終了, dd); }
  else if (drag.mode === "l") { a = dayAdd(b.開始, dd); if (a > z) a = z; }
  else { z = dayAdd(b.終了, dd); if (z < a) z = a; }
  drag.a = a; drag.z = z;
  // 見た目だけ先に動かす（離したときに保存）
  const { from, end } = range(), dw = DW();
  const va = a < from ? from : a, vz = z > end ? end : z;
  drag.el.style.left = dayDiff(from, va) * dw + "px";
  drag.el.style.width = Math.max(dw, (dayDiff(va, vz) + 1) * dw) + "px";
  drag.el.classList.add("drag");
  drag.el.title = `${md(a)}〜${md(z)}（${dayDiff(a, z) + 1}日）`;
  toast(`${md(a)}（${WD[dayDow(a)]}）〜 ${md(z)}（${WD[dayDow(z)]}）　${dayDiff(a, z) + 1}日`, 900);
});
window.addEventListener("pointerup", e => {
  if (!drag) return;
  const d = drag; drag = null;
  if (d.kind === "new") {
    d.el.remove();
    const a = Math.min(d.d0, d.d1), z = Math.max(d.d0, d.d1);
    openBar(d.site, null, { 段: d.lane, 開始: dayAdd(d.from, a), 終了: dayAdd(d.from, z) });
    return;
  }
  if (!d.moved) { openBar(d.site, d.b); return; }
  if (d.a === d.b.開始 && d.z === d.b.終了) { render(); return; }
  const id = d.b.id, a = d.a, z = d.z;
  mutate(fileOfSite(d.site.id), doc => { const b = doc.バー.find(x => x.id === id); if (b) { b.開始 = a; b.終了 = z; stamp(b); } });
});
$("wrap").addEventListener("contextmenu", e => {
  const track = e.target.closest(".track");
  if (!track || track.classList.contains("sum")) return;
  e.preventDefault();
  toggleLaneDay(track.dataset.site, track.dataset.lane, dayAdd(range().from, dayAt(track, e.clientX)));
});
$("wrap").addEventListener("click", e => {
  const fb = e.target.closest("[data-fold]");
  if (fb) { setFold(shownSites().map(s => s.id), fb.dataset.fold === "all"); return; }
  const hd = e.target.closest(".hd[data-day]");
  if (hd) { toggleRest(hd.dataset.day); return; }
  if (e.target.closest(".tg") || e.target.closest("[data-act=fold]")) {
    const id = e.target.closest(".site").dataset.site;
    setFold([id], !folded.has(id)); return;
  }
  if (e.target.closest("[data-act=site]")) openSite(e.target.closest(".site").dataset.site);
});

/* ---------- バーの小窓 ---------- */
let barCtx = null;
$("bColors").innerHTML = COLORS.map(([n, c], i) => `<input type="radio" name="bColor" id="bc${i}" value="${n}"><label for="bc${i}" style="background:${c}" title="${n}"></label>`).join("");
// 名前は「名字　名前」と全角空白を含むので、空白では区切らない
const splitNames = s => s.split(/[、,，・\n]+/).map(x => x.trim()).filter(Boolean);
const toHalf = s => s.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));

function openBar(site, b, init) {
  barCtx = { siteId: site.id, barId: b ? b.id : null, 休み: [...((b && b.休み) || [])], 出勤: [...((b && b.出勤) || [])] };
  const v = b || { 段: init.段, 開始: init.開始, 終了: init.終了, 作業: "", 社員: [], 人数: "", 色: lastColor, メモ: "" };
  $("dBarTtl").textContent = b ? "工程を直す" : "工程を足す";
  $("dBarSite").textContent = site.名前;
  $("bLane").innerHTML = site.段.map(l => `<option value="${esc(l.id)}">${esc(l.名前)}</option>`).join("");
  $("bLane").value = v.段;
  $("bWork").value = v.作業 || "";
  $("bFrom").value = v.開始; $("bTo").value = v.終了;
  $("bStaff").value = joinSplit(v.社員 || [], M().社員 || []).join("、");
  $("bEmp").value = joinSplit(v.従業員 || [], M().従業員 || []).join("、");
  $("bNum").value = v.人数 || "";
  autoNum = SHOW_EMP && (!v.人数 || v.人数 === (v.従業員 || []).length);
  document.querySelectorAll("#dBar .emp").forEach(el => { el.hidden = !SHOW_EMP; });
  drawPick("社員"); drawPick("従業員");
  (document.querySelector(`#bColors input[value="${v.色 || "青"}"]`) || document.querySelector("#bColors input")).checked = true;
  $("bMemo").value = v.メモ || "";
  $("dBarWho").textContent = b && b.更新 ? `最後に直した人：${b.更新.だれ}（${b.更新.いつ}）` : "";
  $("bBarDel").hidden = !b; $("bBarCopy").hidden = !b;
  openDlg($("dBar"));
  showOffInfo();
  if (!b) $("bWork").focus();
}
let lastColor = "青";
function readBarForm() {
  let a = $("bFrom").value, z = $("bTo").value;
  if (z < a) [a, z] = [z, a];
  const n = parseInt(toHalf($("bNum").value), 10);
  return {
    段: $("bLane").value, 作業: $("bWork").value.trim(), 開始: a, 終了: z,
    社員: [...new Set(splitNames($("bStaff").value).map(x => resolveName(x, namesFor("社員"))))],
    従業員: [...new Set(splitNames($("bEmp").value).map(x => resolveName(x, namesFor("従業員"))))], 人数: isNaN(n) ? "" : n,
    色: (document.querySelector("#bColors input:checked") || {}).value || "青", メモ: $("bMemo").value.trim(),
    // 期間の外になった休工日は捨てる
    休み: (barCtx.休み || []).filter(k => k >= a && k <= z).sort(), 出勤: (barCtx.出勤 || []).filter(k => k >= a && k <= z).sort()
  };
}

/* ---------- 休工日のカレンダー ---------- */
let calWork = null;   // カレンダーで直している途中の { 休み:Set, 出勤:Set }
function calLane() { const s = S.files[fileOfSite(barCtx.siteId)].doc; return s.段.find(l => l.id === $("bLane").value); }
function showOffInfo() {
  let a = $("bFrom").value, z = $("bTo").value;
  if (!a || !z) { $("bOffInfo").textContent = ""; return; }
  if (z < a) [a, z] = [z, a];
  const ln = calLane(), b = { 休み: barCtx.休み, 出勤: barCtx.出勤 };
  let n = 0, w = 0;
  for (let k = a; k <= z; k = dayAdd(k, 1)) { n++; if (barWorks(ln, b, k)) w++; }
  const own = barCtx.休み.filter(k => k >= a && k <= z).length;
  $("bOffInfo").textContent = `${n}日のうち作業 ${w}日・休み ${n - w}日` + (own ? `（この工程だけの休工 ${own}日）` : "");
}
["bFrom", "bTo", "bLane"].forEach(id => $(id).addEventListener("change", showOffInfo));
function drawCal() {
  let a = $("bFrom").value, z = $("bTo").value;
  if (z < a) [a, z] = [z, a];
  const ln = calLane(), b = { 休み: [...calWork.休み], 出勤: [...calWork.出勤] };
  const color = colorOf((document.querySelector("#bColors input:checked") || {}).value || "青");
  const h = [];
  let n = 0, w = 0;
  // 期間にかかる月を、日曜はじまりのカレンダーで並べる
  for (let m = a.slice(0, 7) + "-01"; m <= z; m = keyOf(new Date(+m.slice(0, 4), +m.slice(5, 7), 1))) {
    h.push(`<div class="cal"><div class="calm">${+m.slice(0, 4)}年${+m.slice(5, 7)}月</div><div class="calg">` +
      [...WD].map((x, i) => `<div class="calw ${i === 0 ? "sun" : i === 6 ? "sat" : ""}">${x}</div>`).join("") +
      `<div></div>`.repeat(dayDow(m)));
    for (let k = m; k.slice(0, 7) === m.slice(0, 7); k = dayAdd(k, 1)) {
      const hol = holidayOf(k), wd = dayDow(k), num = `<span class="${hol || wd === 0 ? "red" : wd === 6 ? "blue" : ""}">${+k.slice(8)}</span>`;
      if (k < a || k > z) { h.push(`<div class="cald out">${num}</div>`); continue; }
      n++;
      if (isRest(k)) { h.push(`<div class="cald rest" title="会社の休み">${num}<i>休</i></div>`); continue; }
      const on = barWorks(ln, b, k);
      if (on) w++;
      h.push(`<div class="cald ${on ? "con" : "coff"}" data-k="${k}" style="--bc:${color}" title="${esc((hol ? hol + "\n" : "") + (on ? "作業する日（押すと休工）" : "休み（押すと作業する日）"))}">${num}</div>`);
    }
    h.push(`</div></div>`);
  }
  $("calBody").innerHTML = h.join("");
  $("calSum").textContent = `期間 ${n}日　作業 ${w}日　休み ${n - w}日`;
}
$("bOffBtn").onclick = () => {
  if (!$("bFrom").value || !$("bTo").value) { alert("先に開始と終了を入れてください。"); return; }
  calWork = { 休み: new Set(barCtx.休み), 出勤: new Set(barCtx.出勤) };
  $("dCalTtl").textContent = "休工日を設定　" + ($("bWork").value || "（作業名なし）");
  drawCal();
  openDlg($("dCal"));
};
$("calBody").addEventListener("click", e => {
  const c = e.target.closest(".cald[data-k]");
  if (!c) return;
  const k = c.dataset.k, ln = calLane();
  const nowOn = c.classList.contains("con"), base = laneWorks(ln, k);
  calWork.休み.delete(k); calWork.出勤.delete(k);
  if (nowOn && base) calWork.休み.add(k);        // 作業する日 → この工程だけ休工
  if (!nowOn && !base) calWork.出勤.add(k);      // 段の休みの日 → この工程だけ作業
  drawCal();
});
$("bCalReset").onclick = () => { calWork = { 休み: new Set(), 出勤: new Set() }; drawCal(); };
$("bCalCancel").onclick = () => $("dCal").close();
$("fCal").addEventListener("submit", e => {
  e.preventDefault();
  barCtx.休み = [...calWork.休み].sort(); barCtx.出勤 = [...calWork.出勤].sort();
  $("dCal").close();
  showOffInfo();
});

/* ---------- 小窓：見出しをつかんで動かす。枠の外を押しても閉じない（閉じるのは ボタン か Esc） ---------- */
function openDlg(d) {
  d.style.margin = ""; d.style.left = ""; d.style.top = "";
  d.showModal();
}
document.querySelectorAll("dialog h3").forEach(h3 => {
  h3.addEventListener("pointerdown", e => {
    if (e.button !== 0) return;
    const d = h3.closest("dialog"), r = d.getBoundingClientRect();
    const dx = e.clientX - r.left, dy = e.clientY - r.top;
    d.style.margin = "0"; d.style.left = r.left + "px"; d.style.top = r.top + "px";
    try { h3.setPointerCapture(e.pointerId); } catch (x) { }
    const mv = ev => {
      d.style.left = Math.min(Math.max(0, ev.clientX - dx), innerWidth - 80) + "px";
      d.style.top = Math.min(Math.max(0, ev.clientY - dy), innerHeight - 40) + "px";
    };
    const up = () => { h3.removeEventListener("pointermove", mv); h3.removeEventListener("pointerup", up); };
    h3.addEventListener("pointermove", mv); h3.addEventListener("pointerup", up);
    e.preventDefault();
  });
});
// 枠の外（背景）を押しても閉じない：背景で始まった押し下げは何もしない
document.querySelectorAll("dialog").forEach(d => d.addEventListener("click", e => { if (e.target === d) e.stopPropagation(); }));
$("fBar").addEventListener("submit", e => {
  e.preventDefault();
  if (!$("bFrom").value || !$("bTo").value) return;
  const v = readBarForm(), { siteId, barId } = barCtx;
  lastColor = v.色;
  const id = barId || newId("b");
  mutate(fileOfSite(siteId), doc => {
    let b = doc.バー.find(x => x.id === id);
    if (!b) { b = { id }; doc.バー.push(b); }
    Object.assign(b, v); stamp(b);
  });
  $("dBar").close();
});
$("bBarCancel").onclick = () => $("dBar").close();
$("bBarDel").onclick = () => {
  const { siteId, barId } = barCtx;
  if (!confirm("この工程を消します。よろしいですか？")) return;
  mutate(fileOfSite(siteId), doc => { doc.バー = doc.バー.filter(x => x.id !== barId); });
  $("dBar").close();
};
$("bBarCopy").onclick = () => {
  // 「何日から何日まで、少し空けてまた何日から」の2本目を楽に作る
  const v = readBarForm(), len = dayDiff(v.開始, v.終了);
  const site = S.files[fileOfSite(barCtx.siteId)].doc;
  $("dBar").close();
  openBar(site, null, { 段: v.段, 開始: dayAdd(v.終了, 1), 終了: dayAdd(v.終了, 1 + len) });
  $("bWork").value = v.作業; $("bStaff").value = v.社員.join("、"); $("bEmp").value = v.従業員.join("、"); $("bNum").value = v.人数; $("bMemo").value = v.メモ;
  drawPick("社員"); drawPick("従業員");
  (document.querySelector(`#bColors input[value="${v.色}"]`) || {}).checked = true;
  $("bFrom").focus();
};
$("bNum").addEventListener("input", e => { const t = toHalf(e.target.value); if (t !== e.target.value) e.target.value = t; autoNum = false; });

// 社員・従業員を名簿の一覧（フルネーム）から選ぶ。選んである人は ✓。もう一度選ぶと外す
let autoNum = true;
const PICK = { 社員: ["bStaff", "bStaffPick"], 従業員: ["bEmp", "bEmpPick"] };
function drawPick(field) {
  const [inp, box] = PICK[field], list = namesFor(field);
  const cur = splitNames($(inp).value).map(x => resolveName(x, list));
  $(box).innerHTML = `<option value="">＋ 選ぶ</option>` +
    list.map(n => `<option value="${esc(n)}">${cur.includes(n) ? "✓ " : "　 "}${esc(n)}</option>`).join("");
  $(box).value = "";
}
function afterPick(field) {
  if (field === "従業員" && autoNum) { const c = splitNames($("bEmp").value).length; $("bNum").value = c || ""; }
}
Object.entries(PICK).forEach(([field, [inp, box]]) => {
  $(box).addEventListener("change", e => {
    const n = e.target.value;
    if (!n) return;
    const list = namesFor(field);
    let cur = splitNames($(inp).value).map(x => resolveName(x, list));
    cur = cur.includes(n) ? cur.filter(x => x !== n) : [...cur, n];
    $(inp).value = cur.join("、");
    drawPick(field); afterPick(field);
  });
  $(inp).addEventListener("input", () => { drawPick(field); afterPick(field); });
});

/* ---------- 現場の小窓 ---------- */
let siteCtx = null;
const laneRuleText = ln => "作業：平日" + (ln.土 !== false ? "・土" : "") + (ln.日 === true ? "・日" : "") + (ln.祝 === true ? "・祝" : "") +
  ((ln.休み || []).length ? `／その日だけ休み ${ln.休み.length}日` : "") + ((ln.出勤 || []).length ? `／その日だけ出勤 ${ln.出勤.length}日` : "");
function laneRow(l) {
  const ck = (key, label, def) => `<label class="dw"><input type="checkbox" data-k="${key}" ${(l[key] ?? def) ? "checked" : ""}>${label}</label>`;
  return `<div class="ln" data-id="${esc(l.id)}"><input type="text" value="${esc(l.名前)}">` +
    ck("土", "土", true) + ck("日", "日", false) + ck("祝", "祝", false) +
    `<button type="button" class="sub small" data-mv="-1" title="上へ">▲</button><button type="button" class="sub small" data-mv="1" title="下へ">▼</button>` +
    `<button type="button" class="danger small" data-del>✕</button></div>`;
}
function openSite(id) {
  const s = id ? S.files[fileOfSite(id)].doc : null;
  siteCtx = { id };
  $("dSiteTtl").textContent = s ? "現場の設定" : "現場を追加";
  $("sName").value = s ? s.名前 : "";
  $("sCode").value = s ? s.コード || "" : "";
  $("sShort").value = s ? s.略称 || "" : "";
  $("sDone").checked = !!(s && s.完了);
  $("sTanto").value = s ? s.担当 || "" : "";
  const lanes = s ? s.段 : [{ id: newId("l"), 名前: "昼間" }];
  $("sLanes").innerHTML = lanes.map(laneRow).join("");
  $("bSiteDel").hidden = !s; $("bSiteUp").parentElement.hidden = !s;
  openDlg($("dSite"));
  if (!s) $("sName").focus();
}
$("sLanes").addEventListener("click", e => {
  const row = e.target.closest(".ln");
  if (!row) return;
  if (e.target.dataset.mv) {
    const sib = e.target.dataset.mv === "-1" ? row.previousElementSibling : row.nextElementSibling;
    if (sib) e.target.dataset.mv === "-1" ? row.parentNode.insertBefore(row, sib) : row.parentNode.insertBefore(sib, row);
  } else if ("del" in e.target.dataset) {
    const s = siteCtx.id && S.files[fileOfSite(siteCtx.id)].doc;
    if (s && s.バー.some(b => b.段 === row.dataset.id)) { alert("この段にはバーがあるので消せません。先にバーを消すか、別の段へ移してください。"); return; }
    if ($("sLanes").children.length <= 1) { alert("段は1つ以上要ります。"); return; }
    row.remove();
  }
});
$("sName").addEventListener("change", () => {
  const n = $("sName").value.trim(), x = M().現場.find(s => s.名前 === n) || M().現場.find(s => s.略称 === n);
  if (!x) return;
  if (!$("sCode").value) $("sCode").value = x.コード || "";
  if (!$("sTanto").value && x.担当) $("sTanto").value = resolveName(x.担当, namesFor("社員"));
  if (!$("sShort").value && x.略称 && x.略称 !== x.名前) $("sShort").value = x.略称;
});
$("bLaneAdd").onclick = () => { $("sLanes").insertAdjacentHTML("beforeend", laneRow({ id: newId("l"), 名前: "" })); $("sLanes").lastElementChild.querySelector("input").focus(); };
$("fSite").addEventListener("submit", e => {
  e.preventDefault();
  const name = $("sName").value.trim();
  if (!name) return;
  const lanes = [...$("sLanes").children].map(r => {
    const l = { id: r.dataset.id, 名前: r.querySelector("input[type=text]").value.trim() || "（名前なし）" };
    r.querySelectorAll("input[data-k]").forEach(c => { l[c.dataset.k] = c.checked; });
    return l;
  });
  // その日だけの休み・出勤は、保存するときの最新の中身から引き継ぐ
  const keepDays = (doc, l) => { const o = (doc.段 || []).find(x => x.id === l.id); return { ...l, 休み: o ? o.休み || [] : [], 出勤: o ? o.出勤 || [] : [] }; };
  const v = { 名前: name, 完了: $("sDone").checked, 略称: $("sShort").value.trim(), コード: toHalf($("sCode").value.trim()), 担当: resolveName($("sTanto").value.trim(), namesFor("社員")) };
  if (siteCtx.id) {
    mutate(fileOfSite(siteCtx.id), doc => { Object.assign(doc, v); doc.段 = lanes.map(l => keepDays(doc, l)); });
  } else {
    const id = newId("s"), file = fileOfSite(id);
    const maxOrd = Math.max(0, ...sites().map(s => s.並び ?? 0));
    S.files[file] = { doc: { id, ...v, 段: lanes, バー: [], 並び: maxOrd + 10 }, eTag: null, pending: [], saving: false, timer: null, err: null };
    mutate(file, () => { });
  }
  $("dSite").close();
});
$("bSiteCancel").onclick = () => $("dSite").close();
$("bSiteDel").onclick = () => {
  const s = S.files[fileOfSite(siteCtx.id)].doc;
  if (!confirm(`「${s.名前}」を工程表から消します（バー ${s.バー.length} 本も見えなくなります）。よろしいですか？`)) return;
  mutate(fileOfSite(siteCtx.id), doc => { doc.消した = true; });
  $("dSite").close();
};
function moveSite(dir) {
  const list = sites(), i = list.findIndex(s => s.id === siteCtx.id), j = i + dir;
  if (j < 0 || j >= list.length) return;
  // 並びを10おきに振り直してから入れ替える（2つのファイルだけ書く）
  const a = list[i], b = list[j];
  const oa = (j + 1) * 10, ob = (i + 1) * 10;
  list.forEach((s, k) => { if (s !== a && s !== b && (s.並び ?? 0) !== (k + 1) * 10) { const ord = (k + 1) * 10; mutate(fileOfSite(s.id), d => { d.並び = ord; }); } });
  mutate(fileOfSite(a.id), d => { d.並び = oa; });
  mutate(fileOfSite(b.id), d => { d.並び = ob; });
}
$("bSiteUp").onclick = () => moveSite(-1);
$("bSiteDown").onclick = () => moveSite(1);

/* ---------- 上の操作 ---------- */
function setFrom(ym) {
  S.view.from = ym; $("inFrom").value = ym;
  try { localStorage.setItem("kotei_view", JSON.stringify(S.view)); } catch (e) { }
  render();
}
const shiftMonth = (ym, n) => { const d = new Date(+ym.slice(0, 4), +ym.slice(5, 7) - 1 + n, 1); return d.getFullYear() + "-" + pad2(d.getMonth() + 1); };
$("inFrom").onchange = e => e.target.value && setFrom(e.target.value);
$("inMonths").onchange = e => { S.view.months = +e.target.value; setFrom(S.view.from); };
$("bPrev").onclick = () => setFrom(shiftMonth(S.view.from, -1));
$("bNext").onclick = () => setFrom(shiftMonth(S.view.from, 1));
$("bToday").onclick = () => {
  setFrom(todayKey().slice(0, 7));
  const i = dayDiff(range().from, todayKey());
  $("wrap").scrollLeft = Math.max(0, i * DW() - 60);
};
$("bReload").onclick = () => pull().then(() => toast("最新にしました"));
$("tabs").addEventListener("click", e => {
  const t = e.target.dataset.tab;
  if (!t) return;
  S.tab = t;
  try { localStorage.setItem("kotei_tab", t); } catch (e) { }
  document.querySelectorAll("#tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === t));
  $("pChart").hidden = t !== "chart"; $("pFlow").hidden = t !== "flow"; $("pList").hidden = t !== "list";
  render();
});
$("inShowDone").checked = showDone;
$("inShowDone").onchange = e => { showDone = e.target.checked; try { localStorage.setItem("kotei_showDone", showDone ? "1" : "0"); } catch (x) { } render(); };
$("inListRange").onchange = renderList;
$("inListQ").oninput = renderList;
$("flow").addEventListener("click", e => {
  const td = e.target.closest("td[data-bar]");
  if (!td) return;
  const s = S.files[fileOfSite(td.dataset.site)].doc;
  openBar(s, s.バー.find(b => b.id === td.dataset.bar));
});
$("list").addEventListener("click", e => {
  const tr = e.target.closest("tr.r");
  if (!tr) return;
  const s = S.files[fileOfSite(tr.dataset.site)].doc;
  openBar(s, s.バー.find(b => b.id === tr.dataset.bar));
});
window.addEventListener("beforeunload", e => {
  if (Object.values(S.files).some(f => f.pending.length || f.saving)) { e.preventDefault(); e.returnValue = ""; }
});
document.addEventListener("visibilitychange", () => { if (!document.hidden) pull(); });

let toastTimer = null;
function toast(msg, ms = 3000) {
  const el = $("toast"); el.textContent = msg; el.classList.add("on");
  clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove("on"), ms);
}

/* ---------- 試しのデータ（?demo のときだけ。社員名は入れない） ---------- */
function seedDemo() {
  if (Object.keys(DemoStore.all()).length) return;
  const ym = todayKey().slice(0, 7) + "-";
  const mk = (n, name, lanes, bars) => {
    const id = "demo" + n, ls = lanes.map((l, i) => ({ id: "l" + n + i, 名前: l }));
    return [fileOfSite(id), { id, 名前: name, コード: "", 担当: "", 並び: n * 10, 段: ls,
      バー: bars.map(([li, w, a, z, num, c], i) => ({ id: "b" + n + i, 段: ls[li].id, 作業: w, 開始: ym + pad2(a), 終了: ym + pad2(z), 社員: [], 人数: num, 色: c, メモ: "" })) }];
  };
  const a = {};
  [mk(1, "A現場 配水管", ["夜間", "25t夜間"], [[0, "布設工", 3, 10, 3, "青"], [0, "布設工", 15, 22, 3, "青"], [1, "クレーン", 5, 6, "", "橙"]]),
   mk(2, "B現場 道路改良", ["昼間"], [[0, "舗装", 8, 12, 5, "緑"]]),
   mk(3, "C現場 下水", ["昼間"], [[0, "推進工", 1, 25, 4, "紫"]])]
    .forEach(([f, d]) => { a[f] = { data: d, eTag: "d0" }; });
  a["マスタ.json"] = { eTag: "d0", data: { 社員: ["試し　太郎", "試し　次郎", "例題　花子"], 従業員: ["見本　一郎", "見本　二郎", "見本　三郎"],
    現場: [{ 名前: "E現場 新設工事", 略称: "E新設", コード: "520009", 担当: "試し" }] } };
  DemoStore.save(a);
}

/* ---------- 起動 ---------- */
// 版の表示（公開するとき tools\公開する.py が __VER__ を「Vr001 10/07 19:15」のように書き換える）
if ($("ver").textContent === "__VER__") $("ver").textContent = "開発中";
async function start() {
  try { const v = JSON.parse(localStorage.getItem("kotei_view")); if (v && v.from) Object.assign(S.view, v); } catch (e) { }
  if (!S.view.from) S.view.from = todayKey().slice(0, 7);
  $("inFrom").value = S.view.from; $("inMonths").value = S.view.months;

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
    if (!acc) {
      $("signin").hidden = false;
      $("bSignIn").onclick = () => msalApp.loginRedirect({ scopes: SCOPES });
      return;
    }
    msalApp.setActiveAccount(acc);
    me.name = acc.name || acc.username;
    $("who").textContent = me.name;
  }
  $("loading").hidden = false;
  ready = true;
  await pull();
  $("loading").hidden = true;
  showSync();
  let tab = "chart";
  try { tab = localStorage.getItem("kotei_tab") || "chart"; } catch (e) { }
  (document.querySelector(`#tabs [data-tab="${tab}"]`) || document.querySelector("#tabs [data-tab=chart]")).click();
  setInterval(() => { if (!document.hidden) pull(); }, POLL_MS);
}
start();
