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
  },
  // 画像（現場の地図）。上書きでよい（同時に直すことはまず無い）。読むときは data URL で返す
  async putRaw(name, blob) {
    const res = await g(this.base() + ":/" + encodeURIComponent(name) + ":/content", { method: "PUT", headers: { "Content-Type": blob.type || "application/octet-stream" }, body: blob });
    if (!res.ok) throw httpErr(res, "画像の保存");
  },
  async getRaw(name) {
    const res = await g(this.base() + ":/" + encodeURIComponent(name));
    if (res.status === 404) return null;
    if (!res.ok) throw httpErr(res, "画像の読み込み");
    const dl = (await res.json())["@microsoft.graph.downloadUrl"];
    const body = await fetch(dl, { cache: "no-store" });
    if (!body.ok) throw httpErr(body, "画像の読み込み");
    return blobToUrl(await body.blob());
  }
};
const blobToUrl = b => new Promise((ok, ng) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = ng; r.readAsDataURL(b); });

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
  },
  async putRaw(name, blob) { try { localStorage.setItem(this.key + ":" + name, await blobToUrl(blob)); } catch (e) { throw new Error("画像が大きすぎます（試しの画面）"); } },
  async getRaw(name) { try { return localStorage.getItem(this.key + ":" + name); } catch (e) { return null; } }
};
let store = DEMO ? DemoStore : GraphStore;

/* ---------- 状態 ---------- */
// S.files[name] = { doc, eTag, pending:[fn], saving, timer, err }
const S = { files: {}, view: { from: "", months: 2 }, tab: "chart" };
const newId = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const siteFiles = () => Object.keys(S.files).filter(n => n.startsWith("現場_") && S.files[n].doc && !S.files[n].doc.消した);
/* 現場名＝社内の工事名＝略称（2026-10-14 ユーザー指定：略称の欄は現場名にまとめる。正式な名前は「正式工事名」）。
   前のデータで 略称 と 名前 が違う現場は、略称を現場名、名前を正式工事名とみなす（現場の設定を保存し直すと そのように入る） */
const nm = s => s.略称 || s.名前;
const officialOf = s => s.工事名 || (s.略称 && s.略称 !== s.名前 ? s.名前 : "");
const sites = () => siteFiles().map(n => S.files[n].doc).sort((a, b) => (a.並び ?? 0) - (b.並び ?? 0) || String(nm(a)).localeCompare(nm(b), "ja"));
// 終わった現場（完了）は、上の「完了も出す」にチェックしたときだけ出す（その人のブラウザに覚える）
let showDone = false;
try { showDone = localStorage.getItem("kotei_showDone") === "1"; } catch (e) { }
/* 現場の並び（2026-10-13 ユーザー指定：現場名・工事コード・工程の最初の日付でも並べたい）。
   画面の並びだけ変える（その人のブラウザに覚える kotei_sort）。「設定順」は現場の設定の ▲▼ で決めた並び（全員共通） */
let sortBy = "manual";
try { sortBy = localStorage.getItem("kotei_sort") || "manual"; } catch (e) { }
const firstDay = s => s.バー.reduce((m, b) => !m || b.開始 < m ? b.開始 : m, "");
function sortSites(list) {
  const byName = (a, b) => String(nm(a)).localeCompare(String(nm(b)), "ja");
  const last = (x, y, f) => (!x && y) ? 1 : (x && !y) ? -1 : f();   // 空は後ろ
  if (sortBy === "name") return [...list].sort(byName);
  if (sortBy === "code") return [...list].sort((a, b) => { const x = String(a.コード || ""), y = String(b.コード || ""); return last(x, y, () => x.localeCompare(y, "ja", { numeric: true })) || byName(a, b); });
  if (sortBy === "start") return [...list].sort((a, b) => { const x = firstDay(a), y = firstDay(b); return last(x, y, () => x.localeCompare(y)) || byName(a, b); });
  return list;
}
const shownSites = () => sortSites(sites().filter(s => S.print ? S.print.sites.includes(s.id) : (showDone || !s.完了)));
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
/* 開始日は好きな日にできる（2026-10-13 ユーザー指定。前は月の1日だけ）。S.view.from＝"YYYY-MM-DD"。前の "YYYY-MM" は1日とみなす。
   終わりは 開始日から「nか月」先の前の日（10/15 から2か月なら 12/14） */
const fromKey = v => /^\d{4}-\d{2}$/.test(v || "") ? v + "-01" : v;
// n か月先の同じ日（無い日は月末：1/31 の1か月先は 2/28）
function addMonths(k, n) {
  const y = +k.slice(0, 4), m = +k.slice(5, 7) - 1 + n, d = +k.slice(8, 10);
  const last = new Date(y, m + 1, 0).getDate();
  return keyOf(new Date(y, m, Math.min(d, last)));
}
function range() {
  // 印刷の小窓で期間を決めて印刷している間は、その期間（S.print）
  const from = S.print ? S.print.from : fromKey(S.view.from);
  const end = S.print ? S.print.end : dayAdd(addMonths(from, S.view.months), -1);
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
   段.休み／出勤 … その日だけの例外（前は段の上を右クリックで切り替えた。2026-10-08 にやめたが、残っているものは効く） */
function ruleWorks(ln, k) {
  if (holidayOf(k)) return ln.祝 === true;
  const w = dayDow(k);
  if (w === 0) return ln.日 === true;
  if (w === 6) return ln.土 !== false;
  return true;
}
// 段.休み／出勤（前の版の右クリックで入れた「段のその日だけ休み」）は 2026-10-08 から見ない。
// 直す画面が無く、残っていると意味の分からない斜線になるため（ユーザー：10/21 の斜線は何？）。工程ごとの休みは バー.休み
/* 工種（段）の区分と作業名（2026-10-08 ユーザー：「夜間」と書いている所を 作業区分（昼／夜）と 作業名 に分けたい）
   段.区分＝"昼"／"夜"、段.名前＝作業名。区分の無い前のデータは、名前の「昼間／夜間」から区分を決め、その字を除いた残りを作業名とみなす */
function laneKubun(ln) { return !ln ? "昼" : ln.区分 || (/夜/.test(ln.名前 || "") ? "夜" : "昼"); }
function laneWork(ln) { return !ln ? "" : ln.区分 ? (ln.名前 || "") : String(ln.名前 || "").replace(/昼間|夜間/g, "").trim(); }
function laneName(ln) { const w = laneWork(ln), k = laneKubun(ln); return w ? k + "　" + w : k + "間"; }
function laneWorks(ln, k) {
  if (!ln || isRest(k)) return false;
  return ruleWorks(ln, k);
}
// 工程（バー）ごとの休工日（バーの小窓の「休工日を設定」）。段の決まりの上にかける
function barWorks(ln, b, k) {
  if (isRest(k)) return false;
  if (b && (b.休み || []).includes(k)) return false;
  if (b && (b.出勤 || []).includes(k)) return true;
  return laneWorks(ln, k);
}
/* 作業員の人数（2026-10-14 ユーザー：「〇名＋〇名」と書きたい（中身の名前は書かない）・日ごとに変えたい）
   バー.人数＝合計（前の版でも読める）、バー.人数内訳＝"1+3"（＋で分けたときだけ）、
   バー.人数日＝{ "YYYY-MM-DD": "2+4" }（その日だけの人数。無い日はふだんの人数） */
function parseHead(t) {
  const s = String(t ?? "").replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/＋/g, "+").replace(/[名人\s　]/g, "");
  if (!s) return null;
  const parts = s.split("+").filter(x => x !== "");
  if (!parts.length || parts.some(x => !/^\d+$/.test(x))) return null;
  const n = parts.map(Number);
  return { txt: n.join("+"), parts: n, total: n.reduce((a, x) => a + x, 0) };
}
// 工種（段）にも人数を持つ（段.人数＝"1+3"。2026-10-14 ユーザー：現場の設定の工種の行で〇人＋〇人）。バーに人数が無ければ工種の人数
const headOwn = b => b.人数内訳 || (b.人数 !== "" && b.人数 != null ? String(b.人数) : "");
const headTxt = (b, ln) => headOwn(b) || (ln && ln.人数) || "";
const headTxtAt = (b, k, ln) => (b.人数日 && b.人数日[k] != null && b.人数日[k] !== "") ? b.人数日[k] : headTxt(b, ln);
const headAt = (b, k, ln) => { const p = parseHead(headTxtAt(b, k, ln)); return p ? p.total : 0; };
const fmtHead = t => { const p = parseHead(t); return p ? p.parts.map(x => x + "名").join("＋") : ""; };
const dayHeads = b => Object.keys(b.人数日 || {}).sort();
/* 工種は 作業員（自社）か 外注（2026-10-14 ユーザー：作業名の横で 外注 or 作業員(自社)。外注は業者名＋人数、作業員は〇人＋〇人）
   段.種別＝"外注"（無ければ自社）、段.業者＝業者名、段.人数＝"3" */
const isGai = ln => !!ln && ln.種別 === "外注";
const gaiName = ln => isGai(ln) ? (ln.業者 || "外注") : "";
// バーに出す人数：ふだんの人数。日ごとに変えた日があれば ※。外注の工種は業者名を前に
const headLabel = (b, ln) => {
  let t = fmtHead(headTxt(b, ln));
  if (dayHeads(b).length) t = (t || "0名") + "※";
  return isGai(ln) ? [gaiName(ln), t].filter(Boolean).join(" ") : t;
};

function dayClass(k) {
  if (isRest(k)) return "rest";
  if (k === todayKey()) return "today";
  if (holidayOf(k)) return "hol";
  const w = dayDow(k);
  return w === 0 ? "sun" : w === 6 ? "sat" : "";
}

/* ---------- 工程表を描く ---------- */
const DW = () => parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--dw"));

/* バーの下の文字（2026-10-08 ユーザー：文字が多いと見づらい → 空白を省く、収まらなければ名字だけ、重ならないように）
   avail＝同じ行の次のバーまでの幅。長い順に試して、収まる最初の書き方にする */
const measureCtx = document.createElement("canvas").getContext("2d");
function textW(t) { measureCtx.font = '12px "Yu Gothic UI", Meiryo, sans-serif'; return measureCtx.measureText(t).width; }
const noSpace = n => String(n).replace(/[\s　]+/g, "");
// 名字だけ。社員に同じ名字が2人以上いれば空白を省いた氏名
function shortStaff(n) { const sn = surname(n), list = namesFor("社員"); return list.filter(x => surname(x) === sn).length > 1 ? noSpace(n) : sn; }
function barLabel(b, avail, ln) {
  const w = b.作業 || "", num = headLabel(b, ln), st = b.社員 || [];
  const cands = [
    [w, num, st.map(noSpace).join("・")],
    [w, num, st.map(shortStaff).join("・")],
    [w, num, ""], [w, "", ""]
  ];
  let pick = cands[cands.length - 1];
  for (const c of cands) { if (textW(c.filter(Boolean).join(" ")) <= avail) { pick = c; break; } }
  return [esc(pick[0]), esc(pick[1]), pick[2] ? `<span class="st">${esc(pick[2])}</span>` : ""].filter(Boolean).join(" ");
}
function barTitle(s, b) {
  const ln = s.段.find(l => l.id === b.段);
  let work = 0;
  for (let k = b.開始; k <= b.終了; k = dayAdd(k, 1)) if (barWorks(ln, b, k)) work++;
  return `${nm(s)}／${b.作業 || ""}\n${md(b.開始)}〜${md(b.終了)}（${dayDiff(b.開始, b.終了) + 1}日・うち作業 ${work}日）` +
    (b.社員 && b.社員.length ? "\n社員：" + b.社員.join("・") : "") + (isGai(ln) ? "\n外注：" + gaiName(ln) + (headTxt(b, ln) ? " " + fmtHead(headTxt(b, ln)) + (headOwn(b) ? "" : "（工種の人数）") : "") :
      headTxt(b, ln) ? "\n作業員：" + fmtHead(headTxt(b, ln)) + (headOwn(b) ? "" : "（工種の人数）") : "") +
    (dayHeads(b).length ? "\n日ごとの人数：" + dayHeads(b).map(k => md(k) + " " + (fmtHead(b.人数日[k]) || "0名")).join("、") : "") + (b.メモ ? "\n" + b.メモ : "");
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
    days.map((k, i) => { const m = k.endsWith("-01") || (i === 0 && +k.slice(8) < 25); return `<div class="hd ${m ? "mon" : ""}">${m ? (+k.slice(5, 7)) + "月" : ""}</div>`; }).join("") + "</div>");
  h.push(`<div class="hrow d"><div class="corner"></div>` +
    days.map(k => `<div class="hd ${dayClass(k)}" data-day="${k}" title="${esc((holidayOf(k) ? holidayOf(k) + "\n" : "") + "押すと会社の休みにする／戻す")}">${+k.slice(8)}</div>`).join("") + "</div>");
  h.push(`<div class="hrow w"><div class="corner">現場／工種</div>` +
    days.map(k => `<div class="hd ${dayClass(k)}" data-day="${k}" title="押すと会社の休みにする／戻す">${isRest(k) ? "休" : WD[dayDow(k)]}</div>`).join("") + "</div>");
  // 本体
  h.push(`<div id="body">`);
  const left = `calc((var(--sw) + var(--lw)) * var(--lz, 1))`;
  days.forEach((k, i) => { const c = dayClass(k); if (c) h.push(`<div class="col ${c}" style="left:calc(${left} + ${i * dw}px)"></div>`); });
  for (const s of shownSites()) {
    const fold = folded.has(s.id);
    h.push(`<div class="site ${fold ? "folded" : ""} ${s.完了 ? "done" : ""}" data-site="${esc(s.id)}"><div class="sname" data-act="site"><span class="tg" title="${fold ? "ひらく" : "たたむ"}">${fold ? "▸" : "▾"}</span>${s.完了 ? '<span class="donemark">完了</span>' : ""}${esc(nm(s))}` +
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
    } else for (const ln of visibleLanes(s)) {
      if (ln === DONE_ROW) {
        // 完了した工種をたたんだ行（押すと出す／たたむ。その人のブラウザに覚える）
        const n = s.段.filter(l => l.完了).length, open = showLanes.has(s.id);
        h.push(`<div class="lane donerow"><div class="llab" data-act="donelanes" title="${open ? "完了した工種をたたむ" : "完了した工種を出す"}">${open ? "▾" : "▸"} 完了 ${n}工種</div><div class="track sum" style="width:${width}px" data-act="donelanes"></div></div>`);
        continue;
      }
      const bars = s.バー.filter(b => b.段 === ln.id && b.終了 >= from && b.開始 <= end);
      if (laneFold.has(ln.id)) {
        // たたんだ工種：細い1行に色の線だけ（押すとひらく）
        h.push(`<div class="lane lfold" data-lane="${esc(ln.id)}"><div class="llab ${isNight(ln) ? "night" : ""} ${ln.完了 ? "ldone" : ""}" data-act="lanefold" title="押すとひらく">▸ ${esc(laneName(ln))}</div><div class="track sum" data-act="lanefold" style="width:${width}px">` +
          bars.map(b => { const a = b.開始 < from ? from : b.開始, z = b.終了 > end ? end : b.終了; return `<div class="lfbar" style="left:${dayDiff(from, a) * dw}px;width:${(dayDiff(a, z) + 1) * dw}px;background:${colorOf(b.色)}" title="${esc(barTitle(s, b))}"></div>`; }).join("") + `</div></div>`);
        continue;
      }
      const { pos, rows } = packLane(bars);
      h.push(`<div class="lane" data-lane="${esc(ln.id)}" style="height:${rows * rh}px"><div class="llab ${isNight(ln) ? "night" : ""} ${ln.完了 ? "ldone" : ""}" data-act="site" title="${esc(laneName(ln) + "\n" + laneRuleText(ln) + (ln.完了 ? "\n完了した工種" : ""))}"><span class="ltg" data-act="lanefold" title="この工種をたたむ">▾</span>${esc(laneName(ln))}</div><div class="track" data-site="${esc(s.id)}" data-lane="${esc(ln.id)}" style="width:${width}px">`);
      // 工種の休みの日（土曜休みなど）の斜線はやめた（2026-10-13 ユーザー：斜線は日付を押した会社の休みの列だけ）。バーの中の白抜きは残す
      // 文字に使える幅：同じ行で次に始まるバーの手前まで（無ければ表の右端まで）
      const avail = {};
      for (const b of bars) {
        const a = b.開始 < from ? from : b.開始;
        const next = bars.filter(o => o !== b && pos[o.id] === pos[b.id] && o.開始 > b.開始).map(o => o.開始).sort()[0];
        avail[b.id] = Math.max(dw, (next ? dayDiff(a, next) : dayDiff(a, end) + 1) * dw - 6);
      }
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
        h.push(`<div class="bar ${isLight(b.色) ? "dark" : ""} ${b.id === selBar ? "sel" : ""} ${b.開始 < from ? "cut-l" : ""} ${b.終了 > end ? "cut-r" : ""}" data-bar="${esc(b.id)}" ` +
          `style="left:${x}px;top:${pos[b.id] * rh + Math.max(3, (rh - 26) / 2)}px;width:${w}px;--bc:${colorOf(b.色)}" title="${esc(barTitle(s, b))}">` +
          `<span class="stripe"></span>` +
          gaps.map(g => `<span class="gap" style="left:${g.i * dw}px;width:${g.n * dw}px"></span>`).join("") +
          `<span class="h l"></span><span class="lbl" style="max-width:${avail[b.id]}px">${barLabel(b, avail[b.id], ln)}</span><span class="h r"></span></div>`);
      }
      h.push(`</div></div>`);
    }
    h.push(`</div></div>`);
  }
  h.push(`</div><div class="addsite"><button id="bAddSite" class="sub">＋ 現場を追加</button></div>`);
  const hidden = sites().length - shownSites().length;
  if (hidden) h.push(`<div class="addsite note">完了した現場 ${hidden} 件を隠しています（上の「完了も出す」で表示）</div>`);
  $("chart").innerHTML = h.join("");
  fitLaneWidth();
  $("bAddSite").onclick = () => openSite(null);
}

/* 完了した工種（2026-10-13 ユーザー：工種が増えすぎないよう、終わった工種はたたんで見えなくしたい）
   段.完了＝true（全員で共有）。工程表では隠して、現場の一番下に「▸ 完了 n工種」の細い行を出す。押すとその現場だけ出す（kotei_showLanes）。
   人の流れ・一覧は完了の工種も数える（事実なので） */
const DONE_ROW = {};
let showLanes = new Set();
try { showLanes = new Set(JSON.parse(localStorage.getItem("kotei_showLanes")) || []); } catch (e) { }
function visibleLanes(s) {
  const done = s.段.filter(l => l.完了);
  if (!done.length) return s.段;
  return [...(showLanes.has(s.id) ? s.段 : s.段.filter(l => !l.完了)), DONE_ROW];
}
/* 工種を1つずつたたむ／ひらく（2026-10-13 ユーザー指定）。その人のブラウザだけに覚える（kotei_laneFold）。工種名の左の ▾／▸ */
let laneFold = new Set();
try { laneFold = new Set(JSON.parse(localStorage.getItem("kotei_laneFold")) || []); } catch (e) { }
function toggleLaneFold(id) {
  if (!id) return;
  laneFold.has(id) ? laneFold.delete(id) : laneFold.add(id);
  try { localStorage.setItem("kotei_laneFold", JSON.stringify([...laneFold])); } catch (e) { }
  render();
}
function toggleShowLanes(id) {
  showLanes.has(id) ? showLanes.delete(id) : showLanes.add(id);
  try { localStorage.setItem("kotei_showLanes", JSON.stringify([...showLanes])); } catch (e) { }
  render();
}
function setLaneDone(siteId, laneId, on) {
  mutate(fileOfSite(siteId), doc => { const l = doc.段.find(x => x.id === laneId); if (l) { if (on) l.完了 = true; else delete l.完了; } });
}
// 終わった工種：バーがあって、どのバーも今日より前に終わっている
const laneFinished = (s, l) => { const bs = s.バー.filter(b => b.段 === l.id); return bs.length > 0 && bs.every(b => b.終了 < todayKey()); };

/* 工種名の列：いちばん長い工種名が収まる幅にする（2026-10-13 ユーザー指定）。狭すぎ・広すぎは 56〜280px で止める */
function fitLaneWidth() {
  let w = 0;
  document.querySelectorAll("#chart .llab").forEach(el => {
    measureCtx.font = (el.classList.contains("night") ? "bold " : "") + '13px "Yu Gothic UI", Meiryo, sans-serif';
    w = Math.max(w, measureCtx.measureText(el.textContent).width);
  });
  $("chart").style.setProperty("--lw", Math.min(280, Math.max(56, Math.ceil(w) + 12)) + "px");
}

/* ---------- 一覧 ---------- */
function renderList() {
  const { from, end } = range();
  const only = $("inListRange").checked, q = $("inListQ").value.trim();
  const rows = [];
  for (const s of shownSites()) for (const b of s.バー) {
    if (only && (b.終了 < from || b.開始 > end)) continue;
    const ln = s.段.find(l => l.id === b.段);
    const text = [nm(s), officialOf(s), ln && laneName(ln), b.作業, (b.社員 || []).join(" "), (b.従業員 || []).join(" "), b.メモ].join(" ");
    if (q && !q.split(/\s+/).every(w => text.includes(w))) continue;
    rows.push({ s, b, ln });
  }
  rows.sort((x, y) => x.b.開始.localeCompare(y.b.開始) || (x.s.並び ?? 0) - (y.s.並び ?? 0));
  const tot = rows.reduce((a, r) => a + (dayDiff(r.b.開始, r.b.終了) + 1), 0);
  $("list").innerHTML = `<table><tr><th>現場</th><th>段</th><th>作業</th><th>開始</th><th>終了</th><th>日数</th><th>社員</th>${SHOW_EMP ? "<th>従業員</th>" : ""}<th>作業員</th><th>メモ</th><th>更新</th></tr>` +
    rows.map(({ s, b, ln }) => `<tr class="r" data-site="${esc(s.id)}" data-bar="${esc(b.id)}"><td>${esc(nm(s))}</td><td>${esc(ln ? laneName(ln) : "")}</td>` +
      `<td><span class="chip" style="background:${colorOf(b.色)}"></span>${esc(b.作業)}</td><td>${md(b.開始)}（${WD[dayDow(b.開始)]}）</td><td>${md(b.終了)}（${WD[dayDow(b.終了)]}）</td>` +
      `<td class="n">${dayDiff(b.開始, b.終了) + 1}</td><td>${esc((b.社員 || []).join("・"))}</td>${SHOW_EMP ? `<td>${esc((b.従業員 || []).join("・"))}</td>` : ""}<td class="n">${esc(headLabel(b, ln))}</td>` +
      `<td>${esc(b.メモ)}</td><td class="note">${esc(b.更新 ? b.更新.だれ : "")}</td></tr>`).join("") +
    `</table><p class="note">${rows.length} 件（のべ ${tot} 日）</p>`;
}

/* ---------- 人の流れ（工程表のバーから、社員の行き先と作業員の日ごとの人数を出す） ---------- */
const shortOf = nm;
const isNight = ln => laneKubun(ln) === "夜";

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
      else if (siteIds.length > 1) h.push(`<td class="dup" data-site="${esc(hits[0].s.id)}" data-bar="${esc(hits[0].b.id)}" title="${esc(md(k) + " 重複：" + hits.map(x => nm(x.s) + "／" + (x.b.作業 || "")).join("、"))}">${siteIds.length}現場</td>`);
      else { const x = hits[0];
        h.push(`<td class="as ${isLight(x.b.色) ? "dark" : ""}" style="background:${colorOf(x.b.色)}" data-site="${esc(x.s.id)}" data-bar="${esc(x.b.id)}" title="${esc(md(k) + " " + nm(x.s) + "／" + (x.b.作業 || "") + (x.ln ? "（" + laneName(x.ln) + "）" : ""))}">${esc(shortOf(x.s))}</td>`); }
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
    h.push(un.length ? `<td class="un ${dayClass(k)}" data-site="${esc(un[0].s.id)}" data-bar="${esc(un[0].b.id)}" title="${esc(md(k) + " 社員未定：" + un.map(x => nm(x.s) + "／" + (x.b.作業 || "")).join("、"))}">${ss.length > 1 ? ss.length + "件" : esc(ss[0])}</td>` : `<td class="${dayClass(k)}"></td>`);
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
  // 自社の作業員と外注は分けて数える（外注は業者ごと）
  h.push(`<tr class="sec"><th colspan="${cols}">作業員（自社）の人数（工程表のバーの「作業員」の合計。区分が夜なら夜間）</th></tr>`);
  const sum = (k, f) => (at[k] || []).filter(f).reduce((a, x) => a + headAt(x.b, k, x.ln), 0);
  const own = x => !isGai(x.ln);
  const hasHead = (s, f) => s.バー.some(b => { const ln = s.段.find(l => l.id === b.段); return f(ln) && (headTxt(b, ln) || dayHeads(b).length) && b.終了 >= from && b.開始 <= end; });
  const numRow = (cls, label, f) => h.push(`<tr${cls ? ` class="${cls}"` : ""}><th class="nm">${esc(label)}</th>` + days.map(k => { const v = sum(k, f); return `<td class="${v ? "" : "n0"} ${dayClass(k)}">${v || ""}</td>`; }).join("") + `</tr>`);
  for (const s of shownSites()) if (hasHead(s, ln => !isGai(ln))) numRow("", nm(s), x => own(x) && x.s === s);
  numRow("tot", "昼間 計", x => own(x) && !isNight(x.ln));
  numRow("tot", "夜間 計", x => own(x) && isNight(x.ln));
  numRow("tot", "作業員 計", own);
  const gyo = [...new Set(Object.values(at).flat().filter(x => isGai(x.ln)).map(x => gaiName(x.ln)))].sort((a, b) => a.localeCompare(b, "ja"));
  if (gyo.length) {
    h.push(`<tr class="sec"><th colspan="${cols}">外注の人数（業者ごと）</th></tr>`);
    for (const g of gyo) numRow("", g, x => isGai(x.ln) && gaiName(x.ln) === g);
    numRow("tot", "外注 計", x => isGai(x.ln));
  }
  h.push(`</tbody></table>`);
  $("flow").innerHTML = h.join("");
}

function render() {
  if (S.tab === "chart") renderChart(); else if (S.tab === "flow") renderFlow(); else renderList();
  // 候補（マスタ＋入力済み）
  $("staffList").innerHTML = namesFor("社員").map(n => `<option value="${esc(n)}">`).join("");
  $("siteList").innerHTML = M().現場.map(x => `<option value="${esc(x.略称 || x.名前)}">${esc([x.コード, x.略称 && x.略称 !== x.名前 ? x.名前 : "", x.担当 ? "担当 " + x.担当 : ""].filter(Boolean).join("　"))}</option>`).join("");
}

/* ---------- ドラッグ（動かす・延ばす・作る） ----------
   マウス：つかんで動かす／端で延ばす／空いている所をなぞって作る／押すと直す。
   タブレット（2026-10-13 ユーザー指定：スワイプしたのにタップと取られて編集画面が開いた）：
     1本指タップ＝バーを選ぶ／ダブルタップ＝直す（空いている所なら新しいバー）／
     選んだバーは1本指で動かす・端で延ばす／空いている所を1本指でなぞると新しいバー／長押し＝メニュー／
     2本指スワイプ＝スクロール／2本指ピンチ＝拡大縮小（下の「拡大縮小」） */
let drag = null;
let selBar = null;        // タブレットで選んでいるバーの id（描き直しても残す）
// 1日の幅（画面の px）。工程表だけを拡大縮小しているので、その倍率をかける
const pxDay = () => DW() * ZOOM;
function dayAt(track, clientX) {
  const r = track.getBoundingClientRect();
  return Math.floor((clientX - r.left) / pxDay());
}
function setSel(id) {
  selBar = id;
  document.querySelectorAll("#chart .bar.sel").forEach(el => el.classList.remove("sel"));
  if (id) { const el = document.querySelector(`#chart .bar[data-bar="${CSS.escape(id)}"]`); if (el) el.classList.add("sel"); }
}
function startBarDrag(e, barEl, track, mode, x0) {
  const site = S.files[fileOfSite(track.dataset.site)].doc;
  const b = site.バー.find(x => x.id === barEl.dataset.bar);
  if (!b) return;
  drag = { kind: "bar", mode, el: barEl, track, site, b, x0, moved: false, a: b.開始, z: b.終了, from: range().from };
}
function startNewDrag(track, d0) {
  const ghost = document.createElement("div");
  ghost.className = "bar ghost";
  track.appendChild(ghost);
  drag = { kind: "new", track, site: S.files[fileOfSite(track.dataset.site)].doc, lane: track.dataset.lane, d0, d1: d0, el: ghost, from: range().from };
  placeGhost();
}
function cancelDrag() {
  if (!drag) return;
  const d = drag; drag = null;
  if (d.kind === "new") d.el.remove(); else if (d.moved) render();
}
$("wrap").addEventListener("pointerdown", e => {
  if (e.pointerType !== "mouse") return touchDown(e);
  if (e.button !== 0) return;
  const barEl = e.target.closest(".bar"), track = e.target.closest(".track");
  if (!track || track.classList.contains("sum")) return;
  if (barEl) {
    const mode = e.target.classList.contains("l") ? "l" : e.target.classList.contains("r") ? "r" : "move";
    startBarDrag(e, barEl, track, mode, e.clientX);
    barEl.setPointerCapture(e.pointerId);
  } else {
    startNewDrag(track, dayAt(track, e.clientX));
    track.setPointerCapture(e.pointerId);
  }
  e.preventDefault();
});
function placeGhost() {
  const dw = DW(), a = Math.min(drag.d0, drag.d1), z = Math.max(drag.d0, drag.d1);
  drag.el.style.left = a * dw + "px"; drag.el.style.width = (z - a + 1) * dw + "px";
  drag.el.textContent = `${md(dayAdd(drag.from, a))}〜${md(dayAdd(drag.from, z))}（${z - a + 1}日）`;
}
function dragMove(e) {
  if (drag.kind === "new") { drag.d1 = dayAt(drag.track, e.clientX); if (drag.d1 !== drag.d0) drag.moved = true; placeGhost(); return; }
  const dd = Math.round((e.clientX - drag.x0) / pxDay());
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
}
function dragEnd() {
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
}
window.addEventListener("pointermove", e => {
  if (e.pointerType !== "mouse") return touchMove(e);
  if (drag) dragMove(e);
});
window.addEventListener("pointerup", e => {
  if (e.pointerType !== "mouse") return touchUp(e);
  if (drag) dragEnd();
});
window.addEventListener("pointercancel", e => { if (e.pointerType !== "mouse") { if (tp) clearTimeout(tp.timer); tp = null; clearTouchMarks(); cancelDrag(); } });

/* タブレット（指・ペン）の1本指（2026-10-14 ユーザー指定で作り替え。2本指は下の touchstart が引き取る）
   - 1本指で動かす＝画面のスクロール（どこからでも）。2本指＝拡大縮小
   - 長押し（LONG_MS）して、そのまま指を動かす：
       空いている所＝工程線を書く（離すと作る。小窓は出さない）
       工程線の上＝影を付けて「移動」／工程線の端＝矢印を付けて「延長・短縮」
   - 長押しして動かさずに離す＝メニュー（右クリックと同じ：休工日・完了・消す など）
   - ダブルタップ：工程線＝工程の小窓。左の工事名・工種名＝現場の設定（click の方で受ける） */
const LONG_MS = 500;
let tp = null, lastTap = null, lastTouchAt = 0;
const fingers = new Set();
function touchDown(e) {
  fingers.add(e.pointerId);
  lastTouchAt = Date.now();
  if (fingers.size > 1) { if (tp) clearTimeout(tp.timer); tp = null; cancelDrag(); return; }
  const track = e.target.closest(".track"), barEl = e.target.closest(".bar");
  tp = { id: e.pointerId, x0: e.clientX, y0: e.clientY, x: e.clientX, y: e.clientY, target: e.target, track: track && !track.classList.contains("sum") ? track : null,
    barEl: barEl && !barEl.classList.contains("sumbar") ? barEl : null, moved: false, long: false, pan: false, sl: $("wrap").scrollLeft, st: $("wrap").scrollTop };
  tp.timer = setTimeout(touchLong, LONG_MS);
}
function touchLong() {
  if (!tp || tp.moved || !tp.track) return;
  tp.long = true;
  if (navigator.vibrate) try { navigator.vibrate(25); } catch (x) { }
  if (tp.barEl) {
    // 端（指なので広め）なら延ばす・縮める、ほかは移動。状態が分かるよう 影／矢印 を付ける
    const r = tp.barEl.getBoundingClientRect(), edge = Math.min(24, r.width / 3);
    const mode = tp.x0 - r.left < edge ? "l" : r.right - tp.x0 < edge ? "r" : "move";
    startBarDrag(null, tp.barEl, tp.track, mode, tp.x0);
    if (!drag) return;
    tp.barEl.classList.add(mode === "move" ? "lift" : "edge-" + mode);
    toast(mode === "move" ? "移動します：指を左右に動かしてください" : "長さを変えます：指を左右に動かしてください", 1800);
  } else {
    startNewDrag(tp.track, dayAt(tp.track, tp.x0));
    drag.el.classList.add("drawing");
    toast("工程線を書きます：指を左右に動かして線を引き、離すと作ります", 2200);
  }
}
function touchMove(e) {
  if (!tp || e.pointerId !== tp.id) return;
  tp.x = e.clientX; tp.y = e.clientY;
  if (!tp.moved && Math.hypot(e.clientX - tp.x0, e.clientY - tp.y0) > 10) {
    tp.moved = true;
    if (!tp.long) { clearTimeout(tp.timer); tp.pan = true; }   // 長押しの前に動いた＝スクロール
  }
  if (drag) { dragMove(e); return; }
  if (tp.pan) { $("wrap").scrollLeft = tp.sl - (e.clientX - tp.x0); $("wrap").scrollTop = tp.st - (e.clientY - tp.y0); }
}
function touchUp(e) {
  fingers.delete(e.pointerId);
  if (!tp || e.pointerId !== tp.id) return;
  const t = tp; tp = null; clearTimeout(t.timer);
  if (drag) {
    // 長押しして動かさずに離した＝メニュー
    if (t.long && !t.moved) { clearTouchMarks(); cancelDrag(); openCtx(t.track, t.barEl, t.x, t.y); return; }
    touchDragEnd(); return;
  }
  if (t.moved || t.long || !t.track || !t.barEl) return;
  // 工程線のダブルタップ＝工程の小窓
  const key = t.barEl.dataset.bar, now = Date.now();
  if (lastTap && lastTap.key === key && now - lastTap.at < 400) {
    lastTap = null;
    const site = S.files[fileOfSite(t.track.dataset.site)].doc;
    openBar(site, site.バー.find(x => x.id === key));
    return;
  }
  lastTap = { key, at: now };
}
function clearTouchMarks() { document.querySelectorAll("#chart .bar.lift, #chart .bar.edge-l, #chart .bar.edge-r").forEach(el => el.classList.remove("lift", "edge-l", "edge-r")); }
// 指で引いた工程線は、小窓を出さずにそのまま作る（作業名は工種の作業名。ダブルタップで直す）
function touchDragEnd() {
  const d = drag; drag = null;
  clearTouchMarks();
  if (d.kind === "new") {
    d.el.remove();
    const a = dayAdd(d.from, Math.min(d.d0, d.d1)), z = dayAdd(d.from, Math.max(d.d0, d.d1));
    const ln = d.site.段.find(l => l.id === d.lane);
    const b = { id: newId("b"), 段: d.lane, 作業: laneWork(ln), 開始: a, 終了: z, 社員: [], 人数: "", 色: lastColor, メモ: "" };
    stamp(b);
    mutate(fileOfSite(d.site.id), doc => { doc.バー.push({ ...b }); });
    toast(`${md(a)}〜${md(z)}（${dayDiff(a, z) + 1}日）の工程線を作りました。中身はダブルタップで入れます`, 2600);
    return;
  }
  if (!d.moved || (d.a === d.b.開始 && d.z === d.b.終了)) { render(); return; }
  const id = d.b.id, a = d.a, z = d.z;
  mutate(fileOfSite(d.site.id), doc => { const b = doc.バー.find(x => x.id === id); if (b) { b.開始 = a; b.終了 = z; stamp(b); } });
}
/* 右クリックのメニュー（2026-10-08 ユーザー：右クリックの「その日だけ休み」は要らない。
   同じ日に工種を足すと「工程を直す」が開いてしまう → 右クリックで足す・消すを選べるように）
   これまでに右クリックで入れた段の休み（段.休み／出勤）は残っていれば効く。新しくは作らない */
$("wrap").addEventListener("contextmenu", e => {
  // 指の長押しはこちらで出す（touchDown）。ブラウザの長押しメニューは止める
  if (Date.now() - lastTouchAt < 1500) { e.preventDefault(); return; }
  const track = e.target.closest(".track");
  if (!track || track.classList.contains("sum")) return;
  e.preventDefault();
  openCtx(track, e.target.closest(".bar"), e.clientX, e.clientY);
});
function openCtx(track, barEl, cx, cy) {
  const site = S.files[fileOfSite(track.dataset.site)].doc;
  const day = dayAdd(range().from, dayAt(track, cx));
  const b = barEl && site.バー.find(x => x.id === barEl.dataset.bar);
  const lane = track.dataset.lane;
  const ln = site.段.find(l => l.id === lane);
  const laneDoneItem = () => ln && ln.完了
    ? [`「${laneName(ln)}」の完了を取り消す`, () => setLaneDone(site.id, lane, false)]
    : [`「${ln ? laneName(ln) : "工種"}」を完了にしてたたむ`, () => setLaneDone(site.id, lane, true)];
  const items = b ? [
    [`「${b.作業 || "工程"}」を直す`, () => openBar(site, b)],
    ["同じ期間で工種を足す", () => openBar(site, null, { 段: b.段, 開始: b.開始, 終了: b.終了 })],
    [`${md(day)} から工種を足す`, () => openBar(site, null, { 段: lane, 開始: day, 終了: day })],
    ["休工日を設定", () => { openBar(site, b); $("bOffBtn").click(); }],
    laneDoneItem(),
    ["-"],
    [`「${b.作業 || "工程"}」を消す`, () => {
      if (!confirm(`「${b.作業 || "工程"}」（${md(b.開始)}〜${md(b.終了)}）を消します。よろしいですか？`)) return;
      const id = b.id;
      mutate(fileOfSite(site.id), doc => { doc.バー = doc.バー.filter(x => x.id !== id); });
    }, "danger"],
  ] : [
    [`${md(day)} から工種を足す`, () => openBar(site, null, { 段: lane, 開始: day, 終了: day })],
    laneDoneItem(),
    ["現場の設定を開く", () => openSite(site.id)],
  ];
  showCtx(cx, cy, items);
}
function showCtx(x, y, items) {
  const m = $("ctx");
  m.innerHTML = items.map((it, i) => it[0] === "-" ? `<hr>` : `<button type="button" data-i="${i}" class="${it[2] || ""}">${esc(it[0])}</button>`).join("");
  m.hidden = false;
  const r = m.getBoundingClientRect();
  const vw = document.documentElement.clientWidth || innerWidth || 1e5, vh = document.documentElement.clientHeight || innerHeight || 1e5;
  m.style.left = Math.max(0, Math.min(x, vw - r.width - 4)) + "px";
  m.style.top = Math.max(0, Math.min(y, vh - r.height - 4)) + "px";
  m.onclick = ev => { const i = ev.target.dataset.i; if (i === undefined) return; hideCtx(); items[+i][1](); };
}
function hideCtx() { $("ctx").hidden = true; }
document.addEventListener("pointerdown", e => { if (!e.target.closest("#ctx")) hideCtx(); }, true);
document.addEventListener("keydown", e => { if (e.key === "Escape") hideCtx(); });
$("wrap").addEventListener("scroll", hideCtx);
$("wrap").addEventListener("click", e => {
  const fb = e.target.closest("[data-fold]");
  if (fb) { setFold(shownSites().map(s => s.id), fb.dataset.fold === "all"); return; }
  const hd = e.target.closest(".hd[data-day]");
  if (hd) { toggleRest(hd.dataset.day); return; }
  if (e.target.closest("[data-act=lanefold]")) { toggleLaneFold(e.target.closest(".lane").dataset.lane); return; }
  if (e.target.closest("[data-act=donelanes]")) { toggleShowLanes(e.target.closest(".site").dataset.site); return; }
  if (e.target.closest(".tg") || e.target.closest("[data-act=fold]")) {
    const id = e.target.closest(".site").dataset.site;
    setFold([id], !folded.has(id)); return;
  }
  if (e.target.closest("[data-act=site]")) {
    const id = e.target.closest(".site").dataset.site;
    // 指で触ったときは ダブルタップで開く（1本指のスクロールで うっかり開かないように。2026-10-14 ユーザー指定）
    if (Date.now() - lastTouchAt < 800) {
      const now = Date.now();
      if (lastSiteTap && lastSiteTap.id === id && now - lastSiteTap.at < 450) { lastSiteTap = null; openSite(id); }
      else lastSiteTap = { id, at: now };
      return;
    }
    openSite(id);
  }
});
let lastSiteTap = null;

/* ---------- バーの小窓 ---------- */
let barCtx = null;
$("bColors").innerHTML = COLORS.map(([n, c], i) => `<input type="radio" name="bColor" id="bc${i}" value="${n}"><label for="bc${i}" style="background:${c}" title="${n}"></label>`).join("");
// 名前は「名字　名前」と全角空白を含むので、空白では区切らない
const splitNames = s => s.split(/[、,，・\n]+/).map(x => x.trim()).filter(Boolean);
const toHalf = s => s.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));

function openBar(site, b, init) {
  barCtx = { siteId: site.id, barId: b ? b.id : null, 休み: [...((b && b.休み) || [])], 出勤: [...((b && b.出勤) || [])], 人数日: { ...((b && b.人数日) || {}) } };
  const v = b || { 段: init.段, 開始: init.開始, 終了: init.終了, 作業: "", 社員: [], 人数: "", 色: lastColor, メモ: "" };
  $("dBarTtl").textContent = b ? "工程を直す" : "工程を足す";
  $("dBarSite").textContent = nm(site);
  $("bLane").innerHTML = site.段.map(l => `<option value="${esc(l.id)}">${esc(laneName(l))}${l.完了 ? "（完了）" : ""}</option>`).join("");
  $("bLane").value = v.段;
  $("bWork").value = v.作業 || "";
  $("bFrom").value = v.開始; $("bTo").value = v.終了;
  setStaffRows(joinSplit(v.社員 || [], M().社員 || []));
  $("bEmp").value = joinSplit(v.従業員 || [], M().従業員 || []).join("、");
  $("bNum").value = headOwn(v);
  autoNum = SHOW_EMP && (!v.人数 || v.人数 === (v.従業員 || []).length);
  document.querySelectorAll("#dBar .emp").forEach(el => { el.hidden = !SHOW_EMP; });
  drawPick("従業員");
  (document.querySelector(`#bColors input[value="${v.色 || "青"}"]`) || document.querySelector("#bColors input")).checked = true;
  $("bMemo").value = v.メモ || "";
  $("dBarWho").textContent = b && b.更新 ? `最後に直した人：${b.更新.だれ}（${b.更新.いつ}）` : "";
  $("bBarDel").hidden = !b;
  openDlg($("dBar"));
  showOffInfo(); showHeadInfo(); numPlaceholder();
  if (!b) $("bWork").focus();
}
let lastColor = "青";
function readBarForm() {
  let a = $("bFrom").value, z = $("bTo").value;
  if (z < a) [a, z] = [z, a];
  const p = parseHead($("bNum").value), lnp = parseHead((calLane() || {}).人数 || "");
  // その日だけの人数は、期間の中でふだんの人数と違う日だけ残す
  const hd = {};
  for (const [k, t] of Object.entries(barCtx.人数日 || {})) { const q = parseHead(t); if (k >= a && k <= z && q && q.txt !== (p ? p.txt : lnp ? lnp.txt : "")) hd[k] = q.txt; }
  return {
    段: $("bLane").value, 作業: $("bWork").value.trim(), 開始: a, 終了: z,
    社員: [...new Set(getStaffRows().map(x => resolveName(x, namesFor("社員"))))],
    従業員: [...new Set(splitNames($("bEmp").value).map(x => resolveName(x, namesFor("従業員"))))], 人数: p ? p.total : "", 人数内訳: p && p.parts.length > 1 ? p.txt : undefined, 人数日: Object.keys(hd).length ? hd : undefined,
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
  if ($("bNum").value.trim() && !parseHead($("bNum").value)) { alert("作業員は数字で入れてください（例：3、1+3）。"); $("bNum").focus(); return; }
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
// 「続きを作る」は 2026-10-07 にやめた（休みをはさむ工程は「休工日を設定」で1本に書く。ユーザー指定）
$("bNum").addEventListener("input", e => { const t = toHalf(e.target.value).replace(/＋/g, "+"); if (t !== e.target.value) e.target.value = t; autoNum = false; showHeadInfo(); });

/* ---------- 日ごとの人数（2026-10-14 ユーザー：工種の作業員の人数を日ごとに変えたい） ----------
   作業する日ごとに欄。空＝ふだんの人数（バーの小窓の「作業員」）。決めると barCtx.人数日 に入り、バーの「保存」で決まる */
// 作業員の欄の薄い字：工種の人数があればそれ（空ならその人数になる）
function numPlaceholder() { const ln = calLane(), t = ln && ln.人数; $("bNum").placeholder = isGai(ln) ? (t || "") + "（外注 " + gaiName(ln) + "）" : t ? t + "（工種）" : "例：1+3"; }
$("bLane").addEventListener("change", numPlaceholder);
function showHeadInfo() {
  let a = $("bFrom").value, z = $("bTo").value;
  if (z < a) [a, z] = [z, a];
  const ks = Object.keys(barCtx.人数日 || {}).filter(k => k >= a && k <= z);
  $("bHeadInfo").textContent = ks.length ? `日ごとに変えた日 ${ks.length}日（${ks.sort().slice(0, 3).map(md).join("・")}${ks.length > 3 ? "…" : ""}）` : "";
}
["bFrom", "bTo"].forEach(id => $(id).addEventListener("change", showHeadInfo));
function drawHeadCal() {
  let a = $("bFrom").value, z = $("bTo").value;
  if (z < a) [a, z] = [z, a];
  const ln = calLane(), b = { 休み: barCtx.休み, 出勤: barCtx.出勤 };
  const base = parseHead($("bNum").value || (ln && ln.人数) || ""), ph = base ? base.txt : "";
  const h = [];
  for (let m = a.slice(0, 7) + "-01"; m <= z; m = keyOf(new Date(+m.slice(0, 4), +m.slice(5, 7), 1))) {
    h.push(`<div class="cal"><div class="calm">${+m.slice(0, 4)}年${+m.slice(5, 7)}月</div><div class="calg hg">` +
      [...WD].map((x, i) => `<div class="calw ${i === 0 ? "sun" : i === 6 ? "sat" : ""}">${x}</div>`).join("") +
      `<div></div>`.repeat(dayDow(m)));
    for (let k = m; k.slice(0, 7) === m.slice(0, 7); k = dayAdd(k, 1)) {
      const hol = holidayOf(k), wd = dayDow(k), num = `<span class="${hol || wd === 0 ? "red" : wd === 6 ? "blue" : ""}">${+k.slice(8)}</span>`;
      if (k < a || k > z) { h.push(`<div class="cald out">${num}</div>`); continue; }
      if (!barWorks(ln, b, k)) { h.push(`<div class="cald hd hoff" title="休み">${num}<i>休</i></div>`); continue; }
      const v = barCtx.人数日[k] || "";
      h.push(`<div class="cald hd ${v ? "ov" : ""}">${num}<input type="text" data-k="${k}" value="${esc(v)}" placeholder="${esc(ph)}" title="${esc(md(k) + "の人数（空＝ふだん " + (fmtHead(ph) || "なし") + "）")}"></div>`);
    }
    h.push(`</div></div>`);
  }
  $("headBody").innerHTML = h.join("");
  sumHeadCal();
}
function sumHeadCal() {
  const ln = calLane(), base = parseHead($("bNum").value || (ln && ln.人数) || "");
  let tot = 0, n = 0, bad = 0;
  $("headBody").querySelectorAll("input[data-k]").forEach(i => {
    const v = i.value.trim(), p = v ? parseHead(v) : base;
    i.parentElement.classList.toggle("ov", !!v); i.classList.toggle("bad", !!v && !p);
    if (v && !p) bad++;
    if (p) tot += p.total;
    n++;
  });
  $("headSum").textContent = `作業する日 ${n}日　のべ ${tot}人` + (bad ? `　（数字でない欄が ${bad} か所）` : "");
}
$("headBody").addEventListener("input", e => {
  const i = e.target;
  if (!i.dataset.k) return;
  const t = toHalf(i.value).replace(/＋/g, "+");
  if (t !== i.value) i.value = t;
  sumHeadCal();
});
$("bHeadBtn").onclick = () => {
  if (!$("bFrom").value || !$("bTo").value) { alert("先に開始と終了を入れてください。"); return; }
  if ($("bNum").value.trim() && !parseHead($("bNum").value)) { alert("先に「作業員」を数字で入れてください（例：3、1+3）。"); return; }
  $("dHeadTtl").textContent = "日ごとの人数　" + ($("bWork").value || "（作業名なし）");
  drawHeadCal();
  openDlg($("dHead"));
};
$("bHeadReset").onclick = () => { $("headBody").querySelectorAll("input[data-k]").forEach(i => { i.value = ""; }); sumHeadCal(); };
$("bHeadCancel").onclick = () => $("dHead").close();
$("fHead").addEventListener("submit", e => {
  e.preventDefault();
  const ins = [...$("headBody").querySelectorAll("input[data-k]")];
  const bad = ins.find(i => i.value.trim() && !parseHead(i.value));
  if (bad) { alert(md(bad.dataset.k) + " の人数が数字ではありません（例：3、1+3、0）。"); bad.focus(); return; }
  // 欄の無い日（休みの日など）に前から入っていた分は残す（休みを戻したときのため。期間の外はバーの保存で捨てる）
  const keep = Object.fromEntries(Object.entries(barCtx.人数日).filter(([k]) => !ins.some(i => i.dataset.k === k)));
  for (const i of ins) if (i.value.trim()) keep[i.dataset.k] = parseHead(i.value).txt;
  barCtx.人数日 = keep;
  $("dHead").close();
  showHeadInfo();
});

// 社員・従業員を名簿の一覧（フルネーム）から選ぶ。選んである人は ✓。もう一度選ぶと外す
let autoNum = true;
const PICK = { 従業員: ["bEmp", "bEmpPick"] };

/* 社員は1人1欄（2026-10-08 ユーザー：2人以上は欄を増やす。空白が名前の区切りか名字と名前の間か分からなくなるため）
   欄は候補（名簿のフルネーム）から選ぶか打つ。名字だけ打てば名簿のフルネームに寄せる */
function staffRow(n) {
  return `<div class="srow"><input type="text" list="staffList" value="${esc(n || "")}" placeholder="名前（候補から選ぶ）">` +
    `<button type="button" class="danger small" data-del title="この欄を消す">✕</button></div>`;
}
function setStaffRows(list) { $("bStaffRows").innerHTML = (list.length ? list : [""]).map(staffRow).join(""); }
function getStaffRows() { return [...$("bStaffRows").querySelectorAll("input")].map(i => i.value.trim()).filter(Boolean); }
$("bStaffAdd").onclick = () => { $("bStaffRows").insertAdjacentHTML("beforeend", staffRow("")); $("bStaffRows").lastElementChild.querySelector("input").focus(); };
$("bStaffRows").addEventListener("click", e => {
  if (!("del" in e.target.dataset)) return;
  const rows = $("bStaffRows").children;
  if (rows.length <= 1) rows[0].querySelector("input").value = ""; else e.target.closest(".srow").remove();
});
// 名字だけ打って欄を離れたら、名簿のフルネームに直して見せる
$("bStaffRows").addEventListener("change", e => { if (e.target.matches("input")) e.target.value = resolveName(e.target.value.trim(), namesFor("社員")); });
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
const laneRuleText = ln => "作業：平日" + (ln.土 !== false ? "・土" : "") + (ln.日 === true ? "・日" : "") + (ln.祝 === true ? "・祝" : "");
function laneRow(l) {
  const ck = (key, label, def) => `<label class="dw"><input type="checkbox" data-k="${key}" ${(l[key] ?? def) ? "checked" : ""}>${label}</label>`;
  const kb = laneKubun(l), gai = isGai(l);
  return `<div class="ln ${gai ? "gai" : ""}" data-id="${esc(l.id)}"><select class="kb"><option ${kb === "昼" ? "selected" : ""}>昼</option><option ${kb === "夜" ? "selected" : ""}>夜</option></select>` +
    `<input type="text" class="wk" value="${esc(laneWork(l))}" placeholder="作業名（例：立坑工）">` +
    `<select class="sb" title="自社の作業員か外注か"><option value="" ${gai ? "" : "selected"}>作業員</option><option value="外注" ${gai ? "selected" : ""}>外注</option></select>` +
    `<span class="lhd" title="この工種のふだんの作業員の人数（バーで人数を入れなければこの人数）"><input type="text" class="hn" data-h="0" inputmode="numeric" value="${gai ? "" : headPart(l.人数, 0)}">人＋` +
    `<input type="text" class="hn" data-h="1" inputmode="numeric" value="${gai ? "" : headPart(l.人数, 1)}">人</span>` +
    `<span class="lgai" title="外注の業者名とふだんの人数（バーで人数を入れなければこの人数）"><input type="text" class="gy" list="gyoList" value="${esc(l.業者 || "")}" placeholder="業者名">` +
    `<input type="text" class="hn" data-h="g" inputmode="numeric" value="${gai ? esc(String(parseHead(l.人数 || "")?.total ?? "")) : ""}">人</span>` +
    ck("土", "土", true) + ck("日", "日", false) + ck("祝", "祝", false) +
    `<label class="dw ldn" title="終わった工種。工程表ではたたんで隠す"><input type="checkbox" data-done ${l.完了 ? "checked" : ""}>完了</label>` +
    `<button type="button" class="sub small" data-mv="-1" title="上へ">▲</button><button type="button" class="sub small" data-mv="1" title="下へ">▼</button>` +
    `<button type="button" class="danger small" data-del>✕</button></div>`;
}
// 工種の行の 〇人＋〇人。片方だけなら1つの数（"3"）、両方なら "1+3"
function headPart(t, i) { const p = parseHead(t || ""); return p && p.parts[i] != null ? p.parts[i] : ""; }
function laneHeadOf(r) {
  if (r.querySelector("select.sb").value === "外注") { const g = toHalf(r.querySelector('input.hn[data-h="g"]').value).trim(); return g ? String(parseInt(g, 10) || 0) : ""; }
  const [a, b] = [...r.querySelectorAll('.lhd input.hn')].map(i => toHalf(i.value).trim());
  if (!a && !b) return "";
  const x = parseInt(a || "0", 10), y = parseInt(b || "0", 10);
  return b ? (x || 0) + "+" + (y || 0) : String(x || 0);
}
function openSite(id) {
  const s = id ? S.files[fileOfSite(id)].doc : null;
  siteCtx = { id, lanes: s ? s.段.map(l => l.id) : [] };
  $("dSiteTtl").textContent = s ? "現場の設定" : "現場を追加";
  $("sName").value = s ? nm(s) : "";
  $("sCode").value = s ? s.コード || "" : "";
  $("sDone").checked = !!(s && s.完了);
  $("sTanto").value = s ? s.担当 || "" : "";
  // 週間工程表に出す（2026-10-13）
  $("sKouji").value = s ? officialOf(s) : "";
  $("sKiFrom").value = s ? s.工期自 || "" : ""; $("sKiTo").value = s ? s.工期至 || "" : "";
  $("sAddr").value = s ? s.住所 || "" : "";
  siteCtx.map = undefined;   // undefined＝変えない／mapCrop＝新しい画像・切り取りを直した／null＝外す
  showMapBox(null);
  if (s && s.地図) loadMap(s).then(u => { if (siteCtx && siteCtx.id === s.id && siteCtx.map === undefined) showMapBox(u); });
  const lanes = s ? s.段 : [{ id: newId("l"), 区分: "昼", 名前: "" }];
  $("sLanes").innerHTML = lanes.map(laneRow).join("");
  // 業者名の候補：どの現場でも前に入れた業者名
  $("gyoList").innerHTML = [...new Set(Object.values(S.files).flatMap(f => (f.doc && f.doc.段 || []).filter(isGai).map(l => l.業者).filter(Boolean)))]
    .sort((a, b) => a.localeCompare(b, "ja")).map(g => `<option value="${esc(g)}">`).join("");
  $("bSiteDel").hidden = !s; $("bSiteUp").parentElement.hidden = !s;
  // ▲▼ は「設定順」の並びを変える。ほかの並びで見ているときは、その旨を出す
  $("sOrdNote").hidden = sortBy === "manual";
  openDlg($("dSite"));
  if (!s) $("sName").focus();
}
// 使用方法は別の窓（2026-10-14 ユーザー：説明の帯をやめて場所を空ける。パソコン用とスマホ用を分ける）
$("bHelp").onclick = () => {
  const touch = matchMedia("(pointer: coarse)").matches;
  const w = window.open("help.html?v=" + encodeURIComponent($("ver").textContent) + (touch ? "#sp" : "#pc"), "kotei_help", touch ? "" : "width=820,height=900");
  if (w) w.focus();
};
$("sLanes").addEventListener("change", e => {
  if (!e.target.classList.contains("sb")) return;
  const row = e.target.closest(".ln"), gai = e.target.value === "外注";
  row.classList.toggle("gai", gai);
  (gai ? row.querySelector("input.gy") : row.querySelector('.lhd input.hn')).focus();
});
$("sLanes").addEventListener("input", e => { if (e.target.classList.contains("hn")) { const t = toHalf(e.target.value).replace(/[^\d]/g, ""); if (t !== e.target.value) e.target.value = t; } });
$("sLanes").addEventListener("click", e => {
  const row = e.target.closest(".ln");
  if (!row) return;
  if (e.target.dataset.mv) {
    const sib = e.target.dataset.mv === "-1" ? row.previousElementSibling : row.nextElementSibling;
    if (sib) e.target.dataset.mv === "-1" ? row.parentNode.insertBefore(row, sib) : row.parentNode.insertBefore(sib, row);
  } else if ("del" in e.target.dataset) {
    if ($("sLanes").children.length <= 1) { alert("工種は1つ以上要ります。"); return; }
    // バーのある工種も消せる（2026-10-13 ユーザー指定）。バーも一緒に消える。決まるのは「保存」のとき
    const s = siteCtx.id && S.files[fileOfSite(siteCtx.id)].doc;
    const n = s ? s.バー.filter(b => b.段 === row.dataset.id).length : 0;
    if (n && !confirm(`この工種のバー ${n} 本も一緒に消えます（「保存」を押したとき）。よろしいですか？\n残しておくなら「完了」にチェックすると、たたんで隠せます。`)) return;
    row.remove();
  }
});
$("sName").addEventListener("change", () => {
  const n = $("sName").value.trim(), x = M().現場.find(s => s.略称 === n) || M().現場.find(s => s.名前 === n);
  if (!x) return;
  // 現場名は配置表の書き方（略称）、マスタの正式な名前は 正式工事名 へ
  if (x.略称 && n !== x.略称) $("sName").value = x.略称;
  if (!$("sKouji").value && x.名前 && x.名前 !== $("sName").value) $("sKouji").value = x.名前;
  if (!$("sCode").value) $("sCode").value = x.コード || "";
  if (!$("sTanto").value && x.担当) $("sTanto").value = resolveName(x.担当, namesFor("社員"));
  if (!$("sAddr").value && x.場所) $("sAddr").value = x.場所;
});
/* 工種を足すとき、たたむ／そのまま／消す を選ぶ（2026-10-13 ユーザー指定）
   出すのは ①終わった工種（バーがあって、どのバーも今日より前に終わっている）がある ②完了でない工種がもう3つ以上ある（増えすぎないように）とき。
   一覧は完了でない工種ぜんぶ。終わった工種は「たたむ」、まだの工種は「そのまま」をはじめに選んでおく */
const LANE_MAX = 3;
function addLaneRow() { $("sLanes").insertAdjacentHTML("beforeend", laneRow({ id: newId("l"), 区分: "昼", 名前: "" })); $("sLanes").lastElementChild.querySelector("input.wk").focus(); }
$("bLaneAdd").onclick = () => {
  const s = siteCtx.id && S.files[fileOfSite(siteCtx.id)].doc;
  const live = [...$("sLanes").children].filter(r => !r.querySelector("input[data-done]").checked);
  const rows = s ? live.filter(r => s.段.some(x => x.id === r.dataset.id)) : [];
  const fin = rows.filter(r => laneFinished(s, s.段.find(x => x.id === r.dataset.id)));
  if (!fin.length && (live.length < LANE_MAX || !rows.length)) { addLaneRow(); return; }
  $("dOldTtl").textContent = fin.length ? "終わった工種があります" : `工種が ${live.length} つあります`;
  $("oldBody").innerHTML = rows.map(r => {
    const l = s.段.find(x => x.id === r.dataset.id), bs = s.バー.filter(b => b.段 === l.id);
    const last = bs.map(b => b.終了).sort().pop(), done = laneFinished(s, l);
    const note = !bs.length ? "バーなし" : done ? `バー ${bs.length} 本・${md(last)} に終了` : `バー ${bs.length} 本・${md(last)} まで（まだ終わっていない）`;
    const def = done ? 0 : 1;
    return `<div class="oldrow ${done ? "fin" : ""}" data-id="${esc(l.id)}"><div class="onm"><b>${esc(laneName(l))}</b><span class="note">${note}</span></div>` +
      ["たたむ", "そのまま", "消す"].map((v, i) => `<label class="${i === 2 ? "del" : ""}"><input type="radio" name="o_${esc(l.id)}" value="${v}" ${i === def ? "checked" : ""}>${v}</label>`).join("") + `</div>`;
  }).join("");
  openDlg($("dOld"));
};
$("fOld").addEventListener("submit", e => {
  e.preventDefault();
  const picks = [...$("oldBody").children].map(d => ({ id: d.dataset.id, v: d.querySelector("input:checked").value }));
  const dels = picks.filter(p => p.v === "消す");
  if (dels.length) {
    const s = S.files[fileOfSite(siteCtx.id)].doc;
    const n = s.バー.filter(b => dels.some(p => p.id === b.段)).length;
    if (!confirm(`工種 ${dels.length} 件とそのバー ${n} 本を消します（現場の設定の「保存」を押したとき）。よろしいですか？`)) return;
  }
  for (const p of picks) {
    const row = $("sLanes").querySelector(`.ln[data-id="${CSS.escape(p.id)}"]`);
    if (!row) continue;
    if (p.v === "たたむ") row.querySelector("input[data-done]").checked = true;
    if (p.v === "消す") row.remove();
  }
  $("dOld").close();
  addLaneRow();
});
$("bOldCancel").onclick = () => $("dOld").close();
$("fSite").addEventListener("submit", e => {
  e.preventDefault();
  const name = $("sName").value.trim();
  if (!name) return;
  const lanes = [...$("sLanes").children].map(r => {
    const l = { id: r.dataset.id, 区分: r.querySelector("select.kb").value, 名前: r.querySelector("input.wk").value.trim() };
    r.querySelectorAll("input[data-k]").forEach(c => { l[c.dataset.k] = c.checked; });
    if (r.querySelector("input[data-done]").checked) l.完了 = true;
    const hn = laneHeadOf(r);
    if (hn) l.人数 = hn;
    if (r.querySelector("select.sb").value === "外注") { l.種別 = "外注"; const g = r.querySelector("input.gy").value.trim(); if (g) l.業者 = g; }
    return l;
  });
  // その日だけの休み・出勤は、保存するときの最新の中身から引き継ぐ
  const keepDays = (doc, l) => { const o = (doc.段 || []).find(x => x.id === l.id); return { ...l, 休み: o ? o.休み || [] : [], 出勤: o ? o.出勤 || [] : [] }; };
  const v = { 名前: name, 完了: $("sDone").checked, 略称: name, コード: toHalf($("sCode").value.trim()), 担当: resolveName($("sTanto").value.trim(), namesFor("社員")),
    工事名: $("sKouji").value.trim(), 工期自: $("sKiFrom").value, 工期至: $("sKiTo").value, 住所: $("sAddr").value.trim() };
  if (siteCtx.map === null) v.地図 = null;
  const newMap = siteCtx.map;
  if (siteCtx.id) {
    // この小窓で消した工種（開いたときにあって、いま無いもの）のバーも消す。ほかの人が同時に足した工種は残す
    const opened = siteCtx.lanes || [], gone = opened.filter(id => !lanes.some(l => l.id === id));
    if (newMap) { const sid = siteCtx.id; renderMap(newMap).then(m => saveMap(sid, m)); }
    mutate(fileOfSite(siteCtx.id), doc => {
      Object.assign(doc, v);
      const extra = (doc.段 || []).filter(l => !opened.includes(l.id) && !lanes.some(x => x.id === l.id));
      doc.段 = [...lanes.map(l => keepDays(doc, l)), ...extra];
      if (gone.length) doc.バー = doc.バー.filter(b => !gone.includes(b.段));
    });
  } else {
    const id = newId("s"), file = fileOfSite(id);
    const maxOrd = Math.max(0, ...sites().map(s => s.並び ?? 0));
    S.files[file] = { doc: { id, ...v, 段: lanes, バー: [], 並び: maxOrd + 10 }, eTag: null, pending: [], saving: false, timer: null, err: null };
    mutate(file, () => { });
    if (newMap) renderMap(newMap).then(m => saveMap(id, m));
  }
  $("dSite").close();
});
$("bSiteCancel").onclick = () => $("dSite").close();

/* 現場の地図（2026-10-13 ユーザー：週間工程表の右の地図は別に用意して貼り付けたい）
   工程データ\地図_<現場id>.jpg に置く（.json ではないので30秒ごとの読み直しには入らない）。現場.地図＝入れた時刻（版）。
   画像は長い辺 1400px の JPEG に縮めてから保存。読んだものは S.maps に覚える */
S.maps = {};
const mapFile = id => "地図_" + id + ".jpg";
async function loadMap(s) {
  if (!s || !s.地図) return null;
  const c = S.maps[s.id];
  if (c && c.ver === s.地図) return c.url;
  try { const url = await store.getRaw(mapFile(s.id)); S.maps[s.id] = { ver: s.地図, url }; return url; }
  catch (e) { console.error(e); return null; }
}
function saveMap(id, m) {
  const ver = Date.now();
  toast("地図を保存しています…");
  store.putRaw(mapFile(id), m.blob)
    .then(() => { S.maps[id] = { ver, url: m.url }; mutate(fileOfSite(id), d => { d.地図 = ver; }); toast("地図を保存しました"); })
    .catch(e => { console.error(e); toast("地図の保存に失敗：" + e.message); });
}
// 画像を読み終わるまで待つ（decode() は画面が隠れていると終わらないことがあるので onload。5秒で諦める）
const loadImg = src => new Promise(ok => { const i = new Image(); const t = setTimeout(() => ok(null), 5000); i.onload = () => { clearTimeout(t); ok(i); }; i.onerror = () => { clearTimeout(t); ok(null); }; i.src = src; });
/* 地図の切り取り（2026-10-14 ユーザー：Excel の週間工程表にちょうどいい縦横比に加工したい・ドラッグ＆ドロップでも入れたい）
   枠＝Excel の地図の場所（AA〜AJ 列 × 4行目〜記事の下。作業8行のとき 幅 247.5pt・高さ 528pt。出した Excel を実際の Excel で測った）と同じ縦横の比。
   画像は枠の中でつかんで動かす・大きさを変える。はじめは「全体を入れる」（余りは白）。保存は 754×1400px の JPEG。
   mapCrop＝{ img, s:倍率, x, y }（枠の左上からの位置・枠の px）。siteCtx.map に入れると「保存」で書く */
const MAP_RATIO = 247.5 / 528, MAP_FH = 300, MAP_FW = Math.round(MAP_FH * MAP_RATIO), MAP_OUT_H = 1400, MAP_OUT_W = Math.round(MAP_OUT_H * MAP_RATIO);
let mapCrop = null;
const mapFit = img => Math.min(MAP_FW / img.naturalWidth, MAP_FH / img.naturalHeight);
const mapFill = img => Math.max(MAP_FW / img.naturalWidth, MAP_FH / img.naturalHeight);
const mapMax = img => mapFill(img) * 4;
function mapClamp() {
  const c = mapCrop, W = c.img.naturalWidth * c.s, H = c.img.naturalHeight * c.s;
  c.x = W <= MAP_FW ? Math.min(Math.max(c.x, 0), MAP_FW - W) : Math.min(Math.max(c.x, MAP_FW - W), 0);
  c.y = H <= MAP_FH ? Math.min(Math.max(c.y, 0), MAP_FH - H) : Math.min(Math.max(c.y, MAP_FH - H), 0);
}
function mapZoomTo(ns, cx = MAP_FW / 2, cy = MAP_FH / 2) {
  const c = mapCrop, lo = mapFit(c.img), hi = mapMax(c.img);
  ns = Math.min(Math.max(ns, lo), hi);
  c.x = cx - (cx - c.x) * ns / c.s; c.y = cy - (cy - c.y) * ns / c.s; c.s = ns;
  mapClamp(); mapDraw();
}
function mapSetFit(fill) {
  const c = mapCrop; c.s = fill ? mapFill(c.img) : mapFit(c.img);
  c.x = (MAP_FW - c.img.naturalWidth * c.s) / 2; c.y = (MAP_FH - c.img.naturalHeight * c.s) / 2; mapDraw();
}
function mapDraw() {
  const box = $("sMapBox"), c = mapCrop;
  box.style.width = MAP_FW + "px"; box.style.height = MAP_FH + "px";
  box.classList.toggle("has", !!c);
  $("sMapDel").hidden = !c; $("sMapAdj").hidden = !c;
  if (!c) { box.innerHTML = `<span class="note">（地図なし）<br>この枠を押して Ctrl＋V で貼り付け、<br>または画像ファイルをドラッグ</span>`; return; }
  let im = box.querySelector("img");
  if (!im || im.src !== c.img.src) { box.innerHTML = ""; im = c.img; im.alt = "地図"; im.draggable = false; box.appendChild(im); }
  Object.assign(im.style, { left: c.x + "px", top: c.y + "px", width: c.img.naturalWidth * c.s + "px", height: c.img.naturalHeight * c.s + "px" });
  const lo = mapFit(c.img), hi = mapMax(c.img);
  $("sMapZoom").value = Math.round(100 * Math.log(c.s / lo) / Math.log(hi / lo));
}
// 画像を読んで枠に入れる。changed＝新しい画像（保存のとき書く）
async function setMapImage(src, changed) {
  const img = await loadImg(src);
  if (!img) throw new Error("画像を読めません");
  mapCrop = { img, s: 1, x: 0, y: 0 };
  mapSetFit(false);
  if (changed) siteCtx.map = mapCrop;
}
function mapTouched() { if (mapCrop) siteCtx.map = mapCrop; }
async function renderMap(c) {
  const cv = document.createElement("canvas"), k = MAP_OUT_W / MAP_FW;
  cv.width = MAP_OUT_W; cv.height = MAP_OUT_H;
  const x = cv.getContext("2d");
  x.fillStyle = "#fff"; x.fillRect(0, 0, cv.width, cv.height);
  x.imageSmoothingQuality = "high";
  x.drawImage(c.img, c.x * k, c.y * k, c.img.naturalWidth * c.s * k, c.img.naturalHeight * c.s * k);
  const blob = await new Promise(r => cv.toBlob(r, "image/jpeg", 0.85));
  return { blob, url: cv.toDataURL("image/jpeg", 0.85) };
}
function showMapBox(url) {
  mapCrop = null; mapDraw();
  if (url) setMapImage(url, false).catch(e => console.error(e));
}
async function takeMap(file) {
  if (!file || !/^image\//.test(file.type)) { toast("画像のファイルではありません"); return; }
  try { await setMapImage(URL.createObjectURL(file), true); toast("地図を入れました。枠に合わせて動かせます（「保存」で決まります）"); }
  catch (e) { console.error(e); toast("画像を読めませんでした"); }
}
$("sMapFile").onchange = e => { takeMap(e.target.files[0]); e.target.value = ""; };
$("sMapDel").onclick = () => { siteCtx.map = null; mapCrop = null; mapDraw(); };
$("sMapFit").onclick = () => { mapSetFit(false); mapTouched(); };
$("sMapFill").onclick = () => { mapSetFit(true); mapTouched(); };
$("sMapZoom").addEventListener("input", e => {
  const c = mapCrop; if (!c) return;
  const lo = mapFit(c.img), hi = mapMax(c.img);
  mapZoomTo(lo * Math.pow(hi / lo, e.target.value / 100)); mapTouched();
});
$("sMapBox").addEventListener("wheel", e => {
  if (!mapCrop) return;
  e.preventDefault();
  const r = $("sMapBox").getBoundingClientRect();
  mapZoomTo(mapCrop.s * (e.deltaY < 0 ? 1.1 : 1 / 1.1), e.clientX - r.left, e.clientY - r.top); mapTouched();
}, { passive: false });
$("sMapBox").addEventListener("pointerdown", e => {
  $("sMapBox").focus();
  if (!mapCrop || e.button > 0) return;
  e.preventDefault();
  const box = $("sMapBox"), sx = e.clientX, sy = e.clientY, ox = mapCrop.x, oy = mapCrop.y;
  try { box.setPointerCapture(e.pointerId); } catch (x) { }
  const mv = ev => { mapCrop.x = ox + ev.clientX - sx; mapCrop.y = oy + ev.clientY - sy; mapClamp(); mapDraw(); mapTouched(); };
  const up = () => { box.removeEventListener("pointermove", mv); box.removeEventListener("pointerup", up); box.removeEventListener("pointercancel", up); };
  box.addEventListener("pointermove", mv); box.addEventListener("pointerup", up); box.addEventListener("pointercancel", up);
});
// 小窓の中で画像を貼り付け（文字の欄に貼るときは文字のまま）
$("dSite").addEventListener("paste", e => {
  const it = [...(e.clipboardData || {}).items || []].find(i => i.type.startsWith("image/"));
  if (!it) return;
  e.preventDefault(); takeMap(it.getAsFile());
});
// ドラッグ＆ドロップ（現場の小窓のどこへ落としてもよい）。ほかの所へ落としてもブラウザが画像を開いてしまわないように止める
const hasFiles = e => [...((e.dataTransfer || {}).types || [])].includes("Files");
document.addEventListener("dragover", e => { if (hasFiles(e)) e.preventDefault(); });
document.addEventListener("drop", e => { if (hasFiles(e)) e.preventDefault(); });
$("dSite").addEventListener("dragover", e => { if (!hasFiles(e)) return; e.preventDefault(); e.dataTransfer.dropEffect = "copy"; $("sMapBox").classList.add("drop"); });
$("dSite").addEventListener("dragleave", e => { if (!$("dSite").contains(e.relatedTarget)) $("sMapBox").classList.remove("drop"); });
$("dSite").addEventListener("drop", e => {
  $("sMapBox").classList.remove("drop");
  if (!hasFiles(e)) return;
  e.preventDefault();
  const f = [...e.dataTransfer.files].find(x => /^image\//.test(x.type));
  if (f) takeMap(f); else toast("画像のファイルを落としてください（jpg・png など）");
});
$("bSiteDel").onclick = () => {
  const s = S.files[fileOfSite(siteCtx.id)].doc;
  if (!confirm(`「${nm(s)}」を工程表から消します（バー ${s.バー.length} 本も見えなくなります）。よろしいですか？`)) return;
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
function setFrom(k) {
  S.view.from = fromKey(k); $("inFrom").value = S.view.from;
  try { localStorage.setItem("kotei_view", JSON.stringify(S.view)); } catch (e) { }
  render();
}
$("inFrom").onchange = e => e.target.value && setFrom(e.target.value);
$("inMonths").onchange = e => { S.view.months = +e.target.value; setFrom(S.view.from); };
// ◀ ▶ は1か月ずつ（日はそのまま）。今日＝今日から
$("bPrev").onclick = () => setFrom(addMonths(fromKey(S.view.from), -1));
$("bNext").onclick = () => setFrom(addMonths(fromKey(S.view.from), 1));
$("bToday").onclick = () => { setFrom(todayKey()); $("wrap").scrollLeft = 0; };
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
$("inSort").value = sortBy;
$("inSort").onchange = e => { sortBy = e.target.value; try { localStorage.setItem("kotei_sort", sortBy); } catch (x) { } render(); };
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

/* ---------- 拡大縮小（2026-10-13 ユーザー：Ctrl＋ホイールでブラウザごと拡大すると誤動作する → 表だけを拡大縮小）
   Ctrl＋ホイール・Ctrl＋＋／－／0・上の「－ 100% ＋」・2本指ピンチ。上の操作や小窓の大きさは変えない。
   倍率はその人のブラウザに覚える（kotei_zoom）。CSS の zoom を表（#chart・人の流れ・一覧の表）にかける ---------- */
let ZOOM = 1;
try { ZOOM = Math.min(2, Math.max(0.4, +localStorage.getItem("kotei_zoom") || 1)); } catch (e) { }
const scroller = () => S.tab === "chart" ? $("wrap") : S.tab === "flow" ? $("flow") : $("list");
function applyZoom() {
  document.documentElement.style.setProperty("--z", ZOOM);
  // 拡大したときも現場名・工種名の列は広げない（2026-10-14 ユーザー指定）：左の列だけ 1/倍率 をかけて、見た目を 100% のままにする。縮小のときは一緒に縮む
  document.documentElement.style.setProperty("--lz", ZOOM > 1 ? 1 / ZOOM : 1);
  $("bZoom").textContent = Math.round(ZOOM * 100) + "%";
}
// ax, ay：画面上でこの点が動かないように拡大する（指やマウスの位置）。省けば表示の左上
function setZoom(z, ax, ay) {
  z = Math.min(2, Math.max(0.4, Math.round(z * 100) / 100));
  const sc = scroller(), r = sc.getBoundingClientRect();
  const mx = ax == null ? 0 : ax - r.left, my = ay == null ? 0 : ay - r.top;
  const cx = (sc.scrollLeft + mx) / ZOOM, cy = (sc.scrollTop + my) / ZOOM;
  ZOOM = z; applyZoom();
  sc.scrollLeft = cx * z - mx; sc.scrollTop = cy * z - my;
  try { localStorage.setItem("kotei_zoom", z); } catch (e) { }
}
// Ctrl＋ホイール（ノートのタッチパッドのピンチも Ctrl＋ホイールで来る）はブラウザの拡大を止めて表だけ
window.addEventListener("wheel", e => {
  if (!e.ctrlKey) return;
  e.preventDefault();
  if (document.querySelector("dialog[open]")) return;
  const d = Math.max(-100, Math.min(100, e.deltaMode === 1 ? e.deltaY * 33 : e.deltaY));
  setZoom(ZOOM * Math.exp(-d * 0.0015), e.clientX, e.clientY);
}, { passive: false });
window.addEventListener("keydown", e => {
  if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
  const k = e.key;
  if (k === "+" || k === "=" || k === ";") { e.preventDefault(); setZoom(ZOOM * 1.1); }
  else if (k === "-" || k === "_") { e.preventDefault(); setZoom(ZOOM / 1.1); }
  else if (k === "0") { e.preventDefault(); setZoom(1); }
});
$("bZoomIn").onclick = () => setZoom(ZOOM * 1.1);
$("bZoomOut").onclick = () => setZoom(ZOOM / 1.1);
$("bZoom").onclick = () => setZoom(1);
applyZoom();

/* 2本指：スワイプ＝スクロール、ピンチ＝拡大縮小（タブレット）。工程表・人の流れ・一覧で */
let g2 = null;
const dist2 = t => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
const mid2 = t => [(t[0].clientX + t[1].clientX) / 2, (t[0].clientY + t[1].clientY) / 2];
["wrap", "flow", "list"].forEach(id => {
  const sc = $(id);
  sc.addEventListener("touchstart", e => {
    if (e.touches.length !== 2) { if (e.touches.length > 2) g2 = null; return; }
    if (tp) clearTimeout(tp.timer);
    tp = null; clearTouchMarks(); cancelDrag(); hideCtx();
    const [x, y] = mid2(e.touches), r = sc.getBoundingClientRect();
    g2 = { sc, d0: dist2(e.touches), z0: ZOOM, zoomed: false,
      cx: (sc.scrollLeft + x - r.left) / ZOOM, cy: (sc.scrollTop + y - r.top) / ZOOM };
  }, { passive: false });
  sc.addEventListener("touchmove", e => {
    if (!g2 || g2.sc !== sc || e.touches.length !== 2) return;
    if (e.cancelable) e.preventDefault();
    const k = dist2(e.touches) / g2.d0;
    // 指の間が少し動いただけでは拡大しない（2本指で横に流したいとき）
    if (!g2.zoomed && Math.abs(k - 1) > 0.12) g2.zoomed = true;
    const z = Math.min(2, Math.max(0.4, g2.zoomed ? g2.z0 * k : g2.z0));
    if (z !== ZOOM) { ZOOM = Math.round(z * 100) / 100; applyZoom(); }
    const [x, y] = mid2(e.touches), r = sc.getBoundingClientRect();
    sc.scrollLeft = g2.cx * ZOOM - (x - r.left); sc.scrollTop = g2.cy * ZOOM - (y - r.top);
  }, { passive: false });
  sc.addEventListener("touchend", e => {
    if (!g2 || e.touches.length >= 2) return;
    if (g2.zoomed) try { localStorage.setItem("kotei_zoom", ZOOM); } catch (x) { }
    g2 = null;
  });
});
// 工程表の上で動かしているときは、ブラウザに画面を流させない（iPad の Safari 向け。touch-action が効かないときの保険）
$("wrap").addEventListener("touchmove", e => { if ((drag || (tp && tp.pan)) && e.cancelable) e.preventDefault(); }, { passive: false });

/* ---------- 印刷・PDF（2026-10-13 ユーザー指定）
   「印刷・PDF」の小窓で 期間・現場・用紙・ページの分け方 を選んで印刷。Ctrl＋P は画面の期間・現場を1ページで。
   1ページずつ、用紙に収まるよう工程表を作り直して写し（#printPages）を作る。どのページにも上の年月日の行が出る。
   PDF はブラウザの印刷画面で「送信先：PDF に保存」 ---------- */
const PAPER = { A3: [420, 297], A4: [297, 210] };
const PAGE_MM = 7;
const root = document.documentElement.style;
// 用紙の向き・大きさは印刷画面を開く前から決めておく（2026-10-13 ユーザー：A3 にしても A4 で出た）
function setPageStyle() {
  const size = $("inPaper").value || "A3";
  let st = $("printPage");
  if (!st) { st = document.createElement("style"); st.id = "printPage"; document.head.appendChild(st); }
  st.textContent = `@page { size: ${size} landscape; margin: ${PAGE_MM}mm; }`;
}
/* 1ページを用紙いっぱいに（2026-10-13 ユーザー：下と右に余白ができないように／1ページに収めたいのにはみ出す）
   ①行 30px の高さで縦が収まる倍率（文字ごと拡大縮小）→②その倍率で横が埋まる 1日の幅（18px より細くなるなら 18〜24px にして横に合わせる）
   →③それでも縦がはみ出すなら、縦に合わせてさらに縮める（右が少し空く）→④下が余れば行の高さを広げて埋める。返り値＝倍率 */
function fitChart(W, H) {
  const cw = $("chart"), n = range().days.length;
  // 高さは「完了 n工種」の細い行を除いて測る（印刷には出さない）
  const tall = () => { const hw = cw.querySelector(".hrow.w"); let t = hw.offsetTop + hw.offsetHeight + $("body").offsetHeight; cw.querySelectorAll(".lane.donerow").forEach(x => { t -= x.offsetHeight; }); return t; };
  const left = () => cw.querySelector(".corner").offsetWidth;
  root.setProperty("--rh", "30px"); root.setProperty("--dw", "24px"); renderChart();
  let z = Math.min(2.6, H / tall());
  let dw = (W / z - left()) / n;
  if (dw < 18) dw = Math.max(18, Math.min(24, dw));
  if (dw > 90) dw = 90;
  root.setProperty("--dw", dw + "px"); renderChart();
  z = Math.min(W / (left() + dw * n), H / tall());
  if (tall() * z < H) {
    let lo = 30, hi = 240;
    for (let i = 0; i < 12; i++) { const m = (lo + hi) / 2; root.setProperty("--rh", m + "px"); renderChart(); if (tall() * z <= H) lo = m; else hi = m; }
    root.setProperty("--rh", lo + "px"); renderChart();
  }
  return z;
}
function fitFlow(W, H) {
  const tb = () => $("flow").querySelector("table"), n = range().days.length;
  root.setProperty("--fw", "46px"); root.setProperty("--fh", "26px"); renderFlow();
  const nameW = () => tb().querySelector("th.nm").offsetWidth;
  let z = Math.min(1.6, H / tb().offsetHeight);
  let fw = (W / z - nameW()) / n;
  if (fw < 26) fw = Math.max(26, Math.min(46, fw));
  if (fw > 120) fw = 120;
  root.setProperty("--fw", fw + "px"); renderFlow();
  z = Math.min(W / tb().offsetWidth, H / tb().offsetHeight);
  if (tb().offsetHeight * z < H) {
    let lo = 26, hi = 70;
    for (let i = 0; i < 12; i++) { const m = (lo + hi) / 2; root.setProperty("--fh", m + "px"); renderFlow(); if (tb().offsetHeight * z <= H) lo = m; else hi = m; }
    root.setProperty("--fh", lo + "px"); renderFlow();
  }
  return z;
}
// 期間を nx ページに分ける。期間がちょうど nか月で nx で割り切れるなら月の切れ目で、そうでなければ日数で等分
function dayChunks(from, end, nx) {
  const total = dayDiff(from, end) + 1;
  nx = Math.max(1, Math.min(nx, total));
  let months = 0;
  for (let m = 1; m <= 24; m++) if (dayAdd(addMonths(from, m), -1) === end) { months = m; break; }
  const out = [];
  for (let i = 0; i < nx; i++) {
    let a, z;
    if (months && months % nx === 0) { const k = months / nx; a = addMonths(from, i * k); z = dayAdd(addMonths(from, (i + 1) * k), -1); }
    else { a = dayAdd(from, Math.round(total * i / nx)); z = dayAdd(from, Math.round(total * (i + 1) / nx) - 1); }
    out.push({ a, z });
  }
  return out;
}
// 現場を ny ページに分ける（工種の行の数がだいたい同じになるように。並びはそのまま）
function siteGroups(ids, ny) {
  const list = shownSitesOf(ids);
  const wt = s => folded.has(s.id) ? 1 : visibleLanes(s).length;
  const total = list.reduce((a, s) => a + wt(s), 0);
  ny = Math.max(1, Math.min(ny, list.length));
  const out = Array.from({ length: ny }, () => []);
  let acc = 0;
  for (const s of list) { const g = Math.min(ny - 1, Math.floor(acc / total * ny)); out[g].push(s.id); acc += wt(s); }
  return out.filter(g => g.length);
}
const shownSitesOf = ids => sortSites(sites().filter(s => ids.includes(s.id)));
const kd = (k, y) => (y ? k.slice(0, 4) + "年" : "") + `${+k.slice(5, 7)}月${+k.slice(8)}日`;
/* 週間工程表（2026-10-13 ユーザー：Excel の週間工程表の書式でも出したい）。元：現場別\510012金沢線配水\金沢週間工程.xlsx
   1現場1ページ。開始日の週の月曜から3週間（今週・来週・再来週）。行＝作業（バーの作業名。夜間の工種は（夜間））・作業員（その日の人数の合計）・記事（バーのメモ）。
   右に現場の地図（現場の設定で貼り付けた画像）。土日祝・会社の休みの列は灰色。1500×1040px で組んで用紙に合わせて縮める */
const COMPANY = "株式会社 ハシックス";
const WK_W = 1500, WK_H = 1040;
const monday = k => dayAdd(k, -((dayDow(k) + 6) % 7));
const reiwa = (k, wd) => { const y = +k.slice(0, 4) - 2018; return `令和${pad2(y)}年${k.slice(5, 7)}月${k.slice(8, 10)}日` + (wd ? `(${WD[dayDow(k)]})` : ""); };
function weekPage(s, mon) {
  const days = Array.from({ length: 21 }, (_, i) => dayAdd(mon, i)), end = days[20];
  const laneOf = b => s.段.find(l => l.id === b.段);
  const hasNight = s.段.some(isNight);
  // 行：作業ごと（同じ作業名・同じ昼夜はまとめる）
  const rows = [], byKey = {};
  for (const b of s.バー) {
    if (b.終了 < mon || b.開始 > end) continue;
    const ln = laneOf(b);
    const name = (b.作業 || (ln ? laneWork(ln) : "") || "（作業）") + (ln && isNight(ln) ? "（夜間）" : hasNight ? "（昼間）" : "");
    let r = byKey[name];
    if (!r) { r = byKey[name] = { name, on: new Set(), first: b.開始 }; rows.push(r); }
    if (b.開始 < r.first) r.first = b.開始;
    days.forEach(k => { if (k >= b.開始 && k <= b.終了 && barWorks(ln, b, k)) r.on.add(k); });
  }
  rows.sort((a, b) => a.first.localeCompare(b.first));
  const nRows = Math.max(8, rows.length);
  while (rows.length < nRows) rows.push({ name: "", on: new Set() });
  const headH = 104, weekH = 26 + 22 + 30 + 26, workerH = 46, noteH = 92;
  const rowH = Math.max(22, Math.min(70, Math.floor((WK_H - headH - weekH - workerH - noteH) / nRows)));
  const offCls = k => (isRest(k) || dayDow(k) === 0 || dayDow(k) === 6 || holidayOf(k)) ? "wkoff" : "";
  const numCls = k => (dayDow(k) === 0 || holidayOf(k)) ? "red" : dayDow(k) === 6 ? "blue" : "";
  // 作業員：その日に作業するバーの人数の合計
  const workers = days.map(k => s.バー.reduce((a, b) => a + (k >= b.開始 && k <= b.終了 && barWorks(laneOf(b), b, k) ? headAt(b, k, laneOf(b)) : 0), 0));
  // 記事：その週に始まる（1週目はそれより前から続く）バーのメモ
  const notes = [0, 1, 2].map(w => {
    const a = days[w * 7], z = days[w * 7 + 6];
    return s.バー.filter(b => b.メモ && (w === 0 ? b.開始 <= z && b.終了 >= a : b.開始 >= a && b.開始 <= z))
      .sort((x, y) => x.開始.localeCompare(y.開始))
      .map(b => `<div><b>${+(b.開始 < a ? a : b.開始).slice(8)}日</b>　${esc(b.作業 || "")}　${esc(b.メモ)}</div>`).join("");
  });
  // 月の見出し（月が変わる所で区切る）
  const months = [];
  days.forEach(k => { const m = +k.slice(5, 7), last = months[months.length - 1]; if (last && last.m === m) last.n++; else months.push({ m, n: 1 }); });
  const map = S.maps[s.id] && s.地図 && S.maps[s.id].ver === s.地図 ? S.maps[s.id].url : null;
  const totalRows = 4 + nRows + 2;
  // 地図の列は Excel の地図の場所と同じ縦横の比（現場の設定の枠で切り取ってある）
  const mapW = Math.round((weekH + nRows * rowH + workerH + noteH) * MAP_RATIO);
  return `<div class="wkpage">
  <table class="wkhd"><colgroup><col style="width:34px"><col style="width:380px"><col style="width:34px"><col style="width:230px"><col style="width:34px"><col style="width:250px"><col></colgroup>
    <tr><td colspan="5" class="wkttl">週　　間　　工　　程　　表</td><td class="wkmade">${reiwa(todayKey(), true)}作成</td><td class="wkco" rowspan="2">${esc(COMPANY)}</td></tr>
    <tr><td class="vt">工事名</td><td class="wkname"><div>${esc(officialOf(s) || nm(s))}</div><div>工事略称　　${esc(nm(s))}</div></td>
      <td class="vt">工期</td><td class="wkki"><div>自　　${s.工期自 ? reiwa(s.工期自) : ""}</div><div>至　　${s.工期至 ? reiwa(s.工期至) : ""}</div></td>
      <td class="vt">住所</td><td class="wkaddr">${esc(s.住所 || "")}</td></tr>
    <tr><td colspan="6"></td><td class="wktanto">担当：　${esc(s.担当 || "")}</td></tr>
  </table>
  <table class="wk"><colgroup><col style="width:180px">${days.map(() => `<col>`).join("")}<col style="width:${mapW}px"></colgroup>
    <tr style="height:26px"><td rowspan="2" class="wkym">令和${+mon.slice(0, 4) - 2018}年${+mon.slice(5, 7)}月</td>
      <td colspan="7" class="wkw">今　　　　週</td><td colspan="7" class="wkw">来　　　　週</td><td colspan="7" class="wkw">再　来　週</td>
      <td rowspan="${totalRows}" class="wkmap">${map ? `<img src="${map}" alt="">` : `<div class="nomap">地図</div>`}</td></tr>
    <tr style="height:22px">${months.map(x => `<td colspan="${x.n}" class="wkm">${x.m}　月</td>`).join("")}</tr>
    <tr style="height:30px"><td rowspan="2" class="wkdiag"></td>${days.map(k => `<td class="wkd ${numCls(k)}">${+k.slice(8)}</td>`).join("")}</tr>
    <tr style="height:26px">${days.map(k => `<td class="wkd ${numCls(k)}">${isRest(k) ? "休" : WD[dayDow(k)]}</td>`).join("")}</tr>
    ${rows.map(r => `<tr style="height:${rowH}px"><td class="wkl">${esc(r.name)}</td>${days.map(k => `<td class="${offCls(k)}">${r.on.has(k) ? '<div class="wkbar"></div>' : ""}</td>`).join("")}</tr>`).join("")}
    <tr style="height:${workerH}px"><td class="wkl">作業員</td>${days.map((k, i) => `<td class="${offCls(k)} wkn">${workers[i] ? workers[i] + "名" : ""}</td>`).join("")}</tr>
    <tr style="height:${noteH}px"><td class="wkl wknote">記　　　　事</td>${notes.map(n => `<td colspan="7" class="wkmemo">${n}</td>`).join("")}</tr>
  </table></div>`;
}
function preparePrint() {
  setPageStyle();
  const size = $("inPaper").value || "A3";
  const [w, h] = PAPER[size];
  root.setProperty("--z", 1); root.setProperty("--lz", 1);
  const W = ((w - PAGE_MM * 2) * 96 / 25.4) * 0.985;          // 用紙の中の幅（px）。少し控える
  const H = ((h - PAGE_MM * 2) * 96 / 25.4 - 30) * 0.97;      // 上の題の行の分を引き、少し控える（はみ出して2ページになるのを防ぐ）
  const view = range();
  const P = S.print || { from: view.from, end: view.end, sites: shownSites().map(s => s.id), nx: 1, ny: 1 };
  const tabName = { chart: "全体工程表", flow: "人の流れ", list: "工程の一覧" }[S.tab];
  const out = $("printPages");
  out.innerHTML = "";
  if (P.fmt === "week") {
    const z = Math.min(W / WK_W, ((h - PAGE_MM * 2) * 96 / 25.4) * 0.97 / WK_H);
    for (const s of shownSitesOf(P.sites)) {
      const div = document.createElement("div");
      div.className = "ppage";
      div.innerHTML = weekPage(s, monday(P.from));
      div.firstElementChild.style.zoom = z.toFixed(4);
      out.appendChild(div);
    }
  } else if (S.tab === "list") {
    // 一覧は横に合わせて縮めるだけ（縦は続けて次のページへ）
    S.print = P; renderList();
    const t = $("list").querySelector("table").cloneNode(true);
    t.style.zoom = Math.min(1.4, W / $("list").querySelector("table").scrollWidth);
    out.innerHTML = `<div class="ptitle"><b>${tabName}</b>　${kd(P.from, true)} 〜 ${kd(P.end, P.end.slice(0, 4) !== P.from.slice(0, 4))}</div>`;
    out.appendChild(t);
  } else {
    const chunks = dayChunks(P.from, P.end, P.nx || 1);
    const groups = S.tab === "chart" ? siteGroups(P.sites, P.ny || 1) : [P.sites];
    const pages = [];
    for (const g of groups) for (const c of chunks) pages.push({ g, c });
    pages.forEach((pg, i) => {
      S.print = { from: pg.c.a, end: pg.c.z, sites: pg.g };
      const div = document.createElement("div");
      div.className = "ppage";
      div.innerHTML = `<div class="ptitle"><b>${tabName}</b>　${kd(pg.c.a, true)} 〜 ${kd(pg.c.z, pg.c.z.slice(0, 4) !== pg.c.a.slice(0, 4))}${pages.length > 1 ? `<span>${i + 1} / ${pages.length}</span>` : ""}</div>`;
      let el, z;
      if (S.tab === "chart") {
        z = fitChart(W, H);
        el = $("chart").cloneNode(true);
        el.removeAttribute("id"); el.classList.add("pchart");
        const b = el.querySelector("#body"); if (b) { b.removeAttribute("id"); b.classList.add("cbody"); }
        el.querySelectorAll(".addsite, .lane.donerow").forEach(x => x.remove());
      } else {
        z = fitFlow(W, H);
        el = $("flow").querySelector("table").cloneNode(true);
      }
      // 写しにはそのページの幅・高さ・倍率を持たせる（:root の値はページごとに変わるため）
      ["--dw", "--rh", "--fw", "--fh"].forEach(k => { const v = root.getPropertyValue(k); if (v) el.style.setProperty(k, v); });
      el.style.zoom = z.toFixed(4);
      div.appendChild(el);
      out.appendChild(div);
    });
  }
  S.print = P;
  ["--dw", "--rh", "--fw", "--fh"].forEach(k => root.removeProperty(k));
  document.body.classList.add("printing");
}
function afterPrint() {
  document.body.classList.remove("printing");
  $("printPages").innerHTML = "";
  S.print = null;
  ["--dw", "--rh", "--fw", "--fh"].forEach(k => root.removeProperty(k));
  applyZoom(); render();
}
window.addEventListener("beforeprint", preparePrint);
window.addEventListener("afterprint", afterPrint);
/* 印刷の小窓（2026-10-13 ユーザー：印刷する現場と期間を決めてから印刷したい／2ページ以上に分けるのは自分で決めたい）
   期間（はじめは画面の期間）・現場（チェック。はじめは画面に出ている現場）・用紙・ページの分け方（期間を n ページ・現場を n ページ）を選んで「印刷する」 */
const PRINT_OFF_KEY = "kotei_printOff";   // 前に外した現場（その人のブラウザに覚える）
$("bPrint").onclick = () => {
  const { from, end } = range();
  $("pFrom").value = from; $("pTo").value = end;
  let off = [];
  try { off = JSON.parse(localStorage.getItem(PRINT_OFF_KEY)) || []; } catch (e) { }
  const all = sortSites(sites());
  $("pSites").innerHTML = all.map(s => `<label class="${s.完了 ? "done" : ""}"><input type="checkbox" value="${esc(s.id)}" ${(showDone || !s.完了) && !off.includes(s.id) ? "checked" : ""}>` +
    `${s.完了 ? '<span class="donemark">完了</span>' : ""}${esc(nm(s))}${s.コード ? ` <span class="note">${esc(s.コード)}</span>` : ""}</label>`).join("");
  $("pTab").textContent = { chart: "工程表", flow: "人の流れ", list: "一覧" }[S.tab];
  $("pNy").disabled = S.tab !== "chart"; $("pNx").disabled = S.tab === "list";
  snapMonday(); fmtView();
  printCount();
  openDlg($("dPrint"));
};
const isWeek = () => $("pFmt").value === "week";
function fmtView() {
  document.querySelectorAll("#dPrint .fmt-chart").forEach(el => { el.hidden = isWeek(); });
  document.querySelectorAll("#dPrint .fmt-week").forEach(el => { el.hidden = !isWeek(); });
}
// 週間工程表は必ず月曜はじまり（2026-10-13 ユーザー指定）：開始日を月曜に寄せる
function snapMonday() { if (isWeek() && $("pFrom").value) $("pFrom").value = monday($("pFrom").value); }
$("pFrom").addEventListener("change", () => { snapMonday(); printCount(); });
$("pFmt").addEventListener("change", () => { snapMonday(); fmtView(); printCount(); try { localStorage.setItem("kotei_printFmt", $("pFmt").value); } catch (x) { } });
function printCount() {
  const n = $("pSites").querySelectorAll("input:checked").length, a = $("pFrom").value, z = $("pTo").value;
  if (isWeek()) {
    const m = a ? monday(a) : "";
    $("pSum").textContent = `現場 ${n} 件・${m ? md(m) + "（月）〜 " + md(dayAdd(m, 20)) + "（日）の3週間" : "開始日を入れてください"}・${n} ページ`;
    $("bPrintGo").disabled = $("bPrintXlsx").disabled = !n || !a;
    return;
  }
  const days = a && z && z >= a ? dayDiff(a, z) + 1 : 0;
  const pages = S.tab === "list" ? 0 : (+$("pNx").value || 1) * (S.tab === "chart" ? Math.min(n || 1, +$("pNy").value || 1) : 1);
  $("pSum").textContent = `現場 ${n} 件・${days ? days + " 日間" : "期間が正しくありません"}` + (pages ? `・${pages} ページ` : "");
  $("bPrintGo").disabled = !n || !days || days > 366;
  if (days > 366) $("pSum").textContent += "（1年までにしてください）";
}
["pFrom", "pTo", "pNx", "pNy"].forEach(id => $(id).addEventListener("change", printCount));
$("pSites").addEventListener("change", printCount);
$("pAll").onclick = () => { $("pSites").querySelectorAll("input").forEach(c => { c.checked = true; }); printCount(); };
$("pNone").onclick = () => { $("pSites").querySelectorAll("input").forEach(c => { c.checked = false; }); printCount(); };
$("pView").onclick = () => { const v = fromKey(S.view.from); $("pFrom").value = v; $("pTo").value = dayAdd(addMonths(v, S.view.months), -1); printCount(); };
document.querySelectorAll("#dPrint [data-pm]").forEach(b => b.onclick = () => {
  const a = $("pFrom").value || fromKey(S.view.from);
  $("pTo").value = dayAdd(addMonths(a, +b.dataset.pm), -1); printCount();
});
$("bPrintCancel").onclick = () => $("dPrint").close();
$("fPrint").addEventListener("submit", async e => {
  e.preventDefault();
  const boxes = [...$("pSites").querySelectorAll("input")];
  const ids = boxes.filter(c => c.checked).map(c => c.value);
  try {
    localStorage.setItem(PRINT_OFF_KEY, JSON.stringify(boxes.filter(c => !c.checked).map(c => c.value)));
    localStorage.setItem("kotei_printSplit", JSON.stringify([$("pNx").value, $("pNy").value]));
  } catch (x) { }
  $("dPrint").close();
  setPageStyle();
  S.print = { from: $("pFrom").value, end: $("pTo").value, sites: ids, nx: +$("pNx").value || 1, ny: +$("pNy").value || 1, fmt: $("pFmt").value };
  if (S.print.fmt === "week") {
    // 地図は先に読んでおく（印刷の直前には読み込みを待てない）
    toast("地図を読んでいます…", 1500);
    const urls = await Promise.all(shownSitesOf(ids).map(loadMap));
    await Promise.all(urls.filter(Boolean).map(loadImg));
  }
  // 小窓が閉じきってから印刷画面を出す
  setTimeout(() => window.print(), 50);
});
try {
  $("inPaper").value = localStorage.getItem("kotei_paper") || "A3";
  const sp = JSON.parse(localStorage.getItem("kotei_printSplit")) || [];
  if (sp[0]) $("pNx").value = sp[0];
  if (sp[1]) $("pNy").value = sp[1];
  $("pFmt").value = localStorage.getItem("kotei_printFmt") || "chart";
} catch (e) { }
$("inPaper").onchange = e => { try { localStorage.setItem("kotei_paper", e.target.value); } catch (x) { } setPageStyle(); };
setPageStyle();

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
  S.view.from = fromKey(S.view.from) || todayKey().slice(0, 8) + "01";
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
