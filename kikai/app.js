/* 自社機械稼働表（2026-10-23 試作 Vr000 → 2026-10-24 Vr001：全体工程表と同じく OneDrive で共有）
   自社の機械の予約・稼働を、機械×日の表（昼／夜の2行）で持ち、遊んでいる日が見えるようにする。
   画面は GitHub Pages（kkodin/kotei の kikai\）、データは会社の OneDrive（下の「保存先」）。
   予約は全体工程表のバーと同じ形（開始・終了の文字列＋使う曜日＋その日だけの休み）にして、あとでつなげやすくする。 */
"use strict";

const LS_DATA = "kikai_v0_data", LS_VIEW = "kikai_v0_view";
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const lsGet = k => { try { return localStorage.getItem(k); } catch { return null; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, v); return true; } catch { return false; } };

/* ---------- 日付（"YYYY-MM-DD" の文字列で持つ） ---------- */
const pad2 = n => String(n).padStart(2, "0");
const keyOf = d => d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
const dateOf = k => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
const dayAdd = (k, n) => { const d = dateOf(k); d.setDate(d.getDate() + n); return keyOf(d); };
const dayDow = k => dateOf(k).getDay();
const dayDiff = (a, b) => Math.round((dateOf(b) - dateOf(a)) / 86400000);
const todayKey = () => keyOf(new Date());
const md = k => { const d = dateOf(k); return (d.getMonth() + 1) + "/" + d.getDate(); };
const WD = "日月火水木金土";
const addMonths = (k, n) => { const d = dateOf(k), day = d.getDate(); d.setDate(1); d.setMonth(d.getMonth() + n);
  const last = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate(); d.setDate(Math.min(day, last)); return keyOf(d); };

/* 日本の祝日（全体工程表・原価管理と同じ決まり） */
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
const isSunHol = k => dayDow(k) === 0 || !!holidayOf(k);
const isSat = k => dayDow(k) === 6 && !holidayOf(k);
const isWeekday = k => !isSunHol(k) && !isSat(k);          // 遊び・稼働率を数える日（月〜金・祝日を除く）
const dayClass = k => holidayOf(k) ? "hol" : dayDow(k) === 0 ? "sun" : dayDow(k) === 6 ? "sat" : "";

/* 全角の数字・記号を半角に（数字の欄。IME が入っていても受ける） */
const half = s => String(s ?? "").replace(/[０-９．，－＋]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/[,，\s]/g, "");
const num = s => { const v = parseFloat(half(s)); return isFinite(v) ? v : null; };
const yen = v => v == null ? "" : Math.round(v).toLocaleString("ja-JP");

/* ---------- データ ----------
   機械：{ id, 名前, 規格, 分類, 単価昼, 単価夜, メモ, 使わない, 台数(在庫。無ければ1), 単価品(単価表の品の key。無ければ自動) }
     分類・単価は単価表の品に対応していればそちらを使う（下の「単価表」）
   現場：{ 名前, 種類(現場／置場／修理／検査), 色, コード, 完了 }  … 名前＝全体工程表の「社内工事名」（略称）でそろえる
   予約：{ id, 機械, 区分(昼／夜), 現場(名前), 開始, 終了, 土, 日祝, 休み:[日], 出勤:[日], 状態(確定／仮), メモ, 更新, 交渉, 台数, 料金 }
     台数：その予約で使う台数（無ければ1）。料金：{ 方式(日／月／現場／基／m／枚／式), 数量, 金額(一式) }。無ければ単価表どおり
     交渉：予算の無い現場で、機材センター長と決めた特別単価（2026-10-27）。無ければ null
       { 結果(決まった／交渉中), 単価(円／日), 日, 相手, 方法(対面／電話／メール／その他), メモ, 記録(入れた時刻と人) } */
const KINDS = ["現場", "置場", "修理", "検査"];
const KIND_COLOR = { 置場: "#9aa3ad", 修理: "#b23b3b", 検査: "#7b5ea7" };
const PALETTE = ["#3d7cc9", "#3f9a45", "#ef7d1a", "#d6457a", "#16a3a3", "#8d6e63", "#c9a400", "#5c6bc0",
  "#2e7d32", "#e65100", "#6d4c41", "#00838f", "#ad1457", "#558b2f", "#4527a0", "#f9a825"];

const emptyData = () => ({ 形式: "自社機械稼働表", 版: 1, 機械: [], 現場: [
  { 名前: "機材センター", 種類: "置場", 色: KIND_COLOR.置場, コード: "", 完了: false },
  { 名前: "故障又は修理中", 種類: "修理", 色: KIND_COLOR.修理, コード: "", 完了: false },
  { 名前: "特定自主検査", 種類: "検査", 色: KIND_COLOR.検査, コード: "", 完了: false }], 予約: [] });

function normalize(d) {
  if (!d || d.形式 !== "自社機械稼働表") throw new Error("自社機械稼働表のデータではありません");
  d.機械 = d.機械 || []; d.現場 = d.現場 || []; d.予約 = d.予約 || [];
  d.予約.forEach(r => { r.休み = r.休み || []; r.出勤 = r.出勤 || []; r.未使用 = r.未使用 || []; r.状態 = r.状態 || "確定"; });
  return d;
}
/* ---------- 保存先 ----------
   公開した画面：会社の OneDrive「03_自社機械稼働表\稼働データ」フォルダの 稼働データ.json（Microsoft Graph。全体工程表と同じアプリ登録）。
     全員が同じファイルを読み書きする。同時に書いたときは eTag で止め、最新を読み直して自分の変更（関数）をかけ直して保存する（全体工程表の saveFile と同じ）。
   ?demo を付けたとき・ファイルを直接開いたとき（file://）：このブラウザの中だけ（localStorage）。 */
const CFG = window.KIKAI_CONFIG || {};
const DEMO = new URLSearchParams(location.search).has("demo") || location.protocol === "file:";
const SCOPES = ["User.Read", "Files.ReadWrite.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const FILE = "稼働データ.json";
const POLL_MS = 30000;
let msalApp = null, me = { name: "試し" };

async function token() {
  try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken; }
  catch (e) { if (e instanceof msal.InteractionRequiredAuthError) await msalApp.acquireTokenRedirect({ scopes: SCOPES }); throw e; }
}
async function g(path, opt = {}) { return fetch(GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + await token(), ...(opt.headers || {}) } }); }
function httpErr(res, what) { const e = new Error(what + " " + res.status); e.status = res.status; return e; }

let folderBase = "";
async function resolveFolder() {
  const f = CFG.folder || {};
  if (f.itemId) { folderBase = `/drives/${f.driveId}/items/${f.itemId}`; return; }
  /* itemId をまだ config.js に入れていないとき（はじめだけ）：持ち主の OneDrive の道筋で探す（持ち主しか通らない）。
     見つけた ID は「データ」の小窓に出す → config.js に貼って公開し直す */
  const res = await g(`/drives/${f.driveId}/root:/${f.path.split("/").map(encodeURIComponent).join("/")}`);
  if (!res.ok) throw httpErr(res, "稼働データ フォルダ");
  const j = await res.json();
  window.FOLDER_ITEM_ID = j.id;
  folderBase = `/drives/${f.driveId}/items/${j.id}`;
}
const GraphStore = {
  p: () => folderBase + ":/" + encodeURIComponent(FILE),
  async eTag() { const res = await g(this.p() + "?$select=eTag"); if (res.status === 404) return null; if (!res.ok) throw httpErr(res, "確認"); return (await res.json()).eTag; },
  async get() {
    const res = await g(this.p());
    if (res.status === 404) return null;
    if (!res.ok) throw httpErr(res, "読み込み");
    const meta = await res.json(), dl = meta["@microsoft.graph.downloadUrl"];
    const body = dl ? await fetch(dl, { cache: "no-store" }) : await g(this.p() + ":/content");
    if (!body.ok) throw httpErr(body, "中身の読み込み");
    return { data: await body.json(), eTag: meta.eTag };
  },
  // eTag があれば If-Match（違えば 412）。無ければ新規（あれば 409）
  async put(data, eTag) {
    const res = await g(this.p() + ":/content" + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail"),
      { method: "PUT", headers: { "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: JSON.stringify(data) });
    if (!res.ok) throw httpErr(res, "保存");
    return { eTag: (await res.json()).eTag };
  }
};
const LocalStore = {
  async eTag() { return lsGet(LS_DATA + "_tag"); },
  async get() { const t = lsGet(LS_DATA); return t ? { data: JSON.parse(t), eTag: lsGet(LS_DATA + "_tag") } : null; },
  async put(data, eTag) {
    if (lsGet(LS_DATA) && (lsGet(LS_DATA + "_tag") || null) !== (eTag || null)) { const e = new Error("保存 412"); e.status = 412; throw e; }
    const t = "t" + Date.now();
    if (!lsSet(LS_DATA, JSON.stringify(data)) || !lsSet(LS_DATA + "_tag", t)) throw new Error("このブラウザに保存できませんでした（容量・設定）");
    return { eTag: t };
  }
};
const store = DEMO ? LocalStore : GraphStore;

let D = emptyData(), docTag = null, pending = [], saving = false, saveErr = null, saveT = 0, ready = false;

/* 変更は「関数」で持つ（commit）。保存がぶつかったら、読み直した最新の中身に同じ関数をもう一度かける。
   元に戻す・やり直すは、前後で変わった機械・現場・予約だけを戻す関数にする（ほかの人がその間に直した所は消さない） */
const hist = [], fut = [];
function commit(fn, opt = {}) {
  const before = JSON.parse(JSON.stringify(D));
  fn(D);
  if (!opt.noUndo) {
    const after = JSON.parse(JSON.stringify(D));
    hist.push({ undo: patchTo(before, after), redo: patchTo(after, before) }); if (hist.length > 60) hist.shift(); fut.length = 0;
  }
  pending.push(fn); D.更新 = { だれ: me.name, いつ: new Date().toLocaleString("ja-JP") };
  scheduleSave(); renderAll();
}
/* to の中身に戻す関数（from と to で違う所だけ） */
function patchTo(to, from) {
  const keys = { 機械: x => x.id, 予約: x => x.id, 現場: x => x.名前 };
  const ops = [];
  Object.keys(keys).forEach(c => {
    const k = keys[c], A = new Map(from[c].map(x => [k(x), JSON.stringify(x)])), B = new Map(to[c].map(x => [k(x), x]));
    B.forEach((x, id) => { if (A.get(id) !== JSON.stringify(x)) ops.push({ c, id, x: JSON.parse(JSON.stringify(x)) }); });
    A.forEach((_, id) => { if (!B.has(id)) ops.push({ c, id, del: true }); });
  });
  const order = to.機械.map(m => m.id), orderChanged = order.join() !== from.機械.map(m => m.id).join();
  return d => {
    ops.forEach(o => {
      const i = d[o.c].findIndex(x => keys[o.c](x) === o.id);
      if (o.del) { if (i >= 0) d[o.c].splice(i, 1); }
      else if (i >= 0) d[o.c][i] = JSON.parse(JSON.stringify(o.x)); else d[o.c].push(JSON.parse(JSON.stringify(o.x)));
    });
    if (orderChanged) d.機械.sort((a, b) => (order.indexOf(a.id) + 1 || 1e9) - (order.indexOf(b.id) + 1 || 1e9));
  };
}
function undo() { const h = hist.pop(); if (!h) return; fut.push(h); commit(h.undo, { noUndo: true }); }
function redo() { const h = fut.pop(); if (!h) return; hist.push(h); commit(h.redo, { noUndo: true }); }
/* データを丸ごと入れ替える（読み込み・全部消す） */
const replaceAll = data => d => { Object.keys(d).forEach(k => delete d[k]); Object.assign(d, JSON.parse(JSON.stringify(data))); };

function scheduleSave(ms = 700) { clearTimeout(saveT); saveT = setTimeout(saveDoc, ms); showSync(); }
async function saveDoc() {
  if (saving) { scheduleSave(500); return; }
  if (!pending.length) return;
  saving = true; showSync();
  try {
    for (let tries = 0; tries < 5; tries++) {
      const n = pending.length;
      try { const r = await store.put(D, docTag); docTag = r.eTag; pending.splice(0, n); saveErr = null; break; }
      catch (e) {
        if (e.status !== 412 && e.status !== 409) throw e;
        // ほかの人が先に保存した → 最新を読んで、自分の変更を重ね直す
        const fresh = await store.get();
        const d = fresh ? normalize(fresh.data) : D;
        if (fresh) pending.forEach(fn => fn(d));
        D = d; docTag = fresh ? fresh.eTag : null;
        toast("ほかの人の変更を読み込んで、重ねて保存しました");
        renderAll();
      }
    }
  } catch (e) {
    saveErr = e.message; console.error(e);
    toast("保存できませんでした：" + e.message + "（自動でやり直します）");
    setTimeout(() => scheduleSave(0), 10000);
  } finally {
    saving = false; showSync();
    if (pending.length && !saveErr) scheduleSave(300);
  }
}
function showSync() {
  const el = $("sync"); if (!el) return;
  const busy = pending.length || saving;
  el.className = "sync " + (saveErr ? "err" : busy ? "busy" : "ok");
  el.textContent = saveErr ? "保存できていません" : busy ? "保存中…" : DEMO ? "このブラウザに保存済み" : "保存済み";
}
/* ほかの人の書き込みを取り込む（自分の保存待ちがあるときは、保存のときに重ねるので飛ばす） */
let pulling = false;
async function pull() {
  if (pulling || !ready || pending.length || saving) return;
  pulling = true;
  try {
    const t = await store.eTag();
    if (t && t !== docTag) {
      const r = await store.get();
      if (r && !pending.length && !saving) { D = normalize(r.data); docTag = r.eTag; renderAll(); }
    }
  } catch (e) {
    console.error(e);
    if (e.status === 404 || e.status === 403) toast("稼働データ フォルダを開けません。北澤さんに共有（編集可）を頼んでください");
    else toast("最新の読み込みに失敗：" + e.message);
  } finally { pulling = false; }
}
/* 単価表.json（稼働データ と同じフォルダ。tools\単価表を読む.py が NAS から作る）を読む。変わったときだけ描き直す */
let prTag = null;
async function loadPrices() {
  if (DEMO) return;
  try {
    const res = await g(folderBase + ":/" + encodeURIComponent("単価表.json"));
    if (!res.ok) return;
    const meta = await res.json();
    if (meta.eTag === prTag) return;
    const body = await fetch(meta["@microsoft.graph.downloadUrl"], { cache: "no-store" });
    if (!body.ok) return;
    setPrices(await body.json()); prTag = meta.eTag;
    if (ready) renderAll();
  } catch (e) { console.warn("単価表を読めませんでした", e); }
}
let toastT = 0;
function toast(msg) { const t = $("toast"); t.textContent = msg; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => { t.hidden = true; }, 5000); }

const newId = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const machine = id => D.機械.find(m => m.id === id);
const siteOf = name => D.現場.find(s => s.名前 === name) || { 名前: name, 種類: "現場", 色: "#8a96a3" };
const isReal = name => siteOf(name).種類 === "現場";
const isLightColor = hex => { const v = parseInt(String(hex).slice(1), 16); const r = v >> 16, g = (v >> 8) & 255, b = v & 255; return (r * 299 + g * 587 + b * 114) / 1000 > 165; };
function ensureSite(d, name) {
  if (!name || d.現場.some(s => s.名前 === name)) return;
  const used = d.現場.filter(s => s.種類 === "現場").length;
  d.現場.push({ 名前: name, 種類: "現場", 色: PALETTE[used % PALETTE.length], コード: "", 完了: false });
}
const mLabel = m => m ? m.名前 + (m.規格 ? " " + m.規格 : "") : "（消した機械）";

/* その予約がその日に機械を使うか（使う曜日・その日だけの休み／出勤） */
/* planned＝予定で使う日。works＝実際に使う（使った）日＝予定から「使わなかった」と申告した日（未使用）を除く */
function planned(r, k) {
  if (k < r.開始 || k > r.終了) return false;
  if (r.出勤.includes(k)) return true;
  if (r.休み.includes(k)) return false;
  if (isSunHol(k)) return !!r.日祝;
  if (isSat(k)) return !!r.土;
  return true;
}
const works = (r, k) => planned(r, k) && !r.未使用.includes(k);
function workDays(r, from = r.開始, to = r.終了) {
  let n = 0; for (let k = from < r.開始 ? r.開始 : from; k <= to && k <= r.終了; k = dayAdd(k, 1)) if (works(r, k)) n++;
  return n;
}
/* ---------- 単価表（2026-10-27） ----------
   NAS の実行予算単価「機材センター」シートを tools\単価表を読む.py が OneDrive の 稼働データ\単価表.json にする（画面は NAS を読めない）。
   品：{ key, 種別, 名称, 能力, 備考, 単位(施工の単位：現場／基／枚／m), 履歴:[{ から(その日から。""＝はじめから), 日, 施工, 月 }], 無し? }
   対応：{ "名前|規格": 品の key }（道具が名前で決める）。機械の 単価品 があればそちらが先（""＝単価表を使わない）。
   金額は予約の日（月極・施工・一式は開始日）に有効な単価で出す。単価表に無い機械は機械の 単価昼・単価夜。 */
let PR = null, PRI = {};
function setPrices(p) { PR = p || null; PRI = {}; (PR?.品 || []).forEach(i => { PRI[i.key] = i; }); }
setPrices(window.KIKAI_PRICES);
const mKey = m => m.名前 + "|" + (m.規格 || "");
const itemKey = m => m.単価品 !== undefined ? m.単価品 : PR?.対応?.[mKey(m)] || "";
const itemOf = m => m && PRI[itemKey(m)] || null;
const itemLabel = i => i ? i.名称 + (i.能力 ? " " + i.能力 : "") : "";
function priceAt(i, f, k) { let v = null; for (const h of i.履歴) { if (!h.から || h.から <= k) v = h[f]; else break; } return v; }
/* 分類：単価表の品に対応していればその種別、無ければ機械の 分類 */
const OLD_CATS = ["車両", "重機", "小型機械", "測量", "その他"];
const catOf = m => itemOf(m)?.種別 || m.分類 || "その他";
const catList = () => [...new Set([...(PR?.種別 || []), ...D.機械.map(catOf), ...OLD_CATS])];
const catIndex = () => { const L = catList(), o = {}; L.forEach((c, i) => { o[c] = i; }); return c => o[c] ?? 999; };

/* 料金の決め方（予約の 料金＝{ 方式, 数量, 金額 }。無ければ単価表の品から：日の単価があれば日極、無ければ施工の単位） */
const METHODS = { 日: "日極", 月: "月極", 現場: "1現場", 基: "立坑1基", m: "1m", 枚: "1枚", 式: "一式（金額）" };
const UNIT = { 日: "日", 月: "月", 現場: "現場", 基: "基", m: "m", 枚: "枚", 式: "式" };
function methodOf(r, i = itemOf(machine(r.機械))) {
  if (r.料金?.方式) return r.料金.方式;
  const h = i?.履歴[i.履歴.length - 1];
  return h && h.日 == null && h.施工 != null && i.単位 ? i.単位 : "日";
}
const basePrice = (m, kb) => m ? (kb === "夜" ? m.単価夜 : m.単価昼) : null;
const special = r => r.交渉 && r.交渉.結果 === "決まった" && r.交渉.単価 != null;
/* その日の単価（交渉で決まった特別単価が先。noSp＝通常の単価） */
function unitPrice(r, m, k, meth = methodOf(r), noSp = false) {
  if (meth === "式") return null;
  if (!noSp && special(r)) return r.交渉.単価;
  const i = itemOf(m), v = i ? priceAt(i, meth === "日" ? "日" : meth === "月" ? "月" : "施工", k) : null;
  return v == null && meth === "日" ? basePrice(m, r.区分) : v;
}
const monthsOf = r => Math.max(1, Math.ceil((dayDiff(r.開始, r.終了) + 1) / 30));
const qtyOf = (r, meth) => r.料金?.数量 ?? (meth === "月" ? monthsOf(r) : 1);
/* 期間 from〜to の金額。日極は使った日ごと、月極・施工・一式は開始日がその期間にあるときに全部 */
function costOf(r, from = r.開始, to = r.終了, noSp = false, m = machine(r.機械)) {
  const meth = methodOf(r, itemOf(m)), n = r.台数 || 1, days = workDays(r, from, to);
  if (meth === "日") {
    let amt = null;
    for (let k = from < r.開始 ? r.開始 : from; k <= to && k <= r.終了; k = dayAdd(k, 1))
      if (works(r, k)) { const p = unitPrice(r, m, k, meth, noSp); if (p != null) amt = (amt || 0) + p * n; }
    return { meth, n, days, qty: days, unit: unitPrice(r, m, from < r.開始 ? r.開始 : from, meth, noSp), amt };
  }
  const here = r.開始 >= from && r.開始 <= to;
  if (meth === "式") return { meth, n, days, qty: 1, unit: null, amt: here ? r.料金?.金額 ?? null : 0 };
  const unit = unitPrice(r, m, r.開始, meth, noSp), qty = qtyOf(r, meth);
  return { meth, n, days, qty, unit, amt: !here ? 0 : unit == null ? null : unit * qty * n };
}
/* 「8,000円×3日×2台」 */
const costText = c => c.meth === "式" ? "一式" : c.unit == null ? "単価なし" :
  `${yen(c.unit)}円×${c.qty}${UNIT[c.meth]}${c.n > 1 ? "×" + c.n + "台" : ""}`;
const priceOf = (r, m = machine(r.機械)) => unitPrice(r, m, r.開始);
const overlap = (a, b) => a.開始 <= b.終了 && b.開始 <= a.終了;
/* 同じ機械・同じ区分で、同じ日に使う予約（二重の予約） */
function conflictDays(a, b) {
  if (a.id === b.id || a.機械 !== b.機械 || a.区分 !== b.区分 || !overlap(a, b)) return [];
  const out = [], s = a.開始 > b.開始 ? a.開始 : b.開始, e = a.終了 < b.終了 ? a.終了 : b.終了;
  for (let k = s; k <= e; k = dayAdd(k, 1)) if (works(a, k) && works(b, k)) out.push(k);
  return out;
}
/* 在庫の台数を超える日がある予約（1台の機械なら二重の予約） */
const capOf = id => machine(id)?.台数 || 1;
function conflictSet(list) {
  const bad = new Set(), by = {};
  list.forEach(r => (by[r.機械 + "|" + r.区分] ||= []).push(r));
  Object.values(by).forEach(a => {
    const cap = capOf(a[0].機械);
    if (a.length < 2 && (a[0].台数 || 1) <= cap) return;
    const use = {};
    a.forEach(r => { for (let k = r.開始; k <= r.終了; k = dayAdd(k, 1)) if (works(r, k)) (use[k] ||= []).push(r); });
    Object.values(use).forEach(rs => { if (rs.reduce((s, r) => s + (r.台数 || 1), 0) > cap) rs.forEach(r => bad.add(r.id)); });
  });
  return bad;
}
/* 日付の並びを「10/5〜10/9、10/12」の形に */
function spanText(keys) {
  const out = []; let s = null, p = null;
  keys.sort().forEach(k => { if (p && dayAdd(p, 1) === k) { p = k; return; } if (s) out.push(s === p ? md(s) : md(s) + "〜" + md(p)); s = p = k; });
  if (s) out.push(s === p ? md(s) : md(s) + "〜" + md(p));
  return out.join("、");
}

/* ---------- 画面の状態（その人のブラウザに覚える） ---------- */
const S = Object.assign({ tab: "chart", from: "", months: 2, cat: "", q: "", unused: false, paint: "", idle: true, focus: "", sumBy: "machine", sumSort: "order", zoom: 1, showDone: false },
  (() => { try { return JSON.parse(lsGet(LS_VIEW)) || {}; } catch { return {}; } })());
S.sel = null;
if (!/^\d{4}-\d\d-\d\d$/.test(S.from)) S.from = todayKey().slice(0, 8) + "01";
const saveView = () => { const { sel, ...v } = S; lsSet(LS_VIEW, JSON.stringify(v)); };
const rangeEnd = () => dayAdd(addMonths(S.from, S.months), -1);
function rangeDays() { const out = [], e = rangeEnd(); for (let k = S.from; k <= e; k = dayAdd(k, 1)) out.push(k); return out; }
function shownMachines() {
  const q = S.q.trim().toLowerCase();
  const ci = catIndex();   // 分類（単価表の種別）の順にまとめる。分類の中は機械の並び
  return D.機械.filter(m => (S.unused || !m.使わない) && (!S.cat || catOf(m) === S.cat) &&
    (!q || (m.名前 + " " + (m.規格 || "")).toLowerCase().includes(q)))
    .map((m, i) => [m, ci(catOf(m)), i]).sort((a, b) => a[1] - b[1] || a[2] - b[2]).map(x => x[0]);
}
/* 1日の幅：ふだん 24px。表が画面の幅より狭いとき（1か月表示など）は、画面いっぱいになるまで広げる（renderChart で決める） */
let curDW = 24;
const DW = () => curDW;
const cssPx = n => parseFloat(getComputedStyle(document.documentElement).getPropertyValue(n)) || 0;
/* 表の拡大・縮小（CSS の zoom を表だけにかける。上の帯・小窓は大きさが変わらない） */
const ZMIN = 0.5, ZMAX = 2;
function setZoom(z) {
  S.zoom = Math.min(ZMAX, Math.max(ZMIN, Math.round(z * 10) / 10)); saveView();
  const w = $("wrap"), cx = (w.scrollLeft + w.clientWidth / 2) / (w.scrollWidth || 1), cy = (w.scrollTop + w.clientHeight / 2) / (w.scrollHeight || 1);
  renderChart();
  w.scrollLeft = cx * w.scrollWidth - w.clientWidth / 2; w.scrollTop = cy * w.scrollHeight - w.clientHeight / 2;
}
const RH = () => parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--rh")) || 22;

/* ---------- 稼働表 ---------- */
function laneRows(list) {
  /* 同じ行で期間が重なる予約は下の行へ（二重の予約が見えるように） */
  const rows = [], at = {};
  list.slice().sort((a, b) => a.開始 < b.開始 ? -1 : a.開始 > b.開始 ? 1 : a.終了 < b.終了 ? 1 : -1).forEach(r => {
    let i = 0; while (rows[i] && rows[i] >= r.開始) i++;
    rows[i] = r.終了; at[r.id] = i;
  });
  return { n: Math.max(1, rows.length), at };
}

function renderChart() {
  const el = $("chart");
  if (!D.機械.length) {
    el.innerHTML = `<div class="empty">まだ機械がありません。<br>
      ${window.KIKAI_IMPORT ? `<button id="eImport">Excel から取り込んだデータを読む</button>` : `<button id="eImport">ファイルから読み込む（取り込み_自社機械予定表.json）</button>`}　<button class="sub" id="eMaster">機械を1台ずつ足す</button></div>`;
    $("eImport").onclick = () => window.KIKAI_IMPORT ? importBundled() : $("ioFile").click();   // 公開した画面には取り込みデータが無い
    $("eMaster").onclick = () => setTab("master");
    $("warn").textContent = ""; return;
  }
  const days = rangeDays(), N = days.length, from = S.from, to = days[N - 1], rh = RH(), today = todayKey();
  const z = S.zoom || 1;
  const avail = $("wrap").clientWidth / z - cssPx("--nw") - cssPx("--kw") - 2;
  const dw = curDW = Math.max(24, Math.floor(avail / N * 10) / 10);
  el.style.setProperty("--dw", dw + "px"); el.style.setProperty("--z", z);
  $("zReset").textContent = Math.round(z * 100) + "%";
  const ms = shownMachines();
  const inRange = D.予約.filter(r => r.開始 <= to && r.終了 >= from);
  const bad = conflictSet(inRange);
  const byLane = {};
  inRange.forEach(r => (byLane[r.機械 + "|" + r.区分] ||= []).push(r));

  const h = [];
  /* 見出し：月・日・曜 */
  const mo = []; days.forEach((k, i) => { if (i === 0 || k.endsWith("-01")) mo.push({ i, k }); });
  h.push(`<div class="hdr"><div class="corner">機械 ／ 昼・夜</div><div class="days">
    <div class="mo">${mo.map((m, j) => `<div style="width:${((mo[j + 1]?.i ?? N) - m.i) * dw}px">${m.k.slice(0, 4)}年${+m.k.slice(5, 7)}月</div>`).join("")}</div>
    <div class="dn">${days.map(k => `<div class="${dayClass(k)} ${k === today ? "today" : ""}" title="${esc(holidayOf(k))}">${+k.slice(8)}</div>`).join("")}</div>
    <div class="dw">${days.map(k => `<div class="${dayClass(k)} ${k === today ? "today" : ""}">${WD[dayDow(k)]}</div>`).join("")}</div>
  </div></div>`);
  /* 土日祝・今日の列 */
  h.push(`<div id="cols" style="left:calc(var(--nw) + var(--kw))">${days.map((k, i) => {
    const c = k === today ? "today" : dayClass(k); return c ? `<div class="${c}" style="left:${i * dw}px"></div>` : ""; }).join("")}</div>`);

  let lastCat = null;
  ms.forEach(m => {
    const cat = catOf(m);
    if (!S.cat && cat !== lastCat) { h.push(`<div class="grp"><div class="gh">${esc(cat)}</div></div>`); lastCat = cat; }
    const lanes = ["昼", "夜"].map(k => { const list = byLane[m.id + "|" + k] || []; return { k, list, ...laneRows(list) }; });
    /* 稼働率（表示の期間の平日のうち、どちらかの区分で現場にいる日） */
    let wd = 0, used = 0; const idle = [];
    const mine = lanes[0].list.concat(lanes[1].list);
    days.forEach(k => {
      if (!isWeekday(k)) return;
      wd++;
      const on = mine.filter(r => works(r, k));
      if (on.some(r => isReal(r.現場))) used++;
      else if (!on.some(r => ["修理", "検査"].includes(siteOf(r.現場).種類))) idle.push(k);
    });
    const rate = wd ? Math.round(used / wd * 100) : 0;
    h.push(`<div class="mach${m.使わない ? " unused" : ""}" data-m="${m.id}"><div class="mh">
      <div class="mname" title="${esc(mLabel(m) + (m.メモ ? "\n" + m.メモ : "") + "\n稼働 " + used + "日／平日 " + wd + "日")}">
        <b>${esc(m.名前)}</b>${(m.台数 || 1) > 1 ? `<span class="cnt">在庫 ${m.台数}台</span>` : ""}<span class="sp2">${esc(m.規格 || "")}${m.メモ ? "　" + esc(m.メモ) : ""}</span>
        <span class="rate${rate < 50 ? " low" : ""}"><i><s style="width:${rate}%"></s></i>${rate}%</span></div>
      <div class="lks">${lanes.map(l => `<div class="lk${l.k === "夜" ? " n" : ""}" style="height:${l.n * rh}px">${l.k}</div>`).join("")}</div>
    </div><div class="lanes">`);
    lanes.forEach(l => {
      h.push(`<div class="cells${l.k === "夜" ? " n" : ""}" data-m="${m.id}" data-k="${l.k}" style="width:${N * dw}px;height:${l.n * rh}px">`);
      if (l.k === "昼" && S.idle) {
        /* 遊んでいる平日（続いた日はまとめて1つ） */
        const set = new Set(idle);
        for (let i = 0; i < N; i++) {
          if (!set.has(days[i])) continue;
          let e = i; while (e + 1 < N && set.has(days[e + 1])) e++;
          h.push(`<div class="idle${days[e] < today ? " past" : ""}" style="left:${i * dw}px;width:${(e - i + 1) * dw}px;height:${l.n * rh - 2}px"></div>`);
          i = e;
        }
      }
      l.list.forEach(r => h.push(barHtml(r, from, N, dw, rh, l.at[r.id], bad.has(r.id))));
      h.push(`</div>`);
    });
    h.push(`</div></div>`);
  });
  el.innerHTML = h.join("");
  el.style.width = `calc(var(--nw) + var(--kw) + ${N * dw}px)`;
  const nBad = bad.size;
  $("warn").textContent = nBad ? `⚠ 二重の予約（台数を超える予約） ${nBad}本` : "";
  $("warn").title = nBad ? "押すと最初の場所へ" : "";
}

function barHtml(r, from, N, dw, rh, row, conf) {
  const s = Math.max(0, dayDiff(from, r.開始)), e = Math.min(N - 1, dayDiff(from, r.終了));
  const st = siteOf(r.現場), color = st.色 || "#8a96a3";
  const gaps = [], g = [];
  for (let i = s; i <= e; i++) {
    const k = dayAdd(from, i);
    if (!planned(r, k)) gaps.push(i);
    else if (r.未使用.includes(k)) g.push(`<i class="nu" style="left:${(i - s) * dw}px;width:${dw}px"></i>`);   // あとから「使わなかった」と申告した日
  }
  let gs = -1;
  for (let j = 0; j <= gaps.length; j++) {
    if (gs < 0 && j < gaps.length) gs = j;
    else if (gs >= 0 && (j === gaps.length || gaps[j] !== gaps[j - 1] + 1)) {
      g.push(`<i class="gap" style="left:${(gaps[gs] - s) * dw}px;width:${(gaps[j - 1] - gaps[gs] + 1) * dw}px"></i>`);
      gs = j < gaps.length ? j : -1;
    }
  }
  const cls = ["bar"];
  if (isLightColor(color)) cls.push("lt");
  if (r.状態 === "仮") cls.push("kari");
  if (conf) cls.push("conf");
  if (S.sel === r.id) cls.push("sel");
  if (S.focus && r.現場 !== S.focus) cls.push("dim");
  return `<div class="${cls.join(" ")}" data-id="${r.id}" style="left:${s * dw}px;width:${(e - s + 1) * dw}px;top:${row * rh + 2}px;background-color:${color}">` +
    g.join("") + `<span class="lbl">${(r.台数 || 1) > 1 ? `<b class="tk n">${r.台数}台</b>` : ""}${r.交渉 ? `<b class="tk${special(r) ? "" : " ing"}">${special(r) ? "特" : "交渉中"}</b>` : ""}${esc(r.現場)}</span><i class="eg l"></i><i class="eg r"></i></div>`;
}
const ngText = x => x ? (x.結果 === "決まった" && x.単価 != null ? `特別単価 ${yen(x.単価)}円／日` : "特別単価 交渉中") +
  `（${[x.日 ? md(x.日) : "", x.相手, x.方法].filter(Boolean).join("・")}）${x.メモ ? "　" + x.メモ : ""}` : "";
const barTip = r => {
  const m = machine(r.機械), n = workDays(r);
  const c = isReal(r.現場) ? costOf(r) : null;
  return `${r.現場}（${siteOf(r.現場).種類}）${r.状態 === "仮" ? "　仮" : ""}\n${mLabel(m)}　${r.区分}${(r.台数 || 1) > 1 ? "　" + r.台数 + "台" : ""}\n${md(r.開始)}〜${md(r.終了)}　使う日 ${n}日` +
    (c ? `\n${METHODS[c.meth]}　${costText(c)}${c.amt != null ? "＝" + yen(c.amt) + "円" : ""}` : "") +
    `（平日${r.土 ? "・土" : ""}${r.日祝 ? "・日祝" : ""}）${r.休み.length ? `\n使わない日：${spanText(r.休み.slice())}` : ""}${r.未使用.length ? `\n使わなかった（申告）：${spanText(r.未使用.slice())}` : ""}${r.メモ ? "\n" + r.メモ : ""}` +
    (r.交渉 ? "\n" + ngText(r.交渉) : "");
};

/* ---------- 稼働表の操作（マウス・ペン。指はスクロールと、押して小窓） ---------- */
let drag = null;
let lastTap = null;   // 指で線を押した直前の1回（ダブルタップを見分ける）
function laneAt(x, y) { const c = document.elementFromPoint(x, y)?.closest?.(".cells"); return c ? { m: c.dataset.m, k: c.dataset.k, el: c } : null; }
function dayIndexAt(cells, x) {   // 拡大・縮小していても合うように、画面の上の幅÷日数で取る
  const r = cells.getBoundingClientRect(), N = rangeDays().length;
  return Math.max(0, Math.min(N - 1, Math.floor((x - r.left) / (r.width / N))));
}

$("chart").addEventListener("pointerdown", e => {
  hideCtx();
  const cells = e.target.closest(".cells");
  if (!cells || e.button !== 0) return;
  const barEl = e.target.closest(".bar");
  const r = barEl && D.予約.find(x => x.id === barEl.dataset.id);
  const d0 = dayIndexAt(cells, e.clientX);
  if (e.pointerType === "touch") { drag = { touch: true, r, x0: e.clientX, y0: e.clientY, cells, d0 }; return; }
  e.preventDefault();
  if (r) {
    const mode = e.target.classList.contains("eg") ? (e.target.classList.contains("l") ? "l" : "r") : "move";
    drag = { mode, r, el: barEl, cells, d0, dd: 0, lane: { m: r.機械, k: r.区分, el: cells }, moved: false, x0: e.clientX, y0: e.clientY,
      left: parseFloat(barEl.style.left), width: parseFloat(barEl.style.width), top: barEl.style.top };
  } else {
    drag = { mode: "new", cells, d0, d1: d0, lane: { m: cells.dataset.m, k: cells.dataset.k }, moved: false, x0: e.clientX, y0: e.clientY };
  }
  $("chart").setPointerCapture(e.pointerId);
});
$("chart").addEventListener("pointermove", e => {
  if (!drag) { tipMove(e); return; }
  if (drag.touch) return;
  if (!drag.moved && Math.abs(e.clientX - drag.x0) < 4 && Math.abs(e.clientY - drag.y0) < 4) return;
  drag.moved = true; hideTip();
  const dw = DW(), di = dayIndexAt(drag.cells, e.clientX);
  if (drag.mode === "new") {
    drag.d1 = di;
    let g = drag.cells.querySelector(".ghost");
    if (!g) { g = document.createElement("div"); g.className = "ghost"; drag.cells.appendChild(g); }
    const a = Math.min(drag.d0, di), b = Math.max(drag.d0, di), days = rangeDays();
    g.style.left = a * dw + "px"; g.style.width = (b - a + 1) * dw + "px"; g.style.top = "2px";
    g.textContent = (S.paint || "") + " " + md(days[a]) + "〜" + md(days[b]);
    return;
  }
  const dd = di - drag.d0; drag.dd = dd;
  const el = drag.el; el.classList.add("drag");
  if (drag.mode === "move") {
    el.style.left = drag.left + dd * dw + "px";
    const ln = laneAt(e.clientX, e.clientY);
    if (ln && ln.el !== el.parentNode) { ln.el.appendChild(el); el.style.top = "2px"; }
    if (ln) drag.lane = ln;
  } else if (drag.mode === "l") {
    const w = Math.max(dw, drag.width - dd * dw); el.style.left = drag.left + drag.width - w + "px"; el.style.width = w + "px";
  } else {
    el.style.width = Math.max(dw, drag.width + dd * dw) + "px";
  }
});
$("chart").addEventListener("pointerup", e => {
  const g = drag; drag = null;
  if (!g) return;
  if (g.touch) {
    if (Math.abs(e.clientX - g.x0) > 8 || Math.abs(e.clientY - g.y0) > 8) return;   // スクロールだった
    if (g.r) {   // 線：1回で選ぶだけ。続けてもう1回（ダブルタップ）で直す
      const t = lastTap, now = Date.now();
      lastTap = { id: g.r.id, at: now };
      if (t && t.id === g.r.id && now - t.at < 450) { lastTap = null; openRes(g.r); }
      else setSel(g.r.id);
    }
    else if (S.paint) quickAdd(g.cells.dataset.m, g.cells.dataset.k, rangeDays()[g.d0], rangeDays()[g.d0]);
    else openRes(null, { 機械: g.cells.dataset.m, 区分: g.cells.dataset.k, 開始: rangeDays()[g.d0], 終了: rangeDays()[g.d0] });
    return;
  }
  const days = rangeDays();
  if (g.mode === "new") {
    g.cells.querySelector(".ghost")?.remove();
    const a = days[Math.min(g.d0, g.d1)], b = days[Math.max(g.d0, g.d1)];
    if (!g.moved) {   // 押しただけ：選択を外す（ダブルクリックは下の dblclick）
      setSel(null);
      return;
    }
    if (S.paint) quickAdd(g.lane.m, g.lane.k, a, b);
    else openRes(null, { 機械: g.lane.m, 区分: g.lane.k, 開始: a, 終了: b });
    return;
  }
  const r = g.r;
  if (!g.moved) {
    setSel(r.id);   // 描き直さない（ダブルクリックが同じ要素に届くように）
    return;
  }
  const dd = g.dd;
  if (g.mode === "move" && !dd && g.lane.m === r.機械 && g.lane.k === r.区分) { renderChart(); return; }
  commit(d => {
    const x = d.予約.find(y => y.id === r.id);
    if (g.mode === "move") {
      x.開始 = dayAdd(x.開始, dd); x.終了 = dayAdd(x.終了, dd);
      x.休み = x.休み.map(k => dayAdd(k, dd)); x.出勤 = x.出勤.map(k => dayAdd(k, dd)); x.未使用 = x.未使用.map(k => dayAdd(k, dd));
      x.機械 = g.lane.m; x.区分 = g.lane.k;
    } else if (g.mode === "l") {
      x.開始 = dayAdd(x.開始, dd); if (x.開始 > x.終了) x.開始 = x.終了;
    } else {
      x.終了 = dayAdd(x.終了, dd); if (x.終了 < x.開始) x.終了 = x.開始;
    }
    trimDays(x); stamp(x);
    S.sel = x.id;
  });
});
/* ダブルクリック：バー＝直す／空いている所＝線を引く現場を選んでいればその日1日、なければ小窓 */
$("chart").addEventListener("dblclick", e => {
  if ($("dRes").open) return;   // 指のダブルタップで上が開いたあとに届いた分
  const t = document.elementFromPoint(e.clientX, e.clientY);   // 押している間は #chart に捕まえているので、下の要素を取り直す
  const cells = t?.closest(".cells"); if (!cells) return;
  const barEl = t.closest(".bar"), r = barEl && D.予約.find(x => x.id === barEl.dataset.id);
  if (r) { openRes(r); return; }
  const k = rangeDays()[dayIndexAt(cells, e.clientX)];
  if (S.paint) quickAdd(cells.dataset.m, cells.dataset.k, k, k);
  else openRes(null, { 機械: cells.dataset.m, 区分: cells.dataset.k, 開始: k, 終了: k });
});
$("chart").addEventListener("pointercancel",() => { drag = null; renderChart(); });
$("chart").addEventListener("pointerleave", () => hideTip());
function setSel(id) {
  S.sel = id;
  document.querySelectorAll("#chart .bar.sel").forEach(b => b.classList.remove("sel"));
  if (id) document.querySelector(`#chart .bar[data-id="${id}"]`)?.classList.add("sel");
}
const trimDays = x => { const inR = k => k >= x.開始 && k <= x.終了;
  x.休み = x.休み.filter(inR); x.出勤 = x.出勤.filter(inR); x.未使用 = x.未使用.filter(k => inR(k) && planned(x, k)); };
const stamp = x => { x.更新 = new Date().toLocaleString("ja-JP") + " " + me.name; };

function quickAdd(m, k, a, b) {
  const site = S.paint.trim();
  const id = newId("r");
  commit(d => {
    ensureSite(d, site);
    const r = { id, 機械: m, 区分: k, 現場: site, 開始: a, 終了: b, 土: false, 日祝: false, 休み: [], 出勤: [], 未使用: [], 状態: "確定", メモ: "" };
    if (a === b) { if (isSat(a)) r.土 = true; if (isSunHol(a)) r.日祝 = true; }   // 1日だけなら、その曜日でも使う
    stamp(r); d.予約.push(r); S.sel = r.id;
  });
}

/* マウスを載せると中身 */
function tipMove(e) {
  const b = e.target.closest?.(".bar");
  const t = $("tip");
  if (!b) { hideTip(); return; }
  const r = D.予約.find(x => x.id === b.dataset.id); if (!r) return;
  t.textContent = barTip(r); t.hidden = false;
  const x = Math.min(e.clientX + 14, innerWidth - t.offsetWidth - 8), y = Math.min(e.clientY + 16, innerHeight - t.offsetHeight - 8);
  t.style.left = x + "px"; t.style.top = y + "px";
}
const hideTip = () => { $("tip").hidden = true; };

/* 右クリックのメニュー */
$("chart").addEventListener("contextmenu", e => {
  const cells = e.target.closest(".cells"); if (!cells) return;
  e.preventDefault(); hideTip();
  const k = rangeDays()[dayIndexAt(cells, e.clientX)];
  const barEl = e.target.closest(".bar"), r = barEl && D.予約.find(x => x.id === barEl.dataset.id);
  const items = [];
  if (r) {
    S.sel = r.id;
    items.push(["「" + r.現場 + "」を直す…", () => openRes(r)]);
    const ch = fn => () => commit(d => { const x = d.予約.find(y => y.id === r.id); fn(x); trimDays(x); stamp(x); });
    if (planned(r, k)) {
      items.push([md(k) + " は使わない予定にする（休み）", ch(x => togglePlan(x, k))]);
      items.push(r.未使用.includes(k) ? [md(k) + " の「使わなかった」を取り消す", ch(x => { x.未使用 = x.未使用.filter(y => y !== k); })]
        : [md(k) + " は使わなかった（あとから申告）", ch(x => { x.未使用.push(k); })]);
    } else items.push([md(k) + " も使う予定にする", ch(x => togglePlan(x, k))]);
    items.push(["使う日をカレンダーで選ぶ…", () => openRes(r)]);
    items.push([r.交渉 ? "特別単価（交渉の結果）を直す…" : "特別単価（交渉の結果）を入れる…", () => openRes(r, {}, true)]);
    const nid = newId("r");
    if (k > r.開始) items.push([md(k) + " から後を別の予約に分ける", () => commit(d => {
      const x = d.予約.find(y => y.id === r.id), y = JSON.parse(JSON.stringify(x));
      y.id = nid; y.開始 = k; x.終了 = dayAdd(k, -1); trimDays(x); trimDays(y); stamp(x); stamp(y); d.予約.push(y); S.sel = y.id; })]);
    items.push([r.状態 === "仮" ? "確定にする" : "仮（相談中）にする", () => commit(d => { const x = d.予約.find(y => y.id === r.id); x.状態 = x.状態 === "仮" ? "確定" : "仮"; stamp(x); })]);
    items.push(["この現場だけ目立たせる", () => { S.focus = S.focus === r.現場 ? "" : r.現場; saveView(); syncBar(); renderChart(); renderLegend(); }]);
    items.push(null);
    items.push(["消す", () => commit(d => { d.予約 = d.予約.filter(y => y.id !== r.id); S.sel = null; }), "d"]);
  } else {
    const m = cells.dataset.m, kb = cells.dataset.k;
    items.push([md(k) + " から予約を入れる…", () => openRes(null, { 機械: m, 区分: kb, 開始: k, 終了: k })]);
    if (S.paint) items.push([md(k) + " に「" + S.paint + "」を1日", () => quickAdd(m, kb, k, k)]);
    items.push(["この機械の空きを探す", () => { FREE.m = m; setTab("free"); }]);
  }
  showCtx(e.clientX, e.clientY, items);
  renderChart();
});
function showCtx(x, y, items) {
  const c = $("ctx");
  c.innerHTML = items.map((it, i) => it ? `<div data-i="${i}" class="${it[2] || ""}">${esc(it[0])}</div>` : "<hr>").join("");
  c.hidden = false;
  c.style.left = Math.min(x, innerWidth - c.offsetWidth - 6) + "px"; c.style.top = Math.min(y, innerHeight - c.offsetHeight - 6) + "px";
  c.onclick = ev => { const i = ev.target.dataset.i; if (i == null) return; hideCtx(); items[i][1](); };
}
const hideCtx = () => { $("ctx").hidden = true; };
document.addEventListener("pointerdown", e => { if (!e.target.closest("#ctx")) hideCtx(); });

document.addEventListener("keydown", e => {
  if (e.target.closest("input,select,textarea") || document.querySelector("dialog[open]")) return;
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") { e.preventDefault(); e.shiftKey ? redo() : undo(); }
  else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "y") { e.preventDefault(); redo(); }
  else if ((e.key === "Delete" || e.key === "Backspace") && S.sel && S.tab === "chart") {
    e.preventDefault(); const id = S.sel; commit(d => { d.予約 = d.予約.filter(y => y.id !== id); S.sel = null; });
  } else if (e.key === "Escape") { S.sel = null; hideCtx(); renderChart(); }
});

/* 予定で使う日 ⇔ 使わない日（休み）を切り替える */
function togglePlan(x, k) {
  if (planned(x, k)) { x.出勤 = x.出勤.filter(y => y !== k); if (planned(x, k)) x.休み.push(k); x.未使用 = x.未使用.filter(y => y !== k); }
  else { x.休み = x.休み.filter(y => y !== k); if (!planned(x, k)) x.出勤.push(k); }
}

/* 小窓の中のカレンダー（全体工程表の「休工日を設定」と同じ考え方。日曜はじまり）
   plan＝押すと 使う⇔使わない（予定）、done＝押すと「使わなかった」（あとからの申告）を付ける／外す */
const CAL = { mode: "plan" };
function renderCal() {
  const v = resValue(), el = $("rCal");
  if (!v.開始 || !v.終了 || dayDiff(v.開始, v.終了) > 186) { el.innerHTML = `<span class="note">期間が長すぎます（半年まで）</span>`; return; }
  const today = todayKey();
  let k = dayAdd(v.開始, -dayDow(v.開始));
  const end = dayAdd(v.終了, 6 - dayDow(v.終了));
  const h = [`<div class="cal">` + WD.split("").map((w, i) => `<b class="${i === 0 ? "red" : i === 6 ? "blue" : ""}">${w}</b>`).join("")];
  for (; k <= end; k = dayAdd(k, 1)) {
    if (k < v.開始 || k > v.終了) { h.push(`<span class="out"></span>`); continue; }
    const p = planned(v, k), un = v.未使用.includes(k);
    const cls = !p ? "cno" : un ? "cun" : "cyes";
    const tip = (holidayOf(k) ? holidayOf(k) + "　" : "") + (!p ? "使わない" : un ? "使わなかった（申告）" : "使う");
    h.push(`<span class="${cls} ${dayClass(k)} ${k > today ? "fut" : ""}" data-k="${k}" title="${tip}">${k.endsWith("-01") || k === v.開始 ? +k.slice(5, 7) + "/" : ""}${+k.slice(8)}</span>`);
  }
  h.push(`</div>`);
  el.innerHTML = h.join("");
  document.querySelectorAll("input[name=rCalM]").forEach(x => { x.checked = x.value === CAL.mode; });
}
$("rCal").addEventListener("click", e => {
  const k = e.target.dataset?.k; if (!k) return;
  const rest = $("dRes")._rest, v = resValue();
  if (CAL.mode === "plan") { togglePlan(v, k); rest.休み = v.休み; rest.出勤 = v.出勤; rest.未使用 = v.未使用; }
  else {
    if (!planned(v, k)) return;   // 予定で使わない日は申告しない
    rest.未使用 = rest.未使用.includes(k) ? rest.未使用.filter(y => y !== k) : rest.未使用.concat(k);
  }
  resInfo();
});
document.querySelectorAll("input[name=rCalM]").forEach(x => x.addEventListener("change", () => { CAL.mode = x.value; renderCal(); }));

/* ---------- 予約の小窓 ---------- */
let editing = null;   // { id | null }
function openRes(r, init = {}, ng = false) {
  hideTip(); hideCtx();
  editing = r ? r.id : null;
  const v = r ? JSON.parse(JSON.stringify(r)) : Object.assign({ 機械: D.機械[0]?.id, 区分: "昼", 現場: S.paint, 開始: todayKey(), 終了: todayKey(), 土: false, 日祝: false, 休み: [], 出勤: [], 未使用: [], 状態: "確定", メモ: "", 交渉: null }, init);
  /* 特別単価（交渉の結果）。相手ははじめから機材センター長 */
  const x = v.交渉 || { 結果: "決まった", 単価: null, 日: todayKey(), 相手: NG_WHO, 方法: "対面", メモ: "" };
  $("rNg").checked = !!v.交渉 || ng; $("rNgBox").hidden = !$("rNg").checked;
  document.querySelector(`input[name=rNgR][value="${x.結果}"]`).checked = true;
  $("rNgP").value = x.単価 ?? ""; $("rNgD").value = x.日 || ""; $("rNgW").value = x.相手 || ""; $("rNgH").value = x.方法 || "対面"; $("rNgM").value = x.メモ || "";
  $("rNgLog").textContent = v.交渉?.記録 ? "入れた：" + v.交渉.記録 : "";
  $("dRes")._ng = v.交渉 || null;
  $("rN").value = v.台数 || 1;
  $("rPm").value = v.料金?.方式 || ""; $("rPq").value = v.料金?.数量 ?? ""; $("rPa").value = v.料金?.金額 ?? "";
  if (!r && v.開始 === v.終了) { if (isSat(v.開始)) v.土 = true; if (isSunHol(v.開始)) v.日祝 = true; }
  $("rTitle").textContent = r ? "予約を直す" : "予約を入れる";
  $("rM").innerHTML = D.機械.filter(m => !m.使わない || m.id === v.機械).map(m => `<option value="${m.id}">${esc(mLabel(m))}</option>`).join("");
  $("rM").value = v.機械;
  document.querySelector(`input[name=rK][value="${v.区分}"]`).checked = true;
  document.querySelector(`input[name=rSt][value="${v.状態}"]`).checked = true;
  $("rSite").value = v.現場 || ""; $("rFrom").value = v.開始; $("rTo").value = v.終了;
  $("rSat").checked = !!v.土; $("rSun").checked = !!v.日祝; $("rMemo").value = v.メモ || "";
  $("rDel").hidden = !r;
  $("dRes")._rest = { 休み: v.休み.slice(), 出勤: v.出勤.slice(), 未使用: v.未使用.slice() };
  CAL.mode = r && r.開始 <= todayKey() ? CAL.mode : "plan";
  resInfo();
  $("dRes").showModal();
  if (ng) setTimeout(() => { $("rNgBox").scrollIntoView({ block: "nearest" }); $("rNgP").focus(); }, 30);
  else if (!matchMedia("(any-pointer: coarse)").matches) setTimeout(() => (r ? $("rSite") : $("rSite").value ? $("rTo") : $("rSite")).focus(), 30);
}
const NG_WHO = "田村";   // 機材センター長
function ngValue() {
  if (!$("rNg").checked) return null;
  return { 結果: document.querySelector("input[name=rNgR]:checked").value, 単価: num($("rNgP").value), 日: $("rNgD").value,
    相手: $("rNgW").value.trim(), 方法: $("rNgH").value, メモ: $("rNgM").value.trim() };
}
/* 料金の決め方：空＝単価表どおり。数量を空にすると、月極は期間の月数・ほかは1 */
function payValue() {
  const meth = $("rPm").value, q = num($("rPq").value), a = num($("rPa").value);
  if (!meth) return null;
  return { 方式: meth, ...(q != null && meth !== "日" && meth !== "式" ? { 数量: q } : {}), ...(meth === "式" ? { 金額: a } : {}) };
}
$("rPm").innerHTML = `<option value="">単価表どおり</option>` + Object.entries(METHODS).map(([k, t]) => `<option value="${k}">${t}</option>`).join("");
["rN", "rPq", "rPa"].forEach(id => $(id).addEventListener("input", resInfo));
$("rPm").addEventListener("change", resInfo);
$("rNg").addEventListener("change", () => { $("rNgBox").hidden = !$("rNg").checked; if ($("rNg").checked) $("rNgP").focus(); resInfo(); });
["rNgP", "rNgD", "rNgW", "rNgH", "rNgM"].forEach(id => $(id).addEventListener("input", resInfo));
document.querySelectorAll("input[name=rNgR]").forEach(x => x.addEventListener("change", resInfo));
function resValue() {
  const rest = $("dRes")._rest;
  const v = { id: editing || "(new)", 機械: $("rM").value, 区分: document.querySelector("input[name=rK]:checked").value,
    現場: $("rSite").value.trim(), 開始: $("rFrom").value, 終了: $("rTo").value, 土: $("rSat").checked, 日祝: $("rSun").checked,
    休み: rest.休み.slice(), 出勤: rest.出勤.slice(), 未使用: rest.未使用.slice(), 状態: document.querySelector("input[name=rSt]:checked").value, メモ: $("rMemo").value.trim(),
    交渉: ngValue(), 台数: Math.max(1, Math.round(num($("rN").value) || 1)), 料金: payValue() };
  if (v.開始 && v.終了 && v.開始 > v.終了) [v.開始, v.終了] = [v.終了, v.開始];
  trimDays(v);
  return v;
}
function resInfo() {
  const v = resValue();
  renderCal();
  if (!v.開始 || !v.終了) { $("rInfo").textContent = ""; $("rWarn").innerHTML = ""; return; }
  const n = workDays(v), m = machine(v.機械), i = itemOf(m), meth = methodOf(v, i);
  const c = costOf(v), b = costOf(v, v.開始, v.終了, true), cap = m?.台数 || 1;
  /* 料金の欄：単価表の品と、いまの決め方 */
  $("rNcap").textContent = cap > 1 ? `（在庫 ${cap}台）` : "";
  $("rPm").options[0].textContent = "単価表どおり（" + METHODS[methodOf({ ...v, 料金: null }, i)] + "）";
  $("rPqW").hidden = meth === "日" || meth === "式"; $("rPaW").hidden = meth !== "式";
  $("rPu").textContent = UNIT[meth] + (meth === "月" && v.料金?.数量 == null ? `（空なら ${monthsOf(v)}か月）` : meth !== "月" && v.料金?.数量 == null ? "（空なら1）" : "");
  $("rPinfo").textContent = (i ? `単価表：${i.種別}　${itemLabel(i)}　` + [["日", "日"], ["施工", i.単位 || "施工"], ["月", "月"]]
      .map(([f, u]) => { const p = priceAt(i, f, v.開始); return p != null ? `${yen(p)}円／${u}` : ""; }).filter(Boolean).join("・") + (i.備考 ? `（${i.備考}）` : "")
    : m ? "単価表に無い機械（機械・現場 の単価" + (basePrice(m, v.区分) != null ? ` ${yen(basePrice(m, v.区分))}円／日` : "：未設定") + "）" : "");
  $("rNgBase").textContent = `円／${UNIT[meth]}` + (b.unit != null ? `（通常 ${yen(b.unit)}円）` : "（通常の単価なし）");
  $("rInfo").textContent = `${md(v.開始)}〜${md(v.終了)}　使う日 ${n}日` + (v.未使用.length ? `（使わなかった ${v.未使用.length}日を除く）` : "") +
    (isReal(v.現場) ? `　金額の目安 ` + (c.amt != null ? `${yen(c.amt)}円（${special(v) ? "特別単価 " : ""}${costText(c)}` +
      (special(v) && b.amt != null ? `。通常なら ${yen(b.amt)}円、差 ${yen(b.amt - c.amt)}円` : "") + "）" : "—（単価なし）") : "");
  /* 在庫の台数を超える日（1台の機械なら、ほかの予約がある日） */
  const cf = [], over = {}, mine = D.予約.filter(o => o.id !== v.id && o.機械 === v.機械 && o.区分 === v.区分 && overlap(o, v));
  for (let k = v.開始; k <= v.終了; k = dayAdd(k, 1)) {
    if (!works(v, k)) continue;
    const on = mine.filter(o => works(o, k));
    if (on.reduce((s, o) => s + (o.台数 || 1), 0) + (v.台数 || 1) > cap) on.length ? on.forEach(o => (over[o.現場] ||= []).push(k)) : (over["（台数が在庫より多い）"] ||= []).push(k);
  }
  Object.entries(over).forEach(([s, ds]) => cf.push(`${s}（${spanText(ds)}）`));
  /* もう一方の区分も見る（昼夜で別の現場へ行くのは運べるかどうか） */
  const other = D.予約.filter(o => o.id !== v.id && o.機械 === v.機械 && o.区分 !== v.区分 && overlap(o, v) && o.現場 !== v.現場)
    .map(o => { const ds = []; for (let k = v.開始; k <= v.終了; k = dayAdd(k, 1)) if (works(v, k) && works(o, k)) ds.push(k); return ds.length ? `${o.区分}は ${o.現場}（${spanText(ds)}）` : ""; }).filter(Boolean);
  $("rWarn").innerHTML = (cf.length ? `<div class="w">⚠ ${cap > 1 ? `在庫 ${cap}台を超えます。同じ日のこの機械の${esc(v.区分)}` : `同じ日に、この機械の${esc(v.区分)}はもう予約があります`}：${esc(cf.join("、"))}</div>`
    : `<div class="ok">✓ この期間、この機械の${esc(v.区分)}は${cap > 1 ? "在庫の台数の中に収まります" : "空いています"}</div>`) +
    (other.length ? `<div class="w" style="background:#fff7e0;color:#7a5200">参考：${esc(other.join("、"))}</div>` : "");
}
["rM", "rSite", "rFrom", "rTo", "rSat", "rSun"].forEach(id => $(id).addEventListener("input", resInfo));
document.querySelectorAll("input[name=rK]").forEach(x => x.addEventListener("change", resInfo));
$("rCancel").onclick = () => $("dRes").close();
$("rDel").onclick = () => { const id = editing; $("dRes").close(); commit(d => { d.予約 = d.予約.filter(y => y.id !== id); S.sel = null; }); };
$("rSave").onclick = () => {
  const v = resValue();
  if (!v.機械) { alert("機械を選んでください"); return; }
  if (!v.現場) { alert("行き先（現場名）を入れてください"); $("rSite").focus(); return; }
  if (!v.開始 || !v.終了) { alert("期間を入れてください"); return; }
  if (v.交渉 && v.交渉.結果 === "決まった" && v.交渉.単価 == null && methodOf(v) !== "式") { alert("決まった特別単価を入れてください。まだなら「交渉中」に"); $("rNgP").focus(); return; }
  if (v.料金?.方式 === "式" && v.料金.金額 == null) { alert("一式の金額を入れてください"); $("rPa").focus(); return; }
  /* 交渉の中身が変わったときだけ、入れた時刻と人を残す */
  const old = $("dRes")._ng, strip = x => x && JSON.stringify({ ...x, 記録: undefined });
  if (v.交渉) v.交渉.記録 = old && strip(old) === strip(v.交渉) ? old.記録 : new Date().toLocaleString("ja-JP") + " " + me.name;
  $("dRes").close();
  const nid = newId("r");
  commit(d => {
    ensureSite(d, v.現場);
    if (editing) { const x = d.予約.find(y => y.id === editing); Object.assign(x, v, { id: editing }); stamp(x); S.sel = x.id; }
    else { v.id = nid; stamp(v); d.予約.push(v); S.sel = v.id; }
  });
};
/* 小窓の Enter は次の欄へ（閉じない） */
$("dRes").addEventListener("keydown", e => {
  if (e.key !== "Enter" || e.isComposing) return;
  e.preventDefault();
  const f = [...$("dRes").querySelectorAll("input:not([type=radio]):not([type=checkbox]),select")].filter(x => x.offsetParent);
  const i = f.indexOf(e.target); if (i >= 0 && i < f.length - 1) f[i + 1].focus(); else e.target.blur();
});

/* ---------- 空き探し ---------- */
const FREE = { from: "", to: "", k: "昼", cat: "", m: "" };
function renderFree() {
  if (!FREE.from) { FREE.from = todayKey(); FREE.to = dayAdd(todayKey(), 6); }
  const el = $("pFree");
  const cats = catList().filter(c => D.機械.some(m => catOf(m) === c));
  const head = `<h2>空いている機械を探す</h2>
  <div class="form">
    <label>期間 <input type="date" id="fFrom" value="${FREE.from}"> 〜 <input type="date" id="fTo" value="${FREE.to}"></label>
    <label>区分 <select id="fK">${["昼", "夜", "昼夜とも"].map(x => `<option ${FREE.k === x ? "selected" : ""}>${x}</option>`).join("")}</select></label>
    <label>分類 <select id="fCat"><option value="">全部</option>${cats.map(c => `<option ${FREE.cat === c ? "selected" : ""}>${esc(c)}</option>`).join("")}</select></label>
    <label>機械 <select id="fM"><option value="">全部</option>${D.機械.filter(m => !m.使わない).map(m => `<option value="${m.id}" ${FREE.m === m.id ? "selected" : ""}>${esc(mLabel(m))}</option>`).join("")}</select></label>
    <label class="ck"><input type="checkbox" id="fSat"> 土曜も使う</label>
    <span class="note">平日に使う前提で見ます。左の色：緑＝全部空き・黄＝一部空き・赤＝ふさがり</span>
  </div>`;
  const days = []; if (FREE.from && FREE.to && FREE.from <= FREE.to && dayDiff(FREE.from, FREE.to) <= 366)
    for (let k = FREE.from; k <= FREE.to; k = dayAdd(k, 1)) if (isWeekday(k) || (FREE.sat && isSat(k))) days.push(k);
  const ks = FREE.k === "昼夜とも" ? ["昼", "夜"] : [FREE.k];
  const rows = D.機械.filter(m => !m.使わない && (!FREE.cat || catOf(m) === FREE.cat) && (!FREE.m || m.id === FREE.m)).map(m => {
    const mine = D.予約.filter(r => r.機械 === m.id && ks.includes(r.区分)), cap = m.台数 || 1;
    const busy = {}; let nb = 0;
    /* ふさがり＝在庫の台数を全部使っている日（昼夜ともなら、どちらかで） */
    days.forEach(k => { const on = mine.filter(r => works(r, k));
      if (ks.some(kb => on.filter(r => r.区分 === kb).reduce((s, r) => s + (r.台数 || 1), 0) >= cap)) nb++;
      on.forEach(r => (busy[r.現場] ||= []).push(k)); });
    const all = D.予約.filter(r => r.機械 === m.id);
    const before = all.filter(r => r.終了 < FREE.from).sort((a, b) => a.終了 < b.終了 ? 1 : -1)[0];
    const after = all.filter(r => r.開始 > FREE.to).sort((a, b) => a.開始 < b.開始 ? -1 : 1)[0];
    return { m, nb, free: days.length - nb, busy, before, after };
  }).sort((a, b) => a.nb - b.nb || D.機械.indexOf(a.m) - D.機械.indexOf(b.m));
  const tag = n => { const s = siteOf(n); return `<span class="tag${isLightColor(s.色) ? " lt" : ""}" style="background:${s.色}">${esc(n)}</span>`; };
  el.innerHTML = head + (days.length ? `<table class="t"><tr><th>機械</th><th>規格</th><th>空き</th><th>ふさがっている日</th><th>前の予定</th><th>次の予定</th><th></th></tr>` +
    rows.map(x => `<tr class="${x.nb === 0 ? "free" : x.free ? "part" : "busy"}">
      <td>${esc(x.m.名前)}${(x.m.台数 || 1) > 1 ? `<span class="small">（在庫 ${x.m.台数}台）</span>` : ""}</td><td>${esc(x.m.規格 || "")}</td>
      <td class="n">${x.nb === 0 ? "<b>全部空き</b>" : x.free + "／" + days.length + "日"}</td>
      <td>${Object.entries(x.busy).map(([s, ks2]) => tag(s) + `<span class="small">${spanText(ks2)}</span>`).join(" ")}</td>
      <td class="small">${x.before ? esc(x.before.現場) + "（〜" + md(x.before.終了) + "）" : ""}</td>
      <td class="small">${x.after ? esc(x.after.現場) + "（" + md(x.after.開始) + "〜）" : ""}</td>
      <td>${x.free ? `<button class="sub" data-book="${x.m.id}">予約する</button>` : ""}</td></tr>`).join("") + `</table>`
    : `<p class="note">期間を入れてください（1年まで）。</p>`);
  const upd = () => { FREE.from = $("fFrom").value; FREE.to = $("fTo").value; FREE.k = $("fK").value; FREE.cat = $("fCat").value; FREE.m = $("fM").value; FREE.sat = $("fSat").checked; renderFree(); };
  ["fFrom", "fTo", "fK", "fCat", "fM", "fSat"].forEach(id => $(id).onchange = upd);
  $("fSat").checked = !!FREE.sat;
  el.querySelectorAll("[data-book]").forEach(b => b.onclick = () =>
    openRes(null, { 機械: b.dataset.book, 区分: FREE.k === "夜" ? "夜" : "昼", 開始: FREE.from, 終了: FREE.to, 土: !!FREE.sat }));
}

/* ---------- 集計 ---------- */
function renderSum() {
  const el = $("pSum"), days = rangeDays(), from = S.from, to = days[days.length - 1];
  const wd = days.filter(isWeekday).length;
  const list = D.予約.filter(r => r.開始 <= to && r.終了 >= from);
  const ms = shownMachines();
  const head = `<div class="tabs2"><button data-by="machine" class="${S.sumBy === "machine" ? "on" : ""}">機械ごと</button><button data-by="site" class="${S.sumBy === "site" ? "on" : ""}">現場ごと</button><button data-by="nego" class="${S.sumBy === "nego" ? "on" : ""}">特別単価</button>
    <span class="note" style="align-self:center">期間：${from.replace(/-/g, "/")} 〜 ${to.replace(/-/g, "/")}（上の「開始」と「か月」）　平日 ${wd}日　金額は単価表（その日の単価）×使った日・台数の目安。月極・施工・一式は開始日の期間に数える。特別単価が決まった予約はその単価</span>
    <span class="sp"></span><button class="sub" id="sCsv">CSV に書き出す</button></div>`;
  let body = "", csv = [];
  if (S.sumBy === "machine") {
    const rows = ms.map(m => {
      const mine = list.filter(r => r.機械 === m.id);
      let dn = 0, nn = 0, used = 0, idle = 0, rep = 0;
      const amt = mine.filter(r => isReal(r.現場)).reduce((s, r) => s + (costOf(r, from, to, false, m).amt || 0), 0);
      days.forEach(k => {
        const on = mine.filter(r => works(r, k));
        const real = on.filter(r => isReal(r.現場));
        if (real.some(r => r.区分 === "昼")) dn++;
        if (real.some(r => r.区分 === "夜")) nn++;
        if (!isWeekday(k)) return;
        if (real.length) used++;
        else if (on.some(r => ["修理", "検査"].includes(siteOf(r.現場).種類))) rep++;
        else idle++;
      });
      return { m, dn, nn, used, idle, rep, amt, rate: wd ? used / wd : 0 };
    });
    if (S.sumSort === "rate") rows.sort((a, b) => a.rate - b.rate);
    const tot = rows.reduce((a, x) => ({ dn: a.dn + x.dn, nn: a.nn + x.nn, amt: a.amt + x.amt, idle: a.idle + x.idle }), { dn: 0, nn: 0, amt: 0, idle: 0 });
    /* 単価の欄：単価表の品ならその日極・月極（期間の終わりの日の単価）、無ければ機械の単価 */
    const pText = m => { const i = itemOf(m);
      if (!i) return yen(m.単価昼) + (m.単価夜 !== m.単価昼 ? "／" + yen(m.単価夜) : "");
      const d = priceAt(i, "日", to), mo = priceAt(i, "月", to), s = priceAt(i, "施工", to);
      return [d != null ? yen(d) + "／日" : "", mo != null ? yen(mo) + "／月" : "", s != null ? yen(s) + "／" + (i.単位 || "施工") : ""].filter(Boolean).join(" ") || "—"; };
    body = `<div class="form"><label>並び <select id="sSort"><option value="order" ${S.sumSort === "order" ? "selected" : ""}>機械の並び</option><option value="rate" ${S.sumSort === "rate" ? "selected" : ""}>稼働率の低い順（遊んでいる順）</option></select></label></div>
    <table class="t"><tr><th>分類</th><th>機械</th><th>規格</th><th>在庫</th><th>昼 使った日</th><th>夜 使った日</th><th>稼働率（平日）</th><th>遊び（平日）</th><th>修理・検査</th><th>単価（単価表）</th><th>金額の目安</th></tr>` +
      rows.map(x => `<tr><td class="small" style="white-space:nowrap">${esc(catOf(x.m))}</td><td>${esc(x.m.名前)}</td><td>${esc(x.m.規格 || "")}</td><td class="n">${x.m.台数 || 1}</td><td class="n">${x.dn || ""}</td><td class="n">${x.nn || ""}</td>
        <td class="n"><span class="meter${x.rate < .5 ? " low" : ""}"><s style="width:${Math.round(x.rate * 100)}%"></s></span>${Math.round(x.rate * 100)}%</td>
        <td class="n">${x.idle ? x.idle + "日" : ""}</td><td class="n">${x.rep ? x.rep + "日" : ""}</td>
        <td class="n small">${pText(x.m)}</td><td class="n">${yen(x.amt || null)}</td></tr>`).join("") +
      `<tr class="tot"><td colspan="4">計</td><td class="n">${tot.dn}</td><td class="n">${tot.nn}</td><td></td><td class="n">${tot.idle}日</td><td></td><td></td><td class="n">${yen(tot.amt)}</td></tr></table>`;
    csv = [["分類", "機械", "規格", "在庫", "昼 使った日", "夜 使った日", "稼働率", "遊び", "修理・検査", "単価", "金額の目安"],
      ...rows.map(x => [catOf(x.m), x.m.名前, x.m.規格, x.m.台数 || 1, x.dn, x.nn, Math.round(x.rate * 100) + "%", x.idle, x.rep, pText(x.m).replace(/,/g, ""), Math.round(x.amt)])];
  } else if (S.sumBy === "nego") {
    /* 特別単価：機材センター長と交渉した予約の一覧（期間にかかるもの） */
    const ng = list.filter(r => r.交渉).map(r => {
      const m = machine(r.機械), n = workDays(r, from, to), c = costOf(r, from, to), b = costOf(r, from, to, true);
      return { r, m, n, c, base: b.unit, unit: UNIT[c.meth], diff: special(r) && b.amt != null && c.amt != null ? b.amt - c.amt : null };
    }).sort((a, b) => (a.r.交渉.日 || "") < (b.r.交渉.日 || "") ? 1 : -1);
    const tot = ng.reduce((a, x) => a + (x.diff || 0), 0);
    body = ng.length ? `<table class="t"><tr><th>結果</th><th>現場</th><th>機械</th><th>区分</th><th>予約の期間</th><th>使った日</th><th>通常単価</th><th>特別単価</th><th>差額（期間内）</th><th>交渉した日</th><th>相手</th><th>方法</th><th>交渉の中身</th><th>入れた</th><th></th></tr>` +
      ng.map(x => `<tr><td>${x.r.交渉.結果 === "決まった" ? "<b>決まった</b>" : "交渉中"}</td><td>${esc(x.r.現場)}</td><td>${esc(mLabel(x.m))}</td><td>${x.r.区分}</td>
        <td class="small">${md(x.r.開始)}〜${md(x.r.終了)}</td><td class="n">${x.n}</td><td class="n">${x.base != null ? yen(x.base) + "／" + x.unit : ""}</td><td class="n">${special(x.r) ? yen(x.r.交渉.単価) + "／" + x.unit : ""}</td><td class="n">${yen(x.diff)}</td>
        <td class="small">${x.r.交渉.日 ? x.r.交渉.日.replace(/-/g, "/") : ""}</td><td>${esc(x.r.交渉.相手 || "")}</td><td>${esc(x.r.交渉.方法 || "")}</td><td class="small">${esc(x.r.交渉.メモ || "")}</td>
        <td class="small">${esc(x.r.交渉.記録 || "")}</td><td><button class="sub" data-ng="${x.r.id}">直す</button></td></tr>`).join("") +
      `<tr class="tot"><td colspan="8">差額の計（通常の単価との差。期間内の分）</td><td class="n">${yen(tot)}</td><td colspan="6"></td></tr></table>`
      : `<p class="note">この期間に特別単価（交渉）の予約はありません。稼働表の線を右クリック（指は線をダブルタップ）して、予約の小窓の「特別単価」に入れます。</p>`;
    csv = [["結果", "現場", "機械", "区分", "開始", "終了", "使った日", "通常単価", "特別単価", "差額", "交渉した日", "相手", "方法", "交渉の中身", "入れた"],
      ...ng.map(x => [x.r.交渉.結果, x.r.現場, mLabel(x.m), x.r.区分, x.r.開始, x.r.終了, x.n, x.base ?? "", special(x.r) ? x.r.交渉.単価 : "", x.diff ?? "",
        x.r.交渉.日 || "", x.r.交渉.相手 || "", x.r.交渉.方法 || "", x.r.交渉.メモ || "", x.r.交渉.記録 || ""])];
  } else {
    /* 現場ごと：その現場で使った機械と日数（Excel の右の「現場名を選ぶ」の代わり）。決め方・単価・台数・特別単価が違う予約は別の行 */
    const by = {};
    list.forEach(r => {
      const m = machine(r.機械); if (!m) return;
      const c = costOf(r, from, to, false, m), sp = special(r);
      if (!c.days && !c.amt) return;
      const x = (by[r.現場] ||= { site: r.現場, items: {}, days: 0, amt: 0 });
      const it = (x.items[[m.id, r.区分, c.meth, c.unit, c.n, sp].join("|")] ||= { m, k: r.区分, meth: c.meth, unit: c.unit, n: c.n, sp, days: 0, qty: 0, amt: 0 });
      it.days += c.days; x.days += c.days;
      if (c.meth === "日" ? c.days : c.amt) it.qty += c.meth === "日" ? c.days : c.qty;
      if (isReal(r.現場) && c.amt) { it.amt += c.amt; x.amt += c.amt; }
    });
    const sites = Object.values(by).sort((a, b) => (isReal(b.site) - isReal(a.site)) || b.amt - a.amt);
    const itText = it => (it.sp ? "<b>特</b> " : "") + (it.meth === "式" ? "一式" : it.unit == null ? "単価なし" : `${yen(it.unit)}円／${UNIT[it.meth]}`) + (it.meth !== "日" && it.meth !== "式" ? `×${it.qty}` : "") + (it.n > 1 ? `×${it.n}台` : "");
    body = `<table class="t"><tr><th>現場</th><th>機械</th><th>区分</th><th>使った日</th><th>単価（決め方）</th><th>金額の目安</th></tr>` +
      sites.map(x => {
        const its = Object.values(x.items).sort((a, b) => D.機械.indexOf(a.m) - D.機械.indexOf(b.m));
        const s = siteOf(x.site);
        return its.map((it, i) => `<tr>${i === 0 ? `<td rowspan="${its.length + 1}"><span class="tag${isLightColor(s.色) ? " lt" : ""}" style="background:${s.色}">${esc(x.site)}</span><div class="small">${esc(s.種類)}</div></td>` : ""}
          <td>${esc(mLabel(it.m))}</td><td>${it.k}</td><td class="n">${it.days}</td><td class="n small">${isReal(x.site) ? itText(it) : ""}</td><td class="n">${isReal(x.site) ? yen(it.amt) : ""}</td></tr>`).join("") +
          `<tr class="tot"><td colspan="2">計</td><td class="n">${x.days}</td><td></td><td class="n">${isReal(x.site) ? yen(x.amt) : ""}</td></tr>`;
      }).join("") + `</table>`;
    csv = [["現場", "種類", "機械", "区分", "使った日", "決め方", "単価", "数量", "台数", "特別単価", "金額の目安"]];
    sites.forEach(x => Object.values(x.items).forEach(it => csv.push([x.site, siteOf(x.site).種類, mLabel(it.m), it.k, it.days, METHODS[it.meth],
      isReal(x.site) ? it.unit ?? "" : "", it.qty, it.n, it.sp ? "特" : "", isReal(x.site) ? Math.round(it.amt) : ""])));
  }
  el.innerHTML = head + body;
  el.querySelectorAll("[data-by]").forEach(b => b.onclick = () => { S.sumBy = b.dataset.by; saveView(); renderSum(); });
  el.querySelectorAll("[data-ng]").forEach(b => b.onclick = () => { const r = D.予約.find(x => x.id === b.dataset.ng); if (r) openRes(r, {}, true); });
  if ($("sSort")) $("sSort").onchange = () => { S.sumSort = $("sSort").value; saveView(); renderSum(); };
  $("sCsv").onclick = () => download(`自社機械_${{ machine: "機械ごと", site: "現場ごと", nego: "特別単価" }[S.sumBy]}_${from}_${to}.csv`,
    "﻿" + csv.map(r => r.map(v => /[",\n]/.test(String(v ?? "")) ? `"${String(v).replace(/"/g, '""')}"` : (v ?? "")).join(",")).join("\r\n"), "text/csv");
}

/* ---------- 機械・現場（マスタ） ---------- */
/* 期（9月はじまり。2026-09〜2027-08＝52期）。工事コード 510012 なら頭の 51 */
const kiOf = k => +k.slice(0, 4) - (+k.slice(5, 7) >= 9 ? 1974 : 1975);
const siteKi = (s, firstDay) => /^\d{6}$/.test(s.コード || "") ? +s.コード.slice(0, 2) : firstDay ? kiOf(firstDay) : 0;
let MTAB = "machine";
function renderMaster() {
  const el = $("pMaster");
  const keep = cellWant || cellAt(document.activeElement);   // 描き直しても、いた欄（または移る先の欄）へ戻る
  const head = `<div class="tabs2"><button data-mt="machine" class="${MTAB === "machine" ? "on" : ""}">機械</button><button data-mt="site" class="${MTAB === "site" ? "on" : ""}">現場・行き先</button></div>`;
  if (MTAB === "machine") {
    /* 単価表の品を選ぶ欄（種別ごと）。空＝単価表を使わない（機械の単価）。自動＝道具が名前で決めた品 */
    const auto = m => PR?.対応?.[mKey(m)] || "";
    const itemSel = m => `<select data-f="単価品" class="isel"><option value="*" ${m.単価品 === undefined ? "selected" : ""}>自動${auto(m) ? "：" + esc(itemLabel(PRI[auto(m)])) : "（無し）"}</option>
      <option value="" ${m.単価品 === "" ? "selected" : ""}>使わない（右の単価）</option>` +
      (PR?.種別 || []).map(c => `<optgroup label="${esc(c)}">` + PR.品.filter(i => i.種別 === c && !i.無し).map(i =>
        `<option value="${esc(i.key)}" ${m.単価品 === i.key ? "selected" : ""}>${esc(itemLabel(i))}</option>`).join("") + `</optgroup>`).join("") + `</select>`;
    const cats = catList();
    el.innerHTML = head + `<p class="note">変えたら欄を離れたときに保存します。「使わない」にすると稼働表に出ません（予約は残ります）。
      ${PR ? `単価・分類は単価表（実行予算単価「機材センター」、${esc(PR.元の更新 || "")} 更新の版）の品から。単価表に無い機械だけ右の 単価昼・単価夜（1日あたり）を使います。` : "単価表（単価表.json）を読めていません。単価は右の 単価昼・単価夜（1日あたり）。"}
      在庫が何台もある機械（送風機など）は 台数 に台数を入れると、予約ごとに使う台数を入れられ、台数を超えたときだけ二重の予約になります。</p>
    <table class="t grid" id="mt"><tr><th></th><th>分類</th><th>機械名</th><th>規格</th><th>台数</th><th>単価表の品</th><th>単価 昼</th><th>単価 夜</th><th>メモ</th><th>使わない</th><th>予約</th><th></th></tr>` +
      D.機械.map((m, i) => `<tr data-id="${m.id}"><td><button class="sub" tabindex="-1" data-up="${i}" ${i ? "" : "disabled"}>▲</button><button class="sub" tabindex="-1" data-dn="${i}" ${i < D.機械.length - 1 ? "" : "disabled"}>▼</button></td>
        <td>${itemOf(m) ? `<span class="small" title="単価表の種別">${esc(catOf(m))}</span>` : `<select data-f="分類">${cats.map(c => `<option ${(m.分類 || "その他") === c ? "selected" : ""}>${esc(c)}</option>`).join("")}</select>`}</td>
        <td><input data-f="名前" value="${esc(m.名前)}" size="16"></td><td><input data-f="規格" value="${esc(m.規格 || "")}" size="10"></td>
        <td><input data-f="台数" class="num" inputmode="numeric" value="${m.台数 || 1}" size="3"></td><td>${itemSel(m)}</td>
        <td><input data-f="単価昼" class="num" inputmode="numeric" value="${m.単価昼 ?? ""}" ${itemOf(m) ? 'disabled title="単価表の品を使っています"' : ""}></td><td><input data-f="単価夜" class="num" inputmode="numeric" value="${m.単価夜 ?? ""}" ${itemOf(m) ? 'disabled title="単価表の品を使っています"' : ""}></td>
        <td><input data-f="メモ" value="${esc(m.メモ || "")}" size="16"></td>
        <td style="text-align:center"><input type="checkbox" data-f="使わない" ${m.使わない ? "checked" : ""}></td>
        <td class="n small">${D.予約.filter(r => r.機械 === m.id).length}</td>
        <td><button class="danger" tabindex="-1" data-del="${m.id}">消す</button></td></tr>`).join("") + `</table>
    <p><button id="mAdd">＋ 機械を足す</button></p>`;
    el.querySelectorAll("#mt [data-f]").forEach(inp => inp.onchange = () => {
      const id = inp.closest("tr").dataset.id, f = inp.dataset.f;
      commit(d => { const m = d.機械.find(x => x.id === id);
        if (f === "単価品") { if (inp.value === "*") delete m.単価品; else m.単価品 = inp.value; return; }
        if (f === "台数") { m.台数 = Math.max(1, Math.round(num(inp.value) || 1)); return; }
        m[f] = inp.type === "checkbox" ? inp.checked : f.startsWith("単価") ? num(inp.value) : inp.value.trim(); });
    });
    el.querySelectorAll("[data-up],[data-dn]").forEach(b => b.onclick = () => {
      const i = +(b.dataset.up ?? b.dataset.dn), j = b.dataset.up != null ? i - 1 : i + 1;
      const a = D.機械[i].id, b2 = D.機械[j].id;   // 番号でなく id で入れ替える（ほかの人が足した機械があってもずれない）
      commit(d => { const x = d.機械.findIndex(m => m.id === a), y = d.機械.findIndex(m => m.id === b2); if (x >= 0 && y >= 0) [d.機械[x], d.機械[y]] = [d.機械[y], d.機械[x]]; });
    });
    el.querySelectorAll("[data-del]").forEach(b => b.onclick = () => {
      const id = b.dataset.del, n = D.予約.filter(r => r.機械 === id).length, m = machine(id);
      if (!confirm(`「${mLabel(m)}」を消しますか？` + (n ? `\nこの機械の予約 ${n}本も消えます。\n（残したいときは「使わない」にしてください）` : ""))) return;
      commit(d => { d.機械 = d.機械.filter(x => x.id !== id); d.予約 = d.予約.filter(r => r.機械 !== id); });
    });
    $("mAdd").onclick = () => { const id = newId("m"); commit(d => { d.機械.push({ id, 名前: "新しい機械", 規格: "", 分類: S.cat || "小型機械", 単価昼: null, 単価夜: null, メモ: "", 使わない: false }); }); };
  } else {
    const cnt = {}, first = {}, last = {};
    D.予約.forEach(r => { cnt[r.現場] = (cnt[r.現場] || 0) + 1;
      if (!first[r.現場] || r.開始 < first[r.現場]) first[r.現場] = r.開始;
      if (!last[r.現場] || r.終了 > last[r.現場]) last[r.現場] = r.終了; });
    const nDone = D.現場.filter(s => s.完了).length;
    const list = D.現場.filter(s => S.showDone || !s.完了).sort((a, b) => (a.完了 - b.完了) || (siteKi(b, first[b.名前]) - siteKi(a, first[a.名前])) ||
      ((last[b.名前] || "") > (last[a.名前] || "") ? 1 : -1));
    el.innerHTML = head + `<p class="note">現場名は全体工程表の「社内工事名」とそろえておくと、あとでつなげられます。種類が「現場」以外（置場・修理・検査）は稼働率・金額に数えません。
      「完了」にすると、現場を選ぶ欄（線を引く現場・小窓の行き先）に出なくなります。名前を変えると、その現場の予約も全部変わります（同じ名前にすると1つにまとまります）。<br>
      期＝工事コードの頭2けた（無ければ最初に使った日の期。9月はじまり）。</p>
    <div class="form"><label class="ck"><input type="checkbox" id="sDone" ${S.showDone ? "checked" : ""}> 完了した現場も表示（${nDone}件）</label>
      <button class="sub" id="sFill" title="受注台帳等・配置表の名簿と照合した結果（取り込み_自社機械予定表.json）から、現場の情報だけを入れる。予約は変えません">${window.KIKAI_IMPORT ? "取り込みデータ" : "ファイル（取り込み_自社機械予定表.json）"}から 工事コード・元請・地域・工期・完了 を入れる</button>
      <input type="file" id="sFile" accept=".json,application/json" hidden></div>
    <table class="t grid" id="st"><tr><th>期</th><th>現場名・行き先</th><th>工事コード</th><th>元請</th><th>地域</th><th>工期</th><th>種類</th><th>色</th><th>完了</th><th>予約</th><th>最後に使う日</th><th></th></tr>` +
      list.map(s => `<tr data-n="${esc(s.名前)}"${s.完了 ? ' style="color:#888"' : ""}>
        <td class="n">${siteKi(s, first[s.名前]) || ""}</td>
        <td title="${esc([s.正式名, s.担当 && "担当 " + s.担当, s.場所].filter(Boolean).join("\n"))}"><input data-f="名前" value="${esc(s.名前)}" size="18"></td>
        <td><input data-f="コード" value="${esc(s.コード || "")}" size="7" inputmode="numeric"></td>
        <td><input data-f="元請" value="${esc(s.元請 || "")}" size="12"></td>
        <td><input data-f="地域" value="${esc(s.地域 || "")}" size="8"></td>
        <td class="small" style="white-space:nowrap">${s.工期自 || s.工期至 ? `${(s.工期自 || "").replace(/-/g, "/")}〜${(s.工期至 || "").replace(/-/g, "/")}` : ""}</td>
        <td><select data-f="種類">${KINDS.map(k => `<option ${s.種類 === k ? "selected" : ""}>${k}</option>`).join("")}</select></td>
        <td><input type="color" data-f="色" value="${esc(s.色 || "#8a96a3")}"></td>
        <td style="text-align:center"><input type="checkbox" data-f="完了" ${s.完了 ? "checked" : ""}></td>
        <td class="n small">${cnt[s.名前] || 0}</td><td class="small">${last[s.名前] ? last[s.名前].replace(/-/g, "/") : ""}</td>
        <td>${cnt[s.名前] ? "" : `<button class="danger" tabindex="-1" data-del="${esc(s.名前)}">消す</button>`}</td></tr>`).join("") + `</table>
    <p><button id="sAdd">＋ 現場を足す</button></p>`;
    $("sDone").onchange = () => { S.showDone = $("sDone").checked; saveView(); renderMaster(); };
    /* 現場の情報だけを入れる（空いている欄だけ。予約は変えない）。公開した画面には取り込みデータが無いので、ファイルを選ぶ */
    const fillSites = data => {
      const src = {}; (data.現場 || []).forEach(x => { src[x.名前] = x; });
      let n = 0;
      commit(d => d.現場.forEach(s => { const x = src[s.名前]; if (!x) return;
        ["コード", "元請", "地域", "正式名", "工期自", "工期至", "担当", "場所"].forEach(f => { if (!s[f] && x[f]) { s[f] = x[f]; n++; } });
        if (x.完了 && !s.完了) { s.完了 = true; n++; } }));   // しばらく使っていない現場は完了に
      alert(n ? `${n}か所に入れました（空いている欄だけ）。` : "入れる所はありませんでした。");
    };
    $("sFill").onclick = () => window.KIKAI_IMPORT ? fillSites(window.KIKAI_IMPORT) : $("sFile").click();
    $("sFile").onchange = async () => {
      const f = $("sFile").files[0]; $("sFile").value = ""; if (!f) return;
      try { fillSites(normalize(JSON.parse(await f.text()))); } catch (err) { alert("読めませんでした：" + err.message); }
    };
    el.querySelectorAll("#st [data-f]").forEach(inp => inp.onchange = () => {
      const n = inp.closest("tr").dataset.n, f = inp.dataset.f;
      const val = inp.type === "checkbox" ? inp.checked : f === "コード" ? half(inp.value) : inp.value.trim();
      if (f === "名前") {
        if (!val) { renderMaster(); return; }
        const merge = D.現場.some(s => s.名前 === val);
        if (merge && !confirm(`「${val}」はもうあります。「${n}」の予約を「${val}」にまとめますか？`)) { renderMaster(); return; }
        commit(d => { d.予約.forEach(r => { if (r.現場 === n) r.現場 = val; });
          if (merge) d.現場 = d.現場.filter(s => s.名前 !== n); else d.現場.find(s => s.名前 === n).名前 = val;
          if (S.paint === n) S.paint = val; if (S.focus === n) S.focus = val; });
        return;
      }
      commit(d => { const s = d.現場.find(x => x.名前 === n); s[f] = val; if (f === "種類" && val !== "現場" && KIND_COLOR[val]) s.色 = KIND_COLOR[val]; });
    });
    el.querySelectorAll("[data-del]").forEach(b => b.onclick = () => commit(d => { d.現場 = d.現場.filter(s => s.名前 !== b.dataset.del); }));
    $("sAdd").onclick = () => { const n = prompt("現場名（社内工事名）"); if (n && n.trim()) commit(d => ensureSite(d, n.trim())); };
  }
  el.querySelectorAll("[data-mt]").forEach(b => b.onclick = () => { MTAB = b.dataset.mt; renderMaster(); });
  if (keep) focusCell(keep);
}

/* ---------- 表の中の移動（Excel と同じ） ----------
   Tab＝右（行の終わりは次の行の頭）・Shift＋Tab＝左・Enter＝下・Shift＋Enter＝上・↑↓＝上下・←→＝字の端にいるときは左右の欄へ。
   値を変えると保存して表を描き直すので（前は Tab で移る先の欄が消えて、フォーカスが変な所へ飛んだ）、
   移る先は「行のキー（機械の id・現場名）＋欄」で覚え、描き直した表でその欄へ移る。ボタン（▲▼・消す）は Tab で止まらない */
let cellWant = null;
const rowKeyOf = tr => tr.dataset.id || tr.dataset.n;
function cellAt(x) {
  const tbl = x?.closest?.("table.grid"), tr = x?.closest?.("tr");
  return tbl && tr && x.dataset.f ? { tbl: tbl.id, key: rowKeyOf(tr), f: x.dataset.f } : null;
}
function focusCell(w) {
  const tbl = w && document.getElementById(w.tbl); if (!tbl) return false;
  const tr = [...tbl.querySelectorAll("tr[data-id],tr[data-n]")].find(t => rowKeyOf(t) === w.key);
  const x = tr && tr.querySelector(`[data-f="${w.f}"]`); if (!x) return false;
  x.focus();
  if (x.select && x.tagName === "INPUT" && !["checkbox", "color"].includes(x.type)) x.select();
  return true;
}
document.addEventListener("keydown", e => {
  const x = e.target, tbl = x.closest?.("table.grid");
  if (!tbl || !x.dataset.f || e.isComposing || e.altKey || e.ctrlKey) return;
  const text = x.tagName === "INPUT" && !["checkbox", "color"].includes(x.type);
  const all = text && x.selectionStart === 0 && x.selectionEnd === x.value.length;   // 入ったばかり（全部選んである）
  let dr = 0, dc = 0;
  if (e.key === "Tab") dc = e.shiftKey ? -1 : 1;
  else if (e.key === "Enter") dr = e.shiftKey ? -1 : 1;
  else if (e.key === "ArrowDown" && x.tagName !== "SELECT") dr = 1;
  else if (e.key === "ArrowUp" && x.tagName !== "SELECT") dr = -1;
  else if (e.key === "ArrowRight" && x.tagName !== "SELECT" && (!text || all || (x.selectionStart === x.value.length && x.selectionEnd === x.value.length))) dc = 1;
  else if (e.key === "ArrowLeft" && x.tagName !== "SELECT" && (!text || all || (x.selectionStart === 0 && x.selectionEnd === 0))) dc = -1;
  else return;
  e.preventDefault();
  const rows = [...tbl.querySelectorAll("tr[data-id],tr[data-n]")], tr = x.closest("tr");
  let ri = rows.indexOf(tr), ci = [...tr.querySelectorAll("[data-f]")].indexOf(x);
  const cols = r => [...rows[r].querySelectorAll("[data-f]")];
  ri += dr; ci += dc;
  if (e.key === "Tab") {   // 行の端で次（前）の行へ
    if (ci >= cols(ri).length && ri + 1 < rows.length) { ri++; ci = 0; }
    if (ci < 0 && ri > 0) { ri--; ci = cols(ri).length - 1; }
  }
  if (ri < 0 || ri >= rows.length) return;
  const t = cols(ri)[Math.max(0, Math.min(ci, cols(ri).length - 1))];
  let key = rowKeyOf(rows[ri]);
  if (rows[ri] === tr && tbl.id === "st" && x.dataset.f === "名前" && x.value.trim()) key = x.value.trim();   // 現場名を直した行は新しい名前がキー
  cellWant = { tbl: tbl.id, key, f: t.dataset.f };
  x.blur();   // 値が変わっていれば、ここで保存・描き直し（renderMaster が cellWant の欄へ戻す）
  focusCell(cellWant);
  cellWant = null;
});
/* マウスで別の欄を押したときも、描き直しで押した欄が消えないように */
document.addEventListener("pointerdown", e => {
  const w = cellAt(e.target); if (!w) return;
  cellWant = w; setTimeout(() => { cellWant = null; }, 400);
});

/* ---------- 上の帯・凡例 ---------- */
function syncBar() {
  $("vFrom").value = S.from; $("vMonths").value = S.months; $("vQ").value = S.q; $("vUnused").checked = S.unused;
  $("vPaint").value = S.paint; $("vIdle").checked = S.idle;
  const cats = catList().filter(c => D.機械.some(m => catOf(m) === c));
  $("vCat").innerHTML = `<option value="">分類：全部</option>` + cats.map(c => `<option ${S.cat === c ? "selected" : ""}>${esc(c)}</option>`).join("");
  const active = D.現場.filter(s => !s.完了);
  $("siteList").innerHTML = active.map(s => `<option value="${esc(s.名前)}">`).join("");
  const from = S.from, to = rangeEnd();
  const inView = [...new Set(D.予約.filter(r => r.開始 <= to && r.終了 >= from).map(r => r.現場))];
  $("vFocus").innerHTML = `<option value="">―</option>` + inView.map(n => `<option ${S.focus === n ? "selected" : ""}>${esc(n)}</option>`).join("");
}
function renderLegend() {
  const from = S.from, to = rangeEnd();
  const n = {}; D.予約.filter(r => r.開始 <= to && r.終了 >= from).forEach(r => { n[r.現場] = (n[r.現場] || 0) + workDays(r, from, to); });
  const list = Object.keys(n).sort((a, b) => (isReal(b) - isReal(a)) || n[b] - n[a]);
  $("legend").innerHTML = (list.length ? `<b>この期間の行き先：</b>` : "") + list.map(s => `<span data-s="${esc(s)}" class="${S.focus === s ? "on" : ""}" title="押すとこの現場だけ目立たせる／もう一度で戻す。ダブルクリックで「線を引く現場」にする">
    <i style="background:${siteOf(s).色}"></i>${esc(s)} <small>${n[s]}日</small></span>`).join("") +
    `<span class="note">　空いている所：なぞる＝予約の線を引く（線を引く現場を選んでいればその現場、無ければ小窓で選ぶ）・ダブルクリック＝その日1日　バー：動かす・端で延ばす・ダブルクリック＝直す・右クリック＝メニュー　<i style="background:var(--idle)"></i>＝遊んでいる平日　機械名の下の %＝表示している期間の稼働率（平日のうち現場にいた日）</span>`;
  $("legend").querySelectorAll("[data-s]").forEach(x => {
    x.onclick = () => { S.focus = S.focus === x.dataset.s ? "" : x.dataset.s; saveView(); syncBar(); renderChart(); renderLegend(); };
    x.ondblclick = () => { S.paint = x.dataset.s; saveView(); syncBar(); };
  });
}

function setTab(t) {
  S.tab = t; saveView();
  document.body.dataset.tab = t;
  document.querySelectorAll("#tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === t));
  document.querySelectorAll(".page").forEach(p => p.classList.toggle("on", p.id === "p" + t[0].toUpperCase() + t.slice(1)));
  renderAll();
}
function renderAll() {
  syncBar();
  $("bUndo").disabled = !hist.length; $("bRedo").disabled = !fut.length;
  if (S.tab === "chart") { renderChart(); renderLegend(); }
  else if (S.tab === "free") renderFree();
  else if (S.tab === "sum") renderSum();
  else renderMaster();
}

document.querySelectorAll("#tabs button").forEach(b => b.onclick = () => setTab(b.dataset.tab));
$("vFrom").onchange = () => { if ($("vFrom").value) { S.from = $("vFrom").value; saveView(); renderAll(); } };
$("vMonths").onchange = () => { S.months = +$("vMonths").value; saveView(); renderAll(); };
$("vPrev").onclick = () => { S.from = addMonths(S.from, -1); saveView(); renderAll(); };
$("vNext").onclick = () => { S.from = addMonths(S.from, 1); saveView(); renderAll(); };
$("vToday").onclick = () => { S.from = dayAdd(todayKey(), -((dayDow(todayKey()) + 6) % 7) - 7); saveView(); renderAll();   // 先週の月曜から
  requestAnimationFrame(() => { $("wrap").scrollLeft = 0; }); };
$("vCat").onchange = () => { S.cat = $("vCat").value; saveView(); renderAll(); };
$("vQ").oninput = () => { S.q = $("vQ").value; saveView(); renderAll(); };
$("vUnused").onchange = () => { S.unused = $("vUnused").checked; saveView(); renderAll(); };
$("vPaint").onchange = () => { S.paint = $("vPaint").value.trim(); saveView(); };
$("vIdle").onchange = () => { S.idle = $("vIdle").checked; saveView(); renderChart(); };
$("vFocus").onchange = () => { S.focus = $("vFocus").value; saveView(); renderChart(); renderLegend(); };
$("bUndo").onclick = undo; $("bRedo").onclick = redo;
$("zIn").onclick = () => setZoom((S.zoom || 1) + 0.1);
$("zOut").onclick = () => setZoom((S.zoom || 1) - 0.1);
$("zReset").onclick = () => setZoom(1);
$("wrap").addEventListener("wheel", e => { if (!e.ctrlKey) return; e.preventDefault(); setZoom((S.zoom || 1) + (e.deltaY < 0 ? 0.1 : -0.1)); }, { passive: false });
document.addEventListener("keydown", e => {
  if (!(e.ctrlKey || e.metaKey) || S.tab !== "chart" || document.querySelector("dialog[open]")) return;
  if (e.key === "+" || e.key === ";" || e.key === "=") { e.preventDefault(); setZoom((S.zoom || 1) + 0.1); }
  else if (e.key === "-") { e.preventDefault(); setZoom((S.zoom || 1) - 0.1); }
  else if (e.key === "0") { e.preventDefault(); setZoom(1); }
});
let resizeT = 0;
addEventListener("resize", () => { clearTimeout(resizeT); resizeT = setTimeout(() => { if (S.tab === "chart") renderChart(); }, 150); });
$("warn").onclick = () => { const b = document.querySelector(".bar.conf"); if (b) b.scrollIntoView({ block: "center", inline: "center" }); };

/* ---------- データの読み書き ---------- */
function download(name, text, type = "application/json") {
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([text], { type })); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
function loadData(d, how) {
  d = normalize(d);
  const had = D.予約.length || D.機械.length;
  if (!DEMO && had && !confirm("会社の OneDrive の稼働データを、全員分まとめて置き換えます。ほかの人が入れた予約も置き換わります。本当によろしいですか？")) return false;
  if (had && !confirm(`いまのデータ（機械 ${D.機械.length}台・予約 ${D.予約.length}本）を、読み込むデータ（機械 ${d.機械.length}台・予約 ${d.予約.length}本）に置き換えます。よろしいですか？\n（「↶ 戻す」で元に戻せます）`)) return false;
  commit(replaceAll(d));   // 差し替える前を控えに持つので「↶ 戻す」で前のデータに戻る
  $("ioMsg").textContent = `${how}：機械 ${d.機械.length}台・現場 ${d.現場.length}・予約 ${d.予約.length}本を読みました。`;
  return true;
}
async function importBundled() {
  try {
    /* index.html の隣の ../data/取り込み_自社機械予定表.js（ダブルクリックで開いても読める） */
    if (!window.KIKAI_IMPORT) throw new Error("no data");
    if (loadData(JSON.parse(JSON.stringify(window.KIKAI_IMPORT)), "Excel から取り込んだデータ")) {
      S.from = dayAdd(todayKey(), -((dayDow(todayKey()) + 6) % 7) - 7); saveView(); renderAll();
    }
  } catch (err) {
    openDataDlg();
    $("ioMsg").textContent = "data\\取り込み_自社機械予定表.js がありません。tools\\Excelから取り込む.py で作るか、「ファイルから読み込む」で .json を選んでください。";
  }
}
$("bData").onclick = () => openDataDlg();
function openDataDlg() {
  $("ioWhere").innerHTML = DEMO ? "いまは<b>このブラウザの中だけ</b>に覚えています（試し。ほかの人とは共有しません）。"
    : "データは会社の OneDrive（社内共有フォルダ\\03_自社機械稼働表\\稼働データ\\稼働データ.json）に保存し、全員で共有しています。" +
      (window.FOLDER_ITEM_ID ? `<br>稼働データ フォルダの itemId：<b>${esc(window.FOLDER_ITEM_ID)}</b>（config.js に入れて公開し直す）` : "") +
      "<br>" + xlNote();
  $("ioImport").hidden = !window.KIKAI_IMPORT;
  $("ioClear").hidden = !DEMO;   // 共有しているときは全部消すを出さない
  $("ioMsg").textContent = `いま：機械 ${D.機械.length}台・現場 ${D.現場.length}・予約 ${D.予約.length}本` + (D.更新 ? `　最後の変更：${D.更新.いつ} ${D.更新.だれ}` : "");
  $("dData").showModal();
};
$("ioClose").onclick = () => $("dData").close();
$("ioImport").onclick = importBundled;
$("ioOpen").onclick = () => $("ioFile").click();
$("ioFile").onchange = async () => {
  const f = $("ioFile").files[0]; $("ioFile").value = ""; if (!f) return;
  try { loadData(JSON.parse(await f.text()), f.name); } catch (err) { $("ioMsg").textContent = "読めませんでした：" + err.message; }
};
$("ioSave").onclick = () => download(`自社機械稼働表_${todayKey()}.json`, JSON.stringify(D, null, 1));
$("ioPrint").onclick = () => { $("dData").close(); if (S.tab !== "chart") setTab("chart"); setTimeout(() => print(), 50); };
$("ioClear").onclick = () => { if (confirm("データを全部消します（「↶ 戻す」で戻せます）。よろしいですか？")) { commit(replaceAll(emptyData())); $("ioMsg").textContent = "消しました。"; } };

/* 印刷：表の幅を用紙（A3 横）に合わせて縮める */
addEventListener("beforeprint", () => {
  const c = $("chart"); const w = c.scrollWidth, pageW = 1520;   // A3 横の印刷幅（96dpi で約 1530px）
  c.style.setProperty("--pz", Math.min(1, pageW / w).toFixed(3));
});

/* 版の表示（公開するとき tools\公開する.py が __VER__ を「Vr001 10/24 19:15」のように書き換える） */
if ($("ver").textContent === "__VER__") $("ver").textContent = "開発中";
const MY_VER = $("ver").textContent;
async function checkVersion(say) {   // 公開中の index.html の版と比べ、違えば下に「更新する」の帯（全体工程表と同じ）
  if (!/^Vr/.test(MY_VER)) return;
  try {
    const t = await (await fetch("index.html?check=" + Date.now(), { cache: "no-store" })).text();
    const m = t.match(/id="ver"[^>]*>([^<]+)</), v = m && m[1].trim();
    if (v && v !== MY_VER && /^Vr/.test(v)) { $("updVer").textContent = v; $("updBar").hidden = false; }
    else if (say) toast("いまの画面が最新の版です（" + MY_VER + "）");
  } catch (e) { console.warn("版の確認に失敗", e); }
}
$("bUpd").onclick = () => { const q = new URLSearchParams(location.search); q.set("r", Date.now()); location.replace(location.pathname + "?" + q.toString()); };
$("ver").onclick = () => checkVersion(true);
setInterval(checkVersion, 5 * 60 * 1000);
setTimeout(checkVersion, 3000);

async function start() {
  document.body.dataset.tab = S.tab;
  if (DEMO) {
    $("who").textContent = location.protocol === "file:" ? "このパソコンの中だけ（試し）" : "試し（このブラウザの中だけ。OneDrive には書きません）";
  } else {
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: new URL("./", location.href).href },
      cache: { cacheLocation: "localStorage" }
    });
    await msalApp.initialize();
    try { const r = await msalApp.handleRedirectPromise(); if (r) msalApp.setActiveAccount(r.account); }
    catch (e) { $("signinMsg").textContent = "サインインできませんでした：" + e.message; }
    const acc = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
    if (!acc) { $("signin").hidden = false; $("bSignIn").onclick = () => msalApp.loginRedirect({ scopes: SCOPES }); return; }
    msalApp.setActiveAccount(acc);
    me.name = acc.name || acc.username;
    $("who").textContent = me.name;
  }
  $("loading").hidden = false;
  try {
    if (!DEMO) { await resolveFolder(); await loadPrices(); }
    const r = await store.get();
    if (r) { D = normalize(r.data); docTag = r.eTag; }
  } catch (e) {
    console.error(e);
    $("loading").hidden = true;
    $("signin").hidden = false; $("bSignIn").hidden = true;
    $("signinMsg").textContent = (e.status === 404 || e.status === 403) ? "稼働データ フォルダを開けません。北澤さんに共有（編集可）を頼んでください。" : "読み込めませんでした：" + e.message;
    return;
  }
  $("loading").hidden = true;
  ready = true; showSync();
  setTab(S.tab);
  setInterval(() => { if (!document.hidden) pull(); }, POLL_MS);
  setInterval(() => { if (!document.hidden) loadPrices(); }, 5 * 60 * 1000);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) { pull(); checkVersion(); } });
}
start();
