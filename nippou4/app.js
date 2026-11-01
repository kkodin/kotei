/* 工事日報 Vr004（2026-11-02）：Vr003 にクロジカの形のカレンダー・詳細／編集の窓・人別／現場別の集計・請負形態・自社機械稼働表の予約の連動を足した（cal.js）。
   データは Vr003 と同じ（形は足すだけ）。以下は Vr003 の説明。
   工事日報 Vr003（2026-11-01 作り直し）
   会社の標準の日報。サインインした人＝書く人が、毎日「どこで・何をしたか」を書く（私の日報）。
   ・1日を「行った所」のカードで書く：行先（現場・本社など）・始め〜終わり・作業内容・来ていた人（作業員の名前／外注は業者名と人数）。
     本社だけの日は1枚、いくつも回った日はカードを足す。詳しく書かない人はここまで。
   ・現場のカードから「詳しく書く」で常用日報（kotei/jouyou）を開き、この日・この現場に作業内容と人を入れた形で詳しく書く。
   ・画面は PC・iPad・iPhone SE で1つ（幅で並びが変わる）。ボタンは下の3つ＋「…」だけ。ホーム画面に置くとアドレスバーが出ない。
   ・データ：会社の OneDrive の 日報のフォルダ（config の nippou）に 個人_<名前>_<YYYY-MM>.json（1人1か月1ファイル）。マスタは読むだけ。
     365 のアカウントが無い人はゲストに招き、日報のフォルダだけを共有する（工程データは見せない）。ゲストは同じフォルダの 名簿.json
     （工程データのマスタ・現場一覧の名前だけの写し。工程データを読める人が開くたびに作り直す）で名前・現場を選ぶ。
     config に nippou が無いときは、これまでどおり 工程データ\日報\ に書く。 */
"use strict";
const $ = id => document.getElementById(id);
const CFG = window.NIPPOU_CONFIG;
const SCOPES = ["User.Read", "Files.ReadWrite.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const DEMO = new URLSearchParams(location.search).has("demo");
const NDIR = "日報";
const SEP = DEMO || !!(CFG.nippou && CFG.nippou.itemId);   // 日報だけ別のフォルダ（ゲストに共有する）
const ROSTER = "名簿.json";                                  // そのフォルダに置く名簿・現場一覧の写し
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const pad2 = n => String(n).padStart(2, "0");
const keyOf = d => d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
const dateOf = k => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
const dayAdd = (k, n) => { const d = dateOf(k); d.setDate(d.getDate() + n); return keyOf(d); };
const todayKey = () => keyOf(new Date());
const WD = "日月火水木金土";
const monthOf = k => k.slice(0, 7);
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const nowStr = () => new Date().toLocaleString("ja-JP");
// 名前を比べる形（空白・全角半角・株式会社などのちがいを無視）
const nk = s => String(s ?? "").normalize("NFKC").replace(/株式会社|\(株\)|有限会社|\(有\)/g, "").replace(/[\s　・]/g, "").toLowerCase();
const OFFICE = ["本社", "休み"];      // 現場ではない行先（人・外注の欄を出さない）
const isOffice = v => OFFICE.includes(String(v || "").trim()) || /^本社/.test(String(v || "").trim());

/* ブラウザの自動入力（住所・会社名・氏名の候補）を出さない（2026-11-01 ユーザー：表のセルに Chrome の住所の候補が出た）。
   Chrome は autocomplete="off" を無視して住所の候補を出すことがある。1つの form の中の欄が3つ未満だと住所の自動入力をしないので、
   form の外の欄を1つずつ form（display:contents。見た目は変わらない）で包む。あとから足された欄も包む（いま打っている欄なら打つ所を戻す） */
function noAutofill(root) {
  if (!root || !root.querySelectorAll) return;
  const list = root.matches && root.matches("input,textarea,select") ? [root] : root.querySelectorAll("input,textarea,select");
  list.forEach(el => {
    if (el.closest("form") || /^(checkbox|radio|file|button|submit|hidden|range|color)$/.test(el.type)) return;
    const on = document.activeElement === el, s0 = on && el.selectionStart, s1 = on && el.selectionEnd;
    const f = document.createElement("form"); f.className = "nofill"; f.setAttribute("autocomplete", "off");
    f.addEventListener("submit", e => e.preventDefault());
    el.parentNode.insertBefore(f, el); f.appendChild(el); el.setAttribute("autocomplete", "off");
    if (on) { el.focus({ preventScroll: true }); try { el.setSelectionRange(s0, s1); } catch (e) { } }
  });
}
noAutofill(document.body);
new MutationObserver(ms => ms.forEach(m => m.addedNodes.forEach(n => { if (n.nodeType === 1 && n.parentNode) noAutofill(n); }))).observe(document.body, { childList: true, subtree: true });

/* ---------- OneDrive（全体工程表・Vr002 と同じ読み書き） ---------- */
let msalApp = null;
async function token() {
  try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken; }
  catch (e) { if (e instanceof msal.InteractionRequiredAuthError) await msalApp.acquireTokenRedirect({ scopes: SCOPES }); throw e; }
}
async function g(path, opt = {}) { return fetch(GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + await token(), ...(opt.headers || {}) } }); }
function httpErr(res, what) { const e = new Error(what + " " + res.status); e.status = res.status; return e; }
const enc = p => p.split("/").map(encodeURIComponent).join("/");
function GraphFolder(where, sub) {
  const F = CFG[where], base = () => `/drives/${F.driveId}/items/${F.itemId}`, p = name => base() + ":/" + enc((sub ? sub + "/" : "") + name);
  return {
    async list() {
      const out = []; let url = sub ? base() + ":/" + enc(sub) + ":/children?$top=999" : base() + "/children?$top=999";
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
      const meta = await res.json(), dl = meta["@microsoft.graph.downloadUrl"];
      const body = dl ? await fetch(dl, { cache: "no-store" }) : await g(p(name) + ":/content");
      if (!body.ok) throw httpErr(body, "中身の読み込み");
      return { data: await body.json(), eTag: meta.eTag };
    },
    async put(name, data, eTag) {
      const url = p(name) + ":/content" + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
      const send = () => g(url, { method: "PUT", headers: { "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: JSON.stringify(data) });
      let res = await send();
      if (res.status === 404 && sub && !eTag) {   // 日報 フォルダがまだ無い → 作ってからもう一度
        await g(base() + "/children", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: sub, folder: {}, "@microsoft.graph.conflictBehavior": "replace" }) });
        res = await send();
      }
      if (!res.ok) throw httpErr(res, "保存");
      return { eTag: (await res.json()).eTag };
    }
  };
}
const DEMO_KEY = "nippou_demo_v4";
function DemoFolder(prefix) {
  const all = () => { try { return JSON.parse(localStorage.getItem(DEMO_KEY)) || {}; } catch (e) { return {}; } };
  const save = a => { try { localStorage.setItem(DEMO_KEY, JSON.stringify(a)); } catch (e) { } };
  return {
    async list() { const a = all(); return Object.keys(a).filter(n => n.startsWith(prefix)).map(n => ({ name: n.slice(prefix.length), eTag: a[n].eTag })); },
    async get(name) { const x = all()[prefix + name]; return x ? { data: JSON.parse(JSON.stringify(x.data)), eTag: x.eTag } : null; },
    async put(name, data, eTag) {
      const a = all(), cur = a[prefix + name];
      if (eTag ? (!cur || cur.eTag !== eTag) : cur) { const e = new Error("保存 " + (eTag ? 412 : 409)); e.status = eTag ? 412 : 409; throw e; }
      const t = "d" + Date.now() + Math.random().toString(36).slice(2, 6);
      a[prefix + name] = { data: JSON.parse(JSON.stringify(data)), eTag: t }; save(a);
      return { eTag: t };
    }
  };
}
const GUEST_DEMO = DEMO && new URLSearchParams(location.search).has("guest");   // 試し：工程データを読めないゲストのつもり
const koteiStore = DEMO ? DemoFolder(GUEST_DEMO ? "見えない/" : "工程/") : GraphFolder("kotei");
const nippouStore = DEMO ? DemoFolder("日報/") : SEP ? GraphFolder("nippou") : GraphFolder("kotei", NDIR);
const kikaiStore = DEMO ? DemoFolder(GUEST_DEMO ? "見えない/" : "機械/") : GraphFolder("kikai");

/* ---------- 状態 ---------- */
const S = {
  me: { name: "", acct: "" },   // name＝名簿の名前（書く人）。acct＝サインインしたアカウント
  day: todayKey(),
  F: {},                         // 月 → { doc, eTag, pending:[fn], saving, timer, err, loaded }
  M: { 社員: [], 従業員: [], 外注業者: [], 機械: [] },   // 工程データ\マスタ.json（機械は自社機械稼働表の機械の名前）
  sites: [],                     // 稼働データ\現場一覧.json の 現場（コード・名前・正式名・元請）
  view: "me",                    // me＝私の日報、all＝この日の全員の日報、mysum＝私の月の集計、list＝月の一覧、site＝現場別日報（別ウィンドウ）
  who: "",                       // 月の一覧で見ている人（空＝自分）
  K: { sites: [], rest: [] },    // 全体工程表の 現場_*.json と 休日.json（予定を作る。ゲストは読めない）
  kikai: null,                   // 稼働データ.json（機械の稼働の集計。読めなければ出さない）
  site: null,                    // 現場別日報で見ている現場 { place, code }
  tab: "day",                    // 現場別日報：day＝日ごと、sum＝月の集計
  draft: {}                      // 日 → まだ何も書いていない日に出しておく入力欄の id（打つと行になる）
};
const fileName = m => `個人_${nk(S.me.name)}_${m}.json`;
function blankDoc(m) { return { 形式: "私の日報", 版: 1, 名前: S.me.name, アカウント: S.me.acct, 月: m, 日: {} }; }
const fileOf = m => S.F[m] || (S.F[m] = { doc: blankDoc(m), eTag: null, pending: [], saving: false, timer: null, err: null, loaded: false });
const dayOf = (k = S.day) => { const f = S.F[monthOf(k)]; return f && f.doc.日[k] || null; };
const rowsOf = (k = S.day) => (dayOf(k) || {}).行 || [];
// 作業内容は 予定｜報告 の行（作業行）。前の形（作業 だけ）は報告の1行として読む。作業 には報告をつないで入れておく（常用日報・全員の日報が読む）
const workLines = r => r.作業行 && r.作業行.length ? r.作業行 : [{ 予定: "", 報告: r.作業 || "" }];
const syncWork = y => { y.作業 = (y.作業行 || []).map(l => l.報告).filter(Boolean).join("\n"); };
// 参加者・外注・使用機械は 予定（r.予定.人・外注・機械）と 報告（r.人・外注・機械。前からの 人・外注 はそのまま報告）
const listOf = (y, s, kind) => { const o = s === "p" ? (y.予定 ||= {}) : y; return o[kind] ||= []; };
const listIn = (r, s, kind) => (s === "p" ? (r.予定 || {})[kind] : r[kind]) || [];
// 現場の予定を入れる（作業内容は工程表、外注は工程表の外注の段、機械は自社機械稼働表の予約）
function fillPlan(y, k = S.day) {
  const p = planLines(y.行先, y.コード, k);
  if (!workLines(y).some(l => l.予定 || l.報告)) { y.作業行 = p.length ? p.map(t => ({ 予定: t, 報告: "" })) : [{ 予定: "", 報告: "" }]; syncWork(y); }
  if (!["人", "外注", "機械"].some(x => listIn(y, "p", x).length)) y.予定 = { 人: [], 外注: planGai(y.行先, y.コード, k), 機械: planKikai(y.行先, y.コード, k) };
}
const blankRow = (id, start) => {
  const r = { id, 行先: "", コード: "", 始め: start || "", 終わり: "", 作業: "", 作業行: [{ 予定: "", 報告: "" }], 予定: { 人: [], 外注: [], 機械: [] }, 人: [], 外注: [], 機械: [] };
  if (S.view === "site" && S.site) { r.行先 = S.site.place; r.コード = S.site.code; fillPlan(r); }
  return r;
};
const draftId = (k = S.day) => S.draft[k] || (S.draft[k] = uid());

// 変える：fn をいまの中身に当て、保存する。ほかの端末が先に保存していたら、読み直して fn を当て直す
function mutate(k, fn, quiet) {
  const f = fileOf(monthOf(k));
  const wrap = doc => { if (!doc.日[k]) doc.日[k] = { 行: [], 報告: null }; fn(doc.日[k], doc); const d = doc.日[k]; d.更新 = nowStr(); if (d.報告 && !wrap.noMark) d.報告.直した = nowStr(); doc.更新 = nowStr(); };
  wrap.noMark = !!(quiet && quiet.report);
  wrap(f.doc); f.pending.push(wrap);
  clearTimeout(f.timer); f.timer = setTimeout(() => saveMonth(monthOf(k)), 700);
  showSync();
  if (!quiet || quiet.render) render();
}
async function saveMonth(m) {
  const f = S.F[m];
  if (f.saving) { clearTimeout(f.timer); f.timer = setTimeout(() => saveMonth(m), 500); return; }
  if (!f.pending.length) return;
  f.saving = true; showSync();
  try {
    for (let tries = 0; tries < 5; tries++) {
      const n = f.pending.length;
      try { const r = await nippouStore.put(fileName(m), f.doc, f.eTag); f.eTag = r.eTag; f.pending.splice(0, n); f.err = null; break; }
      catch (e) {
        if (e.status !== 412 && e.status !== 409) throw e;
        const fresh = await nippouStore.get(fileName(m)), d = fresh ? fresh.data : blankDoc(m);
        f.pending.forEach(fn => fn(d)); f.doc = d; f.eTag = fresh ? fresh.eTag : null;
        toast("ほかの端末で書いた分を読み込んで、重ねて保存しました"); render();
      }
    }
  } catch (e) {
    f.err = e.message; console.error(e);
    toast("保存できませんでした：" + e.message + "（自動でやり直します）");
    setTimeout(() => { f.err = null; saveMonth(m); }, 10000);
  } finally {
    f.saving = false; showSync();
    if (f.pending.length && !f.err) { clearTimeout(f.timer); f.timer = setTimeout(() => saveMonth(m), 300); }
  }
}
function showSync() {
  const fs = Object.values(S.F), el = $("sync");
  const err = fs.some(f => f.err), busy = fs.some(f => f.pending.length || f.saving);
  el.className = "sync " + (err ? "err" : busy ? "busy" : "ok");
  el.textContent = err ? "保存できていません" : busy ? "保存中…" : (fs.some(f => f.eTag) ? "保存済み" : "");
}
async function loadMonth(m, force) {
  const f = fileOf(m);
  if (f.loaded && !force) return;
  if (f.pending.length || f.saving) return;   // 書きかけを上書きしない
  try { const r = await nippouStore.get(fileName(m)); if (r && !f.pending.length) { f.doc = r.data; f.eTag = r.eTag; } f.loaded = true; }
  catch (e) { toast("日報を読めませんでした：" + e.message); }
}

/* ---------- マスタ（読むだけ）。日報が別のフォルダなら、そこに名前だけの写し（名簿.json）を置き、ゲストはそれを読む ---------- */
const siteMin = x => ({ コード: x.コード || "", 名前: x.名前 || "", 正式名: x.正式名 || "", 元請: x.元請 || "", 発注者: x.発注者 || "" });
async function loadMasters() {
  const [m, s] = await Promise.all([koteiStore.get("マスタ.json").catch(() => null), kikaiStore.get("現場一覧.json").catch(() => null), loadKotei()]);
  if (m) S.M = { 社員: m.data.社員 || [], 従業員: m.data.従業員 || [], 外注業者: m.data.外注業者 || [] };
  if (s) S.sites = (s.data.現場 || []).filter(x => x && (x.名前 || x.正式名));
  if (S.kikai) { const ms = (S.kikai.機械 || []).filter(m => !m.使わない); S.M.機械 = [...new Set(ms.map(machName))]; S.M.機械分類 = Object.fromEntries(ms.map(m => [machName(m), m.分類 || "その他"])); }
  if (SEP) {
    const cur = await nippouStore.get(ROSTER).catch(() => null);
    if (m) await saveRoster(cur, !!s);
    else if (cur) { const c = cur.data; S.M = { 社員: c.社員 || [], 従業員: c.従業員 || [], 外注業者: c.外注業者 || [], 機械: S.kikai ? S.M.機械 : (c.機械 || []), 機械分類: S.kikai ? S.M.機械分類 : (c.機械分類 || {}) }; S.sites = (c.現場 || []).filter(x => x && (x.名前 || x.正式名)); }
    else if (!m) throw new Error("名簿がまだありません（工程データを読める人が一度開くと作られます）");
  }
  try { localStorage.setItem("nippou_master_v3", JSON.stringify({ M: S.M, sites: S.sites })); } catch (e) { }
}
// 工程データを読めた人が開いたとき：名簿.json を今のマスタに合わせる（変わっていなければ書かない）。現場一覧を読めなかったら現場は前の写しのまま
async function saveRoster(cur, gotSites) {
  const c = cur ? cur.data : {};
  const body = { 社員: S.M.社員, 従業員: S.M.従業員, 外注業者: S.M.外注業者, 機械: S.kikai ? S.M.機械 : (c.機械 || []), 機械分類: S.kikai ? S.M.機械分類 : (c.機械分類 || {}), 現場: gotSites ? S.sites.map(siteMin) : (c.現場 || []) };
  if (cur && JSON.stringify(body) === JSON.stringify({ 社員: c.社員, 従業員: c.従業員, 外注業者: c.外注業者, 機械: c.機械, 機械分類: c.機械分類, 現場: c.現場 })) return;
  const doc = { 形式: "工事日報の名簿", 版: 1, 説明: "工程データのマスタ・現場一覧の名前だけの写し。工事日報が自動で作り直す（ここを直しても戻る）", ...body, 更新: nowStr() };
  try { await nippouStore.put(ROSTER, doc, cur ? cur.eTag : null); } catch (e) { console.warn("名簿の写しを保存できませんでした", e); }
}
// 全体工程表（現場_*.json・休日.json）と自社機械稼働表（稼働データ.json）。読めない人（ゲスト）は予定・機械なしで動く
async function loadKotei() {
  try {
    const want = (await koteiStore.list()).filter(f => /^現場_.*\.json$/.test(f.name) || f.name === "休日.json");
    const got = await Promise.all(want.map(f => koteiStore.get(f.name).catch(() => null)));
    const K = { sites: [], rest: [] };
    want.forEach((f, i) => { const d = got[i] && got[i].data; if (!d) return; if (f.name === "休日.json") K.rest = d.休み || []; else if (!d.消した) K.sites.push(d); });
    S.K = K;
  } catch (e) { }
  try { const r = await kikaiStore.get("稼働データ.json"); S.kikai = r ? r.data : null; } catch (e) { S.kikai = null; }
}
function cachedMasters() { try { const c = JSON.parse(localStorage.getItem("nippou_master_v3")); if (c) { S.M = c.M; S.sites = c.sites; } } catch (e) { } }
const people = () => [...S.M.社員, ...S.M.従業員];
const siteName = x => x.名前 || x.正式名;

/* ---------- 全体工程表から、その日・その現場の予定（作業）を作る（全体工程表・Vr002 と同じ決まり） ---------- */
const JP_HOL = {};
const dow = k => dateOf(k).getDay();
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
const holidayOf = k => jpHolidays(+k.slice(0, 4))[k] || "";
function ruleWorks(ln, k) {
  if (holidayOf(k)) return ln.祝 === true;
  const w = dow(k);
  if (w === 0) return ln.日 === true;
  if (w === 6) return ln.土 !== false;
  return true;
}
function barWorks(ln, b, k) {
  if (S.K.rest.includes(k)) return false;
  if (b && (b.休み || []).includes(k)) return false;
  if (b && (b.出勤 || []).includes(k)) return true;
  return !!ln && ruleWorks(ln, k);
}
const laneKubun = ln => !ln ? "昼" : ln.区分 || (/夜/.test(ln.名前 || "") ? "夜" : "昼");
const laneWork = ln => !ln ? "" : ln.区分 ? (ln.名前 || "") : String(ln.名前 || "").replace(/昼間|夜間/g, "").trim();
// 行先（名前・コード）→ 全体工程表の現場
function koteiSite(place, code) {
  const n = nk(place);
  return S.K.sites.find(s => code && s.コード && String(s.コード) === String(code)) ||
    S.K.sites.find(s => n && (nk(s.名前) === n || nk(s.略称) === n)) || null;
}
// その日の予定（段の作業名：バーの作業（夜）／外注 業者）
function planLines(place, code, k) {
  const s = koteiSite(place, code), out = [];
  if (!s) return out;
  for (const ln of s.段 || []) for (const b of (s.バー || []).filter(b => b.段 === ln.id && b.開始 <= k && b.終了 >= k && barWorks(ln, b, k))) {
    const w = laneWork(ln), a = b.作業 || "";
    let t = w && a && w !== a ? `${w}：${a}` : (w || a);
    if (laneKubun(ln) === "夜") t += "（夜）";
    if (ln.種別 === "外注") t += `　外注 ${ln.業者 || ""}`;
    if (t && !out.includes(t)) out.push(t);
  }
  return out;
}
// その日に動く、工程表の外注の段の業者（人数は工程表に無いので空）
function planGai(place, code, k) {
  const s = koteiSite(place, code), out = [];
  if (!s) return out;
  for (const ln of s.段 || []) if (ln.種別 === "外注" && ln.業者 && !out.some(g => g.業者 === ln.業者) &&
    (s.バー || []).some(b => b.段 === ln.id && b.開始 <= k && b.終了 >= k && barWorks(ln, b, k))) out.push({ 業者: ln.業者, 人数: "" });
  return out;
}
const machName = mc => mc.名前 + (mc.規格 ? " " + mc.規格 : "");
const kikaiNames = (place, code) => { const ks = koteiSite(place, code); return [place, ks && ks.名前, ks && ks.略称].filter(Boolean).map(nk); };
// その日、この現場に予約してある自社の機械（仮は入れない）
function planKikai(place, code, k) {
  if (!S.kikai) return [];
  const names = kikaiNames(place, code), out = [];
  (S.kikai.予約 || []).forEach(y => { if (y.状態 === "仮" || !names.includes(nk(y.現場)) || !kikaiWorks(y, k)) return;
    const mc = (S.kikai.機械 || []).find(x => x.id === y.機械); if (mc && !out.includes(machName(mc))) out.push(machName(mc)); });
  return out;
}
// 自社機械稼働表の予約がその日に使うか（自社機械稼働表と同じ）
function kikaiWorks(r, k) {
  if (k < r.開始 || k > r.終了) return false;
  if ((r.未使用 || []).includes(k)) return false;
  if ((r.出勤 || []).includes(k)) return true;
  if ((r.休み || []).includes(k)) return false;
  if (holidayOf(k) || dow(k) === 0) return !!r.日祝;
  if (dow(k) === 6) return !!r.土;
  return true;
}
// 行先 → 現場一覧（受注済みの現場）の現場
function siteOf(place, code) {
  const n = nk(place);
  return S.sites.find(x => code && x.コード && String(x.コード) === String(code)) || S.sites.find(x => n && (nk(x.名前) === n || nk(x.正式名) === n)) || null;
}

/* ---------- 選ぶ欄（打てば絞り込み、押せば選ぶ。スマホでも押しやすい大きさ） ---------- */
let sugBox = null;
function closeSug() { if (sugBox) { sugBox.remove(); sugBox = null; } }
function showSug(input, items, onPick) {
  closeSug();
  const q = nk(input.value), hit = items.filter(x => !q || nk(x.v).includes(q) || nk(x.sub || "").includes(q)).slice(0, 40);
  if (!hit.length) return;
  sugBox = document.createElement("div"); sugBox.className = "sug";
  let grp = null;
  hit.forEach(x => {
    if (x.grp && x.grp !== grp) { grp = x.grp; const h = document.createElement("div"); h.className = "grp"; h.textContent = grp; sugBox.appendChild(h); }
    const b = document.createElement("button"); b.type = "button";
    b.innerHTML = esc(x.v) + (x.sub ? `<small>${esc(x.sub)}</small>` : "");
    b.addEventListener("pointerdown", e => e.preventDefault());   // 押したときに打つ欄から外れない
    b.addEventListener("click", () => { closeSug(); onPick(x); });
    sugBox.appendChild(b);
  });
  input.parentElement.appendChild(sugBox);
}
function bindPick(input, items, onPick) {
  const open = () => showSug(input, items(), onPick);
  input.addEventListener("focus", open);
  input.addEventListener("input", open);
  input.addEventListener("blur", () => setTimeout(closeSug, 150));
  input.addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); const b = sugBox && sugBox.querySelector("button"); if (b && input.value.trim()) b.click(); else { closeSug(); onPick({ v: input.value.trim(), free: true }); } }
    if (e.key === "Escape") closeSug();
  });
}
// 行先の候補：最近行った所（30日）→ 本社・休み → 全部の現場（現場一覧）
function placeItems() {
  const recent = [], seen = new Set();
  for (let i = 0; i <= 30; i++) rowsOf(dayAdd(S.day, -i)).forEach(r => { const k = nk(r.行先); if (r.行先 && !seen.has(k)) { seen.add(k); recent.push({ grp: "最近行った所", v: r.行先, コード: r.コード || "" }); } });
  return [...recent, ...OFFICE.filter(v => !seen.has(nk(v))).map(v => ({ grp: "本社など", v })),
    ...S.sites.filter(x => !seen.has(nk(siteName(x)))).map(x => ({ grp: "現場", v: siteName(x), sub: [x.コード, x.元請 === "自社請け" ? "自社請け " + (x.発注者 || "") : x.元請].filter(Boolean).join(" "), コード: x.コード || "" }))];
}
// 人の候補：この行先に前に来ていた人 → 社員・従業員。PICK.人 で 職員（社員）・作業員（従業員）に絞る
const PICK = { 人: "", 機械: "" };   // 選ぶ欄の絞り込み（参加者＝職員／作業員、機械＝稼働表の分類）
function personItems(place, have) {
  const all = personItems0(place, have), f = PICK.人;
  if (!f) return all;
  const L = new Set((f === "職員" ? S.M.社員 : S.M.従業員).map(nk));
  return all.filter(x => L.has(nk(x.v)));
}
function personItems0(place, have) {
  const hv = new Set(have.map(nk)), out = [], seen = new Set();
  for (let i = 1; i <= 30; i++) rowsOf(dayAdd(S.day, -i)).forEach(r => { if (nk(r.行先) === nk(place)) (r.人 || []).forEach(n => { const k = nk(n); if (!hv.has(k) && !seen.has(k)) { seen.add(k); out.push({ grp: "前にこの現場に来ていた人", v: n }); } }); });
  S.M.社員.forEach(n => { const k = nk(n); if (!hv.has(k) && !seen.has(k)) { seen.add(k); out.push({ grp: "社員", v: n }); } });
  S.M.従業員.forEach(n => { const k = nk(n); if (!hv.has(k) && !seen.has(k)) { seen.add(k); out.push({ grp: "従業員", v: n }); } });
  return out;
}
const gaiItems = () => S.M.外注業者.map(v => ({ v }));
// 機械の候補：前に使った機械 → 自社の機械（自社機械稼働表）。候補にない機械（リースなど）も書ける
function machItems(have) {
  const hv = new Set(have.map(nk)), out = [], seen = new Set();
  const add = (grp, v) => { const k = nk(v); if (v && !hv.has(k) && !seen.has(k)) { seen.add(k); out.push({ grp, v }); } };
  for (let i = 1; i <= 30; i++) rowsOf(dayAdd(S.day, -i)).forEach(r => (r.機械 || []).forEach(v => add("前に使った機械", v)));
  (S.M.機械 || []).forEach(v => add("自社の機械　" + ((S.M.機械分類 || {})[v] || "その他"), v));
  return PICK.機械 ? out.filter(x => ((S.M.機械分類 || {})[x.v] || "その他") === PICK.機械) : out;
}

/* ---------- 時刻（「8」→8:00、「830」→8:30、「21.5」→21:30） ---------- */
function normTime(v) {
  v = String(v || "").normalize("NFKC").replace(/[時]/g, ":").replace(/分/g, "").trim();
  if (!v) return "";
  let m = /^(\d{1,2})[:.](\d{1,2})$/.exec(v);
  if (m && v.includes(".") && m[2].length === 1) return `${+m[1]}:${pad2(Math.round(+("0." + m[2]) * 60))}`;
  if (m) return `${+m[1]}:${pad2(+m[2])}`;
  if (/^\d{1,2}$/.test(v)) return `${+v}:00`;
  if (/^\d{3,4}$/.test(v)) return `${+v.slice(0, -2)}:${v.slice(-2)}`;
  return v;
}

// PC：欄を押してからホイール（左半分で時、右半分で分）か ↑↓ キー。打ってもよい。スマホ・タブレット：押すと回して選ぶ窓
const COARSE = matchMedia("(pointer: coarse)").matches;
function stepTime(v, part, d) {
  let m = tmin(normTime(v)); if (m == null) return null;
  m = part === "h" ? m + d * 60 : Math.round(m / 5) * 5 + d * 5;
  m = ((m % 1440) + 1440) % 1440;
  return `${Math.floor(m / 60)}:${pad2(m % 60)}`;
}
function bindTime(el, def, save) {
  el.addEventListener("change", () => { const v = normTime(el.value); el.value = v; save(v); });
  if (COARSE) { el.readOnly = true; el.setAttribute("inputmode", "none"); el.addEventListener("click", () => drum(el.value || def(), v => { el.value = v; save(v, true); })); return; }
  const mark = p => { const i = el.value.indexOf(":"); if (i < 0) return; try { p === "h" ? el.setSelectionRange(0, i) : el.setSelectionRange(i + 1, el.value.length); } catch (e) { } };
  const bump = (p, d) => { const cur = normTime(el.value), v = cur ? stepTime(cur, p, d) : def(); if (!v) return; el.value = v; mark(p); save(v); };
  el.addEventListener("wheel", e => { if (document.activeElement !== el) return; e.preventDefault(); const b = el.getBoundingClientRect(); bump(e.clientX < b.left + b.width / 2 ? "h" : "m", e.deltaY < 0 ? 1 : -1); }, { passive: false });
  el.addEventListener("keydown", e => { if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return; e.preventDefault(); const i = el.value.indexOf(":"); bump(i < 0 || el.selectionStart <= i ? "h" : "m", e.key === "ArrowUp" ? 1 : -1); });
}
// 回して選ぶ窓（時と分の2列。分は5分ごと）
function drum(v, done) {
  const m = tmin(normTime(v)) ?? 480, H = 44;
  const col = (n, f) => `<div class="dcol">${Array.from({ length: n }, (_, i) => `<div>${f(i)}</div>`).join("")}</div>`;
  dlg(`<h2>時間</h2><div class="drum">${col(24, i => i)}<b>:</b>${col(12, i => pad2(i * 5))}</div>
    <div class="btns"><button id="dClr">消す</button><button data-close>やめる</button><button class="pri" id="dOk">決める</button></div>`);
  const [ch, cm] = document.querySelectorAll("#dlgBody .dcol");
  ch.scrollTop = Math.floor(m / 60) * H; cm.scrollTop = (Math.round((m % 60) / 5) % 12) * H;
  [ch, cm].forEach(c => c.addEventListener("click", e => { const d = e.target.closest(".dcol > div"); if (d) c.scrollTo({ top: [...c.children].indexOf(d) * H, behavior: "smooth" }); }));
  const val = () => `${Math.min(23, Math.round(ch.scrollTop / H))}:${pad2(Math.min(11, Math.round(cm.scrollTop / H)) * 5)}`;
  $("dOk").onclick = () => { const x = val(); closeDlg(); done(x); };
  $("dClr").onclick = () => { closeDlg(); done(""); };
}

/* ---------- 画面 ---------- */
function render() {
  const k = S.day, d = dateOf(k);
  $("dLabel").textContent = `${d.getMonth() + 1}/${d.getDate()}(${WD[d.getDay()]})`;
  $("inDay").value = k;
  $("bToday").hidden = k === todayKey();
  $("bMe").textContent = "👤 " + (S.me.name || "名前");
  const mc = document.querySelector('#menu [data-m="cards"]'); if (mc) mc.textContent = `${S.me.name || ""} の日報（カードで書く）`;
  $("bAdd").textContent = S.view === "site" ? "＋ この現場に書く" : "＋ 現場名(行先)を追加";
  $("bSite").hidden = S.view === "site";
  $("app").classList.toggle("wide", S.view === "cal");
  if (S.view === "cal") $("bAdd").textContent = "＋ 予定・報告を登録";
  if (S.view === "all") return renderAll();
  if (S.view === "mysum") return renderMySum();
  if (S.view === "cal") return renderCal();
  if (S.view === "list") return renderList();
  $("empty").hidden = true;
  if (S.view === "site") renderSite();
  else {
    $("all").hidden = true; $("cards").hidden = false;
    const rows = rowsOf();
    $("dayHead").innerHTML = dayBadge() + (rows.length ? `<span>${rows.length} か所</span>` : "");
    drawCards(rows.length ? rows : [blankRow(draftId())], !rows.length);
  }
  const day = dayOf(), rows = rowsOf();
  const done = day && day.報告 && !day.報告.直した;
  $("bReport").textContent = done ? "報告済み ✓" : day && day.報告 ? "もう一度報告" : "報告する";
  $("bReport").classList.toggle("done", !!done);
  $("bReport").disabled = !rows.length;
}
function dayBadge() {
  const day = dayOf(), rows = rowsOf();
  return day && day.報告 ? (day.報告.直した ? `<span class="badge done">報告済み（あとで直した）</span>` : `<span class="badge done">報告済み</span>`) : rows.length ? `<span class="badge draft">書きかけ</span>` : `<span class="badge none">まだ</span>`;
}
// カードを描く。まだ何も書いていない日は、空の入力欄（draft。打つと行になる）を1枚出しておく
function drawCards(list, draft) {
  const focus = document.activeElement && document.activeElement.dataset.f ? { id: document.activeElement.closest(".card")?.dataset.id, f: document.activeElement.dataset.f } : null;
  $("cards").innerHTML = list.map((r, i) => cardHTML(r, i, draft)).join("");
  list.forEach(r => bindCard(r.id));
  document.querySelectorAll("#cards textarea").forEach(grow);
  if (focus) { const el = document.querySelector(`.card[data-id="${focus.id}"] [data-f="${focus.f}"]`); if (el) el.focus(); }
}
function cardHTML(r, i, draft) {
  const off = isOffice(r.行先), st = !off && r.行先 ? siteOf(r.行先, r.コード) : null, L = workLines(r);
  const info = st ? `<div class="pinfo">${siteInfo(r.行先, r.コード)}</div>` : !off && r.行先 ? `<div class="pinfo">現場一覧にない行先（このまま書けます）</div>` : "";
  const work = L.map((l, j) => `<div class="wr"><textarea data-f="wp${j}" data-wj="${j}" data-wk="予定" rows="1" placeholder="予定">${esc(l.予定 || "")}</textarea><textarea data-f="wr${j}" data-wj="${j}" data-wk="報告" rows="1" placeholder="したこと">${esc(l.報告 || "")}</textarea>${L.length > 1 ? `<button type="button" data-delw="${j}" aria-label="この作業を消す">×</button>` : "<span></span>"}</div>`).join("");
  const hasPR = ["人", "外注", "機械"].some(x => listIn(r, "p", x).length || listIn(r, "r", x).length);
  const side = s => chipsBox(r, s, "人") + gaiBox(r, s) + chipsBox(r, s, "機械");
  const tip = COARSE ? "" : ` title="押してからホイール（左で時・右で分）か ↑↓ キーで変える。打ってもよい"`;
  return `<div class="card" data-id="${r.id}">
    <div class="ch"><span class="no">${i + 1}</span><span class="kind">${esc(r.コード ? "現場 " + r.コード : off ? "" : r.行先 ? "現場" : "")}</span><span class="sp"></span>${draft ? "" : `<button type="button" data-act="del">消す</button>`}</div>
    <div class="row2">
      <div class="f"><label>現場名（行先）</label><div class="pick"><input data-f="place" value="${esc(r.行先 || "")}" placeholder="現場名・本社（打つと候補。自由に書いてもよい）" autocomplete="off"></div>${info}${honshaBox(r)}</div>
      <div class="f"><label>時間</label><div class="time"><input class="tm" data-f="t0" value="${esc(r.始め || "")}" inputmode="decimal" placeholder="始め"${tip}><span>〜</span><input class="tm" data-f="t1" value="${esc(r.終わり || "")}" inputmode="decimal" placeholder="終わり"${tip}></div>
        ${kubunBox(r)}${r.始め || r.終わり ? "" : `<div class="time quick" style="margin-top:6px"><button type="button" data-tq="8:00|17:00">8:00〜17:00</button><button type="button" data-tq="21:00|6:00">21:00〜6:00</button></div>`}</div>
    </div>
    <div class="f"><label>作業内容</label><div class="wl"><div class="wh"><span>予定</span><span>報告（したこと）</span><span></span></div>${work}</div>
      <button type="button" class="addline" data-act="addw">＋ 作業内容を追加</button></div>
    ${off && !hasPR ? "" : `<div class="pr">
      <div class="side plan"><h5>予定</h5>${side("p")}</div>
      <div class="side rep"><h5>報告${["人", "外注", "機械"].some(x => listIn(r, "p", x).length) ? `<button type="button" class="same" data-act="copyp">予定と同じにする</button>` : ""}</h5>${side("r")}</div>
    </div>`}
    ${!off && r.行先 ? `<div class="detail">${S.view === "site" ? `<button type="button" data-act="detail" title="常用日報の形で、始業終業・工種・機械・材料などを書く">常用日報で詳しく書く →</button>` : `<button type="button" data-act="site">この現場の日報・集計（現場別）→</button>`}</div>` : ""}
  </div>`;
}
// 参加者名・使用機械（チップ。打つか選んで増やす）／外注業者名＋人数（1行ずつ増やす）
function chipsBox(r, s, kind) {
  const L = listIn(r, s, kind), man = kind === "人";
  const cats = man ? ["職員", "作業員"] : [...new Set(Object.values(S.M.機械分類 || {}))].sort((a, b) => MCATS.indexOf(a) - MCATS.indexOf(b));
  const sel = cats.length ? `<select class="pf" data-pf="${kind}" aria-label="絞る"><option value="">${man ? "全員" : "全部の機械"}</option>${cats.map(c => `<option${PICK[kind] === c ? " selected" : ""}>${esc(c)}</option>`).join("")}</select>` : "";
  return `<div class="f"><label>${man ? "参加者名" : "使用機械"}</label>${L.length ? `<div class="chips">${L.map((n, j) => `<span class="chip">${esc(n)}<button type="button" data-dc="${s}|${kind}|${j}" aria-label="${esc(n)} を外す">×</button></span>`).join("")}</div>` : ""}` +
    `<div class="pickrow">${sel}<div class="pick"><input data-f="c${s}${kind}" data-ci="${s}|${kind}" placeholder="${man ? "＋ 参加者を追加（名前を打つか選ぶ）" : "＋ 機械を追加（打つか選ぶ）"}" autocomplete="off"></div></div></div>`;
}
const ROLES = ["作業員", "職長", "世話役"];
const MCATS = ["車両", "重機", "小型機械", "測量", "その他"];
function gaiBox(r, s) {
  return `<div class="f"><label>外注業者名＋人数</label>` + listIn(r, s, "外注").map((x, j) => `<div class="gaibox"><div class="gai"><div class="pick"><input data-f="g${s}${j}" data-gi="${s}|${j}" value="${esc(x.業者)}" placeholder="業者名" autocomplete="off"></div><input class="n" data-f="n${s}${j}" data-gn="${s}|${j}" value="${esc(x.人数 ?? "")}" inputmode="numeric" placeholder="人数"><span>名</span><button type="button" data-delg="${s}|${j}" aria-label="外す">×</button></div>
    <div class="gmem">${(x.名前 || []).map((m, k) => `<span class="chip">${esc(m.名前)}<select data-grole="${s}|${j}|${k}" aria-label="役">${ROLES.map(v => `<option${(m.役 || "作業員") === v ? " selected" : ""}>${v}</option>`).join("")}</select><button type="button" data-gdel="${s}|${j}|${k}" aria-label="外す">×</button></span>`).join("")}
      <div class="pick"><input data-f="m${s}${j}" data-gm="${s}|${j}" placeholder="＋ 作業員の名前（わかれば）" autocomplete="off"></div></div></div>`).join("") +
    `<button type="button" class="addline" data-addg="${s}">＋ 外注業者を追加</button></div>`;
}
// 本社：本社だけ／本社＋現場名（本社でその現場の仕事をした）。行.関連現場 = { 名前, コード }
const isHonsha = v => /^本社/.test(String(v || "").trim());
function honshaBox(r) {
  if (!isHonsha(r.行先)) return "";
  const on = !!r.関連現場;
  return `<div class="honsha"><label class="ck"><input type="radio" name="hs${r.id}" data-hs="0"${on ? "" : " checked"}> 本社だけ</label><label class="ck"><input type="radio" name="hs${r.id}" data-hs="1"${on ? " checked" : ""}> 本社＋現場名</label>
    ${on ? `<div class="pick"><input data-f="rel" data-rel="1" value="${esc(r.関連現場.名前 || "")}" placeholder="どの現場の仕事か（打つと候補）" autocomplete="off"></div>` : ""}</div>`;
}
// 外注の作業員名の候補：前にこの業者で来ていた人（あとで NAS の安全書類の作業員名簿も）
function gaiMemItems(gyosha, have) {
  const hv = new Set(have.map(nk)), out = [], seen = new Set();
  for (let i = 0; i <= 60; i++) rowsOf(dayAdd(S.day, -i)).forEach(r => ["p", "r"].forEach(s => listIn(r, s, "外注").forEach(g => { if (nk(g.業者) !== nk(gyosha)) return; (g.名前 || []).forEach(m => { const k = nk(m.名前); if (!hv.has(k) && !seen.has(k)) { seen.add(k); out.push({ grp: "前に来ていた人", v: m.名前, 役: m.役 }); } }); })));
  return out;
}
// 昼勤・夜勤（行.区分。無ければ始めの時刻で決める＝17時〜5時に始めたら夜）
const kubunOf = r => r.区分 || ((t => t != null && (t >= 17 * 60 || t < 5 * 60))(tmin(r.始め)) ? "夜" : "昼");
const kubunBox = r => `<div class="kbn">${["昼", "夜"].map(v => `<label class="ck"><input type="radio" name="kb${r.id}" data-kb="${v}"${kubunOf(r) === v ? " checked" : ""}> ${v}勤</label>`).join("")}</div>`;
// 同じ日・同じ時間帯（昼夜が同じで、時刻があれば重なる）に、ほかの日報や自社機械稼働表の予約に入っている人・機械 → 尋ねる
function spanOf(r) { const a = tmin(r.始め), b = tmin(r.終わり); if (r.終日 || a == null || b == null) return null; return [a, b <= a ? b + 1440 : b]; }
function sameSlot(r1, r2) { if (kubunOf(r1) !== kubunOf(r2)) return false; const a = spanOf(r1), b = spanOf(r2); return !a || !b || (a[0] < b[1] && b[0] < a[1]); }
async function clashes(k, r, what) {
  const out = [], slot = x => `（${kubunOf(x)}勤${x.始め ? " " + x.始め + "〜" + (x.終わり || "") : ""}）`;
  let docs = [];
  try { docs = await monthDocs(monthOf(k)); } catch (e) { }
  evsOn(docs, k).forEach(x => {
    if (x.r.id === r.id || !sameSlot(x.r, r)) return;
    const mem = membersOf(x), mac = listIn(x.r, "r", "機械").length ? listIn(x.r, "r", "機械") : listIn(x.r, "p", "機械");
    (what.人 || []).forEach(n => { if (mem.some(v => nk(v) === nk(n))) out.push(`・${n}：${x.who} の日報「${titleOf(x.r)}」${slot(x.r)}に入っています`); });
    (what.機械 || []).forEach(n => { if (mac.some(v => nk(v) === nk(n))) out.push(`・${n}：${x.who} の日報「${titleOf(x.r)}」${slot(x.r)}で使っています`); });
  });
  if (S.kikai && r.行先) {
    const names = kikaiNames(r.行先, r.コード);
    (what.機械 || []).forEach(n => { const mc = (S.kikai.機械 || []).find(m => nk(machName(m)) === nk(n)); if (!mc) return;
      const y = (S.kikai.予約 || []).find(y => y.機械 === mc.id && y.状態 !== "仮" && (y.区分 || "昼") === kubunOf(r) && kikaiWorks(y, k) && !names.includes(nk(y.現場)));
      if (y) out.push(`・${n}：自社機械稼働表で「${y.現場}」（${y.区分 || "昼"}）に予約があります`); });
  }
  return [...new Set(out)];
}
async function okClash(k, r, what) {
  const m = await clashes(k, r, what);
  return !m.length || confirm(`同じ時間帯に、ほかでも入っています。\n\n${m.join("\n")}\n\nこのままでよいですか？`);
}
// 参加者・機械の絞り込み、外注の作業員名と役、本社＋現場名、昼勤・夜勤。apply(fn, rerender) で行を直す（カードと編集の窓で同じ）
function bindExtras(root, get, apply, refocus) {
  root.querySelectorAll("[data-kb]").forEach(el => el.onchange = () => { const v = el.dataset.kb; apply(y => {
    // 時間が空か、もう一方のいつもの時間（8:00〜17:00／21:00〜6:00）のままなら、いつもの時間に入れ替える
    const t = [y.始め || "", y.終わり || ""].join("|"); y.区分 = v;
    if (!y.終日 && ["|", "8:00|17:00", "21:00|6:00"].includes(t)) { y.始め = v === "夜" ? "21:00" : "8:00"; y.終わり = v === "夜" ? "6:00" : "17:00"; }
  }, true); });
  root.querySelectorAll("[data-pf]").forEach(el => el.onchange = () => { PICK[el.dataset.pf] = el.value; const i = el.parentElement.querySelector("[data-ci]"); if (i) { i.focus(); i.dispatchEvent(new Event("input")); } });
  const gOf = (y, s, j) => listOf(y, s, "外注")[+j];
  root.querySelectorAll("[data-gm]").forEach(el => { const [s, j] = el.dataset.gm.split("|");
    bindPick(el, () => { const g = gOf(get(), s, j) || {}; return gaiMemItems(g.業者, (g.名前 || []).map(m => m.名前)); }, x => {
      const n = String(x.v || "").trim(); if (!n) return;
      apply(y => { const g = gOf(y, s, j); if (!g) return; g.名前 ||= []; if (g.名前.some(m => nk(m.名前) === nk(n))) return; g.名前.push({ 名前: n, 役: x.役 || "作業員" }); if ((parseFloat(g.人数) || 0) < g.名前.length) g.人数 = String(g.名前.length); }, true);
      refocus(`[data-gm="${s}|${j}"]`);
    }); });
  root.querySelectorAll("[data-grole]").forEach(el => el.onchange = () => { const [s, j, k] = el.dataset.grole.split("|"); apply(y => { const g = gOf(y, s, j); if (g && g.名前[+k]) g.名前[+k].役 = el.value; }); });
  root.querySelectorAll("[data-gdel]").forEach(el => el.onclick = e => { e.stopPropagation(); const [s, j, k] = el.dataset.gdel.split("|"); apply(y => { const g = gOf(y, s, j); if (g) g.名前.splice(+k, 1); }, true); });
  root.querySelectorAll("[data-hs]").forEach(el => el.onchange = () => { const on = el.dataset.hs === "1"; apply(y => { if (on) y.関連現場 ||= { 名前: "", コード: "" }; else delete y.関連現場; }, true); if (on) refocus("[data-rel]"); });
  const rel = root.querySelector("[data-rel]");
  if (rel) {
    const setRel = (v, code) => apply(y => { const st = siteOf(v, code); y.関連現場 = { 名前: v, コード: code || (st && st.コード) || "" }; }, true);
    bindPick(rel, () => placeItems().filter(x => x.grp !== "本社など"), x => { rel.value = x.v; setRel(x.v, x.コード); });
    rel.addEventListener("change", () => { const v = rel.value.trim(); if (v !== ((get().関連現場 || {}).名前 || "")) setRel(v, ""); });
  }
}
function grow(t) { t.style.height = "auto"; t.style.height = t.scrollHeight + 2 + "px"; }
const findRow = (d, id) => (d.行 || []).find(x => x.id === id);
// 行を変える。まだ行の無い日の入力欄（draft）なら、ここで行を作る
function setRow(id, fn, rerender) { mutate(S.day, d => { let r = findRow(d, id); if (!r && id === draftId()) { r = blankRow(id); d.行.push(r); } if (r) fn(r); }, rerender ? { render: true } : true); }
// 現場名（行先）を決めた：コードを入れ、作業内容がまだ空なら全体工程表のその日の予定を入れる
function setPlace(id, v, code) {
  setRow(id, y => {
    y.行先 = v; y.コード = code || "";
    fillPlan(y);
  }, true);
}
function bindCard(id) {
  const c = document.querySelector(`.card[data-id="${id}"]`), q = s => c.querySelector(s), r = () => rowsOf().find(x => x.id === id) || (id === draftId() ? blankRow(id) : {});
  const place = q('[data-f="place"]');
  bindPick(place, placeItems, x => { place.value = x.v; setPlace(id, x.v, x.コード); });
  place.addEventListener("change", () => { const v = place.value.trim(); if (v !== r().行先) { const s = siteOf(v), k = koteiSite(v); setPlace(id, v, s ? s.コード : k ? k.コード : ""); } });
  ["t0", "t1"].forEach((f, j) => { const el = q(`[data-f="${f}"]`);
    // ホイール・回す窓で何も入っていないとき：始めは 8:00、終わりは始めの9時間あと（8:00→17:00、21:00→6:00）
    bindTime(el, () => j ? (stepTime(r().始め, "h", 9) || "17:00") : "8:00", (v, re) => setRow(id, y => { y[j ? "終わり" : "始め"] = v; }, re));
    el.addEventListener("keydown", e => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); (j ? q('[data-wk="報告"]') : q('[data-f="t1"]')).focus(); } }); });
  c.querySelectorAll("[data-wj]").forEach(el => el.addEventListener("input", () => {
    grow(el); const j = +el.dataset.wj, f = el.dataset.wk, v = el.value;
    setRow(id, y => { y.作業行 = workLines(y).map(l => ({ ...l })); while (y.作業行.length <= j) y.作業行.push({ 予定: "", 報告: "" }); y.作業行[j][f] = v; syncWork(y); });
  }));
  c.querySelectorAll("[data-ci]").forEach(el => { const [s, kind] = el.dataset.ci.split("|");
    bindPick(el, () => kind === "人" ? personItems(r().行先, listIn(r(), s, kind)) : machItems(listIn(r(), s, kind)), async x => {
      const n = String(x.v || "").trim(); if (!n) return;
      if (listIn(r(), s, kind).some(v => nk(v) === nk(n))) { el.value = ""; return; }
      if (!await okClash(S.day, r(), { [kind]: [n] })) { el.value = ""; return; }
      setRow(id, y => { listOf(y, s, kind).push(n); }, true);
      setTimeout(() => { const p = document.querySelector(`.card[data-id="${id}"] [data-ci="${s}|${kind}"]`); if (p) p.focus(); }, 0);
    }); });
  const gOf = (y, sj) => { const [s, j] = sj.split("|"); return listOf(y, s, "外注")[+j]; };
  c.querySelectorAll("[data-gi]").forEach(el => { const sj = el.dataset.gi;
    bindPick(el, gaiItems, x => { el.value = x.v; setRow(id, y => { const g = gOf(y, sj); if (g) g.業者 = x.v; }); });
    el.addEventListener("change", () => setRow(id, y => { const g = gOf(y, sj); if (g) g.業者 = el.value.trim(); })); });
  c.querySelectorAll("[data-gn]").forEach(el => { const sj = el.dataset.gn; el.addEventListener("change", () => { const v = el.value.normalize("NFKC").replace(/[^\d.]/g, ""); el.value = v; setRow(id, y => { const g = gOf(y, sj); if (g) g.人数 = v; }); }); });
  bindExtras(c, r, (fn, re) => setRow(id, fn, re), sel => setTimeout(() => { const p = document.querySelector(`.card[data-id="${id}"] ${sel}`); if (p) p.focus(); }, 0));
  c.addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    if (b.dataset.tq) { const [a, z] = b.dataset.tq.split("|"); setRow(id, y => { y.始め = a; y.終わり = z; y.区分 = a === "21:00" ? "夜" : "昼"; }, true); }
    if (b.dataset.dc) { const [s, kind, j] = b.dataset.dc.split("|"); setRow(id, y => { listOf(y, s, kind).splice(+j, 1); }, true); }
    if (b.dataset.delg) { const [s, j] = b.dataset.delg.split("|"); setRow(id, y => { listOf(y, s, "外注").splice(+j, 1); }, true); }
    if (b.dataset.addg) { const s = b.dataset.addg; setRow(id, y => { listOf(y, s, "外注").push({ 業者: "", 人数: "" }); }, true); setTimeout(() => { const els = document.querySelectorAll(`.card[data-id="${id}"] [data-gi^="${s}|"]`); els.length && els[els.length - 1].focus(); }, 0); }
    // 予定の参加者・外注・機械を報告に写す（報告に書いてある分は残し、無いものを足す）
    if (b.dataset.act === "copyp") setRow(id, y => {
      ["人", "機械"].forEach(x => { const to = listOf(y, "r", x); listIn(y, "p", x).forEach(n => { if (!to.some(v => nk(v) === nk(n))) to.push(n); }); });
      const to = listOf(y, "r", "外注"); listIn(y, "p", "外注").forEach(g => { if (g.業者 && !to.some(v => nk(v.業者) === nk(g.業者))) to.push({ ...g }); });
    }, true);
    if (b.dataset.act === "addw") { setRow(id, y => { y.作業行 = [...workLines(y).map(l => ({ ...l })), { 予定: "", 報告: "" }]; }, true); setTimeout(() => { const els = document.querySelectorAll(`.card[data-id="${id}"] [data-wk="予定"]`); els.length && els[els.length - 1].focus(); }, 0); }
    if (b.dataset.delw != null) { const j = +b.dataset.delw; setRow(id, y => { y.作業行 = workLines(y).filter((l, i) => i !== j); syncWork(y); }, true); }
    if (b.dataset.act === "site") { const x = r(); openSite(x.行先, x.コード); }
    if (b.dataset.act === "del") { const x = r(); if ((x.行先 || workLines(x).some(l => l.予定 || l.報告) || (x.人 || []).length || (x.機械 || []).length) && !confirm(`「${x.行先 || "この行"}」を消しますか？`)) return; mutate(S.day, d => { d.行 = d.行.filter(y => y.id !== id); if (!d.行.length) d.報告 = null; }); }
    if (b.dataset.act === "detail") toDetail(r());
  });
}
// 打っている欄がキーボードに隠れないように（iPhone）
document.addEventListener("focusin", e => { if (e.target.matches("input,textarea") && window.innerWidth < 760) setTimeout(() => e.target.scrollIntoView({ block: "center", behavior: "smooth" }), 300); });

/* ---------- 下のボタン ---------- */
function addRow() {
  const rows = rowsOf(), last = rows[rows.length - 1], shown = S.view === "site" ? rows.filter(onSite) : rows;
  const focusOn = id => setTimeout(() => { const p = document.querySelector(`.card[data-id="${id}"] [data-f="${S.view === "site" ? "t0" : "place"}"]`); if (p) { p.focus(); p.scrollIntoView({ block: "center" }); } }, 30);
  if (!shown.length && document.querySelector(`.card[data-id="${draftId()}"]`)) return focusOn(draftId());   // 空の入力欄が出ている
  const id = uid();
  // 2つ目からは、前の行の終わりを始めに
  mutate(S.day, d => { d.行.push(blankRow(id, last ? last.終わり || "" : "")); }, { render: true });
  focusOn(id);
}
async function sameAsBefore() {
  await loadMonth(monthOf(dayAdd(S.day, -31)));
  let src = null;
  for (let i = 1; i <= 31 && !src; i++) { const k = dayAdd(S.day, -i); if (rowsOf(k).length) src = k; }
  if (!src) { toast("前の31日に書いた日がありません"); return; }
  const d = dateOf(src), rows = rowsOf(src);
  const list = rows.map(r => `<li>${esc(r.行先 || "（行先なし）")} ${esc([r.始め, r.終わり].filter(Boolean).join("〜"))}${(r.人 || []).length ? `・${r.人.length}名` : ""}${(r.外注 || []).length ? "・外注" : ""}</li>`).join("");
  dlg(`<h2>${d.getMonth() + 1}/${d.getDate()}(${WD[d.getDay()]}) と同じにしますか？</h2><ol>${list}</ol>
    <p style="font-size:14px;color:#5d6b7a">行先・時間・作業内容・参加者・外注・機械（予定と報告）を写します。${rowsOf().length ? "<b>この日に書いてある分は置き換わります。</b>" : ""}</p>
    <div class="btns"><button data-close>やめる</button><button class="pri" id="dOk">同じにする</button></div>`);
  $("dOk").onclick = () => { closeDlg(); const copy = JSON.parse(JSON.stringify(rows)).map(r => ({ ...r, id: uid() })); mutate(S.day, x => { x.行 = copy; x.報告 = null; }); toast("前の日と同じにしました。違う所を直してください"); };
}
function report() {
  const rows = rowsOf(), miss = rows.filter(r => !r.行先 || !r.始め || !r.終わり);
  const go = () => { mutate(S.day, d => { d.報告 = { だれ: S.me.name, いつ: nowStr() }; }, { report: true, render: true }); toast("報告しました"); };
  if (miss.length) { dlg(`<h2>書いていない所があります</h2><p>${miss.map(r => esc(r.行先 || "行先なし") + "：" + [!r.行先 && "行先", !r.始め && "始め", !r.終わり && "終わり"].filter(Boolean).join("・")).join("<br>")}</p><div class="btns"><button data-close>直す</button><button class="pri" id="dOk">このまま報告する</button></div>`); $("dOk").onclick = () => { closeDlg(); go(); }; return; }
  go();
}

/* ---------- 詳しく書く（常用日報へ渡す。同じ kkodin.github.io なのでブラウザの中で渡せる） ---------- */
function toDetail(r) {
  const x = { 日: S.day, 行先: r.行先, コード: r.コード || "", 始め: r.始め || "", 終わり: r.終わり || "", 作業: String(r.作業 || "").split(/\r?\n/).map(v => v.trim()).filter(Boolean), 人: r.人 || [], 外注: r.外注 || [], 機械: r.機械 || [], 書いた人: S.me.name, いつ: Date.now() };
  try { localStorage.setItem("nippou_to_jouyou", JSON.stringify(x)); } catch (e) { }
  const url = new URL("../jouyou/?from=nippou" + (DEMO ? "&demo" : ""), location.href).href;
  if (window.innerWidth < 760 || navigator.standalone) location.href = url; else window.open(url, "_blank");
}

/* ---------- 現場別日報（別ウィンドウ ?site=現場名&code=コード&d=日）
   全員の私の日報から、この現場の分を集めて見せる。自分の分はここでも同じカードで書ける（同じデータなので、私の日報とつながっている）。
   「月の集計」：作業員の出面・外注の人数・機械の稼働（自社機械稼働表）・書いた人の勤務時間・作業内容の予定と報告 ---------- */
const onSite0 = (r, site) => (site.code && r.コード && String(r.コード) === String(site.code)) || nk(r.行先) === nk(site.place);
const onSite = r => !!S.site && (onSite0(r, S.site) || (!!r.関連現場 && onSite0({ 行先: r.関連現場.名前, コード: r.関連現場.コード }, S.site)));
const DOCS = {};   // 月 → { t, docs }（全員の私の日報。20秒で読み直す）
async function monthDocs(m, force) {
  const c = DOCS[m];
  if (!c || force || Date.now() - c.t > 20000) {
    const files = (await nippouStore.list()).filter(f => f.name.startsWith("個人_") && f.name.endsWith(`_${m}.json`));
    const docs = (await Promise.all(files.map(f => nippouStore.get(f.name).catch(() => null)))).filter(Boolean).map(x => x.data);
    DOCS[m] = { t: Date.now(), docs };
  }
  const mine = S.F[m] && S.F[m].loaded ? S.F[m].doc : null;   // 自分の分は書きかけも入れる
  return DOCS[m].docs.filter(x => !mine || nk(x.名前) !== nk(S.me.name)).concat(mine ? [mine] : []);
}
function openSite(place, code) {
  if (!place) { toast("先に現場名（行先）を書いてください"); return; }
  const u = new URL("./", location.href);
  u.searchParams.set("site", place); if (code) u.searchParams.set("code", code); u.searchParams.set("d", S.day);
  if (DEMO) u.searchParams.set("demo", ""); if (GUEST_DEMO) u.searchParams.set("guest", "");
  window.open(u.href, "_blank");
}
// 「現場別詳細工程」：この日の現場が1つならそれ、ほかは選ぶ
// 「現場別詳細工程」（2026-11-02 ユーザー：工程が開くと思った。全体工程表・自社機械稼働表・日報（現場別／人別）を開けるように）
// ほかのアプリは同じ kkodin.github.io/kotei の中（全体工程表＝kotei/、自社機械稼働表＝kikai/、常用日報＝jouyou/）。別のタブで開く
function openApps() {
  const url = p => new URL("../" + p + (DEMO ? "?demo" : ""), location.href).href;
  dlg(`<div class="dh"><h2>開く</h2><button type="button" data-close aria-label="閉じる">×</button></div><div class="apps">
    <button type="button" data-app="kotei"><b>全体工程表</b><small>現場ごとの工程（予定）</small></button>
    <button type="button" data-app="kikai"><b>自社機械稼働表</b><small>機械の予約・稼働</small></button>
    <button type="button" data-app="site"><b>日報（現場別）</b><small>現場を選んで、その現場の全員の日報・月の集計</small></button>
    <button type="button" data-app="person"><b>日報（人別）</b><small>人を選んで、その人の月の一覧</small></button>
    <button type="button" data-app="jouyou"><b>常用日報</b><small>常用の現場の詳しい日報</small></button></div>`);
  $("dlgBody").querySelector(".apps").onclick = e => {
    const b = e.target.closest("[data-app]"); if (!b) return; const a = b.dataset.app;
    if (a === "site") return chooseSite();
    if (a === "person") return choosePerson();
    closeDlg(); window.open(url(a === "kotei" ? "" : a + "/"), "_blank");
  };
}
// 日報（人別）：人を選んで、その人の月の一覧
function choosePerson() {
  const list = [...new Set([S.me.name, ...S.M.社員, ...S.M.従業員])];
  const btns = () => { const q = nk($("pQ").value); return list.filter(n => !q || nk(n).includes(q)).map(n => `<button type="button" data-n="${esc(n)}">${esc(n)}</button>`).join("") || "<p>見つかりません</p>"; };
  dlg(`<h2>だれの日報を開きますか</h2><input id="pQ" placeholder="名前で絞り込み" autocomplete="off"><div class="namelist" id="pL"></div><div class="btns"><button data-close>やめる</button></div>`);
  $("pL").innerHTML = btns(); $("pQ").oninput = () => { $("pL").innerHTML = btns(); };
  $("pL").onclick = e => { const b = e.target.closest("[data-n]"); if (!b) return; closeDlg(); S.who = b.dataset.n; S.view = "list"; render(); };
}
function chooseSite() {
  const today = [], seen = new Set();
  rowsOf().forEach(r => { if (r.行先 && !isOffice(r.行先) && !seen.has(nk(r.行先))) { seen.add(nk(r.行先)); today.push({ grp: "この日の現場", v: r.行先, コード: r.コード || "" }); } });
  if (today.length === 1) return openSite(today[0].v, today[0].コード);
  const items = [...today, ...placeItems().filter(x => x.grp !== "本社など" && !seen.has(nk(x.v)))];
  const btns = () => { const q = nk($("sQ").value); let g = null; return items.filter(x => !q || nk(x.v).includes(q) || nk(x.sub || "").includes(q)).slice(0, 60).map((x, i) =>
    (x.grp !== g ? `<div class="grp">${esc(g = x.grp)}</div>` : "") + `<button type="button" data-i="${items.indexOf(x)}">${esc(x.v)}${x.sub ? `<small> ${esc(x.sub)}</small>` : ""}</button>`).join("") || "<p>見つかりません</p>"; };
  dlg(`<h2>どの現場の日報を開きますか</h2><input id="sQ" placeholder="現場名・コードで絞り込み" autocomplete="off"><div class="namelist" id="sL"></div><div class="btns"><button data-close>やめる</button></div>`);
  $("sL").innerHTML = btns(); $("sQ").oninput = () => { $("sL").innerHTML = btns(); };
  $("sL").onclick = e => { const b = e.target.closest("[data-i]"); if (!b) return; const x = items[+b.dataset.i]; closeDlg(); openSite(x.v, x.コード); };
}
async function renderSite() {
  const st = S.site, info = siteOf(st.place, st.code), m = monthOf(S.day);
  $("dayHead").innerHTML = `<div class="sitehead"><b>${esc(st.place)}</b>${info ? `<small>工事コード ${esc(info.コード || "―")}　元請 ${esc(info.元請 || "―")}</small>` : st.code ? `<small>工事コード ${esc(st.code)}</small>` : ""}</div>
    <div class="tabs"><button type="button" data-tab="day" class="${S.tab === "day" ? "on" : ""}">日ごと</button><button type="button" data-tab="sum" class="${S.tab === "sum" ? "on" : ""}">${+m.slice(5)}月の集計</button></div>`;
  $("dayHead").querySelectorAll("[data-tab]").forEach(b => b.onclick = () => { S.tab = b.dataset.tab; render(); });
  $("all").hidden = false;
  if (S.tab === "sum") { $("cards").hidden = true; return renderSiteSum(); }
  $("cards").hidden = false;
  const own = rowsOf().filter(onSite);
  drawCards(own.length ? own : [blankRow(draftId())], !own.length);
  const k = S.day;
  if (!$("all").innerHTML) $("all").innerHTML = `<p>ほかの人の日報を読んでいます…</p>`;
  try {
    const docs = await monthDocs(m);
    if (S.view !== "site" || S.tab !== "day" || S.day !== k) return;
    const xs = [];
    docs.forEach(doc => { if (nk(doc.名前) === nk(S.me.name)) return; const d = (doc.日 || {})[k]; (d && d.行 || []).filter(onSite).forEach(r => xs.push({ who: doc.名前, r, rep: !!d.報告 })); });
    $("all").innerHTML = `<h3>ほかの人の日報（この現場・この日）</h3>` + (xs.length ? `<div class="allsite">` + xs.map(x => `<div class="row"><span class="who">${esc(x.who)}</span><span class="t">${esc([x.r.始め, x.r.終わり].filter(Boolean).join("〜"))}${x.rep ? "" : "・書きかけ"}</span>` +
      workLines(x.r).filter(l => l.予定 || l.報告).map(l => `<div class="wv"><span>${esc(l.予定 || "")}</span><span>${esc(l.報告 || "")}</span></div>`).join("") +
      prText(x.r) + `</div>`).join("") + `</div>` : `<p class="none">ほかの人はまだ書いていません。</p>`);
  } catch (e) { $("all").innerHTML = `<p class="ng">ほかの人の日報を読めませんでした：${esc(e.message)}</p>`; }
}
const gaiTxt = L => (L || []).filter(g => g.業者).map(g => g.業者 + (g.人数 ? " " + g.人数 + "名" : "")).join("、");
function prText(r) {
  const line = (t, v) => v ? `<small>${t}：${esc(v)}</small><br>` : "";
  return [["参加者", "人", v => v.join("、")], ["外注", "外注", gaiTxt], ["機械", "機械", v => v.join("、")]].map(([t, k, f]) =>
    line(t + "（予定）", f(listIn(r, "p", k))) + line(t, f(listIn(r, "r", k)))).join("");
}
/* 集計 */
const tmin = t => { const m = /^(\d{1,2}):(\d{2})$/.exec(t || ""); return m ? +m[1] * 60 + +m[2] : null; };
function hoursOf(r) { const a = tmin(r.始め), b = tmin(r.終わり); if (a == null || b == null) return 0; let d = b - a; if (d <= 0) d += 1440; return d / 60; }
const fmtH = h => (Math.round(h * 10) / 10) + "時間";
const monthDays = m => { const out = []; for (let k = m + "-01"; monthOf(k) === m; k = dayAdd(k, 1)) out.push(k); return out; };
const dayCls = k => holidayOf(k) || dow(k) === 0 ? "sun" : dow(k) === 6 ? "sat" : "";
// 日×項目の表（rows：[名前, 日→値, 計]）。foot：日ごとの計の行
function gridTable(days, rows, foot) {
  if (!rows.length) return `<p class="none">この月はありません。</p>`;
  return `<div class="gwrap"><table class="grid"><thead><tr><th></th>${days.map(k => `<th class="${dayCls(k)}">${+k.slice(8)}<br><small>${WD[dow(k)]}</small></th>`).join("")}<th>計</th></tr></thead><tbody>` +
    rows.map(([n, f, t]) => `<tr><th>${esc(n)}</th>${days.map(k => `<td class="${dayCls(k)}">${esc(f(k))}</td>`).join("")}<td class="tot">${esc(t)}</td></tr>`).join("") +
    (foot ? `<tr class="foot"><th>${esc(foot[0])}</th>${days.map(k => `<td class="${dayCls(k)}">${esc(foot[1](k))}</td>`).join("")}<td class="tot">${esc(foot[2])}</td></tr>` : "") + `</tbody></table></div>`;
}
// entries：[{ k, who, r }]。by：勤務時間をまとめる（だれ／行先）
function sumHTML(entries, days, by, machines) {
  const ppl = new Map(), pplP = new Map(), mac = new Map(), macP = new Map(), gai = new Map(), hrs = new Map(), work = [];
  const addSet = (mp, n, k) => { if (!mp.has(n)) mp.set(n, new Set()); mp.get(n).add(k); };
  entries.forEach(({ k, who, r }) => {
    listIn(r, "r", "人").forEach(n => addSet(ppl, n, k)); listIn(r, "p", "人").forEach(n => addSet(pplP, n, k));
    listIn(r, "r", "機械").forEach(n => addSet(mac, n, k)); listIn(r, "p", "機械").forEach(n => addSet(macP, n, k));
    (r.外注 || []).forEach(g => { if (!g.業者) return; if (!gai.has(g.業者)) gai.set(g.業者, {}); const o = gai.get(g.業者); o[k] = (o[k] || 0) + (parseFloat(g.人数) || 0); });
    const key = by === "行先" ? (r.行先 || "（行先なし）") : who, o = hrs.get(key) || { days: new Set(), h: 0 }; o.days.add(k); o.h += hoursOf(r); hrs.set(key, o);
    workLines(r).forEach(l => { if (l.予定 || l.報告) work.push({ k, who: by === "行先" ? r.行先 : who, ...l }); });
  });
  const marks = (act, plan) => [...new Set([...act.keys(), ...plan.keys()])].map(n => [n, act.get(n) || new Set(), plan.get(n) || new Set()])
    .sort((a, b) => b[1].size - a[1].size).map(([n, a, p]) => [n, k => a.has(k) ? "○" : p.has(k) ? "予" : "", a.size + "日"]);
  const pplRows = marks(ppl, pplP), macRows = marks(mac, macP), nobe = [...ppl.values()].reduce((a, x) => a + x.size, 0);
  const pplDay = k => { let c = 0; ppl.forEach(set => { if (set.has(k)) c++; }); return c || ""; };
  const gaiRows = [...gai].map(([n, o]) => [n, k => o[k] || "", Object.values(o).reduce((a, x) => a + x, 0) + "人"]);
  const gaiDay = k => { let c = 0; gai.forEach(o => { c += o[k] || 0; }); return c || ""; };
  const sect = (t, h) => `<section class="sum"><h4>${t}</h4>${h}</section>`;
  const key = `<small class="key">○ 報告　予 予定だけ</small>`;
  return sect(`参加者の出面（${ppl.size}人・のべ ${nobe}人日）${key}`, gridTable(days, pplRows, ["人数", pplDay, ""])) +
    sect("外注（業者ごとの人数・報告）", gridTable(days, gaiRows, gaiRows.length > 1 ? ["計", gaiDay, ""] : null)) +
    sect(`使用機械（日報）${key}`, gridTable(days, macRows)) +
    (machines === undefined ? "" : sect("機械の予約（自社機械稼働表）", machines === null ? `<p class="none">自社機械稼働表を読めないため出せません。</p>` : gridTable(days, machines))) +
    sect(by === "行先" ? "行先ごとの時間" : "書いた人の勤務時間（この現場にいた時間）", hrs.size ? `<table class="list"><tr><th>${by === "行先" ? "行先" : "書いた人"}</th><th>日数</th><th>時間</th></tr>` +
      [...hrs].map(([n, o]) => `<tr><td>${esc(n)}</td><td class="num">${o.days.size}日</td><td class="num">${fmtH(o.h)}</td></tr>`).join("") + `</table>` : `<p class="none">この月はありません。</p>`) +
    sect("作業内容（予定と報告）", work.length ? `<table class="list work"><tr><th>日</th><th>${by === "行先" ? "行先" : "だれ"}</th><th>予定</th><th>報告</th></tr>` +
      work.sort((a, b) => a.k.localeCompare(b.k)).map(w => `<tr class="${w.予定 && !w.報告 ? "miss" : ""}"><td>${+w.k.slice(5, 7)}/${+w.k.slice(8)}</td><td>${esc(w.who || "")}</td><td>${esc(w.予定 || "")}</td><td>${esc(w.報告 || "")}</td></tr>`).join("") + `</table>` : `<p class="none">この月はありません。</p>`);
}
// この現場の機械（自社機械稼働表の予約で、その日に使うもの。仮は入れない）
function siteMachines(days) {
  if (!S.kikai) return null;
  const ks = koteiSite(S.site.place, S.site.code), names = [S.site.place, ks && ks.名前, ks && ks.略称].filter(Boolean).map(nk), mach = new Map();
  (S.kikai.予約 || []).forEach(y => {
    if (y.状態 === "仮" || !names.includes(nk(y.現場))) return;
    const mc = (S.kikai.機械 || []).find(x => x.id === y.機械); if (!mc) return;
    const key = mc.名前 + (mc.規格 ? " " + mc.規格 : "") + (y.区分 === "夜" ? "（夜）" : "");
    const o = mach.get(key) || {}; days.forEach(k => { if (kikaiWorks(y, k)) o[k] = (o[k] || 0) + 1; }); mach.set(key, o);
  });
  return [...mach].filter(([, o]) => Object.keys(o).length).map(([n, o]) => [n, k => o[k] || "", Object.keys(o).length + "日"]);
}
async function renderSiteSum() {
  const m = monthOf(S.day);
  $("all").innerHTML = `<p>集計しています…</p>`;
  await loadMonth(m);
  try {
    const docs = await monthDocs(m);
    if (S.view !== "site" || S.tab !== "sum" || monthOf(S.day) !== m) return;
    const entries = [];
    docs.forEach(doc => Object.entries(doc.日 || {}).forEach(([k, d]) => { if (monthOf(k) === m) (d.行 || []).filter(onSite).forEach(r => entries.push({ k, who: doc.名前, r })); }));
    const days = monthDays(m);
    $("all").innerHTML = `<h3>${esc(S.site.place)}　${+m.slice(5)}月の集計<small>（全員の日報から。◀▶ で月を移る）</small></h3>` + sumHTML(entries, days, "だれ", siteMachines(days));
  } catch (e) { $("all").innerHTML = `<p class="ng">集計できませんでした：${esc(e.message)}</p>`; }
}
// 月の一覧（「…」→ この月の一覧）：1日〜月末を縦に、日・曜日・現場・作業内容（要約）。名前を選ぶと、ほかの人の分も見られる
function workSum(rs) {
  const rep = [], plan = [], pick = a => [...new Set(a.map(v => v.trim()).filter(Boolean))].join("、");
  rs.forEach(r => workLines(r).forEach(l => { if (l.報告) rep.push(...String(l.報告).split(/\r?\n/)); else if (l.予定) plan.push(l.予定); }));
  const t = pick(rep), cut = v => v.length > 60 ? v.slice(0, 60) + "…" : v;
  return t ? { t: cut(t) } : { t: cut(pick(plan)), plan: true };
}
async function renderList() {
  $("cards").hidden = true; $("empty").hidden = true; $("all").hidden = false;
  $("dayHead").innerHTML = `<button id="bBack">← 日報に戻る</button>`;
  $("bBack").onclick = () => { S.view = S.site ? "site" : "me"; render(); };
  const m = monthOf(S.day);
  if (!$("all").querySelector(".mlist")) $("all").innerHTML = `<p>読んでいます…</p>`;
  await loadMonth(m);
  let docs = [];
  try { docs = await monthDocs(m); } catch (e) { toast("ほかの人の日報を読めませんでした：" + e.message); }
  if (S.view !== "list" || monthOf(S.day) !== m) return;
  // 人は書いた人だけでなく、ほかの人の日報に参加者として出てくる人（作業員）も選べる。その人が入っている行を全員の日報から集める
  const names = [...new Set([S.me.name, ...docs.map(d => d.名前).filter(Boolean), ...(S.who ? [S.who] : [])])];
  const who = names.find(n => nk(n) === nk(S.who)) || S.me.name, mine = nk(who) === nk(S.me.name);
  const body = monthDays(m).map(k => {
    const xs = evsOn(docs, k).filter(x => membersOf(x).some(v => nk(v) === nk(who))), rs = xs.map(x => x.r), w = workSum(rs);
    const places = [...new Set(rs.map(r => titleOf(r)).filter(Boolean))].join("　"), by = [...new Set(xs.filter(x => nk(x.who) !== nk(who)).map(x => x.who))];
    return `<tr data-k="${k}" class="${dayCls(k)}"><td class="num">${+k.slice(8)}</td><td>${WD[dow(k)]}</td><td>${esc(places)}${by.length ? `<br><small class="plan">記入：${esc(by.join("、"))}</small>` : ""}</td><td>${w.plan && w.t ? `<span class="plan">予定：${esc(w.t)}</span>` : esc(w.t)}${xs.some(x => !x.rep) ? `<span class="mk">書きかけ</span>` : ""}</td></tr>`;
  }).join("");
  $("all").innerHTML = `<h3 class="lhead">${+m.slice(5)}月　<select id="lWho">${names.map(n => `<option${n === who ? " selected" : ""}>${esc(n)}</option>`).join("")}</select><small>◀▶ で月を移る・行を押すとその日</small></h3>
    <table class="list mlist"><thead><tr><th>日</th><th>曜日</th><th>現場</th><th>作業内容</th></tr></thead><tbody>${body}</tbody></table>`;
  $("lWho").onchange = () => { S.who = $("lWho").value; renderList(); };
  $("all").querySelector("tbody").onclick = e => { const tr = e.target.closest("tr[data-k]"); if (!tr) return; S.view = S.site ? "site" : mine ? "me" : "all"; if (S.site) S.tab = "day"; setDay(tr.dataset.k); };
}
// 私の月の集計（「…」→ この月の集計）
async function renderMySum() {
  $("cards").hidden = true; $("empty").hidden = true; $("all").hidden = false;
  $("dayHead").innerHTML = `<button id="bBack">← ${esc(S.me.name)} の日報に戻る</button>`;
  $("bBack").onclick = () => { S.view = "me"; render(); };
  const m = monthOf(S.day);
  await loadMonth(m);
  if (S.view !== "mysum") return;
  const doc = fileOf(m).doc, days = monthDays(m), entries = [];
  let n = 0, total = 0;
  days.forEach(k => { const rs = ((doc.日 || {})[k] || {}).行 || []; if (rs.length) n++; rs.forEach(r => { total += hoursOf(r); entries.push({ k, who: S.me.name, r }); }); });
  $("all").innerHTML = `<h3>${+m.slice(5)}月の集計（${esc(S.me.name)}）<small>書いた日 ${n}日・合計 ${fmtH(total)}</small></h3>` + sumHTML(entries, days, "行先");
}

/* ---------- この日の全員の日報（行先ごと） ---------- */
async function renderAll() {
  $("cards").hidden = true; $("empty").hidden = true; $("all").hidden = false;
  $("dayHead").innerHTML = `<button id="bBack">← ${esc(S.me.name)} の日報に戻る</button>`;
  $("bBack").onclick = () => { S.view = "me"; render(); };
  $("all").innerHTML = `<p>読んでいます…</p>`;
  const k = S.day, m = monthOf(k);
  try {
    const files = (await nippouStore.list()).filter(f => f.name.startsWith("個人_") && f.name.endsWith(`_${m}.json`));
    const docs = (await Promise.all(files.map(f => nippouStore.get(f.name).catch(() => null)))).filter(Boolean).map(x => x.data);
    const by = new Map(), none = [];
    docs.forEach(doc => { const d = (doc.日 || {})[k]; if (!d || !(d.行 || []).length) { none.push(doc.名前); return; }
      d.行.forEach(r => { const key = r.行先 || "（行先なし）"; if (!by.has(key)) by.set(key, []); by.get(key).push({ who: doc.名前, r, rep: !!d.報告 }); }); });
    if (S.view !== "all" || S.day !== k) return;
    const d = dateOf(k);
    $("all").innerHTML = `<h3>${d.getMonth() + 1}/${d.getDate()}(${WD[d.getDay()]}) の全員の日報（${docs.length - none.length} 人）</h3>` +
      ([...by.entries()].sort((a, b) => isOffice(a[0]) - isOffice(b[0]) || a[0].localeCompare(b[0], "ja")).map(([place, xs]) => {
        const ppl = new Set(), gai = []; xs.forEach(x => { (x.r.人 || []).forEach(n => ppl.add(n)); (x.r.外注 || []).forEach(g => g.業者 && gai.push(`${g.業者}${g.人数 ? " " + g.人数 + "名" : ""}`)); });
        return `<div class="allsite"><h4>${esc(place)}${ppl.size ? `　<small>作業員 ${ppl.size}名</small>` : ""}${gai.length ? `　<small>外注 ${esc(gai.join("・"))}</small>` : ""}</h4>` +
          xs.map(x => `<div class="row"><span class="who">${esc(x.who)}</span><span class="t">${esc([x.r.始め, x.r.終わり].filter(Boolean).join("〜"))}${x.rep ? "" : "・書きかけ"}</span><br>${esc(x.r.作業 || "")}${(x.r.人 || []).length ? `<br><small>来ていた人：${esc(x.r.人.join("、"))}</small>` : ""}</div>`).join("") + `</div>`;
      }).join("") || `<p>この日はまだだれも書いていません。</p>`) +
      (none.length ? `<p style="font-size:13px;color:#5d6b7a">この日をまだ書いていない人：${esc(none.join("、"))}</p>` : "");
  } catch (e) { $("all").innerHTML = `<p class="ng">読めませんでした：${esc(e.message)}</p>`; }
}

/* ---------- 窓・知らせ ---------- */
// cls：wide＝広い窓、keep＝外を押しても閉じない（書いている途中の窓）
function dlg(html, cls) { $("dlg").className = cls || ""; $("dlgBody").innerHTML = html; const d = $("dlg"); if (!d.open) d.showModal(); }
function closeDlg() { $("dlg").close(); }
$("dlg").addEventListener("click", e => { if (e.target.closest("[data-close]") || (e.target === $("dlg") && !$("dlg").classList.contains("keep"))) closeDlg(); });
let toastT = null;
function toast(m, ms = 3000) { const t = $("toast"); t.textContent = m; t.classList.add("on"); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("on"), ms); }

/* ---------- 書く人（サインインしたアカウント ↔ 名簿の名前。最初に1回だけ選ぶ。打たない） ---------- */
const ME_KEY = "nippou_me_v3";
function savedMe(acct) { try { return (JSON.parse(localStorage.getItem(ME_KEY)) || {})[acct] || ""; } catch (e) { return ""; } }
function askMe(acctName, first) {
  return new Promise(res => {
    const list = people(), guess = nk(acctName);
    const btns = () => { const q = nk($("meQ").value); return list.filter(n => !q || nk(n).includes(q)).map(n => `<button type="button" data-n="${esc(n)}" class="${nk(n) === guess ? "hit" : ""}">${esc(n)}</button>`).join("") || "<p>名簿に見つかりません</p>"; };
    dlg(`<h2>${first ? "あなたの名前を選んでください" : "名前を選び直す"}</h2><p style="font-size:14px">この名前で日報を書きます（サインイン：${esc(acctName)}）。次からは聞きません。</p>
      <input id="meQ" placeholder="名前で絞り込み" autocomplete="off"><div class="namelist" id="meL"></div>${first ? "" : `<div class="btns"><button data-close>やめる</button></div>`}`);
    $("meL").innerHTML = btns();
    $("meQ").oninput = () => { $("meL").innerHTML = btns(); };
    $("meL").onclick = e => { const b = e.target.closest("[data-n]"); if (!b) return; closeDlg(); res(b.dataset.n); };
    const hit = $("meL").querySelector(".hit"); if (hit) hit.scrollIntoView({ block: "center" });
  });
}
function rememberMe(name) {
  S.me.name = name;
  try { const a = JSON.parse(localStorage.getItem(ME_KEY)) || {}; a[S.me.acct] = name; localStorage.setItem(ME_KEY, JSON.stringify(a)); } catch (e) { }
}
async function setMe(name) { rememberMe(name); S.F = {}; await loadMonth(monthOf(S.day)); render(); }

/* ---------- ホーム画面に置く（アドレスバーを消す） ---------- */
let installEv = null;
window.addEventListener("beforeinstallprompt", e => { e.preventDefault(); installEv = e; });
const standalone = () => window.matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
function installGuide() {
  if (installEv) { installEv.prompt(); installEv = null; return; }
  const ios = /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  dlg(`<h2>ホーム画面に置く</h2><p>ホーム画面（デスクトップ）から開くと、アドレスバーやタブが出ず、入力する所が広くなります。</p>` +
    (ios ? `<ol><li>Safari でこのページを開く</li><li>下（iPad は上）の <b>共有</b>（□に↑）を押す</li><li><b>ホーム画面に追加</b> → <b>追加</b></li><li>ホーム画面の「工事日報」から開く（最初の1回だけサインインし直します）</li></ol>`
      : `<ol><li>Edge か Chrome で、アドレスバーの右の <b>アプリをインストール</b>（□に↓）を押す</li><li>または「…」メニュー → <b>アプリ</b> → <b>このサイトをアプリとしてインストール</b></li><li>デスクトップ・スタートの「工事日報」から開く</li></ol>`) +
    `<div class="btns"><button class="pri" data-close>わかった</button></div>`);
}

/* ---------- 日報のフォルダを調べる（管理者が1回だけ。?setup で開くと「…」に出る） ----------
   OneDrive の共有リンクから driveId・itemId を出し、config.js の nippou に書いてもらう（リンクそのものは config に書かない） */
function findFolder() {
  dlg(`<h2>日報のフォルダを調べる</h2><p style="font-size:14px">OneDrive で日報のフォルダを選び、「共有」→「リンクのコピー」で写したリンクを貼ってください。</p>
    <input id="shUrl" placeholder="https://…sharepoint.com/…" autocomplete="off"><pre id="shOut" style="white-space:pre-wrap;font-size:14px"></pre>
    <div class="btns"><button data-close>閉じる</button><button class="pri" id="shGo">調べる</button></div>`);
  $("shGo").onclick = async () => {
    const u = $("shUrl").value.trim(); if (!u) return;
    const id = "u!" + btoa(unescape(encodeURIComponent(u))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    $("shOut").textContent = "調べています…";
    try {
      const r = await g(`/shares/${id}/driveItem`); if (!r.ok) throw httpErr(r, "リンクを開けません");
      const j = await r.json(); if (!j.folder) throw new Error("フォルダのリンクではありません");
      $("shOut").textContent = `「${j.name}」
nippou: {
  driveId: "${j.parentReference.driveId}",
  itemId: "${j.id}"
}

この4行を Claude に渡してください（config.js に入れます）。`;
    } catch (e) { $("shOut").textContent = "調べられませんでした：" + e.message; }
  };
}

/* ---------- 日を移る・ボタン ---------- */
async function setDay(k) {
  S.day = k; closeSug();
  if (S.view !== "site") try { localStorage.setItem("nippou_view_v3", JSON.stringify({ day: k })); } catch (e) { }   // 現場別日報のウィンドウは私の日報の日を変えない
  render();
  await loadMonth(monthOf(k)); if (S.day === k) render();
}
function bind() {
  const monthView = () => S.view === "mysum" || S.view === "list" || (S.view === "site" && S.tab === "sum");
  const monthStep = n => { const d = dateOf(monthOf(S.day) + "-01"); d.setMonth(d.getMonth() + n); return keyOf(d); };
  $("bPrev").onclick = () => setDay(S.view === "cal" ? calStep(-1) : monthView() ? monthStep(-1) : dayAdd(S.day, -1));
  $("bNext").onclick = () => setDay(S.view === "cal" ? calStep(1) : monthView() ? monthStep(1) : dayAdd(S.day, 1));
  $("bToday").onclick = () => setDay(todayKey());
  $("inDay").onchange = () => { if ($("inDay").value) setDay($("inDay").value); };
  $("bAdd").onclick = () => { if (S.view === "cal") return editEv(S.day, null); if (S.view === "all" || S.view === "mysum" || S.view === "list") { S.view = S.site ? "site" : "me"; render(); } if (S.view === "site" && S.tab !== "day") { S.tab = "day"; render(); } addRow(); };
  $("bSite").onclick = openApps;
  $("bReport").onclick = report;
  $("bMe").onclick = async () => { const n = await askMe(S.me.acct, false); if (n) setMe(n); };
  $("bMore").onclick = e => { e.stopPropagation(); $("menu").hidden = !$("menu").hidden; };
  document.addEventListener("click", e => { if (!e.target.closest("#menu")) $("menu").hidden = true; });
  $("menu").onclick = async e => {
    const m = e.target.closest("[data-m]"); if (!m) return; $("menu").hidden = true;
    if (m.dataset.m === "all") { S.view = "all"; render(); }
    if (m.dataset.m === "reload") { await loadMasters(); await loadMonth(monthOf(S.day), true); render(); toast("最新を読みました"); }
    if (m.dataset.m === "install") installGuide();
    if (m.dataset.m === "me") $("bMe").click();
    if (m.dataset.m === "find") findFolder();
    if (m.dataset.m === "same") sameAsBefore();
    if (m.dataset.m === "kuro") kuroOpen();
    if (m.dataset.m === "list") { S.view = "list"; render(); }
    if (m.dataset.m === "cal") { S.view = "cal"; render(); }
    if (m.dataset.m === "cards") { S.view = "me"; render(); }
    if (m.dataset.m === "mysum") { S.view = S.view === "site" ? "site" : "mysum"; if (S.view === "site") S.tab = "sum"; render(); }
  };
  // PC：Alt+←/→ で日を移る
  document.addEventListener("keydown", e => { if (e.altKey && (e.key === "ArrowLeft" || e.key === "ArrowRight")) { e.preventDefault(); setDay(dayAdd(S.day, e.key === "ArrowLeft" ? -1 : 1)); } });
  window.addEventListener("beforeunload", e => { if (Object.values(S.F).some(f => f.pending.length || f.saving)) { e.preventDefault(); e.returnValue = ""; } });
  document.addEventListener("visibilitychange", () => { if (!document.hidden && !Object.values(S.F).some(f => f.pending.length || f.saving)) loadMonth(monthOf(S.day), true).then(() => { if (!document.activeElement || !document.activeElement.matches("input,textarea")) render(); }); });
}

async function start() {
  if ("serviceWorker" in navigator && location.protocol === "https:") navigator.serviceWorker.register("sw.js").catch(() => { });
  try { const v = JSON.parse(localStorage.getItem("nippou_view_v3")) || {}; if (v.day) S.day = v.day; } catch (e) { }
  const SP = new URLSearchParams(location.search);
  if (SP.get("site")) {
    S.view = "site"; S.site = { place: SP.get("site"), code: SP.get("code") || "" };
    if (/^\d{4}-\d{2}-\d{2}$/.test(SP.get("d") || "")) S.day = SP.get("d");
    document.title = S.site.place + "｜現場別日報";
  }
  cachedMasters();
  if (DEMO) { seedDemo(); S.me.acct = "試し"; }
  else if (location.protocol === "file:") {
    $("signin").hidden = false; $("bSignIn").hidden = true;
    $("signinMsg").innerHTML = 'パソコンのファイルを直接開いているため、サインインできません。<br>公開先から開いてください：<a href="https://kkodin.github.io/kotei/nippou4/">https://kkodin.github.io/kotei/nippou4/</a>';
    return;
  } else {
    msalApp = new msal.PublicClientApplication({ auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: new URL("./", location.href).href }, cache: { cacheLocation: "localStorage" } });
    await msalApp.initialize();
    try { const r = await msalApp.handleRedirectPromise(); if (r) msalApp.setActiveAccount(r.account); } catch (e) { $("signinMsg").textContent = "サインインできませんでした：" + e.message; }
    const acc = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
    if (!acc) { $("signin").hidden = false; $("bSignIn").onclick = () => msalApp.loginRedirect({ scopes: SCOPES }); return; }
    msalApp.setActiveAccount(acc);
    S.me.acct = acc.username || acc.name;
    S.me.acctName = acc.name || acc.username;
  }
  $("app").hidden = false; $("bar").hidden = false;
  if (!DEMO && new URLSearchParams(location.search).has("setup")) $("mFind").hidden = false;
  bind();
  if (!S.site) S.view = "cal";   // いつもカレンダーから開く（クロジカと同じ）
  await loadMasters().catch(e => toast("マスタを読めませんでした（前に読んだ控えで動きます）：" + e.message));
  await loadSets();
  if (S.site && !S.site.code) { const x = siteOf(S.site.place) || koteiSite(S.site.place); if (x && x.コード) S.site.code = String(x.コード); }
  let name = savedMe(S.me.acct);
  if (!name) name = await askMe(S.me.acctName || S.me.acct, true);
  rememberMe(name);
  await setDay(S.day);
  if (!standalone() && !localStorage.getItem("nippou_install_told") && window.innerWidth < 1100) { try { localStorage.setItem("nippou_install_told", "1"); } catch (e) { } setTimeout(installGuide, 800); }
}

/* ---------- 試し（?demo）：架空の名簿・現場・前の日の日報 ---------- */
function seedDemo() {
  try { if (localStorage.getItem(DEMO_KEY)) return; } catch (e) { return; }
  const put = (name, data) => { const all = JSON.parse(localStorage.getItem(DEMO_KEY) || "{}"); all[name] = { data, eTag: "e" + Math.random().toString(36).slice(2) }; localStorage.setItem(DEMO_KEY, JSON.stringify(all)); };
  if (GUEST_DEMO) put("日報/名簿.json", { 形式: "工事日報の名簿", 版: 1, 社員: ["見本　一郎", "見本　二郎", "試験　花子"], 従業員: ["作業　三郎", "作業　四郎", "作業　五郎"], 外注業者: ["第一カッター興業", "見本警備"], 現場: [{ コード: "510001", 名前: "A現場 配水管", 正式名: "A町配水管布設工事", 元請: "見本市水道局" }] });
  put("工程/マスタ.json", { 社員: ["見本　一郎", "見本　二郎", "試験　花子"], 従業員: ["作業　三郎", "作業　四郎", "作業　五郎"], 外注業者: ["第一カッター興業", "見本警備"] });
  put("機械/現場一覧.json", { 現場: [{ コード: "510001", 名前: "A現場 配水管", 正式名: "A町配水管布設工事", 元請: "見本市水道局", 発注者: "見本市水道局" }, { コード: "510002", 名前: "B現場 道路改良", 正式名: "B線道路改良工事", 元請: "見本建設", 発注者: "見本県" },
    { コード: "510003", 名前: "C現場 舗装", 正式名: "C市道舗装修繕工事", 元請: "自社請け", 発注者: "見本市道路局" }, { コード: "510004", 名前: "D現場 下水", 正式名: "D町下水道工事", 元請: "自社請け", 発注者: "見本町" }] });
  const a = n => dayAdd(todayKey(), n);
  put("工程/休日.json", { 休み: [] });
  put("工程/現場_s1.json", { id: "s1", 名前: "A現場 配水管", コード: "510001", 段: [{ id: "l1", 区分: "夜", 名前: "管布設工", 土: false }, { id: "l2", 区分: "夜", 名前: "路面復旧工", 種別: "外注", 業者: "第一カッター興業" }],
    バー: [{ id: "b1", 段: "l1", 作業: "配水管布設", 開始: a(-10), 終了: a(10) }, { id: "b2", 段: "l2", 作業: "舗装切断工", 開始: a(-2), 終了: a(3) }] });
  put("工程/現場_s2.json", { id: "s2", 名前: "B現場 道路改良", コード: "510002", 段: [{ id: "m1", 区分: "昼", 名前: "", 土: true }],
    バー: [{ id: "c1", 段: "m1", 作業: "伐採工", 開始: a(-5), 終了: a(2) }, { id: "c2", 段: "m1", 作業: "掘削工", 開始: a(3), 終了: a(15) }] });
  put("機械/稼働データ.json", { 機械: [{ id: "k1", 名前: "バックホウ", 規格: "0.25㎥", 分類: "重機" }, { id: "k2", 名前: "規制車", 規格: "2t", 分類: "車両" }, { id: "k3", 名前: "4t セルフダンプ", 規格: "", 分類: "車両" }, { id: "k4", 名前: "発電機", 規格: "", 分類: "小型機械" }],
    予約: [{ id: "y1", 機械: "k1", 区分: "昼", 現場: "B現場 道路改良", 開始: a(-5), 終了: a(15), 土: true, 日祝: false, 状態: "確定" },
      { id: "y2", 機械: "k2", 区分: "夜", 現場: "A現場 配水管", 開始: a(-10), 終了: a(10), 土: false, 日祝: false, 状態: "確定" }] });
  put("日報/現場設定.json", { 形式: "工事日報の現場設定", 版: 1, 現場: { "c:510001": { 形態: "常用", 名前: "A現場 配水管" } } });
  const y = dayAdd(todayKey(), -1), m = monthOf(y);
  put(`日報/個人_見本一郎_${m}.json`, { 形式: "私の日報", 版: 1, 名前: "見本　一郎", アカウント: "試し", 月: m, 日: { [y]: { 行: [
    { id: "a1", 行先: "本社", コード: "", 始め: "8:00", 終わり: "10:00", 作業: "朝礼・書類", 人: [], 外注: [] },
    { id: "a2", 行先: "A現場 配水管", コード: "510001", 始め: "10:00", 終わり: "17:00", 作業: "配水管布設 φ150 L=12m", 予定: { 人: ["作業　三郎", "作業　四郎", "作業　五郎"], 外注: [{ 業者: "第一カッター興業", 人数: "2" }], 機械: ["規制車 2t"] }, 人: ["作業　三郎", "作業　四郎"], 外注: [{ 業者: "第一カッター興業", 人数: "2" }], 機械: ["規制車 2t", "発電機（リース）"] }], 報告: { だれ: "見本　一郎", いつ: "" } } } });
  put(`日報/個人_見本二郎_${m}.json`, { 形式: "私の日報", 版: 1, 名前: "見本　二郎", アカウント: "試し2", 月: m, 日: { [y]: { 行: [
    { id: "b1", 行先: "B現場 道路改良", コード: "510002", 始め: "8:00", 終わり: "17:00", 作業: "伐採工", 人: ["作業　五郎"], 外注: [], 機械: ["バックホウ 0.25㎥"] }], 報告: null } } });
  // 今週の分（カレンダーの見本）
  const wk = dayAdd(todayKey(), -((dateOf(todayKey()).getDay() + 6) % 7)), wm = {};
  const add = (who, k, r) => { const f = `日報/個人_${nk(who)}_${monthOf(k)}.json`; wm[f] ||= { 形式: "私の日報", 版: 1, 名前: who, アカウント: "試し", 月: monthOf(k), 日: {} }; (wm[f].日[k] ||= { 行: [], 報告: { だれ: who, いつ: "" } }).行.push({ id: uid(), コード: "", 作業: "", 人: [], 外注: [], 機械: [], ...r }); };
  for (let i = 0; i < 5; i++) { const k = dayAdd(wk, i);
    add("見本　二郎", k, { 行先: "B現場 道路改良", コード: "510002", 始め: "8:00", 終わり: "17:00", 作業: "掘削工", 人: ["作業　五郎"], 機械: ["バックホウ 0.25㎥"] });
    add("試験　花子", k, { 行先: "本社", 始め: "8:30", 終わり: "17:30", 作業: "書類" });
    if (i === 2) add("試験　花子", k, { 行先: "C現場 舗装", コード: "510003", 始め: "13:00", 終わり: "15:00", 作業: "発注者と打合せ" }); }
  Object.entries(wm).forEach(([f, d]) => { const cur = (JSON.parse(localStorage.getItem(DEMO_KEY) || "{}")[f] || {}).data; if (cur) Object.entries(d.日).forEach(([k, v]) => { if (!cur.日[k]) cur.日[k] = v; }); put(f, cur || d); });
}
