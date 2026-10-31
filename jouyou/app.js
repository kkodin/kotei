/* 常用日報 Vr000（2026-10-31 たたき台）
   元の Excel：元請の「作業日報」（1日1シート・A4 横・A〜P 列）
   ・元請業者 → 現場 を選ぶ。現場ごとに日報・作業員名簿・機材を持つ
   ・日報：日付（曜日は自動）・天気・気温・記入者・元請サイン（iPad・スマホだけ書ける）
           上の表＝工種・作業内容・人員（作業員の工種ごとの人数を自動）・使用資機材
           下の表＝作業員名・始業・終業・工種・残業（終業から自動）｜使用機材・規格・数量・累計（自動）｜品名・数量・累計数量（自動）
   ・表のセルの動きは Excel と同じ（grid.js＝日報 Vr002 と同じもの）
   ・データはブラウザ（localStorage）に入る。「保存」でファイル（.json）、「取込」で .json か元の Excel（.xlsx）を読む
   ・出力：Excel（1シート1日＋集計シート）・PDF（印刷。A4 横で1日1ページ） */
"use strict";

const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
const KEY = "jouyou_nippou_v1", UIKEY = "jouyou_nippou_ui";
const WD = "日月火水木金土";
const pad = n => String(n).padStart(2, "0");
const ymd = d => d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
const pd = s => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const addD = (s, n) => { const d = pd(s); d.setDate(d.getDate() + n); return ymd(d); };
const wd = s => pd(s).getDay();
const jpDate = s => { const d = pd(s); return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`; };
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const now = () => { const d = new Date(); return ymd(d) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes()) + ":" + pad(d.getSeconds()); };

/* ================= データ ================= */
// DB = { 形式, 版, 元請:[{id,名前}], 現場:[現場], 職種:[], 天気:[] }
// 現場 = { id, 元請id, 名前, 担当者, 始業, 終業, 締め日, 累計:"期間"|"全体", 作業員:[{名前,職種}], 機材:[{名前,規格}], 日:{ "YYYY-MM-DD": 日 },
//         消した日:{ "YYYY-MM-DD": 消した時刻 }（ほかの端末から消したのを伝える）, 設定更新（決まり・名簿を直した時刻） }
// 日 = { 天気, 気温, 担当者, 上:[[工種,作業内容,人員,資機材]×5], 下:[[12列]], 備考, サイン(dataURL), サイン日時, 人員手(true＝人員を手で直した), 更新 }
// 職種単価：{ 工種: { 昼, 夜, 残業 } }（全部の現場で共通。2026-10-31 ユーザー）。共通更新＝職種・単価・天気を直した時刻（共有の 共通.json と合わせる）
function blankDB() {
  return { 形式: "常用日報", 版: 1, 元請: [], 現場: [], 職種: ["職長", "重機OP", "普通作業員", "管理"], 天気: ["晴", "曇", "雨", "雪", "晴/曇", "曇/晴", "曇/雨", "雨/曇", "晴/雨"] };
}
function fixSite(s) {
  s.id ||= uid(); s.名前 ||= ""; s.担当者 ||= ""; s.始業 ||= "8:00"; s.終業 ||= "17:00";
  if (s.締め日 == null) s.締め日 = 20;
  s.累計 ||= "期間"; s.作業員 ||= []; s.機材 ||= []; s.日 ||= {}; s.消した日 ||= {}; s.設定更新 ||= "";
  s.別 ||= [];   // 揺らぎで「別のもの」とした組 [[ゆるい形a, ゆるい形b]]（2026-10-31：秋元悠斗・秋元和斗は兄弟）
  if (s.単価 && s.単価.length && !s.単価.every(isNewPrice)) s.単価 = migratePrices(s.単価);   // 前の形の単価を分類の形に
  s.勤務 ||= "昼"; s.夜始業 ||= "21:00"; s.夜終業 ||= "6:00";   // 2026-10-31：昼・夜・昼夜の現場
  return s;
}
function fixDB(d) {
  const b = blankDB();
  for (const k in b) if (d[k] == null) d[k] = b[k];
  d.職種単価 ||= {}; d.共通更新 ||= ""; d.会社一覧 ||= [];
  d.職種別名 ||= {}; d.単価たたき台 ||= null;   // 単価のたたき台 { 元: 現場名, 行: [単価マスタの行] }（2026-10-31：ほかの現場はこれを写して見直す）   // まとめた工種の古い名前 → 新しい名前（2026-10-31：土工 → 普通作業員）。打った・読んだ古い名前は新しい名前にする   // 会社・所有者の共通の一覧（2026-10-31）
  d.現場.forEach(fixSite);
  return d;
}
let DB = (() => { try { const s = localStorage.getItem(KEY); if (s) return fixDB(JSON.parse(s)); } catch (e) {} return fixDB(blankDB()); })();
/* 工種の別名を日報に当てる（2026-10-31 ユーザー：土工＝普通作業員に直す）。
   別名（古い名前 → 新しい名前）に「土工 → 普通作業員」をいつも入れる。一覧から古い名前を外し、全部の現場の日報・名簿・単価の工種を新しい名前にする。
   直した日は 更新 を新しくして共有に書く。共通.json の更新日時は変えない（新しい端末の空の一覧で共有を上書きしないため）。 */
const SELF = "ハシックス", OLD_SELF = "ハシックス（自社）";   // 2026-10-31 ユーザー：会社名は「ハシックス」だけに（前の書き方は開いたときに直す）
const KOU_ALIAS0 = { 土工: "普通作業員" };
function aliasKou(x, A = DB.職種別名) {
  let ch = false;
  Object.values(x.日).forEach(d => {
    let c2 = false;
    d.下.forEach(r => { if (r && A[r[3]]) { r[3] = A[r[3]]; c2 = true; } });
    d.下.forEach(r => { if (r) [12, 13].forEach(i => { if (r[i] === OLD_SELF) { r[i] = SELF; c2 = true; } }); });
    (d.車 || []).forEach(r => { if (r && r[5] === OLD_SELF) { r[5] = SELF; c2 = true; } });
    d.上.forEach(r => { if (r && A[r[0]]) { r[0] = A[r[0]]; c2 = true; } });
    if (c2) { d.更新 = now(); ch = true; }
  });
  x.作業員.forEach(w => { if (A[w.職種]) { w.職種 = A[w.職種]; ch = true; } if (w.会社 === OLD_SELF) { w.会社 = SELF; ch = true; } });
  x.機材.forEach(m => { if (m.所有者 === OLD_SELF) { m.所有者 = SELF; ch = true; } });
  (x.単価 || []).forEach(p => { if (p.所有者 === OLD_SELF) { p.所有者 = SELF; ch = true; } });
  (x.単価 || []).forEach(p => { if (p.区分 === "労務" && A[p.名前]) { p.名前 = A[p.名前]; ch = true; } });
  if (compApplyToDays(x, true)) ch = true;   // 会社が空の行に名前一覧の会社
  if (x.職種) { const k = [...new Set(x.職種.map(v => A[v] || v))]; if (k.join("\n") !== x.職種.join("\n")) { x.職種 = k; ch = true; } }
  if (ch) x.設定更新 = now();
  return ch;
}
function kouFix() {
  for (const [a, b] of Object.entries(KOU_ALIAS0)) if (!DB.職種別名[a] && a !== b) DB.職種別名[a] = b;
  const A = DB.職種別名;
  DB.職種 = [...new Set(DB.職種.map(k => A[k] || k))];
  if ((DB.会社一覧 || []).includes(OLD_SELF)) DB.会社一覧 = [...new Set(DB.会社一覧.map(k => k === OLD_SELF ? SELF : k))];
  for (const [a, b] of Object.entries(A)) if (DB.職種単価[a]) { if (!DB.職種単価[b]) DB.職種単価[b] = DB.職種単価[a]; delete DB.職種単価[a]; }
  let n = 0; DB.現場.forEach(x => { if (aliasKou(x, A)) n++; });
  return n;
}
// kouFix() は下の方の定義（nk・companyOf など）を使うので、ファイルの終わり（buildSheet の前）で呼ぶ
let UI = (() => { try { return JSON.parse(localStorage.getItem(UIKEY)) || {}; } catch (e) { return {}; } })();
let saveT = null;
function save() {
  clearTimeout(saveT);
  saveT = setTimeout(() => {
    try { localStorage.setItem(KEY, JSON.stringify(DB)); $("saved").textContent = "ブラウザに保存 " + now().slice(11, 16); }
    catch (e) { toast("ブラウザに保存できません（容量がいっぱいかも）。「保存」でファイルに出してください"); }
    if (typeof syncSoon === "function") syncSoon();
  }, 250);
}
// 日を消す（ほかの端末にも消したことが伝わるように時刻を残す）
function dropDay(s, d) { delete s.日[d]; s.消した日[d] = now(); }
function saveUI() { try { localStorage.setItem(UIKEY, JSON.stringify(UI)); } catch (e) {} }

const site = () => DB.現場.find(s => s.id === UI.site) || null;
const gen = () => DB.元請.find(g => g.id === UI.gen) || null;
const day = () => { const s = site(); return s ? s.日[UI.day] || null : null; };
function ensureDay() {
  const s = site(); if (!s) return null;
  if (!s.日[UI.day]) s.日[UI.day] = { 天気: "", 気温: "", 担当者: "", 上: [], 下: [], 車: [], 備考: "", サイン: "", サイン日時: "", 人員手: false, 更新: "" };
  return s.日[UI.day];
}
// 中身のある日か（空の日は保存しない）
function hasData(d) {
  if (!d) return false;
  if (d.天気 || d.気温 || d.備考 || d.サイン) return true;
  return [...d.上, ...d.下, ...(d.車 || [])].some(r => r && r.some(v => v !== "" && v != null));
}
function tidy(d) {
  // 後ろの空の行は持たない
  const trim = a => { while (a.length && !(a[a.length - 1] || []).some(v => v !== "" && v != null)) a.pop(); };
  trim(d.上); trim(d.下); if (d.車) trim(d.車);
}

/* ================= 締め期間 ================= */
// 締め日 20 → 21日〜翌月20日。締め日 0／31 → 月末締め（1日〜末日）
function period(dateStr, s = site()) {
  const D = s ? +s.締め日 : 20, d = pd(dateStr);
  if (!D || D >= 28) {
    const a = new Date(d.getFullYear(), d.getMonth(), 1), b = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    return [ymd(a), ymd(b)];
  }
  if (d.getDate() > D) return [ymd(new Date(d.getFullYear(), d.getMonth(), D + 1)), ymd(new Date(d.getFullYear(), d.getMonth() + 1, D))];
  return [ymd(new Date(d.getFullYear(), d.getMonth() - 1, D + 1)), ymd(new Date(d.getFullYear(), d.getMonth(), D))];
}
function datesIn(a, b) { const out = []; for (let x = a; x <= b; x = addD(x, 1)) out.push(x); return out; }

/* ================= 入力の直し（IME は Chrome・Edge では切り替えられないので、全角で来ても半角に直して受ける） ================= */
const han = v => String(v ?? "").replace(/[０-９：．－，]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)).replace(/　/g, " ").trim();
function normTime(v) {
  v = han(v).replace(/[^\d:.]/g, "");
  if (!v) return "";
  let m = v.match(/^(\d{1,2}):(\d{1,2})$/);
  if (m) return +m[1] + ":" + pad(+m[2]);
  m = v.match(/^(\d{1,2})\.(\d+)$/);               // 17.5 → 17:30
  if (m) return +m[1] + ":" + pad(Math.round(Number("0." + m[2]) * 60));
  if (/^\d{1,2}$/.test(v)) return +v + ":00";      // 8 → 8:00
  if (/^\d{3,4}$/.test(v)) return +v.slice(0, -2) + ":" + v.slice(-2);   // 1730 → 17:30
  return v;
}
const toMin = t => { const m = String(t || "").match(/^(\d{1,2}):(\d{2})$/); return m ? +m[1] * 60 + +m[2] : null; };
function normNum(v) { v = han(v).replace(/,/g, ""); if (v === "") return ""; const x = Number(v.replace(/[^\d.\-]/g, "")); return isNaN(x) || /[^\d.\-]/.test(v) ? v : String(x); }
function normNin(v) { v = han(v); if (v === "") return ""; const m = v.match(/^(\d+)\s*(名|人)?$/); return m ? m[1] + "名" : v; }
const ninNum = v => { const m = String(v || "").match(/(\d+)/); return m ? +m[1] : 0; };
function normalize(kind, v) {
  if (kind === "time") return normTime(v);
  if (kind === "num" || kind === "hrs") return normNum(v);
  if (kind === "nin") return normNin(v);
  if (kind === "shift") { v = han(v).normalize("NFKC"); return /^(夜|よる|n|ﾖ)/i.test(v) ? "夜" : /^(昼|ひる|d|ﾋ|日)/i.test(v) ? "昼" : ""; }
  return String(v ?? "").replace(/[\t\r\n]+/g, " ").trim();
}
// 残業時間（時間）＝ 定時の終業より遅い分＋定時の始業より早い分。夜をまたぐ終業は＋24時間
// 行の昼夜（空なら現場の勤務。夜の現場は夜、ほかは昼）
const shiftOf = (r, s = site()) => (r && r[14]) || (s && s.勤務 === "夜" ? "夜" : "昼");
const teiji = (s, sh) => sh === "夜" ? [s.夜始業, s.夜終業] : [s.始業, s.終業];
// 早出・残業（分）。夜勤（21:00〜6:00 など日をまたぐ）も、終業が始業より前なら次の日として数える
function otParts(st, en, s = site(), sh = "昼") {
  const [ts, te] = teiji(s, sh);
  const a = toMin(st), b0 = toMin(en), A = toMin(ts), B0 = toMin(te);
  if (a == null || b0 == null || A == null || B0 == null) return null;
  const b = b0 < a ? b0 + 1440 : b0, B = B0 <= A ? B0 + 1440 : B0;
  // 夜勤で 0:30 に始めたような書き方（始業が定時より12時間以上前）は次の日の時刻とみなす
  const a2 = A - a > 720 ? a + 1440 : a, b2 = A - a > 720 ? b + 1440 : b;
  return { 早出: Math.max(0, A - a2), 残業: Math.max(0, b2 - B) };
}
function overtime(st, en, s = site(), sh = "昼") {
  const o = otParts(st, en, s, sh); if (!o) return "";
  const m = o.早出 + o.残業;
  return m ? String(Math.round(m / 60 * 100) / 100) : "";
}

/* ================= 表の列 ================= */
const UP_ROWS = 5, DN_ROWS = 16;
// 上の表は作業内容だけ5行（2026-10-31 ユーザー）。データは前と同じ 上[行] = [工種, 作業内容, 人員, 資機材] の 作業内容（1）に入れる
// 備考は 上[行][4]（2026-10-31 ユーザー：作業内容の隣に備考の列）。右の工種別の集計を小さくした分、左を広げた
const UPC = [
  { t: "作業予定", w: 438, k: "txt", di: 1 },   // 2026-10-31 ユーザー：作業内容 → 作業予定、備考 → 完了報告。幅は同じ
  { t: "完了報告", w: 438, k: "txt", di: 4 }
];
// di＝データの何列目か（2026-10-31 会社の欄を足した。データは後ろ（12）に足し、画面では作業員名の右に出す）
// 幅は紙（#sheet）に収まるように詰めた（2026-10-31 ユーザー：用紙からはみ出た）
// 累計・累計数量の列は出さない（2026-10-31 ユーザー）。空いた幅はほかの列に配った。データの列（di）は前と同じ
const DNC = [
  { t: "作業員名", w: 120, k: "txt", stopFlow: true },
  { t: "会社", w: 110, k: "txt", di: 12, stopFlow: true },
  { t: "昼夜", w: 36, k: "shift", cls: "c", align: "c", di: 14 },
  { t: "始業時間", w: 54, k: "time", cls: "c", align: "c" },
  { t: "終業時間", w: 54, k: "time", cls: "c", align: "c" },
  { t: "工種", w: 78, k: "txt", cls: "c", align: "c" },
  { t: "残", w: 44, k: "hrs", cls: "r", align: "r" },
  { t: "使用機材", w: 126, k: "txt", cls: "sepL" },
  { t: "規格・形式", w: 100, k: "txt" },
  { t: "所有者", w: 110, k: "txt", di: 13, stopFlow: true },
  { t: "数量", w: 42, k: "num", cls: "r", align: "r", di: 7 },
  { t: "材料名", w: 122, k: "txt", cls: "sepL", di: 9 },
  { t: "数量", w: 42, k: "num", cls: "r", align: "r", di: 10 },
  // 実費（円）：高速代（ルートで変わる）・領収書精算など、その日の実際の金額。入れると単価×数量の代わりにこの金額（2026-10-31 ユーザー）
  { t: "実費", w: 56, k: "num", cls: "r", align: "r", di: 15 }
];
/* 高速・燃料（2026-10-31 ユーザー：高速を使ったか（往路・復路・往復）と区間、燃料を何リットル入れたか（軽油・ガソリン）、
   だれ（会社・名前・車両名・車番）を日報に書き、集計に出す。A4 に収めるため既定は3行）。日[日付].車 = [[列0〜9]] */
const CAR_ROWS = 3;   // 1ページの行数（PG.car と同じ）
const CARC = [
  { t: "種類", w: 56, k: "txt", cls: "c", align: "c" },          // 0 高速／軽油／ガソリン
  { t: "往復", w: 50, k: "txt", cls: "c", align: "c" },          // 1 往路／復路／往復（高速）
  { t: "高速の区間", w: 150, k: "txt" },                         // 2
  { t: "給油(L)", w: 54, k: "num", cls: "r", align: "r" },       // 3 燃料のリットル
  { t: "実費", w: 60, k: "num", cls: "r", align: "r" },          // 4 実際の金額（入れると単価の代わり）
  { t: "会社", w: 110, k: "txt", cls: "sepL" },                  // 5
  { t: "名前", w: 110, k: "txt" },                               // 6
  { t: "車両名", w: 120, k: "txt" },                             // 7
  { t: "車番", w: 90, k: "txt" },                                // 8
  { t: "メモ", w: 294, k: "txt" }                                // 9
];
CARC.forEach((c, i) => c.letter = String.fromCharCode(65 + i));
const ARR = { up: "上", dn: "下", car: "車" };   // 表 → 日のデータの名前
/* 1日の入力ページ（2026-10-31 ユーザー：作業内容・人員・機械が多くて1ページに収まらない日は、入力ページを足せる）。
   ページ p は 上[p*5〜]・下[p*16〜]・車[p*3〜] の行（集計はページに関係なく全部の行）。d.頁＝ページの数（足した空のページも数える）。
   何も書いていないページは、ページ・日・現場を移るときに「消しますか」と聞く。 */
const PG = { up: 5, dn: 16, car: 3 };
function pageCount(d) { if (!d) return 1; return Math.max(1, d.頁 || 1, ...Object.entries(ARR).map(([k, a]) => Math.ceil(((d[a] || []).length) / PG[k]))); }
const pgOff = key => (UI.pg || 0) * PG[key];
function pageOf(d, p) { const o = { ...d }; for (const [k, a] of Object.entries(ARR)) o[a] = (d[a] || []).slice(p * PG[k], (p + 1) * PG[k]); return o; }
function pageEmpty(d, p) { const o = pageOf(d, p); return !Object.values(ARR).some(a => o[a].some(r => r && r.some(v => v !== "" && v != null))); }
function removePage(d, p) {
  const n = pageCount(d);
  for (const [k, a] of Object.entries(ARR)) { const arr = d[a] || []; if (arr.length > p * PG[k]) arr.splice(p * PG[k], PG[k]); }
  d.頁 = Math.max(1, n - 1); d.更新 = now();
}
// 何も書いていない足したページを消すか聞く（only：そのページだけ）
function askEmptyPages(only) {
  const s = site(), d = day(); if (!s || !d || !canEdit()) return;
  const ps = [...Array(pageCount(d)).keys()].filter(p => p > 0 && (only == null || p === only) && pageEmpty(d, p));
  if (!ps.length) return;
  if (!confirm(`${jpDate(UI.day)} の ${ps.map(p => p + 1).join("・")}ページ目には何も書いていません。このページを消しますか？`)) return;
  pushUndo(); ps.reverse().forEach(p => removePage(d, p)); tidy(d);
  UI.pg = Math.min(UI.pg || 0, pageCount(d) - 1);
  if (!hasData(d) && pageCount(d) === 1) dropDay(s, UI.day);
  save(); toast("何も書いていないページを消しました");
}
function goPage(p) {
  for (const k of ["up", "dn", "car"]) if (grids[k] && grids[k].editing()) grids[k].commit();
  const s = site(), d = day(), was = UI.pg || 0;
  p = Math.max(0, Math.min(pageCount(d) - 1, p));
  if (p === was) return;
  // 何も書いていない足したページから移るとき
  if (s && d && was > 0 && pageEmpty(d, was) && canEdit() && confirm(`${was + 1}ページ目には何も書いていません。このページを消しますか？`)) {
    pushUndo(); removePage(d, was); tidy(d); save(); if (p > was) p--;
  }
  UI.pg = Math.max(0, Math.min(pageCount(day()) - 1, p)); refreshSheet();
}
UPC.forEach((c, i) => c.letter = String.fromCharCode(65 + i));
// 画面の順のうち、di の無い列はデータの 0〜11 を順に当てる（会社＝12・所有者＝13 は後ろに足した列）
{ let k = 0; DNC.forEach((c, i) => { c.letter = String.fromCharCode(65 + i); if (c.di == null) c.di = k++; }); }
const di = (key, c) => key === "dn" ? DNC[c].di : key === "up" ? (UPC[c].di ?? c) : c;   // 画面の列 → データの列
const SECS = { up: UPC, dn: DNC, car: CARC };

// 累計：同じ「使用機材＋規格」（品名）の数量を、締め期間（または工事の始め）からこの日まで足す
function cumul(kind, r) {
  const s = site(), d = day(); if (!s || !d) return "";
  const row = d.下[r] || [];
  const ki = kind === "kizai" ? 5 : 9, qi = kind === "kizai" ? 7 : 10;
  if (!row[ki]) return "";
  const key = kind === "kizai" ? row[5] + "|" + (row[6] || "") : row[9];
  const from = s.累計 === "全体" ? "0000-00-00" : period(UI.day)[0];
  let sum = 0, any = false;
  for (const [dt, dd] of Object.entries(s.日)) {
    if (dt < from || dt > UI.day) continue;
    dd.下.forEach(x => {
      if (!x || !x[ki]) return;
      const k = kind === "kizai" ? x[5] + "|" + (x[6] || "") : x[9];
      if (k === key) { const q = Number(x[qi]); if (!isNaN(q) && x[qi] !== "") { sum += q; any = true; } }
    });
  }
  return any ? String(Math.round(sum * 1000) / 1000) : "";
}
function getCell(key, r, c) {
  const d = day(); if (!d) return "";
  if (key === "dn" && c === 8) return cumul("kizai", r);
  if (key === "dn" && c === 11) return cumul("hin", r);
  return (((d[ARR[key]] || [])[r] || [])[c]) ?? "";
}

// 作業員の工種（職種）ごとの人数。名簿の職種の順
function countByShoku(d) {
  const m = new Map();
  (d ? d.下 : []).forEach(r => { if (r && r[0]) { const k = r[3] || "（工種なし）"; m.set(k, (m.get(k) || 0) + 1); } });
  const order = [...DB.職種];
  return [...m.entries()].sort((a, b) => { const ia = order.indexOf(a[0]), ib = order.indexOf(b[0]); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib); });
}
// 工種別の集計（右の欄・Excel・PDF）：{ rows: [{ k, n, 昼, 夜 }], 計, 夜あり }
function kouSum(d) {
  const m = new Map();
  (d ? d.下 : []).forEach(r => { if (r && r[0]) { const k = r[3] || "（工種なし）", x = m.get(k) || { 昼: 0, 夜: 0 }; x[shiftOf(r) === "夜" ? "夜" : "昼"]++; m.set(k, x); } });
  const rows = countByShoku(d).map(([k, n]) => ({ k, n, ...m.get(k) }));
  return { rows, 計: rows.reduce((a, x) => a + x.n, 0), 夜あり: rows.some(x => x.夜) };
}
function kouSumHTML(d, title) {
  const K = kouSum(d), nc = K.夜あり ? 4 : 2, tt = title ? `<tr><th class="kt" colspan="${nc}">工種別の集計</th></tr>` : "";
  if (!K.rows.length) return title ? `<table class="ksum"><thead>${tt}<tr><th>工種</th><th>人員</th></tr></thead><tbody><tr><td class="ksE" colspan="2">下の作業員から数えます</td></tr></tbody></table>` : "";
  return `<table class="ksum"><thead>${tt}<tr><th>工種</th>${K.夜あり ? "<th>昼</th><th>夜</th>" : ""}<th>人員</th></tr></thead><tbody>` +
    K.rows.map(x => `<tr><td>${esc(x.k)}</td>${K.夜あり ? `<td class="r">${x.昼 || ""}</td><td class="r night">${x.夜 || ""}</td>` : ""}<td class="r">${x.n}名</td></tr>`).join("") +
    `</tbody><tfoot><tr><td>計</td>${K.夜あり ? `<td class="r">${K.rows.reduce((a, x) => a + x.昼, 0)}</td><td class="r">${K.rows.reduce((a, x) => a + x.夜, 0)}</td>` : ""}<td class="r">${K.計}名</td></tr></tfoot></table>`;
}

/* ================= 元に戻す ================= */
let UNDO = [], REDO = [];
function snap() { const d = day(); return { site: UI.site, day: UI.day, v: d ? JSON.stringify(d) : null }; }
function pushUndo() { UNDO.push(snap()); if (UNDO.length > 100) UNDO.shift(); REDO = []; }
function restore(sn) {
  const s = DB.現場.find(x => x.id === sn.site); if (!s) return;
  if (sn.v) { s.日[sn.day] = JSON.parse(sn.v); s.日[sn.day].更新 = now(); } else dropDay(s, sn.day);   // 戻したものがいちばん新しい
  if (UI.site !== sn.site || UI.day !== sn.day) { UI.site = sn.site; UI.day = sn.day; }
  save(); renderAll();
}
function undo() { if (!canEdit()) return; const sn = UNDO.pop(); if (!sn) return; REDO.push(snap()); restore(sn); }
function redo() { if (!canEdit()) return; const sn = REDO.pop(); if (!sn) return; UNDO.push(snap()); restore(sn); }

/* ================= 会社（2026-10-31 ユーザー：元の Excel は外注の人も名前だけで会社名が無い。
   会社を書く欄を足し、会社ごとの集計と会社別の表で、だれが来ていたか分かるように） =================
   ・作業員名を入れると会社も入る：この現場の名簿（作業員[].会社）→ ほかの現場の名簿 → 配置表の名簿（社員・従業員＝自社、外注業者＝その名前）。
   ・日報の会社の欄を直すと、この現場の名簿の会社も直る（次からその人はその会社で入る）。
   ・「現場・名簿」の「名前一覧・会社の振り分け」で、日報に出てくる名前をまとめて振り分け、日報に入れる。CSV（名前,会社）でも読める。 */
function companyOf(name, s = site()) {
  const k = nk(name); if (!k) return "";
  const hit = list => (list || []).find(w => nk(w.名前) === k && w.会社);
  const w = hit(s && s.作業員) || DB.現場.map(x => hit(x.作業員)).find(Boolean);
  if (w) return w.会社;
  if (MS.社員.some(n => nk(n) === k) || MS.従業員.some(n => nk(n) === k)) return SELF;
  const g = MS.外注業者.find(n => nk(n) === k); if (g) return g;
  return "";
}
function companies(s = site()) {
  const set = new Set([SELF, ...(DB.会社一覧 || [])]);
  DB.現場.forEach(x => x.作業員.forEach(w => w.会社 && set.add(w.会社)));
  if (s) Object.values(s.日).forEach(d => d.下.forEach(r => r && r[12] && set.add(r[12])));
  MS.外注業者.forEach(n => set.add(n));
  return [...set];
}
/* 使用機材の所有者（2026-10-31 ユーザー：使用機材名・規格・形式・所有者名（リース会社、自社、外注業者）。
   機材名の頭の（ハ）（真）は所有者に置き換えるので消す） */
const LEASE = ["東リース", "レント", "セフテック", "第一産商", "アクティオ", "カナモト", "西尾レントオール", "ニッケン", "コクド", "トーコー産業"];
// 頭の字の決まり（自社だけ。外注業者は名簿・会社の一覧から頭の字で当てる。公開するので実在の会社名は書かない）
const ABBR = { "ハ": SELF, "自": SELF };
// 頭の（ハ）などを外す → { 名前, 所有者 }
function splitOwner(name) {
  const m = String(name || "").match(/^\s*[（(]\s*([^）)]{1,6})\s*[）)]\s*(.+)$/);
  if (!m) return { 名前: String(name || "").trim(), 所有者: "" };
  const a = m[1], hit = ABBR[a] || companies().find(c => c.startsWith(a)) || LEASE.find(c => c.startsWith(a)) || a;
  return { 名前: m[2].trim().replace(/^([^（(]*)[）)]+$/, "$1"), 所有者: hit };   // 元の Excel の「（真）通勤車）」のような余分なかっこも外す
}
const abbrOf = o => !o ? "" : o === SELF ? "ハ" : (Object.entries(ABBR).find(([, v]) => v === o) || [o.charAt(0)])[0];
function ownerOf(name, spec, s = site()) {
  const k = nk(name), ks = nk(spec);
  const inList = list => (list || []).find(m => nk(m.名前) === k && nk(m.規格) === ks && m.所有者) || (list || []).find(m => nk(m.名前) === k && m.所有者);
  const m = inList(s && s.機材) || DB.現場.map(x => inList(x.機材)).find(Boolean);
  if (m) return m.所有者;
  if (MS.機械.some(x => nk(x.名前) === k)) return SELF;   // 自社機械稼働表にある機械
  return "";
}
const owners = (s = site()) => [...new Set([SELF, ...companies(s), ...LEASE, ...DB.現場.flatMap(x => x.機材.map(m => m.所有者).filter(Boolean))])];
// 行の使用機材（頭の（ハ）を外して所有者へ）。直したら true
function fixKizaiRow(r, s) {
  if (!r || !r[5]) return false;
  while (r.length < 14) r.push("");
  const sp = splitOwner(r[5]);
  let ch = false;
  if (sp.所有者) { r[5] = sp.名前; if (!r[13]) r[13] = sp.所有者; ch = true; }
  if (!r[13]) { const o = ownerOf(r[5], r[6], s); if (o) { r[13] = o; ch = true; } }
  return ch;
}
function rememberKizai(s, r) {
  if (!r || !r[5]) return;
  let m = s.機材.find(x => x.名前 === r[5] && (x.規格 || "") === (r[6] || "") && (x.所有者 || "") === (r[13] || ""));
  if (!m) m = s.機材.find(x => x.名前 === r[5] && (x.規格 || "") === (r[6] || "") && !x.所有者);
  if (!m) s.機材.push({ 名前: r[5], 規格: r[6] || "", 所有者: r[13] || "" });
  else if (!m.所有者 && r[13]) m.所有者 = r[13];
}
/* ================= 単価マスタ（2026-10-31 ユーザー：機械・労務の集計に金額。単価マスターを作り、工種別・機械種類別） =================
   現場ごとに 単価:[{ 区分:"労務"|"残業"|"機械"|"材料"|"その他", 名前, 規格, 所有者, 単位, 単価, 開始(YYYY-MM-DD・空＝ずっと), メモ }]（共有で一緒に合わせる）。
   労務＝作業員の行の工種（1人・1日）、残業＝工種（無ければ「残業」の行）×時間、機械＝使用機材の名前＋規格＋所有者（無ければ名前＋規格 → 名前）×数量（空なら1台）、材料＝品名×数量。
   単価が変わったら、同じ項目の行を足して 開始 に変わった日を入れる（その日からはその単価）。 */
/* 2026-10-31 ユーザー：単価表を 労務（昼・夜・残業）／機械（車両・重機・小型機械・電動工具）／材料（固定単価・領収書精算）／通勤費（高速代＝区間ごと・通勤車）に分けた。
   行の形：労務 { 区分:"労務", 名前:工種, 昼, 夜, 残業, 開始, 終了, メモ }（工種「（共通）」の行は工種の行に無い単価に使う）
           ほか { 区分, 分類, 名前, 規格, 所有者, 単位, 単価, 開始, 終了, メモ }（材料の領収書精算は単価なし＝日報の数量に金額を入れる。高速代の 名前＝区間）
   前の形（区分 労務／残業／機械／材料／その他 で 単価・昼夜）は開いたときに migratePrices で移す。 */
const PT_K = [["名前", "名前"], ["規格", "規格・形式"], ["所有者", "所有者"], ["単位", "単位"], ["単価", "単価（円）", "n"], ["開始", "開始", "d"], ["終了", "終了", "d"], ["メモ", "メモ"]];
const PT = [
  { 区分: "労務", 分類: "", 題: "労務（工種ごと）", note: "昼・夜は1人1日、残業は1時間。工種「（共通）」の行を作ると、工種の行に無い単価はそちらを使います。夜が空なら昼の単価。", cols: [["名前", "工種"], ["昼", "昼（円/日）", "n"], ["夜", "夜（円/日）", "n"], ["残業", "残業（円/時）", "n"], ["開始", "開始", "d"], ["終了", "終了", "d"], ["メモ", "メモ"]] },
  { 区分: "機械", 分類: "車両", 題: "機械（車両）", cols: PT_K },
  { 区分: "機械", 分類: "重機", 題: "機械（重機）", cols: PT_K },
  { 区分: "機械", 分類: "小型機械", 題: "機械（小型機械）", cols: PT_K },
  { 区分: "機械", 分類: "電動工具", 題: "機械（電動工具）", cols: PT_K },
  { 区分: "材料", 分類: "固定単価", 題: "材料（固定単価）", cols: [["名前", "材料名"], ["規格", "規格"], ["単位", "単位"], ["単価", "単価（円）", "n"], ["開始", "開始", "d"], ["終了", "終了", "d"], ["メモ", "メモ"]] },
  { 区分: "材料", 分類: "領収書精算", 題: "材料（領収書精算）", note: "単価は決めません。日報の材料名にこの名前を書き、「実費」の欄に領収書の金額（円）を入れると、その金額をそのまま足します（同じ品でも日によって金額が違ってよい）。実費が空なら数量の欄を金額とみます（前の書き方）。", cols: [["名前", "材料名"], ["メモ", "メモ"]] },
  { 区分: "通勤費", 分類: "高速代", 題: "通勤費（高速代・区間ごと）", note: "日報の材料名に「高速代（区間）」か区間を書き、数量に回数を入れます。区間が空の行は、区間を書かなかったときに使います。ルートで金額が違う日は、日報の「実費」の欄に実際の金額（円）を入れると、単価×回数の代わりにそちらを使います。", cols: [["名前", "区間"], ["単価", "単価（円/回）", "n"], ["開始", "開始", "d"], ["終了", "終了", "d"], ["メモ", "メモ"]] },
  { 区分: "通勤費", 分類: "通勤車", 題: "通勤費（通勤車）", note: "日報の使用機材に書いた通勤車を、所有者ごとに 台数×単価。", cols: [["名前", "名前"], ["規格", "規格"], ["所有者", "所有者"], ["単価", "単価（円/台日）", "n"], ["開始", "開始", "d"], ["終了", "終了", "d"], ["メモ", "メモ"]] }
];
const PK = ["労務", "機械", "材料", "通勤費"];
const PK_CSV = {
  労務: ["工種", "昼", "夜", "残業", "開始", "終了", "メモ"],
  機械: ["分類", "名前", "規格", "所有者", "単位", "単価", "開始", "終了", "メモ"],
  材料: ["分類", "品名", "規格", "単位", "単価", "開始", "終了", "メモ"],
  通勤費: ["分類", "名前", "規格", "所有者", "単位", "単価", "開始", "終了", "メモ"]
};
// 機械の分類のあたり（名前から）
function guessMach(n) {
  n = String(n || "").normalize("NFKC");
  if (/通勤/.test(n)) return "通勤車";
  if (/バックホ|BH|クレーン|ラフター|ローラー|ブルドーザ|ショベル|フィニッシャ|グレーダ|杭打|ユンボ/i.test(n)) return "重機";
  if (/ダンプ|トラック|ユニック|車|バン|キャリア/.test(n)) return "車両";
  if (/電動|ブレーカ|カッター|セイバー|ドリル|サンダ|グラインダ|インパクト|丸ノコ|丸のこ|チェーンソー|工具/.test(n)) return "電動工具";
  return "小型機械";
}
// この現場で使う工種（2026-10-31 ユーザー：現場ごとに工種を決める。決めていない現場は共通の一覧）
const siteKou = (s = site()) => s && s.職種 && s.職種.length ? s.職種 : DB.職種;
const isNewPrice = p => p.分類 != null || (p.区分 === "労務" && ("昼" in p || "夜" in p || "残業" in p));
function migratePrices(list) {
  if (!list || !list.length || list.every(isNewPrice)) return list || [];
  const out = list.filter(isNewPrice).map(p => ({ ...p }));
  const labor = (name, 開始, 終了) => { let r = out.find(p => p.区分 === "労務" && p.名前 === name && (p.開始 || "") === (開始 || "") && (p.終了 || "") === (終了 || "")); if (!r) { r = { 区分: "労務", 名前: name, 昼: "", 夜: "", 残業: "", 開始: 開始 || "", 終了: 終了 || "", メモ: "" }; out.push(r); } return r; };
  for (const p of list.filter(x => !isNewPrice(x))) {
    const base = { 開始: p.開始 || "", 終了: p.終了 || "", メモ: p.メモ || "" };
    if (p.区分 === "労務") { const r = labor(p.名前, p.開始, p.終了); r[p.昼夜 === "夜" ? "夜" : "昼"] = p.単価; if (p.メモ && !r.メモ) r.メモ = p.メモ; }
    else if (p.区分 === "残業") { const r = labor(p.名前 === "残業" ? "（共通）" : p.名前, p.開始, p.終了); r.残業 = p.単価; if (p.メモ && !r.メモ) r.メモ = p.メモ; }
    else if (p.区分 === "機械") { const g = guessMach(p.名前); out.push({ 区分: g === "通勤車" ? "通勤費" : "機械", 分類: g, 名前: p.名前, 規格: p.規格 || "", 所有者: p.所有者 || "", 単位: p.単位 || "台日", 単価: p.単価, ...base }); }
    else if (p.区分 === "材料") out.push({ 区分: "材料", 分類: "固定単価", 名前: p.名前, 規格: p.規格 || "", 所有者: "", 単位: p.単位 || "", 単価: p.単価, ...base });
    else if (/高速/.test(p.名前)) out.push({ 区分: "通勤費", 分類: "高速代", 名前: "", 規格: "", 所有者: "", 単位: p.単位 || "回", 単価: p.単価, ...base });
    else if (/電動工具|工具/.test(p.名前)) out.push({ 区分: "機械", 分類: "電動工具", 名前: p.名前, 規格: p.規格 || "", 所有者: p.所有者 || "", 単位: p.単位 || "日", 単価: p.単価, ...base });
    else out.push({ 区分: "材料", 分類: "固定単価", 名前: p.名前, 規格: p.規格 || "", 所有者: "", 単位: p.単位 || "", 単価: p.単価, ...base });
  }
  return out;
}
const byStart = (a, b) => (b.開始 || "").localeCompare(a.開始 || "");
const numOk = v => v !== "" && v != null && !isNaN(Number(v));
// 労務・残業の単価（現場の単価表 → 共通の工種の単価）
function laborPrice(s, kou, sh, date, kind) {
  const rows = (s.単価 || []).filter(p => p.区分 === "労務" && inPeriod(p, date));
  const pick = n => rows.filter(p => nk(p.名前) === nk(n)).sort(byStart);
  const val = p => kind === "残業" ? p.残業 : (sh === "夜" ? (numOk(p.夜) ? p.夜 : p.昼) : p.昼);
  for (const p of [...pick(kou), ...pick("（共通）")]) { const v = val(p); if (numOk(v)) return { 単価: Number(v) }; }
  return commonPrice(kou, sh, kind === "残業" ? "残業" : "日", date);
}
// 機械・材料・通勤費の単価：{ 単価, 区分, 分類, 精算 }
function itemPrice(s, kubuns, name, spec, owner, date) {
  const nm = nk(name), area = nk(String(name || "").replace(/^\s*高速(道路)?(代|料金)?\s*[（(]?/, "").replace(/[）)]\s*$/, ""));
  for (const k of kubuns) {
    const rows = (s.単価 || []).filter(p => p.区分 === k && inPeriod(p, date) && (p.分類 === "領収書精算" || numOk(p.単価)));
    const hit = rows.filter(p => p.分類 === "高速代" ? (nk(p.名前) === nm || (p.名前 && nk(p.名前) === area) || (!p.名前 && /高速/.test(name))) : nk(p.名前) === nm);
    if (!hit.length) continue;
    const best = f => hit.filter(f).sort((a, b) => (a.名前 ? 0 : 1) - (b.名前 ? 0 : 1) || byStart(a, b))[0];
    const p = best(p => (p.規格 || p.所有者) && nk(p.規格) === nk(spec) && nk(p.所有者) === nk(owner))
      || best(p => p.規格 && !p.所有者 && nk(p.規格) === nk(spec))
      || best(p => !p.規格 && p.所有者 && nk(p.所有者) === nk(owner))
      || best(p => !p.規格 && !p.所有者);
    if (p) return { 単価: p.分類 === "領収書精算" ? null : Number(p.単価), 区分: p.区分, 分類: p.分類 || "", 精算: p.分類 === "領収書精算" };
  }
  return null;
}
// 工種の略（作業員名別の表のます目に出す）
const SHORT = { "管理": "管", "職長": "職", "重機OP": "OP", "土工": "土", "普通作業員": "普", "草刈り": "草", "溶接工": "溶", "整備": "整", "（工種なし）": "○" };
const SHORT_NOTE = "管＝管理・職＝職長・OP＝重機OP・普＝普通作業員・土＝土工・草＝草刈り（ほかは頭の字）・夜＝夜勤";
const shortK = k => { const n = /（夜）$/.test(k), b = String(k).replace(/（夜）$/, ""); return (SHORT[b] || String(b || "○").normalize("NFKC").slice(0, 1)) + (n ? "夜" : ""); };
// 共通の職種単価（現場の単価マスタに行が無いときに使う）。夜の単価が空なら昼の単価
// 期間：DB.職種単価[工種].期間 = [{ 開始, 終了, 昼, 夜, 残業 }]（2026-10-31 ユーザー：何月何日から何月何日まではこの単価）。開始・終了が空ならその側は限りなし
const inPeriod = (p, date) => !date || ((!p.開始 || p.開始 <= date) && (!p.終了 || date <= p.終了));
function commonPrice(k, sh, kind, date) {
  const base = DB.職種単価[k]; if (!base) return null;
  const per = (base.期間 || []).filter(x => inPeriod(x, date)).sort((a, b) => (b.開始 || "").localeCompare(a.開始 || ""))[0];
  const p = per || base;
  const v = kind === "残業" ? p.残業 : (sh === "夜" ? (p.夜 !== "" && p.夜 != null ? p.夜 : p.昼) : p.昼);
  return v === "" || v == null || isNaN(Number(v)) ? null : { 単価: Number(v), 共通: true };
}
const yen = v => (Math.round(v) || 0).toLocaleString("ja-JP");
const compOfRow = (r, s) => (r && (r[12] || companyOf(r[0], s))) || "（会社未定）";
function setRosterCompany(s, name, comp) {
  let w = s.作業員.find(x => nk(x.名前) === nk(name));
  if (!w) { w = { 名前: masterName(name), 職種: "" }; s.作業員.push(w); }
  if (w.会社 !== comp) { const old = w.会社 || ""; w.会社 = comp; s.設定更新 = now(); compToDays(s, w.名前, old, comp); }
}
// 名前一覧の会社を日報に入れる（2026-10-31 ユーザー：名前一覧・会社の振り分けを日報に反映）。
// その人の行で、会社が空か 前の会社のままの行を新しい会社にする（日報で別の会社を書いた行はそのまま）
function compToDays(s, name, old, comp) {
  let n = 0;
  for (const d of Object.values(s.日)) {
    let ch = false;
    d.下.forEach(r => { if (r && r[0] && nk(r[0]) === nk(name) && (!r[12] || r[12] === old) && r[12] !== comp) { while (r.length < 13) r.push(""); r[12] = comp; n++; ch = true; } });
    if (ch) d.更新 = now();
  }
  return n;
}

/* ================= 揺らぎ対策（2026-10-31 ユーザー：人の名前・使用機材・材料など、同じものを違う名前にしないように） =================
   見る列：作業員名・工種（上と下）・使用機材・規格・品名。
   ・打った言葉を、ゆるい形（全角半角・空白・中黒・かっこ・ハイフン類・英字の大小を無視）で、これまでの書き方（マスタ・この現場の日報）と比べる。
     ゆるい形が同じ → これまでの書き方にそろえる（だまって直し、下に「そろえました」と出す）。
     1字ちがい・片方がもう片方を含む → 「○○のことですか？」と聞く（そろえる／別のものとして入れる）。
   ・「現場・名簿」の「揺らぎを見つけてそろえる」で、これまでの日報の中の揺らぎをまとめて直せる。 */
const YCOLS = { dn: { 0: "作業員名", 12: "会社", 3: "工種", 5: "使用機材", 6: "規格・形式", 13: "所有者", 9: "材料名" }, up: {} };
const loose = v => String(v ?? "").normalize("NFKC").toLowerCase().replace(/[\s　・･()（）\[\]「」【】,，.．]/g, "").replace(/[-‐－―−ｰ]/g, "ー");
function lev(a, b) {
  if (Math.abs(a.length - b.length) > 2) return 9;
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return d[a.length][b.length];
}
// 似ているか（ゆるい形で 1字ちがい（8字以上は2字）・3字以上で片方がもう片方を含む）
const pairKey = (a, b) => [loose(a), loose(b)].sort().join("\u0001");
const isDistinct = (a, b) => { const k = pairKey(a, b); return DB.現場.some(x => (x.別 || []).some(p => p.join("\u0001") === k)); };
function markDistinct(s, list) {
  let n = 0;
  for (let i = 0; i < list.length; i++) for (let j = i + 1; j < list.length; j++) {
    if (loose(list[i]) === loose(list[j]) || isDistinct(list[i], list[j])) continue;
    (s.別 ||= []).push(pairKey(list[i], list[j]).split("\u0001")); n++;
  }
  if (n) s.設定更新 = now();
  return n;
}
function similar(a, b) {
  const x = loose(a), y = loose(b);
  if (!x || !y || x === y) return x === y;
  if (isDistinct(a, b)) return false;
  const n = Math.min(x.length, y.length);
  if (n >= 3 && (x.includes(y) || y.includes(x))) return true;
  return n >= 2 && lev(x, y) <= (n >= 8 ? 2 : 1);
}
// その列のこれまでの書き方（マスタを先に。回数つき）
function canon(key, c, s = site()) {
  const cnt = new Map(), add = (v, w = 1) => { v = String(v || "").trim(); if (v) cnt.set(v, (cnt.get(v) || 0) + w); };
  if (key === "dn" && c === 0) { people().forEach(p => add(p.n, 1000)); s && s.作業員.forEach(w => add(w.名前, 10)); }
  if (c === 3 && key === "dn" || key === "up" && c === 0) siteKou(s).forEach(k => add(k, 1000));
  if (key === "dn" && c === 5) { MS.機械.forEach(m => add(m.名前, 1000)); s && s.機材.forEach(m => add(m.名前, 10)); }
  if (key === "dn" && c === 6) { MS.機械.forEach(m => add(m.規格, 1000)); s && s.機材.forEach(m => add(m.規格, 10)); }
  if (key === "dn" && c === 12) companies(s).forEach(x => add(x, 1000));
  if (key === "dn" && c === 13) owners(s).forEach(x => add(x, 1000));
  if (s) for (const dd of Object.values(s.日)) (dd[ARR[key]] || []).forEach(r => r && add(r[c]));
  return cnt;
}
// 打った言葉を見る：{ v：入れる言葉, fixed：だまってそろえた元の言葉, ask：[似ている書き方] }
function checkWord(key, c, v, cn) {
  v = String(v ?? "").trim();
  if (!v || !(YCOLS[key] && YCOLS[key][c])) return { v };
  if (cn.has(v)) return { v };
  const lv = loose(v);
  const same = [...cn.entries()].filter(([w]) => loose(w) === lv).sort((a, b) => b[1] - a[1]);
  if (same.length) return { v: same[0][0], fixed: v };
  const ask = [...cn.entries()].filter(([w]) => similar(v, w)).sort((a, b) => b[1] - a[1]).slice(0, 5).map(x => x[0]);
  return { v, ask };
}
let YQ = [];   // 聞くこと：{ key, r, c, v, ask }
function askYure() {
  if (!YQ.length) return;
  const q = YQ;
  $("yrBody").innerHTML = `<p class="note">似た書き方がもうあります。同じものなら、これまでの書き方にそろえてください（集計で別々に数えないように）。</p><table class="sm"><thead><tr><th>欄</th><th>入れた言葉</th><th>どれにしますか</th></tr></thead><tbody>` +
    q.map((x, i) => `<tr><td>${esc(YCOLS[x.key][x.c])}</td><td><b>${esc(x.v)}</b></td><td>${x.ask.map((w, j) => `<label class="yr"><input type="radio" name="yr${i}" value="${esc(w)}"${j === 0 ? " checked" : ""}> ${esc(w)}</label>`).join("")}<label class="yr"><input type="radio" name="yr${i}" value=""> 別のもの（このまま・もう聞かない）</label></td></tr>`).join("") + `</tbody></table>`;
  openDlg($("dYure"));
  $("yrOk").onclick = () => {
    const ch = { up: [], dn: [] };
    const s0 = site(); let nd = 0;
    q.forEach((x, i) => { const w = (document.querySelector(`input[name="yr${i}"]:checked`) || {}).value; if (w) ch[x.key].push({ r: x.r, c: x.c, v: w }); else { if (s0) nd += markDistinct(s0, [x.v, ...x.ask]); if (x.isNew && s0 && !(s0.足さない || []).includes(`${x.key}:${x.c}:${loose(x.v)}`)) NQ.push({ key: x.key, c: x.c, v: x.v, 欄: NCOLS[x.key][x.c] }); } });
    if (nd) { save(); toast("別のものとして覚えました（もう聞きません）"); }
    closeDlg($("dYure")); YQ = [];
    for (const k of ["up", "dn"]) if (ch[k].length) applyCells(k, ch[k], true);
    if (NQ.length) setTimeout(askNew, 50);
  };
}

/* ---------- 新しい言葉（2026-10-31 ユーザー：選択肢＋自由入力。自由入力は揺らぎを確かめ、新しいものはマスタに足すか聞く） ----------
   工種 → 工種（職種）の一覧（共通）／作業員名 → この現場の名簿／会社・所有者 → 会社・所有者の一覧（共通）／使用機材 → この現場の機材の一覧。
   「足さない」と答えたものは現場の 足さない に覚えて、もう聞かない。 */
const NCOLS = { dn: { 0: "作業員名", 3: "工種", 5: "使用機材", 9: "材料名", 12: "会社", 13: "所有者" }, up: {} };
function inMaster(key, c, v, s) {
  const k = loose(v); if (!k) return true;
  const has = list => list.some(x => loose(x) === k);
  if (key === "up" || c === 3) return has(siteKou(s));
  if (c === 0) return has(people().map(p => p.n)) || has(s.作業員.map(w => w.名前));
  if (c === 12 || c === 13) return has(owners(s));
  if (c === 5) return has(MS.機械.map(m => m.名前)) || DB.現場.some(x => has(x.機材.map(m => m.名前)));
  if (c === 9) return !!itemPrice(s, ["材料", "通勤費"], v, "", "", "") || (s.単価 || []).some(p => (p.区分 === "材料" || p.区分 === "通勤費") && loose(p.名前) === k);
  return true;
}
let NQ = [];   // { key, c, v, 欄 }
function askNew() {
  if (!NQ.length) return;
  const s = site(), q = NQ.filter((x, i) => NQ.findIndex(y => y.key === x.key && y.c === x.c && loose(y.v) === loose(x.v)) === i);
  NQ = [];
  const where = x => x.key === "up" || x.c === 3 ? (s.職種 && s.職種.length ? "この現場の工種の一覧（単価も聞きます）" : "工種（職種）の一覧（全部の現場）・単価も聞きます") : x.c === 0 ? "この現場の作業員名簿" : x.c === 5 ? "この現場の使用機材の一覧・単価" : x.c === 9 ? "この現場の単価表（単価か領収書精算）" : "会社・所有者の一覧（全部の現場）";
  $("nwBody").innerHTML = `<p class="note">一覧（マスタ）に無い言葉です。次から選べるように、一覧に足しますか？</p><table class="sm"><thead><tr><th>欄</th><th>入れた言葉</th><th>足す先</th><th>足す</th></tr></thead><tbody>` +
    q.map((x, i) => `<tr><td>${esc(x.欄)}</td><td><b>${esc(x.v)}</b></td><td class="note">${esc(where(x))}</td><td class="c"><input type="checkbox" data-nw="${i}" checked></td></tr>`).join("") + `</tbody></table>`;
  openDlg($("dNew"));
  const done = add => { const newKou = [], newItems = [];
    let nA = 0, nN = 0;
    q.forEach((x, i) => {
      const yes = add && document.querySelector(`[data-nw="${i}"]`).checked;
      if (yes) {
        nA++;
        if (x.key === "up" || x.c === 3) { addSiteKou(s, x.v, true); newKou.push(x.v); }
        else if (x.c === 12 || x.c === 13) { if (!(DB.会社一覧 ||= []).some(k => loose(k) === loose(x.v))) { DB.会社一覧.push(x.v); DB.共通更新 = now(); } }
        else if (x.c === 5 || x.c === 9) newItems.push({ 種: x.c === 5 ? "機械" : "材料", 名前: x.v });   // 単価（材料は領収書精算も）を続けて聞く
        // 作業員名・使用機材は日報に入れたときにもう名簿・一覧に入っている
      } else {
        nN++;
        (s.足さない ||= []).push(`${x.key}:${x.c}:${loose(x.v)}`);
        if (x.c === 0 && x.key === "dn") s.作業員 = s.作業員.filter(w => nk(w.名前) !== nk(x.v) || w.会社 || Object.values(s.日).filter(d => d.下.some(r => r && nk(r[0]) === nk(x.v))).length > 1);
        if (x.c === 5 && x.key === "dn") s.機材 = s.機材.filter(m => loose(m.名前) !== loose(x.v) || m.所有者);
        s.設定更新 = now();
      }
    });
    closeDlg($("dNew")); save();
    const next = () => { if (newItems.length) askItemPrice(s, newItems); };
    if (newKou.length) setTimeout(() => askKouPrice(s, newKou, next), 120); else setTimeout(next, 120);
    toast(nA ? `${nA} 件を一覧に足しました` + (nN ? `（${nN} 件は足さない）` : "") : "一覧には足しませんでした（日報にはそのまま入っています）");
  };
  $("nwOk").onclick = () => done(true);
  $("nwNo").onclick = () => done(false);
}

$("dYure").addEventListener("close", () => {
  // 「このままにする」で閉じた（決めるを押さなかった）：入れた言葉はそのまま。一覧に無い言葉なら、続けて足すか聞く
  const s0 = site(), q = YQ; YQ = [];
  q.forEach(x => { if (x.isNew && s0 && !(s0.足さない || []).includes(`${x.key}:${x.c}:${loose(x.v)}`)) NQ.push({ key: x.key, c: x.c, v: x.v, 欄: NCOLS[x.key][x.c] }); });
  if (NQ.length) setTimeout(askNew, 80);
});   // 「このままにする」で閉じても、新しい言葉は続けて聞く

/* ================= セルを変える ================= */
function applyCells(key, ch, noCheck) {
  const s = site(); if (!s) { toast("先に「現場・名簿」で元請と現場を登録してください"); return; }
  if (!canEdit()) { refreshSheet(); return; }
  // 揺らぎ：打つ前のこれまでの書き方と比べる（同じならそろえる・似ていれば後で聞く）
  const fixedMsg = [];
  // まとめた工種の古い名前（土工 など）は新しい名前に
  ch = ch.map(x => { const isK = (key === "dn" && x.c === 3) || (key === "up" && x.c === 0), to = isK && x.v && DB.職種別名[String(x.v).trim()]; if (to) { fixedMsg.push(`「${x.v}」→「${to}」`); return { ...x, v: to }; } return x; });
  // 新しい言葉：打つ前に一覧（マスタ）に無かったもの（揺らぎで聞くものは、その答えの後で）
  const newCand = [];
  if (!noCheck && NCOLS[key]) ch.forEach(x => { if (NCOLS[key][x.c] && x.v && !inMaster(key, x.c, x.v, s) && !(s.足さない || []).includes(`${key}:${x.c}:${loose(x.v)}`)) newCand.push(x); });
  if (!noCheck && YCOLS[key]) {
    const cns = {};
    ch = ch.map(x => {
      if (!YCOLS[key][x.c] || !x.v) return x;
      const r = checkWord(key, x.c, x.v, cns[x.c] ||= canon(key, x.c, s));
      if (r.fixed) fixedMsg.push(`「${r.fixed}」→「${r.v}」`);
      if (r.ask && r.ask.length) YQ.push({ key, r: x.r, c: x.c, v: r.v, ask: r.ask, isNew: NCOLS[key] && NCOLS[key][x.c] && !inMaster(key, x.c, r.v, s) });
      return { ...x, v: r.v };
    });
  }
  pushUndo();
  const d = ensureDay(), arr = (d[ARR[key]] ||= []), NC = key === "up" ? 5 : SECS[key].length;
  const touched = new Set();
  for (const { r, c, v } of ch) {
    if (r >= PG[key] * pageCount(d)) continue;
    while (arr.length <= r) arr.push(Array(NC).fill(""));
    const row = arr[r]; while (row.length < NC) row.push("");
    row[c] = key === "dn" && c === 14 ? normalize("shift", v) : key === "car" && c <= 1 ? carNorm(c, v) : v; touched.add(r);
    if (key === "up" && c === 2) d.人員手 = true;
  }
  if (key === "dn") {
    for (const r of touched) {
      const row = arr[r], c0 = ch.filter(x => x.r === r).map(x => x.c);
      // 名前を入れたら：名簿の工種・定時の始業／終業
      if (c0.includes(0) && row[0]) {
        row[0] = masterName(row[0]);   // 名簿の書き方（「吉田　康之」）にそろえる
        const w = s.作業員.find(x => nk(x.名前) === nk(row[0]));
        if (!row[3] && w) row[3] = w.職種 || "";
        if (!row[14] && !c0.includes(14)) row[14] = s.勤務 === "夜" ? "夜" : "昼";
        { const [ts, te] = teiji(s, shiftOf(row, s)); if (!row[1]) row[1] = ts; if (!row[2]) row[2] = te; }
        if (!row[12] && !c0.includes(12)) row[12] = companyOf(row[0], s);
      }
      if (c0.includes(12) && row[0] && row[12]) setRosterCompany(s, row[0], row[12]);
      // 昼夜を変えたら、定時のままの始業・終業はその昼夜の定時に入れ替える
      if (c0.includes(14) && row[0]) {
        const other = teiji(s, shiftOf(row, s) === "夜" ? "昼" : "夜"), now2 = teiji(s, shiftOf(row, s));
        if (!c0.includes(1) && (!row[1] || row[1] === other[0])) row[1] = now2[0];
        if (!c0.includes(2) && (!row[2] || row[2] === other[1])) row[2] = now2[1];
      }
      if (c0.includes(1) || c0.includes(2) || c0.includes(14) || (c0.includes(0) && row[0])) {
        if (!c0.includes(4)) row[4] = overtime(row[1], row[2], s, shiftOf(row, s));
      }
      // 自社機械の名前を入れたら規格（1つに決まるとき）
      if (c0.includes(5) && row[5]) { const sp = splitOwner(row[5]); if (sp.所有者) { row[5] = sp.名前; if (!c0.includes(13)) row[13] = sp.所有者; } }
      if (c0.includes(5) && row[5] && !row[6]) { const ms = MS.機械.filter(m => nk(m.名前) === nk(row[5])); if (ms.length === 1) row[6] = ms[0].規格; }
      if ((c0.includes(5) || c0.includes(6)) && row[5] && !row[13] && !c0.includes(13)) row[13] = ownerOf(row[5], row[6], s);
      // この現場の名簿・機材に足す（工種を覚える。次から一覧に出る）
      if (row[0] && !s.作業員.some(x => nk(x.名前) === nk(row[0]))) s.作業員.push({ 名前: row[0], 職種: row[3] || "" });
      else if (row[0] && row[3]) { const w = s.作業員.find(x => nk(x.名前) === nk(row[0])); if (w && !w.職種) w.職種 = row[3]; }
      rememberKizai(s, row);
    }
  }
  if (key === "car") for (const r of touched) carFill(s, arr[r], ch.filter(x => x.r === r).map(x => x.c));
  tidy(d); d.更新 = now();
  if (!hasData(d)) dropDay(s, UI.day);
  save(); refreshSheet();
  if (fixedMsg.length) toast("これまでの書き方にそろえました：" + fixedMsg.slice(0, 3).join("・") + (fixedMsg.length > 3 ? " ほか" : ""), 4500);
  // 揺らぎでそろえた・揺らぎで聞くものは除いて、新しい言葉を聞く
  newCand.forEach(x => {
    const now3 = (d[ARR[key]] || [])[x.r], v = now3 && now3[x.c];
    if (!v || inMasterBefore(x, v) || YQ.some(y => y.key === key && y.r === x.r && y.c === x.c)) return;
    NQ.push({ key, c: x.c, v, 欄: NCOLS[key][x.c] });
  });
  function inMasterBefore(x, v) { return loose(v) !== loose(x.v); }   // そろえて別の言葉になった＝もう一覧にある
  if (YQ.length) setTimeout(askYure, 50);
  else if (NQ.length) setTimeout(askNew, 50);
}
// 高速・燃料：種類と往復の書き方をそろえる
function carNorm(c, v) {
  v = String(v ?? "").trim(); if (!v) return v;
  const z = v.normalize("NFKC");
  if (c === 0) return /高速/.test(z) ? "高速" : /混合/.test(z) ? v : /軽油|ディーゼル/.test(z) ? "軽油" : /ガソ|レギュラー|ハイオク/.test(z) ? "ガソリン" : v;
  return /往復|両/.test(z) ? "往復" : /^往|行き/.test(z) ? "往路" : /^復|帰り/.test(z) ? "復路" : v;
}
// 名前を入れたら会社（名前一覧）・前に乗った車両名と車番。車両名を入れたら前の車番
function carFill(s, row, c0) {
  const prev = f => { for (const x of Object.keys(s.日).sort().reverse()) for (const r of (s.日[x].車 || [])) if (r && r !== row && f(r)) return r; return null; };
  if (c0.includes(6) && row[6]) {
    row[6] = masterName(row[6]);
    if (!row[5] && !c0.includes(5)) row[5] = companyOf(row[6], s) || "";
    if (!row[7] && !c0.includes(7)) { const p = prev(r => nk(r[6]) === nk(row[6]) && r[7]); if (p) { row[7] = p[7]; if (!row[8]) row[8] = p[8] || ""; } }
  }
  if (c0.includes(7) && row[7] && !row[8] && !c0.includes(8)) { const p = prev(r => nk(r[7]) === nk(row[7]) && r[8]); if (p) row[8] = p[8]; }
}
// 高速・燃料の1行の金額：実費があればそれ。高速＝区間の単価（通勤費の高速代）×回数（往復は2）、燃料＝材料の単価×リットル
function carAmount(s, r, date) {
  const jv = han(String(r[4] ?? "")).replace(/[,，円\s]/g, ""), jh = jv !== "" && !isNaN(Number(jv)) ? Number(jv) : null;
  if (r[0] === "高速") {
    const n = r[1] === "往復" ? 2 : 1, p = itemPrice(s, ["通勤費"], "高速代" + (r[2] ? "（" + r[2] + "）" : ""), "", "", date);
    return { 種: "高速", 回: n, L: 0, 金額: jh != null ? jh : p ? p.単価 * n : 0, 単価: p ? p.単価 : null, 実費: jh != null, なし: jh == null && !p };
  }
  const L = Number(han(String(r[3] ?? "")).replace(/[,，\s]/g, "")) || 0, p = r[0] ? itemPrice(s, ["材料"], r[0], "", "", date) : null;
  return { 種: r[0] || "", 回: 0, L, 金額: jh != null ? jh : p ? p.単価 * L : 0, 単価: p ? p.単価 : null, 実費: jh != null, なし: jh == null && !p };
}
// だれ（会社・名前・車両名・車番）ごと
function carByWho(C) {
  const by = new Map();
  for (const x of C) {
    const k = x.who.join("\t"), v = by.get(k) || { who: x.who, 往路: 0, 復路: 0, 往復: 0, 高速金額: 0, 軽油: 0, ガソリン: 0, 他L: 0, 燃料金額: 0 };
    if (x.A.種 === "高速") { v[["往路", "復路", "往復"].includes(x.r[1]) ? x.r[1] : "往路"]++; v.高速金額 += x.A.金額; }
    else { if (x.r[0] === "軽油" || x.r[0] === "ガソリン") v[x.r[0]] += x.A.L; else v.他L += x.A.L; v.燃料金額 += x.A.金額; }
    by.set(k, v);
  }
  return [...by.values()].sort((a, b) => a.who.join("\t").localeCompare(b.who.join("\t"), "ja"));
}
function carByArea(C) {
  const by = new Map();
  for (const x of C) if (x.A.種 === "高速") { const k = x.r[2] || "（区間なし）", v = by.get(k) || { 区間: k, 往路: 0, 復路: 0, 往復: 0, 回: 0, 金額: 0 }; v[["往路", "復路", "往復"].includes(x.r[1]) ? x.r[1] : "往路"]++; v.回 += x.A.回; v.金額 += x.A.金額; by.set(k, v); }
  return [...by.values()];
}
function rowOp(key, op, r, n) {
  const s = site(); if (!s || !canEdit()) return;
  if (key === "up" && op === "insert") { toast("作業予定・完了報告は5行です。足りないときは「＋ ページを足す」"); return; }
  r += pgOff(key);   // 表の行 → データの行（いまのページの分ずらす）
  pushUndo();
  const d = ensureDay(), arr = (d[ARR[key]] ||= []), NC = key === "up" ? 5 : SECS[key].length;
  if (op === "insert") { while (arr.length < r) arr.push(Array(NC).fill("")); arr.splice(r, 0, ...Array.from({ length: n }, () => Array(NC).fill(""))); if (key === "up") arr.length = Math.min(arr.length, UP_ROWS); }
  else arr.splice(r, n);
  tidy(d); if (!hasData(d)) dropDay(s, UI.day);
  save(); refreshSheet();
}

// その列の候補（Alt+↓・打ち始めの続き）
function colWords(key, c) {
  const s = site(); if (!s) return [];
  const cnt = new Map(), add = (v, w = 1) => { if (v) cnt.set(v, (cnt.get(v) || 0) + w); };
  if (key === "dn" && c === 0) {
    // マスタ（配置表の名簿：社員・従業員・外注業者）＋この現場で使った人。名前の書き方はマスタにそろえる
    s.作業員.forEach(w => add(masterName(w.名前), 100));
    MS.社員.forEach(n => add(n, 60)); MS.従業員.forEach(n => add(n, 60)); MS.外注業者.forEach(n => add(n, 20));
  }
  if ((key === "dn" && c === 3) || (key === "up" && c === 0)) siteKou(s).forEach(k => add(k, 50));
  if (key === "dn" && c === 5) { s.機材.forEach(m => add(m.名前, 100)); MS.機械.forEach(m => add(m.名前, 50)); }
  if (key === "dn" && c === 12) companies(s).forEach(x => add(x, 100));
  if (key === "dn" && c === 13) owners(s).forEach(x => add(x, 100));
  if (key === "dn" && c === 6) { s.機材.forEach(m => add(m.規格, 100)); MS.機械.forEach(m => add(m.規格, 50)); }
  if (key === "dn" && c === 14) { add("昼", 1000); add("夜", 999); }
  if (key === "car") {
    if (c === 0) ["高速", "軽油", "ガソリン"].forEach((k, i) => add(k, 1000 - i));
    if (c === 1) ["往路", "復路", "往復"].forEach((k, i) => add(k, 1000 - i));
    if (c === 2) (s.単価 || []).forEach(p => { if (p.分類 === "高速代" && p.名前) add(p.名前, 100); });
    if (c === 5) companies(s).forEach(x => add(x, 100));
    if (c === 6) { s.作業員.forEach(w => add(masterName(w.名前), 100)); MS.社員.forEach(n => add(n, 60)); MS.従業員.forEach(n => add(n, 60)); MS.外注業者.forEach(n => add(n, 20)); }
    if (c === 7) { s.機材.forEach(m => add(m.名前, 50)); MS.機械.forEach(m => add(m.名前, 30)); }
  }
  if (key === "dn" && (c === 1 || c === 2)) [s.始業, s.終業, "8:00", "17:00", "18:00", "19:00", "20:00"].forEach(t => add(t, 1));
  if (key === "up" && c === 2) countByShoku(day()).forEach(([k, n]) => add(n + "名", 100));
  for (const dd of Object.values(s.日)) (dd[ARR[key]] || []).forEach(r => r && add(r[c]));
  return [...cnt.entries()].sort((a, b) => b[1] - a[1]).map(x => x[0]).slice(0, 60);
}

/* ================= 画面：日報 ================= */
const grids = {};
let cur = "dn";
function setCur(k) {
  cur = k;
  document.querySelectorAll(".sec").forEach(x => x.classList.toggle("cur", x.dataset.key === k));
}
function buildSheet() {
  $("sheet").innerHTML = `
    <div class="ttlRow">作　業　日　報<span id="pgNo"></span></div>
    <table class="hd">
      <colgroup><col style="width:90px"><col><col style="width:70px"><col style="width:150px"><col style="width:44px"><col style="width:84px"><col style="width:84px"><col style="width:70px"><col style="width:170px"></colgroup>
      <tr><td class="l">元請業者</td><td class="v" id="hGen"></td><td class="l">日　付</td><td class="v dpick" id="hDate" title="押すとカレンダーから日を選べます"></td><td class="v" id="hWd"></td>
          <td class="l">天気</td><td class="l">気温</td><td class="l" rowspan="2">記入者</td><td rowspan="2"><span class="combo"><input id="hTan" autocomplete="off" title="記入者（▾ で一覧から選ぶか、そのまま打つ）"><button type="button" class="cbtn" id="hTanBtn" tabindex="-1" title="一覧から選ぶ">▾</button></span></td></tr>
      <tr><td class="l">現場名</td><td class="v" id="hSite" colspan="4"></td>
          <td><select id="hWx" title="天気（一覧は「現場・名簿」で直せる）"></select></td>
          <td><span class="tmp"><input id="hTmp" type="text" inputmode="decimal" autocomplete="off" title="気温（数字だけ。マイナスと小数点も可）">℃</span></td></tr>
    </table>
    <div id="tanList" class="clist" hidden></div>
    <div class="mid">
      <div class="left sec" data-key="up">
        <div class="secT"><b>作業予定・完了報告</b><span class="cnt" id="cntUp"></span></div>
        <div class="g1" id="g_up"></div>
      </div>
      <div class="right" id="kouSum"></div>
    </div>
    <div class="low sec" data-key="dn">
      <div class="secT"><b>作業員・使用機材・材料</b><span class="cnt" id="cntDn"></span></div>
      <div id="g_dn"></div>
    </div>
    <div class="low sec" data-key="car">
      <div class="secT"><b>高速・燃料</b><span class="note">高速＝往路・復路・往復と区間／燃料＝軽油・ガソリンと給油(L)。だれ（会社・名前・車両名・車番）も。実費を入れるとその金額</span><span class="cnt" id="cntCar"></span></div>
      <div id="g_car"></div>
    </div>`;
  for (const key of ["up", "dn", "car"]) {
    const cols = SECS[key];
    grids[key] = createGrid({
      host: $("g_" + key), scroller: $("sheetWrap"), cols, headRows: 1, tabBlocks: key === "dn" /* 作業員・使用機材・材料名のまとまりごとに Tab で折り返す */, rowBase: 1, rowHeadW: 0, noLetters: true,   // 左端の番号の列は出さない（2026-10-31 ユーザー）
     
      formula: $("fx"), isCur: () => cur === key, onFocus: () => setCur(key),
      rows: () => PG[key],   // 1ページの行数。足りなければページを足す
      head: () => `<tr class="xlhr sh"><th class="xlrh"></th>${cols.map((c, i) => `<th data-col="${i}" class="${(c.cls || "").includes("sepL") ? "sepL" : ""}">${esc(c.t)}</th>`).join("")}</tr>`,
      get: (r, c) => getCell(key, r + pgOff(key), di(key, c)),
      deco: (r, c) => {
        if (key !== "dn") return null;
        const d = day(), row = d && d.下[r + pgOff("dn")];
        if (!row) return null;
        c = di(key, c);
        if (c === 12 && row[0] && !row[12]) return { cls: "nocomp", title: "会社が決まっていません（現場・名簿の「名前一覧・会社の振り分け」で決められます）" };
        if (c === 4 && row[4]) { const sh = shiftOf(row), [ts, te] = teiji(site(), sh), o = otParts(row[1], row[2], site(), sh); return { cls: "ot", title: `${sh}の定時（${ts}〜${te}）を超えた時間` + (o ? `　早出 ${r2(o.早出 / 60)}・残業 ${r2(o.残業 / 60)}` : "") }; }
        if (c === 14 && shiftOf(row) === "夜") return { cls: "night" };
        return null;
      },
      normalize: (c, v) => normalize(cols[c].k, v),
      // 作業内容（上の欄）は自由記入だけ：選択肢・補い・▾ は出さない（2026-10-31 ユーザー）
      complete: (c, v) => { if (key === "up") return null; const lv = v.toLowerCase(), ws = colWords(key, di(key, c)); if (ws.some(x => x.toLowerCase() === lv)) return null; return ws.find(x => x.toLowerCase().startsWith(lv) && x.length > v.length) || null; },   // Excel と同じく、同じ言葉があれば補わない
      list: c => key === "up" ? [] : colWords(key, di(key, c)),
      hasList: c => key === "up" ? false : key === "car" ? [0, 1, 2, 5, 6, 7, 8].includes(c) : [0, 3, 5, 6, 9, 12, 13, 14].includes(di(key, c)),
      apply: ch => applyCells(key, ch.map(x => ({ ...x, r: x.r + pgOff(key), c: di(key, x.c) }))),
      undo, redo,
      onSheet: dir => moveDay(dir),
      onEdge: (dir, c) => {
        const ord = ["up", "dn", "car"], t = ord[ord.indexOf(key) + (dir > 0 ? 1 : -1)] || null;
        if (!t) return false;
        grids[t].enter(dir, Math.min(c, SECS[t].length - 1));
        return true;
      },
      rowOp: (op, r, n) => rowOp(key, op, r, n),
      menu: (x, y, sel) => openMenu(key, x, y, sel),
      print: () => printDays([UI.day]),
      toast,
      cellName: (r, c) => ({ up: "上", dn: "下", car: "高速・燃料" })[key] + (r + 1) + "・" + cols[c].t,
      onSelect: (name, val, stat) => {
        if (cur !== key) return;
        $("nameBox").value = name;
        if (document.activeElement !== $("fx")) $("fx").value = val ?? grids[key]?.editText() ?? "";
        $("stat").textContent = stat || "";
      },
      onEditText: t => { if (document.activeElement !== $("fx")) $("fx").value = t; },
      onMode: m => { $("mode").textContent = m === "edit" ? "編集" : m ? "入力" : "準備完了"; }
    });
  }
  setCur("dn");
  // 記入者：一覧から選ぶ＋自由に打てる（2026-10-31 ユーザー）。ブラウザの候補（datalist）は自動入力が混ざる・iPad で出にくいので自前の一覧
  const tanWords = () => { const s = site(); return [...new Set([s && s.担当者, ...(s ? Object.values(s.日).map(d => d.担当者) : []), ...MS.社員, ...(s ? s.作業員.map(x => masterName(x.名前)) : [])].filter(Boolean))]; };
  const showTan = all => {
    const q = nk($("hTan").value), ws = tanWords().filter(x => all || !q || nk(x).includes(q));
    if (!ws.length) { $("tanList").hidden = true; return; }
    $("tanList").innerHTML = ws.map(x => `<div data-v="${esc(x)}">${esc(x)}</div>`).join("");
    const r = $("hTan").closest("td").getBoundingClientRect();
    Object.assign($("tanList").style, { left: r.left + "px", top: r.bottom + "px", minWidth: r.width + "px" });
    $("tanList").hidden = false;
  };
  $("hTanBtn").addEventListener("pointerdown", e => { e.preventDefault(); if ($("tanList").hidden) { showTan(true); $("hTan").focus(); } else $("tanList").hidden = true; });
  $("hTan").addEventListener("input", () => showTan(false));
  $("hTan").addEventListener("keydown", e => {
    if (e.isComposing) return;
    if (e.key === "ArrowDown" && (e.altKey || $("tanList").hidden)) { e.preventDefault(); showTan(true); }
    else if (e.key === "Escape") $("tanList").hidden = true;
    else if (e.key === "Enter") { $("tanList").hidden = true; $("hTan").blur(); }
  });
  $("hTan").addEventListener("blur", () => setTimeout(() => { $("tanList").hidden = true; }, 150));
  $("tanList").addEventListener("pointerdown", e => {
    const d = e.target.closest("[data-v]"); if (!d) return;
    e.preventDefault(); $("hTan").value = d.dataset.v; $("tanList").hidden = true;
    $("hTan").dispatchEvent(new Event("change"));
  });
  $("sheetWrap").addEventListener("scroll", () => { $("tanList").hidden = true; });
  // 気温は数字だけ：打つたびに数字・小数点・マイナス以外を外す（日本語入力中は確定してから）
  const tmpOnly = v => { v = han(v).replace(/[ー−‐―]/g, "-").replace(/[^\d.\-]/g, ""); return v.replace(/(?!^)-/g, "").replace(/(\..*)\./g, "$1"); };
  let tmpIme = false;
  $("hTmp").addEventListener("compositionstart", () => { tmpIme = true; });
  $("hTmp").addEventListener("compositionend", () => { tmpIme = false; $("hTmp").value = tmpOnly($("hTmp").value); });
  $("hTmp").addEventListener("input", () => { if (!tmpIme) { const v = tmpOnly($("hTmp").value); if (v !== $("hTmp").value) $("hTmp").value = v; } });
  // 上の欄（天気・気温・記入者・備考）
  const hv = (id, f, norm) => $(id).addEventListener("change", () => {
    const s = site(); if (!s) return;
    if (!canEdit()) { refreshSheet(); return; }
    pushUndo();
    const d = ensureDay(), v = norm ? norm($(id).value) : $(id).value.trim();
    $(id).value = v; d[f] = v; d.更新 = now();
    if (f === "担当者" && v) s.担当者 = v;
    if (!hasData(d)) dropDay(s, UI.day);
    save(); renderTabs();
  });
  $("hDate").addEventListener("click", openCal);
  hv("hWx", "天気"); hv("hTmp", "気温", v => { v = tmpOnly(v); return v === "-" || v === "." ? "" : normNum(v); }); hv("hTan", "担当者", v => masterName(v.trim()));
}
function refreshSheet() {
  const s = site(), d = day();
  $("hGen").textContent = gen() ? gen().名前 : "（元請業者を選ぶ）";
  $("hSite").textContent = s ? s.名前 : signedIn() ? "（現場を選ぶ）" : "（右上の「サインイン」をすると現場が出ます）";
  $("hDate").textContent = jpDate(UI.day);
  const w = wd(UI.day);
  $("hWd").textContent = WD[w]; $("hWd").className = "v" + (w === 0 ? " sun" : w === 6 ? " sat" : "");
  $("hDate").className = "v" + (w === 0 ? " sun" : w === 6 ? " sat" : "");
  // 天気の一覧（いまの値が一覧に無いとき＝前に打ったもの・Excel から読んだものは、その値も出す）
  const wx = d ? d.天気 : "", wxs = DB.天気.includes(wx) || !wx ? DB.天気 : [...DB.天気, wx];
  $("hWx").innerHTML = `<option value="">　</option>` + wxs.map(x => `<option value="${esc(x)}">${esc(x)}</option>`).join("");
  $("hWx").value = wx;
  if (document.activeElement !== $("hTmp")) $("hTmp").value = d ? d.気温 : "";
  if (document.activeElement !== $("hTan")) $("hTan").value = d ? d.担当者 : "";
  $("hTan").placeholder = s && s.担当者 ? "前回 " + s.担当者 : "";   // 記入者は日によって替わる（2026-10-31 ユーザー）。前回の人は薄い字で見せるだけ
  { const n = pageCount(d); if ((UI.pg || 0) > n - 1) UI.pg = n - 1;
    $("pgNo").textContent = n > 1 ? `（${(UI.pg || 0) + 1}／${n}ページ）` : "";
    $("pgBar").innerHTML = `<button class="sub" data-pg="-1" title="前のページ"${UI.pg ? "" : " disabled"}>◀</button><span class="pgn">${(UI.pg || 0) + 1}／${n}ページ</span><button class="sub" data-pg="1" title="次のページ"${(UI.pg || 0) < n - 1 ? "" : " disabled"}>▶</button><button class="sub" id="pgAdd" title="作業予定・人員・機械が1ページに収まらないとき、この日の入力ページを足す">＋ ページを足す</button>${n > 1 && UI.pg ? `<button class="sub danger" id="pgDel" title="このページの行を消す">このページを消す</button>` : ""}`; }
  grids.up.render(); grids.dn.render(); grids.car.render();
  { const C = ((d && d.車) || []).filter(r => r && r[0]), hw = C.filter(r => r[0] === "高速").length, L = k => C.filter(r => r[0] === k).reduce((a, r) => a + (Number(r[3]) || 0), 0);
    $("cntCar").textContent = [hw ? `高速 ${hw}件` : "", L("軽油") ? `軽油 ${r2(L("軽油"))}L` : "", L("ガソリン") ? `ガソリン ${r2(L("ガソリン"))}L` : ""].filter(Boolean).join("・"); }
  $("kouSum").innerHTML = kouSumHTML(d, true);
  // 人数のまとめ
  const cnt = countByShoku(d), tot = cnt.reduce((a, x) => a + x[1], 0);
  $("cntDn").textContent = tot ? "作業員 " + cnt.map(([k, n]) => k + " " + n + "名").join("・") + "　計 " + tot + "名" : "";
  $("cntUp").textContent = "";
  $("status").textContent = (d && d.更新 ? "最後に直した " + d.更新 : s ? "まだ書いていない日" : "") + (s && ((s.工期開始 && UI.day < s.工期開始) || (s.工期終了 && UI.day > s.工期終了)) ? `　※ 工期（${s.工期開始 || ""}〜${s.工期終了 || ""}）の外の日です` : "") + (s && s.完了 ? "　※ 完了した現場です" : "");
  renderTabs();
}
function renderTabs() {
  const s = site(); if (!s) { $("tabs").innerHTML = ""; return; }
  const [a, b] = period(UI.day);
  const h = [`<span class="per">${+a.slice(5, 7)}/${+a.slice(8)}〜${+b.slice(5, 7)}/${+b.slice(8)}</span>`];
  for (const x of datesIn(a, b)) {
    const w = wd(x), dd = s.日[x];
    h.push(`<button data-d="${x}" class="${x === UI.day ? "on " : ""}${hasData(dd) ? "has " : ""}${dd && dd.サイン ? "signed " : ""}${w === 0 ? "sun" : w === 6 ? "sat" : ""}" title="${jpDate(x)}(${WD[w]})">${+x.slice(8)}(${WD[w]})</button>`);
  }
  $("tabs").innerHTML = h.join("");
  const on = $("tabs").querySelector(".on"); on && on.scrollIntoView({ block: "nearest", inline: "nearest" });
}
$("tabs").addEventListener("click", e => { const b = e.target.closest("[data-d]"); if (b) setDay(b.dataset.d); });
$("tPrev").onclick = () => setDay(addD(period(UI.day)[0], -1));
$("tNext").onclick = () => setDay(addD(period(UI.day)[1], 1));

function setDay(d) { if (d !== UI.day) { askEmptyPages(); UI.pg = 0; } if (grids.dn && grids.dn.editing()) grids.dn.commit(); if (grids.up && grids.up.editing()) grids.up.commit(); if (grids.car && grids.car.editing()) grids.car.commit(); UI.day = d; saveUI(); showDayBtn(); refreshSheet(); if (!$("pSum").hidden) renderSum(); }
function moveDay(n) { setDay(addD(UI.day, n)); }

/* ---------- 右クリックのメニュー ---------- */
function openMenu(key, x, y, sel) {
  closeMenu();
  const m = document.createElement("div"); m.className = "ctx"; m.id = "ctx";
  const items = [
    ["切り取り", "Ctrl+X", () => document.execCommand("cut")],
    ["コピー", "Ctrl+C", () => document.execCommand("copy")],
    ["貼り付け", "Ctrl+V", async () => { try { const t = await navigator.clipboard.readText(); grids[key].paste(t); } catch (e) { toast("Ctrl+V で貼り付けてください"); } }],
    null,
    ...(key === "up" ? [] : [["行の挿入", "Ctrl+Shift+＋", () => rowOp(key, "insert", sel.r0, sel.r1 - sel.r0 + 1)]]),   // 作業内容は5行だけなので出さない
    ["行の削除", "Ctrl+－", () => rowOp(key, "delete", sel.r0, sel.r1 - sel.r0 + 1)],
    ["数式と値のクリア", "Delete", () => { const ch = []; for (let r = sel.r0; r <= sel.r1; r++) for (let c = sel.c0; c <= sel.c1; c++) if (!SECS[key][c].ro) ch.push({ r: r + pgOff(key), c: di(key, c), v: "" }); applyCells(key, ch); }],
    ...(key === "up" ? [] : [null, ["候補から選ぶ", "Alt+↓", () => grids[key].showList()]])
  ];
  m.innerHTML = items.map((it, i) => it ? `<div data-i="${i}"><span>${it[0]}</span><kbd>${it[1]}</kbd></div>` : "<hr>").join("");
  document.body.appendChild(m);
  m.style.left = Math.min(x, innerWidth - m.offsetWidth - 4) + "px"; m.style.top = Math.min(y, innerHeight - m.offsetHeight - 4) + "px";
  m.addEventListener("mousedown", e => { e.preventDefault(); const d = e.target.closest("[data-i]"); if (d) { closeMenu(); items[+d.dataset.i][2](); grids[key].focus(); } });
}
function closeMenu() { const m = $("ctx"); m && m.remove(); }
document.addEventListener("pointerdown", e => { if (!e.target.closest("#ctx")) closeMenu(); });

/* ---------- 日報を探す（2026-10-31 ユーザー：名前・会社名・作業内容・職種などで探し、当てはまる日の一覧から その日の日報を出す） ---------- */
const FIND_COLS = [["", "どこでも"], ["名前", "作業員名"], ["会社", "会社"], ["作業", "作業予定"], ["工種", "工種（職種）"], ["機材", "使用機材・規格"], ["所有者", "所有者"], ["品名", "材料名"], ["備考", "完了報告"], ["記入者", "記入者"]];
function findIn(s, words, col) {
  const out = [];
  for (const [x, d] of Object.entries(s.日).sort((a, b) => a[0].localeCompare(b[0]))) {
    const hits = [];   // [欄, 値]
    const look = (f, v) => { v = String(v ?? "").trim(); if (v && (!col || col === f)) hits.push([f, v]); };
    d.上.forEach(r => { if (r) { look("作業", r[1]); look("備考", r[4]); } });   // 工種は下の作業員の工種だけ（作業内容から工種を数えない）
    d.下.forEach(r => {
      if (!r) return;
      look("名前", r[0]); look("会社", r[12] || (r[0] ? companyOf(r[0], s) : "")); look("工種", r[3]);
      look("機材", [r[5], r[6]].filter(Boolean).join(" ")); look("所有者", r[13]); look("品名", r[9]);
    });
    look("記入者", d.担当者);
    // 言葉は全部（空白で区切ったものはどれも）入っている日
    const all = hits.map(h => loose(h[1])).join("\n");
    if (!words.every(w => all.includes(w))) continue;
    const got = new Map();
    hits.forEach(([f, v]) => { if (words.some(w => loose(v).includes(w))) { const k = f + "\t" + v; got.set(k, (got.get(k) || 0) + 1); } });
    out.push({ s, x, d, got: [...got.entries()].map(([k, n]) => { const [f, v] = k.split("\t"); return { f, v, n }; }), 人: d.下.filter(r => r && r[0]).length });
  }
  return out;
}
let FIND = { q: "", col: "", all: false };
function renderFind() {
  const words = $("fQ").value.trim().split(/[\s　]+/).map(loose).filter(Boolean);
  FIND = { q: $("fQ").value, col: $("fCol").value, all: $("fAll").checked };
  if (!words.length) { $("fRes").innerHTML = `<p class="note">探したい言葉を入れてください（例：吉田、真創建、U字溝、重機OP）。空白で区切ると、全部が入っている日を出します。</p>`; return; }
  const sites = FIND.all ? DB.現場 : site() ? [site()] : [];
  const res = sites.flatMap(s => findIn(s, words, FIND.col));
  const lab = Object.fromEntries(FIND_COLS);
  $("fRes").innerHTML = `<p class="note">${res.length} 日 見つかりました${res.length ? "（行を押すとその日の日報を出します）" : ""}</p>` + (res.length ? `<table class="sm find"><thead><tr><th>日付</th><th>曜</th>${FIND.all ? "<th>現場</th>" : ""}<th>人員</th><th>当てはまったもの</th></tr></thead><tbody>` +
    res.map(r => `<tr class="${[ "sun", "", "", "", "", "", "sat" ][wd(r.x)]}${r.s.id === UI.site && r.x === UI.day ? " on" : ""}" data-fs="${esc(r.s.id)}" data-fx="${r.x}"><td>${jpDate(r.x)}</td><td class="c">${WD[wd(r.x)]}</td>${FIND.all ? `<td>${esc(siteLabel(r.s))}</td>` : ""}<td class="r">${r.人 || ""}</td><td>${r.got.map(g => `<span class="fchip"><i>${esc(lab[g.f] || g.f)}</i>${esc(g.v)}${g.n > 1 ? `<small>×${g.n}</small>` : ""}</span>`).join("")}</td></tr>`).join("") + `</tbody></table>` : "");
}
$("bFind").onclick = () => {
  $("fCol").innerHTML = FIND_COLS.map(([v, t]) => `<option value="${v}">${t}</option>`).join("");
  $("fQ").value = FIND.q; $("fCol").value = FIND.col; $("fAll").checked = FIND.all;
  renderFind(); openDlg($("dFind")); $("fQ").focus(); $("fQ").select();
};
{ let t = null; $("fQ").addEventListener("input", () => { clearTimeout(t); t = setTimeout(renderFind, 200); }); }
$("fCol").onchange = renderFind; $("fAll").onchange = renderFind;
$("fRes").addEventListener("click", e => {
  const tr = e.target.closest("[data-fx]"); if (!tr) return;
  const s = DB.現場.find(x => x.id === tr.dataset.fs); if (!s) return;
  if (s.id !== UI.site) { UI.gen = s.元請id; UI.site = s.id; UNDO = []; REDO = []; renderAll(); }
  closeDlg($("dFind")); showTab("sheet"); setDay(tr.dataset.fx);
  toast(`${siteLabel(s)}　${jpDate(tr.dataset.fx)} の日報です（もう一度「日報を探す」で一覧に戻れます）`, 4000);
});

/* ---------- ボタン ---------- */
$("bCopyPrev").onclick = () => {
  const s = site(); if (!s || !canEdit()) return;
  const prev = Object.keys(s.日).filter(x => x < UI.day && hasData(s.日[x])).sort().pop();
  if (!prev) { toast("前に書いた日がありません"); return; }
  const d = day();
  if (hasData(d) && !confirm(`この日の内容を ${jpDate(prev)} の内容で置き換えますか？`)) return;
  pushUndo();
  const p = s.日[prev], n = ensureDay();
  n.上 = JSON.parse(JSON.stringify(p.上)); n.下 = JSON.parse(JSON.stringify(p.下));
  n.担当者 = p.担当者; n.人員手 = p.人員手; n.更新 = now();
  save(); refreshSheet(); toast(jpDate(prev) + " を写しました（天気・気温・サイン・備考は写していません）");
};
$("pgBar").addEventListener("click", e => {
  const b = e.target.closest("button"); if (!b) return;
  if (b.dataset.pg) { goPage((UI.pg || 0) + Number(b.dataset.pg)); return; }
  if (!site() || !canEdit()) return;
  if (b.id === "pgAdd") { pushUndo(); const d = ensureDay(); d.頁 = pageCount(d) + 1; d.更新 = now(); UI.pg = d.頁 - 1; save(); refreshSheet(); toast(`${d.頁}ページ目を足しました（何も書かなければ、ほかへ移るときに消すか聞きます）`, 4000); return; }
  if (b.id === "pgDel") {
    const d = day(); if (!d || !UI.pg) return;
    if (!pageEmpty(d, UI.pg) && !confirm(`${UI.pg + 1}ページ目に書いた作業予定・完了報告・作業員・高速・燃料の行も消えます。このページを消しますか？`)) return;
    pushUndo(); removePage(d, UI.pg); tidy(d); UI.pg = Math.min(UI.pg, pageCount(d) - 1); save(); refreshSheet(); return;
  }
});
$("bInsRow").onclick = () => { const g = grids[cur], s = g.selection(); rowOp(cur, "insert", s.r0, s.r1 - s.r0 + 1); g.focus(); };
$("bDelRow").onclick = () => { const g = grids[cur], s = g.selection(); rowOp(cur, "delete", s.r0, s.r1 - s.r0 + 1); g.focus(); };
$("bUndo").onclick = undo; $("bRedo").onclick = redo;
$("bClearDay").onclick = () => {
  const s = site(); if (!s || !s.日[UI.day]) { toast("この日は書いていません"); return; }
  if (!canEdit()) return;
  if (!confirm(jpDate(UI.day) + " の日報を消しますか？（↶ で戻せます）")) return;
  pushUndo(); dropDay(s, UI.day); save(); refreshSheet();
};
$("bXlsDay").onclick = () => exportXlsx([UI.day], false);
$("bPdfDay").onclick = () => printDays([UI.day]);
$("bOut").onclick = () => {
  if (!site()) return;
  const [a, b] = period(UI.day);
  $("outFrom").value = a; $("outTo").value = b; outInfo();
  openDlg($("dOut"));
};
function outDates() {
  const s = site(), a = $("outFrom").value, b = $("outTo").value;
  if (!a || !b || a > b) return [];
  return datesIn(a, b).filter(x => !$("outOnlyData").checked || hasData(s.日[x]));
}
function outInfo() { $("outRange").textContent = `${fmtD($("outFrom").value)} 〜 ${fmtD($("outTo").value)}`; $("outInfo").textContent = outDates().length + " 日分を出します（1シート／1ページに1日）"; }
["outFrom", "outTo", "outOnlyData"].forEach(id => $(id).addEventListener("change", outInfo));
$("outXls").onclick = () => { const ds = outDates(); if (!ds.length) { toast("出す日がありません"); return; } closeDlg($("dOut")); exportXlsx(ds, $("outSum").checked, [$("outFrom").value, $("outTo").value]); };
$("outPdf").onclick = () => { const ds = outDates(); if (!ds.length) { toast("出す日がありません"); return; } closeDlg($("dOut")); printDays(ds, $("outSum").checked ? [$("outFrom").value, $("outTo").value] : null); };

/* ================= 元請サイン（iPad・スマホだけ） ================= */
let signDirty = false;
function openSign() {
  const cv = $("signPad");
  openDlg($("dSign"));
  const r = cv.getBoundingClientRect(), dpr = window.devicePixelRatio || 1;
  cv.width = Math.round(r.width * dpr); cv.height = Math.round(r.height * dpr);
  const g = cv.getContext("2d"); g.setTransform(dpr, 0, 0, dpr, 0, 0);
  g.lineWidth = 2.6; g.lineCap = "round"; g.lineJoin = "round"; g.strokeStyle = "#000";
  signDirty = false;
  $("signNote").textContent = "枠の中に指・ペンでサインしてください（" + jpDate(UI.day) + "）";
}
(() => {
  const cv = $("signPad"); let drawing = null;
  const pt = e => { const r = cv.getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; };
  cv.addEventListener("pointerdown", e => {
    if (e.pointerType === "mouse") { $("signNote").textContent = "マウスでは書けません。iPad・スマホで書いてください"; return; }
    e.preventDefault(); cv.setPointerCapture(e.pointerId);
    drawing = pt(e); const g = cv.getContext("2d"); g.beginPath(); g.arc(drawing[0], drawing[1], 1.2, 0, 7); g.fill(); signDirty = true;
  });
  cv.addEventListener("pointermove", e => {
    if (!drawing) return; e.preventDefault();
    const p = pt(e), g = cv.getContext("2d");
    g.beginPath(); g.moveTo(drawing[0], drawing[1]); g.lineTo(p[0], p[1]); g.stroke(); drawing = p;
  });
  const end = () => { drawing = null; };
  cv.addEventListener("pointerup", end); cv.addEventListener("pointercancel", end);
  $("signClear").onclick = () => { const g = cv.getContext("2d"); g.save(); g.setTransform(1, 0, 0, 1, 0, 0); g.clearRect(0, 0, cv.width, cv.height); g.restore(); signDirty = false; };
  $("signOk").onclick = () => {
    if (!signDirty) { toast("サインが書かれていません"); return; }
    // 小さく（幅 360px）して持つ（ブラウザの保存の容量を食わないように）
    const w = 360, h = Math.round(w * cv.height / cv.width), o = document.createElement("canvas");
    o.width = w; o.height = h; o.getContext("2d").drawImage(cv, 0, 0, w, h);
    pushUndo();
    const d = ensureDay(); d.サイン = o.toDataURL("image/png"); d.サイン日時 = now(); d.更新 = now();
    save(); closeDlg($("dSign")); refreshSheet(); toast("サインを入れました");
  };
})();

/* ================= 小窓：⠿（見出し）で動かす・外を押しても閉じない・後ろをスクロールさせない ================= */
function openDlg(d) {
  d.classList.remove("moved"); d.style.left = d.style.top = "";
  const a = document.activeElement; if (a && a !== document.body) a.blur();
  d.showModal(); document.body.classList.add("modal");
}
function closeDlg(d) { d.close(); }
document.querySelectorAll("dialog").forEach(d => {
  d.addEventListener("close", () => { if (!document.querySelector("dialog[open]")) document.body.classList.remove("modal"); });
  d.addEventListener("click", e => { if (e.target.closest("[data-close]")) closeDlg(d); });
  const hd = d.querySelector(".dhead");
  let drag = null;
  hd.addEventListener("pointerdown", e => {
    if (e.target.closest("button")) return;
    const r = d.getBoundingClientRect();
    drag = { dx: e.clientX - r.left, dy: e.clientY - r.top };
    d.classList.add("moved"); d.style.left = r.left + "px"; d.style.top = r.top + "px";
    hd.setPointerCapture(e.pointerId); e.preventDefault();
  });
  hd.addEventListener("pointermove", e => {
    if (!drag) return;
    const w = d.offsetWidth, h = d.offsetHeight;
    d.style.left = Math.max(-w + 80, Math.min(innerWidth - 80, e.clientX - drag.dx)) + "px";
    d.style.top = Math.max(0, Math.min(innerHeight - 40, e.clientY - drag.dy)) + "px";
  });
  const up = () => { drag = null; };
  hd.addEventListener("pointerup", up); hd.addEventListener("pointercancel", up);
});
// 小窓を開いている間、スワイプで後ろが動かないように（中のスクロールできる所と書く所だけ通す）
document.addEventListener("touchmove", e => {
  if (!document.body.classList.contains("modal")) return;
  const sc = e.target.closest(".dbody");
  if (e.target.closest("canvas") || e.target.closest(".dhead")) { e.preventDefault(); return; }
  if (!sc || sc.scrollHeight <= sc.clientHeight) e.preventDefault();
}, { passive: false });

/* ================= マスタ（ほかのアプリと同じく会社 OneDrive から読む。書かない） =================
   現場一覧.json（稼働データ。NAS の受注台帳・注文書台帳・工事部会資料から毎朝作る）→ 元請・現場（コード・正式名・略称）
   マスタ.json（工程データ。配置表の名簿）→ 社員・従業員・外注業者 → 作業員名の候補
   稼働データ.json（自社機械稼働表）→ 自社機械（名前・規格）→ 使用機材の候補
   読んだものはこのブラウザに控えて、サインインしていないとき・電波の無いときもその控えで動く */
const MKEY = "jouyou_master_v1";
let MS = (() => { try { return JSON.parse(localStorage.getItem(MKEY)) || null; } catch (e) { return null; } })() || { 現場: [], 社員: [], 従業員: [], 外注業者: [], 機械: [], 読んだ: "" };
const CFG = window.JOUYOU_CONFIG || {};
const SCOPES = ["User.Read", "Files.ReadWrite.All"];   // 日報アプリと同じ（同じアプリの登録で、もう同意してあるもの）。書き込みはしない
const GRAPH = "https://graph.microsoft.com/v1.0";
let msalApp = null;
// 名前を比べる形（空白・全角半角・株式会社／(株) のちがいを無視）
const nk = s => String(s ?? "").normalize("NFKC").replace(/株式会社|\(株\)|有限会社|\(有\)|合同会社|\(同\)/g, "").replace(/[\s　・]/g, "").toLowerCase();
const people = () => [...MS.社員.map(n => ({ n, k: "社員" })), ...MS.従業員.map(n => ({ n, k: "従業員" })), ...MS.外注業者.map(n => ({ n, k: "外注" }))];
const masterName = v => { const k = nk(v); const p = people().find(x => nk(x.n) === k); return p ? p.n : v; };
function siteMatch(name, m) {
  const a = nk(name); if (!a) return false;
  const f = nk(m.正式名), r = nk(m.名前);
  return a === f || a === r || (r.length >= 3 && a.includes(r)) || (f.length >= 6 && (a.includes(f) || f.includes(a)));
}
// 現場一覧の現場を、このブラウザの元請・現場に入れる（日報のある現場は名前で結び付ける。入っている値は変えない）
function syncMasterSites() {
  const lim = addD(ymd(new Date()), -400);
  let changed = false;
  const genOf = name => {
    const k = nk(name) || "（元請なし）";
    let g = DB.元請.find(x => nk(x.名前) === k);
    if (!g) { g = { id: "g_" + k, 名前: name || "（元請なし）" }; DB.元請.push(g); changed = true; }
    return g;
  };
  for (const m of MS.現場) {
    let s = DB.現場.find(x => x.コード && x.コード === m.コード) || DB.現場.find(x => !x.コード && siteMatch(x.名前, m));
    const recent = !m.工期至 || m.工期至 >= lim;
    if (!s && !recent) continue;
    const g = genOf(m.元請);
    if (!s) { s = fixSite({ id: "c" + (m.コード || uid()), 元請id: g.id, 名前: m.正式名 || m.名前 }); DB.現場.push(s); changed = true; }
    if (!s.コード || s.略称 !== m.名前) { s.コード = m.コード || ""; s.略称 = m.名前 || ""; changed = true; }
    if (s.元請id !== g.id && nk((DB.元請.find(x => x.id === s.元請id) || {}).名前) !== nk(g.名前)) { s.元請id = g.id; changed = true; }
  }
  // 現場の無くなった元請は消す（手で足したものも、現場が無ければ一覧に出さない）
  const used = new Set(DB.現場.map(s => s.元請id));
  const before = DB.元請.length; DB.元請 = DB.元請.filter(g => used.has(g.id) || !g.id.startsWith("g_"));
  if (DB.元請.length !== before) changed = true;
  if (changed) save();
}
function applyMasters(m) {
  MS = { 現場: m.現場 || [], 社員: m.社員 || [], 従業員: m.従業員 || [], 外注業者: m.外注業者 || [], 機械: m.機械 || [], 読んだ: m.読んだ || now(), 元: m.元 || "" };
  try { localStorage.setItem(MKEY, JSON.stringify(MS)); } catch (e) {}
  syncMasterSites(); showMstat(); renderAll();
}
function showMstat() {
  const el = $("mstat");
  if (!MS.読んだ) { el.textContent = "マスタ未読"; el.className = "mstat"; return; }
  el.textContent = `マスタ ${MS.読んだ.slice(5)}（現場${MS.現場.length}・人${people().length}・機械${MS.機械.length}）`;
  el.className = "mstat " + (MS.読んだ.slice(0, 10) === ymd(new Date()) ? "ok" : "old");
}
let lastToken = "";
async function token() {
  try { return (lastToken = (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken); }
  catch (e) { if (e instanceof msal.InteractionRequiredAuthError) await msalApp.acquireTokenRedirect({ scopes: SCOPES }); throw e; }
}
async function getJSON(where, name) {
  const F = CFG[where], t = await token();
  const res = await fetch(`${GRAPH}/drives/${F.driveId}/items/${F.itemId}:/${encodeURIComponent(name)}:/content`, { headers: { Authorization: "Bearer " + t }, cache: "no-store" });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(name + " " + res.status);
  return res.json();
}
async function loadMasters(quiet) {
  if (!msalApp || !msalApp.getActiveAccount()) return;
  try {
    const [g, k, w] = await Promise.all([getJSON("kikai", "現場一覧.json"), getJSON("kotei", "マスタ.json"), getJSON("kikai", "稼働データ.json").catch(() => null)]);
    const mk = new Map();
    ((w && w.機械) || []).forEach(x => { if (x && x.名前 && !x.使わない) mk.set(x.名前 + "|" + (x.規格 || ""), { 名前: x.名前, 規格: x.規格 || "" }); });
    applyMasters({ 現場: (g && g.現場) || MS.現場, 社員: (k && k.社員) || MS.社員, 従業員: (k && k.従業員) || MS.従業員, 外注業者: (k && k.外注業者) || MS.外注業者,
      機械: mk.size ? [...mk.values()] : MS.機械, 読んだ: now(), 元: g && g.読んだ ? "現場一覧 " + g.読んだ.slice(0, 16).replace("T", " ") : "" });
    if (!quiet) toast("マスタを読みました");
  } catch (e) {
    console.warn(e);
    toast(/40[13]/.test(e.message) ? "マスタのフォルダを開けません（工程データ・稼働データの共有を北澤さんに頼んでください）" : "マスタを読めませんでした（前に読んだ控えで動きます）：" + e.message);
  }
}
async function startAuth() {
  showMstat();
  if (DEMO) { $("who").textContent = "試し（共有はこのブラウザの中）"; startSync(); return; }
  syncMark("", "このブラウザだけ"); $("sync").title = "サインインすると、ほかの PC・iPad と日報を共有します";
  if (location.protocol === "file:" || !window.msal || !CFG.clientId) return;
  msalApp = new msal.PublicClientApplication({
    auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: new URL("./", location.href).href },
    cache: { cacheLocation: "localStorage" }
  });
  try {
    await msalApp.initialize();
    const r = await msalApp.handleRedirectPromise(); if (r) msalApp.setActiveAccount(r.account);
  } catch (e) { console.warn(e); }
  const acc = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if (!acc) { $("bSignIn").hidden = false; toast("ほかの PC・iPad と日報を共有するには「サインイン」してください"); return; }
  msalApp.setActiveAccount(acc);
  $("who").textContent = acc.name || acc.username;
  renderPickers(); refreshSheet();   // 「サインインすると出ます」を消す
  await loadMasters(true);
  setInterval(() => loadMasters(true), 30 * 60 * 1000);
  startSync();
}
function startSync() { sync(); setInterval(() => syncSoon(0), 60 * 1000); LOCK.site = null; takeLock(false); setInterval(beatLock, 45 * 1000); }
$("bSignIn").onclick = () => msalApp && msalApp.loginRedirect({ scopes: SCOPES });

/* ================= 日報の共有（2026-10-31 ユーザー：Excel から取り込んだものがほかの PC で見られない） =================
   工事日報と同じく、全体工程表の 工程データ フォルダの中の「常用日報」フォルダに、現場ごとに 現場_<現場id>.json を置く。
   { 形式:"常用日報の現場", 版:1, 元請名, 現場:{ …現場（日・消した日・設定更新 も）… } }
   ・開いたとき・1分ごと・画面に戻ったときに読み、このブラウザの分と合わせる（日ごとに 更新 の新しいほう。消した日は 消した時刻 が新しければ消す。
     名簿・機材は足し算、決まり（定時・締め日 など）は 設定更新 の新しいほう）
   ・直したら数秒後に書く。同時にほかの人が書いていたら（eTag がちがう）読み直して合わせてから書き直す。
   ・サインインしていないとき・電波の無いときはこのブラウザの中だけ。つながったら書く。 */
const SDIR = "常用日報";
const DEMO = new URLSearchParams(location.search).has("demo");
const SY = { files: {}, busy: false, timer: null, again: false, err: "", last: "" };   // files: 現場id → { eTag, sent }
const signedIn = () => DEMO || !!(msalApp && msalApp.getActiveAccount());
const enc = p => p.split("/").map(encodeURIComponent).join("/");
// 試し（?demo）：OneDrive の代わりにこのブラウザの別の場所（2台目は localStorage を分けて試す）
const DSTORE = "jouyou_demo_store";
const dAll = () => { try { return JSON.parse(localStorage.getItem(DSTORE)) || {}; } catch (e) { return {}; } };
const store = {
  async list() {
    if (DEMO) return Object.entries(dAll()).map(([name, x]) => ({ name, eTag: x.eTag }));
    const F = CFG.kotei, out = [];
    let url = `${GRAPH}/drives/${F.driveId}/items/${F.itemId}:/${enc(SDIR)}:/children?$top=999&$select=name,eTag,file`;
    while (url) {
      const res = await fetch(url, { headers: { Authorization: "Bearer " + await token() } });
      if (res.status === 404) return out;
      if (!res.ok) throw httpErr(res, "一覧");
      const j = await res.json();
      j.value.forEach(x => { if (x.file) out.push({ name: x.name, eTag: x.eTag }); });
      url = j["@odata.nextLink"] || null;
    }
    return out;
  },
  async get(name) {
    if (DEMO) { const x = dAll()[name]; return x ? { data: JSON.parse(x.data), eTag: x.eTag } : null; }
    const F = CFG.kotei, h = { Authorization: "Bearer " + await token() };
    const res = await fetch(`${GRAPH}/drives/${F.driveId}/items/${F.itemId}:/${enc(SDIR + "/" + name)}`, { headers: h });
    if (res.status === 404) return null;
    if (!res.ok) throw httpErr(res, "読み込み");
    const meta = await res.json();
    const body = await fetch(meta["@microsoft.graph.downloadUrl"], { cache: "no-store" });
    if (!body.ok) throw httpErr(body, "中身の読み込み");
    return { data: await body.json(), eTag: meta.eTag };
  },
  async put(name, text, eTag) {
    if (DEMO) {
      const a = dAll(), cur = a[name];
      if (eTag ? (!cur || cur.eTag !== eTag) : cur) throw Object.assign(new Error("保存 " + (eTag ? 412 : 409)), { status: eTag ? 412 : 409 });
      const t = "d" + Date.now() + Math.random().toString(36).slice(2, 6); a[name] = { data: text, eTag: t };
      localStorage.setItem(DSTORE, JSON.stringify(a)); return { eTag: t };
    }
    const F = CFG.kotei, base = `${GRAPH}/drives/${F.driveId}/items/${F.itemId}`;
    const url = `${base}:/${enc(SDIR + "/" + name)}:/content` + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
    const send = async () => fetch(url, { method: "PUT", headers: { Authorization: "Bearer " + await token(), "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: text });
    let res = await send();
    if (res.status === 404 && !eTag) {
      // 常用日報 フォルダがまだ無い → 作ってからもう一度
      await fetch(base + "/children", { method: "POST", headers: { Authorization: "Bearer " + await token(), "Content-Type": "application/json" }, body: JSON.stringify({ name: SDIR, folder: {}, "@microsoft.graph.conflictBehavior": "replace" }) });
      res = await send();
    }
    if (!res.ok) throw httpErr(res, "保存");
    return { eTag: (await res.json()).eTag };
  },
  async del(name) {
    if (DEMO) { const a = dAll(); delete a[name]; localStorage.setItem(DSTORE, JSON.stringify(a)); return; }
    const F = CFG.kotei;
    await fetch(`${GRAPH}/drives/${F.driveId}/items/${F.itemId}:/${enc(SDIR + "/" + name)}`, { method: "DELETE", headers: { Authorization: "Bearer " + await token() } });
  }
};
function httpErr(res, what) { return Object.assign(new Error(what + " " + res.status), { status: res.status }); }
const fileOf = id => "現場_" + id + ".json";
const docOf = s => JSON.stringify({ 形式: "常用日報の現場", 版: 1, 元請名: (DB.元請.find(g => g.id === s.元請id) || {}).名前 || "", 現場: s });
// 共有に出す現場：日報か、消した日か、直した決まりがあるもの（マスタから作っただけの空の現場は出さない）
const shareable = s => Object.keys(s.日).length || Object.keys(s.消した日).length || s.設定更新;
// このブラウザの現場（a）と共有の現場（b）を合わせる
function mergeSite(a, b) {
  fixSite(b);
  const out = b.設定更新 > a.設定更新 ? { ...a, ...b } : { ...b, ...a };
  const bNew = (b.設定更新 || "") > (a.設定更新 || "");
  out.作業員 = a.作業員.map(w => ({ ...w })); b.作業員.forEach(w => { const x = out.作業員.find(y => nk(y.名前) === nk(w.名前)); if (!x) out.作業員.push(w); else if (bNew) Object.assign(x, w); else { if (!x.会社 && w.会社) x.会社 = w.会社; if (!x.職種 && w.職種) x.職種 = w.職種; } });
  out.機材 = [...a.機材]; b.機材.forEach(m => { if (!out.機材.some(x => x.名前 === m.名前 && (x.規格 || "") === (m.規格 || ""))) out.機材.push(m); });
  out.足さない = [...new Set([...(a.足さない || []), ...(b.足さない || [])])];
  out.別 = [...(a.別 || [])]; (b.別 || []).forEach(p => { if (!out.別.some(q => q.join("\u0001") === p.join("\u0001"))) out.別.push(p); });
  out.消した日 = { ...a.消した日 }; for (const [d, t] of Object.entries(b.消した日)) if (!out.消した日[d] || t > out.消した日[d]) out.消した日[d] = t;
  out.日 = {};
  for (const d of new Set([...Object.keys(a.日), ...Object.keys(b.日)])) {
    const x = a.日[d], y = b.日[d], v = !x ? y : !y ? x : ((y.更新 || "") > (x.更新 || "") ? y : x);
    if (v && !(out.消した日[d] && out.消した日[d] >= (v.更新 || ""))) out.日[d] = v;
  }
  return out;
}
function syncMark(cls, text) { const el = $("sync"); el.className = "sync " + cls; el.textContent = text; }
function syncSoon(ms = 2500) { if (!signedIn()) return; clearTimeout(SY.timer); SY.timer = setTimeout(sync, ms); }
const busyEditing = () => (grids.up && grids.up.editing()) || (grids.dn && grids.dn.editing()) || (grids.car && grids.car.editing()) || ["hTan", "hTmp", "hBikou"].includes((document.activeElement || {}).id);
async function sync() {
  if (!signedIn()) return;
  if (SY.busy) { SY.again = true; return; }
  SY.busy = true; syncMark("busy", "共有 …");
  let changed = false;
  try {
    // 1) 読む：ほかの端末が書いたもの（eTag が変わったもの）を合わせる
    let commonIt = null;
    for (const it of await store.list()) {
      if (it.name === "共通.json") { commonIt = it; continue; }
      const m = it.name.match(/^現場_(.+)\.json$/); if (!m) continue;
      const id = m[1], f = SY.files[id];
      if (f && f.eTag === it.eTag) continue;
      const r = await store.get(it.name); if (!r || !r.data || !r.data.現場) continue;
      const rs = r.data.現場;
      let g = DB.元請.find(x => nk(x.名前) === nk(r.data.元請名));
      if (!g) { g = { id: "g_" + (nk(r.data.元請名) || "（元請なし）"), 名前: r.data.元請名 || "（元請なし）" }; DB.元請.push(g); }
      const i = DB.現場.findIndex(x => x.id === id);
      const merged = i < 0 ? fixSite(rs) : mergeSite(DB.現場[i], rs);
      if (!DB.元請.some(x => x.id === merged.元請id)) merged.元請id = g.id;   // ほかの端末の元請 id はこの端末に無いことがある → 名前で合わせる
      aliasKou(merged);
      if (i < 0) DB.現場.push(merged); else DB.現場[i] = merged;
      SY.files[id] = { eTag: r.eTag, sent: JSON.stringify(r.data) };
      changed = true;
    }
    // 共通（工種・職種単価・天気）：新しいほうを使う
    if (commonIt && (!SY.common || SY.common.eTag !== commonIt.eTag)) {
      const r = await store.get("共通.json");
      if (r && r.data) {
        if ((r.data.更新 || "") > (DB.共通更新 || "")) { DB.職種 = r.data.職種 || DB.職種; DB.職種単価 = r.data.職種単価 || {}; DB.天気 = r.data.天気 || DB.天気; DB.会社一覧 = r.data.会社一覧 || DB.会社一覧; DB.職種別名 = r.data.職種別名 || DB.職種別名; DB.単価たたき台 = r.data.単価たたき台 || DB.単価たたき台; DB.共通更新 = r.data.更新; kouFix(); changed = true; }
        SY.common = { eTag: r.eTag, sent: JSON.stringify(r.data) };
      }
    }
    if (DB.共通更新) {
      const text = JSON.stringify({ 形式: "常用日報の共通", 版: 1, 職種: DB.職種, 職種単価: DB.職種単価, 天気: DB.天気, 会社一覧: DB.会社一覧, 職種別名: DB.職種別名, 単価たたき台: DB.単価たたき台, 更新: DB.共通更新 });
      if (!SY.common || SY.common.sent !== text) {
        try { const r = await store.put("共通.json", text, SY.common && SY.common.eTag); SY.common = { eTag: r.eTag, sent: text }; }
        catch (e) { if (e.status !== 412 && e.status !== 409) throw e; SY.common = null; SY.again = true; }   // ほかの人が先に書いた → 次の回に読み直して合わせる
      }
    }
    // 2) 書く：このブラウザで直した現場
    for (const s0 of DB.現場) {
      if (!shareable(s0)) continue;
      for (let n = 0; n < 3; n++) {
        const s = DB.現場.find(x => x.id === s0.id), text = docOf(s), f = SY.files[s.id];
        if (f && f.sent === text) break;
        try { const r = await store.put(fileOf(s.id), text, f && f.eTag); SY.files[s.id] = { eTag: r.eTag, sent: text }; break; }
        catch (e) {
          if (e.status !== 412 && e.status !== 409) throw e;
          // ほかの人が先に書いた → 読み直して合わせてから書き直す
          const r = await store.get(fileOf(s.id));
          if (r && r.data && r.data.現場) { const i = DB.現場.findIndex(x => x.id === s.id); DB.現場[i] = mergeSite(DB.現場[i], r.data.現場); aliasKou(DB.現場[i]); SY.files[s.id] = { eTag: r.eTag, sent: JSON.stringify(r.data) }; changed = true; }
          else delete SY.files[s.id];
        }
      }
    }
    SY.err = ""; SY.last = now();
    syncMark("ok", "共有 ✓ " + SY.last.slice(11, 16));
  } catch (e) {
    console.warn(e); SY.err = e.message;
    syncMark("err", "共有できません");
    $("sync").title = (e.status === 403 || e.status === 404 ? "工程データ フォルダに書けません。北澤さんに共有（編集可）を頼んでください。" : "電波・サインインを確かめてください。") + "（" + e.message + "）このブラウザの中には残っています。";
  } finally {
    SY.busy = false;
    if (changed) {
      try { localStorage.setItem(KEY, JSON.stringify(DB)); } catch (e) {}
      if (!busyEditing()) { renderPickers(); refreshSheet(); if (UI.tab === "sum") renderSum(); if (UI.tab === "master") renderMaster(); }
    }
    if (SY.again) { SY.again = false; syncSoon(500); }
  }
}
/* ================= 使う人の名前・使用中の印・見るだけ（2026-10-31 ユーザー：ログイン時に名前を入れる。
   同時に使おうとしたら、だれが使っているか出す。見るだけ（閲覧のみ）はできる） =================
   ・名前は端末ごとに1回（このブラウザに覚える）。右上の 👤 で直せる。
   ・現場を開くと、共有の「常用日報」フォルダに 使用中_<現場id>.json（名前・端末・始めた時刻・いつ）を置き、45秒ごとに いつ を書き直す。
     3分 いつ が書き直されていなければ、使い終わった（閉じた・電波が切れた）とみなす。
   ・ほかの人の印があれば「○○さんが 12:50 から使っています」と出し、見るだけで開く（表・欄は直せない）。
     「交代して編集」で印を取れる（相手は次の書き直しのときに見るだけになる）。自分から「見るだけにする」もできる。
   ・サインインしていないとき（このブラウザだけ）は印を置かない。 */
const MEKEY = "jouyou_me";
let ME = (() => { try { return JSON.parse(localStorage.getItem(MEKEY)) || null; } catch (e) { return null; } })() || { 名前: "", 端末: uid() };
const saveME = () => { try { localStorage.setItem(MEKEY, JSON.stringify(ME)); } catch (e) {} };
saveME();
const STALE = 3 * 60 * 1000;
const LOCK = { site: null, mine: false, eTag: null, busy: false };
let RO = null;   // null＝直せる／{ by:{名前,開始,いつ} }＝ほかの人が使用中／{ self:true }＝自分で見るだけにした／{ gone:true, by }＝使っていた人が終わった
const lockName = id => "使用中_" + id + ".json";
// 同じ人か：サインインしたアカウントで見る（PC と iPad の両方で同じ人が開いても、別の人とみなさない。2026-10-31 ユーザー）。
// 前の版の印（人 が無い）は名前で見る
const myKey = () => { const a = msalApp && msalApp.getActiveAccount(); return a ? String(a.username || a.homeAccountId || "").toLowerCase() : "名前:" + nk(ME.名前); };
const samePerson = by => !!by && (by.端末 === ME.端末 || (by.人 ? by.人 === myKey() : nk(by.名前) === nk(ME.名前)));
const hm = t => { const d = new Date(t); return pad(d.getHours()) + ":" + pad(d.getMinutes()); };
function canEdit() {
  if (!RO) return true;
  toast(RO.by && !RO.gone ? `${RO.by.名前}さんが ${hm(RO.by.開始)} から使っています。見るだけ（閲覧のみ）で開いています` : "見るだけ（閲覧のみ）で開いています。上の「編集する」で直せます");
  return false;
}
function setRO(v) {
  RO = v;
  document.body.classList.toggle("ro", !!RO);
  ["hWx", "hTmp", "hTan", "hBikou"].forEach(id => { if ($(id)) $(id).disabled = !!RO; });
  const bar = $("roBar");
  if (!RO) { bar.hidden = true; }
  else {
    bar.hidden = false;
    bar.innerHTML = RO.self ? `👁 見るだけ（閲覧のみ）で開いています。<button class="small" id="roEdit">編集する</button>`
      : RO.gone ? `👁 見るだけ：${esc(RO.by.名前)}さんは使い終わりました。<button class="small" id="roEdit">編集する</button>`
      : `👁 見るだけ（閲覧のみ）：<b>${esc(RO.by.名前)}</b>さんが ${hm(RO.by.開始)} から使っています（${hm(RO.by.いつ)} 確認）。直すときは使い終わるのを待つか <button class="small sub" id="roTake">交代して編集</button>`;
    const e1 = $("roEdit"), e2 = $("roTake");
    if (e1) e1.onclick = () => takeLock(false, true);
    if (e2) e2.onclick = () => { if (confirm(`${RO.by.名前}さんが使っています。交代して編集しますか？\n（${RO.by.名前}さんの画面は見るだけになります。同じ日を同時に直していると、新しく直したほうが残ります）`)) takeLock(true); };
  }
  $("bRo").textContent = RO ? "👁 見るだけ" : "✏ 編集中";
  $("bRo").title = RO ? "見るだけで開いています（押すと編集する）" : "押すと見るだけにします（ほかの人が直せるように）";
  $("bRo").className = "sub rotg" + (RO ? " on" : "");
}
async function takeLock(force, fromSelf) {
  const id = UI.site;
  if (LOCK.site && LOCK.site !== id && LOCK.mine) releaseLock(LOCK.site);
  LOCK.site = id; LOCK.mine = false; LOCK.eTag = null;
  if (!id || !signedIn()) { setRO(null); return; }
  if (!ME.名前) { askName(); return; }
  if (LOCK.busy) return; LOCK.busy = true;
  try {
    for (let n = 0; n < 2; n++) {
      const r = await store.get(lockName(id)), by = r && r.data;
      const other = by && !samePerson(by) && Date.now() - by.いつ < STALE;
      if (other && !force) { setRO({ by }); if (!fromSelf) toast(`${by.名前}さんが使っています。見るだけ（閲覧のみ）で開きました`, 5000); return; }
      const doc = { 名前: ME.名前, 人: myKey(), 端末: ME.端末, 開始: samePerson(by) ? by.開始 : Date.now(), いつ: Date.now() };
      try { const w = await store.put(lockName(id), JSON.stringify(doc), r && r.eTag); if (UI.site !== id) return; LOCK.mine = true; LOCK.eTag = w.eTag; setRO(null); return; }
      catch (e) { if (e.status !== 412 && e.status !== 409) throw e; }   // ちょうどほかの人が印を置いた → 読み直す
    }
  } catch (e) { console.warn(e); if (UI.site === id) setRO(null); }   // 印を置けない（電波など）ときは直せるままにする
  finally { LOCK.busy = false; }
}
async function beatLock() {
  const id = LOCK.site; if (!id || !signedIn() || LOCK.busy) return;
  try {
    const r = await store.get(lockName(id)), by = r && r.data;
    if (LOCK.mine) {
      if (by && !samePerson(by)) { LOCK.mine = false; setRO({ by }); toast(`${by.名前}さんが編集を始めたので、見るだけにしました`, 6000); return; }
      const w = await store.put(lockName(id), JSON.stringify({ 名前: ME.名前, 人: myKey(), 端末: ME.端末, 開始: by ? by.開始 : Date.now(), いつ: Date.now() }), r && r.eTag);
      LOCK.eTag = w.eTag;
    } else if (RO && RO.by && !RO.gone) {
      if (!by || Date.now() - by.いつ >= STALE) setRO({ gone: true, by: RO.by });
      else if (samePerson(by)) setRO({ gone: true, by: RO.by });
      else if (by.端末 !== RO.by.端末 || by.いつ !== RO.by.いつ) setRO({ by });
    }
  } catch (e) { console.warn(e); }
}
function releaseLock(id) {
  if (!signedIn()) return;
  if (DEMO) { store.del(lockName(id)); return; }
  // 閉じるときも消せるよう、覚えている token で keepalive の DELETE
  const F = CFG.kotei;
  if (lastToken) fetch(`${GRAPH}/drives/${F.driveId}/items/${F.itemId}:/${enc(SDIR + "/" + lockName(id))}`, { method: "DELETE", headers: { Authorization: "Bearer " + lastToken }, keepalive: true }).catch(() => {});
}
window.addEventListener("pagehide", () => { if (LOCK.mine && LOCK.site) releaseLock(LOCK.site); });
$("bRo").onclick = () => {
  if (RO) { takeLock(false, true); return; }
  if (LOCK.mine && LOCK.site) releaseLock(LOCK.site);
  LOCK.mine = false; setRO({ self: true });
};
// 名前を入れる小窓（名簿の社員から選ぶか打つ）
function askName() {
  const names = [...new Set([...MS.社員, ...MS.従業員])];
  $("nmList").innerHTML = names.map(n => `<button type="button" class="sub" data-n="${esc(n)}">${esc(n)}</button>`).join("") || `<span class="note">（名簿はサインインすると出ます）</span>`;
  const acc = msalApp && msalApp.getActiveAccount();
  $("nmIn").value = ME.名前 || (acc ? masterName(acc.name || "") : "");
  $("nmX").hidden = !ME.名前;
  openDlg($("dName")); setTimeout(() => $("nmIn").focus(), 50);
}
$("nmList").addEventListener("click", e => { const b = e.target.closest("[data-n]"); if (b) $("nmIn").value = b.dataset.n; });
$("nmOk").onclick = () => {
  const v = masterName($("nmIn").value.trim());
  if (!v) { toast("名前を入れてください"); return; }
  ME.名前 = v; saveME(); showMe(); closeDlg($("dName"));
  LOCK.site = null; takeLock(false);
};
$("nmIn").addEventListener("keydown", e => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); $("nmOk").click(); } });
$("dName").addEventListener("cancel", e => { if (!ME.名前) e.preventDefault(); });   // 名前が無いうちは Esc で閉じない
function showMe() { $("bMe").textContent = "👤 " + (ME.名前 || "名前を入れる"); }
$("bMe").onclick = askName;

async function removeShared(id) { if (signedIn() && SY.files[id]) { try { await store.del(fileOf(id)); } catch (e) { console.warn(e); } delete SY.files[id]; } }
document.addEventListener("visibilitychange", () => { if (!document.hidden) syncSoon(300); });

/* ================= 上の選ぶ欄（元請 → 現場） ================= */
const siteLabel = s => s.略称 ? s.略称 + (s.コード ? "（" + s.コード + "）" : "") : s.名前;
function renderPickers() {
  // 現場のある元請だけ・名前の順。現場はコードの新しい順（日報のある現場が先）
  const has = new Set(DB.現場.map(s => s.元請id));
  const gs = DB.元請.filter(g => has.has(g.id)).sort((a, b) => nk(a.名前).localeCompare(nk(b.名前), "ja"));
  if (!gen() && gs.length) UI.gen = (DB.現場.find(s => Object.keys(s.日).length) || {}).元請id || gs[0].id;
  $("inGen").innerHTML = gs.length ? gs.map(g => `<option value="${g.id}"${g.id === UI.gen ? " selected" : ""}>${esc(g.名前)}</option>`).join("") : `<option value="">${signedIn() ? "（元請業者なし）" : "（サインインすると出ます）"}</option>`;
  const nDays = s => Object.keys(s.日).length;
  const ss = DB.現場.filter(s => s.元請id === UI.gen).sort((a, b) => (!!a.完了 - !!b.完了) || (nDays(b) > 0) - (nDays(a) > 0) || String(b.コード || "").localeCompare(String(a.コード || "")));
  if (!ss.some(s => s.id === UI.site)) UI.site = ss.length ? ss[0].id : null;
  $("inSite").innerHTML = ss.length ? ss.map(s => `<option value="${s.id}"${s.id === UI.site ? " selected" : ""} title="${esc(s.名前)}${s.工期開始 || s.工期終了 ? "　工期 " + esc((s.工期開始 || "") + "〜" + (s.工期終了 || "")) : ""}">${esc(siteLabel(s))}${s.完了 ? "（完了）" : ""}</option>`).join("") : `<option value="">${signedIn() ? "（現場なし）" : "（サインインすると出ます）"}</option>`;
  saveUI();
}
$("inGen").onchange = () => { askEmptyPages(); UI.pg = 0; UI.gen = $("inGen").value; UI.site = null; UNDO = []; REDO = []; renderAll(); };
function priceRemind() {
  const s = site(); if (!s || !Object.keys(s.日).length) return;
  UI.pr ||= {}; const k = s.id + ":" + ((s.単価 || []).length ? (s.単価確認 === false ? "未見直し" : "ok") : "なし");
  if (k.endsWith("ok") || UI.pr[k]) return;
  UI.pr[k] = 1; saveUI();
  toast((s.単価 || []).length ? "この現場の単価はたたき台のまま（見直し前）です。現場・名簿 → 単価マスタで見直してください" : "この現場の単価がまだありません。現場・名簿 → 単価マスタで たたき台 を写して見直してください", 6000);
}
$("inSite").onchange = () => { askEmptyPages(); UI.pg = 0; UI.site = $("inSite").value; UNDO = []; REDO = []; renderAll(); priceRemind(); };
/* ---------- 期間を選ぶ（2026-10-31 ユーザー：始まり・終わりの二つのカレンダー。終わりは始まりより前にできない） ---------- */
let RG = null;   // { a, b, ya, yb, done }
const fmtD = d => d ? `${d.slice(0, 4)}/${+d.slice(5, 7)}/${+d.slice(8)}（${WD[wd(d)]}）` : "";
function openRange(a, b, done) {
  RG = { a, b, ya: a.slice(0, 7), yb: b.slice(0, 7), done };
  renderRange(); openDlg($("dRange"));
}
function calMonth(which) {
  const s = site(), ym = which === "a" ? RG.ya : RG.yb, [y, m] = ym.split("-").map(Number);
  const first = new Date(y, m - 1, 1), n = new Date(y, m, 0).getDate(), today = ymd(new Date());
  const h = [`<div class="calnav"><button type="button" class="sub" data-mv="${which}:-1">◀</button><b>${y}年${m}月</b><button type="button" class="sub" data-mv="${which}:1">▶</button></div><div class="calgrid">`];
  ["日", "月", "火", "水", "木", "金", "土"].forEach((x, i) => h.push(`<b class="${i === 0 ? "sun" : i === 6 ? "sat" : ""}">${x}</b>`));
  for (let i = 0; i < first.getDay(); i++) h.push("<span></span>");
  for (let d = 1; d <= n; d++) {
    const k = `${y}-${pad(m)}-${pad(d)}`, w = wd(k), has = s && hasData(s.日[k]);
    const off = which === "b" && k < RG.a;   // 終わりは始まりより前にしない
    const on = which === "a" ? k === RG.a : k === RG.b, inr = k >= RG.a && k <= RG.b;
    h.push(`<button type="button" data-${which}="${k}"${off ? " disabled" : ""} class="${w === 0 ? "sun " : w === 6 ? "sat " : ""}${has ? "has " : ""}${on ? "on " : ""}${inr && !on ? "inper " : ""}${k === today ? "today" : ""}"><span class="dn">${d}</span>${has ? '<span class="mk">●</span>' : ""}</button>`);
  }
  h.push("</div>");
  return h.join("");
}
function renderRange() {
  $("rgA").innerHTML = calMonth("a"); $("rgB").innerHTML = calMonth("b");
  const n = RG.a && RG.b ? datesIn(RG.a, RG.b).length : 0, s = site(), nd = s ? datesIn(RG.a, RG.b).filter(x => hasData(s.日[x])).length : 0;
  $("rgInfo").textContent = `${fmtD(RG.a)} 〜 ${fmtD(RG.b)}　${n}日間（記入済み ${nd}日）`;
}
$("dRange").addEventListener("click", e => {
  const mv = e.target.closest("[data-mv]");
  if (mv) { const [w, d] = mv.dataset.mv.split(":"), key = w === "a" ? "ya" : "yb"; const [y, m] = RG[key].split("-").map(Number); RG[key] = ymd(new Date(y, m - 1 + +d, 1)).slice(0, 7); renderRange(); return; }
  const a = e.target.closest("[data-a]"), b = e.target.closest("[data-b]");
  if (a) { RG.a = a.dataset.a; if (RG.b < RG.a) { RG.b = RG.a; RG.yb = RG.a.slice(0, 7); } renderRange(); }
  if (b && !b.disabled) { RG.b = b.dataset.b; renderRange(); }
});
$("rgPer").onclick = () => { const [a, b] = period(RG.a); RG.a = a; RG.b = b; RG.ya = a.slice(0, 7); RG.yb = b.slice(0, 7); renderRange(); };
$("rgPrevPer").onclick = () => { const [a, b] = period(addD(RG.a, -1)); RG.a = a; RG.b = b; RG.ya = a.slice(0, 7); RG.yb = b.slice(0, 7); renderRange(); };
$("rgNextPer").onclick = () => { const [a, b] = period(addD(RG.b, 1)); RG.a = a; RG.b = b; RG.ya = a.slice(0, 7); RG.yb = b.slice(0, 7); renderRange(); };
$("rgMonth").onclick = () => { const [y, m] = RG.a.split("-").map(Number); RG.a = ymd(new Date(y, m - 1, 1)); RG.b = ymd(new Date(y, m, 0)); RG.ya = RG.a.slice(0, 7); RG.yb = RG.b.slice(0, 7); renderRange(); };
$("rgOk").onclick = () => { closeDlg($("dRange")); RG.done(RG.a, RG.b); };
function showSumRange() { $("sumRange").textContent = `${fmtD($("sumFrom").value)} 〜 ${fmtD($("sumTo").value)}`; }
$("sumRange").onclick = () => openRange($("sumFrom").value, $("sumTo").value, (a, b) => { $("sumFrom").value = a; $("sumTo").value = b; renderSum(); });
$("outRange").onclick = () => openRange($("outFrom").value, $("outTo").value, (a, b) => { $("outFrom").value = a; $("outTo").value = b; outInfo(); });

/* ---------- 日を選ぶカレンダー（2026-10-31 ユーザー：上の日付を押すと、記入済みの日に印が付いたカレンダーから選ぶ） ---------- */
function showDayBtn() { const w = wd(UI.day); $("inDay").textContent = `${UI.day.slice(0, 4)}/${+UI.day.slice(5, 7)}/${+UI.day.slice(8)}（${WD[w]}）`; $("inDay").className = "sub daybtn" + (w === 0 ? " sun" : w === 6 ? " sat" : ""); }
let calYM = "";
function openCal() {
  calYM = UI.day.slice(0, 7);
  renderCal(); openDlg($("dCal"));
}
function renderCal() {
  const s = site(), [y, m] = calYM.split("-").map(Number);
  const first = new Date(y, m - 1, 1), n = new Date(y, m, 0).getDate(), today = ymd(new Date());
  $("calTtl").textContent = `${y}年${m}月`;
  const [pa, pb] = period(UI.day);
  const h = ["日", "月", "火", "水", "木", "金", "土"].map((x, i) => `<b class="${i === 0 ? "sun" : i === 6 ? "sat" : ""}">${x}</b>`);
  for (let i = 0; i < first.getDay(); i++) h.push("<span></span>");
  let cnt = 0;
  for (let d = 1; d <= n; d++) {
    const k = `${y}-${pad(m)}-${pad(d)}`, w = wd(k), dd = s && s.日[k], has = hasData(dd);
    if (has) cnt++;
    const ppl = has ? dd.下.filter(r => r && r[0]).length : 0;
    h.push(`<button type="button" data-d="${k}" class="${w === 0 ? "sun " : w === 6 ? "sat " : ""}${has ? "has " : ""}${k === UI.day ? "on " : ""}${k === today ? "today " : ""}${k >= pa && k <= pb ? "inper" : ""}" title="${jpDate(k)}${has ? "　記入済み" + (ppl ? " " + ppl + "名" : "") + (dd.サイン ? "・サインあり" : "") : ""}">` +
      `<span class="dn">${d}</span>${has ? `<span class="mk">${dd.サイン ? "✍" : "●"}</span>` : ""}${ppl ? `<span class="pp">${ppl}名</span>` : ""}</button>`);
  }
  $("calGrid").innerHTML = h.join("");
  $("calInfo").textContent = s ? `${siteLabel(s)}　この月の記入済み ${cnt} 日` : "現場を選んでいません";
}
$("inDay").onclick = openCal;
$("calGrid").addEventListener("click", e => { const b = e.target.closest("[data-d]"); if (!b) return; closeDlg($("dCal")); setDay(b.dataset.d); });
$("calPrev").onclick = () => { const [y, m] = calYM.split("-").map(Number); const d = new Date(y, m - 2, 1); calYM = ymd(d).slice(0, 7); renderCal(); };
$("calNext").onclick = () => { const [y, m] = calYM.split("-").map(Number); const d = new Date(y, m, 1); calYM = ymd(d).slice(0, 7); renderCal(); };
$("calToday").onclick = () => { closeDlg($("dCal")); setDay(ymd(new Date())); };
// 記入済みの日へ飛ぶ（前・次）
const filledDays = () => { const s = site(); return s ? Object.keys(s.日).filter(k => hasData(s.日[k])).sort() : []; };
$("calPrevHas").onclick = () => { const k = filledDays().filter(x => x < UI.day).pop(); if (!k) { toast("前に記入済みの日はありません"); return; } closeDlg($("dCal")); setDay(k); };
$("calNextHas").onclick = () => { const k = filledDays().find(x => x > UI.day); if (!k) { toast("後に記入済みの日はありません"); return; } closeDlg($("dCal")); setDay(k); };
$("bPrevD").onclick = () => moveDay(-1);
$("bNextD").onclick = () => moveDay(1);
$("bToday").onclick = () => setDay(ymd(new Date()));
document.querySelectorAll("#views button").forEach(b => b.onclick = () => showTab(b.dataset.tab));
function showTab(t) {
  UI.tab = t; saveUI();
  document.querySelectorAll("#views button").forEach(b => b.classList.toggle("on", b.dataset.tab === t));
  $("pSheet").hidden = t !== "sheet"; $("pSum").hidden = t !== "sum"; $("pMaster").hidden = t !== "master";
  if (t === "sum") { const [a, b] = period(UI.day); $("sumFrom").value = a; $("sumTo").value = b; renderSum(); }
  if (t === "master") renderMaster();
  if (t === "sheet") refreshSheet();
}
function renderAll() {
  renderPickers();
  if (UI.site !== LOCK.site) takeLock(false);
  showDayBtn();
  refreshSheet();
  if (UI.tab === "sum") renderSum();
  if (UI.tab === "master") renderMaster();
}

/* ================= 集計 ================= */
function summarize(s, a, b) {
  const dates = datesIn(a, b).filter(x => hasData(s.日[x]));
  const shoku = [], workers = new Map(), kizai = new Map(), hin = new Map(), comp = new Map();
  const money = { 労務: new Map(), 残業: new Map(), 機械: new Map(), 材料: new Map(), 通勤費: new Map() };
  const byShoku = new Map();
  const order = k => { const base = String(k).replace(/（夜）$/, ""), i = DB.職種.indexOf(base); return (i < 0 ? 99 : i) * 2 + (k !== base ? 1 : 0); };
  const car = [];   // 高速・燃料の行：{ date, r, A：carAmount, who：[会社, 名前, 車両名, 車番] }
  const otRows = [];   // 残業のあった行：{ date, 名前, 会社, 工種, 昼夜, 始業, 終業, 早出, 残業, 計, 金額 }
  const perDay = dates.map(x => {
    const d = s.日[x], c = new Map(); let ot = 0, n = 0, nDay = 0, nNight = 0;
    d.下.forEach(r => {
      if (!r) return;
      if (r[0]) {
        const sh = shiftOf(r, s), base = r[3] || "（工種なし）", k = base + (s.勤務 !== "昼" && sh === "夜" ? "（夜）" : ""); c.set(k, (c.get(k) || 0) + 1); n++;
        if (sh === "夜") nNight++; else nDay++;
        if (!shoku.includes(k)) shoku.push(k);
        const o = Number(r[4]) || 0; ot += o;
        const w = workers.get(nk(r[0])) || { 名前: masterName(r[0]), 職種: new Map(), 会社: new Map(), 日: {}, 日工種: {}, 日数: 0, 残業: 0, 金額: 0 };
        w.職種.set(k, (w.職種.get(k) || 0) + 1); if (!(x in w.日)) w.日数++; w.日[x] = (w.日[x] || 0) + o; w.残業 += o; workers.set(nk(r[0]), w);
        (w.日工種[x] ||= []).push(k);
        // 職種別（工種ごとに だれが何日）
        // 職種別は工種ごと（昼・夜は列で分ける。2026-10-31 ユーザー：1人に 昼・夜・残 の3列）
        const zn = sh === "夜" ? "夜" : "昼";
        const SK = byShoku.get(base) || { 日: {}, 昼: {}, 夜: {}, 残: {}, 昼計: 0, 夜計: 0, 計: 0, 残業: 0, 金額: 0, 人: new Map() }; SK.日[x] = (SK.日[x] || 0) + 1; SK[zn][x] = (SK[zn][x] || 0) + 1; SK[zn + "計"]++; SK.計++; SK.残業 += o; if (o) SK.残[x] = (SK.残[x] || 0) + o;
        const sp2 = SK.人.get(nk(r[0])) || { 名前: masterName(r[0]), 日: {}, 昼: {}, 夜: {}, 残: {}, 昼数: 0, 夜数: 0, 日数: 0, 残業: 0 }; if (!(x in sp2.日)) sp2.日数++; sp2.日[x] = (sp2.日[x] || 0) + o; sp2[zn][x] = (sp2[zn][x] || 0) + 1; sp2[zn + "数"]++; if (o) sp2.残[x] = (sp2.残[x] || 0) + o; sp2.残業 += o; SK.人.set(nk(r[0]), sp2); byShoku.set(base, SK);
        const cp = compOfRow(r, s); w.会社.set(cp, (w.会社.get(cp) || 0) + 1);
        const C = comp.get(cp) || { 日: {}, 計: 0, 残業: 0, 金額: 0 }; C.日[x] = (C.日[x] || 0) + 1; C.計++; C.残業 += o; comp.set(cp, C);
        // 金額（労務・残業）
        const pl = laborPrice(s, base === "（工種なし）" ? "" : base, sh, x, "日"), L = money.労務.get(k) || { 数量: 0, 金額: 0, 単価: new Set(), なし: 0 };
        L.数量++; if (pl) { L.金額 += +pl.単価; L.単価.add(+pl.単価); C.金額 += +pl.単価; w.金額 += +pl.単価; SK.金額 += +pl.単価; } else L.なし++; money.労務.set(k, L);
        if (o) {
          const po = laborPrice(s, base, sh, x, "残業"), O = money.残業.get(k) || { 数量: 0, 金額: 0, 単価: new Set(), なし: 0 };
          O.数量 += o; if (po) { O.金額 += o * po.単価; O.単価.add(+po.単価); C.金額 += o * po.単価; w.金額 += o * po.単価; SK.金額 += o * po.単価; } else O.なし++; money.残業.set(k, O);
          // 残業の内訳（早出・残業）。時刻から出せないとき（手で入れた残業だけ）は全部を残業に
          const pt = otParts(r[1], r[2], s, sh), eh = pt ? Math.min(o, pt.早出 / 60) : 0;
          otRows.push({ date: x, 名前: masterName(r[0]), key: nk(r[0]), 会社: cp, 工種: base, 昼夜: sh, 始業: r[1], 終業: r[2], 早出: eh, 残業: o - eh, 計: o, 金額: po ? o * po.単価 : null });
        }
      }
      if (r[5]) {
        const sp = splitOwner(r[5]), own = r[13] || sp.所有者 || ownerOf(sp.名前, r[6], s) || "（所有者未定）"; const k = sp.名前 + (r[6] ? "（" + r[6] + "）" : "") + "\t" + own; const m = kizai.get(k) || { 日: {}, 計: 0 }; const q = Number(r[7]) || 0; m.日[x] = (m.日[x] || 0) + q; m.計 += q; kizai.set(k, m);
        const qq = q || 1, pm = itemPrice(s, ["機械", "通勤費"], sp.名前, r[6] || "", own === "（所有者未定）" ? "" : own, x);
        const MM = pm && pm.区分 === "通勤費" ? money.通勤費 : money.機械, M = MM.get(k) || { 数量: 0, 金額: 0, 単価: new Set(), なし: 0, 分類: pm ? pm.分類 : (guessMach(sp.名前) === "通勤車" ? "通勤車" : "（単価なし）") };
        M.数量 += qq; if (pm) { M.金額 += qq * pm.単価; M.単価.add(+pm.単価); M.分類 = pm.分類; } else M.なし++; MM.set(k, M);
      }
      if (r[9]) {
        const m = hin.get(r[9]) || { 日: {}, 計: 0 }; const q = Number(r[10]) || 0; m.日[x] = (m.日[x] || 0) + q; m.計 += q; hin.set(r[9], m);
        const pz = itemPrice(s, ["材料", "通勤費"], r[9], "", "", x);
        const ZZ = pz && pz.区分 === "通勤費" ? money.通勤費 : money.材料, Z = ZZ.get(r[9]) || { 数量: 0, 金額: 0, 単価: new Set(), なし: 0, 分類: pz ? pz.分類 : (/高速/.test(r[9]) ? "高速代" : "（単価なし）") };
        // 実費の欄に金額があれば、それを使う（高速代のルート違い・領収書の金額）。
        // 領収書精算で実費が空なら、前の書き方どおり数量の欄を金額とみる
        const jv = han(String(r[15] ?? "")).replace(/[,，円\s]/g, ""), jh = jv !== "" && !isNaN(Number(jv)) ? Number(jv) : null;
        if (pz) Z.分類 = pz.分類;
        if (jh != null) { Z.数量 += pz && pz.精算 ? 1 : (q || 1); Z.金額 += jh; Z.実費 = (Z.実費 || 0) + 1; if (pz && pz.精算) Z.精算 = true; }
        else if (pz && pz.精算) { Z.数量 += 1; Z.金額 += q; Z.精算 = true; }
        else { Z.数量 += q; if (pz) { Z.金額 += q * pz.単価; Z.単価.add(+pz.単価); } else Z.なし++; }
        ZZ.set(r[9], Z);
      }
    });
    (d.車 || []).forEach(r => {
      if (!r || !r[0]) return;
      const A = carAmount(s, r, x);
      car.push({ date: x, r, A, who: [r[5] || "", r[6] || "", r[7] || "", r[8] || ""] });
      // 金額：高速 → 通勤費（高速代・区間ごと）、燃料 → 材料（軽油・ガソリン）
      const MM = A.種 === "高速" ? money.通勤費 : money.材料, k = A.種 === "高速" ? "高速代（" + (r[2] || "区間なし") + "）" : r[0];
      const Z = MM.get(k) || { 数量: 0, 金額: 0, 単価: new Set(), なし: 0, 分類: A.種 === "高速" ? "高速代" : "固定単価" };
      Z.数量 += A.種 === "高速" ? A.回 : A.L; Z.金額 += A.金額;
      if (A.実費) Z.実費 = (Z.実費 || 0) + 1; else if (A.単価 != null) Z.単価.add(A.単価); else Z.なし++;
      MM.set(k, Z);
    });
    return { date: x, 天気: d.天気, 気温: d.気温, c, n, ot, 作業: workSummary(d), 昼: nDay, 夜: nNight };
  });
  shoku.sort((p, q) => order(p) - order(q));
  // 会社の順：自社 → 人工の多い順 → 会社未定
  const corder = [...comp.entries()].sort((p, q) => (q[0] === SELF) - (p[0] === SELF) || (p[0] === "（会社未定）") - (q[0] === "（会社未定）") || q[1].計 - p[1].計).map(x => x[0]);
  const ws = [...workers.values()].map(w => ({ ...w, 職種内訳: w.職種, 職種: [...w.職種.entries()].sort((p, q) => q[1] - p[1])[0][0], 会社: [...w.会社.entries()].sort((p, q) => q[1] - p[1])[0][0] }))
    .sort((p, q) => corder.indexOf(p.会社) - corder.indexOf(q.会社) || order(p.職種) - order(q.職種) || q.日数 - p.日数);
  const msum = m => [...m.values()].reduce((a, v) => a + v.金額, 0);
  money.計 = { 労務: msum(money.労務), 残業: msum(money.残業), 機械: msum(money.機械), 材料: msum(money.材料), 通勤費: msum(money.通勤費) };
  money.合計 = money.計.労務 + money.計.残業 + money.計.機械 + money.計.材料 + money.計.通勤費;
  money.なし = [...money.労務.values(), ...money.残業.values(), ...money.機械.values(), ...money.材料.values(), ...money.通勤費.values()].reduce((a, v) => a + v.なし, 0);
  const shokuBase = [...new Set(shoku.map(k => k.replace(/（夜）$/, "")))];   // 職種別の工種の順（昼・夜をまとめる）
  return { dates, shoku, shokuBase, perDay, ws, kizai, hin, comp, corder, money, byShoku, otRows, car, hasNight: perDay.some(p => p.夜) };
}
const r2 = x => String(Math.round(x * 100) / 100);
// その日の作業の要約（2026-10-31 ユーザー：日別に作業内容を要約した列）：作業予定（全部のページ）を重ねずに「・」でつなぐ。完了報告があれば（ ）で
function workSummary(d) {
  const seen = new Set(), out = [];
  (d.上 || []).forEach(r => { const w = String((r && r[1]) || "").trim(), k = String((r && r[4]) || "").trim(); if (!w || seen.has(w)) return; seen.add(w); out.push(w + (k ? `（${k}）` : "")); });
  return out.join("・");
}
function renderSum(all) {
  const s = site(); if (!s) { $("sumBody").innerHTML = `<p class="note">現場を選んでください</p>`; return; }
  const a = $("sumFrom").value, b = $("sumTo").value;
  showSumRange(); renderSumShow();
  if (!a || !b || a > b) return;
  const S = summarize(s, a, b), SM = !UI.hideMoney;   // SM＝金額を出す（2026-10-31 ユーザー：集計表の金額を隠せる）
  if (!S.dates.length) { $("sumBody").innerHTML = `<p class="note">${jpDate(a)}〜${jpDate(b)} に書いた日報がありません</p>`; return; }
  const dl = x => `${+x.slice(5, 7)}/${+x.slice(8)}`;
  const wcls = x => { const w = wd(x); return w === 0 ? "sun" : w === 6 ? "sat" : ""; };
  const h = [];
  // 出す表を選べる（2026-10-31 ユーザー：職種別も出せるように・残業の集計の欄は不要）。区切りの位置を覚えて、最後に選んだ表だけ出す
  const SECT = [], mark = n => SECT.push([n, h.length]);
  mark("日別");
  h.push(`<h3 class="sh">日別（${esc(s.名前)}　${jpDate(a)}〜${jpDate(b)}）</h3>`);
  const NT = S.hasNight || s.勤務 !== "昼";
  h.push(`<table class="sm"><thead><tr><th>日付</th><th>曜</th><th>天気</th><th>気温</th>${S.shoku.map(k => `<th>${esc(k)}</th>`).join("")}${NT ? "<th>昼</th><th>夜</th>" : ""}<th>人員計</th><th>残業計(h)</th><th>作業（作業予定の要約）</th></tr></thead><tbody>`);
  const tot = new Map(); let tn = 0, tot_ot = 0;
  for (const p of S.perDay) {
    h.push(`<tr class="${wcls(p.date)}"><td><a data-go="${p.date}">${dl(p.date)}</a></td><td class="c">${WD[wd(p.date)]}</td><td class="c">${esc(p.天気)}</td><td class="r">${esc(p.気温)}</td>${S.shoku.map(k => `<td class="r">${p.c.get(k) || ""}</td>`).join("")}${NT ? `<td class="r">${p.昼 || ""}</td><td class="r night">${p.夜 || ""}</td>` : ""}<td class="r"><b>${p.n}</b></td><td class="r">${p.ot ? r2(p.ot) : ""}</td><td class="wsum" title="${esc(p.作業)}">${esc(p.作業)}</td></tr>`);
    S.shoku.forEach(k => tot.set(k, (tot.get(k) || 0) + (p.c.get(k) || 0))); tn += p.n; tot_ot += p.ot;
  }
  h.push(`</tbody><tfoot><tr><td colspan="4">計（${S.dates.length}日）</td>${S.shoku.map(k => `<td class="r">${tot.get(k)}</td>`).join("")}${NT ? `<td class="r">${S.perDay.reduce((a2, p) => a2 + p.昼, 0)}</td><td class="r">${S.perDay.reduce((a2, p) => a2 + p.夜, 0)}</td>` : ""}<td class="r">${tn}</td><td class="r">${r2(tot_ot)}</td><td></td></tr></tfoot></table>`);

  mark("会社別");
  // 会社ごとの集計（日ごとの人数＝人工）
  h.push(`<h3 class="sh">会社別（日ごとの人数）</h3>`);
  h.push(`<table class="sm"><thead><tr><th>会社</th>${S.dates.map(x => `<th class="${wcls(x)}">${+x.slice(8)}</th>`).join("")}<th>人工計</th><th>人数</th><th>残業計(h)</th>${SM ? "<th>金額（労務＋残業）</th>" : ""}</tr></thead><tbody>`);
  for (const cp of S.corder) {
    const C = S.comp.get(cp), nppl = S.ws.filter(w => w.会社 === cp).length;
    h.push(`<tr class="${cp === "（会社未定）" ? "nocomp" : ""}"><td><b>${esc(cp)}</b></td>${S.dates.map(x => `<td class="r">${C.日[x] || ""}</td>`).join("")}<td class="r"><b>${C.計}</b></td><td class="r">${nppl}名</td><td class="r">${C.残業 ? r2(C.残業) : ""}</td>${SM ? `<td class="r">${yen(C.金額)}</td>` : ""}</tr>`);
  }
  h.push(`</tbody><tfoot><tr><td>計</td>${S.dates.map(x => `<td class="r">${S.perDay.find(p => p.date === x).n}</td>`).join("")}<td class="r">${tn}</td><td class="r">${S.ws.length}名</td><td class="r">${r2(tot_ot)}</td>${SM ? `<td class="r">${yen(S.money.計.労務 + S.money.計.残業)}</td>` : ""}</tr></tfoot></table>`);
  mark("職種別");
  // 職種別（工種）：工種ごとの日ごとの人数と、その工種で来た人（2026-10-31 ユーザー：日によって役割が違うので職種別と作業員名別に）
  const hasMoney = SM && ((s.単価 || []).length > 0 || Object.values(DB.職種単価).some(p => p.昼 || p.夜 || p.残業));
  // 2026-10-31 ユーザー：1人を 昼・夜・残 の3行にする（日付は1列ずつ）
  h.push(`<h3 class="sh">職種別（工種）の集計（1人を 昼・夜・残 の3行に。昼・夜＝○、残＝残業時間）</h3>`);
  h.push(`<table class="sm shk3"><thead><tr><th>工種</th><th>作業員名</th><th></th>${S.dates.map(x => `<th class="${wcls(x)}">${+x.slice(8)}</th>`).join("")}<th>人工</th>${hasMoney ? "<th>金額</th>" : ""}</tr></thead><tbody>`);
  for (const k of S.shokuBase) {
    const SK = S.byShoku.get(k); if (!SK) continue;
    const ppl = [...SK.人.values()].sort((p, q) => q.日数 - p.日数);
    h.push(`<tr class="cfirst shk"><td rowspan="${1 + ppl.length * 3}"><b>${esc(k)}</b><br><span class="note">${ppl.length}名</span></td><td class="note" colspan="2">日ごとの人数</td>${S.dates.map(x => `<td class="r"><b>${SK.日[x] || ""}</b></td>`).join("")}<td class="r"><b>${SK.計}</b></td>${hasMoney ? `<td class="r"><b>${yen(SK.金額)}</b></td>` : ""}</tr>`);
    ppl.forEach(w => [["昼", x => w.昼[x] ? "○" : "", w.昼数 || ""], ["夜", x => w.夜[x] ? "○" : "", w.夜数 || ""], ["残", x => w.残[x] ? `<span class="ot">${r2(w.残[x])}</span>` : "", w.残業 ? r2(w.残業) : ""]]
      .forEach(([z, f, t], j) => h.push(`<tr class="z${z}${j === 0 ? " pfirst" : ""}">${j === 0 ? `<td rowspan="3">${esc(w.名前)}</td>` : ""}<td class="c zl">${z}</td>${S.dates.map(x => `<td class="c">${f(x)}</td>`).join("")}<td class="r">${t}</td>${hasMoney ? "<td></td>" : ""}</tr>`)));
  }
  h.push(`</tbody></table>`);
  mark("作業員名別");
  // 作業員名別：日ごとにその日の工種（略）、工種ごとの日数（会社ごとにまとめる）
  h.push(`<h3 class="sh">作業員名別の集計（日ごとの工種：${SHORT_NOTE}、赤い数字は残業時間）</h3>`);
  h.push(`<table class="sm"><thead><tr><th>会社</th><th>作業員名</th>${S.dates.map(x => `<th class="${wcls(x)}">${+x.slice(8)}</th>`).join("")}<th>工種ごとの日数</th><th>日数</th><th>残業(h)</th>${hasMoney ? "<th>金額</th>" : ""}</tr></thead><tbody>`);
  for (const cp of S.corder) {
    const list = S.ws.filter(w => w.会社 === cp);
    list.forEach((w, i) => h.push(`<tr class="${i === 0 ? "cfirst" : ""}${cp === "（会社未定）" ? " nocomp" : ""}">${i === 0 ? `<td rowspan="${list.length}"><b>${esc(cp)}</b><br><span class="note">${list.length}名・${S.comp.get(cp).計}人工</span></td>` : ""}<td>${esc(w.名前)}</td>` +
      S.dates.map(x => x in w.日 ? `<td class="c kshort" title="${esc(w.日工種[x].join("・"))}">${w.日工種[x].map(shortK).join("")}${w.日[x] ? `<span class="ot">${r2(w.日[x])}</span>` : ""}</td>` : "<td></td>").join("") +
      `<td class="note">${[...w.職種内訳.entries()].sort((p, q) => q[1] - p[1]).map(([k, n]) => esc(k) + " " + n).join("・")}</td><td class="r"><b>${w.日数}</b></td><td class="r">${w.残業 ? r2(w.残業) : ""}</td>${hasMoney ? `<td class="r">${yen(w.金額)}</td>` : ""}</tr>`));
  }
  h.push(`</tbody></table>`);

  const mt = (ttl, m) => {
    if (!m.size) return;
    const kz = ttl === "使用機材";
    const ent = [...m.entries()].sort((p, q) => kz ? (p[0].split("\t")[1] || "").localeCompare(q[0].split("\t")[1] || "", "ja") : 0);
    h.push(`<h3 class="sh">${ttl}</h3><table class="sm"><thead><tr>${kz ? "<th>所有者</th>" : ""}<th>${kz ? "使用機材（規格）" : "材料名"}</th>${S.dates.map(x => `<th class="${wcls(x)}">${+x.slice(8)}</th>`).join("")}<th>計</th></tr></thead><tbody>`);
    for (const [k, v] of ent) { const [nm, own] = k.split("\t"); h.push(`<tr>${kz ? `<td>${esc(own)}</td>` : ""}<td>${esc(nm)}</td>${S.dates.map(x => `<td class="r">${v.日[x] ? r2(v.日[x]) : ""}</td>`).join("")}<td class="r"><b>${r2(v.計)}</b></td></tr>`); }
    h.push(`</tbody></table>`);
  };
  mark("使用機材");
  mt("使用機材", S.kizai); mt("材料", S.hin);
  if (!S.kizai.size && !S.hin.size) h.push(`<h3 class="sh">使用機材・材料</h3><p class="note">この期間に使用機材・材料の記入はありません。</p>`);
  mark("高速・燃料");
  h.push(SM ? carHTML(S, dl) : dropMoneyCols(carHTML(S, dl)));
  mark("金額");
  if (SM) h.push(moneyHTML(S));   // 金額を隠すときは金額の表を出さない（印刷にも出ない）
  // タブで1つずつ出す（2026-10-31 ユーザー：スクロールを減らす）。all（PDF）のときは全部
  const tab = sumTab(), out = [];
  SECT.forEach(([n, i], k) => { if (all || n === tab) out.push(...h.slice(i, k + 1 < SECT.length ? SECT[k + 1][1] : h.length)); });
  $("sumBody").innerHTML = out.join("");
}
// 高速・燃料の集計（2026-10-31 ユーザー）：だれごと・高速の区間ごと・日ごとの記録
function carHTML(S, dl) {
  const C = S.car, h = [`<h3 class="sh">高速・燃料</h3>`];
  if (!C.length) { h.push(`<p class="note">この期間に高速・燃料の記録はありません（日報の「高速・燃料」の表に書きます）。</p>`); return h.join(""); }
  const rows = carByWho(C), sum = f => rows.reduce((a, v) => a + f(v), 0), hasO = sum(v => v.他L) > 0;
  h.push(`<h4 class="mh">だれが（会社・名前・車両名・車番ごと）</h4><table class="sm"><thead><tr><th>会社</th><th>名前</th><th>車両名</th><th>車番</th><th>高速 往路</th><th>復路</th><th>往復</th><th>高速代</th><th>軽油(L)</th><th>ガソリン(L)</th>${hasO ? "<th>ほかの燃料(L)</th>" : ""}<th>燃料代</th><th>計</th></tr></thead><tbody>` +
    rows.map(v => `<tr>${v.who.map(w => `<td>${esc(w)}</td>`).join("")}<td class="r">${v.往路 || ""}</td><td class="r">${v.復路 || ""}</td><td class="r">${v.往復 || ""}</td><td class="r">${v.高速金額 ? yen(v.高速金額) : ""}</td><td class="r">${v.軽油 ? r2(v.軽油) : ""}</td><td class="r">${v.ガソリン ? r2(v.ガソリン) : ""}</td>${hasO ? `<td class="r">${v.他L ? r2(v.他L) : ""}</td>` : ""}<td class="r">${v.燃料金額 ? yen(v.燃料金額) : ""}</td><td class="r"><b>${yen(v.高速金額 + v.燃料金額)}</b></td></tr>`).join("") +
    `</tbody><tfoot><tr><td colspan="4">計</td><td class="r">${sum(v => v.往路)}</td><td class="r">${sum(v => v.復路)}</td><td class="r">${sum(v => v.往復)}</td><td class="r">${yen(sum(v => v.高速金額))}</td><td class="r">${r2(sum(v => v.軽油))}</td><td class="r">${r2(sum(v => v.ガソリン))}</td>${hasO ? `<td class="r">${r2(sum(v => v.他L))}</td>` : ""}<td class="r">${yen(sum(v => v.燃料金額))}</td><td class="r"><b>${yen(sum(v => v.高速金額 + v.燃料金額))}</b></td></tr></tfoot></table>`);
  const ar = carByArea(C);
  if (ar.length) h.push(`<h4 class="mh">高速の区間ごと</h4><table class="sm"><thead><tr><th>区間</th><th>往路</th><th>復路</th><th>往復</th><th>回数（往復は2）</th><th>金額</th></tr></thead><tbody>` +
    ar.map(v => `<tr><td>${esc(v.区間)}</td><td class="r">${v.往路 || ""}</td><td class="r">${v.復路 || ""}</td><td class="r">${v.往復 || ""}</td><td class="r">${v.回}</td><td class="r"><b>${yen(v.金額)}</b></td></tr>`).join("") +
    `</tbody><tfoot><tr><td>計</td><td class="r">${ar.reduce((a, v) => a + v.往路, 0)}</td><td class="r">${ar.reduce((a, v) => a + v.復路, 0)}</td><td class="r">${ar.reduce((a, v) => a + v.往復, 0)}</td><td class="r">${ar.reduce((a, v) => a + v.回, 0)}</td><td class="r">${yen(ar.reduce((a, v) => a + v.金額, 0))}</td></tr></tfoot></table>`);
  h.push(`<h4 class="mh">記録（日ごと）</h4><table class="sm"><thead><tr><th>日付</th><th>種類</th><th>往復</th><th>区間</th><th>給油(L)</th><th>会社</th><th>名前</th><th>車両名</th><th>車番</th><th>金額</th><th>メモ</th></tr></thead><tbody>` +
    C.map(x => `<tr class="${x.A.なし ? "nocomp" : ""}"><td><a data-go="${x.date}">${dl(x.date)}</a></td><td>${esc(x.r[0])}</td><td>${esc(x.r[1] || "")}</td><td>${esc(x.r[2] || "")}</td><td class="r">${x.A.種 === "高速" ? "" : r2(x.A.L)}</td>${x.who.map(w => `<td>${esc(w)}</td>`).join("")}<td class="r">${x.A.なし ? '<span class="ng">単価なし</span>' : yen(x.A.金額) + (x.A.実費 ? '<span class="note">（実費）</span>' : "")}</td><td>${esc(x.r[9] || "")}</td></tr>`).join("") + `</tbody></table>`);
  return h.join("");
}
// 残業の集計（2026-10-31 ユーザー：もっと分かりやすく）
//   上に まとめ（残業した人・延べ時間＝早出＋残業・金額）、作業員別（日ごとの時間・早出・残業・計・金額）、日ごとの計。ます目に マウスを載せると 始業〜終業
function otHTML(S, wcls) {
  const R = S.otRows, h = [];
  h.push(`<h3 class="sh">残業の集計</h3>`);
  if (!R.length) { h.push(`<p class="note">この期間に残業はありません。</p>`); return h.join(""); }
  const sum = f => R.reduce((a, x) => a + f(x), 0), ppl = new Map();
  R.forEach(x => { const p = ppl.get(x.key) || { 名前: x.名前, 会社: x.会社, 日: {}, 早出: 0, 残業: 0, 計: 0, 金額: 0, なし: 0, 回: 0, 夜: 0 }; const d = p.日[x.date] ||= { 計: 0, t: [] }; d.計 += x.計; d.t.push(`${x.昼夜} ${x.始業 || "?"}〜${x.終業 || "?"}（${x.工種}）`); p.早出 += x.早出; p.残業 += x.残業; p.計 += x.計; p.回++; if (x.昼夜 === "夜") p.夜 += x.計; if (x.金額 == null) p.なし++; else p.金額 += x.金額; ppl.set(x.key, p); });
  const days = S.dates.filter(dt => R.some(x => x.date === dt));
  const money = sum(x => x.金額 || 0), hasM = R.some(x => x.金額 != null);
  h.push(`<div class="otcards"><div><b>${ppl.size}</b>名<span>残業した人</span></div><div><b>${r2(sum(x => x.計))}</b>時間<span>延べ（${R.length}回）</span></div><div><b>${r2(sum(x => x.早出))}</b>時間<span>早出（定時より前）</span></div><div><b>${r2(sum(x => x.残業))}</b>時間<span>残業（定時より後）</span></div>${hasM ? `<div><b>${yen(money)}</b>円<span>残業の金額</span></div>` : ""}<div><b>${days.length}</b>日<span>残業のあった日</span></div></div>`);
  const NT = R.some(x => x.昼夜 === "夜");
  h.push(`<h4 class="mh">作業員別（ます目は その日の残業時間。マウスを載せると始業〜終業）</h4><table class="sm"><thead><tr><th>会社</th><th>作業員名</th>${days.map(x => `<th class="${wcls(x)}">${+x.slice(5, 7)}/${+x.slice(8)}</th>`).join("")}<th>回数</th><th>早出</th><th>残業</th>${NT ? "<th>うち夜勤</th>" : ""}<th>計(h)</th>${hasM ? "<th>金額</th>" : ""}</tr></thead><tbody>`);
  const list = [...ppl.values()].sort((a, b) => S.corder.indexOf(a.会社) - S.corder.indexOf(b.会社) || b.計 - a.計);
  for (const p of list) h.push(`<tr><td>${esc(p.会社)}</td><td>${esc(p.名前)}</td>${days.map(x => p.日[x] ? `<td class="r ot" title="${esc(p.日[x].t.join(" / "))}">${r2(p.日[x].計)}</td>` : "<td></td>").join("")}<td class="r">${p.回}</td><td class="r">${p.早出 ? r2(p.早出) : ""}</td><td class="r">${p.残業 ? r2(p.残業) : ""}</td>${NT ? `<td class="r night">${p.夜 ? r2(p.夜) : ""}</td>` : ""}<td class="r"><b>${r2(p.計)}</b></td>${hasM ? `<td class="r">${yen(p.金額)}${p.なし ? '<span class="ng note">（単価なし）</span>' : ""}</td>` : ""}</tr>`);
  h.push(`</tbody><tfoot><tr><td colspan="2">日ごとの計</td>${days.map(x => `<td class="r">${r2(R.filter(y => y.date === x).reduce((a, y) => a + y.計, 0))}</td>`).join("")}<td class="r">${R.length}</td><td class="r">${r2(sum(x => x.早出))}</td><td class="r">${r2(sum(x => x.残業))}</td>${NT ? `<td class="r">${r2(sum(x => x.昼夜 === "夜" ? x.計 : 0))}</td>` : ""}<td class="r">${r2(sum(x => x.計))}</td>${hasM ? `<td class="r">${yen(money)}</td>` : ""}</tr></tfoot></table>`);
  // 工種・昼夜ごと
  const byK = new Map(); R.forEach(x => { const k = x.工種 + (x.昼夜 === "夜" ? "（夜）" : ""), v = byK.get(k) || { 計: 0, 人: new Set(), 金額: 0 }; v.計 += x.計; v.人.add(x.key); v.金額 += x.金額 || 0; byK.set(k, v); });
  h.push(`<h4 class="mh">工種別</h4><table class="sm"><thead><tr><th>工種</th><th>人数</th><th>時間</th>${hasM ? "<th>金額</th>" : ""}</tr></thead><tbody>${[...byK].map(([k, v]) => `<tr><td>${esc(k)}</td><td class="r">${v.人.size}名</td><td class="r">${r2(v.計)}</td>${hasM ? `<td class="r">${yen(v.金額)}</td>` : ""}</tr>`).join("")}</tbody></table>`);
  return h.join("");
}
// 金額の表（工種別・残業・機械種類別・材料）。単価が無いものは「単価なし」
function moneyHTML(S) {
  const M = S.money, h = [];
  const s = site();
  h.push(`<h3 class="sh">金額（単価マスタから。期間 ${S.dates.length}日）</h3>`);
  if (s.単価確認 === false) h.push(`<p class="note ng">※ この現場の単価は たたき台 から写したまま（見直し前）です。「現場・名簿」の単価マスタで見直してください。</p>`);
  if (!(s.単価 || []).length && DB.単価たたき台) h.push(`<p class="note ng">※ この現場の単価がまだありません。「現場・名簿」の単価マスタで、たたき台を写して見直してください。</p>`);
  if (!(s.単価 || []).length && !Object.values(DB.職種単価).some(p => p.昼 || p.夜 || p.残業)) { h.push(`<p class="note ng">単価がまだありません。「現場・名簿」の「工種（職種）の一覧」か「単価マスタ」で単価を入れると金額が出ます。</p>`); return h.join(""); }
  const CORD = ["車両", "重機", "小型機械", "電動工具", "固定単価", "領収書精算", "高速代", "通勤車", "（単価なし）"];
  const tb = (ttl, m, unit, lab, withCls) => {
    if (!m.size) return;
    const ent = [...m.entries()].sort((a, b) => withCls ? CORD.indexOf(a[1].分類) - CORD.indexOf(b[1].分類) : 0);
    h.push(`<h4 class="mh">${ttl}</h4><table class="sm"><thead><tr>${withCls ? "<th>分類</th>" : ""}<th>${lab}</th><th>数量</th><th>単価</th><th>金額</th></tr></thead><tbody>`);
    for (const [k, v] of ent) {
      const [nm, own] = k.split("\t");
      h.push(`<tr class="${v.なし ? "nocomp" : ""}">${withCls ? `<td>${esc(v.分類 || "")}</td>` : ""}<td>${esc(nm)}${own ? `<span class="note">　${esc(own)}</span>` : ""}</td><td class="r">${v.精算 ? v.数量 + "回" : r2(v.数量) + unit}</td><td class="r">${v.精算 && !v.単価.size ? "領収書精算" : v.単価.size ? [...v.単価].map(yen).join("・") : v.実費 && !v.なし ? "実費" : '<span class="ng">単価なし</span>'}${v.実費 && (v.単価.size || v.精算) ? `<span class="note">（実費 ${v.実費}件）</span>` : ""}</td><td class="r"><b>${yen(v.金額)}</b>${v.なし && v.単価.size ? `<span class="ng note">（${v.なし}件 単価なし）</span>` : ""}</td></tr>`);
    }
    h.push(`</tbody><tfoot><tr><td colspan="${withCls ? 4 : 3}">計</td><td class="r">${yen([...m.values()].reduce((a, v) => a + v.金額, 0))}</td></tr></tfoot></table>`);
  };
  tb("労務（工種別・昼夜）", M.労務, "人工", "工種");
  tb("残業（工種別）", M.残業, "時間", "工種");
  tb("機械（車両・重機・小型機械・電動工具）", M.機械, "台日", "使用機材（規格）・所有者", true);
  tb("材料（固定単価・領収書精算）", M.材料, "", "材料名", true);
  tb("通勤費（高速代・通勤車）", M.通勤費, "", "項目", true);
  h.push(`<table class="sm mtot"><tr><th>労務</th><td class="r">${yen(M.計.労務)}</td><th>残業</th><td class="r">${yen(M.計.残業)}</td><th>機械</th><td class="r">${yen(M.計.機械)}</td><th>材料</th><td class="r">${yen(M.計.材料)}</td><th>通勤費</th><td class="r">${yen(M.計.通勤費)}</td><th>合計</th><td class="r"><b>${yen(M.合計)}</b></td></tr></table>`);
  if (M.なし) h.push(`<p class="note ng">単価の無い項目が ${M.なし} 件あります（薄い黄色）。「現場・名簿」の単価マスタで「日報から項目を足す」と、足りない項目が並びます。</p>`);
  return h.join("");
}
const SUM_TABLES = ["日別", "会社別", "職種別", "作業員名別", "使用機材", "高速・燃料", "金額"];
const sumTab = () => SUM_TABLES.includes(UI.sumTab) && !(UI.hideMoney && UI.sumTab === "金額") ? UI.sumTab : "日別";
// 表の中の金額の列を消す（高速・燃料の表。見出しの最後の段で「金額・高速代・燃料代・計」の列）
function dropMoneyCols(html) {
  const box = document.createElement("div"); box.innerHTML = html;
  const pos = row => { let p = 0; return [...row.cells].map(c => { const a = p; p += c.colSpan; return [c, a]; }); };
  box.querySelectorAll("table").forEach(t => {
    const hr = t.tHead && t.tHead.rows[t.tHead.rows.length - 1]; if (!hr) return;
    const kill = pos(hr).filter(([c]) => /金額|高速代|燃料代|^計$/.test(c.textContent.trim())).map(([, a]) => a);
    if (kill.length) [...t.rows].forEach(row => pos(row).forEach(([c, a]) => { if (kill.includes(a) && c.colSpan === 1) c.remove(); }));
  });
  return box.innerHTML;
}
function renderSumShow() {
  const t = sumTab();
  $("sumShow").innerHTML = `<span class="stabs">` + SUM_TABLES.filter(n => !(UI.hideMoney && n === "金額")).map(n => `<button class="stab${n === t ? " on" : ""}" data-st="${n}">${n === "職種別" ? "職種別（工種）" : n === "作業員名別" ? "作業員別" : n}</button>`).join("") + `</span>` +
    `<span class="rsep"></span><button class="sub small tg hm${UI.hideMoney ? " on" : ""}" id="sumHM" title="集計表の金額を全部隠す／表示する（画面・印刷（PDF）・Excel に反映）">${UI.hideMoney ? "金額を表示" : "金額を隠す"}</button>`;
}
$("sumShow").addEventListener("click", e => {
  if (e.target.closest("#sumHM")) { UI.hideMoney = !UI.hideMoney; saveUI(); renderSumShow(); renderSum(); return; }
  const b = e.target.closest("[data-st]"); if (!b) return;
  UI.sumTab = b.dataset.st; saveUI(); renderSumShow(); renderSum(); $("sumBody").scrollTop = 0;
});
$("sumBody").addEventListener("click", e => { const a = e.target.closest("[data-go]"); if (a) { UI.day = a.dataset.go; showTab("sheet"); setDay(a.dataset.go); } });
["sumFrom", "sumTo"].forEach(id => $(id).addEventListener("change", renderSum));
$("sumPrev").onclick = () => { const [a, b] = period(addD($("sumFrom").value, -1)); $("sumFrom").value = a; $("sumTo").value = b; renderSum(); };
$("sumNext").onclick = () => { const [a, b] = period(addD($("sumTo").value, 1)); $("sumFrom").value = a; $("sumTo").value = b; renderSum(); };
$("bSumXls").onclick = () => { const s = site(); if (!s) return; const a = $("sumFrom").value, b = $("sumTo").value; exportXlsx(datesIn(a, b).filter(x => hasData(s.日[x])), true, [a, b]); };
$("bSumPdf").onclick = () => { if (!site()) return; printDays([], [$("sumFrom").value, $("sumTo").value]); };

/* ================= 現場・名簿 ================= */
function renderMaster() {
  const s = site(), g = gen();
  const h = [];
  h.push(`<h3 class="sh">マスタ（ほかのアプリと同じ所から読む）</h3><div class="mbox"><p class="note">
    現場・元請：現場一覧（NAS の受注台帳・注文書台帳・工事部会資料から毎朝作るもの）${MS.現場.length} 件 ／ 作業員名：配置表の名簿 社員 ${MS.社員.length}・従業員 ${MS.従業員.length}・外注業者 ${MS.外注業者.length} ／ 使用機材：自社機械 ${MS.機械.length}<br>
    ${MS.読んだ ? "読んだ時刻 " + esc(MS.読んだ) + (MS.元 ? "（" + esc(MS.元) + "）" : "") : "まだ読んでいません。上の「サインイン」で会社の Microsoft アカウントに入ると読みます。"}<br>
    サインインしていると開いたとき・30分ごとに読み直します。読めないときは前に読んだ控えで動きます。名簿に無い人・機械は日報に打てば、この現場の分として覚えます。</p>
    <button class="sub" id="mReload"${msalApp && msalApp.getActiveAccount() ? "" : " disabled"}>マスタを読み直す</button></div>`);
  h.push(`<h3 class="sh">元請業者</h3><div class="mrow"><div class="mbox"><table class="ed"><tr><th>元請業者名</th><th>現場数</th><th></th></tr>`);
  for (const x of DB.元請) h.push(`<tr class="${x.id === UI.gen ? "on" : ""}"><td><input data-gen="${x.id}" value="${esc(x.名前)}" style="width:240px"></td><td class="r">${DB.現場.filter(y => y.元請id === x.id).length}</td><td class="btn"><button class="sub small" data-pickgen="${x.id}">選ぶ</button> <button class="sub small danger" data-delgen="${x.id}">消す</button></td></tr>`);
  h.push(`</table><div class="btns"><button id="mAddGen">＋ 元請業者を追加</button></div></div></div>`);
  if (g) {
    // 2026-10-31 ユーザー：現場の工期を決める・完了したらチェック。正式名称は日報・Excel に出る名前、社内名称は選ぶときの名前
    h.push(`<h3 class="sh">${esc(g.名前)} の現場</h3><div class="mrow"><div class="mbox"><table class="ed sites"><tr><th>現場コード</th><th>現場名称（社内）</th><th>現場名称（正式・日報に出る名前）</th><th>工期 始まり</th><th>工期 終わり</th><th>担当者</th><th>完了</th><th>日報の日数</th><th></th></tr>`);
    for (const x of DB.現場.filter(y => y.元請id === g.id).sort((a, b) => (!!a.完了 - !!b.完了) || String(b.コード || "").localeCompare(String(a.コード || "")))) h.push(`<tr class="${x.id === UI.site ? "on" : ""}${x.完了 ? " done" : ""}"><td><input data-sid="${x.id}" data-sx="コード" value="${esc(x.コード || "")}" style="width:70px"></td><td><input data-sid="${x.id}" data-sx="略称" value="${esc(x.略称 || "")}" style="width:120px"></td><td><input data-site="${x.id}" value="${esc(x.名前)}" style="width:300px" title="印刷・Excel に出る現場名"></td>` +
      `<td><input type="date" data-sid="${x.id}" data-sx="工期開始" value="${esc(x.工期開始 || "")}"></td><td><input type="date" data-sid="${x.id}" data-sx="工期終了" value="${esc(x.工期終了 || "")}"></td><td><input data-sid="${x.id}" data-sx="現場担当者" value="${esc(x.現場担当者 || "")}" style="width:100px"></td>` +
      `<td class="c"><input type="checkbox" data-sid="${x.id}" data-sx="完了"${x.完了 ? " checked" : ""} title="工事が終わったらチェック（現場を選ぶ一覧の後ろに回します）"></td><td class="r">${Object.keys(x.日).length}</td><td class="btn"><button class="sub small" data-picksite="${x.id}">選ぶ</button> <button class="sub small danger" data-delsite="${x.id}">消す</button></td></tr>`);
    h.push(`</table><div class="btns"><button id="mAddSite">＋ 現場を追加</button></div></div></div>`);
  }
  if (s) {
    h.push(`<h3 class="sh">${esc(s.名前)} の決まり</h3><div class="mrow"><div class="mbox"><table class="ed">
      <tr><th>定時の始業</th><td><input data-sf="始業" value="${esc(s.始業)}" inputmode="numeric"></td></tr>
      <tr><th>定時の終業</th><td><input data-sf="終業" value="${esc(s.終業)}" inputmode="numeric"></td></tr>
      <tr><th>勤務（昼・夜）</th><td><select data-sf="勤務">${["昼", "夜", "昼夜"].map(k => `<option${k === s.勤務 ? " selected" : ""}>${k}</option>`).join("")}</select></td></tr>
      <tr><th>夜の定時の始業</th><td><input data-sf="夜始業" value="${esc(s.夜始業)}" inputmode="numeric"></td></tr>
      <tr><th>夜の定時の終業</th><td><input data-sf="夜終業" value="${esc(s.夜終業)}" inputmode="numeric"></td></tr>
      <tr><th>締め日</th><td><input data-sf="締め日" value="${esc(s.締め日)}" inputmode="numeric" title="20 → 21日〜翌月20日。0 か 31 → 月末締め"></td></tr>
      <tr><th>累計の数え方</th><td><select data-sf="累計"><option${s.累計 === "期間" ? " selected" : ""}>期間</option><option${s.累計 === "全体" ? " selected" : ""}>全体</option></select></td></tr>
      </table><p class="note">勤務：昼＝昼の現場、夜＝夜の現場、昼夜＝作業員の行の「昼夜」で昼・夜を分ける（名前を入れると昼。夜勤の人は「夜」）。夜の定時は日をまたいでよい（21:00〜6:00）。<br>残業時間＝その人の昼夜の定時の終業より遅い分＋始業より早い分（早出）。作業員の終業時間を入れると自動。<br>累計：期間＝締め期間の中で足す／全体＝この現場のはじめから足す。</p></div>`);
    h.push(rosterHTML(s));
    h.push(`<div class="mbox"><h4>使用機材</h4><table class="ed"><tr><th>使用機材</th><th>規格・形式</th><th>所有者</th><th></th></tr>`);
    const os = owners(s), nAbbr = Object.values(s.日).reduce((a, d) => a + d.下.filter(r => r && r[5] && splitOwner(r[5]).所有者).length, 0) + s.機材.filter(m => splitOwner(m.名前).所有者).length;
    s.機材.forEach((m, i) => h.push(`<tr class="${m.所有者 ? "" : "nocomp"}"><td><input data-mn="${i}" value="${esc(m.名前)}"></td><td><input data-mk="${i}" value="${esc(m.規格)}"></td>` +
      `<td><select data-mo="${i}"><option value="">（未定）</option>${os.map(o => `<option${o === m.所有者 ? " selected" : ""}>${esc(o)}</option>`).join("")}${m.所有者 && !os.includes(m.所有者) ? `<option selected>${esc(m.所有者)}</option>` : ""}<option value="__new">＋ 新しい所有者…</option></select></td>` +
      `<td class="btn"><button class="sub small danger" data-delm="${i}">消す</button></td></tr>`));
    h.push(`<tr><td><input id="mNewM" placeholder="機材を足す"></td><td><input id="mNewMk" placeholder="規格"></td><td></td><td class="btn"><button class="small" id="mAddM">追加</button></td></tr></table>
      <p class="note">所有者＝リース会社・自社・外注業者。機材名の頭の（ハ）（真）などは所有者に置き換えます（（ハ）＝${esc(SELF)}、ほかは頭の字で始まる会社・外注業者）。</p>
      <div class="btns"><button class="sub small" id="mKzFix"${nAbbr ? "" : " disabled"}>（ハ）（真）を所有者に移す${nAbbr ? `（${nAbbr}か所）` : ""}</button><button class="sub small" id="mKzApply">所有者を日報に入れる</button></div></div></div>`);
  }
  if (s) h.push(siteKouHTML(s));
  if (s) h.push(priceHTML(s));
  if (s) h.push(`<h3 class="sh">揺らぎを見つけてそろえる（${esc(siteLabel(s))}）</h3><div class="mbox"><p class="note">作業員名・工種・使用機材・規格・材料名で、同じものらしいのに書き方がちがうもの（全角半角・空白・1字ちがい など）を探します。そろえる書き方を選ぶと、この現場の日報をまとめて直します。</p><button class="sub" id="mYure">揺らぎを探す</button><div id="yureList"></div></div>`);
  h.push(`<h3 class="sh">工種（職種）・天気の一覧（全部の現場で共通）</h3><div class="mrow"><div class="mbox"><h4>工種（職種）と1日あたりの単価</h4>
    <table class="ed shoku"><tr><th></th><th>工種（職種）</th><th>昼（円/日）</th><th>夜（円/日）</th><th>残業（円/時間）</th><th></th></tr>` +
    DB.職種.map((k, i) => { const p = DB.職種単価[k] || {}; return `<tr><td class="btn"><button class="sub small" data-kup="${i}" title="上へ"${i ? "" : " disabled"}>↑</button><button class="sub small" data-kdn="${i}" title="下へ"${i < DB.職種.length - 1 ? "" : " disabled"}>↓</button></td><td><input data-kn="${i}" value="${esc(k)}" autocomplete="off"></td>` +
      ["昼", "夜", "残業"].map(f => `<td><input class="r" data-kp="${i}" data-kf="${f}" value="${esc(p[f] ?? "")}" inputmode="decimal" autocomplete="off" placeholder="${f === "夜" && p.昼 ? esc(p.昼) : ""}"></td>`).join("") +
      `<td class="btn"><button class="sub small" data-kper="${i}" title="何月何日〜何月何日はこの単価、を入れる">期間ごと${(p.期間 || []).length ? "（" + p.期間.length + "）" : ""}</button> <button class="sub small danger" data-kdel="${i}">消す</button></td></tr>`; }).join("") +
    `<tr><td></td><td><input id="mNewK" placeholder="工種を足す" autocomplete="off"></td><td><input class="r" id="mNewK昼" inputmode="decimal" autocomplete="off"></td><td><input class="r" id="mNewK夜" inputmode="decimal" autocomplete="off"></td><td><input class="r" id="mNewK残業" inputmode="decimal" autocomplete="off"></td><td class="btn"><button class="small" id="mAddK">追加</button></td></tr></table>
    <p class="note">この順で集計に並びます。表の単価は いつもの単価、「期間ごと」で 何月何日〜何月何日の単価 を入れると、その期間はそちらを使います。単価は集計の金額に使います（現場の「単価マスタ」に同じ工種の行があれば、そちらが先）。<br>夜が空なら昼の単価。全部の PC・iPad で共通です（サインインしていると共有されます）。</p>
    <div class="btns"><button class="sub small" id="mKFromDays">日報に出てくる工種を足す</button></div></div>
    <div class="mbox"><h4>天気</h4><textarea id="mWx" rows="6" cols="16">${esc(DB.天気.join("\n"))}</textarea></div>
    <div class="mbox"><h4>会社・所有者の一覧</h4><textarea id="mCo" rows="6" cols="22">${esc((DB.会社一覧 || []).join("\n"))}</textarea><p class="note">1行に1つ。日報の会社・所有者の選択肢に出ます<br>（自社・配置表の外注業者・リース会社はいつも出ます）。</p></div></div>`);
  h.push(`<h3 class="sh">データ</h3><div class="mbox"><p class="note">日報はこのブラウザの中に保存しています（ほかのパソコン・iPad とは自動で共有しません）。<br>上の「保存」でこの現場をファイル（.json）に出し、ほかの端末で「取込」すると移せます。「取込」は元の作業日報の Excel（.xlsx）も読めます。</p>
    <button class="sub" id="mSaveAll">全部の現場を保存（.json）</button></div>`);
  $("masterBody").innerHTML = h.join("");
}
$("masterBody").addEventListener("change", e => {
  const t = e.target, s = site(), dsx = t.dataset;
  if (t.id === "mRfC" || t.id === "mRfK") { UI.rFilt = { ...(UI.rFilt || {}), [t.id === "mRfC" ? "会社" : "工種"]: t.value }; saveUI(); renderMaster(); return; }
  if (dsx.sx) {
    const x = DB.現場.find(y => y.id === dsx.sid); if (!x) return;
    if (x.id === UI.site && !canEdit()) { renderMaster(); return; }
    const v = t.type === "checkbox" ? t.checked : dsx.sx === "コード" ? han(t.value.trim()) : t.value.trim();
    const f = dsx.sx === "工期開始" ? v : x.工期開始, to = dsx.sx === "工期終了" ? v : x.工期終了;
    if (f && to && to < f) { toast("工期の終わりが始まりより前です"); renderMaster(); return; }
    x[dsx.sx] = v; x.設定更新 = now(); save(); renderPickers(); renderMaster();
    if (dsx.sx === "完了") toast(v ? `${siteLabel(x)} を完了にしました（現場を選ぶ一覧の後ろに出ます）` : `${siteLabel(x)} の完了を外しました`);
    return;
  }
  if ((dsx.sf || dsx.wn || dsx.wk || dsx.wc || dsx.mn || dsx.mk || dsx.mo || dsx.pi != null || dsx.site === UI.site) && !canEdit()) { renderMaster(); return; }
  if (dsx.gen) { DB.元請.find(x => x.id === dsx.gen).名前 = t.value.trim(); }
  else if (dsx.site) { DB.現場.find(x => x.id === dsx.site).名前 = t.value.trim(); }
  else if (dsx.sf) { let v = t.value.trim(); if (["始業", "終業", "夜始業", "夜終業"].includes(dsx.sf)) v = normTime(v); if (dsx.sf === "締め日") v = Number(han(v)) || 0; s[dsx.sf] = v; t.value = v; }
  else if (dsx.wn) s.作業員[+dsx.wn].名前 = t.value.trim();
  else if (dsx.wk) s.作業員[+dsx.wk].職種 = t.value.trim();
  else if (dsx.wc) {
    let v = t.value;
    if (v === "__new") { v = (prompt("新しい会社の名前") || "").trim(); if (!v) { renderMaster(); return; } }
    const w = s.作業員[+dsx.wc], old = w.会社 || ""; w.会社 = v; s.設定更新 = now(); const n = v ? compToDays(s, w.名前, old, v) : 0; save(); renderMaster(); refreshSheet(); if (n) toast(`日報の ${n} 行の会社を「${v}」にしました`); return;
  }
  else if (dsx.mn) s.機材[+dsx.mn].名前 = t.value.trim();
  else if (dsx.mk) s.機材[+dsx.mk].規格 = t.value.trim();
  else if (dsx.pi != null) {
    const pr = s.単価[+dsx.pi]; let v = t.value.trim();
    if (["単価", "昼", "夜", "残業"].includes(dsx.pf)) { v = han(v).replace(/[,円]/g, ""); if (v !== "" && isNaN(Number(v))) { toast("単価は数字で入れてください"); renderMaster(); return; } }
    pr[dsx.pf] = v; s.設定更新 = now(); save(); return;
  }
  else if (dsx.mo) {
    let v = t.value;
    if (v === "__new") { v = (prompt("新しい所有者（リース会社・外注業者など）の名前") || "").trim(); if (!v) { renderMaster(); return; } }
    s.機材[+dsx.mo].所有者 = v; s.設定更新 = now(); save(); renderMaster(); return;
  }
  else if (dsx.kn != null) {
    const i = +dsx.kn, old = DB.職種[i], v = t.value.trim();
    if (!v || v === old) { if (!v) toast("工種の名前を入れてください"); renderMaster(); return; }
    // 日報の中の古い名前の数
    let n = 0; DB.現場.forEach(x => Object.values(x.日).forEach(d => { d.下.forEach(r => { if (r && r[3] === old) n++; }); d.上.forEach(r => { if (r && r[0] === old) n++; }); }));
    const merge = DB.職種.includes(v);
    if (!confirm(`工種「${old}」を「${v}」${merge ? "にまとめます（「" + v + "」はもう一覧にあります）" : "に名前を変えます"}。\n全部の現場の日報の「${old}」${n} か所も「${v}」に置き換えます。\nこれから「${old}」と打ったり Excel から読んだりしても「${v}」になります。よいですか？`)) { renderMaster(); return; }
    if (!canEdit()) { renderMaster(); return; }
    DB.現場.forEach(x => { let ch = false; Object.values(x.日).forEach(d => { let c2 = false; d.下.forEach(r => { if (r && r[3] === old) { r[3] = v; c2 = true; } }); d.上.forEach(r => { if (r && r[0] === old) { r[0] = v; c2 = true; } }); if (c2) { d.更新 = now(); ch = true; } });
      x.作業員.forEach(w => { if (w.職種 === old) { w.職種 = v; ch = true; } }); (x.単価 || []).forEach(p => { if (p.区分 !== "機械" && p.区分 !== "材料" && p.名前 === old) { p.名前 = v; ch = true; } }); if (ch) x.設定更新 = now(); });
    if (merge) { DB.職種.splice(i, 1); if (!DB.職種単価[v] && DB.職種単価[old]) DB.職種単価[v] = DB.職種単価[old]; delete DB.職種単価[old]; }
    else { DB.職種[i] = v; if (DB.職種単価[old]) { DB.職種単価[v] = DB.職種単価[old]; delete DB.職種単価[old]; } }
    DB.職種別名[old] = v; Object.keys(DB.職種別名).forEach(k => { if (DB.職種別名[k] === old) DB.職種別名[k] = v; }); delete DB.職種別名[v];
    DB.共通更新 = now(); UNDO = []; REDO = []; save(); renderMaster(); refreshSheet(); toast(`「${old}」を「${v}」にしました（日報 ${n} か所）`); return;
  }
  else if (dsx.kp != null) {
    const k = DB.職種[+dsx.kp]; let v = han(t.value).replace(/[,円]/g, "");
    if (v !== "" && isNaN(Number(v))) { toast("単価は数字で入れてください"); renderMaster(); return; }
    (DB.職種単価[k] ||= {})[dsx.kf] = v; DB.共通更新 = now(); save(); renderMaster(); return;
  }
  else if (t.id === "mWx") { DB.天気 = t.value.split(/\n/).map(x => x.trim()).filter(Boolean); DB.共通更新 = now(); }
  else if (t.id === "mCo") { DB.会社一覧 = [...new Set(t.value.split(/\n/).map(x => x.trim()).filter(Boolean))]; DB.共通更新 = now(); }
  else return;
  if (s && (dsx.sf || dsx.wn || dsx.wk || dsx.mn || dsx.mk)) s.設定更新 = now();
  if (dsx.site) DB.現場.find(x => x.id === dsx.site).設定更新 = now();
  save(); renderPickers();
});
let PCSV = null;   // 単価の CSV を読む表（労務・機械・材料・通勤費）
$("masterBody").addEventListener("change", async e => {
  if (e.target.id === "mPrFile") {
    const f = e.target.files[0], s = site(); if (!f || !s || !canEdit()) return;
    const rows = parseCsv(await f.text()), hd = (rows[0] || []).map(x => x.replace(/\s/g, "")), ix = k => hd.indexOf(k);
    const g = (r, ...ks) => { for (const k of ks) if (ix(k) >= 0) return (r[ix(k)] || "").trim(); return ""; };
    const num = v => han(v).replace(/[,，円]/g, ""), dt = v => v.replace(/\//g, "-");
    let list = [];
    if (ix("区分") >= 0) {
      // 前の形（区分 労務／残業／機械／材料／その他 の まとめた CSV）
      list = migratePrices(rows.slice(1).map(r => ({ 区分: g(r, "区分"), 昼夜: g(r, "昼夜"), 名前: g(r, "名前"), 規格: g(r, "規格"), 所有者: g(r, "所有者"), 単位: g(r, "単位"), 単価: num(g(r, "単価")), 開始: dt(g(r, "開始")), 終了: dt(g(r, "終了")), メモ: g(r, "メモ") })).filter(x => x.名前 || x.区分 !== "労務"));
    } else {
      const K = PCSV || (ix("工種") >= 0 ? "労務" : ix("品名") >= 0 ? "材料" : "機械");
      const okCls = PT.filter(T => T.区分 === K).map(T => T.分類);
      list = rows.slice(1).map(r => K === "労務"
        ? { 区分: "労務", 名前: g(r, "工種", "名前"), 昼: num(g(r, "昼")), 夜: num(g(r, "夜")), 残業: num(g(r, "残業")), 開始: dt(g(r, "開始")), 終了: dt(g(r, "終了")), メモ: g(r, "メモ") }
        : { 区分: K, 分類: okCls.includes(g(r, "分類")) ? g(r, "分類") : okCls[0], 名前: g(r, "名前", "品名", "材料名", "区間"), 規格: g(r, "規格"), 所有者: g(r, "所有者"), 単位: g(r, "単位"), 単価: num(g(r, "単価")), 開始: dt(g(r, "開始")), 終了: dt(g(r, "終了")), メモ: g(r, "メモ") })
        .filter(x => x.名前 || x.分類 === "高速代");
    }
    const P = s.単価 ||= []; let n = 0, u = 0;
    for (const x of list) {
      const same = P.find(p => p.区分 === x.区分 && (p.分類 || "") === (x.分類 || "") && nk(p.名前) === nk(x.名前) && nk(p.規格) === nk(x.規格) && nk(p.所有者) === nk(x.所有者) && (p.開始 || "") === (x.開始 || ""));
      if (same) { Object.assign(same, x); u++; } else { P.push(x); n++; }
    }
    PCSV = null; s.設定更新 = now(); save(); renderMaster(); toast(`単価を読みました（足した ${n} 行・直した ${u} 行）`); return;
  }
  if (e.target.id !== "mCompFile") return;
  const f = e.target.files[0], s = site(); if (!f || !s || !canEdit()) return;
  const rows = parseCsv(await f.text());
  const hd = rows[0] || [], iN = Math.max(0, hd.findIndex(x => /名前|氏名/.test(x))), iC = hd.findIndex(x => /会社/.test(x));
  if (iC < 0) { toast("CSV の1行目に「名前」「会社」の見出しが要ります"); return; }
  let n = 0;
  rows.slice(1).forEach(r => { const nm = r[iN], cp = r[iC]; if (nm && cp) { setRosterCompany(s, nm, cp); n++; } });
  save(); renderMaster(); refreshSheet(); toast(`${n}名の会社を読み、日報にも入れました`);
});
$("masterBody").addEventListener("click", e => {
  const s0 = site(), t0 = e.target.closest("button");
  if (s0 && t0 && (t0.id === "mSkAdd" || t0.id === "mSkInit" || t0.id === "mSkAll" || t0.dataset.skdel != null || t0.dataset.skadd != null)) {
    if (!canEdit()) return;
    const s = s0, used = [];
    Object.values(s.日).forEach(d => d.下.forEach(r => { if (r && r[0] && r[3] && !used.includes(r[3])) used.push(r[3]); }));
    if (t0.id === "mSkInit") { s.職種 = DB.職種.filter(k => used.includes(k)); used.forEach(k => { if (!s.職種.includes(k)) s.職種.push(k); }); if (!s.職種.length) s.職種 = [...DB.職種]; toast("日報で使っている工種で一覧を作りました。要らない工種は「外す」で外してください", 5000); }
    else if (t0.id === "mSkAll") { if (!confirm("この現場の工種の一覧をやめて、共通の一覧（全部の現場）を出しますか？")) return; s.職種 = []; }
    else if (t0.dataset.skdel != null) { const k = t0.dataset.skdel; if (s.職種.length <= 1) { toast("工種が無くなります。決めるのをやめるときは下のボタンで"); return; } s.職種 = s.職種.filter(x => x !== k); }
    else {
      const v = (t0.dataset.skadd != null ? t0.dataset.skadd : $("mSkNew").value).trim(); if (!v) { toast("足す工種を入れてください"); return; }
      if (!s.職種 || !s.職種.length) s.職種 = [...DB.職種];
      addSiteKou(s, v, true); setTimeout(() => askKouPrice(s, [DB.職種別名[v] || v]), 120);
    }
    s.設定更新 = now(); save(); renderMaster(); return;
  }
  const t = e.target.closest("button"); if (!t) return;
  const s = site(), dsx = t.dataset;
  if ((dsx.delw != null || dsx.delm != null || t.id === "mAddW" || t.id === "mAddM" || dsx.delsite === UI.site) && !canEdit()) return;
  if (t.id === "mAddGen") return editDlg("元請業者を追加", [["元請業者名", "name", ""]], v => { if (!v.name) return; const g = { id: uid(), 名前: v.name }; DB.元請.push(g); UI.gen = g.id; UI.site = null; save(); renderAll(); });
  if (t.id === "mAddSite") return addSiteDlg({ shime: "20" });
});
// 現場を追加（2026-10-31 ユーザー：正式・社内の名称、現場コード、工期、担当者）。足りない・おかしいときは入れた値のまま開き直す
function addSiteDlg(p) {
  // 元請は一覧から選ぶ（2026-10-31 ユーザー：元請の名前の欄）。記入者は日によって替わるので、ここでは決めない
  const gs = DB.元請.map(g => g.名前), g0 = p.gen || (gen() ? gen().名前 : gs[0] || "");
  if (!gs.length) { toast("先に「＋ 元請業者を追加」で元請を入れてください"); return; }
  editDlg("現場を追加", [["元請業者", "gen", g0, gs], ["現場名称（正式）", "name", p.name || ""], ["現場名称（社内）", "abbr", p.abbr || ""], ["現場コード", "code", p.code || ""], ["工期 始まり", "from", p.from || "", null, "date"], ["工期 終わり", "to", p.to || "", null, "date"], ["担当者", "boss", p.boss || ""], ["締め日", "shime", p.shime ?? "20"]], v => {
    if (!v.name) { toast("現場名称（正式）を入れてください"); addSiteDlg(v); return; }
    if (v.from && v.to && v.to < v.from) { toast("工期の終わりが始まりより前です。直してください"); addSiteDlg(v); return; }
    const g = DB.元請.find(x => x.名前 === v.gen) || gen();
    const x = fixSite({ id: uid(), 元請id: g.id, 名前: v.name, 略称: v.abbr, コード: han(v.code), 工期開始: v.from, 工期終了: v.to, 現場担当者: v.boss, 締め日: Number(han(v.shime)) || 0, 設定更新: now() });
    UI.gen = g.id;
    DB.現場.push(x); UI.site = x.id; save(); renderAll();
  });
}
$("masterBody").addEventListener("click", e => {
  const t = e.target.closest("button"); if (!t) return;
  const s = site(), dsx = t.dataset;
  if (dsx.rsort) { const o = UI.rSort || {}; UI.rSort = { k: dsx.rsort, d: o.k === dsx.rsort ? -(o.d || 1) : 1 }; saveUI(); renderMaster(); return; }
  if (t.id === "mRsClr") { UI.rSort = null; saveUI(); renderMaster(); return; }
  if (t.id === "mRfClr") { UI.rFilt = null; saveUI(); renderMaster(); return; }
  if ((dsx.delw != null || dsx.delm != null || t.id === "mAddW" || t.id === "mAddM" || dsx.delsite === UI.site) && !canEdit()) return;
  if (dsx.pickgen) { UI.gen = dsx.pickgen; UI.site = null; renderAll(); return; }
  if (dsx.picksite) { UI.site = dsx.picksite; renderAll(); return; }
  if (dsx.delgen) { if (DB.現場.some(x => x.元請id === dsx.delgen)) { toast("現場がある元請業者は消せません（先に現場を消す）"); return; } if (!confirm("この元請業者を消しますか？")) return; DB.元請 = DB.元請.filter(x => x.id !== dsx.delgen); save(); renderAll(); return; }
  if (dsx.delsite) { const x = DB.現場.find(y => y.id === dsx.delsite); if (!confirm(`「${x.名前}」と日報 ${Object.keys(x.日).length} 日分を消しますか？\n共有している日報（OneDrive）も消えます（OneDrive のごみ箱から戻せます）。\n（消す前に「保存」でファイルに出しておくと安心です）`)) return; DB.現場 = DB.現場.filter(y => y.id !== dsx.delsite); removeShared(x.id); save(); renderAll(); return; }
  if (dsx.delw != null) { s.作業員.splice(+dsx.delw, 1); s.設定更新 = now(); save(); renderMaster(); return; }
  if (dsx.delm != null) { s.機材.splice(+dsx.delm, 1); s.設定更新 = now(); save(); renderMaster(); return; }
  if (t.id === "mAddW") { addWorkerDlg(s, {}); return; }
  if (t.id === "mAddM") { const n = $("mNewM").value.trim(); if (!n) return; s.機材.push({ 名前: splitOwner(n).名前, 規格: $("mNewMk").value.trim(), 所有者: splitOwner(n).所有者 }); s.設定更新 = now(); save(); renderMaster(); $("mNewM").focus(); return; }
  if (t.id === "mReload") { loadMasters(false); return; }
  if (t.id === "mYure") { findYure(); return; }
  if (dsx.kup != null || dsx.kdn != null) { const i = +(dsx.kup ?? dsx.kdn), j = dsx.kup != null ? i - 1 : i + 1; [DB.職種[i], DB.職種[j]] = [DB.職種[j], DB.職種[i]]; DB.共通更新 = now(); save(); renderMaster(); return; }
  if (dsx.kper != null) { openKPeriod(DB.職種[+dsx.kper]); return; }
  if (dsx.kdel != null) { const k = DB.職種[+dsx.kdel]; if (!confirm(`工種「${k}」を一覧から消しますか？（日報の中の「${k}」はそのまま残ります）`)) return; DB.職種.splice(+dsx.kdel, 1); delete DB.職種単価[k]; DB.共通更新 = now(); save(); renderMaster(); return; }
  if (t.id === "mAddK") {
    const k = $("mNewK").value.trim(); if (!k) { $("mNewK").focus(); return; }
    if (DB.職種.includes(k)) { toast("同じ工種がもうあります"); return; }
    DB.職種.push(k); const p = {}; ["昼", "夜", "残業"].forEach(f => { const v = han($("mNewK" + f).value).replace(/[,円]/g, ""); if (v !== "" && !isNaN(Number(v))) p[f] = v; }); DB.職種単価[k] = p;
    DB.共通更新 = now(); save(); renderMaster(); $("mNewK").focus(); return;
  }
  if (t.id === "mKFromDays") {
    let n = 0; DB.現場.forEach(x => Object.values(x.日).forEach(d => d.下.forEach(r => { if (r && r[0] && r[3] && !DB.職種.includes(r[3])) { DB.職種.push(r[3]); n++; } })));
    if (n) DB.共通更新 = now(); save(); renderMaster(); toast(n ? `${n} 件の工種を足しました` : "足りない工種はありません"); return;
  }
  if (["mCompHint", "mCompApply", "mCompCsvIn", "mKzFix", "mKzApply"].includes(t.id) && !canEdit()) return;
  if (["mPrFill"].includes(t.id) || dsx.prdel != null || dsx.padd != null || dsx.pcsvin != null) { if (!canEdit()) return; }
  if (dsx.padd != null) { const T = PT[+dsx.padd]; (s.単価 ||= []).push(T.区分 === "労務" ? { 区分: "労務", 名前: "", 昼: "", 夜: "", 残業: "", 開始: "", 終了: "", メモ: "" } : { 区分: T.区分, 分類: T.分類, 名前: "", 規格: "", 所有者: "", 単位: T.分類 === "高速代" ? "回" : T.区分 === "材料" ? "" : "台日", 単価: "", 開始: "", 終了: "", メモ: "" }); s.設定更新 = now(); save(); renderMaster(); return; }
  if (dsx.pcsvout) {
    const K = dsx.pcsvout, cols = PK_CSV[K], q = v => /[",\n]/.test(String(v ?? "")) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v ?? "");
    const field = c => c === "工種" || c === "品名" ? "名前" : c;
    const lines = [cols.join(",")].concat((s.単価 || []).filter(p => p.区分 === K).map(p => cols.map(c => q(p[field(c)] ?? "")).join(",")));
    download(`単価_${K}_${fname(s.名前)}.csv`, "\uFEFF" + lines.join("\r\n"), "text/csv"); return;
  }
  if (dsx.pcsvin) { PCSV = dsx.pcsvin; $("mPrFile").value = ""; $("mPrFile").click(); return; }
  if (["mPrTpl", "mPrOk", "mPrSetTpl"].includes(t.id) && !canEdit()) return;
  if (t.id === "mPrTpl") {
    const T = DB.単価たたき台; if (!T) return;
    s.単価 = migratePrices(T.行.map(p => ({ ...p }))).map(p => ({ ...p, メモ: (p.メモ ? p.メモ + "・" : "") + "たたき台から（要見直し）" }));
    s.単価確認 = false; s.設定更新 = now(); save(); renderMaster(); toast("たたき台を写しました。この現場の単価に合わせて直してください"); return;
  }
  if (t.id === "mPrOk") { if (!confirm("この現場の単価を見直しましたか？（「たたき台から」の案内を消します）")) return; s.単価確認 = true; s.単価 = (s.単価 || []).map(p => ({ ...p, メモ: String(p.メモ || "").replace(/・?たたき台から（要見直し）/, "") })); s.設定更新 = now(); save(); renderMaster(); return; }
  if (t.id === "mPrSetTpl") {
    if (!(s.単価 || []).length) { toast("この現場に単価がありません"); return; }
    if (!confirm(`この現場（${siteLabel(s)}）の単価 ${s.単価.length} 行を、ほかの現場のたたき台にします（今のたたき台と入れ替え）。よいですか？`)) return;
    DB.単価たたき台 = { 元: siteLabel(s), 行: s.単価.map(p => ({ ...p })), 更新: now() }; DB.共通更新 = now(); save(); renderMaster(); toast("たたき台にしました"); return;
  }
  if (t.id === "mPrFill") { const n = priceFill(s); if (n) s.設定更新 = now(); save(); renderMaster(); toast(n ? `日報に出てくる項目を ${n} 行足しました（単価を入れてください）` : "足りない項目はありません"); return; }
  if (dsx.prdel != null) { s.単価.splice(+dsx.prdel, 1); s.設定更新 = now(); save(); renderMaster(); return; }
  if (t.id === "mKzFix") {
    if (!confirm("使用機材名の頭の（ハ）（真）などを外して、所有者の欄に移します（この現場の日報すべてと機材の一覧）。よいですか？")) return;
    let n = 0;
    for (const d of Object.values(s.日)) { let ch = false; d.下.forEach(r => { if (r && r[5] && splitOwner(r[5]).所有者 && fixKizaiRow(r, s)) { n++; ch = true; } }); if (ch) d.更新 = now(); }
    const old = s.機材; s.機材 = [];
    old.forEach(m => { const sp = splitOwner(m.名前); const x = { 名前: sp.名前, 規格: m.規格 || "", 所有者: m.所有者 || sp.所有者 }; if (!s.機材.some(y => y.名前 === x.名前 && y.規格 === x.規格 && y.所有者 === x.所有者)) s.機材.push(x); });
    s.設定更新 = now(); UNDO = []; REDO = []; save(); refreshSheet(); renderMaster(); toast(`日報の ${n} か所を直しました`); return;
  }
  if (t.id === "mKzApply") {
    let n = 0;
    for (const d of Object.values(s.日)) { let ch = false; d.下.forEach(r => { if (r && r[5] && !r[13]) { const o = ownerOf(r[5], r[6], s); if (o) { while (r.length < 14) r.push(""); r[13] = o; n++; ch = true; } } }); if (ch) d.更新 = now(); }
    save(); refreshSheet(); renderMaster(); toast(n ? `日報の ${n} か所に所有者を入れました（空いている所だけ）` : "入れる所はありませんでした"); return;
  }
  if (t.id === "mCompHint") {
    let n = 0; s.作業員.forEach(w => { if (!w.会社) { const c = companyOf(w.名前, { 作業員: [] }); if (c) { w.会社 = c; n++; compToDays(s, w.名前, "", c); } } });
    if (n) s.設定更新 = now(); save(); renderMaster(); toast(n ? `${n}名に名簿（配置表）の会社を入れました` : "名簿で決まる人はいませんでした"); return;
  }
  if (t.id === "mCompApply") {
    const nAll = (() => { let k = 0; Object.values(s.日).forEach(d => d.下.forEach(r => { const w = r && r[0] && s.作業員.find(x => nk(x.名前) === nk(r[0])); if (w && w.会社 && r[12] !== w.会社) k++; })); return k; })();
    if (!nAll) { toast("日報の会社はもう一覧のとおりです"); return; }
    if (!confirm(`これまでの日報の会社の欄 ${nAll} か所を、この一覧の会社にします。よいですか？`)) return;
    const n = compApplyToDays(s, false); save(); refreshSheet(); renderMaster(); toast(`${n} か所に会社を入れました`); return;
  }
  if (t.id === "mCompCsvOut") {
    const st = rosterStats(s), q = v => /[",\n]/.test(v) ? '"' + String(v).replace(/"/g, '""') + '"' : v;
    const lines = [["名前", "会社", "工種", "日数", "最初", "最後"].join(",")].concat(s.作業員.map(w => { const x = st.get(nk(w.名前)); return [w.名前, w.会社 || "", w.職種 || "", x ? x.n : 0, x ? x.a : "", x ? x.b : ""].map(q).join(","); }));
    download(`名前一覧_${fname(s.名前)}.csv`, "\uFEFF" + lines.join("\r\n"), "text/csv"); return;
  }
  if (t.id === "mCompCsvIn") { $("mCompFile").value = ""; $("mCompFile").click(); return; }
  if (t.id === "yureFix") { fixYure(); return; }
  if (dsx.ygnot != null) { if (!canEdit()) return; const x = YG[+dsx.ygnot]; markDistinct(s, x.g.map(y => y.w)); save(); findYure(); toast("別のものとして覚えました（もう出しません）"); return; }
  if (t.id === "mSaveAll") { download(`常用日報_全部_${ymd(new Date()).replace(/-/g, "")}.json`, JSON.stringify({ ...DB, 保存日時: now() }, null, 1), "application/json"); }
});
// この現場の日報の中の揺らぎ：列ごとに、似ている書き方をまとめる（マスタの書き方・多く使われている書き方をはじめに選ぶ）
let YG = [];
function findYure() {
  const s = site(); if (!s) return;
  YG = [];
  for (const key of ["dn", "up"]) for (const c of Object.keys(YCOLS[key]).map(Number)) {
    const used = new Map();
    for (const dd of Object.values(s.日)) (dd[ARR[key]] || []).forEach(r => { const v = r && String(r[c] || "").trim(); if (v) used.set(v, (used.get(v) || 0) + 1); });
    const cn = canon(key, c, s), ws = [...used.keys()], seen = new Set();
    for (const w of ws) {
      if (seen.has(w)) continue;
      const g = [w, ...ws.filter(x => x !== w && !seen.has(x) && similar(w, x))];
      // マスタにあって日報に無い書き方も候補に（マスタの書き方にそろえられるように）
      [...cn.keys()].filter(x => cn.get(x) >= 1000 && !g.includes(x) && g.some(y => similar(y, x))).forEach(x => g.push(x));
      g.forEach(x => seen.add(x));
      if (g.length < 2) continue;
      // 日報で使っているのが1通りだけで、その書き方がマスタ（名簿）にあるなら出さない（比嘉健 と 名簿の「比嘉健　真創建」など）
      const usedN = g.filter(x => used.has(x)).length;
      if (usedN < 2 && g.filter(x => used.has(x)).every(x => (cn.get(x) || 0) >= 1000)) continue;
      const best = [...g].sort((a, b) => (cn.get(b) || 0) - (cn.get(a) || 0))[0];
      YG.push({ key, c, g: g.map(x => ({ w: x, n: used.get(x) || 0, m: (cn.get(x) || 0) >= 1000 })), best });
    }
  }
  $("yureList").innerHTML = !YG.length ? `<p class="note">揺らぎは見つかりませんでした。</p>` :
    `<table class="sm"><thead><tr><th>欄</th><th>書き方（日報の中の回数）</th><th>そろえる</th><th></th></tr></thead><tbody>` +
    YG.map((x, i) => `<tr><td>${esc(YCOLS[x.key][x.c])}</td><td>${x.g.map(y => `<label class="yr"><input type="radio" name="yg${i}" value="${esc(y.w)}"${y.w === x.best ? " checked" : ""}> ${esc(y.w)} <span class="note">${y.n}回${y.m ? "・マスタ" : ""}</span></label>`).join("")}</td>` +
      `<td class="c"><input type="checkbox" class="ygOn" data-i="${i}" checked></td><td class="btn"><button class="sub small" data-ygnot="${i}" title="同じものではない（兄弟・別の機械など）。覚えて、もう出さない">別のもの</button></td></tr>`).join("") +
    `</tbody></table><div class="btns"><button id="yureFix">チェックした行をそろえる</button><span class="note">別のもの（兄弟・別の機械など）は「別のもの」を押すと、覚えてもう出しません</span></div>`;
}
function fixYure() {
  const s = site(); if (!s || !canEdit()) return;
  let n = 0;
  const plan = [];
  document.querySelectorAll(".ygOn:checked").forEach(cb => {
    const x = YG[+cb.dataset.i], to = (document.querySelector(`input[name="yg${cb.dataset.i}"]:checked`) || {}).value;
    if (!to) return;
    plan.push({ ...x, to, from: new Set(x.g.map(y => y.w).filter(w => w !== to)) });
  });
  if (!plan.length) return;
  if (!confirm(`${plan.length} 件の揺らぎを、選んだ書き方にそろえます（この現場の日報すべて）。よいですか？`)) return;
  for (const [dt, dd] of Object.entries(s.日)) {
    let ch = false;
    for (const p of plan) (dd[ARR[p.key]] || []).forEach(r => { if (r && p.from.has(String(r[p.c] || "").trim())) { r[p.c] = p.to; n++; ch = true; } });
    if (ch) dd.更新 = now();
  }
  for (const p of plan) {
    if (p.key === "dn" && p.c === 0) { s.作業員 = s.作業員.filter(w => !p.from.has(w.名前)); if (!s.作業員.some(w => w.名前 === p.to)) s.作業員.push({ 名前: p.to, 職種: "" }); }
    if (p.key === "dn" && p.c === 5) s.機材.forEach(m => { if (p.from.has(m.名前)) m.名前 = p.to; });
    if (p.key === "dn" && p.c === 6) s.機材.forEach(m => { if (p.from.has(m.規格)) m.規格 = p.to; });
  }
  s.設定更新 = now(); UNDO = []; REDO = [];
  save(); refreshSheet(); findYure(); toast(`${n} か所をそろえました`);
}
// 名前一覧・会社の振り分け：この現場の名簿＋日報に出てくる名前（日数・最初〜最後）。会社はえらぶ（新しい会社も足せる）
function rosterStats(s) {
  const st = new Map();
  for (const [dt, d] of Object.entries(s.日)) d.下.forEach(r => {
    if (!r || !r[0]) return;
    const k = nk(r[0]), x = st.get(k) || { n: 0, a: dt, b: dt, 未: 0 };
    x.n++; if (dt < x.a) x.a = dt; if (dt > x.b) x.b = dt; if (!r[12]) x.未++; st.set(k, x);
  });
  return st;
}
function rosterHTML(s) {
  // 日報に出てくるのに名簿に無い人は名簿に足す（工種は多いほう）
  const st = rosterStats(s);
  for (const [k] of st) if (!s.作業員.some(w => nk(w.名前) === k)) {
    const kc = new Map(); let nm = "";
    Object.values(s.日).forEach(d => d.下.forEach(r => { if (r && nk(r[0]) === k) { nm = nm || r[0]; kc.set(r[3] || "", (kc.get(r[3] || "") || 0) + 1); } }));
    s.作業員.push({ 名前: masterName(nm), 職種: [...kc.entries()].sort((p, q) => q[1] - p[1])[0][0] });
  }
  const cs = companies(s), hint = w => { const c = companyOf(w.名前, { 作業員: [] }); return !w.会社 && c ? c : ""; };
  // 並べ方：見出しを押した列（もう一度押すと逆）。決めていなければ 会社 → 日数。追加した人は一番下（2026-10-31 ユーザー）
  const RS = UI.rSort || {}, RF = UI.rFilt || {};
  const keyOf = { 名前: p => nk(p.w.名前), 会社: p => p.w.会社 || "～", 工種: p => p.w.職種 || "～", 日数: p => p.x ? p.x.n : 0, 来た日: p => p.x ? p.x.a : "9999" };
  const cmp = (a, b) => typeof a === "number" ? a - b : String(a).localeCompare(String(b), "ja");
  const all = s.作業員.map((w, i) => ({ w, i, x: st.get(nk(w.名前)) }));
  const rows = all.filter(p => (!RF.会社 || (RF.会社 === "（未定）" ? !p.w.会社 : p.w.会社 === RF.会社)) && (!RF.工種 || (p.w.職種 || "（なし）") === RF.工種))
    .sort((p, q) => RS.k ? (RS.d || 1) * cmp(keyOf[RS.k](p), keyOf[RS.k](q)) : (!!p.w.追加 - !!q.w.追加) || (p.w.追加 && q.w.追加 ? p.w.追加.localeCompare(q.w.追加) : 0) || (p.w.会社 || "～").localeCompare(q.w.会社 || "～", "ja") || (q.x ? q.x.n : 0) - (p.x ? p.x.n : 0));
  const th = (k, t) => `<th><button type="button" class="thb${RS.k === k ? " on" : ""}" data-rsort="${k}" title="押すとこの列で並べ替え（もう一度押すと逆の順）">${t}${RS.k === k ? (RS.d === -1 ? " ▼" : " ▲") : ""}</button></th>`;
  const fOpt = (vals, cur) => `<option value="">（すべて）</option>` + vals.map(v => `<option${v === cur ? " selected" : ""}>${esc(v)}</option>`).join("");
  const fComp = [...new Set(all.map(p => p.w.会社 || "（未定）"))].sort((a, b) => a.localeCompare(b, "ja")), fKou = [...new Set(all.map(p => p.w.職種 || "（なし）"))].sort((a, b) => a.localeCompare(b, "ja"));
  const nNo = s.作業員.filter(w => !w.会社).length;
  const md = d => d ? `${+d.slice(5, 7)}/${+d.slice(8)}` : "";
  return `<div class="mbox"><h4>名前一覧・会社の振り分け（${s.作業員.length}名${nNo ? `・<span class="ng">会社未定 ${nNo}名</span>` : ""}）</h4>
    <p class="note">日報に出てくる名前の一覧です。会社を選ぶと、次からその人は日報でその会社で入ります。<br>「日報に入れる」で、これまでの日報の会社の欄もこの一覧のとおりにします。</p>
    <div class="btns"><button class="sub small" id="mCompHint">名簿でおすすめを入れる</button><button class="sub small" id="mCompApply">会社を日報に入れる</button><button class="sub small" id="mCompCsvOut">一覧を CSV で出す</button><button class="sub small" id="mCompCsvIn">CSV（名前,会社）を読む</button><input type="file" id="mCompFile" accept=".csv,text/csv" hidden></div>
    <div class="btns rfilt"><span class="note">絞り込み：</span><label>会社 <select id="mRfC">${fOpt(fComp, RF.会社)}</select></label><label>工種 <select id="mRfK">${fOpt(fKou, RF.工種)}</select></label>${RF.会社 || RF.工種 ? `<span class="note">${rows.length}名 / ${all.length}名</span><button class="sub small" id="mRfClr">絞り込みをやめる</button>` : ""}${RS.k ? `<button class="sub small" id="mRsClr">並べ替えを元に戻す</button>` : ""}</div>
    <table class="ed roster"><tr>${th("名前", "作業員名")}${th("会社", "会社")}${th("工種", "工種（職種）")}${th("日数", "日報の日数")}${th("来た日", "来ていた日")}<th></th></tr>` +
    rows.map(({ w, i, x }) => `<tr class="${w.会社 ? "" : "nocomp"}"><td><input data-wn="${i}" value="${esc(w.名前)}"></td>` +
      `<td><select data-wc="${i}"><option value="">（未定）</option>${cs.map(c => `<option${c === w.会社 ? " selected" : ""}>${esc(c)}</option>`).join("")}${w.会社 && !cs.includes(w.会社) ? `<option selected>${esc(w.会社)}</option>` : ""}<option value="__new">＋ 新しい会社…</option></select>${hint(w) ? `<span class="note"> 名簿では ${esc(hint(w))}</span>` : ""}</td>` +
      `<td><input data-wk="${i}" value="${esc(w.職種 || "")}" list="dlShoku" autocomplete="off"></td><td class="r">${x ? x.n + "日" : ""}</td><td class="note">${x ? md(x.a) + "〜" + md(x.b) : ""}</td>` +
      `<td class="btn"><button class="sub small danger" data-delw="${i}">消す</button></td></tr>`).join("") +
    `</table><div class="btns"><button class="small" id="mAddW">＋ 追加（名前・会社・工種）</button></div>
    <datalist id="dlShoku">${DB.職種.map(k => `<option value="${esc(k)}">`).join("")}</datalist></div>`;
}
// 名前一覧に足す（2026-10-31 ユーザー：追加はフォームで 名前・会社・工種。足した人は一覧の一番下）
function addWorkerDlg(s, p) {
  const cs = companies(s), ks = [...new Set([...siteKou(s), ...DB.職種])];
  editDlg("名前一覧に足す（" + siteLabel(s) + "）", [["名前", "name", p.name || ""], ["会社", "comp", p.comp || "", null, null, cs], ["工種（職種）", "kou", p.kou || "", null, null, ks]], v => {
    const n = masterName(v.name.trim()); if (!n) { toast("名前を入れてください"); addWorkerDlg(s, v); return; }
    if (s.作業員.some(w => nk(w.名前) === nk(n))) { toast(`「${n}」はもう一覧にあります`); addWorkerDlg(s, v); return; }
    const kou = DB.職種別名[v.kou] || v.kou;
    s.作業員.push({ 名前: n, 会社: v.comp, 職種: kou, 追加: now() });
    if (v.comp && !(DB.会社一覧 ||= []).some(k => loose(k) === loose(v.comp)) && v.comp !== SELF) { DB.会社一覧.push(v.comp); DB.共通更新 = now(); }
    if (v.comp) compToDays(s, n, "", v.comp);
    const newK = kou && !siteKou(s).some(k => loose(k) === loose(kou));
    if (newK) addSiteKou(s, kou, true);
    UI.rSort = null; UI.rFilt = null; saveUI();
    s.設定更新 = now(); save(); renderMaster(); toast(`「${n}」を一覧の一番下に足しました`);
    if (newK) setTimeout(() => askKouPrice(s, [kou]), 120);
  });
}
// 単価の見直しの案内：単価が無い現場は「たたき台を写す」、たたき台から写したまま（単価確認＝false）なら「見直して確認した」を押すまで出す
function priceNotice(s) {
  const T = DB.単価たたき台, P = s.単価 || [];
  if (!P.length && T && T.行 && T.行.length) return `<div class="pnote">この現場の単価がまだありません。<b>${esc(T.元 || "ほかの現場")}</b>の単価（${T.行.length} 行）を<b>たたき台</b>として写し、この現場の単価に合わせて直してください。<button class="small" id="mPrTpl">たたき台を写す</button></div>`;
  if (!P.length) return `<div class="pnote">この現場の単価がまだありません。行を足すか、CSV を読んでください。</div>`;
  if (s.単価確認 === false) return `<div class="pnote">この単価は<b>たたき台</b>から写したものです。この現場（元請との取り決め）の単価に合わせて見直し、直し終わったら押してください。<button class="small" id="mPrOk">見直した（この現場の単価にする）</button></div>`;
  return "";
}
// 工種を足す：共通の一覧に無ければ足し、この現場の一覧を決めていればそこにも足す
function addSiteKou(s, v, quiet) {
  v = String(v || "").trim(); if (!v) return false;
  v = DB.職種別名[v] || v;
  if (!DB.職種.some(k => loose(k) === loose(v))) { DB.職種.push(v); DB.共通更新 = now(); }
  if (s.職種 && s.職種.length && !s.職種.some(k => loose(k) === loose(v))) { s.職種.push(v); s.設定更新 = now(); }
  if (!quiet) save();
  return true;
}
// 工種を足したら、この現場の単価（昼・夜・残業）を聞く（2026-10-31 ユーザー：職種を足すときに単価も足すよう促す）
function askKouPrice(s, ks, then) {
  ks = ks.filter(k => !(s.単価 || []).some(p => p.区分 === "労務" && nk(p.名前) === nk(k) && (numOk(p.昼) || numOk(p.夜))));
  if (!ks.length) { if (then) then(); return; }
  const k = ks[0], c = DB.職種単価[k] || {};
  editDlg(`工種「${k}」の単価（${siteLabel(s)}）`, [["昼（円/人日）", "昼", c.昼 || ""], ["夜（円/人日。空なら昼と同じ）", "夜", c.夜 || ""], ["残業（円/時間。空なら（共通）の行）", "残業", c.残業 || ""]], v => {
    const num = x => han(x).replace(/[,，円]/g, "");
    const r = { 昼: num(v.昼), 夜: num(v.夜), 残業: num(v.残業) };
    if (Object.values(r).some(x => x !== "" && isNaN(Number(x)))) { toast("単価は数字で入れてください"); askKouPrice(s, ks, then); return; }
    if (Object.values(r).some(x => x !== "")) {
      (s.単価 ||= []).push({ 区分: "労務", 名前: k, ...r, 開始: "", 終了: "", メモ: "" });
      s.設定更新 = now(); save(); toast(`「${k}」の単価を単価マスタ（労務）に入れました`);
      if (!$("pMaster").hidden) renderMaster();
    } else toast(`「${k}」の単価は入れていません（現場・名簿 → 単価マスタ で入れられます）`, 4000);
    setTimeout(() => askKouPrice(s, ks.slice(1), then), 120);
  });
}
// 手入力した使用機材・材料を一覧に足したら、この現場の単価を聞く（2026-10-31 ユーザー：単価、材料は領収書精算も）
function askItemPrice(s, items) {
  const it = items[0]; if (!it) return;
  const rest = () => { if (items.length > 1) setTimeout(() => askItemPrice(s, items.slice(1)), 120); if (!$("pMaster").hidden) renderMaster(); };
  if ((s.単価 || []).some(p => ["機械", "材料", "通勤費"].includes(p.区分) && loose(p.名前) === loose(it.名前))) { rest(); return; }
  const num = x => han(x || "").replace(/[,，円\s]/g, "");
  const add = (row, msg) => { (s.単価 ||= []).push({ 規格: "", 所有者: "", 単位: "", 単価: "", 開始: "", 終了: "", メモ: "", ...row }); s.設定更新 = now(); save(); toast(msg, 4000); rest(); };
  const bad = u => { if (u === "") { toast(`「${it.名前}」の単価は入れていません（現場・名簿 → 単価マスタで入れられます）`, 4000); rest(); return true; } if (isNaN(Number(u))) { toast("単価は数字で入れてください"); askItemPrice(s, items); return true; } return false; };
  if (it.種 === "機械") {
    editDlg(`使用機材「${it.名前}」の単価（${siteLabel(s)}）`, [["分類", "分類", guessMach(it.名前), ["車両", "重機", "小型機械", "電動工具", "通勤車"]], ["単価（円/台日）", "単価", ""], ["単位", "単位", "台日"]], v => {
      const u = num(v.単価); if (bad(u)) return;
      add({ 区分: v.分類 === "通勤車" ? "通勤費" : "機械", 分類: v.分類, 名前: it.名前, 単位: v.単位 || "台日", 単価: u }, `「${it.名前}」を単価マスタ（${v.分類}）に足しました`);
    });
    return;
  }
  const hw = /高速/.test(it.名前), area = String(it.名前).replace(/^\s*高速(道路)?(代|料金)?\s*[（(]?/, "").replace(/[）)]\s*$/, "");
  editDlg(`材料「${it.名前}」の単価（${siteLabel(s)}）`, [["分類", "分類", hw ? "高速代" : "固定単価", hw ? ["高速代", "固定単価", "領収書精算"] : ["固定単価", "領収書精算"]], ["単価（円。領収書精算なら空）", "単価", ""], ["単位", "単位", hw ? "回" : ""]], v => {
    if (v.分類 === "領収書精算") { add({ 区分: "材料", 分類: "領収書精算", 名前: it.名前 }, `「${it.名前}」を領収書精算の材料に足しました。日報の「実費」に金額を入れます`); return; }
    const u = num(v.単価); if (bad(u)) return;
    if (v.分類 === "高速代") add({ 区分: "通勤費", 分類: "高速代", 名前: area, 単位: v.単位 || "回", 単価: u }, `高速代${area ? "（" + area + "）" : ""}を単価マスタ（通勤費）に足しました`);
    else add({ 区分: "材料", 分類: "固定単価", 名前: it.名前, 単位: v.単位, 単価: u }, `「${it.名前}」を単価マスタ（材料・固定単価）に足しました`);
  });
}
function siteKouHTML(s) {
  const used = new Map();
  Object.values(s.日).forEach(d => d.下.forEach(r => { if (r && r[0] && r[3]) used.set(r[3], (used.get(r[3]) || 0) + 1); }));
  const set = !!(s.職種 && s.職種.length), list = siteKou(s);
  const pr = k => { const p = laborPrice(s, k, "昼", ymd(new Date()), "日"), o = laborPrice(s, k, "昼", ymd(new Date()), "残業"); return (p ? yen(p.単価) : '<span class="ng">単価なし</span>') + (o ? `<span class="note">　残業 ${yen(o.単価)}/時</span>` : ""); };
  const out = [...used.keys()].filter(k => !list.some(x => loose(x) === loose(k)));
  return `<h3 class="sh">この現場の工種（職種）（${esc(siteLabel(s))}）</h3><div class="mbox">
    <p class="note">日報の工種の選択肢に出す工種です。${set ? "この現場で決めた一覧です。" : "<b>まだ決めていないので、共通の一覧（全部の現場）を出しています。</b>"}足すと、この現場の単価（昼・夜・残業）を聞きます。</p>
    <table class="ed"><tr><th>工種</th><th>日報に出てくる数</th><th>単価（昼）</th><th></th></tr>` +
    list.map(k => `<tr><td>${esc(k)}</td><td class="r">${used.get(k) || ""}</td><td>${pr(k)}</td><td class="btn">${set ? `<button class="sub small danger" data-skdel="${esc(k)}">外す</button>` : ""}</td></tr>`).join("") +
    `<tr><td><input id="mSkNew" list="mSkList" placeholder="工種を足す" autocomplete="off"><datalist id="mSkList">${DB.職種.filter(k => !list.includes(k)).map(k => `<option value="${esc(k)}">`).join("")}</datalist></td><td></td><td></td><td class="btn"><button class="small" id="mSkAdd">追加</button></td></tr></table>` +
    (out.length ? `<p class="note ng">一覧に無い工種が日報にあります：${out.map(k => `<button class="sub small" data-skadd="${esc(k)}" title="この現場の一覧に足す">${esc(k)}（${used.get(k)}）を足す</button>`).join(" ")}</p>` : "") +
    `<div class="btns">${set ? `<button class="sub small" id="mSkAll">決めるのをやめる（共通の一覧に戻す）</button>` : `<button class="sub small" id="mSkInit">この現場の工種を決める（日報で使っている工種から）</button>`}</div></div>`;
}
function priceHTML(s) {
  const P = s.単価 || [];
  const h = [`<h3 class="sh">単価マスタ（${esc(siteLabel(s))}）</h3><div class="mbox">${priceNotice(s)}<p class="note">集計の金額に使う単価です（この現場だけ。共有されます）。表ごとに分けています。規格・所有者が空の行は、どの規格・所有者にも使います。<br>単価が期間で変わるときは、同じ項目の行を足して「開始」「終了」に期間を入れます（空ならその側は限りなし）。</p>
    <div class="btns"><button class="sub small" id="mPrFill">日報から項目を足す</button><button class="sub small" id="mPrSetTpl" title="この現場の単価を、ほかの現場のたたき台にする（全部の現場で共通）">この単価をたたき台にする</button><input type="file" id="mPrFile" accept=".csv,text/csv" hidden></div>`];
  for (const K of PK) {
    h.push(`<div class="pgrp"><div class="pgh"><b>${K}</b><span class="sp"></span><button class="sub small" data-pcsvout="${K}">${K}の CSV を出す</button><button class="sub small" data-pcsvin="${K}">${K}の CSV を読む</button></div>`);
    PT.forEach((T, ti) => {
      if (T.区分 !== K) return;
      const rows = P.map((p, i) => ({ p, i })).filter(({ p }) => p.区分 === K && (K === "労務" || p.分類 === T.分類))
        .sort((a, b) => String(a.p.名前).localeCompare(String(b.p.名前), "ja") || (a.p.開始 || "").localeCompare(b.p.開始 || ""));
      const empty = p => K === "労務" ? !numOk(p.昼) && !numOk(p.夜) && !numOk(p.残業) : T.分類 !== "領収書精算" && !numOk(p.単価);
      h.push(`<h4 class="mh">${esc(T.題)}　<span class="note">${rows.length} 行</span></h4>${T.note ? `<p class="note">${esc(T.note)}</p>` : ""}<table class="ed price"><tr>${T.cols.map(c => `<th>${esc(c[1] || c[0])}</th>`).join("")}<th></th></tr>` +
        rows.map(({ p, i }) => `<tr class="${empty(p) ? "nocomp" : ""}">` + T.cols.map(([f, , t]) => t === "d" ? `<td><input type="date" data-pi="${i}" data-pf="${f}" value="${esc(p[f] || "")}"></td>`
          : `<td><input data-pi="${i}" data-pf="${f}" value="${esc(p[f] ?? "")}"${t === "n" ? ' inputmode="decimal" class="r"' : ""} autocomplete="off"></td>`).join("") +
          `<td class="btn"><button class="sub small danger" data-prdel="${i}">消す</button></td></tr>`).join("") +
        `</table><div class="btns"><button class="sub small" data-padd="${ti}">＋ ${esc(T.題)} に行を足す</button></div>`);
    });
    h.push(`</div>`);
  }
  h.push(`</div>`);
  return h.join("");
}
function priceFill(s) {
  const P = s.単価 ||= [];
  let n = 0;
  const hasL = k => P.some(p => p.区分 === "労務" && nk(p.名前) === nk(k));
  const hasI = (K, nm, sp, o) => P.some(p => p.区分 === K && nk(p.名前) === nk(nm) && (nk(p.規格) === nk(sp) || !p.規格) && (nk(p.所有者) === nk(o) || !p.所有者));
  Object.values(s.日).forEach(d => d.下.forEach(r => {
    if (!r) return;
    if (r[0] && r[3] && !hasL(r[3])) { P.push({ 区分: "労務", 名前: r[3], 昼: "", 夜: "", 残業: "", 開始: "", 終了: "", メモ: "" }); n++; }
    if (r[5]) { const sp = splitOwner(r[5]), g = guessMach(sp.名前), K = g === "通勤車" ? "通勤費" : "機械", o = r[13] || sp.所有者 || ""; if (!hasI(K, sp.名前, r[6] || "", o)) { P.push({ 区分: K, 分類: g, 名前: sp.名前, 規格: r[6] || "", 所有者: o, 単位: "台日", 単価: "", 開始: "", 終了: "", メモ: "" }); n++; } }
    if (r[9] && !itemPrice(s, ["材料", "通勤費"], r[9], "", "", "9999-12-31")) { const hw = /高速/.test(r[9]), K = hw ? "通勤費" : "材料"; if (!P.some(p => p.区分 === K && nk(p.名前) === nk(r[9]))) { P.push({ 区分: K, 分類: hw ? "高速代" : "固定単価", 名前: r[9], 規格: "", 所有者: "", 単位: hw ? "回" : "", 単価: "", 開始: "", 終了: "", メモ: "" }); n++; } }
  }));
  return n;
}
// 工種ごとの期間の単価の小窓
let KP = null;
function openKPeriod(k) {
  KP = k; renderKPeriod(); openDlg($("dKp"));
}
function renderKPeriod() {
  const p = DB.職種単価[KP] ||= {}, L = p.期間 ||= [];
  $("kpTtl").textContent = `「${KP}」の期間ごとの単価`;
  $("kpBody").innerHTML = `<p class="note">期間（開始〜終了）に入る日はこの単価を使います。入らない日は いつもの単価（昼 ${p.昼 ? yen(p.昼) : "－"}・夜 ${p.夜 ? yen(p.夜) : "－"}・残業 ${p.残業 ? yen(p.残業) : "－"}）。開始・終了が空ならその側は限りなし。期間が重なったら開始の新しいほう。</p>
    <table class="ed shoku"><tr><th>開始</th><th>終了</th><th>昼（円/日）</th><th>夜（円/日）</th><th>残業（円/時間）</th><th></th></tr>` +
    L.map((x, j) => `<tr><td><input type="date" data-kpj="${j}" data-kpf="開始" value="${esc(x.開始 || "")}"></td><td><input type="date" data-kpj="${j}" data-kpf="終了" value="${esc(x.終了 || "")}"></td>` +
      ["昼", "夜", "残業"].map(f => `<td><input class="r" data-kpj="${j}" data-kpf="${f}" value="${esc(x[f] ?? "")}" inputmode="decimal" autocomplete="off"></td>`).join("") +
      `<td class="btn"><button type="button" class="sub small danger" data-kpdel="${j}">消す</button></td></tr>`).join("") +
    `</table><div class="btns"><button type="button" class="sub small" id="kpAdd">＋ 期間を足す</button><span class="sp"></span><button type="button" data-close>閉じる</button></div>`;
}
$("kpBody").addEventListener("change", e => {
  const t = e.target, j = t.dataset.kpj; if (j == null) return;
  const L = DB.職種単価[KP].期間, f = t.dataset.kpf; let v = t.value.trim();
  if (!["開始", "終了"].includes(f)) { v = han(v).replace(/[,円]/g, ""); if (v !== "" && isNaN(Number(v))) { toast("単価は数字で入れてください"); renderKPeriod(); return; } }
  L[+j][f] = v;
  if (L[+j].開始 && L[+j].終了 && L[+j].終了 < L[+j].開始) { toast("終了が開始より前です"); }
  DB.共通更新 = now(); save();
});
$("kpBody").addEventListener("click", e => {
  const t = e.target.closest("button"); if (!t) return;
  const L = DB.職種単価[KP].期間;
  if (t.id === "kpAdd") { const last = L[L.length - 1]; L.push({ 開始: last && last.終了 ? addD(last.終了, 1) : "", 終了: "", 昼: "", 夜: "", 残業: "" }); DB.共通更新 = now(); save(); renderKPeriod(); return; }
  if (t.dataset.kpdel != null) { L.splice(+t.dataset.kpdel, 1); DB.共通更新 = now(); save(); renderKPeriod(); return; }
});
$("dKp").addEventListener("close", () => { if (UI.tab === "master") renderMaster(); if (UI.tab === "sum") renderSum(); });
function compApplyToDays(s, onlyEmpty) {
  let n = 0;
  for (const d of Object.values(s.日)) {
    let ch = false;
    d.下.forEach(r => {
      if (!r || !r[0]) return;
      const w = s.作業員.find(x => nk(x.名前) === nk(r[0]));
      if (w && w.会社 && r[12] !== w.会社 && (!onlyEmpty || !r[12])) { while (r.length < 13) r.push(""); r[12] = w.会社; n++; ch = true; }
    });
    if (ch) d.更新 = now();
  }
  return n;
}
function parseCsv(t) {
  return t.replace(/^﻿/, "").split(/\r?\n/).map(l => { const out = []; let cur = "", q = false; for (let i = 0; i < l.length; i++) { const ch = l[i]; if (q) { if (ch === '"' && l[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; } else if (ch === '"') q = true; else if (ch === ",") { out.push(cur); cur = ""; } else cur += ch; } out.push(cur); return out.map(x => x.trim()); }).filter(r => r.some(Boolean));
}
function editDlg(title, fields, ok) {
  $("eTtl").textContent = title;
  $("eBody").innerHTML = `<div class="grid">${fields.map(([l, k, v, opts, ty, cand]) => `<label>${esc(l)}</label>` + (opts ? `<select data-k="${k}" style="width:320px">${opts.map(o => `<option${o === v ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>` : `<input data-k="${k}" value="${esc(v)}" style="width:320px"${ty ? ` type="${ty}"` : ""}${cand ? ` list="dl_${k}" autocomplete="off"><datalist id="dl_${k}">${cand.map(o => `<option value="${esc(o)}">`).join("")}</datalist` : ""}>`)).join("")}</div>
    <div class="btns"><span class="sp"></span><button type="button" class="sub" data-close>やめる</button><button type="button" id="eOk">決める</button></div>`;
  openDlg($("dEdit"));
  $("eBody").querySelector("input,select").focus();
  const go = () => { const v = {}; $("eBody").querySelectorAll("[data-k]").forEach(i => v[i.dataset.k] = i.value.trim()); closeDlg($("dEdit")); ok(v); };
  $("eOk").onclick = go;
  $("eBody").onkeydown = e => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); go(); } };
}

/* ================= 保存・取込 ================= */
function download(name, data, type) {
  const b = data instanceof Blob ? data : new Blob([data], { type });
  const a = document.createElement("a"); a.href = URL.createObjectURL(b); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
const fname = s => String(s).replace(/[\\/:*?"<>|]/g, "_").slice(0, 40);
$("bSaveFile").onclick = () => {
  const s = site(); if (!s) { toast("現場を選んでください"); return; }
  download(`常用日報_${fname(s.名前)}_${ymd(new Date()).replace(/-/g, "")}.json`, JSON.stringify({ 形式: "常用日報", 版: 1, 元請: [gen()], 現場: [s], 職種: DB.職種, 天気: DB.天気, 保存日時: now() }, null, 1), "application/json");
  toast("ファイルに保存しました（ダウンロード）");
};
$("bLoadFile").onclick = () => { $("fileIn").value = ""; $("fileIn").click(); };
$("fileIn").onchange = async () => {
  const f = $("fileIn").files[0]; if (!f) return;
  try {
    if (/\.xlsx?$/i.test(f.name)) await importXlsx(await f.arrayBuffer());
    else importJSON(JSON.parse(await f.text()));
  } catch (e) { console.error(e); toast("読めませんでした：" + e.message); }
};
function importJSON(j) {
  if (j.形式 !== "常用日報") { toast("常用日報のファイルではありません"); return; }
  let nNew = 0, nDay = 0, nOver = 0;
  for (const g of j.元請 || []) if (g && !DB.元請.some(x => x.id === g.id)) { const same = DB.元請.find(x => x.名前 === g.名前); if (same) (j.現場 || []).forEach(s => { if (s.元請id === g.id) s.元請id = same.id; }); else DB.元請.push(g); }
  const plans = [];
  for (const s of j.現場 || []) {
    fixSite(s);
    const mine = DB.現場.find(x => x.id === s.id);
    if (!mine) { plans.push(() => { DB.現場.push(s); nNew++; nDay += Object.keys(s.日).length; }); continue; }
    const over = Object.keys(s.日).filter(d => mine.日[d] && JSON.stringify(mine.日[d]) !== JSON.stringify(s.日[d]));
    nOver += over.length;
    plans.push(() => {
      for (const [d, v] of Object.entries(s.日)) { mine.日[d] = v; nDay++; }
      for (const w of s.作業員) if (!mine.作業員.some(x => x.名前 === w.名前)) mine.作業員.push(w);
      for (const m of s.機材) if (!mine.機材.some(x => x.名前 === m.名前 && x.規格 === m.規格)) mine.機材.push(m);
    });
  }
  if (nOver && !confirm(`このブラウザの日報と内容のちがう日が ${nOver} 日あります。ファイルの内容で上書きしますか？`)) return;
  plans.forEach(p => p());
  for (const k of j.職種 || []) if (!DB.職種.includes(k)) DB.職種.push(k);
  if (j.現場 && j.現場[0]) { UI.site = j.現場[0].id; UI.gen = (DB.現場.find(x => x.id === UI.site) || {}).元請id; }
  save(); renderAll(); toast(`取り込みました（日報 ${nDay} 日分${nNew ? "・新しい現場 " + nNew : ""}）`);
}
// 古い Excel（.xls）は ExcelJS が読めないので、SheetJS で .xlsx に直してから同じ読み方で読む（2026-10-31 ユーザー：前の期間の作業日報が .xls）
// ファイルの頭が D0 CF 11 E0（古い Office の形）なら .xls
function xlsToXlsx(buf) {
  const h = new Uint8Array(buf, 0, 4);
  if (!(h[0] === 0xD0 && h[1] === 0xCF && h[2] === 0x11 && h[3] === 0xE0)) return buf;
  if (typeof XLSX === "undefined") throw new Error("古い Excel（.xls）を読む部品が読めていません");
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true, cellFormula: false });
  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}
// 元の作業日報（Excel・1日1シート）を読む。番地は元の作業日報（B3 元請・B4 現場・G3 日付・18行目〜作業員）の形
async function importXlsx(buf) {
  buf = xlsToXlsx(buf);
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
  const tx = v => {
    if (v == null) return "";
    if (v instanceof Date) return v;
    if (typeof v === "object") { if (v.richText) return v.richText.map(x => x.text).join(""); if ("result" in v) return tx(v.result); if (v.text) return String(v.text); return ""; }
    return String(v).trim();
  };
  const got = [];
  wb.eachSheet(ws => {
    const g = a => tx(ws.getCell(a).value);
    let dt = g("G3");
    if (!(dt instanceof Date)) { const m = String(dt || ws.name).match(/(\d{4})[.\/年-](\d{1,2})[.\/月-](\d{1,2})/); if (!m) return; dt = new Date(+m[1], +m[2] - 1, +m[3]); }
    else dt = new Date(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate());
    const d = { 天気: g("J4"), 気温: normNum(String(g("K4")).replace(/℃/g, "")), 担当者: masterName(g("M3")), 上: [], 下: [], 備考: "", サイン: "", サイン日時: "", 人員手: true, 更新: now() };
    for (let i = 0; i < UP_ROWS; i++) { const r = 7 + i * 2; d.上.push([g("A" + r), g("B" + r), normNin(g("D" + r)), g("E" + r)]); }
    for (let r = 18; r <= Math.max(33, ws.rowCount); r++) {
      if (String(g("A" + r)).includes("作業員名")) break;
      const t = c => { const v = ws.getCell(c + r).value; return v instanceof Date ? pad(v.getUTCHours()) + ":" + pad(v.getUTCMinutes()) : tx(v); };
      d.下.push([g("A" + r), normTime(t("C")), normTime(t("D")), g("E" + r), normNum(g("F" + r)), g("G" + r), g("I" + r), normNum(g("J" + r)), "", g("L" + r), normNum(g("M" + r)), ""]);
    }
    d.備考 = ["L8", "L10", "L12", "L14", "L15"].map(g).filter(Boolean).join("\n");
    tidy(d);
    got.push({ date: ymd(dt), gen: g("B3"), site: g("B4"), d });
  });
  if (!got.length) { toast("作業日報のシートが見つかりませんでした"); return; }
  got.sort((p, q) => p.date.localeCompare(q.date));
  openXlsDlg(got);
}
/* ---------- Excel の日を選んで読む（2026-10-31 ユーザー：Excel の特定の日から内容を読み込める） ----------
   読み先：「Excel の日付のまま」＝選んだ日をその日に入れる（現場は Excel の現場名から）
           「いま開いている日へ」＝選んだ1日の内容を、いま開いている現場・日に写す（天気・気温・備考・サインは写さない） */
let XLS = null;
function openXlsDlg(got) {
  XLS = got;
  const a = got[0].date, b = got[got.length - 1].date;
  $("xlsInfo").textContent = `Excel の中に ${got.length} 日分（${jpDate(a)}〜${jpDate(b)}）。現場名：${got[0].site || "（なし）"}`;
  $("xlsFrom").value = a; $("xlsTo").value = b;
  $("xlsCurDay").textContent = site() ? `${siteLabel(site())}　${jpDate(UI.day)}（${WD[wd(UI.day)]}）` : "（現場を選んでいません）";
  $("xlsModeCur").disabled = !site();
  $("xlsModeSame").checked = true; $("xlsRange").hidden = false;
  renderXlsList(true);
  openDlg($("dXls"));
}
function xlsTarget() {
  // Excel の日付のままのとき、入れる先の現場（マスタの現場一覧 → 元請と名前 → 無ければ読むときに作る）
  const sName = XLS[0].site || "（現場）";
  const mm = MS.現場.find(m => siteMatch(sName, m));
  let s = mm && (DB.現場.find(x => x.コード === mm.コード) || (syncMasterSites(), DB.現場.find(x => x.コード === mm.コード)));
  if (!s) { const g = DB.元請.find(x => nk(x.名前) === nk(XLS[0].gen)); s = g && DB.現場.find(x => x.元請id === g.id && nk(x.名前) === nk(sName)); }
  return s || null;
}
function renderXlsList(reset) {
  const cur = $("xlsModeCur").checked, s = cur ? site() : xlsTarget();
  const a = $("xlsFrom").value, b = $("xlsTo").value;
  const keep = reset ? null : new Set([...document.querySelectorAll("#xlsList input:checked")].map(i => i.value));
  $("xlsList").innerHTML = `<table class="sm"><thead><tr><th></th><th>日付</th><th>曜</th><th>作業員</th><th>工種・作業内容</th><th>${cur ? "" : "このブラウザ"}</th></tr></thead><tbody>` +
    XLS.map((x, i) => {
      const n = x.d.下.filter(r => r && r[0]).length, w = wd(x.date);
      const on = keep ? keep.has(String(i)) : (cur ? false : x.date >= a && x.date <= b);
      const has = !cur && s && hasData(s.日[x.date]);
      return `<tr class="${w === 0 ? "sun" : w === 6 ? "sat" : ""}"><td><input type="${cur ? "radio" : "checkbox"}" name="xlsPick" value="${i}"${on ? " checked" : ""}></td><td>${+x.date.slice(5, 7)}/${+x.date.slice(8)}</td><td class="c">${WD[w]}</td><td class="r">${n}名</td>` +
        `<td>${esc(x.d.上.filter(r => r && (r[0] || r[1])).map(r => [r[0], r[1], r[2]].filter(Boolean).join(" ")).join("／").slice(0, 60))}</td><td class="c">${has ? '<span class="ng">書いてある（上書き）</span>' : ""}</td></tr>`;
    }).join("") + `</tbody></table>`;
  xlsCount();
}
function xlsCount() {
  const n = document.querySelectorAll("#xlsList input:checked").length;
  $("xlsOk").textContent = $("xlsModeCur").checked ? "この日に写す" : `${n} 日分を読む`;
  $("xlsOk").disabled = !n;
}
$("xlsList").addEventListener("change", xlsCount);
// 行のどこを押しても選べる（iPad で小さい四角を押さなくてよいように）
$("xlsList").addEventListener("click", e => {
  const tr = e.target.closest("tbody tr"); if (!tr || e.target.tagName === "INPUT") return;
  const i = tr.querySelector("input"); i.checked = i.type === "radio" ? true : !i.checked; xlsCount();
});
["xlsFrom", "xlsTo"].forEach(id => $(id).addEventListener("change", () => { if (!$("xlsModeCur").checked) renderXlsList(true); }));
["xlsModeSame", "xlsModeCur"].forEach(id => $(id).addEventListener("change", () => { $("xlsRange").hidden = $("xlsModeCur").checked; renderXlsList(true); }));
$("xlsAll").onclick = () => { document.querySelectorAll("#xlsList input[type=checkbox]").forEach(i => i.checked = true); xlsCount(); };
$("xlsNone").onclick = () => { document.querySelectorAll("#xlsList input[type=checkbox]").forEach(i => i.checked = false); xlsCount(); };
$("xlsOk").onclick = () => {
  const pick = [...document.querySelectorAll("#xlsList input:checked")].map(i => XLS[+i.value]);
  if (!pick.length) return;
  if ($("xlsModeCur").checked) xlsToCurrent(pick[0]); else xlsSameDays(pick);
};
// 名簿・工種・機材に足す（名前は名簿の書き方にそろえる）
function learn(s, d) {
  d.上.forEach(r => { if (r && DB.職種別名[r[0]]) r[0] = DB.職種別名[r[0]]; });
  d.下.forEach(r => {
    if (r && DB.職種別名[r[3]]) r[3] = DB.職種別名[r[3]];
    if (r[5]) { fixKizaiRow(r, s); rememberKizai(s, r); }
    if (!r[0]) return;
    r[0] = masterName(r[0]);
    if (!r[12]) r[12] = companyOf(r[0], s);
    if (!s.作業員.some(w => nk(w.名前) === nk(r[0]))) s.作業員.push({ 名前: r[0], 職種: r[3] });
    if (r[3] && !DB.職種.includes(r[3])) { DB.職種.push(r[3]); DB.共通更新 = now(); }
  });
}
function xlsToCurrent(x) {
  const s = site(); if (!s || !canEdit()) return;
  if (hasData(day()) && !confirm(`${jpDate(UI.day)} の内容を、Excel の ${jpDate(x.date)} の内容で置き換えますか？（↶ で戻せます）`)) return;
  pushUndo();
  const src = JSON.parse(JSON.stringify(x.d)), d = ensureDay();
  learn(s, src);
  d.上 = src.上; d.下 = src.下; d.人員手 = src.人員手; if (src.担当者) d.担当者 = src.担当者; d.更新 = now();
  closeDlg($("dXls")); save(); refreshSheet();
  toast(`Excel の ${jpDate(x.date)} の内容を ${jpDate(UI.day)} に写しました（天気・気温・備考・サインは写していません）`);
}
function xlsSameDays(pick) {
  const sName = XLS[0].site || "（現場）", gName = XLS[0].gen || "（元請業者）";
  let s = xlsTarget();
  if (!s) {
    let g = DB.元請.find(x => nk(x.名前) === nk(gName));
    if (!g) { g = { id: uid(), 名前: gName }; DB.元請.push(g); }
    s = fixSite({ id: uid(), 元請id: g.id, 名前: sName, 担当者: XLS[0].d.担当者 }); DB.現場.push(s);
  }
  if (s.id === UI.site && !canEdit()) return;
  if (!s.担当者) s.担当者 = masterName(XLS[0].d.担当者);
  const over = pick.filter(x => hasData(s.日[x.date])).length;
  if (over && !confirm(`${s.名前} に、もう書いてある日が ${over} 日あります。Excel の内容で上書きしますか？`)) return;
  for (const x of pick) { const d = JSON.parse(JSON.stringify(x.d)); learn(s, d); s.日[x.date] = d; }
  UI.gen = s.元請id; UI.site = s.id; UI.day = pick[pick.length - 1].date;
  UNDO = []; REDO = [];
  closeDlg($("dXls")); save(); renderAll();
  toast(`Excel から ${pick.length} 日分を読みました（${siteLabel(s)}）`);
}
$("bXlsIn").onclick = () => { $("xlsFile").value = ""; $("xlsFile").click(); };
$("xlsFile").onchange = async () => {
  const f = $("xlsFile").files[0]; if (!f) return;
  try { await importXlsx(await f.arrayBuffer()); } catch (e) { console.error(e); toast("読めませんでした：" + e.message); }
};

/* ================= Excel 出力（元の作業日報と同じ番地・A4 横・1シート1日） ================= */
const B_THIN = { style: "thin", color: { argb: "FF000000" } };
const BORDER = { top: B_THIN, left: B_THIN, bottom: B_THIN, right: B_THIN };
async function exportXlsx(dates, withSum, range) {
  const s = site(); if (!s) { toast("現場を選んでください"); return; }
  if (typeof ExcelJS === "undefined") { toast("Excel を作る部品が読めていません"); return; }
  const wb = new ExcelJS.Workbook(); wb.creator = "常用日報";
  for (const x of [...dates].sort().reverse()) daySheet(wb, s, x, s.日[x] || { 上: [], 下: [] });   // 元の Excel と同じく新しい日が左
  if (withSum && range) sumSheet(wb, s, range[0], range[1]);
  const buf = await wb.xlsx.writeBuffer();
  const nm = dates.length === 1 && !withSum ? `作業日報_${fname(s.名前)}_${dates[0].replace(/-/g, "")}.xlsx`
    : `作業日報_${fname(s.名前)}_${(range ? range[0] : dates[0]).replace(/-/g, "")}-${(range ? range[1] : dates[dates.length - 1]).replace(/-/g, "")}.xlsx`;
  download(nm, new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  toast("Excel を出しました：" + nm);
}
// 1日が何ページでも、ページごとに1シート（工種別の集計はその日の全部）
function daySheet(wb, s, x, d) { const n = pageCount(d); for (let p = 0; p < n; p++) daySheetPage(wb, s, x, pageOf(d, p), d, p, n); }
function daySheetPage(wb, s, x, d, full, pi, pn) {
  const dt = pd(x);
  const ws = wb.addWorksheet(`作業日報${dt.getFullYear()}.${dt.getMonth() + 1}.${dt.getDate()}${pn > 1 ? "-" + (pi + 1) : ""}`, {
    pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true,
      margins: { left: 0.55, right: 0.55, top: 0.6, bottom: 0.5, header: 0.3, footer: 0.3 } },
    views: [{ showGridLines: false }]
  });
  const W = [9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 4.63, 9, 9, 9];
  ws.columns = W.map(w => ({ width: w + 0.71 }));
  // 下の表と高速・燃料の行は合わせて 19 行を目安（高速・燃料が増えたら下の表の空の行を減らす）。A4 1枚に収める
  const nCar = PG.car, nDn = PG.dn, last = 17 + nDn + 1 + nCar;
  const font = { name: "ＭＳ Ｐゴシック", size: 11 };
  for (let r = 1; r <= last; r++) for (let c = 1; c <= 16; c++) { const cl = ws.getCell(r, c); cl.font = font; cl.alignment = { vertical: "middle", horizontal: "center", shrinkToFit: true }; }
  ws.getRow(1).height = 27;
  [3, 4, 5, 6].forEach(r => ws.getRow(r).height = 21);
  for (let r = 18; r <= last; r++) ws.getRow(r).height = 21;
  const put = (a, v, m, extra) => { if (m) ws.mergeCells(m); const cl = ws.getCell(a); cl.value = v; if (extra) Object.assign(cl, extra); return cl; };
  put("A1", "作　業　日　報" + (pn > 1 ? `（${pi + 1}／${pn}）` : ""), "A1:P1").font = { ...font, size: 20 };
  put("A3", "元請業者"); put("B3", gen() ? gen().名前 : "", "B3:D3");
  put("E3", "日　　付", "E3:F3"); put("G3", dt, "G3:H3").numFmt = 'yyyy"年"m"月"d"日"';
  put("I3", WD[dt.getDay()]); put("J3", "天気"); put("K3", "気温");
  put("L3", "記入者", "L3:L4"); put("M3", d.担当者 || "", "M3:P4");
  put("A4", "現場名"); put("B4", s.名前, "B4:I4"); put("J4", d.天気 || ""); put("K4", d.気温 !== "" && d.気温 != null ? d.気温 + "℃" : "℃");
  ws.getCell("K4").alignment = { vertical: "middle", horizontal: "right" };
  // 左：作業内容 5段（2行ずつ）／右：工種別の集計（2026-10-31 ユーザー：サイン・備考をやめた）
  put("A5", "作　業　予　定", "A5:F6"); put("G5", "完　了　報　告", "G5:K6"); put("L5", "工種別の集計", "L5:P6");
  for (let i = 0; i < UP_ROWS; i++) {
    const r = 7 + i * 2, row = d.上[i] || [];
    put("A" + r, row[1] || "", `A${r}:F${r + 1}`).alignment = { vertical: "middle", horizontal: "left", wrapText: true, indent: 1 };
    put("G" + r, row[4] || "", `G${r}:K${r + 1}`).alignment = { vertical: "middle", horizontal: "left", wrapText: true, indent: 1 };
  }
  const KS = kouSum(full), KR = KS.rows.length > 8 ? [...KS.rows.slice(0, 7), { k: "ほか", n: KS.rows.slice(7).reduce((a, x) => a + x.n, 0), 夜: KS.rows.slice(7).reduce((a, x) => a + x.夜, 0) }] : KS.rows;
  const ninTxt = x => x.n + "名" + (KS.夜あり && x.夜 ? `（夜${x.夜}）` : "");
  put("L7", "工種", "L7:N7"); put("O7", "人員", "O7:P7");
  for (let j = 0; j < 8; j++) { const r = 8 + j, x = KR[j]; put("L" + r, x ? x.k : "", `L${r}:N${r}`); put("O" + r, x ? ninTxt(x) : "", `O${r}:P${r}`).alignment = { vertical: "middle", horizontal: "right" }; }
  put("L16", "計", "L16:N16"); put("O16", KS.計 ? ninTxt({ n: KS.計, 夜: KS.rows.reduce((a, x) => a + x.夜, 0) }) : "", "O16:P16").alignment = { vertical: "middle", horizontal: "right" };
  const H = [["A17", "作業員名", "A17:B17"], ["C17", "始業時間"], ["D17", "終業時間"], ["E17", "工種"], ["F17", "残業時間"], ["G17", "使用機材", "G17:H17"], ["I17", "規格・形式"], ["J17", "数量"], ["K17", "材料名", "K17:M17"], ["N17", "数量"], ["O17", "実費", "O17:P17"]];
  H.forEach(([a, v, m]) => put(a, v, m));
  // 累計は画面と同じ計算（この日の現場・日を一時的に当てる）
  const keep = UI.day; UI.day = x;
  for (let i = 0; i < nDn; i++) {
    const r = 18 + i, row = d.下[i] || [];
    const num = v => v === "" || v == null ? "" : isNaN(Number(v)) ? v : Number(v);
    put("A" + r, row[0] || "", `A${r}:B${r}`); put("C" + r, row[0] ? row[1] || "" : ""); put("D" + r, row[0] ? row[2] || "" : "");
    put("E" + r, row[3] || ""); const ot = put("F" + r, num(row[4])); if (row[4]) ot.font = { ...font, color: { argb: "FFC00000" } };
    put("G" + r, row[5] ? (row[13] ? "（" + abbrOf(row[13]) + "）" : "") + row[5] : "", `G${r}:H${r}`); put("I" + r, row[6] || ""); put("J" + r, num(row[7]));
    put("K" + r, row[9] || "", `K${r}:M${r}`); put("N" + r, num(row[10])); const jc = put("O" + r, num(row[15]), `O${r}:P${r}`); if (row[15] !== "" && row[15] != null) jc.numFmt = "#,##0";
  }
  UI.day = keep;
  {
    const hr = 18 + nDn, numx = v => v === "" || v == null ? "" : isNaN(Number(v)) ? v : Number(v);
    [["A", "種類"], ["B", "往復"], ["C", "高速の区間", "E"], ["F", "給油(L)"], ["G", "実費"], ["H", "会社", "I"], ["J", "名前", "K"], ["L", "車両名", "M"], ["N", "車番", "O"], ["P", "メモ"]].forEach(([a, v, e]) => put(a + hr, v, e ? `${a}${hr}:${e}${hr}` : null));
    for (let i = 0; i < nCar; i++) {
      const rr = hr + 1 + i, x = (d.車 || [])[i] || [];
      [["A", x[0]], ["B", x[1]], ["C", x[2], "E"], ["F", numx(x[3])], ["G", numx(x[4])], ["H", x[5], "I"], ["J", x[6], "K"], ["L", x[7], "M"], ["N", x[8], "O"], ["P", x[9]]].forEach(([a, v, e]) => put(a + rr, v ?? "", e ? `${a}${rr}:${e}${rr}` : null));
    }
  }
  for (let r = 3; r <= last; r++) for (let c = 1; c <= 16; c++) {
    ws.getCell(r, c).border = BORDER;
  }
  ws.pageSetup.printArea = `A1:P${last}`;
}
function getCellFor(d, i, c) { return cumul(c === 8 ? "kizai" : "hin", i); }
function sumSheet(wb, s, a, b) {
  const S = summarize(s, a, b);
  const ws = wb.addWorksheet("集計", { pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 }, views: [{ showGridLines: false }] });
  const font = { name: "ＭＳ Ｐゴシック", size: 10 };
  let r = 1;
  const cellAt = (rr, cc, v, o = {}) => { const cl = ws.getCell(rr, cc); cl.value = v; cl.font = { ...font, ...(o.font || {}) }; cl.border = BORDER; cl.alignment = { vertical: "middle", horizontal: o.h || "center" }; if (o.fill) cl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: o.fill } }; return cl; };
  const title = t => { const cl = ws.getCell(r, 1); cl.value = t; cl.font = { ...font, size: 12, bold: true }; r += 1; };
  ws.getCell(1, 1).value = `常用 集計表　${gen() ? gen().名前 : ""}　${s.名前}　${jpDate(a)}〜${jpDate(b)}`; ws.getCell(1, 1).font = { ...font, size: 14, bold: true };
  r = 3;
  title("日別");
  const h1 = ["日付", "曜", "天気", "気温", ...S.shoku, "人員計", "残業計(h)", "作業（作業予定の要約）"];
  h1.forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
  const top = r;
  for (const p of S.perDay) {
    const vals = [pd(p.date), WD[wd(p.date)], p.天気, p.気温 === "" ? "" : Number(p.気温), ...S.shoku.map(k => p.c.get(k) || null), p.n, p.ot || null, p.作業];
    vals.forEach((v, i) => { const cl = cellAt(r, i + 1, v, { h: i === vals.length - 1 ? "left" : i >= 3 ? "right" : "center" }); if (i === 0) cl.numFmt = "m/d"; });
    r++;
  }
  cellAt(r, 1, "計", { fill: "FFF5F5F5" }); [2, 3, 4].forEach(c => cellAt(r, c, "", { fill: "FFF5F5F5" }));
  cellAt(r, h1.length, "", { fill: "FFF5F5F5" }); ws.getColumn(h1.length).width = Math.max(ws.getColumn(h1.length).width || 0, 50);
  for (let i = 5; i <= h1.length - 1; i++) { const col = ws.getColumn(i).letter; cellAt(r, i, { formula: `SUM(${col}${top}:${col}${r - 1})` }, { h: "right", fill: "FFF5F5F5", font: { bold: true } }); }
  r += 2;
  title("会社別（日ごとの人数）");
  ["会社", ...S.dates.map(x => +x.slice(8)), "人工計", "残業計(h)"].forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
  for (const cp of S.corder) {
    const C = S.comp.get(cp);
    cellAt(r, 1, cp, { h: "left", font: { bold: true } });
    S.dates.forEach((x, i) => cellAt(r, 2 + i, C.日[x] || null, { h: "right" }));
    cellAt(r, 2 + S.dates.length, C.計, { h: "right", font: { bold: true } }); cellAt(r, 3 + S.dates.length, C.残業 || null, { h: "right" });
    r++;
  }
  r += 1;
  title("職種別（工種）の集計（1人を 昼・夜・残 の3行に。昼・夜＝○、残＝残業時間）");
  {
    const ND = S.dates.length, gray = "FFF5F5F5";
    ["工種", "作業員名", "", ...S.dates.map(x => +x.slice(8)), "人工"].forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
    for (const k of S.shokuBase) {
      const SK = S.byShoku.get(k); if (!SK) continue;
      const top = r, ppl = [...SK.人.values()].sort((p, q) => q.日数 - p.日数);
      cellAt(r, 1, k, { h: "left", font: { bold: true }, fill: gray }); cellAt(r, 2, "日ごとの人数", { h: "left", fill: gray }); cellAt(r, 3, "", { fill: gray });
      S.dates.forEach((x, i) => cellAt(r, 4 + i, SK.日[x] || null, { h: "right", fill: gray, font: { bold: true } }));
      cellAt(r, 4 + ND, SK.計, { h: "right", fill: gray, font: { bold: true } }); r++;
      for (const w of ppl) {
        const pr = r;
        [["昼", x => w.昼[x] ? "○" : "", w.昼数 || null], ["夜", x => w.夜[x] ? "○" : "", w.夜数 || null], ["残", x => w.残[x] || null, w.残業 || null]].forEach(([z, f, t]) => {
          cellAt(r, 1, ""); cellAt(r, 2, z === "昼" ? w.名前 : "", { h: "left" }); cellAt(r, 3, z, { font: { color: { argb: "FF666666" } } });
          S.dates.forEach((x, i) => cellAt(r, 4 + i, f(x), { h: z === "残" ? "right" : "center", font: z === "残" ? { color: { argb: "FFC00000" } } : {} }));
          cellAt(r, 4 + ND, t, { h: "right" }); r++;
        });
        ws.mergeCells(pr, 2, pr + 2, 2);
      }
      ws.mergeCells(top, 1, r - 1, 1);
      ws.getCell(top, 1).alignment = { vertical: "top", horizontal: "left" };
    }
  }
  r += 1;
  title("作業員名別の集計（日ごとの工種：" + SHORT_NOTE + "）");
  const h2 = ["会社", "作業員名", ...S.dates.map(x => +x.slice(8)), "工種ごとの日数", "日数", "残業計(h)", ...(UI.hideMoney ? [] : ["金額"])];
  h2.forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
  for (const w of S.ws) {
    cellAt(r, 1, w.会社, { h: "left" }); cellAt(r, 2, w.名前, { h: "left" });
    S.dates.forEach((x, i) => cellAt(r, 3 + i, x in w.日 ? w.日工種[x].map(shortK).join("") + (w.日[x] ? " " + w.日[x] : "") : "", { font: x in w.日 && w.日[x] ? { color: { argb: "FFC00000" } } : {} }));
    cellAt(r, 3 + S.dates.length, [...w.職種内訳.entries()].sort((p, q) => q[1] - p[1]).map(([k, n]) => k + " " + n).join("・"), { h: "left" });
    cellAt(r, 4 + S.dates.length, w.日数, { h: "right" }); cellAt(r, 5 + S.dates.length, w.残業 || null, { h: "right" }); if (!UI.hideMoney) cellAt(r, 6 + S.dates.length, w.金額 ? Math.round(w.金額) : null, { h: "right" }).numFmt = "#,##0";
    r++;
  }
  const mt = (ttl, m) => {
    if (!m.size) return;
    r += 1; title(ttl);
    const kz = ttl === "使用機材", o = kz ? 1 : 0;
    [...(kz ? ["所有者"] : []), ttl, ...S.dates.map(x => +x.slice(8)), "計"].forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
    for (const [k, v] of m) { const [nm, own] = k.split("\t"); if (kz) cellAt(r, 1, own, { h: "left" }); cellAt(r, 1 + o, nm, { h: "left" }); S.dates.forEach((x, i) => cellAt(r, 2 + o + i, v.日[x] || null, { h: "right" })); cellAt(r, 2 + o + S.dates.length, v.計, { h: "right", font: { bold: true } }); r++; }
  };
  mt("使用機材", S.kizai); mt("材料", S.hin);
  // 高速・燃料
  if (S.car.length) {
    const hm = !!UI.hideMoney, drop = (vals, idx) => hm ? vals.filter((_, i) => !idx.includes(i)) : vals;
    const rowx0 = (vals, o = {}) => { vals.forEach((v, i) => { const cl = cellAt(r, i + 1, v, { h: typeof v === "number" ? "right" : "left", ...o }); if (typeof v === "number" && !o.raw) cl.numFmt = "#,##0.##"; }); r++; };
    const rowx = (vals, o = {}, idx = []) => rowx0(drop(vals, idx), o);   // 金額を隠すときは idx の列を出さない
    r += 1; title("高速・燃料（だれごと）");
    rowx(["会社", "名前", "車両名", "車番", "高速 往路", "復路", "往復", "高速代", "軽油(L)", "ガソリン(L)", "ほかの燃料(L)", "燃料代", "計"], { fill: "FFEEF3F8", h: "center" }, [7, 11, 12]);
    for (const v of carByWho(S.car)) rowx([...v.who, v.往路 || null, v.復路 || null, v.往復 || null, Math.round(v.高速金額) || null, v.軽油 || null, v.ガソリン || null, v.他L || null, Math.round(v.燃料金額) || null, Math.round(v.高速金額 + v.燃料金額)], {}, [7, 11, 12]);
    const ar = carByArea(S.car);
    if (ar.length) { r += 1; title("高速の区間ごと"); rowx(["区間", "往路", "復路", "往復", "回数（往復は2）", "金額"], { fill: "FFEEF3F8", h: "center" }, [5]); ar.forEach(v => rowx([v.区間, v.往路 || null, v.復路 || null, v.往復 || null, v.回, Math.round(v.金額)], {}, [5])); }
    r += 1; title("高速・燃料（記録）");
    rowx(["日付", "種類", "往復", "区間", "給油(L)", "会社", "名前", "車両名", "車番", "金額", "メモ"], { fill: "FFEEF3F8", h: "center" }, [9]);
    for (const x of S.car) { rowx([x.date.slice(5).replace("-", "/"), x.r[0], x.r[1] || "", x.r[2] || "", x.A.種 === "高速" ? null : x.A.L, ...x.who, x.A.なし ? "単価なし" : Math.round(x.A.金額), x.r[9] || ""], {}, [9]); }
  }
  // 金額
  const M = S.money;
  if ((s.単価 || []).length && !UI.hideMoney) {
    r += 1; title("金額（単価マスタから）");
    const tbx = (ttl, m, unit) => {
      if (!m.size) return;
      ["", ttl, "数量", "単位", "単価", "金額"].slice(1).forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
      const top2 = r;
      for (const [k, v] of m) { const [nm, own] = k.split("\t"); cellAt(r, 1, (v.分類 ? "［" + v.分類 + "］" : "") + nm + (own ? "　" + own : ""), { h: "left" }); cellAt(r, 2, Math.round(v.数量 * 100) / 100, { h: "right" }); cellAt(r, 3, v.精算 ? "回" : unit); cellAt(r, 4, v.精算 && !v.単価.size ? "領収書精算" : !v.単価.size && v.実費 && !v.なし ? "実費" : v.単価.size === 1 ? [...v.単価][0] : v.単価.size ? [...v.単価].join("・") : "単価なし", { h: "right" }).numFmt = "#,##0"; cellAt(r, 5, Math.round(v.金額), { h: "right" }).numFmt = "#,##0"; r++; }
      cellAt(r, 1, "計", { fill: "FFF5F5F5" }); [2, 3, 4].forEach(c => cellAt(r, c, "", { fill: "FFF5F5F5" })); cellAt(r, 5, { formula: `SUM(E${top2}:E${r - 1})` }, { h: "right", fill: "FFF5F5F5", font: { bold: true } }).numFmt = "#,##0"; r += 2;
    };
    tbx("労務（工種別・昼夜）", M.労務, "人工"); tbx("残業（工種別）", M.残業, "時間"); tbx("機械（車両・重機・小型機械・電動工具）", M.機械, "台日"); tbx("材料（固定単価・領収書精算）", M.材料, ""); tbx("通勤費（高速代・通勤車）", M.通勤費, "");
    [["労務", M.計.労務], ["残業", M.計.残業], ["機械", M.計.機械], ["材料", M.計.材料], ["通勤費", M.計.通勤費], ["合計", M.合計]].forEach(([t, v]) => { cellAt(r, 1, t, { fill: "FFEEF3F8" }); cellAt(r, 2, Math.round(v), { h: "right", font: t === "合計" ? { bold: true } : {} }).numFmt = "#,##0"; r++; });
  }
  ws.getColumn(1).width = 18; ws.getColumn(2).width = 13;
  for (let c = 3; c <= Math.max(h1.length, h2.length) + 2; c++) ws.getColumn(c).width = 6.5;
  ws.getColumn(3).width = 8; ws.getColumn(4).width = 7;
}

/* ================= PDF（印刷。A4 横で1日1ページ） ================= */
function dayHTML(s, x, d) { const n = pageCount(d); return [...Array(n).keys()].map(p => dayPageHTML(s, x, pageOf(d, p), d, p, n)).join(""); }
// 字が枠からはみ出るときは横に縮める（2026-10-31 ユーザー）。幅は列の数から見積もる（1列 17.86mm）
function fitP(v, cs = 1) {
  v = String(v ?? ""); if (!v) return "";
  const room = cs * 17.86 - (cs >= 2 ? 8.7 : 0) - 2.5, est = [...v].reduce((a, ch) => a + (ch.charCodeAt(0) > 0xff ? 3.53 : 1.95), 0);
  if (est <= room) return esc(v);
  return `<span class="sq" style="width:${est.toFixed(1)}mm;margin-right:${(room - est).toFixed(1)}mm;transform:scaleX(${(room / est).toFixed(3)})">${esc(v)}</span>`;
}
function dayPageHTML(s, x, d, full, pi, pn) {
  const dt = pd(x), keep = UI.day; UI.day = x;
  const c = (v, cls = "", span = "") => `<td class="${cls}"${span}>${fitP(v, +((/colspan="(\d+)"/.exec(span) || [0, 1])[1]))}</td>`;
  const sp = (cs, rs) => (cs > 1 ? ` colspan="${cs}"` : "") + (rs > 1 ? ` rowspan="${rs}"` : "");
  const up = i => d.上[i] || [];
  const h = [];
  h.push(`<div class="pg"><table class="pt"><colgroup>${[...Array(16)].map((_, i) => `<col style="width:${i === 12 ? 9.2 : 17.86}mm">`).join("")}</colgroup>`);
  h.push(`<tr><td class="t" colspan="16">作　業　日　報${pn > 1 ? `（${pi + 1}／${pn}）` : ""}</td></tr><tr class="b"><td class="nb" colspan="16"></td></tr>`);
  h.push(`<tr class="a">${c("元請業者")}${c(gen() ? gen().名前 : "", "", sp(3))}${c("日　　付", "", sp(2))}${c(jpDate(x), "", sp(2))}${c(WD[dt.getDay()])}${c("天気")}${c("気温")}${c("記入者", "", sp(1, 2))}${c(d.担当者 || "", "", sp(4, 2))}</tr>`);
  h.push(`<tr class="a">${c("現場名")}${c(s.名前, "", sp(8))}${c(d.天気 || "")}${c((d.気温 ?? "") + "℃")}</tr>`);
  // 上の5段（Excel は2行ずつ）。空の行は高さがつぶれるので、1段＝高い1行で組む
  // 左：作業内容 5段／右：工種別の集計（2026-10-31 ユーザー：サイン・備考をやめた）
  h.push(`<tr class="a2">${c("作　業　予　定", "", sp(6))}${c("完　了　報　告", "", sp(5))}${c("工種別の集計", "", sp(5))}</tr>`);
  for (let i = 0; i < UP_ROWS; i++) h.push(`<tr class="a2"><td class="wk"${sp(6)}>${fitP(up(i)[1] || "", 6)}</td><td class="wk"${sp(5)}>${fitP(up(i)[4] || "", 5)}</td>${i === 0 ? `<td class="ks"${sp(5, UP_ROWS)}>${kouSumHTML(full)}</td>` : ""}</tr>`);
  h.push(`<tr class="b2">${c("作業員名", "", sp(2))}${c("始業時間")}${c("終業時間")}${c("工種")}${c("残業時間")}${c("使用機材", "", sp(2))}${c("規格・形式")}${c("数量")}${c("材料名", "", sp(3))}${c("数量")}${c("実費", "", sp(2))}</tr>`);
  const nCar = PG.car, n = PG.dn;
  for (let i = 0; i < n; i++) {
    const r = d.下[i] || [];
    h.push(`<tr class="a">${c(r[0] || "", "", sp(2))}${c(r[0] ? r[1] || "" : "")}${c(r[0] ? r[2] || "" : "")}${c(r[3] || "")}${c(r[4] || "", "ot")}${c(r[5] ? (r[13] ? "（" + abbrOf(r[13]) + "）" : "") + r[5] : "", "", sp(2))}${c(r[6] || "")}${c(r[7] || "")}${c(r[9] || "", "", sp(3))}${c(r[10] || "")}${c(r[15] !== "" && r[15] != null && !isNaN(Number(r[15])) ? yen(Number(r[15])) : r[15] || "", "", sp(2))}</tr>`);
  }
  // 高速・燃料
  h.push(`<tr class="b2">${c("種類")}${c("往復")}${c("高速の区間", "", sp(3))}${c("給油(L)")}${c("実費")}${c("会社", "", sp(2))}${c("名前", "", sp(2))}${c("車両名", "", sp(2))}${c("車番", "", sp(2))}${c("メモ")}</tr>`);
  for (let i = 0; i < nCar; i++) { const r = (d.車 || [])[i] || [], jv = r[4] !== "" && r[4] != null && !isNaN(Number(r[4])) ? yen(Number(r[4])) : r[4] || ""; h.push(`<tr class="a">${c(r[0] || "")}${c(r[1] || "")}${c(r[2] || "", "", sp(3))}${c(r[3] || "")}${c(jv)}${c(r[5] || "", "", sp(2))}${c(r[6] || "", "", sp(2))}${c(r[7] || "", "", sp(2))}${c(r[8] || "", "", sp(2))}${c(r[9] || "", "sm")}</tr>`); }
  h.push(`</table></div>`);
  UI.day = keep;
  return h.join("");
}
function printDays(dates, sumRange) {
  const s = site(); if (!s) { toast("現場を選んでください"); return; }
  const h = dates.map(x => dayHTML(s, x, s.日[x] || { 上: [], 下: [] }));
  if (sumRange) {
    const keepF = $("sumFrom").value, keepT = $("sumTo").value;
    $("sumFrom").value = sumRange[0]; $("sumTo").value = sumRange[1]; renderSum(true);   // PDF は全部の表
    h.push(`<div class="pg sum"><h3>常用 集計表　${esc(gen() ? gen().名前 : "")}　${esc(s.名前)}　${jpDate(sumRange[0])}〜${jpDate(sumRange[1])}</h3>${$("sumBody").innerHTML.replace(/<h3 class="sh">/g, "<h3>")}</div>`);
    if (keepF) { $("sumFrom").value = keepF; $("sumTo").value = keepT; renderSum(); }
  }
  $("printArea").innerHTML = h.join("");
  toast("印刷の画面で「PDF に保存」を選ぶと PDF になります");
  setTimeout(() => window.print(), 300);
}

/* ================= そのほか ================= */
let toastT = null;
function toast(m, ms = 3200) { const t = $("toast"); t.textContent = m; t.classList.add("on"); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("on"), ms); }
document.addEventListener("keydown", e => {
  if (e.altKey && (e.key === "PageUp" || e.key === "PageDown") && !$("pSheet").hidden) { e.preventDefault(); moveDay(e.key === "PageDown" ? 1 : -1); }
});

/* ---------- はじめ ---------- */
UI.day ||= ymd(new Date());
if (kouFix()) setTimeout(save, 0);   // 工種の別名・自社名・会社の空欄を日報に当てる（開いたとき）
buildSheet();
showTab(DB.現場.length ? (UI.tab || "sheet") : "master");
syncMasterSites();
renderAll();
showMe(); setRO(null);
startAuth();
if (!ME.名前) askName();
if (!DB.現場.length) toast("「サインイン」すると現場一覧・名簿を読みます。元の作業日報の Excel は「取込」で読めます");
