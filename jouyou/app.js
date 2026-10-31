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
function blankDB() {
  return { 形式: "常用日報", 版: 1, 元請: [], 現場: [], 職種: ["職長", "重機OP", "土工", "普通作業員", "管理"], 天気: ["晴", "曇", "雨", "雪", "晴/曇", "曇/晴", "曇/雨", "雨/曇", "晴/雨"] };
}
function fixSite(s) {
  s.id ||= uid(); s.名前 ||= ""; s.担当者 ||= ""; s.始業 ||= "8:00"; s.終業 ||= "17:00";
  if (s.締め日 == null) s.締め日 = 20;
  s.累計 ||= "期間"; s.作業員 ||= []; s.機材 ||= []; s.日 ||= {}; s.消した日 ||= {}; s.設定更新 ||= "";
  return s;
}
function fixDB(d) {
  const b = blankDB();
  for (const k in b) if (d[k] == null) d[k] = b[k];
  d.現場.forEach(fixSite);
  return d;
}
let DB = (() => { try { const s = localStorage.getItem(KEY); if (s) return fixDB(JSON.parse(s)); } catch (e) {} return fixDB(blankDB()); })();
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
  if (!s.日[UI.day]) s.日[UI.day] = { 天気: "", 気温: "", 担当者: s.担当者 || "", 上: [], 下: [], 備考: "", サイン: "", サイン日時: "", 人員手: false, 更新: "" };
  return s.日[UI.day];
}
// 中身のある日か（空の日は保存しない）
function hasData(d) {
  if (!d) return false;
  if (d.天気 || d.気温 || d.備考 || d.サイン) return true;
  return [...d.上, ...d.下].some(r => r && r.some(v => v !== "" && v != null));
}
function tidy(d) {
  // 後ろの空の行は持たない
  const trim = a => { while (a.length && !(a[a.length - 1] || []).some(v => v !== "" && v != null)) a.pop(); };
  trim(d.上); trim(d.下);
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
  return String(v ?? "").replace(/[\t\r\n]+/g, " ").trim();
}
// 残業時間（時間）＝ 定時の終業より遅い分＋定時の始業より早い分。夜をまたぐ終業は＋24時間
function overtime(st, en, s = site()) {
  const a = toMin(st), b0 = toMin(en), A = toMin(s.始業), B = toMin(s.終業);
  if (a == null || b0 == null || A == null || B == null) return "";
  const b = b0 < a ? b0 + 1440 : b0;
  const m = Math.max(0, b - B) + Math.max(0, A - a);
  return m ? String(Math.round(m / 60 * 100) / 100) : "";
}

/* ================= 表の列 ================= */
const UP_ROWS = 5, DN_ROWS = 16;
const UPC = [
  { t: "工種", w: 96, k: "txt", cls: "c", align: "c" },
  { t: "作業内容", w: 220, k: "txt", cls: "c", align: "c" },
  { t: "人員", w: 64, k: "nin", cls: "r", align: "r" },
  { t: "使用資機材・重機", w: 318, k: "txt" }
];
const DNC = [
  { t: "作業員名", w: 116, k: "txt", stopFlow: true },
  { t: "始業時間", w: 62, k: "time", cls: "c", align: "c" },
  { t: "終業時間", w: 62, k: "time", cls: "c", align: "c" },
  { t: "工種", w: 82, k: "txt", cls: "c", align: "c" },
  { t: "残業時間", w: 62, k: "hrs", cls: "r", align: "r" },
  { t: "使用機材", w: 128, k: "txt", cls: "sepL" },
  { t: "規格・形式", w: 112, k: "txt" },
  { t: "数量", w: 50, k: "num", cls: "r", align: "r" },
  { t: "累計", w: 54, k: "num", cls: "r ro", align: "r", ro: true },
  { t: "品名", w: 112, k: "txt", cls: "sepL" },
  { t: "数量", w: 50, k: "num", cls: "r", align: "r" },
  { t: "累計数量", w: 64, k: "num", cls: "r ro", align: "r", ro: true }
];
UPC.forEach((c, i) => c.letter = String.fromCharCode(65 + i));
DNC.forEach((c, i) => c.letter = String.fromCharCode(65 + i));
const SECS = { up: UPC, dn: DNC };

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
  return ((d[key === "up" ? "上" : "下"][r] || [])[c]) ?? "";
}

// 作業員の工種（職種）ごとの人数。名簿の職種の順
function countByShoku(d) {
  const m = new Map();
  (d ? d.下 : []).forEach(r => { if (r && r[0]) { const k = r[3] || "（工種なし）"; m.set(k, (m.get(k) || 0) + 1); } });
  const order = [...DB.職種];
  return [...m.entries()].sort((a, b) => { const ia = order.indexOf(a[0]), ib = order.indexOf(b[0]); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib); });
}
// 上の表の「人員」に入れる：作業内容が同じ工種の行 → 工種が同じ行（作業内容が工種名でないもの）→ 空いている行
function fillNinzu(d, force) {
  if (!d || (d.人員手 && !force)) return false;
  const cnt = countByShoku(d);
  if (!cnt.length) return false;
  const up = d.上; while (up.length < UP_ROWS) up.push(["", "", "", ""]);
  up.forEach(r => { while (r.length < 4) r.push(""); });
  const used = new Set(), names = new Set(cnt.map(x => x[0]));
  const before = JSON.stringify(up);
  for (const [k, n] of cnt) {
    let i = up.findIndex((r, j) => !used.has(j) && r[1] === k);
    if (i < 0) i = up.findIndex((r, j) => !used.has(j) && r[0] === k && !names.has(r[1]));
    if (i < 0) i = up.findIndex((r, j) => !used.has(j) && !r[0] && !r[1] && !r[2]);
    if (i < 0) continue;
    if (!up[i][0] && !up[i][1]) up[i][0] = k;
    up[i][2] = n + "名"; used.add(i);
  }
  if (force) d.人員手 = false;
  return JSON.stringify(up) !== before;
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

/* ================= 揺らぎ対策（2026-10-31 ユーザー：人の名前・使用機材・材料など、同じものを違う名前にしないように） =================
   見る列：作業員名・工種（上と下）・使用機材・規格・品名。
   ・打った言葉を、ゆるい形（全角半角・空白・中黒・かっこ・ハイフン類・英字の大小を無視）で、これまでの書き方（マスタ・この現場の日報）と比べる。
     ゆるい形が同じ → これまでの書き方にそろえる（だまって直し、下に「そろえました」と出す）。
     1字ちがい・片方がもう片方を含む → 「○○のことですか？」と聞く（そろえる／別のものとして入れる）。
   ・「現場・名簿」の「揺らぎを見つけてそろえる」で、これまでの日報の中の揺らぎをまとめて直せる。 */
const YCOLS = { dn: { 0: "作業員名", 3: "工種", 5: "使用機材", 6: "規格・形式", 9: "品名" }, up: { 0: "工種" } };
const loose = v => String(v ?? "").normalize("NFKC").toLowerCase().replace(/[\s　・･()（）\[\]「」【】,，.．]/g, "").replace(/[-‐－―−ｰ]/g, "ー");
function lev(a, b) {
  if (Math.abs(a.length - b.length) > 2) return 9;
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return d[a.length][b.length];
}
// 似ているか（ゆるい形で 1字ちがい（8字以上は2字）・3字以上で片方がもう片方を含む）
function similar(a, b) {
  const x = loose(a), y = loose(b);
  if (!x || !y || x === y) return x === y;
  const n = Math.min(x.length, y.length);
  if (n >= 3 && (x.includes(y) || y.includes(x))) return true;
  return n >= 2 && lev(x, y) <= (n >= 8 ? 2 : 1);
}
// その列のこれまでの書き方（マスタを先に。回数つき）
function canon(key, c, s = site()) {
  const cnt = new Map(), add = (v, w = 1) => { v = String(v || "").trim(); if (v) cnt.set(v, (cnt.get(v) || 0) + w); };
  if (key === "dn" && c === 0) { people().forEach(p => add(p.n, 1000)); s && s.作業員.forEach(w => add(w.名前, 10)); }
  if (c === 3 && key === "dn" || key === "up" && c === 0) DB.職種.forEach(k => add(k, 1000));
  if (key === "dn" && c === 5) { MS.機械.forEach(m => add(m.名前, 1000)); s && s.機材.forEach(m => add(m.名前, 10)); }
  if (key === "dn" && c === 6) { MS.機械.forEach(m => add(m.規格, 1000)); s && s.機材.forEach(m => add(m.規格, 10)); }
  if (s) for (const dd of Object.values(s.日)) (key === "up" ? dd.上 : dd.下).forEach(r => r && add(r[c]));
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
    q.map((x, i) => `<tr><td>${esc(YCOLS[x.key][x.c])}</td><td><b>${esc(x.v)}</b></td><td>${x.ask.map((w, j) => `<label class="yr"><input type="radio" name="yr${i}" value="${esc(w)}"${j === 0 ? " checked" : ""}> ${esc(w)}</label>`).join("")}<label class="yr"><input type="radio" name="yr${i}" value=""> 別のもの（このまま）</label></td></tr>`).join("") + `</tbody></table>`;
  openDlg($("dYure"));
  $("yrOk").onclick = () => {
    const ch = { up: [], dn: [] };
    q.forEach((x, i) => { const w = (document.querySelector(`input[name="yr${i}"]:checked`) || {}).value; if (w) ch[x.key].push({ r: x.r, c: x.c, v: w }); });
    closeDlg($("dYure")); YQ = [];
    for (const k of ["up", "dn"]) if (ch[k].length) applyCells(k, ch[k], true);
  };
}

/* ================= セルを変える ================= */
function applyCells(key, ch, noCheck) {
  const s = site(); if (!s) { toast("先に「現場・名簿」で元請と現場を登録してください"); return; }
  if (!canEdit()) { refreshSheet(); return; }
  // 揺らぎ：打つ前のこれまでの書き方と比べる（同じならそろえる・似ていれば後で聞く）
  const fixedMsg = [];
  if (!noCheck && YCOLS[key]) {
    const cns = {};
    ch = ch.map(x => {
      if (!YCOLS[key][x.c] || !x.v) return x;
      const r = checkWord(key, x.c, x.v, cns[x.c] ||= canon(key, x.c, s));
      if (r.fixed) fixedMsg.push(`「${r.fixed}」→「${r.v}」`);
      if (r.ask && r.ask.length) YQ.push({ key, r: x.r, c: x.c, v: r.v, ask: r.ask });
      return { ...x, v: r.v };
    });
  }
  pushUndo();
  const d = ensureDay(), arr = key === "up" ? d.上 : d.下, NC = SECS[key].length;
  const touched = new Set();
  for (const { r, c, v } of ch) {
    if (key === "up" && r >= UP_ROWS) continue;
    while (arr.length <= r) arr.push(Array(NC).fill(""));
    const row = arr[r]; while (row.length < NC) row.push("");
    row[c] = v; touched.add(r);
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
        if (!row[1]) row[1] = s.始業; if (!row[2]) row[2] = s.終業;
      }
      if (c0.includes(1) || c0.includes(2) || (c0.includes(0) && row[0])) {
        if (!c0.includes(4)) row[4] = overtime(row[1], row[2], s);
      }
      // 自社機械の名前を入れたら規格（1つに決まるとき）
      if (c0.includes(5) && row[5] && !row[6]) { const ms = MS.機械.filter(m => nk(m.名前) === nk(row[5])); if (ms.length === 1) row[6] = ms[0].規格; }
      // この現場の名簿・機材に足す（工種を覚える。次から一覧に出る）
      if (row[0] && !s.作業員.some(x => nk(x.名前) === nk(row[0]))) s.作業員.push({ 名前: row[0], 職種: row[3] || "" });
      else if (row[0] && row[3]) { const w = s.作業員.find(x => nk(x.名前) === nk(row[0])); if (w && !w.職種) w.職種 = row[3]; }
      if (row[5] && !s.機材.some(x => x.名前 === row[5] && (x.規格 || "") === (row[6] || ""))) s.機材.push({ 名前: row[5], 規格: row[6] || "" });
    }
    fillNinzu(d, false);
  }
  tidy(d); d.更新 = now();
  if (!hasData(d)) dropDay(s, UI.day);
  save(); refreshSheet();
  if (fixedMsg.length) toast("これまでの書き方にそろえました：" + fixedMsg.slice(0, 3).join("・") + (fixedMsg.length > 3 ? " ほか" : ""), 4500);
  if (YQ.length) setTimeout(askYure, 50);
}
function rowOp(key, op, r, n) {
  const s = site(); if (!s || !canEdit()) return;
  pushUndo();
  const d = ensureDay(), arr = key === "up" ? d.上 : d.下, NC = SECS[key].length;
  if (op === "insert") { while (arr.length < r) arr.push(Array(NC).fill("")); arr.splice(r, 0, ...Array.from({ length: n }, () => Array(NC).fill(""))); if (key === "up") arr.length = Math.min(arr.length, UP_ROWS); }
  else arr.splice(r, n);
  if (key === "dn") fillNinzu(d, false);
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
  if ((key === "dn" && c === 3) || (key === "up" && (c === 0 || c === 1))) DB.職種.forEach(k => add(k, 50));
  if (key === "dn" && c === 5) { s.機材.forEach(m => add(m.名前, 100)); MS.機械.forEach(m => add(m.名前, 50)); }
  if (key === "dn" && c === 6) { s.機材.forEach(m => add(m.規格, 100)); MS.機械.forEach(m => add(m.規格, 50)); }
  if (key === "dn" && (c === 1 || c === 2)) [s.始業, s.終業, "8:00", "17:00", "18:00", "19:00", "20:00"].forEach(t => add(t, 1));
  if (key === "up" && c === 2) countByShoku(day()).forEach(([k, n]) => add(n + "名", 100));
  for (const dd of Object.values(s.日)) (key === "up" ? dd.上 : dd.下).forEach(r => r && add(r[c]));
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
    <div class="ttlRow">作　業　日　報</div>
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
        <div class="secT"><b>工種・作業内容・人員</b><span class="cnt" id="cntUp"></span></div>
        <div class="g1" id="g_up"></div>
      </div>
      <div class="right">
        <div class="rhd"><div class="l">元請<br>サイン</div><div id="signBox" title="iPad・スマホで押すとサインを書けます"></div></div>
        <div class="bikouL">備考</div>
        <textarea id="hBikou"></textarea>
        <div class="akamaru">赤丸は取決め以外</div>
      </div>
    </div>
    <div class="low sec" data-key="dn">
      <div class="secT"><b>作業員・使用機材・品名</b><span class="cnt" id="cntDn"></span></div>
      <div id="g_dn"></div>
    </div>`;
  for (const key of ["up", "dn"]) {
    const cols = SECS[key];
    grids[key] = createGrid({
      host: $("g_" + key), scroller: $("sheetWrap"), cols, headRows: 1, rowBase: 1, rowHeadW: 26, noLetters: true,
      formula: $("fx"), isCur: () => cur === key, onFocus: () => setCur(key),
      rows: () => key === "up" ? UP_ROWS : Math.max(DN_ROWS, (day() ? day().下.length : 0) + 1),
      head: () => `<tr class="xlhr sh"><th class="xlrh"></th>${cols.map((c, i) => `<th data-col="${i}" class="${(c.cls || "").includes("sepL") ? "sepL" : ""}">${esc(c.t)}</th>`).join("")}</tr>`,
      get: (r, c) => getCell(key, r, c),
      deco: (r, c) => {
        if (key !== "dn") return null;
        const d = day(), row = d && d.下[r];
        if (!row) return null;
        if (c === 4 && row[4]) return { cls: "ot", title: "定時（" + site().始業 + "〜" + site().終業 + "）を超えた時間" };
        return null;
      },
      normalize: (c, v) => normalize(cols[c].k, v),
      complete: (c, v) => { const lv = v.toLowerCase(), ws = colWords(key, c); if (ws.some(x => x.toLowerCase() === lv)) return null; return ws.find(x => x.toLowerCase().startsWith(lv) && x.length > v.length) || null; },   // Excel と同じく、同じ言葉があれば補わない
      list: c => colWords(key, c),
      apply: ch => applyCells(key, ch),
      undo, redo,
      onSheet: dir => moveDay(dir),
      onEdge: (dir, c) => {
        const t = key === "up" && dir > 0 ? "dn" : key === "dn" && dir < 0 ? "up" : null;
        if (!t) return false;
        grids[t].enter(dir, Math.min(c, SECS[t].length - 1));
        return true;
      },
      rowOp: (op, r, n) => rowOp(key, op, r, n),
      menu: (x, y, sel) => openMenu(key, x, y, sel),
      print: () => printDays([UI.day]),
      toast,
      cellName: (r, c) => (key === "up" ? "上" : "下") + (r + 1) + "・" + cols[c].t,
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
  hv("hWx", "天気"); hv("hTmp", "気温", v => { v = tmpOnly(v); return v === "-" || v === "." ? "" : normNum(v); }); hv("hTan", "担当者", v => masterName(v.trim())); hv("hBikou", "備考");
  $("signBox").addEventListener("pointerup", e => {
    if (e.target.closest(".del")) return;
    if (!site() || !canEdit()) return;
    if (e.pointerType === "mouse") { toast("元請サインは iPad・スマホで書きます（パソコンでは書けません）"); return; }
    openSign();
  });
}
function refreshSheet() {
  const s = site(), d = day();
  $("hGen").textContent = gen() ? gen().名前 : "（元請業者を選ぶ）";
  $("hSite").textContent = s ? s.名前 : "（現場を選ぶ）";
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
  $("hTan").placeholder = s ? s.担当者 : "";
  if (document.activeElement !== $("hBikou")) $("hBikou").value = d ? d.備考 : "";
  // サイン
  const sb = $("signBox");
  sb.innerHTML = d && d.サイン ? `<img src="${d.サイン}" alt="サイン"><span class="when">${esc(d.サイン日時)}</span><button class="sub del" title="サインを消す">×</button>` : "（iPad・スマホで押してサイン）";
  const del = sb.querySelector(".del");
  if (del) del.onclick = e => { e.stopPropagation(); if (!confirm("この日の元請サインを消しますか？")) return; pushUndo(); const dd = day(); dd.サイン = ""; dd.サイン日時 = ""; save(); refreshSheet(); };
  grids.up.render(); grids.dn.render();
  // 人数のまとめ
  const cnt = countByShoku(d), tot = cnt.reduce((a, x) => a + x[1], 0);
  const upTot = d ? d.上.reduce((a, r) => a + ninNum(r && r[2]), 0) : 0;
  $("cntDn").textContent = tot ? "作業員 " + cnt.map(([k, n]) => k + " " + n + "名").join("・") + "　計 " + tot + "名" : "";
  $("cntUp").innerHTML = d && (tot || upTot) ? `人員の計 ${upTot}名` + (upTot !== tot ? `　<span class="warn">下の作業員（${tot}名）と合いません。「人員を数える」で入れ直せます</span>` : "") + (d.人員手 ? "　（人員は手で直したので自動で変えません）" : "") : "";
  $("status").textContent = d && d.更新 ? "最後に直した " + d.更新 : s ? "まだ書いていない日" : "";
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

function setDay(d) { if (grids.dn && grids.dn.editing()) grids.dn.commit(); if (grids.up && grids.up.editing()) grids.up.commit(); UI.day = d; saveUI(); showDayBtn(); refreshSheet(); if (!$("pSum").hidden) renderSum(); }
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
    ["行の挿入", "Ctrl+Shift+＋", () => rowOp(key, "insert", sel.r0, sel.r1 - sel.r0 + 1)],
    ["行の削除", "Ctrl+－", () => rowOp(key, "delete", sel.r0, sel.r1 - sel.r0 + 1)],
    ["数式と値のクリア", "Delete", () => { const ch = []; for (let r = sel.r0; r <= sel.r1; r++) for (let c = sel.c0; c <= sel.c1; c++) if (!SECS[key][c].ro) ch.push({ r, c, v: "" }); applyCells(key, ch); }],
    null,
    ["候補から選ぶ", "Alt+↓", () => grids[key].showList()]
  ];
  if (key === "up") items.push(["人員を作業員から数える", "", () => doCount()]);
  m.innerHTML = items.map((it, i) => it ? `<div data-i="${i}"><span>${it[0]}</span><kbd>${it[1]}</kbd></div>` : "<hr>").join("");
  document.body.appendChild(m);
  m.style.left = Math.min(x, innerWidth - m.offsetWidth - 4) + "px"; m.style.top = Math.min(y, innerHeight - m.offsetHeight - 4) + "px";
  m.addEventListener("mousedown", e => { e.preventDefault(); const d = e.target.closest("[data-i]"); if (d) { closeMenu(); items[+d.dataset.i][2](); grids[key].focus(); } });
}
function closeMenu() { const m = $("ctx"); m && m.remove(); }
document.addEventListener("pointerdown", e => { if (!e.target.closest("#ctx")) closeMenu(); });

/* ---------- ボタン ---------- */
function doCount() {
  if (!canEdit()) return;
  const d = day(); if (!d || !countByShoku(d).length) { toast("下の表に作業員がいません"); return; }
  pushUndo(); fillNinzu(d, true); d.更新 = now(); save(); refreshSheet(); toast("作業員の工種ごとの人数を「人員」に入れました");
}
$("bCount").onclick = doCount;
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
function outInfo() { $("outInfo").textContent = outDates().length + " 日分を出します（1シート／1ページに1日）"; }
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
  out.作業員 = [...a.作業員]; b.作業員.forEach(w => { if (!out.作業員.some(x => nk(x.名前) === nk(w.名前))) out.作業員.push(w); });
  out.機材 = [...a.機材]; b.機材.forEach(m => { if (!out.機材.some(x => x.名前 === m.名前 && (x.規格 || "") === (m.規格 || ""))) out.機材.push(m); });
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
const busyEditing = () => (grids.up && grids.up.editing()) || (grids.dn && grids.dn.editing()) || ["hTan", "hTmp", "hBikou"].includes((document.activeElement || {}).id);
async function sync() {
  if (!signedIn()) return;
  if (SY.busy) { SY.again = true; return; }
  SY.busy = true; syncMark("busy", "共有 …");
  let changed = false;
  try {
    // 1) 読む：ほかの端末が書いたもの（eTag が変わったもの）を合わせる
    for (const it of await store.list()) {
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
      if (i < 0) DB.現場.push(merged); else DB.現場[i] = merged;
      SY.files[id] = { eTag: r.eTag, sent: JSON.stringify(r.data) };
      changed = true;
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
          if (r && r.data && r.data.現場) { const i = DB.現場.findIndex(x => x.id === s.id); DB.現場[i] = mergeSite(DB.現場[i], r.data.現場); SY.files[s.id] = { eTag: r.eTag, sent: JSON.stringify(r.data) }; changed = true; }
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
      const other = by && by.端末 !== ME.端末 && Date.now() - by.いつ < STALE;
      if (other && !force) { setRO({ by }); if (!fromSelf) toast(`${by.名前}さんが使っています。見るだけ（閲覧のみ）で開きました`, 5000); return; }
      const doc = { 名前: ME.名前, 端末: ME.端末, 開始: by && by.端末 === ME.端末 ? by.開始 : Date.now(), いつ: Date.now() };
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
      if (by && by.端末 !== ME.端末) { LOCK.mine = false; setRO({ by }); toast(`${by.名前}さんが編集を始めたので、見るだけにしました`, 6000); return; }
      const w = await store.put(lockName(id), JSON.stringify({ ...(by || {}), 名前: ME.名前, 端末: ME.端末, 開始: by ? by.開始 : Date.now(), いつ: Date.now() }), r && r.eTag);
      LOCK.eTag = w.eTag;
    } else if (RO && RO.by && !RO.gone) {
      if (!by || Date.now() - by.いつ >= STALE) setRO({ gone: true, by: RO.by });
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
  $("inGen").innerHTML = gs.length ? gs.map(g => `<option value="${g.id}"${g.id === UI.gen ? " selected" : ""}>${esc(g.名前)}</option>`).join("") : `<option value="">（元請業者なし）</option>`;
  const nDays = s => Object.keys(s.日).length;
  const ss = DB.現場.filter(s => s.元請id === UI.gen).sort((a, b) => (nDays(b) > 0) - (nDays(a) > 0) || String(b.コード || "").localeCompare(String(a.コード || "")));
  if (!ss.some(s => s.id === UI.site)) UI.site = ss.length ? ss[0].id : null;
  $("inSite").innerHTML = ss.length ? ss.map(s => `<option value="${s.id}"${s.id === UI.site ? " selected" : ""} title="${esc(s.名前)}">${esc(siteLabel(s))}</option>`).join("") : `<option value="">（現場なし）</option>`;
  saveUI();
}
$("inGen").onchange = () => { UI.gen = $("inGen").value; UI.site = null; UNDO = []; REDO = []; renderAll(); };
$("inSite").onchange = () => { UI.site = $("inSite").value; UNDO = []; REDO = []; renderAll(); };
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
  const shoku = [], workers = new Map(), kizai = new Map(), hin = new Map();
  const order = k => { const i = DB.職種.indexOf(k); return i < 0 ? 99 : i; };
  const perDay = dates.map(x => {
    const d = s.日[x], c = new Map(); let ot = 0, n = 0;
    d.下.forEach(r => {
      if (!r) return;
      if (r[0]) {
        const k = r[3] || "（工種なし）"; c.set(k, (c.get(k) || 0) + 1); n++;
        if (!shoku.includes(k)) shoku.push(k);
        const o = Number(r[4]) || 0; ot += o;
        const w = workers.get(nk(r[0])) || { 名前: masterName(r[0]), 職種: new Map(), 日: {}, 日数: 0, 残業: 0 };
        w.職種.set(k, (w.職種.get(k) || 0) + 1); if (!(x in w.日)) w.日数++; w.日[x] = (w.日[x] || 0) + o; w.残業 += o; workers.set(nk(r[0]), w);
      }
      if (r[5]) { const k = r[5] + (r[6] ? "（" + r[6] + "）" : ""); const m = kizai.get(k) || { 日: {}, 計: 0 }; const q = Number(r[7]) || 0; m.日[x] = (m.日[x] || 0) + q; m.計 += q; kizai.set(k, m); }
      if (r[9]) { const m = hin.get(r[9]) || { 日: {}, 計: 0 }; const q = Number(r[10]) || 0; m.日[x] = (m.日[x] || 0) + q; m.計 += q; hin.set(r[9], m); }
    });
    return { date: x, 天気: d.天気, 気温: d.気温, c, n, ot, sign: !!d.サイン };
  });
  shoku.sort((p, q) => order(p) - order(q));
  const ws = [...workers.values()].map(w => ({ ...w, 職種: [...w.職種.entries()].sort((p, q) => q[1] - p[1])[0][0] }))
    .sort((p, q) => order(p.職種) - order(q.職種) || q.日数 - p.日数);
  return { dates, shoku, perDay, ws, kizai, hin };
}
const r2 = x => String(Math.round(x * 100) / 100);
function renderSum() {
  const s = site(); if (!s) { $("sumBody").innerHTML = `<p class="note">現場を選んでください</p>`; return; }
  const a = $("sumFrom").value, b = $("sumTo").value;
  if (!a || !b || a > b) return;
  const S = summarize(s, a, b);
  if (!S.dates.length) { $("sumBody").innerHTML = `<p class="note">${jpDate(a)}〜${jpDate(b)} に書いた日報がありません</p>`; return; }
  const dl = x => `${+x.slice(5, 7)}/${+x.slice(8)}`;
  const wcls = x => { const w = wd(x); return w === 0 ? "sun" : w === 6 ? "sat" : ""; };
  const h = [];
  h.push(`<h3 class="sh">日別（${esc(s.名前)}　${jpDate(a)}〜${jpDate(b)}）</h3>`);
  h.push(`<table class="sm"><thead><tr><th>日付</th><th>曜</th><th>天気</th><th>気温</th>${S.shoku.map(k => `<th>${esc(k)}</th>`).join("")}<th>人員計</th><th>残業計(h)</th><th>サイン</th></tr></thead><tbody>`);
  const tot = new Map(); let tn = 0, tot_ot = 0;
  for (const p of S.perDay) {
    h.push(`<tr class="${wcls(p.date)}"><td><a data-go="${p.date}">${dl(p.date)}</a></td><td class="c">${WD[wd(p.date)]}</td><td class="c">${esc(p.天気)}</td><td class="r">${esc(p.気温)}</td>${S.shoku.map(k => `<td class="r">${p.c.get(k) || ""}</td>`).join("")}<td class="r"><b>${p.n}</b></td><td class="r">${p.ot ? r2(p.ot) : ""}</td><td class="c">${p.sign ? "✓" : ""}</td></tr>`);
    S.shoku.forEach(k => tot.set(k, (tot.get(k) || 0) + (p.c.get(k) || 0))); tn += p.n; tot_ot += p.ot;
  }
  h.push(`</tbody><tfoot><tr><td colspan="4">計（${S.dates.length}日）</td>${S.shoku.map(k => `<td class="r">${tot.get(k)}</td>`).join("")}<td class="r">${tn}</td><td class="r">${r2(tot_ot)}</td><td class="c">${S.perDay.filter(p => p.sign).length}</td></tr></tfoot></table>`);

  h.push(`<h3 class="sh">作業員別（出面：○、数字は残業時間）</h3>`);
  h.push(`<table class="sm"><thead><tr><th>作業員名</th><th>工種</th>${S.dates.map(x => `<th class="${wcls(x)}">${+x.slice(8)}</th>`).join("")}<th>日数</th><th>残業計(h)</th></tr></thead><tbody>`);
  for (const w of S.ws) h.push(`<tr><td>${esc(w.名前)}</td><td>${esc(w.職種)}</td>${S.dates.map(x => x in w.日 ? `<td class="c">${w.日[x] ? `<span class="ot">${r2(w.日[x])}</span>` : "○"}</td>` : "<td></td>").join("")}<td class="r"><b>${w.日数}</b></td><td class="r">${w.残業 ? r2(w.残業) : ""}</td></tr>`);
  h.push(`</tbody></table>`);

  const mt = (ttl, m) => {
    if (!m.size) return;
    h.push(`<h3 class="sh">${ttl}</h3><table class="sm"><thead><tr><th>${ttl === "使用機材" ? "使用機材（規格）" : "品名"}</th>${S.dates.map(x => `<th class="${wcls(x)}">${+x.slice(8)}</th>`).join("")}<th>計</th></tr></thead><tbody>`);
    for (const [k, v] of m) h.push(`<tr><td>${esc(k)}</td>${S.dates.map(x => `<td class="r">${v.日[x] ? r2(v.日[x]) : ""}</td>`).join("")}<td class="r"><b>${r2(v.計)}</b></td></tr>`);
    h.push(`</tbody></table>`);
  };
  mt("使用機材", S.kizai); mt("品名", S.hin);
  $("sumBody").innerHTML = h.join("");
}
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
    h.push(`<h3 class="sh">${esc(g.名前)} の現場</h3><div class="mrow"><div class="mbox"><table class="ed"><tr><th>コード・略称</th><th>現場名（日報に出る名前）</th><th>日報の日数</th><th></th></tr>`);
    for (const x of DB.現場.filter(y => y.元請id === g.id)) h.push(`<tr class="${x.id === UI.site ? "on" : ""}"><td class="m">${esc(x.コード || "")}${x.略称 ? " " + esc(x.略称) : ""}</td><td><input data-site="${x.id}" value="${esc(x.名前)}" style="width:360px" title="印刷・Excel に出る現場名"></td><td class="r">${Object.keys(x.日).length}</td><td class="btn"><button class="sub small" data-picksite="${x.id}">選ぶ</button> <button class="sub small danger" data-delsite="${x.id}">消す</button></td></tr>`);
    h.push(`</table><div class="btns"><button id="mAddSite">＋ 現場を追加</button></div></div></div>`);
  }
  if (s) {
    h.push(`<h3 class="sh">${esc(s.名前)} の決まり</h3><div class="mrow"><div class="mbox"><table class="ed">
      <tr><th>記入者（いつもの）</th><td><input data-sf="担当者" value="${esc(s.担当者)}"></td></tr>
      <tr><th>定時の始業</th><td><input data-sf="始業" value="${esc(s.始業)}" inputmode="numeric"></td></tr>
      <tr><th>定時の終業</th><td><input data-sf="終業" value="${esc(s.終業)}" inputmode="numeric"></td></tr>
      <tr><th>締め日</th><td><input data-sf="締め日" value="${esc(s.締め日)}" inputmode="numeric" title="20 → 21日〜翌月20日。0 か 31 → 月末締め"></td></tr>
      <tr><th>累計の数え方</th><td><select data-sf="累計"><option${s.累計 === "期間" ? " selected" : ""}>期間</option><option${s.累計 === "全体" ? " selected" : ""}>全体</option></select></td></tr>
      </table><p class="note">残業時間＝定時の終業より遅い分＋始業より早い分（作業員の終業時間を入れると自動）。<br>累計：期間＝締め期間の中で足す／全体＝この現場のはじめから足す。</p></div>`);
    h.push(`<div class="mbox"><h4>この現場で使った作業員（${s.作業員.length}名・工種を覚える）</h4><table class="ed"><tr><th>作業員名</th><th>工種（職種）</th><th></th></tr>`);
    s.作業員.forEach((w, i) => h.push(`<tr><td><input data-wn="${i}" value="${esc(w.名前)}"></td><td><input data-wk="${i}" value="${esc(w.職種)}" list="dlShoku"></td><td class="btn"><button class="sub small danger" data-delw="${i}">消す</button></td></tr>`));
    h.push(`<tr><td><input id="mNewW" placeholder="名前を足す"></td><td><input id="mNewWk" list="dlShoku" placeholder="工種"></td><td class="btn"><button class="small" id="mAddW">追加</button></td></tr></table>
      <datalist id="dlShoku">${DB.職種.map(k => `<option value="${esc(k)}">`).join("")}</datalist>
      <p class="note">日報の作業員名に打った人は自動でここに足されます。<br>名前の候補は配置表の名簿（マスタ）からも出ます。</p></div>`);
    h.push(`<div class="mbox"><h4>使用機材</h4><table class="ed"><tr><th>使用機材</th><th>規格・形式</th><th></th></tr>`);
    s.機材.forEach((m, i) => h.push(`<tr><td><input data-mn="${i}" value="${esc(m.名前)}"></td><td><input data-mk="${i}" value="${esc(m.規格)}"></td><td class="btn"><button class="sub small danger" data-delm="${i}">消す</button></td></tr>`));
    h.push(`<tr><td><input id="mNewM" placeholder="機材を足す"></td><td><input id="mNewMk" placeholder="規格"></td><td class="btn"><button class="small" id="mAddM">追加</button></td></tr></table></div></div>`);
  }
  if (s) h.push(`<h3 class="sh">揺らぎを見つけてそろえる（${esc(siteLabel(s))}）</h3><div class="mbox"><p class="note">作業員名・工種・使用機材・規格・品名で、同じものらしいのに書き方がちがうもの（全角半角・空白・1字ちがい など）を探します。そろえる書き方を選ぶと、この現場の日報をまとめて直します。</p><button class="sub" id="mYure">揺らぎを探す</button><div id="yureList"></div></div>`);
  h.push(`<h3 class="sh">工種（職種）・天気の一覧（全部の現場で共通）</h3><div class="mrow"><div class="mbox"><h4>工種（職種）</h4><textarea id="mShoku" rows="6" cols="22">${esc(DB.職種.join("\n"))}</textarea><p class="note">1行に1つ。この順で集計に並びます。</p></div>
    <div class="mbox"><h4>天気</h4><textarea id="mWx" rows="6" cols="16">${esc(DB.天気.join("\n"))}</textarea></div></div>`);
  h.push(`<h3 class="sh">データ</h3><div class="mbox"><p class="note">日報はこのブラウザの中に保存しています（ほかのパソコン・iPad とは自動で共有しません）。<br>上の「保存」でこの現場をファイル（.json）に出し、ほかの端末で「取込」すると移せます。「取込」は元の作業日報の Excel（.xlsx）も読めます。</p>
    <button class="sub" id="mSaveAll">全部の現場を保存（.json）</button></div>`);
  $("masterBody").innerHTML = h.join("");
}
$("masterBody").addEventListener("change", e => {
  const t = e.target, s = site(), dsx = t.dataset;
  if ((dsx.sf || dsx.wn || dsx.wk || dsx.mn || dsx.mk || dsx.site === UI.site) && !canEdit()) { renderMaster(); return; }
  if (dsx.gen) { DB.元請.find(x => x.id === dsx.gen).名前 = t.value.trim(); }
  else if (dsx.site) { DB.現場.find(x => x.id === dsx.site).名前 = t.value.trim(); }
  else if (dsx.sf) { let v = t.value.trim(); if (dsx.sf === "始業" || dsx.sf === "終業") v = normTime(v); if (dsx.sf === "締め日") v = Number(han(v)) || 0; s[dsx.sf] = v; t.value = v; }
  else if (dsx.wn) s.作業員[+dsx.wn].名前 = t.value.trim();
  else if (dsx.wk) s.作業員[+dsx.wk].職種 = t.value.trim();
  else if (dsx.mn) s.機材[+dsx.mn].名前 = t.value.trim();
  else if (dsx.mk) s.機材[+dsx.mk].規格 = t.value.trim();
  else if (t.id === "mShoku") DB.職種 = t.value.split(/\n/).map(x => x.trim()).filter(Boolean);
  else if (t.id === "mWx") DB.天気 = t.value.split(/\n/).map(x => x.trim()).filter(Boolean);
  else return;
  if (s && (dsx.sf || dsx.wn || dsx.wk || dsx.mn || dsx.mk)) s.設定更新 = now();
  if (dsx.site) DB.現場.find(x => x.id === dsx.site).設定更新 = now();
  save(); renderPickers();
});
$("masterBody").addEventListener("click", e => {
  const t = e.target.closest("button"); if (!t) return;
  const s = site(), dsx = t.dataset;
  if ((dsx.delw != null || dsx.delm != null || t.id === "mAddW" || t.id === "mAddM" || dsx.delsite === UI.site) && !canEdit()) return;
  if (t.id === "mAddGen") return editDlg("元請業者を追加", [["元請業者名", "name", ""]], v => { if (!v.name) return; const g = { id: uid(), 名前: v.name }; DB.元請.push(g); UI.gen = g.id; UI.site = null; save(); renderAll(); });
  if (t.id === "mAddSite") return editDlg("現場を追加（" + gen().名前 + "）", [["現場名", "name", ""], ["記入者", "tan", ""], ["締め日", "shime", "20"]], v => { if (!v.name) return; const x = fixSite({ id: uid(), 元請id: UI.gen, 名前: v.name, 担当者: v.tan, 締め日: Number(han(v.shime)) || 0 }); DB.現場.push(x); UI.site = x.id; save(); renderAll(); });
  if (dsx.pickgen) { UI.gen = dsx.pickgen; UI.site = null; renderAll(); return; }
  if (dsx.picksite) { UI.site = dsx.picksite; renderAll(); return; }
  if (dsx.delgen) { if (DB.現場.some(x => x.元請id === dsx.delgen)) { toast("現場がある元請業者は消せません（先に現場を消す）"); return; } if (!confirm("この元請業者を消しますか？")) return; DB.元請 = DB.元請.filter(x => x.id !== dsx.delgen); save(); renderAll(); return; }
  if (dsx.delsite) { const x = DB.現場.find(y => y.id === dsx.delsite); if (!confirm(`「${x.名前}」と日報 ${Object.keys(x.日).length} 日分を消しますか？\n共有している日報（OneDrive）も消えます（OneDrive のごみ箱から戻せます）。\n（消す前に「保存」でファイルに出しておくと安心です）`)) return; DB.現場 = DB.現場.filter(y => y.id !== dsx.delsite); removeShared(x.id); save(); renderAll(); return; }
  if (dsx.delw != null) { s.作業員.splice(+dsx.delw, 1); s.設定更新 = now(); save(); renderMaster(); return; }
  if (dsx.delm != null) { s.機材.splice(+dsx.delm, 1); s.設定更新 = now(); save(); renderMaster(); return; }
  if (t.id === "mAddW") { const n = $("mNewW").value.trim(); if (!n) return; s.作業員.push({ 名前: n, 職種: $("mNewWk").value.trim() }); s.設定更新 = now(); save(); renderMaster(); $("mNewW").focus(); return; }
  if (t.id === "mAddM") { const n = $("mNewM").value.trim(); if (!n) return; s.機材.push({ 名前: n, 規格: $("mNewMk").value.trim() }); s.設定更新 = now(); save(); renderMaster(); $("mNewM").focus(); return; }
  if (t.id === "mReload") { loadMasters(false); return; }
  if (t.id === "mYure") { findYure(); return; }
  if (t.id === "yureFix") { fixYure(); return; }
  if (t.id === "mSaveAll") { download(`常用日報_全部_${ymd(new Date()).replace(/-/g, "")}.json`, JSON.stringify({ ...DB, 保存日時: now() }, null, 1), "application/json"); }
});
// この現場の日報の中の揺らぎ：列ごとに、似ている書き方をまとめる（マスタの書き方・多く使われている書き方をはじめに選ぶ）
let YG = [];
function findYure() {
  const s = site(); if (!s) return;
  YG = [];
  for (const key of ["dn", "up"]) for (const c of Object.keys(YCOLS[key]).map(Number)) {
    const used = new Map();
    for (const dd of Object.values(s.日)) (key === "up" ? dd.上 : dd.下).forEach(r => { const v = r && String(r[c] || "").trim(); if (v) used.set(v, (used.get(v) || 0) + 1); });
    const cn = canon(key, c, s), ws = [...used.keys()], seen = new Set();
    for (const w of ws) {
      if (seen.has(w)) continue;
      const g = [w, ...ws.filter(x => x !== w && !seen.has(x) && similar(w, x))];
      // マスタにあって日報に無い書き方も候補に（マスタの書き方にそろえられるように）
      [...cn.keys()].filter(x => cn.get(x) >= 1000 && !g.includes(x) && g.some(y => similar(y, x))).forEach(x => g.push(x));
      g.forEach(x => seen.add(x));
      if (g.length < 2) continue;
      const best = [...g].sort((a, b) => (cn.get(b) || 0) - (cn.get(a) || 0))[0];
      YG.push({ key, c, g: g.map(x => ({ w: x, n: used.get(x) || 0, m: (cn.get(x) || 0) >= 1000 })), best });
    }
  }
  $("yureList").innerHTML = !YG.length ? `<p class="note">揺らぎは見つかりませんでした。</p>` :
    `<table class="sm"><thead><tr><th>欄</th><th>書き方（日報の中の回数）</th><th>そろえる</th></tr></thead><tbody>` +
    YG.map((x, i) => `<tr><td>${esc(YCOLS[x.key][x.c])}</td><td>${x.g.map(y => `<label class="yr"><input type="radio" name="yg${i}" value="${esc(y.w)}"${y.w === x.best ? " checked" : ""}> ${esc(y.w)} <span class="note">${y.n}回${y.m ? "・マスタ" : ""}</span></label>`).join("")}</td>` +
      `<td class="c"><input type="checkbox" class="ygOn" data-i="${i}" checked></td></tr>`).join("") +
    `</tbody></table><div class="btns"><button id="yureFix">チェックした行をそろえる</button><span class="note">別のものならチェックを外す</span></div>`;
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
    for (const p of plan) (p.key === "up" ? dd.上 : dd.下).forEach(r => { if (r && p.from.has(String(r[p.c] || "").trim())) { r[p.c] = p.to; n++; ch = true; } });
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
function editDlg(title, fields, ok) {
  $("eTtl").textContent = title;
  $("eBody").innerHTML = `<div class="grid">${fields.map(([l, k, v]) => `<label>${esc(l)}</label><input data-k="${k}" value="${esc(v)}" style="width:320px">`).join("")}</div>
    <div class="btns"><span class="sp"></span><button type="button" class="sub" data-close>やめる</button><button type="button" id="eOk">決める</button></div>`;
  openDlg($("dEdit"));
  $("eBody").querySelector("input").focus();
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
  d.下.forEach(r => {
    if (!r[0]) return;
    r[0] = masterName(r[0]);
    if (!s.作業員.some(w => nk(w.名前) === nk(r[0]))) s.作業員.push({ 名前: r[0], 職種: r[3] });
    if (r[3] && !DB.職種.includes(r[3])) DB.職種.push(r[3]);
    if (r[5] && !s.機材.some(m => m.名前 === r[5] && m.規格 === r[6])) s.機材.push({ 名前: r[5], 規格: r[6] });
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
function daySheet(wb, s, x, d) {
  const dt = pd(x);
  const ws = wb.addWorksheet(`作業日報${dt.getFullYear()}.${dt.getMonth() + 1}.${dt.getDate()}`, {
    pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true,
      margins: { left: 0.55, right: 0.55, top: 0.6, bottom: 0.5, header: 0.3, footer: 0.3 } },
    views: [{ showGridLines: false }]
  });
  const W = [9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 4.63, 9, 9, 9];
  ws.columns = W.map(w => ({ width: w + 0.71 }));
  const nDn = Math.max(DN_ROWS, d.下.length), last = 17 + nDn;
  const font = { name: "ＭＳ Ｐゴシック", size: 11 };
  for (let r = 1; r <= last; r++) for (let c = 1; c <= 16; c++) { const cl = ws.getCell(r, c); cl.font = font; cl.alignment = { vertical: "middle", horizontal: "center", shrinkToFit: true }; }
  ws.getRow(1).height = 27;
  [3, 4, 5, 6].forEach(r => ws.getRow(r).height = 21);
  for (let r = 18; r <= last; r++) ws.getRow(r).height = 21;
  const put = (a, v, m, extra) => { if (m) ws.mergeCells(m); const cl = ws.getCell(a); cl.value = v; if (extra) Object.assign(cl, extra); return cl; };
  put("A1", "作　業　日　報", "A1:P1").font = { ...font, size: 20 };
  put("A3", "元請業者"); put("B3", gen() ? gen().名前 : "", "B3:D3");
  put("E3", "日　　付", "E3:F3"); put("G3", dt, "G3:H3").numFmt = 'yyyy"年"m"月"d"日"';
  put("I3", WD[dt.getDay()]); put("J3", "天気"); put("K3", "気温");
  put("L3", "記入者", "L3:L4"); put("M3", d.担当者 || s.担当者 || "", "M3:P4");
  put("A4", "現場名"); put("B4", s.名前, "B4:I4"); put("J4", d.天気 || ""); put("K4", d.気温 !== "" && d.気温 != null ? d.気温 + "℃" : "℃");
  ws.getCell("K4").alignment = { vertical: "middle", horizontal: "right" };
  put("A5", "工種", "A5:A6"); put("B5", "作業内容", "B5:C6"); put("D5", "人員", "D5:D6");
  put("E5", "使　用　資　機　材　・　重　機", "E5:K6"); put("L5", "元請サイン", "L5:L6"); ws.mergeCells("M5:P6");
  for (let i = 0; i < UP_ROWS; i++) {
    const r = 7 + i * 2, row = d.上[i] || [];
    put("A" + r, row[0] || "", `A${r}:A${r + 1}`); put("B" + r, row[1] || "", `B${r}:C${r + 1}`);
    put("D" + r, row[2] || "", `D${r}:D${r + 1}`).alignment = { vertical: "middle", horizontal: "right" };
    put("E" + r, row[3] || "", `E${r}:J${r + 1}`); ws.mergeCells(`K${r}:K${r + 1}`);
  }
  put("L7", "備考"); ws.mergeCells("M7:P7");
  const bk = put("L8", d.備考 || "", "L8:P15"); bk.alignment = { vertical: "top", horizontal: "left", wrapText: true };
  put("L16", "赤丸は取決め以外", "L16:P16").font = { ...font, size: 10 };
  const H = [["A17", "作業員名", "A17:B17"], ["C17", "始業時間"], ["D17", "終業時間"], ["E17", "工種"], ["F17", "残業時間"], ["G17", "使用機材", "G17:H17"], ["I17", "規格・形式"], ["J17", "数量"], ["K17", "累計"], ["L17", "品名"], ["M17", "数量", "M17:N17"], ["O17", "累計数量", "O17:P17"]];
  H.forEach(([a, v, m]) => put(a, v, m));
  // 累計は画面と同じ計算（この日の現場・日を一時的に当てる）
  const keep = UI.day; UI.day = x;
  for (let i = 0; i < nDn; i++) {
    const r = 18 + i, row = d.下[i] || [];
    const num = v => v === "" || v == null ? "" : isNaN(Number(v)) ? v : Number(v);
    put("A" + r, row[0] || "", `A${r}:B${r}`); put("C" + r, row[0] ? row[1] || "" : ""); put("D" + r, row[0] ? row[2] || "" : "");
    put("E" + r, row[3] || ""); const ot = put("F" + r, num(row[4])); if (row[4]) ot.font = { ...font, color: { argb: "FFC00000" } };
    put("G" + r, row[5] || "", `G${r}:H${r}`); put("I" + r, row[6] || ""); put("J" + r, num(row[7])); put("K" + r, num(row[5] ? getCellFor(d, i, 8) : ""));
    put("L" + r, row[9] || ""); put("M" + r, num(row[10]), `M${r}:N${r}`); put("O" + r, num(row[9] ? getCellFor(d, i, 11) : ""), `O${r}:P${r}`);
  }
  UI.day = keep;
  for (let r = 3; r <= last; r++) for (let c = 1; c <= 16; c++) {
    if (r === 16 && c >= 12) continue;
    ws.getCell(r, c).border = BORDER;
  }
  if (d.サイン) {
    const id = wb.addImage({ base64: d.サイン, extension: "png" });
    ws.addImage(id, { tl: { col: 12.15, row: 4.1 }, br: { col: 15.85, row: 5.9 }, editAs: "oneCell" });
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
  const h1 = ["日付", "曜", "天気", "気温", ...S.shoku, "人員計", "残業計(h)"];
  h1.forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
  const top = r;
  for (const p of S.perDay) {
    const vals = [pd(p.date), WD[wd(p.date)], p.天気, p.気温 === "" ? "" : Number(p.気温), ...S.shoku.map(k => p.c.get(k) || null), p.n, p.ot || null];
    vals.forEach((v, i) => { const cl = cellAt(r, i + 1, v, { h: i >= 3 ? "right" : "center" }); if (i === 0) cl.numFmt = "m/d"; });
    r++;
  }
  cellAt(r, 1, "計", { fill: "FFF5F5F5" }); [2, 3, 4].forEach(c => cellAt(r, c, "", { fill: "FFF5F5F5" }));
  for (let i = 5; i <= h1.length; i++) { const col = ws.getColumn(i).letter; cellAt(r, i, { formula: `SUM(${col}${top}:${col}${r - 1})` }, { h: "right", fill: "FFF5F5F5", font: { bold: true } }); }
  r += 2;
  title("作業員別（○＝出面、数字＝残業時間）");
  const h2 = ["作業員名", "工種", ...S.dates.map(x => +x.slice(8)), "日数", "残業計(h)"];
  h2.forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
  for (const w of S.ws) {
    cellAt(r, 1, w.名前, { h: "left" }); cellAt(r, 2, w.職種);
    S.dates.forEach((x, i) => cellAt(r, 3 + i, x in w.日 ? (w.日[x] ? w.日[x] : "○") : "", { font: x in w.日 && w.日[x] ? { color: { argb: "FFC00000" } } : {} }));
    cellAt(r, 3 + S.dates.length, w.日数, { h: "right" }); cellAt(r, 4 + S.dates.length, w.残業 || null, { h: "right" });
    r++;
  }
  const mt = (ttl, m) => {
    if (!m.size) return;
    r += 1; title(ttl);
    [ttl, ...S.dates.map(x => +x.slice(8)), "計"].forEach((t, i) => cellAt(r, i + 1, t, { fill: "FFEEF3F8" })); r++;
    for (const [k, v] of m) { cellAt(r, 1, k, { h: "left" }); S.dates.forEach((x, i) => cellAt(r, 2 + i, v.日[x] || null, { h: "right" })); cellAt(r, 2 + S.dates.length, v.計, { h: "right", font: { bold: true } }); r++; }
  };
  mt("使用機材", S.kizai); mt("品名", S.hin);
  ws.getColumn(1).width = 16; ws.getColumn(2).width = 9;
  for (let c = 3; c <= Math.max(h1.length, h2.length) + 2; c++) ws.getColumn(c).width = 6.5;
  ws.getColumn(3).width = 8; ws.getColumn(4).width = 7;
}

/* ================= PDF（印刷。A4 横で1日1ページ） ================= */
function dayHTML(s, x, d) {
  const dt = pd(x), keep = UI.day; UI.day = x;
  const c = (v, cls = "", span = "") => `<td class="${cls}"${span}>${esc(v)}</td>`;
  const sp = (cs, rs) => (cs > 1 ? ` colspan="${cs}"` : "") + (rs > 1 ? ` rowspan="${rs}"` : "");
  const up = i => d.上[i] || [];
  const h = [];
  h.push(`<div class="pg"><table class="pt"><colgroup>${[...Array(16)].map((_, i) => `<col style="width:${i === 12 ? 9.2 : 17.86}mm">`).join("")}</colgroup>`);
  h.push(`<tr><td class="t" colspan="16">作　業　日　報</td></tr><tr class="b"><td class="nb" colspan="16"></td></tr>`);
  h.push(`<tr class="a">${c("元請業者")}${c(gen() ? gen().名前 : "", "", sp(3))}${c("日　　付", "", sp(2))}${c(jpDate(x), "", sp(2))}${c(WD[dt.getDay()])}${c("天気")}${c("気温")}${c("記入者", "", sp(1, 2))}${c(d.担当者 || s.担当者 || "", "", sp(4, 2))}</tr>`);
  h.push(`<tr class="a">${c("現場名")}${c(s.名前, "", sp(8))}${c(d.天気 || "")}${c((d.気温 ?? "") + "℃")}</tr>`);
  // 上の5段（Excel は2行ずつ）。空の行は高さがつぶれるので、1段＝高い1行で組む
  h.push(`<tr class="a2">${c("工種")}${c("作業内容", "", sp(2))}${c("人員")}${c("使　用　資　機　材　・　重　機", "", sp(7))}${c("元請サイン")}<td class="s"${sp(4)}>${d.サイン ? `<img src="${d.サイン}">` : ""}</td></tr>`);
  const slot = i => `${c(up(i)[0] || "")}${c(up(i)[1] || "", "", sp(2))}<td style="text-align:right">${esc(up(i)[2] || "")}</td>${c(up(i)[3] || "", "", sp(6))}<td></td>`;
  h.push(`<tr class="a2">${slot(0)}${c("備考")}<td${sp(4)}></td></tr>`);
  h.push(`<tr class="a2">${slot(1)}<td class="bk"${sp(5, 3)}>${esc(d.備考 || "")}</td></tr>`);
  h.push(`<tr class="a2">${slot(2)}</tr><tr class="a2">${slot(3)}</tr>`);
  h.push(`<tr class="a2">${slot(4)}<td class="sm"${sp(5)}>赤丸は取決め以外</td></tr>`);
  h.push(`<tr class="b2">${c("作業員名", "", sp(2))}${c("始業時間")}${c("終業時間")}${c("工種")}${c("残業時間")}${c("使用機材", "", sp(2))}${c("規格・形式")}${c("数量")}${c("累計")}${c("品名")}${c("数量", "", sp(2))}${c("累計数量", "", sp(2))}</tr>`);
  const n = Math.max(DN_ROWS, d.下.length);
  for (let i = 0; i < n; i++) {
    const r = d.下[i] || [];
    h.push(`<tr class="a">${c(r[0] || "", "", sp(2))}${c(r[0] ? r[1] || "" : "")}${c(r[0] ? r[2] || "" : "")}${c(r[3] || "")}${c(r[4] || "", "ot")}${c(r[5] || "", "", sp(2))}${c(r[6] || "")}${c(r[7] || "")}${c(r[5] ? cumul("kizai", i) : "")}${c(r[9] || "")}${c(r[10] || "", "", sp(2))}${c(r[9] ? cumul("hin", i) : "", "", sp(2))}</tr>`);
  }
  h.push(`</table>${d.サイン日時 ? `<div class="ft">サイン ${esc(d.サイン日時)}</div>` : ""}</div>`);
  UI.day = keep;
  return h.join("");
}
function printDays(dates, sumRange) {
  const s = site(); if (!s) { toast("現場を選んでください"); return; }
  const h = dates.map(x => dayHTML(s, x, s.日[x] || { 上: [], 下: [] }));
  if (sumRange) {
    const keepF = $("sumFrom").value, keepT = $("sumTo").value;
    $("sumFrom").value = sumRange[0]; $("sumTo").value = sumRange[1]; renderSum();
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
buildSheet();
showTab(DB.現場.length ? (UI.tab || "sheet") : "master");
syncMasterSites();
renderAll();
showMe(); setRO(null);
startAuth();
if (!ME.名前) askName();
if (!DB.現場.length) toast("「サインイン」すると現場一覧・名簿を読みます。元の作業日報の Excel は「取込」で読めます");
