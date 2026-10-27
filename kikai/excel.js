/* ---------- Excel（月ごとの自社機械予定表）と合わせる（2026-10-27） ----------
   まだ皆が Excel に入れるので、画面が 03_自社機械稼働表\第NN期\R??年??月度自社機械予定表.xlsx を数分ごとに見て、
   ・Excel が変わったら、その月を取り込む（Excel に書いてある日は Excel が勝つ。前に読んだときにあって今は無い日は消す）
   ・画面にあって Excel に無い日（使う日）は、Excel のその枠に現場名を書く
   前に読んだ Excel の中身（控え）を D.Excel に持ち、控えと比べて「Excel で消した」のか「画面で足した」のかを見分ける。
     D.Excel = { 元: { "第52期/R08年10月度自社機械予定表.xlsx": { 月, cTag, シート, 行: { "名前|規格": 昼の行 } } },
                 控え: { "2026-10": { "名前|規格": { 昼: { 日: 現場 }, 夜: {…} } } } }
   読み方は tools\Excelから取り込む.py の read_month と同じ（1台＝昼・夜の2行、A＝機械名・B＝規格、D列から日付）。
   ?demo のとき（OneDrive を使わない）は動かない。 */
const XL_POLL_MS = 3 * 60 * 1000;
const XL_WRITE_MAX = 200;   // 1回に Excel へ書く枠の数（多いときは次の回へ）
const XL_ALIAS = { "ランマ|": ["ランマ1", ""], "プレート|": ["プレート1", ""], "エンジンハイウォッシャー|": ["ハイウォッシャー", "エンジン"] };
const XL_SITE_ALIAS = { "車 検": "車検", "積み置中": "積置き" };
const XL_KIND = { "機材センター": "置場", "機材センター(仙台)": "置場", "積置き": "置場", "車両ヤード": "置場", "資材置き場": "置場",
  "故障又は修理中": "修理", "整備(リコール)": "修理", "特定自主検査": "検査", "車検": "検査", "クレーン検査": "検査", "名義変更": "検査" };
const nfkc = v => String(v ?? "").normalize("NFKC").trim();
const mKey = m => m.名前 + "|" + (m.規格 || "");
const colName = c => { let s = ""; for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s; };
const colNum = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const monthDays = M => new Date(+M.slice(0, 4), +M.slice(5, 7), 0).getDate();
const xlBunrui = name => ([["ダンプ", "車両"], ["バキューム", "車両"], ["ユニック", "車両"], ["ラフター", "車両"], ["バックホウ", "重機"],
  ["バケット", "重機"], ["アーム", "重機"], ["光波", "測量"], ["レベル", "測量"]].find(([k]) => name.includes(k)) || [0, "小型機械"])[1];

/* 1か月の表（式。無い所は値）を読む。F・V は A1 から始まる2次元の配列 */
function xlParse(F, V) {
  const a1 = V[0]?.[0];
  if (typeof a1 !== "number") throw new Error("A1 が日付ではありません");
  const d0 = new Date(1899, 11, 30); d0.setDate(d0.getDate() + Math.floor(a1));
  const M = keyOf(d0).slice(0, 7), n = monthDays(M);
  const cell = (r, c) => F[r - 1]?.[c - 1];   // Excel の行・列（1 から）
  const rows = {}, cells = {}, meta = {};
  for (let r = 3; r <= F.length;) {
    if (nfkc(cell(r, 3)) !== "昼" || nfkc(cell(r + 1, 3)) !== "夜") { r++; continue; }
    let name = nfkc(cell(r, 1)), spec = nfkc(cell(r, 2)), memo = "";
    const spec2 = nfkc(cell(r + 1, 2)), under = nfkc(cell(r + 1, 1));
    if (under) { if (spec) memo = under; else spec = under; }
    if (spec2) spec = (spec + " " + spec2).trim();
    [name, spec] = XL_ALIAS[name + "|" + spec] || [name, spec];
    const key = name + "|" + spec;
    if (!(key in rows)) rows[key] = r;
    meta[key] = { 名前: name, 規格: spec, メモ: memo };
    const c = cells[key] = cells[key] || { 昼: {}, 夜: {} };
    [["昼", r], ["夜", r + 1]].forEach(([kb, rr]) => {
      for (let i = 1; i <= n; i++) {
        const v = cell(rr, 3 + i);
        if (v == null || v === "" || (typeof v === "string" && (!v.trim() || v.startsWith("=")))) continue;
        const s = nfkc(v); c[kb][i] = XL_SITE_ALIAS[s] || s;
      }
    });
    r += 2;
  }
  return { 月: M, n, rows, cells, meta };
}

/* その機械・区分の、その期間に予約で使う日 → [現場…] */
function laneDays(d, mid, kb, from, to, fn = planned) {
  const A = {};
  d.予約.forEach(r => {
    if (r.機械 !== mid || r.区分 !== kb || r.終了 < from || r.開始 > to) return;
    for (let k = r.開始 < from ? from : r.開始; k <= to && k <= r.終了; k = dayAdd(k, 1)) if (fn(r, k)) (A[k] = A[k] || []).push(r.現場);
  });
  return A;
}
function xlMachine(d, key, meta, create) {
  const m = d.機械.find(x => mKey(x) === key);
  if (m || !create) return m?.id;
  const id = newId("m");
  d.機械.push({ id, 名前: meta.名前, 規格: meta.規格, 分類: xlBunrui(meta.名前), 単価昼: null, 単価夜: null, メモ: meta.メモ || "", 使わない: false });
  return id;
}
function xlEnsureSite(d, name) {
  if (!name || d.現場.some(s => s.名前 === name)) return;
  const kind = XL_KIND[name] || "現場";
  if (kind === "現場") ensureSite(d, name);
  else d.現場.push({ 名前: name, 種類: kind, 色: KIND_COLOR[kind], コード: "", 完了: false });
}

/* その月に予約で使う日を、控えと同じ形（"名前|規格" → 区分 → 日 → 現場）で */
function cellsOf(d, M) {
  const n = monthDays(M), from = M + "-01", to = M + "-" + pad2(n), out = {};
  d.機械.forEach(m => ["昼", "夜"].forEach(kb => Object.entries(laneDays(d, m.id, kb, from, to)).forEach(([k, ss]) => {
    const c = out[mKey(m)] = out[mKey(m)] || { 昼: {}, 夜: {} }; c[kb][+k.slice(8)] = ss[0];
  })));
  return out;
}
/* 1か月分の Excel を取り込む（commit の中で使う。ぶつかって掛け直しても同じ結果になるように）。
   seed＝控えがまだ一つも無いとき：いまの画面を「前に読んだ Excel」とみなす（Excel を丸ごと正にする。画面だけの日を Excel に書き戻さない） */
function xlApply(d, P, f, seed) {
  const X = d.Excel = d.Excel || { 元: {}, 控え: {} };
  if (X.元[f.name]?.cTag === f.cTag) return;   // ほかの画面がもう取り込んだ
  const M = P.月, from = M + "-01", to = M + "-" + pad2(P.n), B = X.控え[M] || (seed ? cellsOf(d, M) : {});
  new Set([...Object.keys(P.cells), ...Object.keys(B)]).forEach(key => {
    const mid = xlMachine(d, key, P.meta[key], !!P.cells[key]); if (!mid) return;
    ["昼", "夜"].forEach(kb => {
      const E = P.cells[key]?.[kb] || {}, Bk = B[key]?.[kb] || {}, A = laneDays(d, mid, kb, from, to), Fm = {};
      let same = true;
      for (let i = 1; i <= P.n; i++) {
        const k = M + "-" + pad2(i), a = A[k] || [];
        const v = E[i] ? [E[i]] : a.filter(s => s !== Bk[i]);   // Excel にあれば Excel。無ければ、Excel で消した日だけ消す
        if (v.length) Fm[k] = v;
        if (v.length !== a.length || v.some(s => !a.includes(s))) same = false;
      }
      Object.values(E).forEach(s => xlEnsureSite(d, s));
      if (!same) xlRebuild(d, mid, kb, from, to, Fm);
    });
  });
  X.控え[M] = P.cells;
  X.元[f.name] = { 月: M, cTag: f.cTag, シート: f.sheet || "", 行: P.rows };
}
/* その機械・区分の予約を、その月は Fm（日 → [現場]）、月の外は今のまま、で作り直す。
   同じ現場が続く日（あいだが土日祝だけなら続き）を1本にし、前の予約の id・状態・メモを引き継ぐ */
function xlRebuild(d, mid, kb, from, to, Fm) {
  const old = d.予約.filter(r => r.機械 === mid && r.区分 === kb && r.開始 <= to && r.終了 >= from);
  const days = {};
  const put = (s, k) => (days[s] = days[s] || new Set()).add(k);
  old.forEach(r => { for (let k = r.開始; k <= r.終了; k = dayAdd(k, 1)) if ((k < from || k > to) && planned(r, k)) put(r.現場, k); });
  Object.entries(Fm).forEach(([k, ss]) => ss.forEach(s => put(s, k)));
  d.予約 = d.予約.filter(r => !old.includes(r));
  const used = new Set(), now = new Date().toLocaleString("ja-JP") + " Excel から";
  Object.entries(days).forEach(([site, set]) => {
    const runs = [];
    [...set].sort().forEach(k => {
      const run = runs[runs.length - 1];
      if (run) {
        let g = dayAdd(run[run.length - 1], 1);
        while (g < k && (isSat(g) || isSunHol(g))) g = dayAdd(g, 1);
        if (g === k) { run.push(k); return; }
      }
      runs.push([k]);
    });
    runs.forEach(run => {
      const a = run[0], b = run[run.length - 1], inRun = new Set(run);
      const sat = run.some(isSat), sun = run.some(isSunHol), 休み = [];
      for (let k = a; k <= b; k = dayAdd(k, 1)) if (!inRun.has(k) && (isSunHol(k) ? sun : isSat(k) ? sat : true)) 休み.push(k);
      let best = null, bn = 0;
      old.forEach(r => { if (r.現場 !== site) return; const n = run.filter(k => planned(r, k)).length; if (n > bn) { best = r; bn = n; } });
      const id = best && !used.has(best.id) ? best.id : newId("r"); used.add(id);
      const same = best && best.開始 === a && best.終了 === b && best.土 === sat && best.日祝 === sun && !best.出勤.length && best.休み.join() === 休み.join();
      d.予約.push({ id, 機械: mid, 区分: kb, 現場: site, 開始: a, 終了: b, 土: sat, 日祝: sun, 休み, 出勤: [],
        未使用: best ? best.未使用.filter(k => inRun.has(k)) : [], 状態: best ? best.状態 : "確定", メモ: best ? best.メモ : "",
        ...(same ? (best.更新 ? { 更新: best.更新 } : {}) : { 更新: now }) });
    });
  });
}

/* 画面にあって Excel（控え）に無い日 → Excel に書く枠。Excel に行の無い機械は書けないので数えるだけ */
function xlToWrite(d) {
  const X = d.Excel, out = [], noRow = new Set();
  if (!X) return { out, noRow };
  Object.entries(X.元).forEach(([name, f]) => {
    const M = f.月, n = monthDays(M), from = M + "-01", to = M + "-" + pad2(n), B = X.控え[M] || {};
    d.機械.forEach(m => {
      const key = mKey(m), row = f.行[key];
      ["昼", "夜"].forEach((kb, j) => {
        const A = laneDays(d, m.id, kb, from, to, works), Bk = B[key]?.[kb] || {};
        for (let i = 1; i <= n; i++) {
          const a = A[M + "-" + pad2(i)];
          if (!a || Bk[i]) continue;
          if (!row) { noRow.add(mLabel(m) + "（" + M.replace("-", "/") + "）"); continue; }
          out.push({ name, sheet: f.シート, M, key, kb, i, row: row + j, col: 3 + i, site: a[0] });
        }
      });
    });
  });
  return { out, noRow };
}

/* ---------- OneDrive（Graph）で読む・書く ---------- */
let xlBusy = false, xlRoot = null, xlDrive = "";
const xlState = { いつ: "", 取り込み: [], 書いた: 0, 書けない: [], err: "" };
async function xlJson(path, what) { const r = await g(path); if (!r.ok) throw httpErr(r, what); return r.json(); }
async function xlList() {   // 03_自社機械稼働表（稼働データ フォルダの親）の 第NN期\*自社機械予定表.xlsx
  if (!xlRoot) {
    const p = (await xlJson(folderBase + "?$select=parentReference", "Excel のフォルダ")).parentReference;
    xlDrive = p.driveId; xlRoot = `/drives/${xlDrive}/items/${p.id}`;
  }
  const kids = async p => (await xlJson(p + "/children?$select=id,name,folder,file,cTag&$top=999", "Excel のフォルダ")).value;
  const files = [];
  for (const f of await kids(xlRoot)) {
    if (!f.folder || !/^第\d+期$/.test(f.name)) continue;
    for (const x of await kids(`/drives/${xlDrive}/items/${f.id}`))
      if (x.file && /自社機械予定表\.xlsx$/.test(x.name) && !x.name.startsWith("~$")) files.push({ name: f.name + "/" + x.name, id: x.id, cTag: x.cTag });
  }
  return files;
}
const xlBook = id => `/drives/${xlDrive}/items/${id}/workbook`;
async function xlRead(f) {
  const ws = (await xlJson(xlBook(f.id) + "/worksheets?$select=id,name,position", "Excel のシート")).value.sort((a, b) => a.position - b.position)[0];
  f.sheet = ws.id;
  const u = await xlJson(xlBook(f.id) + "/worksheets/" + encodeURIComponent(ws.id) + "/usedRange?$select=address,formulas,values", "Excel の中身");
  const m = /\$?([A-Z]+)\$?(\d+)/.exec(u.address.split("!").pop()), c0 = colNum(m[1]) - 1, r0 = +m[2] - 1;
  const shift = a => Array.from({ length: r0 }, () => []).concat(a.map(row => Array(c0).fill("").concat(row)));   // A1 から始まるようにずらす
  return xlParse(shift(u.formulas), shift(u.values));
}
async function xlWrite(w, id) {
  const r = await g(xlBook(id) + "/worksheets/" + encodeURIComponent(w.sheet) + `/range(address='${colName(w.col)}${w.row}')`,
    { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ values: [[w.site]] }) });
  if (!r.ok) throw httpErr(r, "Excel への書き込み");
}
async function xlSync() {
  if (DEMO || xlBusy || !ready || pending.length || saving || document.hidden) return;
  xlBusy = true;
  try {
    const files = await xlList(), took = [], seed = !D.Excel;
    for (const f of files) {
      if (D.Excel?.元?.[f.name]?.cTag === f.cTag) continue;
      const P = await xlRead(f);
      commit(d => xlApply(d, P, f, seed && !d.Excel?.元?.[f.name]), { noUndo: true });
      took.push(f.name.replace(/^.*\/|自社機械予定表\.xlsx$/g, ""));
    }
    if (took.length) { xlState.取り込み = took; toast("Excel（" + took.join("・") + "）の変更を取り込みました"); }
    /* 取り込んだばかりの（いまの Excel と控えが同じ）月にだけ書く */
    const ids = new Map(files.map(f => [f.name, f])), { out, noRow } = xlToWrite(D);
    const ws = out.filter(w => ids.get(w.name)?.cTag === D.Excel.元[w.name].cTag && w.sheet).slice(0, XL_WRITE_MAX);
    const done = [];
    for (const w of ws) { await xlWrite(w, ids.get(w.name).id); done.push(w); }
    if (done.length) {
      commit(d => done.forEach(w => { const c = ((d.Excel.控え[w.M] = d.Excel.控え[w.M] || {})[w.key] = d.Excel.控え[w.M][w.key] || { 昼: {}, 夜: {} }); c[w.kb][w.i] = c[w.kb][w.i] || w.site; }), { noUndo: true });
      xlState.書いた += done.length;
      toast(`画面だけにあった ${done.length}日分を Excel に書きました`);
    }
    xlState.書けない = [...noRow]; xlState.いつ = new Date().toLocaleTimeString("ja-JP"); xlState.err = "";
  } catch (e) {
    console.error(e);
    if (!xlState.err) toast("Excel と合わせられませんでした：" + (e.status === 403 || e.status === 404 ? "03_自社機械稼働表 フォルダを開けません" : e.message));
    xlState.err = e.message;
  } finally { xlBusy = false; }
}
/* 「データ」の小窓に出す */
function xlNote() {
  if (DEMO) return "";
  const s = xlState;
  return "Excel（03_自社機械稼働表\\第NN期）と " + XL_POLL_MS / 60000 + "分ごとに合わせています。" +
    (s.いつ ? `最後に見た時刻 ${s.いつ}` : "まだ見ていません") + (s.書いた ? `・Excel に書いた ${s.書いた}日分` : "") +
    (s.err ? `<br>うまくいっていません：${esc(s.err)}` : "") +
    (s.書けない.length ? `<br>Excel に行が無いので書けない機械：${esc(s.書けない.slice(0, 10).join("、"))}${s.書けない.length > 10 ? " ほか" : ""}` : "");
}
setInterval(xlSync, XL_POLL_MS);
setTimeout(xlSync, 8000);
document.addEventListener("visibilitychange", () => { if (!document.hidden) setTimeout(xlSync, 2000); });
