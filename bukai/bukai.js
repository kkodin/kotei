/* 工事部会資料 Vr002（2026-11-13 Claude）
   元のExcel「工事工程計画及び実施報告書」と同じ並びの表で直接入力する。
   - セル：矢印・Enter・Tab・Shift・F2・Delete・Esc はExcelと同じ動き。文字を打つと上書き入力。
   - 右クリック：コピー・切り取り・貼り付け・消去（工事概要は「作業終了の色」も）。Ctrl+C/X/V も使える。
   - 工程線：空いた所をなぞって引く／線を掴んで移動（上下の行へも）／端を掴んで伸縮／Ctrlを押しながら動かすとコピー／
     ダブルクリックで文字／右クリックで色・文字・消す。Deleteで選んだ線を消す。
   - 工程線は部会資料の会議月ごとに保存する（entries[key].lines）。全体工程表の工程原本は書き換えない。
   - 「全体工程表から取り込む」で、工種を外注／直労（従業員）／機材の3行へ大まかにまとめた線を作り、あとは利用者が直す。
   - 入力は1セルごとに自動で共有保存する。同じ項目をほかの人が先に変えていたら止めて知らせる。 */
window.Bukai = (() => {
  "use strict";
  const C = window.BukaiCore, $ = id => document.getElementById(id);
  const e = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const ROWS = 5, DW = 7, LH = 30, COLORS = ["黒", "青", "赤", "緑", "橙", "紫", "茶", "灰", "黄"];
  let ctx, month, archive = "", busy = false, handlerFilter = "", showAll = false;
  let blocks = [], sel = null, anchor = null, editing = null, selLine = null, drag = null, clip = null, sheetActive = false;
  let queue = [], flushing = false, flushTimer = null, importPlan = null;

  /* ---------- データ ---------- */
  const reportName = () => `部会資料_${month}.json`;
  const empty = () => ({ kind: "bukai-report", schemaVersion: 1, yearMonth: month, entries: {} });
  const doc = () => ctx.files()[reportName()]?.doc || empty();
  const archived = () => archive ? ctx.files()[archive]?.doc : null;
  const ro = () => !!archive;
  const itemOf = key => (doc().entries || {})[key] || {};
  const period = () => archived()?.period || C.period(month);
  function linesOf(b) {
    if (b.snap) return b.snap.lines || legacyLines(b.snap);
    const it = itemOf(b.row.key);
    return Array.isArray(it.lines) ? it.lines : [];   // 取り込むまでは空（2026-11-13 利用者指定：取り込んでから編集する）
  }
  // Vr001の配布版（bars）を線として表示する
  function legacyLines(r) {
    const days = period().days;
    return (r.bars || []).map((x, i) => ({ id: "v" + i, row: Math.max(1, C.LANES.indexOf(x.category)), from: x.from, to: x.end, color: x.color || "青", label: x.work || "",
      off: days.filter(k => k >= x.from && k <= x.end && !(x.days || []).includes(k)) }));
  }
  function buildBlocks() {
    const snap = archived();
    if (snap) return (snap.rows || []).map(r => ({ row: { key: r.key, code: r.code, name: r.name }, values: r.values, snap: r, included: true }));
    const d = doc();
    return C.rows(ctx.sites(), ctx.masters()).map(r => ({ row: r, values: C.values(r, d), included: C.selected(r, d) }));
  }
  const handlersOf = b => [b.values.handler, b.values.subHandler].filter(Boolean);
  const nsp = s => String(s).replace(/[\s　]/g, "");
  const matchHandler = b => !handlerFilter || handlersOf(b).some(h => nsp(h).includes(nsp(handlerFilter)));
  const visible = all => all.filter(b => (showAll || b.included) && matchHandler(b));

  /* ---------- セルの定義（Excelの列 A〜G） ---------- */
  // c: 0 工事略称・コード / 1 発注者・工期 / 2 担当 / 3 工事概要 / 4 計画 / 5 実行％ / 6 備考
  const COLS = [["工事略称・コード", 118], ["発注者・工期", 112], ["担当", 54], ["工事概要", 330], ["計画", 36], ["実行%", 38], ["備考", 220]];
  const LANE_W = 34;
  function cellDef(c, r) {
    if (c === 0) return [{ ro: "code" }, { ro: "name" }, { f: "place" }, { ro: "" }, { f: "amount" }][r];
    if (c === 1) return [{ f: "client" }, { f: "clientStaff" }, { f: "start" }, { f: "end" }, { f: "phone" }][r];
    if (c === 2) return [{ f: "handler" }, { f: "subHandler" }, { ro: "" }, { ro: "" }, { ro: "" }][r];
    if (c === 3) return { f: "summary", line: r };
    if (c === 4) return r === 0 ? { f: "plan" } : { ro: "" };
    if (c === 5) return r === 0 ? { f: "execPercent" } : { ro: "" };
    return { f: "remarks", line: r };
  }
  const isRo = (c, r) => cellDef(c, r).ro !== undefined;
  const lineOf = (v, i) => String(v ?? "").split("\n")[i] ?? "";
  function display(b, c, r) {
    const d = cellDef(c, r), v = b.values;
    if (d.ro === "code") return b.row.code || "";
    if (d.ro === "name") return b.row.name || "";
    if (d.ro !== undefined) return "";
    const x = v[d.f];
    if (d.line !== undefined) return lineOf(x, d.line);
    if (d.f === "start") return x ? C.era(x) + "～" : "";
    if (d.f === "end") return x ? C.era(x) : "";
    if (d.f === "amount") return x === "" || x == null ? "" : Number(x).toLocaleString("ja-JP");
    return x ?? "";
  }
  function rawText(b, c, r) {   // 入力欄・コピーに使う文字
    const d = cellDef(c, r);
    if (d.ro !== undefined) return display(b, c, r);
    const x = b.values[d.f];
    if (d.line !== undefined) return lineOf(x, d.line);
    if (d.f === "start" || d.f === "end") return x ? C.era(x) : "";
    return x ?? "";
  }
  const doneSet = b => new Set(String(b.values.summaryDone || "").split(",").filter(Boolean).map(Number));

  /* ---------- 描画 ---------- */
  const md = k => k ? (+k.slice(5, 7)) + "/" + (+k.slice(8)) : "";
  const warekiMonth = m => { const y = +m.slice(0, 4), mo = +m.slice(5); return y >= 2019 ? `令和${y - 2018}年${mo}月` : `${y}年${mo}月`; };
  const localTime = t => { const d = new Date(t); return isNaN(d) ? t : d.toLocaleString("ja-JP"); };
  const dayIdx = (p, k) => { const i = p.days.indexOf(k); return i >= 0 ? i : k < p.from ? -1 : p.days.length; };
  function tableHTML(list, p, forPrint) {
    const chartW = p.days.length * DW, months = [], ticks = [];
    p.days.forEach((k, i) => {
      if (i === 0 || k.endsWith("-01")) months.push(`<span class="mo" style="left:${i * DW}px">${+k.slice(5, 7)} 月</span>`);
      if ([10, 20, 30].includes(+k.slice(8))) ticks.push(`<span class="tk" style="left:${i * DW}px">${+k.slice(8)}</span>`);
    });
    const snap = archived();
    const title = `<tr class="xl-title"><td colspan="7"><b>工事工程計画及び実施報告書</b>　${e(warekiMonth(snap?.yearMonth || month))}工事部会　<small>${e(snap ? "配布版 " + localTime(snap.created.at) : forPrint ? "編集中 " + new Date().toLocaleDateString("ja-JP") + "現在" : "")}</small></td><td colspan="2" class="legend"><span class="donebox"></span> は作業終了分</td></tr>`;
    const head = `<tr class="xl-head">${COLS.map(([t]) => `<th>${t}</th>`).join("")}<th></th><th class="scale"><div class="scale-in" style="width:${chartW}px">${months.join("")}${ticks.join("")}${gridLines(p, false)}</div></th></tr>`;
    const body = list.map((b, bi) => blockHTML(b, bi, p, forPrint)).join("");
    const width = COLS.reduce((a, [, w]) => a + w, 0) + LANE_W + chartW;
    return `<table class="xl" style="width:${width}px"><colgroup>${COLS.map(([, w]) => `<col style="width:${w}px">`).join("")}<col style="width:${LANE_W}px"><col style="width:${chartW}px"></colgroup><thead>${title}${head}</thead><tbody>${body || `<tr><td colspan="9" class="xl-empty">${handlerFilter ? "この担当の掲載現場はありません。" : "掲載する現場がありません。"}</td></tr>`}</tbody></table>`;
  }
  function gridLines(p, shade) {   // 月の境・10日ごとの線・土日祝の網かけ
    let s = "";
    p.days.forEach((k, i) => {
      if (k.endsWith("-01") && i) s += `<i class="mline" style="left:${i * DW}px"></i>`;
      else if ([10, 20].includes(+k.slice(8))) s += `<i class="dline" style="left:${i * DW}px"></i>`;
      if (shade && (ctx.holiday(k) || [0, 6].includes(new Date(k + "T00:00:00").getDay()))) s += `<i class="wkend" style="left:${i * DW}px;width:${DW}px"></i>`;
    });
    return s;
  }
  function blockHTML(b, bi, p, forPrint) {
    const done = doneSet(b);
    let s = "";
    for (let r = 0; r < ROWS; r++) {
      s += `<tr class="r${r}${b.included ? "" : " xl-excl"}" data-b="${bi}">`;
      for (let c = 0; c < 7; c++) {
        const d = cellDef(c, r);
        const cls = ["cell", "c" + c, d.ro !== undefined ? "xl-ro" : "", c === 3 && done.has(r) ? "xl-done" : "", d.ro === "name" ? "xl-name" : "", ["amount", "plan", "execPercent"].includes(d.f) ? "xl-num" : ""].filter(Boolean).join(" ");
        const inc = c === 0 && r === 0 && !forPrint && !archive ? `<label class="inc" title="部会資料に掲載する"><input type="checkbox" data-include="${e(b.row.key)}" ${b.included ? "checked" : ""}></label>` : "";
        s += `<td class="${cls}" data-c="${c}" data-r="${r}">${inc}${e(display(b, c, r))}</td>`;
      }
      s += `<td class="xl-lab">${C.LANES[r] || ""}</td>`;
      if (r === 0) s += `<td class="chart" rowspan="${ROWS}"><div class="chart-in" data-b="${bi}" style="width:${p.days.length * DW}px">${gridLines(p, true)}${linesHTML(b, p)}</div></td>`;
      s += "</tr>";
    }
    return s;
  }
  function linesHTML(b, p) {
    const derived = false, list = linesOf(b);
    if (!list.length && !b.snap && !ro()) return `<span class="xl-hint">工程線なし　右クリック →「全体工程表から取り込む」、または なぞって引く</span>`;
    return list.map(l => {
      if (l.to < p.from || l.from > p.end) return "";
      const a = Math.max(0, dayIdx(p, l.from)), z = Math.min(p.days.length - 1, dayIdx(p, l.to));
      const offs = (l.off || []).map(k => dayIdx(p, k)).filter(i => i >= a && i <= z).map(i => `<i class="xl-off" style="left:${(i - a) * DW}px;width:${DW}px"></i>`).join("");
      const on = selLine && selLine.b === b.row.key && selLine.id === l.id ? " xl-sel" : "";
      return `<div class="xl-ln${derived ? " xl-derived" : ""}${on}" data-id="${e(l.id)}" style="left:${a * DW}px;width:${(z - a + 1) * DW}px;top:${l.row * LH + 14}px;--c:${ctx.color(l.color || "黒")}" title="${e(l.label)} ${md(l.from)}〜${md(l.to)}"><span class="xl-bar">${offs}</span><span class="xl-tx">${e(l.label)}</span><b class="xl-h l"></b><b class="xl-h r"></b></div>`;
    }).join("");
  }
  function render() {
    if (!ctx || ctx.tab() !== "bukai" || editing || drag) return;
    try {
      $("bkMonth").value = month;
      $("bkVersion").innerHTML = '<option value="">編集中（最新の工程を参照）</option>' + archives().map(([name, f]) => `<option value="${e(name)}">配布版 ${e(localTime(f.doc.created.at))}　${e(f.doc.created.by)}</option>`).join("");
      $("bkVersion").value = archive;
      $("bkImport").disabled = ro() || busy; $("bkSnapshot").disabled = ro() || busy; $("bkImportLines").disabled = ro();
      const all = buildBlocks(), hs = [...new Set(all.flatMap(handlersOf))].sort((a, b) => a.localeCompare(b, "ja"));
      $("bkHandler").innerHTML = `<option value="">全員</option>` + hs.map(h => `<option ${h === handlerFilter ? "selected" : ""}>${e(h)}</option>`).join("");
      blocks = visible(all);
      $("bkCount").textContent = `掲載 ${all.filter(b => b.included).length}現場（表示 ${blocks.length}）／工程期間 ${period().from} 〜 ${period().end}`;
      $("bkGrid").innerHTML = tableHTML(blocks, period(), false);
      $("bkSheet").classList.toggle("readonly", ro());
      if (sel && sel.g > maxGr()) sel = anchor = null;
      paintSel();
    } catch (err) { $("bkGrid").innerHTML = `<p class="ng">${e(err.message)}</p>`; onError(err); }
  }
  function archives() { return Object.entries(ctx.files()).filter(([, f]) => f.doc?.kind === "bukai-snapshot" && f.doc.yearMonth === month).sort((a, b) => String(b[1].doc.created.at).localeCompare(String(a[1].doc.created.at))); }

  /* ---------- 選択（g＝表全体の行番号 現場番号×5＋行） ---------- */
  const tdAt = (g, c) => $("bkGrid").querySelector(`tr[data-b="${Math.floor(g / ROWS)}"].r${g % ROWS} td.cell[data-c="${c}"]`);
  const maxGr = () => blocks.length * ROWS - 1;
  const blockAt = g => blocks[Math.floor(g / ROWS)];
  function rect() { const a = anchor || sel; return [Math.min(a.g, sel.g), Math.max(a.g, sel.g), Math.min(a.c, sel.c), Math.max(a.c, sel.c)]; }
  const inRect = (g, c) => { if (!sel) return false; const [g0, g1, c0, c1] = rect(); return g >= g0 && g <= g1 && c >= c0 && c <= c1; };
  function paintSel() {
    $("bkGrid").querySelectorAll(".xl-on,.xl-act").forEach(td => td.classList.remove("xl-on", "xl-act"));
    ed().style.display = sel ? "" : "none";
    if (!sel) { $("bkCellInfo").textContent = ""; return; }
    const [g0, g1, c0, c1] = rect();
    for (let g = g0; g <= g1; g++) for (let c = c0; c <= c1; c++) tdAt(g, c)?.classList.add("xl-on");
    const td = tdAt(sel.g, sel.c);
    if (td) { td.classList.add("xl-act"); if (!drag) td.scrollIntoView({ block: "nearest", inline: "nearest" }); }
    const b = blockAt(sel.g);
    $("bkCellInfo").textContent = b ? `${b.row.name}　${COLS[sel.c][0]}${isRo(sel.c, sel.g % ROWS) ? "（台帳・工程表の値）" : ""}` : "";
    placeEditor();
  }
  function select(g, c, extend) {
    if (!blocks.length) return;
    g = Math.max(0, Math.min(maxGr(), g)); c = Math.max(0, Math.min(6, c));
    sel = { g, c }; if (!extend || !anchor) anchor = { g, c };
    clearLineSel();
    paintSel();
  }

  /* ---------- 入力（選択セルの上に置いた1つの入力欄で受ける。日本語入力も可） ---------- */
  const ed = () => $("bkEd");
  function placeEditor() {
    const t = ed(), td = sel && tdAt(sel.g, sel.c);
    if (!td || editing) return;
    const box = $("bkSheet"), wr = box.getBoundingClientRect(), r = td.getBoundingClientRect();
    Object.assign(t.style, { left: (r.left - wr.left + box.scrollLeft) + "px", top: (r.top - wr.top + box.scrollTop) + "px", width: r.width + "px", height: r.height + "px" });
    t.value = ""; t.classList.remove("editing");
    if (sheetActive && document.activeElement !== t) t.focus({ preventScroll: true });
  }
  function startEdit(mode, initial) {
    if (!sel || ro()) return false;
    if (isRo(sel.c, sel.g % ROWS)) { status("この欄は現場台帳・工程表の値です。全体工程表の現場の設定で直してください。", true); return false; }
    editing = { g: sel.g, c: sel.c, mode, b: blockAt(sel.g) };
    const t = ed(); t.classList.add("editing");
    if (initial != null) t.value = initial; else { t.value = rawText(editing.b, sel.c, sel.g % ROWS); t.setSelectionRange(t.value.length, t.value.length); }
    return true;
  }
  function cancelEdit() { if (!editing) return; editing = null; ed().classList.remove("editing"); ed().value = ""; }
  function commitEdit() {
    if (!editing) return true;
    const { g, c, b } = editing;
    try { setCell(b, c, g % ROWS, ed().value); }
    catch (err) { status(err.message, true); return false; }
    editing = null; ed().classList.remove("editing"); ed().value = "";
    return true;
  }
  // 1つのセルの値を変える（自動保存の列に積む）
  function setCell(b, c, r, text) {
    const d = cellDef(c, r);
    if (d.ro !== undefined || ro()) return;
    const f = d.f, cur = b.values[f];
    let after;
    text = String(text ?? "").replace(/\r/g, "");
    if (d.line !== undefined) {
      const ls = String(cur ?? "").split("\n"); while (ls.length < ROWS) ls.push("");
      ls[d.line] = text.replace(/\n/g, " ");
      after = ls.join("\n").replace(/\n+$/, "");
    } else if (f === "start" || f === "end") after = C.parseDate(text);
    else if (["amount", "plan", "execPercent"].includes(f)) after = text.replace(/[０-９．]/g, ch => ch === "．" ? "." : String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)).replace(/[,，円%％\s]/g, "");
    else after = text;
    after = C.validate(f, after);
    if (C.equal(after, cur ?? "")) return;
    if (f === "start" && after && b.values.end && after > b.values.end) throw Error("工期の開始日が終了日より後です。");
    if (f === "end" && after && b.values.start && b.values.start > after) throw Error("工期の開始日が終了日より後です。");
    change(b, { area: "values", field: f, after });
  }
  // 変更を手元へすぐ反映し、順番に共有保存する
  function change(b, p) {
    const key = b.row.key, it = itemOf(key);
    const before = p.area === "values" ? (it.values || {})[p.field] : it[p.area];
    const ch = { key, area: p.area, field: p.field, before, after: p.after };
    const f = ctx.files()[reportName()] ||= { doc: empty(), eTag: null, pending: [], saving: false, timer: null, err: null };
    f.doc = C.patch(f.doc, [ch]);
    f.saving = true;   // 保存が終わるまで30秒ごとの読み直しで上書きしない
    queue.push(ch);
    b.values = C.values(b.row, f.doc);
    if (p.area === "included") b.included = !!p.after;
    refreshBlock(b);
    clearTimeout(flushTimer); flushTimer = setTimeout(flush, 400);
    status("保存中…");
  }
  async function flush() {
    if (flushing || !queue.length) return;
    flushing = true; ctx.sync();
    const name = reportName(), batch = queue.splice(0);
    try {
      for (let n = 0; ; n++) {
        const fresh = await ctx.store().get(name), latest = fresh?.data || empty();
        if (latest.kind !== "bukai-report" || latest.yearMonth !== month) throw Error("保存先の資料形式を確認してください。");
        const next = C.patch(latest, batch); next.updated = { at: new Date().toISOString(), by: ctx.who() };
        try {
          const saved = await ctx.store().put(name, next, fresh?.eTag || null), f = ctx.files()[name];
          f.eTag = saved.eTag; f.doc = queue.length ? C.patch(next, queue) : next;   // 保存中に打った分も残す
          break;
        } catch (err) { if (![409, 412].includes(err.status) || n >= 2) throw err; }
      }
      status(queue.length ? "保存中…" : "共有保存済み");
    } catch (err) {
      queue = []; console.error(err);
      status("保存できませんでした：" + err.message + "（「最新を読み直す」で確認してください）", true);
    } finally {
      flushing = false;
      const f = ctx.files()[name]; if (f && !queue.length) f.saving = false;
      ctx.sync();
      if (queue.length) flush();
    }
  }
  function refreshBlock(b) {   // 1現場だけ描き直す
    const bi = blocks.indexOf(b); if (bi < 0) return;
    const rows = $("bkGrid").querySelectorAll(`tr[data-b="${bi}"]`);
    if (rows.length !== ROWS) { render(); return; }
    const tmp = document.createElement("tbody"); tmp.innerHTML = blockHTML(b, bi, period(), false);
    [...tmp.children].forEach((tr, i) => rows[i].replaceWith(tr));
    paintSel();
  }

  /* ---------- キーボード（Excelと同じ動き） ---------- */
  function onKey(ev) {
    if (!sel && !selLine) return;
    const k = ev.key, ctrl = ev.ctrlKey || ev.metaKey;
    if (ev.isComposing || k === "Process") return;
    if (!$("bkMenu").hidden) { if (k === "Escape") { ev.preventDefault(); hideMenu(); } return; }
    if (editing) {
      const mv = { Enter: [ev.shiftKey ? -1 : 1, 0], Tab: [0, ev.shiftKey ? -1 : 1] };
      if (editing.mode === "enter") Object.assign(mv, { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] });
      if (k === "Escape") { ev.preventDefault(); cancelEdit(); paintSel(); return; }
      if (mv[k] && !(k === "Enter" && ev.altKey)) { ev.preventDefault(); if (commitEdit()) select(sel.g + mv[k][0], sel.c + mv[k][1]); }
      return;
    }
    if (selLine && (k === "Delete" || k === "Backspace")) { ev.preventDefault(); deleteLine(selLine); return; }
    if (!sel) return;
    const mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [ev.shiftKey ? -1 : 1, 0], Tab: [0, ev.shiftKey ? -1 : 1] }[k];
    if (mv) {
      ev.preventDefault();
      const arrow = k.startsWith("Arrow");
      if (ctrl && arrow) { select(mv[0] ? (mv[0] < 0 ? 0 : maxGr()) : sel.g, mv[1] ? (mv[1] < 0 ? 0 : 6) : sel.c, ev.shiftKey); return; }
      select(sel.g + mv[0], sel.c + mv[1], ev.shiftKey && arrow); return;
    }
    if (k === "Home") { ev.preventDefault(); select(ctrl ? 0 : sel.g, 0, ev.shiftKey); return; }
    if (k === "PageDown" || k === "PageUp") { ev.preventDefault(); select(sel.g + (k === "PageDown" ? ROWS : -ROWS), sel.c, ev.shiftKey); return; }
    if (k === "F2") { ev.preventDefault(); startEdit("edit"); return; }
    if (k === "Delete") { ev.preventDefault(); clearRange(); return; }
    if (k === "Backspace") { ev.preventDefault(); const [g0, g1, c0, c1] = rect(); if (g0 === g1 && c0 === c1) startEdit("enter", ""); else clearRange(); return; }
    if (ctrl && k.toLowerCase() === "c") { ev.preventDefault(); copyRange(false); return; }
    if (ctrl && k.toLowerCase() === "x") { ev.preventDefault(); copyRange(true); return; }
    if (ctrl && k.toLowerCase() === "v") return;   // paste イベントで受ける
    if (ctrl && k.toLowerCase() === "a") { ev.preventDefault(); anchor = { g: 0, c: 0 }; sel = { g: maxGr(), c: 6 }; paintSel(); return; }
    if (!ctrl && !ev.altKey && k.length === 1) {   // 打った文字で上書き入力を始める（文字は入力欄にそのまま入る）
      if (!startEdit("enter", "")) ev.preventDefault();
    }
  }

  /* ---------- コピー・貼り付け・消去 ---------- */
  function rangeCells() { const [g0, g1, c0, c1] = rect(), out = []; for (let g = g0; g <= g1; g++) { const row = []; for (let c = c0; c <= c1; c++) row.push({ g, c }); out.push(row); } return out; }
  async function copyRange(cut) {
    if (!sel) return;
    if (cut && ro()) return;
    const text = rangeCells().map(row => row.map(({ g, c }) => rawText(blockAt(g), c, g % ROWS)).join("\t")).join("\r\n");
    clip = text;
    try { await navigator.clipboard.writeText(text); } catch (err) { /* 使えない環境では画面内のコピーで続ける */ }
    if (cut) clearRange();
    status(cut ? "切り取りました" : "コピーしました");
  }
  function pasteText(text) {
    if (!sel || ro()) return;
    const rows = String(text ?? "").replace(/\r\n?/g, "\n").replace(/\n$/, "").split("\n").map(r => r.split("\t"));
    const [g0, , c0] = rect(); let errors = 0;
    rows.forEach((row, i) => row.forEach((v, j) => {
      const g = g0 + i, c = c0 + j; if (g > maxGr() || c > 6 || isRo(c, g % ROWS)) return;
      try { setCell(blockAt(g), c, g % ROWS, v); } catch (err) { errors++; }
    }));
    anchor = { g: g0, c: c0 }; sel = { g: Math.min(maxGr(), g0 + rows.length - 1), c: Math.min(6, c0 + rows[0].length - 1) }; paintSel();
    if (errors) status(`貼り付けられなかったセルが ${errors} 件あります（日付・数値の形を確認してください）`, true);
  }
  async function pasteFromMenu() { let text = clip ?? ""; try { text = await navigator.clipboard.readText(); } catch (err) { } pasteText(text); }
  function clearRange() {
    if (!sel || ro()) return;
    rangeCells().flat().forEach(({ g, c }) => { if (!isRo(c, g % ROWS)) try { setCell(blockAt(g), c, g % ROWS, ""); } catch (err) { } });
    paintSel();
  }
  function toggleDone() {
    const [g0, g1] = rect(), touched = new Map();
    for (let g = g0; g <= g1; g++) { const b = blockAt(g); if (!touched.has(b)) touched.set(b, []); touched.get(b).push(g % ROWS); }
    touched.forEach((rs, b) => {
      const s = doneSet(b), on = !rs.every(r => s.has(r));
      rs.forEach(r => on ? s.add(r) : s.delete(r));
      change(b, { area: "values", field: "summaryDone", after: [...s].sort().join(",") });
    });
  }

  /* ---------- 右クリックの小窓 ---------- */
  function showMenu(x, y, items) {
    const m = $("bkMenu");
    m.innerHTML = items.map((it, i) => it === "-" ? "<hr>" : it.colors ? `<div class="xl-colors">${COLORS.map(cn => `<button type="button" data-color="${cn}" title="${cn}" style="background:${ctx.color(cn)}"></button>`).join("")}</div>` : `<button type="button" data-i="${i}" class="${it[2] || ""}" ${it[3] ? "disabled" : ""}>${e(it[0])}</button>`).join("");
    m.hidden = false;
    const r = m.getBoundingClientRect();
    m.style.left = Math.max(0, Math.min(x, innerWidth - r.width - 4)) + "px"; m.style.top = Math.max(0, Math.min(y, innerHeight - r.height - 4)) + "px";
    m.onclick = ev => {
      const cb = ev.target.closest("[data-color]"), bt = ev.target.closest("[data-i]");
      if (cb) { hideMenu(); items.find(x => x.colors).colors(cb.dataset.color); }
      else if (bt) { hideMenu(); items[+bt.dataset.i][1](); }
    };
  }
  function hideMenu() { $("bkMenu").hidden = true; if (sheetActive) ed().focus({ preventScroll: true }); }
  function cellMenu(ev) {
    const td = ev.target.closest("td.cell"); if (!td) return;
    ev.preventDefault();
    const g = +td.closest("tr").dataset.b * ROWS + +td.dataset.r, c = +td.dataset.c;
    if (!inRect(g, c)) select(g, c);
    const items = [["コピー　Ctrl+C", () => copyRange(false)], ["切り取り　Ctrl+X", () => copyRange(true), "", ro()], ["貼り付け　Ctrl+V", pasteFromMenu, "", ro()], ["消去　Delete", clearRange, "", ro()]];
    if (sel.c === 3 && !ro()) items.push("-", ["作業終了の色をつける／外す", toggleDone]);
    showMenu(ev.clientX, ev.clientY, items);
  }

  /* ---------- 工程線（全体工程表と同じ操作） ---------- */
  const uid = () => "L" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const addDays = (k, n) => { const d = new Date(k + "T00:00:00"); d.setDate(d.getDate() + n); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };
  function hitChart(el) { const ch = el.closest(".chart-in"); return ch ? { b: blocks[+ch.dataset.b], ch } : null; }
  function dayAt(ch, x) { return Math.max(0, Math.min(period().days.length - 1, Math.floor((x - ch.getBoundingClientRect().left) / DW))); }
  function laneAt(ch, y) { return Math.max(0, Math.min(3, Math.floor((y - ch.getBoundingClientRect().top) / LH))); }
  const editableLines = b => C.clone(linesOf(b));
  function saveLines(b, list) { if (!ro()) change(b, { area: "lines", after: list }); }
  function clearLineSel() { selLine = null; $("bkGrid").querySelectorAll(".xl-ln.xl-sel").forEach(x => x.classList.remove("xl-sel")); }
  function selectLine(b, id) {
    clearLineSel(); selLine = { b: b.row.key, id };
    $("bkGrid").querySelector(`.chart-in[data-b="${blocks.indexOf(b)}"] .xl-ln[data-id="${CSS.escape(id)}"]`)?.classList.add("xl-sel");
    sheetActive = true; ed().focus({ preventScroll: true });
  }
  function onChartDown(ev) {
    if (ev.button !== 0 || ro()) return;
    const hit = hitChart(ev.target); if (!hit) return;
    const { b, ch } = hit, lnEl = ev.target.closest(".xl-ln");
    hideMenu(); if (editing && !commitEdit()) return;
    sel = anchor = null; paintSel();
    if (lnEl) {
      const l = linesOf(b).find(x => x.id === lnEl.dataset.id); if (!l) return;
      const mode = ev.target.matches(".xl-h.l") ? "l" : ev.target.matches(".xl-h.r") ? "r" : "move";
      drag = { kind: "line", b, l, el: lnEl, ch, mode, x0: ev.clientX, y0: ev.clientY, a: l.from, z: l.to, row: l.row, moved: false, copy: false };
      selectLine(b, l.id);
    } else {
      const d0 = dayAt(ch, ev.clientX), row = laneAt(ch, ev.clientY), g = document.createElement("div");
      g.className = "xl-ln xl-ghost"; g.style.top = row * LH + 14 + "px"; ch.appendChild(g);
      drag = { kind: "new", b, ch, d0, d1: d0, row, el: g, moved: false };
      placeGhost(); clearLineSel();
    }
    ev.preventDefault();
  }
  function placeGhost() {
    const a = Math.min(drag.d0, drag.d1), z = Math.max(drag.d0, drag.d1), p = period();
    Object.assign(drag.el.style, { left: a * DW + "px", width: (z - a + 1) * DW + "px" });
    drag.el.innerHTML = `<span class="xl-bar"></span><span class="xl-tx">${md(p.days[a])}〜${md(p.days[z])}（${z - a + 1}日）</span>`;
  }
  function onChartMove(ev) {
    if (!drag) return;
    if (drag.kind === "new") { drag.d1 = dayAt(drag.ch, ev.clientX); if (drag.d1 !== drag.d0) drag.moved = true; placeGhost(); return; }
    if (Math.abs(ev.clientX - drag.x0) > 3 || Math.abs(ev.clientY - drag.y0) > 10) drag.moved = true;
    if (!drag.moved) return;
    const dd = Math.round((ev.clientX - drag.x0) / DW), l = drag.l;
    let a = l.from, z = l.to, row = l.row;
    if (drag.mode === "move") { a = addDays(l.from, dd); z = addDays(l.to, dd); row = laneAt(drag.ch, ev.clientY); }
    else if (drag.mode === "l") { a = addDays(l.from, dd); if (a > z) a = z; }
    else { z = addDays(l.to, dd); if (z < a) z = a; }
    Object.assign(drag, { a, z, row, copy: drag.mode === "move" && (ev.ctrlKey || ev.metaKey) });
    const p = period(), va = Math.max(0, dayIdx(p, a)), vz = Math.min(p.days.length - 1, dayIdx(p, z));
    Object.assign(drag.el.style, { left: va * DW + "px", width: Math.max(DW, (vz - va + 1) * DW) + "px", top: row * LH + 14 + "px" });
    drag.el.classList.add("xl-drag"); drag.el.classList.toggle("xl-copy", drag.copy);
    status(`${drag.copy ? "コピー　" : ""}${md(a)} 〜 ${md(z)}　${Math.round((new Date(z) - new Date(a)) / 864e5) + 1}日`);
  }
  function onChartUp() {
    if (!drag) return;
    const d = drag; drag = null;
    if (d.kind === "new") {
      d.el.remove();
      if (!d.moved) return;   // クリックだけでは引かない
      const p = period(), a = Math.min(d.d0, d.d1), z = Math.max(d.d0, d.d1), list = editableLines(d.b), id = uid();
      list.push({ id, row: d.row, from: p.days[a], to: p.days[z], color: "黒", label: "" });
      saveLines(d.b, list); selectLine(d.b, id);
      status("工程線を引きました。ダブルクリックで文字、右クリックで色を変えられます。");
      return;
    }
    if (!d.moved) return;
    const list = editableLines(d.b), i = list.findIndex(x => x.id === d.l.id); if (i < 0) { render(); return; }
    if (!d.copy && d.a === d.l.from && d.z === d.l.to && d.row === d.l.row) { refreshBlock(d.b); return; }
    const moved = { ...list[i], from: d.a, to: d.z, row: d.row }; delete moved.off;   // 動かした線は休工日の白抜きを外す
    if (d.copy) { moved.id = uid(); list.push(moved); } else list[i] = moved;
    saveLines(d.b, list); selectLine(d.b, moved.id);
  }
  function lineAct(b, id, fn) { const list = editableLines(b), l = list.find(x => x.id === id); if (!l) return; fn(l, list); saveLines(b, list); }
  function editLabel(b, id) {
    const l = linesOf(b).find(x => x.id === id); if (!l) return;
    const t = prompt("工程線の文字（例：推進工、入孔・埋戻し）", l.label || ""); if (t == null) return;
    lineAct(b, id, x => { x.label = t.slice(0, 200); });
  }
  function deleteLine(s) {
    const b = blocks.find(x => x.row.key === s.b); if (!b || ro()) return;
    lineAct(b, s.id, (l, list) => list.splice(list.indexOf(l), 1)); selLine = null;
    status("工程線を消しました");
  }
  function chartMenu(ev) {
    const hit = hitChart(ev.target); if (!hit) return;
    ev.preventDefault(); if (ro()) return;
    const { b } = hit, lnEl = ev.target.closest(".xl-ln"), derived = !Array.isArray(itemOf(b.row.key).lines);
    if (lnEl) {
      const id = lnEl.dataset.id; selectLine(b, id);
      showMenu(ev.clientX, ev.clientY, [{ colors: cn => lineAct(b, id, l => { l.color = cn; }) }, ["文字を入れる・直す", () => editLabel(b, id)], "-", ["この工程線を消す　Delete", () => deleteLine({ b: b.row.key, id }), "danger"]]);
    } else {
      showMenu(ev.clientX, ev.clientY, [
        ["全体工程表から取り込む（工種の振り分けを選ぶ）…", () => openPick(b), "", !b.row.siteId],
        ["全体工程表から取り込む（前回の振り分けで）", () => importLines([b], true), "", !b.row.siteId],
        "-",
        b.row.siteId ? ["全体工程表でこの現場を開く", () => { if (canLeave()) ctx.openSite(b.row.siteId); }] : ["（台帳のみの現場：工程はありません）", () => { }, "", true],
      ]);
    }
  }
  /* ---------- 全体工程表からの取り込み ---------- */
  // list の現場へ、振り分けに従ってまとめた線を入れる。すでに線がある現場は確認してから置き換える
  function importLines(list, ask) {
    if (ro()) return;
    list = list.filter(b => b.row.siteId);
    if (!list.length) { status("全体工程表に工程がある現場がありません。", true); return; }
    const has = list.filter(b => linesOf(b).length);
    if (ask && has.length && !confirm(has.length === 1 && list.length === 1 ? `「${has[0].row.name}」の今の工程線を、全体工程表から取り込んだ線で置き換えます。よろしいですか？` : `${has.length}現場はすでに工程線があります。全体工程表から取り込んだ線で置き換えますか？\n（キャンセル：線が無い現場だけ取り込みます）`)) list = list.filter(b => !has.includes(b));
    let n = 0, empty = 0;
    for (const b of list) {
      const ls = C.deriveLines(b.row, itemOf(b.row.key), period(), ctx.works);
      if (!ls.length) empty++;
      change(b, { area: "lines", after: ls }); n++;
    }
    status(`${n}現場へ全体工程表から取り込みました${empty ? `（${empty}現場はこの3か月に工程がありません）` : ""}。線は自由に直せます。`);
  }
  let pickFor = null;
  function openPick(b) {
    const src = b.row.source || {}, p = period(), it = itemOf(b.row.key);
    const inPeriod = id => (src.バー || []).some(x => x.段 === id && x.開始 <= p.end && x.終了 >= p.from);
    const lanes = (src.段 || []).filter(l => (src.バー || []).some(x => x.段 === l.id));
    if (!lanes.length) { status("この現場は全体工程表に工程がありません。", true); return; }
    pickFor = b;
    $("bkPickTitle").textContent = b.row.name;
    $("bkPickRows").innerHTML = lanes.map(l => { const cs = C.laneCats(l, it); return `<tr data-lane="${e(l.id)}" class="${inPeriod(l.id) ? "" : "xl-out"}"><td>${e(l.名前 || "工種")}${l.種別 === "外注" ? `<small>（外注 ${e(l.業者 || "")}）</small>` : l.業者 ? `<small>（＋外注 ${e(l.業者)}）</small>` : ""}${inPeriod(l.id) ? "" : "<small>（この3か月は工程なし）</small>"}</td>${["外注", "直労", "機材"].map(c => `<td><input type="checkbox" data-cat="${c}" ${cs.includes(c) ? "checked" : ""}></td>`).join("")}</tr>`; }).join("");
    $("dBkPick").showModal();
  }
  function applyPick() {
    const b = pickFor; if (!b) return;
    const cats = { ...(itemOf(b.row.key).categories || {}) };
    $("bkPickRows").querySelectorAll("tr[data-lane]").forEach(tr => { cats[tr.dataset.lane] = [...tr.querySelectorAll("[data-cat]:checked")].map(x => x.dataset.cat); });
    $("dBkPick").close(); pickFor = null;
    change(b, { area: "categories", after: cats });
    importLines([b], true);
  }

  /* ---------- 配布版・印刷・JSON ---------- */
  function capture() {
    const p = C.period(month);
    return { kind: "bukai-snapshot", schemaVersion: 2, yearMonth: month, created: { at: new Date().toISOString(), by: ctx.who() }, period: p,
      rows: buildBlocks().filter(b => b.included).map(b => ({ key: b.row.key, code: b.row.code, name: b.row.name, values: C.clone(b.values), lines: C.clone(linesOf(b)) })) };
  }
  async function snapshot() {
    if (busy || ro() || (editing && !commitEdit())) return;
    if (queue.length || flushing) { status("保存が終わってから配布版を作ってください。", true); return; }
    try {
      const snap = capture(); if (!snap.rows.length) throw Error("掲載する現場を選んでください。");
      if (!confirm(`${month} の ${snap.rows.length}現場を配布版として保存します。保存後も編集中の資料は直せます。`)) return;
      setBusy(true); status("配布版を保存中…");
      const name = `部会配布版_${month}_${ctx.id()}.json`, r = await ctx.store().put(name, snap, null);
      ctx.files()[name] = { doc: C.clone(snap), eTag: r.eTag, pending: [], saving: false, timer: null, err: null }; ctx.sync();
      archive = name; status("配布版を保存しました");
    } catch (err) { onError(err); } finally { setBusy(false); render(); }
  }
  function preparePrint() {
    ctx.paper($("bkPaper").value);   // Ctrl+P で印刷したときも選んだ用紙の大きさで縮める
    const snap = archived() || capture(), p = snap.period || C.period(month);
    const list = (snap.rows || []).map(r => ({ row: { key: r.key, code: r.code, name: r.name }, values: r.values, snap: r, included: true })).filter(matchHandler);
    const out = $("printPages"); out.innerHTML = "";
    const n = +$("bkPerPage").value || 4, total = Math.max(1, Math.ceil(list.length / n)), keep = selLine;
    selLine = null;
    // 画面では #printPages が非表示で大きさが0になり、縮小率が1のまま（5現場目以降が切れた）。測る間だけ表示する
    out.style.cssText = "display:block;position:absolute;left:-99999px;top:0;visibility:hidden";
    for (let i = 0; i < total; i++) {
      const page = document.createElement("div"); page.className = "ppage";
      page.innerHTML = `<div class="bk-print-page">${tableHTML(list.slice(i * n, (i + 1) * n), p, true)}<div class="bk-print-footer"><span>${handlerFilter ? "担当 " + e(handlerFilter) + "　" : ""}工程期間 ${e(p.from)} 〜 ${e(p.end)}　線の白抜きは休工日</span><span>${i + 1} / ${total}</span></div></div>`;
      out.appendChild(page);
      const sheet = page.firstElementChild;
      sheet.style.zoom = Math.min(1, ctx.printWidth() / sheet.scrollWidth, ctx.printHeight() / sheet.scrollHeight).toFixed(4);
    }
    out.style.cssText = "";
    selLine = keep;
    document.body.classList.add("printing");
  }
  function print() {
    if (editing && !commitEdit()) return;
    try { if (!(archived() || capture()).rows.length) throw Error("掲載する現場を選んでください。"); ctx.paper($("bkPaper").value); window.print(); } catch (err) { onError(err); }
  }
  async function readImport(event) {
    const file = event.target.files[0]; event.target.value = ""; if (!file || busy || ro()) return;
    if (queue.length || flushing) { status("保存が終わってから取り込んでください。", true); return; }
    try {
      if (file.size > 2 * 1024 * 1024) throw Error("JSONは2MB以内にしてください。");
      const data = JSON.parse(await file.text()), rows = C.rows(ctx.sites(), ctx.masters());
      importPlan = C.importData(data, month, rows, doc()); $("bkImportError").textContent = "";
      $("bkImportPreview").innerHTML = `<p>${e(month)} の ${importPlan.preview.length}現場へ反映します。指定された空文字は空欄として反映します。</p><table class="bk-preview"><tr><th>現場</th><th>反映する内容</th></tr>${importPlan.preview.map(x => `<tr><td>${e(x.row.code)} ${e(x.row.name)}</td><td>${importPlan.changes.filter(p => p.key === x.row.key && p.area === "values").map(p => `${e(C.fields[p.field])}: ${e(C.values(x.row, doc())[p.field] ?? "")} → ${e(p.after)}`).join("\n")}</td></tr>`).join("")}</table>`;
      $("dBukaiImport").showModal();
    } catch (err) { onError(err); }
  }
  async function applyImport() {
    if (!importPlan || busy) return;
    try {
      setBusy(true);
      const name = reportName();
      for (let n = 0; ; n++) {
        const fresh = await ctx.store().get(name), next = C.patch(fresh?.data || empty(), importPlan.changes);
        try { const r = await ctx.store().put(name, next, fresh?.eTag || null); ctx.files()[name] = { doc: next, eTag: r.eTag, pending: [], saving: false, timer: null, err: null }; break; }
        catch (err) { if (![409, 412].includes(err.status) || n >= 2) throw err; }
      }
      $("dBukaiImport").close(); importPlan = null; status("取り込んだ内容を共有保存しました");
    } catch (err) { onError(err); $("bkImportError").textContent = err.message; } finally { setBusy(false); render(); }
  }
  function exportJSON() {
    try {
      const snap = archived() || capture(), items = snap.rows.map(r => {
        if (!r.code) throw Error("コード未登録の現場があります。現場コードを設定してから書き出してください。");
        const v = r.values;
        return { projectCode: r.code, client: v.client, handler: v.handler, summary: v.summary, plan: v.plan, execPercent: v.execPercent, remarks: v.remarks, start: v.start, end: v.end,
          metadata: { pipeDia: v.pipeDia, length: v.length, count: v.count, soil: v.soil, amount: v.amount, phone: v.phone, place: v.place, clientStaff: v.clientStaff, subHandler: v.subHandler } };
      });
      if (!items.length) throw Error("掲載する現場を選んでください。");
      const payload = { yearMonth: month, updatedAt: snap.created.at, updatedBy: snap.created.by, items };
      const url = URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" })), a = document.createElement("a");
      a.href = url; a.download = `工事部会資料_${month}.json`; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
      status("共有用のJSONを書き出しました");
    } catch (err) { onError(err); }
  }

  /* ---------- 共通 ---------- */
  function setBusy(v) { busy = v; $("bkImportApply").disabled = v; $("bkSnapshot").disabled = v || ro(); }
  function status(text, err = false) { $("bkStatus").textContent = text; $("bkStatus").className = err ? "err" : ""; }
  function onError(err) { console.error(err); status(err.message, true); }
  function canLeave() { if (editing && !commitEdit()) return false; return !busy && ((!queue.length && !flushing) || confirm("保存中の入力があります。移動しますか？")); }

  function init(bridge) {
    ctx = bridge; month = ctx.today().slice(0, 7);
    const sh = $("bkSheet"), t = ed();
    $("bkMonth").onchange = ev => { if (!canLeave()) { ev.target.value = month; return; } try { C.period(ev.target.value); month = ev.target.value; archive = ""; sel = anchor = null; render(); status("会議月を変更しました"); } catch (err) { ev.target.value = month; onError(err); } };
    $("bkVersion").onchange = ev => { if (!canLeave()) { ev.target.value = archive; return; } archive = ev.target.value; sel = anchor = null; render(); };
    $("bkHandler").onchange = ev => { if (editing && !commitEdit()) return; handlerFilter = ev.target.value; sel = anchor = null; render(); };
    $("bkShowAll").onchange = ev => { if (editing && !commitEdit()) return; showAll = ev.target.checked; sel = anchor = null; render(); };
    $("bkRefresh").onclick = async () => { if (!canLeave()) return; try { if (!await ctx.pull()) throw Error("共有データを読み直せませんでした。通信状態を確認してください。"); render(); status("最新の共有データを読み直しました"); } catch (err) { onError(err); } };
    $("bkSnapshot").onclick = snapshot; $("bkPrint").onclick = print; $("bkExport").onclick = exportJSON; $("bkImport").onclick = () => $("bkFile").click(); $("bkFile").onchange = readImport;
    $("bkImportCancel").onclick = () => { $("dBukaiImport").close(); importPlan = null; };
    $("bkImportApply").onclick = applyImport;
    $("bkImportLines").onclick = () => { if (editing && !commitEdit()) return; importLines(blocks.filter(b => !b.snap), true); };
    $("bkPickCancel").onclick = () => { $("dBkPick").close(); pickFor = null; };
    $("bkPickApply").onclick = applyPick;
    sh.addEventListener("pointerdown", ev => {
      if (ev.target.closest("[data-include]") || ev.target === t) return;
      sheetActive = true;
      if (ev.target.closest(".chart-in")) { onChartDown(ev); return; }
      const td = ev.target.closest("td.cell"); if (!td) return;
      if (editing && !commitEdit()) { ev.preventDefault(); return; }
      const g = +td.closest("tr").dataset.b * ROWS + +td.dataset.r, c = +td.dataset.c;
      if (ev.button === 2 && inRect(g, c)) return;
      select(g, c, ev.shiftKey); ev.preventDefault();
      if (ev.button !== 0) return;
      const mv = e2 => { const x = document.elementFromPoint(e2.clientX, e2.clientY)?.closest("#bkGrid td.cell"); if (!x) return; const g2 = +x.closest("tr").dataset.b * ROWS + +x.dataset.r; if (g2 !== sel.g || +x.dataset.c !== sel.c) { sel = { g: g2, c: +x.dataset.c }; paintSel(); } };
      const up = () => { removeEventListener("pointermove", mv); removeEventListener("pointerup", up); };
      addEventListener("pointermove", mv); addEventListener("pointerup", up);
    });
    sh.addEventListener("dblclick", ev => {
      const lnEl = ev.target.closest(".xl-ln"), hit = hitChart(ev.target);
      if (lnEl && hit && !ro()) { editLabel(hit.b, lnEl.dataset.id); return; }
      if (ev.target.closest("td.cell") || ev.target === t) startEdit("edit");
    });
    sh.addEventListener("contextmenu", ev => { if (ev.target.closest(".chart-in")) chartMenu(ev); else if (ev.target === t && sel) { ev.preventDefault(); cellMenu({ target: tdAt(sel.g, sel.c), clientX: ev.clientX, clientY: ev.clientY, preventDefault() { } }); } else cellMenu(ev); });
    sh.addEventListener("change", ev => { const box = ev.target.closest("[data-include]"); if (!box) return; const b = blocks.find(x => x.row.key === box.dataset.include); if (b) change(b, { area: "included", after: box.checked }); if (!showAll && !box.checked) render(); });
    addEventListener("pointermove", onChartMove); addEventListener("pointerup", onChartUp);
    t.addEventListener("keydown", onKey);
    t.addEventListener("compositionstart", () => { if (!editing && sel) startEdit("enter", ""); });
    t.addEventListener("input", () => { if (!editing && sel && t.value) startEdit("enter", t.value); });
    t.addEventListener("paste", ev => { if (editing) return; ev.preventDefault(); pasteText(ev.clipboardData.getData("text/plain")); });
    t.addEventListener("blur", () => setTimeout(() => { if (editing && document.activeElement !== t && $("bkMenu").hidden) { if (!commitEdit()) cancelEdit(); paintSel(); } }, 0));
    document.addEventListener("pointerdown", ev => {
      if (!ev.target.closest("#bkMenu")) $("bkMenu").hidden = true;
      if (!ev.target.closest("#bkSheet") && !ev.target.closest("#bkMenu")) sheetActive = false;
    }, true);
    addEventListener("resize", () => placeEditor());
    window.addEventListener("beforeunload", ev => { if (editing || queue.length || flushing || busy) { ev.preventDefault(); ev.returnValue = ""; } });
  }
  return { init, render, canLeave, preparePrint, isDirty: () => !!editing || queue.length > 0 };
})();
