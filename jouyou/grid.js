/* 常用日報 Vr000 は 日報Vr002 の grid.js を写したもの（タップで直す を足した）
   Excel のように動く表（2026-10-30 ユーザー：使用感・セルの移動方法を EXCEL と同じに）
   ・矢印／Enter（下）／Shift+Enter（上）／Tab（右）／Shift+Tab（左）／Home／Ctrl+Home／Ctrl+End／Ctrl+矢印（データの端）／PageUp・PageDown
   ・選んだセルでそのまま打つ＝上書き（日本語の変換もそのまま入る）／F2・ダブルクリック＝中を直す（矢印は字の中を動く）／Esc＝取り消し
   ・Shift+矢印・ドラッグ＝範囲／Ctrl+A／行番号・列の文字を押すと行・列ごと／Delete＝消す／Backspace＝消して打ち始め
   ・Ctrl+C／X／V（Excel との間でもコピー・貼り付けできる。タブ区切り）／Ctrl+Z／Y／Ctrl+D（下へコピー）／Ctrl+R（右へコピー）
   ・Ctrl+Enter＝選んだ範囲すべてに入れる／Alt+↓＝その列に前に打った言葉の一覧／打ち始めると前に打った言葉を続きに出す（オートコンプリート）
   ・Ctrl+;＝今日の日付／Ctrl+:＝今の時刻／右下の■をドラッグ＝コピー（フィル）／右クリック＝メニュー
   表そのもののデータは持たない。読むのは opt.get(r,c)、変えるのは opt.apply([{r,c,v}])
   Vr002（2026-10-30）：1画面に表を縦にいくつも並べる（入力欄の6種類）。
   ・opt.noLetters：A B C の行を出さず、opt.head() の見出しの行（data-col 付き）を列の見出しにする
   ・opt.onEdge(dir, c)：上の端・下の端から ↑↓／Enter で出たとき、となりの表へ（true を返したら動いたこと）
   ・opt.isCur()：数式バーはいま選んでいる表だけが受ける／opt.onFocus()：この表を選んだ */
"use strict";

function createGrid(opt) {
  const host = opt.host, cols = opt.cols;
  const NC = cols.length;
  let nRows = opt.rows();
  const st = { r: 0, c: 0, ar: 0, ac: 0, edit: null, drag: null, fill: null, list: null };
  // edit: { mode:"enter"|"edit", r, c, orig }
  const el = document.createElement("div");
  el.className = "xl";
  el.tabIndex = -1;
  host.innerHTML = "";
  host.appendChild(el);
  // 表だけを描き直す入れ物（入力欄・枠は外に置いて、描き直してもキーの受け先が消えないように）
  const tw = document.createElement("div");
  el.appendChild(tw);
  const ed = document.createElement("input");   // いつも選んだセルの上にいて、打った字を受ける（変換も）
  ed.className = "xled";
  ed.autocomplete = "off"; ed.spellcheck = false;
  const selBox = document.createElement("div"); selBox.className = "xlsel";
  const actBox = document.createElement("div"); actBox.className = "xlact";
  const handle = document.createElement("div"); handle.className = "xlfill"; handle.title = "ドラッグでコピー（フィル）";
  const fillBox = document.createElement("div"); fillBox.className = "xlfillbox"; fillBox.hidden = true;
  const listBox = document.createElement("div"); listBox.className = "xllist"; listBox.hidden = true;
  let table = null, composing = false, lastInputType = "";

  const top = () => Math.min(st.r, st.ar), bot = () => Math.max(st.r, st.ar);
  const lef = () => Math.min(st.c, st.ac), rig = () => Math.max(st.c, st.ac);
  const clampR = r => Math.max(0, Math.min(nRows - 1, r));
  const clampC = c => Math.max(0, Math.min(NC - 1, c));
  const editable = c => !cols[c].ro;
  const cellEl = (r, c) => table && table.rows[r + opt.headRows] && table.rows[r + opt.headRows].cells[c + 1];

  /* ---------- 描く ---------- */
  function render() {
    nRows = opt.rows();
    const h = [];
    h.push(`<table class="xlt" style="width:${(opt.rowHeadW || 34) + cols.reduce((a, c) => a + c.w, 0)}px"><colgroup><col style="width:${opt.rowHeadW || 34}px">` + cols.map(c => `<col style="width:${c.w}px">`).join("") + `</colgroup>`);
    // 列の文字（A B C …）
    if (!opt.noLetters) h.push(`<thead><tr class="xlch"><th class="xlcorner" data-all="1"></th>` + cols.map((c, i) => `<th data-col="${i}" class="${c.pb ? "pb" : ""}">${c.letter}</th>`).join("") + `</tr></thead>`);
    h.push(`<tbody>`);
    // 上の見出しの行（日付・天候・項目名）は opt.head が html を返す
    h.push(opt.head());
    for (let r = 0; r < nRows; r++) {
      h.push(`<tr><th class="xlrh" data-row="${r}">${r + opt.rowBase}</th>`);
      for (let c = 0; c < NC; c++) {
        const v = opt.get(r, c), d = opt.deco ? opt.deco(r, c) : null;
        h.push(`<td class="${cols[c].cls || ""}${cols[c].pb ? " pb" : ""}${d && d.cls ? " " + d.cls : ""}"${d && d.title ? ` title="${escH(d.title)}"` : ""}><div class="xc">${escH(v)}</div></td>`);
      }
      h.push(`</tr>`);
    }
    h.push(`</tbody></table>`);
    const hadFocus = document.activeElement === ed;
    tw.innerHTML = h.join("");
    table = tw.querySelector("table");
    if (!ed.parentNode) el.append(selBox, actBox, handle, fillBox, ed, listBox);
    if (hadFocus && document.activeElement !== ed) ed.focus({ preventScroll: true });
    overflowText();
    place();
  }
  const escH = s => String(s ?? "").replace(/[&<>"]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
  // Excel と同じく、右のセルが空なら字をはみ出して見せる
  function overflowText() {
    for (let r = 0; r < nRows; r++) {
      const tr = table.rows[r + opt.headRows];
      for (let c = 0; c < NC; c++) {
        const td = tr.cells[c + 1], div = td.firstChild;
        if (!div.textContent || cols[c].align === "r" || cols[c].align === "c") continue;
        let w = td.offsetWidth, k = c + 1;
        while (k < NC && !opt.get(r, k) && !cols[k].stopFlow) { w += tr.cells[k + 1].offsetWidth; k++; }
        if (k > c + 1) { div.style.width = (w - 4) + "px"; td.classList.add("flow"); }
      }
    }
  }
  // 選んだ範囲・選んだセル・入力欄の位置
  function rectOf(r0, c0, r1, c1) {
    const a = cellEl(r0, c0), b = cellEl(r1, c1);
    if (!a || !b) return null;
    return { x: a.offsetLeft, y: a.offsetTop, w: b.offsetLeft + b.offsetWidth - a.offsetLeft, h: b.offsetTop + b.offsetHeight - a.offsetTop };
  }
  function place() {
    if (!table) return;
    st.r = clampR(st.r); st.ar = clampR(st.ar); st.c = clampC(st.c); st.ac = clampC(st.ac);
    const R = rectOf(top(), lef(), bot(), rig()), A = rectOf(st.r, st.c, st.r, st.c);
    if (!R || !A) return;
    Object.assign(selBox.style, { left: R.x - 1 + "px", top: R.y - 1 + "px", width: R.w + 1 + "px", height: R.h + 1 + "px" });
    selBox.classList.toggle("multi", top() !== bot() || lef() !== rig());
    Object.assign(actBox.style, { left: A.x - 1 + "px", top: A.y - 1 + "px", width: A.w + 1 + "px", height: A.h + 1 + "px" });
    Object.assign(handle.style, { left: R.x + R.w - 4 + "px", top: R.y + R.h - 4 + "px" });
    Object.assign(ed.style, { left: A.x + "px", top: A.y + "px", width: Math.max(A.w, st.edit ? 60 : 0) + "px", height: A.h - 1 + "px" });
    ed.classList.toggle("r", cols[st.c].align === "r");
    // 見出しの色（選んでいる行・列）
    table.querySelectorAll(".xlch th.on, .xlrh.on").forEach(x => x.classList.remove("on"));
    for (let c = lef(); c <= rig(); c++) { const th = table.rows[0].cells[c + 1]; th && th.classList.add("on"); }
    for (let r = top(); r <= bot(); r++) { const th = table.rows[r + opt.headRows] && table.rows[r + opt.headRows].cells[0]; th && th.classList.add("on"); }
    if (!st.edit) { ed.value = ""; }
    if (opt.onSelect) opt.onSelect(cellName(st.r, st.c), st.edit ? null : opt.get(st.r, st.c), rangeStat());
  }
  const cellName = (r, c) => opt.cellName ? opt.cellName(r, c) : cols[c].letter + (r + opt.rowBase);
  function rangeStat() {
    if (top() === bot() && lef() === rig()) return "";
    let n = 0, sum = 0, nums = 0;
    for (let r = top(); r <= bot(); r++) for (let c = lef(); c <= rig(); c++) {
      const v = opt.get(r, c); if (v === "") continue; n++;
      const x = Number(String(v).replace(/,/g, "")); if (v !== "" && !isNaN(x)) { sum += x; nums++; }
    }
    return `データの個数: ${n}` + (nums ? `　合計: ${Math.round(sum * 1000) / 1000}` : "");
  }
  // 選んだセルが見えるように（表が縦にいくつも並んでも効くよう、画面の上の位置で比べる）
  function scrollIntoView() {
    const A = cellEl(st.r, st.c), sc = opt.scroller;
    if (!A || !sc) return;
    const s = sc.getBoundingClientRect(), a = A.getBoundingClientRect();
    const top = s.top + (opt.stickyTop ? opt.stickyTop() : 0) + (opt.noLetters ? 0 : table.tHead.offsetHeight + headH());
    const left = s.left + (opt.rowHeadW || 34);
    if (a.top < top) sc.scrollTop -= top - a.top + (st.r === 0 ? headH() + 24 : 0);
    else if (a.bottom > s.top + sc.clientHeight) sc.scrollTop += a.bottom - (s.top + sc.clientHeight) + 2;
    if (a.left < left) sc.scrollLeft -= left - a.left;
    else if (a.right > s.left + sc.clientWidth) sc.scrollLeft += a.right - (s.left + sc.clientWidth) + 2;
  }
  // 上に止めてある見出しの行（日付・項目名）の高さ
  const headH = () => { let h = 0; for (let i = opt.noLetters ? 0 : 1; i < opt.headRows; i++) h += table.rows[i].offsetHeight; return h; };

  /* ---------- 選ぶ・動く ---------- */
  function go(r, c, extend) {
    // 1つずつ上・下へ出たら、となりの表へ
    if (!extend && opt.onEdge && Math.abs(r - st.r) === 1 && (r < 0 || r >= nRows) && opt.onEdge(r < 0 ? -1 : 1, st.c)) return;
    st.r = clampR(r); st.c = clampC(c);
    if (!extend) { st.ar = st.r; st.ac = st.c; }
    place(); scrollIntoView();
  }
  function moveIn(dr, dc) {
    // 範囲を選んでいるときの Enter／Tab は範囲の中を回る（Excel と同じ）
    if (top() !== bot() || lef() !== rig()) {
      let r = st.r, c = st.c;
      if (dr) { r += dr; if (r > bot()) { r = top(); c = c + 1 > rig() ? lef() : c + 1; } if (r < top()) { r = bot(); c = c - 1 < lef() ? rig() : c - 1; } }
      else { c += dc; if (c > rig()) { c = lef(); r = r + 1 > bot() ? top() : r + 1; } if (c < lef()) { c = rig(); r = r - 1 < top() ? bot() : r - 1; } }
      st.r = r; st.c = c; place(); scrollIntoView(); return;
    }
    go(st.r + dr, st.c + dc);
  }
  // Ctrl+矢印：データの端へ
  function jump(dr, dc, extend) {
    let r = st.r, c = st.c;
    const has = (r, c) => opt.get(r, c) !== "";
    const nr = r + dr, ncl = c + dc;
    const inside = (r, c) => r >= 0 && r < nRows && c >= 0 && c < NC;
    if (!inside(nr, ncl)) return;
    if (has(r, c) && has(nr, ncl)) { while (inside(r + dr, c + dc) && has(r + dr, c + dc)) { r += dr; c += dc; } }
    else { r = nr; c = ncl; while (inside(r, c) && !has(r, c) && inside(r + dr, c + dc)) { r += dr; c += dc; } }
    go(r, c, extend);
  }

  /* ---------- 直す ---------- */
  function startEdit(mode, initial) {
    if (!editable(st.c)) { opt.toast && opt.toast("このセルは直せません"); return false; }
    st.edit = { mode, r: st.r, c: st.c, orig: opt.get(st.r, st.c) };
    ed.classList.add("on");
    if (initial !== undefined) ed.value = initial;
    else ed.value = st.edit.orig;
    place();
    ed.focus();
    const n = ed.value.length; ed.setSelectionRange(n, n);
    if (opt.onEditText) opt.onEditText(ed.value);
    if (opt.onMode) opt.onMode(mode);
    return true;
  }
  function commit(all) {
    if (!st.edit) return;
    const v = opt.normalize ? opt.normalize(st.edit.c, ed.value) : ed.value;
    const e = st.edit;
    endEdit();
    if (all && (top() !== bot() || lef() !== rig())) {
      const ch = [];
      for (let r = top(); r <= bot(); r++) for (let c = lef(); c <= rig(); c++) if (editable(c)) ch.push({ r, c, v: opt.normalize ? opt.normalize(c, v) : v });
      opt.apply(ch);
    } else if (v !== e.orig) opt.apply([{ r: e.r, c: e.c, v }]);
  }
  function cancel() { endEdit(); place(); }
  function endEdit() {
    st.edit = null; ed.classList.remove("on"); ed.value = ""; hideList();
    if (opt.onMode) opt.onMode(null);
    el.focus({ preventScroll: true }); ed.focus({ preventScroll: true });
  }
  function clearRange() {
    const ch = [];
    for (let r = top(); r <= bot(); r++) for (let c = lef(); c <= rig(); c++) if (editable(c) && opt.get(r, c) !== "") ch.push({ r, c, v: "" });
    if (ch.length) opt.apply(ch);
  }
  // オートコンプリート（その列で前に打った言葉。打った字で始まるものが1つに決まれば続きを選んだ状態で出す）
  function autoComplete() {
    if (!st.edit || st.edit.mode !== "enter" || composing || !opt.complete) return;
    if (/^delete/.test(lastInputType)) return;
    const v = ed.value;
    if (!v || ed.selectionStart !== v.length) return;
    const hit = opt.complete(st.edit.c, v);
    if (hit && hit.length > v.length) { ed.value = hit; ed.setSelectionRange(v.length, hit.length); }
  }

  /* ---------- 一覧（Alt+↓） ---------- */
  function showList() {
    if (!opt.list || !editable(st.c)) return;
    const items = opt.list(st.c);
    if (!items.length) { opt.toast && opt.toast("この列にはまだ候補がありません"); return; }
    const A = rectOf(st.r, st.c, st.r, st.c);
    st.list = { items, i: Math.max(0, items.indexOf(opt.get(st.r, st.c))) };
    listBox.innerHTML = items.map((t, i) => `<div data-i="${i}" class="${i === st.list.i ? "on" : ""}">${escH(t)}</div>`).join("");
    Object.assign(listBox.style, { left: A.x + "px", top: A.y + A.h + "px", minWidth: A.w + "px" });
    listBox.hidden = false;
    listBox.children[st.list.i] && listBox.children[st.list.i].scrollIntoView({ block: "nearest" });
  }
  function hideList() { st.list = null; listBox.hidden = true; }
  function pickList(i) {
    const v = st.list.items[i]; hideList();
    if (st.edit) { ed.value = v; commit(); } else opt.apply([{ r: st.r, c: st.c, v }]);
    place();
  }
  listBox.addEventListener("mousedown", e => { const d = e.target.closest("[data-i]"); if (d) { e.preventDefault(); pickList(+d.dataset.i); } });

  /* ---------- コピー・貼り付け（タブ区切り。Excel と行き来できる） ---------- */
  function selText() {
    const out = [];
    for (let r = top(); r <= bot(); r++) { const row = []; for (let c = lef(); c <= rig(); c++) row.push(opt.get(r, c)); out.push(row.join("\t")); }
    return out.join("\r\n");
  }
  let clip = null;   // 自分でコピーした範囲（点線の枠）
  function markClip(cut) {
    clip = { r0: top(), c0: lef(), r1: bot(), c1: rig(), cut };
    const R = rectOf(clip.r0, clip.c0, clip.r1, clip.c1);
    let ant = el.querySelector(".xlant");
    if (!ant) { ant = document.createElement("div"); ant.className = "xlant"; el.appendChild(ant); }
    Object.assign(ant.style, { left: R.x - 1 + "px", top: R.y - 1 + "px", width: R.w + 1 + "px", height: R.h + 1 + "px", display: "block" });
  }
  function unmarkClip() { clip = null; const ant = el.querySelector(".xlant"); if (ant) ant.style.display = "none"; }
  function parseTSV(t) {
    t = String(t).replace(/\r\n?/g, "\n").replace(/\n$/, "");
    // 引用符で囲まれたセル（Excel がセルの中の改行を出すとき）
    const rows = []; let row = [], cur = "", q = false;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (q) { if (ch === '"' && t[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch === "\n" ? " " : ch; }
      else if (ch === '"' && cur === "") q = true;
      else if (ch === "\t") { row.push(cur); cur = ""; }
      else if (ch === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; }
      else cur += ch;
    }
    row.push(cur); rows.push(row);
    return rows;
  }
  function pasteText(t) {
    const rows = parseTSV(t);
    const ch = [];
    const one = rows.length === 1 && rows[0].length === 1;
    if (one && (top() !== bot() || lef() !== rig())) {
      for (let r = top(); r <= bot(); r++) for (let c = lef(); c <= rig(); c++) if (editable(c)) ch.push({ r, c, v: opt.normalize ? opt.normalize(c, rows[0][0]) : rows[0][0] });
    } else {
      rows.forEach((row, i) => row.forEach((v, j) => {
        const r = top() + i, c = lef() + j;
        if (c < NC && editable(c)) ch.push({ r, c, v: opt.normalize ? opt.normalize(c, v) : v });
      }));
      st.r = top(); st.c = lef();
      st.ar = Math.min(top() + rows.length - 1, nRows - 1 + rows.length);
      st.ac = clampC(lef() + Math.max(...rows.map(x => x.length)) - 1);
    }
    if (clip && clip.cut) {
      // 切り取り：元の場所を消す（貼った場所と重なるセルは貼った値のまま）
      const dst = new Set(ch.map(x => x.r + ":" + x.c));
      for (let r = clip.r0; r <= clip.r1; r++) for (let c = clip.c0; c <= clip.c1; c++) if (!dst.has(r + ":" + c) && editable(c)) ch.push({ r, c, v: "" });
      unmarkClip();
    }
    if (ch.length) opt.apply(ch, { grow: true });
  }
  document.addEventListener("copy", e => {
    if (!active() || st.edit) return;
    e.preventDefault(); e.clipboardData.setData("text/plain", selText()); markClip(false);
  });
  document.addEventListener("cut", e => {
    if (!active() || st.edit) return;
    e.preventDefault(); e.clipboardData.setData("text/plain", selText()); markClip(true);
  });
  document.addEventListener("paste", e => {
    if (!active() || st.edit) return;
    e.preventDefault(); pasteText(e.clipboardData.getData("text/plain"));
  });
  const active = () => document.activeElement === ed && !host.closest("[hidden]");

  /* ---------- フィル（右下の■をドラッグ） ---------- */
  handle.addEventListener("mousedown", e => {
    e.preventDefault(); e.stopPropagation();
    st.fill = { r0: top(), c0: lef(), r1: bot(), c1: rig(), tr: bot(), tc: rig() };
  });
  function fillTo(r, c) {
    const f = st.fill;
    // 縦か横の、離れているほうへ伸ばす
    const down = r - f.r1, right = c - f.c1;
    if (down >= right && down > 0) { f.tr = r; f.tc = f.c1; } else if (right > 0) { f.tc = c; f.tr = f.r1; } else { f.tr = f.r1; f.tc = f.c1; }
    const R = rectOf(f.r0, f.c0, f.tr, f.tc);
    if (R) { Object.assign(fillBox.style, { left: R.x - 1 + "px", top: R.y - 1 + "px", width: R.w + 1 + "px", height: R.h + 1 + "px" }); fillBox.hidden = false; }
  }
  function fillDone() {
    const f = st.fill; st.fill = null; fillBox.hidden = true;
    const ch = [], h = f.r1 - f.r0 + 1, w = f.c1 - f.c0 + 1;
    for (let r = f.r0; r <= f.tr; r++) for (let c = f.c0; c <= f.tc; c++) {
      if (r <= f.r1 && c <= f.c1) continue;
      if (editable(c)) ch.push({ r, c, v: opt.get(f.r0 + (r - f.r0) % h, f.c0 + (c - f.c0) % w) });
    }
    st.r = f.r0; st.c = f.c0; st.ar = f.tr; st.ac = f.tc;
    if (ch.length) opt.apply(ch); else place();
  }

  /* ---------- マウス ---------- */
  function hit(e) {
    const td = e.target.closest("td, th");
    if (!td || !table.contains(td)) return null;
    const tr = td.parentElement, ri = tr.rowIndex - opt.headRows, ci = td.cellIndex - 1;
    if (td.dataset.col !== undefined) return { col: +td.dataset.col };
    if (td.dataset.row !== undefined) return { row: +td.dataset.row };
    if (td.dataset.all) return { all: true };
    if (ri < 0 || ci < 0) return { head: true };
    return { r: ri, c: ci };
  }
  // 常用日報（2026-10-31）：iPad・スマホは、選んでいるセルをもう一度タップすると中を直す（ダブルタップは画面が拡大されるため）
  let lastPtr = "mouse";
  el.addEventListener("pointerdown", e => { lastPtr = e.pointerType || "mouse"; }, true);
  el.addEventListener("mousedown", e => {
    if (e.target === ed || listBox.contains(e.target)) {
      if (e.target === ed && lastPtr !== "mouse" && !st.edit) startEdit("edit");
      return;
    }
    const h = hit(e);
    if (!h) return;
    if (lastPtr !== "mouse" && !st.edit && h.r === st.r && h.c === st.c && top() === bot() && lef() === rig()) { e.preventDefault(); startEdit("edit"); return; }
    if (e.button === 2) {
      // 右クリック：範囲の外なら そこを選ぶ
      if (h.r !== undefined && (h.r < top() || h.r > bot() || h.c < lef() || h.c > rig())) { if (st.edit) commit(); go(h.r, h.c); }
      if (h.row !== undefined && (h.row < top() || h.row > bot())) { st.r = st.ar = h.row; st.c = NC - 1; st.ac = 0; place(); }
      return;
    }
    e.preventDefault();
    if (st.edit) commit();
    hideList(); unmarkClip();
    if (h.all) { st.ar = 0; st.ac = 0; st.r = nRows - 1; st.c = NC - 1; place(); }
    else if (h.col !== undefined) { st.ar = 0; st.r = nRows - 1; if (!e.shiftKey) st.ac = h.col; st.c = h.col; place(); st.drag = { col: true }; }
    else if (h.row !== undefined) { st.ac = 0; st.c = NC - 1; if (!e.shiftKey) st.ar = h.row; st.r = h.row; place(); st.drag = { row: true }; }
    else if (h.r !== undefined) { if (e.shiftKey) { st.r = h.r; st.c = h.c; place(); } else go(h.r, h.c); st.drag = { cell: true }; }
    ed.focus({ preventScroll: true });
  });
  el.addEventListener("mousemove", e => {
    if (!st.drag && !st.fill) return;
    const h = hit(e);
    if (!h) return;
    if (st.fill) { if (h.r !== undefined) fillTo(h.r, h.c); return; }
    if (st.drag.cell && h.r !== undefined) { st.r = h.r; st.c = h.c; place(); }
    else if (st.drag.col && h.col !== undefined) { st.c = h.col; place(); }
    else if (st.drag.row && h.row !== undefined) { st.r = h.row; place(); }
  });
  document.addEventListener("mouseup", () => { if (st.fill) fillDone(); st.drag = null; });
  el.addEventListener("dblclick", e => {
    // 選んでいるセルの上には打つ欄（見えない）が重なっているので、そこへのダブルクリックもそのセルを直す
    if (e.target === ed) { if (!st.edit) startEdit("edit"); return; }
    const h = hit(e);
    if (h && h.r !== undefined) { go(h.r, h.c); startEdit("edit"); }
  });
  el.addEventListener("contextmenu", e => {
    if (!opt.menu) return;
    e.preventDefault();
    opt.menu(e.clientX, e.clientY, { r0: top(), r1: bot(), c0: lef(), c1: rig() });
  });

  /* ---------- キー ---------- */
  ed.addEventListener("focus", () => { opt.onFocus && opt.onFocus(); });
  ed.addEventListener("compositionstart", () => {
    composing = true;
    if (!st.edit) { if (!startEdit("enter", "")) { ed.blur(); ed.focus(); } }
  });
  ed.addEventListener("compositionend", () => { composing = false; setTimeout(autoComplete, 0); });
  ed.addEventListener("input", e => {
    lastInputType = e.inputType || "";
    if (!st.edit) {
      // 選んだだけの状態で打った → 上書きで入力を始める
      const v = ed.value;
      if (!startEdit("enter", v)) { ed.value = ""; return; }
    }
    if (!composing) autoComplete();
    if (opt.onEditText) opt.onEditText(ed.value);
  });
  ed.addEventListener("keydown", e => {
    if (composing || e.isComposing || e.keyCode === 229) return;
    const k = e.key, ctrl = e.ctrlKey || e.metaKey, sh = e.shiftKey;
    // 一覧が開いているとき
    if (st.list) {
      if (k === "ArrowDown" || k === "ArrowUp") { e.preventDefault(); st.list.i = Math.max(0, Math.min(st.list.items.length - 1, st.list.i + (k === "ArrowDown" ? 1 : -1))); [...listBox.children].forEach((d, i) => d.classList.toggle("on", i === st.list.i)); listBox.children[st.list.i].scrollIntoView({ block: "nearest" }); return; }
      if (k === "Enter" || k === "Tab") { e.preventDefault(); pickList(st.list.i); return; }
      if (k === "Escape") { e.preventDefault(); hideList(); return; }
    }
    if (st.edit) {
      if (k === "Enter") { e.preventDefault(); if (e.altKey) return; commit(ctrl); if (!ctrl) moveIn(sh ? -1 : 1, 0); else place(); return; }
      if (k === "Tab") { e.preventDefault(); commit(); moveIn(0, sh ? -1 : 1); return; }
      if (k === "Escape") { e.preventDefault(); cancel(); return; }
      if (k === "F2") { e.preventDefault(); st.edit.mode = st.edit.mode === "enter" ? "edit" : "enter"; if (opt.onMode) opt.onMode(st.edit.mode); return; }
      if (k === "ArrowDown" && e.altKey) { e.preventDefault(); showList(); return; }
      if (st.edit.mode === "enter" && /^Arrow/.test(k)) {
        e.preventDefault(); commit();
        const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
        go(st.r + d[0], st.c + d[1], sh); return;
      }
      if (ctrl && (k === ";" || k === ":" || k === "*")) { e.preventDefault(); insertNow(k); return; }
      return;   // 字の中を動く・消す などは入力欄のまま
    }
    // ---- 選んでいるだけのとき ----
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) {
      e.preventDefault();
      if (k === "ArrowDown" && e.altKey) { showList(); return; }
      const [dr, dc] = arrows[k];
      if (ctrl) jump(dr, dc, sh); else go(st.r + dr, st.c + dc, sh);
      return;
    }
    if (k === "Enter") { e.preventDefault(); moveIn(sh ? -1 : 1, 0); return; }
    if (k === "Tab") { e.preventDefault(); moveIn(0, sh ? -1 : 1); return; }
    if (k === "Home") { e.preventDefault(); go(ctrl ? 0 : st.r, 0, sh); return; }
    if (k === "End" && ctrl) { e.preventDefault(); let lr = 0, lc = 0; for (let r = 0; r < nRows; r++) for (let c = 0; c < NC; c++) if (opt.get(r, c) !== "") { lr = Math.max(lr, r); lc = Math.max(lc, c); } go(lr, lc, sh); return; }
    if (k === "PageDown" || k === "PageUp") {
      e.preventDefault();
      // Ctrl+PageUp/Down はブラウザがタブの切り替えに使うことがあるので、Alt+PageUp/Down でも日を移る
      if (ctrl || e.altKey) { opt.onSheet && opt.onSheet(k === "PageDown" ? 1 : -1); return; }
      go(st.r + (k === "PageDown" ? 15 : -15), st.c, sh); return;
    }
    if (k === "F2") { e.preventDefault(); startEdit("edit"); return; }
    if (k === "Delete") { e.preventDefault(); clearRange(); return; }
    if (k === "Backspace") { e.preventDefault(); if (startEdit("enter", "")) {} return; }
    if (k === "Escape") { unmarkClip(); hideList(); return; }
    if (ctrl) {
      const lk = k.toLowerCase();
      if (lk === "a") { e.preventDefault(); st.ar = 0; st.ac = 0; st.r = nRows - 1; st.c = NC - 1; place(); return; }
      if (lk === "z") { e.preventDefault(); opt.undo && opt.undo(); return; }
      if (lk === "y") { e.preventDefault(); opt.redo && opt.redo(); return; }
      if (lk === "d") { e.preventDefault(); fillDir("down"); return; }
      if (lk === "r") { e.preventDefault(); fillDir("right"); return; }
      if (k === ";" || k === ":" || k === "*") { e.preventDefault(); if (startEdit("enter", "")) insertNow(k); return; }
      if ((k === "+" || k === "=" || k === ";") && sh) { e.preventDefault(); opt.rowOp && opt.rowOp("insert", top(), bot() - top() + 1); return; }
      if (k === "-") { e.preventDefault(); opt.rowOp && opt.rowOp("delete", top(), bot() - top() + 1); return; }
      if (lk === "p") { e.preventDefault(); opt.print && opt.print(); return; }
      if (lk === "f") { e.preventDefault(); opt.find && opt.find(); return; }
      return;   // Ctrl+C・X・V はブラウザの copy／cut／paste で受ける
    }
    if (e.altKey) return;
    // 普通の字は input に入れさせる（input イベントで入力が始まる）
  });
  function insertNow(k) {
    const d = new Date();
    const s = k === ";" ? (d.getMonth() + 1) + "/" + d.getDate() : d.getHours() + ":" + String(d.getMinutes()).padStart(2, "0");
    const a = ed.selectionStart, b = ed.selectionEnd;
    ed.value = ed.value.slice(0, a) + s + ed.value.slice(b);
    ed.setSelectionRange(a + s.length, a + s.length);
  }
  function fillDir(dir) {
    const ch = [];
    if (dir === "down") {
      if (top() === bot()) { if (st.r === 0) return; for (let c = lef(); c <= rig(); c++) if (editable(c)) ch.push({ r: st.r, c, v: opt.get(st.r - 1, c) }); }
      else for (let r = top() + 1; r <= bot(); r++) for (let c = lef(); c <= rig(); c++) if (editable(c)) ch.push({ r, c, v: opt.get(top(), c) });
    } else {
      if (lef() === rig()) { if (st.c === 0) return; for (let r = top(); r <= bot(); r++) if (editable(st.c)) ch.push({ r, c: st.c, v: opt.get(r, st.c - 1) }); }
      else for (let r = top(); r <= bot(); r++) for (let c = lef() + 1; c <= rig(); c++) if (editable(c)) ch.push({ r, c, v: opt.get(r, lef()) });
    }
    if (ch.length) opt.apply(ch);
  }
  ed.addEventListener("blur", () => {
    // ほかの所（数式バー・ボタン）を押したとき：入力中なら確定する
    setTimeout(() => { if (st.edit && document.activeElement !== ed && document.activeElement !== opt.formula) commit(); }, 0);
  });

  /* ---------- 数式バー（上の入力欄）とつなぐ ---------- */
  if (opt.formula) {
    const fb = opt.formula;
    const mine = () => !opt.isCur || opt.isCur();
    fb.addEventListener("focus", () => {
      if (!mine()) return;
      if (!editable(st.c)) { fb.blur(); return; }
      if (!st.edit) { st.edit = { mode: "edit", r: st.r, c: st.c, orig: opt.get(st.r, st.c) }; ed.value = st.edit.orig; ed.classList.add("on"); place(); }
    });
    fb.addEventListener("input", () => { if (mine() && st.edit) ed.value = fb.value; });
    fb.addEventListener("keydown", e => {
      if (!mine() || e.isComposing || e.keyCode === 229) return;
      if (e.key === "Enter") { e.preventDefault(); ed.value = fb.value; commit(); moveIn(e.shiftKey ? -1 : 1, 0); ed.focus(); }
      if (e.key === "Escape") { e.preventDefault(); cancel(); ed.focus(); }
      if (e.key === "Tab") { e.preventDefault(); ed.value = fb.value; commit(); moveIn(0, e.shiftKey ? -1 : 1); ed.focus(); }
    });
  }

  render();
  return {
    render, place,
    focus() { ed.focus({ preventScroll: true }); },
    select(r, c) { if (st.edit) commit(); go(r, c); ed.focus({ preventScroll: true }); },
    // となりの表から入ってきた（dir 1＝上から下へ → 1行目、-1＝下から上へ → 最後の行）
    enter(dir, c) {
      if (st.edit) commit();
      nRows = opt.rows();
      st.r = st.ar = dir > 0 ? 0 : nRows - 1; st.c = st.ac = clampC(c);
      ed.focus({ preventScroll: true }); place(); scrollIntoView();
    },
    selection: () => ({ r: st.r, c: st.c, r0: top(), r1: bot(), c0: lef(), c1: rig() }),
    setSelection(s) { st.r = s.r; st.c = s.c; st.ar = s.ar ?? s.r; st.ac = s.ac ?? s.c; place(); },
    editing: () => !!st.edit,
    commit: () => commit(),
    editText: () => ed.value,
    paste: pasteText, copyText: selText,
    showList
  };
}
