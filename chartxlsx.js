/* 工程表（いまの画面の形）を Excel（.xlsx）で出す（2026-10-31 ユーザー指定）
   元にした書式：\\192.168.0.15\share\03_社員別フォルダ\北澤\会社全体工程表.xls の「工程表 11-12月」
   （左に No・現場［コード・社内工事名・担当］・工種、右に日の列。月の見出し・日・曜日）
   - 1シートに、印刷の小窓でチェックした現場と期間。工種は画面と同じ（完了の工種は画面で出しているときだけ）。たたみは無視して全部ひらく
   - 工程線は図形（四角。セルに合わせて動く・伸び縮みする）。色は全体工程の線の色。休工日は外枠だけ残して白く抜く（画面と同じ）
   - 1つの工程（行）＝2行：上の行に工程線、下の行に「作業 人数 社員」の文字（右の空いたセルへはみ出して読める）
   - 休みの設定は週間工程表と同じ：表の下（印刷の外）の「休日→」の行で「休」を選ぶ／消すと、その列が灰色のぼかし（日付・曜日は赤字）になる／戻る。
     はじめの「休」は会社の休み・日曜・祝日。土曜の日付は青字
   weekxlsx.js の xlCol・xlRef・xlEsc・xlDate・zipStore・loadLogo・downloadBlob を使う */

const CX_FIRST = 4;   // D 列から日付（A＝No、B＝現場、C＝工種）
const CX_R0 = 6;      // 6 行目から現場（1＝題、2＝期間、3＝月、4＝日、5＝曜）

// 書式は使うたびに番号を振る（同じ書式は同じ番号）
function cxStyles() {
  const fonts = [], borders = [], xfs = [], fi = {}, bi = {}, xi = {};
  const font = (sz, b, color) => {
    const k = [sz, b, color].join("|");
    if (fi[k] == null) { fi[k] = fonts.length; fonts.push(`<font>${b ? "<b/>" : ""}<sz val="${sz}"/>${color ? `<color rgb="${color}"/>` : ""}<name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>`); }
    return fi[k];
  };
  // 線は "thin"／"medium"／"hair:FFBFBFBF"（色つき）の形
  const side = (tag, v) => { if (!v) return `<${tag}/>`; const [st, col] = v.split(":"); return `<${tag} style="${st}">${col ? `<color rgb="${col}"/>` : '<color auto="1"/>'}</${tag}>`; };
  const border = (l, r, t, b, diag) => {
    const k = [l, r, t, b, diag].join("|");
    if (bi[k] == null) { bi[k] = borders.length; borders.push(`<border${diag ? ' diagonalDown="1"' : ""}>${side("left", l)}${side("right", r)}${side("top", t)}${side("bottom", b)}${diag ? side("diagonal", "thin") : "<diagonal/>"}</border>`); }
    return bi[k];
  };
  const style = o => {
    const f = font(o.sz || 9, !!o.b, o.color || ""), bd = border(o.l || "", o.r || "", o.t || "", o.bt || "", !!o.diag);
    const al = `<alignment horizontal="${o.h || "general"}" vertical="${o.v || "center"}"${o.wrap ? ' wrapText="1"' : ""}${o.shrink ? ' shrinkToFit="1"' : ""}/>`;
    const k = [f, bd, al, o.fmt || 0].join("|");
    if (xi[k] == null) { xi[k] = xfs.length; xfs.push(`<xf numFmtId="${o.fmt || 0}" fontId="${f}" fillId="0" borderId="${bd}" applyFont="1" applyBorder="1" applyAlignment="1"${o.fmt ? ' applyNumberFormat="1"' : ""}>${al}</xf>`); }
    return xi[k];
  };
  style({ sz: 11 });   // 0＝標準
  // 条件付き書式：0 灰色のぼかし（休）、1 赤字（日曜・休）、2 青字（土曜）、3 月が替わる所の左の線
  const dxfs = [
    `<dxf><fill><gradientFill><stop position="0"><color rgb="FFBFBFBF"/></stop><stop position="0.5"><color rgb="FFFFFFFF"/></stop><stop position="1"><color rgb="FFBFBFBF"/></stop></gradientFill></fill></dxf>`,
    `<dxf><font><b/><color rgb="FFFF0000"/></font></dxf>`,
    `<dxf><font><b/><color rgb="FF0070C0"/></font></dxf>`,
    `<dxf><border><left style="thin"><color auto="1"/></left></border></dxf>`
  ];
  const xml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="d"/></numFmts>
<fonts count="${fonts.length}">${fonts.join("")}</fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="${borders.length}">${borders.join("")}</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>
<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>
</styleSheet>`;
  return { style, xml };
}

// 令和の年月日（期間の見出し）
const cxWareki = k => { const y = +k.slice(0, 4); return (y >= 2019 ? "令和" + (y - 2018 === 1 ? "元" : y - 2018) : "平成" + (y - 1988)) + "年" + (+k.slice(5, 7)) + "月" + (+k.slice(8, 10)) + "日（" + WD[dayDow(k)] + "）"; };

/* 用紙いっぱいに出す（2026-10-31 ユーザー：極力、印刷時の縮小は使わない・現場が少ないときは行の高さで下の余白を無くす）
   列の幅・行の高さを、用紙（A3／A4 横・余白を除く）に合わせて決める（Excel の画素：1px＝0.75pt、ＭＳ Ｐゴシック 11 で列の幅 1＝8px）。
   Excel は「100%」でも、プリンターによって決めた大きさより 8〜10% 小さく印刷する（この PC の Adobe PDF で 横 0.92・縦 0.90 を測った）。
   その差でできる下・右の余白を無くすため、用紙より 横 8.5%・縦 10% 大きく組み、「1ページに合わせる」で用紙に収める
   （この PC では ほぼ 100%、決めたとおりに印刷するプリンターでは 90% 前後）。
   1日の列が 12px より細くなる・工程の行が 28px より低くなるときは、その大きさで組んで「1ページに合わせる」でもっと縮める。
   小窓のページ分け（期間を nx・現場を ny ページ）は、1ページ＝1シートにする（「ページに合わせる」では手で入れた改ページが効かないため。現場の途中で切らない） */
const CX_PAPER = { A3: [1190.55, 841.89], A4: [841.89, 595.28] };   // 横向きの幅・高さ（pt）
const CX_MARGIN = { l: 0.25, r: 0.25, t: 0.3, b: 0.3 };              // 余白（インチ）
const CX_KW = 1.085, CX_KH = 1.10;                                    // 用紙より大きく組む割合（上の説明）
const CX_LEFT_PX = { 1: 28, 2: 128, 3: 112 };                       // No・現場・工種の列（px）
const CX_HEAD_PX = { 1: 40, 2: 24, 3: 21, 4: 17, 5: 16 };           // 題・期間・月・日・曜の行（px）
const CX_TOT_PX = 20, CX_END_PX = 5;                                  // 合計の行・表の下の線の行
const CX_MIN_DAY = 12, CX_MIN_PAIR = 28, CX_MAX_PAIR = 400;          // 1日の列・工程の行（線＋文字）の下限・上限（px）

/* list＝現場（並び済み）、from〜end＝期間、logo＝{ bytes, w, h } か null */
function chartXlsx(list, from, end, logo) {
  const days = []; for (let k = from; k <= end; k = dayAdd(k, 1)) days.push(k);
  const N = days.length;
  const paperKey = ($("inPaper").value || "A3") === "A3" ? "A3" : "A4", [PW, PH] = CX_PAPER[paperKey];
  const nx = Math.min(+$("pNx").value || 1, N), nyWant = +$("pNy").value || 1;
  const pageWpx = Math.floor((PW - (CX_MARGIN.l + CX_MARGIN.r) * 72) * CX_KW / 0.75);
  const pageHpx = Math.floor((PH - (CX_MARGIN.t + CX_MARGIN.b) * 72) * CX_KH / 0.75);

  // 現場ごとの工種と行の数（1つの工種に重なる工程があれば行を足す：packLane）
  const siteBlocks = list.map(s => {
    const lanes = s.段.filter(l => !l.完了 || showLanes.has(s.id));
    const blocks = (lanes.length ? lanes : [null]).map(ln => {
      const bars = ln ? s.バー.filter(b => b.段 === ln.id && b.終了 >= from && b.開始 <= end) : [];
      const p = packLane(bars);
      return { ln, bars, pos: p.pos, n: p.rows };
    });
    return { s, blocks, pairs: blocks.reduce((a, b) => a + b.n, 0) };
  });
  // 期間を nx に分ける（ちょうど○か月なら月の切れ目：印刷と同じ考え方。ほかは日数で等分）
  const dayRanges = [];
  { const months = (() => { const m = []; let k = from; while (k <= end) { const z = dayAdd(addMonths(k, 1), -1); m.push([k, z > end ? end : z]); k = dayAdd(z, 1); } return m; })();
    if (nx > 1 && months.length % nx === 0 && dayAdd(addMonths(from, months.length), -1) === end) {
      const per = months.length / nx;
      for (let i = 0; i < nx; i++) dayRanges.push([months[i * per][0], months[i * per + per - 1][1]]);
    } else {
      const dp = Math.ceil(N / nx);
      for (let i = 0; i < N; i += dp) dayRanges.push([days[i], days[Math.min(N, i + dp) - 1]]);
    } }
  // 現場を ny に分ける（工程の行の数がだいたい同じになるよう、並びのまま）
  const totalPairs = siteBlocks.reduce((a, x) => a + x.pairs, 0);
  const ny = Math.max(1, Math.min(nyWant, siteBlocks.length));
  const groups = [];
  { let g = [], acc = 0, gi = 0;
    siteBlocks.forEach((x, i) => {
      g.push(x); acc += x.pairs;
      const left = siteBlocks.length - i - 1, need = ny - gi - 1;
      if (need > 0 && (acc >= totalPairs * (gi + 1) / ny || left === need)) { groups.push(g); g = []; gi++; }
    });
    if (g.length || !groups.length) groups.push(g); }

  // 合計（作業員［自社］・外注［外注の工種があるときだけ］・職員［社員の人数。同じ人は1人］）。チェックした現場ぜんぶの、その日に作業する工程（休工日・会社の休みは数えない）
  const hasGai = list.some(s => s.段.some(isGai));
  const totals = ["作業員", ...(hasGai ? ["外注"] : []), "職員"];
  const at = {};
  for (const s of list) for (const b of s.バー) {
    if (b.終了 < from || b.開始 > end) continue;
    const ln = s.段.find(l => l.id === b.段);
    for (let k = b.開始 < from ? from : b.開始; k <= b.終了 && k <= end; k = dayAdd(k, 1)) if (barWorks(ln, b, k)) (at[k] = at[k] || []).push({ b, ln });
  }
  const totalOf = (name, k) => {
    const xs = at[k] || [];
    if (name === "職員") return new Set(xs.flatMap(x => x.b.社員 || [])).size;
    return xs.filter(x => name === "外注" ? isGai(x.ln) : !isGai(x.ln)).reduce((a, x) => a + headAt(x.b, k, x.ln), 0);
  };

  const ST = cxStyles(), noOf = new Map(list.map((s, i) => [s.id, i + 1]));   // No は通し番号
  const sheets = [];
  groups.forEach((g, gi) => dayRanges.forEach(([a, z], xi) => {
    const name = groups.length * dayRanges.length === 1 ? "全体工程表" :
      `${+a.slice(5, 7)}月${+a.slice(8, 10)}日～` + (groups.length > 1 ? ` その${gi + 1}` : "");   // 「(1／2)」の形は Excel が開けずに落ちた（2026-10-31 試し）
    sheets.push({ name, ...cxSheet({ ST, g, a, z, noOf, lastGroup: gi === groups.length - 1, totals, totalOf, pageWpx, pageHpx, logo, from, end }) });
  }));

  const rel = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
  const files = [
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join("") + `</Types>`],
    ["_rels/.rels", rel + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>` +
      sheets.map((x, i) => `<sheet name="${xlEsc(x.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") + `</sheets><definedNames>` +
      sheets.map((x, i) => `<definedName name="_xlnm.Print_Area" localSheetId="${i}">'${xlEsc(x.name)}'!${x.printArea}</definedName>`).join("") + `</definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`],
    ["xl/_rels/workbook.xml.rels", rel + sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") +
      `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml", ST.xml()]
  ];
  sheets.forEach((x, i) => {
    const n = i + 1;
    files.push([`xl/worksheets/sheet${n}.xml`, x.xml]);
    files.push([`xl/worksheets/_rels/sheet${n}.xml.rels`, rel + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${n}.xml"/></Relationships>`]);
    files.push([`xl/drawings/drawing${n}.xml`, x.drawing]);
    files.push([`xl/drawings/_rels/drawing${n}.xml.rels`, rel + (logo ? `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.png"/>` : "") + `</Relationships>`]);
  });
  if (logo) files.push(["xl/media/logo.png", logo.bytes]);   // ロゴは全シートで1つを使う
  return zipStore(files);
}

/* 1ページ＝1シート。g＝このページの現場、a〜z＝このページの期間、lastGroup＝現場の最後のページ（下に合計） */
function cxSheet({ ST, g, a, z, noOf, lastGroup, totals, totalOf, pageWpx, pageHpx, logo, from, end }) {
  const days = []; for (let k = a; k <= z; k = dayAdd(k, 1)) days.push(k);
  const N = days.length, lastCol = CX_FIRST + N - 1, Z = xlCol(lastCol);
  const sty = ST.style;
  const leftPx = CX_LEFT_PX[1] + CX_LEFT_PX[2] + CX_LEFT_PX[3];
  const tot = lastGroup ? totals : [];
  const n = Math.max(1, g.reduce((s, x) => s + x.pairs, 0));
  const fixH = [1, 2, 3, 4, 5].reduce((s, r) => s + CX_HEAD_PX[r], 0) + tot.length * CX_TOT_PX + CX_END_PX;
  // 1日の列・工程の行が下限より小さくなるなら、縦横とも同じ割合で大きく組む（縮めたときに片方だけ余白が出ないように）
  const k = Math.max(1, (leftPx + N * CX_MIN_DAY) / pageWpx, (fixH + n * CX_MIN_PAIR) / pageHpx);
  const W = Math.floor(pageWpx * k), Hh = Math.floor(pageHpx * k);
  // 列の幅：1日の幅は用紙の横いっぱい。割り切れない画素は左の列から 1px ずつ足す
  const dayBase = Math.max(CX_MIN_DAY, Math.floor((W - leftPx) / N)), dayExtra = Math.max(0, W - leftPx - dayBase * N);
  // 行の高さ：題・期間・見出し・（合計）を除いた残りを、工程の行（線＋文字）で等分。余った画素は上の行から 1px ずつ
  const avail = Hh - fixH;
  const base = Math.max(CX_MIN_PAIR, Math.floor(avail / n)), extra = avail - base * n;
  const pairH = Array.from({ length: n }, (_, i) => Math.min(CX_MAX_PAIR, base + (i < extra ? 1 : 0)));
  const colPx = c => CX_LEFT_PX[c] || dayBase + (c - CX_FIRST < dayExtra ? 1 : 0);
  const colPt = c => colPx(c) * 0.75;

  const GRID = "hair:FFBFBFBF";
  const S_ = {
    title: sty({ sz: 18, b: true, h: "left" }),
    period: sty({ sz: 11, h: "left" }),
    made: sty({ sz: 9, h: "right" }),
    corner: sty({ sz: 9, h: "center", l: "thin", r: "thin", t: "medium", bt: "thin" }),
    month: sty({ sz: 10, b: true, h: "left", t: "medium", bt: GRID }),
    monthR: sty({ sz: 10, b: true, h: "left", t: "medium", bt: GRID, r: "thin" }),
    dnum: sty({ sz: 8, h: "center", fmt: 164, shrink: true, l: GRID, r: GRID, bt: GRID }),
    dnumR: sty({ sz: 8, h: "center", fmt: 164, shrink: true, l: GRID, r: "thin", bt: GRID }),
    dwk: sty({ sz: 8, h: "center", shrink: true, l: GRID, r: GRID, bt: "thin" }),
    dwkR: sty({ sz: 8, h: "center", shrink: true, l: GRID, r: "thin", bt: "thin" }),
    help: sty({ sz: 8, color: "FF808080", h: "center" }),
    helpDate: sty({ sz: 8, color: "FF808080", h: "center", fmt: 164, shrink: true }),
    helpBox: sty({ sz: 8, h: "center", l: "thin", r: "thin", t: "thin", bt: "thin", shrink: true }),
    endRow: sty({ t: "medium" })
  };
  // 日の欄：上の線（なし／工種の境 thin／現場の境 medium）、右端の列は右に線。文字の行は上寄せ（線のすぐ下に読める）
  const dayStyle = (top, right, lbl) => sty({ sz: 9, h: "left", v: lbl ? "top" : "center", l: GRID, r: right ? "thin" : GRID, t: top });
  const siteStyle = top => sty({ sz: 10, h: "left", v: "center", wrap: true, l: "thin", r: "thin", t: top, bt: "thin" });
  const noStyle = top => sty({ sz: 9, h: "center", l: "medium", r: "thin", t: top, bt: "thin" });
  const laneStyle = (top, night) => sty({ sz: 9, b: night, color: night ? "FF1F3A5F" : "", h: "left", wrap: true, l: "thin", r: "thin", t: top, bt: "thin" });
  const totLbl = top => sty({ sz: 10, b: true, h: "center", l: "medium", r: "thin", t: top, bt: "thin" });
  const totNum = (top, right) => sty({ sz: 9, b: true, h: "center", shrink: true, l: GRID, r: right ? "thin" : GRID, t: top, bt: "thin" });

  const rows = {};      // r → { c → {s, t, v} }
  const put = (r, c, o) => { const row = rows[r] || (rows[r] = {}); row[c] = Object.assign(row[c] || {}, o); };
  const merges = [];
  const box = (r1, c1, r2, c2, st, val) => {
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) put(r, c, { s: st });
    if (val !== undefined) put(r1, c1, val);
    if (r1 !== r2 || c1 !== c2) merges.push(xlRef(r1, c1) + ":" + xlRef(r2, c2));
  };
  const str = v => ({ t: "s", v });
  const num = v => ({ t: "n", v });
  const rowsPx = { ...CX_HEAD_PX };

  // 1・2 行：題・期間（印刷の小窓の期間全体）・作成日（右上にロゴ）
  put(1, 1, { s: S_.title, ...str("全　体　工　程　表") });
  put(2, 1, { s: S_.period, ...str("期間　" + cxWareki(from) + "　〜　" + cxWareki(end)) });
  put(2, lastCol, { s: S_.made, ...str(cxWareki(todayKey()) + " 作成") });
  // 3〜5 行：月・日・曜
  box(3, 1, 5, 1, S_.corner, str("No"));
  box(3, 2, 5, 3, S_.corner, str("現場　／　工種"));
  days.forEach((k, i) => {
    const c = CX_FIRST + i, R = c === lastCol;
    const m = i === 0 || k.endsWith("-01");
    put(3, c, { s: R ? S_.monthR : S_.month, ...(m ? str((i === 0 || k.slice(5, 7) === "01" ? +k.slice(0, 4) + "年" : "") + (+k.slice(5, 7)) + "月") : {}) });
    put(4, c, { s: R ? S_.dnumR : S_.dnum, ...num(xlDate(k)) });
    put(5, c, { s: R ? S_.dwkR : S_.dwk, ...str(WD[dayDow(k)]) });
  });

  // 現場ごと：工種ごとに、重なる工程の数だけ行（1行＝工程線の行＋文字の行）
  const shapes = [];   // { r, c0, c1, on, color, name }
  let r = CX_R0, pi = 0;
  for (const { s, blocks } of g) {
    const r0 = r, total = blocks.reduce((t, b) => t + b.n * 2, 0);
    const siteTxt = [s.コード || "", nm(s) + (s.完了 ? "（完了）" : ""), s.担当 ? "担当：" + s.担当 : ""].filter(Boolean).join("\n");
    box(r0, 1, r0 + total - 1, 1, noStyle("medium"), num(noOf.get(s.id)));
    box(r0, 2, r0 + total - 1, 2, siteStyle("medium"), str(siteTxt));
    blocks.forEach((bk, bi) => {
      const top = bi === 0 ? "medium" : "thin";
      const night = !!(bk.ln && isNight(bk.ln));
      box(r, 3, r + bk.n * 2 - 1, 3, laneStyle(top, night), bk.ln ? str(laneName(bk.ln) + (isGai(bk.ln) ? "\n（外注 " + gaiName(bk.ln) + "）" : "")) : {});
      for (let j = 0; j < bk.n; j++) {
        const rb = r + j * 2, rt = rb + 1, ph = pairH[pi++];
        rowsPx[rb] = Math.round(ph * 0.42); rowsPx[rt] = ph - rowsPx[rb];   // 上＝工程線、下＝文字
        days.forEach((k, i) => {
          const c = CX_FIRST + i, R = c === lastCol;
          put(rb, c, { s: dayStyle(j === 0 ? top : "", R, false) });
          put(rt, c, { s: dayStyle("", R, true) });
        });
      }
      for (const b of bk.bars) {
        if (b.終了 < a || b.開始 > z) continue;
        const ba = b.開始 < a ? a : b.開始, bz = b.終了 > z ? z : b.終了;
        const rb = r + bk.pos[b.id] * 2, col = colorOf(b.色).slice(1).toUpperCase();
        // 作業する日の続き＝塗った四角、休工日の続き＝外枠だけの白い四角
        let run = null;
        for (let k = ba; k <= bz; k = dayAdd(k, 1)) {
          const on = barWorks(bk.ln, b, k), c = CX_FIRST + dayDiff(a, k);
          if (run && run.on === on) run.c1 = c;
          else { if (run) shapes.push(run); run = { r: rb, c0: c, c1: c, on, color: col, name: (b.作業 || "工程") }; }
        }
        if (run) shapes.push(run);
        // 文字：作業・人数・社員（名字）。工程線の始まりの日（ページの最初の日）の下に
        const label = [b.作業 || "", headLabel(b, bk.ln), (b.社員 || []).map(shortStaff).join("・")].filter(Boolean).join(" ");
        if (label) put(rb + 1, CX_FIRST + dayDiff(a, ba), str(label));
      }
      r += bk.n * 2;
    });
  }
  // 合計（現場の最後のページだけ）
  tot.forEach((name, ti) => {
    const top = ti === 0 ? "medium" : "thin";
    box(r, 1, r, 3, totLbl(top), str(name + "　計"));
    days.forEach((k, i) => { const c = CX_FIRST + i, v = totalOf(name, k); put(r, c, { s: totNum(top, c === lastCol), ...(v ? num(v) : {}) }); });
    rowsPx[r] = CX_TOT_PX;
    r++;
  });
  const rLast = r - 1;
  for (let c = 1; c <= lastCol; c++) put(r, c, { s: S_.endRow });   // 表の下の線
  rowsPx[r] = CX_END_PX;
  const rEnd = r;

  // 下の補助（印刷の外）：日付・「休」・曜日。「休」を選ぶ／消すと、その列の灰色のぼかし（日付〜表の下）が付く／消える（週間工程表と同じ）
  const rHelp = r + 2, H = rHelp + 1;
  put(H, 3, { s: S_.help, ...str("休日→") });
  days.forEach((k, i) => {
    const c = CX_FIRST + i, L = xlCol(c);
    put(rHelp, c, { s: S_.helpDate, f: `${L}4` });
    put(H, c, { s: S_.helpBox, ...((isRest(k) || dayDow(k) === 0 || holidayOf(k)) ? str("休") : {}) });
    put(H + 1, c, { s: S_.help, ...str(WD[dayDow(k)]) });
  });
  rowsPx[H] = 24;
  const maxRow = H + 1;
  const rowPt = rr => (rowsPx[rr] || 18) * 0.75;

  // シートの XML
  const rowXml = [];
  for (let rr = 1; rr <= maxRow; rr++) {
    const row = rows[rr] || {};
    const body = Object.keys(row).map(Number).sort((p, q) => p - q).map(c => {
      const o = row[c], ref = xlRef(rr, c), st = ` s="${o.s || 0}"`;
      if (o.f) return `<c r="${ref}"${st}><f>${xlEsc(o.f)}</f></c>`;
      if (o.t === "n") return `<c r="${ref}"${st}><v>${o.v}</v></c>`;
      if (o.t === "s" && o.v !== "") return `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${xlEsc(o.v)}</t></is></c>`;
      return `<c r="${ref}"${st}/>`;
    }).join("");
    rowXml.push(`<row r="${rr}" ht="${rowPt(rr)}" customHeight="1">${body}</row>`);
  }
  const D = xlCol(CX_FIRST), E = xlCol(CX_FIRST + 1);
  const cf = [
    `<conditionalFormatting sqref="${D}4:${Z}${rLast}"><cfRule type="expression" dxfId="0" priority="1"><formula>${D}$${H}="休"</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="${D}4:${Z}5"><cfRule type="expression" dxfId="1" priority="2"><formula>OR(WEEKDAY(${D}$4)=1,${D}$${H}="休")*(WEEKDAY(${D}$4)&lt;&gt;7)</formula></cfRule><cfRule type="expression" dxfId="2" priority="3"><formula>WEEKDAY(${D}$4)=7</formula></cfRule></conditionalFormatting>`,
    N > 1 ? `<conditionalFormatting sqref="${E}3:${Z}${rLast}"><cfRule type="expression" dxfId="3" priority="4"><formula>DAY(${E}$4)=1</formula></cfRule></conditionalFormatting>` : ""
  ];
  const dv = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${D}${H}:${Z}${H}"><formula1>"休"</formula1></dataValidation></dataValidations>`;
  const paper = ($("inPaper").value || "A3") === "A3" ? 8 : 9;
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<dimension ref="A1:${xlRef(maxRow, lastCol)}"/>
<sheetViews><sheetView showGridLines="0" zoomScale="100" workbookViewId="0"><pane xSplit="3" ySplit="5" topLeftCell="${D}${CX_R0}" activePane="bottomRight" state="frozen"/><selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="${D}${CX_R0}" sqref="${D}${CX_R0}"/></sheetView></sheetViews>
<sheetFormatPr defaultColWidth="9" defaultRowHeight="13.5"/>
<cols>${[1, 2, 3].map(c => `<col min="${c}" max="${c}" width="${CX_LEFT_PX[c] / 8}" customWidth="1"/>`).join("")}${dayExtra ? `<col min="${CX_FIRST}" max="${CX_FIRST + dayExtra - 1}" width="${(dayBase + 1) / 8}" customWidth="1"/>` : ""}${dayExtra < N ? `<col min="${CX_FIRST + dayExtra}" max="${lastCol}" width="${dayBase / 8}" customWidth="1"/>` : ""}</cols>
<sheetData>${rowXml.join("")}</sheetData>
<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>
${cf.join("")}
${dv}
<pageMargins left="${CX_MARGIN.l}" right="${CX_MARGIN.r}" top="${CX_MARGIN.t}" bottom="${CX_MARGIN.b}" header="0.2" footer="0.2"/>
<pageSetup paperSize="${paper}" orientation="landscape" fitToWidth="1" fitToHeight="1"/>
<drawing r:id="rId1"/>
</worksheet>`;

  // 図形（工程線）とロゴ。工程線は「セルに合わせて移動・サイズ変更」（twoCellAnchor）。太さは工程線の行の半分（4〜9pt）で真ん中
  const EMU = 12700, colX = [0]; for (let c = 1; c <= lastCol + 1; c++) colX[c] = colX[c - 1] + colPt(c);
  const rowY = [0]; for (let rr = 1; rr <= maxRow + 1; rr++) rowY[rr] = rowY[rr - 1] + rowPt(rr);
  const anchors = [];
  let id = 2;
  for (const sh of shapes) {
    const rh = rowPt(sh.r), bh = Math.max(4, Math.min(9, rh * 0.5)), off = (rh - bh) / 2;
    const x = colX[sh.c0 - 1], w = colX[sh.c1] - x, y = rowY[sh.r - 1] + off;
    const fill = `<a:solidFill><a:srgbClr val="${sh.on ? sh.color : "FFFFFF"}"/></a:solidFill>`;
    anchors.push(`<xdr:twoCellAnchor><xdr:from><xdr:col>${sh.c0 - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${sh.r - 1}</xdr:row><xdr:rowOff>${Math.round(off * EMU)}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${sh.c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${sh.r - 1}</xdr:row><xdr:rowOff>${Math.round((off + bh) * EMU)}</xdr:rowOff></xdr:to>` +
      `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${xlEsc(sh.name + " " + id)}"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
      `<xdr:spPr><a:xfrm><a:off x="${Math.round(x * EMU)}" y="${Math.round(y * EMU)}"/><a:ext cx="${Math.round(w * EMU)}" cy="${Math.round(bh * EMU)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}<a:ln w="9525"><a:solidFill><a:srgbClr val="${sh.color}"/></a:solidFill></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`);
    id++;
  }
  if (logo) {   // 右上。高さ 20pt。右に 16pt あける（週間工程表と同じく PDF では画像が少し大きく出て切れるため）
    const h = 20, w = h * logo.w / logo.h, x = Math.max(colX[3], colX[lastCol] - w - 16);
    let c = 1; while (c < lastCol && colX[c] <= x) c++;
    anchors.push(`<xdr:oneCellAnchor><xdr:from><xdr:col>${c - 1}</xdr:col><xdr:colOff>${Math.round((x - colX[c - 1]) * EMU)}</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>${5 * EMU}</xdr:rowOff></xdr:from><xdr:ext cx="${Math.round(w * EMU)}" cy="${h * EMU}"/>` +
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="ロゴ"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
      `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(w * EMU)}" cy="${h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`);
  }
  const drawing = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${anchors.join("\n")}</xdr:wsDr>`;
  return { xml, drawing, printArea: `$A$1:$${Z}$${rEnd}` };
}
