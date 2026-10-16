/* 週間工程表を Excel（.xlsx）で出す（2026-10-13 ユーザー指定）
   元にした書式：現場別\510012金沢線配水\金沢週間工程.xlsx（1週＝1シート、月曜はじまり・31日分の列、A4 横）。
   - 作業は1つにつき4行（A:E を結合した作業名・2行目の日の欄に色の文字「青」など → 条件付き書式でその色に塗る。文字の色も同じにして見えなくする）
   - 下の 53〜55 行に 日付・「休」・曜日。「休」を選んだ列は、上の表がグラデーション（灰→白→灰の縦じま）になる条件付き書式
   - 地図：現場の設定で入れた地図を AA4 から 幅 247.5pt×高さ 528pt（AA〜AJ 列・4行目〜記事の下）に置く（2026-10-14 ユーザー指定）。無ければ空けておく
   2026-10-16 ユーザー指定で直した：
   - 3週間（F〜Z 列）にして、右の AA〜AJ は地図の枠（罫線なし・条件付き書式なし。外周の枠線だけ）。地図は AA4〜人員予定の下（セルに付けて置く）
   - 会社名はロゴの画像（logo.png、背景透過）を AA1:AJ1 に。担当は2行目
   - 工事名（縦書き）の右に 工事略称／現場コード（正式工事名は出さない）。工期も縦書き。表の外周に枠線
   - 作業員の行はやめて、人員予定（昼間／夜間）に自社の人数を「1+4」の形で（外注は入れない）。作業の行は9つ以上
   - 記事の右（AA〜AJ）は自由に書く欄。下の補助（日付・「休」・曜日・始まりの日）は F〜Z の3週分（「休」を選ぶ／消すとその列の灰色が付く／消える）
   ライブラリは使わず、OOXML を文字で組んで無圧縮の zip にする（app.js の関数・値を使う） */

const XL_COLORS = { 黒: "FF000000", 青: "FF0070C0", 赤: "FFFF0000", 緑: "FF00A040", 黄: "FFFFC000", 橙: "FFFF7F00", 紫: "FF7030A0", 茶: "FF8D6E63", 灰: "FF8A96A3" };
const XL_DAYS = 21;                 // 列 F〜Z（月曜から3週間）
const XL_FIRST = 6;                 // F 列
const XL_RIGHT = 36;                // AJ 列（AA〜AJ は地図・ロゴ・担当・自由欄）
const XL_ITEMS = 9;                 // 作業の行の数（足りなければ空行）
const XL_MAP_W = 247.5;             // AA〜AJ の幅（pt。実際の Excel で測った）
const XL_MAP_H = 462;                // 地図の枠の高さ（作業9行。実際の Excel で測った）＝ app.js の MAP_RATIO の元
const xlCol = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
const xlRef = (r, c) => xlCol(c) + r;
const xlEsc = v => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
// Excel の日付の通し番号（1899-12-30 から）
const xlDate = k => Math.round((Date.UTC(+k.slice(0, 4), +k.slice(5, 7) - 1, +k.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000);

/* ---- 書式（styles.xml）。s の番号はこの並び ---- */
const XS = {
  def: 0, title: 1, made: 2, box: 3, boxL: 4, boxDate: 5, ym: 6, week: 7, month: 8, dnum: 9, dwk: 10, diag: 11,
  label: 12, gTop: 13, gMid: 14, gBar: 15, gBot: 16, num: 17, note: 18, noteDay: 19, help: 20, helpBox: 21, helpDate: 22, co: 23, boxV: 24,
  titleB: 25, vt: 26, plan: 27, noteEnd: 28, free: 29, monthMid: 30, monthEnd: 31, noteW: 32, noteEndW: 33
};
function xlStyles() {
  const fonts = [
    `<font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>`,                 // 0
    `<font><b/><sz val="20"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>`,             // 1 題
    `<font><sz val="9"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>`,                  // 2 小
    `<font><b/><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>`,             // 3 日付
    `<font><sz val="6"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>`,                  // 4 色の文字（見えなくなる）
    `<font><b/><sz val="18"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>`,             // 5 会社名
    `<font><sz val="8"/><color rgb="FF808080"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font>` // 6 下の補助
  ];
  const T = "thin", M = "medium";
  const bd = (l, r, t, b, extra = "") => `<border${extra}><left${l ? ` style="${l}"` : ""}>${l ? '<color auto="1"/>' : ""}</left><right${r ? ` style="${r}"` : ""}>${r ? '<color auto="1"/>' : ""}</right><top${t ? ` style="${t}"` : ""}>${t ? '<color auto="1"/>' : ""}</top><bottom${b ? ` style="${b}"` : ""}>${b ? '<color auto="1"/>' : ""}</bottom><diagonal${extra ? ` style="${T}"><color auto="1"/></diagonal>` : "/>"}</border>`;
  const borders = [
    `<border><left/><right/><top/><bottom/><diagonal/></border>`, // 0
    bd(T, T, T, T),          // 1 四方
    bd(T, T, T, ""),         // 2 上
    bd(T, T, "", ""),        // 3 左右だけ
    bd(T, T, "", T),         // 4 下
    bd(T, T, T, T, ' diagonalDown="1"'), // 5 斜め
    bd("", "", "", "hair"),  // 6 記事の点線
    bd("", "", "", M),       // 7 担当の下線
    bd("", "", "", T),       // 8 記事の最後の行（外周の下）
    bd("", "", T, ""),       // 9 月の行（上だけ。月が替わる所の左の線は条件付き書式）
    bd("", T, T, ""),        // 10 月の行の右端
    bd(T, "", T, ""),        // 11 月の行の左端
    bd(T, "", "", "hair"),   // 12 記事の点線＋週の境（月曜の左）の縦線
    bd(T, "", "", T)         // 13 記事の最後の行＋週の境
  ];
  const al = (h, v, wrap, rot) => `<alignment horizontal="${h}" vertical="${v}"${wrap ? ' wrapText="1"' : ""}${rot ? ` textRotation="${rot}"` : ""}/>`;
  const xf = (font, border, align, numFmt = 0) => `<xf numFmtId="${numFmt}" fontId="${font}" fillId="0" borderId="${border}" applyFont="1" applyBorder="1" applyAlignment="1"${numFmt ? ' applyNumberFormat="1"' : ""}>${align}</xf>`;
  const xfs = [
    xf(0, 0, al("general", "center")),            // def
    xf(1, 0, al("left", "center")),               // title
    xf(2, 1, al("center", "center"), 164),        // made 作成日
    xf(2, 1, al("center", "center", 1)),          // box
    xf(2, 1, al("left", "center", 1)),            // boxL
    xf(2, 1, al("center", "center"), 165),        // boxDate 工期
    xf(2, 1, al("center", "center"), 166),        // ym 令和○年○月
    xf(2, 1, al("center", "center")),             // week 今週
    xf(2, 11, al("left", "center")),              // month（左端）
    xf(3, 1, al("center", "center"), 167),        // dnum
    xf(3, 1, al("center", "center")),             // dwk
    xf(2, 5, al("center", "center")),             // diag
    xf(2, 1, al("left", "center", 1)),            // label 作業名
    xf(4, 2, al("center", "center")),             // gTop
    xf(4, 3, al("center", "center")),             // gMid
    xf(4, 3, al("center", "center")),             // gBar（色の文字）
    xf(4, 4, al("center", "center")),             // gBot
    xf(2, 3, al("left", "top"), 168),             // num ○名
    xf(2, 6, al("left", "center")),               // note
    xf(2, 6, al("left", "center")),               // noteDay
    xf(6, 0, al("center", "center")),             // help
    xf(2, 1, al("center", "center")),             // helpBox 休
    xf(6, 0, al("center", "center"), 167),        // helpDate
    xf(5, 1, al("right", "center")),              // co 会社名（ロゴが読めないときだけ）
    xf(2, 1, al("center", "center", 1)),          // boxV（住所。横書きで中央）
    xf(1, 1, al("left", "center")),               // titleB 題（外周の枠線）
    xf(2, 1, al("center", "center", 0, 255)),     // vt 縦書き（工事名・工期）
    xf(2, 1, al("center", "center")),             // plan 人員予定「1+4」
    xf(2, 8, al("left", "center")),               // noteEnd 記事の最後の行
    xf(2, 1, al("left", "top", 1)),               // free 記事の右の自由欄
    xf(2, 9, al("left", "center")),               // monthMid 月の行
    xf(2, 10, al("left", "center")),              // monthEnd 月の行の右端
    xf(2, 12, al("left", "center")),              // noteW 記事（週の境。2026-10-16 ユーザー：記事の欄に1週間ごとの縦線）
    xf(2, 13, al("left", "center"))               // noteEndW 記事の最後の行（週の境）
  ];
  // 条件付き書式の書式（dxf）：0..8 色、9 グラデーション、10 日曜の赤字、11 土曜の青字、12 月が替わる所の左の線
  const names = Object.keys(XL_COLORS);
  const dxfs = names.map(n => `<dxf><font><color rgb="${XL_COLORS[n]}"/></font><fill><patternFill><bgColor rgb="${XL_COLORS[n]}"/></patternFill></fill></dxf>`);
  dxfs.push(`<dxf><fill><gradientFill><stop position="0"><color rgb="FFBFBFBF"/></stop><stop position="0.5"><color rgb="FFFFFFFF"/></stop><stop position="1"><color rgb="FFBFBFBF"/></stop></gradientFill></fill></dxf>`);
  dxfs.push(`<dxf><font><b/><color rgb="FFFF0000"/></font></dxf>`);
  dxfs.push(`<dxf><font><b/><color rgb="FF0070C0"/></font></dxf>`);
  dxfs.push(`<dxf><border><left style="thin"><color auto="1"/></left></border></dxf>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="5"><numFmt numFmtId="164" formatCode="[$-ja-JP-x-gannen]ggge&quot;年&quot;mm&quot;月&quot;dd&quot;日&quot;(aaa)&quot;作成&quot;"/><numFmt numFmtId="165" formatCode="[$-ja-JP-x-gannen]ggge&quot;年&quot;mm&quot;月&quot;dd&quot;日&quot;"/><numFmt numFmtId="166" formatCode="[$-ja-JP-x-gannen]ggge&quot;年&quot;m&quot;月&quot;"/><numFmt numFmtId="167" formatCode="d"/><numFmt numFmtId="168" formatCode="0&quot;名&quot;;;"/></numFmts>
<fonts count="${fonts.length}">${fonts.join("")}</fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="${borders.length}">${borders.join("")}</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>
<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>
</styleSheet>`;
}

/* ---- 1現場・1週のシート ---- */
// pics＝{ logo, map }（入れる画像があるか）、mono＝白黒（線はどれも黒。2026-10-16 ユーザー：週間工程はカラー／白黒を選べる）
function xlSheet(s, mon, pics, mono) {
  const days = Array.from({ length: XL_DAYS }, (_, i) => dayAdd(mon, i)), end = days[XL_DAYS - 1];
  const lastCol = XL_FIRST + XL_DAYS - 1;                 // Z
  const R = XL_RIGHT;                                     // AJ
  const laneOf = b => s.段.find(l => l.id === b.段);
  const hasNight = s.段.some(isNight);
  // 作業ごとの行（同じ作業名・同じ昼夜はまとめる）。日ごとにその日のバーの色
  const items = [], byKey = {};
  for (const b of [...s.バー].sort((x, y) => x.開始.localeCompare(y.開始))) {
    if (b.終了 < mon || b.開始 > end) continue;
    const ln = laneOf(b);
    const name = (b.作業 || (ln ? laneWork(ln) : "") || "（作業）") + (ln && isNight(ln) ? "\n(夜間)" : hasNight ? "\n(昼間)" : "");
    let it = byKey[name];
    if (!it) { it = byKey[name] = { name, on: {} }; items.push(it); }
    days.forEach(k => { if (k >= b.開始 && k <= b.終了 && barWorks(ln, b, k)) it.on[k] = mono ? "黒" : (b.色 || "青"); });
  }
  const nItems = Math.max(XL_ITEMS, items.length);
  // 行の位置
  const R0 = 8, rPlan = R0 + nItems * 4, rNote = rPlan + 4, rLast = rNote + 4, rHelp = rLast + 1;   // 人員予定・記事・下の補助（印刷の外）
  const cells = {};       // "r,c" → {s, v, t, f}
  const put = (r, c, o) => { cells[r + "," + c] = Object.assign(cells[r + "," + c] || {}, o); };
  const merges = [];
  const box = (r1, c1, r2, c2, st, val) => {      // 結合して、範囲の全部のセルに罫線の書式
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) put(r, c, { s: st });
    if (val !== undefined) put(r1, c1, val);
    if (r1 !== r2 || c1 !== c2) merges.push(xlRef(r1, c1) + ":" + xlRef(r2, c2));
  };
  const str = v => ({ t: "s", v });
  const num = v => ({ t: "n", v });
  // 1〜3 行：題・作成日・ロゴ（画像）／工事名（工事略称・現場コード）・工期・住所・担当
  box(1, 1, 1, 19, XS.titleB, str("週　　間　　工　　程　　表"));
  box(1, 20, 1, lastCol, XS.made, num(xlDate(todayKey())));
  box(1, lastCol + 1, 1, R, XS.co, pics.logo ? undefined : str(COMPANY));
  box(2, 1, 3, 1, XS.vt, str("工事名"));
  box(2, 2, 2, 4, XS.box, str("工事略称"));
  box(2, 5, 2, 12, XS.boxL, str(nm(s)));
  box(3, 2, 3, 4, XS.box, str("現場コード"));
  box(3, 5, 3, 12, XS.boxL, str(s.コード || ""));
  box(2, 13, 3, 13, XS.vt, str("工期"));
  box(2, 14, 2, 14, XS.box, str("自")); box(3, 14, 3, 14, XS.box, str("至"));
  box(2, 15, 2, 18, XS.boxDate, s.工期自 ? num(xlDate(s.工期自)) : {});
  box(3, 15, 3, 18, XS.boxDate, s.工期至 ? num(xlDate(s.工期至)) : {});
  box(2, 19, 3, 19, XS.boxV, str("住所"));
  box(2, 20, 3, lastCol, XS.boxL, str(s.住所 || ""));
  box(2, lastCol + 1, 2, lastCol + 2, XS.box, str("担当:"));
  box(2, lastCol + 3, 2, R, XS.boxL, str(s.担当 || ""));
  box(3, lastCol + 1, 3, R, XS.box);
  // 4〜7 行：年月・今週／来週／再来週・月・日・曜。右の AA〜AJ は地図の枠（4行目〜人員予定の下。中は罫線なし）
  box(4, 1, 5, 5, XS.ym, { f: xlRef(6, XL_FIRST) });
  box(4, 6, 4, 12, XS.week, str("今週")); box(4, 13, 4, 19, XS.week, str("来週")); box(4, 20, 4, lastCol, XS.week, str("再来週"));
  box(4, lastCol + 1, rPlan + 3, R, XS.box);
  box(6, 1, 7, 5, XS.diag);
  for (let i = 0; i < XL_DAYS; i++) {
    const c = XL_FIRST + i, L = xlCol(c), P = xlCol(c - 1);
    // 月：最初の日と、月が替わる日だけ「12月」（その左の線は条件付き書式。ユーザー：│1月 のように）
    put(5, c, { s: i === 0 ? XS.month : c === lastCol ? XS.monthEnd : XS.monthMid, f: i === 0 ? `MONTH(${L}6)&"月"` : `IF(DAY(${L}6)=1,MONTH(${L}6)&"月","")` });
    // 日付：F6 は下の補助の A 列の始まりの日（ここを変えれば全部ずれる）
    put(6, c, { s: XS.dnum, f: i === 0 ? `$A$${rHelp + 3}` : `${P}6+1` });
    put(7, c, { s: XS.dwk, f: `TEXT(${L}6,"aaa")` });
  }
  // 作業の行（4行で1つ）
  for (let n = 0; n < nItems; n++) {
    const r = R0 + n * 4, it = items[n];
    box(r, 1, r + 3, 5, XS.label, it ? str(it.name) : {});
    days.forEach((k, i) => {
      const c = XL_FIRST + i;
      put(r, c, { s: XS.gTop }); put(r + 1, c, { s: XS.gBar }); put(r + 2, c, { s: XS.gMid }); put(r + 3, c, { s: XS.gBot });
      if (it && it.on[k]) put(r + 1, c, str(it.on[k]));
    });
  }
  // 人員予定（昼間／夜間）：その日に作業する自社の工種のバーの人数を「1+4」の形で足す（外注は入れない）
  const planAt = (k, night) => {
    const ps = [];
    for (const b of s.バー) {
      const ln = laneOf(b);
      if (isGai(ln) || !!(ln && isNight(ln)) !== night || k < b.開始 || k > b.終了 || !barWorks(ln, b, k)) continue;
      const p = parseHead(headTxtAt(b, k, ln));
      if (p && p.total) ps.push(p.parts);
    }
    if (!ps.length) return "";
    const w = Math.max(...ps.map(p => p.length)), sum = Array(w).fill(0);
    ps.forEach(p => p.forEach((x, j) => { sum[w - p.length + j] += x; }));   // 「4」だけのものは右（作業員）に足す
    return sum.join("+");
  };
  box(rPlan, 1, rPlan + 3, 4, XS.label, str("人員予定"));
  box(rPlan, 5, rPlan + 1, 5, XS.week, str("昼間")); box(rPlan + 2, 5, rPlan + 3, 5, XS.week, str("夜間"));
  days.forEach((k, i) => {
    const c = XL_FIRST + i, d = planAt(k, false), n = planAt(k, true);
    box(rPlan, c, rPlan + 1, c, XS.plan, d ? str(d) : {});
    box(rPlan + 2, c, rPlan + 3, c, XS.plan, n ? str(n) : {});
  });
  // 記事（週ごとに、その週に始まるバーのメモ。1週目はそれより前から続くものも）。日付は F/M/T 列、文は次の列から。右の AA〜AJ は自由に書く欄
  box(rNote, 1, rLast, 5, XS.label, str("記　　事"));
  for (let w = 0; w < 3; w++) {
    const a = days[w * 7], z = days[w * 7 + 6], c0 = XL_FIRST + w * 7;
    const list = s.バー.filter(b => b.メモ && (w === 0 ? b.開始 <= z && b.終了 >= a : b.開始 >= a && b.開始 <= z)).sort((x, y) => x.開始.localeCompare(y.開始)).slice(0, 4);
    for (let j = 0; j < 4; j++) {
      const st = j === 3 ? XS.noteEnd : XS.note, stW = j === 3 ? XS.noteEndW : XS.noteW;
      for (let c = c0; c < c0 + 7; c++) put(rNote + 1 + j, c, { s: w && c === c0 ? stW : st });   // 2・3週目の月曜の左に縦線
      const b = list[j];
      if (b) { put(rNote + 1 + j, c0, str(+(b.開始 < a ? a : b.開始).slice(8) + "日")); put(rNote + 1 + j, c0 + 1, str(`${b.作業 || ""}　${b.メモ}`)); }
    }
  }
  for (let c = XL_FIRST; c <= lastCol; c++) put(rNote, c, { s: (c - XL_FIRST) % 7 === 0 && c > XL_FIRST ? XS.noteW : XS.note });
  box(rNote, lastCol + 1, rLast, R, XS.free);
  // 下の補助（印刷の外）：日付・「休」・曜日（F〜Z）と始まりの日（A 列）。「休」を選ぶ／消すと、その列の灰色のぼかし（6行目〜人員予定）が付く／消える
  // （2026-10-16 ユーザー：3週間に直したときに無くなった休日の ON/OFF を戻す。前に消したのは4週目［AA〜AJ］の分だけでよかった）
  const allLanes = s.段;
  put(rHelp + 3, 1, { s: XS.helpDate, ...num(xlDate(mon)) });
  put(rHelp + 1, 1, { s: XS.help, ...str("休日→") });
  put(rHelp + 3, 2, { s: XS.help, ...str("←始まりの日（月曜）") });
  days.forEach((k, i) => {
    const c = XL_FIRST + i, L = xlCol(c);
    put(rHelp, c, { s: XS.helpDate, f: `${L}6` });
    // はじめから「休」：会社の休み・日曜・祝日、工種がどれもその日に作業しないとき（土曜休みの現場の土曜など）
    const off = isRest(k) || (allLanes.length ? allLanes.every(ln => !laneWorks(ln, k)) : (dayDow(k) === 0 || holidayOf(k)));
    put(rHelp + 1, c, { s: XS.helpBox, ...(off ? str("休") : {}) });
    put(rHelp + 2, c, { s: XS.help, f: `TEXT(${L}${rHelp},"aaa")` });
  });
  // シートの XML
  const rowsH = { 1: 30, 2: 24, 3: 24, 4: 18.75, 5: 18.75, 6: 22.5, 7: 22.5 };
  for (let n = 0; n < nItems; n++) { const r = R0 + n * 4; rowsH[r] = 12; rowsH[r + 1] = 6.75; rowsH[r + 2] = 6.75; rowsH[r + 3] = 12; }
  for (let j = 0; j < 4; j++) rowsH[rPlan + j] = 11;     // 人員予定は昼間・夜間の数字が読めるよう4行とも同じ高さ
  rowsH[rNote] = 6; for (let j = 1; j <= 4; j++) rowsH[rNote + j] = 15;
  rowsH[rHelp] = 12; rowsH[rHelp + 1] = 20.25; rowsH[rHelp + 2] = 12; rowsH[rHelp + 3] = 12;
  const maxRow = rHelp + 3;
  const rowXml = [];
  for (let r = 1; r <= maxRow; r++) {
    const cs = Object.keys(cells).filter(k => +k.split(",")[0] === r).map(k => +k.split(",")[1]).sort((a, b) => a - b);
    const body = cs.map(c => {
      const o = cells[r + "," + c], ref = xlRef(r, c), st = ` s="${o.s || 0}"`;
      if (o.f) return `<c r="${ref}"${st}><f>${xlEsc(o.f)}</f></c>`;
      if (o.t === "n") return `<c r="${ref}"${st}><v>${o.v}</v></c>`;
      if (o.t === "s" && o.v !== "") return `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${xlEsc(o.v)}</t></is></c>`;
      return `<c r="${ref}"${st}/>`;
    }).join("");
    rowXml.push(`<row r="${r}"${rowsH[r] ? ` ht="${rowsH[r]}" customHeight="1"` : ""}>${body}</row>`);
  }
  const Z = xlCol(lastCol);
  const barRows = [];
  for (let n = 0; n < nItems; n++) barRows.push(`F${R0 + n * 4 + 1}:${Z}${R0 + n * 4 + 1}`);
  // 「休」の列：日付・曜日（6・7行。ユーザー指定）から人員予定まで灰色のぼかし。日付は日曜か「休」なら赤字（土曜は青のまま）
  const H = rHelp + 1;
  const names = Object.keys(XL_COLORS);
  let pr = 1;
  const cf = [
    `<conditionalFormatting sqref="${barRows.join(" ")}">` + names.map((nm, i) => `<cfRule type="expression" dxfId="${i}" priority="${pr++}"><formula>F${R0 + 1}="${nm}"</formula></cfRule>`).join("") + `</conditionalFormatting>`,
    `<conditionalFormatting sqref="F6:${Z}${rPlan + 3}"><cfRule type="expression" dxfId="9" priority="${pr++}"><formula>F$${H}="休"</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="F6:${Z}7"><cfRule type="expression" dxfId="10" priority="${pr++}"><formula>OR(WEEKDAY(F$6)=1,F$${H}="休")*(WEEKDAY(F$6)&lt;&gt;7)</formula></cfRule><cfRule type="expression" dxfId="11" priority="${pr++}"><formula>WEEKDAY(F$6)=7</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="G5:${Z}5"><cfRule type="expression" dxfId="12" priority="${pr++}"><formula>DAY(G$6)=1</formula></cfRule></conditionalFormatting>`
  ];
  const dv = `<dataValidations count="2"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${barRows.join(" ")}"><formula1>"${names.join(",")}"</formula1></dataValidation>` +
    `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="F${H}:${Z}${H}"><formula1>"休"</formula1></dataValidation></dataValidations>`;
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<dimension ref="A1:${xlRef(maxRow, R)}"/>
<sheetViews><sheetView showGridLines="0" zoomScale="100" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultColWidth="9" defaultRowHeight="13.5"/>
<cols><col min="1" max="5" width="3.625" customWidth="1"/><col min="6" max="${R}" width="4.125" customWidth="1"/></cols>
<sheetData>${rowXml.join("")}</sheetData>
<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>
${cf.join("")}
${dv}
<printOptions horizontalCentered="1"/>
<pageMargins left="0.2" right="0.2" top="0.35" bottom="0.35" header="0.3" footer="0.3"/>
<pageSetup paperSize="${($("inPaper").value || "A3") === "A3" ? 8 : 9}" orientation="landscape" fitToWidth="1" fitToHeight="1"/>
${pics.logo || pics.map ? `<drawing r:id="rId1"/>` : ""}
</worksheet>`;
  return { xml, printArea: `$A$1:$${xlCol(R)}$${rLast}`, rowsH, rEnd: rPlan + 3 };
}

/* ---- 画像：ロゴ（AA1:AJ1 の右寄せ）と地図（AA4 の左上から地図の枠の大きさ）。セルに合わせて伸び縮みしない ---- */
const XL_EMU = 12700;   // 1pt
const XL_COLW = XL_MAP_W / 10;   // AA〜AJ の1列の幅（pt）
function xlPic(id, name, rid, x, row, rowOff, w, h, line) {
  // x は AA 列の左からの pt。列をまたぐ分は列の番号に繰り上げる
  const ci = Math.min(9, Math.floor(x / XL_COLW)), cx = Math.round((x - ci * XL_COLW) * XL_EMU);
  const W = Math.round(w * XL_EMU), H = Math.round(h * XL_EMU);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${26 + ci}</xdr:col><xdr:colOff>${cx}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${Math.round(rowOff * XL_EMU)}</xdr:rowOff></xdr:from><xdr:ext cx="${W}" cy="${H}"/>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${line ? `<a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>` : `<a:noFill/>`}</xdr:spPr></xdr:pic>
<xdr:clientData/></xdr:oneCellAnchor>`;
}
// 地図はセルに付けて置く（始まりと終わりをセルで決める twoCellAnchor）。
// 大きさで置くと、印刷・PDF では行が少し低く出るので地図が人員予定の下へはみ出した（2026-10-16 試し）
function xlPicCells(id, name, rid, x0, y0, x1, y1, rowsH) {
  const col = x => { const ci = Math.min(9, Math.floor(x / XL_COLW)); return [26 + ci, Math.round((x - ci * XL_COLW) * XL_EMU)]; };
  const row = y => { let r = 4; while (y >= (rowsH[r] || 13.5) && rowsH[r + 1]) { y -= rowsH[r] || 13.5; r++; } return [r - 1, Math.round(y * XL_EMU)]; };   // 4行目から
  const [c0, co0] = col(x0), [c1, co1] = col(x1), [r0, ro0] = row(y0), [r1, ro1] = row(y1);
  const W = Math.round((x1 - x0) * XL_EMU), H = Math.round((y1 - y0) * XL_EMU);
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${c0}</xdr:col><xdr:colOff>${co0}</xdr:colOff><xdr:row>${r0}</xdr:row><xdr:rowOff>${ro0}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c1}</xdr:col><xdr:colOff>${co1}</xdr:colOff><xdr:row>${r1}</xdr:row><xdr:rowOff>${ro1}</xdr:rowOff></xdr:to>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr></xdr:pic>
<xdr:clientData/></xdr:twoCellAnchor>`;
}
// logo＝{ w, h }（画像の大きさ）、m＝地図 { w, h }、rowsH＝行の高さ（地図の枠＝4行目〜rEnd 行）
function xlDrawing(logo, m, rowsH, rEnd) {
  const parts = [];
  if (logo) {   // 高さ 20pt（1行目は 30pt）、右に 16pt あける（印刷・PDF では画像が少し大きく出て、6pt だと右の枠で切れた）
    const h = 20, w = Math.min(XL_MAP_W - 12, h * logo.w / logo.h);
    parts.push(xlPic(2, "ロゴ", "rId1", XL_MAP_W - 16 - w, 0, (30 - w * logo.h / logo.w) / 2, w, w * logo.h / logo.w, false));
  }
  if (m) {      // 枠の中に 2pt あけて置く。枠と縦横の比が違う地図（前に保存したもの）は、枠の中に収まるよう縮める（左上に寄せる）
    let FH = 0; for (let r = 4; r <= rEnd; r++) FH += rowsH[r] || 13.5;
    const W = XL_MAP_W - 4, H = FH - 4, k = m.w && m.h ? Math.min(W / m.w, H / m.h) : 0;
    parts.push(xlPicCells(3, "地図", "rId2", 2, 2, 2 + (k ? m.w * k : W), 2 + (k ? m.h * k : H), rowsH));
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${parts.join("\n")}</xdr:wsDr>`;
}
const dataUrlBytes = u => { const b = atob(u.slice(u.indexOf(",") + 1)), a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; };

/* ---- 無圧縮の zip ---- */
const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
function zipStore(files) {
  const enc = new TextEncoder(), parts = [], central = [];
  let off = 0;
  for (const [name, text] of files) {
    const nb = enc.encode(name), db = typeof text === "string" ? enc.encode(text) : text, crc = crc32(db);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
    h.setUint32(14, crc, true); h.setUint32(18, db.length, true); h.setUint32(22, db.length, true); h.setUint16(26, nb.length, true);
    parts.push(new Uint8Array(h.buffer), nb, db);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
    c.setUint32(16, crc, true); c.setUint32(20, db.length, true); c.setUint32(24, db.length, true); c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
    central.push(new Uint8Array(c.buffer), nb);
    off += 30 + nb.length + db.length;
  }
  const csize = central.reduce((a, b) => a + b.length, 0);
  const e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, csize, true); e.setUint32(16, off, true);
  return new Blob([...parts, ...central, new Uint8Array(e.buffer)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

/* ---- ブック ---- */
// maps[i]＝その現場の地図 { url:data URL, w, h }（画像の大きさ）か null。logo＝{ bytes, w, h }（logo.png）か null。mono＝白黒
function weekXlsx(list, mon, maps = [], logo = null, mono = false) {
  const sheets = list.map((s, i) => ({ s, map: maps[i] || null, ...xlSheet(s, mon, { logo: !!logo, map: !!maps[i] }, mono) }));
  const used = new Set();
  const sheetName = s => {
    let n = ((s.略称 || s.名前) + " " + mon.slice(5, 7) + "月" + mon.slice(8, 10) + "日～").replace(/[\[\]:*?\/\\]/g, "").slice(0, 31);
    let k = n, i = 2; while (used.has(k)) k = n.slice(0, 28) + "(" + i++ + ")"; used.add(k); return k;
  };
  const names = sheets.map(x => sheetName(x.s));
  const hasPic = x => !!(logo || x.map);
  const files = [
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}${sheets.map((x, i) => hasPic(x) ? `<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` : "").join("")}</Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${names.map((n, i) => `<sheet name="${xlEsc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets><definedNames>${sheets.map((x, i) => `<definedName name="_xlnm.Print_Area" localSheetId="${i}">'${xlEsc(names[i].replace(/'/g, "''"))}'!${x.printArea}</definedName>`).join("")}</definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml", xlStyles()],
    ...sheets.map((x, i) => [`xl/worksheets/sheet${i + 1}.xml`, x.xml])
  ];
  if (logo) files.push(["xl/media/logo.png", logo.bytes]);   // ロゴは全シートで1つを使う
  const IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  sheets.forEach((x, i) => {
    if (!hasPic(x)) return;
    const n = i + 1, rel = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    files.push([`xl/worksheets/_rels/sheet${n}.xml.rels`, rel + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${n}.xml"/></Relationships>`]);
    files.push([`xl/drawings/drawing${n}.xml`, xlDrawing(logo, x.map, x.rowsH, x.rEnd)]);
    files.push([`xl/drawings/_rels/drawing${n}.xml.rels`, rel + (logo ? `<Relationship Id="rId1" Type="${IMG}" Target="../media/logo.png"/>` : "") + (x.map ? `<Relationship Id="rId2" Type="${IMG}" Target="../media/image${n}.jpeg"/>` : "") + `</Relationships>`]);
    if (x.map) files.push([`xl/media/image${n}.jpeg`, dataUrlBytes(x.map.url)]);
  });
  return zipStore(files);
}
// ロゴ（web\logo.png、背景透過）。大きさは PNG の頭（IHDR）から。読めなければ会社名の文字
async function loadLogo() {
  try {
    const r = await fetch("logo.png", { cache: "no-cache" });
    if (!r.ok) return null;
    const bytes = new Uint8Array(await r.arrayBuffer()), dv = new DataView(bytes.buffer);
    return { bytes, w: dv.getUint32(16), h: dv.getUint32(20) };
  } catch (e) { return null; }
}
function downloadBlob(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
// 印刷の小窓の「Excel で出す」
$("bPrintXlsx").onclick = async () => {
  const ids = [...$("pSites").querySelectorAll("input:checked")].map(c => c.value);
  if (!ids.length || !$("pFrom").value) return;
  const mon = monday($("pFrom").value), list = shownSitesOf(ids);
  const name = "週間工程_" + (list.length === 1 ? (list[0].略称 || list[0].名前) + "_" : "") + mon.replace(/-/g, "") + ".xlsx";
  // 地図は現場の設定で入れたもの（読んでから組む）
  if (list.some(s => s.地図)) toast("地図を読んでいます…", 1500);
  const maps = await Promise.all(list.map(async s => {
    try { const u = await loadMap(s); if (!u || !u.startsWith("data:")) return null; const im = await loadImg(u); return { url: u, w: im ? im.naturalWidth : 0, h: im ? im.naturalHeight : 0 }; }
    catch (e) { return null; }
  }));
  downloadBlob(weekXlsx(list, mon, maps, await loadLogo(), weekMono()), name.replace(/[\\\/:*?"<>|]/g, ""));
  toast("Excel を作りました：" + name);
};
