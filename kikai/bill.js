/* ---------- 請求書（2026-11-13） ----------
   ユーザー：光波距離計・オートレベルを、月に1回、現場別に分けた請求書にして印刷したい（いまはユーザー1人の作業）。
   形は 測量機稼働表.xlsx（令和8年8月分 など）：日付×機械（昼・夜）の表＋現場別の 数量・金額。請求書に載せる現場は、その月に使った現場だけ。
   将来ほかの機械も同じ形にするかもしれない（表の形は変わるかも）→ 機械は「まとまり」に入れるだけで請求書になるようにした。
   - まとまり：{ id, 名前, 題名, 機械:[機械 id] }。D.請求書.まとまり に持つ（全員で同じ。データの形は足すだけ）。無いあいだは機械の名前（レベル・光波）から作る。
   - 印刷の分け方（S.billMode）：group＝機械別の請求（まとまりごとに1枚。Excel と同じ）／site＝現場別の請求（現場ごとに1枚）。
   - 数量＝使った日（昼と夜はそれぞれ1回。同じ日に昼夜使えば2）×台数。使わなかった（申告）の日は数えない（works）。
   - 金額は集計と同じ costOf（単価表・特別単価・料金の決め方）。積置き（請求する）はその現場（billTo）。置場・整備は請求しない。
   - 消費税（S.billTax）：載せない（2026-11-13 ユーザー）。仕組みは残す（現場ごとに 税抜×税率 の1円未満を切り捨て）。
   - 2026-11-13 ユーザー：請求の単価は実行予算（単価表）と別。レベル 200円・光波距離計 1,500円（まとまりの 単価＝{ 機械 id: 円／回 }）。
     請求単価がある機械は 請求単価×数量（交渉で決めた特別単価のある予約はその単価）。無い機械は集計と同じ costOf。
   - 光波距離計は湘南通商から長期リース（月 30,000円）。まとまりの 月額 があれば、集計の一番上に「本社」＝月額−現場の計 の行（金額だけ）。合計＝月額。
   - 印刷は A4 縦（印刷の前に @page を差し替える。稼働表の印刷は A3 横のまま）。 */
const BILL_TAX = 10;
const reiwa = ym => { const y = +ym.slice(0, 4), m = +ym.slice(5, 7); return (y >= 2019 ? "令和" + (y - 2018) : "平成" + (y - 1988)) + "年" + m + "月分"; };
const monthLast = ym => { const d = dateOf(ym + "-01"); d.setMonth(d.getMonth() + 1); d.setDate(0); return keyOf(d); };
const monthMove = (ym, n) => addMonths(ym + "-01", n).slice(0, 7);
if (!/^\d{4}-\d\d$/.test(S.billMonth || "")) S.billMonth = monthMove(todayKey().slice(0, 7), -1);   // はじめは先月（月はじめに先月分を作る）
if (!["group", "site"].includes(S.billMode)) S.billMode = "group";
S.billTax = false;   // 2026-11-13 ユーザー：消費税は入れなくていい（税の行・列は出さない。戻すときはここを true と下のチェックを戻す）

/* 請求単価・月額のはじめの値（2026-11-13 ユーザー：レベル 200円・光波距離計 1,500円、光波は月 30,000円のリース） */
function billDefaults(g) {
  if (!g.単価) { g.単価 = {}; g.機械.forEach(id => { const n = machine(id)?.名前 || ""; if (/レベル/.test(n)) g.単価[id] = 200; else if (/光波/.test(n)) g.単価[id] = 1500; }); }
  if (g.月額 === undefined) g.月額 = g.id === "g_koha" ? 30000 : null;
  return g;
}
/* まとまり：保存したもの、無ければ機械の名前から（レベル／光波） */
function billGroups() {
  const g = D.請求書?.まとまり;
  if (Array.isArray(g)) return g.map(x => billDefaults({ ...x, 機械: (x.機械 || []).filter(id => machine(id)) }));
  const live = D.機械.filter(m => !m.使わない);
  const pick = re => live.filter(m => re.test(m.名前)).map(m => m.id);
  return [
    { id: "g_level", 名前: "レベル（自社機械）", 題名: "測量機稼働表", 機械: pick(/レベル/) },
    { id: "g_koha", 名前: "光波距離計（湘南通商）", 題名: "測量機稼働表", 機械: pick(/光波/) }
  ].filter(x => x.機械.length).map(billDefaults);
}
const saveGroups = fn => commit(d => {
  const L = JSON.parse(JSON.stringify(billGroups()));
  fn(L);
  d.請求書 = Object.assign({}, d.請求書, { まとまり: L });
}, { noUndo: true });

/* ---------- 確認印（2026-11-13） ----------
   ユーザー：いまだに捺印文化があるので、計の横に捺印欄。将来は各担当者が Web 上で確認の印を付けたい。
   ログインしている人の名前の電子印鑑を作れるならそのほうが良い。
   - 集計の各行（本社・現場）の右に「確認印」の欄。紙では空欄＝そのまま判子を押せる。
   - 画面で欄を押す → 「○○ の印で確認しますか」→ ログインしている人（me.name）の名前の丸い印（赤）。もう一度押すと取り消し（押した本人だけ）。
   - D.請求書.確認["YYYY-MM|まとまり id|現場"] = { 名前, 日時, 金額 }。押したときの金額を覚え、金額が変わったら印を薄くして「確かめてください」に出す。 */
const stampKey = (ym, gid, site) => ym + "|" + gid + "|" + site;
const stampOf = key => D.請求書?.確認?.[key] || null;
const sealName = n => { const t = String(n || "").trim().split(/[\s　]+/)[0] || "?"; return t.length > 4 ? t.slice(0, 4) : t; };   // 姓だけ（「北澤 和也」→ 北澤）
function stampCell(key, amt, warn, label) {
  const x = stampOf(key);
  if (!x) return `<td class="bstamp" data-stamp="${esc(key)}" data-amt="${amt}" data-label="${esc(label)}"><span class="bhint">押す</span></td>`;
  const changed = x.金額 !== amt;
  if (changed) warn.push(`${label}：確認印のあと金額が変わりました（印のとき ${yen(x.金額)}円 → いま ${yen(amt)}円）。押し直してください`);
  const nm = sealName(x.名前), d = x.日時 ? x.日時.slice(5, 10).replace("-", "/") : "";
  return `<td class="bstamp on${changed ? " old" : ""}" data-stamp="${esc(key)}" data-amt="${amt}" data-label="${esc(label)}" title="${esc(x.名前 + "　" + (x.日時 || ""))}">
    <span class="seal"><i class="sd">${esc(d)}</i><b class="${nm.length >= 3 ? "l3" : ""}">${esc(nm)}</b><i class="sk">確認</i></span></td>`;
}
function stampClick(td) {
  const key = td.dataset.stamp, amt = +td.dataset.amt, x = stampOf(key), who = me.name;
  if (x && x.名前 !== who && !confirm(`${x.名前} さんの印です。取り消して ${who} の印にしますか？`)) return;
  if (x && x.名前 === who) { if (!confirm(`${td.dataset.label}\n${who} の確認印を取り消しますか？`)) return;
    commit(d => { if (d.請求書?.確認) delete d.請求書.確認[key]; }, { noUndo: true }); return; }
  if (!x && !confirm(`${td.dataset.label}（${yen(amt)}円）\n${who} の印で確認しますか？`)) return;
  const v = { 名前: who, 日時: new Date().toLocaleString("sv-SE").slice(0, 16), 金額: amt };
  commit(d => { d.請求書 = Object.assign({}, d.請求書); d.請求書.確認 = Object.assign({}, d.請求書.確認, { [key]: v }); }, { noUndo: true });
}

/* ---------- 区分（2026-11-13 夜 ユーザー） ----------
   自社のラフタークレーンは昼夜で単価が変わる。表に 昼・昼残業・夜・夜残業 の列。残業した日は表のその日に数量を入れる。
   月末の下に それぞれの列の 合計数量・単価・金額。残業代は 昼 4,000円・夜 5,000円。
   - まとまりの 区分単価[機械 id] = { 昼, 夜, 昼残業, 夜残業 }（「昼・夜・残業で分ける」）。無い機械は 単価[機械 id]（1つの請求単価）。
   - 残業の数量：D.請求書.残業日["YYYY-MM-DD|機械 id|昼／夜"]。その日・その区分に予約のある現場の分として数える。
   - 交渉の特別単価（機材センター長 田村氏が最終確認。予約の 交渉＝決まった）がある予約は、請求書でもその単価（「特」）。 */
const KB2 = ["昼", "夜"], KB4 = ["昼", "昼残業", "夜", "夜残業"];
const isOT = kk => kk.endsWith("残業");
const splitOf = (g, id) => g.区分単価?.[id] || null;
const otOf = (k, mid, kb) => +(D.請求書?.残業日?.[k + "|" + mid + "|" + kb] || 0);
const colsOf = x => x.split ? KB4 : KB2;

/* 1台の機械のその月：日ごとの行き先（表）と、請求する現場ごと・区分ごとの 数量・金額 */
function billMachine(g, m, from, to) {
  const split = splitOf(g, m.id), flat = split ? null : g.単価?.[m.id] ?? null;
  const rs = D.予約.filter(r => r.機械 === m.id && r.開始 <= to && r.終了 >= from);
  const cell = {};   // "日|区分" → [{ name, bill, kari, sp }]
  rs.forEach(r => {
    for (let k = r.開始 < from ? from : r.開始; k <= to && k <= r.終了; k = dayAdd(k, 1)) {
      if (!works(r, k)) continue;
      const name = isTsumi(r.現場) && r.積置き?.現場 ? r.現場 + "（" + r.積置き.現場 + "）" : r.現場;
      (cell[k + "|" + r.区分] ||= []).push({ name, bill: isBill(r), kari: r.状態 === "仮", sp: special(r) });
    }
  });
  const sites = {}, warn = [], col = {}, unset = new Set();
  KB4.forEach(kk => { col[kk] = { q: 0, a: 0, units: new Set() }; });
  const S0 = () => ({ k: Object.fromEntries(KB4.map(kk => [kk, { q: 0, a: 0, units: new Set(), days: [] }])), qty: 0, amt: 0, units: new Set(), sp: false, none: false, days: [] });
  const add = (site, kk, q, a, unit, k) => {
    const x = (sites[site] ||= S0()), y = x.k[kk];
    y.q += q; y.a += a || 0; x.qty += q; x.amt += a || 0; col[kk].q += q; col[kk].a += a || 0;
    if (unit != null) { y.units.add(unit); x.units.add(unit); col[kk].units.add(unit); }
    if (k) { y.days.push(k); x.days.push(k); }
    return x;
  };
  rs.filter(isBill).forEach(r => {
    const site = billTo(r), n = r.台数 || 1;
    if (split || flat != null) {   // 請求単価：使った日ごと（特別単価があればその単価）
      for (let k = r.開始 < from ? from : r.開始; k <= to && k <= r.終了; k = dayAdd(k, 1)) {
        if (!works(r, k)) continue;
        const u = special(r) ? r.交渉.単価 : split ? split[r.区分] ?? null : flat;
        if (u == null) unset.add(r.区分);
        const x = add(site, r.区分, n, u == null ? 0 : u * n, u, k);
        if (special(r)) x.sp = true;
      }
    } else {   // 請求単価が無い機械：集計と同じ（単価表・特別単価・料金の決め方）
      const c = costOf(r, from, to, false, m);
      if (!c.days && !c.amt) return;
      const x = add(site, r.区分, c.meth === "日" ? c.days * c.n : c.qty * c.n, c.amt, c.unit);
      for (let k = r.開始 < from ? from : r.開始; k <= to && k <= r.終了; k = dayAdd(k, 1)) if (works(r, k)) { x.days.push(k); x.k[r.区分].days.push(k); }
      x.meth = c.meth;
      if (c.amt == null) { x.none = true; warn.push(`${mLabel(m)}「${site}」は単価がありません（金額に入っていません）`); }
      if (special(r)) x.sp = true;
    }
    if (r.状態 === "仮") warn.push(`${mLabel(m)} ${r.区分}「${r.現場}」${md(r.開始)}〜${md(r.終了)} が「仮」のまま`);
  });
  /* 残業（区分単価の機械だけ）：その日・その区分の現場の分 */
  const ot = {};
  if (split) Object.entries(D.請求書?.残業日 || {}).forEach(([key, q]) => {
    const [k, mid, kb] = key.split("|"); if (mid !== m.id || k < from || k > to || !q) return;
    ot[k + "|" + kb] = +q;
    const r = rs.find(r => r.区分 === kb && works(r, k) && isBill(r)), kk = kb + "残業", u = split[kk] ?? null;
    if (!r) { warn.push(`${mLabel(m)} ${md(k)} ${kb}：現場の予約が無いので、残業 ${q} を数えていません`); return; }
    if (u == null) unset.add(kk);
    add(billTo(r), kk, +q, u == null ? 0 : u * q, u, k);
  });
  unset.forEach(kk => warn.push(`${mLabel(m)}：${kk}の請求単価がまだ入っていません（0円で数えています）`));
  Object.entries(cell).forEach(([key, a]) => { if (a.length > 1) warn.push(`${mLabel(m)} ${md(key.split("|")[0])} ${key.split("|")[1]} に行き先が2つ（${a.map(x => x.name).join("・")}）`); });
  return { m, split, flat, cell, ot, sites, col, warn };
}
const unitsText = set => set.size === 0 ? "" : set.size === 1 ? yen([...set][0]) : "混在";
const unitText = s => s.meth && s.meth !== "日" ? [...s.units].map(u => yen(u) + "／" + UNIT[s.meth]).join("・") : [...s.units].map(yen).join("・");
const taxOf = v => S.billTax ? Math.floor(v * BILL_TAX / 100) : 0;
const siteCode = name => siteOf(name).コード || "";   // 2026-11-13 ユーザー：正式工事名・元請は要らない。社内の現場名と現場コードだけ
const siteHead = name => siteCode(name) ? "現場コード " + siteCode(name) : "";
const sortSites = names => names.sort((a, b) => (siteOf(a).コード || "999999").localeCompare(siteOf(b).コード || "999999") || a.localeCompare(b, "ja"));

/* 紙の幅（2026-11-13 ユーザー：余白 左 25mm・右 10mm）＝ 210−35＝175mm。日付の列 16mm、残りを列で分ける */
const PAPER_W = 175, DATE_W = 16, OT_W = 13;

/* まとまりごとの1枚（機械別の請求）：日付×機械（昼・夜、区分単価の機械は 昼・昼残業・夜・夜残業）の表＋月末の下に列ごとの 合計数量・単価・金額＋使った現場だけの集計 */
function billGroupPage(g, from, to, made, hidden) {
  const ms = g.機械.map(machine).filter(Boolean), bm = ms.map(m => billMachine(g, m, from, to)), one = ms.length === 1 && !bm[0].split;
  const ym = from.slice(0, 7);
  const days = []; for (let k = from; k <= to; k = dayAdd(k, 1)) days.push(k);
  const names = sortSites([...new Set(bm.flatMap(x => Object.keys(x.sites)))]);
  const nBase = bm.reduce((t, x) => t + 2, 0), nOT = bm.reduce((t, x) => t + (x.split ? 2 : 0), 0);
  const baseW = one ? 33 : Math.min(40, Math.floor((PAPER_W - DATE_W - nOT * OT_W) / nBase));
  const fs = baseW >= 30 ? 1 : baseW >= 24 ? .92 : .82;
  const cellHtml = a => !a ? "" : a.map(x => `<span class="${x.bill ? "" : "nb"}${x.kari ? " kari" : ""}">${x.sp ? '<b class="sp">特</b>' : ""}${esc(x.name)}</span>`).join("・");
  const cols = `<colgroup><col style="width:${DATE_W}mm">${bm.map(x => colsOf(x).map(kk => `<col style="width:${isOT(kk) ? OT_W : baseW}mm">`).join("")).join("")}</colgroup>`;
  const tw = DATE_W + nBase * baseW + nOT * OT_W;
  const foot = (lab, f) => `<tr class="bft"><th class="bd">${lab}</th>${bm.map(x => colsOf(x).map(kk => `<td class="n">${f(x.col[kk], x, kk)}</td>`).join("")).join("")}</tr>`;
  const grid = `<table class="bt bday" style="--bf:${fs};width:${tw}mm">${cols}<thead><tr><th rowspan="2" class="bd">日付</th>${bm.map(x => `<th colspan="${colsOf(x).length}">${esc(mLabel(x.m))}</th>`).join("")}</tr>
    <tr>${bm.map(x => colsOf(x).map(kk => `<th class="${isOT(kk) ? "hot" : ""}">${kk}</th>`).join("")).join("")}</tr></thead><tbody>` +
    days.map(k => `<tr class="${dayClass(k)}"><td class="bd">${+k.slice(8)}<small>（${WD[dayDow(k)]}）</small></td>` +
      bm.map(x => colsOf(x).map(kk => isOT(kk)
        ? `<td class="n co" data-o="${k}|${x.m.id}|${kk.slice(0, 1)}">${x.ot[k + "|" + kk.slice(0, 1)] || ""}</td>`
        : `<td class="cc" data-c="${k}|${x.m.id}|${kk}">${cellHtml(x.cell[k + "|" + kk])}</td>`).join("")).join("") + "</tr>").join("") +
    `</tbody><tfoot>${foot("合計数量", c => c.q || "")}${foot("単価", (c, x, kk) => unitsText(c.units) || (x.split && x.split[kk] != null ? yen(x.split[kk]) : ""))}${foot("金額", c => c.q ? yen(c.a) : "")}</tfoot></table>`;
  let tot = 0; const mTot = bm.map(() => Object.fromEntries(KB4.map(kk => [kk, { q: 0, a: 0 }]).concat([["q", 0], ["a", 0]]))), sw = [];
  const rows = names.map(n => {
    let sub = 0;
    const tds = bm.map((x, i) => {
      const s = x.sites[n];
      if (x.split) {   // 昼・昼残業・夜・夜残業 それぞれ 数量・金額（単価は見出し）
        if (!s) return "<td></td><td></td>".repeat(4);
        sub += s.amt;
        return KB4.map(kk => { const y = s.k[kk]; mTot[i][kk].q += y.q; mTot[i][kk].a += y.a;
          return `<td class="n">${y.q || ""}</td><td class="n">${y.q ? yen(y.a) : ""}</td>`; }).join("");
      }
      if (!s) return `<td class="n"></td><td class="n"></td><td class="n"></td>`;
      sub += s.amt; mTot[i].q += s.qty; mTot[i].a += s.amt;
      return `<td class="n">${unitText(s)}</td><td class="n">${s.qty}</td><td class="n">${s.sp ? '<b class="sp">特</b>' : ""}${s.none ? "単価なし" : yen(s.amt)}</td>`;
    }).join("");
    tot += sub;
    return `<tr><td class="bs">${esc(n)}</td><td class="bcd">${esc(siteCode(n))}</td>${tds}${one ? "" : `<td class="n b">${yen(sub)}</td>`}${stampCell(stampKey(ym, g.id, n), sub, sw, g.名前 + "　" + n)}</tr>`;
  });
  /* 月額（リース料）があれば一番上に「本社」＝月額−現場の計（金額だけ） */
  const lease = g.月額 != null && g.月額 !== "" ? +g.月額 : null, hq = lease != null ? lease - tot : null;
  const blank = bm.map(x => "<td></td>".repeat(x.split ? 8 : 3)).join("");
  const hqRow = lease == null ? "" : `<tr class="bhq"><td class="bs">本社<div class="bsub">月額 ${yen(lease)} − 現場の計 ${yen(tot)}</div></td><td></td>${one ? "<td></td><td></td>" : blank}<td class="n b">${yen(hq)}</td>${stampCell(stampKey(ym, g.id, "本社"), hq, sw, g.名前 + "　本社")}</tr>`;
  const grand = lease != null ? lease : tot;
  const warnHq = lease != null && hq < 0 ? [`${g.名前}：現場の計（${yen(tot)}円）が月額（${yen(lease)}円）を超えています（本社がマイナス）`] : [];
  const tall = names.length + (lease != null ? 1 : 0) <= 8;   // 行が少なければ判子が押せる高さ（11mm）、多ければ 8mm
  const head1 = one ? "<th>単価（円）</th><th>数量</th><th>金額（円）</th>"
    : bm.map(x => `<th colspan="${x.split ? 8 : 3}">${esc(mLabel(x.m))}</th>`).join("") + '<th rowspan="2">計（円）</th>';
  const head2 = one ? "" : `<tr>${bm.map(x => x.split ? KB4.map(kk => `<th>${kk}<div class="bsub bsubh">@${x.split[kk] != null ? yen(x.split[kk]) : "未定"}</div></th><th>金額</th>`).join("") : "<th>単価</th><th>数量</th><th>金額</th>").join("")}</tr>`;
  const totRow = one ? `<td></td><td class="n">${mTot[0].q}</td><td class="n">${yen(grand)}</td>`
    : mTot.map((t, i) => bm[i].split ? KB4.map(kk => `<td class="n">${t[kk].q || ""}</td><td class="n">${t[kk].q ? yen(t[kk].a) : ""}</td>`).join("") : `<td></td><td class="n">${t.q}</td><td class="n">${yen(t.a)}</td>`).join("") + `<td class="n">${yen(grand)}</td>`;
  const sum = names.length || lease != null ? `<table class="bt bsum${tall ? " tall" : ""}${bm.some(x => x.split) ? " wide" : ""}"><thead><tr><th${one ? "" : ' rowspan="2"'}>現場名</th><th${one ? "" : ' rowspan="2"'}>現場コード</th>${head1}<th${one ? "" : ' rowspan="2"'} class="bstampH">確認印</th></tr>
      ${head2}</thead><tbody>${hqRow}${rows.join("")}
      <tr class="btot"><td colspan="2">合計</td>${totRow}<td class="bnost"></td></tr></tbody></table>`
    : `<p class="bnone">この月に請求する現場はありません。</p>`;
  return { html: `<div class="bpage${hidden ? " boff" : ""}">${billHead(g.題名 || "測量機稼働表", g.名前, from, made)}<div class="${one ? "bflex" : ""}">${grid}<div class="bsumw"><h4>現場別の集計<small>（この月に使った現場だけ）</small></h4>${sum}</div></div>
    <p class="bfoot">数量＝使った日（昼・夜はそれぞれ1回）×台数${bm.some(x => x.split) ? "、残業は表に入れた数量" : ""}。単価＝請求の単価（実行予算の単価とは別）。<b class="sp">特</b>＝交渉で決めた特別単価。${lease != null ? "本社＝月額（リース料）から現場の計を引いた額。" : ""}</p></div>`,
    warn: [...warnHq, ...sw, ...bm.flatMap(x => x.warn)], total: grand, sites: names.length };
}
function billHead(title, sub, from, made) {
  return `<div class="bh"><div><div class="bt1">${esc(title)}</div><div class="bt2">${esc(sub)}</div></div>
    <div class="br"><div class="bt1">${reiwa(from.slice(0, 7))}</div><div class="bt3">${from.replace(/-/g, "/")} 〜 ${monthLast(from.slice(0, 7)).replace(/-/g, "/")}　作成 ${made}</div></div></div>`;
}


/* 現場ごとの1枚（現場別の請求）：その現場で使った機械（まとまりごと）・区分・使った日 */
function billSitePages(groups, from, to, made) {
  const per = {}, warn = [];
  groups.forEach(g => g.機械.map(machine).filter(Boolean).forEach(m => {
    const x = billMachine(g, m, from, to); warn.push(...x.warn);
    Object.entries(x.sites).forEach(([n, s]) => ((per[n] ||= {})[g.id] ||= { g, rows: [] }).rows.push({ m, s, x }));
  }));
  const names = sortSites(Object.keys(per));
  let all = 0;
  const pages = names.map((n, pi) => {
    let sub = 0;
    const stamps = [];
    const body = Object.values(per[n]).map(({ g, rows }) => {
      const gs = rows.reduce((t, r) => t + r.s.amt, 0); sub += gs;
      stamps.push(`<table class="bt bsum tall bstampbox"><tr><th>${esc(g.名前)}</th></tr><tr>${stampCell(stampKey(from.slice(0, 7), g.id, n), gs, warn, g.名前 + "　" + n)}</tr></table>`);
      return `<tr class="bg"><td colspan="5">${esc(g.名前)}</td></tr>` + rows.map(({ m, s, x }) => {
        const per1 = (x.split || x.flat != null) ? colsOf(x).filter(kk => s.k[kk].q) : null;
        if (!per1) return `<tr><td>${esc(mLabel(m))}</td><td class="bdays">${spanText([...new Set(s.days)])}</td>
          <td class="n">${s.qty}</td><td class="n">${s.sp ? '<b class="sp">特</b>' : ""}${unitText(s)}</td><td class="n">${s.none ? "単価なし" : yen(s.amt)}</td></tr>`;
        return per1.map(kk => { const y = s.k[kk];
          return `<tr><td>${esc(mLabel(m))}　${kk}</td><td class="bdays">${spanText([...new Set(y.days)])}</td>
          <td class="n">${y.q}</td><td class="n">${s.sp && !isOT(kk) ? '<b class="sp">特</b>' : ""}${unitsText(y.units)}</td><td class="n">${yen(y.a)}</td></tr>`; }).join("");
      }).join("");
    }).join("");
    const tax = taxOf(sub); all += sub + tax;
    return `<div class="bpage">${billHead("機械使用料", n, from, made)}
      <p class="bsite">${esc(siteHead(n)) || "&nbsp;"}</p>
      <table class="bt bone"><thead><tr><th>機械・区分</th><th>使った日</th><th>数量</th><th>単価（円）</th><th>金額（円）</th></tr></thead><tbody>${body}</tbody></table>
      <div class="bstamps"><div class="bstl">確認印</div>${stamps.join("")}</div>
      <table class="bt btax"><tr><th>${S.billTax ? "税抜金額" : "金額"}</th><td class="n">${yen(sub)}</td></tr>${S.billTax ? `<tr><th>消費税（${BILL_TAX}%）</th><td class="n">${yen(tax)}</td></tr><tr class="btot"><th>税込金額</th><td class="n">${yen(sub + tax)}</td></tr>` : ""}</table>
      <p class="bfoot">数量＝使った日（昼・夜はそれぞれ1回）×台数。単価＝請求の単価（実行予算の単価とは別）。<b class="sp">特</b>＝交渉で決めた特別単価。</p></div>`;
  });
  return { pages, names, warn, total: all };
}

/* 請求単価を変えるとき（2026-11-13 ユーザー：請求額は後から確認して決め、記録する。記録したものを変えるときは確認を出す）
   まとまりの 記録["単価|機械 id"・"区分単価|機械 id|区分"・"月額"] = "名前 日時" */
const recNow = () => me.name + " " + new Date().toLocaleString("sv-SE").slice(0, 16);
function priceChange(label, old, v) {
  if (old == null || old === "" || old === v) return true;
  return confirm(`記録してある${label}を変えます。\n${yen(old)}円 → ${v == null ? "（空＝単価表）" : yen(v) + "円"}\nよろしいですか？`);
}

function renderBill() {
  const el = $("pBill"); if (!el) return;
  const ym = S.billMonth, from = ym + "-01", to = monthLast(ym), made = todayKey().replace(/-/g, "/");
  const groups = billGroups();
  let pages = [], warn = [], tabs = [];   // tabs：[{ 名, 金額 }]（ページと同じ順）
  if (S.billMode === "site") {
    const r = billSitePages(groups, from, to, made); pages = r.pages; warn = r.warn;
    tabs = r.names.map(n => ({ 名: n, 金額: null }));
  } else {
    const rs = groups.map(g => billGroupPage(g, from, to, made)); pages = rs.map(x => x.html); warn = rs.flatMap(x => x.warn);
    tabs = rs.map((x, i) => ({ 名: groups[i].名前 || "（名前なし）", 金額: x.total }));
  }
  /* タブで1枚ずつ。印刷は全部（または この1枚） */
  let sel = tabs.findIndex(t => t.名 === S.billSel); if (sel < 0) sel = 0;
  pages = pages.map((h, i) => i === sel ? h : h.replace('<div class="bpage">', '<div class="bpage boff">'));
  const uw = [...new Set(warn)];
  const ci = catIndex();
  const allM = D.機械.filter(m => !m.使わない).map((m, i) => [m, ci(catOf(m)), i]).sort((a, b) => a[1] - b[1] || a[2] - b[2]).map(x => x[0]);
  const rec = (g, key) => g.記録?.[key] ? `<small class="brec" title="記録">${esc(g.記録[key])}</small>` : "";
  const setHtml = groups.map((g, gi) => `<div class="bgset" data-g="${gi}">
      <div><label>まとまりの名前 <input data-f="名前" value="${esc(g.名前)}" size="20"></label>
        <label>題名 <input data-f="題名" value="${esc(g.題名 || "")}" size="14"></label>
        <button type="button" class="sub" data-up="${gi}" ${gi ? "" : "disabled"}>▲</button><button type="button" class="sub" data-dn="${gi}" ${gi < groups.length - 1 ? "" : "disabled"}>▼</button>
        <button type="button" class="danger" data-del="${gi}">このまとまりを消す</button></div>
      <div><label>月額（リース料など） <input data-mo="${gi}" class="num" inputmode="numeric" size="8" value="${g.月額 ?? ""}" placeholder="無し"> 円</label> ${rec(g, "月額")}
        <span class="note">入れると、集計の一番上に「本社」＝月額−現場の計 を出します</span></div>
      <div class="bms">${g.機械.map(id => `<span class="bm">${esc(mLabel(machine(id)))}
        ${splitOf(g, id) ? KB4.map(kk => `${kk} <input data-pk="${gi}|${id}|${kk}" class="num" inputmode="numeric" size="6" value="${splitOf(g, id)[kk] ?? ""}" placeholder="未定">${rec(g, "区分単価|" + id + "|" + kk)}`).join(" ") + " 円"
          : `請求単価 <input data-p="${gi}|${id}" class="num" inputmode="numeric" size="6" value="${g.単価?.[id] ?? ""}" placeholder="単価表"> 円${rec(g, "単価|" + id)}`}
        <label class="ck" title="昼・夜で単価が違う機械（ラフタークレーンなど）。請求書に 昼・昼残業・夜・夜残業 の列"><input type="checkbox" data-sp="${gi}|${id}" ${splitOf(g, id) ? "checked" : ""}> 昼・夜・残業で分ける</label> <button type="button" class="sub" data-rm="${gi}|${id}" title="外す">×</button></span>`).join("") || '<span class="note">機械がありません</span>'}
        <select data-add="${gi}"><option value="">＋ 機械を足す</option>${allM.filter(m => !g.機械.includes(m.id)).map(m => `<option value="${m.id}">${esc(catOf(m) + "：" + mLabel(m))}</option>`).join("")}</select></div>
    </div>`).join("");
  el.innerHTML = `<div class="bctl">
      <button class="sub" id="bmPrev">◀</button><input type="month" id="bmMonth" value="${ym}"><button class="sub" id="bmNext">▶</button>
      <b>${reiwa(ym)}</b>
      <span class="bseg"><button data-mode="group" class="${S.billMode === "group" ? "on" : ""}">機械別の請求</button><button data-mode="site" class="${S.billMode === "site" ? "on" : ""}">現場別の請求</button></span>
      <span class="sp"></span>
      <span class="bzoom" title="Ctrl＋ホイール・Ctrl＋＋／－で拡大・縮小、Ctrl＋0 で画面に合わせる"><button class="sub" id="bmZo">－</button><span id="bmZv">${Math.round((S.billZoom || 1) * 100)}%</span><button class="sub" id="bmZi">＋</button></span>
      ${uw.length ? `<button class="bwbtn" id="bmWarn" title="確かめてください">⚠ ${uw.length}</button>` : ""}
      <button class="sub" id="bmSet">⚙ まとまり・単価</button>
      <button class="sub" id="bmPrint1" ${pages.length ? "" : "disabled"}>この1枚を印刷</button><button id="bmPrint" ${pages.length ? "" : "disabled"}>全部印刷</button>
    </div>
    <div class="btabs">${tabs.map((t, i) => `<button data-tabi="${i}" class="${i === sel ? "on" : ""}">${esc(t.名)}${t.金額 != null ? `<small>${yen(t.金額)}円</small>` : ""}</button>`).join("")}</div>
    <div class="bstage">${pages.join("") || '<p class="note" style="padding:12px">請求書にする機械がありません。右上の「⚙ まとまり・単価」でまとまりを作ってください。</p>'}</div>
    <div class="bpop" id="bmWarnP" hidden><div class="bph"><b>確かめてください</b><button class="sub" data-close>閉じる</button></div><ul>${uw.map(w => `<li>${esc(w)}</li>`).join("")}</ul></div>
    <div class="bpop bset" id="bmSetP" ${S.billSetOpen ? "" : "hidden"}><div class="bph"><b>請求書にする機械（まとまり）と請求単価</b><button class="sub" data-close>閉じる</button></div>
      <p class="note">まとまりごとに1枚の請求書になります（全員で同じ設定）。請求単価は請求書だけの単価で、後から確かめて決めて入れます（機械・現場の画面の単価＝実行予算の単価とは別）。
        空にすると単価表で数えます。一度入れた単価を変えるときは確かめます。交渉で決めた特別単価のある予約は、その単価で数えます。</p>
      ${setHtml}<button type="button" class="sub" id="bmAddG">＋ まとまりを足す</button>
    </div>`;
  el.querySelectorAll("[data-tabi]").forEach(b => b.onclick = () => { S.billSel = tabs[+b.dataset.tabi].名; saveView(); renderBill(); });
  el.querySelectorAll("[data-mode]").forEach(b => b.onclick = () => { S.billMode = b.dataset.mode; saveView(); renderBill(); });
  if ($("bmWarn")) $("bmWarn").onclick = () => { $("bmWarnP").hidden = !$("bmWarnP").hidden; };
  $("bmSet").onclick = () => { S.billSetOpen = $("bmSetP").hidden; $("bmSetP").hidden = !S.billSetOpen; };
  el.querySelectorAll("[data-close]").forEach(b => b.onclick = () => { const p = b.closest(".bpop"); p.hidden = true; if (p.id === "bmSetP") S.billSetOpen = false; });
  $("bmPrint").onclick = billPrint;
  $("bmPrint1").onclick = () => { document.body.classList.add("bprint1"); billPrint(); };
  $("bmZo").onclick = () => billZoom(1 / 1.1); $("bmZi").onclick = () => billZoom(1.1);
  $("bmPrev").onclick = () => { S.billMonth = monthMove(ym, -1); saveView(); renderBill(); };
  $("bmNext").onclick = () => { S.billMonth = monthMove(ym, 1); saveView(); renderBill(); };
  $("bmMonth").onchange = () => { if (/^\d{4}-\d\d$/.test($("bmMonth").value)) { S.billMonth = $("bmMonth").value; saveView(); renderBill(); } };
  el.querySelectorAll("[data-stamp]").forEach(td => td.onclick = () => stampClick(td));
  /* 設定 */
  const keepOpen = () => { S.billSetOpen = true; };
  el.querySelectorAll(".bgset input[data-f]").forEach(x => x.onchange = () => { keepOpen(); const gi = +x.closest(".bgset").dataset.g; saveGroups(L => { L[gi][x.dataset.f] = x.value.trim(); }); });
  el.querySelectorAll("[data-p]").forEach(x => x.onchange = () => {
    const [gi, id] = x.dataset.p.split("|"), v = num(x.value), old = groups[+gi].単価?.[id] ?? null;
    if (!priceChange(`${mLabel(machine(id))} の請求単価`, old, v)) { x.value = old ?? ""; return; }
    keepOpen(); const r = recNow();
    saveGroups(L => { const G = L[+gi]; G.単価 = { ...G.単価 }; if (v == null) delete G.単価[id]; else G.単価[id] = v; G.記録 = { ...G.記録, ["単価|" + id]: r }; });
  });
  el.querySelectorAll("[data-sp]").forEach(x => x.onchange = () => { keepOpen(); const [gi, id] = x.dataset.sp.split("|"); saveGroups(L => { const G = L[+gi]; G.区分単価 = { ...G.区分単価 };
    if (x.checked) G.区分単価[id] = { 昼: G.単価?.[id] ?? null, 夜: null, 昼残業: 4000, 夜残業: 5000 }; else delete G.区分単価[id]; }); });   // 残業代 昼 4,000・夜 5,000（2026-11-13 ユーザー）
  el.querySelectorAll("[data-pk]").forEach(x => x.onchange = () => {
    const [gi, id, kk] = x.dataset.pk.split("|"), v = num(x.value), old = splitOf(groups[+gi], id)?.[kk] ?? null;
    if (!priceChange(`${mLabel(machine(id))} の${kk}の請求単価`, old, v)) { x.value = old ?? ""; return; }
    keepOpen(); const r = recNow();
    saveGroups(L => { const G = L[+gi]; G.区分単価 = { ...G.区分単価 }; G.区分単価[id] = { ...G.区分単価[id], [kk]: v }; G.記録 = { ...G.記録, ["区分単価|" + id + "|" + kk]: r }; });
  });
  el.querySelectorAll("[data-mo]").forEach(x => x.onchange = () => {
    const gi = +x.dataset.mo, v = num(x.value), old = groups[gi].月額 ?? null;
    if (!priceChange("月額", old, v)) { x.value = old ?? ""; return; }
    keepOpen(); const r = recNow(); saveGroups(L => { L[gi].月額 = v; L[gi].記録 = { ...L[gi].記録, 月額: r }; });
  });
  el.querySelectorAll("[data-add]").forEach(x => x.onchange = () => { if (!x.value) return; keepOpen(); const gi = +x.dataset.add; saveGroups(L => { L[gi].機械.push(x.value); }); });
  el.querySelectorAll("[data-rm]").forEach(x => x.onclick = () => { keepOpen(); const [gi, id] = x.dataset.rm.split("|"); saveGroups(L => { L[+gi].機械 = L[+gi].機械.filter(v => v !== id); }); });
  el.querySelectorAll("[data-del]").forEach(x => x.onclick = () => { if (!confirm("このまとまりを消します。よろしいですか？（機械や予約は消えません）")) return; keepOpen(); saveGroups(L => { L.splice(+x.dataset.del, 1); }); });
  el.querySelectorAll("[data-up],[data-dn]").forEach(x => x.onclick = () => { keepOpen(); const i = +(x.dataset.up ?? x.dataset.dn), j = x.dataset.up != null ? i - 1 : i + 1; saveGroups(L => { [L[i], L[j]] = [L[j], L[i]]; }); });
  $("bmAddG").onclick = () => { keepOpen(); saveGroups(L => { L.push({ id: newId("g"), 名前: "新しいまとまり", 題名: "機械使用料", 機械: [], 単価: {}, 月額: null }); }); };
  /* 表の編集 */
  const st = el.querySelector(".bstage");
  st.onclick = e => { const td = e.target.closest("td[data-c],td[data-o]"); if (td) { if (BED) return; bSelect(td); } };
  st.ondblclick = e => { const td = e.target.closest("td[data-c],td[data-o]"); if (td) { bSelect(td); bEdit(); } };
  st.oncontextmenu = e => { const td = e.target.closest("td[data-c],td[data-o]"); if (!td) return; e.preventDefault(); bSelect(td); bMenu(e.clientX, e.clientY); };
  if (BSEL) { const td = st.querySelector(`td[data-${BSEL.t}="${BSEL.key}"]`); if (td) td.classList.add("sel"); else BSEL = null; }
  requestAnimationFrame(fitBill);
}

/* ---------- 請求書の表を直す（2026-11-13 ユーザー：表示されている稼働日が間違っていたら表を編集。
   セルを選ぶ・Del で消す・Alt＋↓ でリスト・Ctrl＋Z やり直し・Ctrl＋Y 繰り返し・右クリックで 削除・コピー・貼り付け・やり直し・繰り返し）
   行き先の欄（data-c＝日|機械|区分）を直すと、稼働表の予約が直る（その日だけ 休み にする／同じ現場の予約に足す／1日の予約を作る）。
   残業の欄（data-o）は D.請求書.残業日。どちらも commit なので Ctrl＋Z で戻る（残業日は patchTo で戻す）。 ---------- */
let BSEL = null, BED = null, BCLIP = null;
function bSelect(td) {
  document.querySelectorAll("#pBill td.sel").forEach(x => x.classList.remove("sel"));
  td.classList.add("sel");
  BSEL = td.dataset.c ? { t: "c", key: td.dataset.c } : { t: "o", key: td.dataset.o };
}
const bTd = () => BSEL && document.querySelector(`#pBill td[data-${BSEL.t}="${BSEL.key}"]`);
function bMove(dr, dc) {
  const td = bTd(); if (!td) return;
  const tr = td.parentElement, cells = [...tr.children], ci = cells.indexOf(td);
  let to = null;
  if (dc) { for (let i = ci + dc; i >= 0 && i < cells.length; i += dc) if (cells[i].matches("[data-c],[data-o]")) { to = cells[i]; break; } }
  else { let r = dr > 0 ? tr.nextElementSibling : tr.previousElementSibling; if (r) to = r.children[ci]; }
  if (to && to.matches("[data-c],[data-o]")) { bSelect(to); to.scrollIntoView({ block: "nearest", inline: "nearest" }); }
}
/* 行き先を空にする関数（その日・その区分で使う予約を、その日だけ使わない日にする。1日だけの予約は消す） */
const bClearFn = (k, mid, kb, keep) => d => {
  d.予約 = d.予約.filter(r => !(r.機械 === mid && r.区分 === kb && r.開始 === k && r.終了 === k && r.現場 !== keep && works(r, k)));
  d.予約.forEach(r => {
    if (r.機械 !== mid || r.区分 !== kb || r.現場 === keep || !works(r, k)) return;
    r.出勤 = r.出勤.filter(x => x !== k); if (planned(r, k)) r.休み.push(k); stamp(r);
  });
};
/* 行き先を site にする関数（ほかの現場は外す → 同じ現場の予約がその日を含めば使う日に戻す → 前後の日に続いていれば延ばす → 無ければ1日の予約） */
const bSetFn = (k, mid, kb, site, nid) => d => {
  bClearFn(k, mid, kb, site)(d);
  const same = d.予約.filter(r => r.機械 === mid && r.区分 === kb && r.現場 === site);
  if (same.some(r => works(r, k))) return;
  const use = r => { r.休み = r.休み.filter(x => x !== k); r.未使用 = r.未使用.filter(x => x !== k); if (!planned(r, k)) r.出勤.push(k); stamp(r); };
  let r = same.find(r => r.開始 <= k && k <= r.終了);
  if (!r) { r = same.find(r => r.終了 === dayAdd(k, -1)); if (r) r.終了 = k; }
  if (!r) { r = same.find(r => r.開始 === dayAdd(k, 1)); if (r) r.開始 = k; }
  if (r) { use(r); return; }
  const x = { id: nid, 機械: mid, 区分: kb, 現場: site, 開始: k, 終了: k, 土: isSat(k), 日祝: isSunHol(k), 休み: [], 出勤: [], 未使用: [], 状態: "確定", メモ: "請求書の表で直した" };
  stamp(x); d.予約.push(x);
};
const bOtFn = (key, v) => d => { d.請求書 = Object.assign({}, d.請求書); d.請求書.残業日 = Object.assign({}, d.請求書.残業日); if (v) d.請求書.残業日[key] = v; else delete d.請求書.残業日[key]; };
function bClear() {
  const td = bTd(); if (!td) return;
  if (BSEL.t === "o") { if (td.textContent.trim()) commit(bOtFn(BSEL.key, null)); return; }
  const [k, mid, kb] = BSEL.key.split("|");
  if (!td.textContent.trim()) return;
  commit(bClearFn(k, mid, kb, null));
}
function bSet(v) {
  if (!BSEL) return;
  if (BSEL.t === "o") { const n = num(v); commit(bOtFn(BSEL.key, n || null)); return; }
  const [k, mid, kb] = BSEL.key.split("|"), site = String(v ?? "").trim();
  if (!site) { commit(bClearFn(k, mid, kb, null)); return; }
  if (!D.現場.some(s => s.名前 === site)) { alert(`「${site}」は登録されていない行き先です。リスト（Alt＋↓）から選んでください。`); return; }
  commit(bSetFn(k, mid, kb, site, newId("r")));
}
function bCopy() {
  const td = bTd(); if (!td) return;
  BCLIP = BSEL.t === "o" ? td.textContent.trim() : (bTd().querySelector("span:not(.nb)")?.textContent || td.textContent).replace(/^特/, "").split("・")[0].trim();
  try { navigator.clipboard?.writeText(BCLIP); } catch { }
  toast("コピー：" + (BCLIP || "（空）"));
}
async function bPaste() {
  let v = BCLIP;
  try { const t = await navigator.clipboard?.readText(); if (t != null && t !== "") v = t.trim().split(/\r?\n/)[0]; } catch { }
  if (v == null) return;
  bSet(v);
}
/* 編集：行き先はリスト、残業は数字の欄 */
function bEdit(first) {
  const td = bTd(); if (!td) return;
  bEditClose();
  const r = td.getBoundingClientRect();
  const box = document.createElement("div"); box.id = "bEdit"; box.style.left = r.left + "px"; box.style.top = r.bottom + 2 + "px";
  if (BSEL.t === "o") {
    box.innerHTML = `<input id="bEditIn" inputmode="decimal" size="6" value="${esc(first ?? td.textContent.trim())}"> <span class="note">Enter で決める・Esc でやめる</span>`;
    document.body.appendChild(box); BED = box;
    const i = $("bEditIn"); i.focus(); if (first == null) i.select();
    i.onkeydown = e => { if (e.key === "Enter") { e.preventDefault(); const v = i.value; bEditClose(); bSet(v); } else if (e.key === "Escape") { e.preventDefault(); bEditClose(); } };
    return;
  }
  const kb = BSEL.key.split("|")[2];
  const cur = td.textContent.trim();
  const list = D.現場.filter(s => !s.完了 && s.種類 !== "整備").sort((a, b) => (a.種類 === "現場" ? 0 : 1) - (b.種類 === "現場" ? 0 : 1) || (b.コード || "").localeCompare(a.コード || ""));
  box.innerHTML = `<select id="bEditIn" size="12"><option value="">（空欄にする）</option>${list.map(s => `<option ${cur.includes(s.名前) ? "selected" : ""} value="${esc(s.名前)}">${esc(s.名前)}${s.種類 !== "現場" ? "（" + s.種類 + "）" : ""}</option>`).join("")}</select>
    <div class="note">${kb}。Enter／ダブルクリックで決める・Esc でやめる</div>`;
  document.body.appendChild(box); BED = box;
  const sl = $("bEditIn"); sl.focus();
  const pick = () => { const v = sl.value; bEditClose(); bSet(v); };
  sl.onkeydown = e => { if (e.key === "Enter") { e.preventDefault(); pick(); } else if (e.key === "Escape") { e.preventDefault(); bEditClose(); } };
  sl.ondblclick = pick;
}
function bEditClose() { if (BED) { BED.remove(); BED = null; } }
document.addEventListener("pointerdown", e => { if (BED && !e.target.closest("#bEdit")) bEditClose(); });
function bMenu(x, y) {
  const o = BSEL.t === "o";
  showCtx(x, y, [
    ["削除（Del）", bClear],
    ["コピー（Ctrl＋C）", bCopy],
    ["貼り付け（Ctrl＋V）", bPaste],
    [o ? "数量を入れる（Enter）" : "リストから選ぶ（Alt＋↓）", () => bEdit()],
    null,
    ["やり直し（Ctrl＋Z）", undo, hist.length ? "" : "off"],
    ["繰り返し（Ctrl＋Y）", redo, fut.length ? "" : "off"]
  ]);
}
document.addEventListener("keydown", e => {
  if (S.tab !== "bill" || !BSEL || BED || e.target.closest("input,select,textarea") || document.querySelector("dialog[open]")) return;
  const c = e.ctrlKey || e.metaKey;
  if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); bClear(); }
  else if (e.altKey && e.key === "ArrowDown") { e.preventDefault(); bEdit(); }
  else if (!c && !e.altKey && e.key.startsWith("Arrow")) { e.preventDefault(); bMove(e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0, e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0); }
  else if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); bEdit(); }
  else if (c && e.key.toLowerCase() === "c") { e.preventDefault(); bCopy(); }
  else if (c && e.key.toLowerCase() === "v") { e.preventDefault(); bPaste(); }
  else if (c && e.key.toLowerCase() === "x") { e.preventDefault(); bCopy(); bClear(); }
  else if (BSEL.t === "o" && !c && /^[0-9.]$/.test(e.key)) { e.preventDefault(); bEdit(e.key); }
  else if (e.key === "Escape") { document.querySelectorAll("#pBill td.sel").forEach(x => x.classList.remove("sel")); BSEL = null; }
});

/* 見本の紙を、見えている所に収まる大きさ×拡大率（Ctrl＋ホイール・Ctrl＋＋／－、Ctrl＋0 で 100%＝画面に合わせる）で出す。印刷は縮めない（app.css） */
if (!(S.billZoom > 0)) S.billZoom = 1;
function fitBill() {
  const st = document.querySelector("#pBill .bstage"), pg = st?.querySelector(".bpage:not(.boff)"); if (!pg) return;
  pg.style.zoom = 1;
  const k = Math.min(1, (st.clientHeight - 12) / pg.offsetHeight, (st.clientWidth - 12) / pg.offsetWidth);
  pg.style.zoom = Math.max(.3, Math.max(.45, k) * S.billZoom).toFixed(3);
  if ($("bmZv")) $("bmZv").textContent = Math.round(S.billZoom * 100) + "%";
}
function billZoom(f) { S.billZoom = f === 0 ? 1 : Math.min(4, Math.max(.5, S.billZoom * f)); saveView(); bEditClose(); fitBill(); }
addEventListener("resize", () => { if (S.tab === "bill") fitBill(); });
addEventListener("wheel", e => { if (e.ctrlKey && S.tab === "bill") { e.preventDefault(); billZoom(e.deltaY < 0 ? 1.1 : 1 / 1.1); } }, { passive: false });
document.addEventListener("keydown", e => {
  if (S.tab !== "bill" || !(e.ctrlKey || e.metaKey) || e.altKey || !["+", ";", "=", "-", "_", "0"].includes(e.key)) return;
  e.preventDefault(); billZoom(e.key === "0" ? 0 : e.key === "-" || e.key === "_" ? 1 / 1.1 : 1.1);
});

/* 印刷：請求書のときだけ A4 縦・余白 左 25mm・右 10mm（2026-11-13 ユーザー）。稼働表の印刷は app.css の A3 横 */
function billPageStyle(on) {
  let st = $("billPageStyle");
  if (!st) { st = document.createElement("style"); st.id = "billPageStyle"; document.head.appendChild(st); }
  st.textContent = on ? "@media print { @page { size: A4 portrait; margin: 9mm 10mm 9mm 25mm; } }" : "";
}
function billPrint() { bEditClose(); billPageStyle(true); setTimeout(() => print(), 50); }
addEventListener("beforeprint", () => billPageStyle(S.tab === "bill"));
addEventListener("afterprint", () => { billPageStyle(false); document.body.classList.remove("bprint1"); });
if (S.tab === "bill" && ready) renderBill();
