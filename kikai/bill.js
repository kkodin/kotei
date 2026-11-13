/* ---------- 請求書（2026-11-13） ----------
   ユーザー：光波距離計・オートレベルを、月に1回、現場別に分けた請求書にして印刷したい（いまはユーザー1人の作業）。
   形は 測量機稼働表.xlsx（令和8年8月分 など）：日付×機械（昼・夜）の表＋現場別の 数量・金額。請求書に載せる現場は、その月に使った現場だけ。
   将来ほかの機械も同じ形にするかもしれない（表の形は変わるかも）→ 機械は「まとまり」に入れるだけで請求書になるようにした。
   - まとまり：{ id, 名前, 題名, 機械:[機械 id] }。D.請求書.まとまり に持つ（全員で同じ。データの形は足すだけ）。無いあいだは機械の名前（レベル・光波）から作る。
   - 印刷の分け方（S.billMode）：group＝まとまりごとに1枚（Excel と同じ）／site＝現場ごとに1枚。
   - 数量＝使った日（昼と夜はそれぞれ1回。同じ日に昼夜使えば2）×台数。使わなかった（申告）の日は数えない（works）。
   - 金額は集計と同じ costOf（単価表・特別単価・料金の決め方）。積置き（請求する）はその現場（billTo）。置場・整備は請求しない。
   - 消費税（S.billTax）：載せない（2026-11-13 ユーザー）。仕組みは残す（現場ごとに 税抜×税率 の1円未満を切り捨て）。
   - 2026-11-13 ユーザー：請求の単価は実行予算（単価表）と別。レベル 200円・光波距離計 1,500円（まとまりの 単価＝{ 機械 id: 円／回 }）。
     請求単価がある機械は 請求単価×数量（特別単価・単価表は使わない）。無い機械は集計と同じ costOf。
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

/* 1台の機械のその月：日ごとの行き先（表）と、請求する現場ごとの 数量・金額 */
function billMachine(m, from, to, price = null) {
  const rs = D.予約.filter(r => r.機械 === m.id && r.開始 <= to && r.終了 >= from);
  const cell = {};   // "日|区分" → [{ name, bill }]
  rs.forEach(r => {
    for (let k = r.開始 < from ? from : r.開始; k <= to && k <= r.終了; k = dayAdd(k, 1)) {
      if (!works(r, k)) continue;
      const name = isTsumi(r.現場) && r.積置き?.現場 ? r.現場 + "（" + r.積置き.現場 + "）" : r.現場;
      (cell[k + "|" + r.区分] ||= []).push({ name, bill: isBill(r), kari: r.状態 === "仮" });
    }
  });
  const sites = {}, warn = [];
  rs.filter(isBill).forEach(r => {
    let c = costOf(r, from, to, false, m);
    const split = !!price && typeof price === "object", p = split ? price[r.区分] ?? null : price;   // 昼・夜で単価が違う機械（ラフタークレーン）は 区分単価
    if (split || price != null) { const q = c.days * c.n; c = { meth: "日", n: c.n, days: c.days, qty: c.days, unit: p, amt: p == null ? (split ? 0 : null) : p * q }; }   // 請求単価（単価表・特別単価は使わない）
    if (!c.days && !c.amt) return;
    const x = (sites[billTo(r)] ||= { qty: 0, amt: 0, units: new Set(), sp: false, none: false, days: [] });
    x.qty += c.meth === "日" ? c.days * c.n : c.qty * c.n;
    if (c.amt == null) x.none = true; else x.amt += c.amt;
    if (c.unit != null) x.units.add(c.unit);
    if (special(r) && price == null) x.sp = true;
    for (let k = r.開始 < from ? from : r.開始; k <= to && k <= r.終了; k = dayAdd(k, 1)) if (works(r, k)) x.days.push(k);
    x.meth = c.meth;
    if (split) { x.k ||= { 昼: 0, 夜: 0, 残業: 0 }; x.k[r.区分] += c.days * c.n; }
    if (r.状態 === "仮") warn.push(`${mLabel(m)} ${r.区分}「${r.現場}」${md(r.開始)}〜${md(r.終了)} が「仮」のまま`);
    if (special(r) && (split || price != null)) warn.push(`${mLabel(m)}「${billTo(r)}」に特別単価（交渉）がありますが、請求書は請求単価で数えています`);
    if (c.amt == null) warn.push(`${mLabel(m)}「${billTo(r)}」は単価がありません（金額に入っていません）`);
  });
  Object.entries(cell).forEach(([key, a]) => { if (a.length > 1) warn.push(`${mLabel(m)} ${md(key.split("|")[0])} ${key.split("|")[1]} に行き先が2つ（${a.map(x => x.name).join("・")}）`); });
  return { m, cell, sites, warn };
}
/* 昼・夜・残業で単価を分ける機械（2026-11-13 ユーザー：自社のラフタークレーンは昼夜で単価が変わる。昼・夜・残業（単価は後で入れる）の欄）
   まとまりの 区分単価[機械 id] = { 昼, 夜, 残業 }。昼・夜の数量は稼働表の予約から、残業の数量は請求書の画面で現場ごとに入れる（D.請求書.残業["YYYY-MM|まとまり|機械|現場"]） */
const KB3 = ["昼", "夜", "残業"];
const splitOf = (g, id) => g.区分単価?.[id] || null;
const otKey = (ym, gid, mid, site) => [ym, gid, mid, site].join("|");
function billMachineG(g, m, from, to) {
  const sp = splitOf(g, m.id), x = billMachine(m, from, to, sp || (g.単価?.[m.id] ?? null));
  if (sp) {
    const ym = from.slice(0, 7);
    Object.entries(x.sites).forEach(([n, s]) => {
      s.k ||= { 昼: 0, 夜: 0, 残業: 0 };
      s.k.残業 = +(D.請求書?.残業?.[otKey(ym, g.id, m.id, n)] || 0);
      s.amt += s.k.残業 * (sp.残業 || 0);
    });
    KB3.forEach(k => { if (sp[k] == null && Object.values(x.sites).some(s => s.k[k])) x.warn.push(`${mLabel(m)}：${k}の請求単価がまだ入っていません（金額は0円で数えています）`); });
  }
  x.split = sp;
  return x;
}
const unitText = s => s.meth && s.meth !== "日" ? [...s.units].map(u => yen(u) + "／" + UNIT[s.meth]).join("・") : [...s.units].map(yen).join("・");
const taxOf = v => S.billTax ? Math.floor(v * BILL_TAX / 100) : 0;
const siteCode = name => siteOf(name).コード || "";   // 2026-11-13 ユーザー：正式工事名・元請は要らない。社内の現場名と現場コードだけ
const siteHead = name => siteCode(name) ? "現場コード " + siteCode(name) : "";
const sortSites = names => names.sort((a, b) => (siteOf(a).コード || "999999").localeCompare(siteOf(b).コード || "999999") || a.localeCompare(b, "ja"));

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

/* まとまりごとの1枚：日付×機械（昼・夜）の表＋使った現場だけの集計 */
function billGroupPage(g, from, to, made) {
  const ms = g.機械.map(machine).filter(Boolean), bm = ms.map(m => billMachineG(g, m, from, to)), one = ms.length === 1 && !bm[0].split;
  const days = []; for (let k = from; k <= to; k = dayAdd(k, 1)) days.push(k);
  const names = sortSites([...new Set(bm.flatMap(x => Object.keys(x.sites)))]);
  const fs = ms.length <= 2 ? 1 : ms.length <= 3 ? .9 : ms.length <= 4 ? .8 : .7;
  const cellHtml = a => !a ? "" : a.map(x => `<span class="${x.bill ? "" : "nb"}${x.kari ? " kari" : ""}">${esc(x.name)}</span>`).join("・");
  /* 列の幅は決めておく（2026-11-13 ユーザー：無駄に列幅が広く、印刷すると両端が切れる）。紙の幅 180mm に収める。
     機械1台なら 日付＋昼夜 で 86mm にして、右に集計を並べる。2台以上は 日付 16mm＋残りを昼夜で等分（1列 40mm まで） */
  const colW = one ? 35 : Math.min(40, Math.floor(164 / (2 * ms.length)));
  const cols = `<colgroup><col style="width:16mm">${ms.map(() => `<col style="width:${colW}mm"><col style="width:${colW}mm">`).join("")}</colgroup>`;
  const grid = `<table class="bt bday" style="--bf:${fs};width:${16 + colW * 2 * ms.length}mm">${cols}<thead><tr><th rowspan="2" class="bd">日付</th>${ms.map(m => `<th colspan="2">${esc(mLabel(m))}</th>`).join("")}</tr>
    <tr>${ms.map(() => "<th>昼</th><th>夜</th>").join("")}</tr></thead><tbody>` +
    days.map(k => `<tr class="${dayClass(k)}"><td class="bd">${+k.slice(8)}<small>（${WD[dayDow(k)]}）</small></td>` +
      bm.map(x => `<td>${cellHtml(x.cell[k + "|昼"])}</td><td>${cellHtml(x.cell[k + "|夜"])}</td>`).join("") + "</tr>").join("") + "</tbody></table>";
  let tot = 0; const mTot = bm.map(() => ({ qty: 0, amt: 0 })), sw = [], ym = from.slice(0, 7);
  const rows = names.map(n => {
    let sub = 0;
    const tds = bm.map((x, i) => {
      const s = x.sites[n];
      if (x.split) {   // 昼・夜・残業 それぞれ 数量・金額（単価は見出し）。残業の数量は画面で入れる
        if (!s) return "<td></td><td></td>".repeat(3);
        sub += s.amt; mTot[i].amt += s.amt; mTot[i].k ||= { 昼: 0, 夜: 0, 残業: 0, a昼: 0, a夜: 0, a残業: 0 };
        return KB3.map(k => { const q = s.k[k], a = q * (x.split[k] || 0); mTot[i].k[k] += q; mTot[i].k["a" + k] += a;
          const qh = k === "残業" ? `<input class="botin" inputmode="decimal" data-ot="${esc(otKey(ym, g.id, x.m.id, n))}" value="${q || ""}">` : q || "";
          return `<td class="n">${qh}</td><td class="n">${q ? yen(a) : ""}</td>`; }).join("");
      }
      if (!s) return `<td class="n"></td><td class="n"></td><td class="n"></td>`;
      sub += s.amt; mTot[i].qty += s.qty; mTot[i].amt += s.amt;
      return `<td class="n">${unitText(s)}</td><td class="n">${s.qty}</td><td class="n">${s.sp ? '<b class="sp">特</b>' : ""}${s.none ? "単価なし" : yen(s.amt)}</td>`;
    }).join("");
    tot += sub;
    return `<tr><td class="bs">${esc(n)}</td><td class="bc">${esc(siteCode(n))}</td>${tds}${one ? "" : `<td class="n b">${yen(sub)}</td>`}${stampCell(stampKey(ym, g.id, n), sub, sw, g.名前 + "　" + n)}</tr>`;
  });
  /* 月額（リース料）があれば一番上に「本社」＝月額−現場の計（金額だけ） */
  const lease = g.月額 != null && g.月額 !== "" ? +g.月額 : null, hq = lease != null ? lease - tot : null;
  const blank = bm.map(x => x.split ? "<td></td>".repeat(6) : `<td></td><td></td><td></td>`).join("");
  const hqRow = lease == null ? "" : one ? `<tr class="bhq"><td class="bs">本社<div class="bsub">月額 ${yen(lease)} − 現場の計 ${yen(tot)}</div></td><td></td><td></td><td></td><td class="n b">${yen(hq)}</td>${stampCell(stampKey(ym, g.id, "本社"), hq, sw, g.名前 + "　本社")}</tr>`
    : `<tr class="bhq"><td class="bs">本社<div class="bsub">月額 ${yen(lease)} − 現場の計 ${yen(tot)}</div></td><td></td>${blank}<td class="n b">${yen(hq)}</td>${stampCell(stampKey(ym, g.id, "本社"), hq, sw, g.名前 + "　本社")}</tr>`;
  const grand = lease != null ? lease : tot;
  const warnHq = lease != null && hq < 0 ? [`${g.名前}：現場の計（${yen(tot)}円）が月額（${yen(lease)}円）を超えています（本社がマイナス）`] : [];
  const tall = names.length + (lease != null ? 1 : 0) <= 8;   // 行が少なければ判子が押せる高さ（11mm）、多ければ 8mm
  const sum = names.length || lease != null ? `<table class="bt bsum${tall ? " tall" : ""}"><thead><tr><th${one ? "" : ' rowspan="2"'}>現場名</th><th${one ? "" : ' rowspan="2"'}>現場コード</th>${one ? "<th>単価（円）</th><th>数量</th><th>金額（円）</th>" : bm.map(x => `<th colspan="${x.split ? 6 : 3}">${esc(mLabel(x.m))}</th>`).join("") + '<th rowspan="2">計（円）</th>'}<th${one ? "" : ' rowspan="2"'} class="bstampH">確認印</th></tr>
      ${one ? "" : `<tr>${bm.map(x => x.split ? KB3.map(k => `<th>${k} 数量<div class="bsub bsubh">@${x.split[k] != null ? yen(x.split[k]) : "未定"}</div></th><th>${k} 金額</th>`).join("") : "<th>単価</th><th>数量</th><th>金額</th>").join("")}</tr>`}</thead><tbody>${hqRow}${rows.join("")}
      <tr class="btot"><td colspan="2">合計</td>${one ? `<td></td><td class="n">${mTot[0].qty}</td><td class="n">${yen(grand)}</td>` : mTot.map((t, i) => bm[i].split ? KB3.map(k => `<td class="n">${t.k?.[k] || ""}</td><td class="n">${t.k?.[k] ? yen(t.k["a" + k]) : ""}</td>`).join("") : `<td></td><td class="n">${t.qty}</td><td class="n">${yen(t.amt)}</td>`).join("") + `<td class="n">${yen(grand)}</td>`}<td class="bnost"></td></tr></tbody></table>`
    : `<p class="bnone">この月に請求する現場はありません。</p>`;
  return { html: `<div class="bpage">${billHead(g.題名 || "測量機稼働表", g.名前, from, made)}<div class="${one ? "bflex" : ""}">${grid}<div class="bsumw"><h4>現場別の集計<small>（この月に使った現場だけ）</small></h4>${sum}</div></div>
    <p class="bfoot">数量＝使った日（昼・夜はそれぞれ1回）×台数。単価＝請求の単価（実行予算の単価とは別）。${lease != null ? "本社＝月額（リース料）から現場の計を引いた額。" : ""}</p></div>`,
    warn: [...warnHq, ...sw, ...bm.flatMap(x => x.warn)], total: grand, sites: names.length };
}
function billHead(title, sub, from, made) {
  return `<div class="bh"><div><div class="bt1">${esc(title)}</div><div class="bt2">${esc(sub)}</div></div>
    <div class="br"><div class="bt1">${reiwa(from.slice(0, 7))}</div><div class="bt3">${from.replace(/-/g, "/")} 〜 ${monthLast(from.slice(0, 7)).replace(/-/g, "/")}　作成 ${made}</div></div></div>`;
}

/* 現場ごとの1枚：その現場で使った機械（まとまりごと）と使った日 */
function billSitePages(groups, from, to, made) {
  const per = {}, warn = [];
  groups.forEach(g => g.機械.map(machine).filter(Boolean).forEach(m => {
    const x = billMachineG(g, m, from, to); warn.push(...x.warn);
    Object.entries(x.sites).forEach(([n, s]) => ((per[n] ||= {})[g.id] ||= { g, rows: [] }).rows.push({ m, s, sp: x.split }));
  }));
  let all = 0;
  const pages = sortSites(Object.keys(per)).map(n => {
    let sub = 0;
    const stamps = [];
    const body = Object.values(per[n]).map(({ g, rows }) => {
      const gs = rows.reduce((t, r) => t + r.s.amt, 0); sub += gs;
      stamps.push(`<table class="bt bsum tall bstampbox"><tr><th>${esc(g.名前)}</th></tr><tr>${stampCell(stampKey(from.slice(0, 7), g.id, n), gs, warn, g.名前 + "　" + n)}</tr></table>`);
      return `<tr class="bg"><td colspan="5">${esc(g.名前)}</td></tr>` + rows.map(({ m, s, sp }) => sp ? KB3.filter(k => s.k[k]).map(k => `<tr><td>${esc(mLabel(m))}　${k}</td><td class="bdays">${k === "残業" ? "" : spanText([...new Set(s.days)])}</td>
        <td class="n">${s.k[k]}</td><td class="n">${sp[k] != null ? yen(sp[k]) : "未定"}</td><td class="n">${yen(s.k[k] * (sp[k] || 0))}</td></tr>`).join("")
        : `<tr><td>${esc(mLabel(m))}</td><td class="bdays">${spanText([...new Set(s.days)])}</td>
        <td class="n">${s.qty}</td><td class="n">${s.sp ? '<b class="sp">特</b>' : ""}${unitText(s)}</td><td class="n">${s.none ? "単価なし" : yen(s.amt)}</td></tr>`).join("");
    }).join("");
    const tax = taxOf(sub); all += sub + tax;
    return `<div class="bpage">${billHead("測量機 使用料", n, from, made)}
      <p class="bsite">${esc(siteHead(n)) || "&nbsp;"}</p>
      <table class="bt bone"><thead><tr><th>機械</th><th>使った日</th><th>数量</th><th>単価（円）</th><th>金額（円）</th></tr></thead><tbody>${body}</tbody></table>
      <div class="bstamps"><div class="bstl">確認印</div>${stamps.join("")}</div>
      <table class="bt btax"><tr><th>${S.billTax ? "税抜金額" : "金額"}</th><td class="n">${yen(sub)}</td></tr>${S.billTax ? `<tr><th>消費税（${BILL_TAX}%）</th><td class="n">${yen(tax)}</td></tr><tr class="btot"><th>税込金額</th><td class="n">${yen(sub + tax)}</td></tr>` : ""}</table>
      <p class="bfoot">数量＝使った日（昼・夜はそれぞれ1回）×台数。単価＝請求の単価（実行予算の単価とは別）。${S.billTax ? "消費税は1円未満切り捨て。" : ""}</p></div>`;
  });
  return { pages, names: sortSites(Object.keys(per)), warn, total: all };
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
    tabs = rs.map((x, i) => ({ 名: groups[i].名前 || "（名前なし）", 金額: x.total, n: x.sites }));
  }
  /* タブで1枚ずつ（2026-11-13 ユーザー：タブで切り替え、スクロールしないで月末まで見たい）。印刷は全部（または この1枚） */
  let sel = tabs.findIndex(t => t.名 === S.billSel); if (sel < 0) sel = 0;
  pages = pages.map((h, i) => i === sel ? h : h.replace('<div class="bpage">', '<div class="bpage boff">'));
  const uw = [...new Set(warn)];
  const ci = catIndex();
  const allM = D.機械.filter(m => !m.使わない).map((m, i) => [m, ci(catOf(m)), i]).sort((a, b) => a[1] - b[1] || a[2] - b[2]).map(x => x[0]);
  const setHtml = groups.map((g, gi) => `<div class="bgset" data-g="${gi}">
      <div><label>まとまりの名前 <input data-f="名前" value="${esc(g.名前)}" size="20"></label>
        <label>題名 <input data-f="題名" value="${esc(g.題名 || "")}" size="14"></label>
        <button type="button" class="sub" data-up="${gi}" ${gi ? "" : "disabled"}>▲</button><button type="button" class="sub" data-dn="${gi}" ${gi < groups.length - 1 ? "" : "disabled"}>▼</button>
        <button type="button" class="danger" data-del="${gi}">このまとまりを消す</button></div>
      <div><label>月額（リース料など） <input data-mo="${gi}" class="num" inputmode="numeric" size="8" value="${g.月額 ?? ""}" placeholder="無し"> 円</label>
        <span class="note">入れると、集計の一番上に「本社」＝月額−現場の計 を出します</span></div>
      <div class="bms">${g.機械.map(id => `<span class="bm">${esc(mLabel(machine(id)))} ${splitOf(g, id) ? KB3.map(k => `${k} <input data-pk="${gi}|${id}|${k}" class="num" inputmode="numeric" size="6" value="${splitOf(g, id)[k] ?? ""}" placeholder="未定">`).join(" ") + " 円"
          : `請求単価 <input data-p="${gi}|${id}" class="num" inputmode="numeric" size="6" value="${g.単価?.[id] ?? ""}" placeholder="単価表"> 円`}
        <label class="ck" title="昼・夜で単価が違う機械（ラフタークレーンなど）。請求書に 昼・夜・残業 の欄"><input type="checkbox" data-sp="${gi}|${id}" ${splitOf(g, id) ? "checked" : ""}> 昼・夜・残業で分ける</label> <button type="button" class="sub" data-rm="${gi}|${id}" title="外す">×</button></span>`).join("") || '<span class="note">機械がありません</span>'}
        <select data-add="${gi}"><option value="">＋ 機械を足す</option>${allM.filter(m => !g.機械.includes(m.id)).map(m => `<option value="${m.id}">${esc(catOf(m) + "：" + mLabel(m))}</option>`).join("")}</select></div>
    </div>`).join("");
  el.innerHTML = `<div class="bctl">
      <button class="sub" id="bmPrev">◀</button><input type="month" id="bmMonth" value="${ym}"><button class="sub" id="bmNext">▶</button>
      <b>${reiwa(ym)}</b>
      <select id="bmMode"><option value="group" ${S.billMode === "group" ? "selected" : ""}>まとまりごとに1枚</option><option value="site" ${S.billMode === "site" ? "selected" : ""}>現場ごとに1枚</option></select>
      <span class="sp"></span>
      ${uw.length ? `<button class="bwbtn" id="bmWarn" title="確かめてください">⚠ ${uw.length}</button>` : ""}
      <button class="sub" id="bmSet">⚙ まとまり・単価</button>
      <button class="sub" id="bmPrint1" ${pages.length ? "" : "disabled"}>この1枚を印刷</button><button id="bmPrint" ${pages.length ? "" : "disabled"}>全部印刷</button>
    </div>
    <div class="btabs">${tabs.map((t, i) => `<button data-tabi="${i}" class="${i === sel ? "on" : ""}">${esc(t.名)}${t.金額 != null ? `<small>${yen(t.金額)}円</small>` : ""}</button>`).join("")}</div>
    <div class="bstage">${pages.join("") || '<p class="note" style="padding:12px">請求書にする機械がありません。右上の「⚙ まとまり・単価」でまとまりを作ってください。</p>'}</div>
    <div class="bpop" id="bmWarnP" hidden><div class="bph"><b>確かめてください</b><button class="sub" data-close>閉じる</button></div><ul>${uw.map(w => `<li>${esc(w)}</li>`).join("")}</ul></div>
    <div class="bpop bset" id="bmSetP" ${S.billSetOpen ? "" : "hidden"}><div class="bph"><b>請求書にする機械（まとまり）と請求単価</b><button class="sub" data-close>閉じる</button></div>
      <p class="note">まとまりごとに1枚の請求書になります（全員で同じ設定）。請求単価は請求書だけの単価です（実行予算の単価表・集計の金額とは別）。空にすると単価表で数えます。</p>
      ${setHtml}<button type="button" class="sub" id="bmAddG">＋ まとまりを足す</button>
    </div>`;
  el.querySelectorAll("[data-tabi]").forEach(b => b.onclick = () => { S.billSel = tabs[+b.dataset.tabi].名; saveView(); renderBill(); });
  if ($("bmWarn")) $("bmWarn").onclick = () => { $("bmWarnP").hidden = !$("bmWarnP").hidden; };
  $("bmSet").onclick = () => { S.billSetOpen = $("bmSetP").hidden; $("bmSetP").hidden = !S.billSetOpen; };
  el.querySelectorAll("[data-close]").forEach(b => b.onclick = () => { const p = b.closest(".bpop"); p.hidden = true; if (p.id === "bmSetP") S.billSetOpen = false; });
  $("bmPrint1").onclick = () => { document.body.classList.add("bprint1"); billPrint(); };
  requestAnimationFrame(fitBill);
  $("bmPrev").onclick = () => { S.billMonth = monthMove(ym, -1); saveView(); renderBill(); };
  $("bmNext").onclick = () => { S.billMonth = monthMove(ym, 1); saveView(); renderBill(); };
  $("bmMonth").onchange = () => { if (/^\d{4}-\d\d$/.test($("bmMonth").value)) { S.billMonth = $("bmMonth").value; saveView(); renderBill(); } };
  $("bmMode").onchange = () => { S.billMode = $("bmMode").value; saveView(); renderBill(); };
  $("bmPrint").onclick = billPrint;
  el.querySelectorAll("[data-stamp]").forEach(td => td.onclick = () => stampClick(td));
  const keepOpen = () => { S.billSetOpen = true; };
  el.querySelectorAll(".bgset input[data-f]").forEach(x => x.onchange = () => { keepOpen(); const gi = +x.closest(".bgset").dataset.g; saveGroups(L => { L[gi][x.dataset.f] = x.value.trim(); }); });
  el.querySelectorAll("[data-p]").forEach(x => x.onchange = () => { keepOpen(); const [gi, id] = x.dataset.p.split("|"), v = num(x.value); saveGroups(L => { L[+gi].単価 = { ...L[+gi].単価 }; if (v == null) delete L[+gi].単価[id]; else L[+gi].単価[id] = v; }); });
  el.querySelectorAll("[data-sp]").forEach(x => x.onchange = () => { keepOpen(); const [gi, id] = x.dataset.sp.split("|"); saveGroups(L => { const G = L[+gi]; G.区分単価 = { ...G.区分単価 };
    if (x.checked) G.区分単価[id] = { 昼: G.単価?.[id] ?? null, 夜: null, 残業: null }; else delete G.区分単価[id]; }); });
  el.querySelectorAll("[data-pk]").forEach(x => x.onchange = () => { keepOpen(); const [gi, id, k] = x.dataset.pk.split("|"), v = num(x.value);
    saveGroups(L => { const G = L[+gi]; G.区分単価 = { ...G.区分単価 }; G.区分単価[id] = { ...G.区分単価[id], [k]: v }; }); });
  el.querySelectorAll("[data-ot]").forEach(x => x.onchange = () => { const v = num(x.value), key = x.dataset.ot;
    commit(d => { d.請求書 = Object.assign({}, d.請求書); d.請求書.残業 = Object.assign({}, d.請求書.残業); if (v) d.請求書.残業[key] = v; else delete d.請求書.残業[key]; }, { noUndo: true }); });
  el.querySelectorAll("[data-mo]").forEach(x => x.onchange = () => { keepOpen(); const v = num(x.value); saveGroups(L => { L[+x.dataset.mo].月額 = v; }); });
  el.querySelectorAll("[data-add]").forEach(x => x.onchange = () => { if (!x.value) return; keepOpen(); const gi = +x.dataset.add; saveGroups(L => { L[gi].機械.push(x.value); }); });
  el.querySelectorAll("[data-rm]").forEach(x => x.onclick = () => { keepOpen(); const [gi, id] = x.dataset.rm.split("|"); saveGroups(L => { L[+gi].機械 = L[+gi].機械.filter(v => v !== id); }); });
  el.querySelectorAll("[data-del]").forEach(x => x.onclick = () => { if (!confirm("このまとまりを消します。よろしいですか？（機械や予約は消えません）")) return; keepOpen(); saveGroups(L => { L.splice(+x.dataset.del, 1); }); });
  el.querySelectorAll("[data-up],[data-dn]").forEach(x => x.onclick = () => { keepOpen(); const i = +(x.dataset.up ?? x.dataset.dn), j = x.dataset.up != null ? i - 1 : i + 1; saveGroups(L => { [L[i], L[j]] = [L[j], L[i]]; }); });
  $("bmAddG").onclick = () => { keepOpen(); saveGroups(L => { L.push({ id: newId("g"), 名前: "新しいまとまり", 題名: "機械使用料", 機械: [], 単価: {}, 月額: null }); }); };
}

/* 見本の紙を、見えている所の高さ・幅に収まるように縮める（月末までスクロールしないで見る）。印刷は縮めない（app.css） */
function fitBill() {
  const st = document.querySelector("#pBill .bstage"), pg = st?.querySelector(".bpage:not(.boff)"); if (!pg) return;
  pg.style.zoom = 1;
  const k = Math.min(1, (st.clientHeight - 12) / pg.offsetHeight, (st.clientWidth - 12) / pg.offsetWidth);
  pg.style.zoom = Math.max(.45, k).toFixed(3);
}
addEventListener("resize", () => { if (S.tab === "bill") fitBill(); });

/* 印刷：請求書のときだけ A4 縦（稼働表の印刷は app.css の A3 横） */
function billPageStyle(on) {
  let st = $("billPageStyle");
  if (!st) { st = document.createElement("style"); st.id = "billPageStyle"; document.head.appendChild(st); }
  st.textContent = on ? "@media print { @page { size: A4 portrait; margin: 9mm 10mm; } }" : "";
}
function billPrint() { billPageStyle(true); setTimeout(() => print(), 50); }
addEventListener("beforeprint", () => billPageStyle(S.tab === "bill"));
addEventListener("afterprint", () => { billPageStyle(false); document.body.classList.remove("bprint1"); });
if (S.tab === "bill" && ready) renderBill();
