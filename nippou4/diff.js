/* 他の記録とくらべる（2026-11-03 ユーザー：工事日報に、他のソースとの違いがある場合、警告を出して修正するために確認する機能）
   ・カレンダーの上に「⚠ 他の記録と違う ○件」（見ている月。今日まで）。押すと人・日ごとの一覧。
   ・くらべる元：配置確定（配置表データ。読むだけ）・常用日報（工程データ\常用日報\現場_*.json。読むだけ）・クロジカ（CSV を選んだときだけ。この画面を開いている間）。
   ・くらべるのは「その人がその日どの現場にいたか」（工事コード、無ければ現場名）。昼夜・作業内容はくらべない。名簿にいる人だけ（外注の人は名前が名簿に無い）。
     配置確定：その日の確定があれば、全員をくらべる。常用日報：その日に人が書いてある現場についてだけ。クロジカ：その人が出てくる日だけ（全員が書いているわけではない）。
   ・直す：その人の日報のその日を、選んだ元に合わせる（合わない行を消し、足りない現場を足す。行.取込 = "check:日|現場"）。ほかの人の日報に参加者として入っている分は消さない。
   ・「このままでよい」：日報のフォルダの 確認済みの違い.json に残す（だれ・いつ・そのときの中身）。元の中身が変わったらまた出す。 */
"use strict";
const DIFF = { m: "", list: null, busy: false, err: [], kuro: null, all: false, H: { years: {} } };
const DONE_FILE = "確認済みの違い.json";
const jouyouStore = DEMO ? DemoFolder("工程/常用日報/") : GraphFolder("kotei", "常用日報");
const DSRC = [["hai", "配置確定"], ["jou", "詳細日報"], ["kuro", "クロジカ"]];
// 現場の形：key が同じなら同じ現場（配置表の名前と現場一覧の名前がちがっても、工事コードが同じなら同じ）
function plOf(name, code, night, work) {
  name = String(name || "").trim();
  if (isOffice(name) || /^機材センター$/.test(name)) return { key: "n" + nk(name), label: name, code: "", night: false, work: work || "" };
  let c = code || "", label = name;
  if (!c) { const g = kuroGuess(name); if (g.how === "一致" || g.how === "候補") { c = g.code || ""; label = g.v; } }
  const st = siteOf(label, c); if (st) label = siteName(st);
  return { key: c ? "c" + c : "n" + kKey(name), label, code: c, night: !!night, work: work || "" };
}
const REST = "休み";
const REST_RE = /有給|有休|休み|休暇|明け|代休|振替|忌引/;
const diffPerson = n => { const p = haiPerson(n); return p && findPerson(p) ? findPerson(p) : ""; };
function dput(map, name, k, src, pl) {
  const id = nk(name) + "|" + k;
  if (!map.has(id)) map.set(id, { id, name, k, app: new Map() });
  const e = map.get(id);
  if (!src) { e.app.set(pl.key, pl); return; }
  (e[src] ||= new Map()).set(pl.key, pl);
}
const dayList = m => { const out = []; for (let k = m + "-01"; monthOf(k) === m && k <= todayKey(); k = dayAdd(k, 1)) out.push(k); return out; };
// 配置確定（予定はくらべない）
async function diffHai(map, m, ctx) {
  const H = DIFF.H; H.from = m + "-01"; H.to = dayAdd(m + "-28", 3); await haiLoad(H);
  dayList(m).forEach(k => {
    const c = ((H.years[k.slice(0, 4)] || {}).確定 || {})[k]; if (!c) return;
    ctx.haiDays.add(k);
    const d = c.data, add = (n, pl) => { const p = diffPerson(n); if (p) dput(map, p, k, "hai", pl); };
    (d.現場 || []).forEach(g => {
      const pl = plOf(g.現場名, "", haiNight(g.区分), g.作業内容);
      (g.作業員 || []).forEach(w => add(w.氏名, pl));
      if (g.担当が行った && tantoName(g.担当)) add(tantoName(g.担当), pl);
    });
    (d.本社 || []).forEach(h => { const t = haiHonsha(h.記載); add(h.氏名, t.off ? plOf(REST, "", false, t.work) : plOf("本社", "", false, t.work || "")); });
    const kc = d.機材センター || {}; (kc.人 || []).forEach(w => add(w.氏名, plOf("機材センター", "", false, String(kc.作業内容 || ""))));
    (Array.isArray(d.休み) ? d.休み : []).forEach(w => { const n = typeof w === "string" ? w : w && w.氏名; if (n) add(n, plOf(REST, "", false, "")); });
  });
}
// 常用日報：その日に人が書いてある現場
async function diffJou(map, m, ctx) {
  const files = (await jouyouStore.list()).filter(f => /^現場_.*\.json$/.test(f.name));
  const docs = await Promise.all(files.map(f => jouyouStore.get(f.name).catch(() => null)));
  docs.filter(Boolean).forEach(x => { const s = x.data.現場 || {};
    dayList(m).forEach(k => { const d = (s.日 || {})[k]; if (!d) return;
      const ppl = [...(d.下 || []), ...(d.外注 || [])].filter(r => r && r[0]); if (!ppl.length) return;
      const L = (d.上 || []).filter(r => r && r.some(Boolean)).map(r => r[1] || r[0] || "").filter(Boolean).join(" ");
      const pl = plOf(s.略称 || s.名前, s.コード, false, L);
      (ctx.jouSites[k] ||= new Set()).add(pl.key);
      ppl.forEach(r => { const p = diffPerson(r[0]); if (p) dput(map, p, k, "jou", pl); });
    }); });
}
// クロジカ（選んだ CSV）
function diffKuro(map, m) {
  (DIFF.kuro || []).forEach(g => { if (monthOf(g.日) !== m || g.日 > todayKey()) return;
    const t = kuroTitle(g.予定), raw = nk(g.予定);
    let pl = null;
    if (RE_KURO_OFF.test(raw)) { if (REST_RE.test(raw)) pl = plOf(REST, "", false, ""); }
    else if (t.honsha) pl = plOf("本社", "", false, g.内容);
    else if (t.base) pl = plOf(t.base, "", t.night, g.内容);
    if (!pl) return;
    const ppl = [...g.名前, ...g.設備.map(v => eqKind(v).人).filter(Boolean)];
    ppl.forEach(n => { const p = diffPerson(n); if (p) dput(map, p, g.日, "kuro", pl); });
  });
}
const dsame = (a, b) => a.size === b.size && [...a.keys()].every(x => b.has(x));
const dtxt = M => M ? (M.size ? [...M.values()].map(p => p.label).join("、") : "（なし）") : "―";
// 違いの一覧を作る
async function diffCalc(m) {
  const map = new Map(), ctx = { haiDays: new Set(), jouSites: {} }; DIFF.err = [];
  for (const doc of await monthDocs(m, true)) Object.entries(doc.日 || {}).forEach(([k, d]) => { if (k > todayKey()) return;
    (d.行 || []).forEach(r => { if (!r.行先) return; const pl = plOf(r.行先, r.コード, kubunOf(r) === "夜", ""); pl.rowOf = true;
      membersOf({ who: doc.名前, r }).forEach(n => { const p = findPerson(n); if (p) dput(map, p, k, null, pl); }); }); });
  await diffHai(map, m, ctx).catch(e => DIFF.err.push("配置確定：" + e.message));
  await diffJou(map, m, ctx).catch(e => DIFF.err.push("常用日報：" + e.message));
  diffKuro(map, m);
  const rest = plOf(REST).key, done = await nippouStore.get(DONE_FILE).catch(() => null), D = done ? done.data.済 || {} : {};
  const out = [];
  map.forEach(e => {
    if (isGone(e.name, m)) return;
    e.diff = {};
    if (ctx.haiDays.has(e.k)) { const h = e.hai || new Map(); if (!dsame(e.app, h) && !(!e.app.size && h.size === 1 && h.has(rest))) e.diff.hai = true; if (!e.hai) e.hai = h; }
    const sites = ctx.jouSites[e.k];
    if (sites) { e.jouAt = sites; const j = e.jou || new Map(); if ([...sites].some(s => e.app.has(s) !== j.has(s))) e.diff.jou = true; if (!e.jou && [...sites].some(s => e.app.has(s))) e.jou = j; }
    if (e.kuro && !dsame(e.app, e.kuro)) e.diff.kuro = true;
    if (!Object.keys(e.diff).length) return;
    e.fp = JSON.stringify(["app", "hai", "jou", "kuro"].map(s => e[s] ? [...e[s].keys()].sort() : null));
    e.done = D[e.id] && D[e.id].中身 === e.fp ? D[e.id] : null;
    out.push(e);
  });
  return out.sort((a, b) => a.k.localeCompare(b.k) || people().findIndex(n => nk(n) === nk(a.name)) - people().findIndex(n => nk(n) === nk(b.name)));
}
// カレンダーの上の警告
async function diffBadge(force) {
  if (DEMO && !window.HAICHI_DEMO) return;
  const m = monthOf(S.day), b = () => $("cDiff");
  const show = () => { const el = b(); if (!el || !DIFF.list || DIFF.m !== m) return; const n = DIFF.list.filter(e => !e.done).length; el.hidden = !n; el.textContent = `⚠ 他の記録と違う ${n}件`; };
  if (DIFF.m === m && DIFF.list && !force) { show(); return; }
  if (DIFF.busy === m) return;
  DIFF.busy = m;
  try { const L = await diffCalc(m); DIFF.m = m; DIFF.list = L; } catch (e) { console.warn("くらべられませんでした", e); }
  DIFF.busy = false;
  show();
}
function diffOpen() {
  const L = (DIFF.list || []).filter(e => DIFF.all || !e.done), m = DIFF.m;
  const cell = (e, s) => `<td class="${e.diff[s] ? "dng" : ""}">${esc(dtxt(e[s]))}</td>`;
  dlg(`<div class="dh"><h2>他の記録との違い（${+m.slice(5)}月 ${L.length}件）</h2><button type="button" data-close aria-label="閉じる">×</button></div>
    <p class="pinfo">その人がその日にいた現場を、工事日報と他の記録でくらべました（色の付いた所が工事日報と違います）。どれが正しいか確かめて、合わせるか「このままでよい」を選んでください。</p>
    ${DIFF.err.length ? `<p class="ng">${esc(DIFF.err.join("／"))}</p>` : ""}
    <div class="dt"><button type="button" id="dKuro">クロジカの CSV もくらべる${DIFF.kuro ? "（読みました）" : ""}</button><label class="ck"><input type="checkbox" id="dAll"${DIFF.all ? " checked" : ""}> 確認済みも出す</label></div>
    <div class="gwrap kuro-wrap"><table class="list askt dift"><tr><th>日</th><th>人</th><th>工事日報</th><th>配置確定</th><th>詳細日報</th><th>クロジカ</th><th>どうするか</th></tr>
    ${L.map((e, i) => `<tr><td>${+e.k.slice(5, 7)}/${+e.k.slice(8)}(${WD[dow(e.k)]})</td><td>${esc(e.name)}</td><td>${esc(dtxt(e.app))}</td>${DSRC.map(([s]) => cell(e, s)).join("")}
      <td class="dact">${DSRC.filter(([s]) => e.diff[s]).map(([s, t]) => `<button type="button" data-fix="${i}|${s}">${t}に合わせる</button>`).join("")}${e.done ? `<small class="pinfo">確認済み（${esc(e.done.だれ)}）</small>` : `<button type="button" data-ok="${i}">このままでよい</button>`}</td></tr>`).join("") || `<tr><td colspan="7">違いはありません。</td></tr>`}</table></div>
    <div class="btns"><button data-close>閉じる</button></div>`, "wide keep");
  $("dAll").onchange = () => { DIFF.all = $("dAll").checked; diffOpen(); };
  $("dKuro").onclick = diffPickKuro;
  const B = $("dlgBody");
  B.querySelectorAll("[data-fix]").forEach(el => el.onclick = () => { const [i, s] = el.dataset.fix.split("|"); diffFix(L[+i], s).catch(e => toast("直せませんでした：" + e.message, 8000)); });
  B.querySelectorAll("[data-ok]").forEach(el => el.onclick = () => diffKeep(L[+el.dataset.ok]).catch(e => toast("残せませんでした：" + e.message, 8000)));
}
async function diffPickKuro() {
  const files = await kuroFolder(); if (!files || !files.length) return;
  const rows = []; for (const f of files) rows.push(...readKuroText(await fileText(f)));
  DIFF.kuro = kuroGroup(rows);
  await diffBadge(true); diffOpen();
}
// その人のその日を、選んだ元に合わせる
async function diffFix(e, src) {
  // 常用日報にその人がいなければ、常用日報の現場の行だけ外す（常用日報はほかの現場を知らない）。いれば、その現場にする
  const t = new Map(src === "jou" && !(e.jou && e.jou.size) ? [...e.app].filter(([key]) => !e.jouAt.has(key)) : []);
  (e[src] || new Map()).forEach((pl, key) => t.set(key, pl));
  const txt = [...t.values()].map(p => p.label).join("、") || "（なし）";
  if (!confirm(`${e.name} の ${+e.k.slice(5, 7)}/${+e.k.slice(8)} を「${txt}」にします（今は「${dtxt(e.app)}」）。よいですか？`)) return;
  const m = monthOf(e.k), name = `個人_${nk(e.name)}_${m}.json`;
  for (let tries = 0; ; tries++) {
    const cur = await nippouStore.get(name), doc = cur ? cur.data : { 形式: "私の日報", 版: 1, 名前: e.name, アカウント: "", 月: m, 日: {} };
    const d = doc.日[e.k] ||= { 行: [], 報告: null };
    d.行 = d.行.filter(r => t.has(plOf(r.行先, r.コード).key));
    const has = new Set(d.行.map(r => plOf(r.行先, r.コード).key));
    t.forEach((pl, key) => { if (has.has(key)) return;
      const off = pl.label === REST;
      d.行.push(haiRow(e.k, pl.label, pl.code, pl.night, pl.work, false, [], e.name, `check:${e.k}|${key}`, off ? { 終日: true, 始め: "", 終わり: "" } : null).r); });
    if (!d.報告) d.報告 = { だれ: S.me.name, いつ: "他の記録に合わせて直した" };
    d.更新 = doc.更新 = nowStr();
    try { await nippouStore.put(name, doc, cur ? cur.eTag : null); break; }
    catch (x) { if ((x.status !== 412 && x.status !== 409) || tries >= 4) throw x; }
  }
  delete DOCS[m];
  if (nk(e.name) === nk(S.me.name)) { const f = S.F[m]; if (f && !f.pending.length && !f.saving) f.loaded = false; await loadMonth(m, true); }
  toast("直しました");
  await diffBadge(true); diffOpen(); render();
}
async function diffKeep(e) {
  for (let tries = 0; ; tries++) {
    const cur = await nippouStore.get(DONE_FILE), doc = cur ? cur.data : { 形式: "工事日報の確認済みの違い", 版: 1, 済: {} };
    doc.済[e.id] = { だれ: S.me.name, いつ: nowStr(), 中身: e.fp };
    try { await nippouStore.put(DONE_FILE, doc, cur ? cur.eTag : null); break; }
    catch (x) { if ((x.status !== 412 && x.status !== 409) || tries >= 4) throw x; }
  }
  e.done = { だれ: S.me.name, 中身: e.fp };
  diffOpen(); diffBadge();
}
