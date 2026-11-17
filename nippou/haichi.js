/* 配置表（配置報告）から取り込む（2026-11-02 ユーザー：クロジカではなく配置報告から過去の作業内容を取り込み、個人のところに反映。
   ○○現場の報告に作業員Aがいたら、作業員Aの工事日報に行を足して、その人の1か月分の日報が作られる形に）
   ・読むのは配置表入力の 配置表データ（会社の OneDrive 社内共有フォルダ\06_配置表\配置表データ\<年>.js と 追加分.js。読むだけ）。
     1日＝{ 確定, 予定 }。その日の確定があれば確定（＝報告）、無ければ予定（＝予定として入れ、報告済みにはしない）。
   ・現場の行の 作業員（と「担当が行った」なら担当の社員）の1人ずつに、その人の日報（個人_<名前>_<月>.json）の行を足す。
     行先＝現場名（現場一覧の名前を確かめる）・区分（昼勤／夜勤 → 8:00〜17:00／21:00〜6:00）・作業内容・同じ現場の人（行.同行。参加者には入れない＝人数を重ねて数えない）。記入者＝その人。
   ・本社内勤の欄：本社内勤＝本社、有給・休み・明け休み＝休み、ほか（健康診断など）＝本社（記載を作業内容に）。機材センターの人＝機材センター。（「本社・休み・機材センターも入れる」で選ぶ）
   ・取り込んだ行は 行.取込 = "haichi:日|現場名|区分"。もう一度取り込んでも重ならない。月ごとの取り込み方は imp.js（クロジカと同じ）。 */
"use strict";
let HAI = null;   // { from, to, hon, years: { 年: { 確定, 予定 } }, map: Map(現場名の形 → { raw, n, v, code, how, on }), pol, def }
const haiDays = H => { const out = []; for (let k = H.from; k <= H.to; k = dayAdd(k, 1)) out.push(k); return out; };
// 配置表データの年のファイル（.js。中身は JSON を変数に入れているだけ）を読む
const jsBody = t => JSON.parse(t.slice(t.indexOf("= {") + 2).trim().replace(/;\s*$/, ""));
async function haiFile(name) {
  if (DEMO) return (window.HAICHI_DEMO || {})[name] || null;
  const c = CFG.haichi;
  for (const p of [c.path, c.oldPath].filter(Boolean)) {
    const res = await g(`/drives/${c.driveId}/root:/${enc(p + "/" + name)}`);
    if (res.status === 404) continue;
    if (!res.ok) throw httpErr(res, name + " を読めません");
    const dl = (await res.json())["@microsoft.graph.downloadUrl"];
    const body = await fetch(dl, { cache: "no-store" });
    if (!body.ok) throw httpErr(body, name + " を読めません");
    return jsBody(await body.text());
  }
  return null;
}
async function haiLoad(H) {
  const ys = [...new Set([H.from.slice(0, 4), H.to.slice(0, 4)])].filter(y => !H.years[y]);
  if (!ys.length) return;
  const add = H.add || (H.add = await haiFile("追加分.js").catch(() => null) || { 確定: {}, 予定: {} });
  for (const y of ys) {
    const d = await haiFile(y + ".js") || { 確定: {}, 予定: {} };
    ["確定", "予定"].forEach(v => Object.entries(add[v] || {}).forEach(([k, x]) => { if (k.startsWith(y)) (d[v] ||= {})[k] = x; }));
    H.years[y] = d;
  }
}
// その日の配置表（確定が無ければ予定）
function haiDay(H, k) {
  const y = H.years[k.slice(0, 4)] || {}, c = (y.確定 || {})[k], p = (y.予定 || {})[k];
  return c ? { d: c.data, plan: false } : p ? { d: p.data, plan: true } : null;
}
// 担当（名字だけ）→ 社員の名前（ひとりに決まるときだけ）
function tantoName(t) { const L = S.M.社員.filter(n => nk(n).startsWith(nk(t))); return nk(t) && L.length === 1 ? L[0] : ""; }
const haiNight = kb => /夜/.test(String(kb || ""));
// 作業員の名前：名簿の名前にそろえる。名簿に無ければ、字数が同じで1字だけちがう名簿の人（配置表の名の字の打ち間違い。例：山田 太朗 → 山田 太郎）。
// 「安全大会」「健康診断」など人でないものは外す（空を返す）
function haiPerson(n) {
  if (findPerson(n)) return findPerson(n);
  const t = String(n || "").trim(); if (!t || RE_KURO_OFF.test(nk(t)) || /大会|診断|会議|研修|講習/.test(t)) return "";
  const a = nk(t), L = people().filter(p => { const b = nk(p); return b.length === a.length && [...b].filter((c, i) => c !== a[i]).length === 1; });
  return L.length === 1 ? L[0] : personName(t);
}
// 本社内勤の欄の書き方：本社内勤／本社内勤（○○＝本社＋現場名／有給・休み・明け休み＝休み／ほか＝本社（書いてあることを作業内容に）
function haiHonsha(t) {
  t = String(t || "").normalize("NFKC").trim();
  if (/有給|有休|休み|^休|明け|代休|振替|忌引/.test(t)) return { off: true, work: t };
  const m = /^本社(?:内勤)?\s*\((.+?)\)?$/.exec(t);
  if (m) { const x = kuroGuess(m[1].trim()); return x.how === "一覧にない" ? { work: m[1].trim() } : { rel: { 名前: x.v, コード: x.code || "" } }; }
  return { work: /^本社(内勤)?$/.test(t) ? "" : t };
}
const haiPlaces = d => [...(d.現場 || []).map(gx => gx.現場名).filter(Boolean)];
// 現場名の対応表を作る（現場一覧の名前に寄せる。クロジカと同じ kuroGuess）
function haiBuild(H) {
  const old = H.map || new Map(), map = new Map();
  haiDays(H).forEach(k => { const x = haiDay(H, k); if (!x) return;
    haiPlaces(x.d).forEach(p => { const key = kKey(p); if (!map.has(key)) map.set(key, old.get(key) || { raw: p, n: 0, ...kuroGuess(p), on: true }); map.get(key).n = 0; }); });
  haiDays(H).forEach(k => { const x = haiDay(H, k); if (x) haiPlaces(x.d).forEach(p => map.get(kKey(p)).n++); });
  H.map = map;
}
function haiRow(k, place, code, night, work, plan, mates, owner, id, extra) {
  const r = { id: uid(), 行先: place, コード: code || "", 始め: night ? "21:00" : "8:00", 終わり: night ? "6:00" : "17:00", 終日: false, 区分: night ? "夜" : "昼",
    作業: "", 作業行: [{ 予定: plan ? work : "", 報告: plan ? "" : work }], 予定: { 人: [], 外注: [], 機械: [] }, 人: [], 外注: [], 機械: [],
    記入者: owner, 取込: id, 更新: nowStr(), ...(extra || {}) };
  if (mates.length) r.同行 = mates;
  if (plan) r.予定のみ = true;
  syncWork(r);
  return { owner, k, r };
}
function haiItems(H) {
  const out = [];
  haiDays(H).forEach(k => {
    const x = haiDay(H, k); if (!x) return;
    const d = x.d, by = new Map();   // 同じ日・同じ現場・同じ区分の行はまとめる（配置表で2行に分かれていることがある）
    (d.現場 || []).forEach(gx => {
      const m = H.map.get(kKey(gx.現場名)); if (!m || !m.on || !m.v) return;
      const night = haiNight(gx.区分), key = kKey(gx.現場名) + "|" + (night ? "夜" : "昼");
      if (!by.has(key)) by.set(key, { m, night, raw: gx.現場名, ppl: [], work: [] });
      const o = by.get(key), add = n => { n = haiPerson(n); if (n && !o.ppl.some(v => nk(v) === nk(n))) o.ppl.push(n); };
      (gx.作業員 || []).forEach(w => add(w.氏名));
      if (gx.担当が行った && tantoName(gx.担当)) add(tantoName(gx.担当));
      if (gx.作業内容 && !o.work.includes(gx.作業内容)) o.work.push(gx.作業内容);
    });
    by.forEach(o => o.ppl.forEach(n => out.push(haiRow(k, o.m.v, o.m.code, o.night, o.work.join("\n"), x.plan, o.ppl.filter(v => v !== n), n,
      `haichi:${k}|${nk(o.raw)}|${o.night ? "夜" : "昼"}`))));
    if (!H.hon) return;
    (d.本社 || []).forEach(h => { const n = haiPerson(h.氏名); if (!n) return; const t = haiHonsha(h.記載);
      out.push(haiRow(k, t.off ? "休み" : "本社", "", false, t.work || "", x.plan, [], n, `haichi:${k}|本社|${nk(n)}`, t.off ? { 終日: true, 始め: "", 終わり: "" } : t.rel ? { 関連現場: t.rel } : null)); });
    const kc = d.機材センター || {}, kp = (kc.人 || []).map(v => haiPerson(v.氏名)).filter(Boolean);
    kp.forEach(n => out.push(haiRow(k, "機材センター", "", false, String(kc.作業内容 || ""), x.plan, kp.filter(v => v !== n), n, `haichi:${k}|機材センター|${nk(n)}`)));
  });
  return out;
}

/* ---------- 窓 ---------- */
function haichiOpen() {
  const t = todayKey(), a = dateOf(monthOf(t) + "-01"); a.setMonth(a.getMonth() - 2);
  HAI = { from: keyOf(a), to: t, hon: true, years: {}, map: null, pol: {}, def: "" };
  haiDraw();
}
async function haiRead() {
  const H = HAI;
  $("hGo") && ($("hGo").disabled = true);
  $("hMsg").textContent = "配置表を読んでいます…";
  try { await haiLoad(H); } catch (e) { $("hMsg").textContent = "読めませんでした：" + e.message; return; }
  haiBuild(H); haiDraw();
}
function haiDraw() {
  const H = HAI, cls = { 一致: "ok", 候補: "maybe", 一覧にない: "ng" };
  const head = `<div class="dh"><h2>配置表から取り込む</h2><button type="button" data-close aria-label="閉じる">×</button></div>
    <p class="pinfo">配置表（配置確定。無い日は配置予定）の現場の作業員ひとりずつに、その人の日報の行を足します。その人の1か月分の日報ができます。<br>
    作業内容は配置表の作業内容、時間は昼勤 8:00〜17:00・夜勤 21:00〜6:00（あとで直せます）。同じ現場の人は「同じ現場」に出します。</p>
    <div class="dt">期間 <input type="date" id="hFrom" value="${H.from}"> 〜 <input type="date" id="hTo" value="${H.to}">
      <label class="ck"><input type="checkbox" id="hHon"${H.hon ? " checked" : ""}> 本社・休み・機材センターも入れる</label></div>`;
  if (!H.map) {
    dlg(head + `<p id="hMsg" class="pinfo"></p><div class="btns"><button data-close>閉じる</button><button class="pri" id="hRead">配置表を読む</button></div>`, "wide keep");
    haiBindHead(); $("hRead").onclick = haiRead; return;
  }
  const items = haiItems(H), ppl = new Set(items.map(i => nk(i.owner))), list = [...H.map.entries()].sort((a, b) => (a[1].how === "一致") - (b[1].how === "一致") || b[1].n - a[1].n);
  const nDays = haiDays(H).filter(k => haiDay(H, k)).length, nPlan = haiDays(H).filter(k => { const x = haiDay(H, k); return x && x.plan; }).length;
  dlg(head + `<p class="kuro-file" id="hMsg">配置表 ${nDays}日分${nPlan ? `（うち ${nPlan}日は確定が無いので予定から。予定として入れ、報告済みにはしません）` : ""}</p>
    <p class="pinfo">現場名ごとに、日報に入れる現場名を確かめてください。<b class="maybe">候補</b>・<b class="ng">一覧にない</b>は打ち直せます（打つと候補）。チェックを外すと取り込みません。</p>
    <div class="gwrap kuro-wrap"><table class="list kuro"><tr><th>取込</th><th>配置表の現場名</th><th>日数</th><th>日報の現場名</th><th></th></tr>
    ${list.map(([key, x], i) => `<tr><td><input type="checkbox" data-hon="${i}"${x.on ? " checked" : ""}></td><td>${esc(x.raw)}</td><td class="num">${x.n}</td>
      <td><div class="pick"><input data-hv="${i}" value="${esc(x.v)}" autocomplete="off"></div>${x.alt && x.alt.length ? `<small class="pinfo">ほかの候補：${esc(x.alt.join("、"))}</small>` : ""}</td>
      <td><span class="kh ${cls[x.how] || ""}">${esc(x.how)}</span></td></tr>`).join("")}</table></div>
    ${polHTML(H, items)}
    <div class="btns"><span style="flex:1"></span><button data-close>閉じる</button><button class="pri" id="hGo"${items.length ? "" : " disabled"}>${items.length}件を取り込む（${ppl.size}人）</button></div>`, "wide keep");
  haiBindHead();
  const B = $("dlgBody");
  B.querySelectorAll("[data-hon]").forEach(el => el.onchange = () => { list[+el.dataset.hon][1].on = el.checked; haiDraw(); });
  B.querySelectorAll("[data-hv]").forEach(el => {
    const x = list[+el.dataset.hv][1];
    const set = (v, code) => { const st = siteOf(v, code), ks = koteiSite(v, code); x.v = v; x.code = code || (st && st.コード) || (ks && ks.コード) || ""; x.how = st || ks ? "一致" : v ? "一覧にない" : x.how; const sc = B.querySelector(".kuro-wrap").scrollTop; haiDraw(); $("dlgBody").querySelector(".kuro-wrap").scrollTop = sc; };
    bindPick(el, () => placeItems().filter(y => y.grp !== "本社など"), y => { el.value = y.v; set(y.v, y.コード); });
    el.addEventListener("change", () => { const v = el.value.trim(); if (v !== x.v) set(v, ""); });
  });
  bindPol(H, haiDraw);
  $("hGo").onclick = () => haiImport(items).catch(e => { toast("取り込めませんでした：" + e.message, 8000); haiDraw(); });
}
function haiBindHead() {
  const H = HAI, re = () => { if (H.map) haiRead(); else haiDraw(); };
  $("hFrom").onchange = () => { const v = $("hFrom").value; if (v) { H.from = v; if (H.to < v) H.to = v; re(); } };
  $("hTo").onchange = () => { const v = $("hTo").value; if (v) { H.to = v; if (H.from > v) H.from = v; re(); } };
  $("hHon").onchange = () => { H.hon = $("hHon").checked; if (H.map) haiDraw(); };
}
async function haiImport(items) {
  $("hGo").disabled = true;
  const st = await importItems(items, { srcs: ["haichi"], pol: HAI.pol, label: "配置表", prog: impProg });
  if (!st) { haiDraw(); return; }
  HAI = null; render();
  importDone(st);
}
