/* 取り込みの共通部分（クロジカ・常用日報・配置表）。2026-11-02 ユーザー：
   「全て取り込むか／アプリの同じ日に予定がある場合どちらを優先するか／個別に確認するかを、月ごとに選択する仕組み」
   ・取り込むものは { owner（だれの日報に入れるか）, k（日）, r（行。r.取込 = "kuro:…"／"jouyou:…"／"haichi:…"） }。
   ・「アプリの予定」＝その人・その日にすでにある行のうち、いま取り込んでいる元（o.srcs）から入れたもの以外（手で書いた行・ほかの元から取り込んだ行）。
   ・月ごとに選ぶ：all＝全て取り込む（両方残す）／app＝アプリを優先（アプリの予定がある日は入れない）／
     imp＝取り込む方を優先（その日のアプリの予定を消して入れる）／ask＝個別に確認（人・日ごとに選ぶ窓を出す）。
   ・もう入っているもの（行.取込・行.取込他 が同じ）は入れない（何度取り込んでも重ならない）。 */
"use strict";
const POLS = [["app", "アプリの予定を優先"], ["imp", "取り込む方を優先"], ["all", "全て取り込む（両方残す）"], ["ask", "個別に確認する"]];
const POL_TIP = { app: "その人のその日にアプリの予定があれば、取り込まない", imp: "その人のその日のアプリの予定を消して、取り込んだものにする（ほかの人の日報に参加者として入っている分は消さない）", all: "アプリの予定はそのまま、取り込んだものも足す", ask: "取り込む前に、重なる人・日ごとに選ぶ窓を出す" };
// 月ごとの選ぶ欄（P.pol = { "YYYY-MM": "app" … }）
function polHTML(P, items) {
  const by = new Map();
  items.forEach(it => { const m = monthOf(it.k); if (!by.has(m)) by.set(m, { n: 0, p: new Set() }); const o = by.get(m); o.n++; o.p.add(nk(it.owner)); });
  const ms = [...by.keys()].sort();
  if (!ms.length) return "";
  ms.forEach(m => { P.pol[m] ||= P.def || "app"; });
  return `<div class="polbox"><b>月ごとの取り込み方</b><small>その人のその日に、アプリの予定（手で書いた行など）があるとき</small>
    <div class="polall">全部の月を <select id="polAll"><option value="">選ぶ…</option>${POLS.map(([v, t]) => `<option value="${v}">${t}</option>`).join("")}</select></div>
    <table class="list pol"><tr><th>月</th><th>件数</th><th>人</th><th>取り込み方</th></tr>${ms.map(m => `<tr><td>${+m.slice(0, 4)}年${+m.slice(5)}月</td><td class="num">${by.get(m).n}</td><td class="num">${by.get(m).p.size}</td>
      <td><select data-pol="${m}" title="${esc(POL_TIP[P.pol[m]])}">${POLS.map(([v, t]) => `<option value="${v}"${P.pol[m] === v ? " selected" : ""}>${t}</option>`).join("")}</select></td></tr>`).join("")}</table></div>`;
}
function bindPol(P, redraw) {
  const B = $("dlgBody");
  B.querySelectorAll("[data-pol]").forEach(el => el.onchange = () => { P.pol[el.dataset.pol] = el.value; el.title = POL_TIP[el.value]; });
  if ($("polAll")) $("polAll").onchange = () => { const v = $("polAll").value; if (!v) return; Object.keys(P.pol).forEach(m => { P.pol[m] = v; }); P.def = v; redraw(); };
}

// 同じ人・同じ日・同じ現場に、ちがう取り込み元（クロジカ／常用日報）の行があれば、その行にまとめる
const sameSite = (a, b) => (a.コード && b.コード ? String(a.コード) === String(b.コード) : nk(a.行先) === nk(b.行先)) && !a.関連現場 && !b.関連現場 && !isOffice(a.行先);
const srcOf = r => String(r.取込 || "").split(":")[0];
const idsOf = y => [y.取込, ...(y.取込他 || [])].filter(Boolean);
function mergeInto(y, r) {
  const uniq = (a, b, key) => { b.forEach(v => { if (!a.some(w => key(w) === key(v))) a.push(v); }); return a; };
  y.人 = uniq(listOf(y, "r", "人"), r.人, nk); y.機械 = uniq(listOf(y, "r", "機械"), r.機械, nk);
  if (srcOf(r) === "jouyou") { y.外注 = r.外注; if (!y.始め || y.終日) { y.始め = r.始め; y.終わり = r.終わり; y.終日 = false; } }
  else y.外注 = uniq(listOf(y, "r", "外注"), r.外注, g => nk(g.業者));
  const has = new Set(workLines(y).map(l => nk(l.予定) + "|" + nk(l.報告)));
  y.作業行 = [...workLines(y).filter(l => l.予定 || l.報告), ...workLines(r).filter(l => (l.予定 || l.報告) && !has.has(nk(l.予定) + "|" + nk(l.報告)))];
  (y.取込他 ||= []).push(r.取込); syncWork(y);
}

// 個別に確認：重なる人・日ごとに「アプリ／取り込み／両方」を選ぶ窓。決めたら Map(owner|k → app/imp/all)、やめたら null
function askEach(list, label) {
  return new Promise(res => {
    const pick = new Map(list.map(c => [c.key, "app"]));
    const tl = r => esc(titleOf(r) + (r.始め || r.終わり ? `  ${r.始め || ""}〜${r.終わり || ""}` : ""));
    dlg(`<div class="dh"><h2>重なる日を確かめる（${list.length}件）</h2><button type="button" data-close aria-label="閉じる">×</button></div>
      <p class="pinfo">その人のその日に、アプリの予定と${esc(label)}の両方があります。どちらを残すか選んでください。</p>
      <div class="btns" style="justify-content:flex-start;margin:0 0 6px">全部を <button type="button" data-allp="app">アプリ</button><button type="button" data-allp="imp">${esc(label)}</button><button type="button" data-allp="all">両方</button></div>
      <div class="gwrap kuro-wrap"><table class="list askt"><tr><th>人</th><th>日</th><th>アプリの予定</th><th>${esc(label)}</th><th>残す</th></tr>
      ${list.map((c, i) => `<tr><td>${esc(c.owner)}</td><td>${+c.k.slice(5, 7)}/${+c.k.slice(8)}(${WD[dow(c.k)]})</td><td>${c.app.map(tl).join("<br>")}</td><td>${c.imp.map(tl).join("<br>")}</td>
        <td class="ch3">${[["app", "アプリ"], ["imp", "取り込み"], ["all", "両方"]].map(([v, t]) => `<label class="ck"><input type="radio" name="ak${i}" value="${v}" data-ak="${i}"${v === "app" ? " checked" : ""}> ${t}</label>`).join("")}</td></tr>`).join("")}</table></div>
      <div class="btns"><button type="button" id="akNo">やめる</button><button class="pri" id="akOk">これで取り込む</button></div>`, "wide keep");
    const B = $("dlgBody");
    B.querySelectorAll("[data-ak]").forEach(el => el.onchange = () => pick.set(list[+el.dataset.ak].key, el.value));
    B.querySelectorAll("[data-allp]").forEach(b => b.onclick = () => { const v = b.dataset.allp; list.forEach(c => pick.set(c.key, v)); B.querySelectorAll(`[data-ak][value="${v}"]`).forEach(el => { el.checked = true; }); });
    $("akNo").onclick = () => { closeDlg(); res(null); };
    $("akOk").onclick = () => res(pick);
    B.querySelector("[data-close]").onclick = () => res(null);
  });
}

/* 取り込む。o = { srcs: ["kuro","jouyou"]（いま取り込んでいる元）, pol: { 月: app/imp/all/ask }, label: "クロジカ", prog: n => 進み具合を出す }
   返す：{ added, merged, skipped（もう入っていた）, kept（アプリを優先して入れなかった）, removed（取り込む方を優先して消したアプリの行）, files } か、やめたら null */
async function importItems(items, o) {
  const SR = new Set(o.srcs), files = new Map();
  items.forEach(it => { const key = it.owner + "|" + monthOf(it.k); if (!files.has(key)) files.set(key, { owner: it.owner, m: monthOf(it.k), list: [] }); files.get(key).list.push(it); });
  const name = F => `個人_${nk(F.owner)}_${F.m}.json`;
  // 1. いまの日報を読む（重なりを確かめるため）
  let i = 0;
  for (const F of files.values()) { o.prog && o.prog(`読んでいます… ${++i}/${files.size}`); F.cur = await nippouStore.get(name(F)); }
  const appRows = d => (d && d.行 || []).filter(y => !SR.has(srcOf(y)));
  // ほかの人の日報に参加者として入っている行も「その人のその日の予定」（クロジカは1人目の日報に入れ、ほかは参加者のため）。
  // 見つけるだけ：取り込む方を優先しても、ほかの人の日報は消さない
  const others = new Map();
  for (const m of new Set([...files.values()].map(F => F.m))) (await monthDocs(m, true).catch(() => [])).forEach(doc => Object.entries(doc.日 || {}).forEach(([k, d]) =>
    (d.行 || []).forEach(r => { if (SR.has(srcOf(r))) return; membersOf({ who: doc.名前, r }).forEach(n => { if (nk(n) === nk(doc.名前)) return; const key = nk(n) + "|" + k; if (!others.has(key)) others.set(key, []); others.get(key).push({ ...r, 行先: r.行先 + "（" + doc.名前 + " の日報）" }); }); })));
  const otherRows = (owner, k) => others.get(nk(owner) + "|" + k) || [];
  // 2. 個別に確認する月の重なり
  const asks = [];
  for (const F of files.values()) {
    if ((o.pol[F.m] || "app") !== "ask") continue;
    const by = new Map();
    F.list.forEach(({ k, r }) => { const d = F.cur && F.cur.data.日[k]; if (d && d.行.some(y => idsOf(y).includes(r.取込))) return; const a = [...appRows(d), ...otherRows(F.owner, k)]; if (!a.length) return;
      if (!by.has(k)) by.set(k, { key: F.owner + "|" + k, owner: F.owner, k, app: a, imp: [] }); by.get(k).imp.push(r); });
    asks.push(...by.values());
  }
  asks.sort((a, b) => a.k.localeCompare(b.k) || a.owner.localeCompare(b.owner, "ja"));
  let pick = new Map();
  if (asks.length) { pick = await askEach(asks, o.label); if (!pick) return null; dlg(`<h2>取り込んでいます</h2><p id="impProg">…</p>`, "keep"); }
  const decide = (owner, k) => { const p = o.pol[monthOf(k)] || "app"; return p === "ask" ? pick.get(owner + "|" + k) || "app" : p; };
  // 3. 書く（ほかの端末が先に保存していたら読み直してやり直す）
  const st = { added: 0, merged: 0, skipped: 0, kept: 0, removed: 0, files: 0 };
  i = 0;
  for (const F of files.values()) {
    o.prog && o.prog(`取り込んでいます… ${++i}/${files.size}`);
    for (let tries = 0; ; tries++) {
      const cur = tries ? await nippouStore.get(name(F)) : F.cur;
      const doc = cur ? cur.data : { 形式: "私の日報", 版: 1, 名前: F.owner, アカウント: "", 月: F.m, 日: {} };
      const s = { added: 0, merged: 0, skipped: 0, kept: 0, removed: 0 }, cleared = new Set();
      F.list.forEach(({ k, r: r0 }) => {
        const r = JSON.parse(JSON.stringify(r0)), d = doc.日[k] ||= { 行: [], 報告: null };
        if (d.行.some(y => idsOf(y).includes(r.取込))) { s.skipped++; return; }
        const a = appRows(d), how = a.length || otherRows(F.owner, k).length ? decide(F.owner, k) : "all";
        if (how === "app") { s.kept++; return; }
        if (how === "imp" && !cleared.has(k)) { cleared.add(k); s.removed += a.length; d.行 = d.行.filter(y => !a.includes(y)); if (!d.行.length) d.報告 = null; }
        const other = d.行.find(y => y.取込 && SR.has(srcOf(y)) && srcOf(y) !== srcOf(r) && sameSite(y, r));
        if (other) { if (srcOf(r) === "jouyou") { const keep = other; Object.assign(r, { id: keep.id }); mergeInto(r, keep); r.取込他 = [...new Set([...(r.取込他 || []).filter(v => v !== keep.取込), keep.取込, ...(keep.取込他 || [])])]; d.行[d.行.indexOf(keep)] = r; } else mergeInto(other, r); s.merged++; }
        else { d.行.push(r); s.added++; }
        if (!d.報告 && !r.予定のみ) d.報告 = { だれ: F.owner, いつ: (o.how ? o.how(r) : o.label) + "から取り込み" };
        delete r.予定のみ;
        d.更新 = nowStr();
      });
      if (!s.added && !s.merged && !s.removed) { st.skipped += s.skipped; st.kept += s.kept; break; }
      doc.更新 = nowStr();
      try { await nippouStore.put(name(F), doc, cur ? cur.eTag : null); Object.keys(s).forEach(x => st[x] += s[x]); st.files++; break; }
      catch (e) { if ((e.status !== 412 && e.status !== 409) || tries >= 4) throw e; }
    }
  }
  // 自分の分は読み直す（書きかけがなければ）
  Object.keys(DOCS).forEach(m => delete DOCS[m]);
  for (const m of new Set([...files.values()].filter(F => nk(F.owner) === nk(S.me.name)).map(F => F.m))) { const f = S.F[m]; if (f && !f.pending.length && !f.saving) f.loaded = false; await loadMonth(m, true); }
  return st;
}
// 進み具合：取り込むボタンの字か、個別に確認したあとの「取り込んでいます」の窓
const impProg = t => { const el = $("impProg") || $("kGo") || $("hGo"); if (el) el.textContent = t; };
function importDone(st, extra) {
  dlg(`<h2>取り込みました</h2><p>${st.added}件を入れました${st.merged ? `（ほかに ${st.merged}件は、同じ人・同じ日・同じ現場の行にまとめました）` : ""}。人・月ごとに ${st.files} のファイルを書きました。
    ${st.skipped ? `<br>${st.skipped}件はもう入っていたので入れていません。` : ""}${st.kept ? `<br>${st.kept}件は、その日にアプリの予定があったので入れていません（アプリを優先）。` : ""}${st.removed ? `<br>取り込む方を優先して、アプリの予定を ${st.removed}件消しました。` : ""}${extra || ""}</p>
    <p class="pinfo">カレンダーの週・月・人別で確かめてください。取り込んだ行は普通の日報と同じに直せます。</p><div class="btns"><button class="pri" data-close>閉じる</button></div>`);
}
