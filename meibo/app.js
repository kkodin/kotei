/* 名簿（2026-11-16 開発環境の整理 4.2：人の名簿を1本にして、確かめる・直す・足す・退職を入れる）
   読む：工程データ\マスタ\人.json（毎朝 01_マスタ抽出\人を作る.py が作る。直さない）
         工程データ\マスタ\人_手直し.json（この画面が書く）・工程データ\退職者.json（前からある退職の一覧）
   書く：人_手直し.json だけ（ほかの人と同時に直しても消し合わないように、書く直前に読み直して1人分だけ入れる・eTag）。
         最後の月（退職）を直したときは、移る間だけ 退職者.json にも同じものを入れる（工事日報がまだそちらを読むため）。
   合わせ方は共通の部品（common/master.js の HX.people）。どのアプリも同じ結果になる。
   この画面には人の名前を書かない（公開の置き場に置くため）。名前はサインインしてから OneDrive から読む。 */
"use strict";
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const DEMO = HX.demo;
const S = { base: null, fix: null, fixTag: null, ret: null, retTag: null, list: [], f: "all", q: "", me: "" };
let app = null, mst = null, root = null;

function toast(t, ms) { const el = $("toast"); el.textContent = t; el.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => el.hidden = true, ms || 3500); }
const ym = d => d.toISOString().slice(0, 7);
const nowStr = () => { const d = new Date(), p = n => String(n).padStart(2, "0"); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; };

/* ---------- 試し（?demo）：ブラウザの中だけ。名前は作りもの ---------- */
function seedDemo() {
  const KEY = "meibo_demo_store";
  try { if (localStorage.getItem(KEY)) return; } catch (e) { return; }
  const P = (名前, ふりがな, 区分, 会社, 職種, 元) => ({ 名前, ふりがな, 区分, 会社, 職種, 元 });
  const base = { 形式: "人", 作成: nowStr(), 人: [
    P("試し　一郎", "ためし　いちろう", "社員", "", "現場監督(土木)", ["配置表", "作業員名簿"]),
    P("試し　二郎", "ためし　じろう", "社員", "", "現場監督(土木)", ["配置表", "作業員名簿"]),
    P("見本　花子", "みほん　はなこ", "従業員", "", "土工", ["配置表", "作業員名簿"]),
    P("見本　三太", "", "従業員", "", "", ["配置表"]),
    P("例題　太郎", "れいだい　たろう", "外注", "Ａ工業", "土工", ["作業員名簿"]),
    P("例題　次郎", "れいだい　じろう", "外注", "Ｂ建設", "オペ", ["作業員名簿"])
  ], 違い: [{ 名前: "見本　三太", 内容: "作業員名簿にいない（ふりがな・職種が無い）" }] };
  const a = { "マスタ/人.json": { data: base, eTag: "d1" }, "退職者.json": { data: { 形式: "退職者", 人: [{ 名前: "試し　二郎", 最後の月: "2026-09" }] }, eTag: "d2" } };
  try { localStorage.setItem(KEY, JSON.stringify(a)); } catch (e) { }
}

/* ---------- 読む ---------- */
async function load() {
  const [b, f, r] = await Promise.all([mst.get("人.json"), mst.get("人_手直し.json"), root.get("退職者.json").catch(() => null)]);
  S.base = b ? b.data : null; S.fix = f ? f.data : { 形式: "人_手直し", 版: 1, 人: {} }; S.fixTag = f ? f.eTag : null;
  S.ret = r ? r.data : { 形式: "退職者", 版: 1, 人: [] }; S.retTag = r ? r.eTag : null;
  S.list = HX.people(S.base, S.fix, S.ret);
  render();
}

/* ---------- 見せる ---------- */
function shown() {
  const q = HX.nameKey(S.q).toLowerCase(), now = ym(new Date());
  return S.list.filter(p => {
    const out = !HX.activeOn(p, now);
    if (S.f === "out") { if (!out) return false; }
    else if (S.f === "fix") { if (!(p.直した.length || p.追加)) return false; }
    else if (S.f !== "all" && p.区分 !== S.f) return false;
    if (q && ![p.名前, p.ふりがな, p.会社, p.職種].some(v => HX.nameKey(v).toLowerCase().includes(q))) return false;
    return true;
  });
}
function render() {
  if (!S.base) { $("info").textContent = "人.json がまだありません（01_マスタ抽出\\人を作る.py で作って、工程データ\\マスタ に置く）。足した人だけ出します。"; }
  const diffs = (S.base && S.base.違い) || [];
  $("diffs").hidden = !diffs.length;
  $("diffs").innerHTML = diffs.length ? `<b>2つの名簿（配置表の名簿・作業員名簿）に違いがあります：${diffs.length}件</b>　名前を押して確かめ、必要なら直してください。<ul>${diffs.map(d => `<li><a href="#" data-n="${esc(d.名前)}">${esc(d.名前)}</a>：${esc(d.内容)}</li>`).join("")}</ul>` : "";
  const rows = shown(), now = ym(new Date()), dk = new Set(diffs.map(d => HX.nameKey(d.名前)));
  if (S.base) $("info").textContent = `${rows.length}人を表示（全部で ${S.list.length}人）。名簿の元：${S.base.作成 || ""} に作ったもの。行を押すと直せます。`;
  $("tbl").tBodies[0].innerHTML = rows.map(p => {
    const out = !HX.activeOn(p, now);
    const tags = (p.追加 ? `<span class="tag add">足した</span>` : "") + (p.直した.filter(c => c !== "外す" && c !== "最後の月").length ? `<span class="tag fix" title="${esc(p.直した.join("・"))}">直した</span>` : "") + (dk.has(HX.nameKey(p.名前)) ? `<span class="tag dif">違い</span>` : "") + (p.外す ? `<span class="tag">外した</span>` : "");
    return `<tr data-k="${esc(HX.nameKey(p.名前))}" class="${out ? "out" : ""}"><td>${esc(p.名前)}${tags}</td><td>${esc(p.ふりがな)}</td><td>${esc(p.区分)}</td><td>${esc(p.会社)}</td><td>${esc(p.職種)}</td><td class="last">${esc(p.最後の月 || "")}</td><td>${esc((p.元 || []).join("・"))}</td></tr>`;
  }).join("") || `<tr><td colspan="7">当てはまる人はいません</td></tr>`;
}

/* ---------- 直す・足す ---------- */
function baseOf(k) { return ((S.base && S.base.人) || []).find(p => HX.nameKey(p.名前) === k) || null; }
function edit(k) {
  const p = k ? S.list.find(x => HX.nameKey(x.名前) === k) : null, b = k ? baseOf(k) : null, isNew = !p;
  const v = p || { 名前: "", ふりがな: "", 区分: "外注", 会社: "", 職種: "", 最後の月: "" };
  const opt = (a, cur) => a.map(x => `<option${x === cur ? " selected" : ""}>${x}</option>`).join("");
  const diff = ((S.base && S.base.違い) || []).filter(d => HX.nameKey(d.名前) === k).map(d => esc(d.内容)).join("<br>");
  $("dlgBody").innerHTML = `<h2>${isNew ? "人を足す" : esc(v.名前)}</h2>
    ${b ? `<div class="src">名簿から作った値：${esc(b.区分)}${b.会社 ? "・" + esc(b.会社) : ""}${b.職種 ? "・" + esc(b.職種) : ""}${b.ふりがな ? "・" + esc(b.ふりがな) : ""}（元：${esc((b.元 || []).join("・"))}）${diff ? `<br><b>違い：</b>${diff}` : ""}</div>` : ""}
    ${p && p.追加 ? `<div class="src">この画面で足した人です。</div>` : ""}
    <label>名前（姓と名の間は全角の空白）<input name="名前" value="${esc(v.名前)}" ${isNew ? "required" : ""} placeholder="例：山田　太郎"></label>
    <div class="row"><label>ふりがな<input name="ふりがな" value="${esc(v.ふりがな)}"></label>
      <label>区分<select name="区分">${opt(["社員", "従業員", "外注"], v.区分)}</select></label></div>
    <div class="row"><label>会社（外注のとき）<input name="会社" value="${esc(v.会社)}"></label>
      <label>職種<input name="職種" value="${esc(v.職種)}"></label></div>
    <div class="row"><label>最後の月（辞めた人。次の月から候補に出さない。前の記録はそのまま）<input name="最後の月" type="month" value="${esc(v.最後の月 || "")}"></label>
      <label>メモ<input name="メモ" value="${esc(p && p.メモ || "")}"></label></div>
    ${isNew ? "" : `<label class="chk"><input type="checkbox" name="外す" ${p.外す ? "checked" : ""}> 名簿から外す（間違って入った人など。どの月にも出さない）</label>`}
    ${p && p.更新 ? `<div class="src">最後に直した：${esc(p.更新)}　${esc(p.更新者 || "")}</div>` : ""}
    <div class="btns">${!isNew && S.fix.人[k] ? `<button type="button" class="left" data-act="undo">直したのをやめる（名簿のとおりに戻す）</button>` : ""}
      <button type="button" data-act="close">閉じる</button><button type="submit" class="pri" data-act="save">${isNew ? "足す" : "直す"}</button></div>`;
  const fm = $("dlgBody");
  fm.onclick = async e => {
    const a = e.target.closest("[data-act]"); if (!a) return;
    if (a.dataset.act === "close") { $("dlg").close(); return; }
    if (a.dataset.act === "undo") { e.preventDefault(); if (!confirm("この人の直しをやめて、名簿のとおりに戻しますか？")) return; await saveOne(k, null, null); $("dlg").close(); }
  };
  fm.onsubmit = async e => {
    e.preventDefault();
    const fd = new FormData(fm), val = c => String(fd.get(c) || "").trim();
    const name = val("名前").replace(/\s+/g, "　");
    if (isNew) {
      if (!/^\S+　\S+$/.test(name)) { toast("名前は「姓　名」の形で入れてください（間に空白）"); return; }
      if (S.list.some(x => HX.nameKey(x.名前) === HX.nameKey(name))) { toast("同じ名前の人がもういます"); return; }
    }
    const nk = isNew ? HX.nameKey(name) : k, ref = b || {};
    const ent = { ...(S.fix.人[nk] || {}) };
    for (const c of ["名前", "ふりがな", "区分", "会社", "職種"]) { const x = c === "名前" ? name : val(c); if (isNew || x !== (ref[c] || "")) ent[c] = x; else delete ent[c]; }
    const last = val("最後の月"), lastChanged = last !== ((p && p.最後の月) || "");
    if (lastChanged) ent.最後の月 = last;
    if (fd.get("外す")) ent.外す = true; else delete ent.外す;
    if (val("メモ")) ent.メモ = val("メモ"); else delete ent.メモ;
    if (isNew) ent.追加 = true;
    const keys = Object.keys(ent).filter(c => !["更新", "更新者"].includes(c));
    await saveOne(nk, keys.length ? ent : null, lastChanged ? { name, last } : null);
    $("dlg").close();
  };
  $("dlg").showModal();
}
// 1人分を書く：書く直前に読み直して、その人の分だけ入れる（ほかの人が同時に直した分を消さない）
async function saveOne(k, ent, mirror) {
  for (let i = 0; i < 3; i++) {
    const cur = await mst.get("人_手直し.json"), data = cur ? cur.data : { 形式: "人_手直し", 版: 1, 説明: "名簿の画面（meibo/）が書く。人.json への直し・足した人・最後の月（退職）", 人: {} };
    data.人 = data.人 || {};
    if (ent) data.人[k] = { ...ent, 更新: nowStr(), 更新者: S.me }; else delete data.人[k];
    try { await mst.put("人_手直し.json", data, cur ? cur.eTag : null); break; }
    catch (e) { if ((e.status === 412 || e.status === 409) && i < 2) continue; toast("保存できませんでした：" + e.message, 6000); throw e; }
  }
  // 移る間だけ：最後の月を 退職者.json にも入れる（工事日報がまだそちらを読む）
  if (mirror) await mirrorRetire(mirror.name, mirror.last).catch(e => toast("退職者.json に入れられませんでした：" + e.message, 6000));
  toast(ent ? "直しました（ほかの端末・アプリにも届きます）" : "名簿のとおりに戻しました（最後の月は 退職者.json の値が残ります）", 5000);
  await load();
}
async function mirrorRetire(name, last) {
  for (let i = 0; i < 3; i++) {
    const cur = await root.get("退職者.json"), data = cur ? cur.data : { 形式: "退職者", 版: 1, 人: [] };
    data.人 = (data.人 || []).filter(r => HX.nameKey(r.名前) !== HX.nameKey(name));
    if (last) data.人.push({ 名前: name, 最後の月: last, メモ: `${nowStr()} 名簿の画面で ${S.me}` });
    try { await root.put("退職者.json", data, cur ? cur.eTag : null); return; }
    catch (e) { if ((e.status === 412 || e.status === 409) && i < 2) continue; throw e; }
  }
}

/* ---------- はじめ ---------- */
async function start() {
  if (DEMO) {
    seedDemo(); S.me = "試し";
    $("who").textContent = "試し（このブラウザの中だけ。OneDrive には書きません）";
    mst = HX.demoStore("meibo_demo_store", "マスタ/"); root = HX.demoStore("meibo_demo_store", "");
  } else {
    if (location.protocol === "file:") { $("signin").hidden = false; $("bSignIn").hidden = true; $("signinMsg").textContent = "パソコンのファイルを直接開いているため、サインインできません。公開先から開いてください。"; return; }
    const s = await HX.signIn(); app = s.app;
    if (s.error) $("signinMsg").textContent = "サインインできませんでした：" + s.error.message;
    if (!s.account) { $("signin").hidden = false; $("bSignIn").onclick = () => HX.login(app); return; }
    S.me = s.account.name || s.account.username; $("who").textContent = S.me;
    mst = HX.store(app, "kotei", "マスタ"); root = HX.store(app, "kotei");
  }
  $("app").hidden = false;
  $("tabs").onclick = e => { const b = e.target.closest("[data-f]"); if (!b) return; S.f = b.dataset.f; $("tabs").querySelectorAll("button").forEach(x => x.classList.toggle("on", x === b)); render(); };
  $("q").oninput = e => { S.q = e.target.value; render(); };
  $("bAdd").onclick = () => edit(null);
  $("tbl").tBodies[0].onclick = e => { const tr = e.target.closest("tr[data-k]"); if (tr) edit(tr.dataset.k); };
  $("diffs").onclick = e => { const a = e.target.closest("[data-n]"); if (!a) return; e.preventDefault(); edit(HX.nameKey(a.dataset.n)); };
  try { await load(); } catch (e) { toast("名簿を読めませんでした：" + e.message, 8000); }
}
start();
