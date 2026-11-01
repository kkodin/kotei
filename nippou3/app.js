/* 工事日報 Vr003（2026-11-01 作り直し）
   会社の標準の日報。サインインした人＝書く人が、毎日「どこで・何をしたか」を書く（私の日報）。
   ・1日を「行った所」のカードで書く：行先（現場・本社など）・始め〜終わり・作業内容・来ていた人（作業員の名前／外注は業者名と人数）。
     本社だけの日は1枚、いくつも回った日はカードを足す。詳しく書かない人はここまで。
   ・現場のカードから「詳しく書く」で常用日報（kotei/jouyou）を開き、この日・この現場に作業内容と人を入れた形で詳しく書く。
   ・画面は PC・iPad・iPhone SE で1つ（幅で並びが変わる）。ボタンは下の3つ＋「…」だけ。ホーム画面に置くとアドレスバーが出ない。
   ・データ：会社の OneDrive の 日報のフォルダ（config の nippou）に 個人_<名前>_<YYYY-MM>.json（1人1か月1ファイル）。マスタは読むだけ。
     365 のアカウントが無い人はゲストに招き、日報のフォルダだけを共有する（工程データは見せない）。ゲストは同じフォルダの 名簿.json
     （工程データのマスタ・現場一覧の名前だけの写し。工程データを読める人が開くたびに作り直す）で名前・現場を選ぶ。
     config に nippou が無いときは、これまでどおり 工程データ\日報\ に書く。 */
"use strict";
const $ = id => document.getElementById(id);
const CFG = window.NIPPOU_CONFIG;
const SCOPES = ["User.Read", "Files.ReadWrite.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const DEMO = new URLSearchParams(location.search).has("demo");
const NDIR = "日報";
const SEP = DEMO || !!(CFG.nippou && CFG.nippou.itemId);   // 日報だけ別のフォルダ（ゲストに共有する）
const ROSTER = "名簿.json";                                  // そのフォルダに置く名簿・現場一覧の写し
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const pad2 = n => String(n).padStart(2, "0");
const keyOf = d => d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
const dateOf = k => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
const dayAdd = (k, n) => { const d = dateOf(k); d.setDate(d.getDate() + n); return keyOf(d); };
const todayKey = () => keyOf(new Date());
const WD = "日月火水木金土";
const monthOf = k => k.slice(0, 7);
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const nowStr = () => new Date().toLocaleString("ja-JP");
// 名前を比べる形（空白・全角半角・株式会社などのちがいを無視）
const nk = s => String(s ?? "").normalize("NFKC").replace(/株式会社|\(株\)|有限会社|\(有\)/g, "").replace(/[\s　・]/g, "").toLowerCase();
const OFFICE = ["本社", "休み"];      // 現場ではない行先（人・外注の欄を出さない）
const isOffice = v => OFFICE.includes(String(v || "").trim()) || /^本社/.test(String(v || "").trim());

/* ブラウザの自動入力（住所・会社名・氏名の候補）を出さない（2026-11-01 ユーザー：表のセルに Chrome の住所の候補が出た）。
   Chrome は autocomplete="off" を無視して住所の候補を出すことがある。1つの form の中の欄が3つ未満だと住所の自動入力をしないので、
   form の外の欄を1つずつ form（display:contents。見た目は変わらない）で包む。あとから足された欄も包む（いま打っている欄なら打つ所を戻す） */
function noAutofill(root) {
  if (!root || !root.querySelectorAll) return;
  const list = root.matches && root.matches("input,textarea,select") ? [root] : root.querySelectorAll("input,textarea,select");
  list.forEach(el => {
    if (el.closest("form") || /^(checkbox|radio|file|button|submit|hidden|range|color)$/.test(el.type)) return;
    const on = document.activeElement === el, s0 = on && el.selectionStart, s1 = on && el.selectionEnd;
    const f = document.createElement("form"); f.className = "nofill"; f.setAttribute("autocomplete", "off");
    f.addEventListener("submit", e => e.preventDefault());
    el.parentNode.insertBefore(f, el); f.appendChild(el); el.setAttribute("autocomplete", "off");
    if (on) { el.focus({ preventScroll: true }); try { el.setSelectionRange(s0, s1); } catch (e) { } }
  });
}
noAutofill(document.body);
new MutationObserver(ms => ms.forEach(m => m.addedNodes.forEach(n => { if (n.nodeType === 1 && n.parentNode) noAutofill(n); }))).observe(document.body, { childList: true, subtree: true });

/* ---------- OneDrive（全体工程表・Vr002 と同じ読み書き） ---------- */
let msalApp = null;
async function token() {
  try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken; }
  catch (e) { if (e instanceof msal.InteractionRequiredAuthError) await msalApp.acquireTokenRedirect({ scopes: SCOPES }); throw e; }
}
async function g(path, opt = {}) { return fetch(GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + await token(), ...(opt.headers || {}) } }); }
function httpErr(res, what) { const e = new Error(what + " " + res.status); e.status = res.status; return e; }
const enc = p => p.split("/").map(encodeURIComponent).join("/");
function GraphFolder(where, sub) {
  const F = CFG[where], base = () => `/drives/${F.driveId}/items/${F.itemId}`, p = name => base() + ":/" + enc((sub ? sub + "/" : "") + name);
  return {
    async list() {
      const out = []; let url = sub ? base() + ":/" + enc(sub) + ":/children?$top=999" : base() + "/children?$top=999";
      while (url) {
        const res = await g(url);
        if (res.status === 404 && sub) return [];
        if (!res.ok) throw httpErr(res, "一覧");
        const j = await res.json();
        j.value.forEach(x => { if (x.file) out.push({ name: x.name, eTag: x.eTag }); });
        url = j["@odata.nextLink"] ? j["@odata.nextLink"].replace(GRAPH, "") : null;
      }
      return out;
    },
    async get(name) {
      const res = await g(p(name));
      if (res.status === 404) return null;
      if (!res.ok) throw httpErr(res, "読み込み");
      const meta = await res.json(), dl = meta["@microsoft.graph.downloadUrl"];
      const body = dl ? await fetch(dl, { cache: "no-store" }) : await g(p(name) + ":/content");
      if (!body.ok) throw httpErr(body, "中身の読み込み");
      return { data: await body.json(), eTag: meta.eTag };
    },
    async put(name, data, eTag) {
      const url = p(name) + ":/content" + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
      const send = () => g(url, { method: "PUT", headers: { "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: JSON.stringify(data) });
      let res = await send();
      if (res.status === 404 && sub && !eTag) {   // 日報 フォルダがまだ無い → 作ってからもう一度
        await g(base() + "/children", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: sub, folder: {}, "@microsoft.graph.conflictBehavior": "replace" }) });
        res = await send();
      }
      if (!res.ok) throw httpErr(res, "保存");
      return { eTag: (await res.json()).eTag };
    }
  };
}
const DEMO_KEY = "nippou_demo_v3";
function DemoFolder(prefix) {
  const all = () => { try { return JSON.parse(localStorage.getItem(DEMO_KEY)) || {}; } catch (e) { return {}; } };
  const save = a => { try { localStorage.setItem(DEMO_KEY, JSON.stringify(a)); } catch (e) { } };
  return {
    async list() { const a = all(); return Object.keys(a).filter(n => n.startsWith(prefix)).map(n => ({ name: n.slice(prefix.length), eTag: a[n].eTag })); },
    async get(name) { const x = all()[prefix + name]; return x ? { data: JSON.parse(JSON.stringify(x.data)), eTag: x.eTag } : null; },
    async put(name, data, eTag) {
      const a = all(), cur = a[prefix + name];
      if (eTag ? (!cur || cur.eTag !== eTag) : cur) { const e = new Error("保存 " + (eTag ? 412 : 409)); e.status = eTag ? 412 : 409; throw e; }
      const t = "d" + Date.now() + Math.random().toString(36).slice(2, 6);
      a[prefix + name] = { data: JSON.parse(JSON.stringify(data)), eTag: t }; save(a);
      return { eTag: t };
    }
  };
}
const GUEST_DEMO = DEMO && new URLSearchParams(location.search).has("guest");   // 試し：工程データを読めないゲストのつもり
const koteiStore = DEMO ? DemoFolder(GUEST_DEMO ? "見えない/" : "工程/") : GraphFolder("kotei");
const nippouStore = DEMO ? DemoFolder("日報/") : SEP ? GraphFolder("nippou") : GraphFolder("kotei", NDIR);
const kikaiStore = DEMO ? DemoFolder(GUEST_DEMO ? "見えない/" : "機械/") : GraphFolder("kikai");

/* ---------- 状態 ---------- */
const S = {
  me: { name: "", acct: "" },   // name＝名簿の名前（書く人）。acct＝サインインしたアカウント
  day: todayKey(),
  F: {},                         // 月 → { doc, eTag, pending:[fn], saving, timer, err, loaded }
  M: { 社員: [], 従業員: [], 外注業者: [] },   // 工程データ\マスタ.json
  sites: [],                     // 稼働データ\現場一覧.json の 現場（コード・名前・正式名・元請）
  view: "me"                     // me＝私の日報、all＝この日のみんなの日報
};
const fileName = m => `個人_${nk(S.me.name)}_${m}.json`;
function blankDoc(m) { return { 形式: "私の日報", 版: 1, 名前: S.me.name, アカウント: S.me.acct, 月: m, 日: {} }; }
const fileOf = m => S.F[m] || (S.F[m] = { doc: blankDoc(m), eTag: null, pending: [], saving: false, timer: null, err: null, loaded: false });
const dayOf = (k = S.day) => { const f = S.F[monthOf(k)]; return f && f.doc.日[k] || null; };
const rowsOf = (k = S.day) => (dayOf(k) || {}).行 || [];

// 変える：fn をいまの中身に当て、保存する。ほかの端末が先に保存していたら、読み直して fn を当て直す
function mutate(k, fn, quiet) {
  const f = fileOf(monthOf(k));
  const wrap = doc => { if (!doc.日[k]) doc.日[k] = { 行: [], 報告: null }; fn(doc.日[k], doc); const d = doc.日[k]; d.更新 = nowStr(); if (d.報告 && !wrap.noMark) d.報告.直した = nowStr(); doc.更新 = nowStr(); };
  wrap.noMark = !!(quiet && quiet.report);
  wrap(f.doc); f.pending.push(wrap);
  clearTimeout(f.timer); f.timer = setTimeout(() => saveMonth(monthOf(k)), 700);
  showSync();
  if (!quiet || quiet.render) render();
}
async function saveMonth(m) {
  const f = S.F[m];
  if (f.saving) { clearTimeout(f.timer); f.timer = setTimeout(() => saveMonth(m), 500); return; }
  if (!f.pending.length) return;
  f.saving = true; showSync();
  try {
    for (let tries = 0; tries < 5; tries++) {
      const n = f.pending.length;
      try { const r = await nippouStore.put(fileName(m), f.doc, f.eTag); f.eTag = r.eTag; f.pending.splice(0, n); f.err = null; break; }
      catch (e) {
        if (e.status !== 412 && e.status !== 409) throw e;
        const fresh = await nippouStore.get(fileName(m)), d = fresh ? fresh.data : blankDoc(m);
        f.pending.forEach(fn => fn(d)); f.doc = d; f.eTag = fresh ? fresh.eTag : null;
        toast("ほかの端末で書いた分を読み込んで、重ねて保存しました"); render();
      }
    }
  } catch (e) {
    f.err = e.message; console.error(e);
    toast("保存できませんでした：" + e.message + "（自動でやり直します）");
    setTimeout(() => { f.err = null; saveMonth(m); }, 10000);
  } finally {
    f.saving = false; showSync();
    if (f.pending.length && !f.err) { clearTimeout(f.timer); f.timer = setTimeout(() => saveMonth(m), 300); }
  }
}
function showSync() {
  const fs = Object.values(S.F), el = $("sync");
  const err = fs.some(f => f.err), busy = fs.some(f => f.pending.length || f.saving);
  el.className = "sync " + (err ? "err" : busy ? "busy" : "ok");
  el.textContent = err ? "保存できていません" : busy ? "保存中…" : (fs.some(f => f.eTag) ? "保存済み" : "");
}
async function loadMonth(m, force) {
  const f = fileOf(m);
  if (f.loaded && !force) return;
  if (f.pending.length || f.saving) return;   // 書きかけを上書きしない
  try { const r = await nippouStore.get(fileName(m)); if (r && !f.pending.length) { f.doc = r.data; f.eTag = r.eTag; } f.loaded = true; }
  catch (e) { toast("日報を読めませんでした：" + e.message); }
}

/* ---------- マスタ（読むだけ）。日報が別のフォルダなら、そこに名前だけの写し（名簿.json）を置き、ゲストはそれを読む ---------- */
const siteMin = x => ({ コード: x.コード || "", 名前: x.名前 || "", 正式名: x.正式名 || "", 元請: x.元請 || "" });
async function loadMasters() {
  const [m, s] = await Promise.all([koteiStore.get("マスタ.json").catch(() => null), kikaiStore.get("現場一覧.json").catch(() => null)]);
  if (m) S.M = { 社員: m.data.社員 || [], 従業員: m.data.従業員 || [], 外注業者: m.data.外注業者 || [] };
  if (s) S.sites = (s.data.現場 || []).filter(x => x && (x.名前 || x.正式名));
  if (SEP) {
    const cur = await nippouStore.get(ROSTER).catch(() => null);
    if (m) await saveRoster(cur, !!s);
    else if (cur) { const c = cur.data; S.M = { 社員: c.社員 || [], 従業員: c.従業員 || [], 外注業者: c.外注業者 || [] }; S.sites = (c.現場 || []).filter(x => x && (x.名前 || x.正式名)); }
    else if (!m) throw new Error("名簿がまだありません（工程データを読める人が一度開くと作られます）");
  }
  try { localStorage.setItem("nippou_master_v3", JSON.stringify({ M: S.M, sites: S.sites })); } catch (e) { }
}
// 工程データを読めた人が開いたとき：名簿.json を今のマスタに合わせる（変わっていなければ書かない）。現場一覧を読めなかったら現場は前の写しのまま
async function saveRoster(cur, gotSites) {
  const c = cur ? cur.data : {};
  const body = { 社員: S.M.社員, 従業員: S.M.従業員, 外注業者: S.M.外注業者, 現場: gotSites ? S.sites.map(siteMin) : (c.現場 || []) };
  if (cur && JSON.stringify(body) === JSON.stringify({ 社員: c.社員, 従業員: c.従業員, 外注業者: c.外注業者, 現場: c.現場 })) return;
  const doc = { 形式: "工事日報の名簿", 版: 1, 説明: "工程データのマスタ・現場一覧の名前だけの写し。工事日報が自動で作り直す（ここを直しても戻る）", ...body, 更新: nowStr() };
  try { await nippouStore.put(ROSTER, doc, cur ? cur.eTag : null); } catch (e) { console.warn("名簿の写しを保存できませんでした", e); }
}
function cachedMasters() { try { const c = JSON.parse(localStorage.getItem("nippou_master_v3")); if (c) { S.M = c.M; S.sites = c.sites; } } catch (e) { } }
const people = () => [...S.M.社員, ...S.M.従業員];
const siteName = x => x.名前 || x.正式名;

/* ---------- 選ぶ欄（打てば絞り込み、押せば選ぶ。スマホでも押しやすい大きさ） ---------- */
let sugBox = null;
function closeSug() { if (sugBox) { sugBox.remove(); sugBox = null; } }
function showSug(input, items, onPick) {
  closeSug();
  const q = nk(input.value), hit = items.filter(x => !q || nk(x.v).includes(q) || nk(x.sub || "").includes(q)).slice(0, 40);
  if (!hit.length) return;
  sugBox = document.createElement("div"); sugBox.className = "sug";
  let grp = null;
  hit.forEach(x => {
    if (x.grp && x.grp !== grp) { grp = x.grp; const h = document.createElement("div"); h.className = "grp"; h.textContent = grp; sugBox.appendChild(h); }
    const b = document.createElement("button"); b.type = "button";
    b.innerHTML = esc(x.v) + (x.sub ? `<small>${esc(x.sub)}</small>` : "");
    b.addEventListener("pointerdown", e => e.preventDefault());   // 押したときに打つ欄から外れない
    b.addEventListener("click", () => { closeSug(); onPick(x); });
    sugBox.appendChild(b);
  });
  input.parentElement.appendChild(sugBox);
}
function bindPick(input, items, onPick) {
  const open = () => showSug(input, items(), onPick);
  input.addEventListener("focus", open);
  input.addEventListener("input", open);
  input.addEventListener("blur", () => setTimeout(closeSug, 150));
  input.addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); const b = sugBox && sugBox.querySelector("button"); if (b && input.value.trim()) b.click(); else { closeSug(); onPick({ v: input.value.trim(), free: true }); } }
    if (e.key === "Escape") closeSug();
  });
}
// 行先の候補：最近行った所（30日）→ 本社・休み → 全部の現場（現場一覧）
function placeItems() {
  const recent = [], seen = new Set();
  for (let i = 0; i <= 30; i++) rowsOf(dayAdd(S.day, -i)).forEach(r => { const k = nk(r.行先); if (r.行先 && !seen.has(k)) { seen.add(k); recent.push({ grp: "最近行った所", v: r.行先, コード: r.コード || "" }); } });
  return [...recent, ...OFFICE.filter(v => !seen.has(nk(v))).map(v => ({ grp: "本社など", v })),
    ...S.sites.filter(x => !seen.has(nk(siteName(x)))).map(x => ({ grp: "現場", v: siteName(x), sub: [x.コード, x.元請].filter(Boolean).join(" "), コード: x.コード || "" }))];
}
// 人の候補：この行先に前に来ていた人 → 社員・従業員
function personItems(place, have) {
  const hv = new Set(have.map(nk)), out = [], seen = new Set();
  for (let i = 1; i <= 30; i++) rowsOf(dayAdd(S.day, -i)).forEach(r => { if (nk(r.行先) === nk(place)) (r.人 || []).forEach(n => { const k = nk(n); if (!hv.has(k) && !seen.has(k)) { seen.add(k); out.push({ grp: "前にこの現場に来ていた人", v: n }); } }); });
  S.M.社員.forEach(n => { const k = nk(n); if (!hv.has(k) && !seen.has(k)) { seen.add(k); out.push({ grp: "社員", v: n }); } });
  S.M.従業員.forEach(n => { const k = nk(n); if (!hv.has(k) && !seen.has(k)) { seen.add(k); out.push({ grp: "従業員", v: n }); } });
  return out;
}
const gaiItems = () => S.M.外注業者.map(v => ({ v }));

/* ---------- 時刻（「8」→8:00、「830」→8:30、「21.5」→21:30） ---------- */
function normTime(v) {
  v = String(v || "").normalize("NFKC").replace(/[時]/g, ":").replace(/分/g, "").trim();
  if (!v) return "";
  let m = /^(\d{1,2})[:.](\d{1,2})$/.exec(v);
  if (m && v.includes(".") && m[2].length === 1) return `${+m[1]}:${pad2(Math.round(+("0." + m[2]) * 60))}`;
  if (m) return `${+m[1]}:${pad2(+m[2])}`;
  if (/^\d{1,2}$/.test(v)) return `${+v}:00`;
  if (/^\d{3,4}$/.test(v)) return `${+v.slice(0, -2)}:${v.slice(-2)}`;
  return v;
}

/* ---------- 画面 ---------- */
function render() {
  const k = S.day, d = dateOf(k);
  $("dLabel").textContent = `${d.getMonth() + 1}/${d.getDate()}(${WD[d.getDay()]})`;
  $("inDay").value = k;
  $("bToday").hidden = k === todayKey();
  $("bMe").textContent = "👤 " + (S.me.name || "名前");
  if (S.view === "all") return renderAll();
  $("all").hidden = true; $("cards").hidden = false;
  const day = dayOf(), rows = rowsOf();
  const st = day && day.報告 ? (day.報告.直した ? `<span class="badge done">報告済み（あとで直した）</span>` : `<span class="badge done">報告済み</span>`) : rows.length ? `<span class="badge draft">書きかけ</span>` : `<span class="badge none">まだ</span>`;
  $("dayHead").innerHTML = st + (rows.length ? `<span>${rows.length} か所</span>` : "");
  $("empty").hidden = rows.length > 0;
  const focus = document.activeElement && document.activeElement.dataset.f ? { id: document.activeElement.closest(".card")?.dataset.id, f: document.activeElement.dataset.f } : null;
  $("cards").innerHTML = rows.map((r, i) => cardHTML(r, i)).join("");
  rows.forEach(r => bindCard(r.id));
  document.querySelectorAll("#cards textarea").forEach(grow);
  if (focus) { const el = document.querySelector(`.card[data-id="${focus.id}"] [data-f="${focus.f}"]`); if (el) el.focus(); }
  const done = day && day.報告 && !day.報告.直した;
  $("bReport").textContent = done ? "報告済み ✓" : day && day.報告 ? "もう一度報告" : "報告する";
  $("bReport").classList.toggle("done", !!done);
  $("bReport").disabled = !rows.length;
}
function cardHTML(r, i) {
  const off = isOffice(r.行先);
  const chips = (r.人 || []).map((n, j) => `<span class="chip">${esc(n)}<button type="button" data-delp="${j}" aria-label="${esc(n)} を外す">×</button></span>`).join("");
  const gai = (r.外注 || []).map((x, j) => `<div class="gai"><div class="pick"><input data-f="gai${j}" data-gi="${j}" value="${esc(x.業者)}" placeholder="業者名" autocomplete="off"></div><input class="n" data-gn="${j}" value="${esc(x.人数 ?? "")}" inputmode="numeric" placeholder="人数"><span>名</span><button type="button" data-delg="${j}" aria-label="外す">×</button></div>`).join("");
  return `<div class="card" data-id="${r.id}">
    <div class="ch"><span class="no">${i + 1}</span><span class="kind">${esc(r.コード ? "現場 " + r.コード : off ? "" : r.行先 ? "現場" : "")}</span><span class="sp"></span><button type="button" data-act="del">消す</button></div>
    <div class="row2">
      <div class="f"><label>行先</label><div class="pick"><input data-f="place" value="${esc(r.行先 || "")}" placeholder="現場名・本社（打つと絞り込み）" autocomplete="off"></div></div>
      <div class="f"><label>時間</label><div class="time"><input data-f="t0" value="${esc(r.始め || "")}" inputmode="decimal" placeholder="始め"><span>〜</span><input data-f="t1" value="${esc(r.終わり || "")}" inputmode="decimal" placeholder="終わり"></div>
        ${r.始め || r.終わり ? "" : `<div class="time quick" style="margin-top:6px"><button type="button" data-tq="8:00|17:00">8:00〜17:00</button><button type="button" data-tq="21:00|6:00">21:00〜6:00</button></div>`}</div>
    </div>
    <div class="f"><label>作業内容</label><textarea data-f="work" rows="1" placeholder="何をしたか">${esc(r.作業 || "")}</textarea></div>
    ${off && !(r.人 || []).length && !(r.外注 || []).length ? "" : `<div class="row3">
      <div class="f"><label>来ていた人（作業員）</label>${chips ? `<div class="chips">${chips}</div>` : ""}<div class="pick"><input data-f="person" placeholder="名前を打つか選ぶ" autocomplete="off"></div></div>
      <div class="f"><label>外注（業者名と人数）</label>${gai}<button type="button" class="addline" data-act="addg">＋ 外注を足す</button></div>
    </div>`}
    ${!off && r.行先 ? `<div class="detail"><button type="button" data-act="detail" title="常用日報の形で、作業予定・完了報告・始業終業・工種・機械・材料などを書く">詳しく書く（現場の日報）→</button></div>` : ""}
  </div>`;
}
function grow(t) { t.style.height = "auto"; t.style.height = t.scrollHeight + 2 + "px"; }
const findRow = (d, id) => (d.行 || []).find(x => x.id === id);
function setRow(id, fn, rerender) { mutate(S.day, d => { const r = findRow(d, id); if (r) fn(r); }, rerender ? { render: true } : true); }
function bindCard(id) {
  const c = document.querySelector(`.card[data-id="${id}"]`), q = s => c.querySelector(s), r = () => rowsOf().find(x => x.id === id) || {};
  const place = q('[data-f="place"]');
  bindPick(place, placeItems, x => { place.value = x.v; setRow(id, y => { y.行先 = x.v; y.コード = x.コード || ""; }, true); });
  place.addEventListener("change", () => { const v = place.value.trim(); if (v !== r().行先) { const s = S.sites.find(x => nk(siteName(x)) === nk(v)); setRow(id, y => { y.行先 = v; y.コード = s ? s.コード || "" : ""; }, true); } });
  ["t0", "t1"].forEach((f, j) => { const el = q(`[data-f="${f}"]`); el.addEventListener("change", () => { const v = normTime(el.value); el.value = v; setRow(id, y => { y[j ? "終わり" : "始め"] = v; }); });
    el.addEventListener("keydown", e => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); (j ? q('[data-f="work"]') : q('[data-f="t1"]')).focus(); } }); });
  const work = q('[data-f="work"]');
  work.addEventListener("input", () => { grow(work); const v = work.value; setRow(id, y => { y.作業 = v; }); });
  const person = q('[data-f="person"]');
  if (person) bindPick(person, () => personItems(r().行先, r().人 || []), x => {
    const n = String(x.v || "").trim(); if (!n) return;
    if ((r().人 || []).some(v => nk(v) === nk(n))) { person.value = ""; return; }
    setRow(id, y => { y.人 = [...(y.人 || []), n]; }, true);
    setTimeout(() => { const p = document.querySelector(`.card[data-id="${id}"] [data-f="person"]`); if (p) p.focus(); }, 0);
  });
  c.querySelectorAll("[data-gi]").forEach(el => { const j = +el.dataset.gi;
    bindPick(el, gaiItems, x => { el.value = x.v; setRow(id, y => { y.外注[j].業者 = x.v; }); });
    el.addEventListener("change", () => setRow(id, y => { if (y.外注[j]) y.外注[j].業者 = el.value.trim(); })); });
  c.querySelectorAll("[data-gn]").forEach(el => { const j = +el.dataset.gn; el.addEventListener("change", () => { const v = el.value.normalize("NFKC").replace(/[^\d.]/g, ""); el.value = v; setRow(id, y => { if (y.外注[j]) y.外注[j].人数 = v; }); }); });
  c.addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    if (b.dataset.tq) { const [a, z] = b.dataset.tq.split("|"); setRow(id, y => { y.始め = a; y.終わり = z; }, true); }
    if (b.dataset.delp != null) { const j = +b.dataset.delp; setRow(id, y => { y.人.splice(j, 1); }, true); }
    if (b.dataset.delg != null) { const j = +b.dataset.delg; setRow(id, y => { y.外注.splice(j, 1); }, true); }
    if (b.dataset.act === "addg") { setRow(id, y => { (y.外注 ||= []).push({ 業者: "", 人数: "" }); }, true); setTimeout(() => { const els = document.querySelectorAll(`.card[data-id="${id}"] [data-gi]`); els.length && els[els.length - 1].focus(); }, 0); }
    if (b.dataset.act === "del") { const x = r(); if ((x.行先 || x.作業 || (x.人 || []).length) && !confirm(`「${x.行先 || "この行"}」を消しますか？`)) return; mutate(S.day, d => { d.行 = d.行.filter(y => y.id !== id); if (!d.行.length) d.報告 = null; }); }
    if (b.dataset.act === "detail") toDetail(r());
  });
}
// 打っている欄がキーボードに隠れないように（iPhone）
document.addEventListener("focusin", e => { if (e.target.matches("input,textarea") && window.innerWidth < 760) setTimeout(() => e.target.scrollIntoView({ block: "center", behavior: "smooth" }), 300); });

/* ---------- 下のボタン ---------- */
function addRow() {
  const rows = rowsOf(), last = rows[rows.length - 1], id = uid();
  // 2つ目からは、前の行の終わりを始めに
  mutate(S.day, d => { d.行.push({ id, 行先: "", コード: "", 始め: last ? last.終わり || "" : "", 終わり: "", 作業: "", 人: [], 外注: [] }); }, { render: true });
  setTimeout(() => { const p = document.querySelector(`.card[data-id="${id}"] [data-f="place"]`); if (p) { p.focus(); p.scrollIntoView({ block: "center" }); } }, 30);
}
async function sameAsBefore() {
  await loadMonth(monthOf(dayAdd(S.day, -31)));
  let src = null;
  for (let i = 1; i <= 31 && !src; i++) { const k = dayAdd(S.day, -i); if (rowsOf(k).length) src = k; }
  if (!src) { toast("前の31日に書いた日がありません"); return; }
  const d = dateOf(src), rows = rowsOf(src);
  const list = rows.map(r => `<li>${esc(r.行先 || "（行先なし）")} ${esc([r.始め, r.終わり].filter(Boolean).join("〜"))}${(r.人 || []).length ? `・${r.人.length}名` : ""}${(r.外注 || []).length ? "・外注" : ""}</li>`).join("");
  dlg(`<h2>${d.getMonth() + 1}/${d.getDate()}(${WD[d.getDay()]}) と同じにしますか？</h2><ol>${list}</ol>
    <p style="font-size:14px;color:#5d6b7a">行先・時間・作業内容・来ていた人・外注を写します。${rowsOf().length ? "<b>この日に書いてある分は置き換わります。</b>" : ""}</p>
    <div class="btns"><button data-close>やめる</button><button class="pri" id="dOk">同じにする</button></div>`);
  $("dOk").onclick = () => { closeDlg(); const copy = JSON.parse(JSON.stringify(rows)).map(r => ({ ...r, id: uid() })); mutate(S.day, x => { x.行 = copy; x.報告 = null; }); toast("前の日と同じにしました。違う所を直してください"); };
}
function report() {
  const rows = rowsOf(), miss = rows.filter(r => !r.行先 || !r.始め || !r.終わり);
  const go = () => { mutate(S.day, d => { d.報告 = { だれ: S.me.name, いつ: nowStr() }; }, { report: true, render: true }); toast("報告しました"); };
  if (miss.length) { dlg(`<h2>書いていない所があります</h2><p>${miss.map(r => esc(r.行先 || "行先なし") + "：" + [!r.行先 && "行先", !r.始め && "始め", !r.終わり && "終わり"].filter(Boolean).join("・")).join("<br>")}</p><div class="btns"><button data-close>直す</button><button class="pri" id="dOk">このまま報告する</button></div>`); $("dOk").onclick = () => { closeDlg(); go(); }; return; }
  go();
}

/* ---------- 詳しく書く（常用日報へ渡す。同じ kkodin.github.io なのでブラウザの中で渡せる） ---------- */
function toDetail(r) {
  const x = { 日: S.day, 行先: r.行先, コード: r.コード || "", 始め: r.始め || "", 終わり: r.終わり || "", 作業: String(r.作業 || "").split(/\r?\n/).map(v => v.trim()).filter(Boolean), 人: r.人 || [], 外注: r.外注 || [], 書いた人: S.me.name, いつ: Date.now() };
  try { localStorage.setItem("nippou_to_jouyou", JSON.stringify(x)); } catch (e) { }
  const url = new URL("../jouyou/?from=nippou" + (DEMO ? "&demo" : ""), location.href).href;
  if (window.innerWidth < 760 || navigator.standalone) location.href = url; else window.open(url, "_blank");
}

/* ---------- この日のみんなの日報（行先ごと） ---------- */
async function renderAll() {
  $("cards").hidden = true; $("empty").hidden = true; $("all").hidden = false;
  $("dayHead").innerHTML = `<button id="bBack">← 自分の日報に戻る</button>`;
  $("bBack").onclick = () => { S.view = "me"; render(); };
  $("all").innerHTML = `<p>読んでいます…</p>`;
  const k = S.day, m = monthOf(k);
  try {
    const files = (await nippouStore.list()).filter(f => f.name.startsWith("個人_") && f.name.endsWith(`_${m}.json`));
    const docs = (await Promise.all(files.map(f => nippouStore.get(f.name).catch(() => null)))).filter(Boolean).map(x => x.data);
    const by = new Map(), none = [];
    docs.forEach(doc => { const d = (doc.日 || {})[k]; if (!d || !(d.行 || []).length) { none.push(doc.名前); return; }
      d.行.forEach(r => { const key = r.行先 || "（行先なし）"; if (!by.has(key)) by.set(key, []); by.get(key).push({ who: doc.名前, r, rep: !!d.報告 }); }); });
    if (S.view !== "all" || S.day !== k) return;
    const d = dateOf(k);
    $("all").innerHTML = `<h3>${d.getMonth() + 1}/${d.getDate()}(${WD[d.getDay()]}) のみんなの日報（${docs.length - none.length} 人）</h3>` +
      ([...by.entries()].sort((a, b) => isOffice(a[0]) - isOffice(b[0]) || a[0].localeCompare(b[0], "ja")).map(([place, xs]) => {
        const ppl = new Set(), gai = []; xs.forEach(x => { (x.r.人 || []).forEach(n => ppl.add(n)); (x.r.外注 || []).forEach(g => g.業者 && gai.push(`${g.業者}${g.人数 ? " " + g.人数 + "名" : ""}`)); });
        return `<div class="allsite"><h4>${esc(place)}${ppl.size ? `　<small>作業員 ${ppl.size}名</small>` : ""}${gai.length ? `　<small>外注 ${esc(gai.join("・"))}</small>` : ""}</h4>` +
          xs.map(x => `<div class="row"><span class="who">${esc(x.who)}</span><span class="t">${esc([x.r.始め, x.r.終わり].filter(Boolean).join("〜"))}${x.rep ? "" : "・書きかけ"}</span><br>${esc(x.r.作業 || "")}${(x.r.人 || []).length ? `<br><small>来ていた人：${esc(x.r.人.join("、"))}</small>` : ""}</div>`).join("") + `</div>`;
      }).join("") || `<p>この日はまだだれも書いていません。</p>`) +
      (none.length ? `<p style="font-size:13px;color:#5d6b7a">この日をまだ書いていない人：${esc(none.join("、"))}</p>` : "");
  } catch (e) { $("all").innerHTML = `<p class="ng">読めませんでした：${esc(e.message)}</p>`; }
}

/* ---------- 窓・知らせ ---------- */
function dlg(html) { $("dlgBody").innerHTML = html; const d = $("dlg"); if (!d.open) d.showModal(); }
function closeDlg() { $("dlg").close(); }
$("dlg").addEventListener("click", e => { if (e.target.closest("[data-close]") || e.target === $("dlg")) closeDlg(); });
let toastT = null;
function toast(m, ms = 3000) { const t = $("toast"); t.textContent = m; t.classList.add("on"); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("on"), ms); }

/* ---------- 書く人（サインインしたアカウント ↔ 名簿の名前。最初に1回だけ選ぶ。打たない） ---------- */
const ME_KEY = "nippou_me_v3";
function savedMe(acct) { try { return (JSON.parse(localStorage.getItem(ME_KEY)) || {})[acct] || ""; } catch (e) { return ""; } }
function askMe(acctName, first) {
  return new Promise(res => {
    const list = people(), guess = nk(acctName);
    const btns = () => { const q = nk($("meQ").value); return list.filter(n => !q || nk(n).includes(q)).map(n => `<button type="button" data-n="${esc(n)}" class="${nk(n) === guess ? "hit" : ""}">${esc(n)}</button>`).join("") || "<p>名簿に見つかりません</p>"; };
    dlg(`<h2>${first ? "あなたの名前を選んでください" : "名前を選び直す"}</h2><p style="font-size:14px">この名前で日報を書きます（サインイン：${esc(acctName)}）。次からは聞きません。</p>
      <input id="meQ" placeholder="名前で絞り込み" autocomplete="off"><div class="namelist" id="meL"></div>${first ? "" : `<div class="btns"><button data-close>やめる</button></div>`}`);
    $("meL").innerHTML = btns();
    $("meQ").oninput = () => { $("meL").innerHTML = btns(); };
    $("meL").onclick = e => { const b = e.target.closest("[data-n]"); if (!b) return; closeDlg(); res(b.dataset.n); };
    const hit = $("meL").querySelector(".hit"); if (hit) hit.scrollIntoView({ block: "center" });
  });
}
function rememberMe(name) {
  S.me.name = name;
  try { const a = JSON.parse(localStorage.getItem(ME_KEY)) || {}; a[S.me.acct] = name; localStorage.setItem(ME_KEY, JSON.stringify(a)); } catch (e) { }
}
async function setMe(name) { rememberMe(name); S.F = {}; await loadMonth(monthOf(S.day)); render(); }

/* ---------- ホーム画面に置く（アドレスバーを消す） ---------- */
let installEv = null;
window.addEventListener("beforeinstallprompt", e => { e.preventDefault(); installEv = e; });
const standalone = () => window.matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
function installGuide() {
  if (installEv) { installEv.prompt(); installEv = null; return; }
  const ios = /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  dlg(`<h2>ホーム画面に置く</h2><p>ホーム画面（デスクトップ）から開くと、アドレスバーやタブが出ず、入力する所が広くなります。</p>` +
    (ios ? `<ol><li>Safari でこのページを開く</li><li>下（iPad は上）の <b>共有</b>（□に↑）を押す</li><li><b>ホーム画面に追加</b> → <b>追加</b></li><li>ホーム画面の「工事日報」から開く（最初の1回だけサインインし直します）</li></ol>`
      : `<ol><li>Edge か Chrome で、アドレスバーの右の <b>アプリをインストール</b>（□に↓）を押す</li><li>または「…」メニュー → <b>アプリ</b> → <b>このサイトをアプリとしてインストール</b></li><li>デスクトップ・スタートの「工事日報」から開く</li></ol>`) +
    `<div class="btns"><button class="pri" data-close>わかった</button></div>`);
}

/* ---------- 日報のフォルダを調べる（管理者が1回だけ。?setup で開くと「…」に出る） ----------
   OneDrive の共有リンクから driveId・itemId を出し、config.js の nippou に書いてもらう（リンクそのものは config に書かない） */
function findFolder() {
  dlg(`<h2>日報のフォルダを調べる</h2><p style="font-size:14px">OneDrive で日報のフォルダを選び、「共有」→「リンクのコピー」で写したリンクを貼ってください。</p>
    <input id="shUrl" placeholder="https://…sharepoint.com/…" autocomplete="off"><pre id="shOut" style="white-space:pre-wrap;font-size:14px"></pre>
    <div class="btns"><button data-close>閉じる</button><button class="pri" id="shGo">調べる</button></div>`);
  $("shGo").onclick = async () => {
    const u = $("shUrl").value.trim(); if (!u) return;
    const id = "u!" + btoa(unescape(encodeURIComponent(u))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    $("shOut").textContent = "調べています…";
    try {
      const r = await g(`/shares/${id}/driveItem`); if (!r.ok) throw httpErr(r, "リンクを開けません");
      const j = await r.json(); if (!j.folder) throw new Error("フォルダのリンクではありません");
      $("shOut").textContent = `「${j.name}」
nippou: {
  driveId: "${j.parentReference.driveId}",
  itemId: "${j.id}"
}

この4行を Claude に渡してください（config.js に入れます）。`;
    } catch (e) { $("shOut").textContent = "調べられませんでした：" + e.message; }
  };
}

/* ---------- 日を移る・ボタン ---------- */
async function setDay(k) {
  S.day = k; closeSug();
  try { localStorage.setItem("nippou_view_v3", JSON.stringify({ day: k })); } catch (e) { }
  render();
  await loadMonth(monthOf(k)); if (S.day === k) render();
}
function bind() {
  $("bPrev").onclick = () => setDay(dayAdd(S.day, -1));
  $("bNext").onclick = () => setDay(dayAdd(S.day, 1));
  $("bToday").onclick = () => setDay(todayKey());
  $("inDay").onchange = () => { if ($("inDay").value) setDay($("inDay").value); };
  $("bAdd").onclick = () => { if (S.view === "all") { S.view = "me"; render(); } addRow(); };
  $("bSame").onclick = sameAsBefore;
  $("bReport").onclick = report;
  $("bMe").onclick = async () => { const n = await askMe(S.me.acct, false); if (n) setMe(n); };
  $("bMore").onclick = e => { e.stopPropagation(); $("menu").hidden = !$("menu").hidden; };
  document.addEventListener("click", e => { if (!e.target.closest("#menu")) $("menu").hidden = true; });
  $("menu").onclick = async e => {
    const m = e.target.closest("[data-m]"); if (!m) return; $("menu").hidden = true;
    if (m.dataset.m === "all") { S.view = "all"; render(); }
    if (m.dataset.m === "reload") { await loadMasters(); await loadMonth(monthOf(S.day), true); render(); toast("最新を読みました"); }
    if (m.dataset.m === "install") installGuide();
    if (m.dataset.m === "me") $("bMe").click();
    if (m.dataset.m === "find") findFolder();
  };
  // PC：Alt+←/→ で日を移る
  document.addEventListener("keydown", e => { if (e.altKey && (e.key === "ArrowLeft" || e.key === "ArrowRight")) { e.preventDefault(); setDay(dayAdd(S.day, e.key === "ArrowLeft" ? -1 : 1)); } });
  window.addEventListener("beforeunload", e => { if (Object.values(S.F).some(f => f.pending.length || f.saving)) { e.preventDefault(); e.returnValue = ""; } });
  document.addEventListener("visibilitychange", () => { if (!document.hidden && !Object.values(S.F).some(f => f.pending.length || f.saving)) loadMonth(monthOf(S.day), true).then(() => { if (!document.activeElement || !document.activeElement.matches("input,textarea")) render(); }); });
}

async function start() {
  if ("serviceWorker" in navigator && location.protocol === "https:") navigator.serviceWorker.register("sw.js").catch(() => { });
  try { const v = JSON.parse(localStorage.getItem("nippou_view_v3")) || {}; if (v.day) S.day = v.day; } catch (e) { }
  cachedMasters();
  if (DEMO) { seedDemo(); S.me.acct = "試し"; }
  else if (location.protocol === "file:") {
    $("signin").hidden = false; $("bSignIn").hidden = true;
    $("signinMsg").innerHTML = 'パソコンのファイルを直接開いているため、サインインできません。<br>公開先から開いてください：<a href="https://kkodin.github.io/kotei/nippou3/">https://kkodin.github.io/kotei/nippou3/</a>';
    return;
  } else {
    msalApp = new msal.PublicClientApplication({ auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: new URL("./", location.href).href }, cache: { cacheLocation: "localStorage" } });
    await msalApp.initialize();
    try { const r = await msalApp.handleRedirectPromise(); if (r) msalApp.setActiveAccount(r.account); } catch (e) { $("signinMsg").textContent = "サインインできませんでした：" + e.message; }
    const acc = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
    if (!acc) { $("signin").hidden = false; $("bSignIn").onclick = () => msalApp.loginRedirect({ scopes: SCOPES }); return; }
    msalApp.setActiveAccount(acc);
    S.me.acct = acc.username || acc.name;
    S.me.acctName = acc.name || acc.username;
  }
  $("app").hidden = false; $("bar").hidden = false;
  if (!DEMO && new URLSearchParams(location.search).has("setup")) $("mFind").hidden = false;
  bind();
  await loadMasters().catch(e => toast("マスタを読めませんでした（前に読んだ控えで動きます）：" + e.message));
  let name = savedMe(S.me.acct);
  if (!name) name = await askMe(S.me.acctName || S.me.acct, true);
  rememberMe(name);
  await setDay(S.day);
  if (!standalone() && !localStorage.getItem("nippou_install_told") && window.innerWidth < 1100) { try { localStorage.setItem("nippou_install_told", "1"); } catch (e) { } setTimeout(installGuide, 800); }
}

/* ---------- 試し（?demo）：架空の名簿・現場・前の日の日報 ---------- */
function seedDemo() {
  try { if (localStorage.getItem(DEMO_KEY)) return; } catch (e) { return; }
  const put = (name, data) => { const all = JSON.parse(localStorage.getItem(DEMO_KEY) || "{}"); all[name] = { data, eTag: "e" + Math.random().toString(36).slice(2) }; localStorage.setItem(DEMO_KEY, JSON.stringify(all)); };
  if (GUEST_DEMO) put("日報/名簿.json", { 形式: "工事日報の名簿", 版: 1, 社員: ["見本　一郎", "見本　二郎", "試験　花子"], 従業員: ["作業　三郎", "作業　四郎", "作業　五郎"], 外注業者: ["第一カッター興業", "見本警備"], 現場: [{ コード: "510001", 名前: "A現場 配水管", 正式名: "A町配水管布設工事", 元請: "見本市水道局" }] });
  put("工程/マスタ.json", { 社員: ["見本　一郎", "見本　二郎", "試験　花子"], 従業員: ["作業　三郎", "作業　四郎", "作業　五郎"], 外注業者: ["第一カッター興業", "見本警備"] });
  put("機械/現場一覧.json", { 現場: [{ コード: "510001", 名前: "A現場 配水管", 正式名: "A町配水管布設工事", 元請: "見本市水道局" }, { コード: "510002", 名前: "B現場 道路改良", 正式名: "B線道路改良工事", 元請: "見本建設" }] });
  const y = dayAdd(todayKey(), -1), m = monthOf(y);
  put(`日報/個人_見本一郎_${m}.json`, { 形式: "私の日報", 版: 1, 名前: "見本　一郎", アカウント: "試し", 月: m, 日: { [y]: { 行: [
    { id: "a1", 行先: "本社", コード: "", 始め: "8:00", 終わり: "10:00", 作業: "朝礼・書類", 人: [], 外注: [] },
    { id: "a2", 行先: "A現場 配水管", コード: "510001", 始め: "10:00", 終わり: "17:00", 作業: "配水管布設 φ150 L=12m", 人: ["作業　三郎", "作業　四郎"], 外注: [{ 業者: "第一カッター興業", 人数: "2" }] }], 報告: { だれ: "見本　一郎", いつ: "" } } } });
  put(`日報/個人_見本二郎_${m}.json`, { 形式: "私の日報", 版: 1, 名前: "見本　二郎", アカウント: "試し2", 月: m, 日: { [y]: { 行: [
    { id: "b1", 行先: "B現場 道路改良", コード: "510002", 始め: "8:00", 終わり: "17:00", 作業: "伐採工", 人: ["作業　五郎"], 外注: [] }], 報告: null } } });
}

start();
