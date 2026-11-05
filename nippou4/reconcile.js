/* 人と現場の照合。一括処理前に全変更を準備し、原本バックアップとeTagで保護する。
   詳細日報は人の列だけを直す。クロジカ本体は変更せず修正一覧を出す。 */
"use strict";
const RC = { choices: new Map(), checked: new Set(), plan: null, running: false };
const RC_SOURCES = [["app", "工事日誌（工事日報）"], ...DSRC];
const rcClone = x => JSON.parse(JSON.stringify(x));
const rcError = text => { throw new Error(text); };
function rcTarget(e, src) {
  if (!e[src]) rcError("選んだ比較元がありません。読み直してください。");
  const out = new Map(src === "jou" ? [...e.app].filter(([k]) => !(e.jouAt || new Set()).has(k)) : []);
  e[src].forEach((p, k) => out.set(k, p)); return out;
}
function reconcileOpen() {
  const list = (DIFF.list || []).filter(e => DIFF.all || !e.done);
  dlg(`<div class="dh"><h2>他の記録との違い：まとめて合わせる</h2><button data-close>×</button></div>
    <p class="pinfo">チェックした人・日ごとに採用する候補を選びます。工事日報・共有の配置確定データ・詳細日報の「人と現場」を合わせます。作業内容・機材・材料・金額は候補から上書きしません。</p>
    ${DIFF.err.length ? `<p class="ng">${esc(DIFF.err.join(" ／ "))}</p>` : ""}
    <div class="dt"><button id="rcAll">全件チェック</button><button id="rcNone">チェック解除</button><label><input id="dAll" type="checkbox"${DIFF.all ? " checked" : ""}>確認済みも表示</label></div>
    <div class="rc-common">チェックした行に同じ候補を指定：${RC_SOURCES.map(([s, t]) => `<button data-rc-common="${s}">${t}</button>`).join("")}</div>
    <div class="gwrap"><table class="list rc-table"><thead><tr><th>対象</th><th>日・人</th>${RC_SOURCES.map(([, t]) => `<th>${t}</th>`).join("")}</tr></thead><tbody>${list.map((e, i) => `<tr><td><input type="checkbox" data-rc-check="${i}"${RC.checked.has(e.id) ? " checked" : ""} aria-label="${esc(e.name)} ${e.k}を処理"></td><td>${esc(e.k)}<br>${esc(e.name)}</td>${RC_SOURCES.map(([s]) => `<td><label class="rc-option${e.diff[s] ? " different" : ""}"><input type="radio" name="rc-${i}" data-rc-source="${i}|${s}"${RC.choices.get(e.id) === s ? " checked" : ""}${!e[s] ? " disabled" : ""}><span>${e[s] ? esc(dtxt(e[s])) : "未取得"}</span></label></td>`).join("")}</tr>`).join("") || '<tr><td colspan="6">違いはありません。</td></tr>'}</tbody></table></div>
    <details class="rc-admin"><summary>管理者向け：クロジカCSV・比較の読み直し</summary><p class="pinfo">CSVの読み込みは管理者が行います。クロジカ本体への更新は行わず、修正一覧を出力します。</p><button id="dKuro">クロジカCSVを読み込んで比較</button><button id="rcReload">比較を読み直す</button></details>
    <p id="rcStatus" role="status"></p><div class="btns"><button id="rcPreview">チェックした行の変更内容を確認</button><button data-close>閉じる</button></div>`, "wide keep reconcile");
  const body = $("dlgBody");
  const status = () => $("rcStatus").textContent = `${list.filter(e => RC.checked.has(e.id)).length}件をチェックしています`;
  body.querySelectorAll("[data-rc-check]").forEach(el => el.onchange = () => { const e = list[+el.dataset.rcCheck]; el.checked ? RC.checked.add(e.id) : RC.checked.delete(e.id); status(); });
  body.querySelectorAll("[data-rc-source]").forEach(el => el.onchange = () => { const [i, s] = el.dataset.rcSource.split("|"); RC.choices.set(list[+i].id, s); RC.checked.add(list[+i].id); reconcileOpen(); });
  $("rcAll").onclick = () => { list.forEach(e => RC.checked.add(e.id)); reconcileOpen(); };
  $("rcNone").onclick = () => { RC.checked.clear(); reconcileOpen(); };
  body.querySelectorAll("[data-rc-common]").forEach(el => el.onclick = () => { const s = el.dataset.rcCommon; let missing = 0; list.filter(e => RC.checked.has(e.id)).forEach(e => e[s] ? RC.choices.set(e.id, s) : missing++); reconcileOpen(); if (missing) toast(`${missing}件は候補が未取得のため指定していません`); });
  $("dAll").onchange = () => { DIFF.all = $("dAll").checked; reconcileOpen(); };
  $("dKuro").onclick = diffPickKuro;
  $("rcReload").onclick = async () => { DIFF.H = { years: {} }; Object.keys(DOCS).forEach(m => delete DOCS[m]); await diffBadge(true); reconcileOpen(); };
  $("rcPreview").onclick = async () => {
    const chosen = list.filter(e => RC.checked.has(e.id));
    if (!chosen.length) { toast("処理する行をチェックしてください"); return; }
    if (chosen.some(e => !RC.choices.has(e.id))) { toast("チェックした行ごとに合わせる候補を選んでください"); return; }
    const b = $("rcPreview"); b.disabled = true; b.textContent = "変更内容を確認中…";
    try { RC.plan = await rcPrepare(chosen.map(e => ({ e, src: RC.choices.get(e.id) }))); rcPreview(); }
    catch (e) { toast(e.message, 10000); if (b.isConnected) { b.disabled = false; b.textContent = "チェックした行の変更内容を確認"; } }
  };
  status();
}
function rcHaichiStore() {
  if (DEMO) return {
    async get() { const d = await DemoFolder("照合配置/").get("追加分.js"); return d || { data: rcClone((window.HAICHI_DEMO || {})["追加分.js"] || { 確定: {}, 予定: {} }), eTag: null }; },
    async put(name, data, tag) { const r = await DemoFolder("照合配置/").put(name, data, tag); (window.HAICHI_DEMO ||= {})[name] = rcClone(data); return r; }
  };
  let location = null;
  return {
    async get() {
      for (const path of [CFG.haichi.path, CFG.haichi.oldPath].filter(Boolean)) {
        const url = `/drives/${CFG.haichi.driveId}/root:/${enc(path + "/追加分.js")}`, res = await g(url);
        if (res.status === 404) continue; if (!res.ok) throw httpErr(res, "配置確定の読み込み");
        const meta = await res.json(), body = await fetch(meta["@microsoft.graph.downloadUrl"], { cache: "no-store" });
        if (!body.ok) throw httpErr(body, "配置確定の中身");
        location = url; return { data: jsBody(await body.text()), eTag: meta.eTag };
      }
      rcError("配置確定の追加分.jsを読めません。管理者が配置データを確認してください。");
    },
    async put(name, data, tag) {
      const res = await g(location + ":/content", { method: "PUT", headers: { "Content-Type": "text/javascript; charset=utf-8", "If-Match": tag }, body: "window.HAICHI_PAST_ADD = " + JSON.stringify(data) + ";\n" });
      if (!res.ok) throw httpErr(res, "配置確定の保存"); return { eTag: (await res.json()).eTag };
    }
  };
}
function rcAdjustPersonal(doc, e, target) {
  const d = (doc.日 || {})[e.k]; if (!d) return;
  const own = nk(doc.名前) === nk(e.name), beforeRows = JSON.stringify(d.行);
  d.行 = (d.行 || []).filter(r => {
    if (target.has(plOf(r.行先, r.コード).key)) return true;
    if (!own && !(r.人 || []).some(n => nk(n) === nk(e.name)) && !((r.予定 || {}).人 || []).some(n => nk(n) === nk(e.name))) return true;
    const had = JSON.stringify(r);
    r.人 = (r.人 || []).filter(n => nk(n) !== nk(e.name));
    if (r.予定 && r.予定.人) r.予定.人 = r.予定.人.filter(n => nk(n) !== nk(e.name));
    if (own) {
      if ((r.人 || []).length || ((r.予定 || {}).人 || []).length || (r.外注 || []).length) r.本人 = false;
      else return false;
    }
    if (had !== JSON.stringify(r)) r.更新 = nowStr(); return true;
  });
  if (beforeRows !== JSON.stringify(d.行)) d.更新 = doc.更新 = nowStr();
}
function rcAdjustHaichi(d, e, target) {
  const match = value => nk(typeof value === "string" ? value : value && value.氏名) === nk(e.name);
  (d.現場 || []).forEach(row => {
    const key = plOf(row.現場名, row.現場コード).key;
    if (!target.has(key)) {
      row.作業員 = (row.作業員 || []).filter(w => !match(w));
      if (row.担当が行った && nk(tantoName(row.担当)) === nk(e.name)) row.担当が行った = false;
    }
  });
  d.本社 = (d.本社 || []).filter(w => !match(w));
  if (d.機材センター) d.機材センター.人 = (d.機材センター.人 || []).filter(w => !match(w));
  d.休み = (Array.isArray(d.休み) ? d.休み : []).filter(w => !match(w));
  target.forEach(pl => {
    if (pl.label === REST) { (d.休み ||= []).push({ 氏名: e.name }); return; }
    if (isOffice(pl.label)) { (d.本社 ||= []).push({ 氏名: e.name, 記載: "本社内勤" }); return; }
    if (pl.label === "機材センター") { ((d.機材センター ||= {}).人 ||= []).push({ 氏名: e.name }); return; }
    const rows = (d.現場 ||= []);
    let row = rows.find(r => plOf(r.現場名, r.現場コード).key === pl.key);
    if (!row) { row = { 現場名: pl.label, 現場コード: pl.code, 区分: pl.night ? "夜勤" : "昼勤", 作業内容: "", 作業員: [] }; rows.push(row); }
    if (!(row.作業員 || []).some(match) && !(row.担当が行った && nk(tantoName(row.担当)) === nk(e.name))) (row.作業員 ||= []).push({ 氏名: e.name });
  });
}
function rcAdjustDetail(s, e, target) {
  const key = plOf(s.略称 || s.名前, s.コード).key, wanted = target.has(key), d = (s.日 || {})[e.k];
  if (!d && !wanted) return;
  const day = (s.日 ||= {})[e.k] ||= { 上: [], 下: [], 外注: [], 車: [], 天気: "", 気温: "", 担当者: "", 更新: "" };
  let found = false, changed = false;
  for (const field of ["下", "外注"]) (day[field] || []).forEach(row => {
    if (!row || nk(row[0]) !== nk(e.name)) return;
    found = true; if (wanted) return;
    // 労務の列だけを空ける。機材(5〜8,13)・材料(9〜11,15)は同じ行でも残す。
    for (const col of [0, 1, 2, 3, 4, 12, 14]) row[col] = "";
    changed = true;
  });
  if (wanted && !found) {
    const pl = target.get(key), row = Array(16).fill(""); row[0] = e.name; row[14] = pl.night ? "夜" : "昼";
    (day.下 ||= []).push(row); changed = true;
  }
  if (changed) day.更新 = nowStr();
}
async function rcPrepare(selected) {
  if (Object.values(S.F).some(f => f.pending.length || f.saving)) rcError("工事日報の保存が終わってから処理してください。");
  DIFF.H = { years: {} };
  const fresh = await diffCalc(DIFF.m);
  selected.forEach(({ e, src }) => { const current = fresh.find(x => x.id === e.id); if (!current || current.fp !== e.fp || !current[src]) rcError("比較元が変わっています。比較を読み直してください。"); });
  const changes = new Map(), get = async (kind, store, name) => {
    const key = kind + "|" + name; if (changes.has(key)) return changes.get(key);
    const before = await store.get(name), entry = { kind, store, name, before: before ? rcClone(before.data) : null, eTag: before && before.eTag, after: before ? rcClone(before.data) : null };
    changes.set(key, entry); return entry;
  };
  const personal = new Map();
  for (const m of new Set(selected.map(x => monthOf(x.e.k)))) {
    const files = (await nippouStore.list()).filter(f => f.name.startsWith("個人_") && f.name.endsWith(`_${m}.json`));
    personal.set(m, await Promise.all(files.map(f => get("工事日報", nippouStore, f.name))));
  }
  const details = await Promise.all((await jouyouStore.list()).filter(f => /^現場_.*\.json$/.test(f.name)).map(f => get("詳細日報", jouyouStore, f.name)));
  const haichiStore = rcHaichiStore(), haichi = await get("配置確定", haichiStore, "追加分.js");
  const summary = [];
  for (const { e, src } of selected) {
    const target = rcTarget(e, src), m = monthOf(e.k);
    const docs = personal.get(m);
    for (const item of docs) rcAdjustPersonal(item.after, e, target);
    let own = docs.find(x => nk(x.after.名前) === nk(e.name));
    if (!own) {
      own = await get("工事日報", nippouStore, `個人_${nk(e.name)}_${m}.json`);
      own.after ||= { 形式: "私の日報", 版: 1, 名前: e.name, 月: m, 日: {} }; docs.push(own);
    }
    const day = (own.after.日 ||= {})[e.k] ||= { 行: [], 報告: null };
    target.forEach(pl => {
      if ((day.行 || []).some(r => plOf(r.行先, r.コード).key === pl.key && r.本人 !== false)) return;
      (day.行 ||= []).push(haiRow(e.k, pl.label, pl.code, pl.night, "", false, [], e.name, `check:${e.k}|${pl.key}`).r);
    }); day.更新 = own.after.更新 = nowStr();
    const hd = haichi.after.確定 && haichi.after.確定[e.k] || (DIFF.H.years[e.k.slice(0, 4)] || {}).確定?.[e.k];
    if (!hd) rcError(`${e.k} の配置確定データがありません。予定から確定を自動生成せずに停止しました。`);
    ((haichi.after.確定 ||= {})[e.k] ||= rcClone(hd)); rcAdjustHaichi(haichi.after.確定[e.k].data, e, target);
    target.forEach(pl => {
      if (isOffice(pl.label) || pl.label === REST || pl.label === "機材センター") return;
      if (!details.some(item => plOf(item.after.現場.略称 || item.after.現場.名前, item.after.現場.コード).key === pl.key)) rcError(`詳細日報に「${pl.label}」の現場がありません。現場マスターから作成してから処理してください。`);
    });
    for (const item of details) rcAdjustDetail(item.after.現場, e, target);
    summary.push({ id: e.id, 日: e.k, 人: e.name, 候補: RC_SOURCES.find(x => x[0] === src)[1], 合わせる現場: dtxt(target), クロジカの現状: e.kuro ? dtxt(e.kuro) : "未取得", クロジカ修正必要: !!e.kuro && !dsame(e.kuro, target) });
  }
  const files = [...changes.values()].filter(x => JSON.stringify(x.before) !== JSON.stringify(x.after));
  for (const file of files.filter(x => x.kind === "詳細日報")) {
    const s = file.after.現場, lock = await jouyouStore.get(`使用中_${s.id}.json`);
    if (lock && Date.now() - new Date(lock.data.いつ).getTime() < 180000) rcError(`詳細日報「${s.名前}」が使用中です。閉じてからもう一度処理してください。`);
  }
  // 権限エラーが起きやすい共有配置の保存を先に行う。複数ファイルは一括原子的には更新できない。
  files.sort((a, b) => ["配置確定", "詳細日報", "工事日報"].indexOf(a.kind) - ["配置確定", "詳細日報", "工事日報"].indexOf(b.kind));
  return { id: "照合_" + Date.now() + "_" + uid(), files, summary, selected, created: nowStr() };
}
function rcExport(plan) {
  const quote = v => '"' + String(v ?? "").replace(/"/g, '""') + '"';
  const rows = [["日", "人", "採用した候補", "合わせる現場", "クロジカの現状", "クロジカ本体での修正"]];
  plan.summary.forEach(s => rows.push([s.日, s.人, s.候補, s.合わせる現場, s.クロジカの現状, s.クロジカ修正必要 ? "要修正（この画面から本体は更新していません）" : "対象外／変更なし"]));
  const a = document.createElement("a"), url = URL.createObjectURL(new Blob(["\ufeff" + rows.map(r => r.map(quote).join(",")).join("\r\n")], { type: "text/csv;charset=utf-8" }));
  a.href = url; a.download = plan.id + "_修正一覧.csv"; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function rcPreview() {
  const p = RC.plan;
  dlg(`<div class="dh"><h2>一括処理の変更内容</h2><button data-close>×</button></div><p>変更前のファイルを日報フォルダにバックアップしてから処理します。詳細日報の作業内容・機材・材料・金額は残し、人の欄だけを調整します。</p>
    <table class="list"><tr><th>日・人</th><th>採用候補</th><th>合わせる現場</th></tr>${p.summary.map(s => `<tr><td>${esc(s.日)} ${esc(s.人)}</td><td>${esc(s.候補)}</td><td>${esc(s.合わせる現場)}</td></tr>`).join("")}</table>
    <p>更新するファイル：${p.files.map(f => esc(f.kind + " / " + f.name)).join("、") || "なし"}</p><p class="pinfo">クロジカは修正一覧の出力のみ。配置確定は共有の配置表データを更新します。元の配置Excelを再出力する際は修正内容も反映してください。途中で競合・権限エラーが起きた場合は停止し、完了したファイルを表示します。</p>
    <div class="btns"><button id="rcRun">バックアップしてまとめて処理</button><button id="rcCsv">修正一覧を出力</button><button id="rcBack">選択に戻る</button></div><p id="rcProgress" role="status"></p>`, "wide keep reconcile");
  $("rcCsv").onclick = () => rcExport(p); $("rcBack").onclick = reconcileOpen;
  $("rcRun").onclick = async () => {
    if (RC.running || !confirm(`${p.summary.length}件を選んだ候補に合わせます。変更前の記録をバックアップして実行しますか？`)) return;
    RC.running = true; $("dlgBody").querySelectorAll("button").forEach(b => b.disabled = true);
    try { const result = await rcExecute(p, msg => { $("rcProgress").textContent = msg; });
      $("rcProgress").textContent = `完了：${result.completed.length}ファイル。バックアップ：${result.backup}。クロジカ本体は未更新です。修正一覧を出力してください。`;
    } catch (e) { $("rcProgress").textContent = e.message; }
    finally { RC.running = false; $("rcCsv").disabled = false; $("rcBack").disabled = false; $("dlgBody").querySelector("[data-close]").disabled = false; }
  };
}
async function rcExecute(p, progress = () => {}) {
  if (Object.values(S.F).some(f => f.pending.length || f.saving)) rcError("工事日報の保存が終わっていません。処理を停止しました。");
  // 全件を保存前に読み直す。変更後のファイルを古い照合結果で上書きしない。
  for (const file of p.files) {
    const cur = await file.store.get(file.name);
    if ((cur && cur.eTag || null) !== (file.eTag || null) || JSON.stringify(cur && cur.data || null) !== JSON.stringify(file.before)) rcError("処理前に記録が変わりました。変更せず停止しました。比較を読み直してください。");
  }
  for (const file of p.files.filter(f => f.kind === "詳細日報")) {
    const lock = await jouyouStore.get(`使用中_${file.after.現場.id}.json`);
    if (lock && Date.now() - new Date(lock.data.いつ).getTime() < 180000) rcError("詳細日報が使用中になりました。変更せず停止しました。");
  }
  const backup = p.id + "_変更前.json";
  await nippouStore.put(backup, { 形式: "記録照合の変更前バックアップ", 版: 1, だれ: S.me.name, いつ: p.created, 選択: p.summary, ファイル: p.files.map(f => ({ 種別: f.kind, 名前: f.name, eTag: f.eTag, 内容: f.before })) }, null);
  const completed = [];
  let logTag = null;
  const log = async state => { const r = await nippouStore.put(p.id + "_実行結果.json", { 形式: "記録照合の実行結果", 照合: p.id, 状態: state, バックアップ: backup, 完了: completed, 対象: p.summary, 更新: nowStr() }, logTag); logTag = r.eTag; };
  await log("実行中");
  try {
    for (const file of p.files) { progress(`${completed.length + 1}/${p.files.length}：${file.kind} ${file.name}`); if (file.kind === "詳細日報") { const lock = await jouyouStore.get(`使用中_${file.after.現場.id}.json`); if (lock && Date.now() - new Date(lock.data.いつ).getTime() < 180000) rcError("詳細日報が使用中になりました"); }
      await file.store.put(file.name, file.after, file.eTag); completed.push(file.kind + "/" + file.name); await log("実行中"); }
    DIFF.H = { years: {} }; Object.keys(DOCS).forEach(m => delete DOCS[m]);
    const fresh = await diffCalc(DIFF.m), done = await nippouStore.get(DONE_FILE), doc = done ? done.data : { 形式: "工事日報の確認済みの違い", 版: 1, 済: {} };
    p.selected.forEach(({ e }) => { const current = fresh.find(x => x.id === e.id); if (current) doc.済[e.id] = { だれ: S.me.name, いつ: nowStr(), 中身: current.fp, 照合: p.id, クロジカ本体: "未更新・修正一覧で確認" }; });
    await nippouStore.put(DONE_FILE, doc, done && done.eTag);
    await diffBadge(true); await loadMonth(monthOf(S.day), true); render();
    await log("完了・クロジカは修正一覧を確認"); RC.checked.clear(); return { completed, backup };
  } catch (e) { DIFF.H = { years: {} }; Object.keys(DOCS).forEach(m => delete DOCS[m]); try { await log("停止：" + e.message); } catch (_) {} rcError(`処理を停止しました：${e.message}。完了したファイル：${completed.join("、") || "なし"}。変更前バックアップ：${backup}。残りは比較を読み直して確認してください。`); }
}
$("dlg").addEventListener("cancel", e => { if (RC.running) e.preventDefault(); });
diffOpen = reconcileOpen;
