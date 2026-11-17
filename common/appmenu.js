/* アプリのタイトルから切り替える（2026-11-17 ユーザー：帯にボタンが多いと使いにくい。左上のタイトルを押すとプルダウンが出て、
   各アプリのタイトルを押すと同じ窓でそのアプリへ移る）。共通の部品の外のアプリ（全体工程表・自社機械稼働表・工事部会資料・配置表）も使えるよう、これだけで動く。
   使い方：<button type="button" class="appTitle" data-appmenu>作業報告カレンダー</button> を帯の左上に置き、このファイルを読む。
   ・アドレスは置き場のいちばん上から数える（このファイルの場所 common/ の1つ上）。試し（?demo）は引き継ぐ。
   ・プルダウンは、何も選ばずに外を押した・ほかのウィンドウへ移ったときに閉じる（フォームの共通ルール 5）。Esc でも閉じる。 */
(function () {
  "use strict";
  const me = document.currentScript && document.currentScript.src;
  const ROOT = me ? new URL("../", me).href : new URL("/", location.href).href;
  // ヒントをポインターに隠れない所に出す部品（common/tip.js。2026-11-18 ユーザー：どのアプリも同じに）
  if (me && !window.__hxTip) { const s = document.createElement("script"); s.src = new URL("tip.js", me).href; document.head.appendChild(s); }
  const APPS = [
    { name: "作業報告カレンダー", path: "nippou/", note: "工事日報（1週間の表・人別・現場別）" },
    { name: "作業日報", path: "nippou/shousai/", note: "詳細日報（作業日報の用紙・集計）" },
    { name: "全体工程表", path: "", note: "現場ごとの工程" },
    { name: "自社機械稼働表", path: "kikai/", note: "機械の予定・稼働" },
    // 2026-11-18 ユーザー：工事部会資料はプルダウンに出さず、代わりに配置表（閲覧のみ）を置く
    { name: "配置表（閲覧のみ）", path: "haichi/", note: "人の配置（見るだけ）" },
    // 2026-11-18 ユーザー：マスターを直す仕組みもここから開く。社員は名簿の人しか選べないので、新しい人はここで足す
    { name: "名簿（人のマスター）", path: "meibo/", note: "人を確かめる・直す・足す・退職" }
  ];
  const demo = new URLSearchParams(location.search).has("demo");
  const urlOf = a => new URL(a.path + (demo ? "?demo" : ""), ROOT).href;
  // いまのアプリ：アドレスがいちばん長く当てはまるもの（nippou/shousai/ は nippou/ より先）
  const here = () => { const p = location.href.split(/[?#]/)[0]; return APPS.filter(a => p.startsWith(new URL(a.path, ROOT).href)).sort((a, b) => b.path.length - a.path.length)[0] || null; };
  const css = `
.appTitle { font: inherit; font-weight: bold; color: inherit; background: transparent; border: 1px solid rgba(255,255,255,.45); border-radius: 6px; padding: 2px 10px; cursor: pointer; white-space: nowrap; display: inline-flex; align-items: center; gap: 6px; }
.appTitle::after { content: "▾"; font-size: .8em; opacity: .85; }
.appTitle:hover { background: rgba(255,255,255,.15); }
.appMenuBox { position: fixed; z-index: 2147483000; background: #fff; color: #1d2321; border: 1px solid #b7c0c8; border-radius: 8px; box-shadow: 0 8px 28px rgba(0,0,0,.25); padding: 4px; min-width: 240px; font-family: "BIZ UDPGothic", "Yu Gothic UI", "Meiryo", sans-serif; }
.appMenuBox a { display: block; padding: 8px 12px; border-radius: 6px; color: #1d2321; text-decoration: none; }
.appMenuBox a b { display: block; font-size: 15px; }
.appMenuBox a small { color: #6b7480; font-size: 12px; }
.appMenuBox a:hover, .appMenuBox a:focus { background: #e8eef5; outline: none; }
.appMenuBox a.cur { background: #1f3a5f; color: #fff; pointer-events: none; }
.appMenuBox a.cur small { color: #d6e0ea; }
@media print { .appTitle::after { display: none; } .appMenuBox { display: none !important; } }`;
  const st = document.createElement("style"); st.textContent = css; document.head.appendChild(st);
  let box = null, from = null;
  function close() { if (box) { box.remove(); box = null; } if (from) from.setAttribute("aria-expanded", "false"); from = null; }
  function open(btn) {
    close(); from = btn; btn.setAttribute("aria-expanded", "true");
    const cur = here();
    box = document.createElement("div"); box.className = "appMenuBox"; box.setAttribute("role", "menu");
    box.innerHTML = APPS.map(a => `<a role="menuitem" href="${urlOf(a)}" class="${cur === a ? "cur" : ""}"><b>${a.name}</b><small>${a.note}</small></a>`).join("");
    // 開いている窓（モーダルの dialog）の中で押したときは、その中に出す（外は押せないため）
    (btn.closest("dialog[open]") || document.body).appendChild(box);
    const r = btn.getBoundingClientRect();
    box.style.left = Math.max(4, Math.min(r.left, innerWidth - box.offsetWidth - 4)) + "px";
    box.style.top = Math.min(r.bottom + 4, innerHeight - box.offsetHeight - 4) + "px";
    const f = box.querySelector("a:not(.cur)"); if (f) f.focus({ preventScroll: true });
  }
  document.addEventListener("click", async e => {
    // 移る前にしておくこと（作業日報：入力中の欄を確定して共有してから）：window.appMenuBefore = async () => { … }
    const a = e.target.closest(".appMenuBox a");
    if (a && typeof window.appMenuBefore === "function") { e.preventDefault(); const to = a.href; close(); try { await window.appMenuBefore(); } catch (err) { } location.href = to; return; }
    const b = e.target.closest("[data-appmenu]");
    if (b) { e.preventDefault(); if (box && from === b) close(); else open(b); }
  });
  document.addEventListener("pointerdown", e => { if (box && !box.contains(e.target) && !e.target.closest("[data-appmenu]")) close(); }, true);
  window.addEventListener("blur", close);
  document.addEventListener("visibilitychange", () => { if (document.hidden) close(); });
  document.addEventListener("keydown", e => { if (e.key === "Escape" && box) { close(); e.stopPropagation(); } }, true);
  window.addEventListener("resize", close);
})();
