/* 全アプリの一覧と、アプリどうしの行き来（2026-11-16 開発環境の整理）
   ・アプリを足すときは、下の APPS に1行足すだけ。アドレスに版の番号は入れない。
   ・HX.appUrl("nippou", "?stay=1") … 置き場のいちばん上から数えたアドレス（どの画面から呼んでも同じ）。試し（?demo）は引き継ぐ。
   ・<nav data-hx-apps></nav> を置くと、ほかのアプリへのボタンを並べる（今のアプリは出さない）。 */
window.HX = window.HX || {};
(function (HX) {
  if (!HX.root) { const me = document.currentScript && document.currentScript.src; HX.root = me ? new URL("../", me).href : new URL("/", location.href).href; }
  // ヒントをポインターに隠れない所に出す部品（2026-11-18 ユーザー。appmenu.js を読まない名簿などのため）
  if (!window.__hxTip) { const s = document.createElement("script"); s.src = new URL("common/tip.js", HX.root).href; document.head.appendChild(s); }
  HX.APPS = [
    { id: "kotei", name: "全体工程", path: "", title: "全体工程表" },
    { id: "nippou", name: "工事日報", path: "nippou/", title: "工事日報（1週間の表・人別の一覧）" },
    { id: "shousai", name: "詳細日報", path: "nippou/shousai/", title: "詳細日報（作業日報の用紙・集計）" },
    { id: "kikai", name: "機械予定", path: "kikai/", title: "自社機械稼働表" },
    { id: "bukai", name: "部会資料", path: "bukai/", title: "工事部会資料" },
    { id: "haichi", name: "配置表", path: "haichi/", title: "配置表（見るだけ）" },
    { id: "meibo", name: "名簿", path: "meibo/", title: "人の名簿（確かめる・直す・足す・退職）" }
  ];
  HX.demo = new URLSearchParams(location.search).has("demo");
  HX.appUrl = function (id, query) {
    const a = HX.APPS.find(x => x.id === id); if (!a) throw new Error("アプリが無い：" + id);
    let q = query || "";
    if (HX.demo && !/[?&]demo(&|=|$)/.test(q)) q += (q.includes("?") ? "&" : "?") + "demo";
    return new URL(a.path + q, HX.root).href;
  };
  HX.here = function () {   // 今のアプリ（いちばん長く当てはまる path）
    const p = location.href.split(/[?#]/)[0];
    return HX.APPS.filter(a => p.startsWith(new URL(a.path, HX.root).href)).sort((a, b) => b.path.length - a.path.length)[0] || null;
  };
  HX.go = (id, query) => { location.href = HX.appUrl(id, query); };   // 同じ窓で（2026-11-03 ユーザー）
  HX.renderApps = function (el, opt) {
    const o = opt || {}, cur = HX.here(), skip = new Set([cur && cur.id, ...(o.skip || [])]);
    // only：出すアプリをこの順で（2026-11-17 ユーザー：工程管理のアプリには 工事日報・全体工程・機械予定 の3つ）。<nav data-hx-apps="nippou,kotei,kikai">
    const list = o.only ? o.only.map(id => HX.APPS.find(a => a.id === id)).filter(Boolean) : HX.APPS;
    el.innerHTML = list.filter(a => !skip.has(a.id)).map(a => `<button type="button" class="${o.cls || "sub"}" data-hx-app="${a.id}" title="${a.title}">${a.name}</button>`).join("");
    el.onclick = async e => {
      const b = e.target.closest("[data-hx-app]"); if (!b) return;
      if (o.before) { try { await o.before(b.dataset.hxApp); } catch (err) { } }
      HX.go(b.dataset.hxApp);
    };
  };
  document.addEventListener("DOMContentLoaded", () => document.querySelectorAll("nav[data-hx-apps]").forEach(el => { const v = el.getAttribute("data-hx-apps"); el.childElementCount || HX.renderApps(el, v ? { only: v.split(",").map(x => x.trim()) } : undefined); }));
})(window.HX);
