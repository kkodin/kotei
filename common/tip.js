/* ヒント（マウスを載せたときの説明）を、マウスポインターに隠れない所に出す（2026-11-18 ユーザー：
   ブラウザのヒントはポインターの後ろに出てほとんど読めない。どのアプリも同じ）。
   ・title のある所にマウスを載せると、title をしまって（ブラウザのヒントを出さない）、ポインターの右下に少し離して出す。
     画面の端に当たるときは左や上に出す。ポインターと一緒に動く。外れたら title を戻す。
   ・title を書き換えるアプリもそのまま動く（載せるたびに今の title を読む）。
   ・タッチだけの端末（iPad など）では何もしない。印刷には出さない。
   ・開いている窓（モーダルの dialog）の上にも出す（popover の一番上の層。使えないブラウザは窓の中に出す）。
   使い方：common/appmenu.js が読み込む（タイトルのプルダウンのあるアプリは何もしなくてよい）。ほかは <script src="…/common/tip.js"></script> */
(function () {
  "use strict";
  if (window.__hxTip) return; window.__hxTip = true;
  if (!(window.matchMedia && matchMedia("(hover: hover) and (pointer: fine)").matches)) return;
  const st = document.createElement("style");
  st.textContent = `.hxTip { position: fixed; z-index: 2147483600; max-width: min(440px, calc(100vw - 16px)); padding: 6px 10px; border-radius: 6px;
  background: #fffde7; color: #1d2321; border: 1px solid #9a8f4a; box-shadow: 0 4px 14px rgba(0,0,0,.25); font: 13px/1.5 "BIZ UDPGothic", "Yu Gothic UI", "Meiryo", sans-serif;
  white-space: pre-line; word-break: break-word; pointer-events: none; text-align: left; font-weight: normal; letter-spacing: 0; }
.hxTip[popover] { inset: auto; margin: 0; overflow: visible; }
@media print { .hxTip { display: none !important; } }`;
  (document.head || document.documentElement).appendChild(st);
  const POP = typeof HTMLElement.prototype.showPopover === "function";
  let tip = null, cur = null, text = "", lx = 0, ly = 0, timer = 0;
  const GAP_X = 16, GAP_Y = 22;   // ポインター（矢印は右下へ約12×20px）より外へ
  function place() {
    if (!tip) return;
    const w = tip.offsetWidth, h = tip.offsetHeight, W = innerWidth, H = innerHeight;
    let x = lx + GAP_X, y = ly + GAP_Y;
    if (x + w > W - 4) x = Math.max(4, lx - w - 10);        // 右に入らなければポインターの左
    if (y + h > H - 4) y = Math.max(4, ly - h - 12);        // 下に入らなければポインターの上
    tip.style.left = x + "px"; tip.style.top = y + "px";
  }
  function show() {
    if (!cur || !text) return;
    if (!tip) { tip = document.createElement("div"); tip.className = "hxTip"; tip.setAttribute("role", "tooltip"); if (POP) tip.setAttribute("popover", "manual"); }
    tip.textContent = text;
    if (POP) {
      // 一番上の層へ（あとから開いた窓より上に出すため、毎回出し直す）
      if (tip.parentNode !== document.body) document.body.appendChild(tip);
      try { if (tip.matches(":popover-open")) tip.hidePopover(); tip.showPopover(); } catch (e) { }
    } else {
      const host = cur.closest("dialog[open]") || document.body;
      if (tip.parentNode !== host) host.appendChild(tip);
    }
    place();
  }
  function hide() { clearTimeout(timer); if (tip) { try { if (POP && tip.matches(":popover-open")) tip.hidePopover(); } catch (e) { } tip.remove(); } }
  function release() {
    hide();
    if (cur && !cur.hasAttribute("title") && cur.dataset.hxTip != null) cur.setAttribute("title", cur.dataset.hxTip);
    if (cur) delete cur.dataset.hxTip;
    cur = null; text = "";
  }
  document.addEventListener("mouseover", e => {
    const t = e.target.closest ? e.target.closest("[title]") : null;
    // 出している所の中で動いた（中に別のヒントのある所が無い）→ そのまま
    if (cur && cur.contains(e.target) && !(t && t !== cur && cur.contains(t))) return;
    release();
    if (!t || t instanceof SVGElement || t === document.documentElement || !t.getAttribute("title")) return;
    cur = t; text = t.getAttribute("title"); t.dataset.hxTip = text; t.removeAttribute("title");
    lx = e.clientX; ly = e.clientY;
    clearTimeout(timer); timer = setTimeout(show, 350);
  }, true);
  document.addEventListener("mousemove", e => { lx = e.clientX; ly = e.clientY; if (tip && tip.parentNode) place(); }, true);
  document.addEventListener("mouseout", e => { if (cur && (!e.relatedTarget || !cur.contains(e.relatedTarget))) release(); }, true);
  ["pointerdown", "wheel", "keydown"].forEach(ev => document.addEventListener(ev, () => { hide(); }, true));
  window.addEventListener("blur", release);
  document.addEventListener("visibilitychange", () => { if (document.hidden) release(); });
})();
