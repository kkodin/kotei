/* マスタを合わせる仕組み（2026-11-16 開発環境の整理：合わせ方はここ1か所だけ）
   人：
     人.json        … 毎朝、01_マスタ抽出\人を作る.py が作る（配置表の名簿＋作業員名簿）。ここは手で直さない（作り直すと戻る）
                      { 形式:"人", 作成, 人:[{ 名前, ふりがな, 区分:"社員"|"従業員"|"外注", 会社, 職種, 元:["配置表","作業員名簿"], 回数 }], 違い:[{ 名前, 内容 }] }
     人_手直し.json … 名簿の画面（meibo/）が書く。名前ごとの直し・足した人・最後の月（退職）
                      { 形式:"人_手直し", 人:{ "<名前（空白なし）>": { 名前, ふりがな, 区分, 会社, 職種, 最後の月, 外す, 追加, メモ, 更新, 更新者 } } }
     退職者.json    … 前からある退職の一覧（{ 人:[{ 名前, 最後の月 }] }）。人_手直し の「最後の月」が無い人だけ使う
   HX.people(人, 手直し, 退職者) → [{ 名前, ふりがな, 区分, 会社, 職種, 最後の月, 外す, 元, 直した:[項目], 追加 }]
   HX.activeOn(p, "YYYY-MM") … その月に候補へ出してよいか（最後の月の次の月から出さない・外した人は出さない） */
window.HX = window.HX || {};
(function (HX) {
  const key = s => String(s || "").replace(/[\s　]/g, "");
  HX.nameKey = key;
  const FIELDS = ["名前", "ふりがな", "区分", "会社", "職種"];
  HX.people = function (base, fix, ret) {
    const out = new Map();
    for (const p of (base && base.人) || []) out.set(key(p.名前), { ...p, 元: p.元 || [], 直した: [] });
    for (const r of (ret && ret.人) || []) { const k = key(r.名前), p = out.get(k); if (p && r.最後の月) p.最後の月 = r.最後の月; }
    const fx = (fix && fix.人) || {};
    for (const [k0, f] of Object.entries(fx)) {
      const k = key(k0);
      let p = out.get(k);
      if (!p) { if (!f.追加) continue; p = { 名前: f.名前 || k0, 元: ["手で足した"], 直した: [], 追加: true }; out.set(k, p); }
      for (const c of FIELDS) if (f[c] != null && f[c] !== "" && f[c] !== p[c]) { p[c] = f[c]; p.直した.push(c); }
      if ("最後の月" in f) { p.最後の月 = f.最後の月 || ""; if (f.最後の月) p.直した.push("最後の月"); }
      if (f.外す) { p.外す = true; p.直した.push("外す"); }
      if (p.追加) p.直した = p.直した.filter(c => c === "外す" || c === "最後の月");   // 足した人は「足した」と出す（直したとは出さない）
      if (f.メモ) p.メモ = f.メモ;
      p.更新 = f.更新; p.更新者 = f.更新者;
    }
    const order = { 社員: 0, 従業員: 1, 外注: 2 };
    return [...out.values()].sort((a, b) => (order[a.区分] ?? 3) - (order[b.区分] ?? 3) || String(a.会社 || "").localeCompare(String(b.会社 || ""), "ja") || String(a.ふりがな || a.名前).localeCompare(String(b.ふりがな || b.名前), "ja"));
  };
  HX.activeOn = (p, ym) => !p.外す && !(p.最後の月 && ym > p.最後の月);
})(window.HX);
