/* OneDrive の読み書き（2026-11-16 開発環境の整理：工事日報 Vr005 の GraphFolder・DemoFolder を共通にしたもの）
   const st = HX.store(msalApp, "kotei", "マスタ");   // 共通の設定（HX_CONFIG）の kotei フォルダの下の「マスタ」
   await st.list()                → [{ name, eTag }]
   await st.get("人.json")        → { data, eTag } ／ 無ければ null
   await st.put("人.json", data, eTag)  → { eTag }。eTag が違えば 412（ほかの人が先に書いた）、eTag 無しで既にあれば 409
   試し（?demo）のときは HX.demoStore(キー, 前置き) を使う（ブラウザの中だけ）。 */
window.HX = window.HX || {};
(function (HX) {
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const enc = p => p.split("/").map(encodeURIComponent).join("/");
  const httpErr = (res, what) => { const e = new Error(what + " " + res.status); e.status = res.status; return e; };
  HX.store = function (app, where, sub) {
    const C = window.HX_CONFIG, F = { driveId: C.driveId, itemId: C[where].itemId };
    const g = async (path, opt = {}) => fetch(GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + await HX.token(app), ...(opt.headers || {}) } });
    const base = () => `/drives/${F.driveId}/items/${F.itemId}`, p = name => base() + ":/" + enc((sub ? sub + "/" : "") + name);
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
        const send = () => g(url, { method: "PUT", headers: { "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: JSON.stringify(data, null, 1) });
        let res = await send();
        if (res.status === 404 && sub && !eTag) {   // フォルダがまだ無い → 作ってからもう一度
          await g(base() + "/children", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: sub, folder: {}, "@microsoft.graph.conflictBehavior": "replace" }) });
          res = await send();
        }
        if (!res.ok) throw httpErr(res, "保存");
        return { eTag: (await res.json()).eTag };
      }
    };
  };
  HX.demoStore = function (KEY, prefix) {
    const all = () => { try { return JSON.parse(localStorage.getItem(KEY)) || {}; } catch (e) { return {}; } };
    const save = a => { try { localStorage.setItem(KEY, JSON.stringify(a)); } catch (e) { } };
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
  };
})(window.HX);
