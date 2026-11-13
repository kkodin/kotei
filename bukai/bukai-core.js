/* 工事部会資料の共通データ。工程原本は参照し、月別の追記と配布版を分ける。 */
(function (root) {
  "use strict";
  const fields = {
    client: "発注者・元請", handler: "担当", summary: "工事概要", plan: "計画％", execPercent: "実行％",
    remarks: "備考", pipeDia: "管径・坑径・管種", length: "推進延長・立坑深", count: "数量", soil: "土質",
    amount: "請負金額", phone: "連絡先", start: "工期 自", end: "工期 至"
  };
  const clone = x => JSON.parse(JSON.stringify(x));
  const equal = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  function period(month) {
    if (!/^(19|20|21)\d{2}-(0[1-9]|1[0-2])$/.test(month || "")) throw Error("会議月を選んでください。");
    const y = +month.slice(0, 4), m = +month.slice(5), days = [];
    const last = new Date(y, m + 2, 0), end = `${last.getFullYear()}-${String(last.getMonth()+1).padStart(2,'0')}-${String(last.getDate()).padStart(2,'0')}`;
    for (let d = new Date(y, m-1, 1); d <= last; d.setDate(d.getDate()+1)) days.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`);
    return { from: month+"-01", end, days };
  }
  function dateOK(v) { return /^\d{4}-\d{2}-\d{2}$/.test(v) && !isNaN(Date.parse(v)) && new Date(v).toISOString().slice(0,10) === v; }
  function validate(field, value) {
    if (!(field in fields)) throw Error("未対応の項目です。");
    if (value === "" || value == null) return "";
    if (["plan", "execPercent", "amount"].includes(field)) {
      if (!/^(\d+(\.\d+)?|\.\d+)$/.test(String(value).trim())) throw Error(fields[field]+"は0以上の数値で入力してください。");
      const n = Number(value);
      if (!Number.isFinite(n) || n < 0 || (field !== "amount" && n > 100)) throw Error(fields[field]+"の値を確認してください。");
      return n;
    }
    if (typeof value !== "string" || value.length > 4000) throw Error(fields[field]+"は4000文字以内で入力してください。");
    if (["start","end"].includes(field) && !dateOK(value)) throw Error("工期は有効な日付にしてください。");
    return value;
  }
  function rows(sites, masters) {
    const out = [], codes = new Map();
    for (const s of sites) {
      const code = String(s.コード || "").trim();
      if (code && codes.has(code)) throw Error(`工事コード ${code} が工程表で重複しています。現場の設定で確認してください。`);
      if (code) codes.set(code, true);
      const m = code ? masters.find(x => String(x.コード || "").trim() === code) || {} : {};
      out.push({key: code ? "code:"+code : "site:"+s.id, siteId:s.id, code, name:s.略称 || s.名前 || "", done:!!s.完了,
        base:{client:m.元請 || m.発注者 || "", handler:s.担当 || m.担当 || "", summary:m.工事概要 || "", plan:"", execPercent:"", remarks:"",
          pipeDia:"", length:"", count:"", soil:"", amount:m.請負金額 ?? "", phone:m.電話 || "", start:s.工期自 || m.工期自 || "", end:s.工期至 || m.工期至 || ""}, source:s});
    }
    for (const m of masters) {
      const code=String(m.コード || "").trim();
      if (!code || codes.has(code)) continue;
      codes.set(code,true);
      out.push({key:"code:"+code,code,name:m.略称 || m.名前 || "",done:true,base:{client:m.元請 || m.発注者 || "",handler:m.担当 || "",summary:m.工事概要 || "",plan:"",execPercent:"",remarks:"",pipeDia:"",length:"",count:"",soil:"",amount:m.請負金額 ?? "",phone:m.電話 || "",start:m.工期自 || "",end:m.工期至 || ""},source:{段:[],バー:[]}});
    }
    return out;
  }
  function values(row, doc) { return {...row.base, ...((doc.entries || {})[row.key] || {}).values}; }
  function selected(row, doc) { const v=(doc.entries || {})[row.key]; return v && typeof v.included === "boolean" ? v.included : !row.done; }
  // 同じ項目が他の人に変更されていたら止める。別の項目の編集は残す。
  function patch(doc, changes) {
    const next=clone(doc); next.entries ||= {};
    for (const p of changes) {
      if (!p.key || !["values","included","categories"].includes(p.area)) throw Error("保存内容を確認してください。");
      const item=next.entries[p.key] ||= {};
      const actual=p.area === "values" ? (item.values || {})[p.field] : item[p.area];
      if (!equal(actual,p.before) && !equal(actual,p.after)) throw Error("同じ項目をほかの人が変更しました。入力内容を控えてから「最新を読み直す」で確認してください。");
      if (p.area === "values") { item.values ||= {}; if (p.after == null) delete item.values[p.field]; else item.values[p.field]=validate(p.field,p.after); }
      else item[p.area]=clone(p.after);
    }
    return next;
  }
  function importData(data, month, sourceRows, doc) {
    if (!data || data.yearMonth !== month || !Array.isArray(data.items) || !data.items.length || data.items.length>1000) throw Error("選択した会議月と一致する yearMonth・items のJSONを選んでください。");
    const seen=new Set(), changes=[], preview=[];
    for (const it of data.items) {
      if (!it || !/^\d{6}$/.test(String(it.projectCode || ""))) throw Error("工事コードは6桁で指定してください。");
      const key="code:"+it.projectCode, row=sourceRows.find(r=>r.key===key);
      if (!row) throw Error(`工事コード ${it.projectCode} が現場台帳・工程表にありません。先に現場を登録してください。`);
      if (seen.has(key)) throw Error(`工事コード ${it.projectCode} がJSON内で重複しています。`);
      seen.add(key);
      const old=(doc.entries || {})[key] || {}, flat={};
      for (const f of Object.keys(fields)) {
        const isMeta=["pipeDia","length","count","soil","amount","phone"].includes(f);
        let v=isMeta ? (it.metadata || {})[f] : it[f];
        if (f === "pipeDia" && it.metadata && ["pipeDia","shaftDia","pipeType"].some(k=>it.metadata[k] != null)) v=[it.metadata.pipeDia,it.metadata.shaftDia,it.metadata.pipeType].filter(x=>x !== "" && x != null).join(" / ");
        if (v == null) continue; // 未指定は保持。空文字は明示的な空欄として扱う。
        flat[f]=validate(f,v);
        changes.push({key,area:"values",field:f,before:(old.values || {})[f],after:flat[f]});
      }
      const checked={...values(row,doc),...flat};
      if (checked.start && checked.end && checked.start>checked.end) throw Error(`工事コード ${it.projectCode} の工期の開始日が終了日より後です。`);
      changes.push({key,area:"included",before:old.included,after:true});
      preview.push({row,values:{...values(row,doc),...flat}});
    }
    return {changes,preview};
  }
  const api={fields,clone,equal,period,dateOK,validate,rows,values,selected,patch,importData};
  if (typeof module !== "undefined" && module.exports) module.exports=api;
  root.BukaiCore=api;
})(typeof window === "undefined" ? globalThis : window);
