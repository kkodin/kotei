/* 全体工程表（2026-10-07）
   画面は GitHub Pages、データは会社の OneDrive「工程データ」フォルダ（Microsoft Graph）。
   現場ごとに1ファイル（現場_<id>.json）。同時に書いたときは eTag で止め、読み直して自分の変更を重ねて保存する。
   ?demo を付けて開くと、サインインせずこのブラウザの中だけで試せる（OneDrive には書かない）。 */
"use strict";

const CFG = window.KOTEI_CONFIG;
const SCOPES = ["User.Read", "Files.ReadWrite.All"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const DEMO = new URLSearchParams(location.search).has("demo");
const POLL_MS = 30000;
const SHOW_EMP = false;   // 従業員の名前を入れる（将来の仕様。いまは作業員は人数だけ）
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------- 日付（"YYYY-MM-DD" の文字列で持つ。時差の狂いを避ける） ---------- */
const pad2 = n => String(n).padStart(2, "0");
const keyOf = d => d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
const dateOf = k => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
const dayAdd = (k, n) => { const d = dateOf(k); d.setDate(d.getDate() + n); return keyOf(d); };
const dayDow = k => dateOf(k).getDay();
const dayDiff = (a, b) => Math.round((dateOf(b) - dateOf(a)) / 86400000);
const todayKey = () => keyOf(new Date());
const md = k => { const d = dateOf(k); return (d.getMonth() + 1) + "/" + d.getDate(); };
const WD = "日月火水木金土";

/* 日本の祝日（原価管理 Vr013 の jpHolidays と同じ決まり） */
const JP_HOL = {};
function jpHolidays(y) {
  if (JP_HOL[y]) return JP_HOL[y];
  const h = {}, k = (m, d) => y + "-" + pad2(m) + "-" + pad2(d);
  const nthMon = (m, n) => { const w = new Date(y, m - 1, 1).getDay(); return 1 + ((8 - w) % 7) + (n - 1) * 7; };
  const add = (m, d, nm) => { h[k(m, d)] = nm; };
  add(1, 1, "元日"); add(1, nthMon(1, 2), "成人の日"); add(2, 11, "建国記念の日");
  if (y >= 2020) add(2, 23, "天皇誕生日");
  add(3, Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4)), "春分の日");
  add(4, 29, "昭和の日"); add(5, 3, "憲法記念日"); add(5, 4, "みどりの日"); add(5, 5, "こどもの日");
  add(7, nthMon(7, 3), "海の日"); add(8, 11, "山の日"); add(10, nthMon(10, 2), "スポーツの日");
  add(9, nthMon(9, 3), "敬老の日");
  add(9, Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4)), "秋分の日");
  add(11, 3, "文化の日"); add(11, 23, "勤労感謝の日");
  Object.keys(h).forEach(x => { const n = dayAdd(x, 2), mid = dayAdd(x, 1);
    if (h[n] && !h[mid] && dayDow(mid) !== 0 && mid.slice(0, 4) === String(y)) h[mid] = "国民の休日"; });
  Object.keys(h).sort().forEach(x => { if (dayDow(x) !== 0) return; let n = dayAdd(x, 1); while (h[n]) n = dayAdd(n, 1); if (n.slice(0, 4) === String(y)) h[n] = "振替休日"; });
  return JP_HOL[y] = h;
}
const holidayOf = k => jpHolidays(+k.slice(0, 4))[k] || "";

/* ---------- 色 ---------- */
const COLORS = [
  ["青", "#3d7cc9"], ["緑", "#3f9a45"], ["黄", "#f2c230"], ["橙", "#ef7d1a"],
  ["赤", "#d93a36"], ["紫", "#8a4fb0"], ["茶", "#8d6e63"], ["灰", "#8a96a3"], ["黒", "#37414b"]
];
const colorOf = n => (COLORS.find(c => c[0] === n) || COLORS[0])[1];
const isLight = n => n === "黄";

/* ---------- 保存先：OneDrive（Graph）／試し（このブラウザの中） ---------- */
let msalApp = null, me = { name: "" };

async function token() {
  try {
    return (await msalApp.acquireTokenSilent({ scopes: SCOPES })).accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) await msalApp.acquireTokenRedirect({ scopes: SCOPES });
    throw e;
  }
}
async function g(path, opt = {}) {
  return fetch(GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + await token(), ...(opt.headers || {}) } });
}
function httpErr(res, what) { const e = new Error(what + " " + res.status); e.status = res.status; return e; }

const GraphStore = {
  base: () => `/drives/${CFG.folder.driveId}/items/${CFG.folder.itemId}`,
  async list() {
    const out = [];
    let url = this.base() + "/children?$top=999";
    while (url) {
      const res = await g(url);
      if (!res.ok) throw httpErr(res, "一覧");
      const j = await res.json();
      j.value.forEach(x => out.push({ name: x.name, eTag: x.eTag, url: x["@microsoft.graph.downloadUrl"] }));
      url = j["@odata.nextLink"] ? j["@odata.nextLink"].replace(GRAPH, "") : null;
    }
    return out;
  },
  async get(name) {
    const p = this.base() + ":/" + encodeURIComponent(name);
    const res = await g(p);
    if (res.status === 404) return null;
    if (!res.ok) throw httpErr(res, "読み込み");
    const meta = await res.json();
    const dl = meta["@microsoft.graph.downloadUrl"];
    const body = dl ? await fetch(dl, { cache: "no-store" }) : await g(p + ":/content");
    if (!body.ok) throw httpErr(body, "中身の読み込み");
    return { data: await body.json(), eTag: meta.eTag };
  },
  // eTag があれば If-Match（違えば 412）。無ければ新規（あれば 409）
  async put(name, data, eTag) {
    const p = this.base() + ":/" + encodeURIComponent(name) + ":/content" + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
    const res = await g(p, { method: "PUT", headers: { "Content-Type": "application/json", ...(eTag ? { "If-Match": eTag } : {}) }, body: JSON.stringify(data, null, 1) });
    if (!res.ok) throw httpErr(res, "保存");
    return { eTag: (await res.json()).eTag };
  },
  // 画像（現場の地図）。上書きでよい（同時に直すことはまず無い）。読むときは data URL で返す
  async putRaw(name, blob) {
    const res = await g(this.base() + ":/" + encodeURIComponent(name) + ":/content", { method: "PUT", headers: { "Content-Type": blob.type || "application/octet-stream" }, body: blob });
    if (!res.ok) throw httpErr(res, "画像の保存");
  },
  async getRaw(name) {
    const res = await g(this.base() + ":/" + encodeURIComponent(name));
    if (res.status === 404) return null;
    if (!res.ok) throw httpErr(res, "画像の読み込み");
    const dl = (await res.json())["@microsoft.graph.downloadUrl"];
    const body = await fetch(dl, { cache: "no-store" });
    if (!body.ok) throw httpErr(body, "画像の読み込み");
    return blobToUrl(await body.blob());
  }
};
const blobToUrl = b => new Promise((ok, ng) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = ng; r.readAsDataURL(b); });

const DemoStore = {
  key: "kotei_demo_v1",
  all() { try { return JSON.parse(localStorage.getItem(this.key)) || {}; } catch (e) { return {}; } },
  save(a) { try { localStorage.setItem(this.key, JSON.stringify(a)); } catch (e) { } },
  async list() { const a = this.all(); return Object.keys(a).map(n => ({ name: n, eTag: a[n].eTag })); },
  async get(name) { const a = this.all(); return a[name] ? { data: JSON.parse(JSON.stringify(a[name].data)), eTag: a[name].eTag } : null; },
  async put(name, data, eTag) {
    const a = this.all(), cur = a[name];
    if (eTag ? (!cur || cur.eTag !== eTag) : cur) { const e = new Error("保存 " + (eTag ? 412 : 409)); e.status = eTag ? 412 : 409; throw e; }
    const t = "d" + Date.now() + Math.random().toString(36).slice(2, 6);
    a[name] = { data: JSON.parse(JSON.stringify(data)), eTag: t }; this.save(a);
    return { eTag: t };
  },
  async putRaw(name, blob) { try { localStorage.setItem(this.key + ":" + name, await blobToUrl(blob)); } catch (e) { throw new Error("画像が大きすぎます（試しの画面）"); } },
  async getRaw(name) { try { return localStorage.getItem(this.key + ":" + name); } catch (e) { return null; } }
};
let store = DEMO ? DemoStore : GraphStore;


const S={files:{}};
const REST_FILE="休日.json";
const isRest=k=>!!S.files[REST_FILE]?.doc?.休み?.includes(k);
function ruleWorks(ln, k) {
  if (holidayOf(k)) return ln.祝 === true;
  const w = dayDow(k);
  if (w === 0) return ln.日 === true;
  if (w === 6) return ln.土 !== false;
  return true;
}
// 段.休み／出勤（前の版の右クリックで入れた「段のその日だけ休み」）は 2026-10-08 から見ない。
// 直す画面が無く、残っていると意味の分からない斜線になるため（ユーザー：10/21 の斜線は何？）。工程ごとの休みは バー.休み
/* 工種（段）の区分と作業名（2026-10-08 ユーザー：「夜間」と書いている所を 作業区分（昼／夜）と 作業名 に分けたい）
   段.区分＝"昼"／"夜"、段.名前＝作業名。区分の無い前のデータは、名前の「昼間／夜間」から区分を決め、その字を除いた残りを作業名とみなす */
function laneKubun(ln) { return !ln ? "昼" : ln.区分 || (/夜/.test(ln.名前 || "") ? "夜" : "昼"); }
function laneWork(ln) { return !ln ? "" : ln.区分 ? (ln.名前 || "") : String(ln.名前 || "").replace(/昼間|夜間/g, "").trim(); }
function laneName(ln) { const w = laneWork(ln), k = laneKubun(ln); return w ? k + "　" + w : k + "間"; }
function laneWorks(ln, k) {
  if (!ln || isRest(k)) return false;
  return ruleWorks(ln, k);
}
// 工程（バー）ごとの休工日（バーの小窓の「休工日を設定」）。段の決まりの上にかける
function barWorks(ln, b, k) {
  if (isRest(k)) return false;
  if (b && (b.休み || []).includes(k)) return false;
  if (b && (b.出勤 || []).includes(k)) return true;
  return laneWorks(ln, k);
}
/* 作業員の人数（2026-10-14 ユーザー：「〇名＋〇名」と書きたい（中身の名前は書かない）・日ごとに変えたい）
   バー.人数＝合計（前の版でも読める）、バー.人数内訳＝"1+3"（＋で分けたときだけ）、
   バー.人数日＝{ "YYYY-MM-DD": "2+4" }（その日だけの人数。無い日はふだんの人数） */
function parseHead(t) {
  const s = String(t ?? "").replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/[＋*＊]/g, "+").replace(/[名人\s　]/g, "");
  if (!s) return null;
  const parts = s.split("+").filter(x => x !== "");
  if (!parts.length || parts.some(x => !/^\d+$/.test(x))) return null;
  // 0 の側は書かない（2026-10-16 ユーザー：「1」とだけ書いたものが「0名＋1名」と出た → 「1名」だけ）
  let n = parts.map(Number);
  if (n.length > 1 && n.some(x => x)) n = n.filter(x => x);
  return { txt: n.join("+"), parts: n, total: n.reduce((a, x) => a + x, 0) };
}
// 工種（段）にも人数を持つ（段.人数＝"1+3"。2026-10-14 ユーザー：現場の設定の工種の行で〇人＋〇人）。バーに人数が無ければ工種の人数
const headOwn = b => b.人数内訳 || (b.人数 !== "" && b.人数 != null ? String(b.人数) : "");
const headTxt = (b, ln) => headOwn(b) || (ln && ln.人数) || "";
const headTxtAt = (b, k, ln) => (b.人数日 && b.人数日[k] != null && b.人数日[k] !== "") ? b.人数日[k] : headTxt(b, ln);
const headAt = (b, k, ln) => { const p = parseHead(headTxtAt(b, k, ln)); return p ? p.total : 0; };
const fmtHead = t => { const p = parseHead(t); return p ? p.parts.map(x => x + "名").join("＋") : ""; };
const dayHeads = b => Object.keys(b.人数日 || {}).sort();
/* 工種は 作業員（自社）か 外注（2026-10-14 ユーザー：作業名の横で 外注 or 作業員(自社)。外注は業者名＋人数、作業員は〇人＋〇人）
   段.種別＝"外注"（無ければ自社）、段.業者＝業者名、段.人数＝"3" */
const isGai = ln => !!ln && ln.種別 === "外注";
/* 作業員と外注の共同作業（2026-11-11 ユーザー：作業員 何名 の隣に 外注 業者名 人数。両方入れられる）
   外注だけの工種は前のまま（段.種別＝"外注"・段.人数＝外注の人数）。両方のときは 段.人数＝作業員、段.業者・段.外注人数＝外注（種別は付けない）。
   バーの人数は、外注だけの工種なら外注の人数、それ以外は作業員の人数。共同の工種の外注の人数は工種の人数を使う */
const jointGai = ln => !!ln && !isGai(ln) && !!(ln.業者 || ln.外注人数) ? { 業者: ln.業者 || "外注", 人数: ln.外注人数 || "" } : null;
const gaiName = ln => isGai(ln) ? (ln.業者 || "外注") : jointGai(ln) ? jointGai(ln).業者 : "";
const hasGai = ln => isGai(ln) || !!jointGai(ln);
// その日の人数：作業員（自社）・外注
const ownAt = (b, k, ln) => isGai(ln) ? 0 : headAt(b, k, ln);
const gaiAt = (b, k, ln) => isGai(ln) ? headAt(b, k, ln) : jointGai(ln) ? ((parseHead(jointGai(ln).人数) || {}).total || 0) : 0;

// 独立APPの起動・共有データ読込。現場や休日の原本への保存機能は持たない。
const query=new URLSearchParams(location.search), upstreamVersion=query.get("sourceVersion") || "";
const upstreamMatch=/^Vr(\d+)/.exec(upstreamVersion);
const entryAllowed=DEMO || (query.get("from")==="kotei" && upstreamMatch && +upstreamMatch[1]>=2);
const sites=()=>Object.entries(S.files).filter(([n,f])=>n.startsWith("現場_") && f.doc && !f.doc.消した).map(([,f])=>f.doc).sort((a,b)=>(a.並び || 0)-(b.並び || 0));
let ready=false,pulling=false;
const uid=()=>"r"+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
function showSync(){const busy=Object.values(S.files).some(f=>f.saving);$("sync").className="sync "+(busy?"busy":"ok");$("sync").textContent=busy?"保存中…":"共有データ";}
function render(){if(ready)Bukai.render();}
async function pull(){
 if(!ready || pulling)return false;pulling=true;
 try{for(const item of await store.list()){
   if(!item.name.endsWith(".json") || item.name==="接続テスト.json")continue;
   const old=S.files[item.name];if(old && (old.eTag===item.eTag || old.saving))continue;
   const r=await store.get(item.name);if(!r || S.files[item.name]!==old || old?.saving)continue;
   S.files[item.name]={doc:r.data,eTag:r.eTag,pending:[],saving:false,timer:null,err:null};
 }render();showSync();return true;}catch(err){$("bkStatus").textContent="共有データを読めませんでした："+err.message;$("bkStatus").className="err";return false;}finally{pulling=false;}
}
let paper="A3";
const papers={A3:[420,297],A4:[297,210]},margin=7;
function pageStyle(){let el=$("bukaiPageStyle");if(!el){el=document.createElement("style");el.id="bukaiPageStyle";document.head.appendChild(el);}el.textContent=`@page { size: ${paper} landscape; margin: ${margin}mm; }`;}
function openKotei(id){const u=new URL("../",location.href);if(id)u.searchParams.set("site",id);if(DEMO)u.searchParams.set("demo","");window.open(u.href,"_blank","noopener");}
$("bHelp").onclick=()=>window.open("help.html#pc","_blank","noopener");
if(DEMO)$("backToKotei").href="../?demo";
Bukai.init({files:()=>S.files,sites,masters:()=>S.files["マスタ.json"]?.doc?.現場 || [],store:()=>store,who:()=>me.name,today:todayKey,tab:()=>"bukai",id:uid,sync:showSync,pull,works:barWorks,own:ownAt,outside:gaiAt,color:colorOf,openSite:openKotei,paper:value=>{paper=value;pageStyle();},printWidth:()=>(papers[paper][0]-margin*2)*96/25.4*0.98,printHeight:()=>(papers[paper][1]-margin*2)*96/25.4*0.96});
window.addEventListener("beforeprint",()=>{if(ready)Bukai.preparePrint();});
function afterPrint(){document.body.classList.remove("printing");$("printPages").innerHTML="";}
window.addEventListener("afterprint",afterPrint);
function seedDemo(){if(Object.keys(DemoStore.all()).length)return;const m=todayKey().slice(0,7),data={};for(let i=1;i<=3;i++){const id="demo"+i;data["現場_"+id+".json"]={eTag:"d"+i,data:{id,コード:"52990"+i,名前:"試験現場 "+i,担当:"見本 担当者",段:[{id:"l"+i,区分:"昼",名前:"管布設工",人数:"3"}],バー:[{id:"b"+i,段:"l"+i,作業:"管布設工",開始:m+"-02",終了:m+"-20",人数:"3",社員:["見本 担当者"],色:"青"}]}};}data["マスタ.json"]={eTag:"m",data:{社員:["見本 担当者"],現場:[]}};DemoStore.save(data);}
async function start(){
 if($("ver").textContent==="__VER__")$("ver").textContent="Vr001 開発中";
 if(!entryAllowed){$("entryOnly").hidden=false;return;}
 if(DEMO){seedDemo();me.name="試し";$("who").textContent="確認用（このブラウザ内だけに保存）";}
 else{
  // 既に登録済みの原版のURIへ戻し、MSALがこのAPPへ復帰する。
  msalApp=new msal.PublicClientApplication({auth:{clientId:CFG.clientId,authority:"https://login.microsoftonline.com/"+CFG.tenantId,redirectUri:new URL("../",location.href).href,navigateToLoginRequestUrl:true},cache:{cacheLocation:"localStorage"}});
  await msalApp.initialize();try{const r=await msalApp.handleRedirectPromise();if(r)msalApp.setActiveAccount(r.account);}catch(err){$("signinMsg").textContent=err.message;}
  const account=msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if(!account){$("signin").hidden=false;$("bSignIn").onclick=()=>msalApp.loginRedirect({scopes:SCOPES,redirectStartPage:location.href});return;}
  msalApp.setActiveAccount(account);me.name=account.name || account.username;$("who").textContent=me.name;
 }
 ready=true;$("loading").hidden=false;const ok=await pull();$("loading").hidden=true;$("pBukai").hidden=false;if(ok)$("bkStatus").textContent="全体工程表と工程データを共有しています";pageStyle();render();setInterval(()=>{if(!document.hidden)pull();},POLL_MS);
}
start().catch(err=>{$("signin").hidden=false;$("signinMsg").textContent="起動できませんでした："+err.message;});
