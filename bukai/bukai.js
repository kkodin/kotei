/* 全体工程表と同じログイン・保存先を使う部会資料画面。 */
window.Bukai = (() => {
  "use strict";
  const C=window.BukaiCore, $=id=>document.getElementById(id), e=s=>String(s ?? "").replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  let ctx, month, active="", archive="", busy=false, dirty=false, draft=null, importPlan=null, filter="", displayed="";
  const reportName=()=>`部会資料_${month}.json`;
  const empty=()=>({kind:"bukai-report",schemaVersion:1,yearMonth:month,entries:{}});
  const doc=()=>ctx.files()[reportName()]?.doc || empty();
  const sourceRows=()=>C.rows(ctx.sites(),ctx.masters());
  const archived=()=>archive ? ctx.files()[archive]?.doc : null;
  const rowList=()=>archived()?.rows || sourceRows();
  const resolved=r=>archived() ? r.values : C.values(r,doc());
  const included=r=>archived() ? true : C.selected(r,doc());
  function setBusy(value) { busy=value;$('pBukai').inert=value;$('bkImportApply').disabled=value; }
  function status(text,err=false) { $("bkStatus").textContent=text; $("bkStatus").className=err?"err":""; }
  function canLeave() { return !busy && (!dirty || confirm("未保存の入力があります。保存せずに移動しますか？")); }
  function stamp() { return {at:new Date().toISOString(),by:ctx.who()}; }
  function cache(name,data,eTag) { ctx.files()[name]={doc:C.clone(data),eTag,pending:[],saving:false,timer:null,err:null};ctx.sync(); }
  async function commit(changes) {
    if (busy) throw Error("保存中です。完了を待ってください。");
    setBusy(true); status("保存中…");
    const name=reportName();
    // 新規作成時も再読込から保護する（保存完了まで表示用データは変更しない）。
    const locked=ctx.files()[name] ||= {doc:empty(),eTag:null,pending:[],saving:false,timer:null,err:null};
    locked.saving=true;
    ctx.sync();
    try {
      for (let n=0;n<3;n++) {
        const fresh=await ctx.store().get(name), latest=fresh?.data || empty();
        if (latest.kind!=="bukai-report" || latest.yearMonth!==month) throw Error("保存先の資料形式を確認してください。");
        const next=C.patch(latest,changes);next.updated=stamp();
        try { const saved=await ctx.store().put(name,next,fresh?.eTag || null); cache(name,next,saved.eTag);status("共有保存済み");return; }
        catch(err) { if (![409,412].includes(err.status)) throw err; }
      }
      throw Error("ほかの人の保存が続いています。少し待って保存し直してください。");
    } finally { setBusy(false); if (locked) locked.saving=false;ctx.sync(); }
  }
  function onError(err) { console.error(err);status(err.message,true); }
  function archives() { return Object.entries(ctx.files()).filter(([,f])=>f.doc?.kind==="bukai-snapshot" && f.doc.yearMonth===month).sort((a,b)=>String(b[1].doc.created.at).localeCompare(String(a[1].doc.created.at))); }
  function render() {
    if (!ctx || ctx.tab()!=="bukai") return;
    try {
      const rows=rowList(), d=doc();
      $("bkMonth").value=month;
      const options=archives();
      $("bkVersion").innerHTML='<option value="">編集中（最新の工程を参照）</option>'+options.map(([name,f])=>`<option value="${e(name)}">配布版 ${e(localTime(f.doc.created.at))}　${e(f.doc.created.by)}</option>`).join("");
      $("bkVersion").value=archive;
      $("bkImport").disabled=!!archive || busy; $("bkSnapshot").disabled=!!archive || busy;
      const selected=rows.filter(included);
      $("bkCount").textContent=`掲載 ${selected.length}現場／工程期間 ${C.period(month).from} 〜 ${C.period(month).end}`;
      if (!rows.some(r=>r.key===active)) active=selected[0]?.key || rows[0]?.key || "";
      $("bkSites").innerHTML=rows.filter(r=>[r.name,r.code,resolved(r).handler].join(" ").includes(filter)).map(r=>`<div class="bk-site ${r.key===active?'active':''}"><input type="checkbox" aria-label="${e(r.name)}を掲載" data-include="${e(r.key)}" ${included(r)?'checked':''} ${archive?'disabled':''}><button type="button" data-row="${e(r.key)}"><small>${e(r.code || 'コード未登録')} ${!archive && !r.siteId?'台帳のみ':''}</small>${e(r.name)}</button></div>`).join("");
      if (dirty) return; // 他の人の更新・30秒の再読込で入力欄を作り直さない。
      const r=rows.find(x=>x.key===active);
      if (!r) { $("bkEditor").innerHTML='<p class="bk-empty">工程表に現場を登録すると、部会資料にも表示されます。</p>';return; }
      const v=resolved(r), item=(d.entries || {})[r.key] || {};
      draft={key:r.key,display:C.clone(v),item:C.clone(item)};displayed=active;
      $("bkEditor").innerHTML=`<h2>${e(r.name)} <span class="note">${e(r.code)}</span></h2>
        <div class="note">${archive?'配布版を表示しています。保存時点の内容です。':'現場台帳・工程表の記入済みデータを参照します。この画面の追記は選択した会議月に保存されます。'}</div>
        <form id="bkForm"><div class="bk-form">${Object.entries(C.fields).map(([f,label])=>{
          const wide=["summary","remarks"].includes(f), num=["plan","execPercent","amount"].includes(f), val=v[f] ?? "";
          return `<label class="${wide?'wide':''}">${e(label)}${wide?`<textarea name="${f}" rows="3" maxlength="4000" ${archive?'readonly':''}>${e(val)}</textarea>`:`<input name="${f}" type="${num?'number':['start','end'].includes(f)?'date':'text'}" ${num?`min="0" step="any" ${f!=='amount'?'max="100"':''}`:'maxlength="4000"'} value="${e(val)}" ${archive?'readonly':''}>`}<span class="bk-origin">${archive?'配布時点':Object.hasOwn(item.values || {},f)?'この会議月の記入':val!==''?'共有データから':'未記入'}</span></label>`;
        }).join("")}</div>${archive?'':`<div class="bk-actions"><button type="submit">この現場を保存</button><button type="button" id="bkInherit" class="sub">追記を外して共有データに戻す</button><span id="bkDirty" class="note">保存済みの内容を表示</span></div>`}</form>
        <div class="bk-source"><h3>工程表と共有する工程</h3><p class="note">工程線・人数・休工日は全体工程表と共通です。部会資料への二重入力は不要です。</p>
        ${!archive && r.siteId?'<button type="button" id="bkOpenChart" class="sub small">全体工程表で開く</button>':''}
        ${!archive?(r.source.段 || []).map(l=>`<label class="bk-lane"><span>${e(l.名前 || '工種')}</span><select data-category="${e(l.id)}"><option value="直労">直労</option><option value="外注">外注</option><option value="機材">機材</option></select></label>`).join(''):''}
        <div class="bk-source-list">${(archive?r.bars:(r.source.バー || [])).map(b=>`<div>${e(b.開始 || b.from)} 〜 ${e(b.終了 || b.end)}　${e(b.作業 || b.work)}</div>`).join('') || '<span class="note">工程は未登録です。現場台帳の情報だけでも資料に掲載できます。</span>'}</div></div>`;
      $("bkEditor").querySelectorAll('[data-category]').forEach(sel=>{ const ln=r.source.段.find(l=>l.id===sel.dataset.category);sel.value=item.categories?.[ln.id] || (ln.種別==='外注'?'外注':'直労'); });
      if (!archive) {
        $("bkForm").onsubmit=saveEntry;
        $("bkInherit").onclick=inherit;
        if ($("bkOpenChart")) $("bkOpenChart").onclick=()=>{if(canLeave()){dirty=false;ctx.openSite(r.siteId);}};
      }
      $("bkEditor").oninput=()=>{if(!archive){dirty=true;if($("bkDirty"))$("bkDirty").textContent='未保存の入力があります';}};
      $("bkEditor").onchange=$("bkEditor").oninput;
    } catch(err) { $("bkEditor").innerHTML=`<p class="ng">${e(err.message)}</p>`;onError(err); }
  }
  async function saveEntry(event) {
    event.preventDefault();if(busy || archive)return;
    try {
      const changes=[], data=new FormData($("bkForm"));
      for(const f of Object.keys(C.fields)) { const v=C.validate(f,data.get(f));if(!C.equal(v,draft.display[f] ?? ''))changes.push({key:draft.key,area:'values',field:f,before:draft.item.values?.[f],after:v}); }
      const categories={...(draft.item.categories || {})};
      $("bkEditor").querySelectorAll('[data-category]').forEach(s=>categories[s.dataset.category]=s.value);
      if(!C.equal(categories,draft.item.categories || {}))changes.push({key:draft.key,area:'categories',before:draft.item.categories,after:categories});
      const dataEnd=data.get('end'),dataStart=data.get('start');if(dataStart && dataEnd && dataStart>dataEnd)throw Error('工期の開始日が終了日より後です。');
      if(changes.length)await commit(changes);
      dirty=false;render();status('共有保存済み');
    } catch(err) {onError(err);}
  }
  async function inherit() {
    if(busy || archive || !confirm('この現場の会議月の追記を外し、共有データの値に戻しますか？'))return;
    try {await commit(Object.keys(draft.item.values || {}).map(field=>({key:draft.key,area:'values',field,before:draft.item.values[field],after:null})));dirty=false;render();}catch(err){onError(err);}
  }
  function localTime(t) {const d=new Date(t);return isNaN(d)?t:d.toLocaleString('ja-JP');}
  function capture() {
    const p=C.period(month), d=doc();
    return {kind:'bukai-snapshot',schemaVersion:1,yearMonth:month,created:stamp(),period:p,
      rows:sourceRows().filter(r=>C.selected(r,d)).map(r=>{
        const cats=(d.entries || {})[r.key]?.categories || {};
        return {key:r.key,code:r.code,name:r.name,values:C.values(r,d),bars:(r.source.バー || []).filter(b=>b.開始<=p.end && b.終了>=p.from).map(b=>{
          const ln=(r.source.段 || []).find(l=>l.id===b.段), days=p.days.filter(k=>k>=b.開始 && k<=b.終了 && ctx.works(ln,b,k));
          return {from:b.開始,end:b.終了,work:b.作業 || ln?.名前 || '',category:cats[b.段] || (ln?.種別==='外注'?'外注':'直労'),color:b.色 || '青',days,
            numbers:Object.fromEntries(days.map(k=>[k,{own:ctx.own(b,k,ln),outside:ctx.outside(b,k,ln),staff:[...(b.社員 || [])]}]))};
        })};
      })};
  }
  async function snapshot() {
    if(busy || archive)return;
    if(dirty){status('配布版を作る前に、この現場の入力を保存してください。',true);return;}
    if(Object.values(ctx.files()).some(f=>f.pending?.length || f.saving || f.err)){status('工程表の保存が完了してから配布版を作ってください。',true);return;}
    try {
      const snap=capture();if(!snap.rows.length)throw Error('掲載する現場を選んでください。');
      if(!confirm(`${month} の ${snap.rows.length}現場を配布版として保存します。保存後も元の工程は編集できます。`))return;
      setBusy(true);status('配布版を保存中…');
      const name=`部会配布版_${month}_${ctx.id()}.json`, r=await ctx.store().put(name,snap,null);
      cache(name,snap,r.eTag);archive=name;dirty=false;render();status('配布版を保存しました');
    }catch(err){onError(err);}finally{setBusy(false);render();}
  }
  function timeline(row,p) {
    const W=640, dayW=594/p.days.length, rows=['計画','外注','直労','機材'], bars=row.bars || [], gap=20;
    const grouped=rows.map(cat=>bars.filter(b=>cat==='計画' || b.category===cat));
    const heights=grouped.map(bs=>Math.max(28,bs.length*gap+4)), height=25+heights.reduce((a,b)=>a+b,0);
    let y=25;
    let s=`<svg class="bk-timeline" viewBox="0 0 ${W} ${height}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="3か月の工程"><rect width="${W}" height="${height}" fill="white"/>`;
    p.days.forEach((k,i)=>{const x=46+i*dayW;if(k.endsWith('-01') || i===0)s+=`<line x1="${x}" x2="${x}" y1="0" y2="${height}" stroke="#555"/><text x="${x+2}" y="11" font-size="10">${+k.slice(5,7)}月</text>`;if([1,10,20].includes(+k.slice(8)))s+=`<text x="${x+1}" y="23" font-size="8">${+k.slice(8)}</text>`;if([0,6].includes(new Date(k+'T00:00:00').getDay()))s+=`<rect x="${x}" y="25" width="${dayW}" height="${height-25}" fill="#f0f0f0"/>`;});
    rows.forEach((cat,idx)=>{s+=`<line x1="0" x2="640" y1="${y}" y2="${y}" stroke="#888"/><text x="1" y="${y+16}" font-size="10">${cat}</text>`;
      grouped[idx].forEach((b,j)=>{const yy=y+4+j*gap,a=Math.max(0,p.days.indexOf(b.from)),start=b.from<p.from?0:a;const z=b.end>p.end?p.days.length-1:p.days.indexOf(b.end);if(z<0)return;
        s+=`<rect x="${46+start*dayW}" y="${yy}" width="${Math.max(dayW,(z-start+1)*dayW)}" height="6" fill="white" stroke="${e(ctx.color(b.color))}" stroke-width="0.6"/>`;
        for(const k of b.days){const n=p.days.indexOf(k);if(n>=0)s+=`<rect x="${46+n*dayW}" y="${yy}" width="${dayW+0.1}" height="6" fill="${e(ctx.color(b.color))}"/>`;}
        const counts=Object.values(b.numbers || {}), range=field=>{const ns=counts.map(n=>n[field]).filter(Number.isFinite);return !ns.length || !Math.max(...ns)?'':Math.min(...ns)===Math.max(...ns)?Math.max(...ns)+'名':Math.min(...ns)+'〜'+Math.max(...ns)+'名';};
        const own=range('own'),outside=range('outside'), staff=[...new Set(counts.flatMap(n=>n.staff || []))];
        const label=[b.work,own?'自社 '+own:'',outside?'外注 '+outside:'',staff.join('・')].filter(Boolean).join('　');
        s+=`<text x="${46+start*dayW}" y="${yy+16}" font-size="8"><title>${e(label)}</title>${e(label)}</text>`;
      });y+=heights[idx];
    });return s+'</svg>';
  }
  function pageHTML(snap,rows,page,total) {
    const money=v=>v==='' || v==null?'':Number(v).toLocaleString('ja-JP')+'円', pct=v=>v==='' || v==null?'':v+'％';
    return `<div class="bk-print-page"><div class="bk-print-title"><b>工事工程計画及び実施報告書　${e(snap.yearMonth)} 工事部会</b><small>${e(archive?'配布版 '+localTime(snap.created.at):'編集中・'+localTime(snap.created.at))}</small></div><table class="bk-report"><colgroup>${[110,130,65,195,45,45,150,150,660].map(w=>`<col style="width:${w}px">`).join('')}</colgroup><thead><tr>${['工事略称・コード','発注者・工期','担当','工事概要','計画','実行','備考','管径・延長・数量等','工程（計画・外注・直労・機材）'].map(s=>`<th>${s}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>{const v=r.values;return `<tr><td>${e(r.code)}<b>${e(r.name)}</b>${e(money(v.amount))}</td><td>${e(v.client)}\n${e(v.start)}\n〜 ${e(v.end)}\n${e(v.phone)}</td><td>${e(v.handler)}</td><td>${e(v.summary)}</td><td class="bk-percent">${e(pct(v.plan))}</td><td class="bk-percent">${e(pct(v.execPercent))}</td><td>${e(v.remarks)}</td><td>${e([v.pipeDia,v.length,v.count && '数量 '+v.count,v.soil && '土質 '+v.soil].filter(Boolean).join('\n'))}</td><td>${timeline(r,snap.period)}</td></tr>`;}).join('')}</tbody></table><div class="bk-print-footer"><span>工程期間 ${e(snap.period.from)} 〜 ${e(snap.period.end)}　白抜きは休工日</span><span>${page} / ${total}</span></div></div>`;
  }
  function preparePrint() {
    const snap=archived() || capture();
    const out=$('printPages');out.innerHTML='';
    const n=+$('bkPerPage').value || 4,total=Math.ceil(snap.rows.length/n);
    for(let i=0;i<total;i++){const page=document.createElement('div');page.className='ppage';page.innerHTML=pageHTML(snap,snap.rows.slice(i*n,(i+1)*n),i+1,total);out.appendChild(page);const sheet=page.firstElementChild;sheet.style.zoom=Math.min(1,ctx.printWidth()/1550,ctx.printHeight()/sheet.scrollHeight).toFixed(4);}
    document.body.classList.add('printing');
  }
  function print() {
    if(dirty){status('入力内容を保存してから印刷してください。',true);return;}
    try {if(!(archived() || capture()).rows.length)throw Error('掲載する現場を選んでください。');ctx.paper($('bkPaper').value);window.print();}catch(err){onError(err);}
  }
  async function readImport(event) {
    const file=event.target.files[0];event.target.value='';if(!file || busy || archive)return;
    if(dirty){status('入力を保存してからJSONを取り込んでください。',true);return;}
    try {if(file.size>2*1024*1024)throw Error('JSONは2MB以内にしてください。');const data=JSON.parse(await file.text());importPlan=C.importData(data,month,sourceRows(),doc());$('bkImportError').textContent='';
      $('bkImportPreview').innerHTML=`<p>${e(month)} の ${importPlan.preview.length}現場へ反映します。指定された空文字は空欄として反映します。</p><table class="bk-preview"><tr><th>現場</th><th>反映する内容</th></tr>${importPlan.preview.map(x=>`<tr><td>${e(x.row.code)} ${e(x.row.name)}</td><td>${importPlan.changes.filter(p=>p.key===x.row.key && p.area==='values').map(p=>`${e(C.fields[p.field])}: ${e(C.values(x.row,doc())[p.field] ?? '')} → ${e(p.after)}`).join('\n')}</td></tr>`).join('')}</table>`;
      $('dBukaiImport').showModal();
    }catch(err){onError(err);}
  }
  function exportJSON() {
    if(dirty){status('入力を保存してから書き出してください。',true);return;}
    try {
      const snap=archived() || capture(), items=snap.rows.map(r=>{
        if(!r.code)throw Error('コード未登録の現場があります。現場コードを設定してから書き出してください。');
        const v=r.values;
        return {projectCode:r.code,client:v.client,handler:v.handler,summary:v.summary,plan:v.plan,execPercent:v.execPercent,remarks:v.remarks,start:v.start,end:v.end,
          metadata:{pipeDia:v.pipeDia,length:v.length,count:v.count,soil:v.soil,amount:v.amount,phone:v.phone}};
      });
      if(!items.length)throw Error('掲載する現場を選んでください。');
      const payload={yearMonth:month,updatedAt:snap.created.at,updatedBy:snap.created.by,items};
      const url=URL.createObjectURL(new Blob([JSON.stringify(payload,null,2)],{type:'application/json'})),a=document.createElement('a');a.href=url;a.download=`工事部会資料_${month}.json`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
      status('共有用のJSONを書き出しました');
    }catch(err){onError(err);}
  }
  function init(bridge) {
    ctx=bridge;month=ctx.today().slice(0,7);
    $('bkMonth').onchange=ev=>{if(!canLeave()){ev.target.value=month;return;}try{C.period(ev.target.value);month=ev.target.value;archive='';active='';dirty=false;render();status('会議月を変更しました');}catch(err){ev.target.value=month;onError(err);}};
    $('bkVersion').onchange=ev=>{if(!canLeave()){ev.target.value=archive;return;}archive=ev.target.value;active='';dirty=false;render();};
    $('bkFilter').oninput=ev=>{filter=ev.target.value;render();};
    $('bkSites').onclick=ev=>{const b=ev.target.closest('[data-row]');if(!b || b.dataset.row===active || !canLeave())return;active=b.dataset.row;dirty=false;render();};
    $('bkSites').onchange=async ev=>{const box=ev.target.closest('[data-include]');if(!box)return;const key=box.dataset.include,old=doc().entries?.[key]?.included;try{await commit([{key,area:'included',before:old,after:box.checked}]);render();}catch(err){box.checked=!box.checked;onError(err);}};
    $('bkRefresh').onclick=async()=>{if(!canLeave())return;try{dirty=false;if(!await ctx.pull())throw Error('共有データを読み直せませんでした。通信状態を確認して、もう一度読み直してください。');render();status('最新の共有データを読み直しました');}catch(err){onError(err);}};
    $('bkSnapshot').onclick=snapshot;$('bkPrint').onclick=print;$('bkExport').onclick=exportJSON;$('bkImport').onclick=()=>$('bkFile').click();$('bkFile').onchange=readImport;
    $('bkImportCancel').onclick=()=>{$('dBukaiImport').close();importPlan=null;};
    $('bkImportApply').onclick=async()=>{if(!importPlan || busy)return;try{await commit(importPlan.changes);$('dBukaiImport').close();importPlan=null;render();}catch(err){onError(err);$('bkImportError').textContent=err.message;}};
    window.addEventListener('beforeunload',ev=>{if(dirty || busy){ev.preventDefault();ev.returnValue='';}});
  }
  return {init,render,canLeave,preparePrint,isDirty:()=>dirty};
})();
