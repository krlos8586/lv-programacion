// ==UserScript==
// @name         LV TowerNex · actualizador
// @namespace    lps-lv-towernex
// @version      1.0.0
// @description  Actualiza la página de líneas de vida LPS/TowerNex con Sitetracker y el SharePoint de TowerNex usando tu propia sesión. SOLO LECTURA: no crea, edita ni borra nada en TowerNex.
// @author       LPS Grupo
// @match        https://krlos8586.github.io/lv-programacion/*
// @match        https://towernex.my.site.com/scp/*
// @run-at       document-idle
// @noframes
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// @grant        GM_openInTab
// @connect      towernexcolombiasas.sharepoint.com
// @updateURL    https://krlos8586.github.io/lv-programacion/lv-actualizador.user.js
// @downloadURL  https://krlos8586.github.io/lv-programacion/lv-actualizador.user.js
// ==/UserScript==
(function(){
'use strict';
const VERSION='1.0.0';
const SP='https://towernexcolombiasas.sharepoint.com/sites/TowerNex-ContratistasOM';
const ST_URL='https://towernex.my.site.com/scp/s/';

/* ---------- 1. En la página de líneas de vida: puente para que el tablero lea las fuentes ---------- */
if(location.hostname==='krlos8586.github.io'){
  document.documentElement.setAttribute('data-lvgm',VERSION);
  const reply=(src,id,ok,data)=>{ try{ src.postMessage({lv:'gm-res',id,ok,data},'*'); }catch(e){} };
  window.addEventListener('message',ev=>{
    const m=ev.data; if(!m||m.lv!=='gm'||!ev.source) return;
    if(ev.origin!==location.origin && ev.origin!=='null') return;          // solo el propio tablero
    if(m.op==='ping') return reply(ev.source,m.id,true,{version:VERSION});
    if(m.op==='sp') return spDocs().then(d=>reply(ev.source,m.id,true,d)).catch(e=>reply(ev.source,m.id,false,String(e&&e.message||e)));
    if(m.op==='st') return stRun().then(d=>reply(ev.source,m.id,true,d)).catch(e=>reply(ev.source,m.id,false,String(e&&e.message||e)));
  });
  // SharePoint: solo GET a la búsqueda del sitio de contratistas O&M (carpeta LPS)
  const spGet=url=>new Promise((res,rej)=>GM_xmlhttpRequest({method:'GET',url,headers:{Accept:'application/json;odata=nometadata'},timeout:60000,
    onload:r=>{ if(r.status===401||r.status===403||/login\.microsoftonline/.test(r.finalUrl||'')) return rej(new Error('login-sp')); if(r.status>=400) return rej(new Error('SP '+r.status));
      try{ res(JSON.parse(r.responseText)); }catch(e){ rej(new Error('login-sp')); } },
    onerror:()=>rej(new Error('red-sp')), ontimeout:()=>rej(new Error('timeout-sp'))}));
  async function spDocs(){
    const q=start=>SP+"/_api/search/query?querytext='"+encodeURIComponent('path:"'+SP+'/LPS" IsDocument:true FileExtension:pdf')+"'&selectproperties='Path,FileType,LastModifiedTime'&rowlimit=500&startrow="+start+"&trimduplicates=false";
    let all=[];
    for(let s=0;s<20000;s+=500){ const j=await spGet(q(s)); const rows=j.PrimaryQueryResult.RelevantResults.Table.Rows.map(x=>Object.fromEntries(x.Cells.map(c=>[c.Key,c.Value])));
      all=all.concat(rows); if(rows.length<500) break; await new Promise(z=>setTimeout(z,800)); }
    const P=SP+'/LPS/', out={};
    for(const x of all){ const p=String(x.Path||'').replace(P,''); if(/CALIBRACI|ACTAS DE APROB|CARTAS GENERADAS|05-EQUIPOS|^CONSOLIDADO/i.test(p)||!/^pdf$/i.test(x.FileType||'')) continue;
      const m=p.match(/CO-[A-Z]{3}\d{4}/); if(!m) continue; const f=p.slice(p.lastIndexOf('/')+1);
      const k=/carta de responsabilidad/i.test(f)?'K':/PROTOCOLO/i.test(f)?'P':/RECERTIFICACION/i.test(f)?'R':/CERTIFICADO DE INSTALACI/i.test(f)?'I':/GARANTIA/i.test(f)?'G':/SUMINISTRO|Certificado LV|Certificado_ LV|Certificado producto/i.test(f)?'F':null;
      if(!k) continue; (out[m[0]]=out[m[0]]||[]).push([k,p,String(x.LastModifiedTime||'').slice(0,10)]); }
    Object.values(out).forEach(v=>v.sort((a,b)=>a[2]<b[2]?1:a[2]>b[2]?-1:(a[1]<b[1]?-1:1)));
    return {total:all.length, docs:out};
  }
  // Sitetracker: abre una pestaña en segundo plano con tu sesión, lee y se cierra
  function stRun(){ return new Promise((res,rej)=>{
    const id=Date.now()+'-'+Math.random().toString(36).slice(2,7); let tab=null, done=false;
    const fin=(f,v)=>{ if(done) return; done=true; try{ GM_removeValueChangeListener(lid); }catch(e){} try{ tab&&tab.close(); }catch(e){} GM_deleteValue('lv_req'); f(v); };
    const lid=GM_addValueChangeListener('lv_res',(n,o,v)=>{ if(!v||v.id!==id) return; GM_deleteValue('lv_res'); v.error?fin(rej,new Error(v.error)):fin(res,v.data); });
    GM_setValue('lv_req',{id,ts:Date.now()});
    tab=GM_openInTab(ST_URL+'?lv='+id,{active:false,insert:true,setParent:true});
    setTimeout(()=>fin(rej,new Error('timeout-st')),120000);
  }); }
  return;
}

/* ---------- 2. En Sitetracker: solo actúa si el tablero lo pidió hace menos de 2 minutos ---------- */
if(location.hostname==='towernex.my.site.com'){
  const req=GM_getValue('lv_req'); if(!req||Date.now()-req.ts>120000) return;
  if(!new URLSearchParams(location.search).get('lv')) return;
  const send=o=>GM_setValue('lv_res',Object.assign({id:req.id},o));
  if(/\/login/i.test(location.pathname)||document.querySelector('input[type=password]')) return send({error:'login-st'});
  window.addEventListener('message',ev=>{ const m=ev.data; if(!m||m.lvst!==req.id) return; m.error?send({error:m.error}):send({data:m.data}); });
  // corre dentro de la página (necesita la sesión); solo usa la acción de LECTURA getListRecordsById
  const code=`(async()=>{const ID=${JSON.stringify(req.id)};const post=o=>window.postMessage(Object.assign({lvst:ID},o),'*');
  try{let t=0;while(!(window.$A&&$A.clientService)&&t<60){await new Promise(r=>setTimeout(r,500));t++}
   if(!window.$A)throw new Error('sin-aura');
   let tok=null;for(let i=0;i<40&&!tok;i++){for(const k of Object.keys($A.clientService)){const v=$A.clientService[k];if(typeof v==='string'&&v.length>40&&/^eyJ/.test(v)){tok=v;break}}if(!tok)await new Promise(r=>setTimeout(r,500))}
   if(!tok)throw new Error('sin-token');
   const h=await (await fetch('/scp/s/',{credentials:'include'})).text();
   const fw=(h.match(/"fwuid":"([^"]+)"/)||[])[1],app=(h.match(/"app":"([^"]+)"/)||[])[1],ld=(h.match(/"loaded":(\\{[^}]*\\})/)||[])[1];
   const ctx=JSON.stringify({mode:'PROD',fwuid:fw,app:app,loaded:ld?JSON.parse(ld):{},dn:[],globals:{},uad:true});
   let seq=300;async function call(p){const b=new URLSearchParams();b.set('message',JSON.stringify({actions:[{id:(++seq)+';a',descriptor:'aura://ListUiController/ACTION$getListRecordsById',callingDescriptor:'UNKNOWN',params:p}]}));b.set('aura.context',ctx);b.set('aura.pageURI','/scp/s/');b.set('aura.token',tok);
    const r=await fetch('/scp/s/sfsites/aura?r='+seq+'&aura.ListUiController.getListRecordsById=1',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded; charset=UTF-8'},body:b.toString(),credentials:'include'});
    const j=JSON.parse((await r.text()).replace(/^while\\(1\\);/,''));if(!j.actions||j.actions[0].state!=='SUCCESS')throw new Error('consulta-rechazada');return j.actions[0].returnValue}
   const P='sitetracker__',J='sitetracker__Job__c.',Sv='sitetracker__Site__c.';
   const jf=['Name',P+'Job_Type__c',P+'Customer_Comments__c',P+'Actual_Start__c',P+'Actual_End__c','Report_Date__c',P+'Job_Status__c',P+'Site__r.Name'];
   const sf=['Name','Lifeline__c','LifelineVendor__c','Lifeline_Manufacture_Date__c','Lifeline_Certification_Date__c'];
   const disp=(rec,f)=>{const ps=f.split('.');let o=rec.fields[ps[0]];for(const p of ps.slice(1)){o=o&&o.value&&o.value.fields&&o.value.fields[p]}return o?(o.displayValue??o.value):null};
   const raw=(rec,f)=>{const o=rec.fields[f];return o?o.value:null};
   const day=v=>{if(!v)return '';const s=String(v);return s.includes('T')?new Date(s).toLocaleDateString('en-CA',{timeZone:'America/Bogota'}):s.slice(0,10)};
   const rj=await call({listViewId:'00Bbn0000099k01EAA',pageSize:2000,optionalFields:jf.map(x=>J+x)});
   const rs=await call({listViewId:'00Bbn0000099Vc5EAE',pageSize:2000,optionalFields:sf.map(x=>Sv+x)});
   const jobs=rj.records.map(rc=>({j:disp(rc,'Name'),site:disp(rc,P+'Site__r.Name'),type:raw(rc,P+'Job_Type__c')||'',est:disp(rc,P+'Job_Status__c')||'',ini:day(raw(rc,P+'Actual_Start__c')),fin:day(raw(rc,P+'Actual_End__c')),carga:day(raw(rc,'Report_Date__c')),com:disp(rc,P+'Customer_Comments__c')||''}));
   const sitios=rs.records.map(rc=>({id:disp(rc,'Name'),lv:disp(rc,'Lifeline__c')||'',ven:disp(rc,'LifelineVendor__c')||'',inst:day(raw(rc,'Lifeline_Manufacture_Date__c')),next:day(raw(rc,'Lifeline_Certification_Date__c'))}));
   post({data:{generado:new Date().toISOString(),jobs,sitios}});
  }catch(e){post({error:String(e&&e.message||e)})}})();`;
  const s=document.createElement('script'); s.src=URL.createObjectURL(new Blob([code],{type:'text/javascript'})); (document.head||document.documentElement).appendChild(s);
  setTimeout(()=>send({error:'timeout-st-pagina'}),100000);
}
})();
