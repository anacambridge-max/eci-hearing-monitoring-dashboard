'use client';
import { useEffect } from 'react';

const clean=(v:any)=>String(v??'').replace(/\s+/g,' ').trim();
const num=(v:any)=>{const n=Number(v);return Number.isFinite(n)?Math.max(n,0):0};
const person=(v:any)=>{const x=clean(v);const m=x.match(/^(.*?)\s*\(\s*(\d{10})\s*\)\s*$/);return m?{name:m[1].trim(),phone:m[2]}:{name:x,phone:''};};

type Hearing={status:string;date:string};
type Row={ps:number;oldPs:any;bloName:string;bloPhone:string;supervisor:string;supervisorPhone:string;officer:string;total:number;generated:number;delivered:number;pending:number;hearingStatus:string;hearingDate:string};

function getHearingMap():Record<string,Hearing>{try{return JSON.parse(localStorage.getItem('eci-hearing-dates')||'{}')}catch{return{}}}

function getRows():Row[]{
  let master:any[]=[],eci:any[]=[];
  try{master=JSON.parse(localStorage.getItem('eci-master')||'[]');eci=JSON.parse(localStorage.getItem('eci-current')||'[]')}catch{return[]}
  if(!master.length||!eci.length)return[];
  const em=new Map<number,any>(eci.map(r=>[Number(r?.ps),r])),hm=getHearingMap();
  return master.map(m=>{
    const ps=num(m?.ps),e=em.get(ps);if(!e)return null;
    const total=num(m?.mapping),rawGen=num(e?.['Notice Generated']),rawDel=num(e?.['Notice Delivered']);
    const generated=Math.min(rawGen,total),delivered=Math.min(rawDel,generated),pending=Math.max(generated-delivered,0);
    if(generated===0&&delivered===0)return null;
    const blo=person(m?.blo),sup=person(m?.supervisor),h=hm[String(ps)]||{status:'',date:''};
    return{ps,oldPs:m?.oldPs,bloName:blo.name,bloPhone:blo.phone,supervisor:sup.name,supervisorPhone:sup.phone||'',officer:person(m?.officer).name,total,generated,delivered,pending,hearingStatus:h.status,hearingDate:h.date};
  }).filter(Boolean).sort((a:any,b:any)=>b.generated-a.generated||b.delivered-a.delivered||a.ps-b.ps) as Row[];
}

function findHeader(a:any[][],names:string[]):{row:number,col:number}|null{
  const wanted=names.map(x=>clean(x).toUpperCase());
  for(let r=0;r<Math.min(a.length,20);r++)for(let c=0;c<(a[r]||[]).length;c++)if(wanted.includes(clean(a[r][c]).toUpperCase()))return{row:r,col:c};
  return null;
}

async function loadHearingFile(file:File){
  const XLSX=await import('xlsx');
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
  const preferred=['All Officers Consolidated','Part Wise Hearing Summary'];
  const sheet=preferred.map(n=>wb.Sheets[n]).find(Boolean)||wb.Sheets[wb.SheetNames.find(s=>/consolidated|part\s*wise/i.test(s))||wb.SheetNames[0]];
  if(!sheet)throw new Error('Sheet not found');
  const a:any[][]=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''});
  const psH=findHeader(a,['PS No.','PS No','Part No.','Part No','Part Number']);
  const dateH=findHeader(a,['Hearing Date(s)','Hearing Date','Hearing Dates']);
  const statusH=findHeader(a,['Hearing Status','Status']);
  if(!psH||!dateH)throw new Error('PS/Part No. or Hearing Date column not found');
  const headerRow=Math.min(psH.row,dateH.row);
  const out:Record<string,Hearing>={};
  const dates:Record<string,string[]>={};
  for(let r=headerRow+1;r<a.length;r++){
    const rawPs=clean(a[r]?.[psH.col]).replace(/\.0$/,'');
    const ps=Number(rawPs);
    if(!Number.isFinite(ps)||ps<=0)continue;
    const date=clean(a[r]?.[dateH.col]);
    const status=statusH?clean(a[r]?.[statusH.col]):'';
    if(!dates[String(ps)])dates[String(ps)]=[];
    if(date&&!dates[String(ps)].includes(date))dates[String(ps)].push(date);
    out[String(ps)]={status:status||'Scheduled',date:dates[String(ps)].join(', ')};
  }
  for(const ps of Object.keys(dates))out[ps]={status:out[ps]?.status||'Scheduled',date:dates[ps].join(', ')};
  localStorage.setItem('eci-hearing-dates',JSON.stringify(out));
  window.dispatchEvent(new Event('eci-hearing-updated'));
  return Object.keys(out).length;
}

async function exportExcel(rows:Row[]){
  const XLSX=await import('xlsx-js-style');
  const headers=['S. No.','PS No.','Old PS','BLO Name','BLO Mobile No.','BLO Supervisor Name','Supervisor Mobile No.','Officer Name','Total No Mapping','No Mapping Notice Generated','No Mapping Notice Delivered','No Mapping Notice Pending','Hearing Status','Hearing Date(s)'];
  const sum=(k:keyof Row)=>rows.reduce((a,r)=>a+Number(r[k]||0),0);
  const out:any[][]=[['NO MAPPING NOTICE — PS WISE','ECI Hearing & Notice Monitoring — AC-34 MATIALA'],['Report generated',new Date().toLocaleString('en-IN')],['Scope','PS having at least 1 No Mapping Notice Generated or Delivered; hearing dates from uploaded PS-wise hearing date Excel'],['',''],headers,...rows.map((r,i)=>[i+1,r.ps,r.oldPs??'—',r.bloName||'—',r.bloPhone||'—',r.supervisor||'—',r.supervisorPhone||'—',r.officer||'—',r.total,r.generated,r.delivered,r.pending,r.hearingStatus||'—',r.hearingDate||'—']),['TOTAL','','','','','','','',sum('total'),sum('generated'),sum('delivered'),sum('pending'),'','']];
  const ws=XLSX.utils.aoa_to_sheet(out),max=headers.length,last=out.length-1,navy='FF17365D',blue='FF1F4E78',pale='FFEAF2F8',white='FFFFFFFF',text='FF17202A',red='FFC00000',redFill='FFFCE4E4',border='FFD0D9E3',totalFill='FFD9EAF7';
  const thin={style:'thin',color:{rgb:border}};
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:max-1}}];
  for(let c=0;c<max;c++){const a=XLSX.utils.encode_cell({r:0,c});if(!ws[a])ws[a]={v:''};ws[a].s={font:{name:'Aptos Display',bold:true,color:{rgb:white},sz:16},fill:{fgColor:{rgb:navy},patternType:'solid'},alignment:{horizontal:'center',vertical:'center'}};const b=XLSX.utils.encode_cell({r:1,c});if(!ws[b])ws[b]={v:''};ws[b].s={font:{name:'Aptos',italic:true,color:{rgb:'FF5B6573'},sz:10},fill:{fgColor:{rgb:pale},patternType:'solid'}};const d=XLSX.utils.encode_cell({r:2,c});if(!ws[d])ws[d]={v:''};ws[d].s={font:{name:'Aptos',bold:true,color:{rgb:'FF5B6573'},sz:9},fill:{fgColor:{rgb:pale},patternType:'solid'}};}
  for(let c=0;c<max;c++){const a=XLSX.utils.encode_cell({r:4,c});ws[a].s={font:{name:'Aptos',bold:true,color:{rgb:white},sz:10},fill:{fgColor:{rgb:blue},patternType:'solid'},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:thin,bottom:thin,left:thin,right:thin}};}
  for(let r=5;r<=last;r++)for(let c=0;c<max;c++){const a=XLSX.utils.encode_cell({r,c}),cell=ws[a];if(!cell)continue;const totalRow=r===last;cell.s=totalRow?{font:{name:'Aptos',bold:true,color:{rgb:navy},sz:10},fill:{fgColor:{rgb:totalFill},patternType:'solid'},border:{top:{style:'medium',color:{rgb:'FF2F75B5'}},bottom:{style:'medium',color:{rgb:'FF2F75B5'}},left:thin,right:thin}}:{font:{name:'Aptos',color:{rgb:text},sz:10},fill:{fgColor:{rgb:r%2?'FFF5F8FC':'FFFFFFFF'},patternType:'solid'},border:{top:thin,bottom:thin,left:thin,right:thin}};if(c===11&&!totalRow)cell.s={...cell.s,fill:{fgColor:{rgb:redFill},patternType:'solid'},font:{name:'Aptos',bold:true,color:{rgb:red},sz:10}};if(typeof cell.v==='number')cell.z='#,##0';if([0,1,2,4,6,8,9,10,11].includes(c))cell.s={...cell.s,alignment:{...(cell.s.alignment||{}),horizontal:'center'}};}
  ws['!cols']=[9,10,10,30,18,28,21,30,20,24,24,24,18,25].map(w=>({wch:w}));ws['!rows']=[{hpt:32},{hpt:20},{hpt:24},{hpt:9},{hpt:40},...Array(Math.max(0,last-4)).fill({hpt:22})];ws['!freeze']={xSplit:0,ySplit:5};ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:4,c:0},e:{r:last-1,c:max-1}})};ws['!pageSetup']={orientation:'landscape',fitToWidth:1,fitToHeight:0,paperSize:9};
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'No Mapping Notice');XLSX.writeFile(wb,`No_Mapping_Notice_PS_Wise_${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}.xlsx`);
}

export default function LowDeliveryReport(){
  useEffect(()=>{
    const setup=()=>{
      const main=document.querySelector('.app main.wrap') as HTMLElement|null,tabs=document.querySelector('.tabs') as HTMLElement|null;if(!main||!tabs)return;
      let custom=tabs.querySelector('[data-notice-not-generated-tab="1"]') as HTMLButtonElement|null;
      if(!custom){custom=document.createElement('button');custom.type='button';custom.className='tab';custom.textContent='No Mapping Notice';custom.title='PS Wise — No Mapping Notice Generated & Delivered';custom.dataset.noticeNotGeneratedTab='1';const report=Array.from(tabs.querySelectorAll('.tab')).find(x=>clean(x.textContent)==='Reports');if(report)report.insertAdjacentElement('afterend',custom);else tabs.appendChild(custom)}
      let panel=main.querySelector('[data-notice-not-generated-panel="1"]') as HTMLElement|null;
      if(!panel){panel=document.createElement('section');panel.className='section';panel.dataset.noticeNotGeneratedPanel='1';panel.style.display='none';main.appendChild(panel)}
      const show=()=>{Array.from(tabs.querySelectorAll('.tab')).forEach(x=>x.classList.remove('active'));custom!.classList.add('active');main.querySelectorAll(':scope > *').forEach(x=>{const el=x as HTMLElement;if(el!==tabs&&el!==panel)el.style.display='none'});panel!.style.display='block';build(panel!)};
      if(custom.dataset.bound!=='1'){custom.dataset.bound='1';custom.addEventListener('click',e=>{e.preventDefault();e.stopPropagation();show()})}
      const onOther=(e:Event)=>{const t=e.target as HTMLElement;if(t.closest('[data-notice-not-generated-tab="1"]'))return;const other=t.closest('.tab');if(other){panel!.style.display='none';main.querySelectorAll(':scope > *').forEach(x=>{const el=x as HTMLElement;if(el!==tabs&&el!==panel)el.style.display=''})}};
      if(tabs.dataset.noticeOtherBound!=='1'){tabs.dataset.noticeOtherBound='1';tabs.addEventListener('click',onOther,true)}
      const active=document.querySelector('.tab.active')?.textContent?.trim()||'';if(active==='No Mapping Notice')show();else panel.style.display='none';
    };
    const build=(panel:HTMLElement)=>{
      const rows=getRows();panel.innerHTML='';
      const head=document.createElement('div');head.className='sectionHead';head.innerHTML='<div><h2 style="margin-bottom:3px">NO MAPPING NOTICE — PS WISE</h2><div class="hint">NO MAPPING NOTICE GENERATED + NO MAPPING NOTICE DELIVERED · Hearing Date(s) from uploaded PS-wise hearing date Excel</div></div><strong style="font-size:18px;color:#24c9ff">'+rows.length.toLocaleString('en-IN')+' PS</strong>';panel.appendChild(head);
      const summary=document.createElement('div');summary.style.cssText='display:flex;gap:18px;flex-wrap:wrap;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.08);font-size:12px;font-weight:700';
      const inp=document.createElement('input');inp.type='file';inp.accept='.xlsx,.xls';inp.style.display='none';
      const load=document.createElement('button');load.className='btn';load.textContent='↑ LOAD PS-WISE HEARING DATES';load.style.cssText='margin-left:auto;padding:7px 12px';
      load.onclick=()=>inp.click();
      inp.onchange=async()=>{const f=inp.files?.[0];if(!f)return;load.disabled=true;load.textContent='LOADING…';try{const n=await loadHearingFile(f);load.textContent='✓ '+n+' PS LOADED';build(panel)}catch{load.textContent='LOAD FAILED — CHECK PS/HEARING DATE COLUMNS'}setTimeout(()=>{load.disabled=false;load.textContent='↑ LOAD PS-WISE HEARING DATES'},3000)};
      summary.innerHTML='<span>PS: <b style="color:#24c9ff">'+rows.length+'</b></span><span>NO MAPPING: <b>'+rows.reduce((a,r)=>a+r.total,0).toLocaleString('en-IN')+'</b></span><span>GENERATED: <b>'+rows.reduce((a,r)=>a+r.generated,0).toLocaleString('en-IN')+'</b></span><span>DELIVERED: <b>'+rows.reduce((a,r)=>a+r.delivered,0).toLocaleString('en-IN')+'</b></span><span>PENDING DELIVERY: <b style="color:#ff8a8a">'+rows.reduce((a,r)=>a+r.pending,0).toLocaleString('en-IN')+'</b></span>';
      summary.appendChild(inp);summary.appendChild(load);
      const btn=document.createElement('button');btn.className='btn';btn.textContent='↓ EXPORT EXCEL';btn.style.cssText='padding:7px 12px';btn.onclick=async()=>{btn.disabled=true;btn.textContent='EXPORTING…';try{await exportExcel(getRows());btn.textContent='✓ EXPORTED'}catch{btn.textContent='EXPORT FAILED'}setTimeout(()=>{btn.disabled=false;btn.textContent='↓ EXPORT EXCEL'},2000)};summary.appendChild(btn);panel.appendChild(summary);
      const wrap=document.createElement('div');wrap.className='tablewrap';const table=document.createElement('table');table.innerHTML='<thead><tr><th>S. No.</th><th>PS No.</th><th>Old PS</th><th>BLO Name</th><th>BLO Mobile No.</th><th>BLO Supervisor Name</th><th>Supervisor Mobile No.</th><th>Officer Name</th><th>Total No Mapping</th><th>No Mapping Notice Generated</th><th>No Mapping Notice Delivered</th><th>No Mapping Notice Pending</th><th>Hearing Status</th><th>Hearing Date(s)</th></tr></thead><tbody></tbody>';
      const body=table.tBodies[0];rows.forEach((r,i)=>{const tr=document.createElement('tr');[i+1,r.ps,r.oldPs??'—',r.bloName||'—',r.bloPhone||'—',r.supervisor||'—',r.supervisorPhone||'—',r.officer||'—',r.total.toLocaleString('en-IN'),r.generated.toLocaleString('en-IN'),r.delivered.toLocaleString('en-IN'),r.pending.toLocaleString('en-IN'),r.hearingStatus||'—',r.hearingDate||'—'].forEach((v,j)=>{const td=document.createElement('td');td.textContent=String(v);if([0,1,2,4,6,8,9,10,11].includes(j))td.style.textAlign='center';if(j===9)td.style.color='#24c9ff';if(j===10)td.style.color='#27db93';if(j===11){td.style.color='#ff8a8a';td.style.fontWeight='800'}if(j===13)td.style.fontWeight='700';tr.appendChild(td)});body.appendChild(tr)});
      if(!rows.length){const tr=document.createElement('tr');const td=document.createElement('td');td.colSpan=14;td.textContent='No PS has any No Mapping Notice Generated/Delivered yet.';td.style.textAlign='center';tr.appendChild(td);body.appendChild(tr)}
      wrap.appendChild(table);panel.appendChild(wrap);
    };
    const observer=new MutationObserver(()=>requestAnimationFrame(setup));observer.observe(document.body,{subtree:true,childList:true});
    const onHearingUpdated=()=>{const p=document.querySelector('[data-notice-not-generated-panel="1"]') as HTMLElement|null;if(p&&p.style.display!=='none')build(p)};
    window.addEventListener('eci-hearing-updated',onHearingUpdated);setup();
    return()=>{observer.disconnect();window.removeEventListener('eci-hearing-updated',onHearingUpdated)};
  },[]);
  return null;
}
