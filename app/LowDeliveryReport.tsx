'use client';
import { useEffect, useState } from 'react';

const supervisorPhone:Record<string,string>={
  'SH. PARVEEN KUMAR':'9953601073','SMT. SHASHI BALA':'9953312984','SH. RAKESH KUMAR':'7011971522','SMT. PARUL GUPTA':'9667881989','SH. SUBHASHISH':'9868252144','SH. VIRENDER':'9868252144',
};
const clean=(v:any)=>String(v??'').replace(/\s+/g,' ').trim();
const num=(v:any)=>{const n=Number(v);return Number.isFinite(n)?Math.max(n,0):0};
const person=(v:any)=>{const x=clean(v);const m=x.match(/^(.*?)\s*\(\s*(\d{10})\s*\)\s*$/);return m?{name:m[1].trim(),phone:m[2]}:{name:x,phone:''};};

type Row={ps:number;oldPs:any;bloName:string;bloPhone:string;supervisor:string;supervisorPhone:string;officer:string;scheduled:number;generated:number;delivered:number;pending:number};
function getRows():Row[]{
  let master:any[]=[],eci:any[]=[];try{master=JSON.parse(localStorage.getItem('eci-master')||'[]');eci=JSON.parse(localStorage.getItem('eci-current')||'[]');}catch{return []}
  if(!master.length||!eci.length)return [];
  const em=new Map<number,any>(eci.map(r=>[Number(r?.ps),r]));
  return master.map(m=>{const ps=num(m?.ps),e=em.get(ps);if(!e)return null;const generated=num(e?.['Notice Generated']);if(generated!==0)return null;const pendingGeneration=num(e?.['Pending for Notice Generation']);const delivered=num(e?.['Notice Delivered']);const scheduled=generated+pendingGeneration;const blo=person(m?.blo),sup=person(m?.supervisor);return {ps,oldPs:m?.oldPs,bloName:blo.name,bloPhone:blo.phone,supervisor:sup.name,supervisorPhone:sup.phone||supervisorPhone[sup.name.toUpperCase()]||'',officer:person(m?.officer).name,scheduled,generated,delivered,pending:Math.max(scheduled-delivered,0)};}).filter(Boolean).sort((a:any,b:any)=>a.ps-b.ps) as Row[];
}

async function exportDirect(rows:Row[]){
  const XLSX=await import('xlsx-js-style');
  const wb=XLSX.utils.book_new();
  const title='NOTICE NOT GENERATED — PS & OFFICER WISE';
  const data:any[][]=[
    ['S. No.','PS No.','Old PS','BLO Name','BLO Mobile No.','BLO Supervisor Name','Supervisor Mobile No.','Officer Name','Notices Scheduled','Notices Generated','Notices Delivered','Notices Pending'],
    ...rows.map((r,i)=>[i+1,r.ps,r.oldPs??'—',r.bloName||'—',r.bloPhone||'—',r.supervisor||'—',r.supervisorPhone||'—',r.officer||'—',r.scheduled,r.generated,r.delivered,r.pending]),
  ];
  const totalScheduled=rows.reduce((a,r)=>a+r.scheduled,0),totalDelivered=rows.reduce((a,r)=>a+r.delivered,0),totalPending=rows.reduce((a,r)=>a+r.pending,0);
  data.push(['TOTAL','','','','','','','',totalScheduled,0,totalDelivered,totalPending]);
  const out:any[][]=[[title,'ECI Hearing & Notice Monitoring — AC-34 MATIALA'],['Report generated',new Date().toLocaleString('en-IN')],['Filter','Only PS where ECI Notice Generated = 0'],['',''],...data];
  const ws=XLSX.utils.aoa_to_sheet(out);const maxCols=12;const last=out.length-1;
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:maxCols-1}}];
  const navy='FF17365D',blue='FF1F4E78',pale='FFEAF2F8',white='FFFFFFFF',text='FF17202A',stripe='FFF5F8FC',red='FFC00000',redFill='FFFCE4E4',border='FFD0D9E3';
  const thin={style:'thin',color:{rgb:border}};const header={font:{name:'Aptos',bold:true,color:{rgb:white},sz:11},fill:{fgColor:{rgb:blue},patternType:'solid'},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:thin,bottom:thin,left:thin,right:thin}};
  const body={font:{name:'Aptos',color:{rgb:text},sz:10},alignment:{vertical:'center'},border:{top:thin,bottom:thin,left:thin,right:thin}};
  const alt={...body,fill:{fgColor:{rgb:stripe},patternType:'solid'}};
  for(let c=0;c<maxCols;c++){const a=XLSX.utils.encode_cell({r:0,c});if(!ws[a])ws[a]={v:''};ws[a].s={font:{name:'Aptos Display',bold:true,color:{rgb:white},sz:16},fill:{fgColor:{rgb:navy},patternType:'solid'},alignment:{horizontal:'center',vertical:'center'}};const b=XLSX.utils.encode_cell({r:1,c});if(!ws[b])ws[b]={v:''};ws[b].s={font:{name:'Aptos',italic:true,color:{rgb:'FF5B6573'},sz:10},fill:{fgColor:{rgb:pale},patternType:'solid'}};const d=XLSX.utils.encode_cell({r:2,c});if(!ws[d])ws[d]={v:''};ws[d].s={font:{name:'Aptos',bold:true,color:{rgb:'FF5B6573'},sz:9},fill:{fgColor:{rgb:pale},patternType:'solid'}};}
  for(let r=4;r<=last;r++){for(let c=0;c<maxCols;c++){const a=XLSX.utils.encode_cell({r,c});const cell=ws[a];if(!cell)continue;if(r===4)cell.s=header;else if(r===last)cell.s={font:{name:'Aptos',bold:true,color:{rgb:navy},sz:10},fill:{fgColor:{rgb:'FFD9EAF7'},patternType:'solid'},border:{top:{style:'medium',color:{rgb:'FF2F75B5'}},bottom:{style:'medium',color:{rgb:'FF2F75B5'}},left:thin,right:thin}};else cell.s=(r-4)%2===0?alt:body;if([9,11].includes(c)){cell.s={...cell.s,fill:{fgColor:{rgb:redFill},patternType:'solid'},font:{name:'Aptos',bold:true,color:{rgb:red},sz:10}};}if(typeof cell.v==='number')cell.z='#,##0';if([0,1,2,4,6,8,9,10,11].includes(c))cell.s={...cell.s,alignment:{...cell.s.alignment,horizontal:'center'}};}}
  ws['!cols']=[9,11,10,32,18,28,20,30,19,18,18,18].map(w=>({wch:w}));ws['!rows']=[{hpt:32},{hpt:20},{hpt:19},{hpt:9},{hpt:36},...Array(Math.max(0,last-4)).fill({hpt:22})];ws['!freeze']={xSplit:0,ySplit:5};ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:4,c:0},e:{r:last-1,c:maxCols-1}})};ws['!pageSetup']={orientation:'landscape',fitToWidth:1,fitToHeight:0,paperSize:9};
  XLSX.utils.book_append_sheet(wb,ws,'Notice Not Generated');
  const stamp=new Date().toISOString().slice(0,19).replace(/[:T]/g,'-');XLSX.writeFile(wb,`Notice_Not_Generated_PS_Wise_${stamp}.xlsx`);
}

export default function LowDeliveryReport(){
  const[busy,setBusy]=useState(false);
  useEffect(()=>{
    const render=()=>{const active=document.querySelector('.tab.active')?.textContent?.trim()||'';const old=document.querySelector('[data-low-delivery-report="1"]');if(active!=='Reports'){old?.remove();return;}if(old)return;const rows=getRows();const host=document.querySelector('.app main.wrap');if(!host)return;const section=document.createElement('section');section.setAttribute('data-low-delivery-report','1');section.className='section';section.style.marginTop='14px';const head=document.createElement('div');head.className='sectionHead';head.innerHTML='<div><h2 style="margin-bottom:3px">NOTICE NOT GENERATED — PS & OFFICER WISE</h2><div class="hint">Only PS with ECI Notice Generated = 0 are listed · Notices Scheduled = Generated + Pending for Notice Generation</div></div><strong style="font-size:18px;color:#ff8a8a">'+rows.length.toLocaleString('en-IN')+' PS</strong>';section.appendChild(head);const zero=rows.filter(r=>r.delivered===0).length,delivered=rows.reduce((a,r)=>a+r.delivered,0),scheduled=rows.reduce((a,r)=>a+r.scheduled,0);const summary=document.createElement('div');summary.style.cssText='display:flex;gap:18px;flex-wrap:wrap;margin:0 0 12px;padding:9px 12px;border:1px solid rgba(255,255,255,.08);border-radius:8px;background:rgba(255,255,255,.025);font-size:12px;font-weight:700';summary.innerHTML='<span>PS NOT GENERATED: <b style="color:#ff8a8a">'+rows.length+'</b></span><span>NOTICES SCHEDULED: <b>'+scheduled.toLocaleString('en-IN')+'</b></span><span>DELIVERED: <b>'+delivered.toLocaleString('en-IN')+'</b></span><span>ZERO DELIVERED: <b style="color:#ff8a8a">'+zero+'</b></span>';section.appendChild(summary);const button=document.createElement('button');button.textContent='↓ EXPORT NOTICE NOT GENERATED REPORT';button.style.cssText='margin:0 0 12px;padding:8px 12px;border-radius:7px;border:1px solid rgba(255,255,255,.18);background:#18243a;color:#fff;font-size:11px;font-weight:800;cursor:pointer';button.onclick=async()=>{if(busy)return;setBusy(true);button.textContent='EXPORTING…';try{await exportDirect(getRows());button.textContent='✓ EXPORTED';setTimeout(()=>button.textContent='↓ EXPORT NOTICE NOT GENERATED REPORT',2500)}catch(e){button.textContent='EXPORT FAILED';setTimeout(()=>button.textContent='↓ EXPORT NOTICE NOT GENERATED REPORT',3000)}finally{setBusy(false)}};section.appendChild(button);const wrap=document.createElement('div');wrap.className='tablewrap';const table=document.createElement('table');table.innerHTML='<thead><tr><th>S. No.</th><th>PS No.</th><th>Old PS</th><th>BLO Name</th><th>BLO Mobile No.</th><th>BLO Supervisor Name</th><th>Supervisor Mobile No.</th><th>Officer Name</th><th>Notices Scheduled</th><th>Notices Generated</th><th>Notices Delivered</th><th>Notices Pending</th></tr></thead><tbody></tbody>';const body=table.tBodies[0];rows.forEach((r,i)=>{const tr=document.createElement('tr');[i+1,r.ps,r.oldPs??'—',r.bloName||'—',r.bloPhone||'—',r.supervisor||'—',r.supervisorPhone||'—',r.officer||'—',r.scheduled.toLocaleString('en-IN'),'0',r.delivered.toLocaleString('en-IN'),r.pending.toLocaleString('en-IN')].forEach((v,j)=>{const td=document.createElement('td');td.textContent=String(v);if([0,1,2,6,8,9,10,11].includes(j))td.style.textAlign='center';if(j===9||j===11){td.style.fontWeight='800';td.style.color='#ff8a8a'}tr.appendChild(td)});body.appendChild(tr)});if(!rows.length){const tr=document.createElement('tr');const td=document.createElement('td');td.colSpan=12;td.textContent='No PS currently has Notice Generated = 0.';td.style.textAlign='center';tr.appendChild(td);body.appendChild(tr)}wrap.appendChild(table);section.appendChild(wrap);host.appendChild(section)};const observer=new MutationObserver(()=>requestAnimationFrame(render));observer.observe(document.body,{subtree:true,childList:true});render();return()=>observer.disconnect()},[busy]);
  return null;
}
