'use client';
import { useEffect } from 'react';

const supervisorPhone:Record<string,string>={
  'SH. PARVEEN KUMAR':'9953601073','SMT. SHASHI BALA':'9953312984','SH. RAKESH KUMAR':'7011971522','SMT. PARUL GUPTA':'9667881989','SH. SUBHASHISH':'9868252144','SH. VIRENDER':'9868252144',
};
const clean=(v:any)=>String(v??'').replace(/\s+/g,' ').trim();
const num=(v:any)=>{const n=Number(v);return Number.isFinite(n)?Math.max(n,0):0};
const person=(v:any)=>{const x=clean(v);const m=x.match(/^(.*?)\s*\(\s*(\d{10})\s*\)\s*$/);return m?{name:m[1].trim(),phone:m[2]}:{name:x,phone:''};};

export default function LowDeliveryReport(){
  useEffect(()=>{
    const render=()=>{
      const active=document.querySelector('.tab.active')?.textContent?.trim()||'';
      const old=document.querySelector('[data-low-delivery-report="1"]');
      if(active!=='Reports'){old?.remove();return;} if(old)return;
      let master:any[]=[],eci:any[]=[];try{master=JSON.parse(localStorage.getItem('eci-master')||'[]');eci=JSON.parse(localStorage.getItem('eci-current')||'[]');}catch{return;}
      if(!master.length||!eci.length)return;
      const em=new Map<number,any>(eci.map(r=>[Number(r?.ps),r]));
      const rows=master.map(m=>{const ps=num(m?.ps),e=em.get(ps);if(!e)return null;
        const generated=num(e?.['Notice Generated']);
        // This report is ONLY for PS where notice generation has not started: ECI Notice Generated = 0.
        if(generated!==0)return null;
        const pendingGeneration=num(e?.['Pending for Notice Generation']);
        const delivered=num(e?.['Notice Delivered']);
        const scheduled=generated+pendingGeneration;
        const blo=person(m?.blo),sup=person(m?.supervisor);
        return {ps,oldPs:m?.oldPs,bloName:blo.name,bloPhone:blo.phone,supervisor:sup.name,supervisorPhone:sup.phone||supervisorPhone[sup.name.toUpperCase()]||'',officer:person(m?.officer).name,scheduled,delivered,pending:Math.max(scheduled-delivered,0)};
      }).filter(Boolean).sort((a:any,b:any)=>a.ps-b.ps) as any[];
      const host=document.querySelector('.app main.wrap');if(!host)return;
      const section=document.createElement('section');section.setAttribute('data-low-delivery-report','1');section.className='section';section.style.marginTop='14px';
      const head=document.createElement('div');head.className='sectionHead';head.innerHTML='<div><h2 style="margin-bottom:3px">NOTICE NOT GENERATED — PS & OFFICER WISE</h2><div class="hint">Only PS with ECI Notice Generated = 0 are listed · Notices Scheduled = Generated + Pending for Notice Generation</div></div><strong style="font-size:18px;color:#ff8a8a">'+rows.length.toLocaleString('en-IN')+' PS</strong>';section.appendChild(head);
      const zero=rows.filter(r=>r.delivered===0).length,delivered=rows.reduce((a,r)=>a+r.delivered,0),scheduled=rows.reduce((a,r)=>a+r.scheduled,0);
      const summary=document.createElement('div');summary.style.cssText='display:flex;gap:18px;flex-wrap:wrap;margin:0 0 12px;padding:9px 12px;border:1px solid rgba(255,255,255,.08);border-radius:8px;background:rgba(255,255,255,.025);font-size:12px;font-weight:700';summary.innerHTML='<span>PS NOT GENERATED: <b style="color:#ff8a8a">'+rows.length+'</b></span><span>NOTICES SCHEDULED: <b>'+scheduled.toLocaleString('en-IN')+'</b></span><span>DELIVERED: <b>'+delivered.toLocaleString('en-IN')+'</b></span><span>ZERO DELIVERED: <b style="color:#ff8a8a">'+zero+'</b></span>';section.appendChild(summary);
      const wrap=document.createElement('div');wrap.className='tablewrap';
      const table=document.createElement('table');table.innerHTML='<thead><tr><th>S. No.</th><th>PS No.</th><th>Old PS</th><th>BLO Name</th><th>BLO Mobile No.</th><th>BLO Supervisor Name</th><th>Supervisor Mobile No.</th><th>Officer Name</th><th>Notices Scheduled</th><th>Notices Generated</th><th>Notices Delivered</th><th>Notices Pending</th></tr></thead><tbody></tbody>';
      const body=table.tBodies[0];rows.forEach((r,i)=>{const tr=document.createElement('tr');[i+1,r.ps,r.oldPs??'—',r.bloName||'—',r.bloPhone||'—',r.supervisor||'—',r.supervisorPhone||'—',r.officer||'—',r.scheduled.toLocaleString('en-IN'),'0',r.delivered.toLocaleString('en-IN'),r.pending.toLocaleString('en-IN')].forEach((v,j)=>{const td=document.createElement('td');td.textContent=String(v);if([0,1,2,6,8,9,10,11].includes(j))td.style.textAlign='center';if(j===9||j===11){td.style.fontWeight='800';td.style.color='#ff8a8a';}tr.appendChild(td)});body.appendChild(tr)});
      if(!rows.length){const tr=document.createElement('tr');const td=document.createElement('td');td.colSpan=12;td.textContent='No PS currently has Notice Generated = 0.';td.style.textAlign='center';tr.appendChild(td);body.appendChild(tr)}
      wrap.appendChild(table);section.appendChild(wrap);host.appendChild(section);
    };
    const observer=new MutationObserver(()=>requestAnimationFrame(render));observer.observe(document.body,{subtree:true,childList:true});render();return()=>observer.disconnect();
  },[]); return null;
}
