'use client';

import { useState } from 'react';

const TAB_NAMES=['Overview','PS Wise','Hearing Centres','Officer Wise','Generated Notices','Pending','Hearings','DEO','Validation','Reports','History','Settings'];

const clean=(v:string)=>v.replace(/\s+/g,' ').trim();

function tableToRows(table:HTMLTableElement){
  const rows:string[][]=[];
  const head=table.querySelector('thead');
  if(head){
    const r=head.querySelector('tr:last-child');
    if(r)rows.push(Array.from(r.cells).map(c=>clean(c.textContent||'')));
  }
  Array.from(table.tBodies).forEach(tb=>Array.from(tb.rows).forEach(r=>{
    if(r.style.display==='none')return;
    rows.push(Array.from(r.cells).map(c=>clean(c.textContent||'')));
  }));
  return rows;
}

async function waitForRender(){await new Promise<void>(resolve=>setTimeout(resolve,180));}

export default function ReportExport(){
  const[busy,setBusy]=useState(false);
  const[message,setMessage]=useState('');

  const exportAll=async()=>{
    if(busy)return;
    setBusy(true);setMessage('Preparing Excel…');
    try{
      const XLSX=await import('xlsx');
      const wb=XLSX.utils.book_new();
      const original=(document.querySelector('.tab.active')?.textContent||'Overview').trim();
      const used=new Set<string>();
      let exported=0;

      for(const tabName of TAB_NAMES){
        if(tabName==='Settings')continue;
        const tabs=Array.from(document.querySelectorAll('.tab')) as HTMLElement[];
        const tab=tabs.find(x=>clean(x.textContent||'')===tabName);
        if(!tab)continue;
        tab.click();
        await waitForRender();
        const tables=Array.from(document.querySelectorAll('table')) as HTMLTableElement[];
        const rows:string[][]=[];
        tables.forEach(t=>{
          const part=tableToRows(t);
          if(part.length){if(rows.length)rows.push([]);rows.push(...part);}
        });
        if(!rows.length)continue;
        let name=tabName.replace(/[\\/?*\[\]:]/g,'').slice(0,31)||'Report';
        let base=name,i=2;while(used.has(name)){name=(base.slice(0,28)+'_'+i++).slice(0,31)}used.add(name);
        const ws=XLSX.utils.aoa_to_sheet(rows);
        const maxCols=Math.max(...rows.map(r=>r.length),1);
        const widths=Array.from({length:maxCols},(_,c)=>Math.min(42,Math.max(10,...rows.map(r=>String(r[c]??'').length+2))));
        ws['!cols']=widths.map(w=>({wch:w}));
        ws['!freeze']={xSplit:0,ySplit:1};
        if(rows[0]?.length)ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:rows.length-1,c:maxCols-1}})};
        XLSX.utils.book_append_sheet(wb,ws,name);
        exported++;
      }

      const restore=Array.from(document.querySelectorAll('.tab')).find(x=>clean(x.textContent||'')===original) as HTMLElement|undefined;
      restore?.click();
      await waitForRender();
      if(!exported)throw Error('No report tables were available to export.');
      const stamp=new Date().toISOString().slice(0,19).replace(/[:T]/g,'-');
      XLSX.writeFile(wb,`ECI_Hearing_Reports_${stamp}.xlsx`);
      setMessage(`${exported} reports exported`);
    }catch(e){setMessage(e instanceof Error?e.message:'Excel export failed');}
    finally{setBusy(false);setTimeout(()=>setMessage(''),3500);}
  };

  return <>
    <button onClick={exportAll} disabled={busy} title="Download all dashboard reports as a formatted Excel workbook" style={{position:'fixed',right:22,top:18,zIndex:99999,border:'1px solid rgba(255,255,255,.18)',borderRadius:8,padding:'9px 14px',fontSize:12,fontWeight:800,cursor:busy?'wait':'pointer',background:'linear-gradient(180deg,#18243a,#101827)',color:'#fff',boxShadow:'0 6px 18px rgba(0,0,0,.28)'}}>{busy?'EXPORTING…':'↓ EXPORT ALL REPORTS'}</button>
    {message&&<div style={{position:'fixed',right:22,top:62,zIndex:99999,padding:'7px 11px',borderRadius:7,fontSize:11,fontWeight:700,background:'#101827',color:'#dbe7ff',border:'1px solid rgba(255,255,255,.12)'}}>{message}</div>}
  </>;
}
