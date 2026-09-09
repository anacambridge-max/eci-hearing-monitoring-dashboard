'use client';

import { useState } from 'react';

const TAB_NAMES=['Overview','PS Wise','Hearing Centres','Officer Wise','Generated Notices','Pending','Hearings','DEO','Validation','Reports','History','Settings'];
type ExportCell=string|number|boolean|null;
const clean=(v:string)=>v.replace(/\s+/g,' ').trim();

function tableToRows(table:HTMLTableElement):string[][]{
  const rows:string[][]=[];
  const head=table.querySelector('thead');
  if(head){const r=head.querySelector('tr:last-child') as HTMLTableRowElement|null;if(r)rows.push(Array.from(r.cells).map(c=>clean(c.textContent||'')));}
  Array.from(table.tBodies).forEach(tb=>Array.from(tb.rows).forEach(r=>{if(r.style.display==='none')return;rows.push(Array.from(r.cells).map(c=>clean(c.textContent||'')));}));
  return rows;
}
function typedValue(value:string):ExportCell{
  const v=clean(value);if(!v)return null;
  if(/^[-+]?\d[\d,]*(\.\d+)?$/.test(v)){const n=Number(v.replace(/,/g,''));if(Number.isFinite(n))return n;}
  if(/^[-+]?\d+(\.\d+)?%$/.test(v)){const n=Number(v.slice(0,-1));if(Number.isFinite(n))return n/100;}
  return v;
}
const line={style:'thin',color:{rgb:'CFD8E3'}};
const header={font:{name:'Aptos',bold:true,color:'FFFFFF',sz:11},fill:{fgColor:{rgb:'1F4E78'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:line,bottom:line,left:line,right:line}};
const title={font:{name:'Aptos Display',bold:true,color:'FFFFFF',sz:16},fill:{fgColor:{rgb:'17365D'}},alignment:{horizontal:'left',vertical:'center'}};
const subtitle={font:{name:'Aptos',italic:true,color:'5B6573',sz:10},fill:{fgColor:{rgb:'EAF0F6'}},alignment:{vertical:'center'}};
const body={font:{name:'Aptos',color:'17202A',sz:10},alignment:{vertical:'center'},border:{top:line,bottom:line,left:line,right:line}};
const bodyAlt={...body,fill:{fgColor:{rgb:'F5F8FC'}}};
const total={font:{name:'Aptos',bold:true,color:'17365D',sz:10},fill:{fgColor:{rgb:'D9EAF7'}},alignment:{vertical:'center'},border:{top:{style:'medium',color:{rgb:'1F4E78'}},bottom:line,left:line,right:line}};

export default function ReportExport(){
  const[busy,setBusy]=useState(false);const[message,setMessage]=useState('');
  const exportAll=async()=>{
    if(busy)return;setBusy(true);setMessage('Preparing professional Excel workbook…');
    try{
      const XLSX=await import('xlsx-js-style');const wb=XLSX.utils.book_new();
      const original=(document.querySelector('.tab.active')?.textContent||'Overview').trim();const used=new Set<string>();let exported=0;const generated=new Date();
      for(const tabName of TAB_NAMES){
        if(tabName==='Settings')continue;
        const tabs=Array.from(document.querySelectorAll('.tab')) as HTMLElement[];const tab=tabs.find(x=>clean(x.textContent||'')===tabName);if(!tab)continue;
        tab.click();await new Promise<void>(r=>setTimeout(r,220));
        const tables=Array.from(document.querySelectorAll('table')) as HTMLTableElement[];const raw:string[][]=[];
        tables.forEach(t=>{const part=tableToRows(t);if(part.length){if(raw.length)raw.push([]);raw.push(...part);}});if(!raw.length)continue;
        let name=tabName.replace(/[\\/?*\[\]:]/g,'').slice(0,31)||'Report';const base=name;let i=2;while(used.has(name)){name=(base.slice(0,27)+'_'+i++).slice(0,31);}used.add(name);
        const data=raw.map(r=>r.map(typedValue));
        const rows:ExportCell[][]=[[tabName,'ECI Hearing & Notice Monitoring'],['Report generated',generated.toLocaleString('en-IN')],['Data source','Current dashboard'],[],...data];
        const ws=XLSX.utils.aoa_to_sheet(rows);const maxCols=Math.max(...raw.map(r=>r.length),2);const last=rows.length-1;
        ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:maxCols-1}}];
        for(let c=0;c<maxCols;c++){const a=XLSX.utils.encode_cell({r:0,c});if(!ws[a])ws[a]={v:''};ws[a].s=title;}
        for(let c=0;c<maxCols;c++){const a=XLSX.utils.encode_cell({r:1,c});if(!ws[a])ws[a]={v:''};ws[a].s=subtitle;}
        for(let c=0;c<maxCols;c++){const a=XLSX.utils.encode_cell({r:2,c});if(!ws[a])ws[a]={v:''};ws[a].s={...subtitle,font:{name:'Aptos',color:{rgb:'6B7280'},sz:9}};}
        let sectionHeader=-1;
        for(let r=4;r<rows.length;r++){
          const source=raw[r-4]||[];
          if(!source.length){sectionHeader=-1;continue;}
          if(sectionHeader<0)sectionHeader=r;
          const first=clean(String(source[0]??''));const isTotal=/^(total|grand total|overall|sub[- ]?total)/i.test(first)||(/total/i.test(first)&&source.length<=3);
          for(let c=0;c<maxCols;c++){
            const a=XLSX.utils.encode_cell({r,c});const cell=ws[a];if(!cell)continue;
            if(r===sectionHeader)cell.s=header;else if(isTotal)cell.s=total;else cell.s=(r-sectionHeader)%2===0?bodyAlt:body;
            if(typeof cell.v==='number')cell.z=/%$/.test(String(source[c]??''))?'0.0%':'#,##0';
            if(c===0)cell.s={...cell.s,alignment:{...cell.s.alignment,horizontal:'center'}};
          }
        }
        const widths=Array.from({length:maxCols},(_,c)=>{const values=raw.map(r=>String(r[c]??''));const max=Math.max(9,...values.map(v=>Math.min(v.length,70)+2));const h=String(raw.find(r=>r.length)?.[c]??'');const cap=/Officer|Location|Centre|Address|Name|Remark|Status|Description/i.test(h)?52:/Date|Time/i.test(h)?22:30;return Math.min(cap,max);});
        ws['!cols']=widths.map(w=>({wch:w}));
        ws['!rows']=[{hpt:30},{hpt:20},{hpt:18},{hpt:8},...Array(Math.max(0,last-3)).fill({hpt:20})];
        ws['!freeze']={xSplit:0,ySplit:4};ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:4,c:0},e:{r:4+raw.length-1,c:maxCols-1}})};
        ws['!pageSetup']={orientation:'landscape',fitToWidth:1,fitToHeight:0,paperSize:9};ws['!margins']={left:0.25,right:0.25,top:0.55,bottom:0.55,header:0.2,footer:0.2};
        XLSX.utils.book_append_sheet(wb,ws,name);exported++;
      }
      const restore=Array.from(document.querySelectorAll('.tab')).find(x=>clean(x.textContent||'')===original) as HTMLElement|undefined;restore?.click();
      if(!exported)throw Error('No report tables were available to export.');
      const stamp=new Date().toISOString().slice(0,19).replace(/[:T]/g,'-');XLSX.writeFile(wb,`ECI_Hearing_Reports_${stamp}.xlsx`);setMessage(`${exported} professional reports exported successfully`);
    }catch(e){setMessage(e instanceof Error?e.message:'Excel export failed');}finally{setBusy(false);setTimeout(()=>setMessage(''),4500);}
  };
  return <><button onClick={exportAll} disabled={busy} title="Download all dashboard reports as a professionally formatted Excel workbook" style={{position:'fixed',right:22,top:18,zIndex:99999,border:'1px solid rgba(255,255,255,.18)',borderRadius:8,padding:'9px 14px',fontSize:12,fontWeight:800,cursor:busy?'wait':'pointer',background:'linear-gradient(180deg,#18243a,#101827)',color:'#fff',boxShadow:'0 6px 18px rgba(0,0,0,.28)'}}>{busy?'EXPORTING…':'↓ EXPORT ALL REPORTS'}</button>{message&&<div style={{position:'fixed',right:22,top:62,zIndex:99999,padding:'7px 11px',borderRadius:7,fontSize:11,fontWeight:700,background:'#101827',color:'#dbe7ff',border:'1px solid rgba(255,255,255,.12)'}}>{message}</div>}</>;
}
