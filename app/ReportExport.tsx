'use client';

import { useState } from 'react';

const TAB_NAMES=['Overview','PS Wise','Hearing Centres','Officer Wise','Generated Notices','Pending','Hearings','DEO','Validation','Reports','History','Settings'];
const clean=(v:string)=>v.replace(/\s+/g,' ').trim();
type ExportCell=string|number|boolean|null;

function tableToRows(table:HTMLTableElement):string[][]{
  const rows:string[][]=[];
  const head=table.querySelector('thead');
  if(head){const r=head.querySelector('tr:last-child') as HTMLTableRowElement|null;if(r)rows.push(Array.from(r.cells).map((c:HTMLTableCellElement)=>clean(c.textContent||'')));}
  Array.from(table.tBodies).forEach(tb=>Array.from(tb.rows).forEach(r=>{if(r.style.display==='none')return;rows.push(Array.from(r.cells).map((c:HTMLTableCellElement)=>clean(c.textContent||'')));}));
  return rows;
}
function typedValue(value:string):ExportCell{
  const v=clean(value);if(v==='')return null;
  if(/^[-+]?\d[\d,]*(\.\d+)?$/.test(v)){const n=Number(v.replace(/,/g,''));if(Number.isFinite(n))return n;}
  if(/^[-+]?\d+(\.\d+)?%$/.test(v)){const n=Number(v.slice(0,-1));if(Number.isFinite(n))return n/100;}
  return v;
}
async function waitForRender(){await new Promise<void>(resolve=>setTimeout(resolve,180));}

const headerStyle={font:{bold:true,color:'FFFFFF',sz:11},fill:{fgColor:{rgb:'17365D'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:{style:'thin',color:{rgb:'B8C7D9'}},bottom:{style:'thin',color:{rgb:'B8C7D9'}},left:{style:'thin',color:{rgb:'B8C7D9'}},right:{style:'thin',color:{rgb:'B8C7D9'}}}};
const bodyBorder={border:{top:{style:'thin',color:{rgb:'D9E2F3'}},bottom:{style:'thin',color:{rgb:'D9E2F3'}},left:{style:'thin',color:{rgb:'D9E2F3'}},right:{style:'thin',color:{rgb:'D9E2F3'}}},alignment:{vertical:'center',wrapText:false}};

export default function ReportExport(){
  const[busy,setBusy]=useState(false);const[message,setMessage]=useState('');
  const exportAll=async()=>{
    if(busy)return;setBusy(true);setMessage('Preparing formatted Excel…');
    try{
      const XLSX=await import('xlsx-js-style');const wb=XLSX.utils.book_new();
      const original=(document.querySelector('.tab.active')?.textContent||'Overview').trim();const used=new Set<string>();let exported=0;const generated=new Date();
      for(const tabName of TAB_NAMES){
        if(tabName==='Settings')continue;
        const tabs=Array.from(document.querySelectorAll('.tab')) as HTMLElement[];const tab=tabs.find(x=>clean(x.textContent||'')===tabName);if(!tab)continue;
        tab.click();await waitForRender();
        const tables=Array.from(document.querySelectorAll('table')) as HTMLTableElement[];const raw:string[][]=[];
        tables.forEach(t=>{const part=tableToRows(t);if(part.length){if(raw.length)raw.push([]);raw.push(...part);}});if(!raw.length)continue;
        let name=tabName.replace(/[\\/?*\[\]:]/g,'').slice(0,31)||'Report';const base=name;let i=2;while(used.has(name)){name=(base.slice(0,28)+'_'+i++).slice(0,31)}used.add(name);
        const data=raw.map(r=>r.map(typedValue));
        const rows:ExportCell[][]=[[tabName,'ECI Hearing & Notice Monitoring'],['Generated On',generated.toLocaleString('en-IN')],[],...data];
        const ws=XLSX.utils.aoa_to_sheet(rows);const maxCols=Math.max(...raw.map(r=>r.length),1);const lastRow=rows.length;const lastCol=maxCols-1;
        ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:Math.max(1,lastCol)}}];
        ws['A1'].s={font:{bold:true,color:'FFFFFF',sz:15},fill:{fgColor:{rgb:'0F243E'}},alignment:{vertical:'center'}};
        ws['A2'].s={font:{italic:true,color:'5B6573',sz:10},alignment:{vertical:'center'}};ws['!rows']=[{hpt:28},{hpt:20},{hpt:8}];
        let sectionStart=3;
        for(let r=3;r<rows.length;r++){
          const sourceRow=raw[r-3]||[];if(sourceRow.length===0){sectionStart=r+1;continue;}
          const isHeader=r===sectionStart;
          for(let c=0;c<maxCols;c++){
            const cell=ws[XLSX.utils.encode_cell({r,c})];if(!cell)continue;
            if(isHeader)cell.s=headerStyle;
            else{cell.s={...bodyBorder,fill:{fgColor:{rgb:(r-sectionStart)%2===0?'F7FAFD':'FFFFFF'}}};if(typeof cell.v==='number'&&/%/.test(String(sourceRow[c]??'')))cell.z='0.0%';else if(typeof cell.v==='number')cell.z='#,##0.##';}
          }
        }
        const widths=Array.from({length:maxCols},(_,c)=>{const values=raw.map(r=>String(r[c]??''));const max=Math.max(8,...values.map(v=>v.length+2));const header=String(raw[0]?.[c]??'');const cap=/Officer|Location|Centre|Address|Name|Remark|Status/i.test(header)?55:34;return Math.min(cap,max);});
        ws['!cols']=widths.map(w=>({wch:w}));ws['!freeze']={xSplit:0,ySplit:4};
        ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:3,c:0},e:{r:Math.max(3,3+raw.length-1),c:lastCol}})};
        ws['!pageSetup']={orientation:'landscape',fitToWidth:1,fitToHeight:0};ws['!margins']={left:0.25,right:0.25,top:0.5,bottom:0.5,header:0.2,footer:0.2};
        XLSX.utils.book_append_sheet(wb,ws,name);exported++;
      }
      const restore=Array.from(document.querySelectorAll('.tab')).find(x=>clean(x.textContent||'')===original) as HTMLElement|undefined;restore?.click();await waitForRender();
      if(!exported)throw Error('No report tables were available to export.');
      const stamp=new Date().toISOString().slice(0,19).replace(/[:T]/g,'-');XLSX.writeFile(wb,`ECI_Hearing_Reports_${stamp}.xlsx`);setMessage(`${exported} reports exported successfully`);
    }catch(e){setMessage(e instanceof Error?e.message:'Excel export failed');}finally{setBusy(false);setTimeout(()=>setMessage(''),4000);}
  };
  return <><button onClick={exportAll} disabled={busy} title="Download all dashboard reports as a professionally formatted Excel workbook" style={{position:'fixed',right:22,top:18,zIndex:99999,border:'1px solid rgba(255,255,255,.18)',borderRadius:8,padding:'9px 14px',fontSize:12,fontWeight:800,cursor:busy?'wait':'pointer',background:'linear-gradient(180deg,#18243a,#101827)',color:'#fff',boxShadow:'0 6px 18px rgba(0,0,0,.28)'}}>{busy?'EXPORTING…':'↓ EXPORT ALL REPORTS'}</button>{message&&<div style={{position:'fixed',right:22,top:62,zIndex:99999,padding:'7px 11px',borderRadius:7,fontSize:11,fontWeight:700,background:'#101827',color:'#dbe7ff',border:'1px solid rgba(255,255,255,.12)'}}>{message}</div>}</>;
}
