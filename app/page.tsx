'use client';
import { useEffect } from 'react';
import * as XLSX from 'xlsx';
import './generated-notices.module.css';
import Dashboard from './EnhancedDashboard';

const uploadRows=(rows:unknown[][],type:'eci'|'master')=>{
  const clean=(v:unknown)=>String(v??'').trim();
  const headerIndex=type==='eci'?rows.findIndex(r=>r.some(v=>clean(v)==='POLLING STATION')):rows.findIndex(r=>r.some(v=>clean(v)==='New P.S. No.'));
  if(headerIndex<0)throw Error(type==='eci'?'ECI header row not found.':'Master header row not found.');
  const heads=rows[headerIndex].map(clean), ix=(x:string)=>heads.indexOf(x);
  const num=(v:unknown)=>{const x=Number(v);return Number.isFinite(x)?x:0};
  if(type==='eci'){
    const p=ix('POLLING STATION');
    const out=rows.slice(headerIndex+1).filter(r=>r.some(v=>clean(v)!=='')).map(r=>{const o:Record<string,unknown>={};heads.forEach((h,i)=>{if(h)o[h]=r[i]});o.ps=Math.trunc(num(r[p]));return o}).filter(r=>Number(r.ps)>=1&&Number(r.ps)<=430);
    const u=new Set(out.map(r=>Number(r.ps)));
    if(out.length!==430||u.size!==430)throw Error(`ECI file has ${out.length} unique valid PS rows; expected 430.`);
    return out;
  }
  const req=['New P.S. No.','Old P.S. No.','Total Anamoly/Discripancy','Total No Mapping','Grand Total','All BLO Name','Supervisor Name','Officer Name (NEW)','Hearing Centre','Address of P.S.','LOCALITY','POLLING AREA','Total Voters'];
  const miss=req.filter(x=>ix(x)<0); if(miss.length)throw Error(`Master sheet is missing: ${miss.join(', ')}`);
  const out=rows.slice(headerIndex+1).filter(r=>r[ix('New P.S. No.')]!==''&&r[ix('New P.S. No.')]!=null).map(r=>({ps:Math.trunc(num(r[ix('New P.S. No.')])),oldPs:r[ix('Old P.S. No.')]===''?null:num(r[ix('Old P.S. No.')]),anomaly:num(r[ix('Total Anamoly/Discripancy')]),mapping:num(r[ix('Total No Mapping')]),grandTotal:num(r[ix('Grand Total')]),blo:clean(r[ix('All BLO Name')]),supervisor:clean(r[ix('Supervisor Name')]),officer:clean(r[ix('Officer Name (NEW)')]),hearingCentre:clean(r[ix('Hearing Centre')]),address:clean(r[ix('Address of P.S.')]),locality:clean(r[ix('LOCALITY')]),pollingArea:clean(r[ix('POLLING AREA')]),voters:num(r[ix('Total Voters')])}));
  const u=new Set(out.map(r=>r.ps));
  if(out.length!==430||u.size!==430)throw Error(`Master file has ${out.length} unique valid PS rows; expected 430.`);
  return out;
};

function DashboardEnhancer(){
  useEffect(()=>{
    const sortState=new WeakMap<HTMLTableElement,{index:number,asc:boolean}>();
    const ensureUploadPanel=()=>{
      if(document.getElementById('eciWorkbookUploadPanel'))return;
      const panel=document.createElement('div');
      panel.id='eciWorkbookUploadPanel';
      panel.style.cssText='position:fixed;top:76px;right:18px;z-index:99999;background:#111827;border:1px solid #334155;border-radius:12px;padding:10px 12px;box-shadow:0 12px 30px rgba(0,0,0,.35);font-family:Arial,sans-serif;min-width:260px;color:#e5e7eb';
      panel.innerHTML=`<div style=\"font-weight:800;font-size:12px;margin-bottom:8px;letter-spacing:.2px\">DATA UPLOAD</div><div style=\"display:flex;gap:7px;flex-wrap:wrap\"><label style=\"cursor:pointer;background:#1d4ed8;color:white;padding:7px 10px;border-radius:7px;font-size:11px;font-weight:700\">UPLOAD MASTER EXCEL<input id=\"eciMasterUploadInput\" type=\"file\" accept=\".xlsx,.xls\" style=\"display:none\"></label><label style=\"cursor:pointer;background:#047857;color:white;padding:7px 10px;border-radius:7px;font-size:11px;font-weight:700\">UPLOAD ECI EXCEL<input id=\"eciEciUploadInput\" type=\"file\" accept=\".xlsx,.xls\" style=\"display:none\"></label></div><div id=\"eciUploadStatus\" style=\"font-size:10px;color:#94a3b8;margin-top:8px\">Master: ${localStorage.getItem('eci-master')?'Loaded':'Not loaded'} · ECI: ${localStorage.getItem('eci-current')?'Loaded':'Not loaded'}</div>`;
      document.body.appendChild(panel);
      const status=panel.querySelector('#eciUploadStatus') as HTMLElement;
      const handle=async(file:File,type:'master'|'eci')=>{
        status.textContent=`Reading ${type==='master'?'Master':'ECI'} workbook…`;
        try{
          const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
          const preferred=type==='eci'?'sirNoticeGenerate':'Rough Data';
          const ws=wb.Sheets[preferred]||wb.Sheets[wb.SheetNames[0]];
          if(!ws)throw Error('Workbook has no readable worksheet.');
          const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''}) as unknown[][];
          const parsed=uploadRows(rows,type);
          if(type==='master'){
            localStorage.setItem('eci-master',JSON.stringify(parsed));
            status.textContent='Master uploaded successfully — reloading dashboard…';
          }else{
            localStorage.setItem('eci-current',JSON.stringify(parsed));
            const now=new Date().toLocaleString('en-IN',{hour12:false});
            localStorage.setItem('eci-updated',now);
            const nums=(v:unknown)=>{const x=Number(v);return Number.isFinite(x)?x:0};
            const item={time:now,kind:'ECI Upload',rows:parsed.length,generated:parsed.reduce((a:any,r:any)=>a+nums(r['Notice Generated']),0),pending:parsed.reduce((a:any,r:any)=>a+nums(r['Pending for Notice Generation']),0),delivered:parsed.reduce((a:any,r:any)=>a+nums(r['Notice Delivered']),0),pendingDelivery:parsed.reduce((a:any,r:any)=>a+nums(r['Notice Pending Delivery']),0)};
            const old=JSON.parse(localStorage.getItem('eci-history')||'[]');
            localStorage.setItem('eci-history',JSON.stringify([item,...(Array.isArray(old)?old:[])].slice(0,20)));
            status.textContent='ECI uploaded successfully — reloading dashboard…';
          }
          window.setTimeout(()=>window.location.reload(),250);
        }catch(err){status.textContent=(err instanceof Error?err.message:'Could not read workbook.');}
      };
      (panel.querySelector('#eciMasterUploadInput') as HTMLInputElement).addEventListener('change',e=>{const f=(e.target as HTMLInputElement).files?.[0];if(f)handle(f,'master')});
      (panel.querySelector('#eciEciUploadInput') as HTMLInputElement).addEventListener('change',e=>{const f=(e.target as HTMLInputElement).files?.[0];if(f)handle(f,'eci')});
    };
    ensureUploadPanel();
    const parse=(v:string)=>{const x=v.replace(/,/g,'').replace(/%/g,'').trim();return x!==''&&Number.isFinite(Number(x))?Number(x):v.toLowerCase()};
    const normalize=(v:string)=>v.replace(/\s+/g,' ').trim().toUpperCase();

    const filterPSWise=(table:HTMLTableElement,value:string)=>{
      Array.from(table.tBodies[0]?.rows||[]).forEach(row=>{
        const status=row.cells[row.cells.length-1]?.dataset.categoryStatus||'';
        row.style.display=value==='All Status'||status.split(' | ').includes(value)?'':'none';
      });
    };

    const enhancePSWiseStatus=(table:HTMLTableElement)=>{
      const headers=Array.from(table.querySelectorAll('thead tr:first-child th')) as HTMLTableCellElement[];
      const labels=headers.map(h=>normalize(h.innerText||''));
      if(!labels.includes('ANOM PENDING GEN')||!labels.includes('NO MAP PENDING GEN'))return;
      const statusHeader=headers[labels.indexOf('STATUS')];
      if(!statusHeader)return;
      let select=statusHeader.querySelector('select') as HTMLSelectElement|null;
      if(!select){
        statusHeader.dataset.statusDropdownEnhanced='1';
        statusHeader.classList.remove('sortable');
        statusHeader.removeAttribute('title');
        select=document.createElement('select');
        select.setAttribute('aria-label','PS Wise Category Status');
        select.style.cssText='min-width:220px;font-size:11px;font-weight:700;padding:5px 8px;border-radius:6px';
        const options=['All Status','ANOMALY · PENDING GENERATION','NO MAPPING · PENDING GENERATION','ANOMALY · GENERATED','NO MAPPING · GENERATED','ANOMALY · CLEAR','NO MAPPING · CLEAR'];
        options.forEach(x=>{const o=document.createElement('option');o.value=x;o.textContent=x;select!.appendChild(o)});
        statusHeader.replaceChildren(select);
        select.addEventListener('click',e=>e.stopPropagation());
        select.addEventListener('change',()=>filterPSWise(table,select!.value));
      }
      Array.from(table.tBodies[0]?.rows||[]).forEach(row=>{
        const cells=row.cells;
        const anPend=Number((cells[6]?.innerText||'0').replace(/,/g,''));
        const nmPend=Number((cells[8]?.innerText||'0').replace(/,/g,''));
        const anGen=Number((cells[5]?.innerText||'0').replace(/,/g,''));
        const nmGen=Number((cells[7]?.innerText||'0').replace(/,/g,''));
        const cell=cells[cells.length-1];
        if(!cell)return;
        const statuses:string[]=[];
        if(anPend>0)statuses.push('ANOMALY · PENDING GENERATION');
        else if(anGen>0)statuses.push('ANOMALY · GENERATED');
        else statuses.push('ANOMALY · CLEAR');
        if(nmPend>0)statuses.push('NO MAPPING · PENDING GENERATION');
        else if(nmGen>0)statuses.push('NO MAPPING · GENERATED');
        else statuses.push('NO MAPPING · CLEAR');
        const value=statuses.join(' | ');
        cell.dataset.categoryStatus=value;
        cell.textContent=value;
      });
      select=statusHeader.querySelector('select') as HTMLSelectElement|null;
      if(select)filterPSWise(table,select.value||'All Status');
    };

    const enhanceGenerated=(table:HTMLTableElement)=>{
      const headers=Array.from(table.querySelectorAll('thead tr:first-child th')) as HTMLTableCellElement[];
      const labels=headers.map(h=>normalize(h.innerText||''));
      if(!labels.includes('ANOMALY')||!labels.includes('NO MAPPING'))return;

      // Generated Notices is strictly a generated-notice delivery register.
      // Never show pending-generation status here; that belongs in PS Wise/Pending.
      Array.from(table.tBodies[0]?.rows||[]).forEach(row=>{
        const cells=row.cells;
        const anGen=Number((cells[5]?.innerText||'0').replace(/,/g,''));
        const anDel=Number((cells[6]?.innerText||'0').replace(/,/g,''));
        const anPdel=Number((cells[7]?.innerText||'0').replace(/,/g,''));
        const nmGen=Number((cells[9]?.innerText||'0').replace(/,/g,''));
        const nmDel=Number((cells[10]?.innerText||'0').replace(/,/g,''));
        const nmPdel=Number((cells[11]?.innerText||'0').replace(/,/g,''));
        const generated=anGen+nmGen;
        const delivered=anDel+nmDel;
        const pendingDelivery=anPdel+nmPdel;
        const cell=cells[cells.length-2];
        if(!cell||generated<=0)return;
        const type=anGen>0&&nmGen>0?'MIXED':anGen>0?'ANOMALY':'NO MAPPING';
        let base='CHECK';
        if(delivered>=generated)base='FULLY DELIVERED';
        else if(delivered>0&&pendingDelivery>0)base='PARTIAL DELIVERY';
        else if(delivered===0&&pendingDelivery>0)base='NOT DELIVERED';
        const next=`${base} · ${type}`;
        if(cell.dataset.generatedStatus!==next){
          cell.dataset.generatedStatus=next;
          const badge=document.createElement('span');
          badge.className='deliveryBadge';
          badge.textContent=next;
          cell.replaceChildren(badge);
        }
      });
    };

    const enhanceGeneratedFilter=()=>{
      const active=document.querySelector('.tab.active')?.textContent?.trim()||'';
      if(active!=='Generated Notices')return;
      const selects=Array.from(document.querySelectorAll('select')) as HTMLSelectElement[];
      const status=selects.find(x=>x.getAttribute('aria-label')==='Status')||selects.find(x=>Array.from(x.options).some(o=>normalize(o.textContent||'').includes('FULLY DELIVERED')));
      if(!status)return;
      const valid=['All','NO MAPPING · FULLY DELIVERED','NO MAPPING · PARTIAL DELIVERY','NO MAPPING · NOT DELIVERED','ANOMALY · FULLY DELIVERED','ANOMALY · PARTIAL DELIVERY','ANOMALY · NOT DELIVERED','MIXED · FULLY DELIVERED','MIXED · PARTIAL DELIVERY','MIXED · NOT DELIVERED'];
      if(status.dataset.generatedFilterEnhanced!=='1'){
        status.dataset.generatedFilterEnhanced='1';
        const current=status.value;
        status.replaceChildren(...valid.map(x=>{const o=document.createElement('option');o.value=x;o.textContent=x==='All'?'All Status':x;return o}));
        status.value=valid.includes(current)?current:'All';
        if(status.value!==current)status.dispatchEvent(new Event('change',{bubbles:true}));
      }
    };

    const enhanceOverviewAnomalyTotal=()=>{
      const active=document.querySelector('.tab.active')?.textContent?.trim()||'';
      if(active!=='Overview')return;
      let masterTotal=0;
      try{
        const raw=localStorage.getItem('eci-master');
        const rows=raw?JSON.parse(raw):[];
        if(Array.isArray(rows))masterTotal=rows.reduce((x,r)=>x+Number(r?.anomaly||0),0);
      }catch{}
      if(!masterTotal)return;
      document.querySelectorAll('[data-anomaly-master-total="1"]').forEach(el=>el.remove());
      const heads=Array.from(document.querySelectorAll('body *')).filter(el=>el.children.length===0 && (el.textContent||'').trim().toUpperCase()==='ANOMALY');
      const heading=heads.find(el=>{let p=el.parentElement;for(let i=0;i<6&&p;i++,p=p.parentElement){const t=p.textContent||'';if(t.includes('Generated')&&t.includes('Pending Gen')&&t.includes('Delivered')&&t.includes('Pending Del'))return true;}return false;});
      if(!heading)return;
      let card=heading.parentElement as HTMLElement|null;
      while(card){const t=card.textContent||'';if(t.includes('Generated')&&t.includes('Pending Gen')&&t.includes('Delivered')&&t.includes('Pending Del'))break;card=card.parentElement;}
      if(!card)return;
      const row=document.createElement('div');
      row.setAttribute('data-anomaly-master-total','1');
      row.style.cssText='display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid rgba(255,255,255,.08);';
      row.innerHTML='<span style="font-weight:700">Total Anomaly</span><strong style="font-size:15px">'+masterTotal.toLocaleString('en-IN')+'</strong>';
      card.insertBefore(row,card.firstElementChild?.nextElementSibling||card.firstElementChild);
    };
    const attach=(table:HTMLTableElement)=>{
      enhancePSWiseStatus(table);
      enhanceGenerated(table);
      const headers=Array.from(table.querySelectorAll('thead tr:first-child th')) as HTMLTableCellElement[];
      const active=document.querySelector('.tab.active')?.textContent?.trim()||'';
      headers.forEach((h,index)=>{
        if(h.dataset.enhancedSort==='1'||h.dataset.statusDropdownEnhanced==='1')return;
        if(h.getAttribute('colspan'))return;
        const label=(h.innerText||'').replace(/\s[↑↓]$/,'').trim();
        if(normalize(label)==='STATUS')return;
        h.dataset.enhancedSort='1';
        h.classList.add('sortable');
        h.title='Click to sort';
        h.addEventListener('click',ev=>{
          ev.stopPropagation();
          let cellIndex=index;
          const currentLabel=(h.innerText||'').replace(/\s[↑↓]$/,'').trim();
          if(active==='Generated Notices'){
            const map:Record<string,number>={PS:0,'Old PS':1,BLO:2,Officer:3,'Hearing Location':4,'Delivery %':16,Hearings:17};
            if(!(currentLabel in map))return;
            cellIndex=map[currentLabel];
          }
          const old=sortState.get(table);
          const asc=old?.index===cellIndex?!old.asc:true;
          sortState.set(table,{index:cellIndex,asc});
          const body=table.tBodies[0];
          if(!body)return;
          const rows=Array.from(body.rows).filter(r=>r.style.display!=='none');
          const hidden=Array.from(body.rows).filter(r=>r.style.display==='none');
          rows.sort((a,b)=>{
            const aa=parse(a.cells[cellIndex]?.innerText||'');
            const bb=parse(b.cells[cellIndex]?.innerText||'');
            const c=typeof aa==='number'&&typeof bb==='number'?aa-bb:String(aa).localeCompare(String(bb),undefined,{numeric:true,sensitivity:'base'});
            return asc?c:-c;
          });
          [...rows,...hidden].forEach(r=>body.appendChild(r));
          headers.forEach(x=>{if(!x.querySelector('select'))x.textContent=x.textContent?.replace(/\s[↑↓]$/,'')||''});
          h.textContent=(h.textContent||'').replace(/\s[↑↓]$/,'')+(asc?' ↑':' ↓');
        });
      });
    };

    let decorating=false;
    const decorate=()=>{
      if(decorating)return;
      decorating=true;
      try{
        document.querySelectorAll('table').forEach(t=>attach(t as HTMLTableElement));
        enhanceGeneratedFilter();
        enhanceOverviewAnomalyTotal();
      }finally{decorating=false}
    };
    const observer=new MutationObserver(decorate);
    observer.observe(document.body,{subtree:true,childList:true});
    decorate();
    return()=>observer.disconnect();
  },[]);
  return null;
}

export default function Page(){return <><Dashboard/><DashboardEnhancer/></>}
