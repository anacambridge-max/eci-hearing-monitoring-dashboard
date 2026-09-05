'use client';
import { useEffect } from 'react';
import './generated-notices.module.css';
import Dashboard from './EnhancedDashboard';

function DashboardEnhancer(){
  useEffect(()=>{
    const sortState=new WeakMap<HTMLTableElement,{index:number,asc:boolean}>();
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
      const statusIndex=labels.indexOf('STATUS');
      if(statusIndex<0)return;
      const statusHeader=headers[statusIndex];
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
        const cell=cells[statusIndex];
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
    const attach=(table:HTMLTableElement)=>{
      enhancePSWiseStatus(table);
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
            const map:Record<string,number>={PS:1,'Old PS':2,BLO:3,Officer:4,'Hearing Location':5,'Delivery %':17,Hearings:18};
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
        const active=document.querySelector('.tab.active')?.textContent?.trim()||'';
        if(active==='Generated Notices'){
          document.querySelectorAll('table tbody tr').forEach(row=>{
            const cells=row.querySelectorAll('td');
            const delivered=Number((cells[15]?.innerText||'0').replace(/,/g,''));
            const pending=Number((cells[16]?.innerText||'0').replace(/,/g,''));
            const anGen=Number((cells[6]?.innerText||'0').replace(/,/g,''));
            const nmGen=Number((cells[10]?.innerText||'0').replace(/,/g,''));
            const badge=cells[19]?.querySelector('.deliveryBadge');
            if(badge){
              const type=anGen>0&&nmGen>0?'MIXED':anGen>0?'ANOMALY':'NO MAPPING';
              const base=pending>0&&delivered>0?'PARTIAL DELIVERY':pending>0?'PENDING DELIVERY':delivered>0?'FULLY DELIVERED':'NOT DELIVERED';
              const next=`${base} · ${type}`;
              if(badge.textContent!==next)badge.textContent=next;
            }
          });
        }
        if(active==='Pending'){
          document.querySelectorAll('table tbody tr').forEach(row=>{
            const cells=row.querySelectorAll('td');
            const pgen=Number((cells[3]?.innerText||'0').replace(/,/g,''));
            const pdel=Number((cells[4]?.innerText||'0').replace(/,/g,''));
            const pct=Number((cells[5]?.innerText||'0').replace(/%/g,''));
            const badge=cells[6]?.querySelector('.deliveryBadge');
            if(badge){
              let next='PENDING GENERATION';
              if(pgen>0&&pdel>0)next='PARTIAL DELIVERY · PENDING GENERATION';
              else if(pdel>0)next='PENDING DELIVERY';
              else if(pgen>0&&pct>0)next='FULLY DELIVERED · PENDING GENERATION';
              if(badge.textContent!==next)badge.textContent=next;
            }
          });
        }
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
