'use client';
import { useEffect } from 'react';
import './generated-notices.module.css';
import Dashboard from './EnhancedDashboard';

function DashboardEnhancer(){
  useEffect(()=>{
    const sortState=new WeakMap<HTMLTableElement,{index:number,asc:boolean}>();
    const parse=(v:string)=>{const x=v.replace(/,/g,'').replace(/%/g,'').trim();return x!==''&&Number.isFinite(Number(x))?Number(x):v.toLowerCase()};
    const attach=(table:HTMLTableElement)=>{
      const headers=Array.from(table.querySelectorAll('thead tr:first-child th')) as HTMLTableCellElement[];
      const active=document.querySelector('.tab.active')?.textContent?.trim()||'';
      headers.forEach((h,index)=>{
        if(h.dataset.enhancedSort==='1')return;
        h.dataset.enhancedSort='1';
        if(h.getAttribute('colspan'))return;
        h.classList.add('sortable');
        h.title='Click to sort';
        h.addEventListener('click',ev=>{
          ev.stopPropagation();
          let cellIndex=index;
          const label=(h.innerText||'').replace(/\s[↑↓]$/,'').trim();
          if(active==='Generated Notices'){
            const map:Record<string,number>={PS:1,'Old PS':2,BLO:3,Officer:4,'Hearing Location':5,'Delivery %':15,Hearings:16,Status:17};
            if(!(label in map))return;
            cellIndex=map[label];
          }
          const old=sortState.get(table);
          const asc=old?.index===cellIndex?!old.asc:true;
          sortState.set(table,{index:cellIndex,asc});
          const body=table.tBodies[0];
          if(!body)return;
          const rows=Array.from(body.rows);
          rows.sort((a,b)=>{
            const aa=parse(a.cells[cellIndex]?.innerText||'');
            const bb=parse(b.cells[cellIndex]?.innerText||'');
            const c=typeof aa==='number'&&typeof bb==='number'?aa-bb:String(aa).localeCompare(String(bb),undefined,{numeric:true,sensitivity:'base'});
            return asc?c:-c;
          });
          rows.forEach(r=>body.appendChild(r));
          headers.forEach(x=>{x.textContent=x.textContent?.replace(/\s[↑↓]$/,'')||''});
          h.textContent=(h.textContent||'').replace(/\s[↑↓]$/,'')+(asc?' ↑':' ↓');
        });
      });
    };
    const decorate=()=>{
      document.querySelectorAll('table').forEach(t=>attach(t as HTMLTableElement));
      const active=document.querySelector('.tab.active')?.textContent?.trim()||'';
      if(active==='Generated Notices'){
        document.querySelectorAll('table tbody tr').forEach(row=>{
          const cells=row.querySelectorAll('td');
          const delivered=Number((cells[13]?.innerText||'0').replace(/,/g,''));
          const pending=Number((cells[14]?.innerText||'0').replace(/,/g,''));
          const anGen=Number((cells[6]?.innerText||'0').replace(/,/g,''));
          const nmGen=Number((cells[9]?.innerText||'0').replace(/,/g,''));
          const badge=cells[17]?.querySelector('.deliveryBadge');
          if(badge){
            const type=anGen>0&&nmGen>0?'MIXED':anGen>0?'ANOMALY':'NO MAPPING';
            const base=pending>0&&delivered>0?'PARTIAL DELIVERY':pending>0?'PENDING DELIVERY':delivered>0?'FULLY DELIVERED':'NOT DELIVERED';
            badge.textContent=`${base} · ${type}`;
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
            badge.textContent=next;
          }
        });
      }
    };
    const observer=new MutationObserver(decorate);
    observer.observe(document.body,{subtree:true,childList:true});
    decorate();
    return()=>observer.disconnect();
  },[]);
  return null;
}

export default function Page(){return <><Dashboard/><DashboardEnhancer/></>}
