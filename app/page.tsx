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
        if(Array.isArray(rows))masterTotal=rows.reduce((a,r)=>a+Number(r?.anomaly||0),0);
      }catch{}
      if(!masterTotal)return;
      const headings=Array.from(document.querySelectorAll('body *')).filter(el=>el.children.length===0 && (el.textContent||'').trim().toUpperCase()==='ANOMALY');
      const heading=headings.find(el=>{const box=el.parentElement?.parentElement;return !!box && (box.textContent||'').includes('Pending Gen') && (box.textContent||'').includes('Delivered');});
      if(!heading)return;
      const box=heading.parentElement?.parentElement;
      if(!box)return;
      if(box.querySelector('[data-anomaly-master-total="1"]'))return;
      const row=document.createElement('div');
      row.setAttribute('data-anomaly-master-total','1');
      row.style.cssText='display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid rgba(255,255,255,.08);';
      row.innerHTML=`<span style="font-weight:600">Total Anomaly (Master)</span><strong style="font-size:15px">${masterTotal.toLocaleString('en-IN')}</strong>`;
      const first=box.querySelector('div');
      if(first?.parentElement===box)box.insertBefore(row,first);else box.prepend(row);
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
