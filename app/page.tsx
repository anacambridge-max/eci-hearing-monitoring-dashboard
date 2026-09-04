'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

type Master = {
  ps: number; oldPs: number | null; anomaly: number; mapping: number; grandTotal: number;
  blo: string; supervisor: string; officer: string; hearingCentre: string; address: string;
  locality: string; pollingArea: string; voters: number;
};
type Eci = Record<string, unknown> & { ps: number };
type MapRow = { centre: string; officer: string; contact: string };
type Cat = {
  anGen:number; nmGen:number; anPend:number; nmPend:number; anDel:number; nmDel:number;
  anPdel:number; nmPdel:number; gen:number; pgen:number; del:number; pdel:number;
};
type Joined = Eci & { m: Master; cat: Cat; map: MapRow | null };

const MAP: Array<[string,number,number,string,string]> = [
 ['GCSSS, SEC-3 DWARKA(P)',1,50,'SH. PARVEEN KUMAR','9953601073'],['GCSSS, SEC-3 DWARKA(P)',55,60,'SH. PARVEEN KUMAR','9953601073'],['GCSSS, SEC-3 DWARKA(P)',78,79,'SH. PARVEEN KUMAR','9953601073'],['GCSSS, SEC-3 DWARKA(P)',91,96,'SH. PARVEEN KUMAR','9953601073'],
 ['GCSSS, SEC-3 DWARKA(S)',51,54,'SMT. SHASHI BALA','9953312984'],['GCSSS, SEC-3 DWARKA(S)',61,77,'SMT. SHASHI BALA','9953312984'],['GCSSS, SEC-3 DWARKA(S)',80,90,'SMT. SHASHI BALA','9953312984'],['GCSSS, SEC-3 DWARKA(S)',97,109,'SMT. SHASHI BALA','9953312984'],['GCSSS, SEC-3 DWARKA(S)',135,145,'SMT. SHASHI BALA','9953312984'],
 ['GCSSC SEC-22 DWARKA(R)',110,134,'SH. RAKESH KUMAR','7011971522'],['GCSSC SEC-22 DWARKA(R)',331,341,'SH. RAKESH KUMAR','7011971522'],['GCSSC SEC-22 DWARKA(R)',343,344,'SH. RAKESH KUMAR','7011971522'],['GCSSC SEC-22 DWARKA(R)',347,349,'SH. RAKESH KUMAR','7011971522'],
 ['VREC MATIALA',146,234,'SMT. PARUL GUPTA','9667881989'],['VREC MATIALA',276,288,'SMT. PARUL GUPTA','9667881989'],
 ['MCD BOYS PRIMARY SCHOOL, QUTUB VIHAR',235,275,'SH. SUBHASHISH','9868252144'],['MCD BOYS PRIMARY SCHOOL, QUTUB VIHAR',289,330,'SH. SUBHASHISH','9868252144'],
 ['GCSSS SEC 22 DWARKA(V)',342,342,'SH. VIRENDER','9868252144'],['GCSSS SEC 22 DWARKA(V)',345,346,'SH. VIRENDER','9868252144'],['GCSSS SEC 22 DWARKA(V)',350,374,'SH. VIRENDER','9868252144'],
 ['GCSSS GHUMANHERA',375,430,'SH. VIRENDER','9868252144']
];
const centres = [...new Set(MAP.map(x=>x[0]))];
const n=(v:unknown)=>{const x=Number(v);return Number.isFinite(x)?x:0};
const s=(v:unknown)=>String(v??'').trim();
const key=(v:unknown)=>s(v).toLowerCase();
const centre=(ps:number):MapRow|null=>{const x=MAP.find(m=>ps>=m[1]&&ps<=m[2]);return x?{centre:x[0],officer:x[3],contact:x[4]}:null};

function parseEci(rows: unknown[][]): Eci[] {
  if(!rows.length) return [];
  const heads=rows[0].map(s); const ix=(name:string)=>heads.indexOf(name);
  const psIx=ix('POLLING STATION');
  if(psIx<0) throw new Error('ECI sheet is missing the POLLING STATION column.');
  return rows.slice(1).filter(r=>r.some(v=>s(v)!=='')).map(r=>{
    const o:Record<string,unknown>={}; heads.forEach((h,i)=>{if(h)o[h]=r[i]});
    o.ps=Math.trunc(n(r[psIx])); return o as Eci;
  }).filter(r=>r.ps>=1&&r.ps<=430);
}

function parseMaster(rows: unknown[][]): Master[] {
  if(rows.length<3) return [];
  const heads=rows[1].map(s); const ix=(name:string)=>heads.indexOf(name);
  const required=['New P.S. No.','Old P.S. No.','Total Anamoly/Discripancy','Total No Mapping','Grand Total','All BLO Name','Supervisor Name','Officer Name (NEW)','Hearing Centre','Address of P.S.','LOCALITY','POLLING AREA','Total Voters'];
  const missing=required.filter(x=>ix(x)<0); if(missing.length) throw new Error(`Master sheet is missing: ${missing.join(', ')}`);
  return rows.slice(2).filter(r=>r[ix('New P.S. No.')]!==''&&r[ix('New P.S. No.')]!=null).map(r=>({
    ps:Math.trunc(n(r[ix('New P.S. No.')])), oldPs:r[ix('Old P.S. No.')]===''?null:n(r[ix('Old P.S. No.')]),
    anomaly:n(r[ix('Total Anamoly/Discripancy')]), mapping:n(r[ix('Total No Mapping')]), grandTotal:n(r[ix('Grand Total')]),
    blo:s(r[ix('All BLO Name')]), supervisor:s(r[ix('Supervisor Name')]), officer:s(r[ix('Officer Name (NEW)')]),
    hearingCentre:s(r[ix('Hearing Centre')]), address:s(r[ix('Address of P.S.')]), locality:s(r[ix('LOCALITY')]),
    pollingArea:s(r[ix('POLLING AREA')]), voters:n(r[ix('Total Voters')])
  }));
}

/* Official allocation rule confirmed by the user:
   notices are allocated to NO MAPPING first; any balance is allocated to ANOMALY. */
function category(e:Eci,m:Master):Cat {
  const gen=Math.max(n(e['Notice Generated']),0);
  const pgen=Math.max(n(e['Pending for Notice Generation']),0);
  const del=Math.max(n(e['Notice Delivered']),0);
  const pdel=Math.max(n(e['Notice Pending Delivery']),0);

  const nmGen=Math.min(gen,Math.max(m.mapping,0));
  const anGen=Math.min(Math.max(gen-nmGen,0),Math.max(m.anomaly,0));

  const nmRemainingForGeneration=Math.max(m.mapping-nmGen,0);
  const nmPend=Math.min(pgen,nmRemainingForGeneration);
  const anPend=Math.min(Math.max(pgen-nmPend,0),Math.max(m.anomaly-anGen,0));

  const nmDel=Math.min(del,nmGen);
  const anDel=Math.min(Math.max(del-nmDel,0),anGen);

  const nmPdel=Math.min(pdel,Math.max(nmGen-nmDel,0));
  const anPdel=Math.min(Math.max(pdel-nmPdel,0),Math.max(anGen-anDel,0));

  return {anGen,nmGen,anPend,nmPend,anDel,nmDel,anPdel,nmPdel,gen,pgen,del,pdel};
}

function status(r:Joined){
  if(n(r['DEO-Status Pending GT 5 Days'])>0||n(r['Hearing Date Lapsed'])>0||n(r['Reschedule Date Lapsed'])>0)return 'CRITICAL';
  if(n(r['Pending for Notice Generation'])>0||n(r['Notice Pending Delivery'])>0||n(r['DEO-Status Total Pending'])>0)return 'PENDING';
  return 'CLEAR';
}

export default function Dashboard(){
  const [master,setMaster]=useState<Master[]>([]); const [eci,setEci]=useState<Eci[]>([]);
  const [tab,setTab]=useState('Overview'); const [q,setQ]=useState(''); const [centreFilter,setCentreFilter]=useState('All');
  const [updated,setUpdated]=useState(''); const [error,setError]=useState('');
  useEffect(()=>{try{const m=localStorage.getItem('eci-master');const e=localStorage.getItem('eci-current');const t=localStorage.getItem('eci-updated');if(m)setMaster(JSON.parse(m));if(e)setEci(JSON.parse(e));if(t)setUpdated(t)}catch{setError('Saved browser data could not be read.')}} ,[]);

  const upload=async(file:File,type:'eci'|'master')=>{
    setError('');
    try{
      const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
      const ws=type==='eci'?(wb.Sheets['sirNoticeGenerate']||wb.Sheets[wb.SheetNames[0]]):(wb.Sheets['Rough Data']||wb.Sheets[wb.SheetNames[0]]);
      if(!ws) throw new Error('Required worksheet was not found.');
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''}) as unknown[][];
      if(type==='eci'){
        const parsed=parseEci(rows); const unique=new Set(parsed.map(x=>x.ps));
        if(parsed.length!==430) throw new Error(`ECI file has ${parsed.length} valid PS rows; expected 430.`);
        if(unique.size!==430) throw new Error('ECI file contains duplicate PS numbers.');
        setEci(parsed); const now=new Date().toLocaleString('en-IN',{hour12:false}); setUpdated(now);
        localStorage.setItem('eci-current',JSON.stringify(parsed)); localStorage.setItem('eci-updated',now);
      }else{
        const parsed=parseMaster(rows); const unique=new Set(parsed.map(x=>x.ps));
        if(parsed.length!==430||unique.size!==430) throw new Error(`Master file has ${parsed.length} unique PS rows; expected 430.`);
        setMaster(parsed); localStorage.setItem('eci-master',JSON.stringify(parsed));
      }
    }catch(err){setError(err instanceof Error?err.message:'Could not read workbook.');}
  };

  const joined=useMemo<Joined[]>(()=>{const mm=new Map(master.map(m=>[m.ps,m]));return eci.map(e=>{const m=mm.get(e.ps)||{ps:e.ps,oldPs:null,anomaly:0,mapping:0,grandTotal:0,blo:'',supervisor:'',officer:'',hearingCentre:'',address:'',locality:'',pollingArea:'',voters:0};return {...e,m,cat:category(e,m),map:centre(e.ps)};});},[eci,master]);
  const filtered=joined.filter(r=>(centreFilter==='All'||r.map?.centre===centreFilter)&&(q===''||String(r.ps).includes(q)||key(r.m.blo).includes(key(q))||key(r.m.locality).includes(key(q))||key(r.m.pollingArea).includes(key(q))));
  const sum=(k:string)=>joined.reduce((a,r)=>a+n(r[k]),0);
  const cats=(k:keyof Cat)=>joined.reduce((a,r)=>a+n(r.cat[k]),0);
  const fmt=(v:number)=>v.toLocaleString('en-IN');
  const kpis=[['Total PS',430],['Notice Generated',sum('Notice Generated')],['Pending Generation',sum('Pending for Notice Generation')],['Notice Delivered',sum('Notice Delivered')],['Pending Delivery',sum('Notice Pending Delivery')],['Hearings Held',sum('Hearings Held')],['Hearing Lapsed',sum('Hearing Date Lapsed')],['Reschedule Lapsed',sum('Reschedule Date Lapsed')],['DEO Pending',sum('DEO-Status Total Pending')],['DEO >5 Days',sum('DEO-Status Pending GT 5 Days')]];

  const generatedTotal=cats('anGen')+cats('nmGen');
  const pendingGenTotal=cats('anPend')+cats('nmPend');
  const deliveredTotal=cats('anDel')+cats('nmDel');
  const pendingDeliveryTotal=cats('anPdel')+cats('nmPdel');
  const capacityViolations=joined.filter(r=>r.cat.anGen+r.cat.nmGen>r.m.grandTotal).length;
  const validation=[
    ['ECI PS count',eci.length,430,eci.length===430],
    ['Master PS count',master.length,430,master.length===430],
    ['Generated split',generatedTotal,sum('Notice Generated'),generatedTotal===sum('Notice Generated')],
    ['Pending generation split',pendingGenTotal,sum('Pending for Notice Generation'),pendingGenTotal===sum('Pending for Notice Generation')],
    ['Delivered split',deliveredTotal,sum('Notice Delivered'),deliveredTotal===sum('Notice Delivered')],
    ['Pending delivery split',pendingDeliveryTotal,sum('Notice Pending Delivery'),pendingDeliveryTotal===sum('Notice Pending Delivery')],
    ['Category capacity',capacityViolations,0,capacityViolations===0]
  ];

  const categoryRows: Array<[string,keyof Cat,keyof Cat,keyof Cat,keyof Cat]>=[['ANOMALY','anGen','anPend','anDel','anPdel'],['NO MAPPING','nmGen','nmPend','nmDel','nmPdel'],['ECI TOTAL','gen','pgen','del','pdel']];

  return <main className="app">
    <header className="topbar"><div><div className="brand">ECI HEARING & NOTICE MONITORING</div><div className="sub">NCT OF DELHI · 430 POLLING STATIONS · LIVE ECI UPLOAD DASHBOARD</div></div><div className="updated">{updated?`Last upload: ${updated}`:'No ECI upload yet'}</div></header>
    <div className="wrap">
      <div className="toolbar"><div className="upload"><label>Upload ECI Excel<input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&upload(e.target.files[0],'eci')}/></label><label>Load Master (one-time)<input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&upload(e.target.files[0],'master')}/></label></div><div className="notice">Recurring: download ECI → upload latest Excel → dashboard updates.</div></div>
      {error&&<div className="notice" style={{background:'#fff1f0',borderColor:'#fecdca',marginBottom:15}}>{error}</div>}
      {!master.length&&<div className="notice" style={{marginBottom:15}}>First time only: load <b>For Hearing Schedule-3.xlsx</b>. It stays in this browser; after that, only ECI uploads are needed.</div>}
      <div className="tabs">{['Overview','PS Wise','Hearing Centres','Pending','Validation'].map(t=><button className={`tab ${tab===t?'active':''}`} onClick={()=>setTab(t)} key={t}>{t}</button>)}</div>

      {tab==='Overview'&&<>
        <div className="grid">{kpis.map(([l,v])=><div className="card" key={l}><div className="label">{l}</div><div className="value">{fmt(Number(v))}</div></div>)}</div>
        <div className="section"><h2>Notice status — separate Anomaly / No Mapping</h2><div className="split">
          {categoryRows.map(x=><div className="mini" key={x[0]}><h3>{x[0]}</h3><div className="metricrow"><span>Generated</span><b>{fmt(cats(x[1]))}</b></div><div className="metricrow"><span>Pending Generation</span><b>{fmt(cats(x[2]))}</b></div><div className="metricrow"><span>Delivered</span><b>{fmt(cats(x[3]))}</b></div><div className="metricrow"><span>Pending Delivery</span><b>{fmt(cats(x[4]))}</b></div></div>)}
        </div><div className="small" style={{padding:'0 14px 14px'}}>Official allocation rule: No Mapping first, then Anomaly. Every category split is independently reconciled to the corresponding ECI aggregate; capacity violations are reported in Validation.</div></div>
      </>}

      {tab==='PS Wise'&&<div className="section"><h2>PS-wise consolidated status</h2><div className="filters"><input placeholder="Search PS / BLO / locality / polling area" value={q} onChange={e=>setQ(e.target.value)}/><select value={centreFilter} onChange={e=>setCentreFilter(e.target.value)}><option>All</option>{centres.map(c=><option key={c}>{c}</option>)}</select></div><div className="tablewrap"><table className="table"><thead><tr><th>PS</th><th>Old PS</th><th>BLO</th><th>Centre</th><th>Anomaly</th><th>No Mapping</th><th>Anom Gen</th><th>NM Gen</th><th>Anom Pend Gen</th><th>NM Pend Gen</th><th>Anom Delivered</th><th>NM Delivered</th><th>Anom Pend Del</th><th>NM Pend Del</th><th>Hearings</th><th>DEO &gt;5</th><th>Status</th></tr></thead><tbody>{filtered.map(r=><tr key={r.ps}><td><b>{r.ps}</b></td><td>{r.m.oldPs??''}</td><td>{r.m.blo.replace(/\n/g,' ')}</td><td>{r.map?.centre||r.m.hearingCentre}</td><td>{r.m.anomaly}</td><td>{r.m.mapping}</td><td>{r.cat.anGen}</td><td>{r.cat.nmGen}</td><td>{r.cat.anPend}</td><td>{r.cat.nmPend}</td><td>{r.cat.anDel}</td><td>{r.cat.nmDel}</td><td>{r.cat.anPdel}</td><td>{r.cat.nmPdel}</td><td>{n(r['Hearings Held'])}</td><td>{n(r['DEO-Status Pending GT 5 Days'])}</td><td>{status(r)}</td></tr>)}</tbody></table></div></div>}

      {tab==='Hearing Centres'&&<div className="section"><h2>Hearing-centre reconciliation</h2><div className="tablewrap"><table className="table"><thead><tr><th>Hearing Centre</th><th>Officer</th><th>Contact</th><th>Allocated PS</th><th>Generated</th><th>Pending Gen</th><th>Delivered</th><th>Pending Delivery</th><th>Hearings</th><th>DEO &gt;5</th></tr></thead><tbody>{centres.map(c=>{const rows=joined.filter(r=>r.map?.centre===c);const map=MAP.find(x=>x[0]===c);return <tr key={c}><td>{c}</td><td>{map?.[3]||''}</td><td>{map?.[4]||''}</td><td>{rows.length}</td><td>{fmt(rows.reduce((a,r)=>a+n(r['Notice Generated']),0))}</td><td>{fmt(rows.reduce((a,r)=>a+n(r['Pending for Notice Generation']),0))}</td><td>{fmt(rows.reduce((a,r)=>a+n(r['Notice Delivered']),0))}</td><td>{fmt(rows.reduce((a,r)=>a+n(r['Notice Pending Delivery']),0))}</td><td>{fmt(rows.reduce((a,r)=>a+n(r['Hearings Held']),0))}</td><td>{fmt(rows.reduce((a,r)=>a+n(r['DEO-Status Pending GT 5 Days']),0))}</td></tr>})}</tbody></table></div></div>}

      {tab==='Pending'&&<div className="section"><h2>Operational pending report</h2><div className="tablewrap"><table className="table"><thead><tr><th>PS</th><th>BLO</th><th>Officer</th><th>Centre</th><th>Pending Gen</th><th>Pending Delivery</th><th>Hearing Lapsed</th><th>Reschedule Lapsed</th><th>DEO Pending</th><th>DEO &gt;5</th><th>Status</th></tr></thead><tbody>{filtered.filter(r=>n(r['Pending for Notice Generation'])>0||n(r['Notice Pending Delivery'])>0||n(r['Hearing Date Lapsed'])>0||n(r['Reschedule Date Lapsed'])>0||n(r['DEO-Status Total Pending'])>0).map(r=><tr key={r.ps}><td><b>{r.ps}</b></td><td>{r.m.blo.replace(/\n/g,' ')}</td><td>{r.map?.officer||r.m.officer}</td><td>{r.map?.centre||r.m.hearingCentre}</td><td>{n(r['Pending for Notice Generation'])}</td><td>{n(r['Notice Pending Delivery'])}</td><td>{n(r['Hearing Date Lapsed'])}</td><td>{n(r['Reschedule Date Lapsed'])}</td><td>{n(r['DEO-Status Total Pending'])}</td><td>{n(r['DEO-Status Pending GT 5 Days'])}</td><td>{status(r)}</td></tr>)}</tbody></table></div></div>}

      {tab==='Validation'&&<div className="section"><h2>Data validation & reconciliation</h2><div className="split">{validation.map(v=><div className="mini" key={v[0]} style={{borderLeft:`4px solid ${v[3]?'#12b76a':'#f04438'}`}}><h3>{v[0]}</h3><div className="small">Actual: {fmt(Number(v[1]))} · Expected: {fmt(Number(v[2]))}</div><div className="value" style={{fontSize:22}}>{v[3]?'PASS':'WARNING'}</div></div>)}</div><div className="small" style={{paddingTop:12}}>Category allocation is deterministic: No Mapping is filled first from generated notices, pending generation and delivered/pending-delivery balances; Anomaly receives the remaining amount. Validation does not falsely compare ECI workflow totals with the master Grand Total.</div></div>}
    </div>
  </main>;
}
