'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

type Master = {
  ps: number; oldPs: number | null; anomaly: number; mapping: number; grandTotal: number;
  blo: string; supervisor: string; officer: string; hearingCentre: string; address: string;
  locality: string; pollingArea: string; voters: number;
};
type Eci = Record<string, unknown> & { ps: number };
type Joined = Eci & { m: Master; cat: Cat; map: MapRow | null };
type Cat = {
  anGen:number; nmGen:number; anPend:number; nmPend:number; anDel:number; nmDel:number;
  anPdel:number; nmPdel:number; gen:number; pgen:number; del:number; pdel:number;
  expectedPendingGen:number; expectedPendingDel:number;
};
type MapRow = { centre:string; officer:string; contact:string };

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
  const heads=rows[0].map(s);
  const ix=(name:string)=>heads.indexOf(name);
  return rows.slice(1).filter(r=>r.some(v=>s(v)!=='')).map(r=>{
    const o:Record<string,unknown>={}; heads.forEach((h,i)=>{if(h)o[h]=r[i]});
    o.ps=Math.trunc(n(r[ix('POLLING STATION')])); return o as Eci;
  }).filter(r=>r.ps>=1&&r.ps<=430);
}

function parseMaster(rows: unknown[][]): Master[] {
  if(rows.length<3) return [];
  const heads=rows[1].map(s); const ix=(name:string)=>heads.indexOf(name);
  return rows.slice(2).filter(r=>r[ix('New P.S. No.')]!==''&&r[ix('New P.S. No.')]!=null).map(r=>({
    ps:Math.trunc(n(r[ix('New P.S. No.')])), oldPs:r[ix('Old P.S. No.')]===''?null:n(r[ix('Old P.S. No.')]),
    anomaly:n(r[ix('Total Anamoly/Discripancy')]), mapping:n(r[ix('Total No Mapping')]), grandTotal:n(r[ix('Grand Total')]),
    blo:s(r[ix('All BLO Name')]), supervisor:s(r[ix('Supervisor Name')]), officer:s(r[ix('Officer Name (NEW)')]),
    hearingCentre:s(r[ix('Hearing Centre')]), address:s(r[ix('Address of P.S.')]), locality:s(r[ix('LOCALITY')]),
    pollingArea:s(r[ix('POLLING AREA')]), voters:n(r[ix('Total Voters')])
  }));
}

function category(e:Eci,m:Master):Cat {
  const gen=n(e['Notice Generated']); const del=n(e['Notice Delivered']);
  // Allocation rule: generated notices are assigned to No Mapping first, then Anomaly.
  const nmGen=Math.min(Math.max(gen,0),Math.max(m.mapping,0)); const anGen=Math.max(gen-nmGen,0);
  const nmDel=Math.min(Math.max(del,0),nmGen); const anDel=Math.min(Math.max(del-nmDel,0),anGen);
  const nmPend=Math.max(m.mapping-nmGen,0); const anPend=Math.max(m.anomaly-anGen,0);
  return {anGen,nmGen,anPend,nmPend,anDel,nmDel,anPdel:Math.max(anGen-anDel,0),nmPdel:Math.max(nmGen-nmDel,0),
    gen,pgen:n(e['Pending for Notice Generation']),del,pdel:n(e['Notice Pending Delivery']),
    expectedPendingGen:n(m.anomaly+m.mapping)-gen, expectedPendingDel:Math.max(gen-del,0)};
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
  const sum=(k:string)=>joined.reduce((a,r)=>a+n(r[k]),0); const cats=(k:keyof Cat)=>joined.reduce((a,r)=>a+n(r.cat[k]),0);
  const fmt=(v:number)=>v.toLocaleString('en-IN');
  const kpis=[['Total PS',430],['Notice Generated',sum('Notice Generated')],['Pending Generation',sum('Pending for Notice Generation')],['Notice Delivered',sum('Notice Delivered')],['Pending Delivery',sum('Notice Pending Delivery')],['Hearings Held',sum('Hearings Held')],['Hearing Lapsed',sum('Hearing Date Lapsed')],['Reschedule Lapsed',sum('Reschedule Date Lapsed')],['DEO Pending',sum('DEO-Status Total Pending')],['DEO >5 Days',sum('DEO-Status Pending GT 5 Days')]];
  const validation=[
    ['ECI PS count',eci.length,430,eci.length===430],['Master PS count',master.length,430,master.length===430],
    ['Anomaly + No Mapping',cats('anGen')+cats('nmGen')+cats('anPend')+cats('nmPend'),master.reduce((a,m)=>a+m.grandTotal,0),cats('anGen')+cats('nmGen')+cats('anPend')+cats('nmPend')===master.reduce((a,m)=>a+m.grandTotal,0)],
    ['Notice delivery',cats('anDel')+cats('nmDel')+cats('anPdel')+cats('nmPdel'),cats('gen'),cats('anDel')+cats('nmDel')+cats('anPdel')+cats('nmPdel')===cats('gen')]
  ];

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
          {([['ANOMALY','anGen','anPend','anDel','anPdel'],['NO MAPPING','nmGen','nmPend','nmDel','nmPdel'],['ECI TOTAL','gen','pgen','del','pdel']] as Array<[string,keyof Cat,keyof Cat,keyof Cat,keyof Cat]>).map(x=><div className="mini" key={x[0]}><h3>{x[0]}</h3><div className="metricrow"><span>Generated</span><b>{fmt(cats(x[1]))}</b></div><div className="metricrow"><span>Pending Generation</span><b>{fmt(cats(x[2]))}</b></div><div className="metricrow"><span>Delivered</span><b>{fmt(cats(x[3]))}</b></div><div className="metricrow"><span>Pending Delivery</span><b>{fmt(cats(x[4]))}</b></div></div>)}
        </div><div className="small" style={{padding:'0 14px 14px'}}>Category allocation uses No-Mapping first, then Anomaly, based on the master category counts. Reconciliation differences are surfaced in Validation.</div></div>
      </>}

      {tab==='PS Wise'&&<div className="section"><h2>PS-wise consolidated status</h2><div className="filters"><input placeholder="Search PS / BLO / locality / polling area" value={q} onChange={e=>setQ(e.target.value)}/><select value={centreFilter} onChange={e=>setCentreFilter(e.target.value)}><option>All</option>{centres.map(c=><option key={c}>{c}</option>)}</select></div><div className="tablewrap"><table className="table"><thead><tr><th>PS</th><th>Old PS</th><th>BLO</th><th>Centre</th><th>Anomaly</th><th>No Mapping</th><th>Generated</th><th>Pending Gen</th><th>Delivered</th><th>Pending Delivery</th><th>Hearings</th><th>DEO &gt;5</th><th>Status</th></tr></thead><tbody>{filtered.map(r=><tr key={r.ps}><td><b>{r.ps}</b></td><td>{r.m.oldPs??''}</td><td>{r.m.blo.replace(/\n/g,' ')}</td><td>{r.map?.centre||r.m.hearingCentre}</td><td>{r.m.anomaly}</td><td>{r.m.mapping}</td><td>{n(r['Notice Generated'])}</td><td>{n(r['Pending for Notice Generation'])}</td><td>{n(r['Notice Delivered'])}</td><td>{n(r['Notice Pending Delivery'])}</td><td>{n(r['Hearings Held'])}</td><td>{n(r['DEO-Status Pending GT 5 Days'])}</td><td className={`status ${status(r)==='CRITICAL'?'critical':status(r)==='PENDING'?'pending':'clear'}`}>{status(r)}</td></tr>)}</tbody></table></div></div>}

      {tab==='Hearing Centres'&&<div className="section"><h2>Hearing-centre reconciliation</h2><div className="tablewrap"><table className="table"><thead><tr><th>Hearing Centre</th><th>Officer</th><th>Contact</th><th>Allocated PS</th><th>Generated</th><th>Pending Gen</th><th>Delivered</th><th>Pending Delivery</th><th>Hearings</th><th>DEO &gt;5</th></tr></thead><tbody>{centres.map(c=>{const rr=joined.filter(r=>r.map?.centre===c);const meta=MAP.find(x=>x[0]===c)!;return <tr key={c}><td>{c}</td><td>{meta[3]}</td><td>{meta[4]}</td><td>{rr.length}</td><td>{rr.reduce((a,r)=>a+n(r['Notice Generated']),0)}</td><td>{rr.reduce((a,r)=>a+n(r['Pending for Notice Generation']),0)}</td><td>{rr.reduce((a,r)=>a+n(r['Notice Delivered']),0)}</td><td>{rr.reduce((a,r)=>a+n(r['Notice Pending Delivery']),0)}</td><td>{rr.reduce((a,r)=>a+n(r['Hearings Held']),0)}</td><td>{rr.reduce((a,r)=>a+n(r['DEO-Status Pending GT 5 Days']),0)}</td></tr>})}</tbody></table></div></div>}

      {tab==='Pending'&&<div className="section"><h2>Operational pending report</h2><div className="tablewrap"><table className="table"><thead><tr><th>PS</th><th>BLO</th><th>Officer</th><th>Centre</th><th>Pending Gen</th><th>Pending Delivery</th><th>Hearing Lapsed</th><th>Reschedule Lapsed</th><th>DEO Pending</th><th>DEO &gt;5</th><th>Priority</th></tr></thead><tbody>{joined.filter(r=>status(r)!=='CLEAR').sort((a,b)=>n(b['DEO-Status Pending GT 5 Days'])-n(a['DEO-Status Pending GT 5 Days'])).map(r=><tr key={r.ps}><td><b>{r.ps}</b></td><td>{r.m.blo.replace(/\n/g,' ')}</td><td>{r.map?.officer||r.m.officer}</td><td>{r.map?.centre||r.m.hearingCentre}</td><td>{n(r['Pending for Notice Generation'])}</td><td>{n(r['Notice Pending Delivery'])}</td><td>{n(r['Hearing Date Lapsed'])}</td><td>{n(r['Reschedule Date Lapsed'])}</td><td>{n(r['DEO-Status Total Pending'])}</td><td>{n(r['DEO-Status Pending GT 5 Days'])}</td><td className={`status ${status(r)==='CRITICAL'?'critical':'pending'}`}>{status(r)}</td></tr>)}</tbody></table></div></div>}

      {tab==='Validation'&&<div className="section"><h2>Data validation & reconciliation</h2><div className="validation">{validation.map(v=><div className={`check ${v[3]?'pass':'warn'}`} key={String(v[0])}><b>{v[0]}</b><div className="small">Actual: {fmt(Number(v[1]))} · Expected: {fmt(Number(v[2]))}</div><strong>{v[3]?'PASS':'WARNING'}</strong></div>)}</div><div className="small" style={{padding:'0 14px 14px'}}>A WARNING is deliberately shown instead of silently forcing category totals to match the ECI download.</div></div>}
      <div className="footer">Source workflow: ECI sirNoticeGenerate upload + fixed 430-PS master/hearing allocation. Data is stored locally in this browser in this version.</div>
    </div>
  </main>;
}
