from pathlib import Path
p=Path('app/page.tsx')
s=p.read_text()
s=s.replace("import * as XLSX from 'xlsx';\n", "")
old="""    const observer=new MutationObserver(decorate);\n    observer.observe(document.body,{subtree:true,childList:true});\n    decorate();\n"""
new="""    let scheduled=false;\n    const scheduleDecorate=()=>{\n      if(scheduled)return;\n      scheduled=true;\n      requestAnimationFrame(()=>{scheduled=false;decorate()});\n    };\n    const observer=new MutationObserver(scheduleDecorate);\n    observer.observe(document.body,{subtree:true,childList:true});\n    scheduleDecorate();\n"""
if old in s:s=s.replace(old,new,1)
p.write_text(s)

p=Path('app/EnhancedDashboard.tsx')
s=p.read_text()
s=s.replace("import * as XLSX from 'xlsx';\n", "")
old="const upload=async(file:File,type:'eci'|'master')=>{setError('');try{const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}),ws=type==='eci'?(wb.Sheets['sirNoticeGenerate']||wb.Sheets[wb.SheetNames[0]]):(wb.Sheets['Rough Data']||wb.Sheets[wb.SheetNames[0]]);"
new="const upload=async(file:File,type:'eci'|'master')=>{setError('');try{const XLSX=await import('xlsx');const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}),ws=type==='eci'?(wb.Sheets['sirNoticeGenerate']||wb.Sheets[wb.SheetNames[0]]):(wb.Sheets['Rough Data']||wb.Sheets[wb.SheetNames[0]]);"
if old in s:s=s.replace(old,new,1)
oldcat="const anPend=Math.min(Math.max(pgen-nmPend,0),Math.max(m.anomaly-anGen,0));"
newcat="const anPend=Math.max(m.anomaly-anGen,0);"
if oldcat not in s: raise SystemExit('anomaly pending formula not found')
s=s.replace(oldcat,newcat,1)
needle="const joined=useMemo<Joined[]>(()=>{"
bridge="""useEffect(()=>{\n  const bindUploadButtons=()=>{\n    const bind=(needle:string,type:'eci'|'master')=>{\n      const els=Array.from(document.querySelectorAll('button,label,[role=\\\"button\\\"]')) as HTMLElement[];\n      const el=els.find(x=>(x.textContent||'').replace(/\\s+/g,' ').trim().toLowerCase().includes(needle));\n      if(!el||el.dataset.uploadBridge==='1')return;\n      el.dataset.uploadBridge='1';\n      const input=document.createElement('input');\n      input.type='file'; input.accept='.xlsx,.xls';\n      input.style.cssText='position:fixed;left:-10000px;top:-10000px;width:1px;height:1px;opacity:0;';\n      input.addEventListener('change',()=>{const file=input.files?.[0];if(file)void upload(file,type);input.value='';});\n      document.body.appendChild(input);\n      el.addEventListener('click',(ev)=>{ev.preventDefault();ev.stopImmediatePropagation();input.click();},true);\n    };\n    bind('upload eci excel','eci'); bind('load master','master');\n  };\n  bindUploadButtons();\n  const id=window.setInterval(bindUploadButtons,1000);\n  return()=>window.clearInterval(id);\n},[upload]);\n"""
if needle not in s: raise SystemExit('joined marker not found')
if bridge not in s:s=s.replace(needle,bridge+needle,1)
p.write_text(s)
print('patched anomaly pending to master-authoritative total')
