import React, { useEffect, useMemo, useRef, useState } from 'react'

type DayEntry = {
  date: string; start: string; end: string; breakMin: number;
  site?: string; task?: string; notes?: string; night?: boolean; travel?: boolean;
};

const pad = (n: number) => (n < 10 ? `0${n}` : `${n}`);
function monthKey(ym: string, profile: string) { return `timesheet:${profile || 'default'}:${ym}`; }
function daysInMonth(ym: string): string[] { const [y,m]=ym.split('-').map(Number); const last=new Date(y,m,0).getDate(); return Array.from({length:last},(_,i)=>`${ym}-${pad(i+1)}`); }
function hmToMinutes(hm: string){ if(!hm) return 0; const [h,m]=hm.split(':').map(Number); return h*60+(m||0); }
function minutesToHM(mins:number){ const s=mins<0?'-':''; const v=Math.abs(mins); const h=Math.floor(v/60); const m=v%60; return `${s}${pad(h)}:${pad(m)}`;}
function computeWorkedMinutes(e: DayEntry){ const start=hmToMinutes(e.start); const end=hmToMinutes(e.end); const raw=end-start-(e.breakMin||0); return Math.max(0,isNaN(raw)?0:raw); }
function isoWeek(d: string){ const date=new Date(d+'T12:00:00'); const t=new Date(date.valueOf()); const dn=(date.getDay()+6)%7; t.setDate(t.getDate()-dn+3); const ft=new Date(t.getFullYear(),0,4); const diff=t.valueOf()-ft.valueOf(); const week=1+Math.round(diff/(7*24*3600*1000)); return {year:t.getFullYear(), week}; }
function toCSV(rows: DayEntry[]){ const header=['Date','Début','Fin','Pause(min)','Heures','Site/Chantier','Tâche','Nuit','Déplacement','Notes']; const body=rows.map(r=>[r.date,r.start,r.end,r.breakMin??0,minutesToHM(computeWorkedMinutes(r)),r.site||'',r.task||'',r.night?'Oui':'Non',r.travel?'Oui':'Non',(r.notes||'').replace(/\n/g,' ')]); return [header,...body].map(a=>a.map(x=>`"${String(x).replace(/"/g,'""')}"`).join(',')).join('\n'); }

interface Profile { employee: string; company: string; defaultBreakMin: number; weeklyThreshold: number; }
const defaultProfile: Profile = { employee:'', company:'E2B', defaultBreakMin:0, weeklyThreshold:35*60 };

export default function App(){
  const today=new Date(); const defaultYM=`${today.getFullYear()}-${pad(today.getMonth()+1)}`;
  const [ym,setYM]=useState<string>(defaultYM);
  const [profile,setProfile]=useState<Profile>(()=>{ try{ const raw=localStorage.getItem('timesheet:profile'); return raw?{...defaultProfile,...JSON.parse(raw)}:defaultProfile;}catch{return defaultProfile;} });
  const [rows,setRows]=useState<DayEntry[]>(()=>{ try{const raw=localStorage.getItem(monthKey(defaultYM,profile.employee)); if(raw) return JSON.parse(raw);}catch{} return daysInMonth(defaultYM).map(d=>({date:d,start:'',end:'',breakMin:profile.defaultBreakMin})); });
  const [compact,setCompact]=useState(false);
  const fileRef=useRef<HTMLInputElement|null>(null);

  useEffect(()=>{ try{ const raw=localStorage.getItem(monthKey(ym,profile.employee)); if(raw) setRows(JSON.parse(raw)); else setRows(daysInMonth(ym).map(d=>({date:d,start:'',end:'',breakMin:profile.defaultBreakMin}))); }catch{} },[ym,profile.employee]);
  useEffect(()=>{ localStorage.setItem('timesheet:profile',JSON.stringify(profile)); localStorage.setItem(monthKey(ym,profile.employee),JSON.stringify(rows)); },[rows,ym,profile]);

  const totals=useMemo(()=>{ const minutes=rows.reduce((a,r)=>a+computeWorkedMinutes(r),0); const weekly=new Map<string,number>(); for(const r of rows){ if(!r.start||!r.end) continue; const {year,week}=isoWeek(r.date); const key=`${year}-W${pad(week)}`; weekly.set(key,(weekly.get(key)||0)+computeWorkedMinutes(r)); } const overtimeByWeek:Record<string,number>={}; weekly.forEach((v,k)=>overtimeByWeek[k]=Math.max(0,v-profile.weeklyThreshold)); const overtime=Object.values(overtimeByWeek).reduce((a,b)=>a+b,0); return {minutes,overtime,weekly:overtimeByWeek}; },[rows,profile.weeklyThreshold]);

  function setCell(i:number, patch: Partial<DayEntry>){ setRows(prev=>prev.map((r,idx)=>idx===i?{...r,...patch}:r)); }
  function clearMonth(){ if(!confirm('Effacer toutes les entrées du mois ?')) return; setRows(daysInMonth(ym).map(d=>({date:d,start:'',end:'',breakMin:profile.defaultBreakMin}))); }
  function copyDown(i:number){ const src=rows[i]; if(!src) return; setRows(prev=>prev.map((r,idx)=> idx>i && !r.start && !r.end ? {...r,start:src.start,end:src.end,breakMin:src.breakMin,site:src.site,task:src.task,night:src.night,travel:src.travel} : r)); }
  function download(filename:string,data:string,mime:string){ const blob=new Blob([data],{type:mime}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename; a.click(); URL.revokeObjectURL(a.href); }
  async function handleImport(file: File){ const reader=new FileReader(); const isXLSX=/xlsx?$/.test(file.name.toLowerCase()); reader.onload=async (e)=>{ const content=e.target?.result; if(!content) return; try{ if(isXLSX && content instanceof ArrayBuffer){ const XLSXmod=await import('xlsx'); const XLSX:any=(XLSXmod as any).default ?? XLSXmod; const wb=XLSX.read(content); const sheet=wb.Sheets[wb.SheetNames[0]]; const aoa=XLSX.utils.sheet_to_json<any[]>(sheet,{header:1}); const text=aoa.map((row:any[])=>row.map((x)=>`"${String(x ?? '').replace(/"/g,'""')}"`).join(',')).join('\n'); mergeImported(parseCSV(text)); } else if(typeof content==='string'){ mergeImported(parseCSV(content)); } }catch(e){ alert("Impossible d'importer le fichier. Vérifiez le format."); } }; if(isXLSX) reader.readAsArrayBuffer(file); else reader.readAsText(file,'utf-8'); }
  function parseCSV(text: string): DayEntry[]{ const lines=text.split(/\r?\n/).filter(Boolean); if(lines.length<=1) return []; const headers=lines[0].split(',').map(h=>h.replace(/^\"|\"$/g,'')); const idx=(n:string)=>headers.findIndex(h=>h.toLowerCase()===n.toLowerCase()); const cols={date:idx('Date'),start:idx('Début'),end:idx('Fin'),pause:idx('Pause(min)'),site:idx('Site/Chantier'),task:idx('Tâche'),nuit:idx('Nuit'),travel:idx('Déplacement'),notes:idx('Notes')}; return lines.slice(1).map(l=>{ const parts=l.match(/\"([^\"]*(?:\"\"[^\"]*)*)\"(?=,|$)/g)?.map(s=>s.replace(/^\"|\"$/g,'').replace(/\"\"/g,'"')) || l.split(',').map(x=>x.replace(/^\"|\"$/g,'')); const get=(i:number)=>(i>=0?parts[i]??'':''); return {date:get(cols.date),start:get(cols.start),end:get(cols.end),breakMin:Number(get(cols.pause)||0),site:get(cols.site),task:get(cols.task),night:/oui/i.test(get(cols.nuit)),travel:/oui/i.test(get(cols.travel)),notes:get(cols.notes)} as DayEntry; }); }
  function mergeImported(imported: DayEntry[]){ if(!imported.length) return; const map=new Map(imported.map(r=>[r.date,r] as const)); setRows(prev=>prev.map(r=> map.get(r.date) ? {...r, ...map.get(r.date)!} : r)); }

  const workdays=rows.filter(r=>r.start&&r.end);
  const todayBadge=(d:string)=>{ const t=new Date(); const ds=`${t.getFullYear()}-${pad(t.getMonth()+1)}-${pad(t.getDate())}`; return ds===d; };

  return (<div className="max-w-4xl mx-auto p-4 md:p-8 space-y-4">
    <header className="flex items-center gap-3">
      <img src="/logo.png" alt="E2B" className="w-10 h-10 rounded-lg shadow" />
      <div><h1 className="text-2xl font-semibold">E2B — Pointage d'Heures</h1><div className="text-sm opacity-80">Mois {ym}</div></div>
      <div className="ml-auto flex items-center gap-2"><label className="text-sm">Compact</label><input type="checkbox" checked={false} onChange={()=>{}} /></div>
    </header>

    <section className="card p-4">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div><label className="mb-1 block text-sm font-medium">Mois</label><input className="input" type="month" value={ym} onChange={(e)=>setYM(e.target.value)} /></div>
        <div className="grid grid-cols-2 gap-3">
          <div><label className="mb-1 block text-sm font-medium">Nom et prénom</label><input className="input" placeholder="Nom et prénom" value={profile.employee} onChange={(e)=>setProfile({...profile, employee:e.target.value})} /></div>
          <div><label className="mb-1 block text-sm font-medium">Société</label><input className="input" placeholder="E2B" value={profile.company} onChange={(e)=>setProfile({...profile, company:e.target.value})} /></div>
          <div><label className="mb-1 block text-sm font-medium">Pause par défaut (min)</label><input className="input" type="number" min={0} value={profile.defaultBreakMin} onChange={(e)=>setProfile({...profile, defaultBreakMin:Number(e.target.value||0)})} /></div>
          <div><label className="mb-1 block text-sm font-medium">Seuil hebdo heures sup (min)</label><input className="input" type="number" min={0} value={profile.weeklyThreshold} onChange={(e)=>setProfile({...profile, weeklyThreshold:Number(e.target.value||0)})} /></div>
        </div>
      </div>
    </section>

    <section className="card">
      <div className="p-4 overflow-x-auto">
        <table className="w-full text-sm">
          <thead><tr className="text-left text-slate-600 dark:text-slate-200">
            <th className="p-2">Date</th><th className="p-2">Jour</th><th className="p-2">Début</th><th className="p-2">Fin</th><th className="p-2">Pause (min)</th><th className="p-2">Heures</th><th className="p-2">Actions</th>
          </tr></thead>
          <tbody>{rows.map((r,i)=>{ const mins=computeWorkedMinutes(r); const dObj=new Date(r.date+'T12:00:00'); const wd=isNaN(dObj.getTime())?-1:dObj.getDay(); const jour=isNaN(dObj.getTime())?'':dObj.toLocaleDateString('fr-FR',{weekday:'long'}); const weekendClass = wd===0?'bg-red-50 dark:bg-red-900/20': wd===6?'bg-amber-50 dark:bg-yellow-900/10':'';
            return (<tr key={r.date} className={`${weekendClass}`}>
              <td className={`p-2 font-medium ${mins===0?'text-slate-500':''}`}><div className="flex items-center gap-2">{/* badge */}{/* todayBadge(r.date) && <span className="badge badge-muted">Aujourd'hui</span> */}<span>{r.date}</span></div></td>
              <td className={`p-2 capitalize ${mins===0?'text-slate-500':''}`}>{jour}</td>
              <td className="p-2"><input className="input" type="time" value={r.start} onChange={(e)=>setCell(i,{start:e.target.value})} /></td>
              <td className="p-2"><input className="input" type="time" value={r.end} onChange={(e)=>setCell(i,{end:e.target.value})} /></td>
              <td className="p-2"><input className="input w-24" type="number" min={0} value={r.breakMin} onChange={(e)=>setCell(i,{breakMin:Number(e.target.value||0)})} /></td>
              <td className="p-2 font-mono">{minutesToHM(mins)}</td>
              <td className="p-2"><div className="flex gap-2">
                <button className="btn btn-outline" onClick={()=>copyDown(i)}>Copier</button>
                <button className="btn btn-raised" onClick={()=>{ const text=toCSV([r]); download(`jour_${r.date}.csv`, text, 'text/csv'); }}>Export</button>
              </div></td>
            </tr>); })}</tbody>
        </table>
      </div>
    </section>

    <div className="flex flex-wrap gap-2">
      <button className="btn btn-raised" onClick={()=>{ const ws=rows.filter(r=>r.start&&r.end); const text=toCSV(ws); const date=ym; const a=new Date(); const stamp=`${a.getFullYear()}-${pad(a.getMonth()+1)}-${pad(a.getDate())}`; download(`E2B_pointage_${date}_${stamp}.csv`, text,'text/csv'); }}>Exporter CSV</button>
      <label className="btn btn-outline cursor-pointer">Importer CSV/XLSX<input ref={fileRef} type="file" accept=".csv,.xls,.xlsx" className="hidden" onChange={(e)=>{const f=e.target.files?.[0]; if(f) handleImport(f); (e.currentTarget as HTMLInputElement).value='';}} /></label>
      <button className="btn btn-outline" onClick={()=>localStorage.setItem(monthKey(ym,profile.employee),JSON.stringify(rows))}>Sauvegarder</button>
      <button className="btn btn-outline" onClick={clearMonth}>Effacer le mois</button>
    </div>

    <footer className="text-xs opacity-70 leading-relaxed">
      <p>Fonctionne hors-ligne, données stockées localement. Ajoutez à l'écran d'accueil pour un usage “application”.</p>
    </footer>
  </div>)
}
