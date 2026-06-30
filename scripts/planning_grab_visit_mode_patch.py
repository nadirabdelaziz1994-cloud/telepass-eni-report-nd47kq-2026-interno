from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

OLD_SELECT = """<div><label>Grab & Go</label><select id="planGrab"><option value="no">No</option><option value="yes">Sì, includili</option></select></div>"""
NEW_SELECT = """<div><label>Periodo planning</label><select id="planPeriodMode" onchange="planTogglePeriodDays()"><option value="days">Numero giorni</option><option value="month">1 mese</option><option value="all">Tutti i PDV</option></select><div class="small-muted" style="margin-top:4px">Numero giorni usa solo quei giorni lavorativi. 1 mese resta nel mese. Tutti i PDV continua anche oltre il mese se serve.</div></div><div><label>Giorni</label><input id="planPeriodDays" type="number" min="1" max="31" value="15"></div><div><label>Planning precedente / JSON</label><input id="planPrev" type="file" accept=".csv,.txt,.xls,.json"><div class="small-muted" style="margin-top:4px">Carica il JSON/CSV già creato: il sito esclude i PV già presenti e riparte dal giorno lavorativo successivo.</div></div><div><label>Grab & Go</label><select id="planGrab"><option value="no">No</option><option value="only_new">Sì, solo non visitati</option><option value="all">Sì, tutti</option></select><div class="small-muted" style="margin-top:4px">Se scegli “Sì, tutti”, i Grab&Go già visitati possono rientrare nel planning e verranno rimessi come non visitati nella pagina Grab & Go.</div></div><div style="align-self:end"><button class="btn light" type="button" onclick="planDownloadJson()">Scarica JSON</button></div>"""

NEW_FUNCTIONS = r'''
let PLAN_PREV = {};
let PLAN_PREV_FILE = null;
function planGrabVisitState(){try{if(typeof grabState==='function')return grabState();return JSON.parse(localStorage.getItem('grabGoVisits')||'{}')||{};}catch(e){return{};}}
function planGrabIsVisited(p){const s=planGrabVisitState();const x=s[String(p&&p.pdv||'')];return !!(x&&Number(x.month));}
async function planEnsureGrabVisitsLoaded(){try{if(typeof loadGrabRemote==='function')await loadGrabRemote();}catch(e){}}
async function planClearGrabVisitedFromPlan(items){const mode=document.getElementById('planGrab')?.value||'no';if(mode!=='all')return;const s=planGrabVisitState();const pdvs=[...new Set((items||[]).filter(p=>p&&p.is_grab&&s[p.pdv]&&Number(s[p.pdv].month)).map(p=>p.pdv))];if(!pdvs.length)return;pdvs.forEach(pdv=>delete s[pdv]);if(typeof saveGrabState==='function')saveGrabState(s);else localStorage.setItem('grabGoVisits',JSON.stringify(s));if(typeof saveGrabRemote==='function'){for(const pdv of pdvs){try{await saveGrabRemote(pdv,0);}catch(e){}}}if(typeof renderGrabGo==='function')renderGrabGo();}
function planIso(d){return new Date(d).toISOString().slice(0,10);}
function planDateLabel(d){return new Date(d).toLocaleDateString('it-IT',{weekday:'short',day:'2-digit',month:'2-digit',year:'numeric'});}
function planDateOnly(d){return new Date(d).toLocaleDateString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric'});}
function planTogglePeriodDays(){const mode=document.getElementById('planPeriodMode')?.value||'days';const box=document.getElementById('planPeriodDays');if(box)box.disabled=mode!=='days';}
function planNextWorkdayAfter(d){const x=new Date(d);x.setDate(x.getDate()+1);while([0,6].includes(x.getDay()))x.setDate(x.getDate()+1);return x;}
function planExtendWorkdaysFrom(startDate,n){const out=[];let d=new Date(startDate);while(out.length<n){if(![0,6].includes(d.getDay()))out.push(new Date(d));d.setDate(d.getDate()+1);}return out;}
function planSameMonth(dateIso,month){return String(dateIso||'').slice(0,7)===String(month||'');}
function planPrevDatesInMonth(month){return Object.values(PLAN_PREV||{}).filter(d=>planSameMonth(d,month)).sort();}
function planPrevDoneSet(month){const s=new Set();Object.entries(PLAN_PREV||{}).forEach(([pdv,d])=>{if(planSameMonth(d,month))s.add(String(pdv).padStart(5,'0'));});return s;}
function planNormDateInput(v){const s=String(v||'').trim();let a=s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](20\d{2})/),b=s.match(/(20\d{2})[\-.](\d{1,2})[\-.](\d{1,2})/);if(a)return a[3]+'-'+a[2].padStart(2,'0')+'-'+a[1].padStart(2,'0');if(b)return b[1]+'-'+b[2].padStart(2,'0')+'-'+b[3].padStart(2,'0');return '';}
function planCollectPrevJson(node,map){if(Array.isArray(node)){node.forEach(x=>planCollectPrevJson(x,map));return;}if(!node||typeof node!=='object')return;const raw=node.pdv||node.PDV||node['n° PV']||node['n PV']||node.n_pv||node.punto_vendita||node['n° pv']||'';const pdv=String(raw||'').replace(/\D+/g,'').padStart(5,'0');const d=planNormDateInput(node.date||node.data||node.DATA||node.dateOnly||node.Data||node.dateLabel||'');if(pdv&&d)map[pdv]=d;Object.values(node).forEach(v=>{if(v&&typeof v==='object')planCollectPrevJson(v,map);});}
function planParsePrev(t){const txt=String(t||''),map={};try{const js=JSON.parse(txt);planCollectPrevJson(js,map);if(Object.keys(map).length)return map;}catch(e){}txt.split(/\n+/).forEach(line=>{const p=(line.match(/\b\d{3,6}\b/)||[])[0];if(!p)return;const d=planNormDateInput(line);if(d)map[p.padStart(5,'0')]=d;});return map;}
async function planLoadPrevIfNeeded(){const f=document.getElementById('planPrev')?.files?.[0]||null;if(!f){PLAN_PREV={};PLAN_PREV_FILE=null;return;}if(PLAN_PREV_FILE===f.name)return;const txt=await f.text();PLAN_PREV=planParsePrev(txt);PLAN_PREV_FILE=f.name;alert('Planning precedente caricato: '+Object.keys(PLAN_PREV).length+' PV trovati');}
function planFirstDay(month,baseDays){const ds=planPrevDatesInMonth(month);if(!ds.length)return baseDays[0]||null;return planNextWorkdayAfter(new Date(ds[ds.length-1]+'T00:00:00'));}
function planSelectedDays(month){const mode=document.getElementById('planPeriodMode')?.value||'days',base=planWorkdays(month);if(!base.length)return[];let start=planFirstDay(month,base)||base[0];let days=base.filter(d=>d>=start);if(mode==='days'){const n=Math.max(1,Number(document.getElementById('planPeriodDays')?.value||15));if(!days.length)days=planExtendWorkdaysFrom(start,n);return days.slice(0,n);}if(mode==='month')return days;if(mode==='all'){if(!days.length)days=[start];return days;}return days;}
function planKm2(a,b){if(!a||!b||a.lat==null||b.lat==null)return 0;const R=6371,dLat=(b.lat-a.lat)*Math.PI/180,dLon=(b.lng-a.lng)*Math.PI/180,la1=a.lat*Math.PI/180,la2=b.lat*Math.PI/180;const x=Math.sin(dLat/2)**2+Math.cos(la1)*Math.cos(la2)*Math.sin(dLon/2)**2;return 2*R*Math.atan2(Math.sqrt(x),Math.sqrt(1-x));}
function planVisitMin2(p){return p&&p.is_grab?15:45;}
function planTravelMin2(k){return Math.round((k/55)*60+8);}
function planMinText(n){const h=Math.floor(n/60),m=n%60;return h+'h '+String(m).padStart(2,'0');}
function planAssignSmart(ordered,days,start,allowOverflow){let di=0,last=start,mins=0,count=0,out=[],curDays=days.slice();if(!curDays.length)curDays=planExtendWorkdaysFrom(new Date(),1);for(const p of ordered){while(di>=curDays.length){if(!allowOverflow)return out;curDays=curDays.concat(planExtendWorkdaysFrom(planNextWorkdayAfter(curDays[curDays.length-1]),5));}let k=planKm2(last,p),add=planTravelMin2(k)+planVisitMin2(p),maxCount=p.is_grab?7:6;let newDay=count>0&&((count>=maxCount)||(mins+add>420&&count>=2)||(mins+add>540));if(newDay){di++;if(di>=curDays.length){if(!allowOverflow)return out;curDays=curDays.concat(planExtendWorkdaysFrom(planNextWorkdayAfter(curDays[curDays.length-1]),5));}mins=0;count=0;last=start;k=planKm2(last,p);add=planTravelMin2(k)+planVisitMin2(p);}const day=curDays[di]||new Date();const st=9*60+mins+planTravelMin2(planKm2(last,p));out.push(Object.assign({},p,{date:planIso(day),dateLabel:planDateLabel(day),dateOnly:planDateOnly(day),time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),travel_km:Math.round(k),visit_min:planVisitMin2(p),day_load:planMinText(mins+add)}));mins+=add;count++;last=p;}return out;}
async function generatePlanning(){const agent=document.getElementById('planAgent')?.value||'',month=document.getElementById('planMonth')?.value||'',grabMode=document.getElementById('planGrab')?.value||'no',startText=document.getElementById('planStart')?.value||'',periodMode=document.getElementById('planPeriodMode')?.value||'days';if(!agent||!month){alert('Scegli agente e mese');return;}planTogglePeriodDays();await planLoadPrevIfNeeded();await planEnsureGrabVisitsLoaded();const done=planPrevDoneSet(month);let pts=planPoints().filter(p=>{if(p.agent!==agent||p.lat==null||p.lng==null||done.has(String(p.pdv).padStart(5,'0')))return false;if(!p.is_grab)return true;if(grabMode==='no')return false;if(grabMode==='only_new')return !planGrabIsVisited(p);return true;});const noCoords=((APP.planning_data&&APP.planning_data.without_coordinates)||[]).filter(p=>p.agent===agent).length;if(!pts.length){alert('Nessun PV con coordinate per questo agente, oppure sono già tutti nel planning precedente caricato');return;}const days=planSelectedDays(month);const start=planStartPoint(pts,startText);const ordered=planOrder(pts,start,month);PLAN.items=planAssignSmart(ordered,days,start,periodMode==='all');PLAN.source=`${agent}_${month}`;await planClearGrabVisitedFromPlan(PLAN.items);renderPlanningTable(noCoords,days.length);}
function planDownloadJson(){if(!PLAN||!PLAN.items||!PLAN.items.length){alert('Prima crea il planning');return;}const data={created_at:new Date().toISOString(),agent:document.getElementById('planAgent')?.value||'',month:document.getElementById('planMonth')?.value||'',period_mode:document.getElementById('planPeriodMode')?.value||'',period_days:document.getElementById('planPeriodDays')?.value||'',items:PLAN.items};const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json;charset=utf-8'}));a.download='planning_'+(data.agent||'agente')+'_'+(data.month||'mese')+'.json';a.click();}
try{setTimeout(planTogglePeriodDays,0);}catch(e){}
'''

GENERATE_RE = re.compile(r"function generatePlanning\(\)\{const agent=.*?renderPlanningTable\(noCoords,days\.length\);\}", re.S)


def patch_html(html: str) -> str:
    changed = False
    if OLD_SELECT in html:
        html = html.replace(OLD_SELECT, NEW_SELECT, 1)
        changed = True
    elif 'id="planGrab"' in html and 'Sì, solo non visitati' not in html:
        print('Attenzione: select Grab & Go trovato ma formato non riconosciuto')

    if 'planGrabIsVisited' not in html or ('planPeriodMode' in html and 'planAssignSmart' not in html):
        html2, n = GENERATE_RE.subn(lambda _m: NEW_FUNCTIONS.strip(), html, count=1)
        if n:
            html = html2
            changed = True
        else:
            print('Attenzione: funzione generatePlanning non trovata')

    return html if changed else html


def main():
    patched = 0
    for name in ["index.html", "Telepass_ENI_sito_v6.html"]:
        path = DOCS_DIR / name
        if not path.exists():
            continue
        old = path.read_text(encoding="utf-8")
        new = patch_html(old)
        if new != old:
            path.write_text(new, encoding="utf-8")
            patched += 1
    print(f"Planning Grab&Go visit mode patch completata: {patched} file aggiornati")


if __name__ == "__main__":
    main()
