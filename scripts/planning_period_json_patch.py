from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

OLD_FORM = '''        <div><label>Mese</label><input id="month" type="month"></div>
        <div><label>Punto di partenza</label><input id="start" placeholder="Città o codice PV"></div>
        <div><label>Grab & Go</label><select id="grab"><option value="no">No</option><option value="yes">Sì, includili</option></select></div>
        <div><label>Planning precedente</label><input id="prev" type="file" accept=".csv,.txt,.xls"></div>'''
NEW_FORM = '''        <div><label>Mese</label><input id="month" type="month"></div>
        <div><label>Periodo planning</label><select id="periodMode" onchange="togglePeriodDays()"><option value="days">Numero giorni</option><option value="month">1 mese</option><option value="all">Tutti i PDV</option></select></div>
        <div><label>Giorni</label><input id="periodDays" type="number" min="1" max="31" value="15"></div>
        <div><label>Punto di partenza</label><input id="start" placeholder="Città o codice PV"></div>
        <div><label>Grab & Go</label><select id="grab"><option value="no">No</option><option value="yes">Sì, includili</option></select></div>
        <div><label>Planning precedente / JSON</label><input id="prev" type="file" accept=".csv,.txt,.xls,.json"></div>'''

OLD_NOTE = '''      <div class="muted" style="margin-top:8px">Export Excel con le stesse colonne del file planning originale. I Grab & Go sono evidenziati, senza colonne extra.</div>'''
NEW_NOTE = '''      <div class="muted" style="margin-top:8px">Periodo: “Numero giorni” usa solo quei giorni lavorativi; “1 mese” resta nel mese; “Tutti i PDV” continua anche oltre il mese se serve. Se carichi un planning precedente/JSON, i PV già presenti nel mese scelto vengono esclusi e il nuovo planning riparte dal giorno lavorativo successivo all’ultima data trovata.</div>
      <div class="muted" style="margin-top:6px">Export Excel con le stesse colonne del file planning originale. I Grab & Go sono evidenziati, senza colonne extra.</div>'''

OLD_ACTION = '''        <button class="btn light" onclick="downloadCsv()">Scarica CSV</button>
        <button class="btn light" onclick="loadRemote(false)">Ricarica modifiche definitive</button>'''
NEW_ACTION = '''        <button class="btn light" onclick="downloadCsv()">Scarica CSV</button>
        <button class="btn light" onclick="downloadJson()">Scarica JSON</button>
        <button class="btn light" onclick="loadRemote(false)">Ricarica modifiche definitive</button>'''

NEW_JS = r'''
function nextWorkdayAfter(d){const x=new Date(d);x.setDate(x.getDate()+1);while(true){const h=holidays(x.getFullYear()),w=x.getDay();if(w!==0&&w!==6&&!h.has(iso(x)))return new Date(x);x.setDate(x.getDate()+1);}}
function extendWorkdaysFrom(startDate,n){const out=[];let d=new Date(startDate);while(out.length<n){const h=holidays(d.getFullYear()),w=d.getDay();if(w!==0&&w!==6&&!h.has(iso(d)))out.push(new Date(d));d.setDate(d.getDate()+1);}return out;}
function togglePeriodDays(){const mode=document.getElementById('periodMode')?.value||'days';const box=document.getElementById('periodDays');if(box)box.disabled=mode!=='days';}
function sameMonthIso(dateIso,mv){return String(dateIso||'').slice(0,7)===String(mv||'');}
function prevDatesInMonth(mv){return Object.values(PREV||{}).filter(d=>sameMonthIso(d,mv)).sort();}
function prevDoneSet(mv){const s=new Set();Object.entries(PREV||{}).forEach(([pdv,d])=>{if(sameMonthIso(d,mv))s.add(normPdv(pdv));});return s;}
function firstPlanningDay(mv,baseDays){const ds=prevDatesInMonth(mv);if(!ds.length)return baseDays[0]||null;return nextWorkdayAfter(new Date(ds[ds.length-1]+'T00:00:00'));}
function periodDays(mv){const mode=document.getElementById('periodMode')?.value||'days',base=workdays(mv);if(!base.length)return[];let start=firstPlanningDay(mv,base)||base[0];let days=base.filter(d=>d>=start);if(mode==='days'){const n=Math.max(1,Number(document.getElementById('periodDays')?.value||15));return days.slice(0,n);}if(mode==='month')return days;if(mode==='all'){if(!days.length)days=[start];return days;}return days;}
function maxDailyCountFor(p){return isGrab(p)?7:6;}
function assignSmart(ordered,days,start,allowOverflow){let di=0,last=start,mins=0,count=0,out=[],curDays=days.slice();if(!curDays.length)curDays=extendWorkdaysFrom(new Date(),1);for(const p of ordered){while(di>=curDays.length){if(!allowOverflow)return out;curDays=curDays.concat(extendWorkdaysFrom(nextWorkdayAfter(curDays[curDays.length-1]),5));}let k=km(last,p),add=travelMin(k)+visitMin(p);let maxCount=maxDailyCountFor(p);let newDay=count>0&&((count>=maxCount)||(mins+add>420&&count>=2)||(mins+add>540));if(newDay){di++;if(di>=curDays.length){if(!allowOverflow)return out;curDays=curDays.concat(extendWorkdaysFrom(nextWorkdayAfter(curDays[curDays.length-1]),5));}mins=0;count=0;last=start;k=km(last,p);add=travelMin(k)+visitMin(p);}const day=curDays[di]||new Date();const st=9*60+mins+travelMin(km(last,p));out.push(Object.assign({},p,{date:iso(day),dateLabel:dateLabel(day),dateOnly:dateOnly(day),time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),travel_km:Math.round(k),visit_min:visitMin(p),day_load:minText(mins+add)}));mins+=add;count++;last=p;}return out;}
function generatePlanning(){const agent=document.getElementById('agent').value,mv=document.getElementById('month').value,inc=document.getElementById('grab').value==='yes',mode=document.getElementById('periodMode')?.value||'days';if(!agent||!mv){alert('Scegli agente e mese');return;}const done=prevDoneSet(mv);let pts=allPoints().filter(p=>p.agent_display===agent&&(inc||!isGrab(p))&&p.lat!=null&&p.lng!=null&&!done.has(normPdv(p.pdv)));if(!pts.length){alert('Nessun PV con coordinate per questo agente, oppure sono già tutti nel planning precedente caricato');return;}const days=periodDays(mv),start=startPoint(pts,document.getElementById('start').value);PLAN=assignSmart(order(pts,start,mv),days,start,mode==='all');renderTable(days.length);}
function recalc(){const days=periodDays(document.getElementById('month').value),start=startPoint(PLAN,document.getElementById('start').value),mode=document.getElementById('periodMode')?.value||'days';PLAN=assignSmart(PLAN,days,start,mode==='all');renderTable(days.length);}
function moveRow(i,d){const j=i+d;if(j<0||j>=PLAN.length)return;[PLAN[i],PLAN[j]]=[PLAN[j],PLAN[i]];recalc();}
function delRow(i){PLAN.splice(i,1);recalc();}
function exportJsonData(){return {created_at:new Date().toISOString(),agent:document.getElementById('agent').value||'',month:document.getElementById('month').value||'',period_mode:document.getElementById('periodMode')?.value||'',period_days:document.getElementById('periodDays')?.value||'',items:PLAN};}
function downloadJson(){if(!PLAN.length){alert('Prima crea il planning');return;}const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(exportJsonData(),null,2)],{type:'application/json;charset=utf-8'}));a.download='planning_'+(document.getElementById('agent').value||'agente')+'_'+(document.getElementById('month').value||'mese')+'.json';a.click();}
function normDateInput(v){const s=String(v||'').trim();let a=s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](20\d{2})/),b=s.match(/(20\d{2})[\-.](\d{1,2})[\-.](\d{1,2})/);if(a)return a[3]+'-'+a[2].padStart(2,'0')+'-'+a[1].padStart(2,'0');if(b)return b[1]+'-'+b[2].padStart(2,'0')+'-'+b[3].padStart(2,'0');return '';}
function collectPrevJson(node,map){if(Array.isArray(node)){node.forEach(x=>collectPrevJson(x,map));return;}if(!node||typeof node!=='object')return;const pdv=normPdv(node.pdv||node.PDV||node['n° PV']||node['n PV']||node.n_pv||node.punto_vendita||'');const d=normDateInput(node.date||node.data||node.DATA||node.dateOnly||node.Data||'');if(pdv&&d)map[pdv]=d;Object.values(node).forEach(v=>{if(v&&typeof v==='object')collectPrevJson(v,map);});}
function parsePrev(t){const txt=String(t||'');const map={};try{const js=JSON.parse(txt);collectPrevJson(js,map);if(Object.keys(map).length)return map;}catch(e){}txt.split(/\n+/).forEach(line=>{const p=(line.match(/\b\d{3,6}\b/)||[])[0];if(!p)return;const d=normDateInput(line);if(d)map[p.padStart(5,'0')]=d;});return map;}
'''


def patch_html(html: str) -> str:
    changed = False
    if 'id="periodMode"' not in html and OLD_FORM in html:
        html = html.replace(OLD_FORM, NEW_FORM, 1)
        changed = True
    if 'downloadJson()' not in html and OLD_ACTION in html:
        html = html.replace(OLD_ACTION, NEW_ACTION, 1)
        changed = True
    if 'Periodo: “Numero giorni”' not in html and OLD_NOTE in html:
        html = html.replace(OLD_NOTE, NEW_NOTE, 1)
        changed = True

    # Replace generate/recalc/parsePrev block while keeping renderTable and export functions.
    if 'function assignSmart(' not in html:
        html, n1 = re.subn(r"function generatePlanning\(\).*?function renderTable", NEW_JS.strip() + "\nfunction renderTable", html, count=1, flags=re.S)
        if n1:
            changed = True
        else:
            print('Attenzione: generatePlanning non trovato')
        # Remove old recalc/move/del definitions if still present before exportRows.
        html, n2 = re.subn(r"function recalc\(\).*?function exportRows", "function exportRows", html, count=1, flags=re.S)
        if n2:
            changed = True
        # Remove old parsePrev because NEW_JS already supplies the new one.
        html, n3 = re.subn(r"function parsePrev\(t\)\{.*?\}\s*\n\s*document\.getElementById\('addPdv'\)", "document.getElementById('addPdv')", html, count=1, flags=re.S)
        if n3:
            changed = True
    if "togglePeriodDays();renderAll();loadRemote(true);" not in html:
        html = html.replace("renderAll();loadRemote(true);", "togglePeriodDays();renderAll();loadRemote(true);", 1)
        changed = True
    return html if changed else html


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato")
        return
    old = path.read_text(encoding="utf-8")
    new = patch_html(old)
    if new != old:
        path.write_text(new, encoding="utf-8")
        print("Planning periodo/JSON patch completata")
    else:
        print("Planning periodo/JSON patch: nessuna modifica")


if __name__ == "__main__":
    main()
