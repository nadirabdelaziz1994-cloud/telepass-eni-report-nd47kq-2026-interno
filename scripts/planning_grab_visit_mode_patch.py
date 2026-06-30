from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

OLD_SELECT = """<div><label>Grab & Go</label><select id="planGrab"><option value="no">No</option><option value="yes">Sì, includili</option></select></div>"""
NEW_SELECT = """<div><label>Grab & Go</label><select id="planGrab"><option value="no">No</option><option value="only_new">Sì, solo non visitati</option><option value="all">Sì, tutti</option></select><div class="small-muted" style="margin-top:4px">Se scegli “Sì, tutti”, i Grab&Go già visitati possono rientrare nel planning e verranno rimessi come non visitati nella pagina Grab & Go.</div></div>"""

NEW_FUNCTIONS = r'''
function planGrabVisitState(){try{if(typeof grabState==='function')return grabState();return JSON.parse(localStorage.getItem('grabGoVisits')||'{}')||{};}catch(e){return{};}}
function planGrabIsVisited(p){const s=planGrabVisitState();const x=s[String(p&&p.pdv||'')];return !!(x&&Number(x.month));}
async function planEnsureGrabVisitsLoaded(){try{if(typeof loadGrabRemote==='function')await loadGrabRemote();}catch(e){}}
async function planClearGrabVisitedFromPlan(items){const mode=document.getElementById('planGrab')?.value||'no';if(mode!=='all')return;const s=planGrabVisitState();const pdvs=[...new Set((items||[]).filter(p=>p&&p.is_grab&&s[p.pdv]&&Number(s[p.pdv].month)).map(p=>p.pdv))];if(!pdvs.length)return;pdvs.forEach(pdv=>delete s[pdv]);if(typeof saveGrabState==='function')saveGrabState(s);else localStorage.setItem('grabGoVisits',JSON.stringify(s));if(typeof saveGrabRemote==='function'){for(const pdv of pdvs){try{await saveGrabRemote(pdv,0);}catch(e){}}}if(typeof renderGrabGo==='function')renderGrabGo();}
async function generatePlanning(){const agent=document.getElementById('planAgent')?.value||'',month=document.getElementById('planMonth')?.value||'',grabMode=document.getElementById('planGrab')?.value||'no',startText=document.getElementById('planStart')?.value||'';if(!agent||!month){alert('Scegli agente e mese');return;}await planEnsureGrabVisitsLoaded();let pts=planPoints().filter(p=>{if(p.agent!==agent||p.lat==null||p.lng==null)return false;if(!p.is_grab)return true;if(grabMode==='no')return false;if(grabMode==='only_new')return !planGrabIsVisited(p);return true;});const noCoords=((APP.planning_data&&APP.planning_data.without_coordinates)||[]).filter(p=>p.agent===agent).length;if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}const days=planWorkdays(month);const start=planStartPoint(pts,startText);const ordered=planOrder(pts,start,month);PLAN.items=planAssign(ordered,days,start);PLAN.source=`${agent}_${month}`;planClearGrabVisitedFromPlan(PLAN.items);renderPlanningTable(noCoords,days.length);}
'''

GENERATE_RE = re.compile(r"function generatePlanning\(\)\{const agent=.*?renderPlanningTable\(noCoords,days\.length\);\}", re.S)


def patch_html(html: str) -> str:
    changed = False
    if OLD_SELECT in html:
        html = html.replace(OLD_SELECT, NEW_SELECT, 1)
        changed = True
    elif 'id="planGrab"' in html and 'Sì, solo non visitati' not in html:
        print('Attenzione: select Grab & Go trovato ma formato non riconosciuto')

    if 'planGrabIsVisited' not in html:
        html2, n = GENERATE_RE.subn(NEW_FUNCTIONS.strip(), html, count=1)
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
