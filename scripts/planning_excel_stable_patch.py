from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

ISGRAB_OLD = "function isGrab(p){return !!p.is_grab;}"
ISGRAB_NEW = """function pvKind(p){const g=!!(p&&p.is_grab),t=!!(p&&p.is_tp);if(g&&t)return 'dual';if(g)return 'grab';return 'tp';}
function isGrabOnly(p){return pvKind(p)==='grab';}
function isDual(p){return pvKind(p)==='dual';}
function isGrab(p){return pvKind(p)==='grab'||pvKind(p)==='dual';}
function typeLabel(p){const k=pvKind(p);return k==='dual'?'Grab & Go + TPoint':(k==='grab'?'Solo Grab & Go':'Solo TPoint');}"""

VISIT_OLD = "function visitMin(p){return isGrab(p)?15:45;}"
VISIT_NEW = "function visitMin(p){const k=pvKind(p);return k==='grab'?15:(k==='dual'?60:45);}"

GENERATE_OLD = "function generatePlanning(){const agent=document.getElementById('agent').value,mv=document.getElementById('month').value,inc=document.getElementById('grab').value==='yes';if(!agent||!mv){alert('Scegli agente e mese');return;}let pts=allPoints().filter(p=>p.agent_display===agent&&(inc||!isGrab(p))&&p.lat!=null&&p.lng!=null);if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}const days=workdays(mv),start=startPoint(pts,document.getElementById('start').value);PLAN=assign(order(pts,start,mv),days,start);renderTable(days.length);}"
GENERATE_NEW = "function generatePlanning(){const agent=document.getElementById('agent').value,mv=document.getElementById('month').value,inc=document.getElementById('grab').value==='yes';if(!agent||!mv){alert('Scegli agente e mese');return;}let pts=allPoints().filter(p=>p.agent_display===agent&&(inc||!isGrabOnly(p))&&p.lat!=null&&p.lng!=null);if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}const days=workdays(mv),start=startPoint(pts,document.getElementById('start').value);PLAN=assign(order(pts,start,mv),days,start);renderTable(days.length);}"

RENDER_NEW = r'''function renderTable(workdayCount){const days=[...new Set(PLAN.map(p=>p.date))].length,grabOnly=PLAN.filter(isGrabOnly).length,dual=PLAN.filter(isDual).length;const rows=PLAN.map((p,i)=>{const k=pvKind(p),cls=k==='dual'?'dual':(k==='grab'?'grab-only':'tp'),pill=k==='dual'?'<span class="pill warn">Grab & Go + TPoint</span>':(k==='grab'?'<span class="pill grab">Solo Grab & Go</span>':'<span class="pill">Solo TPoint</span>');return '<tr class="'+cls+'"><td><button class="btn light small" onclick="moveRow('+i+',-1)">↑</button> <button class="btn light small" onclick="moveRow('+i+',1)">↓</button> <button class="btn bad small" onclick="delRow('+i+')">x</button></td><td>'+esc(p.dateLabel)+'</td><td>'+esc(p.pdv)+'</td><td>'+pill+'</td><td class="city"><b>'+esc(p.city)+'</b><span>'+esc(p.address||'')+'</span></td><td>'+esc(p.province||'')+'</td><td>'+esc(p.region||'')+'</td><td>'+esc(p.rzv||'')+'</td><td>'+esc(p.cr||'')+'</td><td>'+esc(p.focal||'')+'</td><td class="num">'+fmt(p.travel_km||0)+' km</td><td class="num">'+fmt(p.visit_min||0)+' min</td><td>'+esc(p.day_load||'')+'</td></tr>';}).join('');document.getElementById('result').innerHTML='<style>.grab-only{box-shadow:inset 5px 0 0 #2563eb;background:#eff6ff}.dual{box-shadow:inset 5px 0 0 #eab308;background:#fff7d6}</style><section class="metric"><div class="box"><h4>PV pianificati</h4><div class="big">'+fmt(PLAN.length)+'</div></div><div class="box"><h4>Solo Grab & Go</h4><div class="big">'+fmt(grabOnly)+'</div></div><div class="box"><h4>Grab & Go + TPoint</h4><div class="big">'+fmt(dual)+'</div></div><div class="box"><h4>Giorni usati</h4><div class="big">'+fmt(days)+'</div><div class="muted">Lavorativi disponibili: '+fmt(workdayCount||0)+'</div></div></section><section class="card"><div class="muted" style="margin-bottom:8px">Blu = solo Grab & Go · Giallo = Grab & Go + Telepass Point · Normale = solo Telepass Point.</div><div class="table-wrap"><table><thead><tr><th>Modifica</th><th>Data</th><th>PV</th><th>Tipo</th><th>Città / indirizzo</th><th>Prov.</th><th>Regione</th><th>RZV</th><th>CR</th><th>Focal Point ENI</th><th>Km</th><th>Visita</th><th>Carico giorno</th></tr></thead><tbody>'+rows+'</tbody></table></div></section>';}'''


def replace_between(text, start_marker, end_marker, replacement):
    start = text.find(start_marker)
    if start == -1:
        raise RuntimeError(f"start marker not found: {start_marker[:40]}")
    end = text.find(end_marker, start)
    if end == -1:
        raise RuntimeError(f"end marker not found: {end_marker[:40]}")
    return text[:start] + replacement + text[end:]


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, patch saltata")
        return
    html = path.read_text(encoding="utf-8")
    if "function pvKind(p)" not in html:
        html = html.replace(ISGRAB_OLD, ISGRAB_NEW, 1)
        html = html.replace(VISIT_OLD, VISIT_NEW, 1)
        html = html.replace(GENERATE_OLD, GENERATE_NEW, 1)
        html = replace_between(html, "function renderTable(workdayCount){", "function recalc(){", RENDER_NEW)
        path.write_text(html, encoding="utf-8")
        print("Patch stabile tipi applicata")
    else:
        print("Patch stabile tipi già presente")

    try:
        from planning_real_xlsx_patch import main as real_xlsx_main
        real_xlsx_main()
    except Exception as exc:
        raise RuntimeError(f"Errore export XLSX reale: {exc}") from exc

    try:
        from planning_transferte_patch import main as transferte_main
        transferte_main()
    except Exception as exc:
        raise RuntimeError(f"Errore patch trasferte: {exc}") from exc

    try:
        from planning_transferte_limit_patch import main as transferte_limit_main
        transferte_limit_main()
    except Exception as exc:
        raise RuntimeError(f"Errore limite giorni trasferta: {exc}") from exc

    try:
        from planning_overflow_safe_patch import main as overflow_safe_main
        overflow_safe_main()
    except Exception as exc:
        raise RuntimeError(f"Errore overflow safe planning: {exc}") from exc

    try:
        from planning_capacity_patch import main as capacity_main
        capacity_main()
    except Exception as exc:
        raise RuntimeError(f"Errore capacity planning: {exc}") from exc

    try:
        from planning_fill_month_patch import main as fill_month_main
        fill_month_main()
    except Exception as exc:
        raise RuntimeError(f"Errore riempimento mese planning: {exc}") from exc


if __name__ == "__main__":
    main()
