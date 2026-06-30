from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH_JS = r'''
<script>
(function(){
  function monthDaysOnly(days,mv){const pref=mv+'-';return (days||[]).filter(d=>iso(d).startsWith(pref));}
  function usedMonthDays(plan,mv){const pref=mv+'-';return new Set((plan||[]).map(p=>p.date||'').filter(d=>d.startsWith(pref))).size;}
  function shouldFillMonth(ordered,days,start,mv){
    const md=monthDaysOnly(days,mv);
    if(md.length<10 || ordered.length<1)return false;
    const firstPlan=assign(ordered,days,start);
    const used=usedMonthDays(firstPlan,mv);
    return used>0 && used<Math.ceil(md.length*0.75);
  }
  function repeatToFillMonth(ordered,days,start,mv){
    if(!shouldFillMonth(ordered,days,start,mv))return ordered;
    const md=monthDaysOnly(days,mv);
    let out=[];
    for(let cycle=0;cycle<20;cycle++){
      ordered.forEach(p=>out.push(Object.assign({},p,{repeat_cycle:cycle+1})));
      const test=assign(out,days,start);
      if(usedMonthDays(test,mv)>=md.length)break;
    }
    return out;
  }
  window.generatePlanning = generatePlanning = function(){
    const agent=document.getElementById('agent').value,mv=document.getElementById('month').value,inc=document.getElementById('grab').value==='yes';
    if(!agent||!mv){alert('Scegli agente e mese');return;}
    let pts=allPoints().filter(p=>p.agent_display===agent&&(inc||!isGrabOnly(p))&&p.lat!=null&&p.lng!=null);
    if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}
    const days=workdays(mv),start=startPoint(pts,document.getElementById('start').value);
    let ordered=order(pts,start,mv);
    ordered=repeatToFillMonth(ordered,days,start,mv);
    PLAN=assign(ordered,days,start);
    renderTable(days.length);
  };
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, fill-month saltato")
        return
    html = path.read_text(encoding="utf-8")
    if 'repeatToFillMonth' not in html:
        html = html.replace('</body>', PATCH_JS + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Patch riempimento mese applicata")


if __name__ == "__main__":
    main()
