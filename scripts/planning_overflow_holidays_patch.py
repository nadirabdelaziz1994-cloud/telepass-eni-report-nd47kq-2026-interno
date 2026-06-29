from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH_JS = r'''
<script>
(function(){
  function pad(n){return String(n).padStart(2,'0');}
  function iso(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());}
  function parseDateFlex(s, year){
    s=String(s||'').trim();
    if(!s)return null;
    let m=s.match(/^(20\d{2})[-\/\.](\d{1,2})[-\/\.](\d{1,2})$/);
    if(m)return new Date(+m[1],+m[2]-1,+m[3]);
    m=s.match(/^(\d{1,2})[-\/\.](\d{1,2})(?:[-\/\.](20\d{2}))?$/);
    if(m)return new Date(m[3]?+m[3]:year,+m[2]-1,+m[1]);
    return null;
  }
  function easter(y){
    const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),mo=Math.floor((h+l-7*m+114)/31),day=((h+l-7*m+114)%31)+1;
    return new Date(y,mo-1,day);
  }
  function addDays(d,n){const x=new Date(d);x.setDate(x.getDate()+n);return x;}
  window.defaultHolidays = defaultHolidays = function(y){
    const e=easter(y), em=addDays(e,1);
    const arr=[new Date(y,0,1),new Date(y,0,6),em,new Date(y,3,25),new Date(y,4,1),new Date(y,5,2),new Date(y,7,15),new Date(y,10,1),new Date(y,11,8),new Date(y,11,25),new Date(y,11,26)];
    return arr.map(iso);
  };
  window.extraClosedDates = extraClosedDates = function(year){
    const el=document.getElementById('closedDays');
    const txt=el?el.value:'';
    const out=new Set();
    String(txt||'').split(/[\n,;]+/).forEach(part=>{
      part=part.trim(); if(!part)return;
      const bits=part.split(/\s*(?:-|–|—|>)\s*/).filter(Boolean);
      if(bits.length>=2){
        let a=parseDateFlex(bits[0],year), b=parseDateFlex(bits[1],year);
        if(a&&b){if(b<a){const t=a;a=b;b=t;} for(let d=new Date(a);d<=b;d.setDate(d.getDate()+1))out.add(iso(d));}
      }else{
        const d=parseDateFlex(part,year); if(d)out.add(iso(d));
      }
    });
    return out;
  };
  window.closedSetForYear = closedSetForYear = function(year){
    const s=new Set(defaultHolidays(year));
    extraClosedDates(year).forEach(x=>s.add(x));
    extraClosedDates(year+1).forEach(x=>s.add(x));
    return s;
  };
  window.workdays = workdays = function(mv, needed){
    const [y,m]=mv.split('-').map(Number);
    const closed=closedSetForYear(y);
    const days=[];
    let d=new Date(y,m-1,1), guard=0;
    const maxDays=Math.max(31, (needed||0)*2+45);
    while(days.length<maxDays && guard<730){
      const day=d.getDay(), id=iso(d);
      if(day!==0 && day!==6 && !closed.has(id))days.push(id);
      d.setDate(d.getDate()+1); guard++;
    }
    return days;
  };
  const oldGenerate=window.generatePlanning || generatePlanning;
  window.generatePlanning = generatePlanning = function(){
    const agent=document.getElementById('agent').value,mv=document.getElementById('month').value,inc=document.getElementById('grab').value==='yes';
    if(!agent||!mv){alert('Scegli agente e mese');return;}
    let pts=allPoints().filter(p=>p.agent_display===agent&&(inc||!isGrabOnly(p))&&p.lat!=null&&p.lng!=null);
    if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}
    const start=startPoint(pts,document.getElementById('start').value);
    const ordered=order(pts,start,mv);
    const days=workdays(mv, ordered.length);
    PLAN=assign(ordered,days,start);
    renderTable(days.length);
  };
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, patch overflow/ferie saltata")
        return
    html = path.read_text(encoding="utf-8")
    target = '<div><label>Planning precedente</label><input id="prev" type="file" accept=".csv,.txt,.xls"></div>'
    closed = '<div><label>Ferie / chiusure</label><textarea id="closedDays" rows="3" placeholder="10/08/2026-21/08/2026\n24/12/2026"></textarea><p class="muted small">Sabato, domenica e feste nazionali italiane vengono esclusi già in automatico.</p></div>'
    if 'id="closedDays"' not in html and target in html:
        html = html.replace(target, closed + target, 1)
    if 'function defaultHolidays' not in html:
        html = html.replace('</body>', PATCH_JS + '\n</body>', 1)
    path.write_text(html, encoding='utf-8')
    print('Patch overflow giorni e ferie/chiusure applicata')


if __name__ == '__main__':
    main()
