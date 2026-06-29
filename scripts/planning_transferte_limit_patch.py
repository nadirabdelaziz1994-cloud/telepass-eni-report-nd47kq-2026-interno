from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

NEW_EXPORT_ROWS_META = r'''window.exportRowsMeta = exportRowsMeta = function(){
    const header=['n° WEEK','DATA','ORA ','n° PV','AC','RZV','CR','Regione','Provincia','Città ','Indirizzo','FOCAL POINT ENI','MY WORLD','CONFERMA ENI'];
    const meta=[{kind:'header',row:header}];
    const max=maxTripDays();
    function marker(txt){return {kind:'marker',row:[txt,'','','','','','','','','','','','','']};}
    function dataRow(p){return {kind:'data',plan:p,row:[weekNum(p.date),p.dateOnly,'',p.pdv,acVal(p),p.rzv||'',p.cr||'',p.region||'',p.province||'',p.city||'',p.address||'',p.focal||'',p.agent_display||document.getElementById('agent').value,'']};}
    function areaKey(p){return norm(p.region||p.province||p.city||'');}
    function areaLabel(p){return p.region||p.province||p.city||'';}

    const areas={};
    PLAN.forEach(p=>{
      if(!isTrasfertaPoint(p)) return;
      const key=areaKey(p);
      if(!key) return;
      if(!areas[key]) areas[key]={key,label:areaLabel(p),count:0,dates:[],dateSeen:{},distance:0};
      areas[key].count+=1;
      if(!areas[key].dateSeen[p.date]){
        areas[key].dateSeen[p.date]=true;
        areas[key].dates.push(p.date);
      }
      const home=startPoint(PLAN.length?PLAN:allPoints(),document.getElementById('start').value);
      areas[key].distance=Math.max(areas[key].distance, home?km(home,p):0);
    });
    const best=Object.values(areas).sort((a,b)=>{
      if(b.count!==a.count) return b.count-a.count;
      if(b.dates.length!==a.dates.length) return b.dates.length-a.dates.length;
      return b.distance-a.distance;
    })[0];
    const selectedDates = new Set(best ? (max ? best.dates.slice(0,max) : best.dates) : []);
    const selectedArea = best ? best.key : '';

    let first=-1,last=-1;
    PLAN.forEach((p,i)=>{
      if(selectedDates.has(p.date) && areaKey(p)===selectedArea){
        if(first<0) first=i;
        last=i;
      }
    });

    PLAN.forEach((p,idx)=>{
      if(idx===first) meta.push(marker('INIZIO TRASFERTA' + (best && best.label ? ' - ' + best.label : '')));
      meta.push(dataRow(p));
      if(idx===last) meta.push(marker('FINE TRASFERTA' + (best && best.label ? ' - ' + best.label : '')));
    });
    return meta;
  };'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, patch limite trasferte saltata")
        return
    html = path.read_text(encoding="utf-8")
    start_marker = "window.exportRowsMeta = exportRowsMeta = function(){"
    start = html.find(start_marker)
    if start == -1:
        print("exportRowsMeta non trovato: patch limite trasferte saltata")
        return
    end_marker = "\n})();\n</script>"
    end = html.find(end_marker, start)
    if end == -1:
        raise RuntimeError("fine exportRowsMeta non trovata")
    html = html[:start] + NEW_EXPORT_ROWS_META + html[end:]
    path.write_text(html, encoding="utf-8")
    print("Trasferta scelta sulla regione lontana con più PV")


if __name__ == "__main__":
    main()
