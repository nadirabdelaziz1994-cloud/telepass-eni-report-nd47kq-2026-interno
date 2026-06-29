from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

NEW_EXPORT_ROWS_META = r'''window.exportRowsMeta = exportRowsMeta = function(){
    const header=['n° WEEK','DATA','ORA ','n° PV','AC','RZV','CR','Regione','Provincia','Città ','Indirizzo','FOCAL POINT ENI','MY WORLD','CONFERMA ENI'];
    const meta=[{kind:'header',row:header}];
    const max=maxTripDays();
    function marker(txt){return {kind:'marker',row:[txt,'','','','','','','','','','','','','']};}
    function dataRow(p){return {kind:'data',plan:p,row:[weekNum(p.date),p.dateOnly,'',p.pdv,acVal(p),p.rzv||'',p.cr||'',p.region||'',p.province||'',p.city||'',p.address||'',p.focal||'',p.agent_display||document.getElementById('agent').value,'']};}

    const transferDates=[];
    const seen={};
    PLAN.forEach(p=>{
      if(isTrasfertaPoint(p) && !seen[p.date]){
        seen[p.date]=true;
        transferDates.push(p.date);
      }
    });
    const selectedDates = new Set(max ? transferDates.slice(0,max) : transferDates);
    let first=-1,last=-1;
    PLAN.forEach((p,i)=>{
      if(selectedDates.has(p.date)){
        if(first<0) first=i;
        last=i;
      }
    });

    PLAN.forEach((p,idx)=>{
      if(idx===first) meta.push(marker('INIZIO TRASFERTA'));
      meta.push(dataRow(p));
      if(idx===last) meta.push(marker('FINE TRASFERTA'));
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
    print("Limite trasferte applicato: conta giorni totali, non blocchi")


if __name__ == "__main__":
    main()
