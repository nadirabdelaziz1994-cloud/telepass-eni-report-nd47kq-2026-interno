from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH_JS = r'''
<script>
(function(){
  window.visitMin = visitMin = function(p){
    const k=(typeof pvKind==='function')?pvKind(p):(p&&p.is_grab&&!p.is_tp?'grab':'tp');
    return k==='grab'?15:45;
  };
  window.assign = assign = function(ordered,days,start){
    let di=0,last=start,mins=0,count=0,out=[];
    const monthValue=document.getElementById('month')?document.getElementById('month').value:'';
    const monthPrefix=monthValue?monthValue+'-':'';
    const monthDays=days.filter(d=>iso(d).startsWith(monthPrefix));
    const overflowDays=days.filter(d=>!iso(d).startsWith(monthPrefix)).slice(0,15);
    const usableDays=monthDays.concat(overflowDays);
    const list=usableDays.length?usableDays:days;
    const softDay=480;
    const hardDay=560;
    for(const p of ordered){
      const k=km(last,p), add=travelMin(k)+visitMin(p);
      const overSoft=mins>0 && mins+add>softDay && count>=1;
      const overHard=mins>0 && mins+add>hardDay;
      if((overSoft||overHard) && di<list.length-1){
        di++;mins=0;count=0;last=start;
      }
      const day=list[Math.min(di,list.length-1)]||new Date();
      const travel=travelMin(km(last,p));
      const st=9*60+mins+travel;
      out.push(Object.assign({},p,{date:iso(day),dateLabel:dateLabel(day),dateOnly:dateOnly(day),time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),travel_km:Math.round(k),visit_min:visitMin(p),day_load:minText(mins+add)}));
      mins+=add;count++;last=p;
    }
    return out;
  };
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, capacity patch saltata")
        return
    html = path.read_text(encoding='utf-8')
    if 'softDay=480' not in html:
        html = html.replace('</body>', PATCH_JS + '\n</body>', 1)
    path.write_text(html, encoding='utf-8')
    print('Planning capacity patch applicata')


if __name__ == '__main__':
    main()
