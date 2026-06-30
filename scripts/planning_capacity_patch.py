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
  function isTPointForDay(p){
    const k=(typeof pvKind==='function')?pvKind(p):(p&&p.is_grab&&!p.is_tp?'grab':'tp');
    return k!=='grab';
  }
  window.assign = assign = function(ordered,days,start){
    let di=0,last=start,mins=0,count=0,tpCount=0,grabCount=0,out=[];
    const monthValue=document.getElementById('month')?document.getElementById('month').value:'';
    const monthPrefix=monthValue?monthValue+'-':'';
    const monthDays=days.filter(d=>iso(d).startsWith(monthPrefix));
    const overflowDays=days.filter(d=>!iso(d).startsWith(monthPrefix)).slice(0,15);
    const usableDays=monthDays.concat(overflowDays);
    const list=usableDays.length?usableDays:days;
    const softDay=480;
    const hardDay=570;
    const maxTPoint=6;
    const maxTotal=8;
    function dayAddCost(point, from){
      const travel=travelMin(km(from||start,point));
      return {travel, total:travel+visitMin(point), km:km(from||start,point)};
    }
    function mustNewDay(point, cost){
      const isTp=isTPointForDay(point);
      const nextTp=tpCount+(isTp?1:0);
      const nextGrab=grabCount+(isTp?0:1);
      const nextTotal=count+1;
      if(count<=0) return false;
      if(nextTp>maxTPoint) return true;
      if(nextTotal>maxTotal) return true;
      if(nextTp>=5 && nextGrab>2) return true;
      if(nextTp===4 && nextGrab>3) return true;
      if(mins+cost.total>hardDay) return true;
      if(mins+cost.total>softDay && count>=2) return true;
      return false;
    }
    for(const p of ordered){
      let cost=dayAddCost(p,last);
      if(mustNewDay(p,cost) && di<list.length-1){
        di++;mins=0;count=0;tpCount=0;grabCount=0;last=start;
        cost=dayAddCost(p,last);
      }
      const day=list[Math.min(di,list.length-1)]||new Date();
      const st=9*60+mins+cost.travel;
      const isTp=isTPointForDay(p);
      const newLoad=mins+cost.total;
      out.push(Object.assign({},p,{date:iso(day),dateLabel:dateLabel(day),dateOnly:dateOnly(day),time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),travel_km:Math.round(cost.km),visit_min:visitMin(p),day_load:minText(newLoad)}));
      mins=newLoad;count++;if(isTp)tpCount++;else grabCount++;last=p;
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
    start = html.find('<script>\n(function(){\n  window.visitMin = visitMin = function(p){')
    marker = 'Planning capacity patch applicata'
    if start != -1 and 'function isTPointForDay' not in html[start:start+2500]:
        end = html.find('</script>', start)
        if end != -1:
            html = html[:start] + PATCH_JS.strip() + html[end+len('</script>'):]
    elif 'function isTPointForDay' not in html:
        html = html.replace('</body>', PATCH_JS + '\n</body>', 1)
    path.write_text(html, encoding='utf-8')
    print(marker)


if __name__ == '__main__':
    main()
