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
  function isGrabOnlyForDay(p){
    const k=(typeof pvKind==='function')?pvKind(p):(p&&p.is_grab&&!p.is_tp?'grab':'tp');
    return k==='grab';
  }
  function isTPointForDay(p){return !isGrabOnlyForDay(p);}
  function tripLimitActive(){
    try{return typeof maxTripDays==='function' && maxTripDays()>0;}catch(e){return false;}
  }
  function isRemoteDayPoint(p){
    if(!tripLimitActive()) return false;
    try{return typeof isTrasfertaPoint==='function' && isTrasfertaPoint(p);}catch(e){return false;}
  }
  function sameArea(a,b){
    if(!a||!b)return false;
    const ar=norm(a.region||''), br=norm(b.region||'');
    const ap=norm(a.province||''), bp=norm(b.province||'');
    const ac=norm(a.city||''), bc=norm(b.city||'');
    return (ar&&ar===br)||(ap&&ap===bp)||(ac&&ac===bc);
  }
  window.assign = assign = function(ordered,days,start){
    let di=0,last=start,mins=0,count=0,tpCount=0,grabCount=0,out=[];
    let dayRemote=false,dayAnchor=null;
    const monthValue=document.getElementById('month')?document.getElementById('month').value:'';
    const monthPrefix=monthValue?monthValue+'-':'';
    const monthDays=days.filter(d=>iso(d).startsWith(monthPrefix));
    const overflowDays=days.filter(d=>!iso(d).startsWith(monthPrefix)).slice(0,15);
    const usableDays=monthDays.concat(overflowDays);
    const list=usableDays.length?usableDays:days;
    const softDay=450;
    const hardDay=540;
    const remoteHardDay=630;
    const targetTotal=6;
    const maxTPoint=6;
    const maxTotal=7;
    const remoteMaxTotal=11;
    function resetDay(){mins=0;count=0;tpCount=0;grabCount=0;last=start;dayRemote=false;dayAnchor=null;}
    function dayAddCost(point, from){
      const rawKm=km(from||start,point);
      const travel=travelMin(rawKm);
      return {travel, total:travel+visitMin(point), km:rawKm};
    }
    function canAddSeventh(point,cost,nextTp,nextGrab){
      if(count<targetTotal) return true;
      if(count>=maxTotal) return false;
      if(nextTp>5 && nextGrab>0) return false;
      if(cost.km>18 && !(dayAnchor&&sameArea(dayAnchor,point))) return false;
      if(mins+cost.total>hardDay) return false;
      return true;
    }
    function canExtraRemoteGrab(point,cost){
      if(!dayRemote) return false;
      if(!isGrabOnlyForDay(point)) return false;
      if(count>=remoteMaxTotal) return false;
      if(mins+cost.total>remoteHardDay) return false;
      if(cost.km>35 && !(dayAnchor&&sameArea(dayAnchor,point))) return false;
      return true;
    }
    function mustNewDay(point, cost){
      const isTp=isTPointForDay(point);
      const nextTp=tpCount+(isTp?1:0);
      const nextGrab=grabCount+(isTp?0:1);
      const nextTotal=count+1;
      const remoteCandidate=isRemoteDayPoint(point) || dayRemote;
      if(count<=0) return false;
      if(canExtraRemoteGrab(point,cost)) return false;
      if(nextTp>maxTPoint) return true;
      if(remoteCandidate){
        if(nextTotal>remoteMaxTotal) return true;
        if(mins+cost.total>remoteHardDay) return true;
        if(mins+cost.total>hardDay && !isGrabOnlyForDay(point)) return true;
        return false;
      }
      if(nextTotal>maxTotal) return true;
      if(nextTp>=5 && nextGrab>2) return true;
      if(nextTp===4 && nextGrab>3) return true;
      if(!canAddSeventh(point,cost,nextTp,nextGrab)) return true;
      if(mins+cost.total>hardDay) return true;
      if(mins+cost.total>softDay && count>=targetTotal-1) return true;
      return false;
    }
    for(const p of ordered){
      let cost=dayAddCost(p,last);
      if(mustNewDay(p,cost) && di<list.length-1){
        di++;resetDay();cost=dayAddCost(p,last);
      }
      const day=list[Math.min(di,list.length-1)]||new Date();
      const st=9*60+mins+cost.travel;
      const isTp=isTPointForDay(p);
      const newLoad=mins+cost.total;
      out.push(Object.assign({},p,{date:iso(day),dateLabel:dateLabel(day),dateOnly:dateOnly(day),time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),travel_km:Math.round(cost.km),visit_min:visitMin(p),day_load:minText(newLoad)}));
      mins=newLoad;count++;if(isTp)tpCount++;else grabCount++;
      if(isRemoteDayPoint(p)){dayRemote=true;if(!dayAnchor)dayAnchor=p;}
      if(!dayAnchor)dayAnchor=p;
      last=p;
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
    if start != -1:
        end = html.find('</script>', start)
        if end != -1:
            html = html[:start] + PATCH_JS.strip() + html[end+len('</script>'):]
    elif 'function tripLimitActive' not in html:
        html = html.replace('</body>', PATCH_JS + '\n</body>', 1)
    path.write_text(html, encoding='utf-8')
    print(marker)


if __name__ == '__main__':
    main()
