from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH = r'''
<!-- Final planning dates/cap patch: selected month + max 10 workdays -->
<script id="planning-month-plus-10-final-fix">
(function(){
  if(window.__mwMonthPlus10FinalFix)return;
  window.__mwMonthPlus10FinalFix=true;

  function byId(id){return document.getElementById(id);}
  function nrm(v){try{return norm(v);}catch(e){return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}}
  function pdv(v){try{return normPdv(v);}catch(e){var m=String(v||'').match(/\d+/);return m?m[0].padStart(5,'0'):'';}}
  function localIso(d){
    var x=(d instanceof Date)?d:new Date(d);
    return x.getFullYear()+'-'+String(x.getMonth()+1).padStart(2,'0')+'-'+String(x.getDate()).padStart(2,'0');
  }
  function label(d){try{return d.toLocaleDateString('it-IT',{weekday:'short',day:'2-digit',month:'2-digit',year:'numeric'});}catch(e){return localIso(d);}}
  function onlyDate(d){try{return d.toLocaleDateString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric'});}catch(e){return localIso(d);}}
  function sameAgent(p,agent){
    var want=nrm(agent);
    var vals=[p&&p.agent_display,p&&p.agent];
    try{vals.push(visibleAgentName(p&&p.agent));}catch(e){}
    return vals.some(function(v){return nrm(v)===want;});
  }
  function removeExtraTripMode(){var el=byId('tripMode');if(el){var wrap=el.closest('div');if(wrap)wrap.remove();else el.remove();}}

  function holidaySet(y){
    var s={};
    ['01-01','01-06','04-25','05-01','06-02','08-15','11-01','12-08','12-25','12-26'].forEach(function(x){s[y+'-'+x]=true;});
    try{var e=easter(y),p=new Date(e.getFullYear(),e.getMonth(),e.getDate()+1,12);s[localIso(p)]=true;}catch(err){}
    return s;
  }
  function workdaysOfMonth(mv){
    var a=String(mv||'').split('-').map(Number),y=a[0],m=a[1],out=[];
    if(!y||!m)return out;
    var h=holidaySet(y);
    for(var d=new Date(y,m-1,1,12);d.getMonth()===m-1;d.setDate(d.getDate()+1)){
      var w=d.getDay(),k=localIso(d);
      if(w!==0&&w!==6&&!h[k])out.push(new Date(d));
    }
    return out;
  }
  function nextWorkdaysAfterMonth(mv,maxExtra){
    var a=String(mv||'').split('-').map(Number),y=a[0],m=a[1],out=[];
    if(!y||!m)return out;
    var d=new Date(y,m,1,12);
    while(out.length<maxExtra){
      var h=holidaySet(d.getFullYear()),w=d.getDay(),k=localIso(d);
      if(w!==0&&w!==6&&!h[k])out.push(new Date(d));
      d.setDate(d.getDate()+1);
    }
    return out;
  }
  function planningDays(mv){return workdaysOfMonth(mv).concat(nextWorkdaysAfterMonth(mv,10));}

  function fieldText(el){
    var vals=[el.id,el.name,el.placeholder,el.getAttribute('aria-label'),el.getAttribute('data-label')];
    try{if(el.id){var lab=document.querySelector('label[for="'+CSS.escape(el.id)+'"]');if(lab)vals.push(lab.textContent);}var p=el.closest('div,section,fieldset,label');if(p)vals.push(p.textContent);}catch(e){}
    return nrm(vals.filter(Boolean).join(' '));
  }
  function currentTransferDays(){
    var els=[].slice.call(document.querySelectorAll('input,select,textarea')).filter(function(el){var t=fieldText(el);return t.includes('trasf')||t.includes('transfer');});
    var max=0,explicitNo=false;
    els.forEach(function(el){
      var type=String(el.type||'').toLowerCase(),tag=String(el.tagName||'').toLowerCase(),raw=String(el.value||'').trim().toLowerCase(),text=nrm(raw);
      if((type==='checkbox'||type==='radio')&&!el.checked)return;
      if(text==='no'||text==='0'||text==='false'||text==='nessuna'||text==='nessuno')explicitNo=true;
      if((tag==='select'||type==='checkbox'||type==='radio')&&(text==='si'||text==='sì'||text==='yes'||text==='true'))max=Math.max(max,1);
      var m=raw.match(/-?\d+/);if(m){var n=parseInt(m[0],10);if(Number.isFinite(n)&&n>0)max=Math.max(max,n);}
    });
    return explicitNo?0:max;
  }

  function allGeo(points){
    var out=[];
    try{out=out.concat(points||[]);}catch(e){}
    try{out=out.concat(DATA&&DATA.catalog?DATA.catalog:[]);}catch(e){}
    try{out=out.concat(typeof allPoints==='function'?allPoints():[]);}catch(e){}
    var seen={};
    return out.filter(function(p){if(!p||p.lat==null||p.lng==null)return false;var key=(p.pdv||'')+'|'+(p.city||'')+'|'+(p.address||'');if(seen[key])return false;seen[key]=1;return true;});
  }
  function avg(list,raw){
    var pts=(list||[]).filter(function(p){return p&&p.lat!=null&&p.lng!=null;});
    if(!pts.length)return null;
    return {pdv:'PARTENZA',city:(pts[0].city||raw||'Partenza'),address:raw||'',lat:pts.reduce(function(s,p){return s+Number(p.lat||0);},0)/pts.length,lng:pts.reduce(function(s,p){return s+Number(p.lng||0);},0)/pts.length,is_start:true,is_tp:false,is_grab:false};
  }
  function originPoint(points,start){
    var raw=String(start||'').trim(),s=nrm(raw);
    if(!s)return points[0]||null;
    var all=allGeo(points),p=all.find(function(x){return pdv(x.pdv)&&pdv(x.pdv)===pdv(raw);});
    if(p)return Object.assign({},p,{pdv:'PARTENZA',address:raw||p.address||'',is_start:true,is_tp:false,is_grab:false});
    var exact=all.filter(function(x){return nrm((x.city||'')+' '+(x.address||''))===s;});if(exact.length)return avg(exact,raw);
    var contains=all.filter(function(x){var t=nrm((x.city||'')+' '+(x.address||''));return t&&t.includes(s);});if(contains.length)return avg(contains,raw);
    var city=all.filter(function(x){return nrm(x.city)===s;});if(city.length)return avg(city,raw);
    city=all.filter(function(x){var c=nrm(x.city);return c&&s.includes(c);});if(city.length)return avg(city,raw);
    return points[0]||null;
  }

  function agePenalty(p,mv){try{var age=prevAge(p,mv);return age<45?650:(age<90?160:0);}catch(e){return 0;}}
  function choose(left,cur,origin,mv,used){
    var best=null,bestScore=Infinity;
    left.forEach(function(p){var c=nrm(p.city||''),cluster=used&&used[c]?-12:0,score=km(cur,p)+km(p,origin)*0.45+agePenalty(p,mv)+cluster;if(score<bestScore){bestScore=score;best=p;}});
    return best;
  }
  function remove(left,p){var i=left.indexOf(p);if(i>=0)left.splice(i,1);}
  function assignPlan(points,days,origin,mv,transferDays){
    var left=(points||[]).slice(),out=[],previous=null;
    if(!origin)origin=left[0]||null;
    for(var di=0;di<days.length&&left.length;di++){
      var transfer=transferDays>0&&di<transferDays;
      var start=(transfer&&previous)?previous:origin,cur=start,mins=0,count=0,used={},maxCount=transfer?5:4;
      while(left.length&&count<maxCount){
        var p=choose(left,cur,start,mv,used);if(!p)break;
        var leg=km(cur,p),add=travelMin(leg)+visitMin(p),ret=transfer?0:travelMin(km(p,start));
        if(count>0&&mins+add+ret>540)break;
        if(count>=3&&mins+add+ret>480)break;
        remove(left,p);
        var day=days[di],startMin=9*60+mins+travelMin(leg);
        mins+=add;used[nrm(p.city||'')]=1;previous=p;cur=p;count++;
        out.push(Object.assign({},p,{date:localIso(day),dateLabel:label(day),dateOnly:onlyDate(day),time:String(Math.floor(startMin/60)).padStart(2,'0')+':'+String(startMin%60).padStart(2,'0'),travel_km:Math.round(leg),visit_min:visitMin(p),day_load:minText(mins+ret),return_km:transfer?0:Math.round(km(p,start)),planning_mode:transfer?'trasferta':'giornaliero'}));
      }
    }
    window.__mwNotPlannedCount=left.length;
    window.__mwExtraDaysUsed=Math.max(0,[...new Set(out.map(function(p){return p.date;}))].filter(function(d){return d.slice(0,7)!==String(mv);}).length);
    return out;
  }
  function showWarn(){
    var result=byId('result');if(!result)return;
    var old=byId('monthPlus10Warning');if(old)old.remove();
    var n=Number(window.__mwNotPlannedCount||0),extra=Number(window.__mwExtraDaysUsed||0);
    if(n>0||extra>0){
      result.insertAdjacentHTML('afterbegin','<section id="monthPlus10Warning" class="card" style="border-color:#f59e0b;background:#fffbeb"><div class="title" style="color:#92400e">Limite date applicato</div><div class="muted">Il planning usa il mese selezionato più massimo <b>10 giorni lavorativi extra</b>. Giorni extra usati: <b>'+extra+'</b>. PV rimasti fuori: <b>'+n.toLocaleString('it-IT')+'</b>.</div></section>');
    }
  }

  window.generatePlanning = generatePlanning = function(){
    removeExtraTripMode();
    var agent=byId('agent')&&byId('agent').value,mv=byId('month')&&byId('month').value,grabVal=String(byId('grab')&&byId('grab').value||'').toLowerCase();
    var includeGrab=['yes','si','sì','all','only_new'].includes(grabVal);
    if(!agent||!mv){alert('Scegli agente e mese');return;}
    var pts=[];try{pts=allPoints().filter(function(p){return sameAgent(p,agent)&&(includeGrab||!isGrab(p))&&p.lat!=null&&p.lng!=null;});}catch(e){pts=[];}
    if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}
    var days=planningDays(mv);if(!days.length){alert('Nessun giorno lavorativo trovato');return;}
    PLAN=assignPlan(pts,days,originPoint(pts,byId('start')&&byId('start').value),mv,currentTransferDays());
    renderTable(days.length);
    showWarn();
  };

  var oldRender=window.renderTable;
  if(typeof oldRender==='function'){
    window.renderTable = renderTable = function(workdayCount){oldRender(workdayCount);removeExtraTripMode();showWarn();};
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',removeExtraTripMode);else removeExtraTripMode();
  window.addEventListener('pageshow',removeExtraTripMode);
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, month+10 saltato")
        return
    html = path.read_text(encoding="utf-8")
    if 'planning-month-plus-10-final-fix' not in html:
        html = html.replace('</body>', PATCH + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Patch planning mese + 10 giorni applicata")


if __name__ == "__main__":
    main()
