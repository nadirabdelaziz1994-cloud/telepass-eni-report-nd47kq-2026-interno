from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PLANNING_JS = r'''
<!-- Roundtrip planning logic + desktop drag fix -->
<script id="planning-roundtrip-desktop-fix">
(function(){
  if(window.__mwRoundTripPlanningFix)return;
  window.__mwRoundTripPlanningFix=true;
  function byId(id){return document.getElementById(id);}
  function safeNorm(v){try{return norm(v);}catch(e){return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}}
  function safeIso(d){try{return iso(d);}catch(e){return d.toISOString().slice(0,10);}}
  function sameAgent(p,agent){
    if(!agent)return false;
    var vals=[p&&p.agent_display,p&&p.agent];
    try{vals.push(visibleAgentName(p&&p.agent));}catch(e){}
    var want=safeNorm(agent);
    return vals.some(function(v){return safeNorm(v)===want;});
  }
  function ensureTripMode(){
    var start=byId('start');
    if(!start||byId('tripMode'))return;
    var wrap=start.closest('div')||start.parentElement;
    var html='<div><label>Tipo giro</label><select id="tripMode"><option value="home" selected>Giornaliero: parto e torno dal punto di partenza</option><option value="away">Trasferta: giro continuativo tra giorni</option></select><div class="muted" style="font-size:11px;margin-top:4px">Usa “trasferta” solo quando dormi fuori: altrimenti ogni giornata riparte dal punto scritto sopra.</div></div>';
    if(wrap)wrap.insertAdjacentHTML('afterend',html);
  }
  function allKnownGeo(points){
    var out=[];
    try{out=out.concat(points||[]);}catch(e){}
    try{out=out.concat(DATA&&DATA.catalog?DATA.catalog:[]);}catch(e){}
    try{out=out.concat(typeof allPoints==='function'?allPoints():[]);}catch(e){}
    var seen={};
    return out.filter(function(p){
      if(!p||p.lat==null||p.lng==null)return false;
      var key=(p.pdv||'')+'|'+(p.city||'')+'|'+(p.address||'');
      if(seen[key])return false;seen[key]=1;return true;
    });
  }
  function avgPoint(list,label,raw){
    var pts=(list||[]).filter(function(p){return p&&p.lat!=null&&p.lng!=null;});
    if(!pts.length)return null;
    var lat=pts.reduce(function(s,p){return s+Number(p.lat||0);},0)/pts.length;
    var lng=pts.reduce(function(s,p){return s+Number(p.lng||0);},0)/pts.length;
    return {pdv:'PARTENZA',city:label||raw||'Partenza',address:raw||'',lat:lat,lng:lng,is_start:true,is_tp:false,is_grab:false,agent_display:''};
  }
  window.startPoint = startPoint = function(points,start){
    var raw=String(start||'').trim();
    var s=safeNorm(raw);
    if(!s)return (points&&points[0])||null;
    var all=allKnownGeo(points);
    var byPdv=all.find(function(p){return safeNorm(p.pdv)===s||normPdv(p.pdv)===normPdv(raw);});
    if(byPdv)return Object.assign({is_start:true},byPdv,{pdv:'PARTENZA',address:raw||byPdv.address||'',is_tp:false,is_grab:false});
    var exactFull=all.filter(function(p){return safeNorm((p.city||'')+' '+(p.address||''))===s;});
    if(exactFull.length)return avgPoint(exactFull,exactFull[0].city||raw,raw);
    var containsFull=all.filter(function(p){var text=safeNorm((p.city||'')+' '+(p.address||''));return text&&text.includes(s);});
    if(containsFull.length)return avgPoint(containsFull,containsFull[0].city||raw,raw);
    var exactCity=all.filter(function(p){return safeNorm(p.city)===s;});
    if(exactCity.length)return avgPoint(exactCity,exactCity[0].city||raw,raw);
    var cityInStart=all.filter(function(p){var c=safeNorm(p.city);return c&&s.includes(c);});
    if(cityInStart.length)return avgPoint(cityInStart,cityInStart[0].city||raw,raw);
    return (points&&points[0])||null;
  };
  function validWorkdays(mv){
    var ds=[];
    try{ds=(workdays(mv)||[]).map(function(d){return new Date(d);});}catch(e){ds=[];}
    if(!ds.length){
      var a=String(mv||'').split('-').map(Number),y=a[0],m=a[1];
      if(y&&m){for(var d=new Date(y,m-1,1);d.getMonth()===m-1;d.setDate(d.getDate()+1)){var w=d.getDay();if(w!==0&&w!==6)ds.push(new Date(d));}}
    }
    var seen={};
    return ds.filter(function(d){
      var k=safeIso(d),w=d.getDay();
      if(w===0||w===6||seen[k])return false;seen[k]=1;return true;
    }).sort(function(a,b){return a-b;});
  }
  function prevPenalty(p,mv){
    try{var age=prevAge(p,mv);return age<45?650:(age<90?160:0);}catch(e){return 0;}
  }
  function chooseBest(left,cur,home,mv,usedCities){
    var best=null,bestScore=Infinity;
    left.forEach(function(p){
      var city=safeNorm(p.city||'');
      var sameCity=usedCities&&usedCities[city]? -8:0;
      var score=km(cur,p)+(km(p,home)*0.35)+prevPenalty(p,mv)+sameCity;
      if(score<bestScore){bestScore=score;best=p;}
    });
    return best;
  }
  function removePoint(left,p){var i=left.indexOf(p);if(i>=0)left.splice(i,1);}
  function assignRoundTrips(points,days,home,mv,awayMode){
    var left=(points||[]).slice();
    var out=[],previousLast=null;
    if(!home)home=left[0]||null;
    for(var di=0;di<days.length&&left.length;di++){
      var remainingDays=days.length-di;
      var target=Math.max(1,Math.ceil(left.length/remainingDays));
      var maxCount=Math.min(9,Math.max(target,Math.min(target+2,6)));
      var origin=(awayMode&&previousLast)?previousLast:home;
      var cur=origin,mins=0,count=0,usedCities={};
      while(left.length&&count<maxCount){
        var p=chooseBest(left,cur,origin,mv,usedCities);
        if(!p)break;
        var legKm=km(cur,p),add=travelMin(legKm)+visitMin(p),returnHome=travelMin(km(p,origin));
        if(count>0&&mins+add+returnHome>540)break;
        if(count>=target&&mins+add+returnHome>420)break;
        removePoint(left,p);
        var day=days[di]||new Date();
        var startMin=9*60+mins+travelMin(legKm);
        mins+=add;
        usedCities[safeNorm(p.city||'')]=1;
        previousLast=p;cur=p;count++;
        out.push(Object.assign({},p,{
          date:safeIso(day),
          dateLabel:dateLabel(day),
          dateOnly:dateOnly(day),
          time:String(Math.floor(startMin/60)).padStart(2,'0')+':'+String(startMin%60).padStart(2,'0'),
          travel_km:Math.round(legKm),
          visit_min:visitMin(p),
          day_load:minText(mins+returnHome),
          return_km:Math.round(km(p,origin)),
          planning_mode:awayMode?'trasferta':'giornaliero'
        }));
      }
    }
    if(left.length){
      var extra=validWorkdays(mv);var lastDay=extra[extra.length-1]||new Date();
      var origin=awayMode&&previousLast?previousLast:home,cur=origin,mins=0;
      left.forEach(function(p){
        var legKm=km(cur,p),add=travelMin(legKm)+visitMin(p),returnHome=travelMin(km(p,origin));
        var startMin=9*60+mins+travelMin(legKm);mins+=add;cur=p;previousLast=p;
        out.push(Object.assign({},p,{date:safeIso(lastDay),dateLabel:dateLabel(lastDay),dateOnly:dateOnly(lastDay),time:String(Math.floor(startMin/60)).padStart(2,'0')+':'+String(startMin%60).padStart(2,'0'),travel_km:Math.round(legKm),visit_min:visitMin(p),day_load:minText(mins+returnHome),return_km:Math.round(km(p,origin)),planning_mode:awayMode?'trasferta':'giornaliero'}));
      });
    }
    return out;
  }
  window.generatePlanning = generatePlanning = function(){
    ensureTripMode();
    var agent=byId('agent')&&byId('agent').value;
    var mv=byId('month')&&byId('month').value;
    var inc=byId('grab')&&['yes','all','only_new'].includes(byId('grab').value);
    var away=byId('tripMode')&&byId('tripMode').value==='away';
    if(!agent||!mv){alert('Scegli agente e mese');return;}
    var pts=[];
    try{pts=allPoints().filter(function(p){return sameAgent(p,agent)&&(inc||!isGrab(p))&&p.lat!=null&&p.lng!=null;});}catch(e){pts=[];}
    if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}
    var days=validWorkdays(mv);
    if(!days.length){alert('Nessun giorno lavorativo trovato per il mese scelto');return;}
    var home=startPoint(pts,byId('start')&&byId('start').value);
    PLAN=assignRoundTrips(pts,days,home,mv,away);
    renderTable(days.length);
  };
  var oldRenderTable=window.renderTable;
  if(typeof oldRenderTable==='function'){
    window.renderTable = renderTable = function(workdayCount){oldRenderTable(workdayCount);ensureTripMode();};
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',ensureTripMode);else ensureTripMode();
  window.addEventListener('pageshow',ensureTripMode);
})();
</script>
'''

EDIT_JS = r'''
<!-- Desktop drag/drop for individual rows and full day blocks -->
<script id="planning-edit-desktop-drag-fix">
(function(){
  if(window.__mwPlanningEditDesktopFix)return;
  window.__mwPlanningEditDesktopFix=true;
  let rowDragIndex=null,dayDragIndex=null;
  function fmtDateShort(v){try{return new Date(v+'T00:00:00').toLocaleDateString('it-IT',{weekday:'short',day:'2-digit',month:'2-digit'});}catch(e){return v||'';}}
  function pClass(p){if(p.is_grab&&p.is_tp)return 'dual';if(p.is_grab)return 'grab';return 'tp';}
  function pTipo(p){if(p.is_grab&&p.is_tp)return 'Doppio';if(p.is_grab)return 'Grab';return 'TPoint';}
  function groups(){const out=[];PLAN.forEach((p,idx)=>{const key=p.date||'';let g=out.find(x=>x.date===key);if(!g){g={date:key,items:[]};out.push(g);}g.items.push({p,idx});});return out;}
  function orderedDateSlots(gs){return gs.map(g=>g.date).filter(Boolean).sort();}
  function setDateObj(p,d){p.date=d;if(typeof dateLabel==='function')p.dateLabel=dateLabel(d);if(typeof dateOnly==='function')p.dateOnly=dateOnly(d);}
  function applyBlockDates(order,slots){order.forEach((g,i)=>{const nd=slots[i]||g.date;g.items.forEach(it=>setDateObj(it.p,nd));});}
  function moveRowToIndex(from,to,targetDate){
    if(from==null||to==null||from<0||to<0||from>=PLAN.length||to>=PLAN.length)return;
    const item=PLAN[from];
    if(targetDate)setDateObj(item,targetDate);
    PLAN.splice(from,1);
    if(from<to)to--;
    PLAN.splice(to,0,item);
    render();
  }
  function appendRowToDay(from,dayIndex){
    const gs=groups(),g=gs[dayIndex];
    if(!g||from<0||from>=PLAN.length)return;
    const item=PLAN[from];setDateObj(item,g.date);
    PLAN.splice(from,1);
    let insert=PLAN.length;
    for(let i=PLAN.length-1;i>=0;i--){if(PLAN[i].date===g.date){insert=i+1;break;}}
    PLAN.splice(insert,0,item);render();
  }
  window.changeBlockDate=function(dayIndex,newDate){const g=groups()[dayIndex];if(!g||!newDate)return;g.items.forEach(it=>setDateObj(it.p,newDate));render();};
  window.moveDayBlock=function(from,to){const gs=groups();if(from===to||from<0||to<0||from>=gs.length||to>=gs.length)return;const slots=orderedDateSlots(gs);const g=gs.splice(from,1)[0];gs.splice(to,0,g);applyBlockDates(gs,slots);PLAN=[];gs.forEach(x=>x.items.forEach(it=>PLAN.push(it.p)));render();};
  window.dayBlockUp=function(i){moveDayBlock(i,i-1);};
  window.dayBlockDown=function(i){moveDayBlock(i,i+1);};
  window.rowDragStart=function(e,i){rowDragIndex=i;dayDragIndex=null;e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain','row:'+i);};
  window.rowDragOver=function(e){e.preventDefault();e.dataTransfer.dropEffect='move';};
  window.rowDrop=function(e,i){e.preventDefault();const raw=e.dataTransfer.getData('text/plain')||'';const from=raw.startsWith('row:')?Number(raw.slice(4)):rowDragIndex;if(Number.isFinite(from)){const target=PLAN[i];moveRowToIndex(from,i,target&&target.date);}rowDragIndex=null;};
  window.dayDragStartDesktop=function(e,i){dayDragIndex=i;rowDragIndex=null;e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain','day:'+i);};
  window.dayDragOverDesktop=function(e){e.preventDefault();e.dataTransfer.dropEffect='move';};
  window.dayDropDesktop=function(e,i){e.preventDefault();const raw=e.dataTransfer.getData('text/plain')||'';if(raw.startsWith('day:')){moveDayBlock(Number(raw.slice(4)),i);}else if(raw.startsWith('row:')){appendRowToDay(Number(raw.slice(4)),i);}dayDragIndex=null;rowDragIndex=null;};
  window.dayTouchStart=function(e,i){dayDragIndex=i;const b=e.currentTarget.closest('.day-block');if(b)b.classList.add('day-moving');e.preventDefault();};
  window.dayTouchMove=function(e){if(dayDragIndex==null)return;const t=e.touches&&e.touches[0]?e.touches[0]:e;const blocks=[...document.querySelectorAll('.day-block')];let best=null,dist=Infinity;blocks.forEach(b=>{const r=b.getBoundingClientRect(),c=r.top+r.height/2,d=Math.abs(t.clientY-c);if(d<dist){dist=d;best=Number(b.dataset.day);}});if(best!=null&&best!==dayDragIndex){moveDayBlock(dayDragIndex,best);dayDragIndex=best;}e.preventDefault();};
  window.dayTouchEnd=function(e){dayDragIndex=null;document.querySelectorAll('.day-moving').forEach(x=>x.classList.remove('day-moving'));if(e)e.preventDefault();};
  window.render=function(){
    const box=document.getElementById('list');
    if(!PLAN.length){box.innerHTML='<div class="card muted">Nessun planning trovato. Torna alla pagina planning e crealo prima.</div>';return;}
    const gs=groups();let n=0;
    box.innerHTML='<div class="head-row"><div>↕</div><div>#</div><div>Data</div><div>PV</div><div>Comune</div><div>Via</div><div>Tipo</div><div>X</div></div>'+gs.map((g,di)=>{
      const head='<div class="day-block" data-day="'+di+'" draggable="true" ondragstart="dayDragStartDesktop(event,'+di+')" ondragover="dayDragOverDesktop(event)" ondrop="dayDropDesktop(event,'+di+')"><div class="day-controls"><button class="day-drag" draggable="true" ondragstart="dayDragStartDesktop(event,'+di+')" ontouchstart="dayTouchStart(event,'+di+')" ontouchmove="dayTouchMove(event)" ontouchend="dayTouchEnd(event)">↕</button><button class="day-mini" onclick="dayBlockUp('+di+')">↑</button><button class="day-mini" onclick="dayBlockDown('+di+')">↓</button></div><div class="day-title">Giorno '+(di+1)+' · '+esc(fmtDateShort(g.date))+' <span>('+g.items.length+' PV)</span></div><input class="day-date" type="date" value="'+esc(g.date||'')+'" onchange="changeBlockDate('+di+',this.value)"></div>';
      const rows=g.items.map(it=>{const p=it.p,i=it.idx;n++;return '<div class="edit-row '+pClass(p)+'" data-i="'+i+'" draggable="true" ondragstart="rowDragStart(event,'+i+')" ondragover="rowDragOver(event)" ondrop="rowDrop(event,'+i+')"><button class="drag" draggable="true" ondragstart="rowDragStart(event,'+i+')" ontouchstart="touchStart(event,'+i+')" ontouchmove="touchMove(event)" ontouchend="touchEnd(event)">↕</button><div class="idx">'+n+'</div><input class="date" type="date" value="'+esc(p.date||'')+'" onchange="changeDate('+i+',this.value)"><div class="col"><b>'+esc(p.pdv)+'</b></div><div class="col">'+esc(p.city||'')+'</div><div class="col">'+esc(p.address||'')+'</div><div class="col">'+esc(pTipo(p))+'</div><button class="btn bad x" onclick="del('+i+')">×</button></div>';}).join('');
      return head+rows;
    }).join('');
  };
  try{render();}catch(e){}
})();
</script>
'''


def inject_once(html: str, marker: str, snippet: str) -> str:
    if marker in html:
        return html
    if '</body>' in html:
        return html.replace('</body>', snippet + '\n</body>', 1)
    return html + '\n' + snippet


def main():
    planning = DOCS_DIR / 'planning.html'
    if planning.exists():
        html = planning.read_text(encoding='utf-8')
        html = inject_once(html, 'planning-roundtrip-desktop-fix', PLANNING_JS)
        planning.write_text(html, encoding='utf-8')
        print('Planning roundtrip/dates patch applicata')
    else:
        print('planning.html non trovato, roundtrip patch saltata')

    edit = DOCS_DIR / 'planning-edit.html'
    if edit.exists():
        html = edit.read_text(encoding='utf-8')
        html = inject_once(html, 'planning-edit-desktop-drag-fix', EDIT_JS)
        edit.write_text(html, encoding='utf-8')
        print('Planning editor desktop drag patch applicata')
    else:
        print('planning-edit.html non trovato, desktop drag patch saltata')


if __name__ == '__main__':
    main()
