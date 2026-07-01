from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

SNIPPET = r'''
<!-- Smart planning engine + schematic map -->
<script id="planning-smart-engine-map-fix">
(function(){
  if(window.__mwSmartPlanningEngineV1)return;
  window.__mwSmartPlanningEngineV1=true;

  var SETTINGS={
    dayMinutes:480,
    tpCalcMin:30,
    grabCalcMin:10,
    normalMaxKm:300,
    transferMaxKm:350,
    normalSafetyMaxPv:8,
    transferSafetyMaxPv:10
  };
  var selectedMapDate='';

  function byId(id){return document.getElementById(id);}
  function isAdmin(){try{return new URLSearchParams(location.search).get('admin')==='1';}catch(e){return false;}}
  function nrm(v){try{return norm(v);}catch(e){return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}}
  function esc2(v){try{return esc(v);}catch(e){return String(v==null?'':v).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}}
  function sameCity(a,b){return nrm(a&&a.city) && nrm(a&&a.city)===nrm(b&&b.city);}
  function sameProvince(a,b){return nrm(a&&a.province) && nrm(a&&a.province)===nrm(b&&b.province);}
  function dayVisitMin(p){return isGrab(p)?SETTINGS.grabCalcMin:SETTINGS.tpCalcMin;}
  function safetyMaxPv(tr){return tr?SETTINGS.transferSafetyMaxPv:SETTINGS.normalSafetyMaxPv;}
  function maxKm(tr){return tr?SETTINGS.transferMaxKm:SETTINGS.normalMaxKm;}
  function removeHeaderManagementLink(){var old=byId('adminPlanningLink');if(old)old.remove();}

  function ensureManagementButton(formSection){
    removeHeaderManagementLink();
    if(!formSection || byId('managePvWrap'))return;
    var title=formSection.querySelector('.title');
    var wrap=document.createElement('div');
    wrap.id='managePvWrap';
    wrap.className='actions';
    wrap.style.margin='0 0 12px 0';
    wrap.innerHTML='<a class="btn light" id="managePvButton" style="text-decoration:none;display:inline-flex;align-items:center" href="./planning.html?admin=1&v=admin">Aggiungi/Rimuovi un punto vendita</a>';
    if(title)formSection.insertBefore(wrap,title);else formSection.insertBefore(wrap,formSection.firstChild);
  }

  function ensureBackButton(adminSection){
    if(!adminSection || byId('backPlanningWrap'))return;
    var wrap=document.createElement('div');
    wrap.id='backPlanningWrap';
    wrap.className='actions';
    wrap.style.margin='0 0 12px 0';
    wrap.innerHTML='<a class="btn light" style="text-decoration:none;display:inline-flex;align-items:center" href="./planning.html?v=from-admin">← Torna al planning</a>';
    adminSection.insertBefore(wrap,adminSection.firstChild);
  }

  function cleanUiMode(){
    removeHeaderManagementLink();
    var adminSection=byId('apiStatus')?byId('apiStatus').closest('section'):null;
    var sourceSection=byId('sourceBox')?byId('sourceBox').closest('section'):null;
    var formSection=byId('agent')?byId('agent').closest('section'):null;
    var metrics=byId('metricBox'), result=byId('result'), map=byId('planningMapCard');

    if(isAdmin()){
      if(sourceSection)sourceSection.style.display='none';
      if(formSection)formSection.style.display='none';
      if(metrics)metrics.style.display='none';
      if(result)result.style.display='none';
      if(map)map.style.display='none';
      if(adminSection){adminSection.style.display='block';ensureBackButton(adminSection);}
      document.title='Gestione modifiche planning - Telepass';
      var h=document.querySelector('header h1'); if(h)h.textContent='Gestione modifiche planning';
      return;
    }
    if(sourceSection)sourceSection.style.display='none';
    if(adminSection)adminSection.style.display='none';
    if(formSection){formSection.style.display='block';ensureManagementButton(formSection);}
    if(metrics)metrics.style.display='grid';
    if(result)result.style.display='block';
    if(map)map.style.display='block';

    document.querySelectorAll('button').forEach(function(btn){
      var txt=(btn.textContent||'').toLowerCase(), on=String(btn.getAttribute('onclick')||'').toLowerCase();
      if(txt.includes('ricarica modifiche')||on.includes('loadremote(false)'))btn.style.display='none';
    });
    document.querySelectorAll('.card .muted').forEach(function(el){
      var txt=(el.textContent||'').toLowerCase();
      if(txt.includes('pagina separata di sicurezza')||txt.includes('mese selezionato più massimo 10 giorni'))el.style.display='none';
    });
  }

  function scorePoint(p,cur,origin,seed,mv,used){
    var score=km(cur,p)+km(p,origin)*0.28+prevPenalty(p,mv);
    var city=nrm(p.city||''), prov=nrm(p.province||'');
    if(used.city[city])score-=28;
    if(used.province[prov])score-=14;
    if(seed){
      if(sameCity(seed,p))score-=20;
      if(sameProvince(seed,p))score-=10;
      score+=Math.max(0,km(seed,p)-18)*1.3;
    }
    return score;
  }

  function canAddPoint(p,state,tr,origin){
    var cur=state.cur, seed=state.seed, count=state.count;
    var leg=km(cur,p), back=tr?0:km(p,origin);
    var newKm=state.driveKm+leg+back;
    var newMin=state.mins+travelMin(leg)+dayVisitMin(p)+(tr?0:travelMin(back));

    if(count>0 && newMin>SETTINGS.dayMinutes)return false;
    if(count>0 && newKm>maxKm(tr))return false;
    if(count>=safetyMaxPv(tr))return false;

    if(!tr && seed && count>=2){
      var seedDist=km(seed,p);
      if(seedDist>45 && !sameProvince(seed,p))return false;
      if(count>=4 && seedDist>32 && !sameCity(seed,p))return false;
      if(count>=5 && leg>22 && !sameCity(cur,p))return false;
      if(count>=6 && leg>12 && !sameCity(cur,p))return false;
    }
    if(tr && seed && count>=5){
      if(km(seed,p)>65 && !sameProvince(seed,p))return false;
    }
    return true;
  }

  function pickCandidate(left,state,tr,origin,mv){
    var best=null, bestScore=Infinity;
    left.forEach(function(p){
      if(!canAddPoint(p,state,tr,origin))return;
      var s=scorePoint(p,state.cur,origin,state.seed,mv,state.used);
      if(s<bestScore){bestScore=s;best=p;}
    });
    return best;
  }

  function chooseSeed(left,origin,mv){
    var best=null,bestScore=Infinity;
    left.forEach(function(p){
      var s=km(origin,p)+prevPenalty(p,mv)*0.35;
      if(s<bestScore){bestScore=s;best=p;}
    });
    return best;
  }

  function buildSmartDay(left,day,origin,mv,tr,previousLast){
    var start=(tr&&previousLast)?previousLast:origin;
    var seed=chooseSeed(left,start,mv);
    var state={cur:start,seed:seed,mins:0,driveKm:0,count:0,used:{city:{},province:{}}};
    var items=[];

    while(left.length){
      var p=pickCandidate(left,state,tr,start,mv);
      if(!p)break;
      var leg=km(state.cur,p), visitRef=visitMin(p), visitCalc=dayVisitMin(p);
      left.splice(left.indexOf(p),1);
      var st=9*60+state.mins+travelMin(leg);
      state.mins+=travelMin(leg)+visitCalc;
      state.driveKm+=leg;
      state.cur=p;
      state.count++;
      state.used.city[nrm(p.city||'')]=1;
      state.used.province[nrm(p.province||'')]=1;
      if(!state.seed)state.seed=p;
      var backKm=tr?0:km(p,start);
      items.push(Object.assign({},p,{
        date:localIso(day),dateLabel:dateLabel(day),dateOnly:dateOnly(day),
        time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),
        travel_km:Math.round(leg),visit_min:visitRef,
        day_load:minText(state.mins+(tr?0:travelMin(backKm))),
        return_km:tr?0:Math.round(backKm),
        planning_mode:tr?'trasferta':'giornaliero',
        calc_visit_min:visitCalc
      }));
    }
    return {items:items,last:state.cur,mins:state.mins,km:state.driveKm};
  }

  window.assign = assign = function(ordered,days,start,mv){
    var left=(ordered||[]).slice(),out=[],prev=null,transferCount=0;
    try{transferCount=transferDays();}catch(e){transferCount=0;}
    for(var di=0;di<days.length && left.length;di++){
      var tr=transferCount>0 && di<transferCount;
      var res=buildSmartDay(left,days[di],start,mv,tr,prev);
      if(!res.items.length){
        var forced=left.shift();
        if(!forced)break;
        res.items=[Object.assign({},forced,{date:localIso(days[di]),dateLabel:dateLabel(days[di]),dateOnly:dateOnly(days[di]),time:'09:00',travel_km:Math.round(km(start,forced)),visit_min:visitMin(forced),day_load:minText(tr?dayVisitMin(forced):dayVisitMin(forced)+travelMin(km(forced,start))),return_km:tr?0:Math.round(km(forced,start)),planning_mode:tr?'trasferta':'giornaliero',calc_visit_min:dayVisitMin(forced)})];
        res.last=forced;
      }
      res.items.forEach(function(x){out.push(x);});
      prev=res.last;
    }
    window.__leftOut=left.length;
    window.__smartPlanningSettings=SETTINGS;
    return out;
  };

  window.generatePlanning = generatePlanning = function(){
    var agent=byId('agent')&&byId('agent').value;
    var mv=byId('month')&&byId('month').value;
    var grab=byId('grab')&&byId('grab').value;
    var inc=['yes','si','sì','all','only_new'].includes(String(grab||'').toLowerCase());
    if(!agent||!mv){alert('Scegli agente e mese');return;}
    var pts=allPoints().filter(function(p){return p.agent_display===agent && (inc||!isGrab(p)) && p.lat!=null && p.lng!=null;});
    if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}
    var days=workdays(mv);
    var start=startPoint(pts,byId('start')&&byId('start').value);
    PLAN=assign(order(pts,start,mv),days,start,mv);
    selectedMapDate='';
    renderTable(days.length);
    renderPlanningMap();
    cleanUiMode();
  };

  function uniqueDates(){return [...new Set((PLAN||[]).map(function(p){return p.date;}))].filter(Boolean).sort();}
  function mapsUrl(items){
    if(!items.length)return '#';
    var normal=items[0].planning_mode!=='trasferta';
    var startVal=byId('start')&&byId('start').value;
    var first=items[0], last=items[items.length-1];
    var origin=startVal?encodeURIComponent(startVal):encodeURIComponent(first.lat+','+first.lng);
    var destination=normal?(startVal?encodeURIComponent(startVal):encodeURIComponent(first.lat+','+first.lng)):encodeURIComponent(last.lat+','+last.lng);
    var wp=items.map(function(p){return p.lat+','+p.lng;}).join('|');
    return 'https://www.google.com/maps/dir/?api=1&travelmode=driving&origin='+origin+'&destination='+destination+'&waypoints='+encodeURIComponent(wp);
  }

  function renderPlanningMap(date){
    if(!(PLAN||[]).length)return;
    var result=byId('result'); if(!result)return;
    var dates=uniqueDates();
    selectedMapDate=date||selectedMapDate||dates[0];
    var items=PLAN.filter(function(p){return p.date===selectedMapDate && p.lat!=null && p.lng!=null;});
    var card=byId('planningMapCard');
    if(!card){card=document.createElement('section');card.id='planningMapCard';card.className='card';result.parentNode.insertBefore(card,result.nextSibling);}
    var options=dates.map(function(d){var n=PLAN.filter(function(p){return p.date===d;}).length;return '<option value="'+esc2(d)+'" '+(d===selectedMapDate?'selected':'')+'>'+esc2(d)+' · '+n+' PV</option>';}).join('');
    var minLat=Math.min.apply(null,items.map(function(p){return Number(p.lat);}));
    var maxLat=Math.max.apply(null,items.map(function(p){return Number(p.lat);}));
    var minLng=Math.min.apply(null,items.map(function(p){return Number(p.lng);}));
    var maxLng=Math.max.apply(null,items.map(function(p){return Number(p.lng);}));
    if(!isFinite(minLat)||!isFinite(maxLat)||minLat===maxLat){minLat-=0.01;maxLat+=0.01;}
    if(!isFinite(minLng)||!isFinite(maxLng)||minLng===maxLng){minLng-=0.01;maxLng+=0.01;}
    function x(p){return 40+((Number(p.lng)-minLng)/(maxLng-minLng))*520;}
    function y(p){return 320-((Number(p.lat)-minLat)/(maxLat-minLat))*260;}
    var path=items.map(function(p,i){return (i?'L':'M')+x(p).toFixed(1)+' '+y(p).toFixed(1);}).join(' ');
    var dots=items.map(function(p,i){return '<g><circle cx="'+x(p).toFixed(1)+'" cy="'+y(p).toFixed(1)+'" r="13" fill="'+(isGrab(p)?'#7c3aed':'#0d6efd')+'"></circle><text x="'+x(p).toFixed(1)+'" y="'+(y(p)+4).toFixed(1)+'" text-anchor="middle" fill="white" font-size="11" font-weight="700">'+(i+1)+'</text><title>'+esc2((i+1)+'. '+p.pdv+' '+(p.city||''))+'</title></g>';}).join('');
    var list=items.map(function(p,i){return '<div class="small"><b>'+(i+1)+'. '+esc2(p.pdv)+'</b> · '+esc2(p.city||'')+' · '+esc2(p.address||'')+'</div>';}).join('');
    card.innerHTML='<div style="display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap"><div class="title" style="margin:0">Controllo giro giornata</div><div class="actions" style="margin:0"><select id="planningMapDate" onchange="window.renderPlanningMap(this.value)">'+options+'</select><a class="btn light" target="_blank" rel="noopener" href="'+mapsUrl(items)+'">Apri giro su Google Maps</a></div></div><div class="muted" style="margin-top:6px">Mappa schematica: serve a controllare subito se il giro è compatto. La mappa Google reale si collega con la chiave API.</div><div style="display:grid;grid-template-columns:minmax(280px,620px) 1fr;gap:12px;margin-top:10px;align-items:start"><svg viewBox="0 0 600 360" style="width:100%;background:#f8fbff;border:1px solid var(--line);border-radius:14px"><path d="'+path+'" fill="none" stroke="#0f2746" stroke-width="3" stroke-linejoin="round" stroke-linecap="round" opacity="0.55"></path>'+dots+'</svg><div>'+list+'</div></div>';
  }
  window.renderPlanningMap=renderPlanningMap;

  var oldRenderAll=window.renderAll;
  if(typeof oldRenderAll==='function')window.renderAll=function(){oldRenderAll.apply(this,arguments);cleanUiMode();};
  var oldRenderTable=window.renderTable;
  if(typeof oldRenderTable==='function')window.renderTable=function(){oldRenderTable.apply(this,arguments);renderPlanningMap();cleanUiMode();};
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',cleanUiMode);else cleanUiMode();
  window.addEventListener('pageshow',cleanUiMode);
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, smart engine/map patch saltata")
        return
    html = path.read_text(encoding="utf-8")
    markers = [
        '<!-- Planning admin link + higher daily capacity -->',
        '<!-- Smart planning engine + schematic map -->',
    ]
    for marker in markers:
        start = html.find(marker)
        while start != -1:
            end = html.find('</script>', start)
            if end == -1:
                break
            html = html[:start] + html[end + len('</script>'):]
            start = html.find(marker)
    html = html.replace('</body>', SNIPPET + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Smart planning engine e mappa schematica applicati")


if __name__ == "__main__":
    main()
