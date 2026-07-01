from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

SNIPPET = r'''
<!-- Planning admin link + higher daily capacity -->
<script id="planning-admin-link-capacity-fix">
(function(){
  if(window.__mwAdminLinkCapacityFixV2)return;
  window.__mwAdminLinkCapacityFixV2=true;

  function byId(id){return document.getElementById(id);}
  function isAdmin(){
    try{return new URLSearchParams(location.search).get('admin') === '1';}catch(e){return false;}
  }
  function nrm(v){
    try{return norm(v);}catch(e){return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}
  }

  function removeHeaderManagementLink(){
    var old=byId('adminPlanningLink');
    if(old) old.remove();
  }

  function ensureManagementButton(formSection){
    removeHeaderManagementLink();
    if(!formSection || byId('managePvWrap'))return;
    var title=formSection.querySelector('.title');
    var wrap=document.createElement('div');
    wrap.id='managePvWrap';
    wrap.className='actions';
    wrap.style.margin='0 0 12px 0';
    wrap.innerHTML='<a class="btn light" id="managePvButton" style="text-decoration:none;display:inline-flex;align-items:center" href="./planning.html?admin=1&v=admin">Aggiungi/Rimuovi un punto vendita</a>';
    if(title) formSection.insertBefore(wrap,title); else formSection.insertBefore(wrap,formSection.firstChild);
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
    var adminSection=byId('apiStatus') ? byId('apiStatus').closest('section') : null;
    var sourceSection=byId('sourceBox') ? byId('sourceBox').closest('section') : null;
    var formSection=byId('agent') ? byId('agent').closest('section') : null;
    var metrics=byId('metricBox');
    var result=byId('result');

    if(isAdmin()){
      if(sourceSection) sourceSection.style.display='none';
      if(formSection) formSection.style.display='none';
      if(metrics) metrics.style.display='none';
      if(result) result.style.display='none';
      if(adminSection){
        adminSection.style.display='block';
        ensureBackButton(adminSection);
      }
      document.title='Gestione modifiche planning - Telepass';
      var h=document.querySelector('header h1'); if(h) h.textContent='Gestione modifiche planning';
      return;
    }

    if(sourceSection) sourceSection.style.display='none';
    if(adminSection) adminSection.style.display='none';
    if(formSection){
      formSection.style.display='block';
      ensureManagementButton(formSection);
    }
    if(metrics) metrics.style.display='grid';
    if(result) result.style.display='block';

    document.querySelectorAll('button').forEach(function(btn){
      var txt=(btn.textContent||'').toLowerCase();
      var on=String(btn.getAttribute('onclick')||'').toLowerCase();
      if(txt.includes('ricarica modifiche') || on.includes('loadremote(false)')) btn.style.display='none';
    });
    document.querySelectorAll('.card .muted').forEach(function(el){
      var txt=(el.textContent||'').toLowerCase();
      if(txt.includes('pagina separata di sicurezza') || txt.includes('mese selezionato più massimo 10 giorni')) el.style.display='none';
    });
  }

  function chooseMore(left,cur,origin,mv,used){
    var best=null,bestScore=Infinity;
    left.forEach(function(p){
      var city=nrm(p.city||''), cluster=used&&used[city]?-25:0;
      var score=km(cur,p)+(km(p,origin)*0.30)+prevPenalty(p,mv)+cluster;
      if(score<bestScore){bestScore=score;best=p;}
    });
    return best;
  }

  function effectiveVisitMin(p){
    // Nel file resta 45 minuti come riferimento visita.
    // Per il carico giornata usiamo una media più reale, perché molte visite durano 10/15/20 minuti.
    if(isGrab(p))return 10;
    return 25;
  }

  window.assign = assign = function(ordered,days,start,mv){
    var left=(ordered||[]).slice(),out=[],prev=null;
    var transferCount=0;
    try{transferCount=transferDays();}catch(e){transferCount=0;}

    for(var di=0;di<days.length && left.length;di++){
      var tr=transferCount>0 && di<transferCount;
      var origin=(tr&&prev)?prev:start;
      var cur=origin,mins=0,count=0,used={};
      var maxCount=tr?10:9;
      var maxMinutes=tr?720:690;

      while(left.length && count<maxCount){
        var p=chooseMore(left,cur,origin,mv,used);
        if(!p)break;
        var leg=km(cur,p);
        var visitRef=visitMin(p);
        var visitCalc=effectiveVisitMin(p);
        var add=travelMin(leg)+visitCalc;
        var ret=tr?0:travelMin(km(p,origin));

        if(count>0 && mins+add+ret>maxMinutes)break;
        if(count>=7 && mins+add+ret>maxMinutes+45)break;

        left.splice(left.indexOf(p),1);
        var day=days[di],st=9*60+mins+travelMin(leg);
        mins+=add;
        used[nrm(p.city||'')]=1;
        cur=p;
        prev=p;
        count++;

        out.push(Object.assign({},p,{
          date:localIso(day),
          dateLabel:dateLabel(day),
          dateOnly:dateOnly(day),
          time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),
          travel_km:Math.round(leg),
          visit_min:visitRef,
          day_load:minText(mins+ret),
          return_km:tr?0:Math.round(km(p,origin)),
          planning_mode:tr?'trasferta':'giornaliero'
        }));
      }
    }
    window.__leftOut=left.length;
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
    renderTable(days.length);
    cleanUiMode();
  };

  var oldRenderAll=window.renderAll;
  if(typeof oldRenderAll==='function'){
    window.renderAll=function(){oldRenderAll.apply(this,arguments);cleanUiMode();};
  }
  var oldRenderTable=window.renderTable;
  if(typeof oldRenderTable==='function'){
    window.renderTable=function(){oldRenderTable.apply(this,arguments);cleanUiMode();};
  }

  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',cleanUiMode);else cleanUiMode();
  window.addEventListener('pageshow',cleanUiMode);
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, admin/capacity patch saltata")
        return
    html = path.read_text(encoding="utf-8")
    start = html.find('<!-- Planning admin link + higher daily capacity -->')
    while start != -1:
        end = html.find('</script>', start)
        if end == -1:
            break
        html = html[:start] + html[end + len('</script>'):]
        start = html.find('<!-- Planning admin link + higher daily capacity -->')
    html = html.replace('</body>', SNIPPET + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Planning admin link spostato sopra Dati planning e capacità giornaliera aggiornata")


if __name__ == "__main__":
    main()
