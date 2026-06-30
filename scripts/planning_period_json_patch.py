from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH = r'''
<script id="safe-planning-period-json-patch">
(function(){
  var PLAN_PREV_EXTRA = {};
  var PLAN_PREV_FILE = '';
  var GRAB_VISITS_CACHE = null;

  function byId(id){return document.getElementById(id);}
  function normPdvSafe(v){var m=String(v||'').match(/\d+/);return m?m[0].padStart(5,'0'):'';}
  function isoSafe(d){return new Date(d).toISOString().slice(0,10);}
  function nextWorkdayAfterSafe(d){var x=new Date(d);x.setDate(x.getDate()+1);while(x.getDay()===0||x.getDay()===6){x.setDate(x.getDate()+1);}return x;}
  function extendWorkdaysSafe(startDate,n){var out=[],d=new Date(startDate);while(out.length<n){if(d.getDay()!==0&&d.getDay()!==6)out.push(new Date(d));d.setDate(d.getDate()+1);}return out;}
  function dateFromText(v){var s=String(v||'').trim();var a=s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](20\d{2})/);var b=s.match(/(20\d{2})[\-.](\d{1,2})[\-.](\d{1,2})/);if(a)return a[3]+'-'+a[2].padStart(2,'0')+'-'+a[1].padStart(2,'0');if(b)return b[1]+'-'+b[2].padStart(2,'0')+'-'+b[3].padStart(2,'0');return '';}
  function collectJson(node,map){
    if(Array.isArray(node)){node.forEach(function(x){collectJson(x,map);});return;}
    if(!node||typeof node!=='object')return;
    var pdv=normPdvSafe(node.pdv||node.PDV||node['n° PV']||node['n PV']||node.n_pv||node.punto_vendita||'');
    var d=dateFromText(node.date||node.data||node.DATA||node.dateOnly||node.Data||node.dateLabel||'');
    if(pdv&&d)map[pdv]=d;
    Object.keys(node).forEach(function(k){var v=node[k];if(v&&typeof v==='object')collectJson(v,map);});
  }
  function parsePrevSafe(txt){
    var map={};txt=String(txt||'');
    try{var js=JSON.parse(txt);collectJson(js,map);if(Object.keys(map).length)return map;}catch(e){}
    txt.split(/\n+/).forEach(function(line){var p=(line.match(/\b\d{3,6}\b/)||[])[0];if(!p)return;var d=dateFromText(line);if(d)map[p.padStart(5,'0')]=d;});
    return map;
  }
  async function loadPrevFile(){
    var f=byId('prev')&&byId('prev').files?byId('prev').files[0]:null;
    if(!f){PLAN_PREV_EXTRA={};PLAN_PREV_FILE='';return;}
    if(PLAN_PREV_FILE===f.name)return;
    var txt=await f.text();PLAN_PREV_EXTRA=parsePrevSafe(txt);PLAN_PREV_FILE=f.name;
    try{window.PREV=Object.assign({},window.PREV||{},PLAN_PREV_EXTRA);}catch(e){}
    alert('Planning precedente caricato: '+Object.keys(PLAN_PREV_EXTRA).length+' PV trovati');
  }
  function prevDoneSet(month){
    var s=new Set();var all=Object.assign({},window.PREV||{},PLAN_PREV_EXTRA||{});
    Object.keys(all).forEach(function(pdv){var d=all[pdv];if(String(d||'').slice(0,7)===month)s.add(normPdvSafe(pdv));});
    return s;
  }
  function prevLastDate(month){
    var vals=[];var all=Object.assign({},window.PREV||{},PLAN_PREV_EXTRA||{});
    Object.keys(all).forEach(function(pdv){var d=all[pdv];if(String(d||'').slice(0,7)===month)vals.push(d);});
    vals.sort();return vals.length?vals[vals.length-1]:'';
  }
  function selectedDays(month){
    var mode=(byId('periodMode')&&byId('periodMode').value)||'all';
    var base=(typeof workdays==='function'?workdays(month):[]);
    if(!base.length)return [];
    var last=prevLastDate(month);var start=last?nextWorkdayAfterSafe(new Date(last+'T00:00:00')):base[0];
    var days=base.filter(function(d){return d>=start;});
    if(mode==='days'){
      var n=Math.max(1,Number((byId('periodDays')&&byId('periodDays').value)||15));
      if(!days.length)days=extendWorkdaysSafe(start,n);
      return days.slice(0,n);
    }
    if(mode==='month')return days;
    if(mode==='all'){
      var after=extendWorkdaysSafe(days.length?nextWorkdayAfterSafe(days[days.length-1]):start,45);
      return days.concat(after);
    }
    return days;
  }
  async function grabMap(){
    if(GRAB_VISITS_CACHE)return GRAB_VISITS_CACHE;
    var map={};
    try{var res=await fetch('/grab-visite?ts='+Date.now(),{cache:'no-store'});var data=await res.json();if(res.ok&&data.ok){(data.visite||[]).forEach(function(v){var pdv=normPdvSafe(v.pdv);if(pdv&&Number(v.month))map[pdv]=v;});}}catch(e){}
    GRAB_VISITS_CACHE=map;return map;
  }
  function isGrabVisited(p,map){return !!(map&&map[normPdvSafe(p&&p.pdv)]&&Number(map[normPdvSafe(p&&p.pdv)].month));}
  async function clearGrabVisited(items,map){
    var mode=(byId('grab')&&byId('grab').value)||'no';if(mode!=='all')return;
    var seen={};(items||[]).forEach(function(p){if(typeof isGrab==='function'&&isGrab(p)&&isGrabVisited(p,map))seen[normPdvSafe(p.pdv)]=true;});
    var pdvs=Object.keys(seen);if(!pdvs.length)return;
    for(var i=0;i<pdvs.length;i++){
      try{await fetch('/grab-visita',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pdv:pdvs[i],month:0,year:new Date().getFullYear()})});}catch(e){}
    }
    GRAB_VISITS_CACHE=null;
  }
  function injectUi(){
    var month=byId('month'), start=byId('start'), grab=byId('grab'), prev=byId('prev');
    if(!month||!start||!grab||!prev)return;
    if(!byId('periodMode')){
      var wrap=start.closest('div')||start.parentElement;
      if(wrap){wrap.insertAdjacentHTML('beforebegin','<div><label>Periodo planning</label><select id="periodMode" onchange="window.togglePeriodDays&&window.togglePeriodDays()"><option value="days">Numero giorni</option><option value="month">1 mese</option><option value="all" selected>Tutti i PDV</option></select><div class="muted" style="font-size:11px;margin-top:4px">Numero giorni usa solo quei giorni lavorativi. 1 mese resta nel mese. Tutti i PDV continua oltre il mese se serve.</div></div><div><label>Giorni</label><input id="periodDays" type="number" min="1" max="31" value="15"></div>');}
    }
    if(!grab.querySelector('option[value="only_new"]')){
      grab.innerHTML='<option value="no">No</option><option value="only_new">Sì, solo non visitati</option><option value="all">Sì, tutti</option>';
      grab.insertAdjacentHTML('afterend','<div class="muted" style="font-size:11px;margin-top:4px">Se scegli “Sì, tutti”, i Grab&Go già visitati possono rientrare nel planning e verranno rimessi come non visitati nella pagina Grab & Go.</div>');
    }
    try{prev.setAttribute('accept','.csv,.txt,.xls,.json');var lab=prev.closest('div').querySelector('label');if(lab)lab.textContent='Planning precedente / JSON';}catch(e){}
    if(!byId('downloadJsonBtn')){
      var csvBtn=[].slice.call(document.querySelectorAll('button')).find(function(b){return String(b.textContent||'').includes('CSV');});
      var html='<button id="downloadJsonBtn" class="btn light" onclick="window.downloadJson&&window.downloadJson()">Scarica JSON</button>';
      if(csvBtn)csvBtn.insertAdjacentHTML('afterend',html);else prev.closest('section').querySelector('.actions').insertAdjacentHTML('beforeend',html);
    }
    togglePeriodDays();
  }
  window.togglePeriodDays=function(){var mode=(byId('periodMode')&&byId('periodMode').value)||'all';var box=byId('periodDays');if(box)box.disabled=mode!=='days';};
  window.downloadJson=function(){
    var plan=window.PLAN||[];if(!plan.length){alert('Prima crea il planning');return;}
    var data={created_at:new Date().toISOString(),agent:(byId('agent')&&byId('agent').value)||'',month:(byId('month')&&byId('month').value)||'',period_mode:(byId('periodMode')&&byId('periodMode').value)||'',period_days:(byId('periodDays')&&byId('periodDays').value)||'',grab_mode:(byId('grab')&&byId('grab').value)||'',items:plan};
    var a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json;charset=utf-8'}));a.download='planning_'+(data.agent||'agente')+'_'+(data.month||'mese')+'.json';a.click();
  };
  var oldGenerate=window.generatePlanning;
  window.generatePlanning=async function(){
    injectUi();
    var agent=(byId('agent')&&byId('agent').value)||'', month=(byId('month')&&byId('month').value)||'', grabMode=(byId('grab')&&byId('grab').value)||'no';
    if(!agent||!month){alert('Scegli agente e mese');return;}
    if(typeof allPoints!=='function'||typeof startPoint!=='function'||typeof order!=='function'||typeof assign!=='function'){if(typeof oldGenerate==='function')return oldGenerate();return;}
    await loadPrevFile();
    var done=prevDoneSet(month), gm=await grabMap();
    var pts=allPoints().filter(function(p){
      if(p.agent_display!==agent||p.lat==null||p.lng==null||done.has(normPdvSafe(p.pdv)))return false;
      if(!(typeof isGrab==='function'&&isGrab(p)))return true;
      if(grabMode==='no')return false;
      if(grabMode==='only_new')return !isGrabVisited(p,gm);
      return true;
    });
    if(!pts.length){alert('Nessun PV con coordinate per questo agente, oppure sono già tutti nel planning precedente caricato');return;}
    var days=selectedDays(month), start=startPoint(pts,(byId('start')&&byId('start').value)||''), ordered=order(pts,start,month);
    window.PLAN=assign(ordered,days,start);
    await clearGrabVisited(window.PLAN,gm);
    if(typeof renderTable==='function')renderTable(days.length);
  };
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',injectUi);else injectUi();
  setTimeout(injectUi,500);setTimeout(injectUi,1200);
})();
</script>
'''


def main():
    path = DOCS_DIR / 'planning.html'
    if not path.exists():
        print('planning.html non trovato')
        return
    html = path.read_text(encoding='utf-8')
    if 'safe-planning-period-json-patch' not in html:
        html = html.replace('</body>', PATCH + '\n</body>', 1)
        path.write_text(html, encoding='utf-8')
        print('Safe planning period/json patch applicata')
    else:
        print('Safe planning period/json patch già presente')


if __name__ == '__main__':
    main()
