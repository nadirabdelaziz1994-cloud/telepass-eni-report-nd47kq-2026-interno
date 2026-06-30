from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

RUNTIME_PATCH = r'''
<script id="planning-period-runtime-patch">
(function(){
  let PLAN_PREV = {};
  let PLAN_PREV_FILE = null;

  function addStyles(){
    if(document.getElementById('planning-period-runtime-style')) return;
    const st=document.createElement('style');
    st.id='planning-period-runtime-style';
    st.textContent='.plan-period-extra{border:1px dashed var(--line);border-radius:12px;padding:8px;background:#f8fafc}.plan-period-extra label{font-weight:800;display:block;margin-bottom:4px}.plan-period-extra input,.plan-period-extra select{width:100%;padding:9px;border:1px solid var(--line);border-radius:10px}.plan-period-note{color:var(--muted);font-size:12px;line-height:1.3;margin-top:4px}';
    document.head.appendChild(st);
  }
  function normPdv(x){return String(x||'').replace(/\D+/g,'').padStart(5,'0');}
  function iso(d){return new Date(d).toISOString().slice(0,10);}
  function dateLabel(d){return new Date(d).toLocaleDateString('it-IT',{weekday:'short',day:'2-digit',month:'2-digit',year:'numeric'});}
  function dateOnly(d){return new Date(d).toLocaleDateString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric'});}
  function nextWorkdayAfter(d){const x=new Date(d);x.setDate(x.getDate()+1);while([0,6].includes(x.getDay()))x.setDate(x.getDate()+1);return x;}
  function extendWorkdaysFrom(startDate,n){const out=[];let d=new Date(startDate);while(out.length<n){if(![0,6].includes(d.getDay()))out.push(new Date(d));d.setDate(d.getDate()+1);}return out;}
  function sameMonth(dateIso,month){return String(dateIso||'').slice(0,7)===String(month||'');}
  function prevDatesInMonth(month){return Object.values(PLAN_PREV||{}).filter(d=>sameMonth(d,month)).sort();}
  function prevDoneSet(month){const s=new Set();Object.entries(PLAN_PREV||{}).forEach(([pdv,d])=>{if(sameMonth(d,month))s.add(normPdv(pdv));});return s;}
  function normDateInput(v){const s=String(v||'').trim();let a=s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](20\d{2})/),b=s.match(/(20\d{2})[\-.](\d{1,2})[\-.](\d{1,2})/);if(a)return a[3]+'-'+a[2].padStart(2,'0')+'-'+a[1].padStart(2,'0');if(b)return b[1]+'-'+b[2].padStart(2,'0')+'-'+b[3].padStart(2,'0');return '';}
  function collectPrevJson(node,map){if(Array.isArray(node)){node.forEach(x=>collectPrevJson(x,map));return;}if(!node||typeof node!=='object')return;const raw=node.pdv||node.PDV||node['n° PV']||node['n PV']||node.n_pv||node.punto_vendita||node['n° pv']||'';const pdv=normPdv(raw);const d=normDateInput(node.date||node.data||node.DATA||node.dateOnly||node.Data||node.dateLabel||'');if(pdv&&d)map[pdv]=d;Object.values(node).forEach(v=>{if(v&&typeof v==='object')collectPrevJson(v,map);});}
  function parsePrev(t){const txt=String(t||''),map={};try{const js=JSON.parse(txt);collectPrevJson(js,map);if(Object.keys(map).length)return map;}catch(e){}txt.split(/\n+/).forEach(line=>{const p=(line.match(/\b\d{3,6}\b/)||[])[0];if(!p)return;const d=normDateInput(line);if(d)map[normPdv(p)]=d;});return map;}
  async function loadPrevIfNeeded(){const f=document.getElementById('planPrev')?.files?.[0]||null;if(!f){PLAN_PREV={};PLAN_PREV_FILE=null;return;}if(PLAN_PREV_FILE===f.name)return;const txt=await f.text();PLAN_PREV=parsePrev(txt);PLAN_PREV_FILE=f.name;alert('Planning precedente caricato: '+Object.keys(PLAN_PREV).length+' PV trovati');}
  function firstDay(month,baseDays){const ds=prevDatesInMonth(month);if(!ds.length)return baseDays[0]||null;return nextWorkdayAfter(new Date(ds[ds.length-1]+'T00:00:00'));}
  function selectedDays(month){const mode=document.getElementById('planPeriodMode')?.value||'days',base=(typeof planWorkdays==='function'?planWorkdays(month):[]);if(!base.length)return[];let start=firstDay(month,base)||base[0];let days=base.filter(d=>d>=start);if(mode==='days'){const n=Math.max(1,Number(document.getElementById('planPeriodDays')?.value||15));if(!days.length)days=extendWorkdaysFrom(start,n);return days.slice(0,n);}if(mode==='month')return days;if(mode==='all'){if(!days.length)days=[start];return days;}return days;}
  function km(a,b){if(!a||!b||a.lat==null||b.lat==null)return 0;const R=6371,dLat=(b.lat-a.lat)*Math.PI/180,dLon=(b.lng-a.lng)*Math.PI/180,la1=a.lat*Math.PI/180,la2=b.lat*Math.PI/180;const x=Math.sin(dLat/2)**2+Math.cos(la1)*Math.cos(la2)*Math.sin(dLon/2)**2;return 2*R*Math.atan2(Math.sqrt(x),Math.sqrt(1-x));}
  function visitMin(p){return p&&p.is_grab?15:45;}
  function travelMin(k){return Math.round((k/55)*60+8);}
  function minText(n){const h=Math.floor(n/60),m=n%60;return h+'h '+String(m).padStart(2,'0');}
  function assignSmart(ordered,days,start,allowOverflow){let di=0,last=start,mins=0,count=0,out=[],curDays=days.slice();if(!curDays.length)curDays=extendWorkdaysFrom(new Date(),1);for(const p of ordered){while(di>=curDays.length){if(!allowOverflow)return out;curDays=curDays.concat(extendWorkdaysFrom(nextWorkdayAfter(curDays[curDays.length-1]),5));}let k=km(last,p),add=travelMin(k)+visitMin(p),maxCount=p.is_grab?7:6;let newDay=count>0&&((count>=maxCount)||(mins+add>420&&count>=2)||(mins+add>540));if(newDay){di++;if(di>=curDays.length){if(!allowOverflow)return out;curDays=curDays.concat(extendWorkdaysFrom(nextWorkdayAfter(curDays[curDays.length-1]),5));}mins=0;count=0;last=start;k=km(last,p);add=travelMin(k)+visitMin(p);}const day=curDays[di]||new Date();const st=9*60+mins+travelMin(km(last,p));out.push(Object.assign({},p,{date:iso(day),dateLabel:dateLabel(day),dateOnly:dateOnly(day),time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),travel_km:Math.round(k),visit_min:visitMin(p),day_load:minText(mins+add)}));mins+=add;count++;last=p;}return out;}
  function grabStateSafe(){try{if(typeof grabState==='function')return grabState();return JSON.parse(localStorage.getItem('grabGoVisits')||'{}')||{};}catch(e){return{};}}
  function grabIsVisited(p){const s=grabStateSafe();const x=s[String(p&&p.pdv||'')];return !!(x&&Number(x.month));}
  async function clearGrabVisitedFromPlan(items){const mode=document.getElementById('planGrab')?.value||'no';if(mode!=='all')return;const s=grabStateSafe();const pdvs=[...new Set((items||[]).filter(p=>p&&p.is_grab&&s[p.pdv]&&Number(s[p.pdv].month)).map(p=>p.pdv))];if(!pdvs.length)return;pdvs.forEach(pdv=>delete s[pdv]);if(typeof saveGrabState==='function')saveGrabState(s);else localStorage.setItem('grabGoVisits',JSON.stringify(s));if(typeof saveGrabRemote==='function'){for(const pdv of pdvs){try{await saveGrabRemote(pdv,0);}catch(e){}}}if(typeof renderGrabGo==='function')renderGrabGo();}

  function ensurePlanPeriodUI(){
    const grab=document.getElementById('planGrab');
    if(!grab) return;
    addStyles();
    if(grab.options.length<3 || ![...grab.options].some(o=>o.value==='only_new')){
      grab.innerHTML='<option value="no">No</option><option value="only_new">Sì, solo non visitati</option><option value="all">Sì, tutti</option>';
      if(!document.getElementById('planGrabNote')) grab.insertAdjacentHTML('afterend','<div id="planGrabNote" class="plan-period-note">Se scegli “Sì, tutti”, i Grab&Go già visitati possono rientrare nel planning e verranno rimessi come non visitati nella pagina Grab & Go.</div>');
    }
    if(document.getElementById('planPeriodMode')) return;
    const wrap=grab.closest('div')||grab.parentElement;
    const html='<div class="plan-period-extra"><label>Periodo planning</label><select id="planPeriodMode" onchange="window.planTogglePeriodDays&&window.planTogglePeriodDays()"><option value="days">Numero giorni</option><option value="month">1 mese</option><option value="all">Tutti i PDV</option></select><div class="plan-period-note">Numero giorni usa solo quei giorni lavorativi. 1 mese resta nel mese. Tutti i PDV continua anche oltre il mese se serve.</div></div><div class="plan-period-extra"><label>Giorni</label><input id="planPeriodDays" type="number" min="1" max="31" value="15"></div><div class="plan-period-extra"><label>Planning precedente / JSON</label><input id="planPrev" type="file" accept=".csv,.txt,.xls,.json"><div class="plan-period-note">Carica il JSON/CSV già creato: il sito esclude i PV già presenti e riparte dal giorno lavorativo successivo.</div></div><div class="plan-period-extra" style="align-self:end"><button class="btn light" type="button" onclick="window.planDownloadJson&&window.planDownloadJson()">Scarica JSON</button></div>';
    wrap.insertAdjacentHTML('beforebegin',html);
    planTogglePeriodDays();
  }

  async function newGeneratePlanning(){
    ensurePlanPeriodUI();
    const agent=document.getElementById('planAgent')?.value||'',month=document.getElementById('planMonth')?.value||'',grabMode=document.getElementById('planGrab')?.value||'no',startText=document.getElementById('planStart')?.value||'',periodMode=document.getElementById('planPeriodMode')?.value||'days';
    if(!agent||!month){alert('Scegli agente e mese');return;}
    if(typeof planPoints!=='function'||typeof planOrder!=='function'||typeof renderPlanningTable!=='function'){alert('Planning non ancora caricato, aggiorna la pagina.');return;}
    await loadPrevIfNeeded();
    try{if(typeof loadGrabRemote==='function')await loadGrabRemote();}catch(e){}
    const done=prevDoneSet(month);
    let pts=planPoints().filter(p=>{if(p.agent!==agent||p.lat==null||p.lng==null||done.has(normPdv(p.pdv)))return false;if(!p.is_grab)return true;if(grabMode==='no')return false;if(grabMode==='only_new')return !grabIsVisited(p);return true;});
    const noCoords=((window.APP&&APP.planning_data&&APP.planning_data.without_coordinates)||[]).filter(p=>p.agent===agent).length;
    if(!pts.length){alert('Nessun PV con coordinate per questo agente, oppure sono già tutti nel planning precedente caricato');return;}
    const days=selectedDays(month),start=typeof planStartPoint==='function'?planStartPoint(pts,startText):(pts[0]||null),ordered=planOrder(pts,start,month);
    window.PLAN=window.PLAN||{};
    PLAN.items=assignSmart(ordered,days,start,periodMode==='all');
    PLAN.source=agent+'_'+month;
    await clearGrabVisitedFromPlan(PLAN.items);
    renderPlanningTable(noCoords,days.length);
  }

  function planDownloadJson(){
    const items=(window.PLAN&&PLAN.items)||[];
    if(!items.length){alert('Prima crea il planning');return;}
    const data={created_at:new Date().toISOString(),agent:document.getElementById('planAgent')?.value||'',month:document.getElementById('planMonth')?.value||'',period_mode:document.getElementById('planPeriodMode')?.value||'',period_days:document.getElementById('planPeriodDays')?.value||'',items};
    const a=document.createElement('a');
    a.href=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json;charset=utf-8'}));
    a.download='planning_'+(data.agent||'agente')+'_'+(data.month||'mese')+'.json';
    a.click();
  }

  window.planTogglePeriodDays=function(){const mode=document.getElementById('planPeriodMode')?.value||'days';const box=document.getElementById('planPeriodDays');if(box)box.disabled=mode!=='days';};
  window.planDownloadJson=planDownloadJson;
  window.generatePlanning=newGeneratePlanning;
  setInterval(ensurePlanPeriodUI,700);
  document.addEventListener('DOMContentLoaded',()=>setTimeout(ensurePlanPeriodUI,400));
})();
</script>
'''


def patch_html(html: str) -> str:
    if 'planning-period-runtime-patch' in html:
        return html
    marker = '</body>'
    if marker in html:
        return html.replace(marker, RUNTIME_PATCH + '\n' + marker, 1)
    return html + '\n' + RUNTIME_PATCH


def main():
    patched = 0
    for name in ["index.html", "Telepass_ENI_sito_v6.html"]:
        path = DOCS_DIR / name
        if not path.exists():
            continue
        old = path.read_text(encoding="utf-8")
        new = patch_html(old)
        if new != old:
            path.write_text(new, encoding="utf-8")
            patched += 1
    print(f"Planning runtime patch completata: {patched} file aggiornati")


if __name__ == "__main__":
    main()
