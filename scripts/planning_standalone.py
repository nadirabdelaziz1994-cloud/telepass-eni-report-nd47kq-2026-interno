from pathlib import Path
import datetime
import json

from planning_scope_patch import build_planning_data

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"
API_URL = "https://telepass-planning-api.nadirabdelaziz1994.workers.dev"

HTML = r'''<!doctype html>
<html lang="it">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Planning automatico - Telepass ENI</title>
  <style>
    :root{--blue:#0b2d5c;--accent:#0d6efd;--bg:#f5f7fb;--card:#fff;--line:#d9e2ef;--muted:#64748b;--text:#102033;--good:#166534;--warn:#92400e;--bad:#991b1b;--grab:#7c3aed;--grabbg:#f3e8ff}
    *{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font-family:Arial,Helvetica,sans-serif}header{background:var(--blue);color:white;padding:12px 16px;position:sticky;top:0;z-index:2;box-shadow:0 2px 10px #0002}header .row{display:flex;gap:10px;align-items:center;justify-content:space-between;flex-wrap:wrap}header h1{font-size:18px;margin:0}a.top{color:white;text-decoration:none;border:1px solid #ffffff55;border-radius:999px;padding:7px 10px;font-weight:800;font-size:12px}.wrap{max-width:1260px;margin:0 auto;padding:14px}.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:14px;margin:12px 0;box-shadow:0 2px 12px #12376410}.title{font-size:18px;font-weight:900;color:var(--blue);margin:0 0 8px}.muted{color:var(--muted);font-size:13px;line-height:1.35}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px}.form label{font-size:12px;color:var(--muted);font-weight:800;display:block;margin-bottom:4px}.form input,.form select{width:100%;padding:10px;border:1px solid var(--line);border-radius:11px;background:white}.btn{border:0;border-radius:11px;padding:10px 12px;background:var(--accent);color:white;font-weight:900;cursor:pointer}.btn.light{background:#edf4ff;color:var(--blue);border:1px solid #bfd4ef}.btn.bad{background:#fee2e2;color:var(--bad);border:1px solid #fecaca}.btn.good{background:#dcfce7;color:var(--good);border:1px solid #bbf7d0}.actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.metric{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}.metric .box{background:#f8fbff;border:1px solid var(--line);border-radius:14px;padding:12px}.metric h4{margin:0 0 6px;font-size:12px;color:var(--muted)}.big{font-size:26px;font-weight:900;color:var(--blue)}.pill{display:inline-flex;border-radius:999px;padding:4px 8px;font-size:12px;font-weight:900;background:#eef4fb;color:var(--blue);margin:2px}.pill.good{background:#dcfce7;color:var(--good)}.pill.bad{background:#fee2e2;color:var(--bad)}.pill.warn{background:#fef3c7;color:var(--warn)}.pill.grab{background:var(--grabbg);color:var(--grab)}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:14px;background:white}table{border-collapse:collapse;width:100%;min-width:1180px}th,td{border-bottom:1px solid var(--line);padding:9px;text-align:left;font-size:13px;vertical-align:top}th{background:#f1f5f9;color:#334155;font-size:12px;position:sticky;top:0}.num{text-align:right;font-variant-numeric:tabular-nums}.city b{display:block}.city span{color:var(--muted);font-size:12px}.tp{box-shadow:inset 5px 0 0 var(--accent)}.grab{box-shadow:inset 5px 0 0 var(--grab);background:#fbf7ff}.small{font-size:12px}.status-ok{color:var(--good);font-weight:900}.status-bad{color:var(--bad);font-weight:900}
  </style>
</head>
<body>
  <script id="planning-data" type="application/json">__DATA_JSON__</script>
  <header><div class="row"><h1>Planning automatico</h1><a class="top" href="./index.html?v=fix1">← Torna al sito</a></div></header>
  <main class="wrap">
    <section class="card">
      <div class="title">Planning automatico definitivo</div>
      <div class="muted">Pagina separata di sicurezza: non modifica Home/Classifica/Bundle/Grab & Go. Le aggiunte/esclusioni definitive passano da Cloudflare.</div>
      <div id="sourceBox" class="muted" style="margin-top:8px"></div>
    </section>
    <section class="metric" id="metricBox"></section>
    <section class="card">
      <div class="title">Dati planning</div>
      <div class="grid form">
        <div><label>Agente</label><select id="agent"></select></div>
        <div><label>Mese</label><input id="month" type="month"></div>
        <div><label>Punto di partenza</label><input id="start" placeholder="Città o codice PV"></div>
        <div><label>Grab & Go</label><select id="grab"><option value="no">No</option><option value="yes">Sì, includili</option></select></div>
        <div><label>Planning precedente</label><input id="prev" type="file" accept=".csv,.txt,.xls"></div>
      </div>
      <div class="actions">
        <button class="btn" onclick="generatePlanning()">Crea planning</button>
        <button class="btn light" onclick="downloadXls()">Scarica Excel</button>
        <button class="btn light" onclick="downloadCsv()">Scarica CSV</button>
        <button class="btn light" onclick="loadRemote(false)">Ricarica modifiche definitive</button>
      </div>
      <div class="muted" style="margin-top:8px">Export Excel con le stesse colonne del file planning originale. I Grab & Go sono evidenziati, senza colonne extra.</div>
    </section>
    <section class="card">
      <div class="title">Cloudflare definitivo</div>
      <div id="apiStatus" class="muted">Controllo collegamento...</div>
      <div class="grid form" style="margin-top:12px">
        <div><label>Aggiungi PV da anagrafica</label><input id="addPdv" placeholder="Codice PV"></div>
        <div><label>Tipo</label><select id="addType"><option value="tp">Telepass Point</option><option value="grab">Grab & Go</option></select></div>
        <div style="align-self:end"><button class="btn good" onclick="addFromCatalog()">Aggiungi definitivo</button></div>
      </div>
      <div id="addPreview" class="muted" style="margin-top:8px"></div>
      <hr style="border:0;border-top:1px solid var(--line);margin:14px 0">
      <div class="grid form">
        <div><label>Escludi PV</label><input id="removePdv" placeholder="Codice PV da togliere"></div>
        <div style="align-self:end"><button class="btn bad" onclick="removePdvDef()">Escludi definitivo</button></div>
        <div><label>Ripristina PV</label><input id="restorePdv" placeholder="Codice PV da ripristinare"></div>
        <div style="align-self:end"><button class="btn light" onclick="restorePdvDef()">Ripristina</button></div>
      </div>
      <hr style="border:0;border-top:1px solid var(--line);margin:14px 0">
      <div class="title" style="font-size:15px">Aggiunta manuale</div>
      <div class="grid form">
        <div><label>PV</label><input id="manualPdv" placeholder="Codice PV"></div>
        <div><label>Città</label><input id="manualCity" placeholder="Città"></div>
        <div><label>Indirizzo</label><input id="manualAddress" placeholder="Indirizzo"></div>
        <div><label>Latitudine</label><input id="manualLat" placeholder="es. 45.4642"></div>
        <div><label>Longitudine</label><input id="manualLng" placeholder="es. 9.1900"></div>
        <div><label>Tipo</label><select id="manualType"><option value="tp">Telepass Point</option><option value="grab">Grab & Go</option></select></div>
        <div style="align-self:end"><button class="btn good" onclick="addManual()">Aggiungi manuale definitivo</button></div>
      </div>
    </section>
    <section id="result"></section>
  </main>
<script>
const API_URL = "__API_URL__";
const DATA = JSON.parse(document.getElementById('planning-data').textContent || '{}');
let REMOTE = [];
let PLAN = [];
let PREV = {};
let apiReady = false;

function esc(v){return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function fmt(v){return Number(v || 0).toLocaleString('it-IT');}
function normPdv(v){const m=String(v||'').match(/\d+/);return m?m[0].padStart(5,'0'):'';}
function norm(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}
function titleCase(s){return String(s||'').toLowerCase().split(/\s+/).filter(Boolean).map(x=>x.charAt(0).toUpperCase()+x.slice(1)).join(' ');}
function toNum(v){const n=Number(String(v??'').replace(',','.'));return Number.isFinite(n)?n:null;}
function agentKey(s){return norm(s).split(' ').filter(x=>x&&x!=='myworldsrl'&&x!=='it').sort().join('|');}
function rawEmailToName(s){const local=String(s||'').split('@')[0].replace(/[._-]+/g,' ');return titleCase(local);}
function visibleAgentName(a){a=String(a||'').trim();if(!a)return '';if(a.includes('@')){const guessed=rawEmailToName(a);return AGENT_CANON[agentKey(guessed)]||guessed;}return a;}
let AGENT_CANON = {};
function rebuildAgentCanon(){AGENT_CANON={};const raw=[...(DATA.points||[]),...(DATA.catalog||[])].map(p=>String(p.agent||'').trim()).filter(Boolean).filter(a=>!a.includes('@'));raw.forEach(a=>{const k=agentKey(a);if(k&&!AGENT_CANON[k])AGENT_CANON[k]=a;});}

function acVal(p){const r=norm(p.region||'');if(p.ac)return p.ac;if(['lombardia','piemonte','valle d aosta','liguria'].includes(r))return 'NORD OVEST';if(['veneto','friuli venezia giulia','trentino alto adige','emilia romagna'].includes(r))return 'NORD EST';if(['toscana','umbria','marche','lazio','abruzzo','molise'].includes(r))return 'CENTRO';if(['campania','puglia','basilicata','calabria','sicilia','sardegna'].includes(r))return 'SUD';return '';}
function isGrab(p){return !!p.is_grab;}

function getAdminKey(){let k=sessionStorage.getItem('planningAdminKey')||'';if(!k){k=prompt('Inserisci ADMIN_KEY Cloudflare');if(k)sessionStorage.setItem('planningAdminKey',k);}return k;}
async function api(path,opt){const res=await fetch(API_URL+path,opt||{});const data=await res.json().catch(()=>({ok:false,error:'Risposta non valida'}));if(!res.ok||data.ok===false)throw new Error(data.error||('Errore API '+res.status));return data;}
async function loadRemote(silent){const box=document.getElementById('apiStatus');try{const data=await api('/modifiche');REMOTE=data.modifiche||[];apiReady=true;box.innerHTML='<span class="status-ok">Cloudflare collegato</span> · modifiche definitive caricate: <b>'+fmt(REMOTE.length)+'</b><br><span class="small">'+esc(API_URL)+'</span>';renderAll();if(!silent)alert('Modifiche ricaricate');}catch(e){apiReady=false;box.innerHTML='<span class="status-bad">Cloudflare non raggiungibile</span>: '+esc(e.message)+'<br><span class="small">Il planning funziona coi dati base, ma aggiunte/esclusioni definitive no.</span>';renderAll();}}
async function saveRemote(mod){const key=getAdminKey();if(!key)return;await api('/modifica',{method:'POST',headers:{'Content-Type':'application/json','X-Admin-Key':key},body:JSON.stringify(mod)});await loadRemote(true);}
async function restoreRemote(pdv){const key=getAdminKey();if(!key)return;await api('/ripristina',{method:'POST',headers:{'Content-Type':'application/json','X-Admin-Key':key},body:JSON.stringify({pdv})});await loadRemote(true);}

function catalogPoint(pdv){const n=normPdv(pdv);return (DATA.catalog||[]).find(p=>p.pdv===n)||null;}
function basePoints(){return (DATA.points||[]).filter(p=>p&&p.pdv&&p.lat!=null&&p.lng!=null);}
function applyRemote(){const by={};basePoints().forEach(p=>by[p.pdv]=Object.assign({},p));(REMOTE||[]).forEach(m=>{const pnum=normPdv(m.pdv);if(!pnum)return;const action=String(m.action||'').toUpperCase();if(action==='ESCLUDI'){delete by[pnum];return;}if(action==='AGGIUNGI'){const cat=catalogPoint(pnum)||{pdv:pnum};const p=Object.assign({},cat);p.pdv=pnum;p.agent=m.agente||p.agent||'';p.city=m.citta||p.city||'';p.address=m.indirizzo||p.address||'';p.lat=toNum(m.latitudine)!=null?toNum(m.latitudine):p.lat;p.lng=toNum(m.longitudine)!=null?toNum(m.longitudine):p.lng;const tipo=String(m.tipo||'').toLowerCase();p.is_grab=tipo.includes('grab')||p.is_grab;p.is_tp=!p.is_grab||p.is_tp;p.source='Cloudflare definitivo';if(p.lat!=null&&p.lng!=null)by[pnum]=p;}});return Object.values(by).map(p=>Object.assign({},p,{agent_display:visibleAgentName(p.agent)}));}
function allPoints(){return applyRemote();}
function agents(){return [...new Set(allPoints().map(p=>p.agent_display).filter(Boolean))].filter(a=>!a.includes('@')).sort((a,b)=>a.localeCompare(b,'it'));}

function renderAll(){renderSource();renderMetrics();renderAgents();previewAdd();if(PLAN.length)renderTable();}
function renderSource(){document.getElementById('sourceBox').innerHTML='Fonte dati: <b>'+esc(DATA.source_name||'')+'</b><br>Aggiornato: '+esc(DATA.generated_at||'');}
function renderMetrics(){const sum=DATA.summary||{},pts=allPoints(),remEx=(REMOTE||[]).filter(x=>String(x.action||'').toUpperCase()==='ESCLUDI').length,remAdd=(REMOTE||[]).filter(x=>String(x.action||'').toUpperCase()==='AGGIUNGI').length,grabs=pts.filter(isGrab).length;document.getElementById('metricBox').innerHTML='<div class="box"><h4>PV base con coordinate</h4><div class="big">'+fmt(sum.active_with_coordinates||basePoints().length)+'</div></div><div class="box"><h4>PV usabili ora</h4><div class="big">'+fmt(pts.length)+'</div></div><div class="box"><h4>Grab & Go usabili</h4><div class="big">'+fmt(grabs)+'</div></div><div class="box"><h4>Modifiche definitive</h4><div><span class="pill good">Aggiunti '+fmt(remAdd)+'</span><span class="pill bad">Esclusi '+fmt(remEx)+'</span></div></div>';}
function renderAgents(){const sel=document.getElementById('agent');const cur=sel.value;sel.innerHTML='<option value="">Seleziona agente</option>'+agents().map(a=>'<option>'+esc(a)+'</option>').join('');if(cur)sel.value=cur;}
function previewAdd(){const box=document.getElementById('addPreview');const p=catalogPoint(document.getElementById('addPdv')?.value);box.innerHTML=p?'Trovato: <b>'+esc(p.pdv)+'</b> · '+esc(p.city||'')+' · '+esc(p.address||'')+' · coordinate: '+(p.lat&&p.lng?'<span class="status-ok">ok</span>':'<span class="status-bad">mancanti</span>'):'Inserisci un PV per cercarlo in anagrafica.';}

async function addFromCatalog(){const agent=document.getElementById('agent').value;if(!agent){alert('Prima seleziona agente');return;}const p=catalogPoint(document.getElementById('addPdv').value);if(!p){alert('PV non trovato in anagrafica');return;}if(p.lat==null||p.lng==null){alert('PV senza coordinate. Usa aggiunta manuale con latitudine/longitudine.');return;}const tipo=document.getElementById('addType').value==='grab'?'Grab & Go':'Telepass Point';try{await saveRemote({action:'AGGIUNGI',pdv:p.pdv,agente:agent,tipo,citta:p.city||'',indirizzo:p.address||'',latitudine:p.lat,longitudine:p.lng,note:'Aggiunto da planning.html'});alert('PV aggiunto definitivamente');}catch(e){alert('Errore: '+e.message);}}
async function addManual(){const agent=document.getElementById('agent').value;if(!agent){alert('Prima seleziona agente');return;}const pdv=normPdv(document.getElementById('manualPdv').value),lat=toNum(document.getElementById('manualLat').value),lng=toNum(document.getElementById('manualLng').value);if(!pdv||lat==null||lng==null){alert('Servono PV, latitudine e longitudine');return;}const tipo=document.getElementById('manualType').value==='grab'?'Grab & Go':'Telepass Point';try{await saveRemote({action:'AGGIUNGI',pdv,agente:agent,tipo,citta:document.getElementById('manualCity').value,indirizzo:document.getElementById('manualAddress').value,latitudine:lat,longitudine:lng,note:'Aggiunto manuale da planning.html'});alert('PV manuale aggiunto definitivamente');}catch(e){alert('Errore: '+e.message);}}
async function removePdvDef(){const pdv=normPdv(document.getElementById('removePdv').value);if(!pdv){alert('Inserisci PV');return;}if(!confirm('Escludere definitivamente il PV '+pdv+'?'))return;try{await saveRemote({action:'ESCLUDI',pdv,note:'Escluso da planning.html'});PLAN=PLAN.filter(p=>p.pdv!==pdv);alert('PV escluso definitivamente');}catch(e){alert('Errore: '+e.message);}}
async function restorePdvDef(){const pdv=normPdv(document.getElementById('restorePdv').value);if(!pdv){alert('Inserisci PV');return;}try{await restoreRemote(pdv);alert('PV ripristinato');}catch(e){alert('Errore: '+e.message);}}

function km(a,b){if(!a||!b||a.lat==null||b.lat==null)return 0;const R=6371,dLat=(b.lat-a.lat)*Math.PI/180,dLon=(b.lng-a.lng)*Math.PI/180,la1=a.lat*Math.PI/180,la2=b.lat*Math.PI/180;const x=Math.sin(dLat/2)**2+Math.cos(la1)*Math.cos(la2)*Math.sin(dLon/2)**2;return 2*R*Math.atan2(Math.sqrt(x),Math.sqrt(1-x));}
function iso(d){return d.toISOString().slice(0,10);}function dateLabel(d){return d.toLocaleDateString('it-IT',{weekday:'short',day:'2-digit',month:'2-digit',year:'numeric'});}function dateOnly(d){return d.toLocaleDateString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric'});}
function easter(y){const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),mo=Math.floor((h+l-7*m+114)/31),da=((h+l-7*m+114)%31)+1;return new Date(y,mo-1,da);}
function holidays(y){const out=new Set([`${y}-01-01`,`${y}-01-06`,`${y}-04-25`,`${y}-05-01`,`${y}-06-02`,`${y}-08-15`,`${y}-11-01`,`${y}-12-08`,`${y}-12-25`,`${y}-12-26`]);const e=easter(y),p=new Date(e);p.setDate(e.getDate()+1);out.add(iso(p));return out;}
function workdays(mv){const [y,m]=String(mv||'').split('-').map(Number);if(!y||!m)return[];const h=holidays(y),out=[];for(let d=new Date(y,m-1,1);d.getMonth()===m-1;d.setDate(d.getDate()+1)){const w=d.getDay();if(w!==0&&w!==6&&!h.has(iso(d)))out.push(new Date(d));}return out;}
function weekNum(dateIso){const d=new Date(dateIso);const onejan=new Date(d.getFullYear(),0,1);return Math.ceil((((d-onejan)/86400000)+onejan.getDay()+1)/7);}
function visitMin(p){return isGrab(p)?15:45;}function travelMin(k){return Math.round((k/55)*60+8);}function minText(n){const h=Math.floor(n/60),m=n%60;return h+'h '+String(m).padStart(2,'0');}
function startPoint(points,start){const s=norm(start);if(!s)return points[0]||null;return points.find(p=>norm(p.pdv)===s||norm(p.city)===s)||points.find(p=>norm(`${p.pdv} ${p.city} ${p.address}`).includes(s))||points[0]||null;}
function prevAge(p,mv){const d=PREV[p.pdv];if(!d)return 9999;const [y,m]=String(mv||'').split('-').map(Number);return Math.round((new Date(y,m-1,1)-new Date(d))/(1000*3600*24));}
function order(points,start,mv){let left=points.slice(),out=[],cur=start||left[0];while(left.length){left.sort((a,b)=>{const pa=prevAge(a,mv),pb=prevAge(b,mv),pena=pa<45?500:pa<90?120:0,penb=pb<45?500:pb<90?120:0;return (km(cur,a)+pena)-(km(cur,b)+penb);});const n=left.shift();out.push(n);cur=n;}return out;}
function assign(ordered,days,start){let di=0,last=start,mins=0,count=0,out=[];const targetDays=Math.min(days.length,Math.max(1,Math.ceil(ordered.length/3)));const targetPerDay=Math.max(1,Math.ceil(ordered.length/targetDays));for(const p of ordered){const k=km(last,p),add=travelMin(k)+visitMin(p);const newDay=count>0&&((count>=targetPerDay)||(mins+add>420&&count>=2)||(mins+add>540));if(newDay&&di<days.length-1){di++;mins=0;count=0;last=start;}const day=days[Math.min(di,days.length-1)]||new Date();const st=9*60+mins+travelMin(km(last,p));out.push(Object.assign({},p,{date:iso(day),dateLabel:dateLabel(day),dateOnly:dateOnly(day),time:String(Math.floor(st/60)).padStart(2,'0')+':'+String(st%60).padStart(2,'0'),travel_km:Math.round(k),visit_min:visitMin(p),day_load:minText(mins+add)}));mins+=add;count++;last=p;}return out;}
function generatePlanning(){const agent=document.getElementById('agent').value,mv=document.getElementById('month').value,inc=document.getElementById('grab').value==='yes';if(!agent||!mv){alert('Scegli agente e mese');return;}let pts=allPoints().filter(p=>p.agent_display===agent&&(inc||!isGrab(p))&&p.lat!=null&&p.lng!=null);if(!pts.length){alert('Nessun PV con coordinate per questo agente');return;}const days=workdays(mv),start=startPoint(pts,document.getElementById('start').value);PLAN=assign(order(pts,start,mv),days,start);renderTable(days.length);}
function renderTable(workdayCount){const days=[...new Set(PLAN.map(p=>p.date))].length,grab=PLAN.filter(isGrab).length;const rows=PLAN.map((p,i)=>'<tr class="'+(isGrab(p)?'grab':'tp')+'"><td><button class="btn light small" onclick="moveRow('+i+',-1)">↑</button> <button class="btn light small" onclick="moveRow('+i+',1)">↓</button> <button class="btn bad small" onclick="delRow('+i+')">x</button></td><td>'+esc(p.dateLabel)+'</td><td>'+esc(p.pdv)+'</td><td>'+(isGrab(p)?'<span class="pill grab">Grab & Go</span>':'<span class="pill">Telepass Point</span>')+'</td><td class="city"><b>'+esc(p.city)+'</b><span>'+esc(p.address||'')+'</span></td><td>'+esc(p.province||'')+'</td><td>'+esc(p.region||'')+'</td><td>'+esc(p.rzv||'')+'</td><td>'+esc(p.cr||'')+'</td><td>'+esc(p.focal||'')+'</td><td class="num">'+fmt(p.travel_km||0)+' km</td><td class="num">'+fmt(p.visit_min||0)+' min</td><td>'+esc(p.day_load||'')+'</td></tr>').join('');document.getElementById('result').innerHTML='<section class="metric"><div class="box"><h4>PV pianificati</h4><div class="big">'+fmt(PLAN.length)+'</div></div><div class="box"><h4>Grab & Go pianificati</h4><div class="big">'+fmt(grab)+'</div></div><div class="box"><h4>Giorni usati</h4><div class="big">'+fmt(days)+'</div><div class="muted">Lavorativi: '+fmt(workdayCount||0)+'</div></div></section><section class="card"><div class="muted" style="margin-bottom:8px">Sposta le righe: date e ordine vengono ricalcolati. I Grab & Go sono evidenziati.</div><div class="table-wrap"><table><thead><tr><th>Modifica</th><th>Data</th><th>PV</th><th>Tipo</th><th>Città / indirizzo</th><th>Prov.</th><th>Regione</th><th>RZV</th><th>CR</th><th>Focal Point ENI</th><th>Km</th><th>Visita</th><th>Carico giorno</th></tr></thead><tbody>'+rows+'</tbody></table></div></section>';}
function recalc(){const days=workdays(document.getElementById('month').value),start=startPoint(PLAN,document.getElementById('start').value);PLAN=assign(PLAN,days,start);renderTable(days.length);}function moveRow(i,d){const j=i+d;if(j<0||j>=PLAN.length)return;[PLAN[i],PLAN[j]]=[PLAN[j],PLAN[i]];recalc();}function delRow(i){PLAN.splice(i,1);recalc();}
function exportRows(){return [['n° WEEK','DATA','ORA ','n° PV','AC','RZV','CR','Regione','Provincia','Città ','Indirizzo','FOCAL POINT ENI','MY WORLD','CONFERMA ENI']].concat(PLAN.map(p=>[weekNum(p.date),p.dateOnly,'',p.pdv,acVal(p),p.rzv||'',p.cr||'',p.region||'',p.province||'',p.city||'',p.address||'',p.focal||'',p.agent_display||document.getElementById('agent').value,'']));}
function downloadCsv(){if(!PLAN.length){alert('Prima crea il planning');return;}const csv=exportRows().map(r=>r.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(';')).join('\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));a.download='planning_'+(document.getElementById('agent').value||'agente')+'_'+(document.getElementById('month').value||'mese')+'.csv';a.click();}
function downloadXls(){if(!PLAN.length){alert('Prima crea il planning');return;}const rows=exportRows();let html='<html><head><meta charset="UTF-8"><style>table{border-collapse:collapse;font-family:Calibri,Arial;font-size:11pt}th{background:#1f4e78;color:#fff;font-weight:bold;text-align:center}td,th{border:1px solid #999;padding:5px}td.text{mso-number-format:"\\@"}.grab td{background:#eadcf8}.center{text-align:center}</style></head><body><table>';html+=rows.map((r,i)=>{if(i===0)return '<tr>'+r.map(v=>'<th>'+esc(v)+'</th>').join('')+'</tr>';const p=PLAN[i-1];return '<tr class="'+(isGrab(p)?'grab':'')+'">'+r.map((v,j)=>'<td class="'+(j===3?'text ':'')+(j===0||j===1||j===2||j===13?'center':'')+'">'+esc(v)+'</td>').join('')+'</tr>';}).join('');html+='</table></body></html>';const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([html],{type:'application/vnd.ms-excel'}));a.download='planning_'+(document.getElementById('agent').value||'agente')+'_'+(document.getElementById('month').value||'mese')+'.xls';a.click();}
function parsePrev(t){const map={};String(t||'').split(/\n+/).forEach(line=>{const p=(line.match(/\b\d{3,6}\b/)||[])[0];if(!p)return;let d=null;const a=line.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](20\d{2})/),b=line.match(/(20\d{2})[\-.](\d{1,2})[\-.](\d{1,2})/);if(a)d=a[3]+'-'+a[2].padStart(2,'0')+'-'+a[1].padStart(2,'0');if(b)d=b[1]+'-'+b[2].padStart(2,'0')+'-'+b[3].padStart(2,'0');if(d)map[p.padStart(5,'0')]=d;});return map;}

document.getElementById('addPdv').addEventListener('input',previewAdd);
document.getElementById('prev').addEventListener('change',e=>{const f=e.target.files[0];if(!f){PREV={};return;}const r=new FileReader();r.onload=()=>{PREV=parsePrev(r.result);alert('Storico caricato: '+Object.keys(PREV).length+' PV con data trovata');};r.readAsText(f);});
(function init(){rebuildAgentCanon();const now=new Date();document.getElementById('month').value=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');renderAll();loadRemote(true);})();
</script>
</body>
</html>'''


def safe_json_for_html(data):
    return json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")


def main():
    data = build_planning_data()
    data["generated_at"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    html = HTML.replace("__DATA_JSON__", safe_json_for_html(data)).replace("__API_URL__", API_URL)
    DOCS_DIR.mkdir(exist_ok=True)
    (DOCS_DIR / "planning.html").write_text(html, encoding="utf-8")
    print("Planning standalone creato:", data.get("summary", {}))


if __name__ == "__main__":
    main()
