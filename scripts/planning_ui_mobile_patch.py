from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PLAN_CARD_JS = r'''
<script>
(function(){
  window.openPdvManager = openPdvManager = function(){
    const a=document.getElementById('agent') ? document.getElementById('agent').value : '';
    const url='./pdv-manage.html' + (a ? ('?agent='+encodeURIComponent(a)) : '');
    window.location.href=url;
  };
  window.renderTable = renderTable = function(workdayCount){
    const days=[...new Set(PLAN.map(p=>p.date))].length;
    const grabOnly=PLAN.filter(p=>typeof isGrabOnly==='function'&&isGrabOnly(p)).length;
    const dual=PLAN.filter(p=>typeof isDual==='function'&&isDual(p)).length;
    const cards=PLAN.map((p,i)=>{
      const k=(typeof pvKind==='function')?pvKind(p):(p&&p.is_grab?'grab':'tp');
      const cls=k==='dual'?'dual':(k==='grab'?'grab-only':'tp');
      const label=k==='dual'?'Grab & Go + TPoint':(k==='grab'?'Solo Grab & Go':'Solo TPoint');
      const pill=k==='dual'?'warn':(k==='grab'?'grab':'');
      return '<article class="plan-card '+cls+'">'
        +'<div class="plan-head"><div><div class="plan-pdv">PV '+esc(p.pdv)+'</div><div class="plan-date">'+esc(p.dateLabel||'')+(p.time?' · '+esc(p.time):'')+'</div></div><span class="pill '+pill+'">'+esc(label)+'</span></div>'
        +'<div class="plan-place"><b>'+esc(p.city||'')+'</b><span>'+esc(p.address||'')+'</span></div>'
        +'<div class="plan-meta"><div><small>Regione</small><b>'+esc(p.region||'')+'</b></div><div><small>Provincia</small><b>'+esc(p.province||'')+'</b></div><div><small>Auto</small><b>'+fmt(p.travel_km||0)+' km</b></div><div><small>Carico</small><b>'+esc(p.day_load||'')+'</b></div></div>'
        +'<div class="plan-actions"><button class="btn light small" onclick="moveRow('+i+',-1)">↑ Su</button><button class="btn light small" onclick="moveRow('+i+',1)">↓ Giù</button><button class="btn bad small" onclick="delRow('+i+')">Elimina</button></div>'
        +'</article>';
    }).join('');
    document.getElementById('result').innerHTML='<section class="metric result-metrics"><div class="box"><h4>PV</h4><div class="big">'+fmt(PLAN.length)+'</div></div><div class="box"><h4>Giorni usati</h4><div class="big">'+fmt(days)+'</div><div class="muted">Disponibili: '+fmt(workdayCount||0)+'</div></div><div class="box"><h4>Solo Grab</h4><div class="big">'+fmt(grabOnly)+'</div></div><div class="box"><h4>Doppi</h4><div class="big">'+fmt(dual)+'</div></div></section><section class="card"><div class="plan-legend"><span class="legend-dot normal"></span> TPoint <span class="legend-dot grab"></span> Grab <span class="legend-dot dual"></span> Doppio</div><div class="plan-list">'+cards+'</div></section>';
  };
})();
</script>
'''

PLAN_CSS = r'''
<style>
  .hidden-ui{display:none!important}
  .manage-top{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
  .manage-top .btn{font-size:15px;padding:12px 14px;width:100%}
  #metricBox{display:none!important}
  .cloud-hidden{display:none!important}
  .result-metrics{margin-top:12px}.result-metrics .box{padding:10px}.result-metrics .big{font-size:22px}
  .plan-legend{display:flex;gap:10px;align-items:center;flex-wrap:wrap;color:var(--muted);font-weight:800;margin-bottom:10px;font-size:13px}
  .legend-dot{display:inline-block;width:12px;height:12px;border-radius:999px;background:#0d6efd}.legend-dot.grab{background:#7c3aed}.legend-dot.dual{background:#eab308}
  .plan-list{display:grid;gap:10px}
  .plan-card{background:white;border:1px solid var(--line);border-radius:16px;padding:12px;box-shadow:0 2px 8px #1237640d;border-left:6px solid var(--accent)}
  .plan-card.grab-only{border-left-color:#7c3aed;background:#fbf7ff}.plan-card.dual{border-left-color:#eab308;background:#fffaf0}
  .plan-head{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}.plan-pdv{font-size:18px;font-weight:900;color:var(--blue)}.plan-date{font-size:13px;color:var(--muted);font-weight:800;margin-top:2px}
  .plan-place{margin-top:9px}.plan-place b{display:block;font-size:16px}.plan-place span{display:block;color:var(--muted);font-size:13px;margin-top:2px;line-height:1.25}
  .plan-meta{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;margin-top:10px}.plan-meta div{background:#f8fbff;border:1px solid var(--line);border-radius:12px;padding:8px}.plan-meta small{display:block;color:var(--muted);font-size:11px;font-weight:800}.plan-meta b{font-size:13px;color:var(--text)}
  .plan-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.plan-actions .btn{padding:8px 10px;border-radius:10px}
  @media(min-width:720px){.manage-top .btn{width:auto}.plan-meta{grid-template-columns:repeat(4,minmax(0,1fr))}.plan-list{grid-template-columns:repeat(2,minmax(0,1fr))}}
</style>
'''

MANAGE_HTML = r'''<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Gestione PV manuale</title>
<style>
:root{--blue:#0b2d5c;--accent:#0d6efd;--bg:#f5f7fb;--card:#fff;--line:#d9e2ef;--muted:#64748b;--text:#102033;--good:#166534;--bad:#991b1b;--grab:#7c3aed}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font-family:Arial,Helvetica,sans-serif}header{background:var(--blue);color:white;padding:12px 16px;position:sticky;top:0;z-index:2;box-shadow:0 2px 10px #0002}header .row{display:flex;align-items:center;justify-content:space-between;gap:10px}header h1{font-size:18px;margin:0}a.top{color:white;text-decoration:none;border:1px solid #ffffff55;border-radius:999px;padding:7px 10px;font-weight:800;font-size:12px}.wrap{max-width:980px;margin:0 auto;padding:14px}.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:14px;margin:12px 0;box-shadow:0 2px 12px #12376410}.title{font-size:20px;font-weight:900;color:var(--blue);margin:0 0 10px}.muted{color:var(--muted);font-size:13px;line-height:1.35}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px}.form label{font-size:12px;color:var(--muted);font-weight:800;display:block;margin-bottom:4px}.form input,.form select{width:100%;padding:12px;border:1px solid var(--line);border-radius:12px;background:white;font-size:15px}.btn{border:0;border-radius:12px;padding:11px 13px;background:var(--accent);color:white;font-weight:900;cursor:pointer}.btn.light{background:#edf4ff;color:var(--blue);border:1px solid #bfd4ef}.btn.bad{background:#fee2e2;color:var(--bad);border:1px solid #fecaca}.btn.good{background:#dcfce7;color:var(--good);border:1px solid #bbf7d0}.actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.tabs{display:flex;gap:8px;flex-wrap:wrap}.tabs .btn{flex:1;min-width:150px}.pv-list{display:grid;gap:10px;margin-top:12px}.pv-card{background:white;border:1px solid var(--line);border-left:6px solid var(--accent);border-radius:15px;padding:12px}.pv-card.grab{border-left-color:var(--grab);background:#fbf7ff}.pv-top{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}.pv-code{font-size:18px;font-weight:900;color:var(--blue)}.pill{display:inline-flex;border-radius:999px;padding:4px 8px;font-size:12px;font-weight:900;background:#eef4fb;color:var(--blue)}.pill.grab{background:#f3e8ff;color:var(--grab)}.addr{margin-top:8px}.addr b{display:block}.addr span{display:block;color:var(--muted);font-size:13px;margin-top:2px}.empty{padding:20px;border:1px dashed var(--line);border-radius:14px;color:var(--muted);background:#fff}.hide{display:none!important}@media(min-width:720px){.pv-list{grid-template-columns:repeat(2,minmax(0,1fr))}}
</style>
</head>
<body>
<script id="planning-data" type="application/json">__DATA_JSON__</script>
<header><div class="row"><h1>Gestione PV</h1><a class="top" href="./planning.html?v=ui1">← Planning</a></div></header>
<main class="wrap">
<section class="card">
  <div class="title">Aggiungi / rimuovi PV</div>
  <div class="grid form">
    <div><label>Agente</label><select id="agent"></select></div>
  </div>
  <div class="tabs" style="margin-top:12px">
    <button class="btn light" onclick="showTab('list')">Mostra lista PV</button>
    <button class="btn" onclick="showTab('add')">Aggiungi punto vendita</button>
  </div>
  <div id="apiStatus" class="muted" style="margin-top:10px">Controllo Cloudflare...</div>
</section>
<section id="listSec" class="card">
  <div class="title">Lista PV agente</div>
  <div class="muted">Mostra codice PV, via, comune e regione. Il tasto elimina crea una esclusione definitiva su Cloudflare.</div>
  <div id="pvList" class="pv-list"></div>
</section>
<section id="addSec" class="card hide">
  <div class="title">Aggiungi punto vendita</div>
  <div class="grid form">
    <div><label>Codice PV</label><input id="pdv" placeholder="Es. 01480" oninput="previewCatalog()"></div>
    <div><label>Tipo</label><select id="tipo"><option value="tp">Telepass Point</option><option value="grab">Solo Grab & Go</option></select></div>
    <div><label>Città / Comune</label><input id="city" placeholder="Comune"></div>
    <div><label>Indirizzo / Via</label><input id="address" placeholder="Via"></div>
    <div><label>Regione</label><input id="region" placeholder="Regione"></div>
    <div><label>Latitudine</label><input id="lat" placeholder="45.4642"></div>
    <div><label>Longitudine</label><input id="lng" placeholder="9.1900"></div>
  </div>
  <div id="catalogPreview" class="muted" style="margin-top:8px"></div>
  <div class="actions"><button class="btn good" onclick="addPoint()">Salva punto vendita</button><button class="btn light" onclick="fillFromCatalog()">Compila da anagrafica</button></div>
</section>
</main>
<script>
const API_URL="__API_URL__";
const DATA=JSON.parse(document.getElementById('planning-data').textContent||'{}');
let REMOTE=[];let apiReady=false;let AGENT_CANON={};
function esc(v){return String(v == null ? '' : v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function fmt(v){return Number(v||0).toLocaleString('it-IT');}
function normPdv(v){const m=String(v||'').match(/\d+/);return m?m[0].padStart(5,'0'):'';}
function norm(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}
function titleCase(s){return String(s||'').toLowerCase().split(/\s+/).filter(Boolean).map(x=>x.charAt(0).toUpperCase()+x.slice(1)).join(' ');}
function toNum(v){const n=Number(String(v??'').replace(',','.'));return Number.isFinite(n)?n:null;}
function agentKey(s){return norm(s).split(' ').filter(x=>x&&x!=='myworldsrl'&&x!=='it').sort().join('|');}
function rawEmailToName(s){const local=String(s||'').split('@')[0].replace(/[._-]+/g,' ');return titleCase(local);}
function visibleAgentName(a){a=String(a||'').trim();if(!a)return '';if(a.includes('@')){const guessed=rawEmailToName(a);return AGENT_CANON[agentKey(guessed)]||guessed;}return a;}
function rebuildAgentCanon(){AGENT_CANON={};const raw=[...(DATA.points||[]),...(DATA.catalog||[])].map(p=>String(p.agent||'').trim()).filter(Boolean).filter(a=>!a.includes('@'));raw.forEach(a=>{const k=agentKey(a);if(k&&!AGENT_CANON[k])AGENT_CANON[k]=a;});}
function isGrab(p){return !!p.is_grab;}
function catalogPoint(pdv){const n=normPdv(pdv);return (DATA.catalog||[]).find(p=>p.pdv===n)||null;}
function basePoints(){return (DATA.points||[]).filter(p=>p&&p.pdv&&p.lat!=null&&p.lng!=null);}
function applyRemote(){const by={};basePoints().forEach(p=>by[p.pdv]=Object.assign({},p));(REMOTE||[]).forEach(m=>{const pnum=normPdv(m.pdv);if(!pnum)return;const action=String(m.action||'').toUpperCase();if(action==='ESCLUDI'){delete by[pnum];return;}if(action==='AGGIUNGI'){const cat=catalogPoint(pnum)||{pdv:pnum};const p=Object.assign({},cat);p.pdv=pnum;p.agent=m.agente||p.agent||'';p.city=m.citta||p.city||'';p.address=m.indirizzo||p.address||'';p.region=m.regione||p.region||'';p.lat=toNum(m.latitudine)!=null?toNum(m.latitudine):p.lat;p.lng=toNum(m.longitudine)!=null?toNum(m.longitudine):p.lng;const tipo=String(m.tipo||'').toLowerCase();p.is_grab=tipo.includes('grab')||p.is_grab;p.is_tp=!p.is_grab||p.is_tp;p.source='Cloudflare definitivo';if(p.lat!=null&&p.lng!=null)by[pnum]=p;}});return Object.values(by).map(p=>Object.assign({},p,{agent_display:visibleAgentName(p.agent)}));}
function agents(){return [...new Set(applyRemote().map(p=>p.agent_display).filter(Boolean))].filter(a=>!a.includes('@')).sort((a,b)=>a.localeCompare(b,'it'));}
async function api(path,opt){const res=await fetch(API_URL+path,opt||{});const data=await res.json().catch(()=>({ok:false,error:'Risposta non valida'}));if(!res.ok||data.ok===false)throw new Error(data.error||('Errore API '+res.status));return data;}
function getAdminKey(){let k=sessionStorage.getItem('planningAdminKey')||'';if(!k){k=prompt('Inserisci ADMIN_KEY Cloudflare');if(k)sessionStorage.setItem('planningAdminKey',k);}return k;}
async function loadRemote(){const box=document.getElementById('apiStatus');try{const data=await api('/modifiche');REMOTE=data.modifiche||[];apiReady=true;box.innerHTML='<b style="color:#166534">Cloudflare collegato</b> · modifiche: '+fmt(REMOTE.length);renderAll();}catch(e){apiReady=false;box.innerHTML='<b style="color:#991b1b">Cloudflare non raggiungibile</b>: '+esc(e.message);renderAll();}}
async function saveRemote(mod){const key=getAdminKey();if(!key)return;await api('/modifica',{method:'POST',headers:{'Content-Type':'application/json','X-Admin-Key':key},body:JSON.stringify(mod)});await loadRemote();}
function renderAgents(){const sel=document.getElementById('agent');const cur=sel.value||new URLSearchParams(location.search).get('agent')||'';sel.innerHTML='<option value="">Seleziona agente</option>'+agents().map(a=>'<option>'+esc(a)+'</option>').join('');if(cur)sel.value=cur;}
function currentAgent(){return document.getElementById('agent').value;}
function renderList(){const a=currentAgent();const box=document.getElementById('pvList');if(!a){box.innerHTML='<div class="empty">Seleziona prima un agente.</div>';return;}const pts=applyRemote().filter(p=>p.agent_display===a).sort((x,y)=>String(x.region||'').localeCompare(String(y.region||''),'it')||String(x.city||'').localeCompare(String(y.city||''),'it')||String(x.pdv||'').localeCompare(String(y.pdv||''),'it'));if(!pts.length){box.innerHTML='<div class="empty">Nessun PV trovato per questo agente.</div>';return;}box.innerHTML=pts.map(p=>'<article class="pv-card '+(isGrab(p)?'grab':'')+'"><div class="pv-top"><div><div class="pv-code">PV '+esc(p.pdv)+'</div><span class="pill '+(isGrab(p)?'grab':'')+'">'+(isGrab(p)?'Grab & Go':'TPoint')+'</span></div><button class="btn bad" onclick="removePdv(\''+esc(p.pdv)+'\')">Elimina</button></div><div class="addr"><b>'+esc(p.address||'Via non indicata')+'</b><span>'+esc(p.city||'Comune non indicato')+' · '+esc(p.region||'Regione non indicata')+'</span></div></article>').join('');}
function renderAll(){renderAgents();renderList();previewCatalog();}
function showTab(t){document.getElementById('listSec').classList.toggle('hide',t!=='list');document.getElementById('addSec').classList.toggle('hide',t!=='add');if(t==='list')renderList();}
async function removePdv(pdv){if(!confirm('Escludere definitivamente il PV '+pdv+'?'))return;await saveRemote({action:'ESCLUDI',pdv,note:'Escluso da pagina gestione PV'});}
function fillFromCatalog(){const p=catalogPoint(document.getElementById('pdv').value);if(!p){alert('PV non trovato in anagrafica');return;}document.getElementById('city').value=p.city||'';document.getElementById('address').value=p.address||'';document.getElementById('region').value=p.region||'';document.getElementById('lat').value=p.lat??'';document.getElementById('lng').value=p.lng??'';}
function previewCatalog(){const box=document.getElementById('catalogPreview');if(!box)return;const p=catalogPoint(document.getElementById('pdv')?.value);box.innerHTML=p?'Trovato in anagrafica: <b>'+esc(p.pdv)+'</b> · '+esc(p.city||'')+' · '+esc(p.address||''):'Inserisci un codice PV. Se presente in anagrafica puoi compilare i dati automaticamente.';}
async function addPoint(){const agent=currentAgent();if(!agent){alert('Seleziona agente');return;}const pdv=normPdv(document.getElementById('pdv').value);if(!pdv){alert('Inserisci codice PV');return;}let lat=toNum(document.getElementById('lat').value),lng=toNum(document.getElementById('lng').value);if(lat==null||lng==null){const p=catalogPoint(pdv);if(p){lat=p.lat;lng=p.lng;}}if(lat==null||lng==null){alert('Servono latitudine e longitudine');return;}const tipo=document.getElementById('tipo').value==='grab'?'Grab & Go':'Telepass Point';await saveRemote({action:'AGGIUNGI',pdv,agente:agent,tipo,citta:document.getElementById('city').value,indirizzo:document.getElementById('address').value,regione:document.getElementById('region').value,latitudine:lat,longitudine:lng,note:'Aggiunto da pagina gestione PV'});alert('PV salvato');showTab('list');}
document.getElementById('agent').addEventListener('change',renderList);rebuildAgentCanon();loadRemote();
</script>
</body>
</html>'''


def between(text, a, b):
    s = text.find(a)
    if s == -1:
        return None
    e = text.find(b, s)
    if e == -1:
        return None
    return text[s:e+len(b)]


def replace_first_card(html):
    old = between(html, '    <section class="card">\n      <div class="title">Planning automatico definitivo</div>', '    </section>')
    if old:
        new = '    <section class="card manage-top"><button class="btn" onclick="openPdvManager()">Aggiungi / rimuovi PV manualmente</button><div id="sourceBox" class="hidden-ui"></div></section>'
        html = html.replace(old, new, 1)
    html = html.replace('<section class="metric" id="metricBox"></section>', '<section class="metric hidden-ui" id="metricBox"></section>')
    return html


def remove_cloud_card(html):
    start = '    <section class="card">\n      <div class="title">Cloudflare definitivo</div>'
    end = '    <section id="result"></section>'
    s = html.find(start)
    e = html.find(end, s)
    if s != -1 and e != -1:
        hidden = '    <div class="hidden-ui"><div id="apiStatus"></div><input id="addPdv"><div id="addPreview"></div></div>\n'
        html = html[:s] + hidden + html[e:]
    return html


def main():
    path = DOCS_DIR / 'planning.html'
    if not path.exists():
        print('planning.html non trovato, UI patch saltata')
        return
    html = path.read_text(encoding='utf-8')
    data_start = '<script id="planning-data" type="application/json">'
    data_end = '</script>'
    s = html.find(data_start)
    e = html.find(data_end, s)
    data_json = '{}'
    if s != -1 and e != -1:
        data_json = html[s+len(data_start):e]

    if 'openPdvManager' not in html:
        html = html.replace('</head>', PLAN_CSS + '\n</head>', 1)
        html = replace_first_card(html)
        html = remove_cloud_card(html)
        html = html.replace('        <button class="btn light" onclick="loadRemote(false)">Ricarica modifiche definitive</button>\n', '')
        html = html.replace('      <div class="muted" style="margin-top:8px">Export Excel con le stesse colonne del file planning originale. I Grab & Go sono evidenziati, senza colonne extra.</div>\n', '')
        html = html.replace('</body>', PLAN_CARD_JS + '\n</body>', 1)
    else:
        html = replace_first_card(html)
        html = remove_cloud_card(html)
    path.write_text(html, encoding='utf-8')

    manage = MANAGE_HTML.replace('__DATA_JSON__', data_json).replace('__API_URL__', 'https://telepass-planning-api.nadirabdelaziz1994.workers.dev')
    (DOCS_DIR / 'pdv-manage.html').write_text(manage, encoding='utf-8')
    print('UI mobile e pagina gestione PV applicate')


if __name__ == '__main__':
    main()
