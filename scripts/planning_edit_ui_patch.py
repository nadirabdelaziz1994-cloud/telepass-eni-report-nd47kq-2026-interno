from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PLANNING_PATCH = r'''
<script>
(function(){
  function modExtra(m){try{const n=JSON.parse(m.note||'{}');return n&&typeof n==='object'?n:{}}catch(e){return {}}}
  window.applyRemote = applyRemote = function(){
    const by={};basePoints().forEach(p=>by[p.pdv]=Object.assign({},p));
    (REMOTE||[]).forEach(m=>{
      const pnum=normPdv(m.pdv);if(!pnum)return;
      const action=String(m.action||'').toUpperCase();
      if(action==='ESCLUDI'){delete by[pnum];return;}
      if(action==='AGGIUNGI'){
        const extra=modExtra(m),cat=catalogPoint(pnum)||{pdv:pnum};
        const p=Object.assign({},cat);p.pdv=pnum;
        p.agent=m.agente||p.agent||'';
        p.city=m.citta||extra.citta||p.city||'';
        p.address=m.indirizzo||extra.indirizzo||p.address||'';
        p.region=m.regione||extra.regione||p.region||'';
        p.rzv=m.rzv||extra.rzv||p.rzv||'';
        p.cr=m.cr||extra.cr||p.cr||'';
        p.focal=m.focal||extra.focal||p.focal||'';
        p.lat=toNum(m.latitudine)!=null?toNum(m.latitudine):(toNum(extra.latitudine)!=null?toNum(extra.latitudine):p.lat);
        p.lng=toNum(m.longitudine)!=null?toNum(m.longitudine):(toNum(extra.longitudine)!=null?toNum(extra.longitudine):p.lng);
        const tipo=norm(m.tipo||extra.tipo||'');
        if(tipo.includes('grab'))p.is_grab=true;
        if(tipo.includes('telepass')||tipo.includes('tpoint')||tipo==='tp'||tipo.includes('point'))p.is_tp=true;
        if(tipo.includes('solo grab')){p.is_grab=true;p.is_tp=false;}
        if(!p.is_grab&&!p.is_tp)p.is_tp=true;
        p.source='Cloudflare definitivo';
        if(p.lat!=null&&p.lng!=null)by[pnum]=p;
      }
    });
    return Object.values(by).map(p=>Object.assign({},p,{agent_display:visibleAgentName(p.agent)}));
  };
  function savePlanState(){try{if(PLAN&&PLAN.length){sessionStorage.setItem('planningCurrent',JSON.stringify({plan:PLAN,agent:document.getElementById('agent')?.value||'',month:document.getElementById('month')?.value||'',start:document.getElementById('start')?.value||'',grab:document.getElementById('grab')?.value||'',savedAt:Date.now()}));}}catch(e){}}
  window.openPlanningEditor = openPlanningEditor = function(){savePlanState();if(!PLAN||!PLAN.length){alert('Prima crea il planning');return;}location.href='./planning-edit.html?v=edit1';};
  function restoreEdited(){try{const raw=sessionStorage.getItem('planningEdited');if(!raw)return;const data=JSON.parse(raw);if(data&&Array.isArray(data.plan)&&data.plan.length){PLAN=data.plan;sessionStorage.removeItem('planningEdited');renderTable(0);setTimeout(()=>document.getElementById('result')?.scrollIntoView({behavior:'smooth'}),150);}}catch(e){}}
  const oldRender=window.renderTable;
  window.renderTable = renderTable = function(workdayCount){
    oldRender(workdayCount);
    savePlanState();
    const card=document.querySelector('#result .card');
    if(card&&!document.getElementById('editPlanningBtn')){
      card.insertAdjacentHTML('afterbegin','<div class="edit-planning-top"><button id="editPlanningBtn" class="btn" onclick="openPlanningEditor()">Modifica planning</button></div>');
    }
  };
  window.addEventListener('pageshow',restoreEdited);
  window.addEventListener('load',restoreEdited);
})();
</script>
<style>
.edit-planning-top{display:flex;justify-content:flex-end;margin-bottom:12px}.edit-planning-top .btn{width:100%;font-size:15px;padding:12px}@media(min-width:720px){.edit-planning-top .btn{width:auto}}
</style>
'''

EDIT_HTML = r'''<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Modifica planning</title>
<style>
:root{--blue:#0b2d5c;--accent:#0d6efd;--bg:#f5f7fb;--card:#fff;--line:#d9e2ef;--muted:#64748b;--text:#102033;--bad:#991b1b;--good:#166534;--grab:#7c3aed;--dual:#eab308}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font-family:Arial,Helvetica,sans-serif}header{background:var(--blue);color:white;padding:12px 16px;position:sticky;top:0;z-index:5;box-shadow:0 2px 10px #0002}header .row{display:flex;gap:10px;align-items:center;justify-content:space-between}header h1{font-size:18px;margin:0}.wrap{max-width:980px;margin:0 auto;padding:12px}.btn{border:0;border-radius:11px;padding:10px 12px;background:var(--accent);color:white;font-weight:900;cursor:pointer}.btn.light{background:#edf4ff;color:var(--blue);border:1px solid #bfd4ef}.btn.bad{background:#fee2e2;color:var(--bad);border:1px solid #fecaca}.card{background:white;border:1px solid var(--line);border-radius:16px;padding:12px;margin:10px 0;box-shadow:0 2px 12px #12376410}.muted{color:var(--muted);font-size:13px;line-height:1.35}.actions{display:flex;gap:8px;flex-wrap:wrap}.actions .btn{flex:1}.edit-list{display:grid;gap:8px}.edit-row{display:grid;grid-template-columns:42px 38px 135px 1fr 42px;gap:6px;align-items:center;background:white;border:1px solid var(--line);border-left:6px solid var(--accent);border-radius:12px;padding:8px;touch-action:none}.edit-row.grab{border-left-color:var(--grab);background:#fbf7ff}.edit-row.dual{border-left-color:var(--dual);background:#fffaf0}.drag{height:38px;border-radius:10px;border:1px solid var(--line);display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:900;color:var(--blue);background:#f8fbff;cursor:grab}.idx{font-weight:900;color:var(--muted);font-size:13px}.date{width:100%;padding:8px;border:1px solid var(--line);border-radius:9px;background:white;font-size:13px}.main b{display:block;font-size:15px}.main span{display:block;font-size:12px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.x{width:38px;height:38px}.dragging{opacity:.45}.drop-here{outline:3px solid #0d6efd55}@media(max-width:560px){.edit-row{grid-template-columns:38px 28px 116px 1fr 38px;padding:7px;gap:5px}.main b{font-size:14px}.main span{font-size:11px}.date{font-size:12px;padding:7px}.wrap{padding:8px}}
</style>
</head>
<body>
<header><div class="row"><h1>Modifica planning</h1><button class="btn light" onclick="goBack()">Annulla</button></div></header>
<main class="wrap">
<section class="card">
  <div class="muted">Trascina le righe con ↕ per cambiare ordine, usa la X per eliminare, oppure cambia la data dalla riga.</div>
  <div class="actions" style="margin-top:10px"><button class="btn" onclick="saveAndBack()">Salva modifiche</button><button class="btn light" onclick="resetDatesOrder()">Ricalcola date in ordine</button></div>
</section>
<section id="list" class="edit-list"></section>
</main>
<script>
let state={},PLAN=[];let dragIndex=null,touchIndex=null;
function esc(v){return String(v == null ? '' : v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function tipo(p){if(p.is_grab&&p.is_tp)return 'Doppio';if(p.is_grab)return 'Grab';return 'TPoint';}
function cls(p){if(p.is_grab&&p.is_tp)return 'dual';if(p.is_grab)return 'grab';return 'tp';}
function dateLabel(iso){try{return new Date(iso+'T00:00:00').toLocaleDateString('it-IT',{weekday:'short',day:'2-digit',month:'2-digit',year:'numeric'});}catch(e){return iso;}}
function dateOnly(iso){try{return new Date(iso+'T00:00:00').toLocaleDateString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric'});}catch(e){return iso;}}
function load(){try{state=JSON.parse(sessionStorage.getItem('planningCurrent')||'{}');PLAN=Array.isArray(state.plan)?state.plan:[];}catch(e){PLAN=[];}render();}
function render(){const box=document.getElementById('list');if(!PLAN.length){box.innerHTML='<div class="card muted">Nessun planning trovato. Torna alla pagina planning e crealo prima.</div>';return;}box.innerHTML=PLAN.map((p,i)=>'<div class="edit-row '+cls(p)+'" data-i="'+i+'" draggable="true" ondragstart="dragStart(event,'+i+')" ondragover="dragOver(event)" ondrop="dropRow(event,'+i+')"><div class="drag" ontouchstart="touchStart(event,'+i+')" ontouchmove="touchMove(event)" ontouchend="touchEnd(event)">↕</div><div class="idx">'+(i+1)+'</div><input class="date" type="date" value="'+esc(p.date||'')+'" onchange="changeDate('+i+',this.value)"><div class="main"><b>'+esc(p.pdv)+' · '+esc(tipo(p))+'</b><span>'+esc(p.city||'')+' · '+esc(p.address||'')+'</span></div><button class="btn bad x" onclick="del('+i+')">×</button></div>').join('');}
function move(from,to){if(from===to||from<0||to<0||from>=PLAN.length||to>=PLAN.length)return;const [r]=PLAN.splice(from,1);PLAN.splice(to,0,r);render();}
function dragStart(e,i){dragIndex=i;e.dataTransfer.effectAllowed='move';}
function dragOver(e){e.preventDefault();}
function dropRow(e,i){e.preventDefault();move(dragIndex,i);dragIndex=null;}
function touchStart(e,i){touchIndex=i;e.currentTarget.closest('.edit-row').classList.add('dragging');e.preventDefault();}
function touchMove(e){if(touchIndex==null)return;const t=e.touches[0];const el=document.elementFromPoint(t.clientX,t.clientY)?.closest('.edit-row');if(!el)return;const j=Number(el.dataset.i);if(Number.isFinite(j)&&j!==touchIndex){move(touchIndex,j);touchIndex=j;}e.preventDefault();}
function touchEnd(e){touchIndex=null;document.querySelectorAll('.dragging').forEach(x=>x.classList.remove('dragging'));e.preventDefault();}
function changeDate(i,v){if(!PLAN[i]||!v)return;PLAN[i].date=v;PLAN[i].dateLabel=dateLabel(v);PLAN[i].dateOnly=dateOnly(v);}
function del(i){PLAN.splice(i,1);render();}
function resetDatesOrder(){alert('Per ora questa opzione la lasciamo ferma: il ricalcolo automatico lo facciamo nella pagina principale con Crea planning. Qui modifichi manualmente ordine e date.');}
function saveAndBack(){sessionStorage.setItem('planningEdited',JSON.stringify({plan:PLAN,savedAt:Date.now()}));location.href='./planning.html?v=edit-return1';}
function goBack(){location.href='./planning.html?v=edit-cancel1';}
load();
</script>
</body>
</html>'''

PDV_PATCH_JS = r'''
<script>
(function(){
  function modExtra(m){try{const n=JSON.parse(m.note||'{}');return n&&typeof n==='object'?n:{}}catch(e){return {}}}
  window.applyRemote = applyRemote = function(){const by={};basePoints().forEach(p=>by[p.pdv]=Object.assign({},p));(REMOTE||[]).forEach(m=>{const pnum=normPdv(m.pdv);if(!pnum)return;const action=String(m.action||'').toUpperCase();if(action==='ESCLUDI'){delete by[pnum];return;}if(action==='AGGIUNGI'){const extra=modExtra(m),cat=catalogPoint(pnum)||{pdv:pnum};const p=Object.assign({},cat);p.pdv=pnum;p.agent=m.agente||p.agent||'';p.city=m.citta||extra.citta||p.city||'';p.address=m.indirizzo||extra.indirizzo||p.address||'';p.region=m.regione||extra.regione||p.region||'';p.rzv=m.rzv||extra.rzv||p.rzv||'';p.cr=m.cr||extra.cr||p.cr||'';p.focal=m.focal||extra.focal||p.focal||'';p.lat=toNum(m.latitudine)!=null?toNum(m.latitudine):(toNum(extra.latitudine)!=null?toNum(extra.latitudine):p.lat);p.lng=toNum(m.longitudine)!=null?toNum(m.longitudine):(toNum(extra.longitudine)!=null?toNum(extra.longitudine):p.lng);const tipo=norm(m.tipo||extra.tipo||'');if(tipo.includes('grab'))p.is_grab=true;if(tipo.includes('telepass')||tipo.includes('tpoint')||tipo==='tp'||tipo.includes('point'))p.is_tp=true;if(tipo.includes('solo grab')){p.is_grab=true;p.is_tp=false;}if(!p.is_grab&&!p.is_tp)p.is_tp=true;p.source='Cloudflare definitivo';if(p.lat!=null&&p.lng!=null)by[pnum]=p;}});return Object.values(by).map(p=>Object.assign({},p,{agent_display:visibleAgentName(p.agent)}));};
  function setVal(id,v){const el=document.getElementById(id);if(el&&v!=null&&v!=='')el.value=v;}
  window.autoFillFromCatalog = autoFillFromCatalog = function(){previewCatalog();const p=catalogPoint(document.getElementById('pdv')?.value);if(!p)return;setVal('city',p.city||'');setVal('address',p.address||'');setVal('region',p.region||'');setVal('lat',p.lat??'');setVal('lng',p.lng??'');setVal('cr',p.cr||'');setVal('rzv',p.rzv||'');setVal('focal',p.focal||'');};
  window.fillFromCatalog = fillFromCatalog = function(){const p=catalogPoint(document.getElementById('pdv').value);if(!p){alert('PV non trovato in anagrafica');return;}autoFillFromCatalog();};
  window.addPoint = addPoint = async function(){const agent=currentAgent();if(!agent){alert('Seleziona agente');return;}const pdv=normPdv(document.getElementById('pdv').value);if(!pdv){alert('Inserisci codice PV');return;}let lat=toNum(document.getElementById('lat').value),lng=toNum(document.getElementById('lng').value);if(lat==null||lng==null){const p=catalogPoint(pdv);if(p){lat=p.lat;lng=p.lng;}}if(lat==null||lng==null){alert('Servono latitudine e longitudine');return;}const tv=document.getElementById('tipo').value;const tipo=tv==='grab'?'Solo Grab & Go':(tv==='dual'?'TPoint + Grab & Go':'Telepass Point');const extra={tipo,citta:document.getElementById('city').value,indirizzo:document.getElementById('address').value,regione:document.getElementById('region').value,latitudine:lat,longitudine:lng,cr:document.getElementById('cr').value,rzv:document.getElementById('rzv').value,focal:document.getElementById('focal').value};await saveRemote({action:'AGGIUNGI',pdv,agente:agent,tipo,citta:extra.citta,indirizzo:extra.indirizzo,latitudine:lat,longitudine:lng,note:JSON.stringify(extra)});alert('PV salvato');showTab('list');};
})();
</script>
'''


def patch_planning():
    path = DOCS_DIR / 'planning.html'
    html = path.read_text(encoding='utf-8')
    if 'openPlanningEditor' not in html:
        html = html.replace('</body>', PLANNING_PATCH + '\n</body>', 1)
    path.write_text(html, encoding='utf-8')


def patch_manage():
    path = DOCS_DIR / 'pdv-manage.html'
    if not path.exists():
        return
    html = path.read_text(encoding='utf-8')
    html = html.replace('<input id="pdv" placeholder="Es. 01480" oninput="previewCatalog()">','<input id="pdv" placeholder="Es. 01480" oninput="autoFillFromCatalog()">')
    html = html.replace('<option value="tp">Telepass Point</option><option value="grab">Solo Grab & Go</option>', '<option value="tp">Telepass Point</option><option value="dual">TPoint + Grab & Go</option><option value="grab">Solo Grab & Go</option>')
    marker = '<div><label>Regione</label><input id="region" placeholder="Regione"></div>'
    extra = marker + '<div><label>CR</label><input id="cr" placeholder="CR"></div><div><label>RZV</label><input id="rzv" placeholder="RZV"></div><div><label>Focal Point</label><input id="focal" placeholder="Focal Point"></div>'
    if marker in html and 'id="rzv"' not in html:
        html = html.replace(marker, extra, 1)
    if 'autoFillFromCatalog = autoFillFromCatalog' not in html:
        html = html.replace('</body>', PDV_PATCH_JS + '\n</body>', 1)
    path.write_text(html, encoding='utf-8')


def main():
    patch_planning()
    patch_manage()
    (DOCS_DIR / 'planning-edit.html').write_text(EDIT_HTML, encoding='utf-8')
    print('Patch modifica planning e gestione PV avanzata applicata')


if __name__ == '__main__':
    main()
