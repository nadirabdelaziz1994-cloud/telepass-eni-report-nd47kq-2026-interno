from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"
API_URL = "https://telepass-planning-api.nadirabdelaziz1994.workers.dev"

CLOUD_JS = f'''
<script>
window.PLANNING_API_URL = "{API_URL}";
window.PLAN_REMOTE_MODS = [];
window.PLAN_REMOTE_READY = false;

async function planApiFetch(path, options) {{
  const res = await fetch(window.PLANNING_API_URL + path, options || {{}});
  const data = await res.json().catch(() => ({{ok:false,error:'Risposta non valida'} }));
  if(!res.ok || data.ok === false) throw new Error(data.error || ('Errore API ' + res.status));
  return data;
}}

async function planApiLoadRemote(silent) {{
  try {{
    const data = await planApiFetch('/modifiche');
    window.PLAN_REMOTE_MODS = data.modifiche || [];
    window.PLAN_REMOTE_READY = true;
    if(!silent) alert('Modifiche definitive caricate: ' + window.PLAN_REMOTE_MODS.length);
    return window.PLAN_REMOTE_MODS;
  }} catch(e) {{
    window.PLAN_REMOTE_READY = false;
    if(!silent) alert('Errore caricamento modifiche Cloudflare: ' + e.message);
    return [];
  }}
}}

function planGetAdminKey() {{
  let key = sessionStorage.getItem('planningAdminKey') || '';
  if(!key) {{
    key = prompt('Inserisci ADMIN_KEY Cloudflare. Non viene salvata nel sito, solo in questa sessione.');
    if(key) sessionStorage.setItem('planningAdminKey', key);
  }}
  return key || '';
}}

function planPdvNorm(value) {{
  const m = String(value || '').match(/\d+/);
  return m ? m[0].padStart(5,'0') : '';
}}

function planCatalogPoint(pdv) {{
  const n = planPdvNorm(pdv);
  return ((APP.planning_data && APP.planning_data.catalog) || []).find(p => p.pdv === n) || null;
}}

function planApplyRemote(points) {{
  const by = {{}};
  (points || []).forEach(p => {{ if(p && p.pdv) by[p.pdv] = Object.assign({{}}, p); }});
  (window.PLAN_REMOTE_MODS || []).forEach(m => {{
    const pdv = planPdvNorm(m.pdv);
    if(!pdv) return;
    const action = String(m.action || '').toUpperCase();
    if(action === 'ESCLUDI') {{
      delete by[pdv];
      return;
    }}
    if(action === 'AGGIUNGI') {{
      const cat = planCatalogPoint(pdv) || {{pdv}};
      const p = Object.assign({{}}, cat);
      p.pdv = pdv;
      p.agent = m.agente || p.agent || '';
      p.city = m.citta || p.city || '';
      p.address = m.indirizzo || p.address || '';
      p.lat = m.latitudine !== null && m.latitudine !== undefined && m.latitudine !== '' ? Number(m.latitudine) : p.lat;
      p.lng = m.longitudine !== null && m.longitudine !== undefined && m.longitudine !== '' ? Number(m.longitudine) : p.lng;
      const tipo = String(m.tipo || '').toLowerCase();
      p.is_grab = tipo.includes('grab') || p.is_grab;
      p.is_tp = !p.is_grab || p.is_tp;
      p.source = 'Cloudflare definitivo';
      if(p.lat != null && p.lng != null) by[pdv] = p;
    }}
  }});
  return Object.values(by);
}}

const _planBasePointsFn = window.planPoints;
window.planPoints = function() {{
  const base = typeof planBasePoints === 'function' ? planBasePoints() : ((_planBasePointsFn && _planBasePointsFn()) || []);
  const local = typeof planCustomPoints === 'function' ? planCustomPoints() : [];
  const removed = typeof planRemovedMap === 'function' ? planRemovedMap() : {{}};
  const merged = {{}};
  [...base, ...local].forEach(p => {{ if(p && p.pdv && !removed[p.pdv]) merged[p.pdv] = Object.assign(merged[p.pdv] || {{}}, p); }});
  return planApplyRemote(Object.values(merged));
}};

async function planApiSaveModifica(mod) {{
  const key = planGetAdminKey();
  if(!key) return false;
  await planApiFetch('/modifica', {{
    method: 'POST',
    headers: {{'Content-Type':'application/json','X-Admin-Key':key}},
    body: JSON.stringify(mod)
  }});
  await planApiLoadRemote(true);
  return true;
}}

window.addPlanningPdvFromCatalog = async function() {{
  const pv = document.getElementById('planAddPdv')?.value || '';
  const tipo = document.getElementById('planAddType')?.value === 'grab' ? 'Grab & Go' : 'Telepass Point';
  const agent = document.getElementById('planAgent')?.value || '';
  const p = planCatalogPoint(pv);
  if(!p) {{ alert('PV non trovato in anagrafica. Usa aggiunta manuale.'); return; }}
  if(p.lat == null || p.lng == null) {{ alert('PV trovato ma senza coordinate. Prima correggi anagrafica o usa aggiunta manuale con coordinate.'); return; }}
  try {{
    await planApiSaveModifica({{action:'AGGIUNGI', pdv:p.pdv, agente:agent || p.agent || '', tipo, citta:p.city || '', indirizzo:p.address || '', latitudine:p.lat, longitudine:p.lng, note:'Aggiunto da sito'}});
    alert('PV aggiunto definitivamente.');
    renderPlanningAuto();
  }} catch(e) {{ alert('Errore salvataggio definitivo: ' + e.message); }}
}};

window.addPlanningPdvManual = async function() {{
  const pv = planPdvNorm(document.getElementById('manualPdv')?.value || '');
  const lat = Number(String(document.getElementById('manualLat')?.value || '').replace(',','.'));
  const lng = Number(String(document.getElementById('manualLng')?.value || '').replace(',','.'));
  if(!pv || !Number.isFinite(lat) || !Number.isFinite(lng)) {{ alert('Servono almeno PV, latitudine e longitudine.'); return; }}
  const tipo = document.getElementById('manualType')?.value === 'grab' ? 'Grab & Go' : 'Telepass Point';
  const agent = document.getElementById('planAgent')?.value || '';
  try {{
    await planApiSaveModifica({{action:'AGGIUNGI', pdv, agente:agent, tipo, citta:document.getElementById('manualCity')?.value || '', indirizzo:document.getElementById('manualAddress')?.value || '', latitudine:lat, longitudine:lng, note:'Aggiunto manualmente da sito'}});
    alert('PV manuale aggiunto definitivamente.');
    renderPlanningAuto();
  }} catch(e) {{ alert('Errore salvataggio definitivo: ' + e.message); }}
}};

window.removePlanningPdv = async function() {{
  const pv = planPdvNorm(document.getElementById('planRemovePdv')?.value || '');
  if(!pv) {{ alert('Inserisci il codice PV.'); return; }}
  try {{
    await planApiSaveModifica({{action:'ESCLUDI', pdv, note:'Escluso da sito'}});
    PLAN.items = (PLAN.items || []).filter(p => p.pdv !== pv);
    alert('PV escluso definitivamente.');
    renderPlanningAuto();
    if(PLAN.items.length) renderPlanningTable(0,0);
  }} catch(e) {{ alert('Errore salvataggio definitivo: ' + e.message); }}
}};

window.resetPlanningLocalChanges = function() {{
  alert('Le modifiche definitive ora sono su Cloudflare. Per ripristinare un PV escluso servirà il pulsante Ripristina che aggiungeremo nello step successivo.');
}};

const _renderPlanningAutoRemote = window.renderPlanningAuto;
window.renderPlanningAuto = function() {{
  if(!window.PLAN_REMOTE_READY) {{
    planApiLoadRemote(true).then(() => {{
      if(window.PLAN_REMOTE_READY && typeof _renderPlanningAutoRemote === 'function') _renderPlanningAutoRemote();
      decoratePlanningRemote();
    }});
  }}
  if(typeof _renderPlanningAutoRemote === 'function') _renderPlanningAutoRemote();
  decoratePlanningRemote();
}};

function decoratePlanningRemote() {{
  const wrap = document.getElementById('planningAutoWrap');
  if(!wrap || document.getElementById('planningRemoteBox')) return;
  const box = document.createElement('div');
  box.id = 'planningRemoteBox';
  box.className = 'card';
  box.innerHTML = `<div class="section-title">Modifiche definitive Cloudflare</div>
    <div class="small-muted">API collegata: ${window.PLANNING_API_URL}<br>Modifiche definitive caricate: <b>${(window.PLAN_REMOTE_MODS||[]).length}</b>. Aggiunte/esclusioni vengono salvate per tutti.</div>
    <div class="plan-actions"><button class="btn light" onclick="planApiLoadRemote(false).then(()=>renderPlanningAuto())">Ricarica modifiche definitive</button></div>`;
  wrap.parentNode.insertBefore(box, wrap.nextSibling);
}}

planApiLoadRemote(true);
</script>
'''


def patch_html(html):
    if "PLANNING_API_URL" in html:
        return html
    html = html.replace("Queste modifiche restano salvate nel browser di chi usa il sito. Non cambiano i file GitHub.", "Queste modifiche vengono salvate definitivamente su Cloudflare e valgono per tutti.")
    if "</body>" in html:
        return html.replace("</body>", CLOUD_JS + "\n</body>", 1)
    return html + CLOUD_JS


def main():
    for name in ["index.html", "Telepass_ENI_sito_v6.html"]:
        path = DOCS_DIR / name
        if path.exists():
            path.write_text(patch_html(path.read_text(encoding="utf-8")), encoding="utf-8")
    print("Planning Cloudflare patch applicata")


if __name__ == "__main__":
    main()
