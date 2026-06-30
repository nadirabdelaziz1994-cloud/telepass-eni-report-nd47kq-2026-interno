from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"
API_URL = "https://telepass-planning-api.nadirabdelaziz1994.workers.dev"

PATCH_JS = f'''
<script>
(function(){{
  const SITE_REMOTE_API_URL='{API_URL}';
  window.GRAB_REMOTE_VISITS=window.GRAB_REMOTE_VISITS||{{}};
  function siteNormPdv(v){{const m=String(v||'').match(/\d+/);return m?m[0].padStart(5,'0'):'';}}
  function filterPdvArray(arr,excluded){{
    if(!Array.isArray(arr))return 0;
    let removed=0;
    for(let i=arr.length-1;i>=0;i--){{if(excluded.has(siteNormPdv(arr[i]&&arr[i].pdv))){{arr.splice(i,1);removed++;}}}}
    return removed;
  }}
  async function loadMainRemoteExclusions(){{
    try{{
      const res=await fetch(SITE_REMOTE_API_URL+'/modifiche');
      const data=await res.json();
      const mods=data.modifiche||[];
      const excluded=new Set(mods.filter(m=>String(m.action||'').toUpperCase()==='ESCLUDI').map(m=>siteNormPdv(m.pdv)).filter(Boolean));
      if(!excluded.size)return;
      filterPdvArray(APP.rows,excluded);
      filterPdvArray(window.DATA||[],excluded);
      if(APP.custom_report&&Array.isArray(APP.custom_report.rows))filterPdvArray(APP.custom_report.rows,excluded);
      if(APP.grab_go&&Array.isArray(APP.grab_go.rows))filterPdvArray(APP.grab_go.rows,excluded);
      if(typeof refreshAll==='function')refreshAll();
      if(typeof renderGrabGo==='function'&&typeof activePage==='function'&&activePage()==='grab-go')renderGrabGo();
      if(typeof renderGarePdv==='function'&&typeof activePage==='function'&&activePage()==='gare-pdv')renderGarePdv();
    }}catch(e){{console.warn('Modifiche Cloudflare non caricate nel sito principale',e);}}
  }}
  async function loadGrabRemoteVisits(){{
    try{{
      const res=await fetch(SITE_REMOTE_API_URL+'/grab-visite');
      const data=await res.json();
      const rows=data.visite||data.rows||[];
      const out={{}};
      rows.forEach(r=>{{const pdv=siteNormPdv(r.pdv);const m=Number(r.month);if(pdv&&m)out[pdv]={{month:m,year:Number(r.year)||new Date().getFullYear(),saved_at:r.saved_at||r.updated_at||'',source:'cloudflare'}};}});
      window.GRAB_REMOTE_VISITS=out;
      if(typeof renderGrabGo==='function'&&typeof activePage==='function'&&activePage()==='grab-go')renderGrabGo();
    }}catch(e){{console.warn('Visite Grab & Go Cloudflare non disponibili, uso salvataggio browser',e);}}
  }}
  const oldGrabState=window.grabState;
  window.grabState=grabState=function(){{
    let local={{}};
    try{{local=oldGrabState?oldGrabState():JSON.parse(localStorage.getItem('grabGoVisits')||'{{}}')||{{}};}}catch(e){{local={{}};}}
    return Object.assign({{}},local,window.GRAB_REMOTE_VISITS||{{}});
  }};
  const oldSaveGrabState=window.saveGrabState;
  window.saveGrabState=saveGrabState=function(s){{
    if(typeof oldSaveGrabState==='function')return oldSaveGrabState(s);
    localStorage.setItem('grabGoVisits',JSON.stringify(s||{{}}));
  }};
  window.setGrabVisit=setGrabVisit=async function(pdv,m){{
    const p=siteNormPdv(pdv);
    const month=Number(m)||0;
    const year=new Date().getFullYear();
    let local={{}};try{{local=JSON.parse(localStorage.getItem('grabGoVisits')||'{{}}')||{{}};}}catch(e){{local={{}};}}
    if(!month){{delete local[p];delete window.GRAB_REMOTE_VISITS[p];}}else{{local[p]={{month,year,saved_at:new Date().toISOString()}};window.GRAB_REMOTE_VISITS[p]=local[p];}}
    localStorage.setItem('grabGoVisits',JSON.stringify(local));
    if(typeof renderGrabGo==='function')renderGrabGo();
    try{{
      let key=sessionStorage.getItem('planningAdminKey')||'';
      if(!key){{key=prompt('Inserisci ADMIN_KEY Cloudflare per salvare la visita per tutti');if(key)sessionStorage.setItem('planningAdminKey',key);}}
      if(!key)return;
      await fetch(SITE_REMOTE_API_URL+'/grab-visita',{{method:'POST',headers:{{'Content-Type':'application/json','X-Admin-Key':key}},body:JSON.stringify({{pdv:p,month:month||null,year}})}});
      await loadGrabRemoteVisits();
    }}catch(e){{alert('Visita salvata solo su questo browser. Cloudflare non ha risposto: '+e.message);}}
  }};
  function addPlanningMenuButton(){{
    const nav=document.querySelector('.nav');
    if(!nav||nav.querySelector('[data-planning-link]'))return;
    const btn=document.createElement('button');
    btn.type='button';
    btn.dataset.planningLink='1';
    btn.textContent='Planning automatico';
    btn.onclick=function(){{location.href='./planning.html?v=from-main';}};
    const grab=[...nav.children].find(x=>(x.textContent||'').trim().toLowerCase()==='grab & go');
    if(grab&&grab.nextSibling)nav.insertBefore(btn,grab.nextSibling);else nav.appendChild(btn);
  }}
  addPlanningMenuButton();
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',addPlanningMenuButton);else addPlanningMenuButton();
  window.addEventListener('pageshow',addPlanningMenuButton);
  loadMainRemoteExclusions();
  loadGrabRemoteVisits();
}})();
</script>
'''


def patch_html(html: str) -> str:
    # Static fallback insertion, then JS also verifies/creates the button.
    if 'data-planning-link' not in html:
        html = html.replace(
            '<button data-page="grab-go" onclick="showPage(\'grab-go\', this)">Grab & Go</button>',
            '<button data-page="grab-go" onclick="showPage(\'grab-go\', this)">Grab & Go</button>\n      <button data-planning-link="1" onclick="location.href=\'./planning.html?v=from-main\'">Planning automatico</button>',
            1,
        )
        html = html.replace(
            '<button data-page="file-utili" onclick="showPage(\'file-utili\', this)">File utili</button>',
            '<button data-planning-link="1" onclick="location.href=\'./planning.html?v=from-main\'">Planning automatico</button>\n      <button data-page="file-utili" onclick="showPage(\'file-utili\', this)">File utili</button>',
            1,
        )
    if 'loadMainRemoteExclusions' not in html:
        html = html.replace('</body>', PATCH_JS + '\n</body>', 1)
    return html


def main():
    done=0
    for name in ['index.html','Telepass_ENI_sito_v6.html']:
        path=DOCS_DIR/name
        if path.exists():
            path.write_text(patch_html(path.read_text(encoding='utf-8')), encoding='utf-8')
            done+=1
    print(f'Patch sito principale Cloudflare/menu applicata: {done} file')


if __name__ == '__main__':
    main()
