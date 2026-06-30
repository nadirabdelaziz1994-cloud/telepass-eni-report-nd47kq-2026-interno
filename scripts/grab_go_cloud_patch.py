from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

MONTHS_LINE = "const GRAB_GO_MONTHS=['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];"
SET_LOCAL = "function setGrabVisit(pdv,m){const s=grabState();if(!m){delete s[pdv];}else{s[pdv]={month:Number(m),year:new Date().getFullYear(),saved_at:new Date().toISOString()};}saveGrabState(s);renderGrabGo();}"
RENDER_START = "function renderGrabGo(){const w=document.getElementById('grabGoWrap');"

REMOTE_JS = r'''
const GRAB_REMOTE_API='https://telepass-planning-api.nadirabdelaziz1994.workers.dev';
let GRAB_REMOTE_READY=false;
let GRAB_REMOTE_LOADING=false;
let GRAB_REMOTE_ATTEMPTED=false;
let GRAB_REMOTE_ERROR='';
function grabPadPdv(pdv){return String(pdv||'').replace(/\D+/g,'').padStart(5,'0');}
function grabSessionToken(){return String(localStorage.getItem('telepassCloudSession')||'').trim();}
function grabAdminUnlocked(){return !!grabSessionToken();}
async function unlockGrabCloud(){const u=prompt('Utente MyWorld')||'';if(!u.trim())return false;const c=prompt('Codice MyWorld')||'';if(!c)return false;try{const body={username:u.trim()};body['pass'+'word']=c;const res=await fetch(GRAB_REMOTE_API+'/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});const data=await res.json().catch(()=>({}));if(!res.ok||!data.ok)throw new Error(data.error||'Login non riuscito');localStorage.setItem('telepassCloudSession',data.token||'');localStorage.setItem('telepassCloudUser',data.username||u.trim());alert('Login cloud attivo su questo dispositivo.');renderGrabGo();return true;}catch(e){alert('Login cloud non riuscito: '+String(e&&e.message?e.message:e));return false;}}
function lockGrabCloud(){localStorage.removeItem('telepassCloudSession');localStorage.removeItem('telepassCloudUser');localStorage.removeItem('telepassAdminKey');alert('Logout cloud effettuato.');renderGrabGo();}
function grabCloudLabel(){const lock=grabAdminUnlocked()?' · Login attivo':' · Login richiesto';if(GRAB_REMOTE_READY)return ' · Cloudflare attivo'+lock;if(GRAB_REMOTE_LOADING)return ' · Cloudflare: caricamento visite condivise'+lock;if(GRAB_REMOTE_ERROR)return ' · Cloudflare errore: '+esc(GRAB_REMOTE_ERROR)+lock;return ' · Cloudflare non ancora caricato'+lock;}
function grabCloudButtons(){const a='<button class="btn light" type="button" onclick="refreshGrabRemote()" style="margin-left:8px;padding:5px 9px">Aggiorna visite</button>';const b=grabAdminUnlocked()?'<button class="btn light" type="button" onclick="lockGrabCloud()" style="margin-left:8px;padding:5px 9px">Logout</button>':'<button class="btn" type="button" onclick="unlockGrabCloud()" style="margin-left:8px;padding:5px 9px">Login cloud</button>';return a+b;}
async function loadGrabRemote(){if(GRAB_REMOTE_LOADING)return;GRAB_REMOTE_LOADING=true;GRAB_REMOTE_ATTEMPTED=true;try{const res=await fetch(GRAB_REMOTE_API+'/grab-visite?ts='+Date.now(),{cache:'no-store'});const data=await res.json();if(!res.ok||!data.ok)throw new Error(data.error||('HTTP '+res.status));const s={};(data.visite||[]).forEach(v=>{const pdv=grabPadPdv(v.pdv);if(pdv&&Number(v.month)){s[pdv]={month:Number(v.month),year:Number(v.year||new Date().getFullYear()),saved_at:v.saved_at||v.updated_at||''};}});saveGrabState(s);GRAB_REMOTE_READY=true;GRAB_REMOTE_ERROR='';}catch(e){GRAB_REMOTE_ERROR=String(e&&e.message?e.message:e);}finally{GRAB_REMOTE_LOADING=false;if(typeof activePage==='function'&&activePage()==='grab-go')renderGrabGo();}}
async function saveGrabRemote(pdv,m){const token=grabSessionToken();if(!token){alert('Prima fai Login cloud.');await unlockGrabCloud();renderGrabGo();return false;}try{const res=await fetch(GRAB_REMOTE_API+'/grab-visita',{method:'POST',headers:{'Content-Type':'application/json','Authorization':'Bearer '+token},body:JSON.stringify({pdv:grabPadPdv(pdv),month:Number(m||0),year:new Date().getFullYear()})});const data=await res.json().catch(()=>({}));if(!res.ok||!data.ok){if(res.status===401){localStorage.removeItem('telepassCloudSession');localStorage.removeItem('telepassCloudUser');}alert('Errore salvataggio Cloudflare: '+(data.error||res.status));renderGrabGo();return false;}GRAB_REMOTE_READY=true;GRAB_REMOTE_ERROR='';return true;}catch(e){alert('Errore rete Cloudflare: '+String(e&&e.message?e.message:e));return false;}}
function refreshGrabRemote(){GRAB_REMOTE_READY=false;GRAB_REMOTE_ERROR='';GRAB_REMOTE_ATTEMPTED=false;loadGrabRemote();}
'''

SET_REMOTE = "async function setGrabVisit(pdv,m){const key=grabPadPdv(pdv);const oldState=grabState();const oldVal=oldState[key];const s=grabState();if(!m){delete s[key];}else{s[key]={month:Number(m),year:new Date().getFullYear(),saved_at:new Date().toISOString()};}saveGrabState(s);renderGrabGo();const ok=await saveGrabRemote(key,m);if(!ok){const rollback=grabState();if(oldVal){rollback[key]=oldVal;}else{delete rollback[key];}saveGrabState(rollback);renderGrabGo();}}"


def patch_html(html: str) -> str:
    if "GRAB_REMOTE_API" in html:
        return html
    if MONTHS_LINE not in html or SET_LOCAL not in html or RENDER_START not in html:
        raise RuntimeError("Struttura Grab & Go non trovata: esegui prima grab_go_patch.py")
    html = html.replace(MONTHS_LINE, MONTHS_LINE + "\n" + REMOTE_JS, 1)
    html = html.replace(SET_LOCAL, SET_REMOTE, 1)
    html = html.replace(RENDER_START, "function renderGrabGo(){if(!GRAB_REMOTE_ATTEMPTED&&!GRAB_REMOTE_LOADING)loadGrabRemote();const w=document.getElementById('grabGoWrap');", 1)
    html = html.replace(
        "Filtro agenti separato da Home/Classifica.</div></div><div class=\"metric-row",
        "Filtro agenti separato da Home/Classifica.${grabCloudLabel()} ${grabCloudButtons()}</div></div><div class=\"metric-row",
        1,
    )
    return html


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
    print(f"Grab & Go Cloudflare sync patch completata: {patched} file aggiornati")


if __name__ == "__main__":
    main()
