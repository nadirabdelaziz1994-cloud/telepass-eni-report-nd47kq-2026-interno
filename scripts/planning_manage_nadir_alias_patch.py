from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH_JS = r'''
<script>
(function(){
  function isNadirAlias(v){
    const s=norm(String(v||'').replace('@',' ').replace(/[._-]+/g,' '));
    const k=agentKey(String(v||'').replace('@',' ').replace(/[._-]+/g,' '));
    return s==='nadir a'||s==='nadir abdel'||s==='abdel nadir'||s.includes('nadir a ')||s.includes('nadir abdel')||k==='a|nadir'||k==='abdel|nadir';
  }
  const oldVisible=window.visibleAgentName;
  window.visibleAgentName=visibleAgentName=function(a){
    if(isNadirAlias(a))return 'Nadir Abdel';
    if(typeof oldVisible==='function')return oldVisible(a);
    return String(a||'').trim();
  };
  function canonPoint(p){
    if(!p)return p;
    if(isNadirAlias(p.agent)||isNadirAlias(p.agent_display)){
      p.agent='Nadir Abdel';
      p.agent_display='Nadir Abdel';
    }
    return p;
  }
  const oldApply=window.applyRemote;
  if(typeof oldApply==='function'){
    window.applyRemote=applyRemote=function(){return oldApply().map(canonPoint);};
  }
  window.agents=agents=function(){return [...new Set(applyRemote().map(p=>visibleAgentName(p.agent_display||p.agent)).filter(Boolean))].filter(a=>!a.includes('@')).sort((a,b)=>a.localeCompare(b,'it'));};
  const oldCurrent=window.currentAgent;
  window.currentAgent=currentAgent=function(){
    const v=document.getElementById('agent')?document.getElementById('agent').value:'';
    return isNadirAlias(v)?'Nadir Abdel':v;
  };
  const oldRenderList=window.renderList;
  window.renderList=renderList=function(){
    const a=currentAgent();const box=document.getElementById('pvList');
    if(!box)return;
    if(!a){box.innerHTML='<div class="empty">Seleziona prima un agente.</div>';return;}
    const pts=applyRemote().filter(p=>visibleAgentName(p.agent_display||p.agent)===a).sort((x,y)=>String(x.region||'').localeCompare(String(y.region||''),'it')||String(x.city||'').localeCompare(String(y.city||''),'it')||String(x.pdv||'').localeCompare(String(y.pdv||''),'it'));
    if(!pts.length){box.innerHTML='<div class="empty">Nessun PV trovato per questo agente.</div>';return;}
    box.innerHTML=pts.map(p=>'<article class="pv-card '+(isGrab(p)?'grab':'')+'"><div class="pv-top"><div><div class="pv-code">PV '+esc(p.pdv)+'</div><span class="pill '+(isGrab(p)?'grab':'')+'">'+(isGrab(p)?'Grab & Go':'TPoint')+'</span></div><button class="btn bad" onclick="removePdv(\''+esc(p.pdv)+'\')">Elimina</button></div><div class="addr"><b>'+esc(p.address||'Via non indicata')+'</b><span>'+esc(p.city||'Comune non indicato')+' · '+esc(p.region||'Regione non indicata')+'</span></div></article>').join('');
  };
  const oldRenderAgents=window.renderAgents;
  window.renderAgents=renderAgents=function(){
    const sel=document.getElementById('agent');if(!sel)return;
    const raw=sel.value||new URLSearchParams(location.search).get('agent')||'';
    const cur=isNadirAlias(raw)?'Nadir Abdel':raw;
    sel.innerHTML='<option value="">Seleziona agente</option>'+agents().map(a=>'<option>'+esc(a)+'</option>').join('');
    if(cur)sel.value=cur;
  };
  try{renderAgents();renderList();}catch(e){}
})();
</script>
'''


def patch_file(path: Path):
    if not path.exists():
        return False
    html = path.read_text(encoding="utf-8")
    if 'Unify Nadir agent aliases in PV manager' not in html:
        html = html.replace('</body>', '<!-- Unify Nadir agent aliases in PV manager -->\n' + PATCH_JS + '\n</body>', 1)
        path.write_text(html, encoding="utf-8")
    return True


def main():
    ok = patch_file(DOCS_DIR / "pdv-manage.html")
    print("Alias Nadir gestione PV applicato" if ok else "pdv-manage.html non trovato, alias Nadir saltato")


if __name__ == "__main__":
    main()
