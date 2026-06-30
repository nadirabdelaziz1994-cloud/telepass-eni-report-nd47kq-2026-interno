from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH = r'''
<script>
(function(){
  function safeJson(raw){try{return JSON.parse(raw||'{}')}catch(e){return {}}}
  function getCurrentMeta(){return safeJson(sessionStorage.getItem('planningCurrent')||'{}')}
  function setVal(id,v){const el=document.getElementById(id);if(el&&v!=null&&v!=='')el.value=v}
  function restorePlanFromStorage(force){
    const edited=safeJson(sessionStorage.getItem('planningEdited')||'{}');
    const current=getCurrentMeta();
    const source=(edited&&Array.isArray(edited.plan)&&edited.plan.length)?edited:current;
    if(!force && !(edited&&edited.plan&&edited.plan.length))return false;
    if(!source||!Array.isArray(source.plan)||!source.plan.length)return false;
    try{
      if(source.agent)setVal('agent',source.agent);
      if(source.month)setVal('month',source.month);
      if(source.start)setVal('start',source.start);
      if(source.grab)setVal('grab',source.grab);
      PLAN=source.plan;
      sessionStorage.setItem('planningCurrent',JSON.stringify(Object.assign({},current,source,{plan:PLAN,savedAt:Date.now()})));
      sessionStorage.removeItem('planningEdited');
      if(typeof renderTable==='function')renderTable(0);
      const res=document.getElementById('result');
      if(res)res.scrollIntoView({behavior:'smooth',block:'start'});
      return true;
    }catch(e){console.warn('restore planning failed',e);return false;}
  }
  function currentPlanPayload(){
    const cur=getCurrentMeta();
    const plan=(typeof PLAN!=='undefined'&&Array.isArray(PLAN)&&PLAN.length)?PLAN:(Array.isArray(cur.plan)?cur.plan:[]);
    return Object.assign({},cur,{plan:plan,version:1,kind:'telepass-planning-editable',savedAt:Date.now()});
  }
  window.downloadEditablePlanning = downloadEditablePlanning = function(){
    const payload=currentPlanPayload();
    if(!payload.plan||!payload.plan.length){alert('Prima crea o carica un planning');return false;}
    const blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'});
    const a=document.createElement('a');
    const agent=(payload.agent||'planning').replace(/[^a-z0-9]+/gi,'_');
    const month=(payload.month||'mese').replace(/[^a-z0-9]+/gi,'_');
    a.href=URL.createObjectURL(blob);
    a.download='planning_modificabile_'+agent+'_'+month+'.json';
    document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(a.href),1000);
    return true;
  };
  window.downloadExcelAndEditable = downloadExcelAndEditable = function(){
    const payload=currentPlanPayload();
    if(!payload.plan||!payload.plan.length){alert('Prima crea o carica un planning');return;}
    try{sessionStorage.setItem('planningCurrent',JSON.stringify(payload));}catch(e){}
    if(typeof downloadXls==='function')downloadXls();else alert('Funzione Excel non trovata');
    setTimeout(()=>downloadEditablePlanning(),500);
  };
  window.openEditablePlanningFile = openEditablePlanningFile = function(){
    const inp=document.getElementById('editablePlanningFile');
    if(inp)inp.click();
  };
  window.importEditablePlanning = importEditablePlanning = function(input){
    const file=input&&input.files&&input.files[0];
    if(!file)return;
    const r=new FileReader();
    r.onload=function(){
      const data=safeJson(r.result);
      if(!data||!Array.isArray(data.plan)||!data.plan.length){alert('File planning modificabile non valido');return;}
      sessionStorage.setItem('planningCurrent',JSON.stringify(data));
      location.href='./planning-edit.html?v=from-editable-file';
    };
    r.readAsText(file);
    input.value='';
  };
  function ensureButtons(){
    const buttons=[...document.querySelectorAll('button')];
    const excel=buttons.find(b=>(b.textContent||'').toLowerCase().includes('scarica excel'));
    if(excel&&!excel.dataset.comboDone){
      excel.textContent='Scarica Excel + file modificabile';
      excel.onclick=function(e){e.preventDefault();downloadExcelAndEditable();};
      excel.dataset.comboDone='1';
    }
    const oldEditable=[...document.querySelectorAll('button')].find(b=>(b.textContent||'').toLowerCase().includes('scarica planning modificabile'));
    if(oldEditable)oldEditable.remove();
    if(document.getElementById('editablePlanningButtons'))return;
    const box=document.createElement('div');
    box.id='editablePlanningButtons';
    box.style.cssText='display:flex;gap:8px;flex-wrap:wrap;margin-top:10px';
    box.innerHTML='<button class="btn light" type="button" onclick="openEditablePlanningFile()">Modifica planning creato</button><input id="editablePlanningFile" type="file" accept=".json,application/json" style="display:none" onchange="importEditablePlanning(this)">';
    if(excel&&excel.parentElement)excel.parentElement.appendChild(box);else document.body.appendChild(box);
  }
  const oldOpen=window.openPlanningEditor;
  window.openPlanningEditor=function(){
    try{sessionStorage.setItem('planningCurrent',JSON.stringify(currentPlanPayload()));}catch(e){}
    if(typeof oldOpen==='function')oldOpen();else location.href='./planning-edit.html?v=edit';
  };
  window.addEventListener('pageshow',function(){ensureButtons();setTimeout(ensureButtons,250);setTimeout(()=>restorePlanFromStorage(location.href.includes('edit-return')),80);setTimeout(()=>restorePlanFromStorage(location.href.includes('edit-return')),350);});
  window.addEventListener('load',function(){ensureButtons();setTimeout(ensureButtons,300);setTimeout(()=>restorePlanFromStorage(location.href.includes('edit-return')),120);setTimeout(()=>restorePlanFromStorage(location.href.includes('edit-return')),600);});
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, persistence saltato")
        return
    html = path.read_text(encoding="utf-8")
    if 'downloadEditablePlanning' not in html:
        html = html.replace('</body>', PATCH + '\n</body>', 1)
    else:
        # Replace previous persistence patch by appending the newer one last so it wins.
        if 'downloadExcelAndEditable' not in html:
            html = html.replace('</body>', PATCH + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Persistenza planning modificato applicata")


if __name__ == "__main__":
    main()
