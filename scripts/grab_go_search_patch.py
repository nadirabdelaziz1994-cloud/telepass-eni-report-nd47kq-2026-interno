from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

SNIPPET = r'''
<!-- Grab & Go search bar below filters + CR in PV description -->
<script id="grab-go-search-under-filters">
(function(){
  if(window.__mwGrabGoSearchUnderFiltersV2)return;
  window.__mwGrabGoSearchUnderFiltersV2=true;
  window.__mwGrabGoSearchValue='';

  function escAttr(v){
    return String(v == null ? '' : v).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});
  }
  function normPdvLocal(v){
    var m=String(v||'').match(/\d+/);
    return m ? m[0].padStart(5,'0') : '';
  }
  function crForPdv(pdv){
    var n=normPdvLocal(pdv);
    if(!n)return '';
    try{
      var row=(DATA||[]).find(function(x){return normPdvLocal(x&&x.pdv)===n;});
      if(row && row.cr)return String(row.cr).trim();
    }catch(e){}
    try{
      var row2=(APP&&APP.rows?APP.rows:[]).find(function(x){return normPdvLocal(x&&x.pdv)===n;});
      if(row2 && row2.cr)return String(row2.cr).trim();
    }catch(e){}
    return '';
  }

  function ensureGrabSearch(){
    var wrap=document.getElementById('grabGoWrap');
    if(!wrap)return;
    var filters=wrap.querySelector('.grab-filters');
    if(!filters)return;
    var main=filters.querySelector('.filter-main') || filters.firstElementChild;
    if(!main)return;
    var old=document.getElementById('searchText');
    if(old) window.__mwGrabGoSearchValue=old.value||window.__mwGrabGoSearchValue||'';
    var existing=document.getElementById('grabGoSearchRow');
    if(existing) existing.remove();
    var html='<div id="grabGoSearchRow" class="filter-main" style="margin-top:10px">'+
      '<input id="searchText" type="search" placeholder="Cerca PV, comune, indirizzo, agente, CR, contratto, note..." '+
      'value="'+escAttr(window.__mwGrabGoSearchValue||'')+'" '+
      'oninput="window.__mwGrabGoSearchValue=this.value; renderGrabGo()" '+
      'style="width:100%;min-width:280px;padding:10px;border:1px solid var(--line);border-radius:12px;background:#fff">'+
      '</div>';
    main.insertAdjacentHTML('afterend',html);
  }

  function addCrToRows(){
    var wrap=document.getElementById('grabGoWrap');
    if(!wrap)return;
    wrap.querySelectorAll('table.grab-table tbody tr').forEach(function(tr){
      if(tr.__mwCrAdded)return;
      var pdvCell=tr.children && tr.children[1];
      var desc=tr.querySelector('.city-cell');
      if(!pdvCell || !desc)return;
      var cr=crForPdv(pdvCell.textContent||'');
      if(!cr)return;
      var div=document.createElement('div');
      div.className='small-muted';
      div.textContent='CR: '+cr;
      desc.appendChild(div);
      tr.__mwCrAdded=true;
    });
  }

  function patch(){
    if(typeof window.renderGrabGo!=='function')return false;
    if(window.renderGrabGo.__mwSearchWrappedV2)return true;
    var oldRender=window.renderGrabGo;
    window.renderGrabGo=function(){
      var current=document.getElementById('searchText');
      if(current) window.__mwGrabGoSearchValue=current.value||'';
      oldRender.apply(this,arguments);
      ensureGrabSearch();
      addCrToRows();
    };
    window.renderGrabGo.__mwSearchWrappedV2=true;
    try{window.renderGrabGo();}catch(e){}
    return true;
  }

  if(!patch()){
    var tries=0;
    var timer=setInterval(function(){
      tries++;
      if(patch() || tries>50) clearInterval(timer);
    },100);
  }
})();
</script>
'''


def remove_old(html: str) -> str:
    start = html.find('<!-- Grab & Go search bar below filters')
    while start != -1:
        end = html.find('</script>', start)
        if end == -1:
            break
        html = html[:start] + html[end + len('</script>'):]
        start = html.find('<!-- Grab & Go search bar below filters')
    return html


def inject(html: str) -> str:
    if "function renderGrabGo" not in html:
        return html
    html = remove_old(html)
    return html.replace("</body>", SNIPPET + "\n</body>", 1)


def main():
    changed = 0
    for name in ["index.html", "Telepass_ENI_sito_v6.html"]:
        path = DOCS_DIR / name
        if not path.exists():
            continue
        original = path.read_text(encoding="utf-8")
        updated = inject(original)
        if updated != original:
            path.write_text(updated, encoding="utf-8")
            changed += 1
    print(f"Grab & Go search/CR patch applicata su {changed} file")


if __name__ == "__main__":
    main()
