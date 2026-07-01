from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

SNIPPET = r'''
<!-- Grab & Go search bar below filters -->
<script id="grab-go-search-under-filters">
(function(){
  if(window.__mwGrabGoSearchUnderFilters)return;
  window.__mwGrabGoSearchUnderFilters=true;
  window.__mwGrabGoSearchValue='';

  function escAttr(v){
    return String(v == null ? '' : v).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});
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
    var html='<div id="grabGoSearchRow" class="filter-main" style="margin-top:10px">'+
      '<input id="searchText" type="search" placeholder="Cerca PV, comune, indirizzo, agente, contratto, note..." '+
      'value="'+escAttr(window.__mwGrabGoSearchValue||'')+'" '+
      'oninput="window.__mwGrabGoSearchValue=this.value; renderGrabGo()" '+
      'style="width:100%;min-width:280px;padding:10px;border:1px solid var(--line);border-radius:12px;background:#fff">'+
      '</div>';
    main.insertAdjacentHTML('afterend',html);
  }

  function patch(){
    if(typeof window.renderGrabGo!=='function')return false;
    if(window.renderGrabGo.__mwSearchWrapped)return true;
    var oldRender=window.renderGrabGo;
    window.renderGrabGo=function(){
      var current=document.getElementById('searchText');
      if(current) window.__mwGrabGoSearchValue=current.value||'';
      oldRender.apply(this,arguments);
      ensureGrabSearch();
    };
    window.renderGrabGo.__mwSearchWrapped=true;
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


def inject(html: str) -> str:
    if "grab-go-search-under-filters" in html:
        return html
    if "function renderGrabGo" not in html:
        return html
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
    print(f"Grab & Go search patch applicata su {changed} file")


if __name__ == "__main__":
    main()
