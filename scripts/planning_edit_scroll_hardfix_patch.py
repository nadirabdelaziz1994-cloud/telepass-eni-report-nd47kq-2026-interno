from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

HARD_CSS = r'''
<style>
html,body{touch-action:auto!important;overscroll-behavior:auto!important;-webkit-overflow-scrolling:touch!important}
body{min-width:880px!important;overflow-x:auto!important}
.wrap{max-width:none!important;width:880px!important;padding:8px!important}
.edit-list{display:block!important;content-visibility:visible!important}
.edit-row{display:grid!important;grid-template-columns:36px 32px 120px 86px 150px 190px 90px 38px!important;gap:4px!important;align-items:center!important;padding:5px!important;margin:4px 0!important;border-radius:8px!important;touch-action:auto!important;min-height:42px!important;background:#fff!important}
.edit-row.grab{background:#fbf7ff!important}.edit-row.dual{background:#fffaf0!important}
.drag{height:32px!important;width:32px!important;touch-action:none!important;cursor:grab!important;font-size:17px!important;user-select:none!important;-webkit-user-select:none!important}
.idx{font-size:11px!important}.date{font-size:11px!important;padding:5px!important;height:32px!important}.main b{font-size:12px!important}.main span{font-size:10px!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.x{height:32px!important;width:32px!important;padding:0!important}.col{font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.col b{font-size:12px}.head-row{position:sticky;top:55px;z-index:3;background:#eef4ff;border:1px solid #bfd4ef;border-radius:8px;padding:6px 5px;margin-bottom:6px;display:grid;grid-template-columns:36px 32px 120px 86px 150px 190px 90px 38px;gap:4px;color:#0b2d5c;font-size:11px;font-weight:900}.hint{font-size:12px;color:#64748b}
</style>
'''

HARD_JS = r'''
<script>
(function(){
  function pClass(p){if(p.is_grab&&p.is_tp)return 'dual';if(p.is_grab)return 'grab';return 'tp';}
  function pTipo(p){if(p.is_grab&&p.is_tp)return 'Doppio';if(p.is_grab)return 'Grab';return 'TPoint';}
  window.render = function(){
    const box=document.getElementById('list');
    if(!PLAN.length){box.innerHTML='<div class="card muted">Nessun planning trovato. Torna alla pagina planning e crealo prima.</div>';return;}
    box.innerHTML='<div class="head-row"><div>↕</div><div>#</div><div>Data</div><div>PV</div><div>Comune</div><div>Via</div><div>Tipo</div><div>X</div></div>'+
      PLAN.map((p,i)=>'<div class="edit-row '+pClass(p)+'" data-i="'+i+'"><button class="drag" ontouchstart="touchStart(event,'+i+')" ontouchmove="touchMove(event)" ontouchend="touchEnd(event)" onmousedown="dragIndex='+i+'">↕</button><div class="idx">'+(i+1)+'</div><input class="date" type="date" value="'+esc(p.date||'')+'" onchange="changeDate('+i+',this.value)"><div class="col"><b>'+esc(p.pdv)+'</b></div><div class="col">'+esc(p.city||'')+'</div><div class="col">'+esc(p.address||'')+'</div><div class="col">'+esc(pTipo(p))+'</div><button class="btn bad x" onclick="del('+i+')">×</button></div>').join('');
  };
  window.dragStart=function(e,i){};
  window.dragOver=function(e){};
  window.dropRow=function(e,i){};
  window.touchStart=function(e,i){touchIndex=i;const row=e.currentTarget.closest('.edit-row');if(row)row.classList.add('dragging');e.preventDefault();};
  window.touchMove=function(e){if(touchIndex==null)return;const t=e.touches[0];const el=document.elementFromPoint(t.clientX,t.clientY)?.closest('.edit-row');if(!el)return;const j=Number(el.dataset.i);if(Number.isFinite(j)&&j!==touchIndex){move(touchIndex,j);touchIndex=j;}e.preventDefault();};
  window.touchEnd=function(e){touchIndex=null;document.querySelectorAll('.dragging').forEach(x=>x.classList.remove('dragging'));e.preventDefault();};
  document.addEventListener('touchmove',function(e){if(e.target.closest('.drag'))return;}, {passive:true});
  try{render();}catch(e){}
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning-edit.html"
    if not path.exists():
        print("planning-edit.html non trovato, hardfix saltato")
        return
    html = path.read_text(encoding="utf-8")
    import re
    html = re.sub(r'<meta name="viewport"[^>]*>', '<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=10, minimum-scale=0.25, user-scalable=yes">', html, count=1)
    # Remove mobile-draggable row attributes from the original generated rows by overriding render after load.
    if 'Hard fix planning editor scroll and zoom' not in html:
        html = html.replace('</head>', '<!-- Hard fix planning editor scroll and zoom -->\n' + HARD_CSS + '\n</head>', 1)
        html = html.replace('</body>', '<!-- Hard fix planning editor scroll and zoom -->\n' + HARD_JS + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Hard fix scroll/zoom editor planning applicato")


if __name__ == "__main__":
    main()
