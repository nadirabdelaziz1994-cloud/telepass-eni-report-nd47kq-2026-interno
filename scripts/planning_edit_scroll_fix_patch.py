from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

FIX_JS = r'''
<script>
(function(){
  let holdTimer=null;
  window.touchStart = function(e,i){
    touchIndex=i;
    const row=e.currentTarget.closest('.edit-row');
    if(row)row.classList.add('dragging');
    holdTimer=setTimeout(()=>{},120);
    e.preventDefault();
  };
  window.touchMove = function(e){
    if(touchIndex==null)return;
    const t=e.touches[0];
    const el=document.elementFromPoint(t.clientX,t.clientY)?.closest('.edit-row');
    if(!el)return;
    const j=Number(el.dataset.i);
    if(Number.isFinite(j)&&j!==touchIndex){move(touchIndex,j);touchIndex=j;}
    e.preventDefault();
  };
  window.touchEnd = function(e){
    touchIndex=null;
    if(holdTimer)clearTimeout(holdTimer);
    document.querySelectorAll('.dragging').forEach(x=>x.classList.remove('dragging'));
    e.preventDefault();
  };
})();
</script>
'''

FIX_CSS = r'''
<style>
html,body{touch-action:pan-x pan-y;overscroll-behavior:auto;-webkit-overflow-scrolling:touch}
.edit-list{content-visibility:auto;contain-intrinsic-size:1000px}
.edit-row{touch-action:pan-y!important;grid-template-columns:34px 25px 112px minmax(120px,1fr) 34px!important;padding:6px!important;gap:4px!important;border-radius:10px!important}
.drag{touch-action:none!important;height:34px!important;font-size:18px!important;user-select:none;-webkit-user-select:none}
.x{width:34px!important;height:34px!important;padding:0!important}.date{font-size:12px!important;padding:6px!important;border-radius:8px!important}.idx{font-size:11px!important}.main b{font-size:13px!important}.main span{font-size:11px!important;line-height:1.15}.card{padding:10px!important;margin:8px 0!important}.wrap{padding:7px!important}.muted{font-size:12px!important}.btn{padding:9px 10px!important}
@media(max-width:560px){.edit-row{grid-template-columns:32px 22px 104px minmax(100px,1fr) 32px!important}.drag{height:32px!important}.x{width:32px!important;height:32px!important}.date{font-size:11px!important}.main b{font-size:12px!important}.main span{font-size:10px!important}}
</style>
'''


def main():
    path = DOCS_DIR / "planning-edit.html"
    if not path.exists():
        print("planning-edit.html non trovato, fix scroll saltato")
        return
    html = path.read_text(encoding="utf-8")
    html = html.replace(
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        '<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=5, user-scalable=yes">'
    )
    # keep normal page scrolling; dragging is restricted to the handle only
    html = html.replace('padding:8px;touch-action:none}', 'padding:8px;touch-action:pan-y}')
    html = html.replace('padding:7px;gap:5px}', 'padding:7px;gap:5px;touch-action:pan-y}')
    if 'content-visibility:auto' not in html:
        html = html.replace('</head>', FIX_CSS + '\n</head>', 1)
    if 'html,body{touch-action:pan-x pan-y' not in html:
        html = html.replace('</head>', FIX_CSS + '\n</head>', 1)
    # Put the override after the original functions, so it wins.
    if 'Fix mobile scrolling and zoom' not in html and 'window.touchStart = function' not in html:
        html = html.replace('</body>', '<!-- Fix mobile scrolling and zoom -->\n' + FIX_JS + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Fix scroll/zoom editor planning applicato")


if __name__ == "__main__":
    main()
