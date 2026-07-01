from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

SNIPPET = r'''
<!-- Clean colleague UI: hide admin/debug planning details -->
<style id="planning-clean-colleague-ui-style">
  #sourceBox,
  #apiStatus,
  #addPreview { display:none !important; }
</style>
<script id="planning-clean-colleague-ui">
(function(){
  function clean(){
    var source = document.getElementById('sourceBox');
    if(source && source.closest('section')) source.closest('section').style.display = 'none';

    var api = document.getElementById('apiStatus');
    if(api && api.closest('section')) api.closest('section').style.display = 'none';

    document.querySelectorAll('button').forEach(function(btn){
      var txt = (btn.textContent || '').toLowerCase();
      var on = String(btn.getAttribute('onclick') || '').toLowerCase();
      if(txt.includes('ricarica modifiche') || on.includes('loadremote(false)')){
        btn.style.display = 'none';
      }
    });

    document.querySelectorAll('.card .muted').forEach(function(el){
      var txt = (el.textContent || '').toLowerCase();
      if(txt.includes('pagina separata di sicurezza') || txt.includes('mese selezionato più massimo 10 giorni')){
        el.style.display = 'none';
      }
    });
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', clean); else clean();
  window.addEventListener('pageshow', clean);
  var oldRenderAll = window.renderAll;
  if(typeof oldRenderAll === 'function'){
    window.renderAll = function(){
      oldRenderAll.apply(this, arguments);
      clean();
    };
  }
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, clean colleague UI saltata")
        return
    html = path.read_text(encoding="utf-8")
    if 'planning-clean-colleague-ui' not in html:
        html = html.replace('</body>', SNIPPET + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("UI planning colleghi ripulita")


if __name__ == "__main__":
    main()
