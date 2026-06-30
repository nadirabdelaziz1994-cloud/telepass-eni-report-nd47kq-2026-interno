from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

CSS = r'''
<style>
body{padding-bottom:76px!important}
.sticky-save{position:fixed;left:8px;right:8px;bottom:calc(8px + env(safe-area-inset-bottom));z-index:9999;background:#ffffffee;border:1px solid #d9e2ef;border-radius:14px;padding:8px;box-shadow:0 4px 20px #0002;backdrop-filter:blur(8px)}
.sticky-save .btn{width:100%;font-size:14px;padding:11px!important;border-radius:11px!important}
</style>
'''

JS = r'''
<script>
(function(){
  function addStickySave(){
    if(document.getElementById('stickySaveBar'))return;
    const div=document.createElement('div');
    div.id='stickySaveBar';
    div.className='sticky-save';
    div.innerHTML='<button class="btn" onclick="saveAndBack()">Salva modifiche</button>';
    document.body.appendChild(div);
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',addStickySave);else addStickySave();
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning-edit.html"
    if not path.exists():
        print("planning-edit.html non trovato, sticky save saltato")
        return
    html = path.read_text(encoding="utf-8")
    if 'stickySaveBar' not in html:
        html = html.replace('</head>', CSS + '\n</head>', 1)
        html = html.replace('</body>', JS + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Sticky save editor planning applicato")


if __name__ == "__main__":
    main()
