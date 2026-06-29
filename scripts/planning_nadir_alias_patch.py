from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH = r'''
<script>
(function(){
  function nNorm(s){
    return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  }
  const _oldVisibleAgentName = window.visibleAgentName;
  window.visibleAgentName = visibleAgentName = function(a){
    let v = _oldVisibleAgentName ? _oldVisibleAgentName(a) : String(a||'').trim();
    const n = nNorm(v);
    if(n === 'nadir a' || n === 'nadir abdel' || n === 'nadir abdel aziz' || n === 'nadir abdelaziz' || n === 'nadir abdel aziz myworldsrl it') return 'Nadir Abdel';
    return v;
  };
  function refreshNadir(){
    try{
      if(typeof renderAll === 'function') renderAll();
      if(typeof renderAgents === 'function') renderAgents();
    }catch(e){}
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', refreshNadir);
  else refreshNadir();
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, patch saltata")
        return
    html = path.read_text(encoding="utf-8")
    if "Normalize Nadir agent variants" in html or "nadir abdel aziz myworldsrl" in html:
        print("Nadir alias patch già presente")
        return
    html = html.replace("</body>", "<!-- Normalize Nadir agent variants -->\n" + PATCH + "\n</body>", 1)
    path.write_text(html, encoding="utf-8")
    print("Nadir alias patch applicata")


if __name__ == "__main__":
    main()
