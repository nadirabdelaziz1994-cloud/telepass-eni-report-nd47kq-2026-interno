from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH_JS = r'''
<script>
(function(){
  function avgPoint(list,label){
    const pts=(list||[]).filter(p=>p&&p.lat!=null&&p.lng!=null);
    if(!pts.length)return null;
    const lat=pts.reduce((s,p)=>s+Number(p.lat||0),0)/pts.length;
    const lng=pts.reduce((s,p)=>s+Number(p.lng||0),0)/pts.length;
    return {pdv:'PARTENZA',city:label||'Partenza',address:document.getElementById('start')?document.getElementById('start').value:'',lat,lng,is_start:true,is_tp:false,is_grab:false};
  }
  window.startPoint = startPoint = function(points,start){
    const raw=String(start||'').trim();
    const s=norm(raw);
    if(!s)return points[0]||null;
    const byPdv=(points||[]).find(p=>norm(p.pdv)===s);
    if(byPdv)return byPdv;
    const all=[...(points||[]),...((DATA&&DATA.catalog)||[])].filter(p=>p&&p.lat!=null&&p.lng!=null);
    let exactCity=all.filter(p=>norm(p.city)===s);
    if(exactCity.length)return avgPoint(exactCity,exactCity[0].city||raw);
    let containedCity=all.filter(p=>{const c=norm(p.city);return c&&s.includes(c);});
    if(containedCity.length)return avgPoint(containedCity,containedCity[0].city||raw);
    let containedAddr=all.filter(p=>{const text=norm((p.city||'')+' '+(p.address||''));return text&&text.includes(s);});
    if(containedAddr.length)return avgPoint(containedAddr,raw);
    return points[0]||null;
  };
  function fixStartLabel(){
    const input=document.getElementById('start');
    if(!input)return;
    const wrap=input.closest('div');
    const lab=wrap?wrap.querySelector('label'):null;
    if(lab)lab.textContent='Punto di partenza · Città e via';
    input.placeholder='Es. Lissone, Via Roma 10';
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',fixStartLabel);else fixStartLabel();
  window.addEventListener('pageshow',fixStartLabel);
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, start home saltato")
        return
    html = path.read_text(encoding="utf-8")
    html = html.replace('<label>Punto di partenza</label><input id="start" placeholder="Città o codice PV">', '<label>Punto di partenza · Città e via</label><input id="start" placeholder="Es. Lissone, Via Roma 10">')
    if 'Use home address as planning start point' not in html:
        html = html.replace('</body>', '<!-- Use home address as planning start point -->\n' + PATCH_JS + '\n</body>', 1)
    path.write_text(html, encoding="utf-8")
    print("Partenza da abitazione applicata")


if __name__ == "__main__":
    main()
