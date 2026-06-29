from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PATCH_JS = r'''
<script>
(function(){
  function pad(n){return String(n).padStart(2,'0');}
  function iso(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());}
  function addDays(d,n){const x=new Date(d);x.setDate(x.getDate()+n);return x;}
  function easter(y){const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),mo=Math.floor((h+l-7*m+114)/31),day=((h+l-7*m+114)%31)+1;return new Date(y,mo-1,day);}
  function parseDateFlex(s, year){s=String(s||'').trim();let m=s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(20\d{2}))?$/);if(m)return new Date(m[3]?+m[3]:year,+m[2]-1,+m[1]);m=s.match(/^(20\d{2})-(\d{1,2})-(\d{1,2})$/);if(m)return new Date(+m[1],+m[2]-1,+m[3]);return null;}
  function defaultHolidays(y){const e=easter(y),em=addDays(e,1);return [new Date(y,0,1),new Date(y,0,6),em,new Date(y,3,25),new Date(y,4,1),new Date(y,5,2),new Date(y,7,15),new Date(y,10,1),new Date(y,11,8),new Date(y,11,25),new Date(y,11,26)].map(iso);}
  window.addClosedRange = addClosedRange = function(){
    const a=document.getElementById('closedStart'), b=document.getElementById('closedEnd'), out=document.getElementById('closedDays');
    if(!a||!out||!a.value){alert('Seleziona almeno la data di inizio ferie');return;}
    const end=(b&&b.value)?b.value:a.value;
    const line=a.value+' - '+end;
    out.value=(out.value?out.value+'\n':'')+line;
    a.value=''; if(b)b.value='';
  };
  window.clearClosedRanges = clearClosedRanges = function(){const out=document.getElementById('closedDays');if(out)out.value='';};
  function closureSet(year){
    const s=new Set(defaultHolidays(year).concat(defaultHolidays(year+1)));
    const el=document.getElementById('closedDays');
    String(el?el.value:'').split(/[\n,;]+/).forEach(part=>{
      part=part.trim();if(!part)return;
      const bits=part.split(/\s+-\s+/);
      if(bits.length===2){let a=parseDateFlex(bits[0],year),b=parseDateFlex(bits[1],year);if(a&&b){if(b<a){const t=a;a=b;b=t;}for(let d=new Date(a);d<=b;d.setDate(d.getDate()+1))s.add(iso(d));}}
      else{const d=parseDateFlex(part,year);if(d)s.add(iso(d));}
    });
    return s;
  }
  window.workdays = workdays = function(mv){
    const [y,m]=mv.split('-').map(Number);
    const closed=closureSet(y);
    const days=[];
    let d=new Date(y,m-1,1), guard=0;
    while(days.length<260 && guard<540){
      const dow=d.getDay(), id=iso(d);
      if(dow!==0 && dow!==6 && !closed.has(id))days.push(id);
      d.setDate(d.getDate()+1); guard++;
    }
    return days;
  };
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, overflow safe saltato")
        return
    html = path.read_text(encoding="utf-8")
    target = '<div><label>Planning precedente</label><input id="prev" type="file" accept=".csv,.txt,.xls"></div>'
    old_closed = '<div><label>Ferie / chiusure</label><textarea id="closedDays" rows="3" placeholder="10/08/2026 - 21/08/2026\n24/12/2026"></textarea><p class="muted small">Sabato, domenica e feste nazionali italiane sono già esclusi.</p></div>'
    closed = '<div><label>Ferie / chiusure</label><div style="display:grid;grid-template-columns:1fr 1fr;gap:8px"><input id="closedStart" type="date"><input id="closedEnd" type="date"></div><div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap"><button class="btn light" type="button" onclick="addClosedRange()">Aggiungi periodo</button><button class="btn light" type="button" onclick="clearClosedRanges()">Svuota ferie</button></div><textarea id="closedDays" rows="3" readonly placeholder="Periodi ferie aggiunti"></textarea><p class="muted small">Sabato, domenica e feste nazionali italiane sono già esclusi.</p></div>'
    if old_closed in html:
        html = html.replace(old_closed, closed, 1)
    elif 'id="closedDays"' not in html and target in html:
        html = html.replace(target, closed + target, 1)
    if 'function closureSet(year)' not in html:
        html = html.replace('</body>', PATCH_JS + '\n</body>', 1)
    elif 'function addClosedRange' not in html:
        html = html.replace('</body>', PATCH_JS + '\n</body>', 1)
    path.write_text(html, encoding='utf-8')
    print('Overflow safe e calendario ferie applicati')


if __name__ == '__main__':
    main()
