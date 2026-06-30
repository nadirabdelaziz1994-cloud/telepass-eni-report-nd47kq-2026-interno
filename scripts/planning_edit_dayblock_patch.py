from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

CSS = r'''
<style>
.day-block{display:grid;grid-template-columns:116px 1fr 120px;gap:6px;align-items:center;background:#0b2d5c;color:white;border-radius:10px;padding:6px;margin:10px 0 5px;position:relative;z-index:2}
.day-controls{display:flex;gap:4px;align-items:center}.day-drag,.day-mini{height:32px;width:34px;border-radius:9px;border:1px solid #ffffff55;background:#ffffff18;color:white;font-weight:900;font-size:15px;touch-action:none;user-select:none;-webkit-user-select:none}.day-mini{touch-action:auto;cursor:pointer}.day-title{font-weight:900;font-size:13px}.day-title span{font-weight:700;opacity:.8}.day-date{height:32px;border-radius:8px;border:1px solid #ffffff55;background:white;color:#102033;padding:5px;font-size:12px}.day-moving{opacity:.55;outline:3px solid #60a5fa}.edit-row{margin-left:12px!important;width:calc(100% - 12px)!important}.head-row{top:52px!important}
</style>
'''

JS = r'''
<script>
(function(){
  let dayTouchIndex=null;
  function fmtDateShort(v){try{return new Date(v+'T00:00:00').toLocaleDateString('it-IT',{weekday:'short',day:'2-digit',month:'2-digit'});}catch(e){return v||'';}}
  function pClass(p){if(p.is_grab&&p.is_tp)return 'dual';if(p.is_grab)return 'grab';return 'tp';}
  function pTipo(p){if(p.is_grab&&p.is_tp)return 'Doppio';if(p.is_grab)return 'Grab';return 'TPoint';}
  function groups(){
    const out=[];
    PLAN.forEach((p,idx)=>{
      const key=p.date||'';
      let g=out.find(x=>x.date===key);
      if(!g){g={date:key,items:[]};out.push(g);}
      g.items.push({p,idx});
    });
    return out;
  }
  function orderedDateSlots(gs){return gs.map(g=>g.date).filter(Boolean).sort();}
  function applyBlockDates(order, slots){
    order.forEach((g,i)=>{
      const nd=slots[i]||g.date;
      g.items.forEach(it=>{it.p.date=nd;if(typeof dateLabel==='function')it.p.dateLabel=dateLabel(nd);if(typeof dateOnly==='function')it.p.dateOnly=dateOnly(nd);});
    });
  }
  function nearestDayBlockIndex(clientY){
    const blocks=[...document.querySelectorAll('.day-block')];
    if(!blocks.length)return null;
    let best=null,dist=Infinity;
    blocks.forEach(b=>{const r=b.getBoundingClientRect();const c=r.top+r.height/2;const d=Math.abs(clientY-c);if(d<dist){dist=d;best=Number(b.dataset.day);}});
    return Number.isFinite(best)?best:null;
  }
  window.changeBlockDate=function(dayIndex,newDate){
    const gs=groups();const g=gs[dayIndex];if(!g||!newDate)return;
    g.items.forEach(it=>{it.p.date=newDate;if(typeof dateLabel==='function')it.p.dateLabel=dateLabel(newDate);if(typeof dateOnly==='function')it.p.dateOnly=dateOnly(newDate);});
    render();
  };
  window.moveDayBlock=function(from,to){
    const gs=groups();
    if(from===to||from<0||to<0||from>=gs.length||to>=gs.length)return;
    const slots=orderedDateSlots(gs);
    const [g]=gs.splice(from,1);gs.splice(to,0,g);
    applyBlockDates(gs,slots);
    PLAN=[];gs.forEach(x=>x.items.forEach(it=>PLAN.push(it.p)));
    render();
  };
  window.dayBlockUp=function(i){moveDayBlock(i,i-1);};
  window.dayBlockDown=function(i){moveDayBlock(i,i+1);};
  window.dayTouchStart=function(e,i){dayTouchIndex=i;const b=e.currentTarget.closest('.day-block');if(b)b.classList.add('day-moving');e.preventDefault();};
  window.dayTouchMove=function(e){
    if(dayTouchIndex==null)return;
    const t=e.touches&&e.touches[0]?e.touches[0]:e;
    const j=nearestDayBlockIndex(t.clientY);
    if(j!=null&&j!==dayTouchIndex){moveDayBlock(dayTouchIndex,j);dayTouchIndex=j;}
    e.preventDefault();
  };
  window.dayTouchEnd=function(e){dayTouchIndex=null;document.querySelectorAll('.day-moving').forEach(x=>x.classList.remove('day-moving'));e.preventDefault();};
  window.render=function(){
    const box=document.getElementById('list');
    if(!PLAN.length){box.innerHTML='<div class="card muted">Nessun planning trovato. Torna alla pagina planning e crealo prima.</div>';return;}
    const gs=groups();let n=0;
    box.innerHTML='<div class="head-row"><div>↕</div><div>#</div><div>Data</div><div>PV</div><div>Comune</div><div>Via</div><div>Tipo</div><div>X</div></div>'+gs.map((g,di)=>{
      const head='<div class="day-block" data-day="'+di+'"><div class="day-controls"><button class="day-drag" ontouchstart="dayTouchStart(event,'+di+')" ontouchmove="dayTouchMove(event)" ontouchend="dayTouchEnd(event)" onpointerdown="dayTouchStart(event,'+di+')" onpointermove="dayTouchMove(event)" onpointerup="dayTouchEnd(event)">↕</button><button class="day-mini" onclick="dayBlockUp('+di+')">↑</button><button class="day-mini" onclick="dayBlockDown('+di+')">↓</button></div><div class="day-title">Giorno '+(di+1)+' · '+esc(fmtDateShort(g.date))+' <span>('+g.items.length+' PV)</span></div><input class="day-date" type="date" value="'+esc(g.date||'')+'" onchange="changeBlockDate('+di+',this.value)"></div>';
      const rows=g.items.map(it=>{const p=it.p,i=it.idx;n++;return '<div class="edit-row '+pClass(p)+'" data-i="'+i+'"><button class="drag" ontouchstart="touchStart(event,'+i+')" ontouchmove="touchMove(event)" ontouchend="touchEnd(event)" onmousedown="dragIndex='+i+'">↕</button><div class="idx">'+n+'</div><input class="date" type="date" value="'+esc(p.date||'')+'" onchange="changeDate('+i+',this.value)"><div class="col"><b>'+esc(p.pdv)+'</b></div><div class="col">'+esc(p.city||'')+'</div><div class="col">'+esc(p.address||'')+'</div><div class="col">'+esc(pTipo(p))+'</div><button class="btn bad x" onclick="del('+i+')">×</button></div>';}).join('');
      return head+rows;
    }).join('');
  };
  try{render();}catch(e){}
})();
</script>
'''


def main():
    path = DOCS_DIR / "planning-edit.html"
    if path.exists():
        html = path.read_text(encoding="utf-8")
        if 'day-block' not in html:
            html = html.replace('</head>', CSS + '\n</head>', 1)
            html = html.replace('</body>', JS + '\n</body>', 1)
        else:
            html = html.replace('el.dataset-day', 'el.dataset.day')
            marker = '<style>\n.day-block'
            s = html.find(marker)
            if s != -1:
                e = html.find('</style>', s)
                if e != -1:
                    html = html[:s] + CSS.strip() + html[e+len('</style>'):]
            if 'dayBlockUp=function' not in html:
                html = html.replace('</body>', JS + '\n</body>', 1)
        path.write_text(html, encoding="utf-8")
        print("Blocchi data trascinabili applicati")
    else:
        print("planning-edit.html non trovato, dayblock saltato")

    try:
        from planning_edit_persistence_patch import main as persistence_main
        persistence_main()
    except Exception as exc:
        raise RuntimeError(f"Errore persistenza planning modificato: {exc}") from exc


if __name__ == "__main__":
    main()
