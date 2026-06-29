from pathlib import Path
import csv
import datetime
import json
import re
import shutil
import unicodedata

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
GRABEGO_DIR = ROOT / "input" / "grabego"
LISTA_DIR = ROOT / "input" / "lista"
DOCS_DIR = ROOT / "docs"


def norm_header(value):
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").replace("\xa0", " ")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r"[^A-Z0-9%]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def norm_pdv(value):
    nums = re.findall(r"\d+", str(value or ""))
    return nums[0].zfill(5) if nums else ""


def cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    val = row[idx]
    return "" if val is None else str(val).strip()


def find_col(headers, *needles):
    for i, h in enumerate(headers):
        for n in needles:
            if n and n in h:
                return i
    return None


def find_col_exact(headers, *needles):
    wanted = {norm_header(n) for n in needles if n}
    for i, h in enumerate(headers):
        if h in wanted:
            return i
    return None


def latest_file(folder, suffixes=(".xlsx", ".csv")):
    if not folder.exists():
        return None
    files = [p for p in folder.rglob("*") if p.suffix.lower() in suffixes and not p.name.startswith("~$")]
    if not files:
        return None
    preferred = [p for p in files if p.stem.upper().replace(" ", "").replace("_", "").replace("-", "").startswith("GRABEGO")]
    return sorted(preferred or files, key=lambda p: p.stat().st_mtime)[-1]


def load_lista_pdv():
    path = latest_file(LISTA_DIR, (".xlsx",))
    if not path:
        return set()
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    result = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        pdv = norm_pdv(row[0] if row else "")
        if pdv:
            result.add(pdv)
    return result


def agent_from_email(value):
    text = str(value or "").strip()
    if "@" not in text:
        return ""
    local = re.sub(r"[._-]+", " ", text.split("@", 1)[0])
    local = re.sub(r"\s+", " ", local).strip()
    return " ".join(part.capitalize() for part in local.split())


def read_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["PERIMETRO"] if "PERIMETRO" in wb.sheetnames else wb[wb.sheetnames[0]]
    header_row = 1
    for r in range(1, min(ws.max_row, 15) + 1):
        headers = [norm_header(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if any(h in {"PV", "PDV", "PV ENI"} for h in headers) and any("AGENTE" in h for h in headers):
            header_row = r
            break
    headers = [norm_header(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
    rows = ws.iter_rows(min_row=header_row + 1, values_only=True)
    return headers, rows


def read_csv(path):
    text = path.read_text(encoding="utf-8-sig")
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=";,\t|")
    except Exception:
        dialect = csv.excel
        dialect.delimiter = ";"
    raw_rows = list(csv.reader(text.splitlines(), dialect))
    header_idx = 0
    for i, candidate in enumerate(raw_rows[:15]):
        headers = [norm_header(v) for v in candidate]
        if any(h in {"PV", "PDV", "PV ENI"} for h in headers) and any("AGENTE" in h for h in headers):
            header_idx = i
            break
    return [norm_header(v) for v in raw_rows[header_idx]], raw_rows[header_idx + 1:]


def load_grab_go():
    path = latest_file(GRABEGO_DIR)
    lista = load_lista_pdv()
    if not path:
        return {
            "kind": "grab_go",
            "missing_file": True,
            "source_name": "",
            "source_path": "",
            "updated_at": "",
            "rows": [],
            "summary": {"pdv_count": 0, "telepass_point": 0, "not_telepass_point": 0, "agents": 0},
        }

    headers, raw_rows = read_csv(path) if path.suffix.lower() == ".csv" else read_xlsx(path)
    cols = {
        "pdv": find_col_exact(headers, "PV", "PDV", "PV ENI") or find_col(headers, "PV", "PDV"),
        "contract": find_col(headers, "CONTRATTO"),
        "address": find_col(headers, "INDIRIZZO"),
        "city": find_col(headers, "CITTA"),
        "region": find_col(headers, "REGIONE"),
        "agent_email": find_col(headers, "AGENTE"),
        "branch": find_col(headers, "FILIALE"),
        "note": find_col_exact(headers, "NOTE"),
        "training": find_col(headers, "FORMAZIONE"),
        "order": find_col(headers, "ORDINE"),
        "mail_gestore": find_col(headers, "MAIL GESTORE"),
        "pec_gestore": find_col(headers, "PEC GESTORE"),
        "mail_rzv": find_col(headers, "MAIL RZV"),
        "mail_cr": find_col(headers, "MAIL CR"),
        "stand_requested": find_col(headers, "STAND RICHIESTO"),
        "stand": find_col_exact(headers, "STAND"),
        "gestore01": find_col(headers, "GESTORE01"),
        "note_telepass": find_col(headers, "NOTE TELEPASS"),
        "lat": find_col(headers, "LATITUDINE", "LATITUDE", "LAT"),
        "lng": find_col(headers, "LONGITUDINE", "LONGITUDE", "LON", "LNG"),
    }

    rows = []
    seen = set()
    for row in raw_rows:
        pdv = norm_pdv(cell(row, cols["pdv"]))
        if not pdv or pdv in seen:
            continue
        seen.add(pdv)
        mail = cell(row, cols["agent_email"])
        rows.append({
            "pdv": pdv,
            "contract": cell(row, cols["contract"]),
            "address": cell(row, cols["address"]),
            "city": cell(row, cols["city"]),
            "region": cell(row, cols["region"]),
            "agent_email": mail,
            "agent": agent_from_email(mail),
            "branch": cell(row, cols["branch"]),
            "note": cell(row, cols["note"]),
            "training": cell(row, cols["training"]),
            "order": cell(row, cols["order"]),
            "mail_gestore": cell(row, cols["mail_gestore"]),
            "pec_gestore": cell(row, cols["pec_gestore"]),
            "mail_rzv": cell(row, cols["mail_rzv"]),
            "mail_cr": cell(row, cols["mail_cr"]),
            "stand_requested": cell(row, cols["stand_requested"]),
            "stand": cell(row, cols["stand"]),
            "gestore01": cell(row, cols["gestore01"]),
            "note_telepass": cell(row, cols["note_telepass"]),
            "telepass_point": pdv in lista,
            "lat": cell(row, cols["lat"]),
            "lng": cell(row, cols["lng"]),
        })
    rows.sort(key=lambda r: (r.get("agent") or "ZZZ", r.get("region") or "", r.get("city") or "", r.get("pdv") or ""))

    dest_dir = DOCS_DIR / "files" / "GRABEGO"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_name = "GRABEGO" + path.suffix.lower()
    shutil.copy2(path, dest_dir / dest_name)

    return {
        "kind": "grab_go",
        "missing_file": False,
        "source_name": path.name,
        "source_path": f"files/GRABEGO/{dest_name}",
        "updated_at": datetime.datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
        "rows": rows,
        "summary": {
            "pdv_count": len(rows),
            "telepass_point": sum(1 for r in rows if r["telepass_point"]),
            "not_telepass_point": sum(1 for r in rows if not r["telepass_point"]),
            "agents": len({r["agent"] for r in rows if r["agent"]}),
        },
    }


CSS = """.grab-filters{margin-bottom:10px}.grab-agent-card{margin-top:12px}.grab-table{min-width:1280px}.grab-table select{width:150px;padding:7px;border-radius:10px;border:1px solid var(--line);background:#fff}.grab-green{box-shadow:inset 5px 0 0 var(--green)}.grab-red{box-shadow:inset 5px 0 0 var(--red)}.grab-yellow{box-shadow:inset 5px 0 0 #d6a500}.grab-agent-dot{display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:50%;background:#0f2746;color:#fff;font-weight:900}.grab-map-card{margin-top:10px}.grab-map-card iframe{width:100%;height:320px;border:0;border-radius:14px;margin-top:10px;background:#eef4fb}"""

JS = r'''
const GRAB_GO_MONTHS = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
function grabGoRows(){ return (APP.grab_go && APP.grab_go.rows) ? APP.grab_go.rows : []; }
function grabGoState(){ try{return JSON.parse(localStorage.getItem('grabGoVisits')||'{}')||{};}catch(e){return {};} }
function saveGrabGoState(s){ localStorage.setItem('grabGoVisits', JSON.stringify(s||{})); }
function grabGoInitials(a){ return String(a||'?').split(/\s+/).filter(Boolean).slice(0,2).map(x=>x[0]).join('').toUpperCase()||'?'; }
function setGrabGoVisit(pdv, month){ const s=grabGoState(); if(!month){delete s[pdv];}else{s[pdv]={month:Number(month),year:new Date().getFullYear(),saved_at:new Date().toISOString()};} saveGrabGoState(s); renderGrabGo(); }
function grabGoVisitInfo(pdv){ const x=grabGoState()[pdv]; if(!x||!x.month)return{className:'grab-red',label:'Non visitato',rank:2}; const now=new Date(),m=Number(x.month),y=Number(x.year||now.getFullYear()); const diff=(now.getFullYear()-y)*12+((now.getMonth()+1)-m); const label=`Visitato: ${GRAB_GO_MONTHS[m-1]||m} ${y}`; if(diff>4)return{className:'grab-yellow',label:'Visitato più di 4 mesi fa · '+label,rank:1}; return{className:'grab-green',label,rank:0}; }
function grabGoMapsQuery(r){ return `${r.address||''} ${r.city||''} ${r.region||''}`.trim() || r.pdv; }
function openGrabGoMap(q){ const f=document.getElementById('grabGoMapFrame'),l=document.getElementById('grabGoMapLabel'); if(!f)return; f.src='https://maps.google.com/maps?q='+encodeURIComponent(q)+'&output=embed'; if(l)l.textContent=q; }
function grabGoFilteredRows(){ const q=(document.getElementById('searchText')?.value||'').toLowerCase().trim(),agent=document.getElementById('agentFilter')?.value||'',visit=document.getElementById('grabGoVisitFilter')?.value||'',tp=document.getElementById('grabGoTelepassFilter')?.value||'',stand=document.getElementById('grabGoStandFilter')?.value||''; return grabGoRows().filter(r=>{ if(agent&&r.agent!==agent)return false; if(tp==='yes'&&!r.telepass_point)return false; if(tp==='no'&&r.telepass_point)return false; if(stand&&String(r.stand||'')!==stand)return false; const vi=grabGoVisitInfo(r.pdv); if(visit==='not'&&vi.rank!==2)return false; if(visit==='recent'&&vi.rank!==0)return false; if(visit==='old'&&vi.rank!==1)return false; if(q){const hay=`${r.pdv} ${r.city} ${r.address} ${r.region} ${r.agent} ${r.agent_email} ${r.contract} ${r.stand} ${r.note_telepass} ${r.mail_gestore}`.toLowerCase(); if(!hay.includes(q))return false;} return true; }); }
function grabGoOptionSelected(v,s){ return String(v)===String(s)?'selected':''; }
function grabGoSelectOptions(values, placeholder, selected){ return `<option value="">${placeholder}</option>`+values.map(v=>`<option value="${esc(v)}" ${String(v)===String(selected)?'selected':''}>${esc(v)}</option>`).join(''); }
function renderGrabGoMap(rows){ if(!rows.length)return '<div class="empty">Nessun punto vendita da mostrare in mappa con i filtri scelti.</div>'; const q=grabGoMapsQuery(rows[0]); return `<div class="card grab-map-card"><div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;flex-wrap:wrap"><div><h3 style="text-align:left;margin-bottom:4px">Mappa interattiva</h3><div class="small-muted">Mostro il PDV selezionato su Google Maps. I colori visita sono nella lista sotto.</div><div class="small-muted" style="margin-top:4px"><b>Ora in mappa:</b> <span id="grabGoMapLabel">${esc(q)}</span></div></div><a class="btn light" target="_blank" href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(q)}">Apri su Google Maps</a></div><iframe id="grabGoMapFrame" title="Mappa Grab & Go" loading="lazy" src="https://maps.google.com/maps?q=${encodeURIComponent(q)}&output=embed"></iframe></div>`; }
function renderGrabGo(){ const wrap=document.getElementById('grabGoWrap'); if(!wrap)return; const rep=APP.grab_go||{rows:[],summary:{}}; if(rep.missing_file){wrap.innerHTML='<div class="empty">Nessun file GRABEGO.xlsx trovato in input/grabego.</div>';return;} const currentVisit=document.getElementById('grabGoVisitFilter')?.value||'',currentTp=document.getElementById('grabGoTelepassFilter')?.value||'',currentStand=document.getElementById('grabGoStandFilter')?.value||''; const all=grabGoRows(),standValues=[...new Set(all.map(r=>r.stand).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'it')),rows=grabGoFilteredRows(); const tpCount=rows.filter(r=>r.telepass_point).length,stateCounts=rows.reduce((a,r)=>{const v=grabGoVisitInfo(r.pdv).rank;a[v]=(a[v]||0)+1;return a;},{}),byAgent=new Map(); rows.forEach(r=>{const k=r.agent||'Senza agente'; if(!byAgent.has(k))byAgent.set(k,[]); byAgent.get(k).push(r);}); const groups=[...byAgent.entries()].sort((a,b)=>a[0].localeCompare(b[0],'it')); const filterHtml=`<div class="card grab-filters"><div class="filter-main"><select id="grabGoVisitFilter" onchange="renderGrabGo()"><option value="">Tutti gli stati visita</option><option value="not" ${grabGoOptionSelected('not',currentVisit)}>Non visitati</option><option value="recent" ${grabGoOptionSelected('recent',currentVisit)}>Visitati ultimi 4 mesi</option><option value="old" ${grabGoOptionSelected('old',currentVisit)}>Visitati da più di 4 mesi</option></select><select id="grabGoTelepassFilter" onchange="renderGrabGo()"><option value="">Tutti TP / non TP</option><option value="yes" ${grabGoOptionSelected('yes',currentTp)}>Già Telepass Point</option><option value="no" ${grabGoOptionSelected('no',currentTp)}>Non Telepass Point</option></select><select id="grabGoStandFilter" onchange="renderGrabGo()">${grabGoSelectOptions(standValues,'Tutti gli stand',currentStand)}</select><a class="btn light" href="${esc(rep.source_path||'#')}" download>Scarica GRABEGO</a></div><div class="small-muted" style="margin-top:8px">Fonte: ${esc(rep.source_name||'GRABEGO.xlsx')} · Aggiornato: ${esc(rep.updated_at||'')} · Usa anche i filtri globali in alto: agente e ricerca libera.</div></div>`; const kpis=`<div class="metric-row sost" style="margin-top:10px"><div class="metric-card"><h4>PDV Grab & Go filtrati</h4><div class="metric-big">${fmtNum(rows.length)}</div><div class="metric-sub">Totali file: ${fmtNum(all.length)}</div></div><div class="metric-card tone-good"><h4>Già Telepass Point</h4><div class="metric-big">${fmtNum(tpCount)}</div><div class="metric-sub">Segnalati in chiaro</div></div><div class="metric-card tone-bad"><h4>Non visitati</h4><div class="metric-big">${fmtNum(stateCounts[2]||0)}</div><div class="metric-sub">Rosso</div></div></div>`; const monthOptions=sel=>'<option value="">Non visitato</option>'+GRAB_GO_MONTHS.map((m,i)=>`<option value="${i+1}" ${Number(sel)===(i+1)?'selected':''}>${m}</option>`).join(''); const groupsHtml=groups.map(([agent,items])=>{ const body=items.map(r=>{ const vi=grabGoVisitInfo(r.pdv),state=grabGoState()[r.pdv]||{},query=grabGoMapsQuery(r),queryArg=JSON.stringify(query); return `<tr class="${vi.className}"><td><span class="grab-agent-dot">${esc(grabGoInitials(r.agent))}</span></td><td><b>${esc(r.pdv)}</b></td><td><div class="city-cell"><div class="city-main">${esc(r.city)}</div>${r.address?'<div class="city-address">'+esc(r.address)+'</div>':''}</div></td><td>${r.telepass_point?'<span class="badge bene">Già Telepass Point</span>':'<span class="badge male">Non TP</span>'}</td><td>${esc(r.contract||'')}</td><td>${esc(r.stand||'')}</td><td>${esc(r.branch||'')}</td><td>${esc(r.note_telepass||r.note||'')}</td><td><select onchange="setGrabGoVisit('${r.pdv}', this.value)">${monthOptions(state.month)}</select><div class="small-muted">${esc(vi.label)}</div></td><td><button class="btn light small" onclick='openGrabGoMap(${queryArg})'>Mappa</button><a class="btn ghost small" target="_blank" href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}">Google</a></td></tr>`; }).join(''); return `<div class="card grab-agent-card"><div style="display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap"><div class="section-title" style="margin:0">${esc(agent)}</div><div class="small-muted">${fmtNum(items.length)} PDV · ${fmtNum(items.filter(x=>x.telepass_point).length)} già TP</div></div><div class="list-wrap" style="margin-top:10px"><table class="grab-table"><thead><tr><th></th><th>PDV</th><th>Città / indirizzo</th><th>Telepass Point</th><th>Contratto</th><th>Stand</th><th>Filiale</th><th>Note</th><th>Visita</th><th>Mappa</th></tr></thead><tbody>${body}</tbody></table></div></div>`; }).join(''); wrap.innerHTML=filterHtml+kpis+renderGrabGoMap(rows)+(groupsHtml||'<div class="empty">Nessun PDV Grab & Go con i filtri scelti.</div>'); }
'''


def patch_html(html, data):
    if "APP.grab_go" in html:
        return html
    html = html.replace("const DATA = APP.rows || [];", f"APP.grab_go = {json.dumps(data, ensure_ascii=False)};\nconst DATA = APP.rows || [];", 1)
    html = html.replace(".export-note{font-size:12px;color:var(--muted);margin-top:8px;text-align:right}", ".export-note{font-size:12px;color:var(--muted);margin-top:8px;text-align:right}\n" + CSS, 1)
    html = html.replace("<button data-page=\"file-utili\" onclick=\"showPage('file-utili', this)\">File utili</button>", "<button data-page=\"grab-go\" onclick=\"showPage('grab-go', this)\">Grab & Go</button>\n      <button data-page=\"file-utili\" onclick=\"showPage('file-utili', this)\">File utili</button>", 1)
    html = html.replace("  <section id=\"page-file-utili\" class=\"page\">", "  <section id=\"page-grab-go\" class=\"page\">\n    <div class=\"section-title\">Grab & Go</div>\n    <div id=\"grabGoWrap\"></div>\n  </section>\n\n  <section id=\"page-file-utili\" class=\"page\">", 1)
    html = html.replace("function uniq(k){ return [...new Set(DATA.map(r=>r[k]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'it')); }", "function uniq(k){ const vals=DATA.map(r=>r[k]); if(k==='agent' && APP.grab_go && APP.grab_go.rows){ vals.push(...APP.grab_go.rows.map(r=>r.agent)); } return [...new Set(vals.filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'it')); }", 1)
    html = html.replace("function renderFiles(){", JS + "\nfunction renderFiles(){", 1)
    html = html.replace("if(page==='file-utili') renderFiles();", "if(page==='grab-go') renderGrabGo();\n  if(page==='file-utili') renderFiles();", 1)
    html = html.replace("if(activePage()==='gare-pdv') renderGarePdv();", "if(activePage()==='gare-pdv') renderGarePdv();\n  if(activePage()==='grab-go') renderGrabGo();", 1)
    html = html.replace("renderGarePdv();\nrenderGare('gareAgentiWrap',APP.gare_agenti,'Nessuna gara in corso');", "renderGarePdv();\nrenderGrabGo();\nrenderGare('gareAgentiWrap',APP.gare_agenti,'Nessuna gara in corso');", 1)
    return html


def main():
    data = load_grab_go()
    for name in ["index.html", "Telepass_ENI_sito_v6.html"]:
        path = DOCS_DIR / name
        if not path.exists():
            continue
        path.write_text(patch_html(path.read_text(encoding="utf-8"), data), encoding="utf-8")
    print(f"Grab & Go patch completata: {len(data.get('rows', []))} PDV")


if __name__ == "__main__":
    main()
