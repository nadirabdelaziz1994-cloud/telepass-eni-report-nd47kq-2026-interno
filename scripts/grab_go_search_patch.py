from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

OLD_FILTER = """function grabFiltered(){const q=(document.getElementById('searchText')?.value||'').toLowerCase().trim(),agent=document.getElementById('grabAgentFilter')?.value||'',visit=document.getElementById('grabVisitFilter')?.value||'',tp=document.getElementById('grabTpFilter')?.value||'',stand=document.getElementById('grabStandFilter')?.value||'';"""
NEW_FILTER = """function grabFiltered(){const q=String(window.GRAB_SEARCH_TEXT??document.getElementById('grabSearch')?.value??document.getElementById('searchText')?.value??'').toLowerCase().trim(),agent=document.getElementById('grabAgentFilter')?.value||'',visit=document.getElementById('grabVisitFilter')?.value||'',tp=document.getElementById('grabTpFilter')?.value||'',stand=document.getElementById('grabStandFilter')?.value||'';"""

OLD_HTML = """let html=`<div class="card grab-filters"><div class="filter-main">"""
NEW_HTML = """let html=`<div class="card grab-filters"><div style="margin-bottom:10px;display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center"><input id="grabSearch" type="search" value="${esc(window.GRAB_SEARCH_TEXT??document.getElementById('grabSearch')?.value||'')}" onkeydown="if(event.key==='Enter'){window.GRAB_SEARCH_TEXT=this.value;renderGrabGo()}" placeholder="Cerca PV, città, indirizzo, agente, contratto, stand o note" style="width:100%;padding:11px 12px;border:1px solid var(--line);border-radius:12px;background:#fff;font-weight:700"><button class="btn light" type="button" onclick="const el=document.getElementById('grabSearch');window.GRAB_SEARCH_TEXT=el?el.value:'';renderGrabGo()">Cerca</button></div><div class="filter-main">"""


def patch_existing_search(html: str) -> str:
    # Se una versione vecchia della barra è già stata inserita, sostituiamo tutto il blocco.
    new_block = '<div style="margin-bottom:10px;display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center"><input id="grabSearch" type="search" value="${esc(window.GRAB_SEARCH_TEXT??document.getElementById(\'grabSearch\')?.value||\'\')}" onkeydown="if(event.key===\'Enter\'){window.GRAB_SEARCH_TEXT=this.value;renderGrabGo()}" placeholder="Cerca PV, città, indirizzo, agente, contratto, stand o note" style="width:100%;padding:11px 12px;border:1px solid var(--line);border-radius:12px;background:#fff;font-weight:700"><button class="btn light" type="button" onclick="const el=document.getElementById(\'grabSearch\');window.GRAB_SEARCH_TEXT=el?el.value:\'\';renderGrabGo()">Cerca</button></div>'
    pattern = r'<div style="margin-bottom:10px[^`]*?<input id="grabSearch"[^`]*?</div><div class="filter-main">'
    html2 = re.sub(pattern, new_block + '<div class="filter-main">', html, count=1, flags=re.S)
    return html2


def patch_html(html: str) -> str:
    changed = False

    html2 = html.replace(
        "const q=(document.getElementById('grabSearch')?.value||document.getElementById('searchText')?.value||'').toLowerCase().trim(),",
        "const q=String(window.GRAB_SEARCH_TEXT??document.getElementById('grabSearch')?.value??document.getElementById('searchText')?.value??'').toLowerCase().trim(),"
    )
    html2 = patch_existing_search(html2)
    if html2 != html:
        html = html2
        changed = True

    if "id=\"grabSearch\"" in html:
        return html if changed else html

    if OLD_FILTER in html:
        html = html.replace(OLD_FILTER, NEW_FILTER, 1)
        changed = True
    else:
        print("Attenzione: funzione grabFiltered non trovata o già diversa")

    if OLD_HTML in html:
        html = html.replace(OLD_HTML, NEW_HTML, 1)
        changed = True
    else:
        print("Attenzione: blocco filtri Grab & Go non trovato")

    return html if changed else html


def main():
    patched = 0
    for name in ["index.html", "Telepass_ENI_sito_v6.html"]:
        path = DOCS_DIR / name
        if not path.exists():
            continue
        old = path.read_text(encoding="utf-8")
        new = patch_html(old)
        if new != old:
            path.write_text(new, encoding="utf-8")
            patched += 1
    print(f"Grab & Go search patch completata: {patched} file aggiornati")


if __name__ == "__main__":
    main()
