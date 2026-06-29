from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

OLD = "function visibleAgentName(a){a=String(a||'').trim();if(!a)return '';if(a.includes('@')){const guessed=rawEmailToName(a);return AGENT_CANON[agentKey(guessed)]||guessed;}return a;}"

NEW = "function visibleAgentName(a){a=String(a||'').trim();if(!a)return '';const direct=norm(a);if(direct==='nadir a'||direct==='nadir abdel'||direct==='nadir abdel aziz'||direct==='nadir abdelaziz')return 'Nadir Abdel';if(a.includes('@')){const guessed=rawEmailToName(a);const g=norm(guessed);if(g==='nadir a'||g==='nadir abdel'||g==='nadir abdel aziz'||g==='nadir abdelaziz')return 'Nadir Abdel';return AGENT_CANON[agentKey(guessed)]||guessed;}return a;}"


def main():
    path = DOCS_DIR / "planning.html"
    if not path.exists():
        print("planning.html non trovato, patch saltata")
        return
    html = path.read_text(encoding="utf-8")
    html = html.replace("<!-- Normalize Nadir agent variants -->", "")
    start = html.find("<script>\n(function(){\n  function nNorm(s){")
    if start != -1:
        end = html.find("</script>", start)
        if end != -1:
            html = html[:start] + html[end + len("</script>"):]
    if OLD in html:
        html = html.replace(OLD, NEW, 1)
        path.write_text(html, encoding="utf-8")
        print("Nadir alias applicato dentro visibleAgentName")
    elif "direct==='nadir a'" in html:
        print("Nadir alias già presente")
    else:
        print("Funzione visibleAgentName non trovata: nessuna modifica")


if __name__ == "__main__":
    main()
