from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"


def main():
    path = DOCS_DIR / "pdv-manage.html"
    if not path.exists():
        print("pdv-manage.html non trovato, cleanup saltato")
        return
    html = path.read_text(encoding="utf-8")
    start_tag = '<script id="planning-data" type="application/json">'
    end_tag = '</script>'
    s = html.find(start_tag)
    e = html.find(end_tag, s)
    if s != -1 and e != -1:
        before = html[:s].replace('\\"', '"')
        data = html[s:e+len(end_tag)]
        after = html[e+len(end_tag):].replace('\\"', '"')
        html = before + data + after
    else:
        html = html.replace('\\"', '"')
    path.write_text(html, encoding="utf-8")
    print("Cleanup HTML gestione PV applicato")


if __name__ == "__main__":
    main()
