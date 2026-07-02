from pathlib import Path

ROOT = Path(__file__).resolve().parent
DOCS = ROOT / 'docs'
DOCS.mkdir(exist_ok=True)

u = 'https://' + 'myworld-report.pages.dev' + '/'
html = f'<html><head><meta charset="utf-8"><title>MyWorld Report</title></head><body style="font-family:Arial;text-align:center;padding:60px"><h1>MyWorld Report</h1><p>Versione aggiornata disponibile.</p><p><a href="{u}">Apri report</a></p></body></html>'

(DOCS / 'index.html').write_text(html, encoding='utf-8')
(DOCS / '.nojekyll').write_text('', encoding='utf-8')
print('Output: docs/index.html')
