from pathlib import Path

ROOT = Path(__file__).resolve().parent
DOCS = ROOT / 'docs'
DOCS.mkdir(exist_ok=True)

html = '''<!doctype html>
<html lang="it">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MyWorld - Telepass Eni</title>
  <style>
    :root{--bg:#f4f7fb;--card:#fff;--ink:#10243e;--muted:#607089;--line:#d7dfeb;--blue:#0f2746;--blue2:#0d6efd}
    *{box-sizing:border-box}
    body{margin:0;font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--ink)}
    .top{background:var(--blue);color:#fff;padding:18px 20px;box-shadow:0 2px 12px rgba(0,0,0,.16)}
    .wrap{max-width:980px;margin:0 auto}
    .brand{font-size:24px;font-weight:950}
    main{max-width:980px;margin:0 auto;padding:18px}
    .card{background:var(--card);border:1px solid var(--line);border-radius:20px;box-shadow:0 8px 24px rgba(16,36,62,.1);padding:20px;margin-bottom:14px}
    h1,h2{margin:0 0 10px}p{color:var(--muted);line-height:1.45}a.btn{display:inline-flex;background:var(--blue2);color:white;text-decoration:none;border-radius:14px;padding:12px 16px;font-weight:900;margin-top:8px}
    .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.mini{background:#f8fbff;border:1px solid var(--line);border-radius:16px;padding:14px}.mini b{display:block;font-size:20px}.mini span{color:var(--muted);font-size:13px}
    @media(max-width:700px){.grid{grid-template-columns:1fr}.brand{font-size:20px}}
  </style>
</head>
<body>
  <header class="top"><div class="wrap"><div class="brand">MyWorld - Telepass Eni</div></div></header>
  <main>
    <section class="card">
      <h1>Sito vecchio in modalità recupero</h1>
      <p>Ho disattivato temporaneamente le funzioni che stavano rompendo la pagina. Stiamo rifacendo il sito nuovo in modo pulito su MyWorld-Report.</p>
      <p>Questa pagina serve solo a rimettere online il link vecchio senza mostrare codice rotto.</p>
      <a class="btn" href="https://myworld-report.pages.dev">Apri il nuovo sito</a>
    </section>
    <section class="grid">
      <div class="mini"><b>Home</b><span>Da ricostruire nel nuovo sito</span></div>
      <div class="mini"><b>Classifica</b><span>Da importare con dati veri</span></div>
      <div class="mini"><b>Planning</b><span>Da rifare separato e stabile</span></div>
    </section>
  </main>
</body>
</html>'''

(DOCS / 'index.html').write_text(html, encoding='utf-8')
(DOCS / '.nojekyll').write_text('', encoding='utf-8')
print('Modalità emergenza attiva. Output: docs/index.html')
