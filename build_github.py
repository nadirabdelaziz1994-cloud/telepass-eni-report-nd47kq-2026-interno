from pathlib import Path
import subprocess
import sys
import shutil

ROOT = Path(__file__).resolve().parent
LISTA_DIR = ROOT / 'input' / 'lista'
ANAG_DIR = ROOT / 'input' / 'anagrafica'
REPORT_DIR = ROOT / 'input' / 'report_settimanali'
OUT_DIR = ROOT / 'docs'


def first_xlsx(folder: Path):
    files = [p for p in folder.rglob('*.xlsx') if not p.name.startswith('~$')]
    if not files:
        return None
    return sorted(files, key=lambda p: p.stat().st_mtime)[-1]


def main():
    lista = first_xlsx(LISTA_DIR)
    anag = first_xlsx(ANAG_DIR)
    if not lista:
        raise SystemExit("ERRORE: manca il file Lista PDV nella cartella input/lista")
    if not anag:
        raise SystemExit("ERRORE: manca il file Anagrafica nella cartella input/anagrafica")
    if not REPORT_DIR.exists() or not any(REPORT_DIR.rglob('*.xlsx')):
        raise SystemExit("ERRORE: non ci sono report settimana dentro input/report_settimanali")

    OUT_DIR.mkdir(exist_ok=True, parents=True)
    script = ROOT / 'aggiorna_dashboard_github.py'
    if not script.exists():
        raise SystemExit("ERRORE: manca aggiorna_dashboard_github.py")

    cmd = [sys.executable, str(script), str(lista), str(anag), str(REPORT_DIR), str(OUT_DIR)]
    print('Eseguo:', ' '.join(cmd))
    subprocess.run(cmd, check=True)

    generated = OUT_DIR / 'Telepass_ENI_sito_v6.html'
    if generated.exists():
        shutil.copy2(generated, OUT_DIR / 'index.html')

    (OUT_DIR / '.nojekyll').write_text('', encoding='utf-8')
    print('Build completata. Apri docs/index.html o pubblica docs con GitHub Pages.')


if __name__ == '__main__':
    main()
