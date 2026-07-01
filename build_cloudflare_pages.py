import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def run(cmd):
    print('> ' + ' '.join(cmd))
    subprocess.check_call(cmd, cwd=ROOT)


# MODALITA' RIPRISTINO STABILE
# Per riportare online il vecchio sito, generiamo solo la dashboard principale pulita.
# Le patch extra più recenti restano disattivate: Bundle avanzato, Grab&Go patch, Planning patch.
# Appena il nuovo sito è pronto, queste funzioni verranno ricostruite lì in modo separato.
run(['python', 'build_github.py'])

print('Build Cloudflare Pages completata in modalita ripristino. Cartella output: docs')
