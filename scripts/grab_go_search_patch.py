from pathlib import Path

# Patch disattivato per release stabile.
# Motivo: la barra ricerca Grab & Go rompeva il caricamento dati in alcune versioni mobile/browser.
# Lasciamo il sito senza questa modifica per consegnare un link stabile ai colleghi.

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"


def main():
    print("Grab & Go search patch disattivato: sito lasciato stabile")


if __name__ == "__main__":
    main()
