from pathlib import Path

# Patch disattivato temporaneamente.
# Motivo: l'iniezione runtime dei controlli planning ha rotto la pagina principale.
# Il build continua a passare, ma non modifica più il sito principale.

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"


def main():
    print("Planning main runtime patch disattivato: sito principale lasciato invariato")


if __name__ == "__main__":
    main()
