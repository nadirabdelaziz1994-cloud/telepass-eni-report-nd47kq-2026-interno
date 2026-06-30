from pathlib import Path

# Patch disattivato.
# Motivo: la logica 15 giorni / 1 mese / tutti i PDV riduceva i PV trovati nel planning.
# Lasciamo il planning stabile precedente, senza modifiche extra.

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"


def main():
    print("Planning period/json patch disattivato: planning stabile lasciato invariato")


if __name__ == "__main__":
    main()
