from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"


def main():
    print("Planning period/json patch disattivato")
    import planning_roundtrip_desktop_patch
    planning_roundtrip_desktop_patch.main()


if __name__ == "__main__":
    main()
