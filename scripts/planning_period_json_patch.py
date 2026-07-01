from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"


def main():
    print("Planning period/json patch disattivato: generatore planning già corretto")
    import planning_edit_ui_patch
    # Questo import non rilancia la patch UI; serve solo a non lasciare riferimenti rotti in vecchi build.
    try:
        import planning_month_plus_10_patch
        planning_month_plus_10_patch.main()
    except Exception as exc:
        print("Patch mese+10 non applicata su HTML generato:", exc)
    try:
        import planning_clean_colleague_ui_patch
        planning_clean_colleague_ui_patch.main()
    except Exception as exc:
        print("Pulizia UI planning colleghi non applicata:", exc)
    try:
        import planning_admin_link_capacity_patch
        planning_admin_link_capacity_patch.main()
    except Exception as exc:
        print("Pulsante admin/capacità planning non applicati:", exc)


if __name__ == "__main__":
    main()
