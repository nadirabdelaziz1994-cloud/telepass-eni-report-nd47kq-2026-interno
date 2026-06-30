from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = ROOT / "docs"

PLANNING_BUTTON = '<button data-planning-link="1" onclick="location.href=\'./planning.html?v=from-main\'">Planning automatico</button>'


def patch_html(html: str) -> str:
    # Patch sicura: aggiunge solo il pulsante Planning automatico nel menu.
    # Niente sync Cloudflare nel sito principale finché non siamo sicuri del Worker.
    if 'data-planning-link' in html:
        return html

    if '<button data-page="grab-go" onclick="showPage(\'grab-go\', this)">Grab & Go</button>' in html:
        return html.replace(
            '<button data-page="grab-go" onclick="showPage(\'grab-go\', this)">Grab & Go</button>',
            '<button data-page="grab-go" onclick="showPage(\'grab-go\', this)">Grab & Go</button>\n      ' + PLANNING_BUTTON,
            1,
        )

    return html.replace(
        '<button data-page="file-utili" onclick="showPage(\'file-utili\', this)">File utili</button>',
        PLANNING_BUTTON + '\n      <button data-page="file-utili" onclick="showPage(\'file-utili\', this)">File utili</button>',
        1,
    )


def main():
    done = 0
    for name in ['index.html', 'Telepass_ENI_sito_v6.html']:
        path = DOCS_DIR / name
        if path.exists():
            path.write_text(patch_html(path.read_text(encoding='utf-8')), encoding='utf-8')
            done += 1
    print(f'Pulsante Planning automatico applicato: {done} file')


if __name__ == '__main__':
    main()
