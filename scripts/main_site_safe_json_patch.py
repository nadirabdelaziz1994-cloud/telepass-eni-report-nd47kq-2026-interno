from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def safe_payload_expr(module_name: str) -> str:
    return f'''payload = {module_name}.json.dumps(data, ensure_ascii=False).replace("</", "<\\/").replace("\\u2028", "\\\\u2028").replace("\\u2029", "\\\\u2029")
    return tpl.replace("__DATA_JSON__", payload).replace("__CURRENT_WEEK__", f"{{data['meta']['current_week']:02d}}")'''


def patch_template_dashboard():
    path = ROOT / "template_dashboard.html"
    text = path.read_text(encoding="utf-8")
    changed = False

    # Sposta il JSON fuori dallo script JS principale. Così, anche se un dato contiene testo strano,
    # non può tagliare a metà lo script e far vedere codice JavaScript in pagina.
    if '<script id="app-data" type="application/json">__DATA_JSON__</script>' not in text:
        text = text.replace(
            '<script>\nconst APP = __DATA_JSON__;',
            '<script id="app-data" type="application/json">__DATA_JSON__</script>\n<script>\nconst APP = JSON.parse(document.getElementById("app-data").textContent || "{}");',
            1,
        )
        changed = True
    else:
        text = text.replace(
            'const APP = __DATA_JSON__;',
            'const APP = JSON.parse(document.getElementById("app-data").textContent || "{}");',
        )

    if changed:
        path.write_text(text, encoding="utf-8")
        print("Template dashboard isolato: JSON fuori dallo script principale")
    else:
        print("Template dashboard già isolato")


def patch_dashboard_github():
    path = ROOT / "aggiorna_dashboard_github.py"
    text = path.read_text(encoding="utf-8")
    old = 'return tpl.replace("__DATA_JSON__", base.json.dumps(data, ensure_ascii=False)).replace("__CURRENT_WEEK__", f"{data[\'meta\'][\'current_week\']:02d}")'
    new = safe_payload_expr("base")
    if old in text:
        text = text.replace(old, new, 1)
        path.write_text(text, encoding="utf-8")
        print("Safe JSON applicato ad aggiorna_dashboard_github.py")
    elif 'payload = base.json.dumps(data, ensure_ascii=False).replace("</", "<\\/")' in text:
        print("Safe JSON già presente in aggiorna_dashboard_github.py")
    else:
        print("ATTENZIONE: pattern build_html non trovato in aggiorna_dashboard_github.py")


def patch_dashboard_base():
    path = ROOT / "aggiorna_dashboard.py"
    if not path.exists():
        return
    text = path.read_text(encoding="utf-8")
    old = 'return tpl.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False)).replace("__CURRENT_WEEK__", f"{data[\'meta\'][\'current_week\']:02d}")'
    new = safe_payload_expr("json")
    if old in text:
        text = text.replace(old, new, 1)
        path.write_text(text, encoding="utf-8")
        print("Safe JSON applicato ad aggiorna_dashboard.py")
    elif 'payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")' in text:
        print("Safe JSON già presente in aggiorna_dashboard.py")


def patch_grab_go_patch():
    path = ROOT / "scripts" / "grab_go_patch.py"
    if not path.exists():
        return
    text = path.read_text(encoding="utf-8")
    old = 'html = html.replace("const DATA = APP.rows || [];", f"APP.grab_go = {json.dumps(data, ensure_ascii=False)};\\nconst DATA = APP.rows || [];", 1)'
    new = '''payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/").replace("\\u2028", "\\\\u2028").replace("\\u2029", "\\\\u2029")
    html = html.replace("const DATA = APP.rows || [];", f"APP.grab_go = {payload};\\nconst DATA = APP.rows || [];", 1)'''
    if old in text:
        text = text.replace(old, new, 1)
        path.write_text(text, encoding="utf-8")
        print("Safe JSON applicato a grab_go_patch.py")
    elif 'APP.grab_go = {payload}' in text:
        print("Safe JSON già presente in grab_go_patch.py")


def main():
    patch_template_dashboard()
    patch_dashboard_github()
    patch_dashboard_base()
    patch_grab_go_patch()


if __name__ == "__main__":
    main()
