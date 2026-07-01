from pathlib import Path
import datetime
import os
import re
import shutil
import unicodedata

import aggiorna_dashboard as base


def normalize_header(value):
    if value is None:
        return ""
    s = str(value).replace("\r", " ").replace("\n", " ").replace("\xa0", " ")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper()
    s = s.replace("–", "-").replace("—", "-").replace("’", "'").replace("`", "'")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _preferred_total(report_value, fallback_value):
    report_num = base.safe_num(report_value)
    fallback_num = base.safe_num(fallback_value)
    if report_num is None:
        return fallback_num or 0
    if abs(report_num) < 1e-9 and fallback_num not in (None, 0):
        return fallback_num
    return report_num


def parse_report_dynamic(path):
    year, week, sheet = base.extract_week_year(path)
    wb = base.openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]

    headers = [ws.cell(4, c).value for c in range(1, 90)]
    norm_headers = [normalize_header(h) for h in headers]

    def find_exact(txt, after=0):
        target = normalize_header(txt)
        for i, h in enumerate(norm_headers, start=1):
            if i <= after:
                continue
            if h == target:
                return i
        return None

    def find_contains(txt, after=0):
        target = normalize_header(txt)
        for i, h in enumerate(norm_headers, start=1):
            if i <= after:
                continue
            if target and target in h:
                return i
        return None

    def find_tokens(tokens, after=0):
        wanted = [normalize_header(tok) for tok in tokens if tok]
        for i, h in enumerate(norm_headers, start=1):
            if i <= after:
                continue
            if not h:
                continue
            if all(tok in h for tok in wanted):
                return i
        return None

    def first_match(candidates, after=0):
        for cand in candidates:
            if isinstance(cand, (list, tuple)):
                col = find_tokens(cand, after=after)
            else:
                col = find_exact(cand, after=after) or find_contains(cand, after=after)
            if col:
                return col
        return None

    week_2 = f"{int(week):02d}" if week is not None else ""
    week_1 = f"{int(week)}" if week is not None else ""

    c_pdv = first_match(["PV ENI", ["PV", "ENI"], ["PUNTO VENDITA", "ENI"]])
    c_area = first_match(["AREA COMM.", "AREA COMM", "AREA COMMERCIALE"])
    c_reg = first_match(["REGIONE"])
    c_prov = first_match(["PROVINCIA"])
    c_city = first_match(["CITTÀ", "CITTA"])
    c_addr = first_match(["INDIRIZZO"])
    c_attivo = first_match(["ATTIVO"])
    c_data = first_match(["DATA ATTIVAZIONE", ["DATA", "ATTIVAZIONE"]])

    c_vend_week = first_match([
        f"VENDITE {year}/{week_2}",
        f"VENDITE {year}/{week_1}",
        [f"{year}/{week_2}", "VENDITE"],
        [f"{year}/{week_1}", "VENDITE"],
    ])
    c_vend_ly = first_match([
        f"VENDITE {year-1}/{week_2}",
        f"VENDITE {year-1}/{week_1}",
        [f"{year-1}/{week_2}", "VENDITE"],
        [f"{year-1}/{week_1}", "VENDITE"],
    ])
    c_twin = first_match(["DI CUI TWIN", ["TWIN"]], after=(c_vend_ly or 0))
    c_bus_week = first_match(["DI CUI BUSINESS", ["BUSINESS"]], after=(c_twin or c_vend_ly or 0))
    c_ass_week = first_match([
        f"ASS. STRAD. EU VENDITE {year}/{week_2}",
        f"ASS. STRAD. EU VENDITE {year}/{week_1}",
        [f"{year}/{week_2}", "ASS", "STRAD"],
        [f"{year}/{week_1}", "ASS", "STRAD"],
    ])
    c_ass_ly = first_match([
        f"ASS. STRAD. EU VENDITE {year-1}/{week_2}",
        f"ASS. STRAD. EU VENDITE {year-1}/{week_1}",
        [f"{year-1}/{week_2}", "ASS", "STRAD"],
        [f"{year-1}/{week_1}", "ASS", "STRAD"],
    ])
    week_after = (c_ass_ly or c_ass_week or 0)
    c_up_eu_week = first_match(["UPSELL. EU", ["UPSELL", "EU"], ["UP", "EU"]], after=week_after)
    c_sost_week = first_match(["SOST. FAMILY", ["SOST", "FAMILY"]], after=week_after)
    c_prospect = first_match(["PROSPECT", ["PROSPECT"]], after=(c_sost_week or c_up_eu_week or week_after))

    if not c_pdv:
        raise RuntimeError(f"Colonna PDV non trovata in {path}")

    rows = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        pdv = base.norm_pdv(row[c_pdv - 1] if c_pdv and len(row) >= c_pdv else None)
        if not pdv:
            continue
        def val(c):
            return row[c - 1] if c and len(row) >= c else None
        rows.append({
            "pdv": pdv,
            "week_year": year,
            "week_num": week,
            "period": f"{year}-W{int(week):02d}",
            "area_report": val(c_area) or "",
            "regione": val(c_reg) or "",
            "provincia": val(c_prov) or "",
            "citta": val(c_city) or "",
            "indirizzo": val(c_addr) or "",
            "data_attivazione": str(val(c_data) or ""),
            "attivo": val(c_attivo) or "",
            "vendite_settimana": _preferred_total(val(c_vend_week), 0),
            "vendite_anno_prec_stessa_sett": _preferred_total(val(c_vend_ly), 0),
            "twin_settimana": base.safe_num(val(c_twin)) or 0,
            "business_vendite_settimana": base.safe_num(val(c_bus_week)) or 0,
            "prospect_settimana": base.safe_num(val(c_prospect)) or 0,
            "ass_settimana": base.safe_num(val(c_ass_week)) or 0,
            "ass_anno_prec_stessa_sett": base.safe_num(val(c_ass_ly)) or 0,
            "sost_settimana": base.safe_num(val(c_sost_week)) or 0,
            "upgrade_eu_settimana": base.safe_num(val(c_up_eu_week)) or 0,
            "source_file": os.path.basename(path),
        })
    return rows


def enrich_current(records, config):
    current_yearweek = max((r["week_year"], r["week_num"]) for r in records)
    by_pdv = {}
    for r in records:
        by_pdv.setdefault(r["pdv"], []).append(r)
    hist = {}
    current = []
    for pdv, rows in by_pdv.items():
        rows = sorted(rows, key=lambda x: (x["week_year"], x["week_num"]))
        hist[pdv] = rows
        cur = dict(rows[-1])
        ytd_rows = [x for x in rows if x["week_year"] == current_yearweek[0] and x["week_num"] <= current_yearweek[1]]
        cur["tot_vendite_anno"] = sum(x.get("vendite_settimana") or 0 for x in ytd_rows)
        cur["tot_vendite_anno_prec"] = sum(x.get("vendite_anno_prec_stessa_sett") or 0 for x in ytd_rows)
        cur["tot_twin_report"] = sum(x.get("twin_settimana") or 0 for x in ytd_rows)
        cur["tot_business_vendite_anno"] = sum(x.get("business_vendite_settimana") or 0 for x in ytd_rows)
        cur["tot_ass_anno"] = sum(x.get("ass_settimana") or 0 for x in ytd_rows)
        cur["tot_ass_anno_prec"] = sum(x.get("ass_anno_prec_stessa_sett") or 0 for x in ytd_rows)
        cur["tot_sost_family_anno"] = sum(x.get("sost_settimana") or 0 for x in ytd_rows)
        cur["tot_upgrade_eu_anno"] = sum(x.get("upgrade_eu_settimana") or 0 for x in ytd_rows)
        cur["prospect_ytd_calc"] = sum(x.get("prospect_settimana") or 0 for x in ytd_rows)
        cur["twin_ytd_calc"] = cur["tot_twin_report"]
        cur["business_ytd_calc"] = cur["tot_business_vendite_anno"]
        cur["attach_rate_calc"] = (cur["tot_ass_anno"] / cur["tot_vendite_anno"] * 100) if cur["tot_vendite_anno"] else None
        cur["up_eu_rate_calc"] = (cur["tot_upgrade_eu_anno"] / cur["tot_vendite_anno"] * 100) if cur["tot_vendite_anno"] else None
        sales_bad = cur["tot_vendite_anno"] < cur["tot_vendite_anno_prec"] * config.get("sales_bad_threshold", 0.8) if cur["tot_vendite_anno_prec"] else False
        sales_warn = cur["tot_vendite_anno"] < cur["tot_vendite_anno_prec"] if cur["tot_vendite_anno_prec"] else False
        assist_bad = cur["tot_ass_anno"] < cur["tot_ass_anno_prec"] * config.get("assist_bad_threshold", 0.8) if cur["tot_ass_anno_prec"] else False
        assist_warn = cur["tot_ass_anno"] < cur["tot_ass_anno_prec"] if cur["tot_ass_anno_prec"] else False
        cur["trend_note"] = ""
        if len(rows) >= 3:
            last3 = [x.get("vendite_settimana") or 0 for x in rows[-3:]]
            if all(v == 0 for v in last3):
                cur["trend_note"] = "Nessuna vendita nelle ultime 3 settimane"
            elif last3[-1] < last3[0]:
                cur["trend_note"] = "Vendite in calo nelle ultime settimane"
        reasons = []
        if sales_bad or sales_warn:
            reasons.append("Vendite 2026 sotto il 2025")
        if assist_bad or assist_warn:
            reasons.append("Assistenze 2026 sotto il 2025")
        if cur["trend_note"]:
            reasons.append(cur["trend_note"])
        if not reasons:
            reasons.append("Andamento regolare")
        if sales_bad or (sales_warn and assist_warn):
            stato = "Male"
        elif sales_warn or assist_warn or cur["trend_note"]:
            stato = "Da seguire"
        else:
            stato = "Bene"
        cur["stato"] = stato
        cur["motivi"] = reasons
        current.append(cur)
    current.sort(key=lambda r: ((r.get("tot_vendite_anno") or 0), (r.get("prospect_ytd_calc") or 0)), reverse=True)
    total = len(current)
    for i, r in enumerate(current, start=1):
        r["rank_all"] = i
        r["rank_text"] = f"{i} su {total}"
    return current_yearweek, current, hist


def build_html(data):
    tpl_path = Path(base.TEMPLATE_PATH)
    if not tpl_path.exists():
        raise RuntimeError("template_dashboard.html mancante nella repo.")
    tpl = tpl_path.read_text(encoding="utf-8")
    if not tpl.strip():
        raise RuntimeError("template_dashboard.html è vuoto.")
    payload = base.json.dumps(data, ensure_ascii=False)
    # Fondamentale: evita che testo proveniente da Excel/file chiuda il tag <script> della pagina.
    payload = payload.replace("</", "<\\/").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    return tpl.replace("__DATA_JSON__", payload).replace("__CURRENT_WEEK__", f"{data['meta']['current_week']:02d}")


def _find_latest_custom_report(root_dir):
    folder = Path(root_dir) / "input" / "custom_report"
    if not folder.exists():
        return None
    files = [p for p in folder.rglob("*.xlsx") if not p.name.startswith("~$")]
    if not files:
        return None
    return sorted(files, key=lambda p: p.stat().st_mtime)[-1]


def _fmt_custom_col(v):
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m")
    s = str(v or "").strip()
    m = re.search(r"(\d{1,2})[/-](\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}"
    return s


def load_custom_report(root_dir, lista_map, anag_map, out_dir):
    path = _find_latest_custom_report(root_dir)
    if not path:
        return None
    wb = base.openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    if len(header) < 4:
        return None
    dynamic_cols = header[3:-1]
    date_labels = [_fmt_custom_col(v) for v in dynamic_cols]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        pdv = base.norm_pdv(row[0] if row else None)
        if not pdv:
            continue
        daily = [int(round(base.safe_num(v) or 0)) for v in row[3:-1]]
        total = int(round(base.safe_num(row[-1]) or sum(daily)))
        li = lista_map.get(pdv, {})
        an = anag_map.get(pdv, {})
        rows.append({
            "pdv": pdv,
            "city": (row[2] or li.get("lista_citta") or ""),
            "rzv": (row[1] or an.get("rzv") or ""),
            "agent": li.get("agente", "") or "",
            "cr": an.get("cr", "") or "",
            "values": daily,
            "total": total,
        })
    rows.sort(key=lambda r: (r["total"], r["values"][-1] if r["values"] else 0), reverse=True)
    for i, r in enumerate(rows, start=1):
        r["rank"] = i
    dest_dir = Path(out_dir) / "files" / "CUSTOM_REPORT"
    dest_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, dest_dir / path.name)
    return {
        "title": f"Custom Report · {date_labels[0]} - {date_labels[-1]}" if date_labels else "Custom Report",
        "source_name": path.name,
        "source_path": f"files/CUSTOM_REPORT/{path.name}",
        "updated_at": datetime.datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
        "columns": date_labels,
        "rows": rows,
        "summary": {"pdv_count": len(rows), "grand_total": sum(r["total"] for r in rows)},
    }


def main():
    config = base.load_config()
    if len(base.sys.argv) >= 5:
        lista, anag, report_dir, out_dir = base.sys.argv[1:5]
    else:
        lista, anag, report_dir, out_dir = base.pick_inputs()
    base.os.makedirs(out_dir, exist_ok=True)
    lista_map = base.load_lista(lista)
    anag_map = base.load_anag(anag)
    scan = base.scan_report_files(report_dir, year_mode=config.get("year_mode", "latest_year_only"))
    if not scan["selected_paths"]:
        raise RuntimeError("Nessun report ENI valido trovato nella cartella selezionata.")
    records = []
    for path in scan["selected_paths"]:
        for r in parse_report_dynamic(path):
            li = lista_map.get(r["pdv"], {})
            an = anag_map.get(r["pdv"], {})
            r["agente"] = li.get("agente", "") or ""
            r["rzv"] = an.get("rzv", "") or ""
            r["cr"] = an.get("cr", "") or ""
            if not r["citta"]:
                r["citta"] = li.get("lista_citta", "")
            if not r["indirizzo"]:
                r["indirizzo"] = li.get("lista_indirizzo", "")
            records.append(r)
    ded = {}
    for r in records:
        ded[(r["pdv"], r["week_year"], r["week_num"])] = r
    records = sorted(ded.values(), key=lambda x: (x["week_year"], x["week_num"], x["pdv"]))
    current_yearweek, current, hist = enrich_current(records, config)
    summary = base.build_summary(current)
    base.os.makedirs(base.os.path.join(out_dir, "files"), exist_ok=True)
    export_manifest = base.build_export_reports(out_dir, current, current_yearweek[1])
    file_utili = base.copy_file_utili(out_dir)
    data = base.build_data_for_html(current, hist, summary, export_manifest, file_utili, current_yearweek[1], current_yearweek[0])
    data["gare_pdv"] = config.get("gare_pdv", [])
    data["gare_agenti"] = config.get("gare_agenti", [])
    data["custom_report"] = load_custom_report(base.BASE_DIR, lista_map, anag_map, out_dir)
    html_path = base.os.path.join(out_dir, "Telepass_ENI_sito_v6.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(build_html(data))
    master_xlsx = base.os.path.join(out_dir, "Dati_Telepass_ENI_v6.xlsx")
    base.build_master_workbook(master_xlsx, current, records, {"selected_year": scan["selected_year"], "current_week": current_yearweek[1]})
    log_path = base.os.path.join(out_dir, "log_file_usati_v6.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("FILE REPORT TROVATI E USATI\n")
        f.write("==========================\n")
        for p in scan["selected_paths"]:
            y, w, _ = base.extract_week_year(p)
            f.write(f"{y}/W{w:02d} -> {p}\n")
        cp = _find_latest_custom_report(base.BASE_DIR)
        f.write("\nCUSTOM REPORT\n")
        f.write("=============\n")
        f.write((str(cp) if cp else "Nessun custom report trovato") + "\n")
        f.write("\nSETTIMANE MANCANTI\n")
        f.write("==================\n")
        if scan["missing_weeks"]:
            for y, w in scan["missing_weeks"]:
                f.write(f"{y}/W{w:02d}\n")
        else:
            f.write("Nessuna settimana mancante nel blocco usato.\n")
        f.write("\nFILE SCARTATI\n")
        f.write("============\n")
        for p, reason in scan["skipped"]:
            f.write(f"{p} -> {reason}\n")
    print("Creato:", html_path)
    print("Creato:", master_xlsx)
    print("Creato:", log_path)


if __name__ == "__main__":
    main()
