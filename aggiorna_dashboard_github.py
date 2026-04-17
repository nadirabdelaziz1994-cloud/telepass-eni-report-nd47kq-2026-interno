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
    c_sost_family_week = c_sost_week

    c_tot_sales = first_match([
        f"TOTALE VENDITE TELEPASS {year}",
        [str(year), "TOTALE", "VENDITE", "TELEPASS"],
        [str(year), "TOTALE", "VENDITE"],
    ])
    c_tot_sales_prev = first_match([
        f"TOTALE VENDITE TELEPASS {year-1}",
        [str(year-1), "TOTALE", "VENDITE", "TELEPASS"],
        [str(year-1), "TOTALE", "VENDITE"],
    ])
    c_tot_twin = first_match(["TOTALE TWIN", ["TOTALE", "TWIN"]], after=(c_tot_sales_prev or c_tot_sales or 0))
    c_tot_bus = first_match(["DI CUI BUSINESS", ["BUSINESS"]], after=(c_tot_twin or c_tot_sales_prev or c_tot_sales or 0))
    c_tot_ass = first_match([f"TOTALE ASS. STRAD. {year}", [str(year), "TOTALE", "ASS", "STRAD"]])
    c_tot_ass_prev = first_match([f"TOTALE ASS. STRAD. {year-1}", [str(year-1), "TOTALE", "ASS", "STRAD"]])
    totals_after = (c_tot_ass_prev or c_tot_ass or 0)
    c_tot_sost = first_match(["SOST.", ["SOST"]], after=totals_after)
    c_tot_up_eu = first_match(["UPSELL. EU", ["UPSELL", "EU"], ["UP", "EU"]], after=totals_after)
    c_tot_sost_family = first_match(["SOST. FAMILY", ["SOST", "FAMILY"]], after=totals_after)

    recs = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        if not row:
            continue
        pdv = base.norm_pdv(row[c_pdv - 1] if c_pdv else None)
        if not pdv:
            continue
        vend_week = base.safe_num(row[c_vend_week - 1]) if c_vend_week else None
        bus_week = base.safe_num(row[c_bus_week - 1]) if c_bus_week else 0
        if bus_week is not None and (bus_week < 0 or (abs(bus_week - round(bus_week)) > 1e-6 and abs(bus_week) < 1)):
            bus_week = 0
        twin_week = base.safe_num(row[c_twin - 1]) if c_twin else None
        if twin_week is not None and twin_week < 0:
            twin_week = 0
        recs.append({
            "pdv": pdv,
            "week_year": year,
            "week_num": week,
            "period": f"{year}-W{week_2}",
            "area_report": row[c_area - 1] if c_area else "",
            "regione": row[c_reg - 1] if c_reg else "",
            "provincia": row[c_prov - 1] if c_prov else "",
            "citta": row[c_city - 1] if c_city else "",
            "indirizzo": row[c_addr - 1] if c_addr else "",
            "data_attivazione": row[c_data - 1].strftime("%Y-%m-%d") if c_data and hasattr(row[c_data - 1], "strftime") else (str(row[c_data - 1]) if c_data and row[c_data - 1] else ""),
            "attivo": row[c_attivo - 1] if c_attivo else "",
            "vendite_settimana": vend_week or 0,
            "vendite_anno_prec_stessa_sett": base.safe_num(row[c_vend_ly - 1]) if c_vend_ly else 0,
            "twin_settimana": twin_week or 0,
            "business_vendite_settimana": bus_week or 0,
            "prospect_settimana": max((vend_week or 0) - (bus_week or 0), 0),
            "ass_settimana": base.safe_num(row[c_ass_week - 1]) if c_ass_week else 0,
            "ass_anno_prec_stessa_sett": base.safe_num(row[c_ass_ly - 1]) if c_ass_ly else 0,
            "sost_settimana": base.safe_num(row[c_sost_week - 1]) if c_sost_week else 0,
            "upgrade_eu_settimana": base.safe_num(row[c_up_eu_week - 1]) if c_up_eu_week else 0,
            "sost_family_settimana": base.safe_num(row[c_sost_family_week - 1]) if c_sost_family_week else 0,
            "tot_vendite_anno": base.safe_num(row[c_tot_sales - 1]) if c_tot_sales else 0,
            "tot_vendite_anno_prec": base.safe_num(row[c_tot_sales_prev - 1]) if c_tot_sales_prev else 0,
            "tot_twin_report": base.safe_num(row[c_tot_twin - 1]) if c_tot_twin else 0,
            "tot_business_vendite_anno": base.safe_num(row[c_tot_bus - 1]) if c_tot_bus else 0,
            "tot_ass_anno": base.safe_num(row[c_tot_ass - 1]) if c_tot_ass else 0,
            "tot_ass_anno_prec": base.safe_num(row[c_tot_ass_prev - 1]) if c_tot_ass_prev else 0,
            "tot_sost_anno": base.safe_num(row[c_tot_sost - 1]) if c_tot_sost else 0,
            "tot_upgrade_eu_anno": base.safe_num(row[c_tot_up_eu - 1]) if c_tot_up_eu else 0,
            "tot_sost_family_anno": base.safe_num(row[c_tot_sost_family - 1]) if c_tot_sost_family else 0,
            "source_file": base.os.path.basename(path),
        })
    return recs


def enrich_current(records, config):
    hist = base.defaultdict(list)
    for r in records:
        hist[r["pdv"]].append(r)
    for pdv in hist:
        hist[pdv].sort(key=lambda x: (x["week_year"], x["week_num"]))
    current_yearweek = max((r["week_year"], r["week_num"]) for r in records)
    current = []
    th = config["thresholds"]
    for pdv, arr in hist.items():
        cur = next((x for x in arr if (x["week_year"], x["week_num"]) == current_yearweek), None)
        if not cur:
            continue
        cur = cur.copy()
        prev = arr[-2] if len(arr) >= 2 else None
        sales_ytd_sum = sum((x.get("vendite_settimana") or 0) for x in arr)
        sales_prev_ytd_sum = sum((x.get("vendite_anno_prec_stessa_sett") or 0) for x in arr)
        assist_ytd_sum = sum((x.get("ass_settimana") or 0) for x in arr)
        assist_prev_ytd_sum = sum((x.get("ass_anno_prec_stessa_sett") or 0) for x in arr)
        twin_ytd_sum = sum((x.get("twin_settimana") or 0) for x in arr)
        business_ytd_sum = sum((x.get("business_vendite_settimana") or 0) for x in arr)
        prospect_ytd_sum = sum((x.get("prospect_settimana") or 0) for x in arr)
        sost_ytd_sum = sum((x.get("sost_settimana") or 0) for x in arr)
        sost_family_ytd_sum = sum((x.get("sost_family_settimana") or 0) for x in arr)
        up_eu_ytd_sum = sum((x.get("upgrade_eu_settimana") or 0) for x in arr)
        cur["agente"] = cur.get("agente", "")
        cur["cr"] = cur.get("cr", "")
        cur["rzv"] = cur.get("rzv", "")
        cur["business_ytd_calc"] = business_ytd_sum
        cur["twin_ytd_calc"] = twin_ytd_sum
        cur["prospect_ytd_calc"] = prospect_ytd_sum
        cur["tot_vendite_anno"] = _preferred_total(cur.get("tot_vendite_anno"), sales_ytd_sum)
        cur["tot_vendite_anno_prec"] = _preferred_total(cur.get("tot_vendite_anno_prec"), sales_prev_ytd_sum)
        cur["tot_ass_anno"] = _preferred_total(cur.get("tot_ass_anno"), assist_ytd_sum)
        cur["tot_ass_anno_prec"] = _preferred_total(cur.get("tot_ass_anno_prec"), assist_prev_ytd_sum)
        cur["tot_sost_anno"] = _preferred_total(cur.get("tot_sost_anno"), sost_ytd_sum)
        cur["tot_sost_family_anno"] = _preferred_total(cur.get("tot_sost_family_anno"), sost_family_ytd_sum)
        cur["tot_upgrade_eu_anno"] = _preferred_total(cur.get("tot_upgrade_eu_anno"), up_eu_ytd_sum)
        cur["prev_week"] = prev["week_num"] if prev else None
        cur["vendite_week_diff"] = (cur.get("vendite_settimana") or 0) - ((prev or {}).get("vendite_settimana") or 0) if prev else None
        cur["prospect_week_diff"] = (cur.get("prospect_settimana") or 0) - ((prev or {}).get("prospect_settimana") or 0) if prev else None
        cur["ass_week_diff"] = (cur.get("ass_settimana") or 0) - ((prev or {}).get("ass_settimana") or 0) if prev else None
        cur["sales_ytd_diff"] = (cur.get("tot_vendite_anno") or 0) - (cur.get("tot_vendite_anno_prec") or 0)
        cur["assist_ytd_diff"] = (cur.get("tot_ass_anno") or 0) - (cur.get("tot_ass_anno_prec") or 0)
        cur["sales_ytd_pct"] = base.pct(cur.get("tot_vendite_anno"), cur.get("tot_vendite_anno_prec"))
        cur["assist_ytd_pct"] = base.pct(cur.get("tot_ass_anno"), cur.get("tot_ass_anno_prec"))
        ops = (cur.get("tot_vendite_anno") or 0) + (cur.get("tot_sost_family_anno") or 0)
        cur["attach_rate"] = (cur.get("tot_ass_anno") or 0) / ops if ops else None
        cur["up_eu_rate"] = (cur.get("tot_upgrade_eu_anno") or 0) / (cur.get("tot_sost_family_anno") or 0) if (cur.get("tot_sost_family_anno") or 0) else None
        remaining = max(52 - cur["week_num"], 1)
        gap = max((cur.get("tot_vendite_anno_prec") or 0) - (cur.get("tot_vendite_anno") or 0), 0)
        cur["sales_recovery_weekly_need"] = gap / remaining if gap > 0 else 0
        cur["current_weekly_avg"] = (cur.get("tot_vendite_anno") or 0) / max(cur["week_num"], 1)
        trend = ""
        if len(arr) >= 2 and (arr[-1].get("vendite_settimana") or 0) < (arr[-2].get("vendite_settimana") or 0):
            if len(arr) >= 3 and (arr[-2].get("vendite_settimana") or 0) < (arr[-3].get("vendite_settimana") or 0):
                trend = f"In calo da 2 settimane (W{arr[-3]['week_num']:02d} → W{arr[-2]['week_num']:02d} → W{arr[-1]['week_num']:02d})"
            else:
                trend = f"Ultima settimana in calo vs W{arr[-2]['week_num']:02d}"
        cur["trend_note"] = trend
        sp = cur["sales_ytd_pct"] if cur["sales_ytd_pct"] is not None else 0
        ap = cur["assist_ytd_pct"] if cur["assist_ytd_pct"] is not None else 0
        sales_bad = sp <= th["sales_bad_pct"] and cur["sales_ytd_diff"] <= -th["sales_bad_abs"]
        sales_warn = sp <= th["sales_warn_pct"] and cur["sales_ytd_diff"] <= -th["sales_warn_abs"]
        assist_bad = ap <= th["assist_bad_pct"] and cur["assist_ytd_diff"] <= -th["assist_bad_abs"]
        assist_warn = ap <= th["assist_warn_pct"] and cur["assist_ytd_diff"] <= -th["assist_warn_abs"]
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
    return tpl.replace("__DATA_JSON__", base.json.dumps(data, ensure_ascii=False)).replace("__CURRENT_WEEK__", f"{data['meta']['current_week']:02d}")


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


base.parse_report_dynamic = parse_report_dynamic
base.enrich_current = enrich_current
base.build_html = build_html


if __name__ == "__main__":
    main()
