from pathlib import Path
import shutil

import aggiorna_dashboard as base
import aggiorna_dashboard_github as gh

ROOT = Path(__file__).resolve().parent
LISTA_DIR = ROOT / 'input' / 'lista'
ANAG_DIR = ROOT / 'input' / 'anagrafica'
REPORT_DIR = ROOT / 'input' / 'report_settimanali'
OUT_DIR = ROOT / 'docs'


def first_xlsx(folder: Path):
    files = [p for p in folder.rglob('*.xlsx') if not p.name.startswith('~$')]
    if not files:
        return None
    return sorted(files, key=lambda p: p.stat().st_mtime)[-1]


def make_zero_record_from_lista(pdv, lista_info, anag_info, year, week):
    week_2 = f'{int(week):02d}'
    return {
        'pdv': pdv,
        'week_year': year,
        'week_num': week,
        'period': f'{year}-W{week_2}',
        'area_report': '',
        'regione': lista_info.get('lista_regione', '') or '',
        'provincia': lista_info.get('lista_provincia', '') or '',
        'citta': lista_info.get('lista_citta', '') or '',
        'indirizzo': lista_info.get('lista_indirizzo', '') or '',
        'data_attivazione': '',
        'attivo': '',
        'vendite_settimana': 0,
        'vendite_anno_prec_stessa_sett': 0,
        'twin_settimana': 0,
        'business_vendite_settimana': 0,
        'prospect_settimana': 0,
        'ass_settimana': 0,
        'ass_anno_prec_stessa_sett': 0,
        'sost_settimana': 0,
        'upgrade_eu_settimana': 0,
        'sost_family_settimana': 0,
        'tot_vendite_anno': 0,
        'tot_vendite_anno_prec': 0,
        'tot_twin_report': 0,
        'tot_business_vendite_anno': 0,
        'tot_ass_anno': 0,
        'tot_ass_anno_prec': 0,
        'tot_sost_anno': 0,
        'tot_upgrade_eu_anno': 0,
        'tot_sost_family_anno': 0,
        'source_file': 'LISTA PDV ENI - non presente nel report settimana corrente',
        'agente': lista_info.get('agente', '') or '',
        'rzv': anag_info.get('rzv', '') or '',
        'cr': anag_info.get('cr', '') or '',
    }


def add_missing_lista_pdv_to_current_week(records, lista_map, anag_map):
    """Fa vedere anche i PDV presenti in LISTA ma assenti dai report settimana."""
    if not records:
        return records, []

    current_yearweek = max((r['week_year'], r['week_num']) for r in records)
    pdv_in_current_week = {
        r['pdv']
        for r in records
        if (r.get('week_year'), r.get('week_num')) == current_yearweek
    }

    added = []
    for pdv, lista_info in sorted(lista_map.items()):
        if pdv in pdv_in_current_week:
            continue
        anag_info = anag_map.get(pdv, {})
        records.append(make_zero_record_from_lista(pdv, lista_info, anag_info, current_yearweek[0], current_yearweek[1]))
        added.append(pdv)

    records = sorted(records, key=lambda x: (x['week_year'], x['week_num'], x['pdv']))
    return records, added


def mark_lista_only_rows(current, added_pdv):
    added = set(added_pdv)
    if not added:
        return
    note = 'PDV presente in LISTA PDV ENI ma non ancora presente nel report settimanale corrente'
    for row in current:
        if row.get('pdv') in added:
            row['stato'] = 'Da seguire'
            row['motivi'] = [note]


def create_mobile_workbook_with_address(rows, out_path, title, current_week, filter_text):
    wb = base.Workbook()
    ws = wb.active
    ws.title = 'HOME'
    ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:F1')
    ws['A1'] = title
    ws['A1'].fill = base.PatternFill('solid', fgColor=base.BLUE)
    ws['A1'].font = base.Font(size=18, bold=True, color='FFFFFF')
    ws['A1'].alignment = base.Alignment(horizontal='center')
    ws.row_dimensions[1].height = 28
    ws['A3'] = 'Filtro'
    ws['B3'] = filter_text
    ws['A4'] = 'Settimana attuale'
    ws['B4'] = f'W{current_week:02d}'

    cards = [
        ('PDV', len(rows), base.LIGHT),
        ('Vendite 2026', sum(r['tot_vendite_anno'] or 0 for r in rows), 'EAF2FF'),
        ('Prospect 2026', sum(r['prospect_ytd_calc'] or 0 for r in rows), 'E8F7EC'),
        ('Assistenze 2026', sum(r['tot_ass_anno'] or 0 for r in rows), 'FFF4E5'),
    ]
    for idx, (label, val, fill) in enumerate(cards, 1):
        ws.cell(6, idx, label)
        ws.cell(7, idx, val)
        ws.cell(6, idx).font = base.Font(bold=True, color=base.BLUE)
        ws.cell(7, idx).font = base.Font(bold=True, size=20, color=base.BLUE)
        ws.cell(6, idx).fill = base.PatternFill('solid', fgColor=fill.replace('#', ''))
        ws.cell(7, idx).fill = base.PatternFill('solid', fgColor=fill.replace('#', ''))
        ws.cell(6, idx).alignment = base.Alignment(horizontal='center')
        ws.cell(7, idx).alignment = base.Alignment(horizontal='center')

    headers = ['PDV', 'Città', 'Indirizzo', 'Agente', 'CR', 'Vend 2026', 'Vend 2025', 'Ass 2026', 'Ass 2025', 'Prospect', 'Twin', 'Business', 'Sost anno', 'UP EU', 'Stato']
    data = []
    for r in rows:
        data.append([
            r.get('pdv'), r.get('citta'), r.get('indirizzo'), r.get('agente'), r.get('cr'),
            r.get('tot_vendite_anno'), r.get('tot_vendite_anno_prec'), r.get('tot_ass_anno'), r.get('tot_ass_anno_prec'),
            r.get('prospect_ytd_calc'), r.get('twin_ytd_calc'), r.get('business_ytd_calc'), r.get('tot_sost_family_anno'), r.get('tot_upgrade_eu_anno'), r.get('stato')
        ])
    base.add_table(ws, 10, 1, headers, data, 'Report')
    for row in range(11, ws.max_row + 1):
        for col in range(6, 15):
            ws.cell(row, col).number_format = '#,##0'
    ws.freeze_panes = 'A10'
    base.style_sheet(ws)
    base.autosize(ws, max_width=28)
    wb.save(out_path)


def add_address_to_custom_report(custom_report, lista_map):
    if not custom_report:
        return custom_report
    for row in custom_report.get('rows', []):
        row['address'] = lista_map.get(row.get('pdv'), {}).get('lista_indirizzo', '') or ''
    return custom_report


def patch_html_address_display(html):
    patches = [
        (
            '.small-muted{color:var(--muted);font-size:12px;line-height:1.3}',
            '.small-muted{color:var(--muted);font-size:12px;line-height:1.3}\n.city-cell{min-width:120px}.city-main{font-weight:800}.city-address{color:var(--muted);font-size:10px;line-height:1.2;margin-top:2px;max-width:170px}'
        ),
        ('Cerca PDV, città, agente, CR...', 'Cerca PDV, città, via, agente, CR...'),
        (
            '<td>${esc(r.city)}</td>',
            '<td><div class="city-cell"><div class="city-main">${esc(r.city)}</div>${r.address ? \'<div class="city-address">\' + esc(r.address) + \'</div>\' : \'\'}</div></td>'
        ),
        (
            'r.rank_sales, r.pdv, r.city, r.agent, r.cr, r.rzv,',
            "r.rank_sales, r.pdv, r.city, r.address||'', r.agent, r.cr, r.rzv,"
        ),
        (
            "return filteredCustomRows().map(r=>[r.rank,r.pdv,r.city,r.agent||'',r.cr||'',r.rzv||'',...(r.values||[]),r.total]);",
            "return filteredCustomRows().map(r=>[r.rank,r.pdv,r.city,r.address||'',r.agent||'',r.cr||'',r.rzv||'',...(r.values||[]),r.total]);"
        ),
        (
            "return [[\n    r.pdv, r.city, r.agent, r.cr, r.rzv,",
            "return [[\n    r.pdv, r.city, r.address||'', r.agent, r.cr, r.rzv,"
        ),
        ("['#','PDV','Città','Agente','CR','RZV',", "['#','PDV','Città','Indirizzo','Agente','CR','RZV',"),
        ("['PDV','Città','Agente','CR','RZV',", "['PDV','Città','Indirizzo','Agente','CR','RZV',"),
        ("const headers=['#','PDV','Città','Agente','CR','RZV',", "const headers=['#','PDV','Città','Indirizzo','Agente','CR','RZV',"),
    ]
    for old, new in patches:
        html = html.replace(old, new)
    return html


def run_dashboard(lista, anag, report_dir, out_dir):
    config = base.load_config()
    base.os.makedirs(out_dir, exist_ok=True)

    lista_map = base.load_lista(str(lista))
    anag_map = base.load_anag(str(anag))

    scan = base.scan_report_files(str(report_dir), year_mode=config.get('year_mode', 'latest_year_only'))
    if not scan['selected_paths']:
        raise SystemExit('ERRORE: nessun report ENI valido trovato dentro input/report_settimanali')

    records = []
    for path in scan['selected_paths']:
        for r in gh.parse_report_dynamic(path):
            li = lista_map.get(r['pdv'], {})
            an = anag_map.get(r['pdv'], {})
            r['agente'] = li.get('agente', '') or ''
            r['rzv'] = an.get('rzv', '') or ''
            r['cr'] = an.get('cr', '') or ''
            if not r['citta']:
                r['citta'] = li.get('lista_citta', '')
            if not r['indirizzo']:
                r['indirizzo'] = li.get('lista_indirizzo', '')
            records.append(r)

    ded = {}
    for r in records:
        ded[(r['pdv'], r['week_year'], r['week_num'])] = r
    records = sorted(ded.values(), key=lambda x: (x['week_year'], x['week_num'], x['pdv']))

    records, lista_only_pdv = add_missing_lista_pdv_to_current_week(records, lista_map, anag_map)

    current_yearweek, current, hist = gh.enrich_current(records, config)
    mark_lista_only_rows(current, lista_only_pdv)

    summary = base.build_summary(current)
    base.os.makedirs(base.os.path.join(out_dir, 'files'), exist_ok=True)

    base.create_mobile_workbook = create_mobile_workbook_with_address
    export_manifest = base.build_export_reports(str(out_dir), current, current_yearweek[1])
    file_utili = base.copy_file_utili(str(out_dir))

    data = base.build_data_for_html(current, hist, summary, export_manifest, file_utili, current_yearweek[1], current_yearweek[0])
    data['gare_pdv'] = config.get('gare_pdv', [])
    data['gare_agenti'] = config.get('gare_agenti', [])
    data['custom_report'] = add_address_to_custom_report(gh.load_custom_report(base.BASE_DIR, lista_map, anag_map, str(out_dir)), lista_map)

    html_path = base.os.path.join(out_dir, 'Telepass_ENI_sito_v6.html')
    html = patch_html_address_display(gh.build_html(data))
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)

    master_xlsx = base.os.path.join(out_dir, 'Dati_Telepass_ENI_v6.xlsx')
    base.build_master_workbook(master_xlsx, current, records, {'selected_year': scan['selected_year'], 'current_week': current_yearweek[1]})

    log_path = base.os.path.join(out_dir, 'log_file_usati_v6.txt')
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write('FILE REPORT TROVATI E USATI\n')
        f.write('==========================\n')
        for p in scan['selected_paths']:
            y, w, _ = base.extract_week_year(p)
            f.write(f'{y}/W{w:02d} -> {p}\n')

        cp = gh._find_latest_custom_report(base.BASE_DIR)
        f.write('\nCUSTOM REPORT\n')
        f.write('=============\n')
        f.write((str(cp) if cp else 'Nessun custom report trovato') + '\n')

        f.write('\nPDV AGGIUNTI DA LISTA PDV ENI\n')
        f.write('=============================\n')
        if lista_only_pdv:
            for pdv in lista_only_pdv:
                li = lista_map.get(pdv, {})
                an = anag_map.get(pdv, {})
                f.write(f'{pdv} -> {li.get("lista_citta", "")} | {li.get("lista_indirizzo", "")} | Agente: {li.get("agente", "")} | CR: {an.get("cr", "")} | RZV: {an.get("rzv", "")}\n')
        else:
            f.write('Nessuno. Tutti i PDV della lista sono presenti nel report corrente.\n')

        f.write('\nSETTIMANE MANCANTI\n')
        f.write('==================\n')
        if scan['missing_weeks']:
            for y, w in scan['missing_weeks']:
                f.write(f'{y}/W{w:02d}\n')
        else:
            f.write('Nessuna settimana mancante nel blocco usato.\n')

        f.write('\nFILE SCARTATI\n')
        f.write('============\n')
        for p, reason in scan['skipped']:
            f.write(f'{p} -> {reason}\n')

    print('Creato:', html_path)
    print('Creato:', master_xlsx)
    print('Creato:', log_path)
    print('PDV aggiunti da LISTA PDV ENI:', len(lista_only_pdv))


def main():
    lista = first_xlsx(LISTA_DIR)
    anag = first_xlsx(ANAG_DIR)
    if not lista:
        raise SystemExit('ERRORE: manca il file Lista PDV nella cartella input/lista')
    if not anag:
        raise SystemExit('ERRORE: manca il file Anagrafica nella cartella input/anagrafica')
    if not REPORT_DIR.exists() or not any(REPORT_DIR.rglob('*.xlsx')):
        raise SystemExit('ERRORE: non ci sono report settimana dentro input/report_settimanali')

    OUT_DIR.mkdir(exist_ok=True, parents=True)
    run_dashboard(lista, anag, REPORT_DIR, OUT_DIR)

    generated = OUT_DIR / 'Telepass_ENI_sito_v6.html'
    if generated.exists():
        shutil.copy2(generated, OUT_DIR / 'index.html')

    (OUT_DIR / '.nojekyll').write_text('', encoding='utf-8')
    print('Build completata. Apri docs/index.html o pubblica docs con GitHub Pages.')


if __name__ == '__main__':
    main()
