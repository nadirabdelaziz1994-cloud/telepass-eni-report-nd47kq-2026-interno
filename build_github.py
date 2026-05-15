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
    """Fa vedere anche i PDV presenti in LISTA ma assenti dai report settimana.

    Prima la dashboard partiva solo dai report settimanali: un nuovo PDV appena
    inserito in LISTA, ma non ancora produttivo/non ancora presente nel file
    settimana, non veniva mai aggiunto a classifica, filtri ed export.
    Qui creiamo una riga a zero per la settimana corrente, così rimane visibile.
    """
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
    export_manifest = base.build_export_reports(str(out_dir), current, current_yearweek[1])
    file_utili = base.copy_file_utili(str(out_dir))

    data = base.build_data_for_html(current, hist, summary, export_manifest, file_utili, current_yearweek[1], current_yearweek[0])
    data['gare_pdv'] = config.get('gare_pdv', [])
    data['gare_agenti'] = config.get('gare_agenti', [])
    data['custom_report'] = gh.load_custom_report(base.BASE_DIR, lista_map, anag_map, str(out_dir))

    html_path = base.os.path.join(out_dir, 'Telepass_ENI_sito_v6.html')
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(gh.build_html(data))

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
                f.write(f'{pdv} -> {li.get("lista_citta", "")} | Agente: {li.get("agente", "")} | CR: {an.get("cr", "")} | RZV: {an.get("rzv", "")}\n')
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
