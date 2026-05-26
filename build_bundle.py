from pathlib import Path
import datetime
import re
import shutil
import unicodedata

import build_github as site
import aggiorna_dashboard as base

ORIGINAL_PATCH_HTML = site.patch_html_address_display


def norm_header(value):
    if value is None:
        return ''
    text = str(value).replace('\r', ' ').replace('\n', ' ').replace('\xa0', ' ')
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r'[^A-Z0-9%]+', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()


def safe_int(value):
    n = base.safe_num(value)
    return int(round(n or 0))


def find_col(headers, *needles):
    for idx, header in enumerate(headers, start=1):
        for needle in needles:
            if needle and needle in header:
                return idx
    return None


def find_bundle_header(ws):
    for row_idx in range(1, min(ws.max_row, 20) + 1):
        headers = [norm_header(ws.cell(row_idx, col).value) for col in range(1, ws.max_column + 1)]
        has_pdv = any(h in {'PDV', 'PV ENI'} or 'PUNTOEROGAZIONE' in h or 'PUNTO EROGAZIONE' in h for h in headers)
        has_bundle = any(h == 'BUNDLE' or h.startswith('BUNDLE ') or '% BUNDLE' in h for h in headers)
        if not has_pdv or not has_bundle:
            continue
        cols = {
            'product': find_col(headers, 'PRODOTTO'),
            'pdv': find_col(headers, 'PUNTOEROGAZIONE', 'PUNTO EROGAZIONE', 'PV ENI', 'PDV'),
            'area': find_col(headers, 'AREA COMM'),
            'region': find_col(headers, 'REGIONE'),
            'province': find_col(headers, 'PROVINCIA'),
            'city': find_col(headers, 'CITTA'),
            'address': find_col(headers, 'INDIRIZZO'),
            'sales_total': find_col(headers, 'VENDITE TOT', 'VENDITE TOTALI'),
            'twin': find_col(headers, 'DI CUI TWIN', 'TWIN'),
            'prospect': find_col(headers, 'PROSPECT'),
            'bundle': find_col(headers, 'BUNDLE'),
            'bundle_rate': find_col(headers, '% BUNDLE', 'BUNDLE SU PROSPECT'),
        }
        if cols['pdv'] and cols['bundle']:
            return row_idx, cols
    return None, {}


def cell(row, col_idx):
    if not col_idx or col_idx - 1 >= len(row):
        return None
    return row[col_idx - 1]


def bundle_period(ws, header_row, cols):
    if header_row and header_row > 1:
        for col in [cols.get('sales_total'), cols.get('bundle'), cols.get('bundle_rate')]:
            if col:
                value = str(ws.cell(header_row - 1, col).value or '').strip()
                if value:
                    return value
        for col in range(1, ws.max_column + 1):
            value = str(ws.cell(header_row - 1, col).value or '').strip()
            if value:
                return value
    return 'periodo caricato'


def load_bundle_report(root_dir, lista_map, anag_map, out_dir):
    path = site.gh._find_latest_custom_report(root_dir)
    if not path:
        return None
    wb = base.openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row, cols = find_bundle_header(ws)
    if not header_row:
        return None

    period = bundle_period(ws, header_row, cols)
    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        pdv = base.norm_pdv(cell(row, cols.get('pdv')))
        if not pdv:
            continue
        prospect = safe_int(cell(row, cols.get('prospect')))
        bundle = safe_int(cell(row, cols.get('bundle')))
        raw_rate = base.safe_num(cell(row, cols.get('bundle_rate'))) if cols.get('bundle_rate') else None
        if raw_rate is None:
            rate = bundle / prospect * 100 if prospect else None
        else:
            rate = raw_rate * 100 if abs(raw_rate) <= 1.5 else raw_rate
        rate = round(rate, 1) if rate is not None else None

        li = lista_map.get(pdv, {})
        an = anag_map.get(pdv, {})
        sales_total = safe_int(cell(row, cols.get('sales_total')))
        twin = safe_int(cell(row, cols.get('twin')))
        city = cell(row, cols.get('city')) or li.get('lista_citta') or ''
        address = cell(row, cols.get('address')) or li.get('lista_indirizzo') or ''
        rows.append({
            'pdv': pdv,
            'product': cell(row, cols.get('product')) or '',
            'area': cell(row, cols.get('area')) or '',
            'region': cell(row, cols.get('region')) or '',
            'province': cell(row, cols.get('province')) or '',
            'city': city,
            'address': address,
            'agent': li.get('agente', '') or '',
            'cr': an.get('cr', '') or '',
            'rzv': an.get('rzv', '') or '',
            'sales_total': sales_total,
            'twin': twin,
            'prospect': prospect,
            'bundle': bundle,
            'bundle_rate': rate,
            'values': [sales_total, twin, prospect, bundle, rate],
            'total': bundle,
        })

    rows.sort(key=lambda r: ((r['bundle_rate'] if r['bundle_rate'] is not None else -1), r['bundle']), reverse=True)
    for idx, row in enumerate(rows, start=1):
        row['rank'] = idx

    dest_dir = Path(out_dir) / 'files' / 'CUSTOM_REPORT'
    dest_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, dest_dir / path.name)
    total_prospect = sum(r['prospect'] for r in rows)
    total_bundle = sum(r['bundle'] for r in rows)
    return {
        'kind': 'bundle',
        'title': f'Bundle · {period}',
        'period': period,
        'source_name': path.name,
        'source_path': f'files/CUSTOM_REPORT/{path.name}',
        'updated_at': datetime.datetime.fromtimestamp(path.stat().st_mtime).strftime('%d/%m/%Y %H:%M'),
        'columns': ['Vendite tot', 'Twin', 'Prospect', 'Bundle', '% Bundle'],
        'rows': rows,
        'summary': {
            'pdv_count': len(rows),
            'sales_total': sum(r['sales_total'] for r in rows),
            'prospect': total_prospect,
            'bundle': total_bundle,
            'bundle_rate': round(total_bundle / total_prospect * 100, 1) if total_prospect else None,
        },
    }


def keep_address(custom_report, lista_map):
    if not custom_report:
        return custom_report
    for row in custom_report.get('rows', []):
        row['address'] = lista_map.get(row.get('pdv'), {}).get('lista_indirizzo', '') or row.get('address', '') or ''
    return custom_report


BUNDLE_HELPERS = r'''
function bundleReport(){ return APP.custom_report && APP.custom_report.kind==='bundle' ? APP.custom_report : null; }
function bundleRows(){ const rep=bundleReport(); return rep ? (rep.rows||[]) : []; }
function bundleRowForPdv(pdv){ return bundleRows().find(x=>x.pdv===pdv) || null; }
function bundleTone(rate){ if(rate===null || rate===undefined) return ''; if(rate>=60) return 'tone-good'; if(rate<35) return 'tone-bad'; return 'tone-warn'; }
function bundleExportHeaders(){ return ['#','PDV','Città','Indirizzo','Agente','CR','RZV','Vendite tot','Twin','Prospect','Bundle','% Bundle']; }
function bundleExportRows(){
  return filteredCustomRows().slice().sort((a,b)=>((b.bundle_rate ?? -1)-(a.bundle_rate ?? -1)) || ((b.bundle||0)-(a.bundle||0))).map((r,i)=>[
    i+1, r.pdv, r.city, r.address||'', r.agent||'', r.cr||'', r.rzv||'',
    r.sales_total||0, r.twin||0, r.prospect||0, r.bundle||0,
    r.bundle_rate==null ? '' : Number(r.bundle_rate).toFixed(1)+'%'
  ]);
}
function bundleDetailHtml(r){
  const b=bundleRowForPdv(r.pdv), rep=bundleReport();
  if(!b || !rep) return '';
  return `<div class="metric-section">
    <div class="section-title">Bundle</div>
    <div class="metric-row sost">
      <div class="metric-card ${bundleTone(b.bundle_rate)}"><h4>% Bundle<br>su prospect</h4><div class="metric-big">${fmtRate(b.bundle_rate)}</div><div class="metric-sub">${esc(rep.period||'')}</div></div>
      <div class="metric-card"><h4>Bundle<br>completi</h4><div class="metric-big">${fmtNum(b.bundle)}</div><div class="metric-sub">Su ${fmtNum(b.prospect)} prospect</div></div>
      <div class="metric-card"><h4>Vendite periodo</h4><div class="metric-big">${fmtNum(b.sales_total)}</div><div class="metric-sub">Twin: ${fmtNum(b.twin)}</div></div>
    </div>
  </div>`;
}
'''


RENDER_BUNDLE_PAGE = r'''
function renderGarePdv(){
  const w=document.getElementById('garePdvWrap');
  const rep=bundleReport();
  if(!rep){ w.innerHTML='<div class="empty">Nessun file Bundle caricato in input/custom_report.</div>'; return; }
  const rows=filteredCustomRows().slice().sort((a,b)=>((b.bundle_rate ?? -1)-(a.bundle_rate ?? -1)) || ((b.bundle||0)-(a.bundle||0)));
  const sales=rows.reduce((a,r)=>a+(r.sales_total||0),0);
  const prospect=rows.reduce((a,r)=>a+(r.prospect||0),0);
  const bundle=rows.reduce((a,r)=>a+(r.bundle||0),0);
  const rate=prospect ? (bundle/prospect*100) : null;
  const body=rows.map((r,i)=>`<tr>
    <td class="num">${fmtNum(i+1)}</td><td><button class="btn light small" onclick="openDetail('${r.pdv}')">Apri</button></td>
    <td><b>${esc(r.pdv)}</b></td><td><div class="city-cell"><div class="city-main">${esc(r.city)}</div>${r.address ? '<div class="city-address">' + esc(r.address) + '</div>' : ''}</div></td>
    <td>${esc(r.agent||'')}</td><td>${esc(r.cr||'')}</td><td>${esc(r.rzv||'')}</td>
    <td class="num">${fmtNum(r.sales_total)}</td><td class="num">${fmtNum(r.twin)}</td><td class="num">${fmtNum(r.prospect)}</td>
    <td class="num"><b>${fmtNum(r.bundle)}</b></td><td class="num"><b>${fmtRate(r.bundle_rate)}</b></td>
  </tr>`).join('');
  w.innerHTML=`<div class="card">
    <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;flex-wrap:wrap">
      <div><div class="section-title" style="margin:0 0 6px 0">${esc(rep.title||'Bundle')}</div>
      <div class="small-muted">Aggiornato: ${esc(rep.updated_at||'')} · PDV filtrati: ${fmtNum(rows.length)} · File: ${esc(rep.source_name||'')}</div></div>
      <button class="btn light" onclick="downloadFilteredCustomReportCsv()">Scarica Bundle filtrato</button>
    </div>
    <div class="metric-row sost" style="margin-top:12px">
      <div class="metric-card"><h4>Vendite periodo</h4><div class="metric-big">${fmtNum(sales)}</div></div>
      <div class="metric-card"><h4>Bundle completi</h4><div class="metric-big">${fmtNum(bundle)}</div></div>
      <div class="metric-card ${bundleTone(rate)}"><h4>% Bundle filtrata</h4><div class="metric-big">${fmtRate(rate)}</div><div class="metric-sub">Su ${fmtNum(prospect)} prospect</div></div>
    </div>
    <div class="list-wrap" style="margin-top:12px"><table>
      <thead><tr><th class="num">#</th><th></th><th>PDV</th><th>Città</th><th>Agente</th><th>CR</th><th>RZV</th><th class="num">Vendite tot</th><th class="num">Twin</th><th class="num">Prospect</th><th class="num">Bundle</th><th class="num">% Bundle</th></tr></thead>
      <tbody>${body || '<tr><td colspan="12">Nessun PDV sul filtro scelto.</td></tr>'}</tbody>
    </table></div>
  </div>`;
}
'''


def patch_html_bundle(html):
    html = ORIGINAL_PATCH_HTML(html)
    html = html.replace('<button data-page="gare-pdv" onclick="showPage(\'gare-pdv\', this)">Gare PDV</button>', '<button data-page="gare-pdv" onclick="showPage(\'gare-pdv\', this)">Bundle</button>')
    html = html.replace('<button data-page="gare-agenti" onclick="showPage(\'gare-agenti\', this)">Gare agenti</button>', '')
    html = html.replace('<div class="section-title">Gare PDV</div>', '<div class="section-title">Bundle</div>')
    html = html.replace(
        "function pdvObject(){ if(!selectedPdv) return null; return DATA.find(x=>x.pdv===selectedPdv) || null; }",
        "function pdvObject(){ if(!selectedPdv) return null; return DATA.find(x=>x.pdv===selectedPdv) || null; }\n" + BUNDLE_HELPERS,
        1,
    )
    html = html.replace(
        '      </div>\n\n      <div class="metric-section">\n        <div class="section-title">Assistenze stradali</div>',
        '      </div>\n\n      ${bundleDetailHtml(r)}\n\n      <div class="metric-section">\n        <div class="section-title">Assistenze stradali</div>',
        1,
    )
    html = re.sub(r"function renderGarePdv\(\)\{.*?\n\}\n\nfunction csvEscape", RENDER_BUNDLE_PAGE + "\nfunction csvEscape", html, count=1, flags=re.S)
    html = re.sub(r"function customReportExportRows\(\)\{.*?\n\}\nfunction singlePdvSummaryRows", "function customReportExportRows(){ return bundleExportRows(); }\nfunction singlePdvSummaryRows", html, count=1, flags=re.S)
    html = re.sub(r"function downloadFilteredCustomReportCsv\(\)\{.*?\n\}\nfunction openExcelReport", "function downloadFilteredCustomReportCsv(){ if(!bundleReport()){ alert('Nessun file Bundle disponibile.'); return; } downloadCsv('bundle_filtrato.csv', bundleExportHeaders(), bundleExportRows()); }\nfunction openExcelReport", html, count=1, flags=re.S)
    html = re.sub(r"const headers=\['#','PDV','Città','Indirizzo','Agente','CR','RZV',\.\.\.\(\(rep&&rep\.columns\)\|\|\[\]\),'Totale'\];", "const headers=bundleExportHeaders();", html, count=1)
    html = html.replace('Gare PDV filtrate', 'Bundle filtrato')
    html = html.replace('<title>Gare PDV</title>', '<title>Bundle</title>')
    return html


site.gh.load_custom_report = load_bundle_report
site.add_address_to_custom_report = keep_address
site.patch_html_address_display = patch_html_bundle

if __name__ == '__main__':
    site.main()
