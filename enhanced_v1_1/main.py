"""Daily Exception: Power BI + Oracle -> Excel. Python 3.11+."""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import date, datetime
from decimal import Decimal
import getpass
import json
import os
from pathlib import Path
import re
import tempfile
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
import xlsxwriter

from oor_processing import (
    prepare_po,
    consolidate,
    lookup_comments,
    write_po,
    write_summary,
    text,
)

ROOT = Path(__file__).resolve().parent
EXTRA = [
    'REVISION FORMULA', 'LAST FLOOR TRANS. QTY', 'LAST FLOOR TRANS. DATE',
    'DAYS ELAPSED SINCE LAST TRANS.', 'ESTIMATED FLOOR DOH',
    'ESTIMATED FLOOR QOH', 'INVENTORY + LAST TRANS. (WEEKS)',
    'WEEKS ELAPSED SINCE LAST TRANS.', 'REMAINING WEEKS OF COVERAGE (DELTA)',
    'REVISION COMMENT',
]
VISIBLE = {
    'BUYER', 'ORG', 'PRIORITY NUMBER', 'PRIORITY EXCEPTION', 'EXCEPTION TYPE',
    'ITEM', 'DESCRIPTION', 'COMMENT APP', 'COMMENTBY APP', 'ABC', 'PLANNER',
    'ITEM TYPE', 'PREPROCESSING LEAD TIME', 'PROCESSING LEAD TIME',
    'SAFETY STOCK', 'BELOW SAFETY STOCK', 'SHORTAGE', 'DAILY USAGE', 'AWU',
    'FWU', 'PLANNING TIME FENCE', 'TOTAL INVENTORY', 'NON NETTABLE QTY',
    'TOTAL DEMAND CURRENT WEEK', 'TOTAL DEMAND WEEK 2', 'TOTAL DEMAND WEEK 3',
    'TOTAL DEMAND WEEK 4', 'TOTAL DEMAND (4 WEEKS)', 'PO IN RECEIVING',
    'OPEN PO CURRENT WEEK', 'OPEN PO WEEK 2', 'OPEN PO WEEK 3', 'OPEN PO WEEK 4',
    'PURCHASE ORDER DETAILS (DOCK DATE)',
    'PAST LATE SUPPLY PEGGED TO SALES ORDER (LSPSO)', 'LSPSO CURRENT WEEK',
    'LSPSO WEEK 2', 'LSPSO WEEK 3', 'LSPSO WEEK 4', 'LSPSO DETAILS', 'LSPSO ORDERS',
}
SQL_HEADERS = ['ORG', 'ITEM_NAME', 'NEWEST_TRANSACTION_DATE', 'TOTAL_QTY_LAST_24H']


def key(value):
    # Case-insensitive like exact VLOOKUP; preserve spaces and leading zeros.
    return '' if value is None else str(value).casefold()


def read_excel(path, sheet=None):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet] if sheet else wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        headers = [str(x or '').strip() for x in next(it)]
        if not all(headers) or len(set(headers)) != len(headers):
            raise ValueError('Los encabezados deben ser únicos y no estar vacíos.')
        return headers, [dict(zip(headers, row)) for row in it]
    finally:
        wb.close()


def require_headers(headers, required):
    missing = set(required) - set(headers)
    if missing:
        raise ValueError('Faltan columnas: ' + ', '.join(sorted(missing)))


def parse_date(value):
    if value is None or value == '':
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    try:
        result = datetime.fromisoformat(str(value).replace('Z', '+00:00'))
    except ValueError as exc:
        raise ValueError('Fecha Oracle inválida; se esperaba fecha o ISO 8601.') from exc
    if result.tzinfo:
        raise ValueError('La fecha Oracle tiene zona horaria; definir conversión antes de exportar.')
    return result


def number(value):
    return float(value or 0)


def oracle_rows(cfg):
    import oracledb
    settings = cfg['oracle']
    if settings.get('mode', 'thin') == 'thick':
        kwargs = {}
        if settings.get('client_lib_dir'):
            kwargs['lib_dir'] = settings['client_lib_dir']
        if settings.get('config_dir'):
            kwargs['config_dir'] = settings['config_dir']
        oracledb.init_oracle_client(**kwargs)
    dsn = settings.get('dsn')
    if not dsn:
        kind = settings.get('connection_type', 'service_name')
        if kind not in ('sid', 'service_name'):
            raise ValueError('connection_type debe ser sid o service_name.')
        dsn = oracledb.makedsn(settings['host'], int(settings.get('port', 1521)),
                               **{kind: settings[kind]})
    user = os.environ.get('ORACLE_USER') or settings.get('user') or input('Usuario Oracle: ')
    password = os.environ.get('ORACLE_PASSWORD') or getpass.getpass('Contraseña Oracle: ')
    sql = (ROOT / 'transactions.sql').read_text(encoding='utf-8').strip().rstrip(';')
    with oracledb.connect(user=user, password=password, dsn=dsn) as conn:
        with conn.cursor() as cursor:
            cursor.execute('SET TRANSACTION READ ONLY')
            cursor.arraysize = 1000
            cursor.execute(sql)
            headers = [d[0].upper() for d in cursor.description]
            require_headers(headers, SQL_HEADERS)
            return [dict(zip(headers, row)) for row in cursor]


def powerbi_rows(cfg):
    import msal
    import requests
    settings = cfg['powerbi']
    for name in ('tenant_id', 'client_id', 'dataset_id'):
        if not settings.get(name):
            raise ValueError('Configura powerbi.' + name + ' en config.json.')
    query_path = ROOT / settings.get('query_file', 'query.dax')
    if not query_path.exists():
        raise ValueError('Falta query.dax con la consulta REAL del visual de Power BI.')
    query = query_path.read_text(encoding='utf-8').strip()
    if not query:
        raise ValueError('query.dax está vacío.')
    app = msal.PublicClientApplication(
        settings['client_id'],
        authority='https://login.microsoftonline.com/' + settings['tenant_id'],
    )
    token = app.acquire_token_interactive(
        scopes=['https://analysis.windows.net/powerbi/api/Dataset.Read.All'],
        timeout=300,
    )
    if 'access_token' not in token:
        raise ValueError('No se pudo iniciar sesión en Power BI: ' + str(token.get('error', 'sin token')))
    dataset = settings['dataset_id']
    # Only a UUID may enter the API path.
    import uuid
    dataset = str(uuid.UUID(dataset))
    response = requests.post(
        f'https://api.powerbi.com/v1.0/myorg/datasets/{dataset}/executeQueries',
        headers={'Authorization': 'Bearer ' + token['access_token']},
        json={'queries': [{'query': query}], 'serializerSettings': {'includeNulls': True}},
        timeout=(30, 180),
    )
    if not response.ok:
        raise ValueError(f'Power BI HTTP {response.status_code}; revisar permisos y consulta DAX.')
    payload = response.json()
    results = payload.get('results', [])
    if payload.get('error') or len(results) != 1 or results[0].get('error'):
        raise ValueError('Power BI devolvió un error o resultado incompleto.')
    tables = results[0].get('tables', [])
    if len(tables) != 1 or tables[0].get('error'):
        raise ValueError('Se esperaba una sola tabla completa de Power BI.')
    rows = tables[0].get('rows', [])
    if not rows:
        raise ValueError('Power BI no devolvió filas; no se creará un reporte vacío.')
    if len(rows) >= 100000 or sum(len(r) for r in rows) >= 1000000:
        raise ValueError('Resultado al límite de la API; dividir la consulta antes de continuar.')
    expected = json.loads((ROOT / 'raw_headers.json').read_text(encoding='utf-8'))
    mapping = settings.get('column_mapping', {})
    all_keys = set().union(*(r.keys() for r in rows))
    resolved = {}
    for header in expected:
        candidates = [k for k in all_keys if k == header or k == '[' + header + ']'
                      or k.endswith('[' + header + ']')]
        chosen = mapping.get(header)
        if chosen:
            if chosen not in all_keys:
                raise ValueError('No existe el campo DAX configurado para ' + header)
        elif len(candidates) == 1:
            chosen = candidates[0]
        else:
            raise ValueError('Configura column_mapping para la columna: ' + header)
        resolved[header] = chosen
    converted = [{h: row.get(k) for h, k in resolved.items()} for row in rows]
    for row in converted:
        for header in ('PLAN DATE', 'INVENTORY USE UP DATE'):
            if isinstance(row.get(header), str) and row[header]:
                row[header] = parse_date(row[header])
    return expected, converted


def workbook(path):
    return xlsxwriter.Workbook(str(path), {
        'strings_to_formulas': False, 'strings_to_urls': False,
        'default_date_format': 'yyyy-mm-dd hh:mm:ss',
    })


def snapshot(path, headers, rows):
    with workbook(path) as wb:
        ws = wb.add_worksheet('Enhance')
        ws.write_row(0, 0, headers)
        for i, row in enumerate(rows, 1):
            ws.write_row(i, 0, [row.get(h) for h in headers])


def generate(raw_headers, raw_rows, sql_rows, output, lookup="item", missing="blank", po=None, summary=None):
    require_headers(raw_headers, ['ITEM', 'ORG', 'TOTAL INVENTORY', 'PO IN RECEIVING',
                                 'TOTAL DEMAND CURRENT WEEK', 'TOTAL DEMAND WEEK 2', 'AWU'])
    if set(EXTRA) & set(raw_headers):
        raise ValueError('El archivo RAW ya contiene columnas de revisión; usar el original.')
    if lookup not in ('item', 'org_item') or missing not in ('blank', 'na'):
        raise ValueError('Modo de búsqueda o valor sin coincidencia inválido.')
    require_headers(list(sql_rows[0]) if sql_rows else SQL_HEADERS, SQL_HEADERS)
    index = {}
    duplicates = 0
    def make_key(row, source=False):
        item = key(row.get('ITEM_NAME' if source else 'ITEM'))
        return item if lookup == 'item' else (key(row.get('ORG')), item)
    for row in sql_rows:
        if row.get('ITEM_NAME') in (None, ''):
            continue
        k = make_key(row, True)
        if k in index:
            duplicates += 1
        else:
            index[k] = row  # Exact VLOOKUP: first occurrence, no aggregation.
    records = [r for r in raw_rows if r.get('ITEM') not in (None, '')]
    if not records:
        raise ValueError('RAW no contiene ítems.')
    headers = raw_headers + EXTRA

    # CY es la posición final después de eliminar la auxiliar CP.
    if len(headers) > 102:
        raise ValueError(
            "El esquema de Daily Exception cambió: "
            "BUYER COMMENTS en CY desplazaría columnas existentes."
        )

    headers += [""] * (102 - len(headers))
    headers.append("BUYER COMMENTS")
    from xlsxwriter.utility import xl_col_to_name
    columns = {h: i for i, h in enumerate(headers)}
    stats = {'raw_rows': len(records), 'excluded_footer_or_empty': len(raw_rows)-len(records),
             'sql_rows': len(sql_rows), 'matched': 0, 'unmatched': 0,
             'cross_org_matches': 0, 'duplicate_sql_keys': duplicates,
             'lookup': lookup, 'unmatched_value': missing}
    with workbook(output) as wb:
        wb.set_calc_mode('auto')
        ws = wb.add_worksheet('Enhance')
        ws.hide_gridlines(2)
        base = {'font_name': 'Segoe UI', 'font_size': 9, 'border': 1, 'border_color': '#000000'}
        formats = {}
        def fmt(bg=None, white=False, header=False, date_format=False):
            cache_key = (bg, white, header, date_format)
            if cache_key not in formats:
                props = dict(base)
                if bg: props['bg_color'] = bg
                if white: props['font_color'] = '#FFFFFF'
                if header: props['text_wrap'] = True
                if date_format: props['num_format'] = 'm/dd/yyyy   '
                formats[cache_key] = wb.add_format(props)
            return formats[cache_key]
        for i, h in enumerate(headers):
            color = '#4C68A2'
            for token, fill in [('demand', '#00008B'), ('open po', '#006400'),
                                ('lspso', '#E26B0A'), ('shortage', '#0070C0')]:
                if token in h.lower():
                    color = fill
                    break
            if h == EXTRA[0]: color = '#D86DCD'
            elif h in EXTRA[1:6]: color = '#D2B48C'
            elif h in EXTRA[6:] or h == "BUYER COMMENTS":
              color = "#384C70"
            ws.write(0, i, h, fmt(color, True, True))

            hidden_columns = {
                'DAYS ELAPSED SINCE LAST TRANS.',
                'ESTIMATED FLOOR DOH',
                'ESTIMATED FLOOR QOH',
}
            ws.set_column_pixels(
                i,
                i,
                40,
                None,
                {
                    "hidden": (
                        h in hidden_columns
                        or (
                            h not in VISIBLE
                            and h not in EXTRA
                            and h != "BUYER COMMENTS"
                        )
                    )
                },
            )
        ws.set_row(0, 45)
        for rownum, row in enumerate(records, 1):
            ws.set_row(rownum, 15)
            match = index.get(make_key(row))
            stats['matched' if match else 'unmatched'] += 1
            if match and key(match.get('ORG')) != key(row.get('ORG')):
                stats['cross_org_matches'] += 1
            excel_row = rownum + 1
            def ref(h): return f"Enhance!${xl_col_to_name(columns[h])}{excel_row}"            
            for col, h in enumerate(raw_headers):
                value = row.get(h)
                fill, white = None, False
                try:
                    if h == 'PRIORITY NUMBER':
                        m = re.match(r'^[+-]?(?:\d+\.?\d*|\.\d+)', str(value))
                        fill = {1: '#FF0000', 2: '#FFFF00', 3: '#FFA500', 4: '#008000'}.get(float(m[0])) if m else None
                    if h == 'TOTAL INVENTORY' and value is not None and number(value) <= 0:
                        fill, white = '#FF0000', True
                except (ValueError, TypeError):
                    pass
                ws.write(rownum, col, value, fmt(fill, white, date_format=isinstance(value, (date, datetime))))
            for h in EXTRA:
                ws.write_blank(rownum, columns[h], None, fmt())
            qty = -float(match['TOTAL_QTY_LAST_24H']) if match and match.get('TOTAL_QTY_LAST_24H') is not None else None
            txn_date = parse_date(match.get('NEWEST_TRANSACTION_DATE')) if match else None
            for h, value in [(EXTRA[1], qty), (EXTRA[2], txn_date)]:
                if not match and missing == 'na':
                    ws.write_formula(rownum, columns[h], '=NA()', fmt(), '#N/A')
                else:
                    ws.write(rownum, columns[h], value, fmt(date_format=h == EXTRA[2]))
            inv, receiving = number(row.get('TOTAL INVENTORY')), number(row.get('PO IN RECEIVING'))
            demand = number(row.get('TOTAL DEMAND CURRENT WEEK')) + number(row.get('TOTAL DEMAND WEEK 2'))
            awu = number(row.get('AWU'))
            formulas = {
                EXTRA[0]: (f'=IF(({ref("TOTAL INVENTORY")}+{ref("PO IN RECEIVING")})>({ref("TOTAL DEMAND CURRENT WEEK")}+{ref("TOTAL DEMAND WEEK 2")}),"OK","REVIEW")', 'OK' if inv+receiving > demand else 'REVIEW'),
                EXTRA[6]: (f'=IFERROR(({ref("TOTAL INVENTORY")}+ABS({ref(EXTRA[1])}))/{ref("AWU")},0)', (inv+abs(qty or 0))/awu if awu and (match or missing == 'blank') else 0),
                EXTRA[7]: (f'=IFERROR({ref(EXTRA[3])}/5,0)', 0),
                EXTRA[8]: (f'={ref(EXTRA[3])}/5', 0),
            }
            for h, (formula, cached) in formulas.items():
                ws.write_formula(
                    rownum,
                    columns[h],
                    formula,
                    fmt(),
                    cached,
                )

        # Ya terminó el bucle que escribe los artículos de Enhance.
        comment_index = lookup_comments(summary or [])

        comment_format = wb.add_format({
            **base,
            "text_wrap": True,
            "valign": "top",
        })

        for rownum, row in enumerate(records, 1):
            concat_key = key(
                text(row.get("ORG"))
                + text(row.get("ITEM"))
            )

            message = comment_index.get(concat_key)

            # Se guarda texto fijo: equivalente a pegar valores
            # después del VLOOKUP.
            if message is None or message.strip().upper() in {"#N/A", "N/A"}:
                ws.write_blank(
                    rownum,
                    columns["BUYER COMMENTS"],
                    None,
                    comment_format,
                )
            else:
                ws.write_string(
                    rownum,
                    columns["BUYER COMMENTS"],
                    message,
                    comment_format,
                )
        # Agregar las otras dos hojas al mismo libro.
        write_po(wb, po or [])
        write_summary(wb, summary or [])

        stats["po_rows"] = len(po or [])
        stats["summary_rows"] = len(summary or [])

    return stats


def main():
    parser = argparse.ArgumentParser(description=__doc__)

    parser.add_argument(
        '--config',
        type=Path,
        default=ROOT / 'config.json',
    )
    parser.add_argument(
        '--raw',
        type=Path,
        help='RAW local; omitir para consultar Power BI.',
    )
    parser.add_argument(
    "--po",
    type=Path,
    required=True,
    help="Archivo RAW DATA OPEN PO EXCEPTIONS exportado de Power BI.",
)
    parser.add_argument(
        '--enhance',
        type=Path,
        help='ENHANCE local; omitir para consultar Oracle.',
    )
    parser.add_argument(
        '--output-dir',
        type=Path,
        default=Path("C:/Users/E1568382/OneDrive - Emerson/TF Purchasing Analytics - Documents/General Reports/2. Inventory Management/7. Enhanced Daily Exception Messages"),
        help='Carpeta para el Excel final.',
    )
    parser.add_argument(
        '--log-dir',
        type=Path,
        default=ROOT / 'logs',
        help='Carpeta para resultado.json.',
    )
    parser.add_argument(
        '--lookup',
        choices=['item', 'org_item'],
        default='item',
    )
    parser.add_argument(
        '--missing',
        choices=['blank', 'na'],
        default='blank',
    )

    args = parser.parse_args()

    output_dir = args.output_dir.expanduser().resolve()
    log_dir = args.log_dir.expanduser().resolve()

    if output_dir == log_dir:
        raise ValueError(
            'La carpeta de logs debe ser diferente a la del Excel final.'
        )

    cfg = (
        json.loads(args.config.read_text(encoding='utf-8'))
        if args.config.exists()
        else {}
    )

    print('1/4 Obteniendo datos de Power BI o RAW local...')

    if args.raw:
        headers, raw = read_excel(args.raw, 'Export')
    else:
        headers, raw = powerbi_rows(cfg)

    print("Preparando Open PO y consolidando comentarios...")

    po_headers, po_raw = read_excel(args.po, "Export")
    po = prepare_po(po_headers, po_raw)
    summary = consolidate(po)

    print('2/4 Obteniendo datos de Oracle o ENHANCE local...')

    if args.enhance:
        sql_headers, sql = read_excel(args.enhance)
        require_headers(sql_headers, SQL_HEADERS)
    else:
        sql = oracle_rows(cfg)

    now = datetime.now(ZoneInfo('America/Mexico_City'))

    output_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    report_name = (
        f'Revision Enhanced Daily Exception Messages '
        f'{now:%Y-%m-%d}.xlsx'
    )
    final = output_dir / report_name

    # Si ya existe un reporte del día, conservarlo y agregar hora
    # al nombre del nuevo archivo.
    if final.exists():
        final = output_dir / (
            f'Revision Enhanced Daily Exception Messages '
            f'{now:%Y-%m-%d}.xlsx'
        )

    run_log_dir = Path(
        tempfile.mkdtemp(
            prefix=now.strftime('%Y-%m-%d_%H%M%S_'),
            dir=log_dir,
        )
    )

    # Archivo temporal en el mismo destino para publicar el
    # reporte únicamente cuando la generación haya terminado.
    descriptor, temporary_name = tempfile.mkstemp(
        prefix='_pending_',
        suffix='.xlsx',
        dir=output_dir,
    )
    os.close(descriptor)
    stage = Path(temporary_name)

    print('3/4 Creando reporte y aplicando formato...')

    try:
        stats = generate(
            headers,
            raw,
            sql,
            stage,
            args.lookup,
            args.missing,
            po=po,
            summary=summary,
        )
        stage.replace(final)

    finally:
        stage.unlink(missing_ok=True)

    stats['created_at'] = now.isoformat()
    stats['report'] = str(final)

    log_file = run_log_dir / 'resultado.json'
    log_file.write_text(
        json.dumps(stats, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )

    print('4/4 Finalizado:', final)
    print('Log guardado en:', log_file)
    print(json.dumps(stats, ensure_ascii=False, indent=2))

    if stats['cross_org_matches']:
        print(
            'AVISO: búsqueda solo por ITEM produjo coincidencias '
            'entre organizaciones; revisar resultado.json.'
        )
if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        raise SystemExit(f'ERROR: {exc}') from None
