"""Python translation of OOR Formatter and Temporal OOR Comments.

Consolidation is materialized as text: no Excel calculation engine is needed.
"""
from collections import OrderedDict
from datetime import datetime, date, timedelta
from decimal import Decimal, ROUND_HALF_UP
import json
from pathlib import Path
import re

from openpyxl.utils.datetime import to_excel
from xlsxwriter.utility import xl_col_to_name

COLUMNS = json.loads((Path(__file__).parent / 'oor_columns.json').read_text())
FORMATS = {'Text': '@', 'General': 'General', 'Date': 'dd-mmm-yyyy',
           'Number': '#,##0', 'Currency': '$#,##0.00'}
MONTHS = dict(zip('JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split(), range(1, 13)))
MONTHS.update(ENE=1, ABR=4, AGO=8, DIC=12)
MONTH_NAMES = 'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split()


def text(value):
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def norm(value):
    return re.sub(r' +', ' ', text(value).replace('\u00a0', ' ').replace('\t', ' ')).strip().upper()


def date_serial(value):
    if isinstance(value, (datetime, date)):
        return to_excel(datetime.combine(value, datetime.min.time()) if not isinstance(value, datetime) else value)
    if not isinstance(value, str) or not value.strip():
        return value
    part = re.split(r'\s|T(?=\d{2}:)', value.strip())[0]
    p = re.sub(r'[.\-]', '/', part).split('/')
    if len(p) != 3:
        return value
    try:
        if re.fullmatch(r'\d{4}', p[0]):
            y, m, d = map(int, p)
        elif re.fullmatch(r'\d{2}(\d{2})?', p[2]):
            y = int(p[2]) + (2000 if len(p[2]) == 2 else 0)
            if p[1].isalpha():
                d, m = int(p[0]), MONTHS[p[1][:3].upper()]
            else:
                a, b = int(p[0]), int(p[1])
                d, m = (a, b) if a > 12 else (b, a)
        else:
            return value
        if y < 1900:
            return value
        # Match Office Script's 1899-12-30 epoch.
        return (datetime(y, m, d) - datetime(1899, 12, 30)).days
    except (ValueError, KeyError):
        return value


def prepare_po(headers, raw):
    indices = {}
    for h in headers:
        indices.setdefault(norm(h), h)
    missing = [c['source'] for c in COLUMNS if c['source'] and norm(c['source']) not in indices]
    if missing:
        raise ValueError('Open PO: faltan columnas: ' + ', '.join(missing))
    result = []
    for row in raw:
        if not text(row.get(indices['PURCHASE ORDER / REQUISITION'])).strip():
            continue
        output = {}
        for c in COLUMNS:
            v = row.get(indices[norm(c['source'])]) if c['source'] else None
            output[c['name']] = date_serial(v) if c['format'] == 'Date' else v
        result.append(output)
    if not result:
        raise ValueError('Todas las filas Open PO tienen Purchase Order / Requisition vacío.')
    return result


def quantity_text(value):
    try:
        q = Decimal(text(value) or '0')
        if not q.is_finite():
            raise ValueError('Cantidad no finita')
        if q == q.to_integral_value():
            return f'{q:,.0f}'
        return f'{q.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP):,.2f}'.rstrip('0').rstrip('.')
    except Exception as exc:
        raise ValueError(f'PO Qty inválida: {value!r}') from exc


def consolidated_date(value):
    if value is None or value == '':
        return '', float('inf')
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f'Due Date no interpretable: {value!r}. Revisar la fecha en Open PO.')
    day = datetime(1899, 12, 30) + timedelta(days=value)
    return f'{day.day:02d}-{MONTH_NAMES[day.month-1]}-{day.year % 100:02d}', value


def consolidate(rows):
    # Script's unique list is case-sensitive and trimmed, in first-seen order.
    keys = OrderedDict()
    matches = {}
    for row in rows:
        org, item = text(row['Organization']), text(row['Item Name'])
        if org.strip() and item.strip():
            keys.setdefault((org.strip(), item.strip()), None)
        # Excel FILTER equality ignores case, but does not trim source values.
        matches.setdefault((org.casefold(), item.casefold()), []).append(row)
    summary = []
    for org, item in keys:
        entries = []
        for row in matches.get((org.casefold(), item.casefold()), []):
            eta, order = consolidated_date(row['Due Date'])
            message = f"ETA:{eta} QTY:{quantity_text(row['PO Qty'])} C:{text(row['Buyer Comments'])}"
            entries.append((order, message))
        entries.sort(key=lambda entry: entry[0])  # stable for equal dates
        message = ' | \n'.join(entry[1] for entry in entries)
        if len(message) > 32767:
            raise ValueError(f'Comentarios de {org}/{item} exceden el límite de 32767 caracteres de Excel.')
        summary.append((org, item, message))
    return summary


def lookup_comments(summary):
    result = {}
    owners = {}
    for org, item, message in summary:
        k = (org + item).casefold()  # Exact CONCAT + VLOOKUP semantics.
        pair = (org.casefold(), item.casefold())
        if k in owners and owners[k] != pair:
            raise ValueError(f'CONCAT ambiguo: {owners[k]} y {pair}. No se mezclaron comentarios.')
        owners[k] = pair
        result.setdefault(k, message)
    return result


def calculated(row, name, excel_row):
    positions = {c['name']: i for i, c in enumerate(COLUMNS)}
    def ref(n):
        return f'${xl_col_to_name(positions[n])}{excel_row}'
    if name == 'PO ID':
        names = ['Purchase Order (PO) / Requisition', 'PO Line', 'PO Shipment', 'PO Release']
        return '=CONCAT(' + ',"," ,'.join(ref(n) for n in names) + ')', ','.join(text(row[n]) for n in names)
    if name == 'Empty Promise Date':
        return f'=IF({ref("PO Promise Date")}="","YES","NO")', 'YES' if row['PO Promise Date'] in (None, '') else 'NO'
    detail_map = {'LSPSO Date': 'LSPSO', 'LSPF Date': 'LSPF',
                  'Reschedule In Date': 'Reschedule In Details', 'Reschedule Out Date': 'Reschedule Out Details'}
    if name in detail_map:
        details = detail_map[name]
        f = f'=IFERROR(DATEVALUE(LEFT({ref(details)},FIND(" /",{ref(details)})-1)),"")'
        value = text(row[details])
        parsed = date_serial(value.split(' /', 1)[0]) if ' /' in value else ''
        return f, parsed if isinstance(parsed, (int, float)) else ''
    if name.endswith(' (No Postprocessing LT)'):
        base = name.split(' Date')[0]
        date_col = base + ' Date'
        f = f'=IF(ISBLANK({ref(base)}),"",WORKDAY({ref(date_col)},-{ref("Postprocessing LT")}))'
        if row[base] is None:
            return f, ''
        _, start = calculated(row, date_col, excel_row)
        if start == '':
            return f, '#VALUE!'
        try:
            days = int(-float(row['Postprocessing LT'] or 0))
            dt = datetime(1899, 12, 30) + timedelta(days=int(start))
            direction = 1 if days >= 0 else -1
            remaining = abs(days)
            while remaining:
                dt += timedelta(days=direction)
                if dt.weekday() < 5:
                    remaining -= 1
            return f, (dt - datetime(1899, 12, 30)).days
        except (ValueError, TypeError):
            return f, '#VALUE!'
    raise ValueError(f'Fórmula OOR desconocida: {name}')


def write_po(wb, rows):
    ws = wb.add_worksheet('Export')
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)
    ws.set_column_pixels(0, len(COLUMNS)-1, 74)
    base = {'font_name': 'Segoe UI', 'font_size': 9, 'border': 1, 'border_color': '#000000'}
    formats = {}
    ws.set_row(0, 45)
    for col, c in enumerate(COLUMNS):
        header = wb.add_format({**base, 'bg_color': c['fill'], 'font_color': c['font'],
                                'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
        ws.write_string(0, col, c['name'], header)
        formats[col] = wb.add_format({**base, 'num_format': FORMATS[c['format']]})
    for idx, row in enumerate(rows, 1):
        ws.set_row(idx, 15)
        for col, c in enumerate(COLUMNS):
            if c['source']:
                ws.write(idx, col, row[c['name']], formats[col])
            else:
                formula, cache = calculated(row, c['name'], idx + 1)
                ws.write_formula(idx, col, formula, formats[col], cache)
    ws.autofilter(0, 0, len(rows), len(COLUMNS)-1)


def write_summary(wb, summary):
    ws = wb.add_worksheet('Org Item Summary')
    header = wb.add_format({'bg_color': '#1F4E78', 'font_color': '#FFFFFF', 'bold': True})
    wrap = wb.add_format({'text_wrap': True, 'valign': 'top'})
    ws.write_row(0, 0, ['Organization', 'Item Name', 'Consolidated PO Messages'], header)
    for i, (org, item, message) in enumerate(summary, 1):
        ws.write_string(i, 0, org)
        ws.write_string(i, 1, item)
        ws.write_string(i, 2, message, wrap)
    ws.set_column(0, 0, max([len('Organization')] + [len(r[0]) for r in summary]) + 2)
    ws.set_column(1, 1, min(80, max([len('Item Name')] + [len(r[1]) for r in summary]) + 2))
    ws.set_column_pixels(2, 2, 467)
    ws.freeze_panes(1, 0)
