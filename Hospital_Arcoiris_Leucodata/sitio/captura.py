from collections import defaultdict
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from tempfile import NamedTemporaryFile
from urllib.parse import urlsplit, parse_qs
from uuid import uuid4
import json
import os
import webbrowser

from openpyxl import load_workbook # pyright: ignore[reportMissingModuleSource]
import openpyxl.utils.cell # type: ignore

ROOT = Path(__file__).resolve().parent
EXCEL = ROOT / 'Hospital_Infantil_Arcoiris.xlsx'
URL = 'http://127.0.0.1:8765'
FORMATS = {'Fecha': 'dd/mm/yyyy', 'Fecha y hora': 'dd/mm/yyyy hh:mm',
           'Hora': 'hh:mm', 'Importe MXN': '"$"#,##0.00', 'Número': '0.##########'}


def schema(book):
    catalogs = defaultdict(list)
    for group, code, label in book['Catalogos'].iter_rows(min_row=4, values_only=True):
        if group and code:
            catalogs[group].append({'code': code, 'label': label or code})
    tables = {}
    for row in book['Inicio'].iter_rows(values_only=True):
        group, sheet, table, label = row[:4]
        if isinstance(table, str) and table.startswith('tbl_') and sheet in book.sheetnames:
            tables[table] = dict(key=table, group=group, sheet=sheet, label=label, fields=[])
    for row in book['Diccionario'].iter_rows(min_row=4, values_only=True):
        table, name, kind, required, key, ref, catalog, note = row[:8]
        if table in tables:
            tables[table]['fields'].append(dict(name=name, kind=kind, required=required=='Sí',
                primary=key=='Primaria', ref=ref, options=catalogs.get(catalog, []), note=note))
    return tables


def records(book, model):
    sheet = book[model['sheet']]
    table = sheet.tables[model['key']]
    left, top, right, bottom = openpyxl.utils.cell.range_boundaries(table.ref)
    fields = [f['name'] for f in model['fields']]
    if [sheet.cell(top, c).value for c in range(left, right+1)] != fields:
        raise ValueError('Los encabezados de la tabla y el diccionario no coinciden.')
    return (sheet, table, left, top, right, bottom, fields)


def choices(book, model, field):
    sheet, _, left, top, right, bottom, fields = records(book, model)
    key_idx = fields.index(field)
    label_fields = {
        'Pacientes': ('nombres', 'apellidos'), 'Tutores': ('nombres', 'apellidos'),
        'Empleados': ('nombres', 'apellidos'), 'Areas': ('nombre',),
        'Consultorios': ('nombre',), 'Camas': ('codigo', 'habitacion'),
        'Medicamentos': ('nombre_generico', 'presentacion'),
        'Estudios': ('nombre_estudio',), 'Diagnosticos': ('descripcion',),
        'LeucoData_Muestras': ('codigo_muestra',), 'LeucoData_Pruebas': ('nombre_prueba',),
        'Citas': ('inicio', 'estado'), 'Leucemia_Casos': ('tipo_leucemia', 'estado_caso')
    }.get(model['sheet'], ())
    labels = [fields.index(k) for k in label_fields]
    result = []
    for row in sheet.iter_rows(min_row=top+1, max_row=bottom, min_col=left, max_col=right, values_only=True):
        if row[key_idx] is not None:
            code = str(row[key_idx])
            detail = ' · '.join(str(row[i]) for i in labels if row[i] is not None)
            result.append({'code': code, 'label': f'{detail} — {code}' if detail else code})
    return result


def convert(raw, kind):
    if kind == 'Fecha':
        return date.fromisoformat(raw)
    if kind == 'Fecha y hora':
        value = datetime.fromisoformat(raw)
        if value.tzinfo or 'T' not in raw:
            raise ValueError()
        return value
    if kind == 'Hora':
        value = time.fromisoformat(raw)
        if value.tzinfo:
            raise ValueError()
        return value
    if kind in ('Número', 'Importe MXN'):
        value = Decimal(raw)
        if not value.is_finite() or abs(value) > Decimal('1e100'):
            raise ValueError()
        return float(value)
    return raw


def save_record(book, tables, key, data, *, persist=True):
    model = tables[key]
    if not isinstance(data, dict) or set(data) - {f['name'] for f in model['fields']}:
        raise ValueError('El formulario contiene campos desconocidos.')
    values = {}
    for field in model['fields']:
        name = field['name']
        raw = data.get(name, '')
        if not isinstance(raw, str):
            raise ValueError(f'El campo {name} debe enviarse como texto.')
        raw = raw.strip()
        if field['primary']:
            values[name] = uuid4().hex
            record_id = values[name]
            continue
        if not raw:
            if field['required']:
                raise ValueError(f'Completa el campo {name}.')
            continue
        if len(raw) > 32767:
            raise ValueError(f'El campo {name} supera el límite de Excel.')
        if field['ref']:
            target, target_field = field['ref'].split('.', 1)
            if raw not in {o['code'] for o in choices(book, tables[target], target_field)}:
                raise ValueError(f'Selecciona un registro existente en {name}.')
        if field['options'] and raw not in {o['code'] for o in field['options']}:
            raise ValueError(f'Selecciona un valor del catálogo en {name}.')
        try:
            values[name] = convert(raw, field['kind'])
        except (ValueError, InvalidOperation, OverflowError):
            raise ValueError(f'Revisa el formato de {name}.') from None

    if key == 'tbl_Solicitudes_Web' and not (values.get('telefono') or values.get('correo')):
        raise ValueError('Captura un teléfono o correo de contacto.')
    if key in ('tbl_Citas', 'tbl_Horarios'):
        start, end = ('inicio', 'fin') if key == 'tbl_Citas' else ('hora_inicio', 'hora_fin')
        if values[end] <= values[start]:
            raise ValueError('La hora de fin debe ser posterior al inicio.')

    sheet, table, left, top, right, bottom, fields = records(book, model)
    target_row = bottom+1
    if bottom > top and all(sheet.cell(bottom, c).value is None for c in range(left, right+1)):
        target_row = bottom
    if any(sheet.cell(target_row, c).value is not None for c in range(left, right+1)):
        raise ValueError('Hay contenido debajo de la tabla. Libera esa fila en Excel.')
    for column, field in enumerate(model['fields'], left):
        cell = sheet.cell(target_row, column)
        value = values.get(field['name'])
        cell.value = value
        if isinstance(value, str):
            cell.data_type = 's'
        cell.number_format = FORMATS.get(field['kind'], '@')
    if target_row > bottom:
        table.ref = f'{openpyxl.utils.cell.get_column_letter(left)}{top}:{openpyxl.utils.cell.get_column_letter(right)}{target_row}'
        if table.autoFilter:
            table.autoFilter.ref = table.ref

    if persist:
        write_workbook(book)
    return record_id


def write_workbook(book):
    with NamedTemporaryFile(dir=EXCEL.parent, suffix='.xlsx', delete=False) as handle:
        temp = Path(handle.name)
    try:
        book.save(temp)
        os.replace(temp, EXCEL)
    finally:
        temp.unlink(missing_ok=True)


class Handler(BaseHTTPRequestHandler):
    def reply(self, code, body, mime='application/json; charset=utf-8'):
        self.send_response(code)
        self.send_header('Content-Type', mime)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def json(self, code, payload):
        self.reply(code, json.dumps(payload, ensure_ascii=False).encode('utf-8'))

    def local_host(self):
        return self.headers.get('Host') == '127.0.0.1:8765'

    def do_GET(self):
        if not self.local_host():
            return self.json(403, {'error': 'Abre el formulario en ' + URL})
        route = urlsplit(self.path)
        if route.path == '/':
            return self.reply(200, (ROOT/'index.html').read_bytes(), 'text/html; charset=utf-8')
        if route.path != '/api/form':
            return self.json(404, {'error': 'Ruta desconocida.'})
        try:
            with open(EXCEL, 'rb') as source:
                book = load_workbook(source)
            tables = schema(book)
            key = parse_qs(route.query).get('tabla', [''])[0]
            if not key:
                return self.json(200, {'tables': list(tables.values())})
            model = tables[key]
            for field in model['fields']:
                if field['ref']:
                    target, name = field['ref'].split('.', 1)
                    field['options'] = choices(book, tables[target], name)
                    field['targetSheet'] = tables[target]['sheet']
            self.json(200, {'model': model})
        except KeyError:
            self.json(404, {'error': 'No se encontró la tabla o referencia en el archivo.'})
        except (OSError, ValueError):
            self.json(400, {'error': 'Revisa que el Excel original esté en la carpeta y se pueda abrir.'})

    def do_POST(self):
        if self.path != '/api/form':
            return self.json(404, {'error': 'Ruta desconocida.'})
        if not self.local_host() or self.headers.get('Origin') != URL:
            return self.json(403, {'error': 'Abre el formulario desde ' + URL})
        if self.headers.get_content_type() != 'application/json':
            return self.json(415, {'error': 'Formato de solicitud incorrecto.'})
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= 250000:
                raise ValueError('Tamaño de solicitud incorrecto.')
            payload = json.loads(self.rfile.read(length))
            if not isinstance(payload, dict) or not isinstance(payload.get('tabla'), str):
                raise ValueError('Solicitud de captura incorrecta.')
            with open(EXCEL, 'rb') as source:
                book = load_workbook(source)
            key = payload['tabla']
            record_id = save_record(book, schema(book), key, payload.get('datos'))
            self.json(201, {'id': record_id})
        except KeyError:
            self.json(404, {'error': 'No se encontró la tabla o referencia en el archivo.'})
        except ValueError as exc:
            self.json(400, {'error': str(exc)})
        except OSError:
            self.json(409, {'error': 'No se pudo guardar. Cierra el Excel y vuelve a intentar.'})


if __name__ == '__main__':
    with HTTPServer(('127.0.0.1', 8765), Handler) as server:
        print(f'Formulario Arcoíris disponible en {URL}. Presiona Ctrl+C para cerrar.', flush=True)
        webbrowser.open(URL)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
