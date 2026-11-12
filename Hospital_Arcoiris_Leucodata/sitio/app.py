from datetime import datetime
from hashlib import scrypt
from http.server import HTTPServer
from mimetypes import guess_type
from urllib.parse import urlsplit
from uuid import uuid4
from pathlib import Path
import json
import re
import secrets
import unicodedata
import webbrowser

from openpyxl import load_workbook # pyright: ignore[reportMissingModuleSource]
from openpyxl.worksheet.table import Table, TableStyleInfo # pyright: ignore[reportMissingModuleSource]

import captura

ROLES = {'Medico', 'Administrador', 'Personal'}
CREDENTIALS = ('cuenta_id', 'usuario', 'algoritmo', 'sal_hex', 'huella_hex')
PERSONAL = {'area_id', 'nombres', 'apellidos', 'puesto', 'especialidad',
            'cedula_profesional', 'correo_laboral', 'telefono_laboral', 'fecha_alta'}
PATIENT = {'nombres', 'apellidos', 'folio_expediente', 'fecha_nacimiento'}
TUTOR = {'tutor_nombres', 'tutor_apellidos', 'tutor_telefono', 'tutor_correo', 'parentesco'}


def load():
    with captura.EXCEL.open('rb') as file:
        return load_workbook(file)


def table_model(book, sheet_name, fields):
    sheet = book[sheet_name]
    return {'key': 'tbl_' + sheet_name, 'sheet': sheet_name,
            'fields': [{'name': field} for field in fields]}


def append_row(book, model, values):
    sheet, table, left, top, right, bottom, fields = captura.records(book, model)
    row = bottom if all(sheet.cell(bottom, c).value is None for c in range(left, right + 1)) else bottom + 1
    if any(sheet.cell(row, c).value is not None for c in range(left, right + 1)):
        raise ValueError('Libera la fila situada debajo de ' + sheet.title + '.')
    for column, field in enumerate(fields, left):
        cell = sheet.cell(row, column, values.get(field))
        if isinstance(cell.value, str):
            cell.data_type = 's'
            cell.number_format = '@'
        elif isinstance(cell.value, datetime):
            cell.number_format = 'dd/mm/yyyy hh:mm'
    if row > bottom:
        table.ref = f'{captura.get_column_letter(left)}{top}:{captura.get_column_letter(right)}{row}'
        if table.autoFilter:
            table.autoFilter.ref = table.ref


def credential_table(book):
    name = 'Credenciales_Web'
    if name not in book.sheetnames:
        sheet = book.create_sheet(name)
        sheet.append(CREDENTIALS)
        sheet.append([None] * len(CREDENTIALS))
        table = Table(displayName='tbl_' + name, ref='A1:E2')
        table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        sheet.add_table(table)
        sheet.freeze_panes = 'C2'
        for column, width in zip('ABCDE', [37, 37, 25, 35, 70]):
            sheet.column_dimensions[column].width = width
    return table_model(book, name, CREDENTIALS)


def names_for(data, allowed):
    return {key: data[key] for key in allowed if key in data}


def register(book, data):
    if not isinstance(data, dict) or any(not isinstance(v, str) for v in data.values()):
        raise ValueError('Formato de datos incorrecto.')
    data = {key: value.strip() for key, value in data.items()}
    kind = data.get('tipo')
    now = datetime.now().replace(second=0, microsecond=0)
    tables = captura.schema(book)
    account_id = uuid4().hex

    def save(sheet, fields):
        return captura.save_record(book, tables, 'tbl_' + sheet, fields, persist=False)

    if kind == 'personal':
        role = data.get('rol')
        if role not in ROLES:
            raise ValueError('Selecciona el rol del trabajador.')
        person = save('Empleados', {**names_for(data, PERSONAL), 'estado': 'Activo'})
        sheet_name = 'Cuentas_Personal'
        columns = ('cuenta_id', 'empleado_id', 'usuario', 'rol', 'estado', 'creada_en')
        identity = {'empleado_id': person}
        name = data.get('nombres', '') + ' ' + data.get('apellidos', '')
    elif kind == 'paciente':
        dob = captura.date.fromisoformat(data.get('fecha_nacimiento', ''))
        if dob > now.date():
            raise ValueError('La fecha de nacimiento debe ser hoy o anterior.')
        person = save('Pacientes', {**names_for(data, PATIENT), 'estado': 'Activo',
                                    'fecha_registro': now.isoformat()})
        sheet_name = 'Cuentas_Pacientes'
        columns = ('cuenta_id', 'paciente_id', 'tutor_id', 'usuario', 'rol', 'estado', 'creada_en')
        identity = {'paciente_id': person}
        role = 'Paciente'
        name = data.get('nombres', '') + ' ' + data.get('apellidos', '')
        if data.get('titular') == 'tutor':
            if not (data.get('tutor_telefono') or data.get('tutor_correo')):
                raise ValueError('Captura un teléfono o correo del tutor.')
            tutor = save('Tutores', {'nombres': data.get('tutor_nombres', ''),
                'apellidos': data.get('tutor_apellidos', ''), 'telefono': data.get('tutor_telefono', ''),
                'correo': data.get('tutor_correo', ''), 'fecha_registro': now.isoformat()})
            save('Paciente_Tutor', {'paciente_id': person, 'tutor_id': tutor,
                'parentesco': data.get('parentesco', ''), 'es_principal': 'Si',
                'autorizado_portal': 'Si', 'activo': 'Si'})
            identity['tutor_id'] = tutor
            role = 'Tutor'
            name = data['tutor_nombres'] + ' ' + data['tutor_apellidos']
    else:
        raise ValueError('Selecciona un tipo de registro.')

    ascii_name = unicodedata.normalize('NFKD', name.casefold()).encode('ascii', 'ignore').decode()
    base = re.sub(r'[^a-z0-9]+', '.', ascii_name).strip('.')[:30] or 'cuenta'
    account_model = table_model(book, sheet_name, columns)
    credential_model = credential_table(book)
    taken = {option['code'].casefold() for model in (
        account_model, credential_model,
        table_model(book, 'Cuentas_Personal', ('cuenta_id', 'empleado_id', 'usuario', 'rol', 'estado', 'creada_en')),
        table_model(book, 'Cuentas_Pacientes', ('cuenta_id', 'paciente_id', 'tutor_id', 'usuario', 'rol', 'estado', 'creada_en'))
    ) for option in captura.choices(book, model, 'usuario')}
    username = f'{base}.{secrets.token_hex(4)}'
    while username in taken:
        username = f'{base}.{secrets.token_hex(4)}'

    password = secrets.token_urlsafe(20)
    salt = secrets.token_bytes(16)
    digest = scrypt(password.encode(), salt=salt, n=2**17, r=8, p=1, maxmem=256*1024**2)
    append_row(book, account_model, {'cuenta_id': account_id, **identity, 'usuario': username,
                                   'rol': role, 'estado': 'Pendiente', 'creada_en': now})
    append_row(book, credential_model, dict(zip(CREDENTIALS,
        (account_id, username, 'scrypt-131072-8-1', salt.hex(), digest.hex()))))
    captura.write_workbook(book)
    return {'usuario': username, 'contrasena': password, 'persona_id': person, 'rol': role}


class Handler(captura.Handler):
    def do_GET(self):
        if not self.local_host():
            return self.json(403, {'error': 'Usa la dirección local indicada en la terminal.'})
        route = urlsplit(self.path).path
                # Archivos de la página del hospital.
        # ROOT es la carpeta "sitio"; parent es la carpeta principal.
        archivos = {
            '/hospital/': captura.ROOT.parent / 'hospital.html',
            '/hospital/estilos.css': captura.ROOT.parent / 'estilos.css',
            '/hospital/Hospital.jpeg': captura.ROOT.parent / 'Hospital.jpeg',
            '/hospital/LeucoData.jpeg': captura.ROOT.parent / 'LeucoData.jpeg',

            # Imágenes del formulario de registro.
            '/registro/hospital.png': captura.ROOT / 'hospital.png',
            '/registro/leucodatos.png': captura.ROOT / 'leucodatos.png',
        }

        if route in archivos:
            archivo = archivos[route]

            if not archivo.is_file():
                return self.json(404, {
                    'error': 'No se encontró el archivo: ' + archivo.name
                })

            tipo = guess_type(archivo.name)[0] or 'application/octet-stream'
            return self.reply(200, archivo.read_bytes(), tipo)
        if route == '/registro/':
            return self.reply(200, (captura.ROOT/'registro.html').read_bytes(), 'text/html; charset=utf-8')
        if route == '/api/registro':
            try:
                book = load()
                model = captura.schema(book)['tbl_Areas']
                options = captura.choices(book, model, 'area_id')
                return self.json(200, {'areas': options})
            except (OSError, KeyError, ValueError):
                return self.json(409, {'error': 'Revisa el archivo Excel de la carpeta.'})
        if route in ('/registro/sitio/hospital.png', '/registro/sitio/leucodatos.png'):
            route = route.removeprefix('/registro')
        if route.startswith('/sitio/'):
            folder = (captura.ROOT/'sitio').resolve()
            path = (folder/(route.removeprefix('/sitio/') or 'index.html')).resolve()
            if path.is_relative_to(folder) and path.is_file() and path.suffix.lower() in {
                '.html', '.css', '.js', '.png', '.jpg', '.jpeg', '.svg', '.webp'
            }:
                return self.reply(200, path.read_bytes(), guess_type(path.name)[0] or 'application/octet-stream')
            return self.json(404, {'error': 'Archivo del sitio no encontrado.'})
        return super().do_GET()

    def do_POST(self):
        if self.path != '/api/registro':
            return super().do_POST()
        if not self.local_host() or self.headers.get('Origin') != captura.URL:
            return self.json(403, {'error': 'Abre el formulario desde su dirección local.'})
        if self.headers.get_content_type() != 'application/json':
            return self.json(415, {'error': 'Formato de solicitud incorrecto.'})
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= 25000:
                raise ValueError('Tamaño de solicitud incorrecto.')
            data = json.loads(self.rfile.read(length))
            if isinstance(data, dict) and data.get('tipo') == 'paciente':
                try:
                    captura.date.fromisoformat(data.get('fecha_nacimiento', ''))
                except (ValueError, TypeError):
                    raise ValueError('Captura una fecha de nacimiento válida.') from None
            self.json(201, register(load(), data))
        except ValueError as error:
            self.json(400, {'error': str(error)})
        except (OSError, KeyError):
            self.json(409, {'error': 'Revisa que el Excel incluido esté en la carpeta y ciérralo antes de guardar.'})


if __name__ == '__main__':
    with HTTPServer(('127.0.0.1', 8765), Handler) as server:
        url = captura.URL + '/registro/'
        print(f'Registro de cuentas en {url}. Presiona Ctrl+C para cerrar.', flush=True)
        webbrowser.open(url)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
