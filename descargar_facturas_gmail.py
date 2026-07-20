# -*- coding: utf-8 -*-
"""
Revisa el Gmail de tinogas@gmail.com y descarga los adjuntos PDF/XML de
correos que puedan contener facturas (CFDI), desde FECHA_INICIO a la fecha.

Los archivos se guardan en facturas/ para que organizar_facturas.py los
procese después.

Requiere client_secret.json (ya presente en el proyecto) con el scope
gmail.readonly. En el primer uso se abre el navegador para iniciar sesión
con tinogas@gmail.com y se guarda el token en token.json.
"""

import os
import sys
import re
import json
import base64
import time
from datetime import datetime, timezone
from email.header import decode_header

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────
# Configuración
# ──────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CUENTA_GMAIL     = 'tinogas@gmail.com'
CLIENT_SECRET    = os.path.join(BASE_DIR, 'client_secret.json')
TOKEN_PATH       = os.path.join(BASE_DIR, 'token.json')
FACTURAS_DIR     = os.path.join(BASE_DIR, 'facturas')
REVISION_DIR     = os.path.join(BASE_DIR, 'revision')
REGISTRO_JSON    = os.path.join(BASE_DIR, 'gmail_facturas_registro.json')
REPORTE_XLSX     = os.path.join(BASE_DIR, 'gmail_facturas_reporte.xlsx')

SCOPES       = ['https://www.googleapis.com/auth/gmail.readonly']
FECHA_INICIO = '2025/09/01'   # septiembre 2025

# Búsqueda amplia: cualquier correo (bandeja, archivados, etiquetas) con
# adjunto PDF o XML desde FECHA_INICIO. No se restringe por palabra clave
# porque muchas facturas no traen "factura" en el asunto; el filtrado real
# de qué es CFDI válido lo hace organizar_facturas.py al leer cada archivo.
QUERY = f'after:{FECHA_INICIO} has:attachment (filename:pdf OR filename:xml)'

EXTENSIONES_VALIDAS = ('.pdf', '.xml')

COLOR_HDR = '1F4E79'
COLOR_ALT = 'EBF5FB'


# ──────────────────────────────────────────────
# Autenticación
# ──────────────────────────────────────────────
def obtener_servicio_gmail():
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CLIENT_SECRET):
                print(f'ERROR: no se encontró {CLIENT_SECRET}')
                sys.exit(1)
            print(f'Se abrirá el navegador: inicia sesión con {CUENTA_GMAIL} '
                  f'y autoriza el acceso de solo lectura a Gmail.')
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, 'w', encoding='utf-8') as f:
            f.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)


def _con_reintentos(func, *args, **kwargs):
    """Reintenta ante errores transitorios de la API (429/5xx)."""
    espera = 1
    for intento in range(5):
        try:
            return func(*args, **kwargs)
        except HttpError as e:
            if e.resp.status in (429, 500, 502, 503) and intento < 4:
                time.sleep(espera)
                espera *= 2
                continue
            raise


# ──────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────
def _decodificar_header(valor):
    if not valor:
        return ''
    partes = decode_header(valor)
    resultado = []
    for texto, cod in partes:
        if isinstance(texto, bytes):
            resultado.append(texto.decode(cod or 'utf-8', errors='replace'))
        else:
            resultado.append(texto)
    return ' '.join(resultado).strip()


def _limpiar_nombre_archivo(nombre):
    for ch in r'\/:*?"<>|':
        nombre = nombre.replace(ch, ' ')
    return ' '.join(nombre.split()).strip()


def _guardar_sin_sobrescribir(contenido, nombre, carpeta):
    """Guarda contenido en carpeta/nombre; si ya existe, agrega sufijo numérico."""
    os.makedirs(carpeta, exist_ok=True)
    base, ext = os.path.splitext(nombre)
    destino = os.path.join(carpeta, nombre)
    contador = 1
    while os.path.exists(destino):
        destino = os.path.join(carpeta, f'{base}_{contador}{ext}')
        contador += 1
    with open(destino, 'wb') as f:
        f.write(contenido)
    return destino


def _extraer_adjuntos(parts, encontrados):
    """Recorre recursivamente las partes MIME buscando adjuntos PDF/XML."""
    for part in parts or []:
        filename = part.get('filename') or ''
        if filename.lower().endswith(EXTENSIONES_VALIDAS):
            body = part.get('body', {})
            if body.get('attachmentId') or body.get('data'):
                encontrados.append({
                    'filename': _limpiar_nombre_archivo(filename),
                    'attachmentId': body.get('attachmentId'),
                    'data': body.get('data'),
                })
        if part.get('parts'):
            _extraer_adjuntos(part['parts'], encontrados)


def _decodificar_b64(data):
    data = data.replace('-', '+').replace('_', '/')
    falta = len(data) % 4
    if falta:
        data += '=' * (4 - falta)
    return base64.b64decode(data)


# ──────────────────────────────────────────────
# Registro incremental (evita reprocesar correos ya vistos)
# ──────────────────────────────────────────────
def cargar_registro():
    if os.path.exists(REGISTRO_JSON):
        with open(REGISTRO_JSON, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def guardar_registro(registro):
    with open(REGISTRO_JSON, 'w', encoding='utf-8') as f:
        json.dump(registro, f, ensure_ascii=False, indent=2)


# ──────────────────────────────────────────────
# Reporte XLSX
# ──────────────────────────────────────────────
COLS = [
    ('Fecha correo',  16),
    ('Remitente',     40),
    ('Asunto',        50),
    ('Archivo',       45),
    ('Tipo',           6),
    ('Carpeta',       12),
    ('Enlace Gmail',  20),
]

COLOR_FACTURA  = 'D5F5E3'   # verde: pdf + xml juntos
COLOR_REVISION = 'FEF9E7'   # amarillo: falta pdf o xml


def generar_reporte(filas_nuevas):
    hdr_font  = Font(bold=True, color='FFFFFF', size=10)
    hdr_fill  = PatternFill('solid', fgColor=COLOR_HDR)
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin      = Side(style='thin', color='CCCCCC')
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    if os.path.exists(REPORTE_XLSX):
        wb = openpyxl.load_workbook(REPORTE_XLSX)
        ws = wb['Correos'] if 'Correos' in wb.sheetnames else wb.active
        fila_inicio = ws.max_row + 1
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Correos'
        ws.row_dimensions[1].height = 26
        for ci, (nombre, ancho) in enumerate(COLS, 1):
            c = ws.cell(1, ci, nombre)
            c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, hdr_align, border
            ws.column_dimensions[get_column_letter(ci)].width = ancho
        fila_inicio = 2

    for i, fila in enumerate(filas_nuevas):
        ri = fila_inicio + i
        fill = PatternFill('solid', fgColor=COLOR_FACTURA if fila['carpeta'] == 'Factura' else COLOR_REVISION)
        valores = [fila['fecha'], fila['remitente'], fila['asunto'],
                   fila['archivo'], fila['tipo'], fila['carpeta'], 'Abrir correo']
        for ci, val in enumerate(valores, 1):
            c = ws.cell(ri, ci, val)
            c.fill, c.border = fill, border
            c.alignment = Alignment(vertical='center', wrap_text=False)
        link_cell = ws.cell(ri, 7)
        link_cell.hyperlink = fila['enlace']
        link_cell.font = Font(color='1155CC', underline='single')

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(COLS))}1'
    wb.save(REPORTE_XLSX)


# ──────────────────────────────────────────────
# Programa principal
# ──────────────────────────────────────────────
def main():
    print('=' * 60)
    print('  Descarga de Facturas desde Gmail')
    print(f'  Cuenta: {CUENTA_GMAIL}')
    print(f'  Desde:  {FECHA_INICIO}')
    print('=' * 60)

    servicio = obtener_servicio_gmail()
    registro = cargar_registro()

    # ── Listar mensajes que cumplen el filtro ──────────────────────
    print(f'\nBuscando correos: {QUERY}')
    mensajes = []
    page_token = None
    while True:
        resp = _con_reintentos(
            servicio.users().messages().list(
                userId='me', q=QUERY, maxResults=500, pageToken=page_token
            ).execute
        )
        mensajes.extend(resp.get('messages', []))
        page_token = resp.get('nextPageToken')
        if not page_token:
            break
    print(f'  Correos encontrados: {len(mensajes)}')

    nuevos = [m for m in mensajes if m['id'] not in registro]
    ya_procesados = len(mensajes) - len(nuevos)
    print(f'  Ya procesados en corridas anteriores: {ya_procesados}')
    print(f'  Nuevos por revisar: {len(nuevos)}')

    if not nuevos:
        print('\nNada nuevo que descargar.')
        return

    filas_reporte = []
    total_pdf = total_xml = 0

    for i, msg_ref in enumerate(nuevos, 1):
        msg_id = msg_ref['id']
        pct = i / len(nuevos)
        lleno = int(30 * pct)
        barra = '=' * lleno + '.' * (30 - lleno)
        sys.stdout.write(f'\r[{barra}] {pct*100:5.1f}%  ({i}/{len(nuevos)})')
        sys.stdout.flush()

        msg = _con_reintentos(
            servicio.users().messages().get(userId='me', id=msg_id, format='full').execute
        )

        headers = {h['name']: h['value'] for h in msg['payload'].get('headers', [])}
        remitente = _decodificar_header(headers.get('From', ''))
        asunto    = _decodificar_header(headers.get('Subject', '(sin asunto)'))

        ts_ms = int(msg.get('internalDate', 0))
        fecha = datetime.fromtimestamp(ts_ms / 1000, tz=timezone.utc).astimezone()
        fecha_str = fecha.strftime('%Y-%m-%d %H:%M')

        payload = msg['payload']
        partes = payload.get('parts') or [payload]
        adjuntos = []
        _extraer_adjuntos(partes, adjuntos)

        # Un correo solo cuenta como "Factura" si trae PDF y XML juntos;
        # si falta alguno de los dos, todos sus adjuntos van a revisión
        # manual en vez de a facturas/.
        tiene_pdf = any(a['filename'].lower().endswith('.pdf') for a in adjuntos)
        tiene_xml = any(a['filename'].lower().endswith('.xml') for a in adjuntos)
        es_factura = tiene_pdf and tiene_xml
        carpeta_dest = FACTURAS_DIR if es_factura else REVISION_DIR
        etiqueta_carpeta = 'Factura' if es_factura else 'Revisión'

        archivos_guardados = []
        for adj in adjuntos:
            if adj['data']:
                contenido = _decodificar_b64(adj['data'])
            else:
                att = _con_reintentos(
                    servicio.users().messages().attachments().get(
                        userId='me', messageId=msg_id, id=adj['attachmentId']
                    ).execute
                )
                contenido = _decodificar_b64(att['data'])

            destino = _guardar_sin_sobrescribir(contenido, adj['filename'], carpeta_dest)
            archivos_guardados.append(os.path.basename(destino))

            tipo = 'PDF' if destino.lower().endswith('.pdf') else 'XML'
            if tipo == 'PDF':
                total_pdf += 1
            else:
                total_xml += 1

            filas_reporte.append({
                'fecha': fecha_str,
                'remitente': remitente,
                'asunto': asunto,
                'archivo': os.path.basename(destino),
                'tipo': tipo,
                'carpeta': etiqueta_carpeta,
                'enlace': f'https://mail.google.com/mail/u/0/#all/{msg_id}',
            })

        registro[msg_id] = {
            'fecha_procesado': datetime.now().isoformat(timespec='seconds'),
            'fecha_correo': fecha_str,
            'remitente': remitente,
            'asunto': asunto,
            'carpeta': etiqueta_carpeta,
            'archivos': archivos_guardados,
        }

        # Guardar registro cada 25 correos por si el proceso se interrumpe
        if i % 25 == 0:
            guardar_registro(registro)

    print()
    guardar_registro(registro)

    if filas_reporte:
        print(f'\nGenerando reporte ({len(filas_reporte)} archivos nuevos)...')
        generar_reporte(filas_reporte)
        print(f'  Reporte: {REPORTE_XLSX}')
    else:
        print('\nNo se encontraron adjuntos PDF/XML en los correos nuevos.')

    correos_factura   = sum(1 for r in registro.values() if r.get('carpeta') == 'Factura' and r['archivos'])
    correos_revision  = sum(1 for r in registro.values() if r.get('carpeta') == 'Revisión' and r['archivos'])

    print(f'\nArchivos descargados: {total_pdf} PDF, {total_xml} XML')
    print(f'  Correos con PDF+XML completos -> {FACTURAS_DIR}  ({correos_factura} correos)')
    print(f'  Correos incompletos (falta PDF o XML) -> {REVISION_DIR}  ({correos_revision} correos)')
    print('\nSiguiente paso: ejecuta organizar_facturas.py para clasificar lo que quedó en facturas/.')
    print('Revisa manualmente la carpeta revision/ para decidir qué hacer con esos correos.')
    print('=' * 60)


if __name__ == '__main__':
    main()
