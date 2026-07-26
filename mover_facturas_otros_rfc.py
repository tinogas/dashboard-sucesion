# -*- coding: utf-8 -*-
"""
Aparta en revision/otros_rfc/ las facturas ya descargadas que no son tuyas.

Una factura es tuya si tu RFC aparece como emisor o como receptor: las de
renta que le emites a un inquilino cuentan igual que las que te emiten a ti.
Solo se mueven los CFDI donde tu RFC no participa (llegaron por correo pero
son de terceros). El PDF que acompaña al XML se mueve junto con él.

    python mover_facturas_otros_rfc.py                 # simulación (no mueve nada)
    python mover_facturas_otros_rfc.py --aplicar       # mueve de verdad
    python mover_facturas_otros_rfc.py --deshacer      # regresa lo movido

Cada corrida con --aplicar deja registro en movidos_otros_rfc.json, que es lo
que usa --deshacer para devolver cada archivo a su carpeta original.
"""

import os
import sys
import json
import shutil
import argparse
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Dónde buscar: lo pendiente de organizar y lo ya organizado. revision/ queda
# fuera a propósito: eso ya está apartado para revisarse a mano.
CARPETAS_ORIGEN = [
    os.path.join(BASE_DIR, 'facturas'),
    os.path.join(BASE_DIR, 'facturas_organizadas'),
]
DESTINO   = os.path.join(BASE_DIR, 'revision', 'otros_rfc')
BITACORA  = os.path.join(BASE_DIR, 'movidos_otros_rfc.json')

RFC_PROPIO_DEFAULT = 'GALF730909CN0'


def rfcs_de_xml(ruta):
    """(emisor, receptor) del CFDI. (None, None) si el XML no se puede leer."""
    try:
        with open(ruta, 'rb') as f:
            contenido = f.read().lstrip(b'\xef\xbb\xbf')
        root = ET.fromstring(contenido)
    except (OSError, ET.ParseError):
        return None, None

    emisor = receptor = None
    for el in root.iter():
        tag = el.tag.rsplit('}', 1)[-1]
        if tag == 'Emisor' and emisor is None:
            emisor = (el.get('Rfc') or el.get('rfc') or '').strip().upper() or None
        elif tag == 'Receptor' and receptor is None:
            receptor = (el.get('Rfc') or el.get('rfc') or '').strip().upper() or None
        if emisor and receptor:
            break
    return emisor, receptor


def acompanantes(ruta_xml):
    """Archivos que viajan con el XML: los del mismo nombre base (el PDF).

    Se listan las entradas reales de la carpeta en vez de probar extensiones:
    en Windows .pdf y .PDF son el mismo archivo y se contaría dos veces.
    """
    carpeta = os.path.dirname(ruta_xml)
    base = os.path.splitext(os.path.basename(ruta_xml))[0].lower()
    extras = []
    for nombre in os.listdir(carpeta):
        raiz, ext = os.path.splitext(nombre)
        if raiz.lower() == base and ext.lower() != '.xml':
            extras.append(os.path.join(carpeta, nombre))
    return sorted(extras)


def buscar_ajenas(rfcs_propios):
    """[(xml, emisor, receptor, [acompañantes])] de los CFDI de terceros."""
    ajenas = []
    vistos = set()
    for carpeta in CARPETAS_ORIGEN:
        if not os.path.isdir(carpeta):
            continue
        for raiz, _, archivos in os.walk(carpeta):
            if os.path.normcase(raiz).startswith(os.path.normcase(DESTINO)):
                continue
            for nombre in archivos:
                if not nombre.lower().endswith('.xml'):
                    continue
                ruta = os.path.join(raiz, nombre)
                clave = os.path.normcase(os.path.abspath(ruta))
                if clave in vistos:      # facturas/ y Facturas/ son la misma
                    continue             # carpeta en Windows
                vistos.add(clave)

                emisor, receptor = rfcs_de_xml(ruta)
                if not emisor and not receptor:
                    continue             # no es CFDI legible: no se toca
                if {emisor, receptor} & rfcs_propios:
                    continue             # participas en ella: se queda
                ajenas.append((ruta, emisor, receptor, acompanantes(ruta)))
    return ajenas


def destino_sin_sobrescribir(nombre):
    base, ext = os.path.splitext(nombre)
    ruta = os.path.join(DESTINO, nombre)
    n = 1
    while os.path.exists(ruta):
        ruta = os.path.join(DESTINO, f'{base}_{n}{ext}')
        n += 1
    return ruta


def limpiar_carpetas_vacias(carpetas):
    """Borra las carpetas de emisor que quedaron sin archivos tras mover."""
    borradas = []
    for carpeta in sorted(set(carpetas), key=len, reverse=True):
        dentro_de_origen = any(
            os.path.normcase(carpeta).startswith(os.path.normcase(raiz))
            for raiz in CARPETAS_ORIGEN)
        if dentro_de_origen and os.path.isdir(carpeta) and not os.listdir(carpeta):
            os.rmdir(carpeta)
            borradas.append(carpeta)
    return borradas


def mover(ajenas):
    """Mueve los archivos y devuelve la bitácora de lo movido."""
    os.makedirs(DESTINO, exist_ok=True)
    movimientos = []
    for ruta_xml, emisor, receptor, extras in ajenas:
        for origen in [ruta_xml] + extras:
            destino = destino_sin_sobrescribir(os.path.basename(origen))
            shutil.move(origen, destino)
            movimientos.append({
                'origen': os.path.relpath(origen, BASE_DIR),
                'destino': os.path.relpath(destino, BASE_DIR),
                'emisor': emisor,
                'receptor': receptor,
            })

    for carpeta in limpiar_carpetas_vacias([os.path.dirname(a[0]) for a in ajenas]):
        print(f'  carpeta vacía eliminada: {os.path.relpath(carpeta, BASE_DIR)}')
    return movimientos


def guardar_bitacora(movimientos):
    historial = []
    if os.path.exists(BITACORA):
        with open(BITACORA, encoding='utf-8') as f:
            historial = json.load(f)
    historial.extend(movimientos)
    with open(BITACORA, 'w', encoding='utf-8') as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)


def deshacer():
    if not os.path.exists(BITACORA):
        print(f'No hay nada que deshacer: no existe {BITACORA}')
        return 0

    with open(BITACORA, encoding='utf-8') as f:
        historial = json.load(f)

    regresados, fallidos = 0, []
    for mov in reversed(historial):
        origen  = os.path.join(BASE_DIR, mov['origen'])
        destino = os.path.join(BASE_DIR, mov['destino'])
        if not os.path.exists(destino):
            fallidos.append(f"{mov['destino']} (ya no está en otros_rfc/)")
            continue
        if os.path.exists(origen):
            fallidos.append(f"{mov['origen']} (ya existe en su carpeta original)")
            continue
        os.makedirs(os.path.dirname(origen), exist_ok=True)
        shutil.move(destino, origen)
        regresados += 1

    print(f'Archivos regresados a su carpeta original: {regresados}')
    for f_ in fallidos:
        print(f'  omitido: {f_}')
    if regresados and not fallidos:
        os.remove(BITACORA)
        print(f'Bitácora eliminada: {BITACORA}')
    return 0


def main():
    p = argparse.ArgumentParser(
        description='Aparta en revision/otros_rfc/ las facturas de terceros.')
    p.add_argument('--rfc', default=RFC_PROPIO_DEFAULT,
                   help='RFC propio; varios separados por coma. '
                        f'Por defecto {RFC_PROPIO_DEFAULT}.')
    p.add_argument('--aplicar', action='store_true',
                   help='Mueve los archivos. Sin este flag solo simula.')
    p.add_argument('--deshacer', action='store_true',
                   help='Regresa a su carpeta original lo movido antes.')
    args = p.parse_args()

    if args.deshacer:
        return deshacer()

    rfcs_propios = {r.strip().upper() for r in args.rfc.split(',') if r.strip()}
    if not rfcs_propios:
        print('ERROR: hay que indicar al menos un RFC propio.')
        return 1

    print('=' * 70)
    print('  Facturas de terceros -> revision/otros_rfc/')
    print(f'  RFC propio: {", ".join(sorted(rfcs_propios))} (como emisor o receptor)')
    print(f'  Modo: {"APLICAR (se mueven los archivos)" if args.aplicar else "SIMULACIÓN (no se mueve nada)"}')
    print('=' * 70)

    ajenas = buscar_ajenas(rfcs_propios)
    if not ajenas:
        print('\nNo hay facturas de terceros: todo lo descargado te involucra.')
        return 0

    total_archivos = sum(1 + len(extras) for _, _, _, extras in ajenas)
    print(f'\nCFDI de terceros: {len(ajenas)}  ({total_archivos} archivos con sus PDF)\n')
    for ruta, emisor, receptor, extras in ajenas:
        print(f'  {emisor or "?"} -> {receptor or "?"}')
        print(f'      {os.path.relpath(ruta, BASE_DIR)}')
        for extra in extras:
            print(f'      {os.path.relpath(extra, BASE_DIR)}')

    if not args.aplicar:
        print('\nSimulación: no se movió nada.')
        print('Vuelve a ejecutar con --aplicar para moverlos.')
        return 0

    movimientos = mover(ajenas)
    guardar_bitacora(movimientos)
    print(f'\nMovidos {len(movimientos)} archivos a {DESTINO}')
    print(f'Bitácora: {BITACORA}  (--deshacer los regresa a su lugar)')
    print('\nNota: reorganiza con organizar_facturas.py para que el reporte '
          'de facturas refleje el cambio.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
