# -*- coding: utf-8 -*-
"""
Detalle de productos de facturas que llegaron SOLO en PDF (sin XML), para que
el acordeón de Mantenimiento en dashboard.py pueda mostrar su consumo igual
que las facturas con XML.

Cómo agregar una factura nueva:
  1. Abre el PDF y copia Folio Fiscal, fecha de emisión, Subtotal, IVA y Total.
  2. Copia cada concepto con su descripción e importe (sin IVA).
  3. Agrega un dict a FACTURAS_MANUALES con esos datos.

fecha: 'YYYY-MM-DD' (fecha de emisión del CFDI).
total: Subtotal + IVA - Descuento (debe coincidir con el "Total" del PDF).
"""

FACTURAS_MANUALES = [
    # ── FERRETERIA PRIETO (emisor: Jaime Uriel Rodriguez Ibarra) ──────────
    {
        'uuid':       '0375D17C-1F67-43CC-8716-5F7B3D3E3466',
        'emisor':     'Jaime Uriel Rodriguez Ibarra (Ferreteria Prieto)',
        'fecha':      '2025-10-17',
        'total':      1392.00,
        'iva':        192.00,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Trabajo de limpieza', 'importe': 1200.00},
        ],
    },
    {
        'uuid':       '09E16FB7-8215-499E-89B4-1ADEFD25C3AB',
        'emisor':     'Jaime Uriel Rodriguez Ibarra (Ferreteria Prieto)',
        'fecha':      '2025-10-17',
        'total':      2981.20,
        'iva':        411.20,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Material de ferreteria en general', 'importe': 2570.00},
        ],
    },
    {
        'uuid':       '588A5527-EDB7-46BC-BCA6-DEF00BDF5A4E',
        'emisor':     'Jaime Uriel Rodriguez Ibarra (Ferreteria Prieto)',
        'fecha':      '2025-10-24',
        'total':      527.80,
        'iva':        72.80,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Contacto duplex 2 polos',                      'importe': 46.00},
            {'descripcion': 'Placa armada con contacto y 2 interruptores',  'importe': 134.00},
            {'descripcion': 'Placa ciega de aluminio',                      'importe': 25.00},
            {'descripcion': 'Instalacion de accesorios electricos',         'importe': 250.00},
        ],
    },
    {
        'uuid':       'AFE33465-15E7-42D7-B223-EA69165A8164',
        'emisor':     'Jaime Uriel Rodriguez Ibarra (Ferreteria Prieto)',
        'fecha':      '2025-10-20',
        'total':      1635.60,
        'iva':        225.60,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Instalacion de chapa en puerta de herreria',     'importe': 800.00},
            {'descripcion': 'Cerradura de sobreponer tradicional negro',      'importe': 215.00},
            {'descripcion': 'Juego de accesorio de descarga para WC',         'importe': 145.00},
            {'descripcion': 'Instalacion de accesorio para WC',               'importe': 250.00},
        ],
    },
    {
        'uuid':       'C8BAA384-D7C6-463B-A2EF-28831C7BA95E',
        'emisor':     'Jaime Uriel Rodriguez Ibarra (Ferreteria Prieto)',
        'fecha':      '2025-10-17',
        'total':      2088.00,
        'iva':        288.00,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Instalacion de accesorios en lavamanos, lavatrastes y WC', 'importe': 1800.00},
        ],
    },

    # ── SERVICLEAN / Marco Antonio Reyes Diaz — factura de octubre 2025
    # ($7,076.00) pagada en 4 abonos repartidos entre 3 casas distintas.
    # El emparejamiento normal busca UN movimiento cuyo monto sea ~igual al
    # total de la factura, pero aqui cada "Abono" en Movimientos es solo una
    # fraccion del total, asi que se agrega una entrada manual por abono con
    # su propia fecha/monto y su parte proporcional del concepto e IVA
    # (aunque la factura SI tiene XML — vease
    # facturas_organizadas/2025/10_Octubre/MARCO ANTONIO REYES DIAZ/).
    {
        'uuid':       'D5436727-E4D1-4530-80D5-2648317C445A',
        'emisor':     'Marco Antonio Reyes Diaz (Serviclean) - Abono 2, Toledo 2101-A',
        'fecha':      '2025-10-22',
        'total':      600.00,
        'iva':        82.76,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Servicio de limpieza en general (prorrateado, $600.00 de $7,076.00)', 'importe': 517.24},
        ],
    },
    {
        'uuid':       'D5436727-E4D1-4530-80D5-2648317C445A',
        'emisor':     'Marco Antonio Reyes Diaz (Serviclean) - Abono 1, Toledo 2103',
        'fecha':      '2025-10-22',
        'total':      500.00,
        'iva':        68.97,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Servicio de limpieza en general (prorrateado, $500.00 de $7,076.00)', 'importe': 431.03},
        ],
    },
    {
        'uuid':       'D5436727-E4D1-4530-80D5-2648317C445A',
        'emisor':     'Marco Antonio Reyes Diaz (Serviclean) - Abono 3, Toledo 2103',
        'fecha':      '2025-10-23',
        'total':      4000.00,
        'iva':        551.72,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Servicio de limpieza en general (prorrateado, $4,000.00 de $7,076.00)', 'importe': 3448.28},
        ],
    },
    {
        'uuid':       'D5436727-E4D1-4530-80D5-2648317C445A',
        'emisor':     'Marco Antonio Reyes Diaz (Serviclean) - Abono 4, Toledo 2103-B',
        'fecha':      '2025-10-28',
        'total':      1976.00,
        'iva':        272.55,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Servicio de limpieza en general (prorrateado, $1,976.00 de $7,076.00)', 'importe': 1703.45},
        ],
    },

    # ── SERVICLEAN / Marco Antonio Reyes Diaz — factura de enero 2026
    # ($2,900.00). El movimiento junta 2 casas en un solo renglon ("Pesqueira
    # Altos y Toledo 2101-B", ver INMUEBLES_COMBINADOS en dashboard.py), que
    # se reparte 50/50. Aqui va la mitad del concepto e IVA para que cada
    # casa (ya dividida por _expandir_inmuebles_combinados) encuentre su
    # propia parte proporcional.
    {
        'uuid':       'BF10E45D-03A5-4463-93BC-092D720FBA16',
        'emisor':     'Marco Antonio Reyes Diaz (Serviclean) - mitad (50%)',
        'fecha':      '2026-01-14',
        'total':      1450.00,
        'iva':        200.00,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Servicio de limpieza en general (prorrateado 50%, mitad de $2,900.00)', 'importe': 1250.00},
        ],
    },

    # ── MADERAS Y TABLEROS DE NAVOJOA (emisor: Humberta Estrada Duarte) ───
    {
        'uuid':       '7A5BEF8F-49DF-4699-B45E-9FED672060EE',
        'emisor':     'Maderas y Tableros de Navojoa (Humberta Estrada Duarte)',
        'fecha':      '2025-10-27',
        'total':      2290.01,
        'iva':        315.86,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Puertas de tambor (2 pza)',                                       'importe': 1637.94},
            {'descripcion': 'Bisagras Dexter 3"x3" 2.2mm c/baleros c/plana NS c/T (4 pza)',      'importe': 336.21},
        ],
    },
    {
        'uuid':       'C4C29C43-4E20-45F8-BBCD-846287F73DD6',
        'emisor':     'Maderas y Tableros de Navojoa (Humberta Estrada Duarte)',
        'fecha':      '2025-12-04',
        'total':      200.00,
        'iva':        27.59,
        'descuento':  0,
        'conceptos': [
            {'descripcion': 'Chapa manija con llave humo Bitka (1 pza)', 'importe': 172.41},
        ],
    },
]
