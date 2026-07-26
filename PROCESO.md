# Dashboard Sucesión — Proceso y Orden de Ejecución

## Descripción general

Sistema de administración contable para una sucesión de inmuebles. Lee los movimientos registrados en Google Sheets y genera reportes en Excel, PDFs de recibos de renta y un reporte ejecutivo en Word.

---

## Prerrequisitos

- Python 3.x con entorno virtual `.venv` activo
- `credentials.json` — Service Account de Google Cloud con acceso a Google Sheets y Drive
- `firma.png` — Imagen de firma digital (necesaria para los recibos PDF)
- Carpeta `facturas/` con XMLs y/o PDFs de CFDIs (solo para el paso 2)
- Dependencias instaladas:
  ```
  pip install -r requirements.txt
  ```

---

## Orden de ejecución

### Paso 0 — `descargar_facturas_gmail.py` *(opcional)*

Revisa el Gmail de la cuenta configurada y descarga a `facturas/` los adjuntos PDF/XML de correos con posibles CFDI, dentro de un rango de fechas.

| | |
|---|---|
| **Requiere** | `client_secret.json` (credencial OAuth de escritorio de Google Cloud) |
| **Salida** | Archivos en `facturas/` |
| **Salida** | `gmail_facturas_reporte.xlsx` (bitácora de correos revisados) |
| **Salida** | `gmail_facturas_registro.json` (historial interno, evita reprocesar correos) |

- El rango de fechas (`--desde`/`--hasta`, inclusive) se elige en `dashboard_gui.py` o por línea de comandos; sin `--hasta` no hay límite superior.
- Solo se guardan en `facturas/` los CFDI donde el RFC propio (`--rfc`, varios separados por coma) aparece como **emisor o receptor** — las facturas de renta que se emiten a los inquilinos cuentan igual que las que se reciben. Con `--rfc ""` se descarga todo sin filtrar.
- Los CFDI de terceros (el RFC propio no participa) se apartan en `revision/otros_rfc/` sin borrarse.
- Correos incompletos (falta PDF o XML) van a `revision/`.

```
python descargar_facturas_gmail.py --desde 2026-07-01 --hasta 2026-07-24 --rfc GALF730909CN0
```

---

### Paso 1 — `dashboard.py`

Descarga los movimientos de Google Sheets y genera el archivo Excel principal con todas las hojas analíticas.

| | |
|---|---|
| **Fuente** | Google Sheets — hoja `movimientos` |
| **Salida** | `dashboard_sucesion.xlsx` |

**Hojas generadas en el Excel:**

| Hoja | Contenido |
|---|---|
| Movimientos | Copia formateada de los datos originales |
| Dashboard | KPIs y gráficos de ingresos vs. egresos |
| Informe Contable | Resumen agrupado por Registro (expandible) |
| Informe Dinámico | Jerarquía completa: Año › Mes › Registro › Inmueble › Concepto |
| Detalle Registro | Listado detallado con subtotales por Registro |
| Control Rentas | Pivote Inquilino × Mes (ingresos) |
| Control Luz | Pivote Inmueble × Mes (egresos) |
| Control Predial | Pivote Inmueble × Mes (egresos) |
| Control Agua | Pivote Inmueble × Mes (egresos) |
| Control Despacho | Pivote Inquilino × Mes + resumen por concepto |
| Control Impuestos | Pivote Inmueble × Mes (egresos) |
| Detalle por Propiedad | Una hoja por cada inmueble en renta, más los listados en `INMUEBLES_SIN_RENTA` (inmuebles sin ingresos de renta pero con gastos propios, ej. luz/predial de una casa desocupada) |

Al terminar, abre el archivo automáticamente en Excel.

> **Nota:** Cerrar `dashboard_sucesion.xlsx` en Excel antes de volver a ejecutar este script.

---

### Paso 2 — `organizar_facturas.py` *(opcional)*

Ejecutar solo cuando haya facturas CFDI nuevas en la carpeta `facturas/`.

Lee XMLs y PDFs de CFDI, los organiza en carpetas por año/mes/emisor y genera un reporte de cruce con los movimientos.

| | |
|---|---|
| **Requiere** | Carpeta `facturas/` con XMLs/PDFs + `dashboard_sucesion.xlsx` |
| **Salida** | Carpetas `facturas_organizadas/{año}/{MM_Mes}/{Emisor}/` |
| **Salida** | `reporte_facturas.xlsx` con resultado del cruce |

El cruce entre facturas y movimientos usa una tolerancia de ±5 días en fecha y ±2% en monto. El reporte Excel colorea cada fila según la calidad del match:

| Color | Significado |
|---|---|
| Verde | Coincide fecha + monto |
| Amarillo | Coincide solo monto |
| Rojo | Sin coincidencia |

---

### Paso 3 — `generar_recibos.py`

Lee los movimientos con REGISTRO = 'RENTA' y genera un par de PDFs por cada recibo nuevo (original + copia) en tamaño media carta.

| | |
|---|---|
| **Requiere** | `dashboard_sucesion.xlsx` + `firma.png` |
| **Salida** | PDFs en `recibos/{año}/{inmueble}/` |
| **Salida** | `recibos_registro.json` (historial y consecutivo) |

- Cada recibo lleva numeración consecutiva incremental.
- Los recibos ya generados no se vuelven a crear (se detectan por clave `inquilino|mes|año|inmueble`).
- Para reiniciar la numeración desde cero, borrar `recibos_registro.json`.

---

### Paso 4 — `reporte_word.py`

Genera el reporte ejecutivo contable completo en formato Word con tablas, gráficos y análisis por período, registro e inmueble.

| | |
|---|---|
| **Requiere** | Google Sheets + `dashboard_sucesion.xlsx` |
| **Salida** | `reporte_ejecutivo_sucesion.docx` |

**Secciones del documento:**
1. Portada con período y fecha de generación
2. Resumen ejecutivo (KPIs y gráfico anual)
3. Análisis de ingresos (Control de Rentas incluido)
4. Análisis de egresos (Luz, Predial, Agua, Despacho, Impuestos)
5. Reporte detallado por año
6. Resumen de registros y situación de rentas

Al terminar, abre el documento automáticamente en Word.

---

### Paso 5 — `reporte_word_propiedades.py` *(opcional)*

Genera un reporte Word independiente con un salto de página por cada inmueble (en renta o en `INMUEBLES_SIN_RENTA`): su bitácora completa por categoría (Renta, Luz, Agua, Predial, Mantenimiento, Impuestos, etc.) y el detalle de productos de cada factura XML relacionada.

| | |
|---|---|
| **Requiere** | Google Sheets + `dashboard_sucesion.xlsx` |
| **Salida** | `reporte_propiedades.docx` |

---

## Flujo de datos

```
Google Sheets (movimientos)
         │
         ▼
   dashboard.py ─────────────► dashboard_sucesion.xlsx
                                        │
                         ┌──────────────┼──────────────┐
                         │              │              │
                         ▼              ▼              ▼
            organizar_facturas.py  generar_recibos.py  reporte_word.py
                         │              │              │
                         ▼              ▼              ▼
              reporte_facturas.xlsx  recibos/      reporte_ejecutivo_
              facturas_organizadas/  *.pdf         sucesion.docx
```

---

## Utilidades

- **`diagnostico_recibos.py`** — Solo imprime en consola. Útil para verificar datos del Excel y el JSON de recibos sin generar ningún archivo.
- **`mover_facturas_otros_rfc.py`** — Aparta en `revision/otros_rfc/` las facturas ya descargadas (en `facturas/` o `facturas_organizadas/`) donde el RFC propio no participa ni como emisor ni como receptor. Útil para limpiar retroactivamente lo descargado antes de que `descargar_facturas_gmail.py` filtrara por RFC. Corre en modo simulación por defecto; usa `--aplicar` para mover de verdad y `--deshacer` para revertir (se apoya en `movidos_otros_rfc.json`).

---

## Notas de operación

- Siempre ejecutar `dashboard.py` primero; los demás scripts dependen del Excel que genera.
- `organizar_facturas.py` es independiente de `generar_recibos.py` y `reporte_word.py`; puede ejecutarse en cualquier momento después del paso 1.
- Si la API de Google Sheets no responde, los scripts intentan descargar el XLSX completo vía Drive API como alternativa automática.
- Las credenciales en `credentials.json` son de tipo Service Account; no requieren login interactivo.
