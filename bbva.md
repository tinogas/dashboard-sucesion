# Conciliación bancaria BBVA vs. hoja "movimientos"

## Qué es esto

Script exploratorio (`conciliacion_bbva.py`) que cruza el estado de cuenta BBVA de
un mes contra lo capturado en la hoja `movimientos` de Google Sheets (la misma que
alimenta `dashboard.py`), para detectar:

- Movimientos bancarios que no quedaron registrados en la hoja.
- Filas registradas en la hoja que no tienen respaldo en el banco ese mes.
- Diferencias de monto entre lo que dice el banco y lo que se capturó.

No se integra al flujo de `dashboard.py` ni a ningún proceso automático — es una
herramienta de auditoría de uso manual.

---

## Por qué el PDF y no el Excel/Word del banco

Se revisaron las tres muestras que BBVA entrega cada mes:

| Formato | Veredicto |
|---|---|
| **Excel** | Se descarta. Viene roto en ~39 hojas ("Table 1"..."Table 39") con celdas desalineadas, texto multilinea mezclado con el encabezado y montos que caen en columnas inconsistentes. |
| **Word** | No se usó — mismo tipo de exportación tabular poco confiable que el Excel. |
| **PDF** | El que se usa. Trae la sección "Detalle de Movimientos Realizados" en texto plano, con una tabla de columnas consistente (FECHA OPER / FECHA LIQ / DESCRIPCIÓN / CARGOS / ABONOS / SALDO OPERACIÓN / SALDO LIQUIDACIÓN) en todas las páginas. |

## Cómo se parsea el PDF

El texto plano del PDF **no distingue por sí solo si un monto es cargo o abono** —
ambos aparecen como números sueltos en la misma línea. La solución fue usar
`pdfplumber.extract_words()` para obtener la posición horizontal (`x0`) de cada
palabra y clasificar los montos según en qué columna cae:

1. En cada página se busca la fila de encabezado (donde aparecen `CARGOS` y
   `ABONOS`) y se toma el `x0` de cada columna: `CARGOS`, `ABONOS`, `SALDO
   OPERACIÓN`, `SALDO LIQUIDACIÓN`.
2. Se calculan los puntos medios entre columnas consecutivas como límites.
3. Cada línea que empieza con dos fechas (`DD/MES DD/MES`) es un movimiento
   nuevo; las líneas siguientes sin fecha son referencias/continuación de la
   descripción del movimiento anterior.
4. Cada monto de la línea se clasifica como cargo o abono según en qué lado de
   los límites cae su `x0`. Los saldos de operación/liquidación se descartan
   (no hacen falta para el cruce).

**Validación automática:** el propio estado de cuenta imprime sus totales
declarados (`TOTAL IMPORTE CARGOS`, `TOTAL IMPORTE ABONOS` y el número de
movimientos de cada tipo). El script los extrae con regex y compara contra la
suma de lo que parseó — si no coinciden, imprime una advertencia. En los 10
meses procesados (septiembre 2025 a junio 2026) la suma parseada coincidió
exactamente con la declarada en todos los casos.

## Cómo se obtiene "lo registrado"

Se reutilizan funciones ya existentes de `dashboard.py` (sin duplicar lógica):

- `download_movimientos()` — descarga la hoja `movimientos` en vivo desde Google
  Sheets (con su mismo mecanismo de respaldo vía Drive API si gspread falla).
  Si la descarga en vivo falla por completo, el script cae a leer la hoja
  `Movimientos` de la copia local `dashboard_sucesion.xlsx`.
- `detect_columns()` / `prepare_año_mes()` — detectan las columnas
  Fecha/Ingreso/Egreso/Concepto/Registro/Inquilino/Inmueble/Año/Mes igual que el
  resto del proyecto.
- `clean_numeric()` — limpia los montos (quita `$`, comas, etc.).

Se filtra la hoja al año/mes pedido.

## Criterio de emparejamiento

Monto exacto (con dos decimales) + ventana amplia de fecha (mismo mes), tal como
se decidió al diseñar el script. Cuando un monto tiene más de una fila candidata
sin usar, se elige la más cercana en fecha al movimiento bancario (greedy, cada
fila de registro se usa una sola vez).

**Bug encontrado y corregido durante las pruebas:** la columna `Egreso` de la
hoja guarda los montos en **negativo** (ej. `-438.00`), mientras que el cargo
bancario se lee en positivo. La primera versión comparaba los signos tal cual y
por eso *ningún* cargo hacía match. Se corrigió comparando por valor absoluto en
ambos lados.

Los movimientos de "TRASPASO CUENTAS PROPIAS" (transferencias entre cuentas del
mismo titular) se excluyen del cruce — no les corresponde tener fila en la hoja
de registro, así que se marcan aparte como `TRASPASO PROPIO` en vez de `SIN
REGISTRAR`.

---

## Uso

```
python conciliacion_bbva.py --año 2025 --mes 9
```

Requiere que exista `Estados de Cuentas BBVA/{año}/{MM} {MesAbbr}.pdf` (ej. `09
Sep.pdf`). Genera `conciliacion_bbva_{año}_{mes en 3 letras}.xlsx` (ej.
`conciliacion_bbva_2025_sep.xlsx`) con tres hojas:

| Hoja | Contenido |
|---|---|
| **Banco** | Cada movimiento del estado de cuenta, con su estado (`CONCILIADO` / `SIN REGISTRAR` / `TRASPASO PROPIO`) y, si concilió, el concepto/registro/inquilino-inmueble con el que hizo match. |
| **Registro sin banco** | Filas de `movimientos` de ese mes que no encontraron respaldo bancario, con columna `Estado = SIN MOVIMIENTO BANCARIO`. |
| **Resumen** | Totales de cargos/abonos (parseado vs. declarado por el banco), conteos de conciliados/sin registrar de ambos lados. |

Todo lo que **no** coincide entre banco y registro queda marcado en rojo en
ambas hojas, para que salte a la vista sin tener que leer el detalle.

---

## Resultado de la corrida (sep 2025 – jun 2026)

Se corrió para los 10 meses que tienen estado de cuenta disponible en
`Estados de Cuentas BBVA/` (no hay archivos posteriores a junio 2026 todavía).
En todos los meses la suma parseada del PDF coincidió con los totales
declarados por el banco.

| Archivo | Conciliados | Sin registrar (banco) | Sin respaldo bancario (registro) |
|---|---|---|---|
| `conciliacion_bbva_2025_sep.xlsx` | 12 | 6 | 8 |
| `conciliacion_bbva_2025_oct.xlsx` | 25 | 4 | 10 |
| `conciliacion_bbva_2025_nov.xlsx` | 13 | 5 | 10 |
| `conciliacion_bbva_2025_dic.xlsx` | 25 | 3 | 6 |
| `conciliacion_bbva_2026_ene.xlsx` | 26 | 3 | 7 |
| `conciliacion_bbva_2026_feb.xlsx` | 15 | 3 | 6 |
| `conciliacion_bbva_2026_mar.xlsx` | 17 | 4 | 3 |
| `conciliacion_bbva_2026_abr.xlsx` | 22 | 3 | 6 |
| `conciliacion_bbva_2026_may.xlsx` | 13 | 2 | 3 |
| `conciliacion_bbva_2026_jun.xlsx` | 18 | 4 | 4 |

### Hallazgos notables (septiembre 2025, revisados como muestra)

- **Pago registrado a la mitad:** un SPEI recibido de "Abel Páez" el 14/SEP por
  $12,361.15 en el banco aparece registrado como $6,180.575 (justo la mitad) el
  mismo día — posible convención de reparto que vale la pena confirmar.
- **Cargo duplicado en el registro:** un cargo bancario único de $214.00
  ("Manguera para Tomar Agua") está registrado **dos veces** en la hoja, una vez
  contra "Toledo 2101-A" y otra contra "Toledo 2101-B".
- **Diferencia de 2 centavos:** un cargo bancario de $1,922.00 vs. un registro de
  $1,922.02 para el mismo concepto — no concilia por el criterio de monto exacto,
  pero es casi con certeza el mismo movimiento.
- Varios depósitos en efectivo ("DEPOSITO EFECTIVO PRACTIC") y transferencias de
  renta no tienen fila correspondiente el mismo mes — podrían estar registrados
  en otro mes (desfase de captura) o simplemente no capturados.

**Nota:** estos ejemplos son de septiembre 2025; los demás meses (oct 2025 – jun
2026) no se revisaron caso por caso todavía, solo se validaron los totales
agregados de la tabla de arriba.

---

## Limitaciones conocidas / posibles siguientes pasos

- El emparejamiento es por monto exacto — no detecta pagos partidos en varias
  transferencias ni combina varias filas de registro contra un solo movimiento
  bancario (o viceversa).
- No hay tolerancia de centavos — diferencias de $0.01–$0.02 (como la de
  "Pagos Copias R1" arriba) no concilian aunque sean casi con certeza el mismo
  movimiento.
- Si conviene, el siguiente paso natural sería revisar caso por caso los meses
  de octubre 2025 a junio 2026 igual que se hizo con septiembre, o agregar una
  tolerancia pequeña de monto para no perder esos casi-matches.
