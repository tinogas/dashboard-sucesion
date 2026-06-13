#!/usr/bin/env python3
"""
Dashboard Sucesion
Descarga la hoja 'movimientos', genera un Excel local con:
  - Copia de movimientos
  - Dashboard
  - Informe Contable
  - Control Rentas / Luz / Predial / Agua / Despacho
  (filtrados por columna Registro)
"""

import io
import os
import re
import sys
import unicodedata
from datetime import datetime

import pandas as pd
import requests

try:
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import Request as GoogleRequest
    import gspread
    import xlsxwriter
except ImportError:
    print("Dependencias faltantes. Ejecuta:  pip install -r requirements.txt")
    sys.exit(1)

# ── Configuracion ──────────────────────────────────────────────────────────────
SPREADSHEET_ID   = "1YTvNIui0kBSMWRs6mMrZvbM690ILMl_W"
SHEET_NAME_MOV   = "movimientos"
SHEET_GID        = "651730482"
CREDENTIALS_FILE = "credentials.json"
OUTPUT_FILE      = "dashboard_sucesion.xlsx"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

MESES = ["Ene","Feb","Mar","Abr","May","Jun",
         "Jul","Ago","Sep","Oct","Nov","Dic"]
MESES_LARGO = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

# (tab_name, color, valor_Registro, titulo, usar_ingresos, campo_fila)
# campo_fila: "inquilino" → usa columna Inquilino como filas
#             "inmueble"  → usa columna Inmueble como filas
CONTROL_TABS = [
    # (tab_name, color, registro, titulo, use_ingresos, campo_fila, resumen_concepto)
    ("Control Rentas",   "#27AE60", "renta",               "CONTROL DE PAGO DE RENTAS",         True,  "inquilino", False),
    ("Control Luz",      "#F39C12", "luz",                 "CONTROL DE PAGO DE LUZ (CFE)",      False, "inmueble",  False),
    ("Control Predial",  "#8E44AD", "predial",             "CONTROL DE PAGO DE PREDIAL",        False, "inmueble",  False),
    ("Control Agua",     "#2980B9", "agua",                "CONTROL DE PAGO DE AGUA",           False, "inmueble",  False),
    ("Control Despacho",  "#C0392B", "honorarios despacho", "CONTROL DE HONORARIOS — DESPACHO",  False, "inquilino", True),
    ("Control Impuestos", "#7D3C98", "impuestos",           "CONTROL DE PAGO DE IMPUESTOS",      False, "inmueble",  False),
]

# ── Autenticacion ──────────────────────────────────────────────────────────────
def _print_setup():
    print("""
╔══════════════════════════════════════════════════════════════╗
║        CONFIGURACION DE SERVICE ACCOUNT (SIN LOGIN)          ║
╠══════════════════════════════════════════════════════════════╣
║  1. console.cloud.google.com                                 ║
║  2. Habilita: Google Sheets API + Google Drive API           ║
║  3. IAM > Cuentas de servicio > Crear > descargar JSON       ║
║  4. Guarda como credentials.json en esta carpeta             ║
║  5. Comparte el spreadsheet con el email de la cuenta        ║
╚══════════════════════════════════════════════════════════════╝
""")

def get_credentials():
    if not os.path.exists(CREDENTIALS_FILE):
        _print_setup()
        sys.exit(1)
    return Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)


# ── Descarga de la hoja movimientos ───────────────────────────────────────────
def download_movimientos() -> pd.DataFrame:
    creds = get_credentials()

    # Intentar con gspread (Google Sheets nativo)
    try:
        client = gspread.authorize(creds)
        print("Conectado via gspread...")
        ws = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME_MOV)
        print(f"Descargando pestaña '{SHEET_NAME_MOV}'...")
        df = pd.DataFrame(ws.get_all_records())
        df = df.fillna("")
        return df
    except gspread.exceptions.APIError as e:
        if "not supported" not in str(e) and "Office file" not in str(e):
            raise

    # Fallback: descargar xlsx completo y leer solo la hoja movimientos
    print("Archivo formato Office — descargando xlsx completo via Drive API...")
    creds.refresh(GoogleRequest())
    url = (f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
           f"/export?format=xlsx")
    resp = requests.get(
        url, headers={"Authorization": f"Bearer {creds.token}"}, timeout=60
    )
    resp.raise_for_status()

    all_sheets = pd.read_excel(io.BytesIO(resp.content), sheet_name=None, dtype=str)

    # Buscar pestaña movimientos (flexible)
    for name, df in all_sheets.items():
        if SHEET_NAME_MOV.lower() in name.lower():
            print(f"  Pestaña encontrada: '{name}'")
            return df.fillna("")

    # Si no la encuentra, usar la primera
    first = list(all_sheets.keys())[0]
    print(f"  AVISO: usando primera hoja '{first}'")
    return all_sheets[first].fillna("")


# ── Helpers ────────────────────────────────────────────────────────────────────
import re

MESES_MAP = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12,
    "ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
    "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12,
}
# Palabras a eliminar al limpiar Concepto para obtener el inquilino
_STOP_WORDS = re.compile(
    r"\b(renta|pago|de|del|mes|periodo|correspondiente|agua|luz|cfe|predial|honorarios|despacho)\b",
    re.IGNORECASE,
)
_MES_PATTERN = re.compile(
    r"\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|"
    r"octubre|noviembre|diciembre|ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)\b",
    re.IGNORECASE,
)
_YEAR_PATTERN = re.compile(r"\b(19|20)\d{2}\b")


def extract_month_from_text(text: str) -> "int | None":
    """Devuelve el número de mes (1-12) encontrado en el texto, o None."""
    m = _MES_PATTERN.search(str(text))
    return MESES_MAP.get(m.group().lower()) if m else None


def extract_entity_from_concepto(text: str) -> str:
    """Elimina mes, año y palabras genéricas del Concepto para obtener el nombre del inquilino."""
    t = _MES_PATTERN.sub("", str(text))
    t = _YEAR_PATTERN.sub("", t)
    t = _STOP_WORDS.sub("", t)
    t = re.sub(r"[\s\-_/|]+", " ", t).strip(" -_/|")
    return t or str(text).strip()


def clean_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(r"[$,\s%]", "", regex=True),
        errors="coerce",
    )

def parse_dates(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")

def parse_mes_col(series: pd.Series) -> pd.Series:
    """Convierte columna Mes (nombre o número) a entero 1-12."""
    result = []
    for val in series:
        v = str(val).strip().lower()
        # Número directo
        try:
            n = int(float(v))
            result.append(n if 1 <= n <= 12 else None)
            continue
        except (ValueError, TypeError):
            pass
        # Nombre de mes
        num = MESES_MAP.get(v)
        if num is None:
            # Coincidencia parcial (ej. "Enero 2024")
            for name, n in MESES_MAP.items():
                if name in v:
                    num = n
                    break
        result.append(num)
    return pd.array(result, dtype="Int64")

def detect_columns(df: pd.DataFrame):
    """Retorna (fecha, ingreso, egreso, concepto, registro, inquilino, inmueble, año, mes)."""
    def find(keywords):
        for col in df.columns:
            if any(kw in col.lower() for kw in keywords):
                return col
        return None
    def find_exact(names):
        # Normalize to NFC so ñ from Google Sheets (may arrive NFD) matches our literals
        for col in df.columns:
            norm = unicodedata.normalize("NFC", col.strip().lower())
            if norm in names:
                return col
        # Fallback: column name starts with one of the keywords (e.g. "Año Contable")
        for col in df.columns:
            norm = unicodedata.normalize("NFC", col.strip().lower())
            for n in names:
                if re.match(rf"^{re.escape(n)}\b", norm):
                    return col
        return None
    return (
        find({"fecha", "date", "dia"}),
        find({"ingreso", "entrada", "cobro", "haber"}),
        find({"egreso", "salida", "gasto", "debe"}),
        find({"concepto", "descripcion", "detalle"}),
        find({"registro", "cuenta", "rubro", "categoria"}),
        find({"inquilino", "arrendatario", "locatario", "cliente"}),
        find({"inmueble", "propiedad", "local", "predio", "bien"}),
        find_exact({"año", "anio", "year", "ejercicio"}),
        find_exact({"mes", "month", "periodo"}),
    )

def prepare_año_mes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega _año y _mes al DataFrame.
    Fuente primaria: columnas 'Año' y 'Mes'.
    Fallback: columna 'Fecha'.
    """
    fecha_col, _, _, _, _, _, _, año_col, mes_col = detect_columns(df)
    print(f"  [año/mes] fecha='{fecha_col}' | año='{año_col}' | mes='{mes_col}'")
    df = df.copy()

    if año_col:
        _s = pd.to_numeric(df[año_col].astype(str).str.strip(), errors="coerce").round(0)
        _na = _s.isna().to_numpy()
        df["_año"] = pd.arrays.IntegerArray(_s.fillna(0).to_numpy().astype("int64"), _na)
    elif fecha_col:
        df["_año"] = parse_dates(df[fecha_col]).dt.year.astype("Int64")
    else:
        df["_año"] = pd.NA

    if mes_col:
        df["_mes"] = parse_mes_col(df[mes_col])
    elif fecha_col:
        df["_mes"] = parse_dates(df[fecha_col]).dt.month.astype("Int64")
    else:
        df["_mes"] = pd.NA

    return df

def filter_by_registro(df: pd.DataFrame, reg_col: str, value: str) -> pd.DataFrame:
    """Filtra por Registro con coincidencia exacta (sin importar mayusculas)."""
    if not reg_col or reg_col not in df.columns:
        return df.iloc[0:0]
    mask = df[reg_col].astype(str).str.strip().str.lower() == value.lower()
    if mask.sum() == 0:
        # Fallback: coincidencia parcial
        mask = df[reg_col].astype(str).str.strip().str.lower().str.contains(
            value.lower(), regex=False
        )
    return df[mask].copy()


# ── Estilos ────────────────────────────────────────────────────────────────────
BLUE_DARK  = "#1F4E79"
BLUE_MED   = "#2E86C1"
BLUE_LIGHT = "#D6E4F0"
GRAY_ROW   = "#F2F2F2"
GREEN_MED  = "#27AE60"
GREEN_LIGHT= "#D5F5E3"
RED_LIGHT  = "#FADBD8"

def add_formats(wb):
    def f(**kw):
        return wb.add_format(kw)
    return {
        "title":      f(bold=True, font_size=16, font_color="white",
                        bg_color=BLUE_DARK, align="center", valign="vcenter"),
        "subtitle":   f(italic=True, font_size=10, font_color="white",
                        bg_color=BLUE_DARK, align="center"),
        "section":    f(bold=True, font_size=11, font_color="white",
                        bg_color=BLUE_MED),
        "header":     f(bold=True, font_color="white", bg_color=BLUE_MED,
                        border=1, align="center", valign="vcenter", text_wrap=True),
        "hdr_green":  f(bold=True, font_color="white", bg_color="#1E5631",
                        border=1, align="center", valign="vcenter"),
        "cell":       f(border=1, font_size=10, valign="vcenter"),
        "cell_alt":   f(border=1, font_size=10, valign="vcenter", bg_color=GRAY_ROW),
        "money":      f(border=1, num_format="$#,##0.00", font_size=10),
        "money_alt":  f(border=1, num_format="$#,##0.00", font_size=10, bg_color=GRAY_ROW),
        "money_tot":  f(bold=True, border=1, num_format="$#,##0.00",
                        bg_color=BLUE_LIGHT, font_size=10),
        "money_pos":  f(bold=True, border=1, num_format="$#,##0.00",
                        bg_color=GREEN_LIGHT, font_size=10, font_color="#1E5631"),
        "money_neg":  f(bold=True, border=1, num_format="$#,##0.00",
                        bg_color=RED_LIGHT, font_size=10, font_color="#C0392B"),
        "date":       f(border=1, num_format="dd/mm/yyyy", font_size=10),
        "date_alt":   f(border=1, num_format="dd/mm/yyyy", font_size=10, bg_color=GRAY_ROW),
        "kpi_lbl":    f(bold=True, font_size=11, font_color=BLUE_DARK,
                        bg_color=BLUE_LIGHT, border=1, align="center"),
        "kpi_val":    f(bold=True, font_size=18, num_format="$#,##0.00",
                        border=1, align="center", valign="vcenter"),
        "kpi_int":    f(bold=True, font_size=18, num_format="#,##0",
                        border=1, align="center", valign="vcenter"),
        "total_lbl":  f(bold=True, border=1, bg_color=BLUE_DARK,
                        font_color="white", font_size=10),
        "pivot_hdr":  f(bold=True, font_color="white", bg_color=BLUE_MED,
                        border=1, align="center", text_wrap=True, font_size=9),
        "pivot_year": f(bold=True, font_color="white", bg_color=BLUE_DARK,
                        border=1, align="center", font_size=10),
        "pivot_val":  f(border=1, num_format="$#,##0.00", font_size=9, align="right"),
        "pivot_alt":  f(border=1, num_format="$#,##0.00", font_size=9,
                        align="right", bg_color=GRAY_ROW),
        "pivot_tot":  f(bold=True, border=1, num_format="$#,##0.00",
                        font_size=9, align="right", bg_color=BLUE_LIGHT),
        "pivot_zero": f(border=1, font_size=9, align="center",
                        font_color="#BBBBBB"),
    }


# ══════════════════════════════════════════════════════════════════════════════
# HOJAS
# ══════════════════════════════════════════════════════════════════════════════

def write_movimientos(wb, fmt, df: pd.DataFrame):
    """Copia formateada de la hoja movimientos."""
    ws = wb.add_worksheet("Movimientos")
    ws.set_zoom(90)
    ws.freeze_panes(1, 0)
    ws.set_tab_color("#1F4E79")

    fecha_col, ing_col, egr_col, _, _, _, _, _, _ = detect_columns(df)

    # Solo columnas A–I (primeras 11)
    cols = list(df.columns[:11])
    data = df[cols].copy().reset_index(drop=True)

    ws.set_row(0, 22)
    for ci, col in enumerate(cols):
        ws.write(0, ci, col, fmt["header"])
        ws.set_column(ci, ci, max(13, min(len(str(col)) + 6, 40)))

    for ri, row in data.iterrows():
        alt = ri % 2 == 1
        for ci, col in enumerate(cols):
            val = row[col]
            is_money = col in [ing_col, egr_col]
            is_date  = col == fecha_col

            if is_date:
                dt = parse_dates(pd.Series([val])).iloc[0]
                f  = fmt["date_alt"] if alt else fmt["date"]
                if pd.notna(dt):
                    ws.write_datetime(ri+1, ci, dt.to_pydatetime(), f)
                else:
                    ws.write(ri+1, ci, str(val), f)
            elif is_money:
                num = clean_numeric(pd.Series([val])).iloc[0]
                f   = fmt["money_alt"] if alt else fmt["money"]
                ws.write(ri+1, ci, num if pd.notna(num) else "", f)
            else:
                f = fmt["cell_alt"] if alt else fmt["cell"]
                ws.write(ri+1, ci, str(val) if val != "" else "", f)

    ws.autofilter(0, 0, len(data), len(cols)-1)
    print(f"  Hoja copiada: 'Movimientos'  ({len(data)} filas)")


def write_dashboard(wb, fmt, df: pd.DataFrame):
    """Dashboard con KPIs y graficos."""
    ws = wb.add_worksheet("Dashboard")
    ws.activate()
    ws.hide_gridlines(2)
    ws.set_zoom(90)
    ws.set_tab_color(BLUE_MED)

    fecha_col, ing_col, egr_col, conc_col, _, _, _, _, _ = detect_columns(df)

    ingresos = clean_numeric(df[ing_col]) if ing_col else pd.Series(dtype=float)
    egresos  = clean_numeric(df[egr_col]) if egr_col else pd.Series(dtype=float)
    fechas   = parse_dates(df[fecha_col]) if fecha_col else pd.Series(dtype="datetime64[ns]")

    total_ing = ingresos.sum()
    total_egr = egresos.sum()
    balance   = total_ing - total_egr

    ws.set_column(0, 0, 2)
    ws.set_column(1, 12, 11)
    ws.set_column(13, 13, 2)

    r = 1
    ws.set_row(r, 42)
    ws.merge_range(r, 1, r, 12, "DASHBOARD DE MOVIMIENTOS — SUCESION", fmt["title"])
    r += 1
    ws.set_row(r, 16)
    ws.merge_range(r, 1, r, 12,
        f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}   |   {len(df):,} registros",
        fmt["subtitle"])
    r += 2

    # KPIs
    ws.set_row(r, 18)
    ws.merge_range(r, 1, r, 12, "  RESUMEN FINANCIERO", fmt["section"])
    r += 1
    ws.set_row(r, 20); ws.set_row(r+1, 36)
    kpis = [
        ("Total Registros", len(df),     fmt["kpi_int"], 1, 3),
        ("Total Ingresos",  total_ing,   fmt["kpi_val"], 4, 6),
        ("Total Egresos",   total_egr,   fmt["kpi_val"], 7, 9),
        ("Balance",         balance,     fmt["kpi_val"], 10, 12),
    ]
    for lbl, val, vfmt, c1, c2 in kpis:
        ws.merge_range(r,   c1, r,   c2, lbl, fmt["kpi_lbl"])
        ws.merge_range(r+1, c1, r+1, c2, val if pd.notna(val) else 0, vfmt)
    r += 3

    # Datos auxiliares para graficos
    ws_aux = wb.add_worksheet("_aux")
    ws_aux.hide()
    aux_r = 0
    monthly_range = tipo_range = None

    if ing_col and "_año" in df.columns and "_mes" in df.columns:
        tmp = df.copy()
        tmp["_i"] = ingresos
        tmp["_e"] = egresos if egr_col else 0
        tmp = tmp.dropna(subset=["_año", "_mes"])
        # Etiqueta Año-Mes para el eje
        tmp["_periodo"] = (tmp["_año"].astype(str) + "-" +
                           tmp["_mes"].astype(str).str.zfill(2))
        monthly = (tmp.groupby("_periodo")[["_i","_e"]]
                   .sum().reset_index().sort_values("_periodo").tail(24))
        ws_aux.write(aux_r, 0, "Mes")
        ws_aux.write(aux_r, 1, "Ingresos")
        ws_aux.write(aux_r, 2, "Egresos")
        monthly.columns = ["mes", "ing", "egr"]
        for ii, row in enumerate(monthly.itertuples(index=False)):
            ws_aux.write(aux_r+1+ii, 0, row.mes)
            ws_aux.write(aux_r+1+ii, 1, row.ing if pd.notna(row.ing) else 0)
            ws_aux.write(aux_r+1+ii, 2, row.egr if pd.notna(row.egr) else 0)
        n = len(monthly)
        monthly_range = {
            "cats": f"=_aux!$A$2:$A${n+1}",
            "ing":  f"=_aux!$B$2:$B${n+1}",
            "egr":  f"=_aux!$C$2:$C${n+1}",
        }
        aux_r += n + 3

    if conc_col and ing_col:
        tmp2 = df.copy()
        tmp2["_i"] = ingresos
        grp = tmp2.groupby(conc_col)["_i"].sum().reset_index()
        grp = grp.sort_values("_i", ascending=False).head(10)
        ws_aux.write(aux_r, 0, conc_col)
        ws_aux.write(aux_r, 1, "Total")
        grp.columns = ["concepto", "total"]
        for ii, row in enumerate(grp.itertuples(index=False)):
            ws_aux.write(aux_r+1+ii, 0, str(row.concepto))
            ws_aux.write(aux_r+1+ii, 1, row.total if pd.notna(row.total) else 0)
        tipo_range = {
            "cats": f"=_aux!$A${aux_r+2}:$A${aux_r+1+len(grp)}",
            "vals": f"=_aux!$B${aux_r+2}:$B${aux_r+1+len(grp)}",
        }

    # Grafico mensual
    ws.set_row(r, 18)
    ws.merge_range(r, 1, r, 12, "  EVOLUCION MENSUAL (Ingresos vs Egresos)", fmt["section"])
    r += 1
    if monthly_range:
        chart = wb.add_chart({"type": "column"})
        chart.add_series({"name":"Ingresos","categories":monthly_range["cats"],
                          "values":monthly_range["ing"],"fill":{"color":BLUE_MED},"gap":40})
        chart.add_series({"name":"Egresos","categories":monthly_range["cats"],
                          "values":monthly_range["egr"],"fill":{"color":"#E74C3C"},"gap":40})
        chart.set_x_axis({"num_font":{"rotation":-45}})
        chart.set_legend({"position":"bottom"})
        chart.set_chartarea({"border":{"color":"#CCCCCC"}})
        chart.set_size({"width":750, "height":300})
        ws.insert_chart(r, 1, chart)
    r += 17

    # Graficos por concepto
    ws.set_row(r, 18)
    ws.merge_range(r, 1, r, 12, "  INGRESOS POR CONCEPTO (Top 10)", fmt["section"])
    r += 1
    if tipo_range:
        chart2 = wb.add_chart({"type":"bar"})
        chart2.add_series({"name":"Total","categories":tipo_range["cats"],
                           "values":tipo_range["vals"],"fill":{"color":GREEN_MED}})
        chart2.set_legend({"none":True})
        chart2.set_chartarea({"border":{"color":"#CCCCCC"}})
        chart2.set_size({"width":450,"height":320})
        ws.insert_chart(r, 1, chart2)

        chart3 = wb.add_chart({"type":"pie"})
        chart3.add_series({"categories":tipo_range["cats"],"values":tipo_range["vals"],
                           "data_labels":{"percentage":True,"category":True,"font":{"size":9}}})
        chart3.set_legend({"position":"bottom"})
        chart3.set_chartarea({"border":{"color":"#CCCCCC"}})
        chart3.set_size({"width":360,"height":320})
        ws.insert_chart(r, 7, chart3)

    print("  Hoja generada: 'Dashboard'")


def write_informe_dinamico(wb, fmt, df: pd.DataFrame):
    """Informe expandible: Año › Mes › Registro › Inmueble › Concepto."""
    ws = wb.add_worksheet("Informe Dinámico")
    ws.hide_gridlines(2)
    ws.set_tab_color("#16A085")
    ws.set_zoom(90)
    ws.outline_settings(True, False, True, True)   # totales encima, detalle abajo

    _, ing_col, egr_col, conc_col, reg_col, _, inm_col, _, _ = detect_columns(df)

    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 50)
    ws.set_column(2, 4, 17)

    r = 1
    ws.set_row(r, 40)
    ws.merge_range(r, 1, r, 4, "INFORME CONTABLE DINÁMICO", fmt["title"])
    r += 1
    ws.merge_range(r, 1, r, 4,
        f"Expandible: Año › Mes › Registro › Inmueble › Concepto   |   "
        f"Generado: {datetime.now().strftime('%d/%m/%Y')}",
        fmt["subtitle"])
    r += 2

    ws.set_row(r, 20)
    for ci, hdr in enumerate(["Período / Detalle", "Ingresos", "Egresos", "Balance"]):
        ws.write(r, ci+1, hdr, fmt["hdr_green"])
    r += 1

    tmp = df.copy()
    tmp["_ing"] = clean_numeric(tmp[ing_col]) if ing_col else 0
    tmp["_egr"] = clean_numeric(tmp[egr_col]) if egr_col else 0

    def wm(row, col, val, f):
        ws.write(row, col, float(val) if pd.notna(val) else 0.0, f)

    grand_ing = float(tmp["_ing"].sum())
    grand_egr = float(tmp["_egr"].sum())

    # ── Nivel 1: AÑO ─────────────────────────────────────────────────────────
    for año in sorted(tmp["_año"].dropna().unique().astype(int)):
        año_data = tmp[tmp["_año"] == año]
        a_ing = float(año_data["_ing"].sum())
        a_egr = float(año_data["_egr"].sum())

        ws.set_row(r, 22, None, {"level": 1, "collapsed": True})
        ws.write(r, 1, f"  AÑO {año}", fmt["pivot_year"])
        wm(r, 2, a_ing, fmt["money_tot"])
        wm(r, 3, a_egr, fmt["money_tot"])
        wm(r, 4, a_ing - a_egr, fmt["money_tot"])
        r += 1

        # ── Nivel 2: MES ─────────────────────────────────────────────────────
        for mes in sorted(año_data["_mes"].dropna().unique().astype(int)):
            mes_data = año_data[año_data["_mes"] == mes]
            m_ing = float(mes_data["_ing"].sum())
            m_egr = float(mes_data["_egr"].sum())

            ws.set_row(r, 20, None, {"level": 2, "hidden": True, "collapsed": True})
            ws.write(r, 1, f"    {MESES_LARGO[mes-1]} {año}", fmt["pivot_hdr"])
            wm(r, 2, m_ing, fmt["money_alt"])
            wm(r, 3, m_egr, fmt["money_alt"])
            wm(r, 4, m_ing - m_egr, fmt["money_alt"])
            r += 1

            # ── Nivel 3: REGISTRO ─────────────────────────────────────────────
            registros = sorted(mes_data[reg_col].dropna().unique()) if reg_col else []
            for reg in registros:
                reg_data = mes_data[mes_data[reg_col] == reg]
                rg_ing = float(reg_data["_ing"].sum())
                rg_egr = float(reg_data["_egr"].sum())

                ws.set_row(r, 18, None, {"level": 3, "hidden": True, "collapsed": True})
                ws.write(r, 1, f"      {str(reg)}", fmt["section"])
                wm(r, 2, rg_ing, fmt["money_tot"])
                wm(r, 3, rg_egr, fmt["money_tot"])
                wm(r, 4, rg_ing - rg_egr, fmt["money_tot"])
                r += 1

                # ── Nivel 4: INMUEBLE ─────────────────────────────────────────
                inmuebles = sorted(reg_data[inm_col].dropna().unique()) if inm_col else []
                for inm in inmuebles:
                    inm_data = reg_data[reg_data[inm_col] == inm]
                    i_ing = float(inm_data["_ing"].sum())
                    i_egr = float(inm_data["_egr"].sum())

                    ws.set_row(r, 18, None, {"level": 4, "hidden": True, "collapsed": True})
                    ws.write(r, 1, f"        {str(inm)}", fmt["pivot_hdr"])
                    wm(r, 2, i_ing, fmt["money_alt"])
                    wm(r, 3, i_egr, fmt["money_alt"])
                    wm(r, 4, i_ing - i_egr, fmt["money_alt"])
                    r += 1

                    # ── Nivel 5: CONCEPTO ─────────────────────────────────────
                    if conc_col:
                        concs = (inm_data.groupby(conc_col, sort=False)
                                 .agg(ing=("_ing","sum"), egr=("_egr","sum"))
                                 .reset_index()
                                 .sort_values("ing", ascending=False))
                        for idx, (_, crow) in enumerate(concs.iterrows()):
                            alt = idx % 2 == 1
                            c_ing = float(crow["ing"]) if pd.notna(crow["ing"]) else 0.0
                            c_egr = float(crow["egr"]) if pd.notna(crow["egr"]) else 0.0
                            ws.set_row(r, None, None, {"level": 5, "hidden": True})
                            ws.write(r, 1, f"          {str(crow[conc_col])}",
                                     fmt["cell_alt"] if alt else fmt["cell"])
                            ws.write(r, 2, c_ing, fmt["money_alt"] if alt else fmt["money"])
                            ws.write(r, 3, c_egr, fmt["money_alt"] if alt else fmt["money"])
                            ws.write(r, 4, c_ing - c_egr, fmt["money_alt"] if alt else fmt["money"])
                            r += 1

    # ── Gran Total ────────────────────────────────────────────────────────────
    r += 1
    ws.set_row(r, 24)
    ws.write(r, 1, "GRAN TOTAL", fmt["total_lbl"])
    wm(r, 2, grand_ing, fmt["money_tot"])
    wm(r, 3, grand_egr, fmt["money_tot"])
    wm(r, 4, grand_ing - grand_egr, fmt["money_tot"])

    print("  Hoja generada: 'Informe Dinámico'")


def write_informe_contable(wb, fmt, df: pd.DataFrame):
    """Informe contable agrupado por Registro."""
    ws = wb.add_worksheet("Informe Contable")
    ws.hide_gridlines(2)
    ws.set_tab_color(GREEN_MED)
    ws.set_zoom(90)

    _, ing_col, egr_col, _, reg_col, _, _, _, _ = detect_columns(df)

    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 35)
    ws.set_column(2, 4, 16)
    ws.set_column(5, 5, 10)

    r = 1
    ws.set_row(r, 40)
    ws.merge_range(r, 1, r, 5, "INFORME CONTABLE — SUCESION", fmt["title"])
    r += 1
    ws.set_row(r, 16)

    tmp = df.copy()
    tmp["_ing"] = clean_numeric(tmp[ing_col]) if ing_col else 0
    tmp["_egr"] = clean_numeric(tmp[egr_col]) if egr_col else 0
    tmp["_bal"] = tmp["_ing"].fillna(0) - tmp["_egr"].fillna(0)

    años = tmp["_año"].dropna()
    periodo = (f"{int(años.min())} – {int(años.max())}"
               if not años.empty else "—")

    ws.merge_range(r, 1, r, 5,
        f"Periodo: {periodo}   |   Generado: {datetime.now().strftime('%d/%m/%Y')}",
        fmt["subtitle"])
    r += 2

    # Resumen general
    ws.set_row(r, 18)
    ws.merge_range(r, 1, r, 5, "  RESUMEN GENERAL", fmt["section"])
    r += 1
    for lbl, val in [("Total Ingresos", tmp["_ing"].sum()),
                     ("Total Egresos",  tmp["_egr"].sum()),
                     ("Balance Neto",   tmp["_bal"].sum())]:
        fv = fmt["money_pos"] if val >= 0 else fmt["money_neg"]
        ws.write(r, 1, lbl, fmt["kpi_lbl"])
        ws.write(r, 2, val if pd.notna(val) else 0, fv)
        r += 1
    r += 1

    # Detalle por Registro
    if reg_col:
        ws.set_row(r, 18)
        ws.merge_range(r, 1, r, 5, f"  DETALLE POR {reg_col.upper()}", fmt["section"])
        r += 1
        ws.set_row(r, 20)
        for ci, hdr in enumerate(["Registro / Cuenta","Ingresos","Egresos","Balance","% Ingr."]):
            ws.write(r, ci+1, hdr, fmt["hdr_green"])
        r += 1

        tot_ing = tmp["_ing"].sum()
        grouped = tmp.groupby(reg_col).agg(ing=("_ing","sum"),
                                            egr=("_egr","sum")).reset_index()
        grouped["bal"] = grouped["ing"] - grouped["egr"]
        grouped = grouped.sort_values("ing", ascending=False)

        for i, row in grouped.iterrows():
            alt = i % 2 == 1
            fc  = fmt["cell_alt"] if alt else fmt["cell"]
            fm  = fmt["money_alt"] if alt else fmt["money"]
            ws.write(r, 1, str(row[reg_col]), fc)
            ws.write(r, 2, row["ing"] if pd.notna(row["ing"]) else 0, fm)
            ws.write(r, 3, row["egr"] if pd.notna(row["egr"]) else 0, fm)
            ws.write(r, 4, row["bal"] if pd.notna(row["bal"]) else 0, fm)
            pct = (row["ing"] / tot_ing * 100) if tot_ing else 0
            ws.write(r, 5, f"{pct:.1f}%", fc)
            r += 1

        ws.write(r, 1, "TOTALES", fmt["total_lbl"])
        ws.write(r, 2, grouped["ing"].sum(), fmt["money_tot"])
        ws.write(r, 3, grouped["egr"].sum(), fmt["money_tot"])
        ws.write(r, 4, grouped["bal"].sum(), fmt["money_tot"])
        ws.write(r, 5, "100%", fmt["total_lbl"])
        r += 2

    # Desglose por año
    ws.set_row(r, 18)
    ws.merge_range(r, 1, r, 5, "  DESGLOSE POR AÑO", fmt["section"])
    r += 1
    ws.set_row(r, 20)
    for ci, hdr in enumerate(["Año","Ingresos","Egresos","Balance"]):
        ws.write(r, ci+1, hdr, fmt["hdr_green"])
    r += 1
    anual = tmp.groupby("_año").agg(ing=("_ing","sum"),
                                     egr=("_egr","sum")).reset_index()
    anual["bal"] = anual["ing"] - anual["egr"]
    for i, row in anual.iterrows():
        alt = i % 2 == 1
        fc  = fmt["cell_alt"] if alt else fmt["cell"]
        fm  = fmt["money_alt"] if alt else fmt["money"]
        año = int(row["_año"]) if pd.notna(row["_año"]) else ""
        ws.write(r, 1, año, fc)
        ws.write(r, 2, row["ing"] if pd.notna(row["ing"]) else 0, fm)
        ws.write(r, 3, row["egr"] if pd.notna(row["egr"]) else 0, fm)
        ws.write(r, 4, row["bal"] if pd.notna(row["bal"]) else 0, fm)
        r += 1
    ws.write(r, 1, "TOTALES", fmt["total_lbl"])
    ws.write(r, 2, anual["ing"].sum(), fmt["money_tot"])
    ws.write(r, 3, anual["egr"].sum(), fmt["money_tot"])
    ws.write(r, 4, anual["bal"].sum(), fmt["money_tot"])

    print("  Hoja generada: 'Informe Contable'")


def write_control_tab(wb, fmt, df: pd.DataFrame,
                      tab_name: str, color: str,
                      registro_value: str, title: str,
                      use_ingresos: bool, campo_fila: str,
                      resumen_concepto: bool = False):
    """
    Tabla de control filtrada por Registro = registro_value.
    - Año/Mes: columnas dedicadas (pre-computadas por prepare_año_mes)
    - Entidad: columna Inquilino (rentas/despacho) o Inmueble (luz/predial/agua)
    """
    ws = wb.add_worksheet(tab_name)
    ws.hide_gridlines(2)
    ws.set_tab_color(color)
    ws.set_zoom(90)

    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 36)
    ws.set_column(2, 14, 11)

    _, ing_col, egr_col, conc_col, reg_col, inq_col, inm_col, _, _ = detect_columns(df)
    amount_col = (ing_col if use_ingresos and ing_col else egr_col) or ing_col

    # Columna de filas según campo_fila
    if campo_fila == "inquilino":
        fila_col  = inq_col
        fila_hdr  = "Inquilino"
    else:
        fila_col  = inm_col
        fila_hdr  = "Inmueble"

    r = 1
    ws.set_row(r, 40)
    ws.merge_range(r, 1, r, 14, title, fmt["title"])
    r += 1
    ws.set_row(r, 16)
    ws.merge_range(r, 1, r, 14,
        f"Registro: '{registro_value}'   |   Generado: {datetime.now().strftime('%d/%m/%Y')}",
        fmt["subtitle"])
    r += 2

    # ── Filtrar por Registro ───────────────────────────────────────────────────
    filtered = filter_by_registro(df, reg_col, registro_value)

    if filtered.empty:
        ws.merge_range(r, 1, r, 10,
            f"Sin registros con Registro = '{registro_value}'", fmt["section"])
        print(f"  AVISO: sin datos para '{tab_name}' (Registro='{registro_value}')")
        return

    filtered = filtered.copy()

    # ── Monto ──────────────────────────────────────────────────────────────────
    filtered["_monto"] = clean_numeric(filtered[amount_col]) if amount_col else 0

    # ── Filas del pivot: Inquilino o Inmueble ──────────────────────────────────
    if fila_col and fila_col in filtered.columns:
        filtered["_ent"] = filtered[fila_col].astype(str).str.strip()
    else:
        # Si no existe la columna, avisar y usar "Sin datos"
        print(f"  AVISO: columna '{campo_fila}' no encontrada en los datos")
        filtered["_ent"] = "Sin datos"

    filtered = filtered.dropna(subset=["_año", "_mes", "_monto"])
    filtered = filtered[filtered["_monto"] != 0]

    entidades = sorted(filtered["_ent"].replace("", pd.NA).dropna().unique())
    entidades = [e for e in entidades if e and e.lower() != "nan"]
    if not entidades:
        entidades = ["(sin registro)"]

    años = sorted(filtered["_año"].dropna().unique().astype(int))
    total_grand = 0.0

    for año in años:
        año_data = filtered[filtered["_año"] == año]

        ws.set_row(r, 20)
        ws.merge_range(r, 1, r, 14, f"  AÑO {año}", fmt["pivot_year"])
        r += 1

        ws.set_row(r, 18)
        ws.write(r, 1, fila_hdr, fmt["pivot_hdr"])
        for m in range(1, 13):
            ws.write(r, m+1, MESES[m-1], fmt["pivot_hdr"])
        ws.write(r, 14, "Total Año", fmt["pivot_hdr"])
        r += 1

        col_totals = {m: 0.0 for m in range(1, 13)}
        año_total  = 0.0

        for ei, ent in enumerate(entidades):
            alt = ei % 2 == 1
            fp  = fmt["pivot_alt"] if alt else fmt["pivot_val"]
            fz  = fmt["pivot_zero"]
            fc  = fmt["cell_alt"] if alt else fmt["cell"]

            ent_data = año_data[año_data["_ent"] == ent]

            ws.write(r, 1, ent, fc)
            row_total = 0.0
            for m in range(1, 13):
                mes_vals = ent_data[ent_data["_mes"] == m]["_monto"]
                val = float(mes_vals.sum()) if not mes_vals.empty else 0.0
                if val:
                    ws.write(r, m+1, val, fp)
                    col_totals[m] += val
                    row_total     += val
                else:
                    ws.write(r, m+1, "—", fz)
            ws.write(r, 14, row_total if row_total else "—",
                     fmt["pivot_tot"] if row_total else fz)
            año_total += row_total
            r += 1

        ws.write(r, 1, f"TOTAL {año}", fmt["total_lbl"])
        for m in range(1, 13):
            ws.write(r, m+1,
                     col_totals[m] if col_totals[m] else "—",
                     fmt["pivot_tot"] if col_totals[m] else fmt["pivot_zero"])
        ws.write(r, 14, año_total, fmt["pivot_tot"])
        total_grand += año_total
        r += 2

    ws.set_row(r, 20)
    ws.merge_range(r, 1, r, 13, "GRAN TOTAL", fmt["total_lbl"])
    ws.write(r, 14, total_grand, fmt["money_tot"])
    r += 3

    # ── Resumen por Concepto (opcional, solo Despacho) ─────────────────────────
    if resumen_concepto and conc_col:
        ws.set_row(r, 18)
        ws.merge_range(r, 1, r, 6, "  RESUMEN POR CONCEPTO", fmt["section"])
        r += 1
        ws.set_row(r, 18)
        ws.write(r, 1, "Concepto",  fmt["hdr_green"])
        ws.write(r, 2, "Registros", fmt["hdr_green"])
        ws.write(r, 3, "Total",     fmt["hdr_green"])
        r += 1

        resumen = (filtered.groupby(conc_col)
                   .agg(registros=(conc_col, "count"), total=("_monto", "sum"))
                   .reset_index()
                   .sort_values("total", ascending=False))

        for i, row in resumen.iterrows():
            alt = i % 2 == 1
            fc  = fmt["cell_alt"]  if alt else fmt["cell"]
            fm  = fmt["money_alt"] if alt else fmt["money"]
            ws.write(r, 1, str(row[conc_col]), fc)
            ws.write(r, 2, int(row["registros"]), fc)
            ws.write(r, 3, float(row["total"]) if pd.notna(row["total"]) else 0, fm)
            r += 1

        ws.write(r, 1, "TOTAL", fmt["total_lbl"])
        ws.write(r, 2, int(resumen["registros"].sum()), fmt["total_lbl"])
        ws.write(r, 3, float(resumen["total"].sum()), fmt["money_tot"])

    print(f"  Hoja generada: '{tab_name}'  ({len(filtered)} registros)")


# ── Detalle por Registro ───────────────────────────────────────────────────────
def write_detalle_inmuebles(wb, fmt, df: pd.DataFrame):
    """Hoja de detalle ordenada por Registro › Inmueble › Año › Mes.
    Sección por cada Registro con subtotal; dentro de cada sección las filas
    muestran Inmueble, Año, Mes, Ingreso, Egreso, Concepto."""
    _, ing_col, egr_col, conc_col, reg_col, _, inm_col, _, _ = detect_columns(df)

    MESES_C = ["Ene","Feb","Mar","Abr","May","Jun",
               "Jul","Ago","Sep","Oct","Nov","Dic"]

    ws = wb.add_worksheet("Detalle Registro")
    ws.hide_gridlines(2)
    ws.set_tab_color("#16A085")
    ws.set_zoom(90)

    ws.set_column(0, 0,  2)   # indent
    ws.set_column(1, 1, 36)   # Inmueble
    ws.set_column(2, 2,  6)   # Año
    ws.set_column(3, 3,  8)   # Mes
    ws.set_column(4, 4, 14)   # Ingreso
    ws.set_column(5, 5, 14)   # Egreso
    ws.set_column(6, 6, 44)   # Concepto

    data = df.copy()
    data["_reg"] = (data[reg_col].astype(str).str.strip()
                    if reg_col else pd.Series("Sin registro", index=data.index))
    data["_inm"] = (data[inm_col].astype(str).str.strip()
                    if inm_col else pd.Series("Sin inmueble", index=data.index))

    data["_ing_n"] = clean_numeric(data[ing_col]).fillna(0) if ing_col else 0
    data["_egr_n"] = clean_numeric(data[egr_col]).fillna(0) if egr_col else 0

    data = data.sort_values(["_reg", "_inm", "_año", "_mes"], na_position="last")

    r = 1
    ws.set_row(r, 40)
    ws.merge_range(r, 1, r, 6,
        "DETALLE DE MOVIMIENTOS POR REGISTRO", fmt["title"])
    r += 1
    ws.set_row(r, 16)
    ws.merge_range(r, 1, r, 6,
        f"Generado: {datetime.now().strftime('%d/%m/%Y')}",
        fmt["subtitle"])
    r += 2

    registros = sorted(data["_reg"].unique())
    total_ing = 0.0
    total_egr = 0.0
    HDR = ["Inmueble", "Año", "Mes", "Ingreso", "Egreso", "Concepto"]

    for reg in registros:
        df_reg = data[data["_reg"] == reg]

        # Encabezado de sección: Registro
        ws.set_row(r, 20)
        ws.merge_range(r, 1, r, 6, reg, fmt["section"])
        r += 1

        # Encabezados de columna
        for ci, h in enumerate(HDR, start=1):
            ws.write(r, ci, h, fmt["header"])
        r += 1

        reg_ing = 0.0
        reg_egr = 0.0

        for alt, (_, row) in enumerate(df_reg.iterrows()):
            f  = fmt["cell_alt"]  if alt % 2 else fmt["cell"]
            fm = fmt["money_alt"] if alt % 2 else fmt["money"]

            inm_v  = row["_inm"]
            año_v  = str(int(row["_año"])) if pd.notna(row.get("_año")) else ""
            m_idx  = row.get("_mes")
            mes_v  = MESES_C[int(m_idx)-1] if pd.notna(m_idx) and 1 <= int(m_idx) <= 12 else ""
            ing_v  = float(row["_ing_n"]) if row["_ing_n"] else None
            egr_v  = float(row["_egr_n"]) if row["_egr_n"] else None
            conc_v = str(row[conc_col] or "") if conc_col and pd.notna(row.get(conc_col)) else ""

            ws.write(r, 1, inm_v,                            f)
            ws.write(r, 2, año_v,                            f)
            ws.write(r, 3, mes_v,                            f)
            ws.write(r, 4, ing_v if ing_v is not None else "", fm)
            ws.write(r, 5, egr_v if egr_v is not None else "", fm)
            ws.write(r, 6, conc_v,                           f)

            reg_ing += row["_ing_n"]
            reg_egr += row["_egr_n"]
            r += 1

        # Subtotal del registro
        for ci in range(1, 7):
            ws.write(r, ci, "", fmt["total_lbl"])
        ws.write(r, 1, f"SUBTOTAL  {reg}",  fmt["total_lbl"])
        ws.write(r, 4, float(reg_ing) if reg_ing else "", fmt["money_tot"])
        ws.write(r, 5, float(reg_egr) if reg_egr else "", fmt["money_tot"])
        total_ing += reg_ing
        total_egr += reg_egr
        r += 2  # fila en blanco entre registros

    # Gran total
    for ci in range(1, 7):
        ws.write(r, ci, "", fmt["total_lbl"])
    ws.write(r, 1, "GRAN TOTAL",      fmt["total_lbl"])
    ws.write(r, 4, float(total_ing),  fmt["money_tot"])
    ws.write(r, 5, float(total_egr),  fmt["money_tot"])

    ws.freeze_panes(4, 0)
    print(f"  Hoja generada: 'Detalle Registro'  "
          f"({len(data)} registros, {len(registros)} tipos de registro)")


# ── Main ───────────────────────────────────────────────────────────────────────
def _verificar_archivo_cerrado():
    """Cancela el proceso si el archivo de salida esta abierto en Excel."""
    directorio = os.path.dirname(os.path.abspath(__file__))
    lock_file  = os.path.join(directorio, f'~${OUTPUT_FILE}')
    if os.path.exists(lock_file):
        print("\n" + "!" * 60)
        print(f"  AVISO: '{OUTPUT_FILE}' esta abierto en Excel.")
        print("  Cierra el archivo e intenta de nuevo.")
        print("!" * 60)
        sys.exit(1)


def main():
    _verificar_archivo_cerrado()

    print("=" * 60)
    print("  Dashboard Sucesion")
    print("=" * 60)

    print("\nDescargando hoja 'movimientos'...")
    df = download_movimientos()
    print(f"Datos: {len(df)} filas x {len(df.columns)} columnas")
    print(f"Columnas: {list(df.columns)}")

    df = prepare_año_mes(df)

    print("\nGenerando Excel...")
    wb  = xlsxwriter.Workbook(OUTPUT_FILE, {"nan_inf_to_errors": True})
    fmt = add_formats(wb)

    print("\n[1] Copiando hoja Movimientos...")
    write_movimientos(wb, fmt, df)

    print("\n[2] Generando Dashboard...")
    write_dashboard(wb, fmt, df)

    print("\n[3] Generando Informe Contable...")
    write_informe_contable(wb, fmt, df)

    print("\n[4] Generando Informe Dinámico (expandible)...")
    write_informe_dinamico(wb, fmt, df)

    for i, (tab_name, color, reg_val, title, use_ing, campo_fila, res_conc) in enumerate(CONTROL_TABS, 5):
        print(f"\n[{i}] Generando {tab_name} (Registro='{reg_val}', fila='{campo_fila}')...")
        write_control_tab(wb, fmt, df, tab_name, color, reg_val, title, use_ing, campo_fila, res_conc)

    print(f"\n[{len(CONTROL_TABS)+5}] Generando Detalle Inmuebles...")
    write_detalle_inmuebles(wb, fmt, df)

    wb.close()

    print(f"\n{'='*60}")
    print(f"  Archivo generado: {OUTPUT_FILE}")
    print(f"{'='*60}")

    import subprocess
    subprocess.Popen(["start", "", OUTPUT_FILE], shell=True)


if __name__ == "__main__":
    main()
