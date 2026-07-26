#!/usr/bin/env python3
"""
Panel de control (GUI) para el pipeline de Dashboard Sucesión.

Ejecuta en el orden correcto los pasos definidos en PROCESO.md:
  1. descargar_facturas_gmail.py (opcional — baja a facturas/ los adjuntos de
     Gmail del rango de fechas indicado en la GUI)
  2. dashboard.py               (siempre primero — genera dashboard_sucesion.xlsx)
  3. organizar_facturas.py      (opcional — solo si hay CFDIs nuevos en facturas/)
  4. generar_recibos.py         (requiere el Excel del paso 2)
  5. reporte_word.py            (requiere el Excel del paso 2)
  6. reporte_word_propiedades.py (bitácora por inmueble; solo requiere credentials.json)

Cada paso se ejecuta en un proceso hijo independiente (este mismo script,
invocado con --run-step) para aislar sus sys.exit()/errores del proceso
de la GUI y poder capturar su salida en vivo.
"""

import os
import re
import sys
import json
import queue
import threading
import subprocess
import traceback
import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import ttk, messagebox

if getattr(sys, 'frozen', False):
    # PyInstaller onefile: __file__ apunta al bundle temporal, no al .exe real.
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH      = os.path.join(BASE_DIR, 'dashboard_sucesion.xlsx')
XLSX_LOCK      = os.path.join(BASE_DIR, '~$dashboard_sucesion.xlsx')
DOCX_PATH      = os.path.join(BASE_DIR, 'reporte_ejecutivo_sucesion.docx')
DOCX_PROP_PATH = os.path.join(BASE_DIR, 'reporte_propiedades.docx')
RECIBOS_DIR    = os.path.join(BASE_DIR, 'recibos')
RECIBOS_JSON   = os.path.join(BASE_DIR, 'recibos_registro.json')
FACTURAS_DIR   = os.path.join(BASE_DIR, 'facturas')
FACT_ORG_DIR   = os.path.join(BASE_DIR, 'facturas_organizadas')
SIN_XML_XLSX   = os.path.join(BASE_DIR, 'facturas_sin_xml.xlsx')
GMAIL_REPORTE_XLSX = os.path.join(BASE_DIR, 'gmail_facturas_reporte.xlsx')
GMAIL_REGISTRO_JSON = os.path.join(BASE_DIR, 'gmail_facturas_registro.json')
CLIENT_SECRET  = os.path.join(BASE_DIR, 'client_secret.json')
CREDENTIALS    = os.path.join(BASE_DIR, 'credentials.json')

# Formatos aceptados en los campos de fecha del rango de Gmail.
FORMATOS_FECHA = ('%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%d-%m-%Y')

# RFC propio: a facturas/ solo van los CFDI donde aparece como emisor o
# receptor; los de terceros se apartan en revision/otros_rfc/.
# Vacío = descargar todo sin filtrar.
RFC_PROPIO_DEFAULT = 'GALF730909CN0'
RE_RFC = re.compile(r'^[A-ZÑ&]{3,4}\d{6}[A-Z\d]{3}$')

STEPS = [
    {'key': 'facturas_gmail', 'label': 'Descargar Facturas de Gmail',   'default': False},
    {'key': 'dashboard',      'label': 'Actualizar Dashboard (Excel)',   'default': True},
    {'key': 'facturas',       'label': 'Organizar Facturas CFDI',        'default': False},
    {'key': 'facturas_sin_xml', 'label': 'Revisar Facturas sin XML',     'default': False},
    {'key': 'recibos',        'label': 'Generar Recibos de Renta',       'default': True},
    {'key': 'reporte',        'label': 'Reporte Ejecutivo (Word)',       'default': False},
    {'key': 'reporte_propiedades', 'label': 'Reporte de Bitácora por Inmueble (Word)', 'default': False},
]


# ─────────────────────────────────────────────────────────────
# Fechas del rango de descarga de Gmail
# ─────────────────────────────────────────────────────────────
def parsear_fecha(texto, etiqueta='fecha'):
    """Texto → date. Acepta AAAA-MM-DD, AAAA/MM/DD y DD/MM/AAAA."""
    texto = (texto or '').strip()
    if not texto:
        raise ValueError(f'Falta la {etiqueta}.')
    for fmt in FORMATOS_FECHA:
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    raise ValueError(f'{etiqueta[:1].upper()}{etiqueta[1:]} inválida: "{texto}". '
                     'Usa el formato AAAA-MM-DD.')


def parsear_rfcs(texto):
    """'rfc1, rfc2' → lista normalizada en mayúsculas. Vacío = sin filtro."""
    rfcs = []
    for parte in re.split(r'[,;\s]+', (texto or '').strip()):
        if not parte:
            continue
        rfc = parte.upper()
        if not RE_RFC.match(rfc):
            raise ValueError(f'RFC inválido: "{parte}". Un RFC de persona física tiene '
                             '13 caracteres (ej. GALF730909CN0) y uno moral 12.')
        rfcs.append(rfc)
    return rfcs


def fecha_ultima_descarga():
    """Fecha del correo más reciente ya procesado, para continuar donde quedó
    la corrida anterior. Si no hay registro, sugiere los últimos 30 días."""
    try:
        with open(GMAIL_REGISTRO_JSON, encoding='utf-8') as f:
            registro = json.load(f)
        fechas = [r['fecha_correo'][:10] for r in registro.values() if r.get('fecha_correo')]
        if fechas:
            return parsear_fecha(max(fechas))
    except (OSError, ValueError, json.JSONDecodeError, AttributeError, TypeError):
        pass
    return date.today() - timedelta(days=30)


# ─────────────────────────────────────────────────────────────
# Ejecución de un paso (corre dentro del proceso hijo)
# ─────────────────────────────────────────────────────────────
def run_step(key, reset_recibos=False, gmail_desde=None, gmail_hasta=None, gmail_rfc=None):
    if key == 'facturas_gmail':
        import descargar_facturas_gmail
        opciones = {'fecha_inicio': gmail_desde, 'fecha_fin': gmail_hasta}
        if gmail_rfc is not None:   # cadena vacía = descargar sin filtrar
            opciones['rfc_propio'] = gmail_rfc
        descargar_facturas_gmail.main(**opciones)
    elif key == 'dashboard':
        import dashboard
        dashboard.main()
    elif key == 'facturas':
        import organizar_facturas
        organizar_facturas.main()
    elif key == 'facturas_sin_xml':
        import organizar_facturas
        organizar_facturas.revisar_sin_xml()
    elif key == 'recibos':
        import generar_recibos
        generar_recibos.main(reiniciar_numeracion=reset_recibos)
    elif key == 'reporte':
        import reporte_word
        reporte_word.main()
    elif key == 'reporte_propiedades':
        import reporte_word_propiedades
        reporte_word_propiedades.main()
    else:
        raise ValueError(f'Paso desconocido: {key}')


def _arg_valor(flag):
    """Valor que sigue a un flag en sys.argv, o None si no viene."""
    if flag in sys.argv:
        idx = sys.argv.index(flag)
        if idx + 1 < len(sys.argv):
            return sys.argv[idx + 1]
    return None


def _cli_worker_entry():
    """Si se invocó como --run-step <key>, ejecuta ese paso y termina el proceso
    (sys.exit) sin construir la GUI. Si no se pasó ese flag, no hace nada."""
    if '--check-paths' in sys.argv:
        print(f'frozen        = {getattr(sys, "frozen", False)}')
        print(f'BASE_DIR      = {BASE_DIR}')
        print(f'credentials   = {CREDENTIALS} -> existe: {os.path.exists(CREDENTIALS)}')
        print(f'xlsx dashboard= {XLSX_PATH} -> existe: {os.path.exists(XLSX_PATH)}')
        sys.exit(0)
    if '--run-step' not in sys.argv:
        return
    idx = sys.argv.index('--run-step')
    key = sys.argv[idx + 1]
    reset = '--reset' in sys.argv
    desde = _arg_valor('--desde')
    hasta = _arg_valor('--hasta')
    rfc   = _arg_valor('--rfc')
    os.chdir(BASE_DIR)
    try:
        run_step(key, reset_recibos=reset, gmail_desde=desde, gmail_hasta=hasta,
                 gmail_rfc=rfc)
    except SystemExit as e:
        code = e.code if isinstance(e.code, int) else (1 if e.code else 0)
        sys.exit(code)
    except Exception:
        traceback.print_exc()
        sys.exit(1)
    sys.exit(0)


# ─────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Dashboard Sucesión — Panel de Control')
        self.geometry('960x760')
        self.minsize(920, 700)

        self.msg_queue = queue.Queue()
        self.running = False
        self.current_proc = None
        self.step_vars = {}
        self.step_status_labels = {}

        self._build_ui()
        self._refresh_output_buttons()

    # ---------- construcción de la interfaz ----------
    def _build_ui(self):
        pad = {'padx': 10, 'pady': 6}

        header = ttk.Label(
            self, text='Selecciona los pasos a ejecutar (se corren en orden 1→4):',
            font=('Segoe UI', 10, 'bold'))
        header.pack(anchor='w', **pad)

        steps_frame = ttk.Frame(self)
        steps_frame.pack(fill='x', **pad)

        for i, step in enumerate(STEPS, start=1):
            var = tk.BooleanVar(value=step['default'])
            self.step_vars[step['key']] = var

            row = ttk.Frame(steps_frame)
            row.pack(fill='x', pady=2)

            cb = ttk.Checkbutton(row, text=f"{i}. {step['label']}", variable=var)
            cb.pack(side='left')

            status = ttk.Label(row, text='pendiente', foreground='gray')
            status.pack(side='right', padx=6)
            self.step_status_labels[step['key']] = status

            if step['key'] == 'facturas_gmail':
                self._build_rango_gmail(row)
                self._build_rfc_gmail(steps_frame)

            if step['key'] == 'recibos':
                self.reset_recibos_var = tk.BooleanVar(value=False)
                ttk.Checkbutton(
                    row, text='Reiniciar numeración (borra historial)',
                    variable=self.reset_recibos_var,
                ).pack(side='left', padx=16)

        self.step_vars['facturas_gmail'].trace_add(
            'write', lambda *_: self._sync_rango_gmail())
        self._sync_rango_gmail()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill='x', **pad)

        self.run_btn = ttk.Button(btn_frame, text='▶ Ejecutar', command=self._on_run)
        self.run_btn.pack(side='left')

        self.progress = ttk.Progressbar(btn_frame, mode='determinate', maximum=len(STEPS))
        self.progress.pack(side='left', fill='x', expand=True, padx=10)

        ttk.Label(self, text='Registro de ejecución:').pack(anchor='w', padx=10)

        log_frame = ttk.Frame(self)
        log_frame.pack(fill='both', expand=True, padx=10, pady=(0, 6))
        self.log_text = tk.Text(log_frame, wrap='word', state='disabled',
                                 bg='#111', fg='#ddd', font=('Consolas', 9))
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text['yscrollcommand'] = scroll.set
        self.log_text.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        outputs_frame = ttk.LabelFrame(self, text='Abrir resultados')
        outputs_frame.pack(fill='x', padx=10, pady=(0, 10))

        self.btn_excel = ttk.Button(outputs_frame, text='Excel (dashboard)',
                                     command=lambda: self._abrir(XLSX_PATH))
        self.btn_word = ttk.Button(outputs_frame, text='Word (reporte)',
                                    command=lambda: self._abrir(DOCX_PATH))
        self.btn_word_prop = ttk.Button(outputs_frame, text='Word (bitácora inmuebles)',
                                         command=lambda: self._abrir(DOCX_PROP_PATH))
        self.btn_recibos = ttk.Button(outputs_frame, text='Carpeta de recibos',
                                       command=lambda: self._abrir(RECIBOS_DIR))
        self.btn_facturas = ttk.Button(outputs_frame, text='Facturas organizadas',
                                        command=lambda: self._abrir(FACT_ORG_DIR))
        self.btn_sin_xml = ttk.Button(outputs_frame, text='Facturas sin XML',
                                       command=lambda: self._abrir(SIN_XML_XLSX))
        self.btn_gmail = ttk.Button(outputs_frame, text='Reporte descarga Gmail',
                                     command=lambda: self._abrir(GMAIL_REPORTE_XLSX))
        for b in (self.btn_excel, self.btn_word, self.btn_word_prop, self.btn_recibos,
                  self.btn_facturas, self.btn_sin_xml, self.btn_gmail):
            b.pack(side='left', padx=6, pady=6)

    def _build_rango_gmail(self, row):
        """Campos del rango de fechas (inclusive) para la descarga de Gmail.

        Por defecto propone continuar desde el correo más reciente ya
        descargado hasta hoy.
        """
        self.gmail_desde_var = tk.StringVar(value=fecha_ultima_descarga().isoformat())
        self.gmail_hasta_var = tk.StringVar(value=date.today().isoformat())

        ttk.Label(row, text='Rango:').pack(side='left', padx=(18, 4))
        self.gmail_desde_entry = ttk.Entry(
            row, textvariable=self.gmail_desde_var, width=11, justify='center')
        self.gmail_desde_entry.pack(side='left')
        ttk.Label(row, text='→').pack(side='left', padx=4)
        self.gmail_hasta_entry = ttk.Entry(
            row, textvariable=self.gmail_hasta_var, width=11, justify='center')
        self.gmail_hasta_entry.pack(side='left')
        ttk.Label(row, text='AAAA-MM-DD (fin vacío = sin límite)',
                  foreground='gray').pack(side='left', padx=8)

    def _build_rfc_gmail(self, steps_frame):
        """Segunda línea del paso 1: RFC propio para separar las facturas de
        terceros de las que sí son tuyas (las emitas o las recibas)."""
        sub = ttk.Frame(steps_frame)
        sub.pack(fill='x', pady=(0, 2))

        self.gmail_rfc_var = tk.StringVar(value=RFC_PROPIO_DEFAULT)
        ttk.Label(sub, text='Mi RFC:').pack(side='left', padx=(24, 4))
        self.gmail_rfc_entry = ttk.Entry(sub, textvariable=self.gmail_rfc_var, width=32)
        self.gmail_rfc_entry.pack(side='left')
        ttk.Label(sub, text='como emisor o receptor · varios con coma · vacío = todas',
                  foreground='gray').pack(side='left', padx=8)

    def _sync_rango_gmail(self):
        """Los campos del paso 1 solo se editan si ese paso está marcado."""
        estado = 'normal' if self.step_vars['facturas_gmail'].get() else 'disabled'
        self.gmail_desde_entry.config(state=estado)
        self.gmail_hasta_entry.config(state=estado)
        self.gmail_rfc_entry.config(state=estado)

    def _opciones_gmail(self):
        """Rango y RFC ya validados para el paso de Gmail.

        'hasta' es None si no hay límite superior; 'rfc' es una cadena vacía
        cuando el usuario decide no filtrar. Lanza ValueError si algo no sirve.
        """
        desde = parsear_fecha(self.gmail_desde_var.get(),
                              'fecha inicial del rango de Gmail')
        texto_hasta = self.gmail_hasta_var.get().strip()
        hasta = parsear_fecha(texto_hasta, 'fecha final del rango de Gmail') if texto_hasta else None
        if hasta and hasta < desde:
            raise ValueError(
                'El rango de fechas de Gmail está invertido: la fecha final '
                f'({hasta:%Y-%m-%d}) es anterior a la inicial ({desde:%Y-%m-%d}).'
            )
        return {
            'desde': desde.isoformat(),
            'hasta': hasta.isoformat() if hasta else None,
            'rfc': ','.join(parsear_rfcs(self.gmail_rfc_var.get())),
        }

    def _abrir(self, path):
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showinfo('No disponible', f'Aún no existe:\n{path}')

    def _refresh_output_buttons(self):
        self.btn_excel.state(['!disabled'] if os.path.exists(XLSX_PATH) else ['disabled'])
        self.btn_word.state(['!disabled'] if os.path.exists(DOCX_PATH) else ['disabled'])
        self.btn_word_prop.state(['!disabled'] if os.path.exists(DOCX_PROP_PATH) else ['disabled'])
        self.btn_recibos.state(['!disabled'] if os.path.isdir(RECIBOS_DIR) else ['disabled'])
        self.btn_facturas.state(['!disabled'] if os.path.isdir(FACT_ORG_DIR) else ['disabled'])
        self.btn_sin_xml.state(['!disabled'] if os.path.exists(SIN_XML_XLSX) else ['disabled'])
        self.btn_gmail.state(['!disabled'] if os.path.exists(GMAIL_REPORTE_XLSX) else ['disabled'])

    # ---------- validaciones previas ----------
    def _validar_prerrequisitos(self, seleccionados):
        errores = []

        if 'facturas_gmail' in seleccionados:
            if not os.path.exists(CLIENT_SECRET):
                errores.append(
                    'Falta client_secret.json (credencial OAuth de escritorio de Google Cloud), '
                    'necesario para "Descargar Facturas de Gmail".'
                )
            try:
                self._opciones_gmail()
            except ValueError as e:
                errores.append(str(e))

        necesita_credenciales = ('dashboard' in seleccionados or 'reporte' in seleccionados
                                  or 'reporte_propiedades' in seleccionados)
        if necesita_credenciales and not os.path.exists(CREDENTIALS):
            errores.append(
                'Falta credentials.json (Service Account de Google Cloud), '
                'necesario para actualizar el Dashboard o el Reporte Word.'
            )

        if 'dashboard' in seleccionados and os.path.exists(XLSX_LOCK):
            errores.append(
                'dashboard_sucesion.xlsx parece estar abierto en Excel. '
                'Cierra el archivo antes de actualizar el Dashboard.'
            )

        necesita_excel = ('recibos' in seleccionados or 'reporte' in seleccionados)
        if necesita_excel and 'dashboard' not in seleccionados and not os.path.exists(XLSX_PATH):
            errores.append(
                'No existe dashboard_sucesion.xlsx. Marca "Actualizar Dashboard" '
                'o genera el Excel antes de este paso.'
            )

        if 'facturas' in seleccionados and not os.path.isdir(FACTURAS_DIR):
            errores.append(
                'No existe la carpeta facturas/. Se omitirá "Organizar Facturas CFDI".'
            )

        if 'facturas_sin_xml' in seleccionados and not os.path.isdir(FACT_ORG_DIR):
            errores.append(
                'No existe la carpeta facturas_organizadas/. '
                'Se omitirá "Revisar Facturas sin XML".'
            )

        return errores

    # ---------- lanzar ejecución ----------
    def _on_run(self):
        if self.running:
            return

        seleccionados = [s['key'] for s in STEPS if self.step_vars[s['key']].get()]
        if not seleccionados:
            messagebox.showwarning('Nada seleccionado', 'Selecciona al menos un paso.')
            return

        errores = self._validar_prerrequisitos(seleccionados)
        # La falta de facturas/ solo es motivo de omisión, no de bloqueo total.
        bloqueantes = [e for e in errores if 'Se omitirá' not in e]
        if bloqueantes:
            messagebox.showerror('No se puede continuar', '\n\n'.join(bloqueantes))
            return
        if 'facturas' in seleccionados and not os.path.isdir(FACTURAS_DIR):
            seleccionados.remove('facturas')
            self._log('Se omite "Organizar Facturas CFDI": no existe la carpeta facturas/.\n')

        if 'facturas_sin_xml' in seleccionados and not os.path.isdir(FACT_ORG_DIR):
            seleccionados.remove('facturas_sin_xml')
            self._log('Se omite "Revisar Facturas sin XML": no existe la carpeta '
                       'facturas_organizadas/.\n')

        if self.reset_recibos_var.get() and 'recibos' in seleccionados:
            if not messagebox.askyesno(
                'Confirmar reinicio de numeración',
                'Esto borrará recibos_registro.json y reasignará los folios desde 1.\n'
                'Los PDFs ya generados NO se eliminan, pero se perderá el historial '
                'de qué recibos ya se emitieron.\n\n¿Continuar?'
            ):
                return

        # Las opciones se leen aquí (hilo de la GUI); tkinter no es thread-safe.
        gmail = self._opciones_gmail() if 'facturas_gmail' in seleccionados else None

        self._set_running(True, total_steps=len(seleccionados))
        self._clear_log()
        for step in STEPS:
            self.step_status_labels[step['key']].config(text='pendiente', foreground='gray')

        thread = threading.Thread(
            target=self._worker,
            args=(seleccionados, self.reset_recibos_var.get(), gmail),
            daemon=True,
        )
        thread.start()
        self.after(100, self._poll_queue)

    def _set_running(self, running, total_steps=None):
        self.running = running
        self.run_btn.config(state='disabled' if running else 'normal')
        self.progress['value'] = 0
        if total_steps is not None:
            self.progress['maximum'] = total_steps

    # ---------- worker (hilo secundario) ----------
    def _worker(self, seleccionados, reset_recibos, gmail=None):
        ok = True
        for key in seleccionados:
            self.msg_queue.put(('status', key, 'ejecutando…'))
            label = next(s['label'] for s in STEPS if s['key'] == key)
            self.msg_queue.put(('log', f'\n=== {label} ===\n', None))

            cmd = self._build_cmd(key, reset_recibos, gmail)
            try:
                creationflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                proc = subprocess.Popen(
                    cmd, cwd=BASE_DIR,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, bufsize=1, creationflags=creationflags,
                )
                self.current_proc = proc
                for line in proc.stdout:
                    self.msg_queue.put(('log', line, None))
                proc.wait()
            except Exception as e:
                self.msg_queue.put(('log', f'ERROR al lanzar el proceso: {e}\n', None))
                proc = None

            self.current_proc = None
            returncode = proc.returncode if proc else 1

            if returncode == 0:
                self.msg_queue.put(('status', key, 'completado'))
            else:
                self.msg_queue.put(('status', key, f'error (código {returncode})'))
                ok = False
                break

            self.msg_queue.put(('progress', None, None))

        self.msg_queue.put(('finished', ok, None))

    def _build_cmd(self, key, reset_recibos, gmail=None):
        if getattr(sys, 'frozen', False):
            cmd = [sys.executable, '--run-step', key]
        else:
            cmd = [sys.executable, '-u', os.path.abspath(__file__), '--run-step', key]
        if key == 'recibos' and reset_recibos:
            cmd.append('--reset')
        if key == 'facturas_gmail' and gmail:
            if gmail['desde']:
                cmd += ['--desde', gmail['desde']]
            if gmail['hasta']:
                cmd += ['--hasta', gmail['hasta']]
            # Siempre se manda: la cadena vacía significa "sin filtro de RFC".
            cmd += ['--rfc', gmail['rfc']]
        return cmd

    # ---------- cola de mensajes → UI ----------
    def _poll_queue(self):
        while True:
            try:
                kind, a, b = self.msg_queue.get_nowait()
            except queue.Empty:
                break
            try:
                if kind == 'log':
                    self._log(a)
                elif kind == 'status':
                    color = {'ejecutando…': '#0a7', 'completado': '#0a0'}.get(b, '#b00')
                    self.step_status_labels[a].config(text=b, foreground=color)
                elif kind == 'progress':
                    self.progress['value'] += 1
                elif kind == 'finished':
                    self._on_finished(ok=a)
            except Exception as e:
                self._log(f'[GUI] Error procesando mensaje interno: {e}\n')

        if self.running:
            self.after(100, self._poll_queue)

    def _on_finished(self, ok):
        self._set_running(False)
        self._refresh_output_buttons()
        if ok:
            messagebox.showinfo('Listo', 'Todos los pasos seleccionados terminaron correctamente.')
        else:
            messagebox.showerror('Se detuvo con errores',
                                  'Un paso falló. Revisa el registro de ejecución para más detalle.')

    def _log(self, text):
        self.log_text.config(state='normal')
        self.log_text.insert('end', text if text.endswith('\n') else text + '\n')
        self.log_text.see('end')
        self.log_text.config(state='disabled')

    def _clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.config(state='disabled')


def main():
    _cli_worker_entry()  # termina el proceso aquí si es un hijo --run-step
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
