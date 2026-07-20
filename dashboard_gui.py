#!/usr/bin/env python3
"""
Panel de control (GUI) para el pipeline de Dashboard Sucesión.

Ejecuta en el orden correcto los pasos definidos en PROCESO.md:
  1. descargar_facturas_gmail.py (opcional — baja adjuntos nuevos de Gmail a facturas/)
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
import sys
import queue
import threading
import subprocess
import traceback
import tkinter as tk
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
CLIENT_SECRET  = os.path.join(BASE_DIR, 'client_secret.json')
CREDENTIALS    = os.path.join(BASE_DIR, 'credentials.json')

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
# Ejecución de un paso (corre dentro del proceso hijo)
# ─────────────────────────────────────────────────────────────
def run_step(key, reset_recibos=False):
    if key == 'facturas_gmail':
        import descargar_facturas_gmail
        descargar_facturas_gmail.main()
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
    os.chdir(BASE_DIR)
    try:
        run_step(key, reset_recibos=reset)
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

            if step['key'] == 'recibos':
                self.reset_recibos_var = tk.BooleanVar(value=False)
                ttk.Checkbutton(
                    row, text='Reiniciar numeración (borra historial)',
                    variable=self.reset_recibos_var,
                ).pack(side='left', padx=16)

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

        if 'facturas_gmail' in seleccionados and not os.path.exists(CLIENT_SECRET):
            errores.append(
                'Falta client_secret.json (credencial OAuth de escritorio de Google Cloud), '
                'necesario para "Descargar Facturas de Gmail".'
            )

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

        self._set_running(True, total_steps=len(seleccionados))
        self._clear_log()
        for step in STEPS:
            self.step_status_labels[step['key']].config(text='pendiente', foreground='gray')

        thread = threading.Thread(
            target=self._worker, args=(seleccionados, self.reset_recibos_var.get()),
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
    def _worker(self, seleccionados, reset_recibos):
        ok = True
        for key in seleccionados:
            self.msg_queue.put(('status', key, 'ejecutando…'))
            label = next(s['label'] for s in STEPS if s['key'] == key)
            self.msg_queue.put(('log', f'\n=== {label} ===\n', None))

            cmd = self._build_cmd(key, reset_recibos)
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

    def _build_cmd(self, key, reset_recibos):
        if getattr(sys, 'frozen', False):
            cmd = [sys.executable, '--run-step', key]
        else:
            cmd = [sys.executable, '-u', os.path.abspath(__file__), '--run-step', key]
        if key == 'recibos' and reset_recibos:
            cmd.append('--reset')
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
