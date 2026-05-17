"""
Comparador VBA — interfaz gráfica
"""

import sys
import os
import winreg
import threading
import queue
import pythoncom
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from compare import extract_vba, compare as vba_compare, render_excel
from apply import read_diff_excel, apply_vba


# ── Registro de Windows ───────────────────────────────────────────────────────

REG_KEY    = r"Software\ComparadorVBA"
REG_FIELDS = ("excel_a", "excel_b", "diff_out", "diff_in", "dest1", "dest2")

def _load_config() -> dict:
    cfg = {}
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY)
        for f in REG_FIELDS:
            try:
                v, _ = winreg.QueryValueEx(key, f)
                cfg[f] = v
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except FileNotFoundError:
        pass
    return cfg

def _save_config(data: dict) -> None:
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY)
        for f in REG_FIELDS:
            winreg.SetValueEx(key, f, 0, winreg.REG_SZ, data.get(f, ""))
        winreg.CloseKey(key)
    except Exception:
        pass


# ── QueueWriter ───────────────────────────────────────────────────────────────

class QueueWriter:
    def __init__(self, q):
        self.q = q
    def write(self, text):
        if text:
            self.q.put(("normal", text))
    def flush(self):
        pass


# ── Constantes de estilo ──────────────────────────────────────────────────────

FONT_UI    = ("Verdana", 11)
FONT_BOLD  = ("Verdana", 11, "bold")
FONT_TITLE = ("Verdana", 12, "bold")
FONT_LOG   = ("Courier New", 11)

BG_APP      = "#F2F3F4"
BG_COMPARAR = "#2E86C1"
BG_APLICAR  = "#1E8449"
BG_REGISTRO = "#5D6D7E"

C_GENERAR = "#FAD7A0"
C_SIMULAR = "#AED6F1"
C_APLICAR = "#A9DFBF"
C_ABRIR   = "#D7BDE2"
C_SELEC   = "#D5D8DC"
C_ACCION  = "#F0B27A"   # naranja — Borrar datos + Limpiar registro


# ── Aplicación ────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Comparador VBA")
        self.resizable(True, True)
        self.configure(bg=BG_APP)

        self._log_queue: queue.Queue = queue.Queue()
        self._cfg = _load_config()
        self._diff_path: str = ""          # ruta del último diff generado

        self._build_styles()
        self._build_ui()
        self._cargar_config()
        self._centrar_ventana(1600, 870)
        self._poll_log()

    # ── Centrado ──────────────────────────────────────────────────────────────

    def _centrar_ventana(self, w: int, h: int):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = max(0, (self.winfo_screenheight() - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(1200, 700)

    # ── Estilos ───────────────────────────────────────────────────────────────

    def _build_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",               font=FONT_UI, background=BG_APP)
        s.configure("TFrame",          background=BG_APP)
        s.configure("TEntry",          font=FONT_UI, fieldbackground="white")

        # Seleccionar
        s.configure("Sel.TButton",     font=FONT_UI, background=C_SELEC)
        s.map("Sel.TButton",           background=[("active", "#BFC9CA")])

        # Borrar datos / Limpiar registro — naranja, mismo tamaño
        s.configure("Accion.TButton",  font=FONT_UI, background=C_ACCION)
        s.map("Accion.TButton",        background=[("active", "#E59866")])

        # Pasteles de acción
        s.configure("Generar.TButton", font=FONT_BOLD,  background=C_GENERAR)
        s.map("Generar.TButton",       background=[("active", "#F5CBA7")])
        s.configure("Simular.TButton", font=FONT_UI,    background=C_SIMULAR)
        s.map("Simular.TButton",       background=[("active", "#85C1E9")])
        s.configure("AppBtn.TButton",  font=FONT_BOLD,  background=C_APLICAR)
        s.map("AppBtn.TButton",        background=[("active", "#7DCEA0")])
        s.configure("Abrir.TButton",   font=FONT_UI,    background=C_ABRIR)
        s.map("Abrir.TButton",         background=[("active", "#C39BD3")])

    # ── Bloque con título coloreado ───────────────────────────────────────────

    def _bloque(self, parent, titulo: str, bg_titulo: str,
                row: int, pady_top: int = 6) -> tuple[tk.Frame, tk.Frame]:
        """Devuelve (title_bar, content). El caller puede añadir widgets al title_bar."""
        outer = tk.Frame(parent, bg=BG_APP, bd=1, relief="solid",
                         highlightbackground="#AAAAAA", highlightthickness=1)
        outer.grid(row=row, column=0, padx=12, pady=(pady_top, 0), sticky="ew")
        outer.columnconfigure(0, weight=1)

        title_bar = tk.Frame(outer, bg=bg_titulo)
        title_bar.pack(fill=tk.X)
        title_bar.columnconfigure(0, weight=1)   # col 0 = título (crece)
                                                  # col 1+ = botones (fijos)
        tk.Label(title_bar, text=f"  {titulo}",
                 bg=bg_titulo, fg="white",
                 font=FONT_TITLE, anchor="w", pady=5
                 ).grid(row=0, column=0, sticky="ew")

        content = tk.Frame(outer, bg=BG_APP, padx=12, pady=8)
        content.pack(fill=tk.BOTH, expand=True)
        content.columnconfigure(1, weight=1)
        return title_bar, content

    # ── Interfaz ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # ══ Bloque COMPARAR ══
        tb_cmp, cmp = self._bloque(self, "COMPARAR", BG_COMPARAR, row=0, pady_top=12)

        # Borrar datos en la barra de título
        ttk.Button(tb_cmp, text="Borrar datos", style="Accion.TButton",
                   command=self._borrar_comparar
                   ).grid(row=0, column=1, padx=8, pady=4, ipadx=6)

        self.var_a   = tk.StringVar()
        self.var_b   = tk.StringVar()
        self.var_out = tk.StringVar()

        self._fila_abrir(cmp, 0, "Excel 1:",         self.var_a,
                         [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")])
        self._fila_abrir(cmp, 1, "Excel 2:",         self.var_b,
                         [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")])
        self._fila_guardar(cmp, 2, "Guardar diff en:", self.var_out,
                           [("Excel", "*.xlsx")])

        # Fila de botones + resultado
        frm_cb = tk.Frame(cmp, bg=BG_APP)
        frm_cb.grid(row=3, column=0, columnspan=3, pady=(10, 4), sticky="ew")
        frm_cb.columnconfigure(1, weight=1)

        # Izquierda: solo botón GENERAR (Abrir aparece dinámicamente a su derecha)
        frm_cb_izq = tk.Frame(frm_cb, bg=BG_APP)
        frm_cb_izq.grid(row=0, column=0, sticky="w")

        ttk.Button(frm_cb_izq, text="GENERAR VBA_Diff.xlsx",
                   style="Generar.TButton",
                   command=self._accion_comparar).pack(
            side=tk.LEFT, ipadx=12, ipady=4)

        # Derecha: resultado dinámico (stats + Abrir/Idénticos) — oculto hasta GENERAR
        frm_res = tk.Frame(frm_cb, bg=BG_APP)
        frm_res.grid(row=0, column=1, sticky="w", padx=(20, 0))

        self._lbl_stats = tk.Label(frm_res, text="", bg=BG_APP, font=FONT_UI,
                                   fg="#2C3E50")

        self._lbl_identicos = tk.Label(frm_res, text="✓  IDÉNTICOS",
                                       bg=BG_APP, fg="#1E8449", font=FONT_BOLD)

        # Botón "Abrir [filename]" — se muestra siempre tras GENERAR
        self._btn_abrir_res = ttk.Button(frm_res, text="", style="Abrir.TButton",
                                          command=self._abrir_diff_resultado)

        # ══ Bloque APLICAR ══
        tb_app, app = self._bloque(self, "APLICAR", BG_APLICAR, row=1)

        ttk.Button(tb_app, text="Borrar datos", style="Accion.TButton",
                   command=self._borrar_aplicar
                   ).grid(row=0, column=1, padx=8, pady=4, ipadx=6)

        self.var_diff  = tk.StringVar()
        self.var_dest1 = tk.StringVar()
        self.var_dest2 = tk.StringVar()

        self._fila_abrir(app, 0, "VBA_Diff.xlsx:", self.var_diff,
                         [("Excel", "*.xlsx"), ("Todos", "*.*")])
        self._fila_abrir(app, 1, "Destino 1:",    self.var_dest1,
                         [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")])
        self._fila_abrir(app, 2, "Destino 2:",    self.var_dest2,
                         [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")])

        frm_ab = tk.Frame(app, bg=BG_APP)
        frm_ab.grid(row=3, column=0, columnspan=3, pady=(10, 4), sticky="w")

        ttk.Button(frm_ab, text="Simular → Destino 1",
                   style="Simular.TButton",
                   command=lambda: self._accion_aplicar(self.var_dest1, dry_run=True)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab, text="Aplicar → Destino 1",
                   style="AppBtn.TButton",
                   command=lambda: self._accion_aplicar(self.var_dest1, dry_run=False)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab, text="Simular → Destino 2",
                   style="Simular.TButton",
                   command=lambda: self._accion_aplicar(self.var_dest2, dry_run=True)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab, text="Aplicar → Destino 2",
                   style="AppBtn.TButton",
                   command=lambda: self._accion_aplicar(self.var_dest2, dry_run=False)
                   ).pack(side=tk.LEFT, ipadx=6, ipady=3)

        # ══ Bloque REGISTRO ══
        reg_outer = tk.Frame(self, bg=BG_APP, bd=1, relief="solid",
                             highlightbackground="#AAAAAA", highlightthickness=1)
        reg_outer.grid(row=2, column=0, padx=12, pady=(6, 12), sticky="nsew")
        reg_outer.columnconfigure(0, weight=1)
        reg_outer.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)

        tb_reg = tk.Frame(reg_outer, bg=BG_REGISTRO)
        tb_reg.grid(row=0, column=0, sticky="ew")
        tb_reg.columnconfigure(0, weight=1)

        tk.Label(tb_reg, text="  REGISTRO",
                 bg=BG_REGISTRO, fg="white",
                 font=FONT_TITLE, anchor="w", pady=5
                 ).grid(row=0, column=0, sticky="ew")
        ttk.Button(tb_reg, text="Limpiar registro", style="Accion.TButton",
                   command=self._limpiar_log
                   ).grid(row=0, column=1, padx=8, pady=4, ipadx=6)

        reg_content = tk.Frame(reg_outer, bg=BG_APP, padx=6, pady=6)
        reg_content.grid(row=1, column=0, sticky="nsew")
        reg_content.columnconfigure(0, weight=1)
        reg_content.rowconfigure(0, weight=1)

        self.log_widget = scrolledtext.ScrolledText(
            reg_content, height=20, state=tk.NORMAL,
            font=FONT_LOG, wrap=tk.WORD,
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="white")
        self.log_widget.grid(row=0, column=0, sticky="nsew")
        self.log_widget.tag_configure("ok",        foreground="#4EC94E")
        self.log_widget.tag_configure("error",     foreground="#FF6B6B")
        # Diff: fondo completo en cada línea para distinguir incluso líneas en blanco
        self.log_widget.tag_configure("diff_add",
            foreground="#AFFFAF", background="#1A4A1A")   # + verde oscuro bg
        self.log_widget.tag_configure("diff_del",
            foreground="#FFAFAF", background="#4A1A1A")   # - rojo oscuro bg
        self.log_widget.tag_configure("diff_hunk",
            foreground="#F8C471", background="#2C2C00")   # @@ amarillo
        # Cabecera de módulo: spacing1/spacing3 = ancho COMPLETO garantizado por Tk
        self.log_widget.tag_configure("diff_head",
            foreground="white", background="#1A5276",
            font=("Courier New", 12, "bold"),
            spacing1=10, spacing3=10)

        for var in (self.var_a, self.var_b, self.var_out,
                    self.var_diff, self.var_dest1, self.var_dest2):
            var.trace_add("write", lambda *_: self._guardar_config())

        # Limpiar resultado al cambiar cualquiera de los dos Excel de origen
        self.var_a.trace_add("write", lambda *_: self._limpiar_resultado())
        self.var_b.trace_add("write", lambda *_: self._limpiar_resultado())

    # ── Helpers filas ─────────────────────────────────────────────────────────

    def _fila_abrir(self, parent, row, label, var, filetypes):
        tk.Label(parent, text=label, bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(parent, textvariable=var, font=FONT_UI).grid(
            row=row, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_archivo(var, filetypes)
                   ).grid(row=row, column=2, padx=(8, 0), pady=4, ipadx=4)

    def _fila_guardar(self, parent, row, label, var, filetypes):
        tk.Label(parent, text=label, bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(parent, textvariable=var, font=FONT_UI).grid(
            row=row, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._guardar_archivo(var, filetypes)
                   ).grid(row=row, column=2, padx=(8, 0), pady=4, ipadx=4)

    def _seleccionar_archivo(self, var, filetypes):
        inicial = str(Path(var.get()).parent) if var.get() else "/"
        ruta = filedialog.askopenfilename(initialdir=inicial, filetypes=filetypes)
        if ruta:
            var.set(ruta)

    def _guardar_archivo(self, var, filetypes):
        inicial = str(Path(var.get()).parent) if var.get() else "/"
        ruta = filedialog.asksaveasfilename(
            initialdir=inicial, defaultextension=".xlsx", filetypes=filetypes)
        if ruta:
            var.set(ruta)

    # ── Resultado comparación ─────────────────────────────────────────────────

    def _mostrar_resultado(self, total: int, con_dif: int, diff_path: str):
        """Actualiza el área de resultado (llamado desde el hilo principal)."""
        self._lbl_stats.config(
            text=f"{total} módulos  ·  {con_dif} con diferencias")
        self._lbl_stats.pack(side=tk.LEFT, padx=(0, 18))

        self._lbl_identicos.pack_forget()
        self._btn_abrir_res.pack_forget()

        fname = Path(diff_path).name
        self._btn_abrir_res.config(text=f"Abrir  {fname}")
        self._btn_abrir_res.pack(side=tk.LEFT, ipadx=6, ipady=2, padx=(0, 10))

        if con_dif == 0:
            self._lbl_identicos.pack(side=tk.LEFT)

    def _limpiar_resultado(self):
        self._lbl_stats.config(text="")
        self._lbl_stats.pack_forget()
        self._lbl_identicos.pack_forget()
        self._btn_abrir_res.pack_forget()

    # ── Config ────────────────────────────────────────────────────────────────

    def _cargar_config(self):
        self.var_a.set(self._cfg.get("excel_a", ""))
        self.var_b.set(self._cfg.get("excel_b", ""))
        self.var_out.set(self._cfg.get("diff_out", ""))
        self.var_diff.set(self._cfg.get("diff_in", ""))
        self.var_dest1.set(self._cfg.get("dest1", ""))
        self.var_dest2.set(self._cfg.get("dest2", ""))

    def _guardar_config(self):
        _save_config({
            "excel_a":  self.var_a.get(),
            "excel_b":  self.var_b.get(),
            "diff_out": self.var_out.get(),
            "diff_in":  self.var_diff.get(),
            "dest1":    self.var_dest1.get(),
            "dest2":    self.var_dest2.get(),
        })

    def _borrar_comparar(self):
        for v in (self.var_a, self.var_b, self.var_out):
            v.set("")
        self._limpiar_resultado()

    def _borrar_aplicar(self):
        for v in (self.var_diff, self.var_dest1, self.var_dest2):
            v.set("")

    def _abrir_diff_resultado(self):
        if self._diff_path and os.path.exists(self._diff_path):
            os.startfile(self._diff_path)

    # ── Log ───────────────────────────────────────────────────────────────────

    def _poll_log(self):
        while not self._log_queue.empty():
            tag, texto = self._log_queue.get_nowait()
            if tag in ("ok", "error"):
                self.log_widget.insert(tk.END, texto, tag)
            else:
                # Quitar \n iniciales para detectar correctamente el prefijo
                t = texto.lstrip("\n")
                n_nl = len(texto) - len(t)   # nº de \n iniciales

                if t.startswith("+") and not t.startswith("+++ "):
                    self.log_widget.insert(tk.END, texto, "diff_add")
                elif t.startswith("-") and not t.startswith("--- "):
                    self.log_widget.insert(tk.END, texto, "diff_del")
                elif t.startswith("@@"):
                    self.log_widget.insert(tk.END, texto, "diff_hunk")
                elif t.startswith("──"):
                    # Cabecera de módulo: insertar \n iniciales sin color,
                    # luego el texto con fondo azul (spacing1/3 hacen el ancho completo)
                    if n_nl:
                        self.log_widget.insert(tk.END, "\n" * n_nl)
                    self.log_widget.insert(tk.END, t, "diff_head")
                else:
                    self.log_widget.insert(tk.END, texto)
            self.log_widget.see(tk.END)
        self.after(100, self._poll_log)

    def _limpiar_log(self):
        self.log_widget.delete("1.0", tk.END)

    def _log_final(self, msg: str, ok: bool):
        self._log_queue.put(("ok" if ok else "error", "\n" + msg + "\n"))

    # ── Acciones ──────────────────────────────────────────────────────────────

    def _accion_comparar(self):
        a, b, out = (self.var_a.get().strip(),
                     self.var_b.get().strip(),
                     self.var_out.get().strip())
        if not a:   self._log_final("ERROR: Selecciona el Excel 1.", ok=False); return
        if not b:   self._log_final("ERROR: Selecciona el Excel 2.", ok=False); return
        if not out: self._log_final("ERROR: Indica dónde guardar el diff.", ok=False); return
        self._limpiar_log()
        self._limpiar_resultado()
        threading.Thread(
            target=self._thread_comparar, args=(a, b, out), daemon=True).start()

    def _thread_comparar(self, a, b, out):
        pythoncom.CoInitialize()
        writer = QueueWriter(self._log_queue)
        old_stdout = sys.stdout
        sys.stdout = writer
        total = con_dif = 0
        try:
            la, lb = Path(a).name, Path(b).name
            print(f"\nExtrayendo VBA de: {la}")
            mods_a = extract_vba(a)
            print(f"  -> {len(mods_a)} módulos encontrados")
            print(f"Extrayendo VBA de: {lb}")
            mods_b = extract_vba(b)
            print(f"  -> {len(mods_b)} módulos encontrados\n")
            diffs  = vba_compare(mods_a, mods_b)
            render_excel(diffs, la, lb, out)
            total   = len(diffs)
            con_dif = sum(1 for d in diffs if d.status != "equal")
            print(f"Módulos con diferencias: {con_dif} de {total} totales")
        except Exception as e:
            sys.stdout = old_stdout
            self._log_final(f"ERROR: {e}", ok=False)
            return
        finally:
            sys.stdout = old_stdout
            pythoncom.CoUninitialize()

        self._diff_path = out
        self.var_diff.set(out)
        # Actualizar resultado en el hilo principal
        self.after(0, lambda: self._mostrar_resultado(total, con_dif, out))
        self._log_final(
            f"Diff guardado: {Path(out).name}  —  pulsa 'Abrir diff'" if con_dif > 0
            else f"Los ficheros son idénticos. Diff guardado: {Path(out).name}",
            ok=True)

    def _accion_aplicar(self, dest_var, dry_run):
        diff, dest = self.var_diff.get().strip(), dest_var.get().strip()
        if not diff: self._log_final("ERROR: Selecciona el archivo VBA_Diff.xlsx.", ok=False); return
        if not dest: self._log_final("ERROR: Selecciona el archivo destino.", ok=False); return
        self._limpiar_log()
        threading.Thread(
            target=self._thread_aplicar, args=(diff, dest, dry_run), daemon=True).start()

    def _thread_aplicar(self, diff, dest, dry_run):
        pythoncom.CoInitialize()
        writer = QueueWriter(self._log_queue)
        old_stdout = sys.stdout
        sys.stdout = writer
        try:
            modo = "SIMULACIÓN (sin escribir nada)" if dry_run else "APLICANDO cambios"
            print(f"\n{modo} -> {Path(dest).name}\n")
            module_codes = read_diff_excel(diff)
            print(f"  -> {len(module_codes)} módulos con diferencias en el diff\n")
            apply_vba(dest, module_codes, dry_run=dry_run)
            if dry_run:
                print(f"\nPara ver los cambios línea a línea pulsa 'Abrir diff'.")
        except Exception as e:
            sys.stdout = old_stdout
            self._log_final(f"ERROR: {e}", ok=False)
            return
        finally:
            sys.stdout = old_stdout
            pythoncom.CoUninitialize()
        fin = ("Simulación completada  —  pulsa 'Abrir diff' para ver los detalles"
               if dry_run else f"Cambios aplicados en: {Path(dest).name}")
        self._log_final(fin, ok=True)


# ── Entrada ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
