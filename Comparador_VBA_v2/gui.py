"""
Comparador VBA — interfaz gráfica
  Tab A: Comparar Módulos de 2 ficheros Excel
  Tab B: Comparar Módulos de 2 Carpetas
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
from compare import (extract_vba, compare as vba_compare, render_excel,
                     read_module_file, get_vb_name, VBAModule)
from apply import read_diff_excel, apply_vba, apply_module_file, apply_module_folder


# ── Registro de Windows ───────────────────────────────────────────────────────

REG_KEY    = r"Software\ComparadorVBA"
REG_FIELDS = (
    "excel_a", "excel_b", "diff_out", "diff_in", "dest1", "dest2",
    "modo_modulo",
    "folder1", "folder2", "diff_b_out", "diff_b_in", "dest_b1", "dest_b2",
    "tab_activa",
)

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
BG_COMPARAR = "#2E86C1"   # azul     — Tab A activa, bloque COMPARAR Excel
BG_CARPETA  = "#7D3C98"   # violeta  — Tab B activa
BG_COMP_B   = "#6C3483"   # violeta oscuro — bloque COMPARAR Carpetas
BG_APLICAR  = "#1E8449"   # verde    — bloque APLICAR (ambas tabs)
BG_REGISTRO = "#5D6D7E"   # gris     — REGISTRO

C_GENERAR = "#FAD7A0"
C_SIMULAR = "#AED6F1"
C_APLICAR = "#A9DFBF"
C_ABRIR   = "#D7BDE2"
C_SELEC   = "#D5D8DC"
C_ACCION  = "#F0B27A"

_EXTS_MOD = {".bas", ".cls", ".frm"}


# ── Aplicación ────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Comparador VBA")
        self.resizable(True, True)
        self.configure(bg=BG_APP)

        self._log_queue: queue.Queue = queue.Queue()
        self._cfg = _load_config()

        # Tab A state
        self._diff_path: str = ""
        self._modulos_excel1: list[str] = []
        self._modulos_excel2: list[str] = []
        self._carga_session: int = 0

        # Tab B state
        self._diff_b_path: str = ""
        self._mods_carpeta1: list[str] = []   # nombres de fichero con extensión
        self._mods_carpeta2: list[str] = []

        self._build_styles()
        self._build_ui()
        self._cargar_config()
        self._centrar_ventana(1600, 960)
        self._poll_log()

    # ── Centrado ──────────────────────────────────────────────────────────────

    def _centrar_ventana(self, w: int, h: int):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = max(0, (self.winfo_screenheight() - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(1200, 750)

    # ── Estilos ───────────────────────────────────────────────────────────────

    def _build_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",               font=FONT_UI, background=BG_APP)
        s.configure("TFrame",          background=BG_APP)
        s.configure("TEntry",          font=FONT_UI, fieldbackground="white")
        s.configure("Sel.TButton",     font=FONT_UI, background=C_SELEC)
        s.map("Sel.TButton",           background=[("active", "#BFC9CA")])
        s.configure("Accion.TButton",  font=FONT_UI, background=C_ACCION)
        s.map("Accion.TButton",        background=[("active", "#E59866")])
        s.configure("Generar.TButton", font=FONT_BOLD, background=C_GENERAR)
        s.map("Generar.TButton",       background=[("active", "#F5CBA7")])
        s.configure("Simular.TButton", font=FONT_UI,   background=C_SIMULAR)
        s.map("Simular.TButton",       background=[("active", "#85C1E9")])
        s.configure("AppBtn.TButton",  font=FONT_BOLD, background=C_APLICAR)
        s.map("AppBtn.TButton",        background=[("active", "#7DCEA0")])
        s.configure("Abrir.TButton",   font=FONT_UI,   background=C_ABRIR)
        s.map("Abrir.TButton",         background=[("active", "#C39BD3")])

    # ── Bloque coloreado con título (usa pack dentro del parent) ──────────────

    def _bloque(self, parent, titulo: str, bg_titulo: str,
                expand: bool = False, pady_top: int = 6) -> tuple[tk.Frame, tk.Frame]:
        outer = tk.Frame(parent, bg=BG_APP, bd=1, relief="solid",
                         highlightbackground="#AAAAAA", highlightthickness=1)
        outer.pack(fill="both" if expand else "x", expand=expand,
                   padx=12, pady=(pady_top, 0))
        outer.columnconfigure(0, weight=1)

        title_bar = tk.Frame(outer, bg=bg_titulo)
        title_bar.pack(fill=tk.X)
        title_bar.columnconfigure(0, weight=1)
        tk.Label(title_bar, text=f"  {titulo}",
                 bg=bg_titulo, fg="white", font=FONT_TITLE, anchor="w", pady=5
                 ).grid(row=0, column=0, sticky="ew")

        content = tk.Frame(outer, bg=BG_APP, padx=12, pady=8)
        content.pack(fill="both", expand=True)
        content.columnconfigure(1, weight=1)
        return title_bar, content

    # ── Estructura principal ──────────────────────────────────────────────────

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)   # REGISTRO se expande

        # Barra de pestañas
        tab_bar = tk.Frame(self, bg=BG_APP)
        tab_bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))

        self._btn_tab_a = tk.Button(
            tab_bar,
            text="  -A-  Comparar Módulos de 2 ficheros Excel  ",
            command=lambda: self._switch_tab("a"),
            font=FONT_TITLE, fg="white", bg=BG_COMPARAR,
            activebackground=BG_COMPARAR, activeforeground="white",
            relief="flat", bd=0, padx=8, pady=10, cursor="arrow")
        self._btn_tab_a.pack(side=tk.LEFT, padx=(0, 4))

        self._btn_tab_b = tk.Button(
            tab_bar,
            text="  -B-  Comparar Módulos de 2 Carpetas  ",
            command=lambda: self._switch_tab("b"),
            font=FONT_UI, fg="#546E7A", bg="#B8C1CC",
            activebackground=BG_CARPETA, activeforeground="white",
            relief="flat", bd=0, padx=8, pady=10, cursor="hand2")
        self._btn_tab_b.pack(side=tk.LEFT)

        # Área de contenido (frames apilados)
        self._frm_tabs = tk.Frame(self, bg=BG_APP)
        self._frm_tabs.grid(row=1, column=0, sticky="nsew")
        self._frm_tabs.columnconfigure(0, weight=1)
        self._frm_tabs.rowconfigure(0, weight=1)

        self._frm_tab_a = tk.Frame(self._frm_tabs, bg=BG_APP)
        self._frm_tab_a.grid(row=0, column=0, sticky="nsew")
        self._frm_tab_a.columnconfigure(0, weight=1)

        self._frm_tab_b = tk.Frame(self._frm_tabs, bg=BG_APP)
        self._frm_tab_b.grid(row=0, column=0, sticky="nsew")
        self._frm_tab_b.columnconfigure(0, weight=1)

        self._build_tab_a()
        self._build_tab_b()
        self._build_registro()

        self._frm_tab_a.tkraise()

    def _switch_tab(self, tab: str):
        self._tab_activa = tab
        if tab == "a":
            self._frm_tab_a.tkraise()
            self._btn_tab_a.config(fg="white", bg=BG_COMPARAR, font=FONT_TITLE, cursor="arrow")
            self._btn_tab_b.config(fg="#546E7A", bg="#B8C1CC",  font=FONT_UI,    cursor="hand2")
        else:
            self._frm_tab_b.tkraise()
            self._btn_tab_a.config(fg="#546E7A", bg="#B8C1CC",  font=FONT_UI,    cursor="hand2")
            self._btn_tab_b.config(fg="white", bg=BG_CARPETA,  font=FONT_TITLE, cursor="arrow")
        self._guardar_config()

    # ══════════════════════════════════════════════════════════════════════════
    # TAB A — Comparar 2 ficheros Excel (o módulos individuales)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_a(self):
        p = self._frm_tab_a

        # ── COMPARAR ──────────────────────────────────────────────────────────
        tb_cmp, cmp = self._bloque(p, "COMPARAR", BG_COMPARAR, pady_top=8)
        ttk.Button(tb_cmp, text="Borrar datos", style="Accion.TButton",
                   command=self._borrar_comparar
                   ).grid(row=0, column=1, padx=8, pady=4, ipadx=6)

        self.var_a   = tk.StringVar()
        self.var_b   = tk.StringVar()
        self.var_out = tk.StringVar()
        self.var_mod1 = tk.StringVar()
        self.var_mod2 = tk.StringVar()
        self.var_solo_modulo = tk.BooleanVar(value=False)
        self.var_modo_modulo = tk.BooleanVar(value=False)

        self._filetypes_a    = [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")]
        self._filetypes_b_a  = [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")]
        self._filetypes_dest = [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")]

        # Fila 0 — Fuente 1
        self._lbl_fuente1 = tk.Label(cmp, text="Excel 1:", bg=BG_APP, font=FONT_UI, anchor="w")
        self._lbl_fuente1.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(cmp, textvariable=self.var_a, font=FONT_UI).grid(
            row=0, column=1, sticky="ew", pady=4)
        ttk.Button(cmp, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_archivo(self.var_a, self._filetypes_a)
                   ).grid(row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        # Fila 1 — Fuente 2
        self._lbl_fuente2 = tk.Label(cmp, text="Excel 2:", bg=BG_APP, font=FONT_UI, anchor="w")
        self._lbl_fuente2.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(cmp, textvariable=self.var_b, font=FONT_UI).grid(
            row=1, column=1, sticky="ew", pady=4)
        ttk.Button(cmp, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_archivo(self.var_b, self._filetypes_b_a)
                   ).grid(row=1, column=2, padx=(8, 0), pady=4, ipadx=4)

        # Fila 2 — Guardar diff
        self._fila_guardar(cmp, 2, "Guardar diff en:", self.var_out, [("Excel", "*.xlsx")])

        # Fila 3 — Modo Módulo (siempre visible)
        frm_modo = tk.Frame(cmp, bg=BG_APP)
        frm_modo.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Checkbutton(frm_modo, text="Cambiar Selección a Módulos",
                        variable=self.var_modo_modulo,
                        command=self._toggle_modo_modulo).pack(side=tk.LEFT)

        # Fila 4 — Comparar sólo un módulo (dinámico)
        self._frm_chk = tk.Frame(cmp, bg=BG_APP)
        ttk.Checkbutton(self._frm_chk, text="Comparar sólo un módulo",
                        variable=self.var_solo_modulo,
                        command=self._toggle_solo_modulo).pack(side=tk.LEFT)

        # Fila 5 — Comboboxes módulos (dinámico)
        self._frm_mods = tk.Frame(cmp, bg=BG_APP)
        self._frm_mods.columnconfigure(1, weight=1)

        tk.Label(self._frm_mods, text="Módulo Excel 1:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=3)
        self._cmb_mod1 = ttk.Combobox(self._frm_mods, textvariable=self.var_mod1,
                                       state="readonly", font=FONT_UI)
        self._cmb_mod1.grid(row=0, column=1, sticky="ew", pady=3)
        self._cmb_mod1.bind("<<ComboboxSelected>>", self._on_mod1_selected)

        tk.Label(self._frm_mods, text="Módulo Excel 2:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=3)
        self._cmb_mod2 = ttk.Combobox(self._frm_mods, textvariable=self.var_mod2,
                                       state="readonly", font=FONT_UI)
        self._cmb_mod2.grid(row=1, column=1, sticky="ew", pady=3)

        # Fila 6 — Botones + resultado (siempre)
        frm_cb = tk.Frame(cmp, bg=BG_APP)
        frm_cb.grid(row=6, column=0, columnspan=3, pady=(10, 4), sticky="ew")
        frm_cb.columnconfigure(1, weight=1)

        frm_izq = tk.Frame(frm_cb, bg=BG_APP)
        frm_izq.grid(row=0, column=0, sticky="w")
        ttk.Button(frm_izq, text="GENERAR VBA_Diff.xlsx", style="Generar.TButton",
                   command=self._accion_comparar).pack(side=tk.LEFT, ipadx=12, ipady=4)

        frm_res = tk.Frame(frm_cb, bg=BG_APP)
        frm_res.grid(row=0, column=1, sticky="w", padx=(20, 0))
        self._lbl_stats      = tk.Label(frm_res, text="", bg=BG_APP, font=FONT_UI, fg="#2C3E50")
        self._lbl_identicos  = tk.Label(frm_res, text="✓  IDÉNTICOS", bg=BG_APP, fg="#1E8449", font=FONT_BOLD)
        self._btn_abrir_res  = ttk.Button(frm_res, text="", style="Abrir.TButton",
                                           command=self._abrir_diff_resultado)

        # ── APLICAR ───────────────────────────────────────────────────────────
        tb_app, app = self._bloque(p, "APLICAR", BG_APLICAR)
        ttk.Button(tb_app, text="Borrar datos", style="Accion.TButton",
                   command=self._borrar_aplicar
                   ).grid(row=0, column=1, padx=8, pady=4, ipadx=6)

        self.var_diff  = tk.StringVar()
        self.var_dest1 = tk.StringVar()
        self.var_dest2 = tk.StringVar()

        self._lbl_diff_input = tk.Label(app, text="VBA_Diff.xlsx:", bg=BG_APP, font=FONT_UI, anchor="w")
        self._lbl_diff_input.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(app, textvariable=self.var_diff, font=FONT_UI).grid(
            row=0, column=1, sticky="ew", pady=4)
        ttk.Button(app, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_archivo(
                       self.var_diff, [("Excel diff", "*.xlsx"), ("Todos", "*.*")])
                   ).grid(row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        tk.Label(app, text="Destino 1:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(app, textvariable=self.var_dest1, font=FONT_UI).grid(
            row=1, column=1, sticky="ew", pady=4)
        ttk.Button(app, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_archivo(self.var_dest1, self._filetypes_dest)
                   ).grid(row=1, column=2, padx=(8, 0), pady=4, ipadx=4)

        tk.Label(app, text="Destino 2:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(app, textvariable=self.var_dest2, font=FONT_UI).grid(
            row=2, column=1, sticky="ew", pady=4)
        ttk.Button(app, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_archivo(self.var_dest2, self._filetypes_dest)
                   ).grid(row=2, column=2, padx=(8, 0), pady=4, ipadx=4)

        frm_ab = tk.Frame(app, bg=BG_APP)
        frm_ab.grid(row=3, column=0, columnspan=3, pady=(10, 4), sticky="w")
        ttk.Button(frm_ab, text="Simular → Destino 1", style="Simular.TButton",
                   command=lambda: self._accion_aplicar(self.var_dest1, dry_run=True)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab, text="Aplicar → Destino 1", style="AppBtn.TButton",
                   command=lambda: self._accion_aplicar(self.var_dest1, dry_run=False)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab, text="Simular → Destino 2", style="Simular.TButton",
                   command=lambda: self._accion_aplicar(self.var_dest2, dry_run=True)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab, text="Aplicar → Destino 2", style="AppBtn.TButton",
                   command=lambda: self._accion_aplicar(self.var_dest2, dry_run=False)
                   ).pack(side=tk.LEFT, ipadx=6, ipady=3)

        # Traces Tab A
        for var in (self.var_a, self.var_b, self.var_out,
                    self.var_diff, self.var_dest1, self.var_dest2):
            var.trace_add("write", lambda *_: self._guardar_config())
        self.var_a.trace_add("write", lambda *_: self._on_fuente_changed("a"))
        self.var_b.trace_add("write", lambda *_: self._on_fuente_changed("b"))

    # ══════════════════════════════════════════════════════════════════════════
    # TAB B — Comparar Módulos de 2 Carpetas
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_b(self):
        p = self._frm_tab_b

        # ── COMPARAR CARPETAS ─────────────────────────────────────────────────
        tb_cmp, cmp = self._bloque(p, "COMPARAR CARPETAS", BG_COMP_B, pady_top=8)
        ttk.Button(tb_cmp, text="Borrar datos", style="Accion.TButton",
                   command=self._borrar_comparar_b
                   ).grid(row=0, column=1, padx=8, pady=4, ipadx=6)

        self.var_folder1   = tk.StringVar()
        self.var_folder2   = tk.StringVar()
        self.var_out_b     = tk.StringVar()
        self.var_mod_b1    = tk.StringVar()
        self.var_mod_b2    = tk.StringVar()
        self.var_solo_mod_b = tk.BooleanVar(value=False)

        # Fila 0 — Carpeta 1
        tk.Label(cmp, text="Carpeta 1:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(cmp, textvariable=self.var_folder1, font=FONT_UI).grid(
            row=0, column=1, sticky="ew", pady=4)
        ttk.Button(cmp, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_carpeta(self.var_folder1)
                   ).grid(row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        # Fila 1 — Carpeta 2
        tk.Label(cmp, text="Carpeta 2:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(cmp, textvariable=self.var_folder2, font=FONT_UI).grid(
            row=1, column=1, sticky="ew", pady=4)
        ttk.Button(cmp, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_carpeta(self.var_folder2)
                   ).grid(row=1, column=2, padx=(8, 0), pady=4, ipadx=4)

        # Fila 2 — Guardar diff
        self._fila_guardar(cmp, 2, "Guardar diff en:", self.var_out_b, [("Excel", "*.xlsx")])

        # Fila 3 — Comparar sólo un módulo (dinámico)
        self._frm_chk_b = tk.Frame(cmp, bg=BG_APP)
        ttk.Checkbutton(self._frm_chk_b, text="Comparar sólo un módulo",
                        variable=self.var_solo_mod_b,
                        command=self._toggle_solo_mod_b).pack(side=tk.LEFT)

        # Fila 4 — Comboboxes módulos de carpeta (dinámico)
        self._frm_mods_b = tk.Frame(cmp, bg=BG_APP)
        self._frm_mods_b.columnconfigure(1, weight=1)

        tk.Label(self._frm_mods_b, text="Módulo Carpeta 1:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=3)
        self._cmb_mod_b1 = ttk.Combobox(self._frm_mods_b, textvariable=self.var_mod_b1,
                                          state="readonly", font=FONT_UI)
        self._cmb_mod_b1.grid(row=0, column=1, sticky="ew", pady=3)
        self._cmb_mod_b1.bind("<<ComboboxSelected>>", self._on_mod_b1_selected)

        tk.Label(self._frm_mods_b, text="Módulo Carpeta 2:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=3)
        self._cmb_mod_b2 = ttk.Combobox(self._frm_mods_b, textvariable=self.var_mod_b2,
                                          state="readonly", font=FONT_UI)
        self._cmb_mod_b2.grid(row=1, column=1, sticky="ew", pady=3)

        # Fila 5 — Botones + resultado (siempre)
        frm_cb_b = tk.Frame(cmp, bg=BG_APP)
        frm_cb_b.grid(row=5, column=0, columnspan=3, pady=(10, 4), sticky="ew")
        frm_cb_b.columnconfigure(1, weight=1)

        frm_izq_b = tk.Frame(frm_cb_b, bg=BG_APP)
        frm_izq_b.grid(row=0, column=0, sticky="w")
        ttk.Button(frm_izq_b, text="GENERAR Modul_Diff.xlsx", style="Generar.TButton",
                   command=self._accion_comparar_b).pack(side=tk.LEFT, ipadx=12, ipady=4)

        frm_res_b = tk.Frame(frm_cb_b, bg=BG_APP)
        frm_res_b.grid(row=0, column=1, sticky="w", padx=(20, 0))
        self._lbl_stats_b     = tk.Label(frm_res_b, text="", bg=BG_APP, font=FONT_UI, fg="#2C3E50")
        self._lbl_identicos_b = tk.Label(frm_res_b, text="✓  IDÉNTICOS", bg=BG_APP, fg="#1E8449", font=FONT_BOLD)
        self._btn_abrir_res_b = ttk.Button(frm_res_b, text="", style="Abrir.TButton",
                                            command=self._abrir_diff_resultado_b)

        # ── APLICAR CARPETAS ──────────────────────────────────────────────────
        tb_app_b, app_b = self._bloque(p, "APLICAR", BG_APLICAR)
        ttk.Button(tb_app_b, text="Borrar datos", style="Accion.TButton",
                   command=self._borrar_aplicar_b
                   ).grid(row=0, column=1, padx=8, pady=4, ipadx=6)

        self.var_diff_b  = tk.StringVar()
        self.var_dest_b1 = tk.StringVar()
        self.var_dest_b2 = tk.StringVar()

        tk.Label(app_b, text="Modul_Diff.xlsx:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(app_b, textvariable=self.var_diff_b, font=FONT_UI).grid(
            row=0, column=1, sticky="ew", pady=4)
        ttk.Button(app_b, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_archivo(
                       self.var_diff_b, [("Excel diff", "*.xlsx"), ("Todos", "*.*")])
                   ).grid(row=0, column=2, padx=(8, 0), pady=4, ipadx=4)

        tk.Label(app_b, text="Destino 1:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(app_b, textvariable=self.var_dest_b1, font=FONT_UI).grid(
            row=1, column=1, sticky="ew", pady=4)
        ttk.Button(app_b, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_carpeta(self.var_dest_b1)
                   ).grid(row=1, column=2, padx=(8, 0), pady=4, ipadx=4)

        tk.Label(app_b, text="Destino 2:", bg=BG_APP, font=FONT_UI, anchor="w").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(app_b, textvariable=self.var_dest_b2, font=FONT_UI).grid(
            row=2, column=1, sticky="ew", pady=4)
        ttk.Button(app_b, text="Seleccionar", style="Sel.TButton",
                   command=lambda: self._seleccionar_carpeta(self.var_dest_b2)
                   ).grid(row=2, column=2, padx=(8, 0), pady=4, ipadx=4)

        frm_ab_b = tk.Frame(app_b, bg=BG_APP)
        frm_ab_b.grid(row=3, column=0, columnspan=3, pady=(10, 4), sticky="w")
        ttk.Button(frm_ab_b, text="Simular → Destino 1", style="Simular.TButton",
                   command=lambda: self._accion_aplicar_b(self.var_dest_b1, dry_run=True)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab_b, text="Aplicar → Destino 1", style="AppBtn.TButton",
                   command=lambda: self._accion_aplicar_b(self.var_dest_b1, dry_run=False)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab_b, text="Simular → Destino 2", style="Simular.TButton",
                   command=lambda: self._accion_aplicar_b(self.var_dest_b2, dry_run=True)
                   ).pack(side=tk.LEFT, padx=(0, 6), ipadx=6, ipady=3)
        ttk.Button(frm_ab_b, text="Aplicar → Destino 2", style="AppBtn.TButton",
                   command=lambda: self._accion_aplicar_b(self.var_dest_b2, dry_run=False)
                   ).pack(side=tk.LEFT, ipadx=6, ipady=3)

        # Traces Tab B
        for var in (self.var_folder1, self.var_folder2, self.var_out_b,
                    self.var_diff_b, self.var_dest_b1, self.var_dest_b2):
            var.trace_add("write", lambda *_: self._guardar_config())
        self.var_folder1.trace_add("write", lambda *_: self._on_carpeta_changed())
        self.var_folder2.trace_add("write", lambda *_: self._on_carpeta_changed())

    # ── REGISTRO (compartido) ─────────────────────────────────────────────────

    def _build_registro(self):
        reg_outer = tk.Frame(self, bg=BG_APP, bd=1, relief="solid",
                             highlightbackground="#AAAAAA", highlightthickness=1)
        reg_outer.grid(row=2, column=0, padx=12, pady=(6, 12), sticky="nsew")
        reg_outer.columnconfigure(0, weight=1)
        reg_outer.rowconfigure(1, weight=1)

        tb_reg = tk.Frame(reg_outer, bg=BG_REGISTRO)
        tb_reg.grid(row=0, column=0, sticky="ew")
        tb_reg.columnconfigure(0, weight=1)
        tk.Label(tb_reg, text="  REGISTRO",
                 bg=BG_REGISTRO, fg="white", font=FONT_TITLE, anchor="w", pady=5
                 ).grid(row=0, column=0, sticky="ew")
        ttk.Button(tb_reg, text="Limpiar registro", style="Accion.TButton",
                   command=self._limpiar_log
                   ).grid(row=0, column=1, padx=8, pady=4, ipadx=6)

        reg_content = tk.Frame(reg_outer, bg=BG_APP, padx=6, pady=6)
        reg_content.grid(row=1, column=0, sticky="nsew")
        reg_content.columnconfigure(0, weight=1)
        reg_content.rowconfigure(0, weight=1)

        self.log_widget = scrolledtext.ScrolledText(
            reg_content, height=14, state=tk.NORMAL,
            font=FONT_LOG, wrap=tk.WORD,
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="white")
        self.log_widget.grid(row=0, column=0, sticky="nsew")
        self.log_widget.tag_configure("ok",        foreground="#4EC94E")
        self.log_widget.tag_configure("error",     foreground="#FF6B6B")
        self.log_widget.tag_configure("diff_add",
            foreground="#AFFFAF", background="#1A4A1A")
        self.log_widget.tag_configure("diff_del",
            foreground="#FFAFAF", background="#4A1A1A")
        self.log_widget.tag_configure("diff_hunk",
            foreground="#F8C471", background="#2C2C00")
        self.log_widget.tag_configure("diff_head",
            foreground="white", background="#1A5276",
            font=("Courier New", 12, "bold"),
            spacing1=10, spacing3=10)

    # ── Helpers UI ────────────────────────────────────────────────────────────

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

    def _seleccionar_carpeta(self, var):
        inicial = var.get() if var.get() and os.path.isdir(var.get()) else "/"
        ruta = filedialog.askdirectory(initialdir=inicial)
        if ruta:
            var.set(ruta)

    # ── Resultado Tab A ───────────────────────────────────────────────────────

    def _mostrar_resultado(self, total: int, con_dif: int, diff_path: str,
                           filtro: str | None = None):
        if filtro:
            estado = "con diferencias" if con_dif else "sin diferencias"
            txt = f"Módulo '{filtro}'  ·  {estado}"
        else:
            txt = f"{total} módulos  ·  {con_dif} con diferencias"
        self._lbl_stats.config(text=txt)
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

    def _abrir_diff_resultado(self):
        if self._diff_path and os.path.exists(self._diff_path):
            os.startfile(self._diff_path)

    # ── Resultado Tab B ───────────────────────────────────────────────────────

    def _mostrar_resultado_b(self, total: int, con_dif: int, diff_path: str,
                              filtro_b1: str | None = None, filtro_b2: str | None = None):
        if filtro_b1:
            estado = "con diferencias" if con_dif else "sin diferencias"
            n1, n2 = Path(filtro_b1).stem, Path(filtro_b2).stem
            txt = (f"'{n1}'  ·  {estado}" if n1 == n2
                   else f"'{n1}' vs '{n2}'  ·  {estado}")
        else:
            txt = f"{total} módulos  ·  {con_dif} con diferencias"
        self._lbl_stats_b.config(text=txt)
        self._lbl_stats_b.pack(side=tk.LEFT, padx=(0, 18))
        self._lbl_identicos_b.pack_forget()
        self._btn_abrir_res_b.pack_forget()
        self._btn_abrir_res_b.config(text=f"Abrir  {Path(diff_path).name}")
        self._btn_abrir_res_b.pack(side=tk.LEFT, ipadx=6, ipady=2, padx=(0, 10))
        if con_dif == 0:
            self._lbl_identicos_b.pack(side=tk.LEFT)

    def _limpiar_resultado_b(self):
        self._lbl_stats_b.config(text="")
        self._lbl_stats_b.pack_forget()
        self._lbl_identicos_b.pack_forget()
        self._btn_abrir_res_b.pack_forget()

    def _abrir_diff_resultado_b(self):
        if self._diff_b_path and os.path.exists(self._diff_b_path):
            os.startfile(self._diff_b_path)

    # ── Config ────────────────────────────────────────────────────────────────

    def _cargar_config(self):
        cfg = self._cfg
        self.var_a.set(cfg.get("excel_a",   ""))
        self.var_b.set(cfg.get("excel_b",   ""))
        self.var_out.set(cfg.get("diff_out", ""))
        self.var_diff.set(cfg.get("diff_in", ""))
        self.var_dest1.set(cfg.get("dest1",  ""))
        self.var_dest2.set(cfg.get("dest2",  ""))
        self.var_modo_modulo.set(cfg.get("modo_modulo", "") == "1")
        self._aplicar_modo()

        self.var_folder1.set(cfg.get("folder1",   ""))
        self.var_folder2.set(cfg.get("folder2",   ""))
        self.var_out_b.set(cfg.get("diff_b_out",  ""))
        self.var_diff_b.set(cfg.get("diff_b_in",  ""))
        self.var_dest_b1.set(cfg.get("dest_b1",   ""))
        self.var_dest_b2.set(cfg.get("dest_b2",   ""))

        tab = cfg.get("tab_activa", "a")
        if tab in ("a", "b"):
            self._switch_tab(tab)

    def _guardar_config(self):
        _save_config({
            "excel_a":      self.var_a.get(),
            "excel_b":      self.var_b.get(),
            "diff_out":     self.var_out.get(),
            "diff_in":      self.var_diff.get(),
            "dest1":        self.var_dest1.get(),
            "dest2":        self.var_dest2.get(),
            "modo_modulo":  "1" if self.var_modo_modulo.get() else "0",
            "folder1":      self.var_folder1.get(),
            "folder2":      self.var_folder2.get(),
            "diff_b_out":   self.var_out_b.get(),
            "diff_b_in":    self.var_diff_b.get(),
            "dest_b1":      self.var_dest_b1.get(),
            "dest_b2":      self.var_dest_b2.get(),
            "tab_activa":   getattr(self, "_tab_activa", "a"),
        })

    # ── Tab A — Modo Módulo ───────────────────────────────────────────────────

    def _toggle_modo_modulo(self):
        for v in (self.var_a, self.var_b, self.var_out):
            v.set("")
        self._limpiar_resultado()
        self._reset_filtro()
        self._aplicar_modo()
        self._guardar_config()

    def _aplicar_modo(self):
        modo = self.var_modo_modulo.get()
        if modo:
            self._lbl_fuente1.config(text="Módulo 1:")
            self._lbl_fuente2.config(text="Módulo 2:")
            self._lbl_diff_input.config(text="Modul_Diff.xlsx:")
            self._filetypes_a    = [("Módulos VBA", "*.bas *.cls *.frm"), ("Todos", "*.*")]
            self._filetypes_b_a  = [("Módulos VBA", "*.bas *.cls *.frm"), ("Todos", "*.*")]
            self._filetypes_dest = [("Módulos VBA", "*.bas *.cls *.frm"), ("Todos", "*.*")]
        else:
            self._lbl_fuente1.config(text="Excel 1:")
            self._lbl_fuente2.config(text="Excel 2:")
            self._lbl_diff_input.config(text="VBA_Diff.xlsx:")
            self._filetypes_a    = [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")]
            self._filetypes_b_a  = [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")]
            self._filetypes_dest = [("Excel con macros", "*.xlsm *.xlam"), ("Todos", "*.*")]
        self._actualizar_checkbox()

    def _on_fuente_changed(self, which: str):
        self._limpiar_resultado()
        self._reset_filtro()
        self._actualizar_checkbox()
        if self.var_modo_modulo.get():
            if which == "a":
                self._actualizar_filtro_mod2()
            self._sugerir_nombre_salida()

    def _actualizar_filtro_mod2(self):
        a = self.var_a.get().strip()
        if not a or not Path(a).suffix:
            self._filetypes_b_a = [("Módulos VBA", "*.bas *.cls *.frm"), ("Todos", "*.*")]
            return
        ext = Path(a).suffix.lower()
        nombres = {".bas": "Módulo estándar", ".cls": "Módulo de clase", ".frm": "Formulario"}
        self._filetypes_b_a = [(nombres.get(ext, "Módulo VBA"), f"*{ext}"), ("Todos", "*.*")]
        b = self.var_b.get().strip()
        if b and Path(b).suffix.lower() != ext:
            self.var_b.set("")

    def _sugerir_nombre_salida(self):
        a = self.var_a.get().strip()
        b = self.var_b.get().strip()
        if not a or not b:
            return
        out = self.var_out.get().strip()
        if out and not Path(out).name.startswith("Modul_Diff"):
            return
        name_a = get_vb_name(a)
        name_b = get_vb_name(b)
        sufijo = name_a if name_a == name_b else f"{name_a}-{name_b}"
        self.var_out.set(str(Path(a).parent / f"Modul_Diff {sufijo}.xlsx"))

    def _actualizar_checkbox(self):
        if (self.var_a.get().strip() and self.var_b.get().strip()
                and not self.var_modo_modulo.get()):
            self._frm_chk.grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        else:
            self._frm_chk.grid_forget()

    def _toggle_solo_modulo(self):
        if self.var_solo_modulo.get():
            self._frm_mods.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(4, 0))
            self._cargar_modulos()
        else:
            self._frm_mods.grid_forget()
            self.var_mod1.set("")
            self.var_mod2.set("")

    def _reset_filtro(self):
        self.var_solo_modulo.set(False)
        self._frm_mods.grid_forget()
        self.var_mod1.set("")
        self.var_mod2.set("")
        self._modulos_excel1 = []
        self._modulos_excel2 = []

    def _cargar_modulos(self):
        a = self.var_a.get().strip()
        b = self.var_b.get().strip()
        if not a or not b:
            self._log_final("ERROR: Selecciona ambos Excel antes de activar el filtro.", ok=False)
            self.var_solo_modulo.set(False)
            return
        self._carga_session += 1
        session = self._carga_session
        self._cmb_mod1.set("Cargando...")
        self._cmb_mod2.set("Cargando...")
        self._cmb_mod1["values"] = []
        self._cmb_mod2["values"] = []
        self._log_final("Cargando lista de módulos...", ok=True)
        threading.Thread(
            target=self._thread_cargar_mods, args=(a, b, session), daemon=True).start()

    def _thread_cargar_mods(self, path_a: str, path_b: str, session: int):
        pythoncom.CoInitialize()
        try:
            mods_a = sorted(extract_vba(path_a).keys())
            mods_b = sorted(extract_vba(path_b).keys())
        except Exception as e:
            self.after(0, lambda: self._log_final(f"ERROR cargando módulos: {e}", ok=False))
            return
        finally:
            pythoncom.CoUninitialize()
        if session != self._carga_session:
            return
        self._modulos_excel1 = mods_a
        self._modulos_excel2 = mods_b
        self.after(0, lambda: self._actualizar_cmbs(mods_a, mods_b))

    def _actualizar_cmbs(self, mods_a: list[str], mods_b: list[str]):
        self._cmb_mod1["values"] = mods_a
        self._cmb_mod2["values"] = mods_b
        self._cmb_mod1.set("")
        self._cmb_mod2.set("")
        self._limpiar_log()
        self._log_final(
            f"Listo — {len(mods_a)} módulos en Excel 1, {len(mods_b)} en Excel 2. "
            "Selecciona el módulo a comparar.", ok=True)

    def _on_mod1_selected(self, _event=None):
        sel = self.var_mod1.get()
        if sel and sel in self._modulos_excel2:
            self.var_mod2.set(sel)

    # ── Tab A — Borrar / Abrir ────────────────────────────────────────────────

    def _borrar_comparar(self):
        for v in (self.var_a, self.var_b, self.var_out):
            v.set("")
        self._limpiar_resultado()
        self._reset_filtro()

    def _borrar_aplicar(self):
        for v in (self.var_diff, self.var_dest1, self.var_dest2):
            v.set("")

    # ── Tab B — Checkbox / Módulos de carpeta ─────────────────────────────────

    def _on_carpeta_changed(self):
        self._limpiar_resultado_b()
        self.var_solo_mod_b.set(False)
        self._frm_mods_b.grid_forget()
        self.var_mod_b1.set("")
        self.var_mod_b2.set("")
        self._mods_carpeta1 = []
        self._mods_carpeta2 = []
        self._actualizar_checkbox_b()

    def _actualizar_checkbox_b(self):
        if self.var_folder1.get().strip() and self.var_folder2.get().strip():
            self._frm_chk_b.grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))
        else:
            self._frm_chk_b.grid_forget()

    def _toggle_solo_mod_b(self):
        if self.var_solo_mod_b.get():
            self._frm_mods_b.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(4, 0))
            self._cargar_mods_carpeta_b()
        else:
            self._frm_mods_b.grid_forget()
            self.var_mod_b1.set("")
            self.var_mod_b2.set("")

    def _cargar_mods_carpeta_b(self):
        def list_mods(folder: str) -> list[str]:
            try:
                return sorted(
                    fp.name for fp in Path(folder).iterdir()
                    if fp.suffix.lower() in _EXTS_MOD)
            except Exception:
                return []
        f1 = self.var_folder1.get().strip()
        f2 = self.var_folder2.get().strip()
        self._mods_carpeta1 = list_mods(f1)
        self._mods_carpeta2 = list_mods(f2)
        self._cmb_mod_b1["values"] = self._mods_carpeta1
        self._cmb_mod_b2["values"] = self._mods_carpeta2
        self.var_mod_b1.set("")
        self.var_mod_b2.set("")
        self._limpiar_log()
        self._log_final(
            f"Listo — {len(self._mods_carpeta1)} módulos en Carpeta 1, "
            f"{len(self._mods_carpeta2)} en Carpeta 2. Selecciona el módulo a comparar.",
            ok=True)

    def _on_mod_b1_selected(self, _event=None):
        sel = self.var_mod_b1.get()
        if sel and sel in self._mods_carpeta2:
            self.var_mod_b2.set(sel)

    # ── Tab B — Borrar ────────────────────────────────────────────────────────

    def _borrar_comparar_b(self):
        for v in (self.var_folder1, self.var_folder2, self.var_out_b):
            v.set("")
        self._limpiar_resultado_b()
        self._on_carpeta_changed()

    def _borrar_aplicar_b(self):
        for v in (self.var_diff_b, self.var_dest_b1, self.var_dest_b2):
            v.set("")

    # ── Log ───────────────────────────────────────────────────────────────────

    def _poll_log(self):
        while not self._log_queue.empty():
            tag, texto = self._log_queue.get_nowait()
            if tag in ("ok", "error"):
                self.log_widget.insert(tk.END, texto, tag)
            else:
                t = texto.lstrip("\n")
                n_nl = len(texto) - len(t)
                if t.startswith("+") and not t.startswith("+++ "):
                    self.log_widget.insert(tk.END, texto, "diff_add")
                elif t.startswith("-") and not t.startswith("--- "):
                    self.log_widget.insert(tk.END, texto, "diff_del")
                elif t.startswith("@@"):
                    self.log_widget.insert(tk.END, texto, "diff_hunk")
                elif t.startswith("──"):
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

    # ── Tab A — Acciones COMPARAR ─────────────────────────────────────────────

    def _accion_comparar(self):
        a   = self.var_a.get().strip()
        b   = self.var_b.get().strip()
        out = self.var_out.get().strip()
        if not a:   self._log_final("ERROR: Selecciona el Excel 1.", ok=False);          return
        if not b:   self._log_final("ERROR: Selecciona el Excel 2.", ok=False);          return
        if not out: self._log_final("ERROR: Indica dónde guardar el diff.", ok=False);   return

        filtro: str | None = None
        if self.var_solo_modulo.get():
            m1 = self.var_mod1.get().strip()
            m2 = self.var_mod2.get().strip()
            if not m1: self._log_final("ERROR: Selecciona el módulo de Excel 1.", ok=False); return
            if not m2: self._log_final("ERROR: Selecciona el módulo de Excel 2.", ok=False); return
            filtro = m1
            self._filtro_m2 = m2

        self._limpiar_log()
        self._limpiar_resultado()
        threading.Thread(
            target=self._thread_comparar, args=(a, b, out, filtro), daemon=True).start()

    def _thread_comparar(self, a, b, out, filtro: str | None):
        pythoncom.CoInitialize()
        writer = QueueWriter(self._log_queue)
        old_stdout = sys.stdout
        sys.stdout = writer
        total = con_dif = comparados = 0
        try:
            la, lb = Path(a).name, Path(b).name
            if self.var_modo_modulo.get():
                print(f"\nLeyendo módulo: {la}")
                mod_a = read_module_file(a)
                print(f"Leyendo módulo: {lb}\n")
                mod_b = read_module_file(b)
                if mod_a.name != mod_b.name:
                    mod_b = VBAModule(name=mod_a.name, code=mod_b.code, kind=mod_b.kind)
                mods_a = {mod_a.name: mod_a}
                mods_b = {mod_a.name: mod_b}
            else:
                print(f"\nExtrayendo VBA de: {la}")
                mods_a = extract_vba(a)
                print(f"  -> {len(mods_a)} módulos encontrados")
                print(f"Extrayendo VBA de: {lb}")
                mods_b = extract_vba(b)
                print(f"  -> {len(mods_b)} módulos encontrados\n")

                if filtro:
                    m2 = getattr(self, "_filtro_m2", filtro)
                    if m2 != filtro and m2 in mods_b:
                        mods_b[filtro] = VBAModule(
                            name=filtro, code=mods_b[m2].code, kind=mods_b[m2].kind)
                        del mods_b[m2]
                    print(f"Modo filtro: comparando módulo '{filtro}'")
                    if m2 != filtro:
                        print(f"  ('{filtro}' en Excel 1  vs  '{m2}' en Excel 2)")

            diffs = vba_compare(mods_a, mods_b, modulo_filtro=filtro)
            render_excel(diffs, la, lb, out)

            total      = len(diffs)
            skip       = sum(1 for d in diffs if d.status == "skip")
            con_dif    = sum(1 for d in diffs if d.status not in ("equal", "skip"))
            comparados = total - skip

            if self.var_modo_modulo.get():
                print(f"Módulo: {'con diferencias' if con_dif else 'sin diferencias'}")
            elif filtro:
                print(f"Módulo comparado: '{filtro}'")
            else:
                print(f"Módulos con diferencias: {con_dif} de {comparados} comparados")
        except Exception as e:
            sys.stdout = old_stdout
            self._log_final(f"ERROR: {e}", ok=False)
            return
        finally:
            sys.stdout = old_stdout
            pythoncom.CoUninitialize()

        self._diff_path = out
        self.var_diff.set(out)
        self.after(0, lambda: self._mostrar_resultado(comparados, con_dif, out, filtro))
        self._log_final(
            f"Diff guardado: {Path(out).name}  —  pulsa 'Abrir diff'" if con_dif > 0
            else f"Sin diferencias. Diff guardado: {Path(out).name}",
            ok=True)

    # ── Tab A — Acciones APLICAR ──────────────────────────────────────────────

    def _accion_aplicar(self, dest_var, dry_run):
        diff = self.var_diff.get().strip()
        dest = dest_var.get().strip()
        if not diff: self._log_final("ERROR: Selecciona el archivo diff.", ok=False);    return
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
            if self.var_modo_modulo.get():
                for new_code in module_codes.values():
                    apply_module_file(dest, new_code, dry_run=dry_run)
            else:
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

    # ── Tab B — Acciones COMPARAR ─────────────────────────────────────────────

    def _accion_comparar_b(self):
        f1  = self.var_folder1.get().strip()
        f2  = self.var_folder2.get().strip()
        out = self.var_out_b.get().strip()
        if not f1:  self._log_final("ERROR: Selecciona la Carpeta 1.", ok=False);        return
        if not f2:  self._log_final("ERROR: Selecciona la Carpeta 2.", ok=False);        return
        if not out: self._log_final("ERROR: Indica dónde guardar el diff.", ok=False);   return

        filtro_b1 = filtro_b2 = None
        if self.var_solo_mod_b.get():
            filtro_b1 = self.var_mod_b1.get().strip()
            filtro_b2 = self.var_mod_b2.get().strip()
            if not filtro_b1: self._log_final("ERROR: Selecciona el módulo de Carpeta 1.", ok=False); return
            if not filtro_b2: self._log_final("ERROR: Selecciona el módulo de Carpeta 2.", ok=False); return

        self._limpiar_log()
        self._limpiar_resultado_b()
        threading.Thread(
            target=self._thread_comparar_b,
            args=(f1, f2, out, filtro_b1, filtro_b2), daemon=True).start()

    def _thread_comparar_b(self, f1, f2, out, filtro_b1, filtro_b2):
        writer = QueueWriter(self._log_queue)
        old_stdout = sys.stdout
        sys.stdout = writer
        total = con_dif = 0
        try:
            if filtro_b1:
                print(f"\nComparando '{filtro_b1}' (Carpeta 1)  vs  '{filtro_b2}' (Carpeta 2)\n")
                mod_a = read_module_file(str(Path(f1) / filtro_b1))
                mod_b = read_module_file(str(Path(f2) / filtro_b2))
                if mod_a.name != mod_b.name:
                    mod_b = VBAModule(name=mod_a.name, code=mod_b.code, kind=mod_b.kind)
                mods_a = {mod_a.name: mod_a}
                mods_b = {mod_a.name: mod_b}
            else:
                print(f"\nCargando módulos de: {Path(f1).name}")
                mods_a = {fp.stem: read_module_file(str(fp))
                          for fp in sorted(Path(f1).iterdir())
                          if fp.suffix.lower() in _EXTS_MOD}
                print(f"  -> {len(mods_a)} módulos encontrados")
                print(f"Cargando módulos de: {Path(f2).name}")
                mods_b = {fp.stem: read_module_file(str(fp))
                          for fp in sorted(Path(f2).iterdir())
                          if fp.suffix.lower() in _EXTS_MOD}
                print(f"  -> {len(mods_b)} módulos encontrados\n")

            la, lb = Path(f1).name, Path(f2).name
            diffs = vba_compare(mods_a, mods_b)
            render_excel(diffs, la, lb, out)

            total   = len(diffs)
            con_dif = sum(1 for d in diffs if d.status not in ("equal", "skip"))

            if filtro_b1:
                print(f"Módulo: {'con diferencias' if con_dif else 'sin diferencias'}")
            else:
                print(f"Módulos con diferencias: {con_dif} de {total} comparados")
        except Exception as e:
            sys.stdout = old_stdout
            self._log_final(f"ERROR: {e}", ok=False)
            return
        finally:
            sys.stdout = old_stdout

        self._diff_b_path = out
        self.var_diff_b.set(out)
        self.after(0, lambda: self._mostrar_resultado_b(total, con_dif, out, filtro_b1, filtro_b2))
        self._log_final(
            f"Diff guardado: {Path(out).name}  —  pulsa 'Abrir diff'" if con_dif > 0
            else f"Sin diferencias. Diff guardado: {Path(out).name}",
            ok=True)

    # ── Tab B — Acciones APLICAR ──────────────────────────────────────────────

    def _accion_aplicar_b(self, dest_var, dry_run):
        diff = self.var_diff_b.get().strip()
        dest = dest_var.get().strip()
        if not diff: self._log_final("ERROR: Selecciona el Modul_Diff.xlsx.", ok=False);  return
        if not dest: self._log_final("ERROR: Selecciona la carpeta destino.", ok=False);   return
        if not os.path.isdir(dest):
            self._log_final("ERROR: La ruta destino no es una carpeta.", ok=False); return
        self._limpiar_log()
        threading.Thread(
            target=self._thread_aplicar_b, args=(diff, dest, dry_run), daemon=True).start()

    def _thread_aplicar_b(self, diff, dest_folder, dry_run):
        writer = QueueWriter(self._log_queue)
        old_stdout = sys.stdout
        sys.stdout = writer
        try:
            modo = "SIMULACIÓN (sin escribir nada)" if dry_run else "APLICANDO cambios"
            print(f"\n{modo} -> {Path(dest_folder).name}\n")
            apply_module_folder(dest_folder, diff, dry_run=dry_run)
            if dry_run:
                print(f"\nPara ver los cambios línea a línea pulsa 'Abrir diff'.")
        except Exception as e:
            sys.stdout = old_stdout
            self._log_final(f"ERROR: {e}", ok=False)
            return
        finally:
            sys.stdout = old_stdout
        fin = ("Simulación completada  —  pulsa 'Abrir diff' para ver los detalles"
               if dry_run else f"Cambios aplicados en: {Path(dest_folder).name}")
        self._log_final(fin, ok=True)


# ── Entrada ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
