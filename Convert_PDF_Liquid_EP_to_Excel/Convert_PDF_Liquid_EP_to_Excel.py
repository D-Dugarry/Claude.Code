"""
Convert_PDF_Liquid_EP_to_Excel.py — Convierte el PDF "Listado de
Liquidación de Tasas Académicas de Matrícula" (Universidad Permanente
UPUA, Universidad de Alicante) a un fichero Excel con dos hojas: Resumen
(1 fila por alumno, importes agregados) y Detalle (1 fila por línea de
cobro + 1 fila por importe administrativo, para conciliar pago a pago).

Flujo:
  1) Seleccionar el PDF de origen.
  2) Seleccionar la carpeta destino del Excel.
  3) Pulsar "Ejecutar": parsea el PDF (pdf_parser.py), muestra la vista
     previa en las tablas Detalle y Resumen, y genera el .xlsx
     (excel_export.py) en un único paso.

Basado en la plantilla del skill tkinter-app-design (layout de 3 bloques +
panel de configuración). Vendorizado junto a este archivo: data_table.py,
help_tooltips.py, SnailSystem.png.
"""

import os
import sys
import datetime
import threading
import winreg
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

from pdf_parser import parse_pdf, corregir_alineacion, Registro
from excel_export import export_to_excel

# Tooltips + caracolillo: help_tooltips.py vive junto a este archivo.
try:
    from help_tooltips import Tooltip, attach_help_toggle, set_tooltips_enabled
    _HAS_TOOLTIPS = True
except ImportError:
    _HAS_TOOLTIPS = False

    class Tooltip:          # type: ignore[no-redef]
        def __init__(self, *a, **kw):
            pass

    def attach_help_toggle(*a, **kw):
        pass

    def set_tooltips_enabled(v):
        pass

# Tabla reutilizable: data_table.py vive junto a este archivo.
from data_table import install_treeview_style, DataTable

try:
    from PIL import Image, ImageTk
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False


def _resource_path(filename: str) -> str:
    """Ruta a un recurso, compatible con PyInstaller (_MEIPASS) y desarrollo."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)


# ── Constantes de UI ──────────────────────────────────────────────────────────

APP_TITLE = "Convertir Liquidación de Tasas (PDF → Excel)"
APP_MINW = 1100
APP_MINH = 680

BASE = 14
FONT_UI = ("Verdana", BASE)
FONT_BOLD = ("Verdana", BASE, "bold")
FONT_TITLE = ("Verdana", BASE + 1, "bold")
FONT_ENTRY = ("Verdana", BASE - 1)
FONT_LOG = ("Courier New", BASE)
FONT_TABLA = ("Verdana", BASE - 1)
FONT_TABLA_BOLD = ("Verdana", BASE - 1, "bold")

SNAIL_IMG_PX = int(BASE * 2.8)

BG_APP = "#F2F3F4"
BG_SELEC = "#1A5276"   # Bloque 1 — selección (azul oscuro)
BG_TAB1 = "#7D3C98"    # Bloque 2a — vista previa Detalle (morado)
BG_TAB2 = "#1E8449"    # Bloque 2b — vista previa Resumen (verde oscuro)
BG_LOG = "#5D6D7E"     # Bloque 3 — informe (gris azulado)
BG_CFG = "#1B2631"     # Panel configuración (casi negro)

C_SELEC = "#D5D8DC"
C_TODOS = "#AED6F1"
C_NINGUNO = "#F5B7B1"
C_ACCION = "#F0B27A"
C_EXPORTAR = "#A9DFBF"
C_CONFIG = "#AAB7B8"

# Filas con posible desalineación DNI/Apellido pendientes de revisar (ámbar)
C_ROW_REVISAR = "#FDEBD0"      # fondo de fila marcada
C_ROW_REVISAR_SEL = "#F5CBA7"  # fondo de fila marcada seleccionada
C_AVISO = "#F39C12"            # ámbar del texto/botón de aviso


# ── Registro Windows (persistencia de configuración) ──────────────────────────

REG_KEY = r"Software\Dugarry\Convert_PDF_Liquid_EP_to_Excel"

_CFG_KEYS: dict[str, str] = {
    "nombre_fichero": "CfgNombreFichero",
    "abrir_carpeta": "CfgAbrirCarpeta",
}


def _reg_read(name: str) -> str:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY) as k:
            v, _ = winreg.QueryValueEx(k, name)
            return str(v)
    except FileNotFoundError:
        return ""


def _reg_write(name: str, value: str) -> None:
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY) as k:
        winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)


def _cfg_read(key: str) -> str:
    return _reg_read(_CFG_KEYS.get(key, key))


def _cfg_write(key: str, value: str) -> None:
    _reg_write(_CFG_KEYS.get(key, key), value)


# ── Columnas de los Treeview ────────────────────────────────────────────────

# Espejo exacto de las columnas de la hoja "Detalle" del Excel: una fila
# por línea de cobro más una fila por alumno con el importe administrativo
# (esa fila deja vacías Referencia/Fecha/Importe/Plazo/Forma Pago).
COL_DETALLE_IDS = ("exped", "dni", "nombre", "referencia", "fecha",
                    "importe", "imp_adm", "plazo", "forma_pago")
COL_DETALLE_NAMES = ("Exp.", "DNI", "Apellidos y Nombre", "Referencia",
                      "F. Cobro", "I.Acad.", "I.Adm.", "Rec.", "F. Pag.")

COL_RESUMEN_IDS = ("exped", "dni", "nombre", "importe", "neto",
                    "administrativo")
COL_RESUMEN_NAMES = ("Exped", "DNI", "Apellidos y Nombre", "Importe",
                      "I.Acad.", "I.Adm.")


# ══════════════════════════════════════════════════════════════════════════════
# Clase App
# ══════════════════════════════════════════════════════════════════════════════

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(True, True)
        self.minsize(APP_MINW, APP_MINH)
        self.configure(bg=BG_APP)

        self._fuente1_var = tk.StringVar()   # ruta del PDF seleccionado
        self._fuente2_var = tk.StringVar()   # carpeta destino del Excel
        self._config_dirty = False

        self._pdf_path: str | None = None
        self._ultimo_excel: str | None = None
        self._procesando = False

        # Revisión de filas desalineadas DNI/Apellido
        self._registros: list[Registro] = []
        self._iid2reg_res: dict[str, Registro] = {}   # iid Resumen -> Registro
        self._iid_revisar: set[str] = set()           # iids pendientes de revisar

        self._snail_img = None
        if _HAS_PIL:
            try:
                _src = Image.open(_resource_path("SnailSystem.png")).convert("RGBA")
                _src = _src.resize((SNAIL_IMG_PX, SNAIL_IMG_PX), Image.LANCZOS)
                self._snail_img = ImageTk.PhotoImage(_src)
            except Exception:
                pass

        self._build_styles()
        self._build_ui()
        self.update_idletasks()
        self.state("zoomed")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        last_dest = _reg_read("LastCarpetaDestino")
        if last_dest and os.path.isdir(last_dest):
            self._fuente2_var.set(last_dest)

        set_tooltips_enabled(True)

    # ── Estilos ttk ──────────────────────────────────────────────────────────

    def _build_styles(self) -> None:
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", font=FONT_UI, background=BG_APP)
        s.configure("TFrame", background=BG_APP)
        s.configure("TEntry", font=FONT_UI, fieldbackground="white")
        s.configure("TScrollbar", background=BG_APP)
        s.configure("Sel.TButton", font=FONT_UI, background=C_SELEC, padding=(8, 1))
        s.map("Sel.TButton", background=[("active", "#BFC9CA")])
        s.configure("Todos.TButton", font=FONT_UI, background=C_TODOS, padding=(8, 1))
        s.map("Todos.TButton", background=[("active", "#85C1E9")])
        s.configure("Ngno.TButton", font=FONT_UI, background=C_NINGUNO, padding=(8, 1))
        s.map("Ngno.TButton", background=[("active", "#F1948A")])
        s.configure("Accion.TButton", font=FONT_BOLD, background=C_ACCION, padding=(8, 1))
        s.map("Accion.TButton", background=[("active", "#E59866")])
        s.configure("Export.TButton", font=FONT_BOLD, background=C_EXPORTAR, padding=(8, 1))
        s.map("Export.TButton", background=[("active", "#7DCEA0")])
        s.configure("Config.TButton", font=FONT_UI, background=C_CONFIG, padding=(8, 1))
        s.map("Config.TButton", background=[("active", "#808B96")])

        install_treeview_style(self, sep_color=BG_TAB1,
                                font_ui=FONT_TABLA, font_bold=FONT_TABLA_BOLD)

    # ── Helper: bloque con cabecera coloreada ─────────────────────────────────

    def _bloque(self, titulo: str, bg_titulo: str,
                expand: bool = False,
                parent: tk.Widget | None = None,
                content_padx: int = 12,
                header_h: int | None = None) -> tuple[tk.Frame, tk.Frame]:
        p = parent if parent is not None else self
        outer = tk.Frame(p, bg=BG_APP, bd=1, relief="solid",
                          highlightbackground="#AAAAAA", highlightthickness=1)
        outer.pack(fill="both" if expand else "x",
                   expand=expand, padx=10, pady=(6, 0))
        tb = tk.Frame(outer, bg=bg_titulo)
        if header_h is not None:
            tb.configure(height=header_h)
            tb.pack_propagate(False)
        tb.pack(fill="x")
        tb.rowconfigure(0, weight=1)
        tk.Label(tb, text=f"  {titulo}",
                 bg=bg_titulo, fg="white", font=FONT_TITLE, anchor="w"
                 ).grid(row=0, column=0, sticky="w")
        c = tk.Frame(outer, bg=BG_APP, padx=content_padx, pady=8)
        c.pack(fill="both", expand=True)
        c.columnconfigure(1, weight=1)
        return tb, c

    # ── Construcción de la UI ─────────────────────────────────────────────────

    def _build_ui(self) -> None:
        sel_tb = self._build_seleccion()
        self.update_idletasks()
        self._header_h = sel_tb.winfo_reqheight()
        self._build_tablas()
        self._build_informe()
        tk.Frame(self, bg=BG_APP, height=8).pack()

    # ── BLOQUE 1: Selección ───────────────────────────────────────────────────

    def _build_seleccion(self) -> tk.Frame:
        tb, c = self._bloque("Selección", BG_SELEC)

        tb.columnconfigure(1, weight=1)
        tk.Label(tb, text="", bg=BG_SELEC).grid(row=0, column=1, sticky="ew")

        _bd = SNAIL_IMG_PX
        self._logo_lbl = tk.Canvas(tb, width=_bd, height=_bd, bg=BG_SELEC,
                                    highlightthickness=0, bd=0, cursor="hand2")
        if self._snail_img:
            self._logo_lbl.create_image(_bd // 2, _bd // 2, image=self._snail_img)
        else:
            self._logo_lbl.create_text(_bd // 2, _bd // 2, text="🐚",
                                        font=("Verdana", BASE + 2))
        self._logo_lbl.grid(row=0, column=2, padx=(0, 10))

        attach_help_toggle(
            self._logo_lbl,
            signature="Dugarry'26",
            hint_on="Ocultar textos explicativos de los botones",
            hint_off="Mostrar textos explicativos de los botones",
            on_secondary=self._show_ayuda)

        # ── PDF a convertir ─────────────────────────────────────────────────
        tk.Label(c, text="PDF a convertir:", bg=BG_APP, font=FONT_UI, anchor="w"
                 ).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(4, 2))
        tk.Entry(c, textvariable=self._fuente1_var, font=FONT_ENTRY,
                 bg="white", fg="#222222", relief="sunken", bd=1, state="readonly"
                 ).grid(row=0, column=1, sticky="ew", pady=(4, 2), ipady=4)
        _btn1 = ttk.Button(c, text="Seleccionar", style="Sel.TButton",
                           command=self._pick_pdf)
        _btn1.grid(row=0, column=2, padx=(8, 0), pady=(4, 2), ipadx=4)
        Tooltip(_btn1, "Selecciona el PDF de liquidación de tasas a convertir")

        # ── Carpeta destino del Excel ─────────────────────────────────────
        tk.Label(c, text="Carpeta destino Excel:", bg=BG_APP, font=FONT_UI, anchor="w"
                 ).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(2, 4))
        tk.Entry(c, textvariable=self._fuente2_var, font=FONT_ENTRY,
                 bg="white", fg="#222222", relief="sunken", bd=1
                 ).grid(row=1, column=1, sticky="ew", pady=(2, 4), ipady=4)
        _btn2 = ttk.Button(c, text="Seleccionar", style="Sel.TButton",
                           command=self._pick_carpeta_destino)
        _btn2.grid(row=1, column=2, padx=(8, 0), pady=(2, 4), ipadx=4)
        Tooltip(_btn2, "Selecciona la carpeta donde se guardará el Excel generado")

        return tb

    # ── BLOQUE 2: Tablas ──────────────────────────────────────────────────────

    def _build_tablas(self) -> None:
        self._tables_frame = tk.Frame(self, bg=BG_APP)
        self._tables_frame.pack(fill="both", expand=True)

        _tbl_grid = tk.Frame(self._tables_frame, bg=BG_APP)
        _tbl_grid.pack(fill="both", expand=True)
        # Detalle tiene más columnas (9) que Resumen (6): más ancho relativo.
        _tbl_grid.columnconfigure(0, weight=3)   # vista previa Detalle
        _tbl_grid.columnconfigure(1, weight=2)   # vista previa Resumen
        _tbl_grid.rowconfigure(0, weight=1)

        _col_tab1 = tk.Frame(_tbl_grid, bg=BG_APP)
        _col_tab1.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        _col_tab2 = tk.Frame(_tbl_grid, bg=BG_APP)
        _col_tab2.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        # ── Tabla 1: vista previa (Detalle) ────────────────────────────────
        tb1, c1 = self._bloque("Vista previa (Detalle)", BG_TAB1,
                                expand=True, parent=_col_tab1,
                                header_h=self._header_h)
        tb1.columnconfigure(0, weight=1)

        self._lbl_count1 = tk.Label(tb1, text="", bg=BG_TAB1, fg="white", font=FONT_UI)
        self._lbl_count1.grid(row=0, column=1, padx=(0, 6))

        c1.rowconfigure(0, weight=1)
        c1.columnconfigure(0, weight=1)
        self._tree_detalle = DataTable(
            c1, COL_DETALLE_IDS, COL_DETALLE_NAMES,
            font_ui=FONT_TABLA, font_bold=FONT_TABLA_BOLD,
            on_select=lambda s, t, lbl=self._lbl_count1:
                lbl.configure(text=f"{s}/{t}" if t else ""))
        self._tree_detalle.pack(fill="both", expand=True)
        self._tree_detalle.set_anchor("importe", "e")
        self._tree_detalle.set_anchor("imp_adm", "e")
        self._tree_detalle.set_anchor("plazo", "center")
        self._tree_detalle.set_anchor("forma_pago", "center")

        # ── Tabla 2: vista previa (Resumen) ────────────────────────────────
        tb2, c2 = self._bloque("Vista previa (Resumen)", BG_TAB2,
                                expand=True, parent=_col_tab2,
                                header_h=self._header_h)
        tb2.columnconfigure(0, weight=1)

        # Controles de revisión de desalineación DNI/Apellido (aviso en la
        # barra de título del bloque). Ocultos mientras no haya filas marcadas.
        self._frm_revisar = tk.Frame(tb2, bg=BG_TAB2)
        self._frm_revisar.grid(row=0, column=1, padx=(0, 10))
        self._lbl_revisar = tk.Label(
            self._frm_revisar, text="", bg=C_AVISO, fg="#1B2631",
            font=FONT_TABLA_BOLD, padx=8, pady=1)
        self._lbl_revisar.pack(side="left", padx=(0, 6))
        _bopts = dict(bg=C_AVISO, fg="#1B2631", activebackground="#E67E22",
                      relief="raised", bd=1, font=FONT_TABLA_BOLD, cursor="hand2")
        tk.Button(self._frm_revisar, text="◀", width=2,
                  command=lambda: self._nav_revisar(-1), **_bopts
                  ).pack(side="left", padx=1)
        tk.Button(self._frm_revisar, text="Revisar…",
                  command=self._abrir_revision, **_bopts
                  ).pack(side="left", padx=1)
        tk.Button(self._frm_revisar, text="▶", width=2,
                  command=lambda: self._nav_revisar(+1), **_bopts
                  ).pack(side="left", padx=1)
        self._frm_revisar.grid_remove()

        self._lbl_count2 = tk.Label(tb2, text="", bg=BG_TAB2, fg="white", font=FONT_UI)
        self._lbl_count2.grid(row=0, column=2, padx=(0, 6))

        c2.rowconfigure(0, weight=1)
        c2.columnconfigure(0, weight=1)
        self._tree_resumen = DataTable(
            c2, COL_RESUMEN_IDS, COL_RESUMEN_NAMES,
            font_ui=FONT_TABLA, font_bold=FONT_TABLA_BOLD,
            on_select=lambda s, t, lbl=self._lbl_count2:
                lbl.configure(text=f"{s}/{t}" if t else ""))
        self._tree_resumen.pack(fill="both", expand=True)
        self._tree_resumen.set_anchor("importe", "e")
        self._tree_resumen.set_anchor("administrativo", "e")
        self._tree_resumen.set_anchor("neto", "e")
        self._tree_resumen.set_tag("revisar", C_ROW_REVISAR, C_ROW_REVISAR_SEL)
        self._tree_detalle.set_tag("revisar", C_ROW_REVISAR, C_ROW_REVISAR_SEL)

    # ── BLOQUE 3: Informe ─────────────────────────────────────────────────────

    def _build_informe(self) -> None:
        tb, c = self._bloque("Informe", BG_LOG, content_padx=4,
                              header_h=self._header_h)
        tb.columnconfigure(0, weight=1)
        tb.columnconfigure(3, weight=1)

        self._btn_accion = ttk.Button(tb, text="▶  Ejecutar", style="Export.TButton",
                                      command=self._ejecutar)
        self._btn_accion.grid(row=0, column=1, padx=(0, 100), pady=2, ipadx=6)
        Tooltip(self._btn_accion,
                "Parsea el PDF seleccionado y genera el Excel de una vez")

        self._btn_abrir = ttk.Button(tb, text="📂 Abrir Excel", style="Sel.TButton",
                                     command=self._abrir_excel_generado)
        self._btn_abrir.grid(row=0, column=2, padx=(0, 4), pady=2, ipadx=6)
        Tooltip(self._btn_abrir, "Abre el último Excel generado en esta sesión")

        tk.Label(tb, text="", bg=BG_LOG).grid(row=0, column=3, sticky="ew")

        _btn_cfg = ttk.Button(tb, text="⚙  Configuración", style="Config.TButton",
                               command=self._show_config)
        _btn_cfg.grid(row=0, column=4, padx=(0, 8), pady=2, ipadx=4)
        Tooltip(_btn_cfg, "Abre el panel de configuración de la aplicación")

        c.rowconfigure(0, weight=1)
        c.columnconfigure(0, weight=1)
        c.columnconfigure(1, weight=0)

        log_frame = tk.Frame(c, bg="#1E1E1E")
        log_frame.grid(row=0, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        # Alto reducido un 30% respecto al valor original (6 -> 4 líneas).
        self._log = tk.Text(
            log_frame, height=4, state="disabled", font=FONT_LOG, wrap="word",
            spacing1=2, spacing3=2,
            bg="#1E1E1E", fg="#D4D4D4", insertbackground="white",
            relief="flat", bd=0)
        log_vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self._log.yview)
        self._log.configure(yscrollcommand=log_vsb.set)
        self._log.grid(row=0, column=0, sticky="nsew")
        log_vsb.grid(row=0, column=1, sticky="ns")
        self._log.bind("<MouseWheel>",
                       lambda e: self._log.yview_scroll(
                           -1 if e.delta > 0 else 1, "units"))

        self._log.tag_configure("ok", foreground="#4EC94E")
        self._log.tag_configure("error", foreground="#FF6B6B")
        self._log.tag_configure("info", foreground="#85C1E9")
        self._log.tag_configure("warn", foreground="#F0B27A")

        _lf = tkfont.Font(family=FONT_UI[0], size=FONT_UI[1])
        _lw = _lf.metrics("linespace") + 14
        _lcvs = tk.Canvas(c, width=_lw, bg=C_ACCION,
                          highlightthickness=1, highlightbackground="#D4895A",
                          cursor="hand2", relief="flat")
        _lcvs.grid(row=0, column=1, sticky="ns", padx=(3, 0))
        _ltxt = _lcvs.create_text(_lw // 2, 50, text="Limpiar",
                                   angle=270, font=FONT_UI, fill="#222222")
        _lcvs.bind("<Configure>",
                   lambda e, cv=_lcvs, t=_ltxt, w=_lw:
                       cv.coords(t, w // 2, e.height // 2))
        _lcvs.bind("<Enter>", lambda e: _lcvs.configure(bg="#E59866"))
        _lcvs.bind("<Leave>", lambda e: _lcvs.configure(bg=C_ACCION))
        _lcvs.bind("<Button-1>", lambda e: self._log_clear())
        Tooltip(_lcvs, "Borra el contenido del panel Informe")

    # ── Panel de configuración (overlay lateral) ──────────────────────────────

    def _show_config(self) -> None:
        if hasattr(self, "_cfg_panel") and self._cfg_panel.winfo_ismapped():
            self._hide_config()
            return
        if not hasattr(self, "_cfg_panel"):
            self._build_config_panel()
        self._cfg_panel.place(relx=1.0, rely=0.0, anchor="ne", relwidth=0.34,
                              relheight=0.98, x=-10, y=8)
        self._cfg_panel.lift()

    def _hide_config(self) -> None:
        if hasattr(self, "_cfg_panel"):
            self._cfg_panel.place_forget()

    def _build_config_panel(self) -> None:
        p = tk.Frame(self, bg=BG_CFG, bd=2, relief="ridge")
        self._cfg_panel = p

        hdr = tk.Frame(p, bg=BG_CFG)
        hdr.pack(fill="x", padx=10, pady=(10, 4))
        tk.Label(hdr, text="⚙  Configuración", bg=BG_CFG, fg="white",
                 font=FONT_BOLD).pack(side="left")
        ttk.Button(hdr, text="✕ Cerrar", style="Config.TButton",
                   command=self._hide_config).pack(side="right")

        tk.Frame(p, bg="#4A5568", height=1).pack(fill="x", padx=10, pady=(0, 8))

        inner = tk.Frame(p, bg=BG_CFG)
        inner.pack(fill="both", expand=True)

        _FG = "white"
        _PAD = {"padx": 14, "pady": (6, 2)}

        tk.Label(inner, text="Nombre del fichero Excel (sin extensión):",
                 bg=BG_CFG, fg=_FG, font=FONT_UI, anchor="w"
                 ).pack(fill="x", **_PAD)
        _default_nombre = _cfg_read("nombre_fichero") or "Liquidacion_Tasas"
        self._var_nombre_fichero = tk.StringVar(value=_default_nombre)
        ttk.Entry(inner, textvariable=self._var_nombre_fichero, font=FONT_ENTRY
                  ).pack(fill="x", padx=14, pady=(0, 4))

        self._var_abrir_carpeta = tk.BooleanVar(
            value=(_cfg_read("abrir_carpeta") != "0"))
        tk.Checkbutton(inner, text="Abrir la carpeta destino al terminar",
                       variable=self._var_abrir_carpeta, bg=BG_CFG, fg=_FG,
                       selectcolor=BG_CFG, activebackground=BG_CFG,
                       activeforeground=_FG, font=FONT_UI, anchor="w"
                       ).pack(fill="x", padx=14, pady=(6, 2))

        btns = tk.Frame(inner, bg=BG_CFG)
        btns.pack(fill="x", padx=14, pady=12)
        _b_save = ttk.Button(btns, text="¡Guardar!", style="Export.TButton",
                             command=self._save_config)
        _b_save.pack(side="left", ipadx=10)
        Tooltip(_b_save, "Guarda la configuración en el Registro de Windows")
        _b_rev = ttk.Button(btns, text="Revertir", style="Sel.TButton",
                            command=self._revert_config)
        _b_rev.pack(side="left", padx=(8, 0), ipadx=6)
        Tooltip(_b_rev, "Descarta los cambios y restaura los valores guardados")

    def _save_config(self) -> None:
        _cfg_write("nombre_fichero", self._var_nombre_fichero.get().strip())
        _cfg_write("abrir_carpeta", "1" if self._var_abrir_carpeta.get() else "0")
        self._config_dirty = False
        self._log_write("Configuración guardada.", "ok")

    def _revert_config(self) -> None:
        self._var_nombre_fichero.set(_cfg_read("nombre_fichero") or "Liquidacion_Tasas")
        self._var_abrir_carpeta.set(_cfg_read("abrir_carpeta") != "0")
        self._config_dirty = False

    # ── Panel de ayuda ──────────────────────────────────────────────────────

    def _show_ayuda(self) -> None:
        messagebox.showinfo(
            "Ayuda",
            "1) Selecciona el PDF de liquidación de tasas.\n"
            "2) Selecciona la carpeta donde quieres guardar el Excel.\n"
            "3) Pulsa Ejecutar.\n\n"
            "Se genera un único Excel con dos hojas: 'Resumen' (una fila "
            "por alumno, con importes agregados) y 'Detalle' (una fila "
            "por línea de cobro más una fila con el importe "
            "administrativo de cada alumno).")

    # ── Selección de fuentes ──────────────────────────────────────────────────

    def _pick_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar PDF de liquidación de tasas",
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._pdf_path = path
        self._fuente1_var.set(path)
        self._log_write(f"PDF seleccionado: {path}", "info")
        self._ejecutar()

    def _pick_carpeta_destino(self) -> None:
        path = filedialog.askdirectory(title="Seleccionar carpeta destino del Excel")
        if path:
            self._fuente2_var.set(path)
            _reg_write("LastCarpetaDestino", path)
            self._log_write(f"Carpeta destino: {path}", "info")

    # ── Carga de tablas ───────────────────────────────────────────────────────

    def _load_tabla_detalle(self, registros: list[Registro]) -> None:
        rows, tags = [], []
        i = 0

        def _tag(marcada: bool, pos: int) -> str:
            return "revisar" if marcada else ("row_a" if pos % 2 == 0 else "row_b")

        for r in registros:
            for ref in r.referencias:
                rows.append((r.exped, r.dni, r.nombre, ref.referencia,
                             ref.fecha, f"{ref.importe:.2f}", "",
                             ref.plazo, ref.forma_pago))
                tags.append(_tag(r.revisar, i))
                i += 1
            # Fila con el importe administrativo: sin datos de cobro.
            rows.append((r.exped, r.dni, r.nombre, "", "", "",
                        f"{r.administrativo:.2f}", "", ""))
            tags.append(_tag(r.revisar, i))
            i += 1
        self._tree_detalle.load(rows, tags=tags)

    def _load_tabla_resumen(self, registros: list[Registro]) -> None:
        rows, tags = [], []
        for i, r in enumerate(registros):
            rows.append((r.exped, r.dni, r.nombre, f"{r.importe:.2f}",
                         f"{r.importe_neto:.2f}", f"{r.administrativo:.2f}"))
            tags.append("revisar" if r.revisar
                        else ("row_a" if i % 2 == 0 else "row_b"))
        iids = self._tree_resumen.load(rows, tags=tags)
        self._iid2reg_res = dict(zip(iids, registros))
        self._iid_revisar = {iid for iid, r in zip(iids, registros) if r.revisar}

    def _resumen_row(self, r: Registro) -> tuple:
        """Fila de la tabla Resumen para un Registro (mismo orden de columnas)."""
        return (r.exped, r.dni, r.nombre, f"{r.importe:.2f}",
                f"{r.importe_neto:.2f}", f"{r.administrativo:.2f}")

    # ── Revisión de desalineación DNI/Apellido ─────────────────────────────────

    def _post_correccion(self, n_rev: int) -> None:
        """Tras parsear+corregir: refresca el contador de filas a revisar y
        muestra u oculta los controles de revisión en la barra de título."""
        self._actualizar_contador_revisar()
        if n_rev:
            self._log_write(
                f"⚠ {n_rev} fila(s) con posible desalineación DNI/Apellido: "
                "separadas por heurística, revísalas con los botones ◀ Revisar ▶ "
                "de la Vista Resumen.", "warn")

    def _actualizar_contador_revisar(self) -> None:
        pendientes = sum(1 for r in self._registros if r.revisar)
        if pendientes:
            self._lbl_revisar.configure(text=f"⚠ Filas a revisar: {pendientes}")
            self._frm_revisar.grid()
        else:
            self._frm_revisar.grid_remove()

    def _marcadas_visual(self) -> list[str]:
        """iids de las filas marcadas (pendientes), en el orden visual actual."""
        return [iid for iid in self._tree_resumen.tree.get_children()
                if iid in self._iid_revisar]

    def _nav_revisar(self, paso: int) -> None:
        """Desplaza el foco a la siguiente/anterior fila marcada, centrándola."""
        marc = self._marcadas_visual()
        if not marc:
            return
        foco = self._tree_resumen.tree.focus()
        if foco in marc:
            pos = (marc.index(foco) + paso) % len(marc)
        else:
            pos = 0 if paso >= 0 else len(marc) - 1
        iid = marc[pos]
        self._tree_resumen.tree.selection_set(iid)
        self._tree_resumen.tree.focus(iid)
        self._tree_resumen.scroll_to(iid, center=True)

    def _abrir_revision(self) -> None:
        """Abre el diálogo de revisión sobre la fila marcada con foco (o la
        primera marcada si no hay ninguna con foco)."""
        marc = self._marcadas_visual()
        if not marc:
            return
        foco = self._tree_resumen.tree.focus()
        iid = foco if foco in marc else marc[0]
        self._tree_resumen.scroll_to(iid, center=True)
        self._dialog_revision(iid, self._iid2reg_res[iid])

    def _dialog_revision(self, iid: str, reg: Registro) -> None:
        """Diálogo modal para validar/ajustar el corte DNI|Apellido de una
        fila. Muestra el texto crudo pegado con el corte propuesto; al hacer
        clic en un carácter se marca dónde empieza el apellido, con vista
        previa en vivo. Al aceptar, aplica el corte, desmarca la fila,
        reordena la tabla, actualiza el contador y regenera el Excel."""
        raw = reg.dni_raw
        # Corte inicial: el sugerido por la heurística o, si no hay, el inicio
        # de la última racha de letras (punto de partida razonable a ajustar).
        if reg.split_idx is not None:
            idx0 = reg.split_idx
        else:
            idx0 = len(raw)
            while idx0 > 0 and (raw[idx0 - 1].isalpha() or raw[idx0 - 1] == "�"):
                idx0 -= 1
        sel = tk.IntVar(value=idx0)

        # Vecinos (contexto alfabético) según el orden visual de la tabla.
        prev_iid = self._tree_resumen.tree.prev(iid)
        next_iid = self._tree_resumen.tree.next(iid)
        prev_txt = self._tree_resumen.tree.set(prev_iid, "nombre") if prev_iid else "—"
        next_txt = self._tree_resumen.tree.set(next_iid, "nombre") if next_iid else "—"

        dlg = tk.Toplevel(self)
        dlg.title("Revisar separación DNI / Apellido")
        dlg.configure(bg=BG_APP)
        dlg.transient(self)
        dlg.resizable(False, False)

        tk.Frame(dlg, bg=C_AVISO, height=6).pack(fill="x")
        cont = tk.Frame(dlg, bg=BG_APP, padx=18, pady=14)
        cont.pack(fill="both", expand=True)

        tk.Label(cont, text="Revisar separación DNI / Apellido", bg=BG_APP,
                 fg="#1B2631", font=FONT_BOLD, anchor="w").pack(fill="x")
        tk.Label(cont, text="Haz clic en el carácter donde empieza el APELLIDO.",
                 bg=BG_APP, fg="#555555", font=FONT_ENTRY, anchor="w"
                 ).pack(fill="x", pady=(2, 10))

        # Fila de caracteres clicables
        chars = tk.Frame(cont, bg="white", bd=1, relief="sunken", padx=6, pady=6)
        chars.pack(fill="x")
        char_lbls: list[tk.Label] = []
        for k, ch in enumerate(raw):
            lb = tk.Label(chars, text=ch, font=("Consolas", BASE + 2),
                          padx=2, pady=2, cursor="hand2")
            lb.pack(side="left")
            lb.bind("<Button-1>", lambda e, k=k: sel.set(k))
            char_lbls.append(lb)

        prev_row = tk.Label(cont, bg=BG_APP, fg="#555555", font=FONT_ENTRY,
                            anchor="w", justify="left")
        prev_row.pack(fill="x", pady=(12, 0))
        lbl_dni = tk.Label(cont, bg=BG_APP, fg="#1B2631", font=FONT_UI, anchor="w")
        lbl_dni.pack(fill="x", pady=(8, 0))
        lbl_nom = tk.Label(cont, bg=BG_APP, fg="#1B2631", font=FONT_UI, anchor="w")
        lbl_nom.pack(fill="x", pady=(2, 0))

        def _reconstruir(idx: int) -> tuple[str, str]:
            frag = raw[idx:]
            base = reg.nombre_raw
            if base.lstrip().startswith(","):
                nombre = frag + base
            elif frag:
                nombre = f"{frag} {base}"
            else:
                nombre = base
            return raw[:idx], nombre

        def _redibujar(*_a) -> None:
            idx = sel.get()
            for k, lb in enumerate(char_lbls):
                if k < idx:
                    lb.configure(bg="#D6EAF8", fg="#1B4F72")   # DNI (azul)
                else:
                    lb.configure(bg="#D5F5E3", fg="#1E6B3A")   # apellido (verde)
            dni, nombre = _reconstruir(idx)
            lbl_dni.configure(text=f"DNI:                        {dni}")
            lbl_nom.configure(text=f"Apellidos y Nombre:  {nombre}")
            prev_row.configure(
                text=f"Contexto (orden alfabético):\n   anterior:  {prev_txt}\n"
                     f"   siguiente: {next_txt}")

        sel.trace_add("write", _redibujar)
        _redibujar()

        btns = tk.Frame(cont, bg=BG_APP)
        btns.pack(fill="x", pady=(16, 0))

        def _aceptar() -> None:
            idx = sel.get()
            antes = (reg.dni, reg.nombre)
            reg.aplicar_split(idx)
            reg.revisar = False
            dlg.destroy()
            self._aplicar_revision(iid, reg, cambiado=(antes != (reg.dni, reg.nombre)))

        ttk.Button(btns, text="Aceptar", style="Export.TButton",
                   command=_aceptar).pack(side="right")
        ttk.Button(btns, text="Cancelar", style="Config.TButton",
                   command=dlg.destroy).pack(side="right", padx=(0, 8))

        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        dlg.grab_set()
        dlg.wait_window()

    def _aplicar_revision(self, iid: str, reg: Registro, cambiado: bool) -> None:
        """Aplica el resultado de una revisión: actualiza la fila en las dos
        tablas, la desmarca, reordena Resumen alfabéticamente, refresca el
        contador y (si cambió el dato) regenera el Excel."""
        self._iid_revisar.discard(iid)
        self._tree_resumen.set_values(iid, self._resumen_row(reg))
        self._tree_resumen.unmark(iid)
        self._tree_resumen.sort("nombre")                 # reordena tras revisar
        self._load_tabla_detalle(self._registros)         # refresca marcas/valores
        self._actualizar_contador_revisar()

        if cambiado and self._ultimo_excel:
            try:
                export_to_excel(self._registros, self._ultimo_excel)
                self._log_write(
                    f"Fila revisada (Exped {reg.exped}): {reg.dni} | "
                    f"{reg.nombre}. Excel actualizado.", "ok")
            except Exception as exc:
                self._log_write(f"No se pudo regenerar el Excel: {exc}", "error")
        else:
            self._log_write(
                f"Fila revisada (Exped {reg.exped}): sin cambios.", "info")

        # Deja el foco en la siguiente fila pendiente, si queda alguna.
        if self._iid_revisar:
            self._nav_revisar(+1)

    # ── Acción principal ──────────────────────────────────────────────────────

    def _ejecutar(self) -> None:
        if self._procesando:
            self._log_write("Ya hay un proceso en marcha, espera a que termine.", "warn")
            return
        if not self._pdf_path:
            self._log_write("Selecciona un PDF.", "warn")
            return
        destino = self._fuente2_var.get().strip()
        if not destino or not os.path.isdir(destino):
            self._log_write("Selecciona una carpeta destino válida.", "warn")
            return

        nombre_base = (getattr(self, "_var_nombre_fichero", None).get().strip()
                       if hasattr(self, "_var_nombre_fichero") else "") \
            or _cfg_read("nombre_fichero") or "Liquidacion_Tasas"
        marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(destino, f"{nombre_base}_{marca}.xlsx")

        self._procesando = True
        self._btn_accion.state(["disabled"])
        self._log_write(f"Procesando {os.path.basename(self._pdf_path)}…", "info")
        threading.Thread(target=self._ejecutar_worker,
                         args=(self._pdf_path, output_path), daemon=True).start()

    def _ejecutar_worker(self, pdf_path: str, output_path: str) -> None:
        try:
            try:
                registros, paginas = parse_pdf(pdf_path)
            except Exception as exc:
                self.after(0, self._log_write, f"Error al leer el PDF: {exc}", "error")
                return

            self.after(0, self._log_write,
                      f"{os.path.basename(pdf_path)}: {paginas} páginas, "
                      f"{len(registros)} registros.", "ok")

            if not registros:
                self.after(0, self._log_write,
                          "No se ha extraído ningún registro.", "warn")
                return

            # Detección + separación heurística de filas desalineadas
            # (DNI que se comió el inicio del apellido). No afecta importes.
            n_rev = corregir_alineacion(registros)

            export_to_excel(registros, output_path)
            self._ultimo_excel = output_path
            self._registros = registros
            self.after(0, self._load_tabla_detalle, registros)
            self.after(0, self._load_tabla_resumen, registros)
            self.after(0, self._log_write,
                      f"Excel generado: {output_path}", "ok")
            self.after(0, self._post_correccion, n_rev)
            # Confirmación modal de fin de grabación (nombre + carpeta); al
            # cerrarla abre la carpeta destino si la opción está activada.
            self.after(0, self._confirmar_grabacion, output_path)
        finally:
            self.after(0, self._fin_ejecutar)

    def _fin_ejecutar(self) -> None:
        self._procesando = False
        self._btn_accion.state(["!disabled"])

    def _abrir_carpeta_destino(self) -> None:
        destino = self._fuente2_var.get().strip()
        if destino and os.path.isdir(destino):
            os.startfile(destino)  # noqa: S606 (app de escritorio Windows)
        else:
            self._log_write("No hay carpeta destino seleccionada.", "warn")

    def _abrir_excel_generado(self) -> None:
        """Abre el último Excel generado (el botón '📂 Abrir Excel'). Antes
        abría la carpeta destino, cosa que ya hace el auto-abrir configurable
        al terminar; ahora abre directamente el fichero, que es lo útil."""
        if self._ultimo_excel and os.path.isfile(self._ultimo_excel):
            os.startfile(self._ultimo_excel)  # noqa: S606 (app de escritorio)
        else:
            self._log_write(
                "Todavía no se ha generado ningún Excel en esta sesión.", "warn")

    def _confirmar_grabacion(self, output_path: str) -> None:
        """Aviso modal de fin de grabación del Excel: confirma el nombre de
        archivo y la carpeta. Tras cerrarlo, abre la carpeta destino si la
        opción de configuración está activada."""
        nombre = os.path.basename(output_path)
        carpeta = os.path.dirname(output_path)
        messagebox.showinfo(
            "Excel generado",
            f"Se ha grabado correctamente el Excel:\n\n{nombre}\n\n"
            f"en la carpeta:\n{carpeta}",
            parent=self)
        if getattr(self, "_var_abrir_carpeta", None) is None or \
           self._var_abrir_carpeta.get():
            self._abrir_carpeta_destino()

    # ── Log ───────────────────────────────────────────────────────────────────

    def _log_write(self, msg: str, tag: str = "") -> None:
        self._log.configure(state="normal")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log.insert("end", f"[{ts}] {msg}\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _log_clear(self) -> None:
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    # ── Cierre ────────────────────────────────────────────────────────────────

    def _on_close(self) -> None:
        if self._config_dirty:
            if not messagebox.askyesno(
                    "Cambios sin guardar",
                    "Hay cambios sin guardar. ¿Salir de todos modos?"):
                return
        self.destroy()


# ── Punto de entrada ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()
