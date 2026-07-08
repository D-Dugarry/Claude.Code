"""
Convert_PDF_Liquid_EP_to_Excel.py — Convierte el PDF "Listado de
Liquidación de Tasas Académicas de Matrícula" (Universidad Permanente
UPUA, Universidad de Alicante) a un fichero Excel con dos hojas: Resumen
(1 fila por alumno, importes agregados) y Detalle (1 fila por línea de
cobro + 1 fila por importe administrativo, para conciliar pago a pago).

Flujo:
  1) Seleccionar el PDF de origen: se parsea automáticamente
     (pdf_parser.py) y se muestra la vista previa en las tablas Detalle
     y Resumen. La carpeta destino se autorrellena con la del propio PDF.
  2) Pulsar "Guardar Excel" para generar el .xlsx (excel_export.py) con
     los datos ya parseados; solo entonces se puede "Abrir Excel".

Basado en la plantilla del skill tkinter-app-design (layout de 3 bloques +
panel de configuración). Vendorizado junto a este archivo: data_table.py,
help_tooltips.py, SnailSystem.png.

Incluye además dos paneles de COMPARACIÓN (overlays a pantalla completa,
mismo patrón que el de configuración), cada uno con su botón en Selección:
  - "Comparar Resumen": selecciona dos PDF (A = anterior, B = nuevo), los
    parsea con pdf_parser y muestra en una única tabla de diff los alumnos
    cuyo importe cambió entre ambos (motor en pdf_compare.py).
  - "Comparar Detalle": muestra las tablas de Detalle de ambos PDF lado a
    lado (A izquierda, B derecha), con el mismo formato, anclajes y
    auto-ajuste que la vista previa (Detalle) de la ventana principal,
    pero sin las columnas Rec. y F. Pag. La selección se sincroniza entre
    A y B por Exped, y un checkbox permite ocultar en B los alumnos nuevos
    y colorear en ambas tablas los alumnos cuyo importe cambió.
"""

# Última actualización: 2026-07-08 13:27

import os
import sys
import datetime
import threading
import winreg
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

from pdf_parser import parse_pdf, corregir_alineacion, Registro
from pdf_compare import comparar, Comparacion, _norm_texto
from excel_export import (export_to_excel, export_rows_to_excel,
                          export_multi_sheet_to_excel)

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


def _col_indices(ids: tuple[str, ...], *names: str) -> tuple[int, ...]:
    """Índices (0-based) de `names` dentro de `ids`, en ese orden. Función
    aparte (y no una generator expression suelta en el cuerpo de la clase
    App) porque una comprensión/generator dentro de un cuerpo de clase NO ve
    las variables de esa clase salvo en el iterable más externo (gotcha de
    scoping de Python 3) — como llamada de función normal, sí las ve."""
    return tuple(ids.index(n) for n in names)


def _exe_dir() -> str:
    """Carpeta donde vive el ejecutable (PyInstaller, sys.executable) o el
    script (modo desarrollo); se usa como carpeta propuesta por defecto en
    diálogos de guardado que no dependen de una carpeta destino ya elegida."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


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

# Banda por defecto del tamaño de letra de las dos tablas. El auto-ajuste elige,
# dentro de este rango, el mayor tamaño con el que AMBAS tablas quepan enteras a
# lo ancho de su recuadro (ver App._autofit_tablas). Configurable en el panel.
FONT_TABLA_MIN_DEF = 8
FONT_TABLA_MAX_DEF = BASE      # = FONT_UI, para que no supere la letra de la UI

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

# Panel de comparación de dos PDF: colores de fila del diff
BG_COMP = "#154360"            # cabecera del panel (azul más oscuro que Selección)
C_ROW_SUBE = "#D5F5E3"         # el importe neto sube en B (verde claro)
C_ROW_SUBE_SEL = "#A9DFBF"
C_ROW_BAJA = "#FADBD8"         # el importe neto baja en B (rojo claro)
C_ROW_BAJA_SEL = "#F1948A"

# Panel Comparar Detalle: marcado fila a fila (Referencia/Imp.Adm.), no por
# alumno. "nueva" = fila presente solo en B (amarillo), "del" = fila presente
# solo en A, ya no está en B (azul). sube/baja reusan los colores de arriba.
C_ROW_NUEVA = "#FCF3CF"        # recibo/Imp.Adm. nuevo en B (amarillo claro)
C_ROW_NUEVA_SEL = "#F7DC6F"
C_ROW_DEL = "#D6EAF8"          # recibo/Imp.Adm. de A que ya no está en B (azul claro)
C_ROW_DEL_SEL = "#AED6F1"


# ── Registro Windows (persistencia de configuración) ──────────────────────────

REG_KEY = r"Software\Dugarry\Convert_PDF_Liquid_EP_to_Excel"

_CFG_KEYS: dict[str, str] = {
    "nombre_fichero": "CfgNombreFichero",
    "abrir_carpeta": "CfgAbrirCarpeta",
    "font_min": "CfgFontMin",
    "font_max": "CfgFontMax",
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


def _cfg_int(key: str, default: int) -> int:
    """Lee una clave de configuración como entero; si no existe o no es un
    entero válido, devuelve `default`."""
    try:
        return int(_cfg_read(key))
    except (TypeError, ValueError):
        return default


# ── Columnas de los Treeview ────────────────────────────────────────────────

# Espejo exacto de las columnas de la hoja "Detalle" del Excel: una fila
# por línea de cobro más una fila por alumno con el importe administrativo
# (esa fila deja vacías Referencia/Fecha/Importe/Plazo/Forma Pago).
COL_DETALLE_IDS = ("exped", "dni", "nombre", "referencia", "fecha",
                    "importe", "imp_adm", "plazo", "forma_pago")
COL_DETALLE_NAMES = ("Exp.", "DNI", "Apellidos y Nombre", "Referencia",
                      "F. Cobro", "I.Acad.", "I.Adm.", "Rec.", "F. Pag.")

COL_RESUMEN_IDS = ("exped", "dni", "nombre", "num_ref", "importe", "neto",
                    "administrativo")
COL_RESUMEN_NAMES = ("Exped", "DNI", "Apellidos y Nombre", "Rec.", "Importe",
                      "I.Acad.", "I.Adm.")

# Panel Comparar Detalle: mismas columnas centrales que la vista Detalle
# (sin Plazo ni F. Pag., que ahí se llaman "Rec."/"F. Pag."), más dos propias
# del panel: "Rec." (aquí: nº de recibo del Exped, 1/2/3... o "Adm." en la
# fila de Imp.Adm. — concepto distinto del "Rec." de la vista Detalle) antes
# de Referencia, y "Estado" al final (↑/↓ si cambió el importe de esa fila,
# "New"/"Del" si la fila solo existe en un lado).
COL_COMPDET_IDS = ("exped", "dni", "nombre", "rec") + COL_DETALLE_IDS[3:7] + ("estado",)
COL_COMPDET_NAMES = (("Exp.", "DNI", "Apellidos y Nombre", "Rec.")
                     + COL_DETALLE_NAMES[3:7] + ("Estado",))
# Índices (0-based) de columnas de importe/DNI, para exportar Comparar
# Detalle como Excel Table (formato numérico + fila de totales, igual
# criterio que COL_DIFF_MONEY_COLS/COL_DIFF_DNI_COL de Comparar Resumen).
COL_COMPDET_MONEY_COLS = _col_indices(COL_COMPDET_IDS, "importe", "imp_adm")
COL_COMPDET_DNI_COL = COL_COMPDET_IDS.index("dni")

# Hoja "Detalle A+B" (solo en la exportación, no hay tabla en pantalla):
# une A y B lado a lado en una única fila por recibo (o Imp.Adm.). Bloque de
# identificación + Rec. (compartido, un recibo de A y su pareja en B son la
# misma fila), bloque A (Referencia/F.Cobro/I.Acad./I.Adm./Estado, sufijo A),
# bloque B (ídem, sufijo B) y una columna Estado final compartida.
COL_COMPDET_UNI_IDS = (
    "exped", "dni", "nombre", "rec",
    "referencia_a", "fecha_a", "importe_a", "imp_adm_a", "estado_a",
    "referencia_b", "fecha_b", "importe_b", "imp_adm_b", "estado_b",
    "estado")
COL_COMPDET_UNI_NAMES = (
    "Exp.", "DNI", "Apellidos y Nombre", "Rec.",
    "Referencia A", "F. Cobro A", "I.Acad. A", "I.Adm. A", "Estado A",
    "Referencia B", "F. Cobro B", "I.Acad. B", "I.Adm. B", "Estado B",
    "Estado")
COL_COMPDET_UNI_MONEY_COLS = _col_indices(
    COL_COMPDET_UNI_IDS, "importe_a", "imp_adm_a", "importe_b", "imp_adm_b")
COL_COMPDET_UNI_DNI_COL = COL_COMPDET_UNI_IDS.index("dni")


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

        # Banda de tamaño de letra de las tablas (auto-ajuste dentro de ella).
        self._var_font_min = tk.IntVar(
            value=_cfg_int("font_min", FONT_TABLA_MIN_DEF))
        self._var_font_max = tk.IntVar(
            value=_cfg_int("font_max", FONT_TABLA_MAX_DEF))
        self._font_tabla_actual = FONT_TABLA[1]   # tamaño aplicado ahora mismo

        self._pdf_path: str | None = None
        self._ultimo_excel: str | None = None
        self._procesando = False
        self._guardando = False

        # Comparación de dos PDF (panel overlay a pantalla completa)
        self._comp_a_var = tk.StringVar()    # ruta del PDF A (anterior)
        self._comp_b_var = tk.StringVar()    # ruta del PDF B (nuevo)
        self._comparando = False
        self._exportando_diff = False
        self._comp_actual: Comparacion | None = None   # última comparación
        # Filtro: ocultar filas donde solo varía I.Adm. (en las muestras hay
        # un cambio sistemático de -0,60 € en casi todos los alumnos que
        # taparía los cambios reales de cobros).
        self._var_comp_solo_acad = tk.BooleanVar(value=False)
        # Filtros por Estado (columna final del diff): qué categorías mostrar.
        self._var_est_altas = tk.BooleanVar(value=True)
        self._var_est_bajas = tk.BooleanVar(value=True)
        self._var_est_cambios = tk.BooleanVar(value=True)
        self._var_est_sin = tk.BooleanVar(value=True)

        # Comparación de Detalle (segundo panel: dos tablas lado a lado)
        self._compdet_a_var = tk.StringVar()   # ruta del PDF A (anterior)
        self._compdet_b_var = tk.StringVar()   # ruta del PDF B (nuevo)
        self._comparando_det = False
        self._exportando_det = False
        # Filtros por color/Estado (fila a fila: qué categorías mostrar),
        # últimos datos parseados (para re-filtrar sin re-parsear) y mapas
        # Exped -> iids para sincronizar la selección entre las tablas A y B.
        self._var_cd_sube = tk.BooleanVar(value=True)
        self._var_cd_baja = tk.BooleanVar(value=True)
        self._var_cd_nueva = tk.BooleanVar(value=True)
        self._var_cd_del = tk.BooleanVar(value=True)
        self._var_cd_sin = tk.BooleanVar(value=True)
        self._compdet_datos: tuple | None = None
        self._exped2iid_det_a: dict[str, list[str]] = {}
        self._exped2iid_det_b: dict[str, list[str]] = {}

        # Revisión de filas desalineadas DNI/Apellido
        self._registros: list[Registro] = []
        self._iid2reg_res: dict[str, Registro] = {}   # iid Resumen -> Registro
        self._iid_revisar: set[str] = set()           # iids pendientes de revisar

        # Sincronización de selección por DNI entre Detalle y Resumen
        self._dni2iid_detalle: dict[str, list[str]] = {}   # dni -> iids Detalle
        self._dni2iid_resumen: dict[str, str] = {}         # dni -> iid Resumen
        self._sync_guard = False   # evita eco al reflejar la selección cruzada

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

        # Re-auto-ajusta las tablas cuando cambia el ANCHO de la ventana (al
        # acabar de maximizarse tras el arranque, o si el usuario la redimensiona
        # luego). Con debounce para no recalcular en cada píxel del arrastre.
        self._last_win_w = 0
        self._resize_job = None
        self.bind("<Configure>", self._on_resize)

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

        # ── Comparar dos PDF (abren los paneles de comparación) ─────────────
        _btn3 = ttk.Button(c, text="⚖ Comparar Resumen", style="Todos.TButton",
                           command=self._show_comparar)
        _btn3.grid(row=0, column=3, padx=(16, 0), pady=(4, 2),
                   ipadx=4, sticky="ew")
        Tooltip(_btn3, "Compara dos PDF de liquidación (anterior y nuevo) y "
                       "muestra los alumnos cuyo importe cambió")
        _btn4 = ttk.Button(c, text="⚖ Comparar Detalle", style="Todos.TButton",
                           command=self._show_comparar_det)
        _btn4.grid(row=1, column=3, padx=(16, 0), pady=(2, 4),
                   ipadx=4, sticky="ew")
        Tooltip(_btn4, "Compara dos PDF de liquidación mostrando la tabla de "
                       "Detalle de cada uno lado a lado (A = anterior a la "
                       "izquierda, B = nuevo a la derecha)")

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
        self._config_tabla_detalle(self._tree_detalle)

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
        self._tree_resumen.set_anchor("num_ref", "center")
        self._tree_resumen.set_anchor("importe", "e")
        self._tree_resumen.set_anchor("administrativo", "e")
        self._tree_resumen.set_anchor("neto", "e")
        self._tree_resumen.set_tag("revisar", C_ROW_REVISAR, C_ROW_REVISAR_SEL)

        # Sincronización de selección por DNI: al seleccionar una fila en una
        # tabla, se selecciona (y se hace visible) la fila / filas con el
        # mismo DNI en la otra. add="+" para no pisar el binding interno de
        # DataTable (franjas/tags de selección + contador).
        self._tree_detalle.tree.bind(
            "<<TreeviewSelect>>", self._on_sel_detalle, add="+")
        self._tree_resumen.tree.bind(
            "<<TreeviewSelect>>", self._on_sel_resumen, add="+")

    def _config_tabla_detalle(self, tabla: DataTable) -> None:
        """Anclajes y tag de revisión comunes a toda tabla con columnas de la
        vista Detalle (la principal y las dos del panel Comparar Detalle, que
        no llevan Rec./F. Pag.): aplica solo los de columnas presentes."""
        for cid, anc in (("importe", "e"), ("imp_adm", "e"),
                         ("plazo", "center"), ("forma_pago", "center"),
                         ("rec", "center"), ("estado", "center")):
            if cid in tabla.col_ids:
                tabla.set_anchor(cid, anc)
        tabla.set_tag("revisar", C_ROW_REVISAR, C_ROW_REVISAR_SEL)

    # ── BLOQUE 3: Informe ─────────────────────────────────────────────────────

    def _build_informe(self) -> None:
        tb, c = self._bloque("Informe", BG_LOG, content_padx=4,
                              header_h=self._header_h)
        tb.columnconfigure(0, weight=1)
        tb.columnconfigure(3, weight=1)

        self._btn_guardar = ttk.Button(tb, text="💾 Guardar Excel", style="Export.TButton",
                                       command=self._guardar_excel)
        self._btn_guardar.grid(row=0, column=1, padx=(0, 100), pady=2, ipadx=6)
        Tooltip(self._btn_guardar,
                "Exporta a Excel los datos ya parseados del PDF actual; abre "
                "un diálogo para elegir carpeta y nombre antes de guardar")

        self._btn_abrir = ttk.Button(tb, text="📂 Abrir Excel", style="Sel.TButton",
                                     command=self._abrir_excel_generado)
        self._btn_abrir.grid(row=0, column=2, padx=(0, 4), pady=2, ipadx=6)
        Tooltip(self._btn_abrir, "Abre el último Excel generado en esta sesión")
        self._btn_abrir.grid_remove()   # visible solo tras guardar un Excel

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

        tk.Frame(p, bg="#4A5568", height=1).pack(fill="x", padx=10, pady=(6, 6))

        tk.Label(inner, text="Tamaño de letra de las tablas:",
                 bg=BG_CFG, fg=_FG, font=FONT_UI, anchor="w"
                 ).pack(fill="x", **_PAD)
        _fila = tk.Frame(inner, bg=BG_CFG)
        _fila.pack(fill="x", padx=14, pady=(0, 2))
        tk.Label(_fila, text="Mín.", bg=BG_CFG, fg=_FG, font=FONT_UI
                 ).pack(side="left")
        ttk.Spinbox(_fila, from_=6, to=24, width=4, font=FONT_ENTRY,
                    textvariable=self._var_font_min,
                    command=self._on_font_cfg_change
                    ).pack(side="left", padx=(4, 14))
        tk.Label(_fila, text="Máx.", bg=BG_CFG, fg=_FG, font=FONT_UI
                 ).pack(side="left")
        ttk.Spinbox(_fila, from_=6, to=24, width=4, font=FONT_ENTRY,
                    textvariable=self._var_font_max,
                    command=self._on_font_cfg_change
                    ).pack(side="left", padx=(4, 0))
        tk.Label(inner,
                 text="La letra se ajusta sola dentro de este rango para que "
                      "cada tabla quepa entera en su recuadro.",
                 bg=BG_CFG, fg="#AEB6BF", font=FONT_ENTRY, anchor="w",
                 justify="left", wraplength=360
                 ).pack(fill="x", padx=14, pady=(0, 4))

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

    def _on_font_cfg_change(self) -> None:
        """El usuario tocó un Spinbox del tamaño de letra: marca dirty (el
        cambio se aplica al pulsar «¡Guardar!», que re-ejecuta el auto-ajuste)."""
        self._config_dirty = True

    def _save_config(self) -> None:
        _cfg_write("nombre_fichero", self._var_nombre_fichero.get().strip())
        _cfg_write("abrir_carpeta", "1" if self._var_abrir_carpeta.get() else "0")
        smin, smax = self._banda_font()
        _cfg_write("font_min", str(smin))
        _cfg_write("font_max", str(smax))
        self._config_dirty = False
        self._log_write("Configuración guardada.", "ok")
        # Reaplica el nuevo rango de tamaño a las tablas ya cargadas.
        if self._tree_detalle.tree.get_children():
            self._autofit_tablas()

    def _revert_config(self) -> None:
        self._var_nombre_fichero.set(_cfg_read("nombre_fichero") or "Liquidacion_Tasas")
        self._var_abrir_carpeta.set(_cfg_read("abrir_carpeta") != "0")
        self._var_font_min.set(_cfg_int("font_min", FONT_TABLA_MIN_DEF))
        self._var_font_max.set(_cfg_int("font_max", FONT_TABLA_MAX_DEF))
        self._config_dirty = False

    # ── Panel de comparación de dos PDF (overlay a pantalla completa) ──────────

    # Columnas "_fl" (flecha): estrechas, contienen solo ↑/↓ del cambio de la
    # columna numérica que las precede — separadas del número para no romper
    # su alineación a la derecha (ver _fmt_delta_num/_fmt_delta_arrow).
    COL_DIFF_IDS = ("exped", "dni", "nombre", "rec_a", "rec_b", "rec_b_fl",
                    "imp_a", "imp_b", "d_imp", "d_imp_fl",
                    "adm_a", "adm_b", "d_adm", "d_adm_fl",
                    "d_neto", "d_neto_fl", "estado")
    COL_DIFF_NAMES = ("Exped", "DNI", "Apellidos y Nombre", "Rec. A", "Rec. B", "",
                      "I.Acad. A", "I.Acad. B", "Δ Acad.", "",
                      "I.Adm. A", "I.Adm. B", "Δ Adm.", "",
                      "Δ Neto", "", "Estado")
    # Índices (0-based) de columnas de importe/recuento del diff, para la
    # exportación a Excel (Table real con formato numérico + fila de
    # totales: suma en las de importe, recuento de alumnos en DNI).
    COL_DIFF_MONEY_COLS = _col_indices(
        COL_DIFF_IDS, "imp_a", "imp_b", "d_imp", "adm_a", "adm_b",
        "d_adm", "d_neto")
    COL_DIFF_INT_COLS = _col_indices(COL_DIFF_IDS, "rec_a", "rec_b")
    COL_DIFF_DNI_COL = COL_DIFF_IDS.index("dni")

    def _nombres_export_diff(self) -> tuple[str, ...]:
        """Cabeceras de exportación del diff: igual que COL_DIFF_NAMES pero
        sin huecos. Las columnas "_fl" (flecha) llevan cabecera vacía en
        pantalla a propósito, para que DataTable.autosize() las deje
        estrechas (ver COL_DIFF_IDS); pero una Excel Table exige cabeceras
        no vacías y únicas, así que aquí se llaman "F_<columna anterior>"
        (p. ej. la flecha de "Δ Acad." pasa a "F_Δ Acad.")."""
        out: list[str] = []
        prev = ""
        for n in self.COL_DIFF_NAMES:
            if n:
                out.append(n)
                prev = n
            else:
                out.append(f"F_{prev}")
        return tuple(out)

    def _show_comparar(self) -> None:
        """Muestra el panel de comparación (mismo patrón que Configuración,
        pero ocupando toda la ventana)."""
        if not hasattr(self, "_comp_panel"):
            self._build_comparar_panel()
        self._comp_panel.place(relx=0.5, rely=0.5, anchor="center",
                               relwidth=0.99, relheight=0.98)
        self._comp_panel.lift()

    def _hide_comparar(self) -> None:
        if hasattr(self, "_comp_panel"):
            self._comp_panel.place_forget()

    def _build_comparar_panel(self) -> None:
        p = tk.Frame(self, bg=BG_APP, bd=2, relief="ridge",
                     highlightbackground="#AAAAAA", highlightthickness=1)
        self._comp_panel = p

        # Cabecera del panel
        hdr = tk.Frame(p, bg=BG_COMP)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  ⚖  Comparar Resumen de dos liquidaciones "
                           "(A = anterior, B = nuevo)",
                 bg=BG_COMP, fg="white", font=FONT_TITLE, anchor="w"
                 ).pack(side="left", pady=4)
        ttk.Button(hdr, text="✕ Cerrar", style="Config.TButton",
                   command=self._hide_comparar).pack(side="right", padx=8)
        self._btn_exportar_diff = ttk.Button(
            hdr, text="💾 Exportar Excel", style="Export.TButton",
            command=self._exportar_diff)
        self._btn_exportar_diff.pack(side="right", padx=8)
        Tooltip(self._btn_exportar_diff,
                "Exporta a Excel las filas mostradas en la tabla, con los "
                "filtros de Estado actualmente activos")

        # Selección de los dos PDF
        sel = tk.Frame(p, bg=BG_APP, padx=12, pady=8)
        sel.pack(fill="x")
        sel.columnconfigure(1, weight=1)
        for fila, (texto, var, cmd) in enumerate((
                ("PDF A (anterior):", self._comp_a_var,
                 lambda: self._pick_pdf_comparar("A")),
                ("PDF B (nuevo):", self._comp_b_var,
                 lambda: self._pick_pdf_comparar("B")))):
            tk.Label(sel, text=texto, bg=BG_APP, font=FONT_UI, anchor="w"
                     ).grid(row=fila, column=0, sticky="w", padx=(0, 10), pady=2)
            tk.Entry(sel, textvariable=var, font=FONT_ENTRY, bg="white",
                     fg="#222222", relief="sunken", bd=1, state="readonly"
                     ).grid(row=fila, column=1, sticky="ew", pady=2, ipady=4)
            ttk.Button(sel, text="Seleccionar", style="Sel.TButton", command=cmd
                       ).grid(row=fila, column=2, padx=(8, 0), pady=2, ipadx=4)

        # Resumen de la comparación (contadores + delta del total)
        self._lbl_comp_resumen = tk.Label(p, text="", bg=BG_APP,
                                          fg="#1B2631", font=FONT_BOLD,
                                          anchor="w")
        self._lbl_comp_resumen.pack(fill="x", padx=12, pady=(0, 4))

        # Filtros por Estado (qué categorías de fila se muestran en la tabla)
        fil = tk.Frame(p, bg=BG_APP)
        fil.pack(fill="x", padx=12, pady=(0, 4))
        tk.Label(fil, text="Mostrar:", bg=BG_APP, fg="#1B2631",
                 font=FONT_ENTRY).pack(side="left")
        for texto, var in (("Altas", self._var_est_altas),
                           ("Bajas", self._var_est_bajas),
                           ("Con cambios", self._var_est_cambios),
                           ("Sin cambios", self._var_est_sin)):
            tk.Checkbutton(fil, text=texto, variable=var, bg=BG_APP,
                           fg="#1B2631", activebackground=BG_APP,
                           font=FONT_ENTRY, command=self._refiltrar_diff
                           ).pack(side="left", padx=(10, 0))

        # Tabla única de diff
        cont = tk.Frame(p, bg=BG_APP, padx=12)
        cont.pack(fill="both", expand=True, pady=(0, 10))
        self._tree_diff = DataTable(
            cont, self.COL_DIFF_IDS, self.COL_DIFF_NAMES,
            font_ui=FONT_TABLA, font_bold=FONT_TABLA_BOLD)
        self._tree_diff.pack(fill="both", expand=True)
        for cid in ("imp_a", "imp_b", "d_imp", "adm_a", "adm_b",
                    "d_adm", "d_neto"):
            self._tree_diff.set_anchor(cid, "e")
        for cid in ("rec_a", "rec_b", "rec_b_fl", "d_imp_fl", "d_adm_fl",
                    "d_neto_fl", "estado"):
            self._tree_diff.set_anchor(cid, "center")
        self._tree_diff.set_tag("sube", C_ROW_SUBE, C_ROW_SUBE_SEL)
        self._tree_diff.set_tag("baja", C_ROW_BAJA, C_ROW_BAJA_SEL)
        self._tree_diff.set_tag("aviso", C_ROW_REVISAR, C_ROW_REVISAR_SEL)

        # Leyenda de colores + filtro (misma fila, bajo la tabla)
        ley = tk.Frame(p, bg=BG_APP)
        ley.pack(fill="x", padx=12, pady=(0, 8))
        _chk = tk.Checkbutton(
            ley, text="Ocultar filas donde solo cambia I.Adm.",
            variable=self._var_comp_solo_acad, bg=BG_APP, fg="#1B2631",
            activebackground=BG_APP, font=FONT_ENTRY,
            command=self._refiltrar_diff)
        _chk.pack(side="right")
        Tooltip(_chk, "Quita del diff los alumnos cuyo único cambio es el "
                      "importe administrativo (p. ej. un ajuste sistemático "
                      "aplicado a todos), dejando solo los cambios de cobros")
        for color, texto in ((C_ROW_SUBE, "el importe neto sube en B"),
                             (C_ROW_BAJA, "el importe neto baja en B"),
                             (C_ROW_REVISAR, "mismo Exped con DNI distinto "
                                             "en A y B (revisar)")):
            tk.Label(ley, text="   ", bg=color, relief="solid", bd=1
                     ).pack(side="left", padx=(0, 4))
            tk.Label(ley, text=texto, bg=BG_APP, fg="#555555", font=FONT_ENTRY
                     ).pack(side="left", padx=(0, 18))

    def _pick_pdf_comparar(self, cual: str) -> None:
        """Selecciona el PDF A o B. Cuando ya están los dos, lanza la
        comparación automáticamente (mismo criterio que _pick_pdf)."""
        path = filedialog.askopenfilename(
            title=f"Seleccionar PDF {cual} "
                  f"({'anterior' if cual == 'A' else 'nuevo'})",
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        (self._comp_a_var if cual == "A" else self._comp_b_var).set(path)
        if self._comp_a_var.get() and self._comp_b_var.get():
            self._comparar_ejecutar()

    def _comparar_ejecutar(self) -> None:
        if self._comparando:
            self._log_write("Ya hay una comparación en marcha.", "warn")
            return
        pa, pb = self._comp_a_var.get().strip(), self._comp_b_var.get().strip()
        if not pa or not pb:
            self._lbl_comp_resumen.configure(
                text="Selecciona los dos PDF a comparar.", fg=C_AVISO)
            return
        self._comparando = True
        self._lbl_comp_resumen.configure(text="Comparando…", fg="#555555")
        self._log_write(
            f"Comparando {os.path.basename(pa)} ↔ {os.path.basename(pb)}…",
            "info")
        threading.Thread(target=self._comparar_worker, args=(pa, pb),
                         daemon=True).start()

    def _comparar_worker(self, path_a: str, path_b: str) -> None:
        try:
            try:
                regs_a, pag_a = parse_pdf(path_a)
                regs_b, pag_b = parse_pdf(path_b)
            except Exception as exc:
                self.after(0, self._log_write,
                           f"Error al leer los PDF: {exc}", "error")
                self.after(0, self._lbl_comp_resumen.configure,
                           {"text": f"Error al leer los PDF: {exc}",
                            "fg": "#C0392B"})
                return
            # Corrige la desalineación DNI/Apellido antes de comparar (no
            # toca importes; mejora el casado por DNI y los nombres del diff).
            corregir_alineacion(regs_a)
            corregir_alineacion(regs_b)
            comp = comparar(regs_a, regs_b)
            self.after(0, self._log_write,
                       f"A: {pag_a} pág., {len(regs_a)} reg. · "
                       f"B: {pag_b} pág., {len(regs_b)} reg.", "ok")
            self.after(0, self._load_tabla_diff, comp)
        finally:
            self.after(0, self._fin_comparar)

    def _fin_comparar(self) -> None:
        self._comparando = False

    def _refiltrar_diff(self) -> None:
        """Reaplica el filtro sobre la última comparación (sin re-parsear)."""
        if self._comp_actual is not None:
            self._load_tabla_diff(self._comp_actual)

    def _exportar_diff(self) -> None:
        """Exporta a Excel las filas actualmente mostradas en la tabla de
        diff (según los filtros de Estado y el de 'solo cambia I.Adm.'
        activos). Propone como carpeta la del propio .exe/script y como
        nombre 'Diff_Resumen_<fecha_hora>', pero deja decidir al usuario en
        el diálogo 'Guardar como' (igual que 'Guardar Excel')."""
        if self._exportando_diff:
            self._log_write("Ya se está exportando el Excel.", "warn")
            return
        if self._comp_actual is None:
            self._log_write(
                "Todavía no hay una comparación que exportar.", "warn")
            return

        marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = filedialog.asksaveasfilename(
            title="Exportar diff de Comparar Resumen",
            initialdir=_exe_dir(), initialfile=f"Diff_Resumen_{marca}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")])
        if not output_path:
            self._log_write("Exportación a Excel cancelada.", "info")
            return

        filas = [t[1] for t in self._filas_diff_visibles(self._comp_actual)]
        self._exportando_diff = True
        self._btn_exportar_diff.state(["disabled"])
        self._log_write(f"Exportando diff a {output_path}…", "info")
        threading.Thread(target=self._exportar_diff_worker,
                         args=(filas, output_path), daemon=True).start()

    def _exportar_diff_worker(self, filas: list[tuple], output_path: str
                              ) -> None:
        try:
            try:
                export_rows_to_excel(
                    self._nombres_export_diff(), filas, output_path,
                    sheet_name="Diff Resumen", table_name="Diff_Resumen",
                    money_cols=self.COL_DIFF_MONEY_COLS,
                    int_cols=self.COL_DIFF_INT_COLS,
                    count_col=self.COL_DIFF_DNI_COL)
            except Exception as exc:
                self.after(0, self._log_write,
                          f"No se pudo exportar el Excel: {exc}", "error")
                return
            self.after(0, self._log_write,
                       f"Excel exportado: {output_path}", "ok")
        finally:
            self.after(0, self._fin_exportar_diff)

    def _fin_exportar_diff(self) -> None:
        self._exportando_diff = False
        self._btn_exportar_diff.state(["!disabled"])

    @staticmethod
    def _fmt_delta(delta: float) -> str:
        """Celda de una columna Δ: vacía si no hay cambio; si lo hay, el
        valor con signo y una flecha ↑/↓ a la derecha. Usado por el panel
        Comparar Detalle (columna Estado, texto suelto sin problema de
        alineación); la tabla de diff de Comparar Resumen usa en su lugar
        _fmt_delta_num/_fmt_delta_arrow, con la flecha en su propia columna
        estrecha para no romper la alineación a la derecha de los números."""
        if not delta:
            return ""
        return f"{delta:+.2f} {'↑' if delta > 0 else '↓'}"

    @staticmethod
    def _fmt_delta_num(delta: float) -> str:
        """Valor con signo de una columna Δ de la tabla de diff, sin flecha."""
        return f"{delta:+.2f}" if delta else ""

    @staticmethod
    def _fmt_delta_arrow(delta: float) -> str:
        """Flecha ↑/↓ de una columna Δ, para la columna estrecha contigua."""
        return ("↑" if delta > 0 else "↓") if delta else ""

    @staticmethod
    def _estado_cambios(f) -> str:
        """Texto de la columna Estado para un alumno con importe cambiado
        (FilaDiff en comp.modificados): qué importe cambió (I.Acad. y/o
        I.Adm., puede ser uno de los dos o ambos) y en qué sentido, en vez
        del genérico "Con cambios" anterior."""
        partes = []
        if f.delta_importe:
            partes.append(f"Acad. {'↑' if f.delta_importe > 0 else '↓'}")
        if f.delta_administrativo:
            partes.append(f"Adm. {'↑' if f.delta_administrativo > 0 else '↓'}")
        return " / ".join(partes)

    def _fila_diff(self, f, estado: str) -> tuple:
        """Valores de una fila del diff para un alumno presente en A y B
        (FilaDiff), con la columna Estado al final. Cada Δ (y el cambio de
        nº de recibos) va seguido de su columna "_fl" con solo la flecha."""
        delta_rec = f.num_refs_b - f.num_refs_a
        rec_b_fl = ("↑" if delta_rec > 0 else "↓") if delta_rec else ""
        return (f.exped, f.dni, f.nombre, f.num_refs_a, f.num_refs_b, rec_b_fl,
                f"{f.importe_a:.2f}", f"{f.importe_b:.2f}",
                self._fmt_delta_num(f.delta_importe),
                self._fmt_delta_arrow(f.delta_importe),
                f"{f.administrativo_a:.2f}", f"{f.administrativo_b:.2f}",
                self._fmt_delta_num(f.delta_administrativo),
                self._fmt_delta_arrow(f.delta_administrativo),
                self._fmt_delta_num(f.delta_neto),
                self._fmt_delta_arrow(f.delta_neto), estado)

    def _fila_solo(self, r: Registro, alta: bool) -> tuple:
        """Valores de una fila del diff para un alumno presente solo en un
        listado (Alta = solo en B, Baja = solo en A): las columnas del lado
        ausente van vacías y los Δ reflejan su aportación al total. En una
        Alta, Rec. B pasa de "nada" (el alumno no estaba en A) a `rec`
        recibos: es un incremento, así que lleva su flecha ↑ igual que un
        cambio de recibos entre A y B (en una Baja, Rec. B queda vacío —sin
        valor que anotar—, así que no lleva flecha)."""
        signo = 1 if alta else -1
        rec = len(r.referencias)
        imp, adm = f"{r.importe:.2f}", f"{r.administrativo:.2f}"
        d_imp = round(signo * r.importe, 2)
        d_adm = round(signo * r.administrativo, 2)
        d_neto = round(signo * r.importe_neto, 2)
        if alta:
            rec_b_fl = "↑" if rec else ""
            return (r.exped, r.dni, r.nombre, "", rec, rec_b_fl,
                    "", imp, self._fmt_delta_num(d_imp),
                    self._fmt_delta_arrow(d_imp),
                    "", adm, self._fmt_delta_num(d_adm),
                    self._fmt_delta_arrow(d_adm),
                    self._fmt_delta_num(d_neto), self._fmt_delta_arrow(d_neto),
                    "Alta")
        return (r.exped, r.dni, r.nombre, rec, "", "",
                imp, "", self._fmt_delta_num(d_imp),
                self._fmt_delta_arrow(d_imp),
                adm, "", self._fmt_delta_num(d_adm),
                self._fmt_delta_arrow(d_adm),
                self._fmt_delta_num(d_neto), self._fmt_delta_arrow(d_neto),
                "Baja")

    def _filas_diff_visibles(self, comp: Comparacion
                             ) -> list[tuple[str, tuple, str]]:
        """Filas del diff que corresponden a los filtros por Estado y el de
        'solo cambia I.Adm.' actualmente activos, ordenadas alfabéticamente
        por nombre. Devuelve (clave de orden, valores de fila, tag de color);
        la usan tanto _load_tabla_diff (tabla en pantalla) como _exportar_diff
        (Excel), para que exportar refleje exactamente lo que se ve."""
        filas: list[tuple[str, tuple, str]] = []   # (clave orden, valores, tag)
        if self._var_est_cambios.get():
            vis = comp.modificados
            if self._var_comp_solo_acad.get():
                vis = [f for f in vis if f.delta_importe != 0]
            for f in vis:
                delta = f.delta_neto or f.delta_importe
                tag = ("aviso" if f.dni_distinto
                       else "sube" if delta > 0 else "baja")
                filas.append((_norm_texto(f.nombre),
                              self._fila_diff(f, self._estado_cambios(f)),
                              tag))
        if self._var_est_sin.get():
            for f in comp.iguales:
                filas.append((_norm_texto(f.nombre),
                              self._fila_diff(f, "Sin cambios"),
                              "aviso" if f.dni_distinto else ""))
        if self._var_est_bajas.get():
            for r in comp.solo_a:
                filas.append((_norm_texto(r.nombre), self._fila_solo(r, False),
                              "sube" if r.importe_neto < 0 else "baja"))
        if self._var_est_altas.get():
            for r in comp.solo_b:
                filas.append((_norm_texto(r.nombre), self._fila_solo(r, True),
                              "baja" if r.importe_neto < 0 else "sube"))
        filas.sort(key=lambda t: t[0])
        return filas

    def _load_tabla_diff(self, comp: Comparacion) -> None:
        """Rellena la tabla de diff y el resumen con una Comparacion,
        aplicando los filtros por Estado y el de 'solo cambia I.Adm.'."""
        self._comp_actual = comp
        filas = self._filas_diff_visibles(comp)
        self._tree_diff.load([t[1] for t in filas],
                             tags=[t[2] for t in filas])

        total = (len(comp.modificados) + len(comp.iguales)
                 + len(comp.solo_a) + len(comp.solo_b))
        filtro = (f" · Mostrados: {len(filas)} de {total}"
                  if len(filas) != total else "")
        self._lbl_comp_resumen.configure(
            text=f"Modificados: {len(comp.modificados)} · "
                 f"Solo en A (bajas): {len(comp.solo_a)} · "
                 f"Solo en B (altas): {len(comp.solo_b)} · "
                 f"Sin cambios: {len(comp.iguales)}{filtro}",
            fg="#1B2631")
        self._log_write(
            f"Comparación: {len(comp.modificados)} modificados, "
            f"{len(comp.solo_a)} solo en A, {len(comp.solo_b)} solo en B, "
            f"{len(comp.iguales)} sin cambios. Δ total {comp.delta_total:+.2f} €.",
            "ok")

    # ── Panel de comparación de Detalle (dos tablas lado a lado) ──────────────

    def _show_comparar_det(self) -> None:
        """Muestra el panel Comparar Detalle (mismo overlay a pantalla
        completa que el de Comparar Resumen)."""
        if not hasattr(self, "_compdet_panel"):
            self._build_comparar_det_panel()
        self._compdet_panel.place(relx=0.5, rely=0.5, anchor="center",
                                  relwidth=0.99, relheight=0.98)
        self._compdet_panel.lift()

    def _hide_comparar_det(self) -> None:
        if hasattr(self, "_compdet_panel"):
            self._compdet_panel.place_forget()
            # Las tablas comparten el estilo ttk "Treeview": al cerrar el
            # panel, devuelve a las tablas principales su tamaño óptimo.
            if self._tree_detalle.tree.get_children():
                self._autofit_tablas()

    def _build_comparar_det_panel(self) -> None:
        p = tk.Frame(self, bg=BG_APP, bd=2, relief="ridge",
                     highlightbackground="#AAAAAA", highlightthickness=1)
        self._compdet_panel = p

        # Cabecera del panel
        hdr = tk.Frame(p, bg=BG_COMP)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  ⚖  Comparar Detalle de dos liquidaciones "
                           "(A = anterior, B = nuevo)",
                 bg=BG_COMP, fg="white", font=FONT_TITLE, anchor="w"
                 ).pack(side="left", pady=4)
        ttk.Button(hdr, text="✕ Cerrar", style="Config.TButton",
                   command=self._hide_comparar_det).pack(side="right", padx=8)
        self._btn_exportar_det = ttk.Button(
            hdr, text="💾 Exportar Excel", style="Export.TButton",
            command=self._exportar_compdet)
        self._btn_exportar_det.pack(side="right", padx=8)
        Tooltip(self._btn_exportar_det,
                "Exporta a Excel las tablas A y B actualmente mostradas "
                "(con los filtros 'Mostrar:' activos), una por hoja")

        # Selección de los dos PDF
        sel = tk.Frame(p, bg=BG_APP, padx=12, pady=8)
        sel.pack(fill="x")
        sel.columnconfigure(1, weight=1)
        for fila, (texto, var, cmd) in enumerate((
                ("PDF A (anterior):", self._compdet_a_var,
                 lambda: self._pick_pdf_comparar_det("A")),
                ("PDF B (nuevo):", self._compdet_b_var,
                 lambda: self._pick_pdf_comparar_det("B")))):
            tk.Label(sel, text=texto, bg=BG_APP, font=FONT_UI, anchor="w"
                     ).grid(row=fila, column=0, sticky="w", padx=(0, 10), pady=2)
            tk.Entry(sel, textvariable=var, font=FONT_ENTRY, bg="white",
                     fg="#222222", relief="sunken", bd=1, state="readonly"
                     ).grid(row=fila, column=1, sticky="ew", pady=2, ipady=4)
            ttk.Button(sel, text="Seleccionar", style="Sel.TButton", command=cmd
                       ).grid(row=fila, column=2, padx=(8, 0), pady=2, ipadx=4)

        # Resumen de la comparación (páginas/registros/neto de cada PDF)
        self._lbl_compdet_resumen = tk.Label(p, text="", bg=BG_APP,
                                             fg="#1B2631", font=FONT_BOLD,
                                             anchor="w")
        self._lbl_compdet_resumen.pack(fill="x", padx=12, pady=(0, 4))

        # Filtros por color/Estado (qué categorías de fila se muestran en
        # AMBAS tablas; mismo patrón "Mostrar:" que el panel Comparar Resumen).
        fil_est = tk.Frame(p, bg=BG_APP)
        fil_est.pack(fill="x", padx=12, pady=(0, 4))
        tk.Label(fil_est, text="Mostrar:", bg=BG_APP, fg="#1B2631",
                 font=FONT_ENTRY).pack(side="left")
        for texto, var in (("Suben", self._var_cd_sube),
                           ("Bajan", self._var_cd_baja),
                           ("Nuevos", self._var_cd_nueva),
                           ("Eliminados", self._var_cd_del),
                           ("Sin cambios", self._var_cd_sin)):
            tk.Checkbutton(fil_est, text=texto, variable=var, bg=BG_APP,
                           fg="#1B2631", activebackground=BG_APP,
                           font=FONT_ENTRY, command=self._refiltrar_compdet
                           ).pack(side="left", padx=(10, 0))

        # Dos tablas de Detalle lado a lado (mismos bloques con cabecera
        # coloreada y contador que la vista previa de la ventana principal).
        grid = tk.Frame(p, bg=BG_APP)
        grid.pack(fill="both", expand=True)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        grid.rowconfigure(0, weight=1)
        col_a = tk.Frame(grid, bg=BG_APP)
        col_a.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        col_b = tk.Frame(grid, bg=BG_APP)
        col_b.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        self._tree_det_a, self._lbl_fecha_det_a = self._bloque_tabla_detalle(
            col_a, "Detalle — PDF A (anterior)")
        self._tree_det_b, self._lbl_fecha_det_b = self._bloque_tabla_detalle(
            col_b, "Detalle — PDF B (nuevo)")

        # Sincronización de selección A <-> B por Exped (misma mecánica que
        # la sincronización por DNI entre Detalle y Resumen; add="+" para no
        # pisar el binding interno de DataTable).
        self._tree_det_a.tree.bind(
            "<<TreeviewSelect>>", self._on_sel_det_a, add="+")
        self._tree_det_b.tree.bind(
            "<<TreeviewSelect>>", self._on_sel_det_b, add="+")

        # Leyenda de colores (bajo las tablas)
        fil = tk.Frame(p, bg=BG_APP)
        fil.pack(fill="x", padx=12, pady=(4, 8))
        for color, texto in ((C_ROW_SUBE, "sube (↑)"),
                             (C_ROW_BAJA, "baja (↓)"),
                             (C_ROW_NUEVA, "nuevo en B"),
                             (C_ROW_DEL, "eliminado en B"),
                             (C_ROW_REVISAR, "desalineación DNI/Apellido")):
            tk.Label(fil, text="   ", bg=color, relief="solid", bd=1
                     ).pack(side="left", padx=(0, 4))
            tk.Label(fil, text=texto, bg=BG_APP, fg="#555555", font=FONT_ENTRY
                     ).pack(side="left", padx=(0, 18))

    def _bloque_tabla_detalle(self, parent: tk.Widget, titulo: str
                              ) -> tuple[DataTable, tk.Label]:
        """Bloque con cabecera coloreada + fecha de último cobro + contador +
        DataTable idéntico al de la vista previa (Detalle) de la ventana
        principal. Devuelve (tabla, etiqueta de última fecha de cobro)."""
        tb, c = self._bloque(titulo, BG_TAB1, expand=True, parent=parent,
                             header_h=self._header_h)
        tb.columnconfigure(0, weight=1)
        lbl_fecha = tk.Label(tb, text="", bg=BG_TAB1, fg="white",
                             font=FONT_ENTRY)
        lbl_fecha.grid(row=0, column=1, padx=(0, 16))
        lbl_count = tk.Label(tb, text="", bg=BG_TAB1, fg="white", font=FONT_UI)
        lbl_count.grid(row=0, column=2, padx=(0, 6))
        c.rowconfigure(0, weight=1)
        c.columnconfigure(0, weight=1)
        tabla = DataTable(
            c, COL_COMPDET_IDS, COL_COMPDET_NAMES,
            font_ui=FONT_TABLA, font_bold=FONT_TABLA_BOLD,
            on_select=lambda s, t, lbl=lbl_count:
                lbl.configure(text=f"{s}/{t}" if t else ""))
        tabla.pack(fill="both", expand=True)
        self._config_tabla_detalle(tabla)
        # Colores del marcado fila a fila A-B (ver _estados_detalle)
        tabla.set_tag("sube", C_ROW_SUBE, C_ROW_SUBE_SEL)
        tabla.set_tag("baja", C_ROW_BAJA, C_ROW_BAJA_SEL)
        tabla.set_tag("nueva", C_ROW_NUEVA, C_ROW_NUEVA_SEL)
        tabla.set_tag("del", C_ROW_DEL, C_ROW_DEL_SEL)
        return tabla, lbl_fecha

    def _pick_pdf_comparar_det(self, cual: str) -> None:
        """Selecciona el PDF A o B del panel Comparar Detalle. Cuando ya
        están los dos, lanza la comparación automáticamente."""
        path = filedialog.askopenfilename(
            title=f"Seleccionar PDF {cual} "
                  f"({'anterior' if cual == 'A' else 'nuevo'})",
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        (self._compdet_a_var if cual == "A" else self._compdet_b_var).set(path)
        if self._compdet_a_var.get() and self._compdet_b_var.get():
            self._comparar_det_ejecutar()

    def _comparar_det_ejecutar(self) -> None:
        if self._comparando_det:
            self._log_write("Ya hay una comparación en marcha.", "warn")
            return
        pa = self._compdet_a_var.get().strip()
        pb = self._compdet_b_var.get().strip()
        if not pa or not pb:
            self._lbl_compdet_resumen.configure(
                text="Selecciona los dos PDF a comparar.", fg=C_AVISO)
            return
        self._comparando_det = True
        self._lbl_compdet_resumen.configure(text="Comparando…", fg="#555555")
        self._log_write(
            f"Comparando Detalle {os.path.basename(pa)} ↔ "
            f"{os.path.basename(pb)}…", "info")
        threading.Thread(target=self._comparar_det_worker, args=(pa, pb),
                         daemon=True).start()

    def _comparar_det_worker(self, path_a: str, path_b: str) -> None:
        try:
            try:
                regs_a, pag_a = parse_pdf(path_a)
                regs_b, pag_b = parse_pdf(path_b)
            except Exception as exc:
                self.after(0, self._log_write,
                           f"Error al leer los PDF: {exc}", "error")
                self.after(0, self._lbl_compdet_resumen.configure,
                           {"text": f"Error al leer los PDF: {exc}",
                            "fg": "#C0392B"})
                return
            # Igual que en Comparar Resumen: separa DNI/Apellido desalineados
            # antes de mostrar (no toca importes).
            corregir_alineacion(regs_a)
            corregir_alineacion(regs_b)
            # Comparación (motor de pdf_compare, no modifica los Registro):
            # aporta los alumnos nuevos en B y los cambios de importe para
            # el filtro/marcado del checkbox.
            comp = comparar(regs_a, regs_b)
            self.after(0, self._log_write,
                       f"A: {pag_a} pág., {len(regs_a)} reg. · "
                       f"B: {pag_b} pág., {len(regs_b)} reg.", "ok")
            self.after(0, self._load_tablas_compdet,
                       regs_a, regs_b, pag_a, pag_b, comp)
        finally:
            self.after(0, self._fin_comparar_det)

    def _fin_comparar_det(self) -> None:
        self._comparando_det = False

    def _load_tablas_compdet(self, regs_a: list[Registro],
                             regs_b: list[Registro],
                             pag_a: int, pag_b: int,
                             comp: Comparacion) -> None:
        """Guarda los datos parseados y pinta las dos tablas + el resumen.
        El re-filtrado del checkbox reusa estos datos sin re-parsear."""
        self._compdet_datos = (regs_a, regs_b, pag_a, pag_b, comp)
        self._render_tablas_compdet()
        self._log_write(
            f"Comparación Detalle cargada. Δ neto {comp.delta_total:+.2f} €.",
            "ok")
        # Con las dos tablas rellenas, ajusta letra/columnas a sus recuadros.
        self.after(0, self._autofit_compdet)

    def _refiltrar_compdet(self) -> None:
        """Reaplica los filtros 'Mostrar:' sobre los últimos datos parseados
        (sin re-parsear los PDF)."""
        if self._compdet_datos is not None:
            self._render_tablas_compdet()
            self.after(0, self._autofit_compdet)

    def _compdet_mostrar_filtros(self) -> dict[str, bool]:
        """Estado actual de los checkboxes 'Mostrar:' del panel Comparar
        Detalle, como dict color -> se_muestra (clave '' = sin cambios)."""
        return {
            "sube": self._var_cd_sube.get(),
            "baja": self._var_cd_baja.get(),
            "nueva": self._var_cd_nueva.get(),
            "del": self._var_cd_del.get(),
            "": self._var_cd_sin.get(),
        }

    def _exportar_compdet(self) -> None:
        """Exporta a Excel las dos tablas de Detalle actualmente mostradas
        (A y B, con los filtros 'Mostrar:' activos), una en cada hoja del
        mismo libro, más una tercera hoja 'Detalle A+B' que las une lado a
        lado (ver `_filas_detalle_unificado`). Propone como carpeta la del
        propio .exe/script y como nombre 'Comparar_Detalle_<fecha_hora>',
        pero deja decidir al usuario en el diálogo 'Guardar como' (igual que
        'Guardar Excel')."""
        if self._exportando_det:
            self._log_write("Ya se está exportando el Excel.", "warn")
            return
        if self._compdet_datos is None:
            self._log_write(
                "Todavía no hay una comparación de Detalle que exportar.",
                "warn")
            return

        marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = filedialog.asksaveasfilename(
            title="Exportar Comparar Detalle",
            initialdir=_exe_dir(),
            initialfile=f"Comparar_Detalle_{marca}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")])
        if not output_path:
            self._log_write("Exportación a Excel cancelada.", "info")
            return

        regs_a, regs_b, _pag_a, _pag_b, comp = self._compdet_datos
        mostrar = self._compdet_mostrar_filtros()
        pareja_de_a = {id(ra): rb for ra, rb in comp.pares}
        pareja_de_b = {id(rb): ra for ra, rb in comp.pares}
        rows_a, _tags_a, _dnis_a, _tot_a = self._filas_detalle_estado(
            regs_a, pareja_de_a, True, mostrar)
        rows_b, _tags_b, _dnis_b, _tot_b = self._filas_detalle_estado(
            regs_b, pareja_de_b, False, mostrar)
        rows_uni = self._filas_detalle_unificado(mostrar)
        sheets = [
            ("Detalle A", COL_COMPDET_NAMES, rows_a,
             COL_COMPDET_MONEY_COLS, (), COL_COMPDET_DNI_COL),
            ("Detalle B", COL_COMPDET_NAMES, rows_b,
             COL_COMPDET_MONEY_COLS, (), COL_COMPDET_DNI_COL),
            ("Detalle A+B", COL_COMPDET_UNI_NAMES, rows_uni,
             COL_COMPDET_UNI_MONEY_COLS, (), COL_COMPDET_UNI_DNI_COL),
        ]

        self._exportando_det = True
        self._btn_exportar_det.state(["disabled"])
        self._log_write(f"Exportando Comparar Detalle a {output_path}…",
                        "info")
        threading.Thread(target=self._exportar_compdet_worker,
                         args=(sheets, output_path), daemon=True).start()

    def _exportar_compdet_worker(
            self, sheets: list[tuple[str, tuple, list[tuple],
                              tuple[int, ...], tuple[int, ...], int | None]],
            output_path: str) -> None:
        try:
            try:
                export_multi_sheet_to_excel(sheets, output_path)
            except Exception as exc:
                self.after(0, self._log_write,
                          f"No se pudo exportar el Excel: {exc}", "error")
                return
            self.after(0, self._log_write,
                       f"Excel exportado: {output_path}", "ok")
        finally:
            self.after(0, self._fin_exportar_compdet)

    def _fin_exportar_compdet(self) -> None:
        self._exportando_det = False
        self._btn_exportar_det.state(["!disabled"])

    @staticmethod
    def _ultima_fecha_cobro(registros: list[Registro]) -> str:
        """Fecha de cobro (F. Cobro) más reciente entre todas las Referencia
        de los registros dados, en formato dd/mm/aaaa. Cadena vacía si no hay
        ninguna referencia con fecha reconocible."""
        mejor_d = None
        mejor_txt = ""
        for r in registros:
            for ref in r.referencias:
                try:
                    d = datetime.datetime.strptime(ref.fecha, "%d/%m/%Y")
                except ValueError:
                    continue
                if mejor_d is None or d > mejor_d:
                    mejor_d, mejor_txt = d, ref.fecha
        return mejor_txt

    def _cmp_importe(self, va: float, vb: float) -> tuple[str, str]:
        """Compara dos importes (A, B) de una misma fila: tag de color +
        texto de la columna Estado (el delta B-A con signo, misma cifra que
        `_fmt_delta` en la tabla de diff de Comparar Resumen, seguido de la
        flecha ↑/↓). Sin cambio: sin tag ni texto."""
        diff = round(vb - va, 2)
        if diff > 0:
            return "sube", self._fmt_delta(diff)
        if diff < 0:
            return "baja", self._fmt_delta(diff)
        return "", ""

    def _filas_detalle_estado(
            self, registros: list[Registro], parejas: dict[int, Registro],
            es_a: bool, mostrar: dict[str, bool]
            ) -> tuple[list[tuple], list[str], list[str], int]:
        """Filas del panel Comparar Detalle para `registros` (de A o de B),
        una por Referencia más la de Imp.Adm. por alumno, con el marcado
        fila a fila del otro lado: `parejas` mapea id(Registro) -> Registro
        emparejado del otro listado (ausente si el alumno no existe allí).
        La columna Rec. numera los recibos de cada Exped por orden (1, 2...,
        el mismo de `r.referencias`, independiente de qué filtro esté
        activo) y pone "Adm." en la fila del importe administrativo. Cada
        fila se compara con su equivalente en la pareja (misma Referencia, o
        el Imp.Adm.): sube (verde) / baja (rojo) si el importe cambió, nueva
        (amarillo, solo en B) / del (azul, solo en A) si la fila no tiene
        equivalente, o sin cambio (sin tag). `mostrar` mapea cada categoría
        ("sube"/"baja"/"nueva"/"del"/"") a si se debe generar esa fila; el
        marcado tiene prioridad sobre el ámbar de revisión. Devuelve (filas,
        tags, dnis, total de filas SIN aplicar `mostrar`, para poder
        informar cuántas quedan ocultas por los filtros)."""
        rows, tags, dnis = [], [], []
        i = 0
        total = 0

        def _tag(color: str) -> str:
            nonlocal i
            final = color or ("revisar" if r.revisar else
                              ("row_a" if i % 2 == 0 else "row_b"))
            i += 1
            return final

        for r in registros:
            pareja = parejas.get(id(r))
            refs_otro = ({ref.referencia: ref for ref in pareja.referencias}
                         if pareja is not None else {})

            for rec_i, ref in enumerate(r.referencias, start=1):
                otro = refs_otro.get(ref.referencia)
                if otro is None:
                    color, estado = ("del", "Del") if es_a else ("nueva", "New")
                else:
                    va, vb = (ref.importe, otro.importe) if es_a \
                        else (otro.importe, ref.importe)
                    color, estado = self._cmp_importe(va, vb)
                total += 1
                if not mostrar.get(color, True):
                    continue
                rows.append((r.exped, r.dni, r.nombre, str(rec_i),
                            ref.referencia, ref.fecha,
                            f"{ref.importe:.2f}", "", estado))
                tags.append(_tag(color))
                dnis.append(r.dni)

            if pareja is None:
                color, estado = ("del", "Del") if es_a else ("nueva", "New")
            else:
                va, vb = (r.administrativo, pareja.administrativo) if es_a \
                    else (pareja.administrativo, r.administrativo)
                color, estado = self._cmp_importe(va, vb)
            total += 1
            if not mostrar.get(color, True):
                continue
            rows.append((r.exped, r.dni, r.nombre, "Adm.", "", "",
                        "", f"{r.administrativo:.2f}", estado))
            tags.append(_tag(color))
            dnis.append(r.dni)

        return rows, tags, dnis, total

    def _filas_detalle_unificado(self, mostrar: dict[str, bool]
                                 ) -> list[tuple]:
        """Filas de la hoja de exportación 'Detalle A+B': une A y B lado a
        lado en una única fila por recibo (o Imp.Adm.), en vez de una tabla
        por PDF (esta hoja no tiene tabla en pantalla, solo existe al
        exportar). Reutiliza `_compdet_datos`/`comp.pares` igual que
        `_filas_detalle_estado`, pero en vez de una fila por lado por
        Referencia, junta ambos lados de la MISMA Referencia en una fila.
        Si el recibo no existe en B, el bloque B queda en blanco (Estado
        final 'Del'); si es nuevo en B, el bloque A queda en blanco (Estado
        final 'New'). El Rec. es compartido (mismo recibo en ambos bloques)
        y sigue el orden de A primero, recibos nuevos de B después, y por
        último los alumnos enteramente nuevos en B (altas)."""
        regs_a, _regs_b, _pag_a, _pag_b, comp = self._compdet_datos
        pareja_de_a = {id(ra): rb for ra, rb in comp.pares}
        BLANCO = ("", "", "", "", "")

        def _bloque_ref(ref, estado: str) -> tuple:
            return (ref.referencia, ref.fecha, f"{ref.importe:.2f}", "",
                    estado)

        def _bloque_adm(r: Registro, estado: str) -> tuple:
            return ("", "", "", f"{r.administrativo:.2f}", estado)

        rows: list[tuple] = []

        def _emitir(exped, dni, nombre, rec, bloque_a, bloque_b,
                    color, estado) -> None:
            if mostrar.get(color, True):
                rows.append((exped, dni, nombre, rec,
                            *bloque_a, *bloque_b, estado))

        for ra in regs_a:
            rb = pareja_de_a.get(id(ra))
            refs_b = ({ref.referencia: ref for ref in rb.referencias}
                     if rb is not None else {})
            refs_a_ids = {ref.referencia for ref in ra.referencias}
            rec_i = 0

            for ref in ra.referencias:
                rec_i += 1
                otro = refs_b.get(ref.referencia)
                if otro is None:
                    color, estado = "del", "Del"
                    bloque_a, bloque_b = _bloque_ref(ref, estado), BLANCO
                else:
                    color, estado = self._cmp_importe(ref.importe,
                                                      otro.importe)
                    bloque_a = _bloque_ref(ref, estado)
                    bloque_b = _bloque_ref(otro, estado)
                _emitir(ra.exped, ra.dni, ra.nombre, str(rec_i),
                       bloque_a, bloque_b, color, estado)

            if rb is not None:
                for ref in rb.referencias:
                    if ref.referencia in refs_a_ids:
                        continue    # ya emparejado arriba
                    rec_i += 1
                    color, estado = "nueva", "New"
                    _emitir(ra.exped, ra.dni, ra.nombre, str(rec_i),
                           BLANCO, _bloque_ref(ref, estado), color, estado)

            if rb is None:
                color, estado = "del", "Del"
                bloque_a, bloque_b = _bloque_adm(ra, estado), BLANCO
            else:
                color, estado = self._cmp_importe(ra.administrativo,
                                                  rb.administrativo)
                bloque_a = _bloque_adm(ra, estado)
                bloque_b = _bloque_adm(rb, estado)
            _emitir(ra.exped, ra.dni, ra.nombre, "Adm.",
                   bloque_a, bloque_b, color, estado)

        # Alumnos enteramente nuevos en B (altas de comp.solo_b): todo su
        # bloque A vacío, Estado final "New".
        for rb in comp.solo_b:
            rec_i = 0
            for ref in rb.referencias:
                rec_i += 1
                _emitir(rb.exped, rb.dni, rb.nombre, str(rec_i),
                       BLANCO, _bloque_ref(ref, "New"), "nueva", "New")
            _emitir(rb.exped, rb.dni, rb.nombre, "Adm.",
                   BLANCO, _bloque_adm(rb, "New"), "nueva", "New")

        return rows

    def _render_tablas_compdet(self) -> None:
        """Rellena las dos tablas de Detalle (A izquierda, B derecha) con el
        marcado fila a fila (Referencia/Imp.Adm.) contra el listado
        emparejado del otro lado, vía `_filas_detalle_estado`. Los filtros
        "Mostrar:" (uno por categoría de color) deciden qué filas se generan
        en cada tabla; ambas tablas usan los mismos filtros."""
        if self._compdet_datos is None:
            return
        regs_a, regs_b, pag_a, pag_b, comp = self._compdet_datos
        mostrar = self._compdet_mostrar_filtros()

        self._lbl_fecha_det_a.configure(
            text=f"Última fecha de cobro: "
                 f"{self._ultima_fecha_cobro(regs_a) or '—'}")
        self._lbl_fecha_det_b.configure(
            text=f"Última fecha de cobro: "
                 f"{self._ultima_fecha_cobro(regs_b) or '—'}")

        # Emparejamiento (por Exped/DNI, ya resuelto por pdf_compare.comparar)
        # en ambos sentidos, por identidad de Registro.
        pareja_de_a = {id(ra): rb for ra, rb in comp.pares}
        pareja_de_b = {id(rb): ra for ra, rb in comp.pares}

        mostradas: dict[str, tuple[int, int]] = {}
        for tabla, regs, parejas, es_a in (
                (self._tree_det_a, regs_a, pareja_de_a, True),
                (self._tree_det_b, regs_b, pareja_de_b, False)):
            rows, tags, _dnis, total = self._filas_detalle_estado(
                regs, parejas, es_a, mostrar)
            mostradas["A" if es_a else "B"] = (len(rows), total)
            iids = tabla.load(rows, tags=tags)
            # Mapa Exped -> iids para la sincronización de selección A <-> B.
            mapa: dict[str, list[str]] = {}
            for iid, row in zip(iids, rows):
                mapa.setdefault(row[0], []).append(iid)
            if es_a:
                self._exped2iid_det_a = mapa
            else:
                self._exped2iid_det_b = mapa

        texto = (f"A: {pag_a} pág. · {len(regs_a)} alumnos · "
                 f"neto {comp.total_neto_a:,.2f} €   |   "
                 f"B: {pag_b} pág. · {len(regs_b)} alumnos · "
                 f"neto {comp.total_neto_b:,.2f} € "
                 f"(Δ {comp.delta_total:+,.2f} €)")
        extra = [f"{lado}: {n}/{total}"
                for lado, (n, total) in mostradas.items() if n != total]
        if extra:
            texto += "   |   Filas mostradas " + " · ".join(extra)
        self._lbl_compdet_resumen.configure(text=texto, fg="#1B2631")

    # Sincronización de selección A <-> B por Exped (clave estable entre
    # versiones del mismo curso, la misma que usa el motor de comparación).

    def _on_sel_det_a(self, _event=None) -> None:
        if self._sync_guard:
            return
        iid = self._fila_activa(self._tree_det_a)
        if iid is None:
            return
        exped = self._tree_det_a.tree.set(iid, "exped")
        objetivos = self._exped2iid_det_b.get(exped)
        if objetivos:
            self._sync_select(self._tree_det_b, objetivos)

    def _on_sel_det_b(self, _event=None) -> None:
        if self._sync_guard:
            return
        iid = self._fila_activa(self._tree_det_b)
        if iid is None:
            return
        exped = self._tree_det_b.tree.set(iid, "exped")
        objetivos = self._exped2iid_det_a.get(exped)
        if objetivos:
            self._sync_select(self._tree_det_a, objetivos)

    def _autofit_compdet(self, _try: int = 0) -> None:
        """Auto-ajuste del tamaño de letra para las dos tablas del panel
        Comparar Detalle (mismo criterio que _autofit_tablas: el mayor tamaño
        de la banda con el que ambas quepan a lo ancho de su recuadro)."""
        if not hasattr(self, "_tree_det_a"):
            return
        smin, smax = self._banda_font()
        self.update_idletasks()
        w1 = self._tree_det_a.tree.winfo_width()
        w2 = self._tree_det_b.tree.winfo_width()
        if (w1 < 50 or w2 < 50):        # todavía sin geometría real: reintenta
            if _try < 20:
                self.after(150, lambda: self._autofit_compdet(_try + 1))
            return
        best = smax
        for tabla, ancho in ((self._tree_det_a, w1), (self._tree_det_b, w2)):
            if tabla.tree.get_children():
                best = min(best, tabla.fit_font_size(smin, smax, ancho - 2))
        self._apply_table_font(best)
        self._log_write(
            f"Auto-ajuste de tablas (Comparar Detalle): letra {best} pt "
            f"(rango {smin}–{smax}).", "info")

    # ── Panel de ayuda ──────────────────────────────────────────────────────

    def _show_ayuda(self) -> None:
        messagebox.showinfo(
            "Ayuda",
            "1) Selecciona el PDF de liquidación de tasas: se parsea solo "
            "y aparece la vista previa (la carpeta destino se rellena con "
            "la del propio PDF; puedes cambiarla).\n"
            "2) Pulsa 'Guardar Excel' para generarlo.\n\n"
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
        # Autorrelleno: por defecto el Excel se guarda en la misma carpeta
        # que el PDF de origen (el usuario puede cambiarla después con
        # "Seleccionar" antes de pulsar "Guardar Excel").
        self._fuente2_var.set(os.path.dirname(path))
        self._log_write(f"PDF seleccionado: {path}", "info")
        self._parsear()

    def _pick_carpeta_destino(self) -> None:
        path = filedialog.askdirectory(title="Seleccionar carpeta destino del Excel")
        if path:
            self._fuente2_var.set(path)
            self._log_write(f"Carpeta destino: {path}", "info")

    # ── Carga de tablas ───────────────────────────────────────────────────────

    @staticmethod
    def _filas_detalle(registros: list[Registro]
                       ) -> tuple[list[tuple], list[str], list[str]]:
        """Filas de una tabla con las columnas de la vista Detalle: una fila
        por línea de cobro más una fila por alumno con el importe
        administrativo. Devuelve (filas, tags franja/revisar, dni por fila).
        La usan la vista previa principal y el panel Comparar Detalle."""
        rows, tags, dnis = [], [], []
        i = 0

        def _tag(marcada: bool, pos: int) -> str:
            return "revisar" if marcada else ("row_a" if pos % 2 == 0 else "row_b")

        for r in registros:
            for ref in r.referencias:
                rows.append((r.exped, r.dni, r.nombre, ref.referencia,
                             ref.fecha, f"{ref.importe:.2f}", "",
                             ref.plazo, ref.forma_pago))
                tags.append(_tag(r.revisar, i))
                dnis.append(r.dni)
                i += 1
            # Fila con el importe administrativo: sin datos de cobro.
            rows.append((r.exped, r.dni, r.nombre, "", "", "",
                        f"{r.administrativo:.2f}", "", ""))
            tags.append(_tag(r.revisar, i))
            dnis.append(r.dni)
            i += 1
        return rows, tags, dnis

    def _load_tabla_detalle(self, registros: list[Registro]) -> None:
        rows, tags, dnis = self._filas_detalle(registros)
        iids = self._tree_detalle.load(rows, tags=tags)
        self._dni2iid_detalle = {}
        for iid, dni in zip(iids, dnis):
            self._dni2iid_detalle.setdefault(dni, []).append(iid)

    def _load_tabla_resumen(self, registros: list[Registro]) -> None:
        rows, tags = [], []
        for i, r in enumerate(registros):
            rows.append((r.exped, r.dni, r.nombre, len(r.referencias),
                         f"{r.importe:.2f}", f"{r.importe_neto:.2f}",
                         f"{r.administrativo:.2f}"))
            tags.append("revisar" if r.revisar
                        else ("row_a" if i % 2 == 0 else "row_b"))
        iids = self._tree_resumen.load(rows, tags=tags)
        self._iid2reg_res = dict(zip(iids, registros))
        self._iid_revisar = {iid for iid, r in zip(iids, registros) if r.revisar}
        self._dni2iid_resumen = {r.dni: iid for iid, r in zip(iids, registros)}

    # ── Sincronización de selección por DNI entre Detalle y Resumen ─────────

    @staticmethod
    def _fila_activa(tabla: DataTable) -> str | None:
        """iid de la fila con foco si sigue seleccionada, si no la primera
        fila seleccionada; None si no hay selección."""
        sel = tabla.tree.selection()
        if not sel:
            return None
        foco = tabla.tree.focus()
        return foco if foco in sel else sel[0]

    def _sync_select(self, tabla: DataTable, iids: list[str]) -> None:
        self._sync_guard = True
        try:
            tabla.tree.selection_set(iids)
            tabla.tree.focus(iids[0])
            tabla.scroll_to(iids[0], center=True)
        finally:
            self.after_idle(lambda: setattr(self, "_sync_guard", False))

    def _on_sel_detalle(self, _event=None) -> None:
        if self._sync_guard:
            return
        iid = self._fila_activa(self._tree_detalle)
        if iid is None:
            return
        dni = self._tree_detalle.tree.set(iid, "dni")
        objetivo = self._dni2iid_resumen.get(dni)
        if objetivo:
            self._sync_select(self._tree_resumen, [objetivo])

    def _on_sel_resumen(self, _event=None) -> None:
        if self._sync_guard:
            return
        iid = self._fila_activa(self._tree_resumen)
        if iid is None:
            return
        dni = self._tree_resumen.tree.set(iid, "dni")
        objetivos = self._dni2iid_detalle.get(dni)
        if objetivos:
            self._sync_select(self._tree_detalle, objetivos)

    def _resumen_row(self, r: Registro) -> tuple:
        """Fila de la tabla Resumen para un Registro (mismo orden de columnas)."""
        return (r.exped, r.dni, r.nombre, len(r.referencias),
                f"{r.importe:.2f}", f"{r.importe_neto:.2f}",
                f"{r.administrativo:.2f}")

    # ── Tamaño de letra de las tablas (control manual + auto-ajuste) ────────────

    def _banda_font(self) -> tuple[int, int]:
        """Rango [mín, máx] configurado, saneado: enteros en 6–24 y mín ≤ máx.
        (El Spinbox puede quedar vacío o con texto no numérico; se tolera.)"""
        try:
            smin = int(self._var_font_min.get())
        except (tk.TclError, ValueError):
            smin = FONT_TABLA_MIN_DEF
        try:
            smax = int(self._var_font_max.get())
        except (tk.TclError, ValueError):
            smax = FONT_TABLA_MAX_DEF
        smin = max(6, min(24, smin))
        smax = max(6, min(24, smax))
        if smin > smax:
            smin, smax = smax, smin
        return smin, smax

    def _apply_table_font(self, size: int) -> None:
        """Fija el tamaño de letra RENDERIZADO de las dos tablas (comparten el
        estilo ttk 'Treeview'), escala la altura de fila y re-mide las columnas
        para que su ancho vuelva a cuadrar con el nuevo tamaño."""
        fam = FONT_TABLA[0]
        font_n = (fam, size)
        font_b = (fam, size, "bold")
        rowh = int(size * 2.4) + 1
        st = ttk.Style(self)
        st.configure("Treeview", font=font_n, rowheight=rowh)
        st.configure("Treeview.Heading", font=font_b)
        tablas = [self._tree_detalle, self._tree_resumen]
        if hasattr(self, "_tree_det_a"):     # tablas del panel Comparar Detalle
            tablas += [self._tree_det_a, self._tree_det_b]
        for tabla in tablas:
            tabla.set_measure_fonts(font_n, font_b)
            tabla.autosize()
        self._font_tabla_actual = size

    def _on_resize(self, event) -> None:
        """Debounce del redimensionado de la ventana: solo reacciona a cambios
        de ANCHO del propio toplevel (no de sus hijos ni a cambios de alto, que
        no afectan a si una tabla cabe a lo ancho). Reprograma el auto-ajuste
        200 ms después del último evento."""
        if event.widget is not self:
            return
        w = self.winfo_width()
        if w == self._last_win_w:
            return
        self._last_win_w = w
        if self._resize_job is not None:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(200, self._autofit_on_resize)

    def _autofit_on_resize(self) -> None:
        self._resize_job = None
        # Si el panel Comparar Detalle está visible con datos, se re-ajusta
        # ese (es lo que se ve); si no, las tablas principales.
        if (hasattr(self, "_compdet_panel")
                and self._compdet_panel.winfo_ismapped()
                and self._tree_det_a.tree.get_children()):
            self._autofit_compdet()
        elif (hasattr(self, "_tree_detalle")
                and self._tree_detalle.tree.get_children()):
            self._autofit_tablas()

    def _autofit_tablas(self, _try: int = 0) -> None:
        """Tras rellenar las tablas, elige el mayor tamaño de letra de la banda
        configurada con el que AMBAS quepan enteras a lo ancho de su recuadro
        (sin scroll horizontal) y lo aplica a las dos. Es un único tamaño
        compartido: manda la tabla más ancha (Detalle). Si una tabla es estrecha
        y sobra hueco, el tamaño sube hasta el máximo o hasta que la otra tabla
        deje de caber, lo que ocurra antes."""
        if not hasattr(self, "_tree_detalle"):
            return
        smin, smax = self._banda_font()
        self.update_idletasks()
        w1 = self._tree_detalle.tree.winfo_width()
        w2 = self._tree_resumen.tree.winfo_width()
        if (w1 < 50 or w2 < 50):        # todavía sin geometría real: reintenta
            if _try < 20:
                self.after(150, lambda: self._autofit_tablas(_try + 1))
            return
        best = smax
        for tabla, ancho in ((self._tree_detalle, w1), (self._tree_resumen, w2)):
            if tabla.tree.get_children():
                best = min(best, tabla.fit_font_size(smin, smax, ancho - 2))
        self._apply_table_font(best)
        self._log_write(
            f"Auto-ajuste de tablas: letra {best} pt (rango {smin}–{smax}).",
            "info")

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
        # El DNI de esta fila pudo cambiar (separación DNI/Apellido): el mapa
        # de sincronización Resumen -> Detalle quedaría con la clave vieja.
        self._dni2iid_resumen = {r.dni: i for i, r in self._iid2reg_res.items()}
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

    # ── Parseo automático (al seleccionar el PDF) ─────────────────────────────

    def _parsear(self) -> None:
        if self._procesando:
            self._log_write("Ya hay un proceso en marcha, espera a que termine.", "warn")
            return
        if not self._pdf_path:
            self._log_write("Selecciona un PDF.", "warn")
            return

        self._registros = []
        self._ultimo_excel = None
        self._btn_abrir.grid_remove()   # el Excel del PDF anterior ya no vale
        self._procesando = True
        self._log_write(f"Procesando {os.path.basename(self._pdf_path)}…", "info")
        threading.Thread(target=self._parsear_worker,
                         args=(self._pdf_path,), daemon=True).start()

    def _parsear_worker(self, pdf_path: str) -> None:
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

            self._registros = registros
            self.after(0, self._load_tabla_detalle, registros)
            self.after(0, self._load_tabla_resumen, registros)
            # Con las dos tablas ya rellenas, ajusta letra y ancho para que
            # entren enteras en sus recuadros (dentro de la banda configurada).
            self.after(0, self._autofit_tablas)
            self.after(0, self._post_correccion, n_rev)
        finally:
            self.after(0, self._fin_parsear)

    def _fin_parsear(self) -> None:
        self._procesando = False

    # ── Guardar Excel (acción explícita del usuario) ──────────────────────────

    def _guardar_excel(self) -> None:
        if self._guardando:
            self._log_write("Ya se está guardando el Excel.", "warn")
            return
        if not self._registros:
            self._log_write(
                "Todavía no hay datos que exportar: selecciona un PDF.", "warn")
            return

        destino = self._fuente2_var.get().strip()
        if not destino or not os.path.isdir(destino):
            destino = os.path.expanduser("~")
        nombre_base = (getattr(self, "_var_nombre_fichero", None).get().strip()
                       if hasattr(self, "_var_nombre_fichero") else "") \
            or _cfg_read("nombre_fichero") or "Liquidacion_Tasas"
        marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_propuesto = f"{nombre_base}_{marca}.xlsx"

        # El propio diálogo "Guardar como" propone carpeta y nombre, pero deja
        # que el usuario los cambie antes de confirmar.
        output_path = filedialog.asksaveasfilename(
            title="Guardar Excel de liquidación de tasas",
            initialdir=destino, initialfile=nombre_propuesto,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")])
        if not output_path:
            self._log_write("Guardado cancelado.", "info")
            return
        self._fuente2_var.set(os.path.dirname(output_path))

        self._guardando = True
        self._btn_guardar.state(["disabled"])
        self._log_write(f"Guardando Excel en {output_path}…", "info")
        threading.Thread(target=self._guardar_excel_worker,
                         args=(list(self._registros), output_path),
                         daemon=True).start()

    def _guardar_excel_worker(self, registros: list[Registro],
                              output_path: str) -> None:
        try:
            try:
                export_to_excel(registros, output_path)
            except Exception as exc:
                self.after(0, self._log_write,
                          f"No se pudo generar el Excel: {exc}", "error")
                return
            self._ultimo_excel = output_path
            self.after(0, self._log_write, f"Excel generado: {output_path}", "ok")
            self.after(0, self._btn_abrir.grid)
            # Confirmación modal de fin de grabación (nombre + carpeta); al
            # cerrarla abre la carpeta destino si la opción está activada.
            self.after(0, self._confirmar_grabacion, output_path)
        finally:
            self.after(0, self._fin_guardar)

    def _fin_guardar(self) -> None:
        self._guardando = False
        self._btn_guardar.state(["!disabled"])

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
