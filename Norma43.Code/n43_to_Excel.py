# ═══════════════════════════════════════════════════════════════════
#  n43_to_Excel — Conversor de ficheros Norma 43 (AEB 43) a Excel
#  Versión: v03
# ═══════════════════════════════════════════════════════════════════
#
#  Descripción:
#    Aplicación de escritorio con interfaz gráfica (tkinter) que
#    lee ficheros bancarios en formato Norma 43 (.n43) y genera
#    libro(s) Excel (.xlsx) con los movimientos formateados.
#
#  Formato Norma 43 — tipos de registro:
#    11  Cabecera de cuenta  (entidad, oficina, cuenta, fechas, saldo)
#    22  Movimiento          (fecha op., fecha valor, importe, doc, ref)
#    23  Concepto complementario (hasta 5 líneas por movimiento)
#    33  Cierre de cuenta    (saldo final, nº de registros)
#    88  Cierre de fichero
#
#  Codificación de los ficheros N43: ISO-8859-1 (latin-1)
#
#  Dependencias externas:
#    - openpyxl       → generación del fichero Excel
#    - Pillow (PIL)   → escalado de imagen (opcional; hay fallback)
#
#  Cambios respecto a v02:
#    - Selección múltiple de ficheros .n43
#    - Opción para permitir cuentas diferentes
#    - Opción para guardar cada cuenta en Excel separados
#    - Mismo cuenta → movimientos encadenados en una sola hoja
#    - Listbox para nombres de ficheros n43 (10 filas + scrollbar)
#    - Listbox dinámico para nombres Excel en modo multi separado
#    - Filtrado automático al mezclar cuentas sin permiso explícito
#
# ═══════════════════════════════════════════════════════════════════
#
#  ÍNDICE DE BLOQUES                                         Aprox.
#  ─────────────────────────────────────────────────────────────────
#   1. Imports y carga condicional de Pillow ................   65
#   2. read_n43_header()  — lectura de cabecera N43 ........   89
#   3. fmt_saldo()        — formato numérico español .......  139
#   4. Funciones de parseo N43 → Excel
#      _read_n43_rows()   — lectura encadenable ............  155
#      _write_rows_to_sheet() — escritura con estilos .......  205
#      _write_n43_sheet() — hoja simple ....................  295
#      _write_combined_to_sheet() — hoja encadenada ........  305
#      parse_n43()        — fichero único → Excel ...........  320
#      parse_n43_combined() — misma cuenta → un Excel ......  330
#      parse_n43_accounts() — varias cuentas → un Excel ....  345
#   5. Constantes de la GUI ................................  370
#   6. run_gui()  — interfaz gráfica principal .............  380
#      6a-6g. Widgets principales ..........................  420
#      6h. Variables de estado .............................  500
#      6h-bis. Panel de configuración ......................  510
#      6i. seleccionar() .................................  620
#      6j. convertir() ...................................  710
#      6k. Imagen Caracolillo y crédito ....................  810
#      6l. Inicialización y arranque .......................  865
#   7. Punto de entrada (__main__) .........................  880
#
# ═══════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────────
#  1. IMPORTS Y CARGA CONDICIONAL DE PILLOW
# ─────────────────────────────────────────────────────────────────

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import OrderedDict
from datetime import date
import os
import sys
import winreg

try:
    from PIL import Image, ImageTk
    RESAMPLE = getattr(Image, "Resampling", Image).LANCZOS
    HAS_PIL = True
except (ImportError, AttributeError):
    HAS_PIL = False


def resource_path(filename):
    """Resuelve rutas de recursos tanto en script como en .exe (PyInstaller)."""
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


_REG_KEY = r"Software\n43_to_Excel"


def load_config() -> dict:
    data = {}
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_KEY)
        for name in ("last_dir", "last_output_dir"):
            try:
                val, _ = winreg.QueryValueEx(key, name)
                data[name] = val
            except FileNotFoundError:
                pass
        for name in ("auto_update_n43_dir", "auto_update_output_dir",
                     "multi_select", "multi_account", "separate_excel"):
            try:
                val, _ = winreg.QueryValueEx(key, name)
                data[name] = bool(val)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except FileNotFoundError:
        pass
    return data


def save_config(data: dict):
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, _REG_KEY)
        for name, val in data.items():
            if isinstance(val, bool):
                winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, int(val))
            else:
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(val))
        winreg.CloseKey(key)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────
#  2. READ_N43_HEADER — Lectura de cabecera del fichero N43
# ─────────────────────────────────────────────────────────────────

def read_n43_header(path_n43: str) -> dict:
    info = {"num_movimientos": 0}
    with open(path_n43, encoding="latin-1") as f:
        for linea in f:
            reg = linea[:2]
            if reg == "11":
                info["entidad"] = linea[2:6].strip()
                info["oficina"] = linea[6:10].strip()
                info["cuenta"] = linea[10:20].strip()
                info["fecha_ini"] = date(
                    int("20" + linea[20:22]), int(linea[22:24]), int(linea[24:26]))
                info["fecha_fin"] = date(
                    int("20" + linea[26:28]), int(linea[28:30]), int(linea[30:32]))
                signo = -1 if linea[32] == "1" else 1
                info["saldo_ini"] = int(linea[33:47]) * signo * 0.01
                info["divisa"] = linea[47:50].strip()
            elif reg == "22":
                info["num_movimientos"] += 1
            elif reg == "33":
                signo = -1 if linea[58] == "1" else 1
                info["saldo_fin"] = int(linea[59:73]) * signo * 0.01
    return info


# ─────────────────────────────────────────────────────────────────
#  3. FMT_SALDO — Formato numérico estilo español
# ─────────────────────────────────────────────────────────────────

def fmt_saldo(valor: float) -> str:
    return f"{valor:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


# ─────────────────────────────────────────────────────────────────
#  4. FUNCIONES DE PARSEO N43 → EXCEL
# ─────────────────────────────────────────────────────────────────

def _read_n43_rows(path_n43: str, state: dict = None) -> tuple:
    """
    Lee movimientos de un fichero N43 con estado encadenable entre ficheros.

    state: dict con claves 'saldo', 'num_orden', 'anualidad'.
      - saldo=None → usa el saldo_ini del propio registro 11 (primer fichero).
      - saldo=float → continúa acumulando desde ese valor (ficheros encadenados).
    Devuelve (rows, state_actualizado).
    """
    if state is None:
        state = {"saldo": None, "num_orden": 0, "anualidad": ""}

    saldo_acum  = state["saldo"]
    num_orden   = state["num_orden"]
    anualidad   = state["anualidad"]
    usar_propio = saldo_acum is None
    saldo_ini   = saldo_acum   # para ficheros encadenados ya es correcto
    rows        = []
    current_row = None

    with open(path_n43, encoding="latin-1") as f:
        for linea in f:
            reg = linea[:2]
            if reg == "11" and usar_propio:
                signo      = -1 if linea[32] == "1" else 1
                saldo_acum = int(linea[33:47]) * signo * 0.01
                saldo_ini  = saldo_acum   # saldo antes del primer movimiento
                usar_propio = False
            elif reg == "22":
                año = "20" + linea[10:12]
                if linea[10:12] != anualidad:
                    num_orden = 0
                    anualidad = linea[10:12]
                num_orden += 1
                signo   = -1 if linea[27] == "1" else 1
                importe = int(linea[28:42]) * signo * 0.01
                saldo_acum += importe
                current_row = [
                    f"{año}{num_orden:06d}",
                    linea[6:10].strip(),
                    date(int("20" + linea[10:12]),
                         int(linea[12:14]), int(linea[14:16])),
                    date(int("20" + linea[16:18]),
                         int(linea[18:20]), int(linea[20:22])),
                    importe,
                    round(saldo_acum, 2),
                    linea[42:52].strip(),
                    linea[52:102].strip(),
                    "", "", "", "", ""
                ]
                rows.append(current_row)
            elif reg == "23" and current_row:
                idx = int(linea[2:4])
                if 1 <= idx <= 5:
                    current_row[7 + idx] = linea[4:79].strip()

    state["saldo"]     = saldo_acum
    state["num_orden"] = num_orden
    state["anualidad"] = anualidad
    return rows, state, saldo_ini


def _write_rows_to_sheet(ws, rows: list, saldo_ini: float = 0.0):
    """Escribe en ws una lista de filas ya leídas, con cabecera y estilos completos."""
    headers = [
        "Orden", "Oficina", "F.Operación", "F.Valor", "Importe",
        "Saldo", "Documento", "Referencia",
        "RCM01", "RCM02", "RCM03", "RCM04", "RCM05"
    ]

    header_fill  = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin         = Side(style="thin", color="CCCCCC")
    cell_border  = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.row_dimensions[1].height = 30
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = cell_border

    fill_even    = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
    fill_odd     = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    font_data    = Font(name="Arial", size=9)
    align_center = Alignment(horizontal="center", vertical="center")
    align_right  = Alignment(horizontal="right",  vertical="center")
    align_left   = Alignment(horizontal="left",   vertical="center")
    date_fmt     = "DD/MM/YYYY"
    money_fmt    = '#,##0.00 €;[RED]-#,##0.00 €'
    money_pos    = '#,##0.00 €'

    for r_idx, row in enumerate(rows, 2):
        fill = fill_even if r_idx % 2 == 0 else fill_odd
        for c_idx, val in enumerate(row, 1):
            cell        = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font   = font_data
            cell.fill   = fill
            cell.border = cell_border
            if c_idx in (3, 4):
                cell.number_format = date_fmt
                cell.alignment     = align_center
            elif c_idx == 5:
                cell.number_format = money_fmt
                cell.alignment     = align_right
            elif c_idx == 6:
                cell.number_format = money_pos
                cell.alignment     = align_right
            elif c_idx in (1, 2, 7):
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    col_widths = [14, 8, 12, 12, 14, 14, 12, 30, 25, 25, 25, 25, 25]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    last_data_row = len(rows) + 1
    total_row     = last_data_row + 1
    ws.row_dimensions[total_row].height = 20

    total_fill = PatternFill("solid", start_color="2E75B6", end_color="2E75B6")
    total_font = Font(name="Arial", bold=True, color="FFFFFF", size=9)

    for c in range(1, len(headers) + 1):
        cell      = ws.cell(row=total_row, column=c)
        cell.fill = total_fill
        cell.font = total_font
        cell.border = cell_border

    ws.cell(row=total_row, column=1, value="TOTAL").alignment = \
        Alignment(horizontal="center")
    ws.cell(row=total_row, column=5,
            value=f"=SUM(E2:E{last_data_row})").number_format = money_fmt
    ws.cell(row=total_row, column=5).alignment = align_right
    ws.cell(row=total_row, column=6,
            value=f"=E{total_row}+{saldo_ini:.2f}"
            ).number_format = money_pos

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_data_row}"


def _write_n43_sheet(ws, path_n43: str) -> int:
    """Lee un único fichero N43 y lo escribe en ws. Devuelve nº de filas."""
    rows, _, saldo_ini = _read_n43_rows(path_n43)
    _write_rows_to_sheet(ws, rows, saldo_ini)
    return len(rows)


def _write_combined_to_sheet(ws, paths: list, infos: list = None) -> int:
    """Varios ficheros de la misma cuenta → una hoja con saldo y numeración encadenados."""
    # Ordenar por fecha_ini si tenemos infos
    if infos and len(infos) == len(paths):
        pairs = sorted(zip(paths, infos), key=lambda x: x[1]["fecha_ini"])
        paths = [p for p, _ in pairs]

    state = {"saldo": None, "num_orden": 0, "anualidad": ""}
    all_rows = []
    saldo_ini_total = None
    for path in paths:
        rows, state, saldo_ini = _read_n43_rows(path, state)
        if saldo_ini_total is None:
            saldo_ini_total = saldo_ini
        all_rows.extend(rows)

    _write_rows_to_sheet(ws, all_rows, saldo_ini_total or 0.0)
    return len(all_rows)


def parse_n43(path_n43: str, path_out: str) -> int:
    """Fichero único → Excel."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracto"
    count = _write_n43_sheet(ws, path_n43)
    wb.save(path_out)
    return count


def parse_n43_combined(paths: list, path_out: str, infos: list = None) -> int:
    """Varios ficheros de la misma cuenta → un Excel, una hoja, movimientos encadenados."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracto"
    count = _write_combined_to_sheet(ws, paths, infos)
    wb.save(path_out)
    return count


def parse_n43_accounts(groups: OrderedDict, path_out: str) -> dict:
    """
    Varias cuentas → un Excel con una hoja por cuenta.
    groups: OrderedDict { cuenta: {"paths": [...], "infos": [...]} }
    Devuelve {cuenta: nº_filas}.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    results = {}

    for cuenta, grp in groups.items():
        paths = grp["paths"]
        infos = grp["infos"]
        if infos:
            first_info = min(infos, key=lambda x: x["fecha_ini"])
            sheet_name = f"Cta{cuenta[-4:]}_{first_info['fecha_ini'].strftime('%m%y')}"
        else:
            sheet_name = f"Cta{cuenta[-4:]}"
        sheet_name = sheet_name[:31]

        existing = [s.title for s in wb.worksheets]
        if sheet_name in existing:
            base, n = sheet_name[:28], 2
            while f"{base}_{n}" in existing:
                n += 1
            sheet_name = f"{base}_{n}"

        ws = wb.create_sheet(title=sheet_name)
        if len(paths) == 1:
            count = _write_n43_sheet(ws, paths[0])
        else:
            count = _write_combined_to_sheet(ws, paths, infos)
        results[cuenta] = count

    wb.save(path_out)
    return results


def _excel_name_for_group(cuenta: str, infos: list) -> str:
    """Genera el nombre ENTIDAD_CUENTA4_FECHAINI_FECHAFIN.xlsx para un grupo."""
    first = min(infos, key=lambda x: x["fecha_ini"])
    last  = max(infos, key=lambda x: x["fecha_fin"])
    return (f"{first['entidad']}_{cuenta[-4:]}_"
            f"{first['fecha_ini'].strftime('%Y%m%d')}_"
            f"{last['fecha_fin'].strftime('%Y%m%d')}.xlsx")


def _default_multi_name(groups: OrderedDict) -> str:
    """Nombre por defecto para un Excel que agrupa varias cuentas."""
    first_info = list(groups.values())[0]["infos"][0]
    return f"{first_info['entidad']}_multi.xlsx"


def _parse_group(grp: dict, p_out: str):
    """Escribe un grupo (una cuenta, uno o varios ficheros) en un Excel."""
    if len(grp["paths"]) == 1:
        parse_n43(grp["paths"][0], p_out)
    else:
        parse_n43_combined(grp["paths"], p_out, grp["infos"])


# ─────────────────────────────────────────────────────────────────
#  5. CONSTANTES DE LA GUI
# ─────────────────────────────────────────────────────────────────

WIN_W          = 1100
GAP            = 20
BOTTOM_RESERVE = 75

# Colores para checkboxes según estado
CB_FG_ON   = "#1a1a1a"   # texto activo (negro)
CB_FG_OFF  = "#BBBBBB"   # texto desactivado (gris claro)
CB_SEL_CLR = "#D0E8F7"   # fondo del cuadrito cuando está marcado


# ─────────────────────────────────────────────────────────────────
#  6. RUN_GUI — Interfaz gráfica principal
# ─────────────────────────────────────────────────────────────────

def run_gui():
    BG = "#F0F4F8"
    año_actual = date.today().year

    root = tk.Tk()
    root.title("Conversor Norma 43 → Excel")
    root.resizable(False, False)
    root.configure(bg=BG)

    multi_select_var   = tk.BooleanVar(value=False)
    multi_account_var  = tk.BooleanVar(value=False)
    separate_excel_var = tk.BooleanVar(value=False)

    # ── 6a. Ventana y función de centrado ────────────────────────

    def center_window(w, h):
        sx = root.winfo_screenwidth()
        sy = root.winfo_screenheight()
        x  = (sx - w) // 2
        y  = (sy - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")

    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    # ── 6b. Título y subtítulo ───────────────────────────────────

    tk.Label(root, text="Conversor Archivo en formato Norma 43 → Excel",
             font=("Arial", 18, "bold"), bg="#1F4E79", fg="white",
             pady=15).pack(fill="x")

    frame_body   = tk.Frame(root, bg=BG)
    frame_body.pack(fill="x")
    frame_config = tk.Frame(root, bg=BG)

    tk.Label(frame_body, text="Selecciona un fichero con extensión .n43 "
                        "para convertirlo a Excel .xlsx",
             font=("Arial", 14), bg=BG, fg="#333333").pack(pady=(GAP, 0))

    # ── 6c. Botones y barra de progreso ──────────────────────────

    frame_btns = tk.Frame(frame_body, bg=BG)
    frame_btns.pack(pady=(GAP, 0))

    btn_seleccionar = tk.Button(
        frame_btns, text="📂  Seleccionar fichero .n43",
        font=("Arial", 14, "bold"), bg="#2E75B6", fg="white",
        activebackground="#1F4E79", activeforeground="white",
        bd=0, padx=36, pady=18, cursor="hand2")
    btn_seleccionar.pack(side="left", padx=(0, 25))

    btn_convertir = tk.Button(
        frame_btns, text="⚡  Convertir a Excel",
        font=("Arial", 14, "bold"), bg="#70AD47", fg="white",
        activebackground="#507E35", activeforeground="white",
        bd=0, padx=36, pady=18, cursor="hand2",
        state="disabled")
    btn_convertir.pack(side="left", padx=(25, 0))

    progress = ttk.Progressbar(frame_body, orient="horizontal",
                               length=700, mode="indeterminate")
    progress.pack(pady=(GAP, 0))

    # ── 6d. Listbox ficheros n43 (10 filas + scrollbar) ──────────

    frame_path = tk.Frame(frame_body, bg="#E2EAF4", bd=1, relief="sunken")
    frame_path.pack(fill="x", padx=30, pady=(GAP, 0))
    frame_path_inner = tk.Frame(frame_path, bg="#E2EAF4")
    frame_path_inner.pack(fill="x", padx=18, pady=6)
    tk.Label(frame_path_inner, text="Nombre del fichero Norma-43:",
             font=("Arial", 12, "bold"), bg="#E2EAF4", fg="#555555",
             anchor="w").pack(anchor="w")

    frame_path_lbx = tk.Frame(frame_path_inner, bg="#E2EAF4")
    frame_path_lbx.pack(fill="x", pady=(4, 0))
    path_scrollbar = tk.Scrollbar(frame_path_lbx, orient="vertical")
    path_listbox = tk.Listbox(frame_path_lbx, height=10,
                              font=("Arial", 11), bg="white", fg="#333333",
                              selectmode="extended", bd=1, relief="sunken",
                              yscrollcommand=path_scrollbar.set,
                              activestyle="none")
    path_scrollbar.config(command=path_listbox.yview)
    path_scrollbar.pack(side="right", fill="y")
    path_listbox.pack(side="left", fill="x", expand=True)

    # ── 6e. Recuadro de datos de cabecera ────────────────────────

    frame_header = tk.LabelFrame(
        frame_body, text=" Datos de cabecera ",
        font=("Arial", 13, "bold"), bg=BG, fg="#1F4E79",
        bd=2, relief="groove", padx=10, pady=8)
    frame_header.pack(fill="x", padx=30, pady=(GAP, 0))

    header_labels = {}
    campos = [
        ("Entidad:",       "entidad",          0, 0),
        ("Oficina:",       "oficina",          0, 2),
        ("Cuenta:",        "cuenta",           0, 4),
        ("Movimientos:",   "num_movimientos",  0, 6),
        ("Fecha inicio:",  "fecha_ini",        1, 0),
        ("Saldo inicial:", "saldo_ini",        1, 2),
        ("Fecha fin:",     "fecha_fin",        1, 4),
        ("Saldo final:",   "saldo_fin",        1, 6),
    ]
    for texto, key, row, col in campos:
        tk.Label(frame_header, text=texto,
                 font=("Arial", 13, "bold"), bg=BG, fg="#333333",
                 padx=6, pady=4).grid(row=row, column=col, sticky="e")
        lbl = tk.Label(frame_header, text="—",
                       font=("Arial", 13), bg=BG, fg="#555555",
                       padx=6, pady=4)
        lbl.grid(row=row, column=col + 1, sticky="w")
        header_labels[key] = lbl

    for c in (1, 3, 5, 7):
        frame_header.columnconfigure(c, weight=1)

    # ── 6f. Recuadro nombre fichero(s) Excel ─────────────────────

    frame_excel = tk.Frame(frame_body, bg="#E2EAF4", bd=1, relief="sunken")
    frame_excel.pack(fill="x", padx=30, pady=(GAP, 0))
    frame_excel_inner = tk.Frame(frame_excel, bg="#E2EAF4")
    frame_excel_inner.pack(fill="x", padx=18, pady=6)
    tk.Label(frame_excel_inner, text="Nombre del fichero Excel:",
             font=("Arial", 12, "bold"), bg="#E2EAF4", fg="#555555",
             anchor="w").pack(anchor="w")

    frame_excel_single = tk.Frame(frame_excel_inner, bg="#E2EAF4")
    excel_name_var = tk.StringVar(value="")
    entry_excel_name = tk.Entry(frame_excel_single, textvariable=excel_name_var,
                                font=("Arial", 12), bg="white", fg="#333333",
                                relief="sunken", bd=1, insertbackground="#333333",
                                state="disabled")
    entry_excel_name.pack(fill="x", pady=(4, 0))
    frame_excel_single.pack(fill="x")

    frame_excel_multi = tk.Frame(frame_excel_inner, bg="#E2EAF4")
    excel_scrollbar = tk.Scrollbar(frame_excel_multi, orient="vertical")
    excel_listbox = tk.Listbox(frame_excel_multi, height=10,
                               font=("Arial", 11), bg="white", fg="#333333",
                               bd=1, relief="sunken",
                               yscrollcommand=excel_scrollbar.set,
                               activestyle="none")
    excel_scrollbar.config(command=excel_listbox.yview)
    excel_scrollbar.pack(side="right", fill="y")
    excel_listbox.pack(side="left", fill="x", expand=True)

    def update_excel_display():
        if separate_excel_var.get():
            frame_excel_single.pack_forget()
            frame_excel_multi.pack(fill="x", pady=(4, 0))
        else:
            frame_excel_multi.pack_forget()
            frame_excel_single.pack(fill="x")
        root.update_idletasks()
        root.geometry(f"{WIN_W}x{root.winfo_reqheight()}")

    # ── 6g. Pie de página ────────────────────────────────────────

    tk.Label(frame_body,
             text="El Excel se guarda en la misma carpeta que el fichero .n43",
             font=("Arial", 14), bg=BG, fg="#888888").pack(pady=(GAP, 0))
    tk.Frame(frame_body, bg=BG, height=BOTTOM_RESERVE).pack()

    # ── 6h. Variables de estado ──────────────────────────────────

    selected_paths  = []
    n43_header      = {"infos": [], "groups": OrderedDict(), "excel_names": []}
    n43_dir_var     = tk.StringVar()
    output_dir_var  = tk.StringVar()
    auto_n43_var    = tk.BooleanVar(value=True)
    auto_output_var = tk.BooleanVar(value=True)

    # ── 6h-bis. Panel de configuración ───────────────────────────

    tk.Label(frame_config, text="⚙  Configuración",
             font=("Arial", 16, "bold"), bg=BG, fg="#1F4E79"
             ).pack(pady=(GAP * 2, GAP))

    frm_n43 = tk.LabelFrame(frame_config, text="  Carpeta de ficheros .n43  ",
                             font=("Arial", 12, "bold"), bg=BG, fg="#333333",
                             bd=2, relief="groove", padx=10, pady=8)
    frm_n43.pack(fill="x", padx=30, pady=(0, GAP))
    frm_n43_row = tk.Frame(frm_n43, bg=BG)
    frm_n43_row.pack(fill="x")
    tk.Entry(frm_n43_row, textvariable=n43_dir_var, font=("Arial", 11),
             bg="white", fg="#333333", relief="sunken",
             state="readonly").pack(side="left", fill="x", expand=True)
    tk.Button(frm_n43_row, text="Cambiar", font=("Arial", 11),
              bg="#2E75B6", fg="white", bd=0, padx=14, pady=6, cursor="hand2",
              command=lambda: cambiar_dir(n43_dir_var, "last_dir")
              ).pack(side="left", padx=(8, 0))
    tk.Checkbutton(frm_n43, text="Actualizar Ruta con la última seleccionada",
                   variable=auto_n43_var, bg=BG, fg="#333333",
                   activebackground=BG, font=("Arial", 10),
                   selectcolor=CB_SEL_CLR,
                   command=lambda: save_config({"auto_update_n43_dir": auto_n43_var.get()})
                   ).pack(anchor="w", pady=(4, 0))

    frm_out = tk.LabelFrame(frame_config, text="  Carpeta de salida Excel  ",
                             font=("Arial", 12, "bold"), bg=BG, fg="#333333",
                             bd=2, relief="groove", padx=10, pady=8)
    frm_out.pack(fill="x", padx=30, pady=(0, GAP))
    frm_out_row = tk.Frame(frm_out, bg=BG)
    frm_out_row.pack(fill="x")
    tk.Entry(frm_out_row, textvariable=output_dir_var, font=("Arial", 11),
             bg="white", fg="#333333", relief="sunken",
             state="readonly").pack(side="left", fill="x", expand=True)
    tk.Button(frm_out_row, text="Cambiar", font=("Arial", 11),
              bg="#2E75B6", fg="white", bd=0, padx=14, pady=6, cursor="hand2",
              command=lambda: cambiar_dir(output_dir_var, "last_output_dir")
              ).pack(side="left", padx=(8, 0))
    tk.Checkbutton(frm_out, text="Actualizar Ruta con la última seleccionada",
                   variable=auto_output_var, bg=BG, fg="#333333",
                   activebackground=BG, font=("Arial", 10),
                   selectcolor=CB_SEL_CLR,
                   command=lambda: save_config({"auto_update_output_dir": auto_output_var.get()})
                   ).pack(anchor="w", pady=(4, 0))

    # LabelFrame de opciones de selección
    frm_sel = tk.LabelFrame(frame_config, text="  Opciones de selección  ",
                             font=("Arial", 12, "bold"), bg=BG, fg="#333333",
                             bd=2, relief="groove", padx=10, pady=10)
    frm_sel.pack(fill="x", padx=30, pady=(0, GAP))

    cb_multi_select = tk.Checkbutton(
        frm_sel,
        text="Selección múltiple de ficheros",
        variable=multi_select_var,
        bg=BG, fg=CB_FG_ON,
        activebackground=BG, activeforeground=CB_FG_ON,
        disabledforeground=CB_FG_OFF,
        selectcolor=CB_SEL_CLR,
        font=("Arial", 11, "bold"),
        command=lambda: on_multi_select_toggle())
    cb_multi_select.pack(anchor="w", pady=(0, 4))

    # Separador visual entre nivel 1 y nivel 2
    tk.Frame(frm_sel, bg="#CCCCCC", height=1).pack(fill="x", padx=8, pady=(0, 6))

    cb_multi_account = tk.Checkbutton(
        frm_sel,
        text="Permitir cuentas diferentes",
        variable=multi_account_var,
        bg=BG, fg=CB_FG_OFF,
        activebackground=BG, activeforeground=CB_FG_ON,
        disabledforeground=CB_FG_OFF,
        selectcolor=CB_SEL_CLR,
        font=("Arial", 11),
        state="disabled",
        command=lambda: on_multi_account_toggle())
    cb_multi_account.pack(anchor="w", padx=(20, 0), pady=(0, 4))

    cb_separate_excel = tk.Checkbutton(
        frm_sel,
        text="Guardar cada cuenta en un Excel separado",
        variable=separate_excel_var,
        bg=BG, fg=CB_FG_OFF,
        activebackground=BG, activeforeground=CB_FG_ON,
        disabledforeground=CB_FG_OFF,
        selectcolor=CB_SEL_CLR,
        font=("Arial", 11),
        state="disabled",
        command=lambda: on_separate_excel_toggle())
    cb_separate_excel.pack(anchor="w", padx=(44, 0))

    tk.Button(frame_config, text="✅  Cerrar configuración",
              font=("Arial", 13, "bold"), bg="#70AD47", fg="white",
              activebackground="#507E35", bd=0, padx=30, pady=14,
              cursor="hand2", command=lambda: show_body()
              ).pack(pady=(GAP, 0))
    tk.Frame(frame_config, bg=BG, height=BOTTOM_RESERVE).pack()

    # ── Funciones del panel de configuración ─────────────────────

    def cambiar_dir(var, config_key):
        d = filedialog.askdirectory(title="Seleccionar carpeta",
                                    initialdir=var.get() or "")
        if d:
            var.set(d)
            save_config({config_key: d})

    def show_body():
        frame_config.pack_forget()
        frame_body.pack(fill="x")

    def _refresh_cascade():
        """Actualiza estado visual y funcional de los checkboxes dependientes."""
        if multi_select_var.get():
            cb_multi_account.config(state="normal", fg=CB_FG_ON)
            if multi_account_var.get():
                cb_separate_excel.config(state="normal", fg=CB_FG_ON)
            else:
                cb_separate_excel.config(state="disabled", fg=CB_FG_OFF)
        else:
            cb_multi_account.config(state="disabled", fg=CB_FG_OFF)
            cb_separate_excel.config(state="disabled", fg=CB_FG_OFF)

    def show_config():
        cfg = load_config()
        n43_dir_var.set(cfg.get("last_dir", ""))
        output_dir_var.set(cfg.get("last_output_dir", ""))
        auto_n43_var.set(cfg.get("auto_update_n43_dir", True))
        auto_output_var.set(cfg.get("auto_update_output_dir", True))
        multi_select_var.set(cfg.get("multi_select", False))
        multi_account_var.set(cfg.get("multi_account", False))
        separate_excel_var.set(cfg.get("separate_excel", False))
        _refresh_cascade()
        frame_body.pack_forget()
        frame_config.pack(fill="x")

    def on_multi_select_toggle():
        if not multi_select_var.get():
            multi_account_var.set(False)
            separate_excel_var.set(False)
            update_excel_display()
        _refresh_cascade()
        save_config({
            "multi_select":   multi_select_var.get(),
            "multi_account":  multi_account_var.get(),
            "separate_excel": separate_excel_var.get(),
        })

    def on_multi_account_toggle():
        if not multi_account_var.get():
            separate_excel_var.set(False)
            update_excel_display()
        _refresh_cascade()
        save_config({
            "multi_account":  multi_account_var.get(),
            "separate_excel": separate_excel_var.get(),
        })

    def on_separate_excel_toggle():
        save_config({"separate_excel": separate_excel_var.get()})
        update_excel_display()
        # Actualizar nombres en listbox si ya hay ficheros cargados
        _refresh_excel_names()

    def _refresh_excel_names():
        """Recalcula y muestra los nombres Excel según la configuración actual."""
        groups = n43_header.get("groups", OrderedDict())
        excel_names = n43_header.get("excel_names", [])
        if not groups:
            return
        if separate_excel_var.get():
            excel_listbox.delete(0, "end")
            for name in excel_names:
                excel_listbox.insert("end", name)
        else:
            if len(groups) == 1:
                excel_name_var.set(excel_names[0] if excel_names else "")
                entry_excel_name.config(state="normal")
            else:
                excel_name_var.set(_default_multi_name(groups))
                entry_excel_name.config(state="normal")

    # ── 6i. seleccionar() — callback del botón seleccionar ───────

    def seleccionar():
        cfg = load_config()

        if multi_select_var.get():
            paths = list(filedialog.askopenfilenames(
                title="Selecciona ficheros N43",
                initialdir=cfg.get("last_dir", ""),
                filetypes=[("Ficheros N43", "*.n43 *.N43"),
                           ("Todos los ficheros", "*.*")]
            ))
        else:
            p = filedialog.askopenfilename(
                title="Selecciona el fichero N43",
                initialdir=cfg.get("last_dir", ""),
                filetypes=[("Ficheros N43", "*.n43 *.N43"),
                           ("Todos los ficheros", "*.*")]
            )
            paths = [p] if p else []

        if not paths:
            return

        # Leer todas las cabeceras
        try:
            infos = [read_n43_header(p) for p in paths]
        except Exception as e:
            messagebox.showerror("Error al leer cabeceras", str(e))
            return

        # Filtrar automáticamente si hay cuentas mezcladas sin permiso
        if len(paths) > 1 and not multi_account_var.get():
            primera_cuenta = infos[0]["cuenta"]
            descartados = [(p, i) for p, i in zip(paths, infos)
                           if i["cuenta"] != primera_cuenta]
            if descartados:
                n_desc = len(descartados)
                ctas_desc = ", ".join(
                    sorted(set(i["cuenta"][-4:] for _, i in descartados)))
                messagebox.showinfo(
                    "Ficheros filtrados automáticamente",
                    f"Se han excluido {n_desc} fichero(s) porque pertenecen\n"
                    f"a cuentas distintas a la primera seleccionada "
                    f"({primera_cuenta[-4:]}).\n\n"
                    f"Cuentas excluidas: {ctas_desc}\n\n"
                    f"Activa 'Permitir cuentas diferentes' en la configuración\n"
                    f"si deseas trabajar con varias cuentas a la vez.")
                pairs = [(p, i) for p, i in zip(paths, infos)
                         if i["cuenta"] == primera_cuenta]
                paths = [p for p, _ in pairs]
                infos = [i for _, i in pairs]

        selected_paths.clear()
        selected_paths.extend(paths)

        if cfg.get("auto_update_n43_dir", True):
            save_config({"last_dir": os.path.dirname(paths[0])})

        # Listbox de ficheros n43
        path_listbox.delete(0, "end")
        for p in paths:
            path_listbox.insert("end", os.path.basename(p))

        if len(paths) == 1:
            btn_seleccionar.config(text="📂  Cambiar el fichero .n43",
                                   fg="#FFE066")
        else:
            btn_seleccionar.config(
                text=f"📂  Cambiar los {len(paths)} ficheros .n43",
                fg="#FFE066")
        btn_convertir.config(state="normal", bg="#70AD47")

        # Limpiar display Excel
        excel_name_var.set("")
        entry_excel_name.config(state="disabled")
        excel_listbox.delete(0, "end")

        # Agrupar por cuenta (OrderedDict mantiene orden de aparición)
        groups = OrderedDict()
        for path, info in zip(paths, infos):
            cuenta = info["cuenta"]
            if cuenta not in groups:
                groups[cuenta] = {"paths": [], "infos": []}
            groups[cuenta]["paths"].append(path)
            groups[cuenta]["infos"].append(info)

        # Ordenar cada grupo por fecha_ini
        for g in groups.values():
            pairs = sorted(zip(g["paths"], g["infos"]),
                           key=lambda x: x[1]["fecha_ini"])
            g["paths"] = [p for p, _ in pairs]
            g["infos"] = [i for _, i in pairs]

        n43_header["infos"]  = infos
        n43_header["groups"] = groups

        # Generar nombres Excel (uno por cuenta)
        excel_names = [_excel_name_for_group(c, g["infos"])
                       for c, g in groups.items()]
        n43_header["excel_names"] = excel_names

        # Actualizar cabecera resumen
        try:
            all_infos = infos
            if len(groups) == 1:
                cuenta   = list(groups.keys())[0]
                grp_infos = list(groups.values())[0]["infos"]
                header_labels["entidad"].config(text=grp_infos[0]["entidad"])
                header_labels["oficina"].config(text=grp_infos[0]["oficina"])
                header_labels["cuenta"].config(text=cuenta)
                header_labels["num_movimientos"].config(
                    text=str(sum(i["num_movimientos"] for i in grp_infos)))
                header_labels["fecha_ini"].config(
                    text=min(i["fecha_ini"] for i in grp_infos).strftime("%d/%m/%Y"))
                header_labels["fecha_fin"].config(
                    text=max(i["fecha_fin"] for i in grp_infos).strftime("%d/%m/%Y"))
                # Saldo sólo tiene sentido para un único fichero
                if len(grp_infos) == 1:
                    header_labels["saldo_ini"].config(
                        text=fmt_saldo(grp_infos[0]["saldo_ini"]))
                    header_labels["saldo_fin"].config(
                        text=fmt_saldo(grp_infos[0].get("saldo_fin", 0)))
                else:
                    header_labels["saldo_ini"].config(text="—")
                    header_labels["saldo_fin"].config(text="—")
            else:
                ctas = list(groups.keys())
                resumen = ", ".join(c[-4:] for c in ctas[:4])
                if len(ctas) > 4:
                    resumen += "…"
                header_labels["entidad"].config(text=all_infos[0]["entidad"])
                header_labels["oficina"].config(text="—")
                header_labels["cuenta"].config(text=resumen)
                header_labels["num_movimientos"].config(
                    text=str(sum(i["num_movimientos"] for i in all_infos)))
                header_labels["fecha_ini"].config(
                    text=min(i["fecha_ini"] for i in all_infos).strftime("%d/%m/%Y"))
                header_labels["fecha_fin"].config(
                    text=max(i["fecha_fin"] for i in all_infos).strftime("%d/%m/%Y"))
                header_labels["saldo_ini"].config(text="—")
                header_labels["saldo_fin"].config(text="—")
        except Exception:
            for lbl in header_labels.values():
                lbl.config(text="—")

        _refresh_excel_names()

    # ── 6j. convertir() — callback del botón convertir ───────────

    def convertir():
        if not selected_paths:
            return

        cfg    = load_config()
        groups = n43_header.get("groups", OrderedDict())
        if not groups:
            return

        n_accounts = len(groups)

        def _finish_conversion(open_path):
            progress.stop()
            progress["value"] = 100
            btn_seleccionar.config(text="📂  Seleccionar otro fichero .n43", fg="white")
            btn_convertir.config(state="disabled", bg="#A8D08D")
            try:
                os.startfile(open_path)
            except Exception:
                pass

        if not separate_excel_var.get():
            # ── Un único fichero Excel ────────────────────────────
            if n_accounts == 1:
                cuenta     = list(groups.keys())[0]
                grp        = list(groups.values())[0]
                excel_name = excel_name_var.get().strip()
                if not excel_name:
                    excel_name = _excel_name_for_group(cuenta, grp["infos"])
                elif not excel_name.lower().endswith(".xlsx"):
                    excel_name += ".xlsx"
            else:
                excel_name = excel_name_var.get().strip() or _default_multi_name(groups)
                if not excel_name.lower().endswith(".xlsx"):
                    excel_name += ".xlsx"

            p_out = filedialog.asksaveasfilename(
                title="Guardar Excel como...",
                initialdir=cfg.get("last_output_dir",
                                   os.path.dirname(selected_paths[0])),
                initialfile=excel_name,
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"),
                           ("Todos los ficheros", "*.*")]
            )
            if not p_out:
                return

            progress.start(10)
            root.update()
            try:
                if n_accounts == 1:
                    _parse_group(list(groups.values())[0], p_out)
                else:
                    parse_n43_accounts(groups, p_out)

                if cfg.get("auto_update_output_dir", True):
                    save_config({"last_output_dir": os.path.dirname(p_out)})

                n_src = (os.path.basename(selected_paths[0])
                         if len(selected_paths) == 1
                         else f"{len(selected_paths)} ficheros")
                messagebox.showinfo(
                    "✅ Conversión completada",
                    f"Se ha importado:\n{n_src}\n\n"
                    f"1º. He interpretado su contenido,\n"
                    f"2º. Lo he convertido en una Tabla,\n"
                    f"3º. Y guardado en la ruta:\n"
                    f"{os.path.dirname(p_out)}\n\n"
                    f"Con el nombre:\n{os.path.basename(p_out)}")

                excel_name_var.set(os.path.basename(p_out))
                entry_excel_name.config(state="disabled")
                _finish_conversion(os.path.dirname(p_out))

            except Exception as e:
                progress.stop()
                messagebox.showerror("Error en la conversión", str(e))

        else:
            # ── Un Excel por cuenta ───────────────────────────────
            out_dir = filedialog.askdirectory(
                title="Selecciona la carpeta donde guardar los Excel",
                initialdir=cfg.get("last_output_dir",
                                   os.path.dirname(selected_paths[0]))
            )
            if not out_dir:
                return

            excel_names = n43_header.get("excel_names", [])
            out_paths   = [os.path.join(out_dir, name) for name in excel_names]

            progress.start(10)
            root.update()
            try:
                for grp, p_out in zip(groups.values(), out_paths):
                    _parse_group(grp, p_out)

                if cfg.get("auto_update_output_dir", True):
                    save_config({"last_output_dir": out_dir})

                excel_listbox.delete(0, "end")
                for name in excel_names:
                    excel_listbox.insert("end", name)

                messagebox.showinfo(
                    "✅ Conversión completada",
                    f"Se han generado {n_accounts} fichero(s) Excel\n"
                    f"en la carpeta:\n{out_dir}")

                _finish_conversion(out_dir)

            except Exception as e:
                progress.stop()
                messagebox.showerror("Error en la conversión", str(e))

    btn_seleccionar.config(command=seleccionar)
    btn_convertir.config(command=convertir)

    # ── 6k. Imagen Caracolillo y crédito ─────────────────────────

    IMG_H = 40
    right_block = tk.Frame(root, bg=BG)
    img_path = resource_path("Caracolillo_Fósil.png")

    img_ok = False
    if os.path.isfile(img_path):
        if HAS_PIL:
            try:
                pil_img = Image.open(img_path)
                ratio   = IMG_H / pil_img.height
                pil_img = pil_img.resize(
                    (int(pil_img.width * ratio), IMG_H), RESAMPLE)
                root._caracolillo = ImageTk.PhotoImage(pil_img)
                img_ok = True
            except Exception:
                pass
        else:
            try:
                raw    = tk.PhotoImage(file=img_path)
                factor = max(1, raw.height() // IMG_H)
                root._caracolillo = raw.subsample(factor, factor)
                img_ok = True
            except Exception:
                pass

    lbl_hint = tk.Label(right_block, text="⚙  Configurar",
                        font=("Arial", 9, "italic"), bg=BG, fg="#555555",
                        cursor="hand2")
    lbl_hint.grid(row=0, column=0, sticky="e")
    lbl_hint.grid_remove()

    if img_ok:
        lbl_img = tk.Label(right_block, image=root._caracolillo,
                           bg=BG, cursor="hand2")
    else:
        lbl_img = tk.Canvas(right_block, width=IMG_H, height=IMG_H,
                            bg="white", highlightthickness=1,
                            highlightbackground="red", cursor="hand2")
    lbl_img.grid(row=1, column=0, sticky="e")

    tk.Label(right_block, text=f"Dugarry'{año_actual}",
             font=("Arial", 8), bg=BG, fg="#AAAAAA").grid(row=2, column=0, sticky="e")

    def _on_enter(_):
        lbl_hint.grid()

    def _on_leave(_):
        lbl_hint.grid_remove()

    lbl_img.bind("<Enter>",    _on_enter)
    lbl_img.bind("<Leave>",    _on_leave)
    lbl_hint.bind("<Enter>",   _on_enter)
    lbl_hint.bind("<Leave>",   _on_leave)
    lbl_img.bind("<Button-1>",  lambda _: show_config())
    lbl_hint.bind("<Button-1>", lambda _: show_config())

    right_block.place(relx=1.0, rely=1.0, x=-20, y=-20, anchor="se")

    # ── 6l. Inicialización desde config y arranque ───────────────

    startup_cfg = load_config()
    multi_select_var.set(startup_cfg.get("multi_select", False))
    multi_account_var.set(startup_cfg.get("multi_account", False))
    separate_excel_var.set(startup_cfg.get("separate_excel", False))
    _refresh_cascade()
    update_excel_display()

    root.update_idletasks()
    center_window(WIN_W, root.winfo_reqheight())
    root.mainloop()


# ─────────────────────────────────────────────────────────────────
#  7. PUNTO DE ENTRADA
# ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    run_gui()
