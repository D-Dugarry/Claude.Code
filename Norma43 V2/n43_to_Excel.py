# ═══════════════════════════════════════════════════════════════════
#  n43_to_Excel — Conversor de ficheros Norma 43 (AEB 43) a Excel
#  Versión: v02
# ═══════════════════════════════════════════════════════════════════
#
#  Descripción:
#    Aplicación de escritorio con interfaz gráfica (tkinter) que
#    lee un fichero bancario en formato Norma 43 (.n43) y genera
#    un libro Excel (.xlsx) con los movimientos formateados.
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
#  Cambios respecto a v01:
#    - Ancho de ventana ampliado a 1100px
#    - Todos los elementos siempre visibles (sin pack_forget)
#    - Botones y barra de progreso bajo el subtítulo
#    - Nombre del Excel con recuadro, debajo de los datos de cabecera
#    - Leyenda "Nombre del fichero Excel:" siempre visible
#    - Altura de ventana calculada una sola vez al arrancar
#
# ═══════════════════════════════════════════════════════════════════
#
#  ÍNDICE DE BLOQUES                                         Línea
#  ─────────────────────────────────────────────────────────────────
#   1. Imports y carga condicional de Pillow ................   63
#   2. read_n43_header()  — lectura de cabecera N43 ........   87
#   3. fmt_saldo()        — formato numérico español .......  137
#   4. parse_n43()        — parser completo N43 → Excel ....  153
#      4a. Estilos y cabecera de la hoja ...................  182
#      4b. Lectura línea a línea del fichero ...............  197
#      4c. Escritura de filas con estilos alternados .......  255
#      4d. Fila de totales y ajustes finales ...............  296
#   5. Constantes de la GUI ................................  339
#   6. run_gui()  — interfaz gráfica principal .............  349
#      6a. Ventana y función de centrado ...................  387
#      6b. Título y subtítulo ..............................  402
#      6c. Botones y barra de progreso .....................  412
#      6d. Ruta del fichero seleccionado ...................  439
#      6e. Recuadro de datos de cabecera ...................  450
#      6f. Recuadro nombre fichero Excel ...................  487
#      6g. Pie de página ...................................  509
#      6h. Variables de estado .............................  521
#      6i. seleccionar()  — callback botón seleccionar .....  526
#      6j. convertir()    — callback botón convertir .......  589
#      6k. Imagen Caracolillo y crédito ....................  648
#      6l. Cálculo del alto y arranque .....................  702
#   7. Punto de entrada (__main__) .........................  716
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
from datetime import date
import os
import sys

# Pillow se usa para escalar la imagen del caracolillo con calidad
# (LANCZOS). Si no está instalado, se usa el subsample nativo de
# tkinter como fallback (menor calidad pero funcional).
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


# ─────────────────────────────────────────────────────────────────
#  2. READ_N43_HEADER — Lectura de cabecera del fichero N43
# ─────────────────────────────────────────────────────────────────

def read_n43_header(path_n43: str) -> dict:
    """
    Lee los datos de cabecera de un fichero Norma 43.

    Extrae información de dos tipos de registro:
      - Registro 11 (cabecera de cuenta):
            Pos  2-5   → Código entidad (4 dígitos)
            Pos  6-9   → Código oficina (4 dígitos)
            Pos 10-19  → Número de cuenta (10 dígitos)
            Pos 20-25  → Fecha inicio (AAMMDD)
            Pos 26-31  → Fecha fin    (AAMMDD)
            Pos 32     → Signo saldo inicial (1=deudor, 2=acreedor)
            Pos 33-46  → Saldo inicial en céntimos (14 dígitos)
            Pos 47-49  → Código divisa (978=EUR)
      - Registro 33 (cierre de cuenta):
            Pos 58     → Signo saldo final
            Pos 59-72  → Saldo final en céntimos (14 dígitos)

    Args:
        path_n43: Ruta al fichero .n43

    Returns:
        dict con claves: entidad, oficina, cuenta, fecha_ini (date),
        fecha_fin (date), saldo_ini (float), divisa, saldo_fin (float)
    """
    info = {}
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
            elif reg == "33":
                signo = -1 if linea[58] == "1" else 1
                info["saldo_fin"] = int(linea[59:73]) * signo * 0.01
    return info


# ─────────────────────────────────────────────────────────────────
#  3. FMT_SALDO — Formato numérico estilo español
# ─────────────────────────────────────────────────────────────────

def fmt_saldo(valor: float) -> str:
    """
    Formatea un valor monetario al estilo español.
    Ejemplo: 1234.50 → '1.234,50 €'

    Usa un triple-replace para convertir la salida de f-string
    (que usa coma para miles y punto para decimales en inglés)
    al formato español (punto para miles, coma para decimales).
    """
    return f"{valor:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


# ─────────────────────────────────────────────────────────────────
#  4. PARSE_N43 — Parser completo del fichero N43 → Excel
# ─────────────────────────────────────────────────────────────────

def parse_n43(path_n43: str, path_out: str) -> int:
    """
    Lee un fichero Norma 43, extrae todos los movimientos y genera
    un libro Excel (.xlsx) con formato profesional.

    Columnas generadas:
      Orden | Oficina | F.Operación | F.Valor | Importe |
      Saldo | Documento | Referencia | RCM01..RCM05

    Args:
        path_n43: Ruta al fichero .n43 de entrada
        path_out: Ruta del fichero .xlsx de salida

    Returns:
        Número de movimientos importados (registros tipo 22)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracto"

    headers = [
        "Orden", "Oficina", "F.Operación", "F.Valor", "Importe",
        "Saldo", "Documento", "Referencia",
        "RCM01", "RCM02", "RCM03", "RCM04", "RCM05"
    ]

    # ── 4a. Estilos y cabecera de la hoja ────────────────────────
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

    # ── 4b. Lectura línea a línea del fichero ────────────────────
    #
    #  El saldo se calcula acumulativamente: parte del saldo inicial
    #  (registro 11) y suma cada importe de movimiento (registro 22).
    #
    #  El campo "Orden" se construye como AÑO + número secuencial
    #  de 6 dígitos, reiniciándose cuando cambia la anualidad de
    #  la fecha de valor.
    #
    #  Los registros 23 contienen conceptos complementarios (RCM01
    #  a RCM05). Se asocian al último movimiento leído.

    rows, saldo_acum, num_orden = [], 0.0, 0
    current_row = None
    anualidad = ""

    with open(path_n43, encoding="latin-1") as f:
        for linea in f:
            reg = linea[:2]

            if reg == "11":
                # Cabecera: obtener saldo inicial para acumular
                signo = -1 if linea[32] == "1" else 1
                saldo_acum = int(linea[33:47]) * signo * 0.01

            elif reg == "22":
                # Movimiento: extraer datos y acumular saldo
                año = "20" + linea[10:12]
                if linea[16:18] != anualidad:
                    num_orden = 0
                    anualidad = linea[16:18]
                num_orden += 1
                signo = -1 if linea[27] == "1" else 1
                importe = int(linea[28:42]) * signo * 0.01
                saldo_acum += importe
                current_row = [
                    f"{año}{num_orden:06d}",               # Orden
                    linea[6:10].strip(),                    # Oficina origen
                    date(int("20" + linea[10:12]),          # F.Operación
                         int(linea[12:14]),
                         int(linea[14:16])),
                    date(int("20" + linea[16:18]),          # F.Valor
                         int(linea[18:20]),
                         int(linea[20:22])),
                    importe,                                # Importe
                    round(saldo_acum, 2),                   # Saldo acumulado
                    linea[42:52].strip(),                   # Documento
                    linea[52:102].strip(),                  # Referencia
                    "", "", "", "", ""                      # RCM01..RCM05
                ]
                rows.append(current_row)

            elif reg == "23" and current_row:
                # Concepto complementario (1 a 5)
                idx = int(linea[2:4])
                if 1 <= idx <= 5:
                    current_row[7 + idx] = linea[4:79].strip()

    # ── 4c. Escritura de filas con estilos alternados ────────────
    #
    #  Filas pares: fondo azul claro (#EBF3FB)
    #  Filas impares: fondo blanco
    #  Formato de fecha: DD/MM/YYYY
    #  Formato de importes: #,##0.00 € (negativos en rojo)
    #  Formato de saldo: #,##0.00 € (siempre positivo en display)

    fill_even    = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
    fill_odd     = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    font_data    = Font(name="Arial", size=9)
    align_center = Alignment(horizontal="center", vertical="center")
    align_right  = Alignment(horizontal="right",  vertical="center")
    align_left   = Alignment(horizontal="left",   vertical="center")

    date_fmt  = "DD/MM/YYYY"
    money_fmt = '#,##0.00 €;[RED]-#,##0.00 €'
    money_pos = '#,##0.00 €'

    for r_idx, row in enumerate(rows, 2):
        fill = fill_even if r_idx % 2 == 0 else fill_odd
        for c_idx, val in enumerate(row, 1):
            cell        = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font   = font_data
            cell.fill   = fill
            cell.border = cell_border

            if c_idx in (3, 4):          # Fechas
                cell.number_format = date_fmt
                cell.alignment     = align_center
            elif c_idx == 5:             # Importe (puede ser negativo)
                cell.number_format = money_fmt
                cell.alignment     = align_right
            elif c_idx == 6:             # Saldo
                cell.number_format = money_pos
                cell.alignment     = align_right
            elif c_idx in (1, 2, 7):     # Orden, Oficina, Documento
                cell.alignment = align_center
            else:                        # Referencia, RCMs
                cell.alignment = align_left

    # ── 4d. Fila de totales y ajustes finales ────────────────────

    # Anchos de columna predefinidos
    col_widths = [14, 8, 12, 12, 14, 14, 12, 30, 25, 25, 25, 25, 25]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Fila de totales con fondo azul y texto blanco
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

    # Suma de importes (fórmula Excel)
    ws.cell(row=total_row, column=5,
            value=f"=SUM(E2:E{last_data_row})").number_format = money_fmt
    ws.cell(row=total_row, column=5).alignment = align_right

    # Saldo final = suma importes + saldo inicial
    ws.cell(row=total_row, column=6,
            value=f"=E{total_row}+{rows[0][5] - rows[0][4] if rows else 0:.2f}"
            ).number_format = money_pos

    # Inmovilizar primera fila y activar autofiltro
    ws.freeze_panes  = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_data_row}"

    wb.save(path_out)
    return len(rows)


# ─────────────────────────────────────────────────────────────────
#  5. CONSTANTES DE LA GUI
# ─────────────────────────────────────────────────────────────────

WIN_W          = 1100    # Ancho fijo de la ventana (px)
GAP            = 20      # Separación uniforme entre elementos (px)
BOTTOM_RESERVE = 75      # Reserva inferior: caracolillo (40px)
                         #   + crédito (~15px) + margen (20px)


# ─────────────────────────────────────────────────────────────────
#  6. RUN_GUI — Interfaz gráfica principal
# ─────────────────────────────────────────────────────────────────

def run_gui():
    """
    Construye y ejecuta la interfaz gráfica del conversor.

    Disposición de elementos (de arriba a abajo):
      1. Barra de título (azul oscuro)
      2. Subtítulo con instrucciones
      3. Botones (seleccionar + convertir) y barra de progreso
      4. Ruta del fichero .n43 seleccionado (frame sunken)
      5. Datos de cabecera N43 (LabelFrame con grid 2×8)
      6. Nombre del fichero Excel generado (frame sunken)
      7. Pie de página (texto gris)
      8. Espaciador + Caracolillo y crédito (esquina inf-dcha)

    Todos los elementos están siempre visibles. Los campos de
    datos se rellenan al seleccionar un fichero, y el nombre
    del Excel se muestra tras la conversión.

    Ciclo de estados de los botones:
      ┌──────────────────┬─────────────────────────┬───────────────────┐
      │ Estado           │ Botón Seleccionar       │ Botón Convertir   │
      ├──────────────────┼─────────────────────────┼───────────────────┤
      │ Inicio           │ #2E75B6 "Seleccionar…"  │ disabled          │
      │ Fichero cargado  │ #8FACCF "Cambiar…"      │ #70AD47 enabled   │
      │ Tras convertir   │ #2E75B6 "Seleccionar…"  │ #A8D08D disabled  │
      └──────────────────┴─────────────────────────┴───────────────────┘
    """
    BG = "#F0F4F8"
    año_actual = date.today().year

    root = tk.Tk()
    root.title("Conversor Norma 43 → Excel")
    root.resizable(False, False)
    root.configure(bg=BG)

    # ── 6a. Ventana y función de centrado ────────────────────────

    def center_window(w, h):
        """Centra la ventana en la pantalla según resolución."""
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

    tk.Label(root, text="Selecciona un fichero con extensión .n43 "
                        "para convertirlo a Excel .xlsx",
             font=("Arial", 14), bg=BG, fg="#333333").pack(pady=(GAP, 0))

    # ── 6c. Botones y barra de progreso ──────────────────────────
    #    Colocados justo debajo del subtítulo.
    #    Separación de 50px entre ambos botones (25px + 25px).
    #    Los comandos se asignan más abajo, tras definir las funciones.

    frame_btns = tk.Frame(root, bg=BG)
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

    progress = ttk.Progressbar(root, orient="horizontal",
                               length=700, mode="indeterminate")
    progress.pack(pady=(GAP, 0))

    # ── 6d. Ruta del fichero seleccionado ────────────────────────
    #    Frame sunken que muestra el nombre del fichero .n43.
    #    Texto inicial: "Ningún fichero seleccionado"

    path_var = tk.StringVar(value="Ningún fichero seleccionado")
    frame_path = tk.Frame(root, bg="#E2EAF4", bd=1, relief="sunken")
    frame_path.pack(fill="x", padx=30, pady=(GAP, 0))
    tk.Label(frame_path, textvariable=path_var, font=("Arial", 12),
             bg="#E2EAF4", fg="#555555", anchor="w",
             wraplength=520, justify="left", padx=18, pady=6).pack(fill="x")

    # ── 6e. Recuadro de datos de cabecera ────────────────────────
    #    LabelFrame con grid de 2 filas, siempre visible:
    #      Fila 0: Entidad | Oficina | Cuenta
    #      Fila 1: Fecha inicio | Saldo inicial | Fecha fin | Saldo final
    #    Los valores muestran "—" hasta que se selecciona un fichero.

    frame_header = tk.LabelFrame(
        root, text=" Datos de cabecera ",
        font=("Arial", 13, "bold"), bg=BG, fg="#1F4E79",
        bd=2, relief="groove", padx=10, pady=8)
    frame_header.pack(fill="x", padx=30, pady=(GAP, 0))

    header_labels = {}
    campos = [
        # (etiqueta,         clave,       fila, columna)
        ("Entidad:",       "entidad",   0, 0),
        ("Oficina:",       "oficina",   0, 2),
        ("Cuenta:",        "cuenta",    0, 4),
        ("Fecha inicio:",  "fecha_ini", 1, 0),
        ("Saldo inicial:", "saldo_ini", 1, 2),
        ("Fecha fin:",     "fecha_fin", 1, 4),
        ("Saldo final:",   "saldo_fin", 1, 6),
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

    # Distribuir columnas de valores equitativamente
    for c in (1, 3, 5, 7):
        frame_header.columnconfigure(c, weight=1)

    # ── 6f. Recuadro nombre fichero Excel ────────────────────────
    #    Frame sunken (mismo estilo que la ruta del .n43).
    #    La leyenda "Nombre del fichero Excel:" siempre visible.
    #    El valor se rellena tras la conversión y se limpia al
    #    seleccionar otro fichero.

    frame_excel = tk.Frame(root, bg="#E2EAF4", bd=1, relief="sunken")
    frame_excel.pack(fill="x", padx=30, pady=(GAP, 0))

    excel_name_var = tk.StringVar(value="")
    frame_excel_inner = tk.Frame(frame_excel, bg="#E2EAF4")
    frame_excel_inner.pack(fill="x", padx=18, pady=6)

    tk.Label(frame_excel_inner, text="Nombre del fichero Excel:",
             font=("Arial", 12, "bold"), bg="#E2EAF4", fg="#555555",
             anchor="w").pack(side="left")

    lbl_excel_name = tk.Label(frame_excel_inner, textvariable=excel_name_var,
                              font=("Arial", 12), bg="#E2EAF4", fg="#333333",
                              anchor="w", padx=8)
    lbl_excel_name.pack(side="left", fill="x", expand=True)

    # ── 6g. Pie de página ────────────────────────────────────────

    lbl_footer = tk.Label(root,
             text="El Excel se guarda en la misma carpeta que el fichero .n43",
             font=("Arial", 14), bg=BG, fg="#888888")
    lbl_footer.pack(pady=(GAP, 0))

    # Espaciador inferior: reserva espacio para el bloque caracolillo
    # que está posicionado con place() y no participa del pack.
    spacer = tk.Frame(root, bg=BG, height=BOTTOM_RESERVE)
    spacer.pack()

    # ── 6h. Variables de estado ──────────────────────────────────

    selected_path = {"value": None}
    n43_header    = {"info": None, "excel_name": None}

    # ── 6i. seleccionar() — callback del botón seleccionar ───────
    #
    #  Flujo:
    #    1. Abrir diálogo de selección de fichero (.n43)
    #    2. Actualizar la ruta mostrada en pantalla
    #    3. Cambiar estado visual del botón a "Cambiar…" (atenuado)
    #    4. Restaurar botón convertir (por si venía de conversión)
    #    5. Limpiar nombre del Excel (se mostrará tras convertir)
    #    6. Leer cabecera del N43 y rellenar el grid
    #    7. Calcular el nombre del fichero Excel de salida

    def seleccionar():
        p = filedialog.askopenfilename(
            title="Selecciona el fichero N43",
            filetypes=[("Ficheros N43", "*.n43 *.N43"),
                       ("Todos los ficheros", "*.*")]
        )
        if p:
            # 1-2. Guardar ruta y mostrar nombre
            selected_path["value"] = p
            path_var.set(os.path.basename(p))

            # 3. Atenuar botón seleccionar
            btn_seleccionar.config(text="📂  Cambiar el fichero .n43",
                                   bg="#8FACCF")

            # 4. Restaurar botón convertir
            btn_convertir.config(state="normal", bg="#70AD47")

            # 5. Limpiar nombre del Excel (leyenda se mantiene visible)
            excel_name_var.set("")

            # 6-7. Leer cabecera y generar nombre Excel
            try:
                info = read_n43_header(p)
                n43_header["info"] = info

                header_labels["entidad"].config(text=info["entidad"])
                header_labels["oficina"].config(text=info["oficina"])
                header_labels["cuenta"].config(text=info["cuenta"])
                header_labels["fecha_ini"].config(
                    text=info["fecha_ini"].strftime("%d/%m/%Y"))
                header_labels["fecha_fin"].config(
                    text=info["fecha_fin"].strftime("%d/%m/%Y"))
                header_labels["saldo_ini"].config(
                    text=fmt_saldo(info["saldo_ini"]))
                header_labels["saldo_fin"].config(
                    text=fmt_saldo(info.get("saldo_fin", 0)))

                # Nombre: ENTIDAD_CUENTA(4últ)_FECHAINI_FECHAFIN.xlsx
                n43_header["excel_name"] = (
                    f"{info['entidad']}_"
                    f"{info['cuenta'][-4:]}_"
                    f"{info['fecha_ini'].strftime('%Y%m%d')}_"
                    f"{info['fecha_fin'].strftime('%Y%m%d')}"
                    f".xlsx"
                )
            except Exception:
                for lbl in header_labels.values():
                    lbl.config(text="—")
                n43_header["info"] = None
                n43_header["excel_name"] = None

    # ── 6j. convertir() — callback del botón convertir ───────────
    #
    #  Flujo:
    #    1. Determinar ruta de salida (nombre personalizado o default)
    #    2. Iniciar barra de progreso
    #    3. Ejecutar parse_n43()
    #    4. Mostrar mensaje de éxito con nº de movimientos
    #    5. Actualizar estados de botones (seleccionar restaurado,
    #       convertir atenuado)
    #    6. Mostrar nombre del fichero Excel generado
    #    7. Abrir la carpeta del fichero en el explorador

    def convertir():
        if not selected_path["value"]:
            return
        p_in = selected_path["value"]

        # 1. Ruta de salida
        if n43_header.get("excel_name"):
            p_out = os.path.join(os.path.dirname(p_in),
                                 n43_header["excel_name"])
        else:
            p_out = os.path.splitext(p_in)[0] + ".xlsx"

        # 2. Barra de progreso
        progress.start(10)
        root.update()

        try:
            # 3. Conversión
            n = parse_n43(p_in, p_out)
            progress.stop()
            progress["value"] = 100

            # 4. Mensaje de éxito
            messagebox.showinfo(
                "✅ Conversión completada",
                f"Se han importado {n} movimientos.\n\n"
                f"Fichero guardado en:\n{p_out}")

            # 5. Actualizar estados de botones
            btn_seleccionar.config(
                text="📂  Seleccionar otro fichero .n43", bg="#2E75B6")
            btn_convertir.config(state="disabled", bg="#A8D08D")

            # 6. Mostrar nombre del Excel generado en el recuadro
            excel_name_var.set(n43_header.get("excel_name", ""))

            # 7. Abrir carpeta en explorador
            os.startfile(os.path.dirname(p_out))

        except Exception as e:
            progress.stop()
            messagebox.showerror("Error en la conversión", str(e))

    # Asignar comandos a los botones (definidos después de las funciones)
    btn_seleccionar.config(command=seleccionar)
    btn_convertir.config(command=convertir)

    # ── 6k. Imagen Caracolillo y crédito ─────────────────────────
    #    Bloque posicionado con place() en la esquina inferior
    #    derecha (20px de margen). No participa del flujo pack,
    #    así que permanece fijo independientemente del contenido.
    #
    #    La imagen se busca como "Caracolillo_Fósil.png" en la
    #    misma carpeta del script. Se escala a 40px de alto.
    #    Fallback si no existe o no se puede cargar: recuadro
    #    blanco con borde rojo.

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
            # Fallback: subsample nativo de tkinter (sin Pillow)
            try:
                raw    = tk.PhotoImage(file=img_path)
                factor = max(1, raw.height() // IMG_H)
                root._caracolillo = raw.subsample(factor, factor)
                img_ok = True
            except Exception:
                pass

    if img_ok:
        tk.Label(right_block, image=root._caracolillo,
                 bg=BG).pack(anchor="e")
    else:
        # Placeholder: recuadro rojo si la imagen no se encuentra
        tk.Canvas(right_block, width=IMG_H, height=IMG_H,
                  bg="white", highlightthickness=1,
                  highlightbackground="red").pack(anchor="e")

    # Crédito: Dugarry'AñoActual
    tk.Label(right_block, text=f"Dugarry'{año_actual}",
             font=("Arial", 8), bg=BG, fg="#AAAAAA").pack(anchor="e")

    # Anclar a esquina inferior derecha con 20px de margen
    right_block.place(relx=1.0, rely=1.0, x=-20, y=-20, anchor="se")

    # ── 6l. Cálculo del alto y arranque ──────────────────────────
    #
    #  Todos los elementos están siempre visibles, así que el alto
    #  necesario se obtiene directamente de winfo_reqheight().
    #  Se fija como constante y se centra la ventana.

    root.update_idletasks()
    WIN_H = root.winfo_reqheight()
    center_window(WIN_W, WIN_H)

    root.mainloop()


# ─────────────────────────────────────────────────────────────────
#  7. PUNTO DE ENTRADA
# ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    run_gui()
