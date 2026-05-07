# =============================================================================
# n43_to_excel.py  —  Conversor Norma AEB 43 → Excel (.xlsx)
# =============================================================================
# Descripción:
#   Lee un fichero .n43 (Norma AEB 43), extrae los registros de movimientos
#   bancarios y genera un Excel formateado con cabecera, filas alternadas,
#   fila de totales con fórmulas, panel congelado, autofiltro e imagen de
#   marca. Incluye una interfaz gráfica (Tkinter) para seleccionar el fichero
#   de entrada; el Excel de salida se guarda en la misma carpeta.
#
# Dependencias externas:  openpyxl  (pip install openpyxl)
#
# Índice de bloques principales:
#   Línea  31  — Imports y dependencias
#   Línea  42  — resource_path()   : Resolución de rutas (script / .exe)
#   Línea  51  — parse_n43()       : Parseo N43 y generación del Excel
#     Línea  69  —   Libro y estilos de cabecera de columnas
#     Línea  93  —   Parseo del fichero N43 (registros 11, 22, 23)
#     Línea 134  —   Estilos y formato de filas de datos
#     Línea 168  —   Anchos de columna
#     Línea 173  —   Fila de totales, panel fijo y autofiltro
#     Línea 198  —   Inserción de imagen de marca (Caracolillo_Fósil.png)
#   Línea 212  — run_gui()         : Interfaz gráfica Tkinter
#     Línea 220  —   Ventana principal y título
#     Línea 243  —   seleccionar() : diálogo de apertura de fichero
#     Línea 254  —   convertir()   : llama a parse_n43 y gestiona el resultado
#     Línea 277  —   Widgets: botones, barra de progreso, etiquetas
#   Línea 303  — Punto de entrada (__main__)
# =============================================================================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XlImage
from openpyxl.utils import get_column_letter
from datetime import date
import os
import sys


def resource_path(filename):
    """Resuelve la ruta de recursos tanto en script como en .exe (PyInstaller)."""
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def parse_n43(path_n43: str, path_out: str):
    """
    Parsea un fichero N43 y genera un Excel con los movimientos.

    Parámetros:
        path_n43 : ruta al fichero de entrada (.n43 / .N43)
        path_out : ruta del Excel de salida (.xlsx)

    Devuelve:
        int — número de movimientos (registros tipo 22) procesados.

    Estructura de registros Norma 43:
        Reg. 11 — Cabecera de cuenta: banco, oficina, cuenta, saldo inicial
        Reg. 22 — Movimiento: fecha operación/valor, importe, referencia
        Reg. 23 — Conceptos complementarios del último movimiento (hasta 5)
        Reg. 33 — Fin de cuenta  (no procesado)
        Reg. 88 — Fin de fichero (no procesado)
    """
    # ---- Libro y estilos de cabecera ----------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracto"

    headers = [
        "Orden", "Oficina", "F.Operación", "F.Valor", "Importe",
        "Saldo", "Documento", "Referencia", "RCM01", "RCM02", "RCM03", "RCM04", "RCM05"
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

    # ---- Parseo del fichero N43 (registros 11, 22, 23) ----------------------
    rows, saldo_acum, num_orden = [], 0.0, 0
    current_row = None
    anualidad   = ""

    with open(path_n43, encoding="latin-1") as f:
        for linea in f:
            reg = linea[:2]
            if reg == "11":
                # Saldo inicial: posición 32 = signo (1=debe), 33-46 = importe en céntimos
                signo      = -1 if linea[32] == "1" else 1
                saldo_acum = int(linea[33:47]) * signo * 0.01
            elif reg == "22":
                # Movimiento: año yy en posición 10-11 (F.Operación) y 16-17 (F.Valor)
                año = "20" + linea[10:12]
                if linea[16:18] != anualidad:   # reinicia contador al cambiar de año
                    num_orden = 0
                    anualidad = linea[16:18]
                num_orden += 1
                # Importe: posición 27 = signo (1=debe/cargo), 28-41 = importe en céntimos
                signo   = -1 if linea[27] == "1" else 1
                importe = int(linea[28:42]) * signo * 0.01
                saldo_acum += importe
                current_row = [
                    f"{año}{num_orden:06d}",
                    linea[6:10].strip(),          # oficina
                    date(int("20"+linea[10:12]), int(linea[12:14]), int(linea[14:16])),  # F.Operación
                    date(int("20"+linea[16:18]), int(linea[18:20]), int(linea[20:22])),  # F.Valor
                    importe,
                    round(saldo_acum, 2),
                    linea[42:52].strip(),         # documento
                    linea[52:102].strip(),        # referencia (concepto principal)
                    "", "", "", "", ""            # RCM01..05 — se rellenan con reg. 23
                ]
                rows.append(current_row)
            elif reg == "23" and current_row:
                # Concepto complementario: posición 2-3 = índice (01-05)
                idx = int(linea[2:4])
                if 1 <= idx <= 5:
                    current_row[7 + idx] = linea[4:79].strip()

    # ---- Estilos y formato de filas de datos --------------------------------
    fill_even    = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
    fill_odd     = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    font_data    = Font(name="Arial", size=9)
    align_center = Alignment(horizontal="center", vertical="center")
    align_right  = Alignment(horizontal="right",  vertical="center")
    align_left   = Alignment(horizontal="left",   vertical="center")

    date_fmt  = "DD/MM/YYYY"
    money_fmt = '#,##0.00 €;[RED]-#,##0.00 €'   # importes negativos en rojo
    money_pos = '#,##0.00 €'

    for r_idx, row in enumerate(rows, 2):
        fill = fill_even if r_idx % 2 == 0 else fill_odd
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font   = font_data
            cell.fill   = fill
            cell.border = cell_border

            if c_idx in (3, 4):      # Fechas: F.Operación y F.Valor
                cell.number_format = date_fmt
                cell.alignment     = align_center
            elif c_idx == 5:         # Importe (puede ser negativo)
                cell.number_format = money_fmt
                cell.alignment     = align_right
            elif c_idx == 6:         # Saldo acumulado (siempre positivo en extractos normales)
                cell.number_format = money_pos
                cell.alignment     = align_right
            elif c_idx in (1, 2, 7):
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # ---- Anchos de columna --------------------------------------------------
    col_widths = [14, 8, 12, 12, 14, 14, 12, 30, 25, 25, 25, 25, 25]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- Fila de totales, panel fijo y autofiltro ---------------------------
    last_data_row = len(rows) + 1
    total_row     = last_data_row + 1
    ws.row_dimensions[total_row].height = 20

    total_fill = PatternFill("solid", start_color="2E75B6", end_color="2E75B6")
    total_font = Font(name="Arial", bold=True, color="FFFFFF", size=9)

    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=total_row, column=c)
        cell.fill   = total_fill
        cell.font   = total_font
        cell.border = cell_border

    ws.cell(row=total_row, column=1, value="TOTAL").alignment = Alignment(horizontal="center")
    ws.cell(row=total_row, column=5,
            value=f"=SUM(E2:E{last_data_row})").number_format = money_fmt
    ws.cell(row=total_row, column=5).alignment = align_right
    # Saldo final = suma de importes + saldo antes del primer movimiento
    ws.cell(row=total_row, column=6,
            value=f"=E{total_row}+{rows[0][5] - rows[0][4] if rows else 0:.2f}").number_format = money_pos

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_data_row}"

    # ---- Inserción de imagen de marca (Caracolillo_Fósil.png) ---------------
    img_path = resource_path("Caracolillo_Fósil.png")
    if os.path.isfile(img_path):
        img = XlImage(img_path)
        img.width  = 57   # ~1.5 cm
        img.height = 57
        anchor_col = get_column_letter(len(headers))  # columna M (última)
        anchor_row = total_row + 2
        ws.add_image(img, f"{anchor_col}{anchor_row}")

    wb.save(path_out)
    return len(rows)


def run_gui():
    """Lanza la interfaz gráfica Tkinter para seleccionar y convertir el N43."""
    root = tk.Tk()
    root.title("Conversor N43 → Excel")
    root.geometry("480x320")
    root.resizable(False, False)
    root.configure(bg="#F0F4F8")

    # ---- Ventana principal y título -----------------------------------------
    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    tk.Label(root, text="Conversor Norma 43 → Excel",
             font=("Arial", 14, "bold"), bg="#1F4E79", fg="white",
             pady=12).pack(fill="x")

    tk.Label(root, text="Selecciona un fichero .n43 para convertirlo a .xlsx",
             font=("Arial", 10), bg="#F0F4F8", fg="#333333",
             pady=8).pack()

    path_var   = tk.StringVar(value="Ningún fichero seleccionado")
    frame_path = tk.Frame(root, bg="#E2EAF4", bd=1, relief="sunken")
    frame_path.pack(fill="x", padx=20, pady=(0, 10))
    tk.Label(frame_path, textvariable=path_var, font=("Arial", 9),
             bg="#E2EAF4", fg="#555555", anchor="w",
             wraplength=420, justify="left", padx=8, pady=6).pack(fill="x")

    selected_path = {"value": None}

    # ---- seleccionar(): diálogo de apertura de fichero ----------------------
    def seleccionar():
        p = filedialog.askopenfilename(
            title="Selecciona el fichero N43",
            filetypes=[("Ficheros N43", "*.n43 *.N43"), ("Todos los ficheros", "*.*")]
        )
        if p:
            selected_path["value"] = p
            path_var.set(os.path.basename(p))
            btn_convertir.config(state="normal")

    # ---- convertir(): llama a parse_n43 y gestiona el resultado -------------
    def convertir():
        if not selected_path["value"]:
            return
        p_in  = selected_path["value"]
        p_out = os.path.splitext(p_in)[0] + ".xlsx"

        progress.start(10)
        root.update()

        try:
            n = parse_n43(p_in, p_out)
            progress.stop()
            progress["value"] = 100
            messagebox.showinfo(
                "✅ Conversión completada",
                f"Se han importado {n} movimientos.\n\nFichero guardado en:\n{p_out}"
            )
            os.startfile(os.path.dirname(p_out))   # abre la carpeta de destino
        except Exception as e:
            progress.stop()
            messagebox.showerror("Error en la conversión", str(e))

    # ---- Widgets: botones, barra de progreso, etiquetas ---------------------
    frame_btns = tk.Frame(root, bg="#F0F4F8")
    frame_btns.pack(pady=10)

    tk.Button(frame_btns, text="📂  Seleccionar fichero .n43",
              font=("Arial", 10, "bold"), bg="#2E75B6", fg="white",
              activebackground="#1F4E79", activeforeground="white",
              bd=0, padx=16, pady=8, cursor="hand2",
              command=seleccionar).pack(side="left", padx=8)

    btn_convertir = tk.Button(frame_btns, text="⚡  Convertir a Excel",
                              font=("Arial", 10, "bold"), bg="#70AD47", fg="white",
                              activebackground="#507E35", activeforeground="white",
                              bd=0, padx=16, pady=8, cursor="hand2",
                              state="disabled", command=convertir)
    btn_convertir.pack(side="left", padx=8)

    progress = ttk.Progressbar(root, orient="horizontal", length=440, mode="indeterminate")
    progress.pack(pady=(0, 8))

    tk.Label(root, text="El Excel se guarda en la misma carpeta que el fichero .n43",
             font=("Arial", 8), bg="#F0F4F8", fg="#888888").pack()

    root.mainloop()


if __name__ == "__main__":
    run_gui()
