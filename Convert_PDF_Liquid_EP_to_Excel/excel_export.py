"""
excel_export.py — Vuelca una lista de pdf_parser.Registro a un fichero
.xlsx con dos hojas:

  - "Resumen": una fila por alumno/expediente, con los importes agregados
    y las fechas de cobro concatenadas (para cuadrar totales rápidamente).
  - "Detalle": una fila por línea de referencia de cobro (mismo Exped/DNI
    repetido si el alumno tiene varios pagos), más una fila adicional por
    alumno con el importe administrativo (columna "Imp.Adm."; el resto de
    columnas de esa fila — Referencia, Fecha Cobro, Importe, Plazo, Forma
    Pago — quedan vacías, ya que el cargo administrativo no es una línea
    de cobro). Así cada importe (cobros + administrativo) tiene su propia
    fila para conciliar uno a uno.
"""

# Última actualización: 2026-07-08 13:11

from __future__ import annotations

import re

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from pdf_parser import Registro

_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_FILL = "1A5276"
_MONEY_FMT = "#,##0.00 €"
_INT_FMT = "#,##0"

_RESUMEN_HEADERS = (
    "Exped", "DNI", "Apellidos y Nombre", "Nº Refs.",
    "Importe Cobrado", "Administrativo", "Importe Neto",
    "Fechas de Cobro",
)
_DETALLE_HEADERS = (
    "Exped", "DNI", "Apellidos y Nombre", "Referencia",
    "Fecha Cobro", "Importe", "Plazo", "Forma Pago", "Imp.Adm.",
)


def _write_header(ws: Worksheet, headers: tuple[str, ...]) -> None:
    fill = PatternFill(start_color=_HEADER_FILL,
                        end_color=_HEADER_FILL, fill_type="solid")
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=text)
        cell.font = _HEADER_FONT
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def _autosize(ws: Worksheet, headers: tuple[str, ...]) -> None:
    widths = [len(h) for h in headers]
    for row in ws.iter_rows(min_row=2, values_only=True):
        for i, value in enumerate(row):
            w = len(str(value)) if value is not None else 0
            if w > widths[i]:
                widths[i] = w
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 60)


def _write_resumen(ws: Worksheet, registros: list[Registro]) -> None:
    _write_header(ws, _RESUMEN_HEADERS)
    money_cols = (5, 6, 7)
    for row_i, r in enumerate(registros, start=2):
        fechas = "; ".join(ref.fecha for ref in r.referencias)
        ws.cell(row=row_i, column=1, value=r.exped)
        ws.cell(row=row_i, column=2, value=r.dni)
        ws.cell(row=row_i, column=3, value=r.nombre)
        ws.cell(row=row_i, column=4, value=len(r.referencias))
        ws.cell(row=row_i, column=5, value=r.importe)
        ws.cell(row=row_i, column=6, value=r.administrativo)
        ws.cell(row=row_i, column=7, value=r.importe_neto)
        ws.cell(row=row_i, column=8, value=fechas)
        for c in money_cols:
            ws.cell(row=row_i, column=c).number_format = _MONEY_FMT
    _autosize(ws, _RESUMEN_HEADERS)


def _write_detalle(ws: Worksheet, registros: list[Registro]) -> None:
    _write_header(ws, _DETALLE_HEADERS)
    row_i = 2
    for r in registros:
        for ref in r.referencias:
            ws.cell(row=row_i, column=1, value=r.exped)
            ws.cell(row=row_i, column=2, value=r.dni)
            ws.cell(row=row_i, column=3, value=r.nombre)
            ws.cell(row=row_i, column=4, value=ref.referencia)
            ws.cell(row=row_i, column=5, value=ref.fecha)
            ws.cell(row=row_i, column=6, value=ref.importe)
            ws.cell(row=row_i, column=7, value=ref.plazo)
            ws.cell(row=row_i, column=8, value=ref.forma_pago)
            ws.cell(row=row_i, column=6).number_format = _MONEY_FMT
            row_i += 1
        # Fila adicional con el importe administrativo del alumno: solo
        # los datos identificativos + Imp.Adm.; el resto de columnas de
        # cobro quedan vacías porque no es una línea de cobro.
        ws.cell(row=row_i, column=1, value=r.exped)
        ws.cell(row=row_i, column=2, value=r.dni)
        ws.cell(row=row_i, column=3, value=r.nombre)
        ws.cell(row=row_i, column=9, value=r.administrativo)
        ws.cell(row=row_i, column=9).number_format = _MONEY_FMT
        row_i += 1
    _autosize(ws, _DETALLE_HEADERS)


def export_to_excel(registros: list[Registro], output_path: str) -> None:
    """Genera el .xlsx con las hojas Resumen y Detalle a partir de los
    registros ya parseados. Sobrescribe el fichero si ya existe."""
    wb = Workbook()
    ws_resumen = wb.active
    ws_resumen.title = "Resumen"
    _write_resumen(ws_resumen, registros)

    ws_detalle = wb.create_sheet("Detalle")
    _write_detalle(ws_detalle, registros)

    wb.save(output_path)


def _to_number(value):
    """Convierte a número real un valor de celda que puede venir como texto
    ya formateado para pantalla (p. ej. "+67.40", "-0.60", "238.39", "") o ya
    numérico (int/float). Vacío -> None (celda en blanco: no cuenta en el
    SUM/COUNT de la fila de totales). float() admite el "+" inicial tal
    cual, no hace falta despojarlo."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip()
    return float(s) if s else None


def _escape_structref(name: str) -> str:
    """Escapa un nombre de columna para usarlo dentro de una referencia
    estructurada de Excel Table, p. ej. Tabla[Nombre]: duplica los
    caracteres especiales (comilla simple, almohadilla) según OOXML."""
    return name.replace("'", "''").replace("#", "##")


def _sanitize_table_name(name: str, used: set[str]) -> str:
    """Nombre válido y único (dentro del libro) para una Excel Table: solo
    letras/dígitos/guion bajo, no empieza por dígito, sin espacios."""
    base = re.sub(r"\W+", "_", name).strip("_") or "Tabla"
    if base[0].isdigit():
        base = f"T_{base}"
    candidate = base
    n = 1
    while candidate in used:
        n += 1
        candidate = f"{base}_{n}"
    used.add(candidate)
    return candidate


def _write_table(ws: Worksheet, table_name: str, headers: tuple[str, ...],
                 rows: list[tuple], *, money_cols: tuple[int, ...] = (),
                 int_cols: tuple[int, ...] = (),
                 count_col: int | None = None) -> None:
    """Vuelca headers+rows como una Excel Table (ListObject) real, no como
    celdas sueltas: las columnas de `money_cols`/`int_cols` (índices
    0-based) se escriben como número real (no texto) con su formato (€ o
    entero), imprescindible para que la fila de totales pueda sumarlas. Esa
    fila de totales lleva un recuento (`COUNTA`, vía fórmula SUBTOTAL) en
    `count_col` y una suma (`SUM`, vía SUBTOTAL) en cada columna de
    `money_cols`; el resto de columnas quedan sin función (celda vacía).
    Los headers deben ser todos no vacíos y únicos (lo exige una Excel
    Table): quien llame debe resolver antes cualquier cabecera en blanco."""
    _write_header(ws, headers)
    numeric_cols = set(money_cols) | set(int_cols)
    for row_i, row in enumerate(rows, start=2):
        for col_i, value in enumerate(row, start=1):
            idx = col_i - 1
            if idx in numeric_cols:
                value = _to_number(value)
            cell = ws.cell(row=row_i, column=col_i, value=value)
            if idx in money_cols:
                cell.number_format = _MONEY_FMT
            elif idx in int_cols:
                cell.number_format = _INT_FMT
    _autosize(ws, headers)

    n_cols = len(headers)
    totals_row = 1 + len(rows) + 1
    for col_i in range(1, n_cols + 1):
        idx = col_i - 1
        if idx == count_col:
            code = 103   # SUBTOTAL 103 = COUNTA (cuenta texto y números)
        elif idx in money_cols:
            code = 109   # SUBTOTAL 109 = SUM
        else:
            continue
        ref = _escape_structref(headers[idx])
        ws.cell(row=totals_row, column=col_i,
                value=f"=SUBTOTAL({code},{table_name}[{ref}])")

    tab = Table(displayName=table_name,
               ref=f"A1:{get_column_letter(n_cols)}{totals_row}")
    tab.tableColumns = [TableColumn(id=i + 1, name=h)
                        for i, h in enumerate(headers)]
    if count_col is not None:
        tab.tableColumns[count_col].totalsRowFunction = "count"
    for idx in money_cols:
        tab.tableColumns[idx].totalsRowFunction = "sum"
    tab.totalsRowShown = True
    tab.totalsRowCount = 1
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2",
                                        showRowStripes=True)
    ws.add_table(tab)


def export_rows_to_excel(headers: tuple[str, ...], rows: list[tuple],
                         output_path: str, *, sheet_name: str = "Datos",
                         table_name: str = "Tabla",
                         money_cols: tuple[int, ...] = (),
                         int_cols: tuple[int, ...] = (),
                         count_col: int | None = None) -> None:
    """Genera un .xlsx de una sola hoja a partir de filas ya formateadas
    para pantalla, volcadas como una Excel Table real (ver `_write_table`).
    Usado para exportar tablas de la UI que no vienen de Registro, p. ej. el
    diff de Comparar Resumen. Sobrescribe el fichero si ya existe."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    _write_table(ws, _sanitize_table_name(table_name, set()), headers, rows,
                money_cols=money_cols, int_cols=int_cols, count_col=count_col)
    wb.save(output_path)


def export_multi_sheet_to_excel(
        sheets: list[tuple[str, tuple[str, ...], list[tuple],
                          tuple[int, ...], tuple[int, ...], int | None]],
        output_path: str) -> None:
    """Genera un .xlsx con una hoja por cada tupla `sheets` de
    (nombre, headers, rows, money_cols, int_cols, count_col) — mismo
    significado de cada campo que en `export_rows_to_excel`, pero por hoja:
    cada una puede tener su propio layout de columnas (caso de uso real:
    las tablas A/B de Comparar Detalle y la hoja unificada "Detalle A+B",
    con distinto nº y orden de columnas). Cada hoja se escribe como su
    propia Excel Table (ver `_write_table`). Sobrescribe el fichero si ya
    existe."""
    wb = Workbook()
    used_names: set[str] = set()
    for i, (sheet_name, headers, rows, money_cols, int_cols, count_col) \
            in enumerate(sheets):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = sheet_name
        table_name = _sanitize_table_name(sheet_name, used_names)
        _write_table(ws, table_name, headers, rows, money_cols=money_cols,
                    int_cols=int_cols, count_col=count_col)
    wb.save(output_path)
