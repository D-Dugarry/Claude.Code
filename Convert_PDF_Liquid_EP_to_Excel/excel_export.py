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

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from pdf_parser import Registro

_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_FILL = "1A5276"
_MONEY_FMT = "#,##0.00 €"

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
