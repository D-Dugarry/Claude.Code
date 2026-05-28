"""
VBA Comparator — compara módulos VBA de dos archivos .xlsm
Salida: consola (rich) + Excel (openpyxl)
Extracción: win32com (requiere Excel instalado)
"""

import sys
import os
import tempfile
import difflib
import re
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

import io
# Forzar UTF-8 en stdout para evitar errores de codificación en Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

try:
    import win32com.client
except ImportError:
    sys.exit("ERROR: pywin32 no instalado. Ejecuta: pip install pywin32")

try:
    from rich.console import Console
    from rich.table import Table
    from rich.text import Text
    from rich import box
except ImportError:
    sys.exit("ERROR: rich no instalado. Ejecuta: pip install rich")

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("ERROR: openpyxl no instalado. Ejecuta: pip install openpyxl")


# ── Colores ──────────────────────────────────────────────────────────────────

FILL_ADD       = PatternFill("solid", fgColor="C6EFCE")  # verde claro
FILL_DEL       = PatternFill("solid", fgColor="FFC7CE")  # rojo claro
FILL_CHANGE    = PatternFill("solid", fgColor="FFEB9C")  # amarillo
FILL_HEADER    = PatternFill("solid", fgColor="2F75B6")  # azul oscuro
FILL_MODULE    = PatternFill("solid", fgColor="D6E4F0")  # azul claro
FILL_RESULTADO = PatternFill("solid", fgColor="EAF2FB")  # azul muy claro (columna editable)
FILL_HDR_RES   = PatternFill("solid", fgColor="D35400")  # naranja oscuro (cabecera Resultado)
FILL_EQUAL     = None

TAB_GREEN = "70AD47"   # pestaña módulo sin cambios
TAB_RED   = "FF0000"   # pestaña módulo con cambios
TAB_BLUE  = "4472C4"   # pestaña índice


# ── Estructuras ───────────────────────────────────────────────────────────────

@dataclass
class VBAModule:
    name: str
    code: str
    kind: str  # "Module", "ClassModule", "Document", "Form"


@dataclass
class ModuleDiff:
    name: str
    kind: str
    status: str          # "equal" | "modified" | "only_a" | "only_b"
    lines_a: list[str] = field(default_factory=list)
    lines_b: list[str] = field(default_factory=list)
    opcodes: list = field(default_factory=list)  # difflib SequenceMatcher opcodes


# ── Extracción VBA via win32com ───────────────────────────────────────────────

COMPONENT_TYPE = {1: "Module", 2: "ClassModule", 3: "Form", 100: "Document"}


def extract_vba(filepath: str) -> dict[str, VBAModule]:
    """Abre el .xlsm con Excel y extrae todos los módulos VBA."""
    abs_path = str(Path(filepath).resolve())
    if not os.path.exists(abs_path):
        sys.exit(f"ERROR: no se encuentra el archivo '{abs_path}'")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    modules: dict[str, VBAModule] = {}
    tmp_dir = tempfile.mkdtemp()

    try:
        wb = excel.Workbooks.Open(abs_path, ReadOnly=True)
        try:
            vbp = wb.VBProject
        except Exception:
            sys.exit(
                f"ERROR: no se puede acceder al VBProject de '{filepath}'.\n"
                "Asegúrate de que 'Confiar en el acceso al modelo de objetos VBA' "
                "está habilitado en Opciones de Excel > Centro de confianza."
            )

        for comp in vbp.VBComponents:
            name = comp.Name
            kind = COMPONENT_TYPE.get(comp.Type, "Unknown")
            lines = comp.CodeModule.CountOfLines
            code = comp.CodeModule.Lines(1, lines) if lines > 0 else ""
            modules[name] = VBAModule(name=name, code=code, kind=kind)

        try:
            wb.Close(SaveChanges=False)
        except Exception:
            # Algunos ficheros tienen un macro Workbook_BeforeClose que falla en
            # modo COM (p.ej. intentan interactuar con el usuario). El VBA ya se
            # extrajo, así que ignoramos el error y dejamos que excel.Quit() limpie.
            pass
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

    return modules


# ── Comparación ───────────────────────────────────────────────────────────────

def _normalize(code: str) -> list[str]:
    """
    Normaliza el código antes de comparar:
    - Sin espacios/tabs al final de cada línea
    - Sin líneas en blanco al inicio ni al final del módulo
    - Múltiples líneas en blanco consecutivas colapsadas a una sola
    """
    lines = [line.rstrip() for line in code.splitlines()]

    # Colapsar múltiples blancos consecutivos en uno
    result: list[str] = []
    prev_blank = False
    for line in lines:
        if line == "":
            if not prev_blank:
                result.append(line)
            prev_blank = True
        else:
            result.append(line)
            prev_blank = False

    # Eliminar blancos al inicio y al final del módulo
    while result and result[0] == "":
        result.pop(0)
    while result and result[-1] == "":
        result.pop()

    return result


def compare(
    modules_a: dict[str, VBAModule],
    modules_b: dict[str, VBAModule],
    modulo_filtro: str | None = None,
) -> list[ModuleDiff]:
    all_names = sorted(set(modules_a) | set(modules_b))
    diffs: list[ModuleDiff] = []

    for name in all_names:
        mod_a = modules_a.get(name)
        mod_b = modules_b.get(name)
        kind = (mod_a or mod_b).kind

        # Si hay filtro y este módulo no es el seleccionado → marcar como "skip"
        if modulo_filtro and name != modulo_filtro:
            diffs.append(ModuleDiff(name=name, kind=kind, status="skip"))
            continue

        if mod_a is None:
            diffs.append(ModuleDiff(name=name, kind=kind, status="only_b",
                                    lines_b=_normalize(mod_b.code)))
        elif mod_b is None:
            diffs.append(ModuleDiff(name=name, kind=kind, status="only_a",
                                    lines_a=_normalize(mod_a.code)))
        else:
            la = _normalize(mod_a.code)
            lb = _normalize(mod_b.code)
            sm = difflib.SequenceMatcher(None, la, lb, autojunk=False)
            opcodes = sm.get_opcodes()
            status = "equal" if all(tag == "equal" for tag, *_ in opcodes) else "modified"
            diffs.append(ModuleDiff(name=name, kind=kind, status=status,
                                    lines_a=la, lines_b=lb, opcodes=opcodes))

    return diffs


# ── Salida consola (rich) ─────────────────────────────────────────────────────

STATUS_ICON = {
    "equal":    ("[=]", "dim"),
    "modified": ("[~]", "yellow"),
    "only_a":   ("[-]", "red"),
    "only_b":   ("[+]", "green"),
}

STATUS_LABEL = {
    "equal":    "sin cambios",
    "modified": "modificado",
    "only_a":   "solo en A",
    "only_b":   "solo en B",
}


def render_console(diffs: list[ModuleDiff], label_a: str, label_b: str,
                   show_equal: bool = False) -> None:
    console = Console(highlight=False)

    # ── resumen ──
    console.rule("[bold blue]RESUMEN DE MÓDULOS")
    summary = Table(box=box.SIMPLE_HEAD, show_header=True, header_style="bold")
    summary.add_column("Módulo", style="bold")
    summary.add_column("Tipo")
    summary.add_column("Estado")

    for d in diffs:
        icon, style = STATUS_ICON[d.status]
        summary.add_row(d.name, d.kind, Text(f"{icon} {STATUS_LABEL[d.status]}", style=style))

    console.print(summary)

    totals = {s: sum(1 for d in diffs if d.status == s)
              for s in ("equal", "modified", "only_a", "only_b")}
    console.print(
        f"  [dim]Sin cambios:[/] {totals['equal']}  "
        f"[yellow]Modificados:[/] {totals['modified']}  "
        f"[red]Solo en A:[/] {totals['only_a']}  "
        f"[green]Solo en B:[/] {totals['only_b']}"
    )
    console.print()

    # ── detalle por módulo ──
    for d in diffs:
        if d.status == "equal" and not show_equal:
            continue

        icon, style = STATUS_ICON[d.status]
        console.rule(f"[{style}]{icon} {d.name}[/]  [dim]({d.kind})[/]")

        if d.status == "only_a":
            _print_only(console, d.lines_a, label_a, "red")
        elif d.status == "only_b":
            _print_only(console, d.lines_b, label_b, "green")
        elif d.status == "modified":
            _print_diff(console, d, label_a, label_b)
        else:
            console.print(f"  [dim]Idéntico ({len(d.lines_a)} líneas)[/]")
        console.print()


def _print_only(console: Console, lines: list[str], label: str, color: str) -> None:
    prefix = "+" if color == "green" else "-"
    console.print(f"  [bold {color}]{prefix} {label}[/]")
    for i, line in enumerate(lines, 1):
        console.print(f"  [{color}]{i:>4} {prefix} {line}[/]")


def _print_diff(console: Console, d: ModuleDiff, label_a: str, label_b: str) -> None:
    table = Table(box=box.SIMPLE, show_header=True, padding=(0, 1),
                  header_style="bold", expand=True)
    table.add_column(f"#{label_a}", style="dim", width=5, no_wrap=True)
    table.add_column(label_a, ratio=1)
    table.add_column(f"#{label_b}", style="dim", width=5, no_wrap=True)
    table.add_column(label_b, ratio=1)

    for tag, i1, i2, j1, j2 in d.opcodes:
        if tag == "equal":
            for ia, ib in zip(range(i1, i2), range(j1, j2)):
                table.add_row(str(ia + 1), Text(d.lines_a[ia]),
                              str(ib + 1), Text(d.lines_b[ib]))
        elif tag == "replace":
            la_block = d.lines_a[i1:i2]
            lb_block = d.lines_b[j1:j2]
            for k in range(max(len(la_block), len(lb_block))):
                ca = (str(i1 + k + 1), Text(la_block[k], style="yellow")) if k < len(la_block) else ("", Text(""))
                cb = (str(j1 + k + 1), Text(lb_block[k], style="yellow")) if k < len(lb_block) else ("", Text(""))
                table.add_row(ca[0], ca[1], cb[0], cb[1])
        elif tag == "delete":
            for ia in range(i1, i2):
                table.add_row(str(ia + 1), Text(d.lines_a[ia], style="red"), "", Text(""))
        elif tag == "insert":
            for ib in range(j1, j2):
                table.add_row("", Text(""), str(ib + 1), Text(d.lines_b[ib], style="green"))

    console.print(table)


# ── Salida Excel (openpyxl) ───────────────────────────────────────────────────

def render_excel(diffs: list[ModuleDiff], label_a: str, label_b: str,
                 output_path: str, incluir_iguales: bool = False) -> None:
    wb = openpyxl.Workbook()

    # mapa nombre original → nombre de hoja seguro
    # "skip" nunca tiene hoja (ni con incluir_iguales)
    # "equal" solo tiene hoja si incluir_iguales=True
    sheet_name_map: dict[str, str] = {}
    for d in diffs:
        if d.status == "skip":
            continue
        if d.status != "equal" or incluir_iguales:
            safe = re.sub(r"[\\/*?:\[\]]", "_", d.name)[:31]
            sheet_name_map[d.name] = safe

    # hoja índice (siempre con los 102 módulos)
    ws_sum = wb.active
    ws_sum.title = "Indice"
    ws_sum.sheet_properties.tabColor = TAB_BLUE
    _build_summary_sheet(ws_sum, diffs, label_a, label_b, sheet_name_map)

    # hojas solo para módulos con diferencias (o todos si --incluir-iguales)
    for d in diffs:
        if d.name not in sheet_name_map:
            continue
        safe_name = sheet_name_map[d.name]
        ws = wb.create_sheet(title=safe_name)
        ws.sheet_properties.tabColor = TAB_GREEN if d.status == "equal" else TAB_RED
        _build_module_sheet(ws, d, label_a, label_b)

    wb.save(output_path)


def _header_font():
    return Font(bold=True, color="FFFFFF")


def _thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def _build_summary_sheet(ws, diffs: list[ModuleDiff], label_a: str, label_b: str,
                          sheet_name_map: dict[str, str]):
    headers = ["Módulo", "Tipo", "Estado", "Líneas A", "Líneas B", "Cambios"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = FILL_HEADER
        cell.font = _header_font()
        cell.alignment = Alignment(horizontal="center")

    FILL_SKIP  = PatternFill("solid", fgColor="D3D3D3")  # gris — Sin Comparar
    STATUS_ES = {"equal": "Sin cambios", "modified": "Modificado",
                 "only_a": "Solo en A", "only_b": "Solo en B",
                 "skip":  "Sin Comparar"}
    STATUS_FILL = {"equal": None, "modified": FILL_CHANGE,
                   "only_a": FILL_DEL, "only_b": FILL_ADD,
                   "skip":  FILL_SKIP}

    for d in diffs:
        changes = sum(1 for tag, *_ in d.opcodes if tag != "equal") if d.opcodes else (
            len(d.lines_a) if d.status == "only_a" else len(d.lines_b)
        )
        row = [d.name, d.kind, STATUS_ES[d.status],
               len(d.lines_a), len(d.lines_b), changes]
        ws.append(row)
        row_idx = ws.max_row
        fill = STATUS_FILL[d.status]
        if fill:
            for cell in ws[row_idx]:
                cell.fill = fill

        # hipervínculo solo si el módulo tiene hoja propia (no los idénticos)
        name_cell = ws.cell(row_idx, 1)
        if d.name in sheet_name_map:
            safe = sheet_name_map[d.name]
            name_cell.hyperlink = f"#'{safe}'!A1"
            name_cell.font = Font(color="0563C1", underline="single", bold=True)
        else:
            name_cell.font = Font(color="000000")

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    for col in ("D", "E", "F"):
        ws.column_dimensions[col].width = 10

    ws.freeze_panes = "A2"


def _build_module_sheet(ws, d: ModuleDiff, label_a: str, label_b: str):
    # fila 1: botón volver al índice (A1:D1) + nombre original en E1 + estado en F1
    ws.append(["<< Indice", "", "", "", d.name, d.status, d.kind])
    ws.merge_cells("A1:D1")
    back_cell = ws.cell(1, 1)
    back_cell.hyperlink = "#'Indice'!A1"
    back_cell.font = Font(color="0563C1", underline="single", bold=True)
    back_cell.fill = FILL_MODULE
    back_cell.alignment = Alignment(horizontal="left")
    name_ref = ws.cell(1, 5)
    name_ref.font = Font(size=8, color="888888", italic=True)
    name_ref.alignment = Alignment(horizontal="right")
    ws.cell(1, 6).font = Font(size=8, color="888888", italic=True)

    # fila 2: cabecera de columnas (5 columnas: #A, CódigoA, #B, CódigoB, Resultado)
    ws.append([f"# {label_a}", label_a, f"# {label_b}", label_b, "Resultado"])
    for i, cell in enumerate(ws[2], 1):
        cell.fill = FILL_HDR_RES if i == 5 else FILL_HEADER
        cell.font = _header_font()
        cell.alignment = Alignment(horizontal="center")

    # datos — col E (Resultado) pre-rellena con la versión "ganadora" por defecto:
    #   equal   → valor común
    #   replace → valor de B (versión más nueva)
    #   delete  → valor de A (se mantiene por defecto; borrar celda para eliminar)
    #   insert  → valor de B (se incluye por defecto)
    if d.status == "only_a":
        for i, line in enumerate(d.lines_a, 1):
            ws.append([i, line, "", "", line])
            ws.cell(ws.max_row, 2).fill = FILL_DEL
            ws.cell(ws.max_row, 5).fill = FILL_RESULTADO
    elif d.status == "only_b":
        for i, line in enumerate(d.lines_b, 1):
            ws.append(["", "", i, line, line])
            ws.cell(ws.max_row, 4).fill = FILL_ADD
            ws.cell(ws.max_row, 5).fill = FILL_RESULTADO
    else:
        for tag, i1, i2, j1, j2 in d.opcodes:
            if tag == "equal":
                for ia, ib in zip(range(i1, i2), range(j1, j2)):
                    ws.append([ia + 1, d.lines_a[ia], ib + 1, d.lines_b[ib], d.lines_a[ia]])
            elif tag == "replace":
                la_b, lb_b = d.lines_a[i1:i2], d.lines_b[j1:j2]
                for k in range(max(len(la_b), len(lb_b))):
                    resultado = lb_b[k] if k < len(lb_b) else la_b[k]
                    r = [i1 + k + 1 if k < len(la_b) else "",
                         la_b[k] if k < len(la_b) else "",
                         j1 + k + 1 if k < len(lb_b) else "",
                         lb_b[k] if k < len(lb_b) else "",
                         resultado]
                    ws.append(r)
                    ws.cell(ws.max_row, 2).fill = FILL_CHANGE
                    ws.cell(ws.max_row, 4).fill = FILL_CHANGE
                    ws.cell(ws.max_row, 5).fill = FILL_RESULTADO
            elif tag == "delete":
                for ia in range(i1, i2):
                    ws.append([ia + 1, d.lines_a[ia], "", "", d.lines_a[ia]])
                    ws.cell(ws.max_row, 2).fill = FILL_DEL
                    ws.cell(ws.max_row, 5).fill = FILL_RESULTADO
            elif tag == "insert":
                for ib in range(j1, j2):
                    ws.append(["", "", ib + 1, d.lines_b[ib], d.lines_b[ib]])
                    ws.cell(ws.max_row, 4).fill = FILL_ADD
                    ws.cell(ws.max_row, 5).fill = FILL_RESULTADO

    # anchos
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 6
    ws.column_dimensions["D"].width = 55
    ws.column_dimensions["E"].width = 60
    ws.freeze_panes = "A3"

    # fuente monoespaciada en columnas de código (datos desde fila 3)
    for col in (2, 4, 5):
        for row in ws.iter_rows(min_row=3, min_col=col, max_col=col):
            for cell in row:
                cell.font = Font(name="Courier New", size=9)


# ── Módulos sueltos (.bas / .cls / .frm) ─────────────────────────────────────

def split_module_header(content: str) -> tuple[str, str]:
    """Divide el contenido de un fichero .bas/.cls/.frm en (cabecera, código).
    La cabecera = todas las líneas hasta la última que empiece por 'Attribute '."""
    lines = content.splitlines(keepends=True)
    last_attr = -1
    for i, line in enumerate(lines):
        if line.strip().startswith("Attribute "):
            last_attr = i
    if last_attr == -1:
        return "", content
    return "".join(lines[: last_attr + 1]), "".join(lines[last_attr + 1 :])


def get_vb_name(filepath: str) -> str:
    """Devuelve el valor de Attribute VB_Name del fichero; usa el stem como fallback."""
    try:
        content = Path(filepath).read_text(encoding="cp1252", errors="replace")
        for line in content.splitlines():
            m = re.match(r'\s*Attribute VB_Name\s*=\s*"(.+)"', line)
            if m:
                return m.group(1)
    except Exception:
        pass
    return Path(filepath).stem


def read_module_file(filepath: str) -> "VBAModule":
    """Lee un .bas/.cls/.frm y devuelve VBAModule con la cabecera eliminada."""
    ext = Path(filepath).suffix.lower()
    kind_map = {".bas": "Module", ".cls": "ClassModule", ".frm": "Form"}
    name = Path(filepath).stem
    content = Path(filepath).read_text(encoding="cp1252", errors="replace")
    _, code = split_module_header(content)
    return VBAModule(name=name, code=code, kind=kind_map.get(ext, "Module"))


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Compara el código VBA de dos archivos .xlsm"
    )
    parser.add_argument("archivo_a", help="Primer archivo .xlsm")
    parser.add_argument("archivo_b", help="Segundo archivo .xlsm")
    parser.add_argument(
        "-o", "--output",
        default="vba_diff.xlsx",
        help="Ruta del Excel de salida (default: vba_diff.xlsx)"
    )
    parser.add_argument(
        "--show-equal",
        action="store_true",
        help="Mostrar también módulos idénticos en consola"
    )
    parser.add_argument(
        "--solo-consola",
        action="store_true",
        help="No generar archivo Excel"
    )
    parser.add_argument(
        "--solo-excel",
        action="store_true",
        help="No mostrar diff en consola"
    )
    parser.add_argument(
        "--incluir-iguales",
        action="store_true",
        help="Generar también hojas para módulos sin diferencias"
    )
    args = parser.parse_args()

    label_a = Path(args.archivo_a).name
    label_b = Path(args.archivo_b).name

    console = Console()
    console.print(f"\n[bold]Extrayendo VBA de:[/] {label_a}")
    modules_a = extract_vba(args.archivo_a)
    console.print(f"  -> {len(modules_a)} modulos encontrados")

    console.print(f"[bold]Extrayendo VBA de:[/] {label_b}")
    modules_b = extract_vba(args.archivo_b)
    console.print(f"  -> {len(modules_b)} modulos encontrados\n")

    diffs = compare(modules_a, modules_b)

    if not args.solo_excel:
        render_console(diffs, label_a, label_b, show_equal=args.show_equal)

    if not args.solo_consola:
        render_excel(diffs, label_a, label_b, args.output, args.incluir_iguales)
        console.print(f"[bold green]Excel guardado:[/] {args.output}\n")


if __name__ == "__main__":
    main()
