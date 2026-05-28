"""
VBA Apply — lee la columna Resultado del vba_diff.xlsx y la escribe en el .xlsm destino.

Flujo:
  1. compare.py genera vba_diff.xlsx con columna Resultado pre-rellena
  2. El usuario edita la columna Resultado en Excel
  3. apply.py lee esa columna y sobreescribe el VBA en el archivo destino

Uso:
  python apply.py vba_diff.xlsx destino.xlsm
  python apply.py vba_diff.xlsx destino.xlsm --dry-run
"""

import sys
import os
import re
import difflib
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("ERROR: openpyxl no instalado. Ejecuta: pip install openpyxl")

try:
    import win32com.client
except ImportError:
    sys.exit("ERROR: pywin32 no instalado. Ejecuta: pip install pywin32")


SKIP_SHEETS   = {"Indice"}
COL_RESULTADO = 5   # columna E  →  código final editable
COL_MOD_NAME  = 5   # celda E1   →  nombre original del módulo (escrito por compare.py)
COL_STATUS    = 6   # celda F1   →  estado ("equal", "modified", "only_a", "only_b")


# ── Lectura del diff Excel ────────────────────────────────────────────────────

def read_diff_excel(diff_path: str) -> dict[str, str]:
    """
    Lee el Excel de diff y devuelve {nombre_módulo_original: código_resultado}.
    El nombre original del módulo se lee de E1 en cada hoja.
    Si E1 está vacío, usa el nombre de la hoja como fallback.
    """
    if not os.path.exists(diff_path):
        sys.exit(f"ERROR: no se encuentra el diff '{diff_path}'")

    wb = openpyxl.load_workbook(diff_path)
    results: dict[str, str] = {}

    for sheet_name in wb.sheetnames:
        if sheet_name in SKIP_SHEETS:
            continue

        ws = wb[sheet_name]

        # nombre original del módulo desde E1
        orig_name = ws.cell(1, COL_MOD_NAME).value
        if not orig_name:
            orig_name = sheet_name

        # saltear módulos sin diferencias o no comparados
        status = ws.cell(1, COL_STATUS).value
        if status in ("equal", "skip"):
            continue

        # reconstruir código desde columna E, filas 3+
        lines: list[str] = []
        for row in ws.iter_rows(min_row=3, min_col=COL_RESULTADO,
                                 max_col=COL_RESULTADO, values_only=True):
            val = row[0]
            lines.append(str(val) if val is not None else "")

        # eliminar líneas vacías del final
        while lines and lines[-1].strip() == "":
            lines.pop()

        results[str(orig_name)] = "\n".join(lines)

    return results


# ── Escritura en el .xlsm ─────────────────────────────────────────────────────

def apply_vba(xlsm_path: str, module_codes: dict[str, str],
              dry_run: bool = False) -> None:
    abs_path = str(Path(xlsm_path).resolve())
    if not os.path.exists(abs_path):
        sys.exit(f"ERROR: no se encuentra el archivo destino '{abs_path}'")

    if dry_run:
        # Importar extract_vba de compare.py para leer el código actual del destino
        from compare import extract_vba

        print(f"\nComparando columna Resultado con código actual del destino...\n")

        try:
            current_modules = extract_vba(abs_path)
        except Exception as e:
            print(f"\nERROR al leer el destino: {e}")
            return

        total_add = total_del = modules_changed = 0

        for name, resultado_code in module_codes.items():
            current_mod  = current_modules.get(name)
            current_lines  = current_mod.code.splitlines() if current_mod else []
            resultado_lines = resultado_code.splitlines()

            diff = list(difflib.unified_diff(
                current_lines,
                resultado_lines,
                fromfile=f"Destino   ({name})",
                tofile=f"Resultado ({name})",
                lineterm="",
                n=2
            ))

            adds = sum(1 for l in diff if l.startswith("+") and not l.startswith("+++"))
            dels = sum(1 for l in diff if l.startswith("-") and not l.startswith("---"))

            if not diff:
                print(f"── {name}   ✓ sin cambios")
            else:
                modules_changed += 1
                total_add += adds
                total_del += dels

                # Una sola línea de cabecera:
                # spacing1 y spacing3 en el tag "diff_head" añaden espacio
                # de ancho COMPLETO (garantizado por Tk) arriba y abajo.
                stats = f"+{adds} se añadirían  ·  -{dels} se eliminarían"
                print(f"\n──  {name}   {stats}")

                for line in diff:
                    if line.startswith("--- ") or line.startswith("+++ "):
                        continue
                    print(line)
                print()

        sep = "─" * 55
        print(sep)
        if modules_changed == 0:
            print("✓ La columna Resultado es idéntica al código del destino.")
        else:
            print(f"Resumen: {modules_changed} módulos con cambios"
                  f"  ·  +{total_add} añadidas  ·  -{total_del} eliminadas")
        print("⚠ Ningún archivo modificado  —  pulsa Aplicar para escribir los cambios")
        return

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_path, ReadOnly=False)
        try:
            vbp = wb.VBProject
        except Exception:
            sys.exit(
                "ERROR: no se puede acceder al VBProject.\n"
                "Habilita 'Confiar en el acceso al modelo de objetos VBA' en\n"
                "Opciones de Excel > Centro de confianza > Configuración de macros."
            )

        applied: list[str] = []
        not_found: list[str] = []

        # mapa de nombres de módulos en el .xlsm para búsqueda rápida
        xlsm_modules = {comp.Name: comp for comp in vbp.VBComponents}

        for mod_name, new_code in module_codes.items():
            comp = xlsm_modules.get(mod_name)
            if comp is None:
                not_found.append(mod_name)
                continue

            cm = comp.CodeModule
            total_lines = cm.CountOfLines

            # borrar todas las líneas actuales del módulo
            if total_lines > 0:
                cm.DeleteLines(1, total_lines)

            # insertar las nuevas líneas una a una
            if new_code.strip():
                for i, line in enumerate(new_code.splitlines(), 1):
                    cm.InsertLines(i, line)

            applied.append(mod_name)

        wb.Save()
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass  # macro BeforeClose puede fallar en modo COM; el guardado ya se hizo

        print(f"\nAplicados correctamente: {len(applied)} módulos")
        if not_found:
            print(f"No encontrados en el .xlsm:  {len(not_found)}")
            for n in not_found:
                print(f"  - {n}")

    finally:
        try:
            excel.Quit()
        except Exception:
            pass


# ── Escritura en fichero de módulo (.bas / .cls / .frm) ──────────────────────

_KIND_EXT = {"Module": ".bas", "ClassModule": ".cls", "Form": ".frm", "Document": ".cls"}


def apply_module_folder(dest_folder: str, diff_path: str, dry_run: bool = False) -> None:
    """Lee Modul_Diff.xlsx y escribe cada módulo al fichero correspondiente en dest_folder."""
    from compare import split_module_header

    dest = Path(dest_folder).resolve()
    if not dest.is_dir():
        sys.exit(f"ERROR: la carpeta destino no existe: '{dest}'")

    wb = openpyxl.load_workbook(diff_path)
    total_changed = total_skipped = 0

    for sheet_name in wb.sheetnames:
        if sheet_name in SKIP_SHEETS:
            continue
        ws = wb[sheet_name]
        orig_name = ws.cell(1, COL_MOD_NAME).value or sheet_name
        status    = ws.cell(1, COL_STATUS).value
        kind      = ws.cell(1, 7).value or "Module"   # G1 — escrito por compare.py

        if status in ("equal", "skip"):
            total_skipped += 1
            continue

        # Reconstruir código desde columna Resultado
        lines: list[str] = []
        for row in ws.iter_rows(min_row=3, min_col=COL_RESULTADO,
                                 max_col=COL_RESULTADO, values_only=True):
            val = row[0]
            lines.append(str(val) if val is not None else "")
        while lines and lines[-1].strip() == "":
            lines.pop()
        new_code = "\n".join(lines)

        # Buscar el fichero en dest_folder (cualquier extensión válida)
        ext = _KIND_EXT.get(kind, ".bas")
        dest_file = dest / f"{orig_name}{ext}"
        if not dest_file.exists():
            for e in (".bas", ".cls", ".frm"):
                alt = dest / f"{orig_name}{e}"
                if alt.exists():
                    dest_file = alt
                    break

        if not dest_file.exists():
            # El módulo no existe en destino → crearlo
            if not dry_run:
                dest_file.write_text(new_code + "\n", encoding="cp1252")
                print(f"  + {orig_name}{ext}  (creado)")
            else:
                print(f"  + {orig_name}{ext}  (se crearía en destino)")
            total_changed += 1
            continue

        abs_dest = str(dest_file)
        content = dest_file.read_text(encoding="cp1252", errors="replace")
        header, current_code = split_module_header(content)

        if dry_run:
            import difflib
            diff = list(difflib.unified_diff(
                current_code.splitlines(), new_code.splitlines(),
                fromfile=f"Destino   ({dest_file.name})",
                tofile=f"Resultado ({dest_file.name})",
                lineterm="", n=2,
            ))
            adds = sum(1 for l in diff if l.startswith("+") and not l.startswith("+++"))
            dels = sum(1 for l in diff if l.startswith("-") and not l.startswith("---"))
            if not diff:
                print(f"── {dest_file.name}   ✓ sin cambios")
            else:
                total_changed += 1
                print(f"\n──  {dest_file.name}   +{adds} se añadirían  ·  -{dels} se eliminarían")
                for line in diff:
                    if line.startswith("--- ") or line.startswith("+++ "):
                        continue
                    print(line)
                print()
        else:
            new_content = header
            if header and not header.endswith("\n"):
                new_content += "\n"
            new_content += new_code
            if new_code and not new_code.endswith("\n"):
                new_content += "\n"
            dest_file.write_text(new_content, encoding="cp1252")
            print(f"  ✓ {dest_file.name}  aplicado")
            total_changed += 1

    sep = "─" * 55
    print(sep)
    if dry_run:
        if total_changed == 0:
            print("✓ El resultado es idéntico al código de destino.")
        else:
            print(f"Resumen: {total_changed} módulos con cambios")
        print("⚠ Ningún archivo modificado  —  pulsa Aplicar para escribir los cambios")
    else:
        print(f"Aplicados: {total_changed} módulos")


def apply_module_file(dest_path: str, new_code: str, dry_run: bool = False) -> None:
    """Escribe new_code en un .bas/.cls/.frm conservando la cabecera original."""
    from compare import split_module_header

    abs_path = str(Path(dest_path).resolve())
    if not os.path.exists(abs_path):
        sys.exit(f"ERROR: no se encuentra el destino '{abs_path}'")

    content = Path(abs_path).read_text(encoding="cp1252", errors="replace")
    header, current_code = split_module_header(content)

    if dry_run:
        diff = list(difflib.unified_diff(
            current_code.splitlines(), new_code.splitlines(),
            fromfile=f"Destino   ({Path(dest_path).name})",
            tofile=f"Resultado ({Path(dest_path).name})",
            lineterm="", n=2,
        ))
        adds = sum(1 for l in diff if l.startswith("+") and not l.startswith("+++"))
        dels = sum(1 for l in diff if l.startswith("-") and not l.startswith("---"))
        if not diff:
            print(f"── {Path(dest_path).name}   ✓ sin cambios")
        else:
            print(f"\n──  {Path(dest_path).name}   +{adds} se añadirían  ·  -{dels} se eliminarían")
            for line in diff:
                if line.startswith("--- ") or line.startswith("+++ "):
                    continue
                print(line)
            print()
        sep = "─" * 55
        print(sep)
        if not diff:
            print("✓ El resultado es idéntico al código del destino.")
        else:
            print(f"Resumen: +{adds} añadidas  ·  -{dels} eliminadas")
        print("⚠ Ningún archivo modificado  —  pulsa Aplicar para escribir los cambios")
        return

    new_content = header
    if header and not header.endswith("\n"):
        new_content += "\n"
    new_content += new_code
    if new_code and not new_code.endswith("\n"):
        new_content += "\n"
    Path(abs_path).write_text(new_content, encoding="cp1252")
    print(f"\nAplicado correctamente: {Path(dest_path).name}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Aplica la columna Resultado del diff Excel a un archivo .xlsm"
    )
    parser.add_argument("diff_excel",   help="Archivo vba_diff.xlsx generado por compare.py")
    parser.add_argument("target_xlsm",  help="Archivo .xlsm donde se aplicarán los cambios")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Muestra qué módulos se modificarían sin escribir nada en el .xlsm"
    )
    args = parser.parse_args()

    print(f"\nLeyendo diff:  {Path(args.diff_excel).name}")
    module_codes = read_diff_excel(args.diff_excel)
    print(f"  -> {len(module_codes)} módulos encontrados en el diff")

    print(f"Destino:       {Path(args.target_xlsm).name}")
    apply_vba(args.target_xlsm, module_codes, dry_run=args.dry_run)

    if not args.dry_run:
        print("\nArchivo guardado correctamente.\n")


if __name__ == "__main__":
    main()
