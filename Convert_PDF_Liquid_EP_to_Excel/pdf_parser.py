"""
pdf_parser.py — Extrae los registros del PDF "Listado de Liquidación de
Tasas Académicas de Matrícula" (Universidad Permanente UPUA, Universidad de
Alicante) y los convierte en una lista de Registro (uno por alumno/expediente).

Formato del PDF (ver CLAUDE.md para el detalle completo): informe paginado
con cabecera/pie repetidos en cada página y, por cada alumno, un bloque de
líneas:
  - línea de alumno:       Exped  Dni  Apellidos y Nombre
  - 1..n líneas de cobro:  Referencia  Fecha Cobro  Importe  Plazo  Forma Pago
  - Importe:                <suma de los importes cobrados>
  - Administrativo:         <gastos administrativos>
  - Importe Total :         <importe neto> €  (a veces en blanco en el propio
                             PDF — ver Registro.importe_total_pdf)

Extracción: fitz (PyMuPDF) con get_text("text", sort=True), que reordena las
palabras por posición y produce texto por líneas ya alineado con la tabla
visual (a diferencia del orden de lectura por defecto, que mezcla la
cabecera de columnas). Validado contra el PDF de muestra completo (206
páginas, 1645 registros): 0 líneas sin clasificar y la suma de importes
netos calculados cuadra exactamente con la línea "IMPORTE TOTAL POR TASAS"
del propio PDF.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import fitz  # PyMuPDF


# ── Patrones de línea ────────────────────────────────────────────────────

_REF_RE = re.compile(
    r'^(\d{13})\s+(\d{2}/\d{2}/\d{4})\s+([\d.,]+)\s+(\d+)\s+(\d+)\s*$')
_STUDENT_RE = re.compile(r'^(\d{1,6})\s+(\S+)\s+(.+?)\s*$')
_IMPORTE_RE = re.compile(r'^Importe:\s+([\d.,]+)\s*$')
_ADMIN_RE = re.compile(r'^Administrativo:\s*-?\s*([\d.,]+)\s*$')
_TOTAL_RE = re.compile(r'^Importe Total\s*:\s*([\d.,]*)\s*.?\s*$')

_NOISE_MARKERS = (
    "Curso Acad", "Listado de Liquidaci", "UNIVERSIDAD PERMANENTE",
    "Forma de Pago", "Devoluci", "IMPORTE TOTAL POR TASAS",
)


def _is_noise(line: str) -> bool:
    """Líneas de cabecera/pie de página que no forman parte de un registro."""
    # Pie "Página N de <total_páginas>": el nº total de páginas puede variar
    # entre documentos, así que se detecta por el patrón, no por un valor fijo.
    if re.match(r"^P.gina\s+\d+\s+de\s+\d+", line):
        return True
    if line.startswith("Exped") or "Apellidos y Nombre" in line:
        return True
    if line.startswith("Fecha") and "Importe" in line and "Forma" in line:
        return True
    if any(marker in line for marker in _NOISE_MARKERS):
        return True
    return False


def _to_float(value: str) -> float:
    """Convierte '1.234,56' (formato español) a float. '' -> 0.0."""
    value = value.strip()
    if not value:
        return 0.0
    return float(value.replace(".", "").replace(",", "."))


# ── Estructuras de datos ─────────────────────────────────────────────────

@dataclass
class Referencia:
    referencia: str
    fecha: str          # dd/mm/aaaa, tal cual aparece en el PDF
    importe: float
    plazo: int
    forma_pago: int


@dataclass
class Registro:
    exped: str
    dni: str
    nombre: str
    referencias: list[Referencia] = field(default_factory=list)
    importe: float = 0.0             # suma de importes cobrados ("Importe:")
    administrativo: float = 0.0      # gastos administrativos
    importe_total_pdf: float | None = None  # "Importe Total" según el PDF

    @property
    def importe_neto(self) -> float:
        """Importe neto de matrícula. Se calcula siempre como importe -
        administrativo en vez de confiar en el campo 'Importe Total' del
        PDF, porque ese campo viene en blanco en ~0,6% de los registros
        (defecto del propio informe origen). Verificado: cuando el PDF sí
        trae el valor, coincide exactamente con este cálculo en el 100%
        de los casos."""
        return round(self.importe - self.administrativo, 2)


# ── Parseo de un PDF ───────────────────────────────────────────────────────

def parse_pdf(path: str, on_page=None) -> tuple[list[Registro], int]:
    """Parsea un PDF de liquidación de tasas y devuelve
    (registros, nº_páginas).

    on_page: callback opcional `on_page(pagina_actual, total_paginas)`,
    útil para refrescar una barra de progreso desde la UI.
    """
    doc = fitz.open(path)
    total_paginas = doc.page_count
    records: list[Registro] = []
    cur: Registro | None = None

    for pno in range(total_paginas):
        text = doc[pno].get_text("text", sort=True)
        for raw in text.split("\n"):
            line = raw.strip()
            if not line or _is_noise(line):
                continue

            m = _REF_RE.match(line)
            if m:
                if cur is not None:
                    ref_str, fecha, importe, plazo, forma = m.groups()
                    cur.referencias.append(Referencia(
                        referencia=ref_str, fecha=fecha,
                        importe=_to_float(importe),
                        plazo=int(plazo), forma_pago=int(forma)))
                continue

            m = _IMPORTE_RE.match(line)
            if m:
                if cur is not None:
                    cur.importe = _to_float(m.group(1))
                continue

            m = _ADMIN_RE.match(line)
            if m:
                if cur is not None:
                    cur.administrativo = _to_float(m.group(1))
                continue

            m = _TOTAL_RE.match(line)
            if m:
                if cur is not None:
                    val = m.group(1)
                    cur.importe_total_pdf = _to_float(val) if val else None
                continue

            m = _STUDENT_RE.match(line)
            if m and m.group(1).isdigit():
                if cur is not None:
                    records.append(cur)
                cur = Registro(exped=m.group(1), dni=m.group(2),
                                nombre=m.group(3))
                continue
            # Línea no reconocida: se ignora en silencio (no debería
            # ocurrir sobre el formato validado, pero un PDF distinto
            # podría traer variaciones que aún no cubran los patrones
            # de arriba).

        if on_page:
            on_page(pno + 1, total_paginas)

    if cur is not None:
        records.append(cur)

    doc.close()
    return records, total_paginas
