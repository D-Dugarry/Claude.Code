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
import unicodedata
from dataclasses import dataclass, field

import fitz  # PyMuPDF


# ── Patrones de línea ────────────────────────────────────────────────────

_REF_RE = re.compile(
    r'^(\d{13})\s+(\d{2}/\d{2}/\d{4})\s+(-?[\d.,]+)\s+(\d+)\s+(\d+)\s*$')
_STUDENT_RE = re.compile(r'^(\d{1,6})\s+(\S+)\s+(.+?)\s*$')
_IMPORTE_RE = re.compile(r'^Importe:\s+([\d.,]+)\s*$')
_ADMIN_RE = re.compile(r'^Administrativo:\s*-?\s*([\d.,]+)\s*$')
_TOTAL_RE = re.compile(r'^Importe Total\s*:\s*(-?[\d.,]*)\s*.?\s*$')

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

    # ── Corrección de errores de alineación DNI/Apellido ──────────────────
    # Cuando el PDF origen desalinea la fila del alumno, el campo DNI se
    # "come" el inicio del apellido (p. ej. dni="DEUL5Z83KFMARTIN",
    # nombre=", KRODEL"). corregir_alineacion() detecta el caso, guarda el
    # valor crudo aquí y aplica una separación heurística; revisar=True hasta
    # que el usuario la valide en la UI. Ver corregir_alineacion().
    revisar: bool = False            # fila con posible desalineación pendiente
    dni_raw: str = ""                # DNI+apellido pegados, tal cual el PDF
    nombre_raw: str = ""             # nombre original (suele empezar por ",")
    split_idx: int | None = None     # corte aplicado dentro de dni_raw

    @property
    def importe_neto(self) -> float:
        """Importe neto de matrícula. Se calcula siempre como importe -
        administrativo en vez de confiar en el campo 'Importe Total' del
        PDF, porque ese campo viene en blanco en ~0,6% de los registros
        (defecto del propio informe origen). Verificado: cuando el PDF sí
        trae el valor, coincide exactamente con este cálculo en el 100%
        de los casos."""
        return round(self.importe - self.administrativo, 2)

    def aplicar_split(self, idx: int) -> None:
        """Separa dni_raw en DNI (primeros `idx` caracteres) + fragmento de
        apellido (el resto) y reconstruye el nombre completo. Idempotente:
        siempre parte de dni_raw/nombre_raw, no del valor ya corregido."""
        frag = self.dni_raw[idx:]
        self.dni = self.dni_raw[:idx]
        base = self.nombre_raw
        if base.lstrip().startswith(","):
            self.nombre = frag + base           # "MARTIN" + ", KRODEL"
        elif frag:
            self.nombre = f"{frag} {base}"
        else:
            self.nombre = base
        self.split_idx = idx


# ── Detección y corrección de desalineación DNI/Apellido ───────────────────

def _norm_apellido(s: str) -> str:
    """Normaliza la parte de apellidos (antes de la primera coma) para poder
    compararla alfabéticamente: mayúsculas, sin acentos, conservando el
    carácter de reemplazo U+FFFD como comodín (los nombres del PDF traen
    glyphs corruptos que no se pueden recuperar)."""
    s = s.split(",")[0]
    out = []
    for ch in unicodedata.normalize("NFKD", s):
        if unicodedata.combining(ch):
            continue
        if ch == "�":
            out.append("�")
        elif ch.isalpha():
            out.append(ch.upper())
        elif ch == " ":
            out.append(" ")
    return "".join(out)


def _es_sospechosa(r: Registro) -> bool:
    """Señales de desalineación: DNI anormalmente largo (>10; los válidos
    miden 8-9) o el nombre empieza por coma (el apellido se lo comió el DNI)."""
    return len(r.dni) > 10 or r.nombre.lstrip().startswith(",")


def _vecino_apellido(records: list[Registro], i: int, paso: int) -> str:
    """Apellido normalizado del primer vecino NO sospechoso en la dirección
    `paso` (-1 anterior, +1 siguiente). El PDF va ordenado por 'Apellidos y
    Nombre', así que estos vecinos acotan alfabéticamente el apellido real."""
    j = i + paso
    while 0 <= j < len(records):
        if not records[j].revisar:
            return _norm_apellido(records[j].nombre)
        j += paso
    return ""


def _prefijo_comodin(a: str, b: str) -> int:
    """Longitud del prefijo común de a y b, tratando U+FFFD como comodín."""
    n = 0
    for x, y in zip(a, b):
        if x == y or x == "�" or y == "�":
            n += 1
        else:
            break
    return n


def _sugerir_split(records: list[Registro], i: int) -> int | None:
    """Devuelve el índice de corte más probable dentro de dni_raw (longitud
    del DNI), o None si la heurística no encuentra un candidato fiable.

    Idea: el fragmento de apellido es un sufijo alfabético del texto pegado.
    Se prueba cada corte y se acepta aquel cuyo fragmento encaja como prefijo
    completo del apellido de un vecino (o viceversa), aprovechando que el PDF
    está ordenado alfabéticamente. No es fiable al 100% (por eso la fila queda
    marcada para revisión), pero acierta en los casos observados."""
    merged = records[i].dni_raw
    vecinos = [_vecino_apellido(records, i, +1),
               _vecino_apellido(records, i, -1)]
    mejor: tuple[int, int] | None = None            # (score, idx)
    for idx in range(4, len(merged)):
        frag = merged[idx:]
        if not frag or not all(c.isalpha() or c == "�" for c in frag):
            continue
        fnorm = _norm_apellido(frag)
        if not fnorm:
            continue
        for vecino in vecinos:
            if not vecino:
                continue
            m = _prefijo_comodin(fnorm, vecino)
            comun = min(len(fnorm), len(vecino))
            if m == comun and comun >= 3:           # el más corto es prefijo
                if mejor is None or m > mejor[0]:
                    mejor = (m, idx)
    return mejor[1] if mejor else None


def corregir_alineacion(records: list[Registro]) -> int:
    """Detecta filas con posible desalineación DNI/Apellido y aplica una
    separación heurística basada en el orden alfabético de los vecinos. Marca
    cada fila afectada con revisar=True (para validación manual en la UI) y
    guarda el valor crudo en dni_raw/nombre_raw. No toca importes, así que no
    afecta al checksum de 'IMPORTE TOTAL POR TASAS'. Devuelve el nº de filas
    marcadas. Se llama tras parse_pdf (no dentro, para no alterar el test de
    parseo)."""
    sospechosas = [i for i, r in enumerate(records) if _es_sospechosa(r)]
    for i in sospechosas:                           # 1ª pasada: marcar todas
        r = records[i]
        r.dni_raw = r.dni
        r.nombre_raw = r.nombre
        r.revisar = True
    for i in sospechosas:                           # 2ª pasada: separar
        idx = _sugerir_split(records, i)
        if idx is not None:
            records[i].aplicar_split(idx)
    return len(sospechosas)


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
