"""
pdf_compare.py — Motor de comparación entre dos listados de liquidación de
tasas (dos versiones del mismo informe en momentos distintos), sobre las
listas de Registro que devuelve pdf_parser.parse_pdf.

Foco: CAMBIOS DE IMPORTE de un mismo alumno entre el listado A (anterior) y
el B (nuevo). Los alumnos presentes solo en uno de los dos ficheros no son
filas de diff (no son "cambios de importe"), pero se cuentan y se devuelven
para poder avisar en la UI — un importe puede "cambiar" simplemente porque
el alumno causó baja, y eso no debe pasar en silencio.

Emparejamiento: por Exped (clave primaria, estable entre versiones del mismo
curso), con el DNI como verificación — si un mismo Exped trae DNI distinto
en A y B (posible corrupción de fuente del PDF origen o desalineación), la
fila se marca con dni_distinto=True para avisar, comparando con la misma
normalización comodín (U+FFFD) que usa pdf_parser.corregir_alineacion.

Este módulo no conoce nada de tkinter ni de Excel: recibe dos list[Registro]
y devuelve una Comparacion. No modifica los Registro que recibe.
"""

# Última actualización: 2026-07-07 11:36

from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field

from pdf_parser import Registro


def _norm_texto(s: str) -> str:
    """Normaliza un texto (DNI o nombre) para compararlo entre A y B:
    mayúsculas, sin acentos ni espacios, conservando U+FFFD como comodín
    (glyphs corruptos del PDF origen)."""
    out = []
    for ch in unicodedata.normalize("NFKD", s):
        if unicodedata.combining(ch) or ch.isspace():
            continue
        out.append(ch if ch == "�" else ch.upper())
    return "".join(out)


def _iguales_comodin(a: str, b: str) -> bool:
    """True si a y b coinciden carácter a carácter tratando U+FFFD como
    comodín (misma longitud exigida)."""
    if len(a) != len(b):
        return False
    return all(x == y or x == "�" or y == "�" for x, y in zip(a, b))


@dataclass
class FilaDiff:
    """Un alumno presente en ambos listados cuyo importe cambió."""
    exped: str
    dni: str                 # el del listado B (el más reciente)
    nombre: str              # el del listado B
    importe_a: float         # "Importe:" (I.Acad. bruto) en A
    importe_b: float
    administrativo_a: float  # gastos administrativos en A
    administrativo_b: float
    dni_distinto: bool = False   # mismo Exped con DNI distinto en A y B

    @property
    def delta_importe(self) -> float:
        return round(self.importe_b - self.importe_a, 2)

    @property
    def delta_administrativo(self) -> float:
        return round(self.administrativo_b - self.administrativo_a, 2)

    @property
    def delta_neto(self) -> float:
        return round((self.importe_b - self.administrativo_b)
                     - (self.importe_a - self.administrativo_a), 2)


@dataclass
class Comparacion:
    """Resultado de comparar dos listados (A = anterior, B = nuevo)."""
    modificados: list[FilaDiff] = field(default_factory=list)
    solo_a: list[Registro] = field(default_factory=list)   # bajas
    solo_b: list[Registro] = field(default_factory=list)   # altas
    iguales: int = 0
    total_neto_a: float = 0.0    # Σ importe_neto del fichero A completo
    total_neto_b: float = 0.0

    @property
    def delta_total(self) -> float:
        return round(self.total_neto_b - self.total_neto_a, 2)

    @property
    def delta_modificados(self) -> float:
        """Δ neto agregado solo de los alumnos con importe cambiado."""
        return round(sum(f.delta_neto for f in self.modificados), 2)


def _indexar(records: list[Registro]) -> dict[str, list[Registro]]:
    """Índice Exped -> registros con ese Exped (lista por si el informe
    trajera un Exped repetido, aunque no se ha observado en las muestras)."""
    idx: dict[str, list[Registro]] = {}
    for r in records:
        idx.setdefault(r.exped, []).append(r)
    return idx


def _emparejar(la: list[Registro], lb: list[Registro]
               ) -> tuple[list[tuple[Registro, Registro]],
                          list[Registro], list[Registro]]:
    """Empareja los registros de un mismo Exped. Caso normal: 1 y 1. Si el
    Exped viniera repetido, casa primero por DNI normalizado (con comodín) y
    el resto por orden de aparición."""
    if len(la) == 1 and len(lb) == 1:
        return [(la[0], lb[0])], [], []
    pares: list[tuple[Registro, Registro]] = []
    resto_b = list(lb)
    resto_a = []
    for ra in la:
        na = _norm_texto(ra.dni)
        casado = next((rb for rb in resto_b
                       if _iguales_comodin(na, _norm_texto(rb.dni))), None)
        if casado is not None:
            resto_b.remove(casado)
            pares.append((ra, casado))
        else:
            resto_a.append(ra)
    while resto_a and resto_b:              # remanentes: por orden
        pares.append((resto_a.pop(0), resto_b.pop(0)))
    return pares, resto_a, resto_b


def comparar(records_a: list[Registro],
             records_b: list[Registro]) -> Comparacion:
    """Compara dos listados y devuelve los alumnos cuyo importe cambió
    (I.Acad. bruto o I.Adm., redondeados a céntimos), más los presentes solo
    en uno de los dos y los totales netos de cada fichero."""
    res = Comparacion(
        total_neto_a=round(sum(r.importe_neto for r in records_a), 2),
        total_neto_b=round(sum(r.importe_neto for r in records_b), 2))

    idx_a = _indexar(records_a)
    idx_b = _indexar(records_b)

    for exped in idx_a:
        if exped not in idx_b:
            res.solo_a.extend(idx_a[exped])
            continue
        pares, sobra_a, sobra_b = _emparejar(idx_a[exped], idx_b[exped])
        res.solo_a.extend(sobra_a)
        res.solo_b.extend(sobra_b)
        for ra, rb in pares:
            cambio = (round(ra.importe, 2) != round(rb.importe, 2)
                      or round(ra.administrativo, 2) != round(rb.administrativo, 2))
            if cambio:
                res.modificados.append(FilaDiff(
                    exped=exped, dni=rb.dni, nombre=rb.nombre,
                    importe_a=ra.importe, importe_b=rb.importe,
                    administrativo_a=ra.administrativo,
                    administrativo_b=rb.administrativo,
                    dni_distinto=not _iguales_comodin(
                        _norm_texto(ra.dni), _norm_texto(rb.dni))))
            else:
                res.iguales += 1

    for exped in idx_b:
        if exped not in idx_a:
            res.solo_b.extend(idx_b[exped])

    # Orden alfabético por nombre (mismo criterio visual que las tablas).
    res.modificados.sort(key=lambda f: _norm_texto(f.nombre))
    return res
