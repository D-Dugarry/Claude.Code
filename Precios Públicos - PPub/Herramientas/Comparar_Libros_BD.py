"""Compara dos copias del libro PPub (ANTES y DESPUÉS de un cambio de código) sin abrir Excel.

Pensado para la "doble ejecución": se ejecuta el mismo proceso, sobre el mismo fichero de entrada,
con el código antiguo y con el nuevo, se guarda cada resultado como una copia del .xlsm y este
script dice si las dos son idénticas.

Compara, celda a celda y dentro del rango de cada tabla:
  - el valor (texto exacto; los números tal como los guarda Excel, así que distingue hasta el
    último decimal),
  - el color de relleno (lo que pinta M_112 en los duplicados),
  - el formato de número.
Y además el log del proceso (nombre APP_Task_Inf) y el informe de cada botón del Ribbon (columna
Informe_Rut de Lo_RibbonUI, que guarda el log de la última ejecución de cada botón: así se comparan los
de M_110, M_210 y M_310 aunque APP_Task_Inf solo tenga el último), quitando las horas y los tiempos.
Las líneas de tiempos de cada paso del formateo ("Col. 14 F_Emi (F)....1,34 seg.") se quitan enteras,
porque su texto cambia con el código (p. ej. "Col. 11 Ref (N, bloque 11-13)").

Con --por, las diferencias de cada hoja se cuentan también por el valor de esa columna en la fila
(p. ej. --por C_Acad: cuántas diferencias de cada columna caen en cada curso).

Con --sin-hora, las columnas indicadas (por su cabecera) se comparan sin la hora: 46237,0034 y 46237
cuentan como iguales, y las celdas que solo difieren en la hora se cuentan aparte. Es para comparar
con un ANTES de cuando las fechas se guardaban con hora.

Uso:
    python Comparar_Libros_BD.py ANTES.xlsm DESPUES.xlsm
    python Comparar_Libros_BD.py ANTES.xlsm DESPUES.xlsm --hojas Sht__BD --max 50
    python Comparar_Libros_BD.py ANTES.xlsm DESPUES.xlsm --sin-hora F_Emi,F_Vto,F_Cob
    python Comparar_Libros_BD.py ANTES.xlsm DESPUES.xlsm --por C_Acad

Las hojas se buscan por su CodeName (Sht__BD...), no por el nombre de la pestaña.
Lee las hojas en streaming: con 177.000 filas tarda un par de minutos y no carga el libro en memoria.
"""
# Last Rev. 2026-10-04 17:10

import argparse
import math
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": M, "r": R}
T_ROW, T_C, T_V, T_IS, T_T = ("{%s}%s" % (M, x) for x in ("row", "c", "v", "is", "t"))

HOJAS_DEFECTO = ["Sht__BD", "Sht__BD_Dupl", "Sht__BD_ErrDate"]
NOMBRE_LOG = "APP_Task_Inf"


def col2num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def num2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ref2rc(ref):
    m = re.match(r"\$?([A-Z]+)\$?(\d+)", ref)
    return int(m.group(2)), col2num(m.group(1))


def texto_si(el):
    """Texto de un <si> o <is> (con o sin rich text)."""
    return "".join(t.text or "" for t in el.iter(T_T))


class Libro:
    def __init__(self, ruta):
        self.ruta = ruta
        self.z = zipfile.ZipFile(ruta)
        wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))}
        self.hojas = {}                                     # nombre pestaña -> fichero xml
        for s in wb.find("m:sheets", NS):
            destino = rels[s.get("{%s}id" % R)]
            self.hojas[s.get("name")] = "xl/" + destino.replace("/xl/", "").lstrip("/")
        self.nombres = {}
        dn = wb.find("m:definedNames", NS)
        if dn is not None:
            for d in dn:
                if d.get("localSheetId") is None:
                    self.nombres[d.get("name")] = d.text
        self.codename = {}                                  # CodeName -> nombre pestaña
        for nombre, f in self.hojas.items():
            cab = self.z.open(f).read(4000).decode("utf-8", "ignore")
            m = re.search(r'codeName="([^"]+)"', cab)
            if m:
                self.codename[m.group(1)] = nombre
        self._cargar_cadenas()
        self._cargar_estilos()

    def _cargar_cadenas(self):
        self.ss = []
        if "xl/sharedStrings.xml" in self.z.namelist():
            for ev, el in ET.iterparse(self.z.open("xl/sharedStrings.xml"), events=("end",)):
                if el.tag == "{%s}si" % M:
                    self.ss.append(texto_si(el))
                    el.clear()

    def _cargar_estilos(self):
        st = ET.fromstring(self.z.read("xl/styles.xml"))
        fmts = {}
        nf = st.find("m:numFmts", NS)
        if nf is not None:
            for f in nf:
                fmts[f.get("numFmtId")] = f.get("formatCode")
        fills = [ET.tostring(f, encoding="unicode").replace(' xmlns="%s"' % M, "")
                 for f in st.find("m:fills", NS)]
        self.estilo = []                                    # índice s -> (relleno, formato de número)
        for xf in st.find("m:cellXfs", NS):
            fid = int(xf.get("fillId", "0"))
            nid = xf.get("numFmtId", "0")
            relleno = fills[fid] if fid > 1 else ""         # 0 y 1 son los rellenos por defecto (none / gray125)
            self.estilo.append((relleno, fmts.get(nid, "builtin:" + nid)))

    def tabla_de(self, hoja):
        """(nombre, fila1, col1, fila2, col2, cabeceras) de la primera tabla de la hoja."""
        f = self.hojas[hoja]
        relf = f.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if relf not in self.z.namelist():
            return None
        for r in ET.fromstring(self.z.read(relf)):
            if "tables/" in r.get("Target"):
                t = ET.fromstring(self.z.read("xl/tables/" + r.get("Target").split("tables/")[1]))
                a, b = t.get("ref").split(":")
                f1, c1 = ref2rc(a)
                f2, c2 = ref2rc(b)
                f2 -= int(t.get("totalsRowCount", "0"))     # la fila de totales no es de datos (y depende de ShowTotals)
                cab = [tc.get("name") for tc in t.find("m:tableColumns", NS)]
                return t.get("name"), f1, c1, f2, c2, cab
        return None

    def filas(self, hoja, f1, c1, f2, c2):
        """Genera (nº fila, {col: (valor, relleno, formato)}) de las filas del rango, en orden."""
        for ev, el in ET.iterparse(self.z.open(self.hojas[hoja]), events=("end",)):
            if el.tag != T_ROW:
                continue
            nf = int(el.get("r"))
            if f1 <= nf <= f2:
                celdas = {}
                for c in el.findall(T_C):
                    fila, col = ref2rc(c.get("r"))
                    if not (c1 <= col <= c2):
                        continue
                    t = c.get("t", "n")
                    v = c.find(T_V)
                    if t == "s" and v is not None:
                        val = "s:" + self.ss[int(v.text)]
                    elif t == "inlineStr":
                        val = "s:" + texto_si(c.find(T_IS))
                    elif v is not None:
                        val = t + ":" + (v.text or "")
                    else:
                        val = None
                    relleno, fmt = self.estilo[int(c.get("s", "0"))]
                    if val is None and not relleno:
                        continue                            # celda vacía sin color: igual que no existir
                    celdas[col] = (val, relleno, fmt)
                yield nf, celdas
            el.clear()

    def valor_nombre(self, nombre):
        ref = self.nombres.get(nombre)
        if not ref:
            return None
        m = re.match(r"'?(.+?)'?!(\$?[A-Z]+\$?\d+)", ref)
        hoja, celda = m.group(1), m.group(2).replace("$", "")
        fila, col = ref2rc(celda)
        for nf, celdas in self.filas(hoja, fila, col, fila, col):
            if col in celdas and celdas[col][0]:
                return celdas[col][0][2:]
        return ""


def quitar_hora(val):
    """'n:46237.003368055557' -> 'n:46237' (la fecha sin la hora); el resto, tal cual."""
    if val and val.startswith("n:"):
        x = float(val[2:])
        if x != math.floor(x):
            return "n:%d" % math.floor(x)
    return val


def comparar_hoja(A, B, codename, maximo, sin_hora=(), por=None):
    print("\n" + "=" * 100 + "\n" + codename)
    if codename not in A.codename or codename not in B.codename:
        print("   No existe en los dos libros.")
        return 1
    ha, hb = A.codename[codename], B.codename[codename]
    ta, tb = A.tabla_de(ha), B.tabla_de(hb)
    if not ta or not tb:
        print("   Sin tabla en alguno de los libros.")
        return 1
    print("   ANTES:   %s  %s%d:%s%d  (%d filas)" % (ta[0], num2col(ta[2]), ta[1], num2col(ta[4]), ta[3], ta[3] - ta[1]))
    print("   DESPUÉS: %s  %s%d:%s%d  (%d filas)" % (tb[0], num2col(tb[2]), tb[1], num2col(tb[4]), tb[3], tb[3] - tb[1]))
    if ta[5] != tb[5]:
        print("   ¡Las cabeceras de las tablas no coinciden!")
    cab = ta[5]
    f1, c1 = min(ta[1], tb[1]), min(ta[2], tb[2])
    f2, c2 = max(ta[3], tb[3]), max(ta[4], tb[4])

    col_por = cab.index(por) + ta[2] if por and por in cab else None
    if por and col_por is None:
        print("   (--por: la tabla no tiene la columna %r)" % por)
    difs = {"valor": 0, "relleno": 0, "formato": 0}
    solo_hora = {}                                          # columna -> celdas que solo difieren en la hora
    por_col = {}
    por_grupo = {}                                          # (columna, tipo, valor de --por) -> celdas
    mostradas = 0
    ga, gb = A.filas(ha, f1, c1, f2, c2), B.filas(hb, f1, c1, f2, c2)
    ra, rb = next(ga, None), next(gb, None)
    while ra or rb:
        if rb is None or (ra and ra[0] < rb[0]):
            fila, ca, cb = ra[0], ra[1], {}
            ra = next(ga, None)
        elif ra is None or rb[0] < ra[0]:
            fila, ca, cb = rb[0], {}, rb[1]
            rb = next(gb, None)
        else:
            fila, ca, cb = ra[0], ra[1], rb[1]
            ra, rb = next(ga, None), next(gb, None)
        if ca == cb:
            continue
        for col in sorted(set(ca) | set(cb)):
            va, vb = ca.get(col, (None, "", "")), cb.get(col, (None, "", ""))
            if va == vb:
                continue
            nombre = cab[col - ta[2]] if 0 <= col - ta[2] < len(cab) else "?"
            if nombre in sin_hora:
                va2, vb2 = (quitar_hora(va[0]),) + va[1:], (quitar_hora(vb[0]),) + vb[1:]
                if va[0] != vb[0] and va2[0] == vb2[0]:
                    solo_hora[nombre] = solo_hora.get(nombre, 0) + 1
                va, vb = va2, vb2
                if va == vb:
                    continue
            for i, tipo in enumerate(("valor", "relleno", "formato")):
                if va[i] != vb[i] and not (tipo == "formato" and (va[0] is None or vb[0] is None)):
                    difs[tipo] += 1
                    por_col[(nombre, tipo)] = por_col.get((nombre, tipo), 0) + 1
                    if col_por is not None:
                        g = (ca.get(col_por) or cb.get(col_por) or (None,))[0]
                        g = g[2:] if g else "(vacío)"
                        por_grupo[(nombre, tipo, g)] = por_grupo.get((nombre, tipo, g), 0) + 1
                    if mostradas < maximo:
                        mostradas += 1
                        print("   %s%-7d %-22s %-8s ANTES=%r  DESPUÉS=%r" % (
                            num2col(col), fila, nombre[:22], tipo, va[i], vb[i]))
    total = sum(difs.values())
    for nombre, n in solo_hora.items():
        print("   %-30s %d celdas iguales salvo la hora (no cuentan como diferencia)" % (nombre, n))
    if total == 0:
        print("   IDÉNTICAS (valores, rellenos y formatos de número).")
    else:
        print("   DIFERENCIAS: %s" % difs)
        for (nombre, tipo), n in sorted(por_col.items(), key=lambda x: -x[1])[:20]:
            print("      %-30s %-8s %d" % (nombre, tipo, n))
        if por_grupo:
            print("   Por %s:" % por)
            for (nombre, tipo, g), n in sorted(por_grupo.items(), key=lambda x: (x[0][0], x[0][1], -x[1])):
                print("      %-30s %-8s %-12s %d" % (nombre, tipo, g, n))
    return total


# Líneas de tiempos de cada paso del formateo (Rut_Lo_Format_LoData_LoDefColData): fuera enteras
RE_TIEMPOS_FORMATEO = re.compile(r"^\s*(Quitar formatos y validaciones|Col\. \d+ .*\((T|F|N)(,[^)]*)?\) <t> seg\.|Anchos, alineaci)")


def limpiar_log(txt):
    txt = re.sub(r"\d\d:\d\d:\d\d Lap: +[\d.,]+ seg\. ", "", txt or "")
    txt = re.sub(r"\.*\s*[\d.]*\d,\d+ seg\.", " <t> seg.", txt)    # tiempos de cada paso: "Col. 14 F_Emi (F)....1,34 seg."
    txt = re.sub(r"\d\d-\w{3}-\d\d \d\d:\d\d", "<fecha>", txt)
    txt = re.sub(r"\d\d/\d\d/\d{4} \d{1,2}:\d\d:\d\d", "<fecha>", txt)
    return [l.rstrip() for l in txt.replace("\r", "").replace("_x000D_", "").split("\n")
            if not RE_TIEMPOS_FORMATEO.search(l)]


def informes_botones(L):
    """{tag: informe} de la columna Informe_Rut de Lo_RibbonUI (hoja con CodeName Prog__RibbonUI)."""
    hoja = L.codename.get("Prog__RibbonUI")
    t = L.tabla_de(hoja) if hoja else None
    if not t or "Uribbon-Tags" not in t[5] or "Informe_Rut" not in t[5]:
        return {}
    c_tag, c_inf = t[5].index("Uribbon-Tags") + t[2], t[5].index("Informe_Rut") + t[2]
    out = {}
    for nf, celdas in L.filas(hoja, t[1] + 1, t[2], t[3], t[4]):
        tag = (celdas.get(c_tag) or (None,))[0]
        if tag:
            inf = (celdas.get(c_inf) or (None,))[0]
            out[tag[2:]] = inf[2:] if inf else ""
    return out


def comparar_logs(la, lb, titulo):
    """Compara dos logs ya limpios; devuelve 0 si son iguales."""
    if la == lb:
        print("   %-40s IDÉNTICO (%d líneas)." % (titulo, len(la)))
        return 0
    import difflib
    print("   %-40s DIFERENTE:" % titulo)
    for l in list(difflib.unified_diff(la, lb, "ANTES", "DESPUÉS", lineterm="", n=0))[:80]:
        print("      " + l)
    return 1


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("antes")
    ap.add_argument("despues")
    ap.add_argument("--hojas", default=",".join(HOJAS_DEFECTO), help="CodeNames separados por comas")
    ap.add_argument("--max", type=int, default=30, help="diferencias a mostrar por hoja")
    ap.add_argument("--sin-hora", default="", help="cabeceras de fecha a comparar sin la hora, separadas por comas")
    ap.add_argument("--por", default=None, help="cabecera por cuyo valor se cuentan las diferencias (p. ej. C_Acad)")
    args = ap.parse_args()

    A, B = Libro(args.antes), Libro(args.despues)
    sin_hora = {c.strip() for c in args.sin_hora.split(",") if c.strip()}
    total = 0
    for h in args.hojas.split(","):
        total += comparar_hoja(A, B, h.strip(), args.max, sin_hora, args.por)

    print("\n" + "=" * 100 + "\nLogs, sin horas ni tiempos")
    total += comparar_logs(limpiar_log(A.valor_nombre(NOMBRE_LOG)), limpiar_log(B.valor_nombre(NOMBRE_LOG)), NOMBRE_LOG)
    ia, ib = informes_botones(A), informes_botones(B)
    for tag in sorted(set(ia) | set(ib)):
        if ia.get(tag) or ib.get(tag):
            total += comparar_logs(limpiar_log(ia.get(tag)), limpiar_log(ib.get(tag)), "Informe_Rut de " + tag)

    print("\nRESULTADO: " + ("IDÉNTICOS" if total == 0 else "HAY DIFERENCIAS"))
    sys.exit(0 if total == 0 else 1)


if __name__ == "__main__":
    main()
