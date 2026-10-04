# -*- coding: utf-8 -*-
"""Vuelca a CSV las dos tablas de configuracion del menu y del Ribbon de PPub_BDatos_2026.xlsm.

- Tb_Tareas   (hoja _Menu_Aux, CodeName Prog__Menu_Aux) -> CustomUI/Tb_Tareas_snapshot.csv
- Lo_RibbonUI (hoja RibbonUI,  CodeName Prog__RibbonUI) -> CustomUI/Lo_RibbonUI_snapshot.csv

Las dos viven dentro del .xlsm (que no se versiona): sin estos volcados, sus cambios son
invisibles para git. Lee el libro como zip, sin abrir Excel; funciona aunque el libro este
abierto (lee lo ultimo GUARDADO). Localiza las hojas por CodeName, no por nombre de pestana.

Uso:  python CustomUI/Volcar_Tareas_y_Ribbon.py [tareas|ribbon] [ruta_libro.xlsm]
"""
# Last Rev. 2026-10-04 13:00
import csv
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
LIBRO = os.path.join(RAIZ, 'PPub_BDatos_2026.xlsm')
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
TABLAS = {  # clave -> (CodeName de la hoja, fichero de salida)
    'tareas': ('Prog__Menu_Aux', 'Tb_Tareas_snapshot.csv'),
    'ribbon': ('Prog__RibbonUI', 'Lo_RibbonUI_snapshot.csv'),
}


def _col2num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def _num2col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _limpiar(txt):
    """Deshace los escapes _xHHHH_ de Excel y aplana los saltos de linea a ' / '."""
    txt = re.sub(r'_x([0-9A-Fa-f]{4})_', lambda m: chr(int(m.group(1), 16)), txt)
    txt = txt.replace('\r\n', '\n').replace('\r', '\n')
    return txt.replace('\n', ' / ')


def leer_tabla(z, codename):
    """Devuelve (nombre_tabla, ref, filas) de la 1a tabla de la hoja con ese CodeName."""
    ss = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall('{%s}si' % NS):
            ss.append(''.join(t.text or '' for t in si.iter('{%s}t' % NS)))
    hoja = None
    for n in z.namelist():
        if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
            cab = z.read(n)[:4000].decode('utf-8', 'replace')
            m = re.search(r'codeName="([^"]+)"', cab)
            if m and m.group(1) == codename:
                hoja = n
                break
    if hoja is None:
        raise SystemExit(f'No hay ninguna hoja con CodeName {codename} en el libro')
    rels = 'xl/worksheets/_rels/' + os.path.basename(hoja) + '.rels'
    tabla = re.search(r'Target="\.\./tables/(table\d+\.xml)"', z.read(rels).decode('utf-8')).group(1)
    tx = z.read('xl/tables/' + tabla).decode('utf-8')
    nombre = re.search(r' name="([^"]+)"', tx).group(1)
    ref = re.search(r' ref="([^"]+)"', tx).group(1)
    a, b = ref.split(':')
    ca, ra = re.match(r'([A-Z]+)(\d+)', a).groups()
    cb, rb = re.match(r'([A-Z]+)(\d+)', b).groups()
    c1, c2, r1, r2 = _col2num(ca), _col2num(cb), int(ra), int(rb)
    datos = {}
    for row in ET.fromstring(z.read(hoja)).iter('{%s}row' % NS):
        r = int(row.get('r'))
        if not r1 <= r <= r2:
            continue
        for c in row.findall('{%s}c' % NS):
            col = _col2num(re.match(r'([A-Z]+)', c.get('r')).group(1))
            if not c1 <= col <= c2:
                continue
            v = c.find('{%s}v' % NS)
            if c.get('t') == 's' and v is not None:
                val = ss[int(v.text)]
            elif c.get('t') == 'inlineStr':
                val = ''.join(t.text or '' for t in c.iter('{%s}t' % NS))
            else:
                val = v.text if v is not None else ''
            datos[(r, col)] = _limpiar(val or '')
    filas = [[r] + [datos.get((r, col), '') for col in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
    cab = ['Fila_Excel'] + [_num2col(col) for col in range(c1, c2 + 1)]
    return nombre, ref, [cab] + filas


def main():
    args = sys.argv[1:]
    claves = [a for a in args if a in TABLAS] or list(TABLAS)
    libro = next((a for a in args if a.lower().endswith('.xlsm')), LIBRO)
    with zipfile.ZipFile(libro) as z:
        for clave in claves:
            codename, salida = TABLAS[clave]
            nombre, ref, filas = leer_tabla(z, codename)
            ruta = os.path.join(AQUI, salida)
            with open(ruta, 'w', newline='', encoding='utf-8-sig') as f:
                csv.writer(f, delimiter=';').writerows(filas)
            print(f'{nombre} ({codename}, {ref}): {len(filas) - 2} filas de datos -> {os.path.relpath(ruta, RAIZ)}')


if __name__ == '__main__':
    main()
