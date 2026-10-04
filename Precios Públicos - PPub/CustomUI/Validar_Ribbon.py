# -*- coding: utf-8 -*-
"""Valida el Ribbon de PPub_BDatos_2026.xlsm contra el VBA exportado y sus dos tablas.

Adaptado del Validar_Ribbon.py de Jornadas_y_Congresos (2026-10-04). Lee del .xlsm (como zip,
sin abrir Excel; vale con el libro abierto: lee lo ultimo GUARDADO) el customUI14.xml y las
tablas Lo_RibbonUI (Prog__RibbonUI) y Tb_Tareas (Prog__Menu_Aux); el VBA, de VBA_Moduls/.

ERRORES (salida con codigo 1):
  E1  Control con getVisible="GetVsbl_CtrlTab" o getSupertip="getStip_CtrlTab" cuyo tag no
      tiene fila en Lo_RibbonUI (se quedaria oculto / sin supertip). Salvo TAGS_SIN_FILA.
  E2  Tag repetido en Lo_RibbonUI (solo manda la primera fila).
  E3  Callback del XML que no existe en el VBA (RibbonX falla en silencio). Salvo CALLBACKS_PENDIENTES.
  E4  Rutina de Tb_Tareas (Nombre_Rut, partida por " + ") que no existe en el VBA.
  E5  Rutina de Tb_Tareas que tambien es la de un boton (Nom_Rut de Lo_RibbonUI): las rutinas
      con boton no deben verse en Form_Menu.
AVISOS:
  A1  Fila de Tb_Tareas con Uribbon-Tags (desde el 2026-10-04 los botones viven en Lo_RibbonUI).
  A2  Fila de Lo_RibbonUI cuyo tag no esta en el XML.
  A3  Usuario que no es un Usuario_ID de Tb_Usuarios, u hoja de SheetsNames que no existe.
  A4  Nom_Rut de Lo_RibbonUI que no existe en el VBA.
  A5  Tags sin fila a proposito (TAGS_SIN_FILA) y callbacks pendientes (CALLBACKS_PENDIENTES).

Uso:  python CustomUI/Validar_Ribbon.py [ruta_libro.xlsm]
"""
# Last Rev. 2026-10-04 13:00
import glob
import os
import re
import sys
import zipfile

from Volcar_Tareas_y_Ribbon import LIBRO, RAIZ, leer_tabla

VBA_DIR = os.path.join(RAIZ, 'VBA_Moduls')
# Controles que hoy quedan ocultos a proposito por no tener fila (menus contextuales sin uso).
TAGS_SIN_FILA = {'change_usuario', 'liq_titprop_cctxt_helpcomments', 'liq_titprop_helpcomments'}
# Callbacks del XML sin rutina, pendientes de decidir (ver memoria del proyecto, 2026-10-01).
CALLBACKS_PENDIENTES = {'getlbl_cctxtbtnsw_pruebaonoff', 'onact_change_usuario', 'onact_groupsavetimer_usb'}
HOJAS_COMODIN = {'no'}          # elemento de SheetsNames que oculta el control en todas las hojas
ATRIB_CALLBACK = ('onAction', 'getVisible', 'getLabel', 'getSupertip', 'getContent', 'getText',
                  'onChange', 'getPressed', 'getEnabled', 'getImage', 'onLoad', 'getScreentip')


def _lista(txt):
    return [e.strip() for e in re.split(r'[,;]', txt or '') if e.strip()]


def _rutinas(txt):
    return [e.strip() for e in re.split(r'\s\+\s|[,;]', txt or '') if e.strip()]


def main():
    libro = next((a for a in sys.argv[1:] if a.lower().endswith('.xlsm')), LIBRO)
    errores, avisos = [], []

    defs = set()
    for f in glob.glob(os.path.join(VBA_DIR, '*.*')):
        if f.lower().endswith(('.bas', '.cls', '.frm')):
            txt = open(f, 'rb').read().decode('cp1252')
            defs.update(m.lower() for m in re.findall(
                r'^[ \t]*(?:Public |Private )?(?:Sub|Function)[ \t]+(\w+)', txt, re.M))

    with zipfile.ZipFile(libro) as z:
        xml = z.read('customUI/customUI14.xml').decode('utf-8')
        _, _, rib = leer_tabla(z, 'Prog__RibbonUI')
        _, _, tar = leer_tabla(z, 'Prog__Menu_Aux')
        _, _, usu = leer_tabla(z, 'Prog__Usuarios')
        hojas = {h.lower() for h in re.findall(r'<sheet [^>]*name="([^"]+)"', z.read('xl/workbook.xml').decode('utf-8'))}
    xml = re.sub(r'<!--.*?-->', '', xml, flags=re.S)
    controles = []
    for m in re.finditer(r'<(\w+)\b([^>]*)>', xml):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(2)))
        if 'id' in a or 'idMso' in a:
            controles.append(a)
    ids_usuario = {r[1].lower() for r in usu[2:] if r[1]}
    cab_tar = tar[1][1:]                              # cabecera de Tb_Tareas (Tarea, Usuario, ...)
    rib = [r for r in rib[2:] if any(r[1:])]          # [Fila, Tag, Usuario, Sheets, Nom_Rut, Desc, Inf, Grupo]
    tar = [r for r in tar[2:] if any(r[1:])]          # [Fila, Tarea, Usuario, Nombre_Rut, ...]

    # --- E1 / A5: controles sin fila
    tags_rib = {}
    for r in rib:
        t = r[1].strip().lower()
        if t in tags_rib:
            errores.append(f'E2  Tag repetido en Lo_RibbonUI: "{r[1]}" (filas {tags_rib[t]} y {r[0]})')
        else:
            tags_rib[t] = r[0]
    tags_xml = set()
    for a in controles:
        tg = a.get('tag', '')
        if tg:
            tags_xml.add(tg.lower())
        if a.get('getVisible') == 'GetVsbl_CtrlTab' or a.get('getSupertip') == 'getStip_CtrlTab':
            if tg.lower() not in tags_rib:
                if tg.lower() in TAGS_SIN_FILA:
                    avisos.append(f'A5  {a["id"]}: tag "{tg}" sin fila a proposito (queda oculto)')
                else:
                    errores.append(f'E1  {a["id"]}: tag "{tg}" sin fila en Lo_RibbonUI')

    # --- E3: callbacks
    for a in controles:
        for at in ATRIB_CALLBACK:
            cb = a.get(at)
            if cb and cb.lower() not in defs:
                quien = a.get('id') or a.get('idMso')
                if cb.lower() in CALLBACKS_PENDIENTES:
                    avisos.append(f'A5  {quien}: {at}="{cb}" no existe en el VBA (pendiente conocido)')
                else:
                    errores.append(f'E3  {quien}: {at}="{cb}" no existe en el VBA')

    # --- A2 / A3 / A4: filas de Lo_RibbonUI
    ruts_boton = set()
    for r in rib:
        if r[1].strip().lower() not in tags_xml:
            avisos.append(f'A2  Lo_RibbonUI fila {r[0]}: el tag "{r[1]}" no esta en el XML')
        for u in _lista(r[2]):
            if u.lower() not in ids_usuario:
                avisos.append(f'A3  Lo_RibbonUI fila {r[0]} ({r[1]}): "{u}" no es un Usuario_ID')
        for h in _lista(r[3]):
            if h.lower() not in hojas and h.lower() not in HOJAS_COMODIN:
                avisos.append(f'A3  Lo_RibbonUI fila {r[0]} ({r[1]}): la hoja "{h}" no existe')
        for rut in _rutinas(r[4]):
            ruts_boton.add(rut.lower())
            if rut.lower() not in defs:
                avisos.append(f'A4  Lo_RibbonUI fila {r[0]} ({r[1]}): Nom_Rut "{rut}" no existe en el VBA')

    # --- A1 / E4 / E5: filas de Tb_Tareas
    i_tag = cab_tar.index('Uribbon-Tags') + 1 if 'Uribbon-Tags' in cab_tar else None
    for r in tar:
        if i_tag and r[i_tag].strip():
            avisos.append(f'A1  Tb_Tareas fila {r[0]} ("{r[1]}") tiene Uribbon-Tags = "{r[i_tag]}"')
        for rut in _rutinas(r[3]):
            if rut.lower() not in defs:
                errores.append(f'E4  Tb_Tareas fila {r[0]} ("{r[1]}"): la rutina "{rut}" no existe en el VBA')
            if rut.lower() in ruts_boton:
                errores.append(f'E5  Tb_Tareas fila {r[0]} ("{r[1]}"): "{rut}" es la rutina de un boton del Ribbon')

    for linea in errores + avisos:
        print(linea)
    print(f'\n{len(errores)} errores, {len(avisos)} avisos  '
          f'(Lo_RibbonUI: {len(rib)} filas, Tb_Tareas: {len(tar)} filas, controles: {len(controles)})')
    sys.exit(1 if errores else 0)


if __name__ == '__main__':
    main()
