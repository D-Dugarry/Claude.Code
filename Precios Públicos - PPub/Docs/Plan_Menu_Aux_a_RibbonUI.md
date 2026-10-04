# Plan: botones fuera de `_Menú_Aux`, metadatos del Ribbon en `Lo_RibbonUI`

**Proyecto:** `PPub_BDatos_2026.xlsm` · **Fecha:** 2026-10-04 · **Modelo:** Jornadas y Congresos
(fases F3–F4 de agosto de 2026 y cambio de `Form_Running_Rut` del 2026-08-17; ver
`../Jornadas_y_Congresos/docs/Informe_Ribbon.md` e `Informe_Rutinas_Botones_Unicos.md`).

## Estado: HECHO (2026-10-04)

Las seis fases están aplicadas en el libro, probadas por el usuario en Excel y validadas con
`CustomUI/Validar_Ribbon.py` (0 errores). Commits: `fa59c6fa` (fase 0), `318a97b2` (fases 1–4),
`9fe2300a` (volcado tras la fase 1), `fa1ce549` (fase 5) y el de cierre, con los 3 callbacks de los
menús contextuales que faltaban y la documentación. `Tb_Tareas` pasó de 65 a 8 filas; `Lo_RibbonUI`
tiene 26. El módulo de un solo uso `Z_Migrar_RibbonUI` se quitó del libro y del repo.

## Objetivo

1. Quitar de la tabla `Tb_Tareas` (hoja `_Menú_Aux`, CodeName `Prog__Menu_Aux`) las filas que
   corresponden a un botón del Ribbon.
2. Que una rutina con botón no pueda verse ni lanzarse desde `Form_Menu`, y que
   `Form_Running_Rut` deje de buscar rutinas en `_Menú_Aux` y de lanzarlas.

## Cómo queda

- **`Lo_RibbonUI`** (hoja `RibbonUI`, CodeName `Prog__RibbonUI`, importada de JyC) es la única
  fuente de los metadatos del Ribbon: una fila por `tag`, con `Usuario`, `SheetsNames`, `Nom_Rut`,
  `Descripion`, `Informe_Rut` y `Group-Tag`. La lee `M___RibbonUI_Rules` (motor portado de
  `M_003_RibbonUI_Rules` de JyC), con el `tag` comparado entero y sin distinguir mayúsculas.
  Un `tag` sin fila queda oculto; si la tabla no se puede leer, se ve todo.
- **`Tb_Tareas`** se queda solo con las tareas que no tienen botón (8).
- **`Form_Running_Rut`** ya no lanza nada: el botón lo abre con `Rut_Progreso_Abrir`, llama a la
  rutina directamente y lo cierra con `Rut_Progreso_Cerrar` (`M_90_Rutinas_Menu_Aux`).
- **El XML del Ribbon no cambia**: `GetVsbl_CtrlTab` y `getStip_CtrlTab` conservan su nombre.

## Dependencias de `_Menú_Aux` que había que cortar antes de borrar filas

| Para qué | Dónde | Si se borraba la fila |
|---|---|---|
| Visibilidad | `Rut_Filtrar_Tareas` (columna `Visible`, 19 llamadas) + `Func_CtrlTab_View` (por prefijo `Like Tag*`) | el botón desaparecía |
| Supertips | `Func_STip_CtrlTab_Value` | "¡ Control.Tag, NO encontrado !" |
| Lanzar por nombre | `RuT_Ejecutar_Rut` (4 callbacks), `Form_Running_Rut` (9 botones), `Func_Rut_CtrlTab_Value` | "No existe la Tarea" / "la Rutina NO Existe" |
| Informe | `RuT_Load_Task_Data`, también desde `Workbook_Open` | error 13 al abrir el libro |

## Fallos previos encontrados (y qué se hace)

1. **Export e Import BDatos no funcionaban**: los callbacks pedían `Rut_LstObj_Export_Bdatos` /
   `Rut_LstObj_Import_Bdatos`, que no existen (las reales son `Rut_Lo_*_Hist_Bdatos`). → Se
   arreglan en la fase 4.
2. **ReCalculate** ejecutaba `"Rut_Recalcular_Tabla_" & ActiveSheet.Name`: solo se veía en
   `Inf_Recibos_TIO`, la única hoja sin rutina con ese nombre. → Elige la rutina por CodeName y se
   ve en `Inf_Recibos_TIO`, `Inf_RSm` y `JIs_AE4`.
3. ~~Usuarios `Susan`/`Fanny` mal escritos~~ — **falso**: lo leí en un volcado recortado a 10
   caracteres. Las listas dicen `Boss,SusanaC,FannyR;RafaG`, todo IDs correctos.
4. El separador de *Restore* estaba oculto: la fila decía `RestoreBdatos` y `Like` distingue
   mayúsculas. → Desaparece con la comparación nueva.
5. 4 filas de botones que ya no existen en el XML y 18 tareas sin botón que apuntan a rutinas
   inexistentes. → Se borran (quedan en el historial de `CustomUI/Tb_Tareas_snapshot.csv`).

## Decisiones (2026-10-04, el usuario: "de acuerdo con todo")

1. Se borran también las filas sin tag cuya rutina tiene botón (grupo C): Activar Modo
   Programación, Copia USB, Show RibbonX, Hide/Restore Context Menu, Protect/UnProtect,
   SW-Probando, Sheets Show all y JI's ActivEco4.
2. El informe de la última ejecución se sigue guardando, ahora en `Lo_RibbonUI.Informe_Rut`
   (lo muestra el supertip de Boss).
3. Sin efecto (ver fallo 3).
4. Se arreglan los fallos 1 y 2.
5. Se borran las 18 tareas muertas.

## Fases

Cada tanda termina con: reimportar, compilar (Depuración → Compilar), probar en Excel y commit.

- **Fase 0. Red de seguridad.** `CustomUI/Volcar_Tareas_y_Ribbon.py` (volcados CSV de las dos
  tablas) y `CustomUI/Validar_Ribbon.py` (cruce XML ↔ VBA ↔ tablas, adaptado de JyC).
- **Fase 1. Rellenar `Lo_RibbonUI`.** Macro de un solo uso `Z_Fase1_Rellenar_RibbonUI`
  (módulo `Z_Migrar_RibbonUI`): borra las filas de JyC y crea las 26 de PPub, copiando de
  `Tb_Tareas` la descripción y el informe.
- **Fase 2. Motor de reglas.** `M___RibbonUI_Rules.bas` + constantes `Rib_*`;
  `GetVsbl_CtrlTab`/`getStip_CtrlTab` leen de él; `RefreshRibbon` llama a `Rules_Reset`.
- **Fase 3. Callbacks directos.** Cambiar usuario, vista del Ribbon, Ribbon Refresh, Buscar
  vínculos, Modo Programación y Reset dejan de pasar por `_Menú_Aux`.
- **Fase 4. `Form_Running_Rut`.** `Rut_Progreso_Abrir`/`Rut_Progreso_Cerrar`; fuera los
  `Call_RuT_*` y el camino viejo del formulario.
- **Fase 5. Podar** (segunda tanda, tras probar la primera): borrar las 57 filas, quitar
  `Rut_Filtrar_Tareas` y sus llamadas, la columna `Visible` y `Task_Visible`; `Form_Menu` y el
  menú dinámico "9_ " pasan a `Fnc_Lista_Contiene`; borrar las rutinas que queden sin llamador.
- **Fase 6. Documentar**: `CLAUDE.md`, memoria y volcados finales.
