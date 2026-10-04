# Plan: botones fuera de `_Menú_Aux`, metadatos del Ribbon en `Lo_RibbonUI`

**Proyecto:** `PPub_BDatos_2026.xlsm` · **Fecha:** 2026-10-04 · **Revisado:** 2026-10-04 17:15 ·
**Modelo:** Jornadas y Congresos (fases F3–F4 de agosto de 2026 y cambio de `Form_Running_Rut`
del 2026-08-17; ver `../Jornadas_y_Congresos/docs/Informe_Ribbon.md` e
`Informe_Rutinas_Botones_Unicos.md`).

**Etiquetas:** ✅ **Hecho** · ⏳ **Pendiente** · ⏸️ **Postergado**

## Estado: ✅ HECHO (2026-10-04)

Las seis fases están aplicadas en el libro, probadas por el usuario en Excel y validadas con
`CustomUI/Validar_Ribbon.py` (0 errores). Commits: `fa59c6fa` (fase 0), `318a97b2` (fases 1–4),
`9fe2300a` (volcado tras la fase 1), `fa1ce549` (fase 5), `72a673aa` (cierre: 3 callbacks de los
menús contextuales que faltaban y documentación) y `b32f24af` (borrada
`Rut_Right_Click_Control_KK`). `Tb_Tareas` pasó de 65 a 8 filas; `Lo_RibbonUI` tiene 26. El módulo
de un solo uso `Z_Migrar_RibbonUI` se quitó del libro y del repo.

**Comprobación de la revisión (16:35):** el libro guardado a las 16:25 tiene el mismo código que
`VBA_Moduls/` en sus 113 módulos (leído con olevba). Las únicas diferencias son mayúsculas que el
editor de VBA cambia por su cuenta, como `TBx_Informe` → `Tbx_Informe` o `allowFiltering` →
`AllowFiltering`. `Validar_Ribbon.py` da 0 errores y 4 avisos, los cuatro intencionados. Los
volcados de las dos tablas siguen al día: al regenerarlos solo cambia la hora de un informe.

### Resumen

| Qué | Estado |
|---|---|
| Objetivos 1 y 2 | ✅ Hecho |
| Fases 0 a 6 | ✅ Hecho |
| Fallos previos 1, 2, 4 y 5 (el 3 resultó no ser un fallo) | ✅ Hecho |
| Decisiones 1 a 5 | ✅ Hecho |
| 3 callbacks de los menús contextuales | ✅ Hecho |
| Reimportar `M_000_Ini_APP` en el libro | ✅ Hecho |
| Borrar el temporal de Excel `A1C24500` | ✅ Hecho |
| Confirmar que el libro compila tras la última reimportación | ✅ Hecho (el usuario, 2026-10-04) |
| Revisar lo que sigue vigente de `Informes/Informe_RibbonX.html` | ⏸️ Postergado |
| `M_520` (Recalcular JIs 303): restaurar sus hojas o borrar el módulo | ⏸️ Postergado |

## Objetivo

1. ✅ **Hecho** — Quitar de la tabla `Tb_Tareas` (hoja `_Menú_Aux`, CodeName `Prog__Menu_Aux`)
   las filas que corresponden a un botón del Ribbon.
2. ✅ **Hecho** — Que una rutina con botón no pueda verse ni lanzarse desde `Form_Menu`, y que
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

| Para qué | Dónde | Si se borraba la fila | Estado |
|---|---|---|---|
| Visibilidad | `Rut_Filtrar_Tareas` (columna `Visible`, 19 llamadas) + `Func_CtrlTab_View` (por prefijo `Like Tag*`) | el botón desaparecía | ✅ Hecho: fases 2 y 5 |
| Supertips | `Func_STip_CtrlTab_Value` | "¡ Control.Tag, NO encontrado !" | ✅ Hecho: fase 2 (ahora lee `Lo_RibbonUI`) |
| Lanzar por nombre | `RuT_Ejecutar_Rut` (4 callbacks), `Form_Running_Rut` (9 botones), `Func_Rut_CtrlTab_Value` | "No existe la Tarea" / "la Rutina NO Existe" | ✅ Hecho: fases 3 y 4 |
| Informe | `RuT_Load_Task_Data`, también desde `Workbook_Open` | error 13 al abrir el libro | ✅ Hecho: fase 4 (ahora `Rut_RibbonUI_Guardar_Informe`) |

## Fallos previos encontrados (y qué se hizo)

1. ✅ **Hecho** (fase 4) — **Export e Import BDatos no funcionaban**: los callbacks pedían
   `Rut_LstObj_Export_Bdatos` / `Rut_LstObj_Import_Bdatos`, que no existen (las reales son
   `Rut_Lo_*_Hist_Bdatos`).
2. ✅ **Hecho** (fase 4) — **ReCalculate** ejecutaba `"Rut_Recalcular_Tabla_" & ActiveSheet.Name`:
   solo se veía en `Inf_Recibos_TIO`, la única hoja sin rutina con ese nombre. Ahora elige la
   rutina por CodeName y se ve en `Inf_Recibos_TIO`, `Inf_RSm` y `JIs_AE4`.
3. ✅ **Hecho** (descartado) — ~~Usuarios `Susan`/`Fanny` mal escritos~~: **no era un fallo**. Lo
   leí en un volcado recortado a 10 caracteres. Las listas dicen `Boss,SusanaC,FannyR;RafaG`, todo
   IDs correctos.
4. ✅ **Hecho** — El separador de *Restore* estaba oculto: la fila decía `RestoreBdatos` y `Like`
   distingue mayúsculas. Desapareció con la comparación nueva.
5. ✅ **Hecho** (fase 5) — 4 filas de botones que ya no existen en el XML y 18 tareas sin botón
   que apuntaban a rutinas inexistentes: borradas (quedan en el historial de
   `CustomUI/Tb_Tareas_snapshot.csv`).

## Decisiones (2026-10-04, el usuario: "de acuerdo con todo")

1. ✅ **Hecho** — Se borran también las filas sin tag cuya rutina tiene botón (grupo C): Activar
   Modo Programación, Copia USB, Show RibbonX, Hide/Restore Context Menu, Protect/UnProtect,
   SW-Probando, Sheets Show all y JI's ActivEco4.
2. ✅ **Hecho** — El informe de la última ejecución se sigue guardando, ahora en
   `Lo_RibbonUI.Informe_Rut` (lo muestra el supertip de Boss).
3. ✅ **Hecho** — Sin efecto (ver fallo 3).
4. ✅ **Hecho** — Se arreglan los fallos 1 y 2.
5. ✅ **Hecho** — Se borran las 18 tareas muertas.

## Fases

Cada tanda termina con: reimportar, compilar (Depuración → Compilar), probar en Excel y commit.

- ✅ **Hecho** (`fa59c6fa`) — **Fase 0. Red de seguridad.** `CustomUI/Volcar_Tareas_y_Ribbon.py`
  (volcados CSV de las dos tablas) y `CustomUI/Validar_Ribbon.py` (cruce XML ↔ VBA ↔ tablas,
  adaptado de JyC).
- ✅ **Hecho** (`318a97b2`, volcado en `9fe2300a`) — **Fase 1. Rellenar `Lo_RibbonUI`.** Macro de
  un solo uso `Z_Fase1_Rellenar_RibbonUI` (módulo `Z_Migrar_RibbonUI`): borra las filas de JyC y
  crea las 26 de PPub, copiando de `Tb_Tareas` la descripción y el informe. El módulo ya se quitó.
- ✅ **Hecho** (`318a97b2`) — **Fase 2. Motor de reglas.** `M___RibbonUI_Rules.bas` + constantes
  `Rib_*`; `GetVsbl_CtrlTab`/`getStip_CtrlTab` leen de él; `RefreshRibbon` llama a `Rules_Reset`.
- ✅ **Hecho** (`318a97b2`) — **Fase 3. Callbacks directos.** Cambiar usuario, vista del Ribbon,
  Ribbon Refresh, Buscar vínculos, Modo Programación y Reset dejan de pasar por `_Menú_Aux`.
- ✅ **Hecho** (`318a97b2`) — **Fase 4. `Form_Running_Rut`.** `Rut_Progreso_Abrir`/
  `Rut_Progreso_Cerrar`; fuera los `Call_RuT_*` y el camino viejo del formulario.
- ✅ **Hecho** (`fa1ce549`) — **Fase 5. Podar** (segunda tanda, tras probar la primera): 57 filas
  borradas (65 → 8), fuera `Rut_Filtrar_Tareas` y sus llamadas, la columna `Visible` y
  `Task_Visible`; `Form_Menu` y el menú dinámico "9_ " pasan a `Fnc_Lista_Contiene`. Rutinas
  borradas por quedar sin llamador: `Rut_Chg_Usuario`, `Rut_OnOff_SW_WB_Deactivate`,
  `Rut_RibbonX_ShowAll`, `Rut_OnOff_SW_Probando` y `Rut_ProtectUnProtect_ActivSheet`. Se quedan
  `Rut_RibbonRefresh` (la llama `M___RibbonUI`) y `Rut_Btn_Menu_Aux` (la llaman dos formas de
  hoja).
- ✅ **Hecho** (`72a673aa`) — **Fase 6. Documentar**: `CLAUDE.md`, memoria y volcados finales.

## Después del cierre

- ✅ **Hecho** (`72a673aa`) — **3 callbacks de los menús contextuales** que el XML pedía y no
  existían: `GetLbl_CCtxtBtnSW_PruebaONOFF`, `OnAct_Change_Usuario` y `OnAct_GroupSaveTimer_USB`.
  El clic derecho daba "No se puede ejecutar la macro" y el menú salía sin ninguno de los
  elementos propios. Probado: los dos menús (tabla y celda) salen completos.
- ✅ **Hecho** (`b32f24af`) — Borrada `Rut_Right_Click_Control_KK` (`M_000_Ini_APP`), copia sin
  uso de la ya borrada `Rut_OnOff_SW_WB_Deactivate`.
- ✅ **Hecho** — `M_000_Ini_APP` reimportado en el libro (comprobado con olevba en el libro
  guardado a las 16:25).
- ✅ **Hecho** — Borrado el temporal de Excel `A1C24500` (108 MB) de la carpeta del proyecto.
- ✅ **Hecho** (decisión) — Avisos intencionados de `Validar_Ribbon.py`: los tags
  `Change_Usuario`, `Liq_TitProp_CCtxt_HelpComments` y `Liq_TitProp_HelpComments` no tienen fila
  en `Lo_RibbonUI` a propósito (quedan ocultos). `RunRutPrueba` apunta a `RuT_kkkk`, que no
  existe; es el botón de pruebas y su `Nom_Rut` se cambia según lo que se quiera probar.
- ✅ **Hecho** — El libro compila (Depuración → Compilar) tras la última reimportación:
  confirmado por el usuario el 2026-10-04.

## Relacionado, fuera de este plan

- ⏸️ **Postergado** — **`Informes/Informe_RibbonX.html`**, sin revisar entero. Este plan ya
  resolvió parte: R3, R4 y R5 con los 3 callbacks contextuales; R6 el 2026-09-30; y las mejoras
  sobre `Func_CtrlTab_View`, `Rut_Filtrar_Tareas` y la visibilidad duplicada dejaron de aplicar con
  `Lo_RibbonUI`. Falta ver si siguen vigentes R1, R2 y R7 a R10, y el resto de las mejoras. Del
  patrón de reintento sin límite (`Resume` tras el error -2147417848) queda una sola ocurrencia en
  `M___RibbonUI`, en `OnAct_MenuAux`; las demás desaparecieron al reescribir los callbacks.
- ⏸️ **Postergado** — **`M_520_Calcular_JIs_PPub_1303`**: su fila de `Tb_Tareas` ("Recalcular JIs
  303") se borró en la fase 5 como tarea muerta. El módulo sigue comentado entero y sin llamador,
  porque las hojas `Sht__BD_AdmP` y `Sht__BD_JIs_303` no existen. Hay que decidir si se restauran
  las hojas o se borra el módulo.
