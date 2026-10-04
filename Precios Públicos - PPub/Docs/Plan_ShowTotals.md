# Plan: control de la fila de totales (`ShowTotals`) de las tablas

**Proyecto:** `PPub_BDatos_2026.xlsm` · **Fecha:** 2026-10-04 17:40 · **Revisado:** 2026-10-04 22:25 ·
**Modelos:** Jornadas y Congresos (patrón de escritura en tablas con fila de totales, nacido del
fallo de `Lo_TPV` del 2026-08-26; ver `../Jornadas_y_Congresos/CLAUDE.md`, apartado «Escribir en
un `ListObject` que tiene FILA DE TOTALES») y Enseñanzas Propias (`Rut_Lo_WrkSht_Preparar` oculta
los totales, commit `1f9cb9cc` del 2026-09-27; `Fnc_Lo_Contar_Visibles`, commit `f79d4c4d` del
2026-09-28).

**Etiquetas:** ✅ **Hecho** · 🧪 **Aplicado, falta probarlo en Excel** · ⏳ **Pendiente** · ⏸️ **Postergado**

## Estado: ✅ FASES 1 A 5 HECHAS Y VERIFICADAS (commit del 2026-10-04 23:35); fase 6 postergada

Pruebas 5 a 7 de «Cómo probarlo» no hechas: el usuario las dio por opcionales (cambios menores, ya
cubiertos por la doble ejecución). `M_310` ya no vuelve a mostrar los totales de `Tb_INSS`.

Están escritas y verificadas en el texto (bytes cp1252/CRLF, mismo balance de bloques que en `HEAD`
en los 28 módulos), reimportadas en el libro (olevba: libro = repo, `M_310` incluido) y con la
**doble ejecución superada** (libro de las 23:04: `M_110` Sí/Sí/Sí/No a las 22:56 y `M_210` a las
23:02, contra `PPub_BDatos_2026 - Antes_ShowTotals.xlsm`). `BDatos`, `BD_Dupl`, `BD_ErrDate`,
`BD_ImpAdm` y `BD_INSS` salen **idénticas** (valores, rellenos y formatos), y los logs también,
salvo el bloque de `M_215` en el de `M_110`: en ANTES, `M_110` se ejecutó antes de reimportar
`BD_ImpAdm` (1.463 filas) y aquí después (1.584); el resultado final es el mismo. La política de la
fase 4 actuó: `Tb_DefCols172146` pasó a tener totales y `Tb_DefCols_BD` los conserva tras `M_110`.
Faltan las pruebas 5 a 7 de «Cómo probarlo» (`M_195`, botón de columnas, cancelar `M_110`).

| Fase | Qué | ¿Cambia el resultado? | Estado |
|---|---|---|---|
| 0 | Cerrar la tanda pendiente del 2026-10-04 | — | ✅ Hecho (otra sesión: `68c48d66` y siguientes, doble ejecución superada) |
| 1 | `M_195` (error 91), el «flag» muerto de las `DefCol` y su único `ShowTotals = False` | Sí: `M_195` deja de fallar | ✅ Hecho |
| 2 | Escrituras y copias que se protegen solas (patrón de JyC) | No | ✅ Hecho |
| 3 | `Fnc_Lo_Contar_Visibles` en los 57 recuentos | Solo en `M_115` sin recibos del curso Pos | ✅ Hecho |
| 4 | Política de cierre: 10 tablas siempre con totales | Sí: `DefCol` e informe siempre con totales | ✅ Hecho |
| 5 | `Rut_Lo_WrkSht_Preparar` oculta los totales (pieza de EP) y limpieza | No | ✅ Hecho |
| 6 | Estudio: estado de la fila de totales configurable para las 39 tablas | — | ⏸️ Postergado (plan propio) |
| — | Exportación a `.xlsx` con la fila de totales (`Rut_Lo_Export_XlsX`) | — | ⏸️ Postergado |

### Decisiones del usuario (2026-10-04)

| # | Decisión | Respuesta |
|---|---|---|
| 1 | Qué tablas quedan siempre con totales | ✅ **Solo `BDatos`, `Inf_Recibos_TIO` y todas las `DefCol`** («los únicos totales visibles que se requieren por ahora»). Detalle en la fase 4 |
| 2 | ¿Quitar los `ShowTotals = True` del final de las rutinas? | ✅ **De momento no.** Antes, estudiar cómo configurar el estado de la fila de totales para todas las tablas (fase 6) |
| 3 | ¿Se hace la fase 5? | ✅ **Sí** |

## Qué hay en cada proyecto

**Jornadas y Congresos.** No hay una rutina central: es una **convención**, documentada en su
`CLAUDE.md` y aplicada en 9 sitios. Al escribir en una tabla con fila de totales:

```vba
TotalesVisibles = Lo.ShowTotals
Lo.ShowTotals = False
If Not Lo.DataBodyRange Is Nothing Then Lo.DataBodyRange.Delete
Lo.HeaderRowRange.Offset(1, 0).Resize(N, N_Cols).Value = Datos
Lo.Resize Lo.HeaderRowRange.Resize(N + 1, N_Cols)
Lo.ShowTotals = TotalesVisibles
```

Se adoptó tras el 2026-08-26: escribir con la fila de totales visible la pisó y dejó 386 pagos de
`Lo_TPV` fuera de la tabla. Los recuentos de filas visibles compensan el estado con
`Count - 1 + .ShowTotals` (`True` vale -1).

**Enseñanzas Propias** (mismo framework `Rut_Lo` que PPub). Dos piezas: `Rut_Lo_WrkSht_Preparar`
oculta la fila de totales al preparar la hoja, y `Fnc_Lo_Contar_Visibles(Lo, Col)` cuenta sobre
`DataBodyRange` y devuelve 0 si `SpecialCells` falla.

**PPub antes de este plan.**

- Unas 30 rutinas tocaban `ShowTotals`. Casi todas los apagaban al empezar, y las de proceso los
  encendían con `= True` al terminar: forzaban el valor, no restauraban el que había.
- Guardar y restaurar el estado solo lo hacían 2 rutinas de `Rut_Lo`.
- `Rut_Lo_WrkSht_Preparar` no tocaba los totales.
- 57 recuentos en 16 módulos usaban
  `.Range.Columns(X).SpecialCells(xlCellTypeVisible).Cells.Count - 1`, que solo acierta con los
  totales ocultos. `M_130` restaba 2 porque allí los tiene visibles.

### Lo que hace distinto a PPub: la fila de totales alimenta los paneles resumen

En PPub la fila de totales no es estética. Las celdas de encima de varias tablas son fórmulas que
la leen, y con los totales ocultos dan `#REF!`:

| Hoja | Tabla | Celdas que leen `[#Totals]` |
|---|---|---|
| `BDatos` | `Tb_BDatos` | 17 (`ACont_Emi`, `Ref`, `Imp_Rec`, `Rec_Imp_*`, `Emitido`...) |
| `BD_Ant` | `Tb_BD_Ant` | 9 |
| `BD_ImpAdm_2025-26` | `Tabla21` | 5 |
| `BD_INSS` | `Tb_INSS` | 2 |
| `Inf_Recibos_TIO` | `Tb_JIs_Tasas425762` | 1 |
| `BD_AE4x4` | `Tb_AE4x4` | 1 |

Además, las hojas `_DefCol_*` generan las líneas `Public Const` con
`Tb_DefCols_BD37[[#Totals],[LongNombre]]` (6 hojas), `Tb_DefCols_BD[[#Totals],[LongNombre]]` (3
hojas) y `Tb_Menu_Aux`/`Tabla22` (`_DefCol_BD_Aux`).

Por eso a PPub no le basta el «restaurar el estado previo» de JyC: si una tabla estaba oculta, la
deja oculta y su panel sigue en `#REF!`. Las rutinas compartidas restauran, porque no les toca
imponer nada; pero al terminar cada tarea se aplica una **política**: las tablas de la lista quedan
con los totales visibles. Es lo que ya hace `Rut_On_Functions`, que aplica la política de la app y
no el estado previo.

## Hallazgos (2026-10-04)

1. **`M_195` daba error 91 si `Tb_DefCols_BD` tenía la fila de totales oculta.** Escribía
   `Lo_DefCol_BD.TotalsRowRange(DefC_HiddenCol) = False`. `M_110` ocultaba esa fila en cada
   ejecución, y solo la volvían a mostrar `M_210`, `M_90` o el botón de mostrar/ocultar columnas
   (en el libro de las 17:24 estaba oculta; en el de las 20:36, visible). Ese «flag» se escribía en
   6 sitios y **no lo leía nadie**: el botón decide con el switch `Sw_Col_Hide_<CodeName>`, y
   ninguna fórmula lo usa. Es el mismo estado muerto que JyC retiró el 2026-08-11. → Fase 1.
2. **`M_115` contaba con los totales visibles.** `M_114` termina con `Lo_BD.ShowTotals = True` y
   `M_110` llama a `M_115` justo después: el recuento de su línea 47 salía 1 de más. Sin recibos
   del curso Pos habría dado 1, y el `ClearContents` sobre `SpecialCells`, error 1004. → Fase 3.
3. **El fallo de `Lo_TPV`, latente en las importaciones.** `Rut_Lo_Import_WorkSheet` y
   `Rut_Lo_Import_LoData_LoDefCol` vacían la tabla y pegan con
   `Lo_Data.Range.Offset(1, 0).PasteSpecial`. Con la fila de totales visible pegarían encima de
   ella. No pasaba solo porque los 5 llamadores apagaban antes los totales.
   `Rut_Lo_DataBodyRange_Filtered_Copy` los apagaba y no los restauraba. → Fase 2.
4. **Paneles en `#REF!` tras algunas ejecuciones.** `M_110` apaga los totales de `Tb_BD_Ant` y no
   los vuelve a encender; al cancelar el diálogo de fichero en `M_110`, `M_195`, `M_210` o `M_310`
   el `GoTo Restablecer_Valores` salta el `ShowTotals = True`; `M_180` los apaga y no los restaura.
   → Fase 4 para `BDatos`. En `BD_Ant`, `BD_ImpAdm`, `BD_INSS` y `BD_AE4x4` no se corrige: según la
   decisión 1, sus totales no se requieren por ahora.
5. **`M_130` necesitaba encontrar los totales visibles al entrar** (línea 60:
   `Lo_Inf.TotalsRowRange.Row`, error 91 si una ejecución anterior se cortó con la fila oculta).
   → Fase 2.
6. **`M_130` copiaba la leyenda de tipos de recibo con la fila de totales incluida, si la hubiera**
   (`ListColumns(1).Range` abarca también la fila de totales de `Tb_TipoRec`). → Fase 2.
7. **Exportar a `.xlsx`.** `Rut_Lo_Export_XlsX` (línea 53) copia `Lo_Data.Range`, con la fila de
   totales como una fila de valores más si se ve. JyC lo resolvió en `M_825`. → Postergado.
8. **`M_510` falla al empezar, sin relación con este plan.** `Set Lo_PlanAE4 =
   Sht__BD_JIs_AE4.ListObjects(1)`, pero esa hoja no tiene ninguna tabla en el libro (comprobado en
   el `.xlsm`). No se ha tocado; queda como pendiente aparte.

## Fases

### Fase 0: prerrequisito ✅

La hizo otra sesión el 2026-10-04: tanda commiteada (`68c48d66`, `1e155491`, `cf557e59`,
`4114938b`) y pusheada, con la doble ejecución de `M_110` + `M_210` superada. El libro de las 20:36
coincide con el repo.

### Fase 1: `M_195`, el flag muerto y el `ShowTotals = False` de las `DefCol` ✅

- Quitadas las escrituras a `TotalsRowRange(DefC_HiddenCol)` (también las ya comentadas): `M_195`
  (2), `M_210` (2), `M_90_Rutinas_Menu_Aux` (1), `M_110` (1, comentada), `M_114` (1, comentada) y
  `Rut_Lo_Columns_Show_Hide` (la del `Reset` y el bloque que alternaba el flag).
- Quitado `Lo_DefCol_BD.ShowTotals = False` de `M_110`: era el único sitio que ocultaba los
  totales de una `DefCol`.
- Se quedan los `ShowTotals = True` de las `DefCol` (`M_210`, `M_90`, `Rut_Lo_Columns_Show_Hide`).
- El código que lee las `DefCol` usa `DataBodyRange` y `ListRows.Count`: no le afecta que la fila
  de totales esté visible.

### Fase 2: escrituras y copias que se protegen solas (patrón de JyC) ✅

Dos ayudantes nuevos en `Rut_Lo`: `Fnc_Lo_Totales_Ocultar(Lo)` (oculta y devuelve cómo estaba) y
`Rut_Lo_Totales_Restaurar(Lo, Visibles)`. Los dos **solo cambian `ShowTotals` si hace falta**:
tras `M_110`/`M_210`, cambiarlo en `Tb_INSS` da -2147417848 (fallo abierto, ver `CLAUDE.md`).

- `Rut_Lo_Import_WorkSheet` y `Rut_Lo_Import_LoData_LoDefCol`: ocultan justo antes del
  `DataBodyRange.Delete` y restauran en sus 5 salidas posteriores.
- `Rut_Lo_DataBodyRange_Filtered_Copy`: restaura al final y en la salida temprana.
- `Rut_Lo_DataBodyRange_Filter_y_DEL` y `..._Filter_x2Crit_Copy_ColSource_to_ColTarget`: ya
  guardaban y restauraban; ahora con los ayudantes (sin cambiar el valor si no hace falta), y sin
  el segundo `.ShowTotals = False` redundante de la primera.
- `M_130`, línea 60: `UltFila` se calcula con `Lo_Inf.Range`, sin depender de que se vea la fila
  de totales.
- `M_130`, línea 438: copia cabecera y datos de `Tb_TipoRec` (`Resize(TRows_LoTipoRec + 1, 2)`).
- `M_180` no se toca aquí: ya oculta antes de pegar (con `Preparar`, fase 5), y lo que le faltaba,
  volver a mostrar los totales de `BDatos`, lo hace la fase 4.
- `M_311` y `M_411` escriben en su tabla con `Rut_Lo_DataBodyRange_Filtered_Copy`, así que quedan
  cubiertas.

### Fase 3: recuentos independientes de los totales ✅

- Nueva `Fnc_Lo_Contar_Visibles(Lo, Col)` en `Rut_Lo`, **adaptada**, no copiada de EP: cuenta sobre
  `.Range` y descuenta la cabecera y, si se ve, la fila de totales (la fórmula de JyC). La de EP
  cuenta sobre `DataBodyRange` y devuelve 0 si `SpecialCells` falla. En PPub hay un fallo de Excel
  documentado que hace fallar a `SpecialCells` en todo Excel (`Tb_INSS`, ver `CLAUDE.md`). Con 0,
  `M_311` no borraría los recibos del curso y luego copiaría los nuevos, y quedarían duplicados sin
  avisar. Así, si `SpecialCells` falla, el error salta igual que antes.
- Tabla sin filas: 0 (el patrón viejo contaba 1, la fila de inserción).
- Sustituidos los 57 recuentos: `M_111` 4, `M_112` 5, `M_113` 18, `M_114` 10, `M_115` 2, `M_130` 1
  (el que restaba 2), `M_132` 1, `M_133` 1, `M_211` 2, `M_215` 2, `M_311` 3, `M_315` 1, `M_411` 3,
  `M_413` 1, `M_415` 1 y `Rut_Lo` 2.

### Fase 4: política de cierre, 10 tablas siempre con totales ✅

Nueva `Rut_Lo_Totales_Mostrar` (en `Rut_Lo`): muestra la fila de totales de las tablas de la lista
que la tengan oculta.

**Qué tablas entran (decisión 1).** El libro tiene 39 tablas; 24 tienen fila de totales
configurada y 15 no.

*A. Siempre visibles: entran en la política (10 tablas).*

| Tabla | Hoja (CodeName) | Por qué |
|---|---|---|
| `Tb_BDatos` | `BDatos` (`Sht__BD`) | Panel de 17 celdas |
| `Tb_JIs_Tasas425762` | `Inf_Recibos_TIO` (`Sht__Inf_Recibos_TIO`) | Informe de `M_130`/`M_132`/`M_133`; 1 celda la lee |
| `Tb_DefCols_BD` | `_DefCol_BDatos` (`Prog_DefCol_BD`) | `DefCol`; `LongNombre` de 3 generadores |
| `Tb_DefCols_BD37` | `_DefCol_INSS` (`Prog_DefCol_LSace06`) | `DefCol`; `LongNombre` de 6 generadores |
| `Tb_Menu_Aux` | `_DefCol_BD_Aux` (`Prog_DefCol_Aux`) | `DefCol`; generador de su hoja |
| `Tabla22` | `_DefCol_BD_Aux` (`Prog_DefCol_Aux`) | `DefCol`; generador de su hoja |
| `Tabla6` | `_DefCol_Inf_Recibos` (`Prog_DefCol_Inf_Recibos`) | `DefCol` |
| `Tb_DefCols172146` | `_DefCol_Inf_Peter` (`Prog_DefCol_Inf_Peter`) | `DefCol` (hoy oculta; la fila de debajo está libre) |
| `Tb_DefCols17214658` | `_DefCol_Inf_PeterINSS` (`Prog_DefCol_Inf_PeterINSS`) | `DefCol` |
| `Tabla10` | `_DefCol_JIs_AE4` (`Prog_DefCol_JIs_AE4`) | `DefCol` |

- Las dos de datos van por el CodeName de su hoja (constante `Lo_Totales_Hojas` de `Rut_Lo`). Las
  `DefCol` van por **regla**: todas las tablas de las hojas cuyo CodeName empieza por
  `Prog_DefCol_`, así que entra sola cualquier `DefCol` nueva.
- La decisión vive en un único sitio, `Fnc_Lo_Totales_Siempre_Visibles`, para que la fase 6 solo
  cambie de dónde la lee.

*B. Sin gestionar: se quedan como las dejen sus rutinas (14 tablas).*

| Tabla | Hoja (CodeName) | Nota |
|---|---|---|
| `Tb_BD_Ant` | `BD_Ant` (`Sht__BD_Ant`) | Panel de 9 celdas: en `#REF!` mientras la fila esté oculta (decisión 1) |
| `Tabla21` | `BD_ImpAdm_2025-26` (`Sht__BD_IAdm_CAcadAnt`) | Panel de 5 celdas; `M_210`/`M_215` la dejan visible al terminar |
| `Tb_INSS` | `BD_INSS` (`Sht__BD_INSS`) | Panel de 2 celdas; es la del fallo -2147417848. Desde el 2026-10-04 22:51 `M_310` ya no vuelve a mostrar sus totales al terminar (decisión del usuario), así que, una vez ocultos, nadie los vuelve a cambiar |
| `Tb_AE4x4` | `BD_AE4x4` (`Sht__BD_AE4x4`) | Panel de 1 celda; `M_410`/`M_415` la dejan visible al terminar |
| `Tb_BD_Duplic`, `Tb_BD_Err`, `Tb_BD_Err51` | `BD_Dupl`, `BD_ErrDate`, `BD_RegAnul` | Reciben registros del proceso; ninguna fórmula lee su fila de totales |
| `Tb_JIs_Tasas4257625`, `Tb_JIs_Tasas425762617`, `Tabla5` | `Inf_RecibosJIs`, `Inf_Felipe`, `Inf_Rec_Mov` | Informes que no toca ningún código |
| `Tb_JIs_Tasas4257623` | `Inf_Recibos_TIO y EP` (`Hoja1`) | Hoja residual |
| `Tb_Conceptos`, `Tb_Clasif_Eco` | catálogos | El código solo lee su `DataBodyRange` |
| `Tb_TipoRec` | `Tb_TiposRec` (`Prog_TipoRec`) | Desde la fase 2, `M_130` ya no copia su fila de totales aunque se vea |

*C. Fuera del alcance: sin fila de totales configurada (15 tablas).* `Tb_APP`, `Tb_Usuarios`,
`Tb_Tareas`, `Lo_RibbonUI`, `Lo_SwitchsAPP`, `Tabla94` (`Sheet_Buffer`), `Tb_Sheets_List`,
`Tabla13`, `Tb_Bancos`, `Tabla42`, `Table1`, `Tb_Tipo_Ensenanzas`, `Tabla252`, `Tab_Organicas` y
`Lo_Num_Formulas`.

**Cómo funciona.**

- En hoja protegida desprotege con `Rut_Prot_Save` y reprotege con `Rut_Prot_Restore`, con los
  mismos permisos.
- Cada tabla con su propio control de errores: un fallo se anota en Inmediato
  (`!!! Rut_Lo_Totales_Mostrar: ...`) y sigue con las demás.
- Se llama en los cuatro sitios de la red de seguridad de `Rut_Reset_State`:
  `Rut_Progreso_Cerrar` (`M_90`), el final de tarea de `Form_Menu`, el del menú dinámico
  (`OnAction_Dynamic_Task`, `M___RibbonUI`) y `RuT_Al_Abrir_WorkBook`, donde sustituye el
  `ShowTotals = True` de `BDatos`.
- Los `ShowTotals = True` del final de cada rutina se quedan (decisión 2), salvo el de `Tb_INSS` en
  `M_310`, que se quitó a petición del usuario (22:51).

### Fase 5: `Rut_Lo_WrkSht_Preparar` oculta los totales (pieza de EP) ✅

- `Rut_Lo_WrkSht_Preparar` oculta la fila de totales, **solo si se ve**. Nunca se llama sobre las
  `DefCol`.
- Quitadas 28 líneas `ShowTotals = False` redundantes (misma tabla, `Preparar` en la misma rutina
  y nada entre medias que la vuelva a mostrar): `M_110` 4, `M_112` 1, `M_114` 1, `M_130` 2, `M_132`
  3, `M_133` 3, `M_180` 1, `M_195` 2, `M_210` 1, `M_215` 2, `M_310` 1, `M_410` 1, `M_415` 1, `M_602`
  4 y `M_90` 1. En `M_215`, `M_415` y `M_602` era el final de una línea con varias sentencias.
- Se quedan los de las rutinas que reciben la tabla ya preparada (`M_111`, `M_113`, `M_115`,
  `M_211`, `M_311`, `M_411`, `M_413`), el de `M_315` (no llama a `Preparar`), los de la tabla del
  libro importado (`Rut_Lo_Import`), y los de `M_118` (su llamada está comentada) y `M_510`
  (hallazgo 8), que no se han tocado.

### Fase 6: estudio, estado de la fila de totales configurable para las 39 tablas ⏸️

Pedido por el usuario al decidir el punto 2. Tendrá su propio plan; aquí van los puntos que
debería resolver.

**Objetivo.** Un único sitio que diga, para cada tabla del libro, cómo debe quedar su fila de
totales fuera de los procesos: **Visible**, **Oculta** o **Libre** (no se toca). Con eso la lista
de la fase 4 pasa a ser datos, y se pueden quitar los `ShowTotals = True` del final de las rutinas
(decisión 2).

**Puntos a estudiar.**

1. **Dónde.** Una tabla nueva, p. ej. `Lo_Totales` en una hoja `Prog_*` (como `Lo_SwitchsAPP` o
   `Lo_RibbonUI`), con una fila por tabla: `Tabla`, `Hoja` (CodeName), `Totales`
   (Visible/Oculta/Libre) y `Motivo`. No sirve `Tb_Sheets_List` (`Prog_HojasName`): va por hoja, y
   hay hojas con dos tablas, y `RuT_WrkBook_Sheets_List` la borra y la rehace entera, así que
   perdería cualquier columna escrita a mano.
2. **Clave.** El nombre de la tabla, que es único en el libro. Conviene renombrar antes las que
   conservan el nombre por defecto (`Tabla21`, `Tabla22`, `Tabla6`, `Tabla10`, `Tabla5`,
   `Tb_DefCols172146`...). Excel actualiza solo las fórmulas que las usan; el código que las
   nombra (`ListObjects("Tb_Conceptos")`...) hay que revisarlo a mano.
3. **Alta.** Una rutina que recorra todas las tablas y añada las que falten **sin tocar** las
   filas existentes (fusionar, no reconstruir), con un valor propuesto: Visible si alguna fórmula
   lee su `[#Totals]` o es una `DefCol`; Oculta si el código copia rangos que incluirían la fila de
   totales; Libre en el resto.
4. **Aplicación.** `Fnc_Lo_Totales_Siempre_Visibles` pasa a leer la tabla. ¿Se aplica también
   **Oculta**, o solo Visible? (`Tb_INSS`: mejor Libre o Oculta, por su fallo.)
5. **Reglas además de filas.** ¿Se mantiene la regla «todas las `Prog_DefCol_*`» o se convierte en
   filas explícitas?
6. **Validación sin abrir Excel.** Un `Herramientas/Validar_Totales.py` (o una ampliación de
   `CustomUI/Validar_Ribbon.py`) que cruce `xl/tables/*.xml` del `.xlsm` con la tabla de
   configuración: tablas sin fila, filas de tablas que ya no existen, tablas con fórmulas que leen
   `[#Totals]` que no están en Visible, y tablas en Visible sin funciones de totales.
7. **Volcado a CSV para git**, como `Tb_Tareas` y `Lo_RibbonUI` (`Volcar_Tareas_y_Ribbon.py`).
8. **Después**, quitar los `ShowTotals = True` del final de las rutinas, empezando por los pasos
   intermedios de `M_110` (`M_112` a `M_114`), que encienden y apagan los totales de `BDatos`
   (160.000 filas) entre paso y paso.
9. **Portabilidad.** El mismo diseño serviría en JyC y EP; candidato a skill, como
   `vba-ws-protect-status`.

## Cómo probarlo

**Módulos a reimportar (28):** `Form_Menu` (con su `.frx`), `M_000_Ini_APP`,
`M_110_Update_LSGES04_ACont`, `M_111_Del_Null_Reg_G04_ACont`, `M_112_Manage_Duplicates`,
`M_113_Assign_Concept_Eco`, `M_114_Clasif_Recibos`, `M_115_Assign_ImpAdm_CAcad`,
`M_130_Gen_Inf_Recibos`, `M_132_Anadir_Num_JIs_al_Inf`, `M_133_Anadir_Num_JIs_a_BDatos`,
`M_180_Restituir_Datos_Tabla`, `M_195_Add_UXXIdata_in_BDatos`, `M_210_Update_LSGES04_C_Acad_Ant`,
`M_211_Del_Null_Reg_CAcadAnt`, `M_215_Copy_ImpAdm_CAcadAnt`, `M_310_Update_LSace06_INSS`,
`M_311_Import_LSace06_CAcad`, `M_315_Copy_INSS_a_BD`, `M_410_Update_BD_AE4x4`,
`M_411_Import_AE4x1`, `M_413_Assign_ImpAdm_AE4x1`, `M_415_Copy_AE4_a_BD`,
`M_602_Cierre_Contable_PLANES`, `M_90_Rutinas_Menu_Aux`, `M___RibbonUI`, `Rut_Lo` y `Rut_Lo_Import`.

1. **Antes de reimportar**, copiar el libro tal como está (ya procesado con el código actual) como
   `PPub_BDatos_2026 - Antes_ShowTotals.xlsm`. La copia `PPub_BDatos_2026 - Antes.xlsm` de las 20:13
   se conserva para el fallo de `M_310`.
2. Reimportar los 28 módulos y compilar (Depuración → Compilar).
3. Guardar, cerrar y volver a abrir: `BDatos` y `Inf_Recibos_TIO` con totales, y
   `_DefCol_Inf_Peter` también (hasta ahora oculta). En Inmediato, ningún `!!! Rut_Lo_Totales_Mostrar`.
4. **Doble ejecución:** `M_110` (Sí/Sí/Sí/No) y `M_210` con los mismos ficheros que la última vez,
   guardar y comparar:
   `python Herramientas/Comparar_Libros_BD.py "PPub_BDatos_2026 - Antes_ShowTotals.xlsm" PPub_BDatos_2026.xlsm --hojas Sht__BD,Sht__BD_Dupl,Sht__BD_ErrDate,Sht__BD_IAdm_CAcadAnt,Sht__BD_INSS --por C_Acad`.
   Resultado esperado: **idénticos**, datos y logs.
5. `M_195`: ocultar la fila de totales de `Tb_DefCols_BD` (en Inmediato:
   `Prog_DefCol_BD.ListObjects(1).ShowTotals = False`) y lanzar la tarea con el formulario de
   progreso abierto, también desde Inmediato y línea a línea: `Rut_Progreso_Abrir "Prueba M_195"`,
   `RuT_Add_UXXIdata_in_BDatos` (contestar «No» a las dos preguntas) y `Rut_Progreso_Cerrar`. No
   sirve `Form_Menu`: no lista las tareas con «INTERNO» en el nombre, y `M_195` necesita un
   formulario abierto para escribir su progreso. Debe terminar sin error, y tras
   `Rut_Progreso_Cerrar` la `DefCol` vuelve a tener totales. Ojo: aun contestando «No», `M_195`
   escribe la fecha actual en `APP_Last_Import` y en `BD_Ant!D1` (comportamiento de siempre).
6. El botón de mostrar/ocultar columnas de `BDatos` sigue alternando.
7. Cancelar el diálogo de fichero de `M_110`: al cerrar, `BDatos` con totales (antes se quedaba sin
   ellos y el panel en `#REF!`).
