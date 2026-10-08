# Plan: importar el LSace06 INSS directamente desde la carpeta del Robot

**Proyecto:** `PPub_BDatos_2026.xlsm` · **Fecha:** 2026-10-08 · **Afecta a:** `M_310`, `M_311`, `M_314` (se retira),
`M_315` (solo comprobar), `M_000_Ini_Var_APP`, `Rut_File_Folder_NEXE`, documentación.

**Etiquetas:** ✅ **Hecho** · 🧪 **Aplicado, falta probarlo en Excel** · ⏳ **Pendiente**

## Estado: 🧪 APLICADO Y COMPILA (2026-10-08); PROBADO UNA VEZ CON LOS DOS CURSOS, FALTAN LOS FALLOS Y EL CASO DE UN SOLO CURSO

Fase 0 hecha y decisiones del usuario recogidas (sección 6). Faltan las fases 2 a 5. Comprobado el 2026-10-08:

- `M_000_Ini_Var_APP.bas` **del repo** ya trae las 16 constantes `LS06_*` nuevas (sin `LS06_ImpRec`, `LS06_RecINSS` ni
  `LS06_ACont_Cob`), pero está **sin commitear** y el libro **guardado** (08:47) aún tiene el antiguo de 29 columnas:
  hay que reimportarlo. `M_310`, `M_311`, `M_314` y `M_315` del libro son **idénticos** a los del repo (olevba), así que el
  «Úlitma/Última» de la captura eran solo datos viejos.
- `Func_LstObj_ListColumns_DefCol_Check_OK` **no da problema**: sale como correcta en cuanto encuentra una fila de la `DefCol`
  con `TitColGenInf` vacío (`RecFound`), así que 15 columnas en el fichero frente a 16 en la `DefCol` valen.
- `Rut_Lo_DataBodyRange_Filtered_Copy` **no da problema**: pega por posición y redimensiona `Lo_Target` con **su** ancho,
  así que `Rec found` queda dentro de la tabla y vacía.

## Cambios posteriores a este plan (2026-10-08, tras la primera prueba)

- `Lo_INSS` se **vacía entera antes de importar** (en `M_310`, con el recuento por curso en el informe); `M_311` ya no borra por curso.
- `M_310` **localiza los dos ficheros antes de vaciar**: si falta uno, pregunta si abortar o quedarse con un solo curso
  (`Fnc_LSace06_Elegir_Fichero`).
- El recuento de columnas «NO necesarias» usa `DefC_ProtectData` (contaba `DefC_HiddenCol`).
- Informe final de `M_315` **por curso**, y apariencia ajustada (líneas en blanco, sin «Lap» en la 1ª línea por curso).
- `M_315` **ya no vacía `Rec_Imp_INSS` entera** (solo los recibos de Ant/Pos) y carga esa columna en RAM para devolverla con sus valores.
- Primera prueba (Ant 2025-26: 24.392 reg; Pos 2026-27: 23.970 reg): importa, y el cruce da 1.156 y 22.358 recibos encontrados.

## 1. Qué se pide

1. Los datos se leen **directamente de una ruta configurada**: el nombre definido `APP_Ruta_Robot_LSc06_Cacad`
   + el **curso académico** que toque + la **carpeta de fecha más reciente**.
2. El fichero se llama `LSace06_C_Acad_2026-27_INSS_(2026-10-02).xlsx`: cambian el curso y la fecha según
   lo requerido.
3. Si falla la ruta o el fichero: **mensaje que diga cuál ha sido el fallo** y proponer **abortar** o **abrir el
   explorador** para elegir el fichero a mano.
4. El fichero **ha cambiado de estructura y ya viene filtrado**: no hace falta `M_314`. La tabla `Lo_INSS` tiene
   **una columna más** que el fichero (`Rec found`), que se usa en otro módulo (`M_315`).

## 2. Cómo está hoy (comprobado)

| Pieza | Hoy |
|---|---|
| `M_310` `RuT_Update_LSace06_CAcad_ImpAdm_INSS` | Pregunta Sí/No, y llama **dos veces** a la importación (curso Ant y Pos) con `"LSACE06_<curso>\|LSace06_C_Acad_<curso>"`; después ofrece `M_315`. |
| `M_311` `Rut_Lo_Import_LoData_LoDefCol_LSace06` | Abre el diálogo de archivos (`Rut_File_Select_V2`, que siempre arranca en la carpeta del libro), abre el libro, comprueba cabecera, quita `ORIGEN`, formatea, **borra** por curso, por nombre `*INSS*` y por importe negativo, llama a `M_314`, vacía columnas no necesarias, borra de `Lo_INSS` el/los cursos y copia. |
| `M_314` `RuT_Find_Rec_INSS_C_Acad` | Ordena, se queda con el **primer** recibo de cada `Plan_DNI`, pregunta si exportar el `-INSS`, borra los cobrados en `AnoCont-1` del curso Ant. |
| `M_315` `Rut_Copy_ImpINSS_en_BDatos` | Cruza `BD` con `Lo_INSS` por `Ref` en RAM y escribe `Found / Not Found` en **`LS06_RecFound`**. |
| Nombres en `ConfigAPP` | `APP_Ruta_Robot_Ges04_Acont` (C11), `APP_Ruta_Robot_Ges04_CAcad` (C12) y **`APP_Ruta_Robot_LSc06_Cacad` (C13)** ya existen en el libro, pero **ningún módulo los usa todavía**. |

### El fichero nuevo (`LSace06_C_Acad_2026-27_INSS_(2026-10-02).xlsx`, leído con openpyxl)

- Una sola hoja, `Data`, **sin `ListObject`**, **15 columnas**: `REFERENCIA, CURSO_ACADEMICO, PLAN, EXPEDIENTE,
  NUM_RECIBO, COD_ACTIV, DESC_CONCEPTO, COD_CONCEPTO, DNI, APELLIDOS_Y_NOMBRE, IMP_UNITARIO, CANTIDAD,
  TOTAL_CONCEPTO, COD_DESCUENTO, DESC_DESCUENTO`. **Sin `ORIGEN`.**
- Todo **texto** (`"1.12"`, `"2026248517338"`), como los demás ficheros del Robot.
- `Prog_DefCol_LSace06` (captura) ya tiene **16 filas**: esas 15 (columna `TitColGenInf`) más `RecFound`,
  que no tiene `TitColGenInf` porque no viene en el fichero.
- Carpetas reales (en `Robot_PPub\LSace06_C_Acad\<curso>\<AAAA-MM-DD>\`): hay 4 ficheros por fecha
  (`…_(fecha).xlsx`, `…_Evol_…`, `…_INSS_…`, `…_Inf_…`); **solo nos interesa el `_INSS_`**.

## 3. Hallazgo importante: el repo y el libro ya no coinciden en `LS06_*`

`M_000_Ini_Var_APP.bas` del repo todavía define **29** columnas `LS06_*` (`LS06_ACont_Cob`, `LS06_ImpRec`,
`LS06_RecINSS`…; `LS06_RecFound = 29`). La captura de `Prog_DefCol_LSace06` ya genera **16**
(`LS06_RecFound = 16`). Consecuencias:

- `M_311` usa `LS06_ImpRec` y `M_314` usa `LS06_RecINSS` y `LS06_ACont_Cob`, que **dejan de existir**: mientras
  no se quiten, el libro **no compila** (`Variable no definida`).
- `M_315` solo usa `LS06_Ref`, `LS06_Concept_Imp` y `LS06_RecFound`, que se conservan (con los índices nuevos).
- La captura de `BD_INSS` dice «**Última** Importación…» y el código escribe «**Úlitma** Importación…» (con
  otro formato de fecha): el libro parece llevar un `M_311` distinto del repo. **Antes de editar nada hay que
  exportar del libro `M_000_Ini_Var_APP`, `M_310`, `M_311` y `M_314`** y comparar (olevba), para trabajar sobre
  lo que de verdad hay.

## 4. Diseño

### 4.1 Función nueva: localizar el fichero

`Fnc_LSace06_Robot_Localizar(CAcad As String, ByRef Ruta As String, ByRef Fallo As String) As Boolean`
(en `M_311`, o en un módulo propio si crece). Pasos, cada uno con **su** mensaje de fallo:

1. **Leer `Prog__APP.Range("APP_Ruta_Robot_LSc06_Cacad")`** (se comprueba primero dónde vive el nombre:
   `ConfigAPP`, hoja `Prog__APP`; por el 1004 de nombres de otra hoja). Fallo: «La ruta no está configurada».
2. **Convertir la ruta** con `Fnc_Format_Ruta` (si es `https://nexe.ua.es/…` pasa a letra de unidad). Con URLs
   `Dir$` da el error 52.
3. **Carpeta del curso**: `<ruta>\<CAcad>` (p. ej. `…\LSace06_C_Acad\2026-27`). Fallo: «No existe la carpeta del curso …».
4. **Carpeta de fecha más reciente**: recorrer las subcarpetas (`Dir$(…, vbDirectory)`), quedarse con las que
   empiezan por `AAAA-MM-DD` válido y elegir la **fecha mayor** (se compara la fecha, no el texto; se ignoran
   sufijos como `-r`). Fallo: «No hay carpetas de fecha en …».
5. **Fichero**: `LSace06_C_Acad_<CAcad>_INSS_(<AAAA-MM-DD>).xlsx` dentro de esa carpeta. Fallo: «No existe el fichero … en la carpeta más reciente (<fecha>)».

Devuelve `True` y la ruta completa, o `False` y el texto del fallo.

### 4.2 Qué pasa si falla (punto 3 de la petición)

Aviso con el **fallo concreto** y hasta **tres opciones** (hace falta un cuadro de 3 botones: `Func_MsgBox_vbYesNo` solo
tiene dos; `Form_MsgBox` admite hasta 3, se mira al implementar):

1. **Usar la fecha anterior**: solo si el fallo es «a la carpeta más reciente le falta el `_INSS_`» y existe una carpeta de
   fecha **anterior** que sí lo tiene. El aviso **dice cuál es** («Más reciente: 2026-10-08, sin `_INSS_`. Anterior con fichero:
   2026-10-02») y, si se acepta, usa esa. Se busca hacia atrás hasta la primera que lo tenga.
2. **Abrir el explorador** para elegir a mano: `Rut_File_Select_V2` empezando en la carpeta que sí existe (hoy el diálogo
   arranca siempre en la del libro: se usa el parámetro `PathFile`, que existe y no se usaba).
3. **Abortar**: como una cancelación (`Arch_New_Name = "Cancel"`, mismo cierre de siempre).

Con otros fallos (ruta sin configurar, sin carpeta de curso, sin carpetas de fecha) no hay fecha anterior que ofrecer: solo 2 y 3.
Lo que se elija a mano **debe llamarse** `LSace06_C_Acad_<curso>_INSS_…` (el curso pedido **y** `_INSS_`): se valida con
`Fnc_Nombre_Fichero_Valido` y, si no cuadra, se avisa y se vuelve a ofrecer. Con la fecha anterior no hace falta: ya cumple el nombre.

### 4.3 Cambios en `M_311` (`Rut_Lo_Import_LoData_LoDefCol_LSace06`)

| Paso actual | Qué se hace |
|---|---|
| Diálogo `Rut_File_Select_V2` como primer paso | Primero `Fnc_LSace06_Robot_Localizar`; el diálogo solo como alternativa (4.2). |
| Parámetro `Arch_New_Name` con `"A\|B"` | Pasa a recibir el **curso** (o el nombre completo ya resuelto). Se retira la lista de prefijos `\|`. |
| Comprobar cabecera (`Func_LstObj_ListColumns_DefCol_Check_OK`) | **Se mantiene**, pero la tabla del fichero tiene 15 columnas y la `DefCol` 16: hay que ver cómo compara (comprobado: para en la primera fila sin `TitColGenInf`; vale tal cual). |
| Quitar `ORIGEN` | Se puede dejar (inofensivo) o quitar; el fichero nuevo no lo trae. |
| Formateo (`Rut_Lo_Format_LoData_LoDefColData`) | **Se mantiene**: sigue siendo texto con punto decimal; la conversión en RAM ya lo resuelve. |
| Borrar por `C_Acad <> Ant y Pos` | **Se sustituye por una comprobación**: si el `CURSO_ACADEMICO` del fichero no es el pedido, avisar y abortar (el fichero ya es de un solo curso). |
| Borrar `<>*INSS*` | Se retira (ya viene filtrado). |
| Borrar importes `< 0` | **Se mantiene, por seguridad** (decisión del usuario), pero sobre `LS06_Concept_Imp` (`LS06_ImpRec` ya no existe). Con el mismo aviso en el log y sin quedarse sin recibos. |
| `RuT_Find_Rec_INSS_C_Acad` (`M_314`) | **Se retira** (petición 4). |
| `CantCAcad` / `CursoArray` con 1 ó 2 cursos | Se simplifica: **un fichero = un curso**. Las fechas `APP_Last_LSace06_CAcad*` se anotan según el curso importado. |
| Vaciar columnas no necesarias (`DefC_ProtectData`) | Se mantiene si la `DefCol` lo sigue marcando; si no queda ninguna, se retira. |
| Borrar de `Lo_INSS` los recibos del curso y copiar | **Se mantiene** (incluido el `DoEvents` antes de `Rut_Lo_WrkSht_Preparar`). La copia lleva 15 columnas a una tabla de 16: comprobado, copia por posición y deja `Rec found` vacía. |

### 4.4 Cambios en `M_310`

Pasa a llamar a la importación con el **curso** (`C_Acad_Ant`, `C_Acad_Pos`) en vez de la lista de nombres. El
resto del flujo (Sí/No, `M_315`, informe) no cambia. El texto del informe incluirá la **ruta completa** del
fichero que se ha leído (también `\|` desaparece de los avisos).

### 4.5 Retirar `M_314`

Es lo que pide el punto 4, pero conviene saber lo que se pierde (las tres cosas, también retiradas):

1. **Quedarse con el primer recibo de cada `Plan_DNI`**: el fichero ya viene filtrado (decisión del usuario).
2. **El `MsgBox` «¿Exportar los recibos INSS a un Excel para Felipe?»**: se prescinde (decisión del usuario; el Robot ya
   deja el `_INSS_` en su carpeta).
3. **Borrar los del curso Ant cobrados en `AnoCont-1`** (`LS06_ACont_Cob`): la columna ya no existe y el Robot **no** los
   excluye, pero el usuario dice que **no es necesario** → la regla se retira.
4. De los importes negativos, en cambio, **sí** se conserva el borrado (4.3).

Se borra `M_314_Find_Rec_INSS.bas` del repo y del libro, y se actualizan los comentarios que lo citan
(`M_1___`, `M_3___`, cabecera de `M_311`).

### 4.6 `M_315` y `M_000_Ini_Var_APP`

- `M_315`: no cambia, pero se **comprueba** con los índices nuevos (`LS06_RecFound = 16`) y que el cruce
  siga funcionando con `Ref` ordenada.
- `M_000_Ini_Var_APP`: se **reexporta del libro** (las constantes las generan las hojas `_DefCol_*`) para que
  repo y libro coincidan.

## 5. Orden de trabajo

| Fase | Qué | Estado |
|---|---|---|
| 0 | Comparar libro y repo (olevba), cabecera con la col. extra y copia por posición (ver «Estado»). Pendiente solo mirar el **valor** de `APP_Ruta_Robot_LSc06_Cacad` en `ConfigAPP!C13`. | ✅ Hecho |
| 1 | Resolver las preguntas de la sección 6. | ✅ Hecho |
| 2 | `Fnc_LSace06_Robot_Localizar` + el aviso «fecha anterior / explorador / abortar» (4.1 y 4.2), con `vbYesNoCancel` nativo (no hizo falta `Form_MsgBox`). `Rut_File_Select_V2` usa ya `PathFile`. | 🧪 |
| 3 | `M_311` simplificado (4.3), `M_310` (4.4) y `M_314` borrado del repo (4.5). Script cp1252/CRLF, sello `Last Rev.`, bytes verificados (0 `U+FFFD`, sin LF sueltos). | 🧪 |
| 4 | Dar la **lista de componentes a reimportar** (`M_000_Ini_Var_APP`, `M_310`, `M_311`, y borrar `M_314`); el método lo elige el usuario. | ⏳ |
| 5 | Probar (sección 7), actualizar `CLAUDE.md`, `Docs/` y los comentarios de `M_1___`/`M_3___`. | ⏳ |

## 6. Decisiones del usuario (2026-10-08)

| # | Pregunta | Respuesta |
|---|---|---|
| 1 | ¿Qué hacer si a la carpeta más reciente le falta el `_INSS_`? | Avisar y dar a elegir: **abortar**, o **decir cuál es la fecha anterior** y ofrecerla (4.2). |
| 2 | ¿El Robot excluye importes negativos y cobrados en `AnoCont-1`? | **No excluye ninguno.** Solo es necesario, por seguridad, el borrado de **importes negativos**; se conserva. |
| 3 | ¿Se prescinde de exportar los recibos INSS «para Felipe»? | **Sí.** |
| 4 | ¿Qué nombre debe tener un fichero elegido a mano? | **El esperado**: con el curso y con `_INSS_`. Si no, se rechaza. |
| 5 | `M_000_Ini_Var_APP` y la cabecera/copia con la columna de más | Constantes ya actualizadas en el repo; cabecera y copia comprobadas, sin problema. |

## 7. Cómo probarlo

1. **Camino feliz**: `M_310` con Ant y Pos → localiza `…\2025-26\2026-10-02\…_INSS_(2026-10-02).xlsx` y el del
   `2026-27`; el log enseña la ruta; `BD_INSS` queda con los dos cursos y `Rec found` vacía; después, `M_315`
   rellena `Found/Not Found`.
2. **Doble ejecución** (como en los pasos a RAM): libro de antes, importando a mano el mismo fichero, contra el
   libro nuevo con la ruta automática; `Herramientas/Comparar_Libros_BD.py` sobre `Sht__BD` (columna `Rec_Imp_INSS`)
   y la hoja `BD_INSS` (con `--hojas`). Las diferencias esperables: filas que antes borraba `M_314` (cobrados en `AnoCont-1` del curso Ant) y ahora
   se quedan, ya que se retira esa regla. Son justo lo que hay que revisar.
3. **Fallos** (uno por uno): nombre `APP_Ruta_Robot_LSc06_Cacad` vacío; ruta que no existe; curso sin carpeta;
   curso sin carpetas de fecha; fecha sin `_INSS_` (con y sin fecha anterior disponible); en cada uno, probar las opciones que ofrezca
   el aviso. Y un fichero con importes negativos: se borran y queda anotado en el log.
4. **Fichero elegido a mano** con otro curso o sin `_INSS_` en el nombre → aviso y nueva oferta, sin tocar `BD_INSS`.
5. **Ruta NEXE** (`https://nexe.ua.es/…`) y ruta de unidad: las dos deben localizar el fichero.
6. **Trabajo y casa**: el nombre `APP_Ruta_Robot_LSc06_Cacad` vive en el libro (viaja con él), pero la **unidad
   `Y:`** no tiene por qué existir igual en los dos ordenadores: se prueba en ambos.

## 8. Riesgos

- **Libro sin compilar entre fases**: las constantes `LS06_*` nuevas dejan sin definir las antiguas hasta que
  `M_311` y `M_314` se actualicen; por eso se reimportan **a la vez**.
- **Excel abierto con el fichero**: la captura muestra el `_INSS_` abierto en Excel (existe `~$LSace06…`).
  Se abre con `ReadOnly:=True`, así que no estorba, pero el libro de datos y el de programación no pueden tener
  el **mismo nombre** (Excel no admite dos iguales): el nombre `LSace06_C_Acad_…` es distinto del de `PPub_BDatos_…`, así
  que no pasa, pero conviene recordar que `Workbooks.Open` devolvería `Nothing` en silencio con un homónimo.
- **Orden `Dir$` de carpetas**: no se fía del orden alfabético; se calcula la fecha mayor.
