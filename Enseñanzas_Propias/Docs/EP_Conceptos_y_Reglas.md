# EP - Conceptos y Reglas

Liquidación de Enseñanzas Propias / Títulos Propios de la Universidad de Alicante (`EP_202x-2x_BaseDatos-LIQ _V2.6.xlsm`). Este documento recoge, a partir de una lectura completa del código VBA exportado en `VBA_Moduls/`, los conceptos de negocio y las reglas que el código realmente aplica — no es un manual de usuario, es una referencia para quien vaya a tocar el código.

## Ámbito

EFP (Enseñanzas de Formación Permanente) y CFC/AFC (Cursos y Actividades de Formación Complementaria) comparten el mismo libro y el mismo código; se distinguen por el switch `APP_EFP_o_CFC` y por el Plan de estudios activo. El libro se reutiliza cada año: `M79` genera, a partir de esta plantilla, una copia real por combinación Tipo×Curso Académico (ver *M79 — copia anual del libro*).

## Glosario

| Término | Significado |
|---|---|
| **LSGES04** | Exportación de recibos del sistema contable de la UA; es la fuente de datos que se importa periódicamente a `Prog_BD`. |
| **EFP** | Enseñanzas de Formación Permanente. |
| **CFC / AFC** | Cursos de Formación Continua / Actividades de Formación Complementaria — se tratan como un mismo grupo (`CFCyAFC`) a efectos de fichero/copia anual. |
| **Plan** | Código del plan de estudios (título propio) al que pertenece un recibo. |
| **Curso Académico (Ant/Pos)** | Cada copia anual del libro trabaja con dos ediciones: la del curso académico "Anterior" y la "Posterior" al año contable en curso. |
| **Año Contable (`APP_AñoCont`)** | Ejercicio contable activo; distinto del Curso Académico (un recibo puede cobrarse en un año contable distinto al de su curso). |
| **ImpRec / ImpCob / ImpAdm / ImpAcad** | Importe del Recibo (emitido), Importe Cobrado, Importe Administrativo (tasa de gestión, concepto 1303) e Importe Académico (docencia, conceptos 1310/1311). `Acad` se obtiene siempre por resta (`Total − Adm`), nunca con un `SumIfs` directo. |
| **Coef_VRI** | Porcentaje de retención que aplica el Vicerrectorado de Investigación sobre el importe académico neto de un Plan. |
| **JI** | Justificante de Ingreso. Hay tres pares de columnas en `Prog_BD` (`_Adm`/`_Acad`): `JI_Emi`, `AD_0010`, `JI_443`. |
| **ExpAdm** | Expediente Administrativo al que se imputa la operación. |
| **RDT** | En este proyecto, campo de control/consolidación asociado a la Liquidación (no confundir con el "RDT = Redistribución de Crédito" de otros proyectos hermanos de la UA). |
| **Orgánica** | Unidad orgánica a la que se imputa el importe; un Plan puede tener más de una asociada (se concatenan con `" - "`). |
| **Ajuste de Matrícula** | Recibo con `ImpAdm < 0`; se excluye de casi toda la contabilidad (`AutoFilter Field:=BD_ImpAdm, Criteria1:=">=0"` para excluirlo). |
| **AE4** | Código de Actividad Económica 4 = Enseñanzas Propias (`M72` conserva solo `AE=4` o `AE=300`). |
| **AE4x4 / AE4x1** | Ficheros/informes contables generados a partir de las 4 copias anuales EFP/CFC × Curso Ant/Pos (ver más abajo). |
| **Deleted / DeletedConJI** | Marca que recibe en `Prog_BD` un recibo que ha dejado de aparecer en la última importación de LSGES04; `DeletedConJI` (con aviso obligatorio) si ya tenía un JI asignado — es la salvaguarda para no perder silenciosamente un recibo ya justificado administrativamente. |

## Qué es una Liquidación y cómo se identifica

Una Liquidación es el conjunto de recibos de `Prog_BD` (tabla `Lo_BD`) de un **Plan + Curso Académico** concretos. La genera `Rut_00_Liquid_TitProp(Liq_Plan, CursAcad)` en `M10__Liquid_EP.bas`: filtra `Lo_BD` por `BD_Plan = Liq_Plan` y `BD_C_Acad = CursAcad`, excluye los recibos de Ajuste de Matrícula (`BD_ImpAdm < 0`), y, salvo que el Plan sea de Microcredencial, excluye también los recibos negativos (salvo switch `Sw_VerRecNeg`) y los recibos con `ImpCob = 0` (salvo switch `Sw_VerRecNoCob`). El resultado se vuelca a la tabla `Lo_Liq` de la hoja `Wk_TitP_Liquid`. Si no hay ningún recibo pagado, aborta con `MsgBox`.

> *Implementado en:* `Rut_00_Liquid_TitProp` (`M10__Liquid_EP.bas`).

## El ciclo de un recibo: de LSGES04 a `Prog_BD`

```
        Exportación LSGES04 (sistema contable UA)
                        │
                        ▼
   ┌─────────────────────────────────────────┐
   │ M01 — Importar a Prog_LsGes04            │
   └─────────────────────────────────────────┘
                        │
                        ▼
   ┌─────────────────────────────────────────┐
   │ M02 — Depurar: fuera de curso, Matrícula │
   │ = N, ActivEco ≠ 4, matrícula a coste 0,  │
   │ subvencionados al 100%                   │
   └─────────────────────────────────────────┘
                        │
                        ▼
   ┌─────────────────────────────────────────┐
   │ M02_Manage_Duplicates — detectar y       │
   │ resolver referencias repetidas           │
   └─────────────────────────────────────────┘
                        │
                        ▼
   ┌─────────────────────────────────────────┐
   │ M03 — Asignar cta. de ingreso (CCC)      │
   │ M04 — Asignar Concepto Económico y       │
   │       Tipo de Enseñanza                  │
   │ M05 — Asignar Tipo de Recibo             │
   │ M06 — Primera Tasa Admin. por matrícula  │
   └─────────────────────────────────────────┘
                        │
                        ▼
   ┌─────────────────────────────────────────┐
   │ M07 — Merge-join con Prog_BD: altas,     │
   │ actualizaciones, bajas (Deleted)         │
   └─────────────────────────────────────────┘
                        │
                        ▼
   ┌─────────────────────────────────────────┐
   │ M08 — Actualizar Coef_VRI por Plan       │
   └─────────────────────────────────────────┘
```

### M02 — Depuración inicial

Solo se procesan recibos del Curso Académico configurado, con `Matrícula ≠ N` y `ActivEco = 4`, descartando también las matrículas a coste cero y los recibos subvencionados al 100 % (`ImpRec = 0`). `M02_Del_Reg_EFP_o_CFCyAFC.bas` separa además el subconjunto EFP del CFC/AFC.

### Duplicados (`M02_Manage_Duplicates.bas`)

`RuT_Duplicates_Search` ordena por `BD_Ref` y compara cada fila con la anterior: si coinciden, marca la primera como `"RpN"` y la siguiente como `"RepeN+1"` en `BD_Incidencias`, y llama a `RuT_Duplicates_Search_Mark_DIFF`, que compara columna a columna (solo las marcadas `Compare=True` en `Prog_DefCol_BD`) y anota en `BD_H_Incidencias` el detalle `"En Col. nºX (Nombre)<>[valor]"`. Después borra de la tabla de trabajo todos los repetidos salvo el último de cada serie, copia los "finalistas" a `Prog_BD_Dupl` (histórico), y en esa tabla purga los que ya estaban repetidos de tandas anteriores (`RpIdem`) o los que no tuvieron ningún cambio real (`"(en 0 Cols)"`).

> *Implementado en:* `RuT_Duplicates_Search` / `RuT_Duplicates_Search_Mark_DIFF` (`M02_Manage_Duplicates.bas`).

### Tipo de Recibo (`M05_Asign_Tipo_Recibo.bas`)

Se clasifica cada fila de `Prog_LsGes04` mediante una cascada de `AdvancedFilter` contra rangos con nombre definidos en la hoja (`Tb_CriT_*`, no en el código VBA), en este orden:

| Orden | Tipo asignado | Criterio (rango) |
|---|---|---|
| 1 | `_ERR_Date_` | `Tb_CriT_Reg_Err` — fechas inconsistentes |
| 2 | `Emitido` | `Tb_CriT_Emitido` |
| 3 | `EjeAnt` | `Tb_CriT_EjeAnt` — del ejercicio contable anterior |
| 4 | `Añejo` | `Tb_CriT_Añeja` — anteriores al ejercicio anterior |
| 5 | `Aplazado` | `Tb_CriT_Aplazado` |
| 6 | `ADxAplz` | `Tb_CriT_ADxAplz` — anulado por aplazamiento, a cobrar el año siguiente |
| 7 | `_Devol_` | `ImpRec < 0` |
| 8 | `_AñoCont'XX_` | `ACont_Cob < APP_AñoCont` (cobrado en un ejercicio anterior al de contabilización) |
| 9 | `_Ajust_Matríc_` | `ImpAdm < 0` |

El resultado se escribe en `BD_Tipo_Rec` (columna 36). Un contador final (`RegsSinTipo`) avisa si quedan recibos sin clasificar tras pasar las 9 reglas.

> *Implementado en:* `RuT_Determinar_Tipo_Recibo` (`M05_Asign_Tipo_Recibo.bas`). Nota: el texto de informe del paso 9 está actualmente copiado del paso 8 (ver `Informe_Bugs.md`, hallazgo A7) — el orden y el criterio de clasificación en sí son correctos, solo el mensaje de log está mal etiquetado.

### Coef_VRI (`M08_Actualizar_Tb_Coef_VRI.bas`)

Es un porcentaje de retención (entero, p. ej. 15 o 20) sobre el importe académico neto:

```
Rg_VRI = Round((Imp_Cob − Imp_Adm) × Coef_VRI / 100, 2)
```

Se guarda por Plan en `Prog_Coef_Ret_VRI`. `RuT_Lo_Coef_VRI_Actualizar` recorre `Prog_BD` agrupado por Plan y, si el Plan no existe aún en `Prog_Coef_Ret_VRI`, crea una fila nueva con `ORden="x"` y le asigna el máximo `Coef_VRI` visto en las filas de ese Plan; si es 0, usa el valor por defecto de la fila 1 (EFP) o 2 (CFC) de esa misma tabla. Cuando se da de alta un recibo nuevo (`M07`), se hace la búsqueda inversa: si el Plan no está en `Prog_Coef_Ret_VRI`, se usa igualmente ese valor por defecto 1/2 según EFP o CFC/AFC.

> *Implementado en:* `RuT_Lo_Coef_VRI_Actualizar` (`M08_Actualizar_Tb_Coef_VRI.bas`), reutilizado en `M07_Actualiz_BDatos_con_LsGes04.bas` y en el cálculo de `M10__Liquid_EP.bas`/`M39_Resumen_Planes_C_Acad.bas`.

### Sincronización `Prog_BD` ↔ `Prog_LsGes04` (`M07`)

Merge-join ordenado por `BD_Ref`: refs iguales → actualiza y registra incidencias de cambio de importe (`Chg_ImpRec`/`ImpCob`/`ImpAdm`); ref nueva → alta en `Prog_BD` con búsqueda de `Coef_VRI`; ref que ya no aparece en la última importación → se marca `"Deleted"` o, si ya tiene JI asignado, `"DeletedConJI"` con aviso obligatorio al usuario — la salvaguarda contra pérdida silenciosa de recibos ya justificados administrativamente.

### Primera Tasa Administrativa por matrícula (`M06`)

`Rut_Assign_Imp_AdmAcad_C_Acad` ordena por Plan+DNI+NumRec y, para cada combinación Plan+DNI **nueva**, copia el importe académico/administrativo/descuento del **primer** recibo a `BD_Rec_Imp_Acad/Adm/Dto` — los recibos repetidos del mismo alumno no vuelven a aportar tasa administrativa.

## Botones por rango con nombre — la interfaz de `Wk_TitP_Liquid`

No hay Ribbon custom ni UserForm de pantalla principal: la interacción diaria es directamente sobre la única hoja visible (`Wk_TitP_Liquid`), donde cada rango con nombre actúa como "botón" al hacer clic o escribir en él.

**Al hacer clic (`Worksheet_SelectionChange`):**

| Rango | Acción |
|---|---|
| `Liquid_Del_Filtro` | Quita todos los filtros de la tabla |
| `Liquid_Help` | Muestra/oculta los comentarios de ayuda de celda |
| `Liquid_Sw_VerRecNeg` | Alterna ver recibos negativos |
| `Liquid_Sw_VerRecNOCob` | Alterna ver recibos no cobrados |
| `Liquid_RDT_Totales` | Rota entre las 3 vistas de resumen RDT (`Tot:` / `Tot=` / `Tot-`) |

**Al escribir (`Worksheet_Change`):**

| Rango | Acción |
|---|---|
| `Liquid_Plan` | Cambia de Liquidación (avisa si hay cambios sin guardar, vía `Sw_Cmb`) |
| `Liquid_Num_JI_Emi`, `Liquid_Num_AD_0010`, `Liquid_Num_JI_443`, `Liquid_Num_ExpAdm`, `Liquid_Num_RDT`, `Liquid_Coef_VRI`, `Liquid_Orgánica` | Escriben ese valor en todas las filas visibles/filtradas de la Liquidación |
| `Liquid_Filtro_N_Liquid`, `Liquid_Filtro_N_Recibo`, `Liquid_Filtro_Año_Vto`, `Liquid_Filtro_Año_Emi`, `Liquid_Filtro_Año_Cob` | Filtran la tabla por la columna correspondiente |

> *Implementado en:* `Wk_TitP_Liquid.cls`.

### Guardar la Liquidación sobre `Prog_BD`

Al guardar, `Rut_2_Actualizar_Dat_TitPropHist_con_Dat_Liquid` (`M10__Liquid_EP.bas`) hace un segundo merge-join, esta vez entre `Lo_TPLiquid` (lo que se ve en pantalla) y `Prog_BD`, ambos ordenados por Ref, y traslada a `Prog_BD` los campos JI/AD-0010/JI-443/ExpAdm/RDT/Coef_VRI/Orgánica/Observaciones y la Tasa Administrativa (ajustada si `Liq_Ajst_Tadm` no está vacío).

> ⚠️ Este merge-join tiene un `Case Else` ausente que puede desincronizar los dos punteros a mitad de recorrido — ver `Informe_Bugs.md`, hallazgo A3, antes de asumir que "guardar" siempre actualiza todo lo que debería.

## Cierre Contable de Planes (`M38`)

`Rut_Cierre_Contable_AñoCont` primero verifica que la hoja activa coincide con el Curso Académico/Año Contable actuales (Ant vs Pos), abortando si no. Luego prepara `Prog_BD` (desprotege, ordena por `C_Acad`/`ACont_Emi`/`Plan`), filtra `ImpAdm >= 0` (excluye Ajustes de Matrícula) y `Tipo_Rec <> "_ERR_Date_"`, y copia el rango visible resultante a una tabla temporal en la hoja `Sht__Buffer` — patrón "filtrar → copiar a Buffer → ListObject temporal" que se repite en `M22_NEW`, `M32`, `M39` y `M40`. Según sea Curso Académico Ant o Pos, recorre cada Plan (vía una `Collection` de valores únicos) calculando por `Application.SumIfs` los importes Emitido/Anulado/Cobrado/RDT/Pendiente-RDT/ADxAplz/Aplazado/EjeAnt/Pendiente-Cobro, separando siempre la parte Administrativa (concepto 1303) de la Académica (1310/1311) por resta.

> *Implementado en:* `Rut_Cierre_Contable_AñoCont`, `RuT_Cierre_Contable_Planes_CAcad_Ant`/`_Pos` (`M38_Cierre_Contable_PLANES.bas`).

## Números de JI (`M41`/`M42`)

Cada recibo de `Prog_BD` tiene 6 columnas de JI (`JI_Emi`, `AD_0010`, `JI_443`, cada una con variante `_Adm`/`_Acad`). El diseño previsto es:

- **M41** (`Rut_Añadir_a_Inf_Recibos_JIs_de_BDatos`) recorre la tabla-resumen del informe contable y, para cada combinación Tipo_Recibo+Concepto_Económico, filtra `Prog_BD` y vuelca (deduplicando con `Scripting.Dictionary`) los JI únicos encontrados, unidos con `", "`.
- **M42** (`Rut_Añadir_a_BDatos_JIs_de_Inf_Recibos`) hace el camino inverso: toma el JI tecleado a mano en el informe-resumen y lo "explota" hacia todas las filas de detalle que comparten Tipo_Rec/Concepto en `Prog_BD`.

> ⚠️ **Este flujo no compila en el estado actual de este libro**: M41/M42 referencian `Sht__BD`/`Sht__Inf_Recibos_TIO`, hojas que no existen en Enseñanzas_Propias (parecen código traído sin adaptar del proyecto hermano PPub). Ver `Informe_Bugs.md`, hallazgo B3, antes de dar por buena esta descripción como comportamiento activo — es el diseño previsto, no necesariamente el que se ejecuta hoy.

## AE4x4 / AE4x1 y "las 4 copias"

Las "4 copias EFP/CFC de curso Ant/Pos" son los 4 ficheros `*_BaseDatos_Liq_*.xlsm` que `M79` genera cada año (uno por Tipo de Enseñanza EFP/CFCyAFC × uno por Curso Académico Ant/Pos). `M50` (`RuT_Inf_Contable_Recibos_AE4x4`) los reimporta secuencialmente en 4 pasos — `EFP_<CAcadAnt>`, `EFP_<CAcadPos>`, `CFCyAFC_<CAcadAnt>`, `CFCyAFC_<CAcadPos>` — cada uno vía `M51` (`Rut_Lo_Import_AE4x1`), que abre el Excel elegido en solo lectura, valida por nombre que es el fichero correcto y copia su tabla `BDatos` a un `ListObject` intermedio para generar los informes contables `Inf_Cont_AE4x4`/`AE4x41`.

### Relación con el proyecto hermano "Precios Públicos - PPub"

`M50_Inf_Cont_AE4x4.bas` lleva literalmente el comentario `'- M_410_Update_Lo_AE4` referenciando esa fase del proyecto `Precios Públicos - PPub` (mismo repo padre, carpeta `../Precios Públicos - PPub/`). Este libro EP es el que **genera** (vía `M79`) los ficheros AE4x4/AE4x1 que PPub luego importa en su pipeline 410–415 — antes de tocar la lógica de AE4/JI's/Coef_VRI en cualquiera de los dos libros, conviene revisar el módulo equivalente en el otro.

> ⚠️ El flujo M50/M51 tal como está exportado hoy tampoco compila de forma limpia (Subs públicas duplicadas entre `M50_Inf_Cont_AE4x4`/`_AE4x41` y entre `M51_Import_AE4x1`/`_AE4x11`, variable `Lo_AE4x1` no declarada en una de las dos ramas) — ver `Informe_Bugs.md`, hallazgos B1/B2/B5/B6. Todo apunta a que `_AE4x41`/`_AE4x11` son ramas de desarrollo en curso, no el camino que se ejecuta a diario.

## Restituir BDatos (`M71`)

`Mod_Restituir_Tabla_Prog_BD` recupera el contenido de `Prog_BD` (y de `Prog_BD_Deleted`/`Prog_BD_Dupl`) desde un Excel de backup elegido a mano, e importa también los parámetros de `Prog__APP` (Año Contable, Curso Académico, EFP/CFC, Plan activo...) **desde el backup**, sobrescribiendo los del libro actual antes de mostrar ningún resumen. Tras la restitución ofrece, con dos `MsgBox` de confirmación, borrar registros no válidos (llama a `M72`) y/o generar directamente el Excel anual (llama a `M79`).

> ⚠️ A diferencia de `M51` (que valida el nombre del fichero elegido), `M71` acepta cualquier `*.xls?` sin comprobar que sea un backup de EP, y la sobrescritura de `Prog_BD`/`Prog__APP` ocurre sin ninguna confirmación previa — solo se informa después, ya consumada. Ver `Informe_Bugs.md`, hallazgo B12.

## M79 — copia anual del libro

`Rut_Crear_WB_EFP_o_CFCyAFC_de_CAcad_Ant_o_Pos` construye el nombre `<TipoEP>_<CursoAcad>_BaseDatos_Liq_<VersiónApp>.xlsm` (`TipoEP` = `EFP_` o `CFCyAFC_`), propone la ruta vía `Application.GetSaveAsFilename` (el usuario puede cambiarla) y ejecuta `ThisWorkbook.SaveCopyAs` — copia binaria íntegra del libro actual (código y datos), sin reiniciar nada. Es la rutina que produce los 4 ficheros que `M50`/`M51` reimportan al año siguiente.

## `Rut_Lo.bas` — framework genérico de tablas

Caja de herramientas que el resto del proyecto usa para no reescribir lógica de `ListObject`:

| Rutina | Qué hace |
|---|---|
| `Rut_Lo_DataBodyRange_Copy` | Copia el `DataBodyRange` completo de una tabla origen a una destino, opcionalmente vaciando antes el destino |
| `Rut_Lo_DataBodyRange_Filtered_Copy` | Igual, pero solo con las filas **visibles** (filtradas) del origen; opcionalmente borra del origen las filas copiadas |
| `Rut_Lo_WrkSht_Preparar` | Muestra todas las filas/columnas ocultas y quita filtros de `.ListObjects(1)` de una hoja |
| `Rut_Lo_Sort` | Ordena una tabla por una columna (ascendente/descendente) |
| `Rut_Lo_Filtro` / `Rut_Lo_Filtros_Quitar` | Aplican o retiran el `AutoFilter` de una tabla |
| `Rut_Lo_DataBodyRange_Filter_y_DEL` | Filtra por 1-2 criterios y **borra** las filas visibles resultantes |

`Prog_DefCol*` (definición de columnas: nombre/orden/tipo/formato) **no** las consume `Rut_Lo.bas` — las usa una rutina distinta, `Rut_X_Format_LoData_LoDefCol` (`M90_Rut_Format_Colmns.bas`), y en la práctica solo `Prog_DefCol_BD` tiene consumidores reales (`M01`, `M50`); `Prog_DefCol1/2/3`, `Prog_DefCol_Aux`, `Prog_DefCol_InfRec_ACont` y `Prog_DefCol_Plan_Rsm` son catálogos sin ningún consumidor detectado en el código exportado.

## Menú contextual personalizado

`Rut__Right_Click_VBA.bas` sustituye el menú nativo de clic derecho en dos barras distintas de Excel (`"Cell"` y `"List Range Popup"`): oculta todas las opciones nativas dejando solo las propias (conversión Mayús/Minús/Propia, "Fill Red", insertar rectángulo), todas etiquetadas con un Tag propio para poder borrarlas limpiamente al cerrar el libro. La limpieza al cierre vive duplicada dentro de `Rut_ConfigExcel_RESTABLECER` (`M00_Ini_APP.bas`), que sí cuelga de `ThisWorkbook.Workbook_BeforeClose`.

## `Form_Menu` — menú de tareas auxiliares

UserForm que se autoajusta a pantalla completa vía `Application.Top/Left/Width/Height` + `Zoom` (nunca redimensiona controles a mano — el patrón "seguro" frente al bug de hit-test ya documentado en las notas globales de UserForms). Carga desde `Prog__Menú_Aux.ListObjects("Tb_Tareas")` una lista de tareas filtrada por usuario autorizado (o vacía = para todos) y, al pulsar "Ejecutar", despacha dinámicamente por nombre con `Application.Run` — una tarea puede encadenar varias rutinas separadas por `" + "`. Si el nombre no existe, captura el error y avisa con `MsgBox` en vez de reventar. También aloja los switches de depuración (Probando/DisplayAlerts/EnableEvents/WB_Deactivate), visibles solo para el usuario "Boss".

## `Form_MsgBox` — cuadro de mensaje propio

Título opcional (con o sin barra de ventana de Windows, vía API), icono seleccionable por control (`Lb_Img_Ask/Stop/Exclam/Ok/...`), marco rojo opcional, 1 a 3 botones configurables por cadena (`"Ok+Cancel+Stop"`, autoajustados en ancho) y alto autoajustado al mensaje. Se pilota con 4 variables `Public` (`MsgBx_Title`, `MsgBx_TitleBar`, `MsgBx_Msg`, `MsgBx_Answer`) declaradas en `M00_Ini_Var_APP.bas`.

## Las hojas `Prog_*` están vacías a propósito

Confirmado por lectura completa: las ~29 hojas de catálogo (`Prog_BD`, `Prog_Concept`, `Prog_TipoRec`, `Prog_Coef_Ret_VRI`, `Prog_Orgánicas`...) y las de configuración (`Prog__APP`, `Prog__Usuarios`, `Prog__Menú_Aux`) tienen el code-behind vacío (como mucho un `Option Explicit` suelto). El contenido funcional vive siempre en el `ListObject` de la hoja o en los módulos `M*`/`Rut_*` que la referencian — no busques lógica de negocio dentro de un `Prog_*.cls`. Los `Wk_*.cls` son la excepción real a este patrón: `Wk_TitP_Liquid` (y en menor medida sus hermanas `Wk_TitP_LIQx_PDF`, `Wk_TitP_LiqPDF`, `Wk_TitP_UNO`) sí contienen lógica de eventos activa.
