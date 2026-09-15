# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué es este proyecto

Libro Excel con VBA (`EP_202x-2x_BaseDatos-LIQ _V3.xlsm`, ~5 MB) que gestiona la **Liquidación de Enseñanzas Propias / Títulos Propios de la Universidad de Alicante** (EFP: Enseñanzas de Formación Permanente; CFC/AFC: Cursos y Actividades de Formación Complementaria). El código VBA vive exportado como texto plano en `VBA_Moduls/` (131 ficheros `.bas`/`.cls`/`.frm`/`.frx`); el `.xlsm` es el binario real donde se ejecuta y que contiene además los datos (tablas Excel).

`202x-2x` en el nombre del fichero es un placeholder genérico: `M79_Crear_WB_EP_CAcad.bas` (`Rut_Crear_WB_EFP_o_CFCyAFC_de_CAcad_Ant_o_Pos`) genera, a partir de esta plantilla, una copia real por combinación EFP/CFCyAFC × curso académico (`<Tipo>_<CursoAcad>_BaseDatos_Liq_<versión>.xlsm`) — este `.xlsm` es la plantilla/versión de programación, no un curso concreto.

No hay build/compilación: el ciclo de trabajo es editar los módulos en `VBA_Moduls/` (o directamente en el editor VBA del `.xlsm`) y mantener ambos sincronizados manualmente.

### Relación con "Precios Públicos - PPub"

Este libro comparte linaje de código con el proyecto hermano `../Precios Públicos - PPub/` (ver su `CLAUDE.md`): `M50_Inf_Cont_AE4x4.bas` lleva literalmente el comentario `'- M_410_Update_Lo_AE4` como referencia a la fase `M_410_Update_Lo_AE4` de PPub. Este libro **importa las 4 copias EFP/CFC de curso Ant/Pos** generadas por `M79` y genera los informes contables (`Inf_Cont_AE4x4`/`AE4x41`) y ficheros AE4x4/AE4x1 que después importa PPub en su pipeline 410–415. Antes de tocar la lógica de AE4/JI's/Coef_VRI en cualquiera de los dos libros, revisa el módulo equivalente en el otro.

## Encoding — reglas obligatorias antes de tocar `VBA_Moduls/`

Los ficheros `.bas`/`.cls`/`.frm` son **Windows-1252 (CP1252), no UTF-8**, con finales de línea **CRLF**. Editarlos con las herramientas Write/Edit estándar (que asumen UTF-8) corrompe silenciosamente tildes/ñ existentes a `U+FFFD`, incluso si el cambio en sí es solo ASCII.

- Nunca uses el Write tool ni el Edit tool directamente sobre estos ficheros.
- Para modificarlos, usa un script que lea/escriba explícitamente `cp1252` preservando CRLF (p. ej. Python: `open(p,'rb').read().decode('cp1252')` … `.encode('cp1252')`), localizando la línea a tocar por un **substring único** (nunca por conteo de espacios).
- Tras cualquier escritura, verifica a nivel de byte: 0 secuencias `0xEF 0xBF 0xBD`, sin BOM, sin LF sueltos (todo CRLF). Los `.frx` son binarios (blobs de controles de UserForm): no aplican estas comprobaciones de texto.

## Re-exportar los módulos VBA — el proyecto está protegido con contraseña

`VBProject.Protection = 1`. En el Office actual, abrir el libro por **Automation/COM** (`Workbooks.Open` + `.VBProject.VBComponents`) falla de inmediato con "No se puede ejecutar la operación porque el proyecto está protegido", **sin llegar a mostrar ningún diálogo de contraseña** aunque la ventana del editor VBA esté visible — el truco de rellenar el diálogo por `win32gui`/`win32api` (el que usa la app `Import_Export_VBA_Moduls`) ya no funciona en esta build de Office para proyectos protegidos.

Vía que sí funciona: el propio libro trae una macro de auto-exportación, **`Rut_VBA_Export_Moduls`** (en `Rut_VBA_Moduls_Export.bas`). Con el libro abierto normalmente en Excel (contraseña introducida a mano si el VBE la pide al expandir el árbol del proyecto), ejecútala desde el editor VBA (F5, o `Rut_VBA_Export_Moduls` en la ventana Inmediato). Exporta los 131 componentes a `[RutaLibro]\[NombreLibro]_VBA_Moduls\`; copia el resultado sobre `VBA_Moduls/` en este repo.

## Arquitectura

### Convención de nombres de módulos

- **`MNN_Descripción.bas`** — módulos de **flujo de proceso**, numerados por fase del pipeline (dos dígitos, sin ceros de relleno tipo PPub). Documentados con logs reales de ejecución en los propios comentarios de cabecera.
- **`MNN__________________.bas`** — módulos "separador" (uno por rango de decena: `M01`, `M10`, `M20`...). Solo `M01__________________.bas` tiene contenido real (documenta el pipeline de importación LSGES04 con logs reales); el resto están vacíos.
- **`M0999_*.bas`** — herramientas de depuración/desarrollo puntual, no forman parte del pipeline regular.
- **`Prog_*.cls`** — code-behind de **hojas de configuración/catálogo** (`VB_Name` = CodeName de la hoja), **vacío a propósito** salvo casos puntuales. El contenido funcional vive en el `ListObject` de esa hoja. Incluye `Prog_DefCol*` (definición de columnas para las rutinas genéricas `Rut_Lo_*`), catálogos de dominio (`Prog_Concept`, `Prog_TipoRec`, `Prog_CodActiv`, `Prog_Coef_Ret_VRI`, `Prog_Orgánicas`...) y `Prog__APP`/`Prog__Usuarios`/`Prog__Menú_Aux` (doble guion bajo) — configuración de la app, switches, usuario activo y tabla de tareas del menú auxiliar.
- **`Wk_*.cls`** — code-behind de las **hojas de trabajo visibles**, a diferencia de `Prog_*`/`Sht__*` estas **sí tienen lógica real**: `Wk_TitP_Liquid` (la única hoja visible de la app) implementa `Worksheet_SelectionChange`/`Worksheet_Change` como manejadores de "botones" — cada rango con nombre (`Liquid_Plan`, `Liquid_Filtro_*`, `Liquid_Num_JI_Emi`...) actúa como control al hacer clic o escribir en él. `Wk_Lista_Panes*`, `Wk_TitP_LIQx_PDF`, `Wk_TitP_UNO`, `Wk_Inf_Anulados` son hojas de trabajo auxiliares del mismo tipo.
- **`Sht__*.cls`** — code-behind de las **hojas de informe** (`Sht__Inf_*`, `Sht__Buffer`), igual de vacío que `Prog_*`: son hojas ocultas que solo reciben resultados generados por rutinas `M*`/`Rut_*`.
- **`Rut_*.bas`** — librería transversal de utilidades agrupada por objeto Excel: `Rut_Lo*` (ListObject: copiar/ordenar/filtrar/formatear — núcleo del framework de tablas), `Rut_WB`/`Rut_Wb_CopSegTimed_USB_HD` (Workbook, incluye copia de seguridad con marca de tiempo en USB+HD), `Rut_WS`/`Rut_Ws_*Stratistics` (Worksheet), `Rut_Hipervinculos`, `Rut_NEXE`, `Rut_Ranges`, `Rut__Right_Click_VBA` (menú contextual custom), `Rut_VBA_Moduls_Export` (ver arriba).
- **`Form_*.frm/.frx`** — UserForms: `Form_Menu` (menú de tareas auxiliares), `Form_Usuario` (login), `Form_MsgBox`/`Mensaje` (cuadros de mensaje con estilo propio).
- **`Módulo1.bas`, `Módulo3.bas`, `Módulo4.bas`, `Módulo5.bas`, `Hoja1.cls`, `Módulo_Filtro_Avanzado_Prueba.bas`** — residuales/plantilla por defecto de Excel o pruebas sueltas, sin renombrar; no forman parte de la arquitectura activa.

### Flujo de arranque (`ThisWorkbook.cls` → `Workbook_Open`)

1. Resetea el switch `SW_EnableEvents` y los switches de usuario (`SW_Boss`, `SW_Probando`).
2. `Rut_ConfigExcel_Establecer` (pantalla completa, oculta barra de fórmulas/estado/Ribbon nativo — no hay Ribbon custom, se oculta el nativo vía `ExecuteExcel4Macro "show.toolbar"`).
3. Carga `Lo_Tareas` desde `Prog__Menú_Aux.ListObjects("Tb_Tareas")` y llama `Rut_Usuario_Chg` (muestra `Form_Usuario`, login).
4. Si han pasado más de 7 días desde `APP_CopSeg_HD_Date`, lanza `Rut_WrkBooK_CopSegTimed_WB_HD` (backup con marca de tiempo en USB+HD; **nota**: `M90_CopSeg_USB_HD.bas` es una versión antigua íntegramente comentada — la activa es `Rut_Wb_CopSegTimed_USB_HD.bas`).
5. `RuT_Al_Abrir_WorkBook` (en `M00_Ini_APP.bas`): oculta todas las hojas salvo `Wk_TitP_Liquid` (`xlSheetVeryHidden`), la protege con `UserInterfaceOnly:=True`, recalcula la liquidación del plan/curso activo (`Rut_00_Liquid_TitProp`) e inmoviliza paneles.

No hay menú Ribbon ni UserForm de pantalla principal: la interacción diaria es **directamente sobre la hoja `Wk_TitP_Liquid`**, haciendo clic o escribiendo en celdas con nombre que actúan como controles (filtros por año/liquidación/recibo, añadir Núm. JI/AD/RDT/Coef_VRI/Orgánica a los registros visibles, alternar vista de tabla RDT). `Form_Menu` (tareas auxiliares) se abre desde el menú contextual personalizado (`Rut__Right_Click_VBA.bas` sustituye las opciones nativas del clic derecho por las propias).

- **`M00_Ini_APP.bas`** — entorno Excel: `Rut_ConfigExcel_Establecer/RESTABLECER`, `Rut_Off_Functions`/`Rut_On_Functions` (pausar cálculo/eventos durante procesos largos).
- **`M00_Ini_Var_APP.bas`** — sin lógica: es el **esquema de datos central**, define como `Public Const` el índice de columna de cada tabla clave (`BD_*` en `Prog_BD`, `CLiq_*` en la Liquidación, `CSol_*` en Solicitudes, `DefC_*`, `InfRec_*`). Referencia obligada antes de tocar cualquier rutina que lea/escriba columnas de una tabla.
- **`Mód_Menú_Usuario.bas`** — cambio de usuario (`Rut_Usuario_Chg`, muestra `Form_Usuario`).

### Pipeline de proceso (módulos `MNN_*`, por rango de decena)

| Rango | Fase |
|---|---|
| 00 | Arranque, config Excel, esquema de datos |
| 01–08 | Importación/depuración de **LSGES04** (exportación de recibos del sistema contable UA): importar → borrar registros no válidos/EFP-CFCyAFC de otro curso → gestionar duplicados → asignar cta. de ingreso → asignar concepto económico/tipo de enseñanza → asignar tipo de recibo (Emitido/EjeAnt/Añejo/Aplazado) → actualizar `Prog_BD` → actualizar tabla de coeficientes de retención VRI |
| 09 | Importar Solicitudes de Liquidación (`Sol_Liq`) |
| 10 | **`M10__Liquid_EP.bas`** — núcleo: genera la Liquidación por Plan/Curso Académico en `Wk_TitP_Liquid` (`Rut_00_Liquid_TitProp`), tabla RDT (Retención sobre Rendimientos del Trabajo) por Núm. RDT o Núm. Liquidación |
| 12 | Generar Liquidación en PDF |
| 15 | Exportar Liquidación a `.xlsx` |
| 20–22 | Informes EP para UXXI, Resumen de Títulos Propios |
| 31–33 | Listados de Planes (por Orgánica, Anulados, Devoluciones) |
| 38–39 | Cierre Contable de Planes + exportación, Resumen Planes por Curso Académico |
| 40–44 | Informe Contable de Recibos, añadir Núm. JI's al informe/a `Prog_BD`, exportar informe de acontecimientos por curso académico |
| 50–51 | Informes contables **AE4x4/AE4x41** (importa las 4 copias EFP/CFC Ant/Pos de `M79`), importar AE4x1 — ver [relación con PPub](#relación-con-precios-públicos---ppub) |
| 71–72 | Restituir `Prog_BD` desde una versión de trabajo anterior, borrar registros no válidos de `Prog_BD` |
| 79 | Crear copia anual del libro (`Rut_Crear_WB_EFP_o_CFCyAFC_de_CAcad_Ant_o_Pos`) |
| 80 | Enviar correo |
| 90 | Utilidades transversales del menú auxiliar (formato de columnas, rutinas varias) |

## Notas al modificar código

- Los `.cls` de `Prog_*`/`Sht__*` están vacíos a propósito: no busques lógica de negocio dentro, está en el `ListObject` de la hoja o en los módulos `M*`/`Rut_*` que la referencian. Los `.cls` de `Wk_*` son la excepción: ahí sí vive lógica real (manejadores de eventos de hoja).
- `M90_CopSeg_USB_HD.bas` es código muerto (comentado en su totalidad); la copia de seguridad activa es `Rut_Wb_CopSegTimed_USB_HD.bas`.
- Antes de escribir `Hoja.Range("NombreDefinido")`, comprueba en qué hoja vive realmente el nombre (Administrador de Nombres, Ctrl+F3) — el ámbito "Libro" de un nombre no exime de cualificar con la hoja física donde vive la celda.
- Antes de tocar la lógica de AE4/AE4x4/JI's/Coef_VRI, revisa también el módulo equivalente en `../Precios Públicos - PPub/VBA_Moduls/` (mismo linaje de código, ver arriba).
