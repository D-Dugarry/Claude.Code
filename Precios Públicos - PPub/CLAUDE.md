# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué es este proyecto

Libro Excel con VBA (`PPub_BDatos_2026.xlsm`, ~68 MB) que gestiona los **recibos de Precios Públicos de la Universidad de Alicante** (matrículas, tasas académicas/administrativas, Enseñanzas Propias, cuotas INSS...). El código VBA vive exportado como texto plano en `VBA_Moduls/` (120 ficheros `.bas`/`.cls`/`.frm`/`.frx`); el `.xlsm` es el binario real donde se ejecuta y que contiene además los datos (tablas Excel).

No hay build/compilación: el ciclo de trabajo es editar los módulos en `VBA_Moduls/` (o directamente en el editor VBA del `.xlsm`) y mantener ambos sincronizados manualmente (import/export de módulos vía el editor VBA, Alt+F11 → clic derecho → Exportar/Importar archivo).

## Encoding — reglas obligatorias antes de tocar `VBA_Moduls/`

Los ficheros `.bas`/`.cls`/`.frm` son **Windows-1252 (CP1252), no UTF-8**, con finales de línea **CRLF**. Editarlos con las herramientas Write/Edit estándar (que asumen UTF-8) corrompe silenciosamente tildes/ñ existentes a `U+FFFD`, incluso si el cambio en sí es solo ASCII.

- Nunca uses el Write tool ni el Edit tool directamente sobre estos ficheros.
- Para modificarlos, usa un script que lea/escriba explícitamente `cp1252` preservando CRLF (p. ej. Python: `open(p,'rb').read().decode('cp1252')` … `.encode('cp1252')`), localizando la línea a tocar por un **substring único** (nunca por conteo de espacios).
- Tras cualquier escritura, verifica a nivel de byte: 0 secuencias `0xEF 0xBF 0xBD`, sin BOM, sin LF sueltos (todo CRLF).
- Anclajes por substring: cuidado con caracteres que "parecen" otros (p. ej. el ordinal `ª` es `\xAA`, no una `á`) — verifica el byte real en vez de asumirlo por cómo se ve.

## Arquitectura

### Convención de nombres de módulos

- **`M_NNN_Descripción.bas`** — módulos de **flujo de proceso**, numerados por rango de cientos (fase del pipeline anual) y decenas/unidades (paso dentro de la fase). Cada uno suele tener un wrapper `Call_Rut_...` (gestiona errores, muestra `Form_Running_Rut`) y la rutina de negocio real `Rut_...`.
- **`M_D___________.bas`** (p. ej. `M_1____`, `M_2____`, `M_9____`) — módulos "separador": no son código ejecutable, solo comentarios que documentan en detalle cada fase del pipeline (incluso con logs reales de ejecución). Son la mejor fuente para entender el negocio de cada rango de cientos.
- **`Prog_*.cls`** — code-behind de **hojas de configuración/catálogo** (`VB_Name` = CodeName de la hoja), casi siempre vacío salvo `Worksheet_Activate/Deactivate`. El contenido funcional vive en el `ListObject` (tabla Excel) de esa hoja, no en la clase. Dos subtipos:
  - `Prog_DefCol_*` → definición de columnas (nombre, formato, ancho, orden, ocultar/proteger) usada por las rutinas genéricas de formateo/import de `Rut_Lo_*`.
  - Catálogos de dominio (`Prog_Concept`, `Prog_ClasifEco`, `Prog_TipoRec`, `Prog_Bco`, `Prog_Orgánicas`, `Prog_CodActiv`, `Prog_Tipo_EPE`, `Prog_RetVRI`, `Prog_TitOf_Plazos`...) → tablas maestras de lookup.
  - `Prog__APP`, `Prog__Usuarios`, `Prog__Menú_Aux` (doble guion bajo) → configuración de la aplicación: switches (`SW_*`), rutas, usuario activo, textos de `Form_MsgBox`, tabla de tareas del menú/Ribbon con visibilidad por usuario.
- **`Sht__*.cls`** — mismo patrón que `Prog_*` (wrapper vacío de hoja), pero para las **hojas de datos**: `Sht__BD` (tabla maestra de recibos), `Sht__BD_Ant` (versión/ejercicio anterior), `Sht__BD_Dupl`, `Sht__BD_ErrDate`, `Sht__BD_RegAnul`, `Sht__BD_AE4x4`, `Sht__BD_INSS`, `Sht__BD_JIs_AE4`, `Sht__BD_IAdm_CAcadAnt`, hojas `Sht__Inf_*` (informes), `Sht__Buffer` (hoja de trabajo temporal). Se referencian por CodeName (`Sht__BD.ListObjects(1)`) para no depender del nombre visible de la pestaña.
- **`Rut_*.bas`** (y variante `RuT_*.bas`) — librería transversal de utilidades reutilizables agrupada por objeto Excel: `Rut_Wb*` (Workbook — incluye `Rut_Wb_CopSegTimed_USB_HD.bas`, copias de seguridad con marca de tiempo en USB+disco local), `Rut_Ws*` (Worksheet), `Rut_Lo*` (ListObject: import/export/formato/orden/duplicados — el núcleo del framework de tablas), `Rut_UserForms`, `Rut_Hipervinculos`, `Rut_Filtro_Avanzado_VBA`, `Rut_File_Folder_NEXE`.
- **`Form_*.frm/.frx`** — UserForms.
- **`Hoja1.cls`, `Hoja4.cls`, `Hoja17.cls`, `Módulo1.bas`** — residuales/plantilla por defecto de Excel, sin renombrar; no forman parte de la arquitectura activa.

### Flujo de arranque (`ThisWorkbook.cls` → `Workbook_Open`)

1. Minimiza otras ventanas/instancias de Excel, oculta la app y desactiva pantalla.
2. Activa eventos, fija switches iniciales (`SW_Events`, `SW_Test`, `SW_RightClickMenú_Visible`, `SW_WB_Deactivate`) en `Prog__APP`.
3. Muestra `Form_Usuario` (login) y carga datos de tareas (`RuT_Load_Task_Data`).
4. Oculta el menú contextual nativo dejando solo las opciones custom (`Rut_Context_Buttons_Hide`) y oculta el Ribbon nativo.
5. Activa `Sht__BD`, quita filtros, muestra todas las hojas, inmoviliza paneles.
6. Ejecuta `RuT_Al_Abrir_WorkBook` (en `M_000_Ini_APP.bas`) y `Rut_Menú_ShowAll` (muestra el Ribbon custom).

No hay un menú de UserForm como pantalla principal: la interacción diaria es vía **Ribbon personalizado (RibbonX, backend en `M___RibbonUI.bas`)**, con dos pestañas: `TabUserMenu` (import/export de BD, importaciones LSGES04/LSace06/AE4, recalcular tablas JIs, exportar hoja, mostrar/ocultar columnas, navegación entre hojas) y `TabProgMenu` (solo visible para el usuario "Boss": herramientas de depuración, proteger hoja, refrescar ribbon, auditar hipervínculos). Los UserForms se reservan para login, progreso y diálogos puntuales.

- **`M_000_Ini_APP.bas`** — rutinas de entorno Excel: `Rut_ConfigExcel_Establecer/RESTABLECER`, `Rut_Off_Functions`/`Rut_On_Functions` (pausar cálculo/eventos durante procesos largos), `Rut_Context_Buttons_Hide/Restore`, `Rut_Menú_HideAll/ShowAll/ShowAll_Short`.
- **`M_000_Ini_Var_APP.bas`** — sin lógica: es el **esquema de datos central**. Define como `Public Const` el índice de columna de cada tabla clave (`BD_*`, `AE4_*`, `LSace06_*`, `EPplazos_*`, `JIs_*`, `InfRec_*`, `DR_*`, `DefCol_*`, `Task_*`) más variables públicas de estado. Es la referencia obligada antes de tocar cualquier rutina que lea/escriba columnas de una tabla.
- **`M_000_Menú_Usuario.bas`** — cambio de usuario (`Rut_Usuario_Chg`) y filtrado de qué botones del Ribbon/menú son visibles según usuario activo y hoja activa (`Rut_Filtrar_Tareas`, lee `Prog__Menú_Aux`).

### Pipeline de proceso anual (módulos `M_NNN_*`, por rango de cientos)

El negocio central es un ETL/contable anual sobre la tabla maestra `Sht__BD` (recibos):

| Rango | Fase |
|---|---|
| 000 | Arranque, config Excel, esquema de datos, menú de usuario |
| 90 | Utilidades transversales de la app (protección de hojas, ejecución genérica de tareas, carga/guardado de datos de tarea) |
| 110–118 | Importación/depuración de **LSGES04** (Año Contable actual): importar → borrar registros no válidos → gestionar duplicados → asignar concepto económico/año vto. → clasificar recibos (Emitido/Aplazado/EjeAnt/Añejo/ADxAplz) → asignar Importe Académico/Administrativo → actualizar BDatos con BDatos_Ant |
| 130 | Generación del Informe de Recibos (Inf_Recibos/TIO); añadir números de JI's al informe y a BDatos |
| 180 | Restituir datos de `Sht__BD` desde una versión anterior del Excel (backup/recovery manual) |
| 193/195 | Importar datos de un Wb de mes anterior / añadir datos UXXI a BDatos |
| 210–215 | Mismo proceso que 110–118 pero para el **Curso Académico Anterior** (matrículas del curso previo cobradas en el año contable actual) |
| 310–315 | Importación de **LSace06**, identificación de recibos con cuota **INSS** y copia a BD |
| 410–415 | Importación de **AE4** (Enseñanzas Propias: EFP/CFC/AFC/TUP...), ficheros AE4x4/AE4x1, asignación de Importe Admin. y copia a BD |
| 510–520 | Cálculo de **JI's** (Justificantes de Ingreso, contabilidad) — de AE4 y de PPub genérico "1303" |
| 600–602 | Informes finales (Informe EP/RSm) y cierre contable por Planes |

Los módulos separadores `M_D_____________` de cada rango documentan el detalle exacto de la fase con logs reales — consúltalos antes de modificar el pipeline correspondiente.

## Notas al modificar código

- Los `.cls` de `Prog_*`/`Sht__*` están vacíos a propósito: no busques lógica de negocio dentro, está en el `ListObject` de la hoja o en los módulos `M_*`/`Rut_*` que la referencian.
- Antes de escribir `Hoja.Range("NombreDefinido")`, comprueba en qué hoja vive realmente el nombre (Administrador de Nombres, Ctrl+F3) — el ámbito "Libro" de un nombre no exime de cualificar con la hoja física donde vive la celda.
- `Rut_Filtro_Avanzado_VBA.bas` tiene referencias a variables no declaradas (`Ws`, `TablaDatos`): parece código en desarrollo/borrador, no confirmado como activo.
