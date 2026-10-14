# Plan: `Inf_Conta_TIO`, el Informe de Recibos sin el INSS

**Proyecto:** `PPub_BDatos_2026.xlsm` · **Fecha:** 2026-10-13 (decisiones, 2026-10-14) · **Afecta a:**
`M_140_Gen_Inf_Conta_TIO`, `M_141_Anadir_Num_JIs_a_BDatos` y `M_142_Mail_Inf_Conta_TIO` (nuevos); `M_130`, `M_133` y
`M_135` (pasan a ser el motor común, con el mismo resultado de hoy); `M_000_Ini_Var_APP` (constantes `InfCon_*`);
`M___RibbonUI` (callbacks); `customUI14.xml` (dos `screentip` pasan a `getScreentip`); `Lo_RibbonUI` (seis filas);
`ConfigAPP` y `Lo_ListObjAPP` (nombres y una fila); documentación.

**Etiquetas:** ✅ **Hecho** · 🧪 **Aplicado, falta probarlo en Excel** · ⏳ **Pendiente**

## Estado (2026-10-14 12:55): ✅ TERMINADO. Fases 1-4 en el libro y probadas por el usuario en Excel ("funciona todo")

- ✅ **Fases 1-3 en el repo** (`M_000_Ini_Var_APP`, `M_130`, `M_133`, `M_135`, `Rut_File_Folder_NEXE`, `M___RibbonUI` y
  los nuevos `M_140`, `M_141`, `M_142`), **probadas por COM** en tres copias del libro de las 09:26 (sección 5.4).
- ✅ **Módulos importados por el usuario** (guardado de las 11:51): el código de los 9 módulos del libro es el del repo
  (comprobado con `olevba`; solo cambian mayúsculas, que el editor de VBA unifica en todo el proyecto).
- ✅ **Fase 4 aplicada al libro (11:56)**, con Excel y el Custom UI Editor cerrados (confirmado por el usuario): XML,
  6 filas de `Lo_RibbonUI` y la fila de `Lo_ListObjAPP`. Copia previa:
  `CopiaSeguridad/PPub_BDatos_2026_antes_Inf_Conta_TIO_Ribbon_2026-10-14_1155.xlsm`. Antes de sustituirlo se comprobó la
  copia: BDatos, BD_Dupl, BD_ErrDate y los dos informes idénticos (valores y formatos); el mismo código VBA (119
  módulos); `Validar_Ribbon.py`, 0 errores. Excel, al guardar, recalculó el alto automático de las filas 6 y 16 de
  `Inf_Conta_TIO` (28 → 15,5): el ábaco los vuelve a ajustar.
- ✅ **Probado por el usuario en Excel** (2026-10-14, «funciona todo»): los botones de las dos hojas (5.3).
- ✅ `App_Last_Calc_Inf_Conta_TIO` creado; `App_Last_Calc_Inf_Recibos` y `App_Mail_Inf_Recibos` renombrados con `_TIO`
  (decisión 9); «Red:» lo entiende el código (decisión 10, `Fnc_Ruta_Red`). Guardado a las 09:26.
- ✅ Decisiones de la sección 7, contestadas por el usuario.
- ✅ **En el libro (guardado de las 08:00):** hoja `Inf_Conta_TIO` (CodeName `Sht__Inf_Conta_TIO`), tabla
  `Lo_Inf_Conta_TIO`, las dos formas borradas e `Inf_Recibos_TIO (2)` borrada.
- ✅ **Nombres de `ConfigAPP`** (guardado de las 09:26): `App_Mail_Inf_Recibos_TIO` (`B7`) y `App_Mail_Inf_Conta_TIO`
  (`B8`), los dos con felipe.martinez@gcloud.ua.es; `App_RutaRed_Inf_Conta_TIO` (`C11`) y `App_RutaRed_Inf_Recibos_TIO`
  (`C12`), las dos con la misma carpeta (`Red:\…\Tesoreria_Felipe\Aa_TESORERIA_FELIPE\DOCUMENTOS DUGARRY`);
  `App_Last_Calc_Inf_Conta_TIO` (`B45`) y `App_Last_Calc_Inf_Recibos_TIO` (`B46`). Ya no existen `App_Ruta_Inf_Recibos`,
  `App_RutaRed_Inf_Recibos`, `App_Last_Calc_Inf_Recibos` ni `App_Mail_Inf_Recibos`.
- ⚠️ **Hasta que se importen los módulos nuevos, el ábaco y el correo de `Inf_Recibos_TIO` fallan en el libro**, porque
  el código de hoy busca los nombres antiguos: el ábaco calcula, pero da error al anotar la fecha (`App_Last_Calc_Inf_Recibos`), y el
  correo no encuentra los destinatarios ni la carpeta.

## 1. Qué se pide (2026-10-13)

1. Una variante de `Inf_Recibos_TIO` que se llama **`Inf_Conta_TIO`**, con dos diferencias:
   - **No tiene las tres columnas del INSS** (`INSS Emi.`, `INSS Cob.` e `INSS Pdte`).
   - **Emitido, Cobrado y Pendiente (E, F y G) son la suma** de los importes administrativos y los académicos.
2. Tres módulos nuevos, en el rango 140 (el separador `M_14___…` ya lo creó el usuario): `M_140_Gen_Inf_Conta_TIO`,
   `M_141_Anadir_Num_JIs_a_BDatos` y `M_142_Mail_Inf_Conta_TIO`.
3. **Los mismos botones del Ribbon:** el grupo del ábaco ve en qué hoja está y hace lo que toca en cada una.

## 2. Cómo está

### 2.1 La hoja (guardado del 2026-10-14 08:00, leído como zip)

- `Inf_Conta_TIO`, CodeName `Sht__Inf_Conta_TIO`, **tabla `Lo_Inf_Conta_TIO`** en `B4:AG20`, con fila de totales
  (`SUBTOTAL` sobre su propia tabla) y **sin formas**. Es una copia de `Inf_Recibos_TIO` sin las columnas del INSS: sus
  valores son los de la copia (`B2` «Informe Recibos 2026», `D2` del 13-oct a las 13:27, y E:G con el INSS dentro).
- Tiene 32 columnas: las 35 de `Inf_Recibos_TIO` menos las 7-9.

  | Columnas de la tabla | `Inf_Recibos_TIO` | `Inf_Conta_TIO` |
  |---|---|---|
  | Tipo Rec., Enseñanza, Concepto, Emitido, Cobrado, Pendiente | 1-6 (B:G) | 1-6 (B:G) |
  | INSS Emi., Cob., Pdte. | 7-9 (H:J) | — |
  | Cta. Adm. … Acad. Pdte. | 10-17 (K:R) | 7-14 (H:O) |
  | 6 columnas de JI's y AD's | 18-23 (S:X) | 15-20 (P:U) |
  | Descripción para el JI, Descripción Mov. Contable | 24-25 (Y:Z) | 21-22 (V:W) |
  | Importes académicos por curso | 26-35 (AA:AJ) | 23-32 (X:AG) |

  Desde Cta. Adm., **todas las columnas se corren 3 a la izquierda**.
- Títulos de grupo combinados en la fila 2: `E2:G2`, `H2:K2`, `L2:O2`, `P2:U2` (UXXI-EC) y `X2:AG2`. Debajo de la
  tabla, la leyenda de tipos de recibo (U:V) y el bloque EURLE (`E31:G31`, `H31:K31`, `L31:O31`), también de la copia.
- **Sin fila** en `Lo_ListObjAPP` (su fila de totales queda «libre») ni en `Prog_Hojas_Names`, y sin nombre
  `App_Last_Calc_…`. El menú de hojas del Ribbon ya la lista sola (lista todos los CodeName `Sht__Inf*`).

### 2.2 Lo que hoy está atado a `Inf_Recibos_TIO`

| Dónde | Qué |
|---|---|
| `M_130` | La hoja (`Sht__Inf_Recibos_TIO`), las constantes `InfRec_*`, la cabecera del bloque EURLE (`Range("e2:r2")`, fija), el título de `B2` y `App_Last_Calc_Inf_Recibos`. Además, **lee de las celdas del INSS** para calcular Adm. y Acad. |
| `M_133` | La hoja y las constantes `InfRec_TipRec`, `InfRec_ConcptEco` e `InfRec_JI_Emi_Adm`. |
| `M_135` | La hoja, la última columna del adjunto (`InfRec_Descrip_Contab`), el nombre del `.xlsx`, `App_Last_Calc_Inf_Recibos`, `App_Mail_Inf_Recibos`, `App_Ruta_Inf_Recibos` (renombrado por el usuario) y los textos (pregunta, asunto y cuerpo, que habla del INSS). |
| `M___RibbonUI` | `OnAct_ReCalculate_ListObj_Sht` elige la rutina por CodeName; `GetLbl_ReCalculate_ListObj_Sht` tiene un caso especial para `Inf_Recibos_TIO`; `OnAct_JIs_Inf_a_BD` y `OnAct_Mail_Inf_Recibos` llaman a una rutina fija. El rótulo del grupo ya sale de la hoja («Calc_» y el nombre). |
| `customUI14.xml` | Los dos botones tienen un `screentip` fijo: «Añadir a BDatos los JI's del informe.» y «Enviar por correo el Informe de Recibos.». |
| `Lo_RibbonUI` | `SheetsNames` = `Inf_Recibos_TIO` en las filas de los tres botones y del grupo (con `Inf_RSm` en las del ábaco). |
| `Rut_Lo` | `Lo_Totales_Hojas` (la regla por defecto de la fila de totales; manda `Lo_ListObjAPP`). |

### 2.3 Las cuentas

En `Inf_Recibos_TIO`, **Emitido = Imp.Rec. = INSS + Adm. + Acad.** `M_130` calcula Adm. como `Rec_Imp_Adm` menos el INSS,
y Acad. como Imp.Rec. menos los dos. En `Inf_Conta_TIO`, **Emitido = Adm. + Acad.**, es decir, el de `Inf_Recibos_TIO`
menos el INSS. Comprobado con los totales de la captura del 13-oct a las 23:14; cuadra al céntimo:

| | `Inf_Recibos_TIO` | − INSS | = `Inf_Conta_TIO` | = Adm. + Acad. |
|---|---|---|---|---|
| Emitido | 26.940.388,22 | 26.312,16 | **26.914.076,06** | 1.870.463,19 + 25.043.612,87 |
| Cobrado | 15.118.652,76 | 24.828,16 | **15.093.824,60** | 1.634.118,72 + 13.459.705,88 |
| Pendiente | 11.821.735,46 | 1.484,00 | **11.820.251,46** | 236.344,47 + 11.583.906,99 |

Con los mismos datos, esos serán los totales de control de la prueba (sección 5). **Todo lo demás** (Adm., Acad., JI's,
descripciones e importes por curso) **es igual en los dos informes.**

Dos detalles:

- En las filas **ADM y TNCT**, `M_130` pone Adm. = Emitido entero y deja Acad. vacía. Si una de esas filas llevara INSS,
  iría dentro de Adm. en los dos informes. Hoy ninguna lleva.
- `M_130` borra al final las filas con Emitido a cero. En `Inf_Conta_TIO`, el borrado mirará el Emitido **sin INSS**, así
  que una fila que solo tuviera INSS desaparecería (en el otro informe se queda). Hoy no hay ninguna así.

## 3. Cómo se hace: un motor común y tres módulos finos (decisión 1)

### 3.1 La idea

`M_130`, `M_133` y `M_135` siguen haciendo lo mismo, pero reciben **qué informe** tratan: la hoja, dónde está cada
columna, sus nombres de Excel y sus textos. `M_140`, `M_141` y `M_142` solo guardan lo propio de `Inf_Conta_TIO` y
llaman al motor. Así no se duplican las sumas en memoria, las reglas de los JI's ni la copia a `.xlsx`, que ya están
probadas con doble ejecución.

- **Mapa de columnas.** Un array `Col(1 To 35)` indexado por las constantes `InfRec_*`: dice en qué columna de **este**
  informe está cada dato (0 = no la tiene). Para `Inf_Recibos_TIO`, `Col(i) = i`. Para `Inf_Conta_TIO`, sale de las
  constantes nuevas **`InfCon_*`** (1-32), que van en `M_000_Ini_Var_APP` junto a las `InfRec_*`, y las del INSS valen 0.
  El motor escribe `RwJI.Range(Def.Col(InfRec_Adm_Emi))` donde hoy escribe `RwJI.Range(InfRec_Adm_Emi)`.
- **Descripción del informe:** `Public Type T_Inf_Def` en `M_130`. Lleva:
  - la hoja, el mapa de columnas y si lleva INSS;
  - el nombre del informe para los textos («Informe de Recibos» / «Informe Conta») y el título de `B2`;
  - los nombres de Excel de la fecha del cálculo, de los destinatarios y de la carpeta de red;
  - el nombre del `.xlsx`, y el asunto y el cuerpo del correo.

  La rellenan `Rut_Inf_Def_Recibos` (`M_130`) y `Rut_Inf_Def_Conta` (`M_140`). Una tercera, `Fnc_Inf_Def_Hoja(Ws, Def)`,
  rellena la de la hoja que se le pase y devuelve `False` si no es ninguna de las dos: es la que usan los callbacks.
- **Las rutinas públicas de hoy no cambian de nombre**, así que los callbacks, `Lo_RibbonUI` y lo que las llame siguen
  igual. Cada una pasa a ser una línea que llama al motor:

  | Módulo | Rutina pública (botón) | Llama al motor |
  |---|---|---|
  | `M_130` | `Rut_Recalcular_Tabla_Inf_Recibos` (sin cambios) | `Rut_Inf_Calcular(Def)` (el cuerpo de hoy) |
  | `M_140` | `Rut_Recalcular_Tabla_Inf_Conta` (nueva) | `Rut_Inf_Calcular(Def)` |
  | `M_133` | `Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos` (sin cambios) | `Rut_Inf_JIs_a_BDatos(Def)` |
  | `M_141` | `Rut_Anadir_a_BDatos_JIs_de_Inf_Conta` (nueva) | `Rut_Inf_JIs_a_BDatos(Def)` |
  | `M_135` | `Rut_Mail_Inf_Recibos` y `Rut_Inf_Recibos_Probar_Xlsx` (sin cambios) | `Rut_Mail_Inf(Def)` y `Fnc_Inf_A_Xlsx(Def, ErrMsg)` |
  | `M_142` | `Rut_Mail_Inf_Conta` y `Rut_Inf_Conta_Probar_Xlsx` (nuevas) | Los mismos |

  Las funciones auxiliares de `M_135` (ruta libre, formas con macro, nombres que sobran, fecha del Robot, carpeta de red)
  siguen privadas: las usa el motor, que vive en el mismo módulo.

### 3.2 Nombres de Excel de cada informe

| Para qué | `Inf_Recibos_TIO` | `Inf_Conta_TIO` |
|---|---|---|
| Fecha del último cálculo (rótulo del ábaco y pregunta del correo) | `App_Last_Calc_Inf_Recibos` (`B45`). Propuesta: renombrarlo a `App_Last_Calc_Inf_Recibos_TIO` (decisión 9) | **`App_Last_Calc_Inf_Conta_TIO`, por crear** |
| Destinatarios del correo | `App_Mail_Inf_Recibos` (`B7`). Propuesta: `App_Mail_Inf_Recibos_TIO` (decisión 9) | `App_Mail_Inf_Conta_TIO` (`B8`) |
| Carpeta de red | `App_RutaRed_Inf_Recibos_TIO` (`C12`; el antiguo `App_Ruta_Inf_Recibos`) | `App_RutaRed_Inf_Conta_TIO` (`C11`) |

Con `App_Last_Calc_<nombre de la hoja>`, el rótulo del ábaco funciona sin casos especiales: es lo que
`GetLbl_ReCalculate_ListObj_Sht` busca por defecto.

### 3.3 Cálculo: `M_130` y `M_140`

- Todas las columnas pasan por el mapa. Las del INSS solo se escriben si el informe las tiene.
- **Los importes se calculan en variables** (desde `Sumas`) y luego se escriben. Hoy Adm. y Acad. se calculan leyendo
  las celdas del INSS ya escritas, que en `Inf_Conta_TIO` no existen. La cuenta es la misma, así que `Inf_Recibos_TIO`
  da lo mismo que hoy.
- **En `Inf_Conta_TIO`, E/F/G = Adm. + Acad. de cada fila, ya redondeados**, después del redondeo a 2 decimales y antes
  de borrar las filas sin Emitido. Así cuadran al céntimo con la suma de las columnas. En `Inf_Recibos_TIO`, Emitido sigue
  siendo Imp.Rec.
- **Cabecera del bloque EURLE:** de la columna de Emitido a la de Acad. Pdte. de la fila 2, según el mapa (`E2:R2` en un
  informe y `E2:O2` en el otro), en lugar del `Range("e2:r2")` fijo.
- El título de `B2` («Informe Conta 2026»), la fecha del cálculo, los rótulos de las cabeceras de JI's y de «Año
  Contable…», la leyenda (U:V en `Inf_Conta_TIO`) y el aviso de error salen de la descripción del informe.
- **`M_140`:** `Rut_Recalcular_Tabla_Inf_Conta` y `Rut_Inf_Def_Conta`.

### 3.4 JI's del informe a BDatos: `M_133` y `M_141`

- El motor recibe la hoja y las columnas de `Tipo_Rec`, concepto y primer JI. **Los dos informes escriben en las mismas
  6 columnas de BDatos (AR:AW)**, en el mismo orden que sus columnas de JI's:

  | BDatos (constante) | Cabecera en BDatos | Columna en los informes (`Inf_Recibos_TIO` / `Inf_Conta_TIO`) |
  |---|---|---|
  | AR (`BD_JI_Emi_Adm`, 44) | JI_Emi Adm | S / P: «JI-2025 o JI-2026, 1303.00, Adm.» (Cta. Aplic. 7400) |
  | AS (`BD_JI_Emi_Acad`, 45) | JI_Emi Acad | T / Q: «JI-2025 o JI-2026, 1310.xx 00/01/02, Acad.» (7401) |
  | AT (`BD_AD_Emi_Adm`, 46) | AD_Emi Adm | U / R: «AD-2026, 1303.00, Adm.» (7400) |
  | AU (`BD_AD_Emi_Acad`, 47) | AD_Emi Acad | V / S: «AD-2026, 1310.xx 00/01/02, Acad.» (7401) |
  | AV (`BD_JI_443_Adm`, 48) | JI_443 Adm | W / T: «JI-2026, 1303.00, Adm.» (4430) |
  | AW (`BD_JI_443_Acad`, 49) | JI_443 Acad | X / U: «JI-2026, 1310.xx 00/01/02, Acad.» (4430) |

- Las reglas no cambian: solo filas con algún JI; en ellas, una celda vacía o con «·-·» vacía esa columna en sus
  recibos; y se salta el 311.
- El mensaje de `App_Task_Inf` dice de qué informe salen los JI's.
- Los JI's de BDatos llegan a cada informe al recalcularlo (`M_130`/`M_140`, en la pasada en memoria), igual que hoy.

### 3.5 Correo: `M_135` y `M_142`

- El motor recibe la hoja, la última columna del adjunto (`Descrip_Contab`: **B:Z** en un informe y **B:W** en el otro),
  el nombre del fichero (**`Inf_Conta_TIO (datos del AAAA-MM-DD).xlsx`**), los nombres de 3.2 y los textos.
- **Asunto:** «PPub - Informe Conta, año contable 2026, cierre 31/12/2026».
- **Cuerpo:** «Adjunto el **Informe Conta** de Precios Públicos del año contable 2026 (cursos 2025-26 y 2026-27): importes
  emitidos, cobrados y pendientes por tipo de recibo y concepto, desglosados en administrativos y académicos, **sin el
  seguro obligatorio (INSS)**, con los justificantes de ingreso (JI) y la descripción de cada movimiento contable. Las
  filas de EURLE van aparte, debajo de la tabla.» Y debajo, como hoy, la fecha de descarga, el cierre y el último
  cálculo.
- Las preguntas y el informe del formulario dicen «Informe Conta» donde hoy dicen «Informe de Recibos». El resto del
  flujo es el mismo: modo prueba, «¡ ENVÍO REAL !» con No por defecto y copia en `<carpeta>\<año>\`.

### 3.6 Ribbon: los mismos botones

- **`M___RibbonUI`:**
  - `OnAct_ReCalculate_ListObj_Sht`: un `Case Sht__Inf_Conta_TIO.CodeName` → `Rut_Recalcular_Tabla_Inf_Conta`.
  - `GetLbl_ReCalculate_ListObj_Sht`: fuera el caso especial de `Inf_Recibos_TIO` si se renombra
    `App_Last_Calc_Inf_Recibos` (decisión 9).
  - `OnAct_JIs_Inf_a_BD` y `OnAct_Mail_Inf_Recibos`: un `Select Case ActiveSheet.CodeName` que llama a la rutina de cada
    hoja; el título del formulario de progreso, también según la hoja.
  - **Nuevo `GetScrTip_Inf`** (decisión 8): el `screentip` de los dos botones según la hoja activa. Por ejemplo,
    «Enviar por correo el Informe Conta.» en `Inf_Conta_TIO`. Toma el nombre del informe de `Fnc_Inf_Def_Hoja`. Al
    cambiar de hoja se vuelve a pedir, porque el `Worksheet_Activate` de las dos hojas llama a `RefreshRibbon`.
- **`customUI14.xml`:** en `BtnJIs_Inf_a_BD` y `BtnMail_Inf_Recibos`, `screentip="…"` pasa a
  `getScreentip="GetScrTip_Inf"`. ⚠️ **Aviso antes y después**: con Excel y el Custom UI Editor cerrados.
- **`Lo_RibbonUI`:** en las filas `ReCalculate_ListObj_Sht_Group`, `ReCalculate_ListObj_Sht`, `JIs_Inf_a_BD`,
  `Mail_Inf_Recibos` y `ExportWorkSheet_Group`, añadir `Inf_Conta_TIO` a `SheetsNames`. En las de los botones, además,
  las rutinas nuevas en `Nom_Rut` y textos que valgan para los dos informes en `Descripion` (el supertip). Mismos
  usuarios. Si el botón de exportar tiene que verse también para todos ellos, la fila `ExportWorkSheet` (hoy solo
  `Boss`).
- El informe de la última ejecución (`Informe_Rut`) es uno por botón: enseñará el de la última hoja en que se usó. El
  texto dice de qué informe es.
- **Orden al importar:** `M___RibbonUI` nombra `Sht__Inf_Conta_TIO` y las rutinas de `M_140`-`M_142`. No compila hasta
  que estén importados los módulos nuevos (la hoja con ese CodeName ya está guardada).

### 3.7 El libro

- ✅ Hoja `Inf_Conta_TIO` / `Sht__Inf_Conta_TIO`, tabla `Lo_Inf_Conta_TIO` y formas borradas (usuario, 2026-10-14).
- ✅ Nombres de 3.2, creados y renombrados por el usuario (guardado de las 09:26).
- 🧪 **Probado en una copia, falta aplicarlo al libro** (fase 4):
  - **`Lo_RibbonUI`:** `Inf_Conta_TIO` en `SheetsNames` de `ReCalculate_ListObj_Sht_Group`, `ReCalculate_ListObj_Sht`,
    `JIs_Inf_a_BD`, `Mail_Inf_Recibos`, `ExportWorkSheet_Group` y `ExportWorkSheet`; las rutinas nuevas en `Nom_Rut` y
    supertips que valen para los dos informes.
  - **`Lo_ListObjAPP`:** fila de `Lo_Inf_Conta_TIO` detrás de la de `Inf_Recibos_TIO`, con `Sw_TRow_Hide` = FALSO (fila de
    totales visible) y `Sw_Col_Hide` vacío (sin mostrar/ocultar columnas, decisión 7).
  - **`customUI14.xml`:** `getScreentip="GetScrTip_Inf"` en los dos botones y el comentario del grupo.

  Se hace con `f4_libro.ps1` (COM, Excel aparte, eventos y macros apagados) y `f4_xml.py` (zip reconstruido: de 330
  entradas solo cambia `customUI/customUI14.xml`). En la copia, `Validar_Ribbon.py`: 0 errores, los 5 avisos de antes. El
  libro guardado por COM pesa menos (53 → 44 MB) solo por la compresión: el contenido sin comprimir es el mismo.

## 4. Fases

1. ✅ **(usuario)** Decisiones 9 y 10; `App_Last_Calc_Inf_Conta_TIO` creado; guardado a las 09:26.
2. ✅ **(yo) Fase 1, cálculo:** constantes `InfCon_*` e `InfRec_N_Col` (`M_000_Ini_Var_APP`), `T_Inf_Def`,
   `Rut_Inf_Def_Recibos`, `Fnc_Inf_Def_Hoja` y el motor `Rut_Inf_Calcular` (`M_130`), y `M_140`. Probada (5.4).
3. ✅ **(yo) Fase 2, JI's:** motor `Rut_Inf_JIs_a_BDatos` (`M_133`) y `M_141`. Probada (5.4).
4. ✅ **(yo) Fase 3, correo:** motor `Rut_Mail_Inf`, `Fnc_Inf_A_Xlsx` y `Rut_Inf_Probar_Xlsx` (`M_135`), `M_142` y
   `Fnc_Ruta_Red` (`Rut_File_Folder_NEXE`). Probada (5.4).
5. ✅ **(yo) Fase 4, Ribbon:** callbacks en `M___RibbonUI`; filas, XML y fila de `Lo_ListObjAPP` aplicados al libro
   el 2026-10-14 a las 11:56 (3.7), con la confirmación del usuario de que Excel y el Custom UI Editor estaban cerrados.
6. ✅ **(usuario)** Módulos importados y compilados (guardado de las 11:51; el libro tiene el código del repo).
   Botones probados en las dos hojas (5.3): «funciona todo».
7. ✅ **(yo) Cierre:** `CLAUDE.md`, volcado de las tablas (`CustomUI/Volcar_Tareas_y_Ribbon.py`), `Validar_Ribbon.py`
   (0 errores), este plan y commit (2026-10-14).

## 5. Pruebas

### 5.1 `Inf_Recibos_TIO` no cambia (doble ejecución por COM, en dos copias del libro)

- **Ábaco:** módulos de hoy contra los nuevos, con los mismos datos. `Herramientas/Comparar_Libros_BD.py --hojas
  Sht__Inf_Recibos_TIO`: 0 diferencias, salvo la hora de `D2`.
- **JI's:** los mismos JI's sembrados en el informe; `M_133` viejo contra el nuevo. BDatos idéntica.
- **Correo:** con las preguntas y el envío simulados. El mismo `.xlsx` y los mismos textos (salvo los nombres de 3.2).

### 5.2 `Inf_Conta_TIO`

- H:AG de `Inf_Conta_TIO` = K:AJ de `Inf_Recibos_TIO`, fila a fila (mismas filas y en el mismo orden).
- En cada fila, y en el bloque EURLE: E = I + M, F = J + N y G = K + O, exactos.
- Totales = los de control de 2.3 (con los mismos datos).
- `B2`, `D2`, cabeceras de JI's y «Año Contable…», leyenda en U:V y cabecera del bloque EURLE en E:O.
- `M_141` desde `Inf_Conta_TIO` deja BDatos igual que `M_133` desde `Inf_Recibos_TIO` con los mismos JI's.
- `M_142` (simulado): adjunto B:W llamado `Inf_Conta_TIO (datos del …).xlsx`, con sus textos, destinatarios y carpeta.

### 5.3 En Excel (usuario)

En cada hoja: el rótulo del grupo («Calc_Inf_Recibos_TIO» / «Calc_Inf_Conta_TIO»), el ábaco y su fecha, la ayuda de
los dos botones, los JI's y el correo con «Prueba en ON».

### 5.4 Resultado de las pruebas por COM (2026-10-14, ✅)

Tres copias del libro guardado a las 09:26, en un Excel aparte (el del usuario, intacto), con los `MsgBox` y el envío
cambiados por respuestas de guion:

- **A**, el código del libro (el de hoy, con la fecha del cálculo apuntando a `App_Last_Calc_Inf_Recibos_TIO`): ábaco,
  `.xlsx`, 4 semillas de JI's en `Inf_Recibos_TIO` y `M_133`.
- **C**, el código nuevo: ábaco de `Inf_Recibos_TIO`, las mismas semillas y `M_133` nuevo.
- **B**, el código nuevo: los dos ábacos, los dos `.xlsx`, las semillas en `Inf_Conta_TIO`, `M_141` y tres envíos
  simulados.

| Comprobación | Resultado |
|---|---|
| `Inf_Recibos_TIO` tras el ábaco, A frente a B (814 celdas: valor, formato, combinadas y altos) | **0 diferencias** |
| `Inf_Recibos_TIO` tras los JI's, A frente a C | **0 diferencias** |
| Adjunto `Inf_Recibos_TIO (datos del 2026-10-02).xlsx`, viejo frente a nuevo (B:Z, 614 celdas) | **0 diferencias** |
| BDatos (160.538 filas × 69 columnas), A frente a C y A frente a B (`M_141` desde `Inf_Conta_TIO`) | **Idéntica.** Solo cambia el texto de `App_Task_Inf`, a propósito: «JI's de Inf_Conta_TIO a BDatos: …» |
| `Inf_Conta_TIO`: B:D y H:AG frente a `Inf_Recibos_TIO`, 15 filas | **0 diferencias** |
| `Inf_Conta_TIO`: E = I + M, F = J + N, G = K + O, en las 15 filas y en las 3 de EURLE | **0 errores** |
| Totales de `Inf_Conta_TIO` | **26.914.076,06 / 15.093.824,60 / 11.820.251,46**, los de control de 2.3 |
| `B2`, cabeceras, celdas combinadas (E31:O31 de EURLE) | «Informe Conta 2026»; las mismas combinadas que antes del cálculo |
| Adjunto `Inf_Conta_TIO (datos del 2026-10-02).xlsx` | Tabla `B4:W20`, la misma comprobación de sumas: 0 errores |
| Rótulo del ábaco en `Inf_Conta_TIO` | La fecha de `App_Last_Calc_Inf_Conta_TIO` |
| `Fnc_Inf_Def_Hoja` | Conta y Recibos, Verdadero; BDatos y `Nothing`, Falso |
| `Fnc_Ruta_Red` | `Red:\Espai\x` y `RED:\Espai\x` → `Y:\Espai\x`; `\Espai\x` y `F:\x`, sin cambios |
| Correo 1: Recibos, Sí y Sí | Asunto y cuerpo de siempre; crea `2026\` y copia el `.xlsx` |
| Correo 2: Conta, Sí y Sí | «PPub - Informe Conta, año contable 2026, cierre 31/12/2026», cuerpo «sin el seguro obligatorio (INSS)», copia `Inf_Conta_TIO (…).xlsx` en `2026\` |
| Correo 3: Conta con `Red:\Espai_ServConta\no_existe_prueba` | «No se encuentra la carpeta Y:\Espai_ServConta\no_existe_prueba»; con «No», no se envía nada |
| Contador de `Rut_Off_Functions` al acabar cada paso | 0 |

Tiempos: ábaco de `Inf_Recibos_TIO` 2,8 s (viejo) y 2,4 s (nuevo); `Inf_Conta_TIO` 2,1 s; JI's, unos 2 s.

## 6. Visto y sin tocar

- En `Inf_Recibos_TIO`, el título de grupo de E:G dice «Importes Recibos: Acad. + Adm.», pero esas columnas llevan
  también el INSS. En `Inf_Conta_TIO`, ese título es exacto.
- **El botón de exportar no anota nunca su fecha.** Su rótulo lee `App_Last_Exp_<hoja>`, pero ninguna rutina lo escribe:
  `Rut_Lo_Export_WorkSheet` no lo hace, y el único que se escribe es `App_Last_Exp_JIsPPub` (`M_90`). En
  `Inf_Recibos_TIO` el rótulo dice «Export WS Exp.», porque su nombre es `App_Last_Exp_Inf_Recibos`, sin `_TIO`. Por eso
  no se crea `App_Last_Exp_Inf_Conta_TIO`.
- `Inf_Recibos_TIO (2)` la ha borrado el usuario (2026-10-14). Siguen `Inf_Recibos_TIO y EP` e `Inf_RecibosJIs`. Se
  corregirá en el `CLAUDE.md` en el cierre.
- `ExportWorkSheet_Group` lleva en `SheetsNames` la hoja `_ConfigAPP`, que no existe (aviso A3 de `Validar_Ribbon.py`,
  de antes).

## 7. Decisiones

1. ✅ **Motor común** (2026-10-14). `M_130`, `M_133` y `M_135` reciben qué informe tratan; `M_140`-`M_142` son finos.
2. ✅ **JI's: las mismas 6 columnas de BDatos (AR:AW)** para los dos informes (tabla de 3.4). Son los JI's de los
   importes Adm. y Acad., que son iguales en los dos. Pasarlos desde cualquiera de los dos escribe en las mismas celdas,
   y el otro informe los recoge al recalcularlo.
3. ✅ **Correo de `Inf_Conta_TIO`:** destinatarios en `App_Mail_Inf_Conta_TIO` y carpeta en `App_RutaRed_Inf_Conta_TIO`
   (los dos, creados por el usuario); la de `Inf_Recibos_TIO` pasa a `App_RutaRed_Inf_Recibos_TIO`. Asunto y cuerpo, los
   propuestos (3.5).
4. ✅ **Título de `B2`:** «Informe Conta 2026».
5. ✅ **Formas:** borradas por el usuario.
6. ✅ **Tabla:** `Lo_Inf_Conta_TIO` (renombrada por el usuario).
7. ✅ **Exportar la hoja, sí; mostrar/ocultar columnas, no** por ahora.
8. ✅ **Ayuda de los botones variable según la hoja:** `getScreentip` (3.6).
9. ✅ **Nombres con `_TIO`** (2026-10-14, renombrados por el usuario): `App_Last_Calc_Inf_Recibos_TIO` y
   `App_Mail_Inf_Recibos_TIO`. El rótulo del ábaco ya no tiene caso especial.
10. ✅ **El prefijo «Red:» de las rutas de `ConfigAPP`** (2026-10-14, a propósito, para distinguir las rutas de la unidad
    de red): «Red» se cambia por `App_User_Unid_Red` (`Red:\Espai…` → `Y:\Espai…`). Lo hace `Fnc_Ruta_Red`, en
    `Rut_File_Folder_NEXE`, para que lo pueda usar cualquier rutina; de momento lo usa la carpeta de red del correo
    (`Fnc_Inf_Carpeta_Red`, `M_135`). Las rutas que empiezan por `\` siguen valiendo como antes. `App_RutaRed_Robot_PPub`
    también lleva «Red:», pero no la lee ningún módulo.
