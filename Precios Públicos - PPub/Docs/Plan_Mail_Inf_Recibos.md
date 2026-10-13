# Plan: enviar por correo el Informe de Recibos (`Inf_Recibos_TIO`) como `.xlsx`, y sus botones al Ribbon

**Proyecto:** `PPub_BDatos_2026.xlsm` · **Fecha:** 2026-10-12 · **Afecta a:** `Rut_Mail_Send` (nuevo),
`M_135_Mail_Inf_Recibos` (nuevo), `Prog__APP_Mail` (hoja nueva, falta exportarla), `M___RibbonUI` (tres callbacks),
`customUI14.xml` (un grupo con tres botones), `Lo_RibbonUI` (cuatro filas), nombres rotos (limpieza), documentación.
Además, lo de la sección 7: `M_130`, `M_132`, `M_133` y `Rut_Lo` (`Rut_Names_Borrar_Rotos`).

**Etiquetas:** ✅ **Hecho** · 🧪 **Aplicado, falta probarlo en Excel** · ⏳ **Pendiente**

## Estado: FASE 1 ✅; FASES 2 Y 3 🧪 PROBADAS POR COM; FASE 4 Y ETIQUETAS «Prueba: NO/SÍ» 🧪 PUESTAS EN EL LIBRO (2026-10-13 noche); pendiente: importar `M___RibbonUI`, `M_130` y `M_135` y probar los botones (fase 5)

- ✅ Configuración de correo (hoja `ConfigMail`) hecha por el usuario y exportada (`Prog__APP_Mail.cls`).
- ✅ Sección 7 arreglada el 2026-10-13 en `M_130`, `M_132` y `M_133`, y nueva `Rut_Names_Borrar_Rotos` en `Rut_Lo`.
  **Importados en el libro**: comprobado con `olevba` en el guardado de las 00:38.
- ✅ Fase 1: `Rut_Mail_Send.bas` importado en el libro y probado con envío real (ver la fase 1 de la sección 4).
- ✅ Los 18 nombres rotos, borrados a mano por el usuario (2.2). Comprobado en el guardado de las 00:33: 0 nombres con
  `#REF!` o vínculo externo, y sin vínculos externos.

Modelos: **Jornadas_y_Congresos** (`M_801_Mail_Send_New`, `M_810_Mail_Valid`, `M_815_Mail_HTML`, `M_825_Sheet_TO_Xlsx` y
`M_715`/`M_720`/`M_725`, que mandan un informe como adjunto) y **Certificados_de_Pago** (`M_800_Mail_Send_New`,
`M_4_Mandar_Correo`, y `M_5`, que copia una hoja a un libro nuevo y la guarda como `.xlsx`).

## 1. Qué se pide

1. Sacar **una parte** de la hoja `Inf_Recibos_TIO` a un `.xlsx` y mandarlo por correo.
2. Pasar las **dos formas con macro** de la hoja («Añadir JI's de BD al Inf.» y «Añadir JI's del Inf. a BD») a
   **botones del Ribbon**. Va en el mismo plan porque toca el mismo `customUI14.xml`: así se edita una sola vez.
3. Decisiones del usuario (2026-10-12):
   - **Parte:** columnas **B:Z**, todas sus filas: títulos, tabla con totales, leyenda de tipos de recibo y bloque EURLE.
     Fuera: los importes académicos por curso (AA:AJ) y las marcas `x` de la fila 1.
   - **Título:** «Informe Recibos 2026», ya puesto en la hoja (`B2:C2`); la copia no lo toca.
   - **Destinatarios:** `Prog__APP.Range("App_Mail_Inf_Recibos")` (direcciones separadas por `;`).
   - **Remitente:** `ingresos@ua.es`.
   - **Cuerpo:** solo texto, **sin** los totales.
   - **Botones:** visibles solo en `Inf_Recibos_TIO`, para **todos** los usuarios de *Recalcular*
     (`Boss,SusanaC,FannyR;RafaG`).
   - **Configuración del correo:** en la hoja nueva `ConfigMail` (CodeName `Prog__APP_Mail`), como JyC y Certificados.

## 2. Cómo está hoy (comprobado en el `.xlsm` guardado a las 23:48, leído como zip)

### 2.1 La hoja `Inf_Recibos_TIO` (CodeName `Sht__Inf_Recibos_TIO`)

| Filas | Columnas (B:Z) | Qué hay |
|---|---|---|
| 1 | B, C, E:R | Marcas `x` que usa el botón *Hide Cols_Row1* (también en AA:AJ). |
| 2 | B:X | **`B2:C2` combinadas: «Informe Recibos 2026»** (desde el 2026-10-13 lo escribe `M_130`); `D2` «Último Cálculo: … / Fecha de cierre contable: …» (lo escribe `M_130`); títulos de grupo **combinados** `E2:G2`, `H2:J2`, `K2:N2`, `O2:R2` y `S2:X2` («UXXI-EC»). |
| 3 | S:X | Cuentas de aplicación (`Cta.Aplic. 7400`…). |
| 4–20 | B:AJ | Tabla **`Lo_Inf_Recibos_TIO`**: cabecera, 15 filas y fila de totales (`SUBTOTAL`). Y:Z son las descripciones del JI y del movimiento contable (anchos 101 y 103). |
| 21–22 | — | Vacías. |
| 23–28 | X:Y | Leyenda de tipos de recibo (`M_130` la copia de `Prog_TipoRec`). |
| 30 | — | Vacía. |
| 31–36 | B:R | **Bloque EURLE**: rango suelto, no tabla. `M_130` filtra `EURLE*`, lo copia aquí y lo borra de la tabla: títulos de grupo combinados (fila 31), cabecera, 3 filas y su «Total» (valores). Solo existe si hay recibos de EURLE. |

Además:

- **2 formas con macro** sobre `S2:X2`: «Añadir JI's de BD al Inf.» (`Rut_Anadir_a_Inf_Recibos_JIs_de_BDatos`, `M_132`)
  y «Añadir JI's del Inf. a BD» (`Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos`, `M_133`). Pasan al Ribbon (3.4).
- **Código de hoja** (`Sht__Inf_Recibos_TIO.cls`): `Worksheet_Activate`/`Deactivate` llaman a `Rut_Context_Buttons_*` y a
  `RefreshRibbon`, que viven en módulos de PPub.
- **Ningún nombre definido** apunta a la hoja, y no hay columnas ocultas.
- La copia visible `Inf_Recibos_TIO (3)` ya no está. Sigue la antigua `Inf_Recibos_TIO (2)` (muy oculta, de febrero,
  CodeName `Sht__Inf_Recibos_TIO1`); al envío no le afecta.

### 2.2 Configuración del correo

**Hoja nueva `ConfigMail`** (CodeName **`Prog__APP_Mail`**, como en JyC y Certificados):

| Fila | Nombre | Valor |
|---|---|---|
| 2 | `APP_MailCta` | `ingresos@gcloud.ua.es` |
| 3 | `APP_MailClau` | 16 caracteres (contraseña de aplicación) ✅ |
| 4 | `APP_MailFrom` | `ingresos@ua.es` |
| 5 | `APP_Servicio` | `Servei de Comptabilitat / Servicio de Contabilidad` |
| 6 | `APP_Unidad` | `Unitat d'Ingressos / Unidad de Ingresos` |
| 7 | `APP_Web_es` | `https://sc.ua.es/es/tesoreria/unidad-ingresos/` |
| 8 | `APP_Web_va` | `https://sc.ua.es/va/tesoreria/unitat-ingressos/` |
| 9 | `APP_MailFirm` | Firma HTML de JyC, con sus **5 huecos**: `ext. ____`, `Servei _____`, `Unitat _____`, `https://sc.ua.es/es/_____` y `https://sc.ua.es/va/_____` |

**En `ConfigAPP`** (`Prog__APP`): `APP_User_Mail` (C4, `dugarry@ua.es`), `APP_User_Ext` (C5), `APP_User_Name` (C3),
`APP_Last_Calc_Inf_Recibos` (B42) y **`App_Mail_Inf_Recibos` (B46, `felipe.martinez@gcloud.ua.es`)**. No hay
`APP_User_Rubrica` (JyC la usa como rúbrica; aquí se usará `APP_User_Name`).

⚠️ **Nombres rotos.** Al copiar `ConfigAPP` para crear `ConfigMail` y borrar después las filas de correo de `ConfigAPP`,
quedaron 3 `APP_Mail*` de ámbito libro con `ConfigAPP!#REF!`; el usuario los borró el 2026-10-12, junto con
`APP_AñoCont` (de libro, con la ñ de antes de la normalización, apuntando a otro libro). **Quedan 18** (guardado del
2026-10-13 00:18): `APP_C_Acad_Ant`, `APP_C_Acad_Pos` y `APP_FechCierreCont` **con ámbito de hoja** en `ConfigMail`,
`DefCol_LoDefCols`, `DefCol_MenuAux`, `ListObjAPP`, `RibbonUI` y `SwitchsAPP`, todos apuntando a
`[EP_202x-2x_BaseDatos-LIQ _V3.xlsm]_ConfigAPP` (el libro de Enseñanzas Propias, vínculo externo `externalLink1`).

- El Administrador de nombres no los enseñaba porque **las hojas estaban ocultas**: los nombres de ámbito hoja de una
  hoja oculta no salen ahí (lo comprobó el usuario; anotado en el CLAUDE.md global). Al mostrar las hojas aparecieron y
  los borró a mano (2026-10-13). Para otra vez queda **`Rut_Names_Borrar_Rotos`** (`Rut_Lo`), que los encuentra estén
  o no visibles: sin argumento solo los lista en Inmediato; con `False` los borra. Respeta los internos de Excel
  (`_FilterDatabase`…, 4 con `#REF!`) y `T_Organicas` (`=Lo_Organicas[]`).
- Borrarlos es seguro: las fórmulas que usan esos nombres (`Filtros_Tipo_Rec`, `Tb_TiposRec`, `BD_ImpAdm_2025-26`)
  están en hojas **sin** copia rota, así que ya leen los de ámbito libro (`ConfigAPP!B17`, `B18`, `B16`).
- ✅ Al borrarlos desapareció también el vínculo externo.

El código nuevo **siempre** cualifica con la hoja (`Prog__APP_Mail.Range("APP_MailCta")`).

### 2.3 Lo que hacen JyC y Certificados

| | JyC | Certificados |
|---|---|---|
| Motor | `M_801` `Rut_Email_Send` (Sub; `Enviado` ByRef) | `M_800` `Fnc_Email_Send` (Function Boolean + `ErrMsg` ByRef) |
| Envío | CDO con enlace tardío (`CreateObject("CDO.Message")`), `smtp.gmail.com:465` SSL, contraseña de aplicación de Google | Igual |
| Configuración | `Prog__APP_Mail` + `APP_User_Ext`/`APP_User_Rubrica` de `Prog__APP`; firma con 5 huecos | Igual, firma con el hueco `ext. ____` |
| Modo prueba | `Sw_Test` → constante `DIR_PRUEBA` y asunto «[PRUEBA] …» | `Sw_Test` → `APP_User_Mail_Cta` |
| Errores | `MsgBox` dentro del motor | `MsgBox` dentro del motor, y además `ErrMsg`; trata los errores 52/53/76 de ruta del adjunto |
| Validación | `M_810` `Fnc_Valid_Email_Multi`: pregunta con `MsgBox` si quitar cada dirección mala | Igual, con `ErrMsg` |
| Hoja → `.xlsx` | `M_825` `Fnc_SheetToXlsx`: copia el `UsedRange` **visible** a un libro nuevo con `PasteSpecial` (anchos, valores, formatos), recrea las tablas a mano (`Rut_Xlsx_Recrear_Tablas`), guarda en `%TEMP%` | `M_5`: `ActiveSheet.Copy` a un libro nuevo y `SaveAs` |
| Flujo | El botón genera el informe y pregunta «¿Enviamos…?» → `Rut_Mail_Inf_*`: adjunto, asunto con la fecha de corte, cuerpo HTML, envío, `Kill` del temporal, confirmación | — |

PPub ya tiene su propia exportación de hoja: `Rut_Lo_Export_WS_Xlsx` (`Rut_Ws_Export.bas`, botón *Export WS*), que hace
`WrkSht.Copy` + `SaveAs` con diálogo, como Certificados.

## 3. Diseño

### 3.1 El adjunto: copiar la hoja entera y recortarla (no el `PasteSpecial` de JyC)

| | Copiar y recortar (**elegido**) | `PasteSpecial` de un rango (JyC `M_825`) |
|---|---|---|
| Tabla | Viaja como tabla, con estilo, autofiltro y totales vivos | Hay que recrearla a mano (y aquí sería parte de la tabla, B:Z de B:AJ) |
| Altos de fila | Se conservan (cabecera de 3 líneas, leyenda de 80) | **No se copian** |
| Combinadas, formatos de número, colores | Se conservan | Se conservan |
| Columnas ocultas | Se muestran en la copia | **No viajan** (`SpecialCells(xlCellTypeVisible)`). La fila 1 marca con `x` justo B, C y E:R: con *Hide Cols_Row1* puesto, el adjunto saldría sin los importes |
| Precedente | `Rut_Lo_Export_WS_Xlsx` (PPub) y `M_5` (Certificados) | JyC |

Con B:Z **no se borra ninguna fila**: todo lo que hay en esas columnas va al adjunto tal cual (título, tabla, leyenda y
bloque EURLE).

Pasos de `Fnc_Inf_Recibos_A_Xlsx` (en `M_135`), **sin tocar nunca la hoja origen**:

1. **Comprobar** en la hoja origen que la tabla tiene filas y que la fila de totales está visible (si no,
   `Rut_Lo_Totales_Mostrar`). Si algo no cuadra, no se genera nada y se dice por qué.
2. **Apagar eventos** (`Rut_Off_Functions` y `Sw_WB_Deactivate = False`, como `M_120`/`M_311`). Es obligatorio:
   la copia lleva el código de la hoja, y su `Worksheet_Activate` llamaría a rutinas de PPub que en el libro nuevo no
   existen (error de compilación); y el `Workbook_Deactivate` de PPub restablecería Excel a mitad.
3. `Sht__Inf_Recibos_TIO.Copy` y en seguida `Set Wb_New = ActiveWorkbook`.
4. En la copia:
   - mostrar todas las filas y columnas y quitar filtros;
   - borrar las columnas a la derecha de Z (de `Lo.Range.Column + InfRec_Descrip_Contab` en adelante): la tabla queda en
     B:Z y sus `SUBTOTAL` siguen valiendo; con ellas se van el título `AA2:AJ2` y los importes académicos;
   - **vaciar** las celdas de la fila 1 que contengan `x`, sin borrar la fila (no se toca `B2:C2`);
   - borrar las formas **con macro** (`OnAction` no vacío). Cuando los botones de JIs estén en el Ribbon y se quiten las
     formas, este paso ya no hará nada, pero se deja como red de seguridad;
   - red de seguridad: romper vínculos (`LinkSources` / `BreakLink`) y borrar nombres con `#REF!` o `[`. Hoy no hay,
     pero si alguna fórmula de la hoja usara un nombre, la copia lo arrastraría.
5. Guardar en `%TEMP%\Inf_Recibos_<AnoCont>_<cierre aaaa-mm-dd>.xlsx` (`FileFormat:=51`, con `DisplayAlerts = False`,
   que descarta en silencio el código de hoja), cerrar y devolver la ruta.
6. En **todas** las salidas: cerrar `Wb_New` si quedó abierto, `Sw_WB_Deactivate = True`, `Rut_On_Functions`, y volver a
   `ThisWorkbook` y a la hoja.

Si más adelante otro informe necesita lo mismo, la parte genérica (pasos 2, 3, 4 menos el recorte, y 6) se sube a
`Rut_Ws_Export` como función propia. `Rut_Lo_Export_WS_Xlsx` no se toca.

### 3.2 Motor de correo: `Rut_Mail_Send.bas` (nuevo)

Prefijo `Rut_` porque es librería transversal (convención de PPub) y servirá para otros envíos. Con la hoja `ConfigMail`,
la configuración y la firma son las de JyC, así que se porta **el motor de JyC** (`M_801`: lectura de
`Prog__APP_Mail` y sustitución de los 5 huecos de la firma) **con la forma de función del de Certificados**
(`Fnc_Email_Send`, que devuelve `Boolean` y `ErrMsg`). Cambios:

- **Nombres:** todos los de correo con `Prog__APP_Mail.Range(...)`; `APP_User_Ext`, `APP_User_Name` y `APP_User_Mail` con
  `Prog__APP.Range(...)`; `Sw_Test` con `Prog__APP_Switch.Range(...)`.
- **Rúbrica:** `APP_User_Name` (PPub no tiene `APP_User_Rubrica`).
- **Modo prueba:** a `APP_User_Mail` (el usuario activo), no a una constante.
- **Sin `MsgBox` dentro.** Devuelve `False` y la causa en `ErrMsg` (credenciales, sin conexión, adjunto no encontrado…),
  con los mismos textos de ayuda que JyC/Certificados. Motivo: se ejecuta con `Form_Running_Rut` abierto a pantalla
  completa, y un `MsgBox` normal se abre **detrás** y Excel parece colgado (pasó el 2026-10-11). El llamador escribe
  `ErrMsg` en el informe y avisa con `vbMsgBoxSetForeground + vbSystemModal`.
- **Validación sin preguntas:** `Fnc_Valid_Email` (la expresión regular de `M_810`) y `Fnc_Valid_Email_Lista`, que
  limpia la lista (espacios, `;` sobrantes) y devuelve en `ErrMsg` las direcciones mal escritas. No se pregunta si
  quitarlas, como en `M_810`: la lista es de configuración y se corrige en `ConfigAPP`.
- `Fnc_HTML_Texto` (de `M_815`: el euro como `&euro;`, saltos de línea como `<br>`).
- `Rut_Cambiar_Contrasena_Email` (escribe la contraseña nueva en `APP_MailClau` sin los espacios que añade Google), como
  tarea **sin botón** en `Tb_Tareas` para Boss.
- CDO con enlace tardío: no hace falta marcar ninguna referencia en el VBE.

### 3.3 Rutina de negocio: `M_135_Mail_Inf_Recibos.bas` (nuevo)

`Rut_Mail_Inf_Recibos` (va en el rango 130, junto a `M_130`, `M_132` y `M_133`):

1. `Set ActivForm = Form_Running_Rut`; cada paso al informe con `Rut_TimeLap_Inf`.
2. Lee `Prog__APP.Range("App_Mail_Inf_Recibos")`; vacía o con direcciones malas → no envía y lo dice.
3. Pregunta con `Func_MsgBox_vbYesNo` (`M_110`, que ya sale por delante del formulario): «¿Enviar el Informe de
   Recibos (último cálculo: …, cierre contable: …) a …?». Con `Sw_Test`, dice que irá a `APP_User_Mail`. **No regenera
   el informe**, como JyC: se manda lo que hay en la hoja, y la pregunta enseña la fecha del último cálculo
   (`APP_Last_Calc_Inf_Recibos`).
4. `Fnc_Inf_Recibos_A_Xlsx` → ruta del adjunto.
5. **Asunto:** `PPub - Informe de Recibos, año contable 2026, cierre dd/mm/aaaa`.
   **Cuerpo** (HTML, estilo JyC, sin totales): saludo; «Adjunto el Informe de Recibos de Precios Públicos del año
   contable 2026 (cursos 2025-26 y 2026-27): importes emitidos, cobrados y pendientes por tipo de recibo y concepto,
   desglosados en seguro obligatorio (INSS), administrativos y académicos, con los justificantes de ingreso (JI) y la
   descripción de cada movimiento contable. Las filas de EURLE van aparte, debajo.»; fecha de cierre y del último
   cálculo.
6. `Fnc_Email_Send(...)`, y después `Kill` del temporal **siempre** (con `On Error Resume Next`). Con `Sw_Test`, el
   adjunto se conserva en `%TEMP%` y su ruta va al informe, para poder revisarlo.
7. Informe final: «Enviado a …» o la causa del fallo. Lo guarda `Rut_Progreso_Cerrar` en `Informe_Rut`, y el supertip de
   Boss enseña así el último envío sin ningún nombre nuevo.

### 3.4 Ribbon: grupo «Informe de Recibos» con tres botones

> **Actualizado el 2026-10-13:** al final son **dos** botones (correo y «JI's del Inf. a BD»), porque los JI's de BDatos
> al informe los pone `M_130` y `M_132` se retiró (`Docs/Plan_Botones_JIs_Ribbon.md`). Lo puesto en el libro está en la
> fase 4 de la sección 4. Lo de abajo es el diseño original.

⚠️ **`customUI14.xml`: se avisa antes de tocarlo y después.** El Custom UI Editor no deja fichero de bloqueo; si lo
tienes abierto, al guardar machacaría el cambio (o al revés).

- **XML**: un grupo nuevo junto a `ReCalculate_ListObj_Sht_Group`, visible solo en `Inf_Recibos_TIO`. Las imágenes de
  los botones de JIs se eligen entre las que ya tiene el libro (`Table_Add`, `Import_Blue`, `Export_JIs`…):

  ```xml
  <group id="Inf_Recibos_Group"    tag="Inf_Recibos_Group"    label="Informe de Recibos"
         getVisible="GetVsbl_CtrlTab">
      <button id="BtnJIs_BD_a_Inf"         tag="JIs_BD_a_Inf"         image="Table_Add"    size="large"
              label="JI's de BD al Inf."    getVisible="GetVsbl_CtrlTab"    getSupertip="getStip_CtrlTab"
              onAction="OnAct_JIs_BD_a_Inf"/>
      <button id="BtnJIs_Inf_a_BD"         tag="JIs_Inf_a_BD"         image="Import_Blue"  size="large"
              label="JI's del Inf. a BD"    getVisible="GetVsbl_CtrlTab"    getSupertip="getStip_CtrlTab"
              onAction="OnAct_JIs_Inf_a_BD"/>
      <button id="BtnMail_Inf_Recibos"     tag="Mail_Inf_Recibos"     image="Mail"         size="large"
              label="Enviar por correo"     getVisible="GetVsbl_CtrlTab"    getSupertip="getStip_CtrlTab"
              onAction="OnAct_Mail_Inf_Recibos"/>
  </group>
  ```

- **Callbacks** en `M___RibbonUI`:

  ```vba
  Sub OnAct_Mail_Inf_Recibos(control As IRibbonControl)       '- tarea larga: con el formulario de progreso
      Call Rut_Progreso_Abrir("Enviar el Informe de Recibos por correo")
      Call Rut_Mail_Inf_Recibos
      Call Rut_Progreso_Cerrar(control.Tag)
  End Sub

  Sub OnAct_JIs_BD_a_Inf(control As IRibbonControl)           '- como la forma de hoy, sin formulario
      Call Rut_Anadir_a_Inf_Recibos_JIs_de_BDatos
      Call RefreshRibbon
  End Sub

  Sub OnAct_JIs_Inf_a_BD(control As IRibbonControl)
      Call Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos
      Call RefreshRibbon
  End Sub
  ```

  Los de JIs llaman a la rutina **igual que la forma**, sin `Rut_Progreso_Abrir`: `M_132` pregunta a mitad con un
  `MsgBox` («¿Quieres los JI's de las Enseñanzas Propias…?»), que con el formulario abierto saldría detrás y Excel
  parecería colgado; y ninguna de las dos escribe en el informe del formulario. Si más adelante se quieren con
  formulario, primero hay que cambiar ese `MsgBox` por `Func_MsgBox_vbYesNo`.
- **Filas en `Lo_RibbonUI`** (sin fila, el control queda oculto), todas con `Usuario` = `Boss,SusanaC,FannyR;RafaG`,
  `SheetsNames` = `Inf_Recibos_TIO` y `Group-Tag` = `Inf_Recibos_Group`:

  | `Uribbon-Tags` | `Nom_Rut` |
  |---|---|
  | `Inf_Recibos_Group` | — |
  | `JIs_BD_a_Inf` | `Rut_Anadir_a_Inf_Recibos_JIs_de_BDatos` |
  | `JIs_Inf_a_BD` | `Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos` |
  | `Mail_Inf_Recibos` | `Rut_Mail_Inf_Recibos` |

- Cuando los botones funcionen, **borrar las dos formas** de la hoja.
- Después: `CustomUI/Volcar_Tareas_y_Ribbon.py` y `CustomUI/Validar_Ribbon.py` (con el libro guardado). Ninguna de las
  tres rutinas está en `Tb_Tareas`, así que la regla E5 (rutina con botón y con fila en `Tb_Tareas`) no salta.

## 4. Orden de trabajo

0. ⏳ **Preparar.**
   - ✅ Hoja `ConfigMail` con la contraseña de aplicación, el remitente y la firma; `App_Mail_Inf_Recibos` en `ConfigAPP`.
   - ✅ Hoja nueva exportada a `VBA_Moduls/Prog__APP_Mail.cls` (falta commitearla).
   - ✅ Los cambios de la otra sesión (`M_130`, separadores, `M___RibbonUI`) ya están commiteados (`cd97ee759`,
     `b3c2fb75c`).
   - ✅ Los 18 nombres rotos, borrados.
   - ✅ Importados `M_130_Gen_Inf_Recibos`, `M_132_Anadir_Num_JIs_al_Inf`, `M_133_Anadir_Num_JIs_a_BDatos` y `Rut_Lo`
     (sección 7; `olevba`, guardado de las 00:38).
   - ⏳ `Sw_Test` a VERDADERO para las pruebas (`Workbook_Open` lo deja en FALSO al abrir el libro).
1. 🧪 **`Rut_Mail_Send.bas`** (2026-10-13). Lo escrito, respecto a 3.2:
   - `Fnc_Email_Send(Dest, Asunto, [Adjunto], [CuerpoHTML], [ErrMsg], [DestReal])`: `DestReal` devuelve a quién va el
     correo de verdad (con `Sw_Test`, `APP_User_Mail`), para el informe. La lista de destinatarios se valida también
     en modo prueba. Si el adjunto no existe, **no envía** (Certificados avisaba y lo mandaba sin adjunto).
     `CuerpoHTML` es un fragmento: el motor pone `<html><body>`, la rúbrica (`APP_User_Name`) y la firma.
   - `Fnc_Valid_Email_Lista` acepta `;` y `,`, quita espacios (también el de no separación), separadores sobrantes y
     direcciones repetidas.
   - `Fnc_HTML_Texto` escapa además `&`, `<` y `>`.
   - Errores: los de JyC/Certificados, y además `0x80040211` con `0x80040217` en la descripción (no inicia sesión:
     credenciales). Se sabrá en la prueba 5 qué código da Gmail con una contraseña mala.
   - **Prueba:** la línea de Inmediato que proponía este plan (`? Fnc_Email_Send(..., Msg), Msg`) no compila
     ("El tipo de argumento ByRef no coincide": en Inmediato `Msg` es `Variant`). En su lugar, `Call Rut_Mail_Probar`
     (con adjunto, `Call Rut_Mail_Probar(True)`): primero `Fnc_Mail_Probar_Sin_Enviar` (8 listas de direcciones,
     `Fnc_HTML_Texto` y 2 envíos que no deben salir: sin destinatarios y con un adjunto que no existe) y después un correo al usuario activo con tildes,
     eñe y euro en el asunto y en el cuerpo.
   - **Probado por COM** (copia del guardado de las 00:38, instancia de Excel aparte): compila;
     `Fnc_Mail_Probar_Sin_Enviar` sin fallos; con `SMTP_SERVIDOR` apuntando a `127.0.0.1` en la copia, el envío
     recorre CDO entero (mensaje, configuración, adjunto) y falla al conectar con el texto de «Sin conexión», con
     `DestReal` = `a@ua.es;b@ua.es` (lista limpia) y, con `Sw_Test`, `dugarry@ua.es`.
   - **Primer envío real** (2026-10-13, `Rut_Mail_Probar(True)` en el libro): llegó a `dugarry@ua.es` con el cuerpo,
     la rúbrica y la firma bien, y la eñe y las tildes del asunto bien, pero **el € del asunto llegó como `?`**.
     Causa, vista en el MIME que genera CDO (`GetStream`, sin enviar): CDO usa por defecto `iso-8859-1`, que no tiene
     el euro, y el HTML iba sin juego de caracteres declarado (Gmail lo adivinó). Con `.BodyPart.Charset = "utf-8"`
     antes del asunto y `.HTMLBodyPart.Charset = "utf-8"` tras el cuerpo, asunto, texto y HTML salen en `utf-8` y se
     decodifican bien. Aplicado (11:26) y probado por COM. JyC y Certificados tienen lo mismo (sin tocar).
   - ✅ **Segundo envío real** (reimportado, `Rut_Mail_Probar(True)`): asunto con su €, cuerpo, rúbrica, firma y el
     adjunto `PPub_Prueba_Correo.txt`, todo bien. **Fase 1 cerrada.**
2. 🧪 **`Fnc_Inf_Recibos_A_Xlsx([ErrMsg])`** (en `M_135_Mail_Inf_Recibos`, nuevo, 2026-10-13), y la prueba
   `Rut_Inf_Recibos_Probar_Xlsx` (genera el `.xlsx` sin enviarlo y lo señala en el Explorador). Respecto a 3.1:
   - El filtro de la tabla se quita desde la tabla (`LoN.AutoFilter.ShowAllData`, como `Rut_Lo_Filtros_Quitar`): con
     `WsN.ShowAllData` daba **error 1004** en la copia de una hoja filtrada.
   - **Altos de fila:** casi todas las filas tienen alto automático (`M_130` hace `AutoFit`) y en el libro nuevo Excel
     los recalcula a su manera (filas de 15,5 a 30). Se copian los de la hoja (los que se ven en PPub); las filas
     ocultas por un filtro dan alto 0 y se quedan con el suyo.
   - **Nombres:** Excel copiaba con la hoja `RibbonPtr` (constante de PPub). Se borra todo nombre que no apunte a un
     rango de la propia copia (también los `#REF!` y los de otros libros).
   - Si la hoja está oculta, se muestra un momento para copiarla y se vuelve a ocultar. Se vuelve a la hoja activa.
     `Sw_WB_Deactivate` recupera su valor anterior (no se fuerza a VERDADERO).
   - **Probado por COM** (copia del guardado de las 00:38): (A) la hoja tal cual y (B) muy oculta, con E:R ocultas y
     la tabla filtrada. Los dos `.xlsx`: una hoja, tabla `B4:Z20` con 25 columnas, totales y botones de filtro sin
     filtro, **0 celdas distintas** de la hoja en B:Z (valor y fórmula), combinadas iguales (sin `AA2:AJ2`), anchos
     iguales, fila 1 sin `x`, sin formas, sin vínculos, sin nombres, sin macros, `calcPr` **sin** `calcMode` (=
     Automático). La hoja origen, igual que antes (oculta, columnas ocultas, filtro, hoja activa), contador de
     `Rut_Off_Functions` a 0. Riesgo 3: pasar PPub a Automático un momento cuesta 0,01 s (un `CalculateFull`, 0,04 s).
   - **Decisiones del usuario (2026-10-13):** la fila 1, que queda vacía, con **alto normal** (`StandardHeight`; en la
     hoja mide 59 puntos), y el adjunto al **100 % de zoom** (la hoja está al 55 %). Aplicado y probado por COM.
   - `D2` dirá «cierre» cuando se vuelva a calcular el informe con el `M_130` nuevo (en el guardado de las 00:38 aún
     decía «cierro»).
   - **Nombre del adjunto (decisión del usuario, 2026-10-13):** `Inf_Recibos_TIO (datos del AAAA-MM-DD).xlsx`, con la
     **fecha de la carpeta de descarga del Robot**, que se guarda en `App_Last_Import_Robot_PPub` (`ConfigAPP!B23`,
     creado por el usuario). La anota `M_110` al importar el LSGES04 del año contable: `Fnc_Fecha_Carpeta_Robot`
     (`Rut_Lo_Import`, junto a `Fnc_Ruta_Inicial_Robot` y con su misma regla) saca la fecha de la carpeta que contiene
     el fichero elegido (`…\2026\2026-10-02\LsGes04_AñoCont_2026_(2026-10-02).xlsx`); si no es una carpeta de fecha,
     deja la celda vacía y el adjunto se llama `Inf_Recibos_TIO.xlsx`. `M_210`, `M_310` y `M_120` no la tocan.
     Probado por COM: el proyecto compila; rutas del Robot, de la Red Nexe (`/`), «2026-10-02 bis», sin carpeta de
     fecha y «2026-02-30»; y el nombre con la celda con fecha, número de serie, texto, vacía y sin el nombre.
   - **Falta:** importar `M_135_Mail_Inf_Recibos`, `M_110_Update_LSGES04_ACont` y `Rut_Lo_Import`, y que el usuario
     genere y revise el `.xlsx` en su Excel.
3. 🧪 **`Rut_Mail_Inf_Recibos`** (en `M_135`) y el callback **`OnAct_Mail_Inf_Recibos`** (`M___RibbonUI`, tras el del
   ábaco) (2026-10-13):
   - El callback: `Rut_Progreso_Abrir` → `Rut_Mail_Inf_Recibos` → `Rut_Progreso_Cerrar(control.Tag, True)`.
   - La rutina: valida `App_Mail_Inf_Recibos` (`Fnc_Valid_Email_Lista`); pregunta con `Func_MsgBox_vbYesNo`
     (último cálculo, cierre, descarga de los recibos, destinatarios y, con `Sw_Test`, «MODO PRUEBA: irá solo a
     `App_User_Mail`»); genera el adjunto; asunto `PPub - Informe de Recibos, año contable 2026, cierre dd/mm/aaaa`;
     cuerpo HTML al estilo de JyC `M_720` («Hola», qué lleva el informe, EURLE aparte, fechas de descarga, cierre y
     último cálculo; sin totales); envía; borra el temporal salvo con `Sw_Test`; todo al informe del formulario. Los
     fallos, al informe y en un `MsgBox` por delante del formulario (`Rut_Mail_Inf_Recibos_Aviso`).
   - **Los callbacks de JI's ya no son de este plan:** según `Docs/Plan_Botones_JIs_Ribbon.md` (otra sesión, mismo
     día), «Añadir JI's de BD al Inf.» lo hace `M_130` al calcular y queda un solo botón, «JI's del Inf. a BD», con su
     callback en ese plan. El grupo «Informe de Recibos» (fase 4) llevará el botón de correo y ese.
   - **`D2` del informe (petición del usuario):** `M_130` añade una tercera línea, `"Descarga de los Recibos del " &
     Prog__APP.Range("App_Last_Import_Robot_PPub")`. La fila 2 ya mide 63,75 (la amplió el usuario): caben las tres.
   - **Probado por COM** (copia del guardado de las 12:44, con los 6 módulos del repo: `Rut_Lo_Import`, `M_110`,
     `M_130`, `M___RibbonUI`, `Rut_Mail_Send` y `M_135`): el proyecto compila; `D2` sale con sus tres líneas; y
     `Rut_Mail_Inf_Recibos` de principio a fin con y sin `Sw_Test` (en la copia: pregunta respondida «Sí», sin
     formulario, `MsgBox` a Inmediato y servidor `127.0.0.1`): pregunta, informe, adjunto `Inf_Recibos_TIO (datos del
     2026-10-02).xlsx` (con `D2` de tres líneas), fallo «Sin conexión» y, con `Sw_Test`, el adjunto conservado.
     **Falta:** el envío real desde el botón (fase 4) o desde Inmediato.
   - ⚠️ **Envío real por error (2026-10-13, en la primera prueba del usuario desde Inmediato):** el informe llegó a
     `felipe.martinez@…` (asunto sin «[PRUEBA]») porque `Sw_Test` estaba en FALSO. El usuario lo creía activo: en el
     Ribbon se veía el botón **«Prueba ON»**, que es el que se ve con la prueba **apagada** y significa «pulsa para
     activarla» (`BtnSW_PruebaON`, `getVisible` = `Not Sw_Test`; con la prueba encendida se ve «Prueba OFF»). El motor
     hizo lo correcto; la pregunta no avisaba de que el envío era real. Arreglos (decisión del usuario):
     - ✅ **La pregunta** (en el repo, probada por COM, falta importarla): con un envío real empieza por
       «¡ ENVÍO REAL !   Va a:  …», con icono de aviso y **«No» por defecto** (`vbDefaultButton2`: un Intro sin mirar no
       envía nada); con `Sw_Test`, «MODO PRUEBA (Sw_Test): irá solo a dugarry@ua.es (no a …)». Ya no usa
       `Func_MsgBox_vbYesNo` (que no tiene botón por defecto). El informe del formulario anota «Envío real a …» o «Modo
       prueba…». (El ⚠ no cabe: el `MsgBox` de VBA solo muestra caracteres de Windows-1252.)
     - 🧪 **Las etiquetas del Ribbon** (puestas en el libro el 2026-10-13 noche, junto con la fase 4; falta verlas en
       Excel), que digan el estado: `BtnSW_PruebaON` → **«Prueba: NO»** y `BtnSW_PruebaOFF` →
       **«Prueba: SÍ»** (solo el `label`; las ayudas «Activar/Desactivar SW_Prueba» y las del menú contextual ya
       describen bien estado y acción). Pendiente: el usuario guarda y cierra Excel y confirma que no tiene el
       `customUI14.xml` abierto en el Custom UI Editor; entonces copia del libro en `CopiaSeguridad\`, cambio de las dos
       etiquetas en el zip, verificación de que el resto queda idéntico y aviso al usuario.
   - 🧪 **Copia en la carpeta de red** (petición del usuario, 2026-10-13, escrita desde la sesión de
     `Docs/Plan_Botones_JIs_Ribbon.md`):
     - **Dónde:** `App_Ruta_Inf_Recibos`, en `ConfigAPP!C10`; `App_RutaRed_Inf_Recibos` apunta a la misma celda. La
       celda guarda la ruta **sin unidad** (`\Espai_ServConta\_Unidad_Tesorería\Tesoreria_Felipe\Aa_TESORERIA_FELIPE\DOCUMENTOS DUGARRY`),
       como `App_RutaRed_Robot_PPub`.
     - **Ruta completa:** `Fnc_Inf_Recibos_Carpeta_Red` le pone delante la unidad de red del usuario
       (`App_User_Unid_Red`, que `Form_Usuario` cambia con el teletrabajo; en el trabajo, `Y:`, que es Nexe por
       WebDAV). Una ruta completa (`Y:\…`, `\\servidor\…`) se usa tal cual, y una de la Red Nexe (`https://…`) pasa por
       `Fnc_Format_Ruta` y `Fnc_URL_Decodificar`.
     - **Pregunta:** una propia, después de la del envío: «¿ Copiar también el informe en la carpeta de red ?», con la
       carpeta y el nombre del fichero.
       - Si ya hay uno con ese nombre, la pregunta lo dice («se sustituirá»).
       - Con `Sw_Test`, avisa de que la carpeta es la de verdad, y «No» es el botón por defecto.
       - Si la carpeta no se encuentra o falta la ruta o la unidad, pregunta si se envía el correo de todos modos.
     - **La copia:** `FileCopy` del `.xlsx` de `%TEMP%`, antes del envío. Si falla, avisa por delante del formulario
       (con el error 70, «¿ lo tiene alguien abierto ?») y el correo se envía igual.
     - **Nombre:** el mismo que el adjunto, `Inf_Recibos_TIO (datos del AAAA-MM-DD).xlsx`, ahora en
       `Fnc_Inf_Recibos_Nombre`, que usa también `Fnc_Inf_Recibos_A_Xlsx`. Nunca lleva la hora detrás, aunque la lleve
       el temporal.
     - **Probado por COM** (copia del guardado de las 13:23, preguntas contestadas por código, envío simulado: no
       salió ningún correo). Salieron como se esperaba los 9 casos:
       - copia en una carpeta de prueba;
       - repetida (se sustituye);
       - destino de solo lectura (error 75, aviso, el correo sale igual);
       - «No» a copiar;
       - «No» al envío;
       - carpeta inexistente, con «No» y con «Sí» a enviar de todos modos;
       - `C10` vacía;
       - la ruta real de `C10`: `Y:\Espai_ServConta\…\DOCUMENTOS DUGARRY` existe y se pregunta bien. Se contestó «No»:
         no se escribió nada en Nexe.
4. 🧪 **Ribbon** (2026-10-13 noche, con Excel y el Custom UI Editor cerrados; copia previa en
   `CopiaSeguridad/PPub_BDatos_2026_antes_Ribbon_Inf_Recibos_2026-10-13_2028.xlsm`):
   - **XML:** grupo `Inf_Recibos_Group` («Informe de Recibos») detrás del ábaco, con dos botones: «JI's del Inf. a BD»
     (`BtnJIs_Inf_a_BD`, icono `Copy_JIs`, `OnAct_JIs_Inf_a_BD`, de `Docs/Plan_Botones_JIs_Ribbon.md`) y «Enviar por
     correo» (`BtnMail_Inf_Recibos`, icono `Mail`, `OnAct_Mail_Inf_Recibos`). Ya no hay botón «JI's de BD al Inf.»: lo
     hace `M_130` al calcular, y `M_132` se retiró. Zip reconstruido y verificado (solo cambia `customUI14.xml`).
   - **Filas en `Lo_RibbonUI`:** `Inf_Recibos_Group`, `JIs_Inf_a_BD` y `Mail_Inf_Recibos`, con `Boss,SusanaC,FannyR;RafaG`
     e `Inf_Recibos_TIO`. Volcado (`Lo_RibbonUI_snapshot.csv`) y `Validar_Ribbon.py`: 0 errores.
   - **Formas:** borrada la de `M_132`; la de `M_133` se borra cuando se pruebe el botón.
5. ⏳ **Pruebas** (sección 5), en los dos ordenadores.
6. ⏳ **Documentación:** `CLAUDE.md` del proyecto (motor de correo, `ConfigMail`, `M_135`, botones de JIs) y una línea
   en el separador del rango 13 cuando los separadores renombrados estén commiteados. Fila de
   `Rut_Cambiar_Contrasena_Email` en `Tb_Tareas`.

Componentes a importar en el libro: `Rut_Mail_Send` y `M_135_Mail_Inf_Recibos` (nuevos) y `M___RibbonUI`.

## 5. Cómo probarlo

Con `Sw_Test` = VERDADERO (el correo va a `APP_User_Mail` con «[PRUEBA]»):

| # | Caso | Debe pasar |
|---|---|---|
| 1 | Normal, con EURLE | Adjunto con B:Z: título, tabla como tabla con totales, leyenda y bloque EURLE; sin formas ni `x` |
| 2 | Con *Hide Cols_Row1* puesto (B, C y E:R ocultas en la hoja) | El adjunto sale completo y la hoja sigue con sus columnas ocultas |
| 3 | Con un filtro en la tabla | El adjunto sale completo y la hoja sigue filtrada |
| 4 | `App_Mail_Inf_Recibos` vacía, o con una dirección mal escrita | No envía y dice cuál |
| 5 | Contraseña mala | Mensaje de credenciales **por delante** del formulario; Excel no se queda «colgado» |
| 6 | Sin conexión | Mensaje de conexión, igual |
| 7 | Responder «No» a la pregunta | No se genera nada |
| 8 | Después de cada caso | La hoja origen igual que antes, el libro sin vínculos nuevos, contador de `Rut_Off_Functions` a 0 (`? Fnc_Get_NestLevel()`) |
| 9 | Con `Sw_Test` = FALSO | Llega a los destinatarios reales y el temporal ya no está en `%TEMP%` |
| 10 | Botones de JIs | Hacen lo mismo que las formas (también la pregunta de EP de `M_132`), solo se ven en `Inf_Recibos_TIO` |
| 11 | «Sí» a copiar en la carpeta de red | El `.xlsx` queda en `App_Ruta_Inf_Recibos` (unidad del usuario + `C10`) con el nombre del adjunto, y el informe del formulario lo dice |
| 12 | Repetir el envío | La pregunta avisa de que el fichero de la red se sustituirá |
| 13 | El fichero de la red abierto por otro (o la unidad desconectada) | Aviso por delante del formulario; el correo se envía igual (desconectada: pregunta si enviarlo de todos modos) |

## 6. Decisiones del usuario (2026-10-12)

1. **Título:** hecho en la hoja («Informe Recibos 2026», `B2:C2`); la copia no lo toca.
2. **Totales en el cuerpo del correo:** no.
3. **Quién ve los botones:** todos los usuarios de *Recalcular* (`Boss,SusanaC,FannyR;RafaG`).
4. **Destinatarios:** `Prog__APP.Range("App_Mail_Inf_Recibos")`.
5. **Remitente:** `ingresos@ua.es`.
6. **Supuesto desfase de los títulos del bloque EURLE:** no existe. Salía de un fallo del script con el que leí el
   `.xlsm` (se comía las filas vacías y corría los textos una fila); la captura y la relectura confirman que los
   títulos están en la fila 31, combinados sobre su cabecera. No hay nada que arreglar.
7. **Rango:** B:Z en vez de B:R.
8. **Formas de la hoja:** pasan a botones del Ribbon.
9. **Configuración del correo:** en la hoja `ConfigMail` (`Prog__APP_Mail`).

## 7. Visto al preparar el plan: 🧪 arreglado el 2026-10-13 (falta importar y probar)

- 🧪 **Errata que llegaría al destinatario:** `D2` decía «Fecha de **cierro** contable». `M_130` escribe ya «cierre».
- 🧪 **Título fijo:** «Informe Recibos 2026» estaba escrito a mano en `B2`. `M_130` lo escribe ahora con `APP_AnoCont`
  al final del cálculo, junto a `D2`, en la ancla `B2` de `B2:C2`.
- 🧪 **`M_132` y `M_133`** guardaban `Application.Calculation` en un `Boolean` (`Sw_Calculation`) y al terminar no
  recuperaban el modo de cálculo. Ahora es `XlCalculation`. `M_120`, `M_210`, `M_310` y `M_602` tienen lo mismo y siguen sin tocar.

## 8. Riesgos

1. **Nombres en la hoja equivocada** (2.2): los rotos ya no están, pero el código nuevo cualifica siempre con la hoja
   donde vive el nombre (`Prog__APP_Mail` para el correo, `Prog__APP` para usuario y destinatarios).
2. **Eventos encendidos durante la copia:** error de compilación del código de hoja copiado o Excel restablecido a
   mitad (3.1, paso 2).
3. **Modo de cálculo guardado en el adjunto.** El `.xlsx` se guarda con Excel en Manual (`Rut_Off_Functions`), y un
   libro guarda el modo de cálculo de la aplicación: si el destinatario lo abre el primero, su Excel podría quedar en
   Manual. Comprobarlo en la fase 2 (`calcPr` de `xl/workbook.xml`) y, si sale `manual`, poner Automático solo
   alrededor del `SaveAs`. `Rut_Lo_Export_WS_Xlsx` tendría lo mismo.
4. **Libro homónimo abierto:** si hay abierto un `Inf_Recibos_<…>.xlsx` con el mismo nombre (p. ej. un adjunto anterior
   abierto desde el correo), el `SaveAs` falla. Comprobarlo antes y, si está abierto, añadir la hora al nombre.
5. **`MsgBox` detrás del formulario:** ningún `MsgBox` sin `vbMsgBoxSetForeground + vbSystemModal` mientras
   `Form_Running_Rut` está abierto (3.2, 3.4).
6. **Cambios en `M_130`:** el recorte solo depende de `InfRec_Descrip_Contab` (última columna, Z). Si `M_130` mueve
   columnas, hay que revisarlo.
7. **La contraseña viaja dentro del `.xlsm`**, y con él en las copias de seguridad (`CopiaSeguridad\`, USB), igual que
   en JyC y Certificados. El `.xlsm` no está en git (`.gitignore`), y los volcados `CustomUI/*.csv` no leen `ConfigMail`.
