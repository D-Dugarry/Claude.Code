# Informe de Bugs — EP_202x-2x_BaseDatos-LIQ _V3.xlsm

Auditoría del código VBA exportado en `VBA_Moduls/` (Enseñanzas Propias — Liquidación de Títulos Propios, Universidad de Alicante). Generado el 14/09/2026 mediante lectura íntegra de los 133 módulos exportados, con verificación cruzada de los hallazgos más graves contra el fichero real.

**Resumen:** 10 Críticos · 11 Altos · 10 Medios · 7 Bajos — 39 hallazgos (20 corregidos, 4 desactivados/aparcados, 1 descartado).

> **Actualización 2026-09-19 (3ª):** el proyecto **compila limpio** tras esta sesión. Añadidos **C13** y **C14** (dos errores de compilación más, detectados y corregidos sobre la marcha: `Dictionary` ambiguo entre Scripting Runtime/Word, y `ListColumns.Add` con un argumento `Name:=` inexistente). **B3 pasa a desactivado** (módulos `M41`/`M42` comentados enteros, mismo patrón que B5/B6/B19), confirmado por el usuario como importación pendiente de adaptar desde otra app.
>
> **Actualización 2026-09-18:** añadido **B16**, detectado al compilar el proyecto tras corregir A3. Además de los bugs, se ha hecho una **limpieza de código muerto** en los módulos `Rut_*` (ver el apéndice al final).

> **Nota sobre codificación:** los ficheros `.bas`/`.cls`/`.frm` están en CP1252, no UTF-8. Antes de aplicar cualquier corrección directamente sobre `VBA_Moduls/`, edítalos siempre con un script que preserve CP1252/CRLF — nunca con el Editor de texto plano ni herramientas UTF-8 (ver `CLAUDE.md` de este proyecto).

## Antes de tocar nada: compila el proyecto

Los 8 hallazgos Críticos de este informe son, cada uno por separado, un error de compilación VBA (`Option Explicit` roto, nombre duplicado, tipo incompatible, `Sub`/`End Sub` descuadrados). **Un solo error de compilación en cualquier módulo bloquea la ejecución de todas las macros del libro**, no solo la rutina afectada — VBA compila el proyecto entero, no módulo a módulo. Antes de corregir nada, abre el editor VBA del `.xlsm` real y ejecuta **Depurar → Compilar VBA Project**, y anota qué marca primero: puede que algunos de estos 8 puntos ya estén corregidos ahí y esta exportación esté desactualizada, o puede que el libro lleve tiempo sin compilar limpio (varios de los módulos implicados — `M50_Inf_Cont_AE4x41`, `M51_Import_AE4x11`, `M41`/`M42` — parecen ramas de desarrollo/copias aún no terminadas, no el camino que se ejecuta a diario).

## Índice

**Leyenda:** ✅ Corregido · ⏸️ Desactivado/aparcado (módulo comentado, decisión pendiente) · ❌ Descartado (no era bug) · ○ Pendiente

- **Bloque A — Arranque, pipeline LSGES04 y núcleo de Liquidación**
  - ✅ A1 · Importe cobrado truncado a `Long`
  - ✅ A2 · `Workbooks(IntialName).Close` falla si el usuario cambia el nombre al guardar
  - ✅ A3 · Merge-join sin `Case Else` — desincronización silenciosa al guardar la Liquidación
  - ✅ A4 · Constante equivocada al ordenar la tabla de Coeficientes VRI
  - ✅ A5 · Columnas de flags de Tipo_Recibo hardcodeadas (52–56)
  - ✅ A6 · `.EntireRow.Delete` en vez de `.Delete`, único caso del pipeline
  - ✅ A7 · Mensaje de informe copiado y mal etiquetado
  - ○ A8 · Cualificación inconsistente de `Range("Sw_VerRecNeg")` (a verificar)
  - ❌ A9 · `Coef_VRI` declarado `Integer` — descartado
- **Bloque B — Informes, Cierre Contable y AE4x4/AE4x1**
  - ✅ B1 · Sub pública duplicada: `RuT_Inf_Contable_Recibos_AE4x4`
  - ✅ B2 · Sub pública duplicada: `Rut_Lo_Import_AE4x1`
  - ⏸️ B3 · Objetos de hoja inexistentes en este libro (código de PPub sin adaptar) — módulos desactivados
  - ✅ B4 · `End Sub` huérfano en `M90_Rutinas_X.bas`
  - ⏸️ B5 · Variable de objeto `Lo_AE4x1` no declarada — módulo desactivado
  - ⏸️ B6 · `Lo_AE4x1` usado sin inicializar — módulo desactivado
  - ○ B7 · Límite de filas hardcodeado a 5000
  - ○ B8 · Filtro roto por referencia sin cualificar y variable de bucle equivocada
  - ○ B9 · `Application.Calculation` guardado en variable `Boolean`
  - ○ B10 · Año "2025" hardcodeado en nombres de fichero exportado
  - ○ B11 · Módulos M22 OLD/NEW con agrupación distinta sobre la misma hoja destino (a verificar)
  - ○ B12 · `M71_Restituir` sobrescribe datos y parámetros sin validar ni confirmar
  - ○ B13 · `Wk_Lista_Panes1.cls` sin ninguna referencia en el código (a verificar)
  - ○ B14 · Bucle `Do While` sin cota superior en M31/M32/M33
  - ○ B15 · Contraseña de correo en texto plano
  - ✅ B16 · Cuatro rutinas de otro libro en `M90_Rutinas_X.bas` (hallazgo posterior)
  - ✅ B17 · `ActivForm` usada en 6 módulos y declarada en ninguno
  - ✅ B18 · `AñoContAnt` sin declarar en `M20_Resumen_Tit_Propios.bas`
  - ⏸️ B19 · `Cod_Plan` sin declarar y vaciado antes de usarse (`M21_Resumen_Tit_Propios_UNO.bas`) — módulo desactivado
- **Bloque C — Librería transversal `Rut_*`/`Prog_*` y formularios**
  - ○ C1 · Variable no declarada `LoTb` (typo de `Lo_Tb`)
  - ○ C2 · `.calcMode` no es un miembro de `Application`
  - ✅ C3 · `String` pasado donde se espera `Worksheet` por referencia (ruta en producción)
  - ✅ C4 · Llamada a una rutina que no existe: `Rut_Actualizar_1_LS_VAL`
  - ○ C5 · `Rut_WrkSheet_ReducirPeso` opera sobre la hoja activa, no sobre la recibida
  - ○ C6 · Pérdida de datos en `M0999_Modif_Cols_BDatos.bas` si se reejecuta
  - ○ C7 · Valor mágico `-0.86` escrito sobre datos reales sin confirmación
  - ○ C8 · `For Each` que ignora la variable de iteración
  - ○ C9 · `Módulo3.bas` sin `Option Explicit`, con variables casi homónimas
  - ○ C10 · Asimetría Private/Public en rutinas invocadas por nombre (a verificar)
  - ○ C11 · Ruta de disco hardcodeada como fallback silencioso
  - ✅ C12 · `.UsedRange` como instrucción suelta — propiedad usada como si fuera un método
  - ✅ C13 · `Dictionary` ambiguo entre Scripting Runtime y Word Object Library
  - ✅ C14 · `ListColumns.Add` con un argumento `Name:=` que no existe

---

## Bloque A — Arranque, pipeline LSGES04 y núcleo de Liquidación
*(ThisWorkbook, M00, M01–M09, M10, M12, M15, Wk_TitP_Liquid)*

### A1 · Importe cobrado truncado a `Long`
**Severidad:** Alto · **Fichero:** `M09_Importar_Sol_Liq.bas` — líneas 129, 151, 187 · **Estado:** ✅ Corregido (2026-09-14)

```vba
    Dim Cobrado     As Long:        Cobrado = 0
...
                    Cobrado = Cobrado + .Cells(LinLiq, CLiq_Imp_Cob)
...
        Right("__________" & Cobrado, 8) & "  Importe Cobrado en esta Liquidación." & vbCrLf
```

`Cobrado` acumula un importe monetario (con decimales) en una variable `Long`. Cada suma trunca los céntimos del `Currency`/`Double` de la celda, así que el total mostrado en el informe final ("Importe Cobrado en esta Liquidación") puede no coincidir con la suma real.

**Impacto:** el informe que ve el usuario tras importar una Solicitud de Liquidación muestra un importe cobrado incorrecto (redondeado a la baja), sin que nada avise del desajuste.

**Arreglo:** declarar `Cobrado As Currency` (o `Double`). Aplicado en `VBA_Moduls/M09_Importar_Sol_Liq.bas` (línea 130 tras el sello `Last Rev.`); las líneas de acumulación y de informe no necesitaron cambios.

### A2 · `Workbooks(IntialName).Close` falla si el usuario cambia el nombre al guardar
**Severidad:** Alto · **Fichero:** `M12_Genera_LIQx_PDF.bas` — líneas 574, 582, 590-591 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-15)

```vba
    IntialName = "LIQxPDF_" & Wk_TitP_Liquid.Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"
    FullName = FPath & "LIQxPDF_" & Wk_TitP_Liquid.Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(FullName, "Excel Files (*.xlsx), *.xlsx")
...
            On Error GoTo 0
...
'    ActiveWorkbook.Close savechanges:=True
    Workbooks(IntialName).Close SaveChanges:=True
```

`IntialName` es solo el nombre **sugerido** para el diálogo `GetSaveAsFilename`; si el usuario lo cambia al guardar, el libro copiado queda con otro nombre y `Workbooks(IntialName)` no existe → error 9 "Subscript out of range", **sin gestor de errores activo** (el `On Error GoTo GestError` ya se ha desactivado con `On Error GoTo 0` dos líneas antes). La línea correcta original (`ActiveWorkbook.Close savechanges:=True`) sigue comentada justo encima, confirmando que es una regresión: las rutinas hermanas `Rut_Exportar_La_Liquidación` (M15) y `Rut_Exportar_La_LIQxn_PDF` (M12) sí usan `ActiveWorkbook.Close`.

**Impacto:** generar el PDF de Liquidación falla con un error críptico en cuanto el usuario cambia el nombre propuesto en el diálogo de guardado, y el libro temporal queda abierto sin cerrarse.

**Arreglo:** restaurada `ActiveWorkbook.Close SaveChanges:=True` en `VBA_Moduls/M12_Genera_LIQx_PDF.bas` (línea 591 tras el sello `Last Rev.`), eliminando la línea rota y la comentada que quedaba redundante.

### A3 · Merge-join sin `Case Else` — desincronización silenciosa al guardar la Liquidación
**Severidad:** Alto · **Fichero:** `M10__Liquid_EP.bas` — líneas 742-778 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-18)

```vba
    For F_Liquid = 1 To Lo_TPLiquid.DataBodyRange.Rows.Count   '--- Bucle para recorrer todas la filas de la Liquidación

        Select Case Lo_TPLiquid.DataBodyRange.Cells(F_Liquid, CLiq_Ref)

            '--- Saltar al siguiente Lo_TitPHist   <<<...
            Case Is > .Cells(F_TitPH, BD_Ref)
                F_Liquid = F_Liquid - 1
            '--- ACTUALIZAR CON Datos de la Liquidación <<<...
            Case Is = .Cells(F_TitPH, BD_Ref)

                Set RowLiq = Lo_TPLiquid.ListRows(F_Liquid)
                Set RowDat = Lo_TitPHist.ListRows(F_TitPH)

                RowDat.Range(BD_Liquidado) = RowLiq.Range(CLiq_NumLiquid)
                RowDat.Range(BD_JI_Emi_Acad) = RowLiq.Range(CLiq_JI_Emi)
                ...
        End Select

        F_TitPH = F_TitPH + 1

    Next F_Liquid
```

Es un merge-join de dos tablas ordenadas por `Ref` (línea 738/740), pero falta el caso `Case Is < .Cells(F_TitPH, BD_Ref)` (o `Case Else`). Cuando la Ref de la Liquidación es **menor** que la de `Prog_BD` en la posición actual, el `Select Case` no entra en ninguna rama — pero `F_TitPH = F_TitPH + 1` se ejecuta igualmente, **fuera** del `Select`, avanzando el puntero de `Prog_BD` sin haber comparado nada. Esto desalinea ambos punteros para el resto del recorrido.

**Impacto:** al guardar una Liquidación, dejan de actualizarse silenciosamente en `Prog_BD` los campos JI, AD-0010, JI-443, ExpAdm, RDT, Coef_VRI, Orgánica y Tasa Administrativa de registros que sí deberían actualizarse — **sin ningún aviso ni contador de "no actualizados"**, a diferencia de `M07` (importación LSGES04), que sí avisa con `MsgBox` cuando algo no cuadra.

**Objeción considerada — "el `Case Else` nunca puede darse":** es cierto que, por diseño, todas las filas de la Liquidación salen del Histórico y no se añaden filas nuevas, es decir `Ref_Liquid ⊆ Ref_BD`. Pero el `Select Case` no compara conjuntos, compara **la posición de dos punteros que avanzan a distinto ritmo**, y hay dos vías por las que se llega al caso `<` con la premisa intacta:

1. **Desbordamiento de `F_TitPH` (verificado por simulación).** `F_TitPH` se incrementa en cada iteración del `For`, incondicionalmente, y nada comprueba que siga dentro de la tabla. `.Cells(F_TitPH, BD_Ref)` con `F_TitPH` mayor que el número de filas **no da error**: lee celdas vacías de debajo del `DataBodyRange`. Y `"R99" > Empty` es `True`, así que entra en `Case Is >`, `F_Liquid` retrocede, `F_TitPH` sigue subiendo → **bucle infinito**. Basta con que el último `Ref` de la Liquidación sea mayor que el último del Histórico.
2. **Orden no idéntico entre las dos tablas.** `Rut_Lo_Sort` ordena con `xlSortNormal`; si la columna `Ref` no es homogénea de tipo entre `Prog_BD` (col. 11) y la Liquidación (col. 19) — unos valores numéricos y otros texto —, Excel coloca los números antes que el texto y el mismo `Ref` cae en posición relativa distinta en cada tabla.

**Demostración del impacto silencioso:** simulando el código original con `Prog_BD = [R01..R06]` y `Liquidación = [R02, R03b, R04]`, sólo se actualiza `R02`: **`R04` existe en ambas tablas y no se actualiza**, y el `MsgBox` final sigue diciendo "¡¡¡Hecho!!!". Con el arreglo se actualizan `R02` y `R04`, y se reporta únicamente `R03b`.

**Arreglo aplicado:** no sólo se recupera el paso del merge-join, sino que se convierte el fallo silencioso en un fallo visible:

- Guard `If F_TitPH > N_TitPH` al principio del bucle, que corta el bucle infinito y contabiliza como "no emparejado" el resto de la Liquidación.
- `Case Else` con `F_TitPH = F_TitPH - 1` (compensa el `+1` de abajo: no se avanza en el Histórico) que anota la `Ref` no encontrada y continúa.
- Contador `Reg_NoEmparejados` + lista de las 10 primeras `Ref`, mostrados en un `MsgBox` de advertencia **antes** del "¡¡¡Hecho!!!", en la línea de lo que ya hace `M07`.

Si la premisa de diseño se cumple siempre, el coste en ejecución es cero y ninguno de los dos avisos llega a aparecer.

### A4 · Constante equivocada al ordenar la tabla de Coeficientes VRI
**Severidad:** Medio · **Fichero:** `M08_Actualizar_Tb_Coef_VRI.bas` — línea 96 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-18)

```vba
    ' --- líneas 48-49, correctas ---
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)
    ...
    ' --- líneas 95-96 y 105-106, repetidas igual en ambos bloques finales ---
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, BD_Plan, xlAscending, False)
```

`Lo_RetVRI` es `Prog_Coef_Ret_VRI`, cuyo esquema de columnas es `CoefVRI_*` (con `Plan` en la columna 1); `BD_Plan` (=3) es la constante de la tabla `Prog_BD`, no la de esta tabla. Al reutilizarla aquí por error se ordena en realidad por `CoefVRI_Concepto` (columna 3) en vez de por `CoefVRI_Plan`. El bloque de preparación inicial sí lo hace bien; el error, con toda pinta de copia-pega, solo está en los dos bloques finales.

**Impacto:** acotado — solo afecta al desempate de filas nuevas (todas con `ORden="x"`) — pero es un fallo real y reproducible.

**Arreglo aplicado:** sustituido `BD_Plan` por `CoefVRI_Plan` en la línea 96. La línea 106 (bloque `Restablecer_Valores:`) ya usaba `CoefVRI_Plan` correctamente, así que la única ocurrencia errónea era la 96.

### A5 · Columnas de flags de Tipo_Recibo hardcodeadas (52–56)
**Severidad:** Medio · **Ficheros:** `M05_Asign_Tipo_Recibo.bas` — líneas 35-40, 59-65 · `M00_Ini_Var_APP.bas` · **Estado:** ✅ Corregido (2026-09-18)

```vba
        .DataBodyRange.Columns(BD_Tipo_Rec).ClearContents
        .DataBodyRange.Columns(52).Resize(, 56).ClearContents
        '.DataBodyRange.Columns(BD_CriT_Emi).Resize(, BD_CriT_ErrDate - BD_CriT_Emi + 1).ClearContents
...
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
            .DataBodyRange.Columns(52).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
```

Las columnas 52-56 de `Prog_LsGes04` se usan como "flags" internos por tipo (Emitido=52, EjeAnt=53, Añejo=54, Aplazado=55, ADxAplz=56), con números literales en vez de constantes con nombre. La línea comentada `BD_CriT_Emi`/`BD_CriT_ErrDate` demuestra que antes existían constantes específicas para este rango, ya no presentes en `M00_Ini_Var_APP.bas`. Además, esos mismos números **coinciden** con `BD_H_Incidencias` (52), `BD_EP_Ctrl` (53) y `BD_EP_GestReg` (54) del esquema de `Prog_BD` — aquí se reutilizan como scratch space en `Prog_LsGes04` (una tabla distinta) después de que `M02_Manage_Duplicates` ya haya usado la columna 52 como historial de diferencias de duplicados.

**Impacto:** si en el futuro se añade o quita una columna a `Lo_Ges04` (el propio pipeline añade y borra columnas temporales en varios puntos), estos números fijos apuntarán a la columna equivocada sin que salte ningún error.

**Bug adicional detectado al arreglarlo (línea 39):** `.DataBodyRange.Columns(52).Resize(, 56).ClearContents` **no limpia las columnas 52-56**, sino las **52-107**: el segundo argumento de `Resize` es el *número de columnas* del rango resultante (el ancho), no la columna final. Limpiaba por tanto 56 columnas a partir de la 52, arrasando todo lo que hubiera más allá del rango de flags dentro de la tabla. El equivalente correcto es `Resize(, 5)`.

**Arreglo aplicado:** definidas las 5 constantes con nombre en `M00_Ini_Var_APP.bas`, justo tras el esquema de `Prog_BD`, más dos derivadas para el `ClearContents` (así el rango a limpiar se recalcula solo si algún día cambia el bloque):

```vba
'--- Flags internos de Tipo_Recibo en Prog_LsGes04 (M05_Asign_Tipo_Recibo) ------
'    Columnas de marca por tipo; el valor definitivo va en BD_Tipo_Rec.
Public Const G04_Flag_Emitido      As Integer = 52   ' col: az
Public Const G04_Flag_EjeAnt       As Integer = 53   ' col: ba
Public Const G04_Flag_Anejo        As Integer = 54   ' col: bb
Public Const G04_Flag_Aplazado     As Integer = 55   ' col: bc
Public Const G04_Flag_ADxAplz      As Integer = 56   ' col: bd
Public Const G04_Flag_Primera      As Integer = G04_Flag_Emitido
Public Const G04_Flag_Cuantas      As Integer = G04_Flag_ADxAplz - G04_Flag_Emitido + 1
```

En `M05` se han sustituido los 6 literales por las constantes, y el `ClearContents` pasa a ser `.DataBodyRange.Columns(G04_Flag_Primera).Resize(, G04_Flag_Cuantas).ClearContents` — que ahora sí limpia exactamente las 5 columnas de flags, corrigiendo de paso el `Resize(, 56)`. Se han eliminado también las dos líneas comentadas obsoletas que referenciaban `BD_CriT_Emi`/`BD_CriT_ErrDate` (constantes ya inexistentes).

**Nota de diseño:** estas 5 columnas **solo se escriben, nunca se leen** en ninguno de los 131 módulos (el valor operativo es `BD_Tipo_Rec`, que recibe el mismo texto una línea más arriba en cada bloque). Se han conservado —en vez de eliminarlas— por decisión del usuario, al servir de marca visible en la propia hoja. Los números **coinciden** con `BD_H_Incidencias` (52), `BD_EP_Ctrl` (53) y `BD_EP_GestReg` (54) del esquema de `Prog_BD`, pero se trata de tablas distintas: los nuevos nombres `G04_*` dejan explícito que pertenecen a `Prog_LsGes04`.

### A6 · `.EntireRow.Delete` en vez de `.Delete`, único caso del pipeline
**Severidad:** Medio · **Fichero:** `M02_Del_Reg_No_Válidos.bas` — línea 86 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-14)

```vba
        .Range.AdvancedFilter xlFilterInPlace, Prog_Filtros_TipRec.Range("Tb_CriT_ImpMatCeroLsGes04")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        If rowfind > 0 Then
'            On Error Resume Next
            .DataBodyRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete
'            On Error GoTo 0
                TxtMsg1 = "Borrados Recibos Matrícula de Actividad Acad. a Coste CERO"
```

Todos los demás borrados de filas filtradas del mismo pipeline usan `.Delete` sobre el rango de la tabla, no `.EntireRow.Delete`: confirmado al comparar contra las 4 ocurrencias hermanas del propio `M02_Del_Reg_No_Válidos.bas` (líneas 33, 49, 68, 103 antes del arreglo) y la de `M02_Del_Reg_EFP_o_CFCyAFC.bas`, todas `.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete` sin `.EntireRow`. Este último borra la fila **completa de la hoja** (todas las columnas, no solo las de la tabla) y desplaza el resto de filas de la hoja entera, no solo de la tabla.

**Arreglo aplicado:** quitado `.EntireRow` — ahora `.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete`, igual que las otras 5 ocurrencias del mismo patrón en el pipeline.

### A7 · Mensaje de informe copiado y mal etiquetado
**Severidad:** Bajo · **Fichero:** `M05_Asign_Tipo_Recibo.bas` — líneas 136-159 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-14)

```vba
        '------------------ Filtra Cobradas en Años anteriores al de Emisión -------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .Range.AutoFilter Field:=BD_ACont_Cob, Criteria1:="<" & APP_AñoCont
        ...
        TxtProgreso = TxtProgreso & vbCrLf & ... & " Registros _Contab_Ant_, Cobrados anteriormente y por lo tanto, ya Contabilizado."

        '------------------ Filtra Cobradas en Años anteriores al de Emisión -------------------------------------   ← comentario IDÉNTICO
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .Range.AutoFilter Field:=BD_ImpAdm, Criteria1:="<0"                                                          ← pero el filtro real es OTRO
        ...
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_Ajust_Matríc_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & ... & " Registros _Contab_Ant_, Cobrados anteriormente y por lo tanto, ya Contabilizado."   ← texto también copiado
```

El segundo bloque filtra en realidad `BD_ImpAdm < 0` (Ajustes de Matrícula), pero conserva sin tocar el comentario de cabecera y el texto de informe del bloque anterior. El usuario que lee el log de ejecución ve una descripción que no corresponde al filtro realmente aplicado.

**Arreglo aplicado:** cabecera de comentario `'-_Ajust_Matríc_-...`, descripción `'--- Filtra Ajustes de Matrícula (ImpAdm negativo) ---...` y mensaje de progreso `"Registros _Ajust_Matríc_, Ajustes de Matrícula (Imp. Admin. negativo)."` — ya no repiten el texto del bloque `_Contab_Ant_` anterior.

### A8 · Cualificación inconsistente de `Range("Sw_VerRecNeg")` (a verificar)
**Severidad:** Medio, a verificar · **Fichero:** `M10__Liquid_EP.bas`

El módulo mezcla, en distintos puntos, `Prog__APP.Range("Sw_VerRecNeg")` (cualificado con la hoja) con `Range("Sw_VerRecNeg")` sin cualificar dentro de `Rut_03_Generar_Tabla_RDT_x_NumLiquid_Con_Devoluciones`, que se dispara desde `Worksheet_SelectionChange` de `Wk_TitP_Liquid` con esa hoja como activa, no `Prog__APP`. El propio `CLAUDE.md` de este proyecto documenta que `Hoja.Range("Nombre")` revienta con error 1004 si el nombre vive en otra hoja aunque sea de ámbito Libro; el caso simétrico (`Range()` sin cualificar resolviendo por la hoja activa) tiene el mismo riesgo si `Sw_VerRecNeg` resultara tener ámbito de hoja. No se ha podido verificar el ámbito real del nombre definido (vive en el `.xlsm`, no en el texto exportado): queda como sospecha, no como bug confirmado.

### A9 · `Coef_VRI` declarado `Integer` — descartado
**Severidad:** Bajo · **Ficheros:** `M00_Ini_Var_APP.bas` línea 232, `M08_Actualizar_Tb_Coef_VRI.bas` línea 11 · **Estado:** ❌ Descartado (2026-09-15) — no es un bug

`Coef_VRI` (porcentaje de retención VRI) se declara `Integer` en ambos sitios. Con los valores observados en el código (15, 20) es correcto, y quedaba como sospecha por si algún Plan tuviera un coeficiente no entero en `Prog_Coef_Ret_VRI` (p. ej. 17,5%), que se truncaría/redondearía silenciosamente al leerlo.

**Verificado por el usuario:** `Coef_VRI` solo contiene números enteros — es una retención fijada por el Vicerrectorado de Investigación en puntos porcentuales enteros, nunca con decimales. `Integer` es el tipo correcto; no hay truncamiento posible en la práctica.

---

## Bloque B — Informes, Cierre Contable y AE4x4/AE4x1
*Ficheros: M20–M22, M31–M39, M40–M44, M50–M51, M71–M72, M79–M80, M90, `Sht__Inf_*`, `Wk_Lista_Panes*`.*

### B1 · Sub pública duplicada: `RuT_Inf_Contable_Recibos_AE4x4`
**Severidad:** Crítico · **Ficheros:** `M50_Inf_Cont_AE4x4.bas:8` y `M50_Inf_Cont_AE4x41.bas:8` · **Estado:** ✅ Corregido (2026-09-15)

```vba
Sub RuT_Inf_Contable_Recibos_AE4x4()  '- Importar los 4 WB: EFP y CFC de AñoCon_Ant/Pos
```

Dos módulos estándar distintos declaran, ambos sin `Private`, una `Sub` pública con el nombre exacto — verificado con `grep` sobre los dos ficheros. Es un error de compilación de VBA ("Ambiguous name detected") que impide compilar el proyecto **entero**, no solo estas rutinas.

**Arreglo:** eliminado el módulo `M50_Inf_Cont_AE4x41.bas` (la rama de desarrollo: le faltaba el log de progreso `Rut_TimeLap_Inf` y el `Set ActivForm`), tanto del proyecto VBA real (`vbaProject.bin`, confirmado con `oletools`) como de `VBA_Moduls/`. Antes de borrar se auditaron macros de shapes, Ribbon y la tabla `Tb_Tareas` del menú auxiliar: ninguno de los dos módulos estaba enganchado a la UI.

### B2 · Sub pública duplicada: `Rut_Lo_Import_AE4x1`
**Severidad:** Crítico · **Ficheros:** `M51_Import_AE4x1.bas:8-13` y `M51_Import_AE4x11.bas:8-13` · **Estado:** ✅ Corregido (2026-09-15)

```vba
Sub Rut_Lo_Import_AE4x1(Ws_AE4x1 As Worksheet, _
                                  Arch_New_Name As String, _
                                  SheetNom As String, _
                                  Optional NameFileAE4 As String, _
                                  Optional PathFileAE4 As String, _
                                  Optional SW_Inicilizar_Ws As Boolean = False)
```

Misma firma, misma Sub pública, en dos módulos distintos — verificado. Segundo "Ambiguous name detected" independiente del anterior, agravando el bloqueo de compilación.

**Arreglo:** eliminado el módulo `M51_Import_AE4x11.bas` (misma rama de desarrollo que B1), tanto del proyecto VBA real como de `VBA_Moduls/`, manteniendo `M51_Import_AE4x1.bas`. ⚠️ Ese módulo superviviente sigue teniendo un bug de compilación independiente y ahora más urgente — ver B5.

### B3 · Objetos de hoja inexistentes en este libro (código de PPub sin adaptar)
**Severidad:** Crítico · **Ficheros:** `M41_Añadir_Núm_JIs_al_Inf.bas:13,15` y `M42_Añadir_Núm_JIs_a_BDatos.bas:13,15,32,40` · **Estado:** ⏸️ Módulos desactivados (2026-09-19) — el bug de fondo sigue sin resolver

```vba
        Dim Sht_Inf         As Worksheet:       Set Sht_Inf = Sht__Inf_Recibos_TIO
        Dim Lo_Inf          As ListObject:      Set Lo_Inf = Sht_Inf.ListObjects(1)
        Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
```

Ni `Sht__Inf_Recibos_TIO` ni `Sht__BD` existen como CodeName de ninguna hoja de este proyecto (la tabla de datos aquí se llama `Prog_BD`, no `Sht__BD`, y no hay ninguna hoja "TIO"). Con `Option Explicit`, "Variable no definida" en compilación. Todo apunta a que `M41`/`M42` son código copiado literalmente del proyecto hermano *Precios Públicos - PPub* (que sí tiene esas hojas) y nunca adaptado a Enseñanzas_Propias — coherente con que M42 también reescriba `Prog_BD`/`Sht_Inf` en la línea 32-40 mezclando nombres de ambos proyectos.

**Impacto:** el flujo de "añadir números de JI al informe / a la base de datos", tal como está exportado hoy, no compila en este libro.

**Decisión aplicada (2026-09-19):** mismo patrón que B5/B6/B19 — el usuario confirmó que ambos módulos son una importación a medio adaptar desde otra app, y prefirió **comentarlos enteros** (0 líneas de código vivo) en vez de intentar adivinar la hoja correcta. De cada fichero solo quedan sin comentar `Attribute VB_Name`, `Option Explicit` y una cabecera con el sello `Last Rev.` explicando el motivo. Auditados por las 4 vías: sin invocación viva en código, shapes ni `Tb_Tareas`.

**Para reactivar en el futuro:** sustituir `Sht_Inf_Recibos_TIO` por la hoja EP correspondiente (`Sht__Inf_EFP_ACont1_CAcad` / `_2_CAcad` / CFC) y `Sht__BD` por `Prog_BD`, verificando antes con un volcado real de `ThisWorkbook.VBComponents` u otra confirmación directa — no adivinar por el nombre.

### B4 · `End Sub` huérfano en `M90_Rutinas_X.bas`
**Severidad:** Crítico · **Fichero:** `M90_Rutinas_X.bas` — líneas 193-203 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-18)

```vba
    .AutoFilter Field:=CTa_Ref, Criteria1:=Application.Transpose(Lo_Busca_JI.DataBodyRange.Columns(2)), Operator:=xlFilterValues

    End With

End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' ==================================================================================================================================
End Sub
' ==================================================================================================================================
Sub Rut_Quita_Ascii_160(ByRef Lo_Tb As ListObject, columna As Integer)
```

La línea 197 ya cierra correctamente `Rut_Filtro_NumJI_2_en_Tasas`; la línea 200 es un segundo `End Sub` sin ningún `Sub`/`Function` abierto que cerrar. Error de sintaxis, impide compilar el módulo.

**Arreglo aplicado:** borrada la línea 200. El módulo tenía 19 `Sub` / 20 `End Sub`; ahora 15/15 tras esta corrección y la de B16.

### B5 · Variable de objeto `Lo_AE4x1` no declarada
**Severidad:** Crítico · **Fichero:** `M51_Import_AE4x1.bas` — líneas 25, 93 · **Estado:** ⏸️ Módulo desactivado (2026-09-18) — el bug sigue sin resolver

```vba
    Dim Ws_AE4          As Worksheet:   Set Ws_AE4 = Lo_AE4x1.Parent
...
    Else
        Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_AE4x1, False)
    End If
```

`Lo_AE4x1` no se declara en ningún punto de este módulo. Nota (2026-09-15): tras resolver B2, el módulo gemelo `M51_Import_AE4x11.bas` —que tenía el único `Dim Lo_AE4x1 As ListObject` local del proyecto, aunque no aplicaba aquí por ser otro módulo— ya no existe. Este hallazgo sigue exactamente igual de crítico: `Option Explicit` sigue rompiendo la compilación de `M51_Import_AE4x1.bas` con "Variable no definida".

### B6 · `Lo_AE4x1` usado sin inicializar
**Severidad:** Alto · **Fichero:** `M51_Import_AE4x11.bas` — líneas 90-98

```vba
    If SW_Inicilizar_Ws Then       '- Es el 1º, sólo copiar Lo_ClsBk en Ws_AE4x1
        Call Rut_WrkSheet_Vaciar(Ws_AE4x1.Name)
        Lo_ClsBk.Range.Copy Destination:=Ws_AE4x1.Range("A1")
        Set Lo_AE4x1 = Ws_AE4x1.ListObjects(1)
    Else
        Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_AE4x1, False)
    End If
```

`Lo_AE4x1` solo se asigna dentro de la rama `If SW_Inicilizar_Ws`; la rama `Else` lo usa sin haberlo asignado nunca. Los 8 puntos de llamada (4 en `M50_Inf_Cont_AE4x4.bas` y 4 en `M50_Inf_Cont_AE4x41.bas`) pasan siempre `False` como último argumento, así que `Lo_AE4x1` permanecería `Nothing` en tiempo de ejecución. Aunque B1/B2/B3/B5 ya impiden compilar, esto confirma que el flujo AE4 está incompleto también a nivel lógico: falta pasar `True` en la primera de las 4 importaciones.

Nota (2026-09-15): tras resolver B2, este fichero (`M51_Import_AE4x11.bas`) ya no existe. El problema de fondo persiste, y de forma más grave, en el módulo superviviente `M51_Import_AE4x1.bas`: ahí la rama `If SW_Inicilizar_Ws` tampoco asigna `Set Lo_AE4x1 = ...` (solo copia el rango), así que `Lo_AE4x1` no se asigna en NINGUNA rama — coincide con B5. Las 4 llamadas restantes (solo desde `M50_Inf_Cont_AE4x4.bas`) siguen pasando `False`.

### B5/B6 · Decisión: módulos `M50` y `M51` desactivados por completo (2026-09-18)

Ante la imposibilidad de compilar el proyecto con el flujo AE4x4/AE4x1 en este estado, se han **comentado íntegramente** los dos módulos implicados:

| Módulo | Rutina pública única | Líneas |
|---|---|---|
| `M50_Inf_Cont_AE4x4.bas` | `RuT_Inf_Contable_Recibos_AE4x4` | 114 |
| `M51_Import_AE4x1.bas` | `Rut_Lo_Import_AE4x1` | 127 |

De cada fichero solo quedan sin comentar la línea `Attribute VB_Name` (obligatoria para poder reimportar el módulo), el sello `Last Rev.` y una cabecera nueva que explica el motivo de la desactivación y qué hay que resolver antes de reactivarlo. `Option Explicit` queda comentado por ser una instrucción del módulo. Verificado que **no queda ni una sola línea de código viva** en ninguno de los dos.

**Auditoría previa por las 4 vías** — ninguna invocación viva de las dos rutinas:

1. **Código VBA** — `Rut_Lo_Import_AE4x1` solo la llamaba `M50` (4 veces, líneas 45/51/57/63), que queda comentado a la vez; `RuT_Inf_Contable_Recibos_AE4x4` no la llamaba nadie.
2. **Macros asignadas a shapes** — ninguno de los 66 shapes con macro del `.xlsm` apunta a ellas.
3. **Tabla `Tb_Tareas` del menú auxiliar** — vector crítico aquí, porque `Form_Menu.frm:266` ejecuta las tareas con `Application.Run Rut_Name` leyendo el **nombre de la rutina de la columna 3** de la tabla: una invocación dinámica que ningún grep del código ve. Leídas las 40 tareas reales del libro (`xl/tables/table12.xml` → `sheet10.xml`): **ninguna** apunta a estas dos rutinas.
4. **Resto del XML** del libro — las dos cadenas no aparecen en ninguna celda; solo dentro de `vbaProject.bin`, que es el propio código compilado, no una invocación.

**Efecto colateral revisado:** al comentar `M50` desaparece el único `Set ActivForm` del proyecto (ver B17), con lo que `ActivForm` queda declarada en `M00` pero sin ningún uso vivo. Es inofensivo para compilar, y la declaración **se mantiene**: `M03`, `M04`, `M05` y `M90_Rut_Format_Colmns` la siguen usando en líneas comentadas, y cualquiera de ellas que se reactive la necesitará (con su propio `Set`, ya que no cuelgan de `M50`).

**Para reactivar el flujo AE4** hay que resolver antes B5/B6: declarar `Lo_AE4x1` (probablemente `Public` en `M00_Ini_Var_APP.bas`, como el resto de objetos del proyecto), añadir su `Set Lo_AE4x1 = Ws_AE4x1.ListObjects(1)` tras la copia del rango en la rama `If SW_Inicilizar_Ws`, y cambiar a `True` la primera de las 4 llamadas de `M50` para que esa rama llegue a ejecutarse.

### B7 · Límite de filas hardcodeado a 5000
**Severidad:** Alto · **Fichero:** `M21_Resumen_Tit_Propios_UNO.bas` — línea 101

```vba
    For Fila_DR = 1 To 5000 'Lo_TPH.ListRows.Count
```

El bucle está limitado a 5000 iteraciones fijas en vez de usar `Lo_TPH.ListRows.Count`, que queda comentado justo al lado — delatando la intención original. Si `Prog_BD` supera las 5000 filas, se truncan datos silenciosamente; si tiene menos, `Lo_TPH.ListRows(Fila_DR)` lanza "Subscript out of range" al superar el número real de filas.

**Arreglo:** `For Fila_DR = 1 To Lo_TPH.ListRows.Count`.

### B8 · Filtro roto por referencia sin cualificar y variable de bucle equivocada
**Severidad:** Alto · **Fichero:** `M33_List_PLANES_Anulados.bas` — líneas 74-77

```vba
    For Cont = 2 To Lo_TitPH.ListRows.Count
        If .Cells(Cont, BD_Tipo_Rec) = "Deleted" Then GoTo Reg_Siguiente
        If Cells(ContIni, BD_ImpAdm) < 0 Then GoTo Reg_Siguiente
        If .Cells(Cont, BD_Plan) <> Cod_Plan Then
```

La línea 76 debería ser `.Cells(Cont, BD_ImpAdm)` — cualificada con el `.` del `With Lo_TitPH.DataBodyRange` en curso, y usando `Cont` (la variable del bucle) — tal como hacen los módulos hermanos `M31_List_PLANES.bas:90` y `M32_List_PLANES_Devoluc.bas:67`. Al faltar el punto, `Cells(...)` referencia implícitamente la hoja **activa**, no `Lo_TitPH.DataBodyRange`; y al usar `ContIni` (constante fijada antes del bucle) en vez de `Cont`, siempre mira la misma celda de la primera fila.

**Impacto:** el filtro de "excluir Ajustes de Matrícula" de este informe de Anulados nunca actúa como se pretende.

**Arreglo:** `If .Cells(Cont, BD_ImpAdm) < 0 Then GoTo Reg_Siguiente`.

### B9 · `Application.Calculation` guardado en variable `Boolean`
**Severidad:** Alto · **Ficheros:** `M40_Inf_Contab_Recibos.bas:170,601`, `M41_Añadir_Núm_JIs_al_Inf.bas:27,228`, `M42_Añadir_Núm_JIs_a_BDatos.bas:27,91`

```vba
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
...
    Application.Calculation = Sw_Calculation
```

`Application.Calculation` devuelve una constante `XlCalculation` (p. ej. `xlCalculationAutomatic = -4105`), no un booleano. Al guardarla en `Boolean`, cualquier valor no-cero se convierte en `True`, perdiendo qué modo había realmente; al restaurar, `Application.Calculation = Sw_Calculation` escribe `True` (-1), que no es ningún miembro válido de `XlCalculation`. El mismo patrón aparece también en `Rut_WS.bas` (ver C5) y `M10__Liquid_EP.bas`.

**Arreglo:** `Dim Sw_Calculation As XlCalculation` (o `Long`).

### B10 · Año "2025" hardcodeado en nombres de fichero exportado
**Severidad:** Medio · **Ficheros:** `M38x_Export_Cierre_Contable.bas:21-22`, y de forma idéntica `M22_Inf_EPs_para_UXXI_NEW.bas:489-490`

```vba
    FichName = "NUEVO_" & TipoEP & Prog__APP.Range("APP_CursAcad") _
               & "_Cierre_2025 " & Format(Now, "(yyyy-mm-dd_hhmm)") & ".xlsx"
```

El "2025" no depende de `APP_CursAcad`/`APP_AñoCont`; en cursos posteriores el fichero exportado seguirá llamándose "..._Cierre_2025..." aunque se genere en otro año contable.

**Arreglo:** sustituir el literal `"_Cierre_2025"` por `"_Cierre_" & Prog__APP.Range("APP_AñoCont")`.

### B11 · Módulos M22 OLD/NEW con agrupación distinta sobre la misma hoja destino (a verificar)
**Severidad:** Medio, a verificar · **Ficheros:** `M22_Inf_EPs_para_UXXI.bas` (`Rut_Informe_EPs_para_UXXI_OLD`) y `M22_Inf_EPs_para_UXXI_NEW.bas` (`Rut_Informe_EPs_para_UXXI`)

No colisionan en compilación (nombres de Sub distintos), pero ambas escriben en la misma hoja destino `Sht__Inf_EPs_UXXI`, y la agrupación cambió: la versión OLD agrupa por `Plan & Curso_Acad_Ant & Año_Emi_Ant` (tres claves), la NEW agrupa solo por `Cod_Plan & Año_Emi` (dos claves — se eliminó `Curso_Acad`), y los `SumIfs` que siguen tampoco filtran ya por `BD_C_Acad`. Si el mismo código de Plan se reutiliza en dos ediciones (Curso_Acad) distintas que comparten año contable de emisión, la versión NEW fusionaría en una sola fila-resumen datos que la OLD mantenía separados. No se ha confirmado si los códigos de Plan son siempre únicos por Curso_Acad en los datos reales, ni si `_OLD` sigue invocándose desde algún botón/Ribbon no auditado en este trabajo (por grep no aparece ninguna llamada en el código exportado).

### B12 · `M71_Restituir` sobrescribe datos y parámetros sin validar ni confirmar
**Severidad:** Medio · **Fichero:** `M71_Restituir_BDatos_EP_Work.bas` — líneas 28-44, 56-65, 109

```vba
        .Filters.Add "Sólo Ficheros Excel", "*.xls?", 1
        If .Show <> -1 Then
            ...
        Else
            Arch__EP_New = .SelectedItems(1)
...
    Prog__APP.Range("APP_AñoCont") = ClsBk.Sheets(Prog__APP.Name).Range("APP_AñoCont")
    Prog__APP.Range("APP_CursAcad") = ClsBk.Sheets(Prog__APP.Name).Range("APP_CursAcad")
    ...
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_BD, True)
```

A diferencia de `M51` (que valida que el fichero elegido empiece por el nombre esperado), el selector aquí acepta cualquier `*.xls?` sin comprobar que sea realmente un backup de Enseñanzas Propias. Los parámetros de `Prog__APP` (Año Contable, Curso Académico, EFP/CFC, Plan activo...) se sobrescriben **desde el backup** antes de mostrar ningún resumen, y la tabla `Prog_BD` completa se sustituye (`True` = vaciar destino primero) sin ningún `MsgBox` de confirmación previo — el único aviso llega después, ya consumado, en el log de `Form_Menu.TB_Informe`.

**Impacto:** elegir por error el fichero equivocado cambia silenciosamente el Año Contable/Curso Académico activos de la aplicación y sustituye los datos de trabajo, sin posibilidad de cancelar tras verlo.

**Arreglo:** validar el nombre del fichero elegido (como hace M51) y pedir confirmación explícita antes de sobrescribir `Prog__APP`/`Prog_BD`.

### B13 · `Wk_Lista_Panes1.cls` sin ninguna referencia en el código (a verificar)
**Severidad:** Bajo, a verificar · **Fichero:** `Wk_Lista_Panes1.cls`

Solo `Wk_Lista_Panes` se usa (`M31_List_PLANES.bas:29`); `Wk_Lista_Panes1` (hoja vacía, sin código propio) no aparece en ningún otro `.bas`/`.cls` del proyecto. Probable hoja duplicada/huérfana, candidata a limpieza — pero esta tarea no ha auditado macros asignadas a shapes ni el XML del Ribbon, así que un grep negativo en el código no es prueba suficiente de "huérfano" (ver la nota ya recogida en el `CLAUDE.md` global sobre este mismo riesgo).

### B14 · Bucle `Do While` sin cota superior en M31/M32/M33
**Severidad:** Bajo · **Ficheros:** `M31_List_PLANES.bas:59-61`, `M32_List_PLANES_Devoluc.bas:40-42`, `M33_List_PLANES_Anulados.bas:47-49` (mismo patrón en los tres)

```vba
    Do While .Cells(ContIni, BD_Tipo_Rec) = "Deleted" Or .Cells(ContIni, BD_ImpAdm) < 0
        ContIni = ContIni + 1
    Loop
    Cont = ContIni
    Cod_Plan = .Cells(Cont, BD_Plan)
```

Si todas las filas de `Prog_BD` cumplieran la condición de salto, `ContIni` superaría `Lo_TitPH.ListRows.Count` y la siguiente lectura fallaría con un error de ejecución en vez de un mensaje claro tipo "no hay datos". Caso límite, baja probabilidad en uso real.

### B15 · Contraseña de correo en texto plano
**Severidad:** Bajo · **Fichero:** `M80_Mandar_Correo.bas` — línea 126 y `Rut_Cambiar_Contraseña_Email` (líneas 142-151)

La contraseña SMTP se lee/escribe/muestra en una celda (`Range("APP_MailClau")`) y en `MsgBox`/`InputBox` sin ningún enmascarado. No es un bug funcional, pero es una práctica de higiene de datos a revisar.

---

### B16 · Cuatro rutinas de otro libro en `M90_Rutinas_X.bas` (hallazgo posterior)
**Severidad:** Crítico · **Fichero:** `M90_Rutinas_X.bas` — líneas 137-226 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-18)

Detectado al compilar el proyecto tras corregir A3: el VBE marca `No se ha definido la variable` sobre `Hp_BuscarJI_1`. Es el mismo patrón que B3 (código importado de otro libro sin adaptar), pero en un fichero distinto y con origen distinto: **no viene de PPub**, donde estos identificadores tampoco existen, sino de un tercer libro, el de Tasas/Aplicaciones.

Siete identificadores usados y declarados en ningún sitio del proyecto:

| Identificador | Qué sería | Dónde se usaba |
|---|---|---|
| `Hp_BuscarJI_1`, `Hp_BuscarJI_2`, `Hp_T_Aplic` | CodeName de hoja | `M90_Rutinas_X.bas` |
| `Lo_Aplic` | `ListObject` | `M90_Rutinas_X.bas` |
| `CTa_Ref` | `Const` de columna | `M90_Rutinas_X.bas` |
| `Lo_Prog_Colns`, `LastCol_Tb_Solicitudes` | `ListObject` / `Const` | `M90_Rutinas_X.bas` + `M90_Rutinas_Menú_Aux.bas` |

Tres indicios de que es código ajeno a este libro: el prefijo `Hp_` no es la convención de aquí (`Prog_*`, `Wk_*`, `Sht__*`); los comentarios hablan de *Tasas*, *Aplicaciones* y *Solicitudes*, no de Liquidación de Títulos Propios; y ninguno de los identificadores aparece tampoco en PPub.

**Arreglo aplicado:** eliminadas las 4 rutinas que dependían de los cinco primeros identificadores (`Rut_Filtro_NumJI_1_en_Tasas`, `Rut_Filtro_NumJI_2_en_Tasas`, `Rut_Ajustar_V_H_Alignment`, `Rut_Filas_Ajustar_Alto`), tras auditar las 4 vías de invocación: sin llamadas en el código VBA, sin macro asignada en ninguno de los 68 shapes del `.xlsm`, sin apariciones en `sharedStrings` ni en ninguna otra parte XML (el libro no tiene Ribbon custom). Módulo de 365 a 254 líneas.

**Resuelto provisionalmente (2026-09-18):** `Lo_Prog_Colns` y `LastCol_Tb_Solicitudes` seguían sin declarar, usados en `Rut_Columnas_Ajustar_Ancho` y `Rut_Columnas_Mostrar` de este módulo y en `M90_Rutinas_Menú_Aux.bas:89-90`. El usuario **comentó esas rutinas** para desbloquear la compilación (ver el apéndice, sección "Comentadas por el usuario"). Sigue en pie el criterio de que **estas rutinas sí parecen propias de este libro** (`Rut_Columnas_Mostrar` escribe en `Form_Menu.TB_Informe`, el menú auxiliar): si se quieren recuperar, lo correcto es **declarar las variables apuntando a la tabla real** — probablemente el `ListObject` de alguna hoja `Prog_DefCol*`, cuyas filas 4, 5, 7 y 11 se usan como ancho, alineación horizontal, marca "Ocultar" y alineación vertical — en vez de dejarlas comentadas para siempre.

### B17 · `ActivForm` usada en 6 módulos y declarada en ninguno
**Severidad:** Crítico · **Ficheros:** `M00_Ini_Var_APP.bas` (la declaración que faltaba), `M50_Inf_Cont_AE4x4.bas`, `M51_Import_AE4x1.bas` · **Estado:** ✅ Corregido (2026-09-18)

Detectado al compilar tras los arreglos de C3: el VBE marca `No se ha definido la variable` sobre `ActivForm` en `M51_Import_AE4x1.bas:20`.

```vba
    Dim TxT_ProgIni     As String:      TxT_ProgIni = ActivForm.Controls("TBx_Informe")
```

`ActivForm` es el UserForm activo, sobre el que las rutinas largas van escribiendo el progreso vía `Rut_TimeLap_Inf(ActivForm, "TBx_Informe", …)`. Aparece **41 veces en 6 módulos** (`M03`, `M04`, `M05`, `M50`, `M51`, `M90_Rut_Format_Colmns`) y tiene su `Set` en `M50_Inf_Cont_AE4x4.bas:24`:

```vba
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.
```

…pero **no se declara en ningún punto del proyecto**. No es código ajeno importado (a diferencia de B3 y B16): el libro hermano **PPub sí la declara**, en el módulo equivalente a nuestro `M00_Ini_Var_APP.bas`, exactamente una línea antes de `Lo_Tareas`:

```vba
' Precios Públicos - PPub, M_000_Ini_Var_APP.bas:355
    Public ActivForm        As Object      '- Identificamos qué Formulario está Activo.  ----------
    Public Lo_Tareas        As ListObject
```

Nuestro `M00` conserva `Lo_Tareas` y todo el bloque de alrededor (`MsgBx_*`, etc.) en el mismo orden, pero le falta justo la línea de `ActivForm`: se perdió al derivar este libro de PPub.

**Arreglo aplicado:** restituida `Public ActivForm As Object` en `M00_Ini_Var_APP.bas`, en la misma posición que ocupa en PPub (justo antes de `Public Lo_Tareas`). `Public` es el ámbito correcto: el `Set` vive en `M50` y el uso en `M51`, módulos distintos.

**Verificado que no queda un Error 91 latente:** las 4 llamadas a `Rut_Lo_Import_AE4x1` (`M50:45,51,57,63`) son posteriores al `Set` de `M50:24`, así que `ActivForm` llegaba poblado a `M51`. En `M03`, `M04`, `M05` y `M90_Rut_Format_Colmns` todas las apariciones están **en líneas comentadas**, así que no hay ninguna ruta viva que la use sin asignar. Cuidado si se reactiva alguna: habría que añadirle su propio `Set` (esas rutinas no cuelgan de `M50`).

**Actualización (2026-09-18, posterior):** al desactivarse por completo `M50` y `M51` (ver B5/B6), el único `Set ActivForm` del proyecto queda comentado, así que `ActivForm` pasa a estar **declarada pero sin ningún uso vivo**. La declaración se conserva a propósito: es la correcta según PPub, no estorba a la compilación, y la necesitará cualquiera de las rutinas comentadas de `M03`/`M04`/`M05`/`M90` o el propio flujo AE4 cuando se reactiven.

### B18 · `AñoContAnt` sin declarar en `M20_Resumen_Tit_Propios.bas`
**Severidad:** Crítico · **Fichero:** `M20_Resumen_Tit_Propios.bas` — línea 183 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-19)

Detectado al compilar: `No se ha definido la variable` sobre `AñoContAnt`, dentro del bloque "Importe Aplz Aplazado" de `Rut_Resumen_Tab_TitPropios`:

```vba
            '--- Importe Aplz Aplazado --------
            If RwPH.Range(BD_ImpCob) > 0 And RwPH.Range(BD_ACont_Vto) > AñoContAnt Or _
```

A diferencia de B3/B16/B17, aquí **no hay nada ajeno ni exótico**: `AñoContAnt` es una variable normal del dominio —el primer año del curso académico, "2025" en "2025-26"— que **otros 5 módulos declaran localmente con la misma línea** (`M20_Inf_EP_UXXI`, `M22_Inf_EPs_para_UXXI`, `M22_..._NEW`, `M32_List_Anul_y_Devoluciones`, `M39_Resumen_Planes_C_Acad`):

```vba
    Dim AñoContAnt      As Integer:     AñoContAnt = Left(CursoAcad, 4)   '- El 1º Año de Curso
```

Simplemente se olvidó en este módulo, que la usa una única vez. Su fuente (`APP_CursAcad`) ya estaba leída en la rutina.

**Arreglo aplicado:** añadida la declaración junto a las demás, al principio de `Rut_Resumen_Tab_TitPropios`:

```vba
 Dim AñoContAnt            As Integer:         AñoContAnt = Left(CursAcad, 4)      '- El 1º Año del Curso Académico
```

**Un detalle que había que mirar antes de copiar y pegar:** en este módulo la variable del curso académico **no se llama `CursoAcad` sino `CursAcad`** (el nombre local, ya declarado en la línea 9). Copiar literalmente la línea de los otros módulos habría dejado un segundo "variable no definida" en su lugar. `Integer` es el tipo correcto: se compara contra `BD_ACont_Vto`, un año contable numérico, igual que en los 5 módulos hermanos.

**⚠️ Inconsistencia preexistente detectada de paso (no corregida):** en este módulo `AñoCont` se declara **`As String`**, mientras que en los otros cuatro es `As Integer`:

| Módulo | `AñoCont` |
|---|---|
| `M20_Resumen_Tit_Propios.bas` | **`String`** |
| `M20_Inf_EP_UXXI.bas`, `M22_Inf_EPs_para_UXXI.bas`, `M32_List_Anul_y_Devoluciones.bas`, `M39_Resumen_Planes_C_Acad.bas` | `Integer` |

Y aquí se usa en comparaciones y aritmética numéricas (`> AñoCont`, `AñoCont + 1`, líneas 161-247). VBA lo resuelve por coerción implícita, así que "funciona", pero un `String` en una comparación de orden puede comparar como texto en vez de como número en según qué contexto. **No se ha tocado**: no bloquea la compilación y cambiar el tipo podría alterar resultados de un informe en producción. Conviene verificarlo con datos reales antes de unificarlo.

### B19 · `Cod_Plan` sin declarar y vaciado antes de usarse (`M21_Resumen_Tit_Propios_UNO.bas`)
**Severidad:** Crítico · **Fichero:** `M21_Resumen_Tit_Propios_UNO.bas` — líneas 94-95 y 106 · **Estado:** ⏸️ Módulo desactivado (2026-09-19) — decisión de diseño pendiente

Detectado al compilar: `No se ha definido la variable` sobre `Cod_Plan`, en el filtro principal del bucle de `Rut_Resumen_Tab_TitPropios_UNO`:

```vba
    Range("TP_Cod_Plan") = ""                                    ' línea 94  ← lo VACÍA
    Range("TP_Cod_Plan").Offset(0, 1) = Range("APP_CursAcad")    ' línea 95
    ...
        If RwPH.Range(BD_Plan) <> Cod_Plan Then GoTo Siguiente_Fila   ' línea 106  ← y aquí filtra por él
```

**No basta con declarar la variable**, y por eso no se ha aplicado un arreglo mecánico: hay un problema de diseño detrás. El módulo es la variante "un solo plan" de `M20_Resumen_Tit_Propios`, y `Cod_Plan` es el plan que hay que resumir. El rastro de dónde debería salir está en el code-behind de la hoja, `Wk_TitP_UNO.cls:22`, **entero comentado**:

```vba
'        Rut_Resumen_Tab_TitPropios_UNO (Range("TP_Cod_Plan"))
```

Es decir: en el diseño original la rutina **recibía el código de plan como parámetro**, disparada por el `Worksheet_Change` de la hoja al escribir en la celda. Ese evento está desactivado, y la firma actual no tiene parámetros. Además, la línea 94 **vacía `TP_Cod_Plan`** justo antes del bucle que lo necesita, con lo que leerlo ahí devolvería siempre cadena vacía.

**Dato relevante sobre el rango** (verificado en `xl/workbook.xml`): `TP_Cod_Plan` es un nombre de **ámbito de hoja, definido tres veces** en tres hojas distintas:

| Nombre | Apunta a | Hoja (CodeName) |
|---|---|---|
| `TP_Cod_Plan` | `EP_Resumen_1!$F$1` | — |
| `TP_Cod_Plan` | `EP_UXXI!$F$2` | — |
| `TP_Cod_Plan` | `Tit_Propio_UNO!$F$1` | **`Wk_TitP_UNO`** ← el de este módulo |

Los `Range("TP_Cod_Plan")` **sin cualificar** de las líneas 94-95 resuelven por la hoja activa, exactamente el riesgo que documenta el `CLAUDE.md` del proyecto (mismo patrón que A8). Al reactivar el módulo hay que cualificar: `Wk_TitP_UNO.Range("TP_Cod_Plan")`.

**⚠️ A diferencia de `M50`/`M51`, esta rutina SÍ estaba viva.** Auditadas las 4 vías: no la llama ningún código VBA (la única referencia, `Wk_TitP_UNO.cls:22`, está comentada) ni ningún shape, pero **sí figura en la tabla `Tb_Tareas`** del menú auxiliar — fila 31, tarea *"2_ Exportar Tabla Resumen de UN Tit.Propio"* —, que `Form_Menu.frm:266` ejecuta con `Application.Run`. Mientras el módulo esté comentado, **esa entrada del menú fallará al pulsarla**: conviene quitar o marcar esa fila en `Tb_Tareas` mientras tanto.

**Decisión aplicada:** comentar el módulo entero (289 → 305 líneas, 0 de código vivo), a la espera de decidir de dónde sale el plan. Las dos salidas razonables, cuando se retome:

1. **Leer el plan del rango antes de vaciarlo** — coherente con que la tarea se lance desde el menú sin argumentos: el usuario escribe el plan en `Tit_Propio_UNO!F1` y luego lanza la tarea. Requiere `Dim Cod_Plan As String: Cod_Plan = Wk_TitP_UNO.Range("TP_Cod_Plan")` al principio, control de vacío, y mover el vaciado de la línea 94 a después de la lectura.
2. **Restaurar el diseño original** — devolver el parámetro a la firma (`Sub Rut_Resumen_Tab_TitPropios_UNO(Cod_Plan As String)`) y reactivar el `Worksheet_Change` de `Wk_TitP_UNO.cls`. Esto es **incompatible** con lanzarla desde `Tb_Tareas` sin argumentos, así que habría que quitarla del menú.

## Bloque C — Librería transversal `Rut_*`/`Prog_*` y formularios
*Ficheros: `Rut_Lo`, `Rut_WB`, `Rut_WS`, `Rut__Right_Click_VBA`, `Form_*`, `M0999_*`, `Módulo*`.*

### C1 · Variable no declarada `LoTb` (typo de `Lo_Tb`)
**Severidad:** Crítico · **Fichero:** `Rut_Lo.bas` — línea 136 (`Option Explicit` activo)

```vba
Sub Rut_Lo_Filtro(ByRef Lo_Tb As ListObject, columna As Integer, Criterio As String, Optional SW_Clear As Boolean = False)
    If SW_Clear And Not Lo_Tb.AutoFilter Is Nothing Then LoTb.AutoFilter.ShowAllData
    Lo_Tb.Range.AutoFilter Field:=columna, Criteria1:=Criterio
End Sub
```

El parámetro se llama `Lo_Tb`, pero la línea usa `LoTb` (sin guion bajo), no declarada en ningún sitio del proyecto (confirmado con grep global). Con `Option Explicit`, "Variable no definida" en compilación. Esta `Sub` en concreto no la llama nadie (existe un duplicado correcto y sí usado, `Rut_x_Filtro_LoTb` en `M90_Rutinas_X.bas`, consistente con `LoTb`), pero el error de compilación afecta igualmente a todo el proyecto.

**Arreglo:** cambiar `LoTb` por `Lo_Tb` (o borrar la Sub, ya que está duplicada y sin uso).

### C2 · `.calcMode` no es un miembro de `Application`
**Severidad:** Crítico · **Fichero:** `Rut_Lo_Export_XlsX.bas` — línea 44 (`Rut_Lo_Export_to_New_WB`)

```vba
    Dim calcMode    As XlCalculation
...
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .calcMode = .Calculation
        .Calculation = xlCalculationManual
    End With
...
    With Application
        .ScreenUpdating = True
        .Calculation = calcMode
    End With
```

`calcMode` se declara como variable local (línea 20) para guardar el modo de cálculo y restaurarlo después (eso sí, correcto al final). Pero dentro del `With Application`, `.calcMode = .Calculation` lleva el punto delante de `calcMode`, así que VBA lo interpreta como `Application.calcMode` — propiedad que no existe → "Método o miembro de datos no encontrado" en compilación. El patrón correcto (`calcMode = .Calculation`, sin punto inicial) sí se usa de forma consistente 4 veces en `Rut__Right_Click_VBA.bas`.

**Arreglo:** quitar el punto inicial → `calcMode = .Calculation`.

### C3 · `String` pasado donde se espera `Worksheet` por referencia (ruta en producción)
**Severidad:** Crítico · **Fichero:** `Rut_WS.bas` — líneas 42-49, 58-70 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-18)

```vba
Sub Rut_WrkSheet_Preparar(WrkSht As Worksheet)  '- Mostrar todas las Filas y Columnas, y Quitar Filtros.
    With WrkSht
        .Columns.EntireColumn.Hidden = False
        .Rows.EntireRow.Hidden = False
        Call Rut_Lo_Filtros_Quitar(.ListObjects(1))
    End With
End Sub
...
Sub Rut_WrkSheet_Vaciar(ByVal WrkSht As String)      '--- Borra Toda la Hoja incluso los objetos (Shapes) ---
    ...
    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
        ...
        .Unprotect
        Call Rut_WrkSheet_Preparar(WrkSht)
```

`Rut_WrkSheet_Preparar` exige un `Worksheet` por parámetro (ByRef implícito, línea 42); `Rut_WrkSheet_Vaciar` le pasa su propio parámetro `WrkSht`, que es `String` (línea 58) → "Error de compilación: Tipo de argumento ByRef incompatible". A diferencia de otros hallazgos de este bloque, `Rut_WrkSheet_Vaciar` **está activamente en uso**: la llaman `M09_Importar_Sol_Liq.bas:76`, `M12_Genera_LIQx_PDF.bas:24,303`, `M51_Import_AE4x1.bas:90` y `Rut_Hipervinculos.bas:13` (nota 2026-09-15: `M51_Import_AE4x11.bas:92` ya no existe, ver B2).

**Arreglo aplicado (2026-09-18):** en vez del parche mínimo que se proponía aquí (envolver el argumento en `Application.Workbooks(...).Sheets(WrkSht)`), se ha cambiado **la firma** de la rutina por coherencia con `Rut_WrkSheet_Preparar` y con el resto de la librería `Rut_*`, que ya trabaja con objetos:

```vba
Sub Rut_WrkSheet_Vaciar(WrkSht As Worksheet)      '--- Borra Toda la Hoja incluso los objetos (Shapes) ---
    ...
    With WrkSht                                   '- antes: Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
        ...
        Call Rut_WrkSheet_Preparar(WrkSht)        '- ahora ya recibe el Worksheet que espera
```

Así desaparece además la resolución por nombre contra `ThisWorkbook`, que era una limitación implícita (la hoja tenía que vivir sí o sí en este libro) y una fuente de error en ejecución si el nombre no existía.

**Llamadas adaptadas (5):**
- `M09_Importar_Sol_Liq.bas:77`, `M12_Genera_LIQx_PDF.bas:25` y `:304`, `M51_Import_AE4x1.bas:90` — pasaban `<Hoja>.Name`; ahora pasan el propio objeto de hoja (`Prog_Sol_Liq`, `Wk_TitP_LIQx_PDF`, `Ws_AE4x1`).
- `Rut_Hipervinculos.bas:13` — era el único caso que pasaba un **literal** (`"Hipervínculos"`). Se ha reordenado el bloque para obtener antes la referencia y pasarla:

```vba
    If Fnc_WrkSheet_Exist("Hipervínculos") Then
        Set hojaResultado = ThisWorkbook.Sheets("Hipervínculos")
        Call Rut_WrkSheet_Vaciar(hojaResultado)
```

`WrkSht_Activa` (el nombre de la hoja activa que se guarda para restaurarla al final) sigue siendo `String` a propósito: ahí sí se quiere el nombre.

### C4 · Llamada a una rutina que no existe: `Rut_Actualizar_1_LS_VAL`
**Severidad:** Alto · **Fichero:** `Mensaje.frm` — línea 20 (`UserForm_Activate`) · **Estado:** ✅ Corregido (2026-09-19) — UserForm eliminado

```vba
Sub UserForm_Activate()
Rut_Actualizar_1_LS_VAL
Unload Me
End Sub
```

`Rut_Actualizar_1_LS_VAL` no existe en ningún módulo del proyecto (comprobado con grep global, no solo en el bloque revisado). Al ser una llamada directa (no `Application.Run` con cadena), VBA la resuelve en compilación: "Sub o función no definida". El propio formulario `Mensaje` tampoco lo muestra nadie (ni `Mensaje.Show` ni `Load Mensaje` aparecen en ningún módulo) — código huérfano por ambos lados.

**Arreglo aplicado (2026-09-19):** eliminados `Mensaje.frm` y `Mensaje.frx` del repo. El UserForm entero eran 23 líneas, de las cuales 4 de código, y su única acción al abrirse era llamar a la rutina inexistente y descargarse a sí mismo — no mostraba nada ni tenía lógica propia.

**Auditado por las 4 vías antes de borrarlo**, todas negativas: ningún `Mensaje.Show`/`Load Mensaje` en el código, ningún shape con macro que lo invoque, ninguna entrada en `Tb_Tareas` ni en ninguna celda del libro. Se comprobó además que `Rut_Actualizar_1_LS_VAL` **tampoco existe en PPub**, así que no es un caso de código importado a medias (a diferencia de B3, B16 o B17, donde el libro hermano sí tenía el original): es una llamada a algo que no ha existido nunca en ninguno de los dos libros.

Verificado tras el borrado que no queda ninguna referencia colgando. Ojo con un falso positivo al buscar: `Form_MsgBox.frm` contiene cuatro apariciones de `Lb_Mensaje`, que es un **control** de ese formulario, sin ninguna relación con el UserForm eliminado. Quedan 3 formularios: `Form_Menu`, `Form_MsgBox` y `Form_Usuario`.

**⚠️ Pendiente en el `.xlsm`:** el componente sigue dentro del libro. Hay que quitarlo a mano en el editor VBA — clic derecho sobre `Mensaje` en el árbol del proyecto → *Quitar Mensaje…* → *No* cuando pregunte si exportar. Reimportar los módulos no lo elimina solo.

### C5 · `Rut_WrkSheet_ReducirPeso` opera sobre la hoja activa, no sobre la recibida
**Severidad:** Alto · **Fichero:** `Rut_WS.bas` — líneas 20-38

```vba
Sub Rut_WrkSheet_ReducirPeso(ByVal WrkSht As String, Optional Sw_Del_DataBodyRange As Boolean = False)
        Dim ws      As Worksheet:       Set ws = Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
        Dim Lo      As ListObject:      Set Lo = ws.ListObjects(1)
        Dim Sw_Calculation  As Boolean:  Sw_Calculation = Application.Calculation:   Application.Calculation = xlManual
    With ws
        ...
        With Lo.Range
            Range(.Cells(.Rows.Count, .Columns.Count).Address).Select   ' Selecciona la última celda de la tabla
            ActiveCell.Offset(1, 1).Select
        End With
        .Range(ActiveCell.Address & ":" & Cells(Rows.Count, 1).Address).EntireRow.Delete
        .Range(ActiveCell.Address & ":" & Cells(1, Columns.Count).Address).EntireColumn.Delete
```

`Range(...)`, `ActiveCell` y `Cells(...)` van sin cualificar: aunque están dentro de `With ws`/`With Lo.Range`, al no llevar el punto delante actúan sobre la hoja **activa** en ese momento, no sobre `ws` (la hoja recibida por parámetro). Además, el mismo defecto de C9/B9 se repite aquí: `Sw_Calculation` guarda `Application.Calculation` (un `XlCalculation`) en una variable `Boolean`.

**Impacto:** si en el momento de la llamada la hoja activa es distinta de `WrkSht`, el `.Select`/`EntireRow.Delete`/`EntireColumn.Delete` borraría filas/columnas de la hoja equivocada.

**Arreglo:** cualificar todo con `ws.` (`ws.Range(...)`, sustituyendo `ActiveCell` por una variable `Range` propia calculada dentro del `With Lo.Range`); y declarar `Sw_Calculation As XlCalculation`.

### C6 · Pérdida de datos en `M0999_Modif_Cols_BDatos.bas` si se reejecuta
**Severidad:** Alto · **Fichero:** `M0999_Modif_Cols_BDatos.bas` — líneas 97-102 (`Reorganizar_ListObject_Solo`)

```vba
    Lo.ListColumns("Col_55").DataBodyRange.Copy
    Lo.ListColumns("Col_56").DataBodyRange.Copy
    Lo.ListColumns("Col_57").DataBodyRange.Copy
    Lo.ListColumns("Col_55").Delete: Lo.ListColumns("Col_56").Delete: Lo.ListColumns("Col_57").Delete
    Lo.ListColumns.Add.Name = "Col_55": Lo.ListColumns.Add.Name = "Col_56": Lo.ListColumns.Add.Name = "Col_57"
    Lo.ListColumns("Col_55").DataBodyRange.PasteSpecial xlPasteValues
```

Cada `.Copy` sobrescribe el portapapeles del anterior: al llegar al `PasteSpecial` solo sobrevive la copia de `Col_57`, y ni siquiera se pega en `Col_57` sino en `Col_55`. Se borran las 3 columnas y solo se "recupera" un pegado, y el equivocado. Es una migración "de un solo uso" (nombrada por fecha, como sus hermanas `Reorganizar_Lo_BDatos_Cambio_Ene_26`/`_Nov_2025`), que además usa `ActiveWorkbook`/`Worksheets(...)` sin cualificar por libro.

**Impacto:** si alguien vuelve a lanzar este módulo por error (Alt+F8) sobre `Prog_BD` real, puede corromper o perder columnas de la base de datos.

**Sugerencia:** cualificar con `ThisWorkbook.Worksheets(...)`, y pegar cada columna justo después de copiarla, no encadenar 3 `.Copy` seguidos; o retirar el módulo del libro de producción si el cambio ya se aplicó.

### C7 · Valor mágico `-0.86` escrito sobre datos reales sin confirmación
**Severidad:** Alto · **Fichero:** `M0999_Busca_Planes_ImpAdmERR.bas` — líneas 56-57

```vba
                    RowData.Range(BD_Rec_Imp_INSS) = RowData.Range(BD_Rec_Imp_Adm)
                    RowData.Range(BD_Rec_Imp_Adm) = -0.86
```

Sin `MsgBox` de confirmación previa y sin comprobar si la fila ya fue "arreglada" antes: no es idempotente — si se ejecuta dos veces, la segunda sobrescribiría de nuevo con `-0.86` filas que ya tuvieran ese valor por otro motivo. Opera directamente sobre `Prog_BD.ListObjects(1)`, la tabla de producción.

**Sugerencia:** añadir una comprobación de "ya aplicado" y una confirmación explícita antes de escribir, o mover el módulo fuera del libro de producción si ya cumplió su propósito puntual.

### C8 · `For Each` que ignora la variable de iteración
**Severidad:** Medio · **Fichero:** `Rut_WS.bas` — líneas 72-75 (dentro de `Rut_WrkSheet_Vaciar`)

```vba
        Dim tbl As ListObject
        For Each tbl In .ListObjects
            If .ListObjects(1).ShowAutoFilter Then .ListObjects(1).AutoFilter.ShowAllData   '- Quitar filtro Tabla
        Next tbl
```

El bucle recorre `tbl`, pero el cuerpo siempre opera sobre `.ListObjects(1)` (la primera tabla), nunca sobre `tbl`. En una hoja con una sola tabla no se nota; en una con varias, solo se le quita el filtro a la primera, tantas veces como tablas haya.

**Arreglo:** sustituir `.ListObjects(1)` por `tbl` dentro del bucle.

### C9 · `Módulo3.bas` sin `Option Explicit`, con variables casi homónimas
**Severidad:** Medio · **Fichero:** `Módulo3.bas` (sin línea `Option Explicit` tras el `Attribute VB_Name`)

El fichero declara a nivel de módulo unas 80 variables acumuladoras del resumen contable de planes (`Imp_Emis`, `Imp__ADx`, `TImpADxAdm`...) y redeclara localmente variantes muy parecidas dentro de `RuT_Estadística_Contable_Planes_CAcad_Pos` (p. ej. `Imp_ADx` frente a `Imp__ADx`, con y sin doble guion bajo, conviviendo en el mismo fichero). Sin `Option Explicit`, un typo futuro en cualquiera de esos nombres tan parecidos no daría error de compilación: crearía silenciosamente una `Variant` nueva (siempre 0/vacía), y el total contable saldría mal sin aviso. Es el único módulo de "prueba" revisado sin `Option Explicit` — todos sus hermanos (`M0999_*`, `Módulo_Filtro_Avanzado_Prueba`, `Módulo1/4/5`) sí lo tienen.

**Arreglo:** añadir `Option Explicit` y compilar para detectar cualquier variable ya mal escrita.

### C10 · Asimetría Private/Public en rutinas invocadas por nombre (a verificar)
**Severidad:** Medio, a verificar · **Fichero:** `Rut__Right_Click_VBA.bas` línea 202 vs línea 283, invocadas desde `M00_Ini_APP.bas` líneas 85-86

```vba
Private Sub DelMenúRightClickCell()      ' Rut__Right_Click_VBA.bas:202
...
Sub DelMenúRightClickList()              ' Rut__Right_Click_VBA.bas:283  (pública)
```

```vba
    Run ("DelMenúRightClickCell") '- Elimina otros posible Menús XML     ' M00_Ini_APP.bas:85
    Run ("DelMenúRightClickList") '- Elimina otros posible Menús XML     ' M00_Ini_APP.bas:86
```

`Rut_ConfigExcel_RESTABLECER` (que cuelga de `ThisWorkbook.Workbook_BeforeClose`, es decir, se ejecuta cada vez que se cierra el libro) llama por nombre, vía `Run`, a una Sub `Private` declarada en **otro módulo** — patrón conocido como problemático en VBA (`Run`/`Application.Run` sobre un procedimiento `Private` de otro módulo puede fallar con error 1004). Tampoco hay ningún `On Error` activo todavía en esas dos líneas concretas. No se puede confirmar sin ejecutar el VBA real, así que queda como sospecha razonable, no como bug confirmado — pero la asimetría entre las dos Subs gemelas (una `Private`, la otra pública) ya llama la atención por sí sola.

**Sugerencia:** hacer `DelMenúRightClickCell` pública (como su gemela) para eliminar la duda, y probar cerrando el libro para confirmar que no salta ningún error.

### C11 · Ruta de disco hardcodeada como fallback silencioso
**Severidad:** Bajo · **Fichero:** `Rut_Wb_CopSegTimed_USB_HD.bas` — líneas 192-194

```vba
        If Not Fnc_Range_Exist("APP_CopSeg_Usb_Path") Then
            MsgBox "¡¡¡ Falta crear el Range('APP_CopSeg_Usb_Path') !!!", vbExclamation, "Procedimiento: Copia de Seguridad"
            FichPath = "F:\__CopSeg Versiones Programas\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
```

Solo se usa si falta el rango con nombre `APP_CopSeg_Usb_Path` (si existe, se usa ese en su lugar), pero el fallback da por hecho que en la máquina donde se ejecute existe `F:\__CopSeg Versiones Programas\`. Dado que este libro se usa entre varios ordenadores, si el rango de configuración se pierde en una copia del libro, el fallback podría apuntar a una ruta que no existe en esa máquina — al menos hay un `MsgBox` previo que avisa de la falta del Range.

**Sugerencia:** si el rango no existe, pedir la ruta con `Application.GetSaveAsFilename`/`FileDialog` en vez de asumir una ruta fija.

### C12 · `.UsedRange` como instrucción suelta — propiedad usada como si fuera un método
**Severidad:** Crítico · **Fichero:** `Rut_WS.bas` — líneas 30 y 64 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-18)

Detectado al compilar tras los arreglos de C3/B17: `El uso de la propiedad no es válido`, señalando `.UsedRange` en `Rut_WrkSheet_Vaciar`.

```vba
    With WrkSht
        ...
            .Columns.Delete     ' --- con esto se borran hasta los "Shapes"
            .UsedRange          ' ← línea suelta: NO es una instrucción válida
```

La intención es legítima y el truco es conocido: tras borrar filas y columnas, **leer `UsedRange` fuerza a Excel a recalcular el rango usado** de la hoja, para que la barra de desplazamiento no siga creyendo que la hoja es enorme. El problema es que `UsedRange` es una **propiedad que devuelve un `Range`**, no un método: una línea que solo nombra la propiedad, sin leer su valor ni asignar nada, no es una instrucción válida en VBA.

Aparecía **dos veces en el mismo módulo**, y llamativamente solo una rompía la compilación:

| Línea | Código | ¿Compila? |
|---|---|---|
| 30, en `Rut_WrkSheet_ReducirPeso` | `ActiveSheet.UsedRange` | Sí — VBA lo admite como expresión con el objeto cualificado |
| 64, en `Rut_WrkSheet_Vaciar` | `.UsedRange` (colgando del `With`) | **No** — "El uso de la propiedad no es válido" |

Esa asimetría es la que hizo que el fallo pasara desapercibido: la variante de la línea 30 lleva años en el módulo sin dar guerra, aunque **tampoco surtía efecto** (nombrar la propiedad sin leerla no fuerza nada).

**Arreglo aplicado:** en ambos casos se lee la propiedad a una variable descartable, que es el idiom habitual y deja constancia de la intención:

```vba
    Dim Dummy_UsedRange As String   '- Solo para forzar la lectura de UsedRange (ver mas abajo)
    ...
    Dummy_UsedRange = .UsedRange.Address   ' Para restablecer el rango de celdas en uso (hay que LEER la propiedad para que surta efecto)
```

Así se corrige de paso la línea 30, que compilaba pero no hacía nada. Barrido del resto del proyecto en busca del mismo patrón (líneas que empiezan por `.` con una propiedad conocida y sin `=` ni paréntesis): **no hay más casos**. Queda la gemela comentada en `M90_Rutinas_X.bas:25`, que si alguna vez se reactiva arrastrará el mismo error.

### C13 · `Dictionary` ambiguo entre Scripting Runtime y Word Object Library
**Severidad:** Crítico · **Fichero:** `Rut_Lo_Col_Format_Date.bas` — línea 312 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-19)

```vba
Private Sub MostrarReporteEstadisticas( _
    ByRef estadisticas As Dictionary, _
    ByVal nombreTabla As String, _
    ByVal nombreColumna As String)
    ...
    For Each clave In estadisticas.Keys
```

Detectado al compilar: "No se encontró el método o el dato miembro" sobre `.Keys`. El proyecto tiene referenciadas a la vez **Microsoft Scripting Runtime** y **Microsoft Word 16.0 Object Library**, y ambas exponen un tipo llamado `Dictionary` (colección clave/valor en Scripting Runtime; diccionario de corrección ortográfica en Word). Al declarar el parámetro `As Dictionary` sin cualificar, el compilador lo resolvió contra `Word.Dictionary`, que no tiene `.Keys`. La variable que se le pasa (`estadisticas`, creada con `CreateObject("Scripting.Dictionary")` en la rutina llamante) está declarada `As Object` ahí, por eso no daba el mismo error en ese punto — el fallo solo aparecía en el parámetro tipado.

**Impacto:** `MostrarReporteEstadisticas` no compilaba, bloqueando el proyecto entero (como todo error de compilación VBA).

**Arreglo aplicado:** cualificado el tipo con su librería: `ByRef estadisticas As Scripting.Dictionary`.

**Nota general:** cualquier otro `As Dictionary` sin cualificar en este proyecto corre el mismo riesgo mientras ambas referencias sigan activas; no se ha encontrado ninguna otra ocurrencia en el barrido de esta sesión.

### C14 · `ListColumns.Add` con un argumento `Name:=` que no existe
**Severidad:** Crítico · **Fichero:** `M0999_Modif_Cols_BDatos.bas` — líneas 67, 68, 85, 86, 87, 93 (antes del arreglo) · **Estado:** ✅ Corregido (2026-09-19)

```vba
Lo.ListColumns.Add Position:=iCol33, Name:="Nueva_Col1"
```

Detectado al compilar: "No se encontró el argumento con nombre", señalando `Name:=`. `ListColumns.Add` solo admite el argumento con nombre `Position`; no existe ningún argumento `Name` en su firma. `Add` devuelve el objeto `ListColumn` recién creado, así que el nombre hay que asignarlo en una instrucción aparte sobre ese objeto.

**Impacto:** `Reorganizar_ListObject_Solo` (rutina de migración puntual, ver también C6 en el mismo módulo) no compilaba.

**Arreglo aplicado:** las 6 ocurrencias pasan de `Add Position:=X, Name:="Y"` a `Add(Position:=X).Name = "Y"` (paréntesis obligatorios para encadenar `.Name` sobre el valor devuelto por la función en la misma línea).

---

## Apéndice · Limpieza de código muerto en los módulos `Rut_*` (2026-09-18)

Auditoría independiente de los bugs: qué rutinas de los 15 módulos `Rut_*` no las llama nadie. **75 rutinas analizadas, 29 sin ninguna invocación.**

### Método

Un grep negativo sobre el código exportado **no es prueba suficiente** en un libro VBA: el punto de invocación puede vivir fuera del texto fuente. Se cruzaron las cuatro vías:

1. **Código VBA** — los 131 módulos exportados, descartando líneas comentadas, y capturando también las llamadas indirectas por `Run("...")` y `.OnAction` (cadenas de texto que un grep de nombres no ve).
2. **Macros asignadas a shapes** — `xl/drawings/*.xml` del `.xlsm` leído como zip: **68 shapes con macro**, 22 nombres distintos.
3. **`sharedStrings.xml`** — por si la rutina se invoca desde una tabla de configuración o el menú dinámico.
4. **Resto de partes XML** del `.xlsm`. El libro no tiene Ribbon custom, así que esa vía no aplica.

Un detalle que casi provoca un borrado erróneo: hay nombres que son **prefijo de otros vivos** (`Rut_Lo_Filtro` vs `Rut_Lo_Filtros_Quitar`, con 69 usos). Los anclajes deben exigir el nombre exacto seguido de `(` o fin de palabra.

### Eliminadas (9)

| Módulo | Rutina | Motivo |
|---|---|---|
| `Rut_Lo.bas` | `Rut_Lo_DataBodyRange_Filtered_Copy_OLD____` | Existe la versión viva (9 usos) |
| `Rut_Lo.bas` | `Rut_Lo_Filtro` | Huérfana **y rota**: usaba `LoTb` en vez de `Lo_Tb` |
| `Rut_Lo.bas` | `Rut_Copiar_EntireRow_LstObjct` | Huérfana **y rota**: tablas `Tab_INI`/`Tab_FIN` inexistentes |
| `Rut_Lo_Export_XlsX.bas` | `Rut_Lo_Export_KKKK…K` | Nombre aporreado |
| `Rut_Ranges.bas` | `Ejemplo_Selección_Múltiple` | "Ejemplo" |
| `Rut_Wb_CopSegTimed_USB_HD.bas` | `EJEMPLO_ActualizarTabla` | "EJEMPLO" |
| `Rut_WS.bas` | `Rut_WrkSheet_ReducirPeso_xx` | Existe la viva (3 usos) |
| `Rut_WS.bas` | `Rut_WrkSheet_Vaciar_xx` | Existe la viva (6 usos) |
| `Rut_WS.bas` | `Rutxxxx_Exportar_La_Liquidación` | Prefijo `Rutxxxx` |

**−225 líneas.** Verificado tras el borrado: `Sub`/`End Sub` equilibrado en los 5 módulos, CP1252 sin `U+FFFD` ni BOM, CRLF íntegro, y **ninguna rutina viva pasó a huérfana** (prueba de que no se borró nada que fuera llamado).

### Conservadas a propósito

**Herramientas de diagnóstico manual** — huérfanas por diseño, se lanzan con F5 cuando hacen falta: `ListarHipervinculos`, `Rut_Ranges_List_ALL`, `RevisarFormulasEnHojas`, `Rut_Wb_Shapes_Statistics_List`, `RuT_AllSheets_Crear_Lista`, `RuT_AllSheets_Visible_OrNot`, `Rut_WrkBook_CopSegTimed_List_Organize`, `Rut_WrkBook_CopSegTimed_List_Selected_Del`.

**⚠️ `Rut_VBA_Export_Moduls`** — aparece como huérfana, pero es la macro de auto-exportación documentada en el `CLAUDE.md` del proyecto. **No borrar:** al estar el proyecto VBA protegido con contraseña, es la única vía de exportar los módulos.

**`NewMenúRightClickCell` / `NewMenúRightClickList`** — construyen el menú contextual personalizado. Sus gemelas `DelMenú*` **sí** se llaman (`M00_Ini_APP.bas:85-86`), pero las `New*` solo se invocan desde líneas comentadas: el menú custom puede estar a medio desactivar. Requiere decisión, no limpieza automática.

**`RuT_Antes_de_Cerrar_WorkBook`** — versión alternativa de `Workbook_BeforeClose`; el evento real (`ThisWorkbook.cls:26`) hace otra cosa. Antes de borrarla conviene decidir si su contenido debería estar en el evento.

### Comentadas por el usuario para poder compilar (2026-09-18)

El usuario comentó 12 rutinas que impedían compilar el proyecto y reexportó los módulos. Todas fallaban por **identificadores que no existen en este libro** — `Lo_Prog_Colns`, `Solicitudes`, `LastCol_Tb_Solicitudes` —, residuos del libro hermano **PPub** sin adaptar, la misma familia que B3 y B16:

| Módulo | Rutinas comentadas |
|---|---|
| `M90_Rutinas_Menú_Aux.bas` | `Rut_Mostrar_Col_Ocultas` |
| `M90_Rutinas_X.bas` | `Rut_Columnas_Ajustar_Ancho`, `Rut_Columnas_Mostrar`, `Rut_Visible_Hidde_Tablas_Prog` |
| `Rut_Ranges.bas` | `Rut_Ranges_List_ALL`, `RuT_UsedRange_Save_New_WorkBook_Liq_TPV`, `RuT_UsedRange2_Save_New_WorkBook_Liq_TPV`, `RuT_Range_Save_New_WorkBook_Liq_TPV`, `RuT_Range1_Save_Liquidación_New_WorkBook` |
| `Rut_Wb_CopSegTimed_USB_HD.bas` | `Rut_WrkBook_CopSegTimed_List_Organize`, `Rut_WrkBook_Folder_List_File`, `Rut_WrkBook_CopSegTimed_List_Selected_Del` |

**Verificado por las 4 vías** (código vivo, macros de shapes en `xl/drawings/*.xml`, `sharedStrings.xml`, sin Ribbon custom): **ninguna de las 12 tiene invocación viva**, así que comentarlas no rompe nada en ejecución. Único falso positivo del grep: `Rut_Columnas_Mostrar` aparece dentro de `Rut_Columnas_Mostrar_WrkSht`, que es una rutina **distinta y viva**.

Ojo a la contradicción con la sección anterior: tres de ellas (`Rut_Ranges_List_ALL`, `Rut_WrkBook_CopSegTimed_List_Organize`, `Rut_WrkBook_CopSegTimed_List_Selected_Del`) figuraban como *"conservadas a propósito — herramientas de diagnóstico manual"*. Estaban **rotas de todos modos** (no compilaban), así que conservarlas no aportaba nada real. Si alguna se quiere recuperar, hay que adaptarla primero a los objetos de este libro; si no, procede borrarlas del todo en vez de dejarlas comentadas.

### Pendiente

- **Decidir sobre las 12 rutinas comentadas arriba**: borrarlas definitivamente o adaptarlas a los objetos de este libro.
- **Grupo 2 — 8 variantes `ByHand`/`_01`** de rutinas vivas, sin decidir: `Rut_Lo_Export_to_New_WB_ByHand`, `Rut_WrkSheet_To_PDF_ByHand`, `RuT_Sort_Sheets_ByHand`, `Rut_Ws_All_Stratistics_01` y las 4 variantes casi idénticas de `RuT_*Save_New_WorkBook_Liq_TPV` en `Rut_Ranges.bas` (estas 4 últimas son justamente parte de las comentadas ahora).
- **Los módulos `M*` no se han auditado** — previsiblemente tienen más código muerto (`M90_CopSeg_USB_HD.bas` está entero comentado).

---

*Auditoría realizada mediante lectura completa de los 133 módulos exportados (no solo búsqueda de patrones), con verificación cruzada directa contra el fichero real de los 6 hallazgos más críticos (B1, B2, B4, C1, C3, A3) antes de incluirlos en este informe. Los hallazgos marcados "a verificar" señalan sospechas razonables que no se han podido confirmar por completo con el código en texto plano — antes de invertir tiempo en corregirlos, confírmalos contra el `.xlsm` real.*
