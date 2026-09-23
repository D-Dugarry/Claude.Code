Attribute VB_Name = "M03_Asig_nCta_CCC_Ing"
' Last Rev. 2026-09-23 18:56
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M03_Asig_nCta_CCC_Ing - Asignar la cuenta bancaria de ingreso de cada recibo
' =================================================================================================
'
' PROPOSITO
'  Rellena la columna Cta_Ingreso: en que cuenta de la UA entro realmente el
'  dinero de cada recibo. Parte de la cuenta de pago que trae LSGES04 y la
'  corrige por casos especiales (pagos FLY WIRE, regularizaciones manuales
'  anotadas en BD_InfRegulariz, recibos negativos o no cobrados).
'
' INDICE DE RUTINAS Y FUNCIONES
'  RuT_Determinar_Cta_Ingreso_ByHand ... Lanzadera manual (cuerpo comentado).
'  RuT_Determinar_Cta_Ingreso(Lo_Data, Col_Ref, Col_CtaPag, Col_CtaIng)
'                                        Rutina principal.
'
' TRAMOS DE PROGRAMACION
'    0. Limpia Col_CtaIng y copia sobre ella Col_CtaPag (valor de partida).
'
'    Despues, una pasada por cada caso especial. Todas siguen el mismo patron:
'    quitar filtros -> (ordenar) -> filtrar -> contar sobre Col_Ref -> escribir
'    el valor en las celdas visibles de Col_CtaIng -> anotar en el informe.
'
'    1. BD_ImpRec < 0                  -> 'Imp_Rec <0'  (no es un ingreso).
'    2. BD_ACont_Cob vacio             -> 'No Cobrado'.
'    3. Col_CtaPag = 'FLY WIRE*'       -> 0049 6659 07 2416175503.
'    4. BD_InfRegulariz = 'FLY*'       -> 0049 6659 07 2416175503.
'    5. BD_InfRegulariz = '0049 '      -> 0049 6659 07 2416175503 (G.Acad).
'    6. BD_InfRegulariz = '6659072416125620*' -> 0049 6659 07 2416125620.
'    7. BD_InfRegulariz = '*(0049)'    -> 0049 6659 07 2416125620.
'    8. BD_InfRegulariz = '*(2100)'    -> 2100 8984 16 0200003529.
'    9. BD_InfRegulariz = '*(0081)'    -> 0081 3191 42 0001068211.
'   10. BD_InfRegulariz = '*(0014)'    -> 9000 0005 00 0260000014 (Bco. Espana).
'
'    11. Resumen: cuantos quedan SIN cuenta, cuantos venian ya con cuenta
'        asignada por el sistema y el total de la tabla.
'
'  El orden importa: las reglas posteriores SOBRESCRIBEN a las anteriores,
'  asi que las regularizaciones manuales (BD_InfRegulariz) mandan sobre lo que
'  trajera LSGES04.
'
' NOTAS
'  El informe solo se vuelca a Form_Menu si el switch Sw_Boss esta activo:
'  es detalle tecnico pensado para el desarrollador, no para el usuario.
'
'  Las cuentas bancarias estan escritas a pelo en el codigo (no en una tabla
'  de configuracion): un cambio de cuenta de la UA se corrige aqui.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2026-01-03
Option Explicit

            Sub RuT_Determinar_Cta_Ingreso_ByHand()
'                Prog_LsGes04.Unprotect
'                Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)          '- Quita filtros, filas y columnas ocultas
'                Call RuT_Determinar_Cta_Ingreso(Prog_LsGes04.ListObjects(1), BD_FVto, BD_ACont_Vto)
            End Sub
'- -------------------------------------------------------------------------------------------------
'- Rellenar Col Cta_Ingreso con nº Cta. correspondiente --------------------------------------------
'- -------------------------------------------------------------------------------------------------
Sub RuT_Determinar_Cta_Ingreso(Lo_Data As ListObject, _
                               Col_Ref As Integer, _
                               Col_CtaPag As Integer, _
                               Col_CtaIng As Integer)

Debug.Print ">>> RuT_Determinar_Cta_Ingreso"
    Dim rowfind     As Variant
    Dim Sw_Boss     As Boolean:     Sw_Boss = Prog__APP_Switch.Range("Sw_Boss")
    Dim TxtProgreso As String:      TxtProgreso = Form_Menu.TB_Informe
    Lo_Data.ShowTotals = False
        
    '- ---------------------------------------------------------------------------------------------
    '- Determinar Cta-CCC Ingreso de cada Rec. -----------------------------------------------------
    '- ---------------------------------------------------------------------------------------------
    With Lo_Data
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .DataBodyRange.Columns(Col_CtaIng).ClearContents   '- Se supone que está vacía...
        
        '-Copy Col Col_CtaPag en Col Col_CtaIng ----------------------------------------------------
        .ListColumns(Col_CtaPag).DataBodyRange.Copy
        .ListColumns(Col_CtaIng).DataBodyRange.PasteSpecial Paste:=xlPasteValues

        '-Filtra Recibos "Imp_Rec <0"  -------------------------------------------------------------
        .AutoFilter.ShowAllData            ' Elimina los filtros
        Call Rut_Lo_Sort(Lo_Data, G04_ImpRec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=G04_ImpRec, Criteria1:="<0"
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "Imp_Rec <0"
            '- Visualizo el progreso ---------------------------------------------------------------
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "Imp_Rec < 0"
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "Imp_Rec <0", 0)
        End If

        '-Filtra Recibos "No Cobrado"  -------------------------------------------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        Call Rut_Lo_Sort(Lo_Data, G04_ACont_Cob, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=G04_ACont_Cob, Criteria1:="="
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "No Cobrado"
            '- Visualizo el progreso ---------------------------------------------------------------
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "No Cobrado"
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "No Cobrado", 0)
        End If
       
        '-Filtra Recibos Col_CtaPag = "FLY WIRE    "  ----------------------------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        Call Rut_Lo_Sort(Lo_Data, Col_CtaPag, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=Col_CtaPag, Criteria1:="=FLY WIRE*"
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416175503"
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "FLY WIRE"
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "FLY WIRE", 0)
        End If
       
        '-Filtra Recibos Col_CtaPag Inf-Regularizado= "FLY"  ---------------------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        Call Rut_Lo_Sort(Lo_Data, G04_InfRegulariz, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=G04_InfRegulariz, Criteria1:="=FLY*"
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416175503"
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "FLY Regularizado"
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "FLY Regularizado", 0)
        End If
       
        '-Filtra Recibos Col_CtaPag Inf-Regularizado= "0049 "  -------------------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        .Range.AutoFilter Field:=G04_InfRegulariz, Criteria1:="=0049 "
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416175503"
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "Regularizado G.Acad"
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "Regularizado G.Acad", 0)
        End If
       
        '-Filtra Recibos Col_CtaPag Inf-Regularizado= "6659072416125620 "  -------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        .Range.AutoFilter Field:=G04_InfRegulariz, Criteria1:="=6659072416125620*"
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416125620"
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "Regularizado G.Acad ???"
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "Regularizado G.Acad ???", 0)
        End If
       
        '-Filtra Recibos Col_CtaPag Inf-Regularizado= "(0049)"  ------------------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        .Range.AutoFilter Field:=G04_InfRegulariz, Criteria1:="=*(0049)"
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416125620"
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "Regularizado S.Inf."
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "Regularizado S.Inf.", 0)
        End If
       
        '-Filtra Recibos Col_CtaPag Inf-Regularizado= "(2100)"  ------------------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        .Range.AutoFilter Field:=G04_InfRegulariz, Criteria1:="=*(2100)"
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "2100 8984 16 0200003529"
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "Regularizado S.Inf."
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "Regularizado S.Inf.", 0)
        End If
       
        '-Filtra Recibos Col_CtaPag Inf-Regularizado= "(0081)"  ------------------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        .Range.AutoFilter Field:=G04_InfRegulariz, Criteria1:="=*(0081)"
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "0081 3191 42 0001068211"
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "Regularizado S.Inf."
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "Regularizado S.Inf.", 0)
        End If
       
        '-Filtra Recibos Col_CtaPag Inf-Regularizado= "(0014)"  ------------------------------------
        .AutoFilter.ShowAllData         ' Elimina los filtros
        .Range.AutoFilter Field:=G04_InfRegulariz, Criteria1:="=*(0014)"
        rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(Col_CtaIng).SpecialCells(xlCellTypeVisible).Cells.Value = "9000 0005 00 0260000014)"   '- Bco.Esp.
            TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Cta_CCC " & "Regularizado S.Inf."
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "Regularizado S.Inf.", 0)
        End If
       
        .AutoFilter.ShowAllData         ' Elimina los filtros
    
        '- Visualizo el progreso --------
        rowfind = Application.WorksheetFunction.CountIf(Lo_Data.DataBodyRange.Columns(Col_CtaIng), "")
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Rec. SIN Cta-CCC de Ingreso Asignados."
        rowfind = Application.WorksheetFunction.CountIf(Lo_Data.DataBodyRange.Columns(G04_CtaPag), "<>")
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Rec. CON Cta-CCC de Ingreso Asignados por el sistema."
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(.ListRows.Count, "#,##0"), 8) & " Rec. en BDatos"
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Rec. SIN Cta-CCC de Ingreso Asignados.", 0)
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & "x.xxx", 8) & " Rec. CON Cta-CCC de Ingreso Asignados previamente.", 0)
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(.ListRows.Count, "#,##0"), 8) & " Rec. en BDatos", 0)
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación de Cta-CCC de Ingreso.", TimeLapSub)
    
    End With    '- Lo_Data.
    
    If Sw_Boss Then Form_Menu.TB_Informe = TxtProgreso
    Lo_Data.ShowTotals = True
        
Call Rut_Lo_Filtros_Quitar(Lo_Data)
Debug.Print "<<< RuT_Determinar_Cta_Ingreso"
End Sub





