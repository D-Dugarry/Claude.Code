Attribute VB_Name = "M05_Asign_Tipo_Recibo"
' Last Rev. 2026-09-18 19:19
'2026-01-06
Option Explicit

            Sub RuT_Determinar_Tipo_Recibo_ByHand()
                Prog_LsGes04.Unprotect
                Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)          '- Quita filtros, filas y columnas ocultas
                Call RuT_Determinar_Tipo_Recibo
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Clasificar Recibos en Emitidos, Remesados, EjeAnt, ADxAplz, Añejas ---------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Determinar_Tipo_Recibo()
Debug.Print ">>> RuT_Determinar_Tipo_Recibo"
    Dim rowfind         As Variant
    Dim RegsCanTot      As Long
    Dim RegsSinTipo     As Long
    Dim APP_AñoCont     As String:          APP_AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim TxtProgreso     As String:          TxtProgreso = Form_Menu.TB_Informe
    Dim Sw_Boss         As Boolean:         Sw_Boss = Prog__APP_Switch.Range("Sw_Boss")
    
    Dim Lo_G04          As ListObject:      Set Lo_G04 = Prog_LsGes04.ListObjects(1)
    Application.ScreenUpdating = False
    Prog_LsGes04.Select
    Lo_G04.ShowTotals = False
    Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)
    
        '- Visualizo el progreso ---------------------------------------------------------------------------------------
        'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificación de Recibos, Estadística:", 0, , , , , , 2)
    With Lo_G04
    
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, BD_ACont_Emi, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_G04, BD_ACont_Vto, xlAscending, False)
        Call Rut_Lo_Sort(Lo_G04, BD_ACont_Cob, xlAscending, False)
            
        .DataBodyRange.Columns(BD_Tipo_Rec).ClearContents
        .DataBodyRange.Columns(G04_Flag_Primera).Resize(, G04_Flag_Cuantas).ClearContents

        TxtProgreso = TxtProgreso & vbCrLf & String(8, " ") & " Tipificado de Recibos. " & String(8, "_")
        
        '-Filtra Recibos ErrDate - -----------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, BD_FEmi, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_Err")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Incidencias).SpecialCells(xlCellTypeVisible).Cells.Value = "_ERR_Date_"
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_ERR_Date_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros con errores de fechas."
    
        '-Filtra Recibos Emitidos -------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .AutoFilter.ShowAllData            ' Elimina los filtros
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Emitido")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
            .DataBodyRange.Columns(G04_Flag_Emitido).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros Emitidos."

        '-Filtra Recibos EjeAnt -------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_EjeAnt")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "EjeAnt"
            .DataBodyRange.Columns(G04_Flag_EjeAnt).SpecialCells(xlCellTypeVisible).Cells.Value = "EjeAnt"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " Registros del Ejercicio Anterior. " & APP_AñoCont - 1
    
        '-Filtra Recibos Añejos -------------------------------------------------------------------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Añeja")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Añejo"
            .DataBodyRange.Columns(G04_Flag_Anejo).SpecialCells(xlCellTypeVisible).Cells.Value = "Añejo"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " Registros Añejos, anteriores a " & APP_AñoCont - 1
    
        '-Filtra Recibos Aplazado -------------------------------------------------------------------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Aplazado")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Aplazado"
            .DataBodyRange.Columns(G04_Flag_Aplazado).SpecialCells(xlCellTypeVisible).Cells.Value = "Aplazado"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " Registros Aplazados, Anulados por Aplazamiento en " & APP_AñoCont - 1
    
        '-Filtra Recibos ADxAplz -------------------------------------------------------------------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_ADxAplz")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz"
            .DataBodyRange.Columns(G04_Flag_ADxAplz).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " Registros ADxAplz, Anulados por Aplazamiento en " & APP_AñoCont & ", a cobrar en " & APP_AñoCont + 1
        
'        '-Filtra Recibos ADxAplz_SinCob -------------------------------------------------------------------------------
'        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_ADxAplz_SinCob")
'        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
'        If RowFind > 0 Then
'            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz_SinCob"
'            .DataBodyRange.Columns(G04_Flag_ADxAplz).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz_SinCob"
'            RegsCanTot = RegsCanTot + RowFind
'        End If
'        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
'                    " Registros ADxAplz_SinCob, Anulados por Aplazamiento en " & APP_AñoCont - 1 & "  y  ¡¡¡ SIN Cobrar en " & APP_AñoCont & " !!!"
        
        
        '---------------------------------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------------------
        '-Filtra Recibos Devolución - ----------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, BD_ImpRec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_ImpRec, Criteria1:="<0"
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_Devol_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros de devolución."
        
        '---------------------------------------------------------------------------------------------------------
        '-_Contab_Ant_--------------------------------------------------------------------------------------------
        '------------------ Filtra Cobradas en Años anteriores al de Emisión -------------------------------------
        '---------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .Range.AutoFilter Field:=BD_ACont_Cob, Criteria1:="<" & APP_AñoCont
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_AñoCont'" & Right(APP_AñoCont - 1, 2) & "_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros _Contab_Ant_, Cobrados anteriormente y por lo tanto, ya Contabilizado."
        
        '---------------------------------------------------------------------------------------------------------
        '-_Ajust_Matríc_------------------------------------------------------------------------------------------
        '------------------ Filtra Ajustes de Matrícula (ImpAdm negativo) ----------------------------------------
        '---------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .Range.AutoFilter Field:=BD_ImpAdm, Criteria1:="<0"
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_Ajust_Matríc_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros _Ajust_Matríc_, Ajustes de Matrícula (Imp. Admin. negativo)."
        
        '--------------------------------------------------------------------------------------------------------
        '-Filtra Recibos Sin Tipo ---------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        RegsSinTipo = Application.WorksheetFunction.CountIf(.DataBodyRange.Columns(BD_Tipo_Rec), "")

        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(RegsSinTipo, "#,##0"), 8) & " Registros SIN Tipificar, de" & Lo_G04.ListRows.Count & "reg."
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "=") & " " & Format(Lo_G04.ListRows.Count, "#,##0"), 8) & " Total Registros, Tipificados: " & RegsCanTot & "reg." & vbLf


        '---------------------------------------------------------------------------------------------------------
        '-Filtra Recibos Anul - -------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, BD_Anul, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_Anul")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Incidencias).SpecialCells(xlCellTypeVisible).Cells.Value = "_Reg_Anul_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros Anulados."
    
        '--------------------------------------------------------------------------------------------------------
        '-Filtra Recibos Sin Tipo ---------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        rowfind = Application.WorksheetFunction.CountIf(.DataBodyRange.Columns(BD_Tipo_Rec), "")
        If rowfind > 0 Then
'''            MsgBox "¡¡¡ Recibos SIN identificar su Tipo_Recibos !!! " & RowFind & "reg.", vbOKOnly + vbExclamation
        End If
        
        TxtProgreso = TxtProgreso & vbLf & _
            Right(String(8, "_") & Format(Application.WorksheetFunction.CountIf(.DataBodyRange.Columns(BD_Tipo_Rec), "="), "#,##0"), 8) & " 'Sin Identificar' Rec. SIN Determinar su Tipo (Si > 0, HAY QUE VERIFICAR FILTROS)" & vbCrLf & _
                  String(8, " ") & " Son Reg. que se han quedado fuera de todos los filtros, o filtros que duplican tipo." & vbCrLf & _
            Right(String(8, "=") & " " & Format(.ListRows.Count, "#,##0"), 8) & " Reg. en BDatos." & vbCrLf
       
        '- Oculta columnas para poder verificar los resultados, SÓLO CUANDO SE ESTÁ COMPROBANDO LA PROGRAMACIÓN..............
        .DataBodyRange.Columns(3).Resize(, 3).Hidden = True
        .DataBodyRange.Columns(9).Resize(, 5).Hidden = True
        .DataBodyRange.Columns(18).Resize(, 7).Hidden = True
        .DataBodyRange.Columns(26).Resize(, 8).Hidden = True
        .DataBodyRange.Columns(38).Resize(, 10).Hidden = True
        .DataBodyRange.Columns(49).Resize(, 3).Hidden = True
        
    End With    '- Lo_G04
    
    If Sw_Boss Then Form_Menu.TB_Informe = TxtProgreso
    Call Rut_Lo_Filtros_Quitar(Lo_G04)
    Lo_G04.ShowTotals = True
Application.ScreenUpdating = True
'MsgBox "Finn"
Debug.Print "<<< RuT_Determinar_Tipo_Recibo"
End Sub


