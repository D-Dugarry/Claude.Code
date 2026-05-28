Attribute VB_Name = "Rut_ListObjects"
Option Explicit

Sub Rut_LstObj_Columns_Ajustar_Ancho_ByHand()
    Call Rut_LstObj_Columns_Ajustar_Ancho(H_Liq_TPV, Prog_DefColTPVLiq)
End Sub
'###################################################################################################################################
Sub Rut_LstObj_Columns_Show_Hide(WrkSht As Worksheet, WsDefCol As Worksheet, Optional Reset As Boolean = False)
    Dim Lo_DefCol       As ListObject:      Set Lo_DefCol = WsDefCol.ListObjects(1)
    Dim Lo_Table        As ListObject:      Set Lo_Table = WrkSht.ListObjects(1)
    Dim Cont_Col        As Integer
    Dim Col_HiddenSw    As Integer:         Col_HiddenSw = Lo_DefCol.ListColumns("HiddenCol").Range.Column
    Dim ColEnBlco       As Integer:         ColEnBlco = Lo_Table.Range.Columns(1).Column - 1    '- Por si hay columnas en blanco a la derecha de la ListObject.
    Dim HiddenCol       As Boolean
    '- Si activo el Reset, lo único que hago es volver Mostrar/Ocultar tal como deberían estar (por si alguien las ha tocado)
    If Reset Then Lo_DefCol.TotalsRowRange(Col_HiddenSw) = Not Lo_DefCol.TotalsRowRange(Col_HiddenSw)
    '- Procedo a mostrar u ocultar las columnas.
    If Lo_DefCol.TotalsRowRange(Col_HiddenSw) = True Then
        WrkSht.Columns.Hidden = False
    Else
        For Cont_Col = 1 To Lo_Table.ListColumns.Count
            HiddenCol = Lo_DefCol.DataBodyRange.Cells(Cont_Col, Col_HiddenSw)      '.Value2
            WrkSht.Columns(ColEnBlco + Cont_Col).Hidden = HiddenCol
        Next Cont_Col
    End If
    Lo_DefCol.TotalsRowRange(Col_HiddenSw) = Not Lo_DefCol.TotalsRowRange(Col_HiddenSw)
End Sub
'###################################################################################################################################
Sub Rut_LstObj_Columns_Show(WrkSht As Worksheet, WsDefCol As Worksheet)
    Dim Lo_DefCol   As ListObject:      Set Lo_DefCol = WsDefCol.ListObjects(1)
    Dim Lo_Table    As ListObject:      Set Lo_Table = WrkSht.ListObjects(1)
    Dim Ctrl_Col    As Integer:         Ctrl_Col = Lo_DefCol.ListColumns("HiddenCol").Range.Column
        
    WrkSht.Columns.Hidden = False
    Lo_DefCol.TotalsRowRange(Ctrl_Col) = False

End Sub
'###################################################################################################################################
Sub Rut_LstObj_Columns_Ajustar_Ancho(WrkSht As Worksheet, WsDefCol As Worksheet)
    Dim Lo_DefCol   As ListObject:      Set Lo_DefCol = WsDefCol.ListObjects(1)
    Dim Lo_Table    As ListObject:      Set Lo_Table = WrkSht.ListObjects(1)
    Dim Cont_Col    As Integer
    Dim ColEnBlco   As Integer:         ColEnBlco = Lo_Table.Range.Columns(1).Column - 1
    Dim Ancho       As Integer
        
    For Cont_Col = 1 To Lo_Table.ListColumns.Count
        Ancho = Lo_DefCol.DataBodyRange.Cells(Cont_Col, Lo_DefCol.ListColumns("Ancho_Col").Range.Column)      '.Value2
        WrkSht.Columns(ColEnBlco + Cont_Col).ColumnWidth = Ancho
    Next Cont_Col

End Sub
'###################################################################################################################################
    ' Call Rut_LstObj_DataBodyRange_Filtered_DEL(LstObj, ColSearch, Criterio)
Sub Rut_LstObj_DataBodyRange_Filtered_DEL(ByRef LstObj As ListObject, ColSearch As Integer, Criterio As String)
    
    Dim Cant_Find  As Integer
    Application.DisplayAlerts = False
    Call Rut_LstObj_Filtros_Quitar(LstObj)                '- Quitar filtros
    Call Rut_LstObj_Sort(LstObj, ColSearch, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---------------------
    With LstObj
        .Range.AutoFilter Field:=ColSearch, Criteria1:=Criterio       '- Filtrar
        Cant_Find = .Range.Columns(ColSearch).SpecialCells(xlCellTypeVisible).Cells.Count - 2 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
        If Cant_Find > 0 Then .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete          '- Borrar Filas visibles
        .Range.AutoFilter                                                                   '- Quitar filtros
    End With
    Application.DisplayAlerts = True
    End Sub

'###################################################################################################################################
    ' Call Rut_LstObj_DataBodyRange_Copy(Lo_Source, Lo_Target, False/True)
Sub Rut_LstObj_DataBodyRange_Copy(ByRef Lo_Source As ListObject, ByRef Lo_Target As ListObject, Optional Borrar As Boolean = False)
    
    If Lo_Source.DataBodyRange Is Nothing Then MsgBox "Tabla de Origen Sin Datos": GoTo Fin
    If Not Lo_Target.DataBodyRange Is Nothing And Borrar Then Lo_Target.DataBodyRange.Delete
    Lo_Source.DataBodyRange.Copy
    If Lo_Target.DataBodyRange Is Nothing Then
        Lo_Target.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    Else
        Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0).PasteSpecial Paste:=xlPasteValues
    End If
    Application.CutCopyMode = False
Fin:
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
'###################################################################################################################################
    ' Call Rut_LstObj_DataBodyRange_Filtered_Copy(Lo_Source, Lo_Target, False/True)
Sub Rut_LstObj_DataBodyRange_Filtered_Copy(ByRef Lo_Source As ListObject, ByRef Lo_Target As ListObject, Optional Borrar As Boolean = False)
    
    If Lo_Source.DataBodyRange Is Nothing Then MsgBox "Tabla de Origen Sin Datos": GoTo Fin
    If Not Lo_Target.DataBodyRange Is Nothing And Borrar Then Lo_Target.DataBodyRange.Delete
    Lo_Source.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
    If Lo_Target.DataBodyRange Is Nothing Then
        Lo_Target.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    Else
        Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0).PasteSpecial Paste:=xlPasteValues
    End If
    Application.CutCopyMode = False
Fin:
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
'###################################################################################################################################
    ' Call Rut_LstObj_DataBodyRange_Filtered_Copy(Lo_Source, Lo_Target, ColSearch, Criterio, False/True)
Sub Rut_LstObj_DataBodyRange_Filter_And_Copy(ByRef Lo_Source As ListObject, ByRef Lo_Target As ListObject, ColSearch As Integer, Criterio As String, Optional Borrar As Boolean = False)
    
    Dim Cant_Find  As Integer
    Application.DisplayAlerts = False
    If Lo_Source.DataBodyRange Is Nothing Then MsgBox "Tabla de Origen Sin Datos": GoTo Fin
    If Not Lo_Target.DataBodyRange Is Nothing And Borrar Then Lo_Target.DataBodyRange.Delete
    Call Rut_LstObj_Filtros_Quitar(Lo_Source)                '- Quitar filtros
    Call Rut_LstObj_Sort(Lo_Source, ColSearch, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---------------------
    With Lo_Source
        .Range.AutoFilter Field:=ColSearch, Criteria1:=Criterio       '- Filtrar
        Cant_Find = .Range.Columns(ColSearch).SpecialCells(xlCellTypeVisible).Cells.Count - 1 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
        If Cant_Find > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
            If Lo_Target.DataBodyRange Is Nothing Then
                Lo_Target.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
            Else
                Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0).PasteSpecial Paste:=xlPasteValues
            End If
        End If
        .Range.AutoFilter                                                                  '- Quitar filtros
    End With
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
Fin:
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
            Sub Rut_LstObj_Buscar_Duplicados_ByHand()
                    Call Rut_LstObj_Buscar_Duplicados(ActiveSheet.ListObjects(1), 8, 31, True)  '- (ListObject, Colbusc, [ColNota]=0, [ClearColNota]=False)
            End Sub
'###################################################################################################################################
Sub Rut_LstObj_Buscar_Duplicados(ByRef Lo_Tb As ListObject, ColBusc As Integer, Optional ColNota As Integer = 0, Optional ClearColNota As Boolean = False)
Dim Lin         As Long
Dim ContDupl    As Long:    ContDupl = 0
Dim CantReg     As Long:    CantReg = Lo_Tb.ListRows.Count
Dim Ref_Ant     As Variant      '- Así da igual que sea Núm o Texto....

    Call Rut_LstObj_Sort(Lo_Tb, ColBusc, xlAscending, True)
    Debug.Print "CantReg: ", CantReg
    With Lo_Tb.DataBodyRange
    
        Ref_Ant = .Cells(1, ColBusc)
        If ColNota > 0 Then
            If ClearColNota Then .Columns(ColNota).ClearContents
            For Lin = 2 To CantReg
                If .Cells(Lin, ColBusc) = Ref_Ant Then
                    ContDupl = ContDupl + 1
                    .Cells(Lin, ColBusc).Interior.ColorIndex = 34
                    .Cells(Lin - 1, ColBusc).Interior.ColorIndex = 35
                        .Cells(Lin - 1, ColNota) = "Ojo Duplicado_1"
                        .Cells(Lin, ColNota) = "Ojo Duplicado_2"
                Else
                    Ref_Ant = .Cells(Lin, ColBusc)
                End If
    
                '- Visualizo el progreso ---------------------------------------------------------------------------------------
                If Lin Mod 100 = 0 Then
                    Debug.Print Format(Lin, "#,##0") & " de " & Format(CantReg, "#,##0")
    '                Form_Menu.TBx_Tarea_Inform = TxT_Progreso & "Incorporando LSGES04:  " & Format(LinPag, "#,##0") & " de " & Format(Lo_TPVpag.ListRows.Count, "#,##0")
    '                Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
    '                Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
                End If
            Next Lin
        Else
            For Lin = 2 To CantReg
                If .Cells(Lin, ColBusc) = Ref_Ant Then
                    ContDupl = ContDupl + 1
                    .Cells(Lin, ColBusc).Interior.ColorIndex = 34
                    .Cells(Lin - 1, ColBusc).Interior.ColorIndex = 35
                Else
                    Ref_Ant = .Cells(Lin, ColBusc)
                End If
                '- Visualizo el progreso ---------------------------------------------------------------------------------------
                If Lin Mod 100 = 0 Then
                    Debug.Print Format(Lin, "#,##0") & " de " & Format(CantReg, "#,##0")
    '                Form_Menu.TBx_Tarea_Inform = TxT_Progreso & "Incorporando LSGES04:  " & Format(LinPag, "#,##0") & " de " & Format(Lo_TPVpag.ListRows.Count, "#,##0")
    '                Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
    '                Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
                End If
            Next Lin
        End If
    End With
Debug.Print ContDupl
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
'###################################################################################################################################
Sub Rut_LstObj_WrkSht_Preparar(WrkSht As Worksheet)
' ----------------------------------------------------------------------------------------------------------------------------------
    With WrkSht
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        Call Rut_LstObj_Filtros_Quitar(.ListObjects(1))
    End With
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

'###################################################################################################################################
'   call Rut_LstObj_Sort(Lo_Tb , Columna , VarOrden , [SW_Clear]= False)
Sub Rut_LstObj_Sort(ByRef Lo_Tb As ListObject, Columna As Integer, VarOrden As String, Optional SW_Clear As Boolean = False)
' ----------------------------------------------------------------------------------------------------------------------------------
    With Lo_Tb.Sort
        If SW_Clear Then .sortFields.Clear
        .sortFields.Add Key:=Lo_Tb.ListColumns(Columna).Range, SortOn:=xlSortOnValues, Order:=VarOrden, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

'###################################################################################################################################
    '-  Call Rut_LstObj_Filtro(Lo_Tb , Columna , Criterio, [SW_Clear] = False)
Sub Rut_LstObj_Filtro(ByRef Lo_Tb As ListObject, Columna As Integer, Criterio As String, Optional SW_Clear As Boolean = False)
' ----------------------------------------------------------------------------------------------------------------------------------
    If SW_Clear And Not Lo_Tb.AutoFilter Is Nothing Then Lo_Tb.AutoFilter.ShowAllData
    Lo_Tb.Range.AutoFilter Field:=Columna, Criteria1:=Criterio
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

                            '#######################################################################################################
                            Sub Rut_Filtros_Quitar_ActivSheet_LstObj()      ' Muestra Todas las Solicitudes y Activar Filtros >>>>>>
                            Dim Protect_Status   As Boolean
                            Protect_Status = ActiveSheet.ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
                            Call Rut_LstObj_Filtros_Quitar(ActiveSheet.ListObjects(1))
                            If Protect_Status Then ActiveSheet.Protect
                            End Sub     ' ---------------------------------------------------------------------------------------<<<
'###################################################################################################################################
    '   call Rut_LstObj_Filtros_Quitar(Lo_Tb)
Sub Rut_LstObj_Filtros_Quitar(ByRef Lo_Tb As ListObject)      ' Muestra Todas las Solicitudes y Activar Filtros >>>>>>>>>>>>>>>>>>>>
' ----------------------------------------------------------------------------------------------------------------------------------
    With Lo_Tb
        If .ShowAutoFilter Then                         '--- Compruebo que la Tabla tiene los Filtros Activos --------
            With .AutoFilter
                 If .FilterMode Then .ShowAllData
            End With
        Else                                            '--- Si NO tiene los Filtros Activos, los Activo -------------
            .ShowAutoFilter = True
        End If
    End With
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

Sub Rut_Copiar_EntireRow_LstObjct() '- Rut_Añade_Row_entera, DataBodyRange.Rows(2) ES OBLIGATORIO sino no funciona
                
ActiveSheet.ListObjects("Tab_INI").DataBodyRange.Rows(2).Copy ActiveSheet.ListObjects("Tab_FIN").ListRows(3).Range '
                
End Sub






