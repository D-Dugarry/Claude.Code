Attribute VB_Name = "Rut_Lo"
' Last Rev. 2026-09-30 18:25
Option Explicit

'###################################################################################################################################
Sub Rut_Lo_Columns_Show_Hide(WrkSht As Worksheet, WsDefCol As Worksheet, Col_HiddenSw As Integer, Optional Reset As Boolean = False)
Debug.Print "Rut_Lo_Columns_Show_Hide"
    Dim Lo_DefCol       As ListObject:      Set Lo_DefCol = WsDefCol.ListObjects(1)
    Dim Lo_Table        As ListObject:      Set Lo_Table = WrkSht.ListObjects(1)
    Dim Cont_Col        As Integer
    Dim ColEnBlco       As Integer:         ColEnBlco = Lo_Table.Range.Columns(1).Column - 1    '- Por si hay columnas en blanco a la derecha de la ListObject.
    Dim HiddenCol       As Boolean
    Dim SetHidden       As Variant
    Dim SW_Col_Hide_Name  As String:        SW_Col_Hide_Name = "SW_Col_Hide_" & WrkSht.CodeName
    Dim SW_Col_Hide     As Boolean:         SW_Col_Hide = Prog__APP.Range(SW_Col_Hide_Name).Value2
    'Rango completo de columnas afectadas
    Dim RngCols         As Range:           Set RngCols = WrkSht.Range(WrkSht.Columns(ColEnBlco + 1), WrkSht.Columns(ColEnBlco + Lo_Table.ListColumns.Count))

    If Reset Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        WrkSht.Columns.Hidden = False
        Lo_DefCol.TotalsRowRange(Col_HiddenSw).Value = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    If SW_Col_Hide = True Then
        'Un solo disparo para mostrar todas
        RngCols.EntireColumn.Hidden = False
    Else
        'Leer de una vez la columna de definición en una matriz
        SetHidden = Lo_DefCol.DataBodyRange.Columns(Col_HiddenSw).Value2
        'Aplicar Hidden por filas de la matriz
        For Cont_Col = 1 To Lo_Table.ListColumns.Count
            HiddenCol = SetHidden(Cont_Col, 1)
            WrkSht.Columns(ColEnBlco + Cont_Col).Hidden = HiddenCol
        Next Cont_Col
    End If

    'Actualizar flags solo una vez
    Lo_DefCol.ShowTotals = True
    With Lo_DefCol.TotalsRowRange(Col_HiddenSw)
        .Value = Not .Value
    End With
    Prog__APP.Range(SW_Col_Hide_Name).Value = Not SW_Col_Hide

End Sub

' ==================================================================================================================================
Function Func_LstObj_ListColumns_DefCol_Check_OK(ByRef Lo_Data As ListObject, ByRef LoDefCol As ListObject, Col_TitColCompare As Integer) As Boolean
' ==================================================================================================================================
    '- Compara el Nombre de las columnas de una Tabla con el Nombre que debería tener según la tabla LoDefCol ----------------------
    Dim Ccol                As Integer
    Func_LstObj_ListColumns_DefCol_Check_OK = False
    '--- Verificar las Columnas --------------------------------------------------------------------------------------
    For Ccol = 1 To LoDefCol.DataBodyRange.Rows.Count
    
        If LoDefCol.DataBodyRange.Cells(Ccol, Col_TitColCompare) = "" Then GoTo Finalizar
        
        If LoDefCol.DataBodyRange.Cells(Ccol, Col_TitColCompare) <> Lo_Data.HeaderRowRange.Cells(Ccol) Then
        
            MsgBox " En la Col. nº " & Ccol & vbLf & _
                   " Nombre Col. debe ser: " & LoDefCol.DataBodyRange.Cells(Ccol, Col_TitColCompare) & vbLf & _
                   " Nombre Col. es . . .: " & Lo_Data.HeaderRowRange.Cells(Ccol) _
                   , , "Procedimiento: Comparación Fila de Títulos de tablas"
            
            Exit Function
        End If
     Next
     
Finalizar:
    Func_LstObj_ListColumns_DefCol_Check_OK = True
End Function
' ==================================================================================================================================
' ==================================================================================================================================
Function Func_LstObj_HeaderRow_Check_2Lo_OK(ByRef Lo1 As ListObject, ByRef Lo2 As ListObject) As Boolean
' ==================================================================================================================================
    '- Compara el Nombre de las columnas de una Tabla con el Nombre que debería tener según la tabla LoDefCol ----------------------
    Dim Ccol                As Integer
    Func_LstObj_HeaderRow_Check_2Lo_OK = False
    '--- Verificar las Columnas --------------------------------------------------------------------------------------
    For Ccol = 1 To Lo1.ListColumns.Count
    
        If Lo1.HeaderRowRange.Cells(Ccol) <> Lo2.HeaderRowRange.Cells(Ccol) Then
        
            MsgBox " En la Col. nº " & Ccol & vbLf & _
                   " Nombre Col. Lo1: " & Lo1.HeaderRowRange.Cells(Ccol) & vbLf & _
                   " Nombre Col. Lo2: " & Lo2.HeaderRowRange.Cells(Ccol) _
                   , , "Procedimiento: Comparación Fila de Títulos de 2 tablas"
            
            Exit Function
        End If
     Next
     
Finalizar:
    Func_LstObj_HeaderRow_Check_2Lo_OK = True
End Function
' ==================================================================================================================================
'###################################################################################################################################
Sub Rut_Lo_Columns_Ajustar_Ancho_con_DefCol()
    Dim NameSheet     As String:    NameSheet = Right(ActiveSheet.Name, Len(ActiveSheet.Name) - 8)
    Dim i       As Integer
    Dim AnchoCol    As Integer
    Dim LoDefCol    As ListObject:      Set LoDefCol = ActiveSheet.ListObjects(1)
    Sheets(NameSheet).Visible = xlSheetVisible
    Dim LoSheet     As ListObject:      Set LoSheet = Sheets(NameSheet).ListObjects(1)
    With LoDefCol
        For i = 1 To .ListRows.Count
            AnchoCol = .DataBodyRange.Cells(i, .ListColumns("Ancho col").Index)
            LoSheet.ListColumns(i).Range.ColumnWidth = AnchoCol
        Next i
    End With
'    Dim Cont_Col As Integer
'    Dim Ancho   As Integer
'    For Cont_Col = 1 To LastCol_Tb_Solicitudes
'        If Not Columns(Cont_Col).Hidden And Lo_Prog_Colns.DataBodyRange.Cells(7, Cont_Col) <> "Ocultar" Then
'            Ancho = Lo_Prog_Colns.DataBodyRange.Cells(4, Cont_Col)      '.Value2
'            Columns(Cont_Col).ColumnWidth = Ancho
'        End If
'    Next Cont_Col
End Sub
' ==================================================================================================================================
Sub Rut_Lo_ListColumns_ClearContents_DefC_ProtectData(Lo_Data As ListObject, _
                                                      LoDefCol As ListObject, _
                                                      Col_DefC_ProtectData As Integer)
' ==================================================================================================================================
    '- Borra el contenido (delicado y no estrictamente necesario) de las columnas de una Tabla según su tabla LoDefCol ----------------------
    Dim Ccol                As Integer
    For Ccol = 1 To LoDefCol.DataBodyRange.Rows.Count
    
        If LoDefCol.DataBodyRange.Cells(Ccol, Col_DefC_ProtectData) Then Lo_Data.ListColumns(Ccol).DataBodyRange.ClearContents
     
     Next
End Sub
' ==================================================================================================================================
'''    '###################################################################################################################################
'''    ' Copia el DataBodyRange FILTRADO de Lo_Source en Lo_Target,      Opcional: Borrar primero contenido de Lo_Target
'''    ' Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Source,Lo_Target,[DelFirstLoTarget=False])
'''    Sub Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Source As ListObject, _
'''                                            Lo_Target As ListObject, _
'''                                            Optional DelFirstLoTarget As Boolean = False)
'''    ' ----------------------------------------------------------------------------------------------------------------------------------
'''    Debug.Print "Rut_Lo_DataBodyRange_Filtered_Copy"
'''        Lo_Target.ShowTotals = False
'''        If DelFirstLoTarget And Not Lo_Target.DataBodyRange Is Nothing Then Lo_Target.DataBodyRange.Delete
'''        If Lo_Target.DataBodyRange Is Nothing Then
'''            Lo_Source.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy _
'''                    Destination:=Lo_Target.Range.Cells(1, 1).Offset(1, 0)
'''        Else
'''            Lo_Source.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy _
'''                    Destination:=Lo_Target.DataBodyRange.Cells(Lo_Target.DataBodyRange.Rows.Count, 1).Offset(1, 0)
'''        End If
'''        Application.CutCopyMode = False
'''    End Sub
'''    ' ----------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
' Copia el DataBodyRange FILTRADO de Lo_Source en Lo_Target,      Opcional: Borrar primero contenido de Lo_Target y Borrar los registros filtrados de Lo_Source
Sub Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Source As ListObject, _
                                       Lo_Target As ListObject, _
                                       Optional DelFirstLoTarget As Boolean = False, _
                                       Optional DelLoSourceFilteredRows As Boolean = False)
' ----------------------------------------------------------------------------------------------------------------------------------
Debug.Print "Rut_Lo_DataBodyRange_Filtered_Copy"
    Dim Ws As Worksheet
    Set Ws = Lo_Target.Parent
    
    If Lo_Target.ShowTotals Then Lo_Target.ShowTotals = False
    If DelFirstLoTarget And Not Lo_Target.DataBodyRange Is Nothing Then Lo_Target.DataBodyRange.Delete
    
    Dim RangoACopiar    As Range
    On Error Resume Next
    Set RangoACopiar = Lo_Source.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If RangoACopiar Is Nothing Then Exit Sub
    
    '- Pega datos justo debajo de la Lo_Target
    Dim StartRowAdd     As Long
    If Lo_Target.DataBodyRange Is Nothing Then
'        Lo_Target.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
        StartRowAdd = Lo_Target.Range.Offset(1, 0).Row
        RangoACopiar.Copy Destination:=Lo_Target.Range.Offset(1, 0)
   Else
'        Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0).PasteSpecial Paste:=xlPasteValues
        StartRowAdd = Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0).Row
        RangoACopiar.Copy Destination:=Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0)
    End If
    
    '- Como copio un Rango, Lo_Target NO se expande, Sólo se copia a continuación y forman parte de la Listobject.
    '- Tengo que Redimensionar la tabla para incluir las nuevas filas en la Listobject
    Dim RowsACopiar         As Long:        RowsACopiar = RangoACopiar.Rows.Count
    Dim NuevoRangoAmpliado  As Range       '- Defino un NuevoRango que abarca la Lo_Target más el Rango Copiado.
    Set NuevoRangoAmpliado = Ws.Range(Lo_Target.Range.Cells(1, 1), _
                             Ws.Cells(StartRowAdd + RowsACopiar - 1, Lo_Target.Range.Column + Lo_Target.Range.Columns.Count - 1))
    Lo_Target.Resize NuevoRangoAmpliado   '- Redefino Lo_Target con el tamaño de Lo_Target más el Rango Copiado: NuevoRangoAmpliado
    
    If DelLoSourceFilteredRows Then RangoACopiar.Delete     '- Como estamos dentro del "IF Not RangoACopiar Is Nothing" el Rango tiene datos y los podemos Borrar
    Application.CutCopyMode = False
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
    ' Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, ColSearch1, Criterio1)    ¡¡¡ QUITA FILTROS SI HAY  !!!
Sub Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data As ListObject, _
                                         ColSearch1 As Integer, _
                                         Criterio1 As String, _
                                Optional ColSearch2 As Integer = 1, _
                                Optional Criterio2 As String = "")
    
    If Lo_Data.DataBodyRange Is Nothing Then Exit Sub
    Dim RowsFind  As Variant
    Dim Sw_ShowTotals    As Boolean:    Sw_ShowTotals = Lo_Data.ShowTotals:    Lo_Data.ShowTotals = False
    Call Rut_Lo_Filtros_Quitar(Lo_Data)                '- Quitar filtros
    With Lo_Data
        If Criterio2 = "" Then                                          '- 1 criterio
            Call Rut_Lo_Sort(Lo_Data, ColSearch1, xlAscending, True)
            .Range.AutoFilter Field:=ColSearch1, Criteria1:=Criterio1
        Else                                                            '- 2 criterios
            Call Rut_Lo_Sort(Lo_Data, ColSearch1, xlAscending, True)
            Call Rut_Lo_Sort(Lo_Data, ColSearch2, xlAscending)
            .Range.AutoFilter Field:=ColSearch1, Criteria1:=Criterio1
            .Range.AutoFilter Field:=ColSearch2, Criteria1:=Criterio2, Operator:=xlAnd
        End If
        .ShowTotals = False
        RowsFind = .Range.Columns(ColSearch1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 '2 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
        If RowsFind > 0 Then .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete          '- Borrar Filas visibles
    End With
    Call Rut_Lo_Filtros_Quitar(Lo_Data)                '- Quitar filtros
    Lo_Data.ShowTotals = Sw_ShowTotals
Debug.Print "Rut_Lo_DataBodyRange_Filter_y_DEL: ColSearch1: " & ColSearch1 & ", Criterio1: _" & Criterio1 & ", Criterio2: _" & Criterio2 & "_ RowsFind: _" & RowsFind & " reg."
End Sub
'###################################################################################################################################
'---------- Copia ColSource en ColTarget de los valores Filtrados ------------------------------------------------------------------2026-01-24
Sub Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(Lo_Data As ListObject, _
                                                                   ColSource As Integer, _
                                                                   ColTarget As Integer, _
                                                                   ColCrit_1 As Integer, _
                                                                   Criterio1 As String, _
                                                                   Optional ColCrit_2 As Integer = 0, _
                                                                   Optional Criterio2 As String = "")
    If Lo_Data.DataBodyRange Is Nothing Then Exit Sub
    Dim RowsFind     As Variant
    Dim Sw_ShowTotals    As Boolean:    Sw_ShowTotals = Lo_Data.ShowTotals:    Lo_Data.ShowTotals = False
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    With Lo_Data
        If Criterio2 = "" Then                                          '- 1 criterio
            Call Rut_Lo_Sort(Lo_Data, ColCrit_1, xlAscending, True)
            .Range.AutoFilter Field:=ColCrit_1, Criteria1:=Criterio1
        Else                                                            '- 2 criterios
            If ColCrit_1 = ColCrit_2 Then                               '- 2 Criterios en la misma Columna
                Call Rut_Lo_Sort(Lo_Data, ColCrit_1, xlAscending, True)
                .Range.AutoFilter Field:=ColCrit_1, Criteria1:=Criterio1, Operator:=xlAnd, _
                                                    Criteria1:=Criterio2
            Else                                                        '- 2 Criterios en Distintas Columnas
                Call Rut_Lo_Sort(Lo_Data, ColCrit_1, xlAscending, True)
                Call Rut_Lo_Sort(Lo_Data, ColCrit_2, xlAscending)
                .Range.AutoFilter Field:=ColCrit_1, Criteria1:=Criterio1
                .Range.AutoFilter Field:=ColCrit_2, Criteria1:=Criterio2, Operator:=xlAnd
            End If
        End If
        RowsFind = .Range.Columns(ColSource).SpecialCells(xlCellTypeVisible).Cells.Count - 1 '2 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
        If RowsFind > 0 Then
            Dim RngSource As Range
            Dim RngTarget As Range
            Set RngSource = Lo_Data.ListColumns(ColSource).DataBodyRange
            Set RngSource = RngSource.SpecialCells(xlCellTypeVisible)
            Set RngTarget = Lo_Data.ListColumns(ColTarget).DataBodyRange
            Set RngTarget = RngTarget.SpecialCells(xlCellTypeVisible)
            RngSource.Copy
            RngTarget.PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If
    End With
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Lo_Data.ShowTotals = Sw_ShowTotals
Debug.Print "Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget" & _
            ": ColSource: " & ColSource & ", ColTarget: " & ColTarget & _
            ", ColCrit_1: _" & ColCrit_1 & ", Criterio1: _" & Criterio1 & _
            ", ColCrit_2: _" & ColCrit_2 & ", Criterio2: _" & Criterio2 & _
            "_ RowsFind: _" & RowsFind & " reg."
End Sub
'- ------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
Sub Rut_Lo_WrkSht_Preparar(WrkSht As Worksheet)
' ----------------------------------------------------------------------------------------------------------------------------------
    With WrkSht
        .Unprotect
        .Columns.EntireColumn.Hidden = False        '-1º Mostrar todas las Columnas
        If Fnc_Range_Exist("SW_Col_Hide_" & WrkSht.CodeName) Then Prog__APP.Range("SW_Col_Hide_" & WrkSht.CodeName) = False
        .Rows.EntireRow.Hidden = False              '-2º Mostrar todas las Filas
        Call Rut_Lo_Filtros_Quitar(.ListObjects(1)) '-3º Quitar Filtros
'        .Protect allowFiltering:=True, DrawingObjects:=True, allowSorting:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
    End With
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

'###################################################################################################################################
Sub Rut_Lo_Sort(ByRef Lo_Tb As ListObject, Columna As Integer, VarOrden As String, Optional SW_Clear As Boolean = False)
' ----------------------------------------------------------------------------------------------------------------------------------
    With Lo_Tb.Sort
        If SW_Clear Then .SortFields.Clear
        .SortFields.Add Key:=Lo_Tb.ListColumns(Columna).Range, SortOn:=xlSortOnValues, Order:=VarOrden, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
' ==================================================================================================================================
Sub Rut_Lo_FreezePanes_InmovilizaFxCx(Optional XCol As Integer = 1)
    'Call Rut_Lo_FreezePanes_InmovilizaFxCx(2)  '- Inmoviliza la Fila de Cabecera y la Xcol de la LstObj de la ActiveSheet.
Debug.Print "Rut_Lo_FreezePanes_InmovilizaFxCx"
    Dim Ws As Worksheet
    Dim Tbl As ListObject
    Dim filaInicio As Long
    Dim columnaInicio As Long
    ' Asigna la hoja activa a la variable ws
    Set Ws = ActiveSheet
    ' Verifica si hay tablas en la hoja
    If Ws.ListObjects.Count = 0 Then
        MsgBox "No hay tablas en esta hoja.", vbExclamation, "Error"
        Exit Sub
    End If
    ' Obtiene la primera tabla en la hoja
    Set Tbl = Ws.ListObjects(1)
    ' Obtiene la fila y columna de inicio de la tabla
    filaInicio = Tbl.HeaderRowRange.Row + 1
    columnaInicio = Tbl.Range.Column
    ' Desactiva cualquier inmovilización actual
    If Not ActiveWindow Is Nothing Then
        If ActiveWindow.FreezePanes Then
            ActiveWindow.FreezePanes = False
        End If
    End If
    ' Activa la inmovilización en la fila y columna de inicio de la tabla
    Ws.Cells(filaInicio, columnaInicio + XCol).Select
    ActiveWindow.FreezePanes = True
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

'###################################################################################################################################
Sub Rut_Lo_Filtros_Quitar(Lo_Tb As ListObject)      ' Muestra Todas las Solicitudes y Activar Filtros >>>>>>>>>>>>>>>>>>>>
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






