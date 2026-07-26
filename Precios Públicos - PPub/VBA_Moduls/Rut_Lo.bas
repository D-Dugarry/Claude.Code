Attribute VB_Name = "Rut_Lo"
Option Explicit

'###################################################################################################################################
            Sub Rut_Lo_Buscar_Duplicados_ByHand()
                                        Call Rut_Lo_Buscar_Duplicados(ActiveSheet.ListObjects(1), 11, 44, True) '- (ListObject, Colbusc, [ColNota]=0, [ClearColNota]=False)
            End Sub
'###################################################################################################################################
Sub Rut_Lo_Buscar_Duplicados(ByRef Lo_Tb As ListObject, ColBusc As Integer, Optional ColNota As Integer = 0, Optional ClearBeforeColNota As Boolean = False)
Dim Lin         As Long
Dim ContDupl    As Long:    ContDupl = 0
Dim CantReg     As Long:    CantReg = Lo_Tb.ListRows.Count
Dim Ref_Ant     As Variant      '- Así da igual que sea Núm o Texto....

    Call Rut_Lo_Sort(Lo_Tb, ColBusc, xlAscending, True)
    Debug.Print "CantReg: " & CantReg
    With Lo_Tb.DataBodyRange
        Ref_Ant = .Cells(1, ColBusc)
        If ColNota > 0 Then
            If ClearBeforeColNota Then .Columns(ColNota).ClearContents
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
            Next Lin
        End If
    End With
MsgBox "FIN"
Debug.Print "ContDupl: " & ContDupl
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

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

'###################################################################################################################################
Sub Rut_Lo_Columns_Show_Hide_OLD(WrkSht As Worksheet, WsDefCol As Worksheet, Col_HiddenSw As Integer, Optional Reset As Boolean = False)
Debug.Print "Rut_Lo_Columns_Show_Hide"
    Dim Lo_DefCol       As ListObject:      Set Lo_DefCol = WsDefCol.ListObjects(1)
    Dim Lo_Table        As ListObject:      Set Lo_Table = WrkSht.ListObjects(1)
    Dim Cont_Col        As Integer
    Dim ColEnBlco       As Integer:         ColEnBlco = Lo_Table.Range.Columns(1).Column - 1    '- Por si hay columnas en blanco a la derecha de la ListObject.
    Dim HiddenCol       As Boolean
    '- Si activo el Reset, lo único que hago es volver Mostrar/Ocultar tal como deberían estar (por si alguien las ha tocado)
    If Reset Then
        WrkSht.Columns.Hidden = False  ' WrkSht.Columns.EntireColumn.Hidden = False
        Lo_DefCol.TotalsRowRange(Col_HiddenSw).Value = False
    End If
    '- Procedo a mostrar u ocultar las columnas.
    If Prog__APP.Range("SW_Col_Hide_" & WrkSht.CodeName) = True Then
        For Cont_Col = 1 To Lo_Table.ListColumns.Count
            WrkSht.Columns(ColEnBlco + Cont_Col).Hidden = False
        Next Cont_Col
    Else
        For Cont_Col = 1 To Lo_Table.ListColumns.Count
            HiddenCol = Lo_DefCol.DataBodyRange.Cells(Cont_Col, Col_HiddenSw)      '.Value2
            WrkSht.Columns(ColEnBlco + Cont_Col).Hidden = HiddenCol
        Next Cont_Col
    End If
    Lo_DefCol.TotalsRowRange(Col_HiddenSw) = Not Lo_DefCol.TotalsRowRange(Col_HiddenSw)
    Prog__APP.Range("SW_Col_Hide_" & WrkSht.CodeName) = Not Prog__APP.Range("SW_Col_Hide_" & WrkSht.CodeName)
FinSub:
End Sub
'###################################################################################################################################
Sub Rut_Lo_Columns_Show_Hide_VeryOLD(WrkSht As Worksheet, WsDefCol As Worksheet, Optional Reset As Boolean = False)
Debug.Print "Rut_Lo_Columns_Show_Hide"
    Dim Lo_DefCol       As ListObject:      Set Lo_DefCol = WsDefCol.ListObjects(1)
    Dim Lo_Table        As ListObject:      Set Lo_Table = WrkSht.ListObjects(1)
    Dim Cont_Col        As Integer
'    Dim Col_HiddenSw    As Integer:         Col_HiddenSw = Lo_DefCol.ListColumns("HiddenCol").Range.Column
    Dim Col_HiddenSw    As Integer:         Col_HiddenSw = Lo_DefCol.ListColumns("HiddenCol").Range.Column
    Dim ColEnBlco       As Integer:         ColEnBlco = Lo_Table.Range.Columns(1).Column - 1    '- Por si hay columnas en blanco a la derecha de la ListObject.
    Dim HiddenCol       As Boolean
    '- Si activo el Reset, lo único que hago es volver Mostrar/Ocultar tal como deberían estar (por si alguien las ha tocado)
    If Reset Then Lo_DefCol.TotalsRowRange(Col_HiddenSw) = Not Lo_DefCol.TotalsRowRange(Col_HiddenSw)
    '- Procedo a mostrar u ocultar las columnas.
    If Lo_DefCol.TotalsRowRange(Col_HiddenSw) = True Then
        For Cont_Col = 1 To Lo_Table.ListColumns.Count
            WrkSht.Columns(ColEnBlco + Cont_Col).Hidden = False
        Next Cont_Col
    Else
        For Cont_Col = 1 To Lo_Table.ListColumns.Count
            HiddenCol = Lo_DefCol.DataBodyRange.Cells(Cont_Col, Col_HiddenSw)      '.Value2
            WrkSht.Columns(ColEnBlco + Cont_Col).Hidden = HiddenCol
        Next Cont_Col
    End If
    Lo_DefCol.TotalsRowRange(Col_HiddenSw) = Not Lo_DefCol.TotalsRowRange(Col_HiddenSw)
End Sub
'###################################################################################################################################
        Sub Func_LstObj_ListColumns_DefCol_Check_OK_ByHand()  '--- Solo hace falta cambiar las variables de la Rutina -----------------------
            MsgBox "La comparación de columnas de la tabla ha resultado ser: " & vbLf & vbLf & _
                   "                        " & Func_LstObj_ListColumns_DefCol_Check_OK(Sht__BD.ListObjects(1), Prog_DefCol_BD.ListObjects(1), DefC_TitColLstObj) _
                   , , "Procedimiento: Comparación Fila de Títulos de tablas"
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
'###################################################################################################################################
' Copia el DataBodyRange de Lo_Source en Lo_Target,      Opcional: DelFirstLoTarget primero contenido de Lo_Target
Sub Rut_Lo_DataBodyRange_Copy(Lo_Source As ListObject, _
                                  Lo_Target As ListObject, _
                                  Optional DelFirstLoTarget As Boolean = False, _
                                  Optional DelAfterLoSource As Boolean = False, _
                                  Optional WithFormat As Boolean = False)
' ----------------------------------------------------------------------------------------------------------------------------------
    If Not Lo_Target.DataBodyRange Is Nothing And DelFirstLoTarget Then Lo_Target.DataBodyRange.Delete
    Lo_Source.DataBodyRange.Copy
    If WithFormat Then
        If Lo_Target.DataBodyRange Is Nothing Then
            Lo_Target.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteAll
        Else
            Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0).PasteSpecial Paste:=xlPasteAll
        End If
    Else
        If Lo_Target.DataBodyRange Is Nothing Then
            Lo_Target.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
        Else
            Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0).PasteSpecial Paste:=xlPasteValues
        End If
    End If
    '- Del Source Range if required -----------------------------------
    If DelAfterLoSource Then Lo_Source.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
    Application.CutCopyMode = False
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
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
    ' Call Rut_Lo_DataBodyRange_Filter_Copy_Del(Lo_Source,Lo_Target,ColSearch1, Criterio1,ColSearch2, Criterio2,RngCriteria,DelFirstLoTarget,DelAfterLoSource)
Sub Rut_Lo_DataBodyRange_Filter_Copy_Del(Lo_Source As ListObject, _
                                           Lo_Target As ListObject, _
                                           Optional ColSearch1 As Integer = 0, Optional Criterio1 As String = "", _
                                           Optional ColSearch2 As Integer = 0, Optional Criterio2 As String = "", _
                                           Optional RngCriteria As String = "", _
                                           Optional DelFirstLoTarget As Boolean = False, _
                                           Optional DelAfterLoSource As Boolean = False)
'------------------------------------------------------------------------------------------------------------------------------------
Debug.Print "Rut_Lo_DataBodyRange_Filter_Copy_Del"
    Dim RowsFind  As Variant
    If Lo_Source.DataBodyRange Is Nothing Then Exit Sub
    Lo_Source.ShowTotals = False
    Lo_Target.ShowTotals = False
    '- Borrar Target before copy ---------------------------
    If Not Lo_Target.DataBodyRange Is Nothing And DelFirstLoTarget Then Lo_Target.DataBodyRange.Delete
    
    Call Rut_Lo_Filtros_Quitar(Lo_Source)                '- Quitar filtros
    If ColSearch1 <> 0 Then
                                Call Rut_Lo_Sort(Lo_Source, ColSearch1, xlAscending, True)     '- Ordenar primero accelera un montón el borrado -----
        If ColSearch2 <> 0 Then Call Rut_Lo_Sort(Lo_Source, ColSearch2, xlAscending, False)
    End If
    With Lo_Source
        '- Hay Rango-Criterio ---------------------------------
        If RngCriteria <> "" Then
            .Range.AdvancedFilter xlFilterInPlace, Range(RngCriteria)
            'RowsFind = .Range.SpecialCells(xlCellTypeVisible).Rows.Count - 1 '- Falla a veces y no he averiguado porque, habiendo celdas filtradas, no funciona!!!
            RowsFind = .Range.Columns(ColSearch1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 '2 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
            If RowsFind = 0 Then GoTo Terminar
        '- NO Hay Rango-Criterio ---------------------------------
        Else
            .Range.AutoFilter Field:=ColSearch1, Criteria1:=Criterio1       '- Filtrar
            If Criterio2 <> "" Then
                .Range.AutoFilter Field:=ColSearch2, Criteria1:=Criterio2       '- Filtrar
            End If
            'RowsFind = .Range.SpecialCells(xlCellTypeVisible).Rows.Count - 1 '- Falla a veces y no he averiguado porque, habiendo celdas filtradas, no funciona!!!
            RowsFind = .Range.Columns(ColSearch1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 '2 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
            If RowsFind = 0 Then GoTo Terminar
        End If
        '- Copy Filtered Range --------------------------------------------
        If Lo_Target.DataBodyRange Is Nothing Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=Lo_Target.Range.Cells(1, 1).Offset(1, 0)
        Else
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=Lo_Target.DataBodyRange.Cells(Lo_Target.DataBodyRange.Rows.Count, 1).Offset(1, 0)
        End If
    End With    '- Lo_Source
    '- Del Source Range if required -----------------------------------
    If DelAfterLoSource Then Lo_Source.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
    Application.CutCopyMode = False
    
Terminar:
    Call Rut_Lo_Filtros_Quitar(Lo_Source)                '- Quitar filtros
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

            Sub Rut_Lo_DataBodyRange_Filter_y_DEL_ByHand()
                Call Rut_Lo_DataBodyRange_Filter_y_DEL(ActiveSheet.ListObjects(1), _
                                LS06_C_Acad, "<>2025-26", LS06_C_Acad, "<>2024-25")
            End Sub
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
'###################################################################################################################################
    'call Rut_Lo_Sort_By2Cols( LstObj , Columna1 , Columna2 , xlAscending/xlDescending ,  True/False )
Sub Rut_Lo_Sort_By2Cols(ByRef LstObj As ListObject, Columna1 As Integer, Columna2 As Integer, VarOrden As Long, Optional SW_Clear As Boolean = False)
    
    With LstObj.Sort
        If SW_Clear Then .SortFields.Clear
        .SortFields.Add Key:=LstObj.ListColumns(Columna1).Range, SortOn:=xlSortOnValues, Order:=VarOrden, DataOption:=xlSortNormal
        .SortFields.Add Key:=LstObj.ListColumns(Columna2).Range, SortOn:=xlSortOnValues, Order:=VarOrden, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

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
Sub Rut_Lo_Filtro(ByRef Lo_Tb As ListObject, Columna As Integer, Criterio As String, Optional SW_Clear As Boolean = False)
' ----------------------------------------------------------------------------------------------------------------------------------
    If SW_Clear And Not Lo_Tb.AutoFilter Is Nothing Then Lo_Tb.AutoFilter.ShowAllData
    Lo_Tb.Range.AutoFilter Field:=Columna, Criteria1:=Criterio
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

                            '#######################################################################################################
                            Sub Rut_Filtros_Quitar_ActivSheet_LstObj()      ' Muestra Todas las Solicitudes y Activar Filtros >>>>>>
                                    Call Rut_Lo_Filtros_Quitar(ActiveSheet.ListObjects(1))
                            End Sub     ' ---------------------------------------------------------------------------------------<<<
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

Sub Rut_Copiar_EntireRow_LstObjct() '- Rut_Añade_Row_entera, DataBodyRange.Rows(2) ES OBLIGATORIO sino no funciona
                
ActiveSheet.ListObjects("Tab_INI").DataBodyRange.Rows(2).Copy ActiveSheet.ListObjects("Tab_FIN").ListRows(3).Range '
                
End Sub






