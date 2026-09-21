Attribute VB_Name = "Rut_Lo"
' Last Rev. 2026-09-21 12:12
Option Explicit


'###################################################################################################
' Copia el DataBodyRange de TabSource en TabTarget,      Opcional: Borrar primero contenido de TabTarget
    Sub Rut_Lo_DataBodyRange_Copy(Lo_Source As ListObject, _
                              Lo_Target As ListObject, _
                              Optional DelFirstLoTarget As Boolean = False)
    
    If DelFirstLoTarget And Not Lo_Target.DataBodyRange Is Nothing Then Lo_Target.DataBodyRange.Delete
    If Lo_Source.DataBodyRange Is Nothing Then MsgBox "Tabla de Origen Sin Datos": Exit Sub
    Lo_Source.DataBodyRange.Copy
    If Lo_Target.DataBodyRange Is Nothing Then
        Lo_Target.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    Else
        Lo_Target.DataBodyRange.Offset(Lo_Target.DataBodyRange.Rows.Count, 0).PasteSpecial Paste:=xlPasteValues
    End If
    Application.CutCopyMode = False
End Sub
' --------------------------------------------------------------------------------------------------
'###################################################################################################
' Copia el DataBodyRange FILTRADO de Lo_Source en Lo_Target,      Opcional: Borrar primero contenido de Lo_Target y Borrar los registros filtrados de Lo_Source
Sub Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Source As ListObject, _
                                       Lo_Target As ListObject, _
                                       Optional DelFirstLoTarget As Boolean = False, _
                                       Optional DelLoSourceFilteredRows As Boolean = False)
' --------------------------------------------------------------------------------------------------
Debug.Print "Rut_Lo_DataBodyRange_Filtered_Copy"
    Dim ws As Worksheet
    Set ws = Lo_Target.Parent
    Dim Sw_ShowTotals       As Boolean:     Sw_ShowTotals = Lo_Target.ShowTotals:               Lo_Target.ShowTotals = False
    Dim Sw_DisplayAlerts    As Boolean:     Sw_DisplayAlerts = Application.DisplayAlerts:       Application.DisplayAlerts = False
    
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
    Set NuevoRangoAmpliado = ws.Range(Lo_Target.Range.Cells(1, 1), _
                             ws.Cells(StartRowAdd + RowsACopiar - 1, Lo_Target.Range.Column + Lo_Target.Range.Columns.Count - 1))
    Lo_Target.Resize NuevoRangoAmpliado   '- Redefino Lo_Target con el tamaño de Lo_Target más el Rango Copiado: NuevoRangoAmpliado
    
    If DelLoSourceFilteredRows Then RangoACopiar.Delete     '- Como estamos dentro del "IF Not RangoACopiar Is Nothing" el Rango tiene datos y los podemos Borrar
    Application.CutCopyMode = False
    Lo_Target.ShowTotals = Sw_ShowTotals
    Application.DisplayAlerts = Sw_DisplayAlerts
End Sub
' --------------------------------------------------------------------------------------------------
'###################################################################################################
'###################################################################################################
' Copia el DataBodyRange FILTRADO de Lo_Source en Lo_Target,      Opcional: Borrar primero contenido de Lo_Target y Borrar los registros filtrados de Lo_Source
' --------------------------------------------------------------------------------------------------
'###################################################################################################
Sub Rut_Lo_WrkSht_Preparar(WrkSht As Worksheet)
' --------------------------------------------------------------------------------------------------
    With WrkSht
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        Call Rut_Lo_Filtros_Quitar(.ListObjects(1))
    End With
End Sub
' --------------------------------------------------------------------------------------------------

'###################################################################################################
Sub Rut_Lo_Sort(ByRef Lo_Tb As ListObject, columna As Integer, VarOrden As String, Optional SW_Clear As Boolean = False)
' --------------------------------------------------------------------------------------------------
    With Lo_Tb.Sort
        If SW_Clear Then .sortFields.Clear
        .sortFields.Add Key:=Lo_Tb.ListColumns(columna).Range, SortOn:=xlSortOnValues, Order:=VarOrden, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
' --------------------------------------------------------------------------------------------------

                            '#######################################################################
                            Sub Rut_Filtros_Quitar_ActivSheet_LstObj()      ' Muestra Todas las Solicitudes y Activar Filtros
                                    Call Rut_Lo_Filtros_Quitar(ActiveSheet.ListObjects(1))
                            End Sub     ' ----------------------------------------------------------
'###################################################################################################
Sub Rut_Lo_Filtros_Quitar(ByRef Lo_Tb As ListObject)      ' Muestra Todas las Solicitudes y Activar Filtros
' --------------------------------------------------------------------------------------------------
    With Lo_Tb
        If .ShowAutoFilter Then                         '--- Compruebo que la Tabla tiene los Filtros Activos
            With .AutoFilter
                 If .FilterMode Then .ShowAllData
            End With
        Else                                            '--- Si NO tiene los Filtros Activos, los Activo
            .ShowAutoFilter = True
        End If
    End With
End Sub


'###################################################################################################
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
'###################################################################################################




