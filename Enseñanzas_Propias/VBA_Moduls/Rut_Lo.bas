Attribute VB_Name = "Rut_Lo"
' Last Rev. 2026-09-27 23:42
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
    
    '- Pega datos justo debajo de la Lo_Target ---------------------------------------------------
    '-  OJO (bug corregido 2026-09-23): NO usar
    '-      Lo_Target.DataBodyRange.Offset(DataBodyRange.Rows.Count, 0)
    '-  como destino. Con la tabla en 1 registro (fila 5) y Totales en la 6, ese Offset apunta
    '-  justo A LA FILA DE TOTALES, y el Resize de mas abajo dejaba la tabla una fila corta: lo
    '-  copiado quedaba FUERA del ListObject (visto en BD_Deleted, con miles de filas huerfanas).
    '-  Se calcula todo con COORDENADAS NUMERICAS, ya con los Totales apagados.
    Dim FilCabecera     As Long:    FilCabecera = Lo_Target.HeaderRowRange.Row
    Dim ColIniTgt       As Long:    ColIniTgt = Lo_Target.Range.Column
    Dim ColFinTgt       As Long:    ColFinTgt = ColIniTgt + Lo_Target.Range.Columns.Count - 1
    Dim FilasTgt        As Long
    Dim StartRowAdd     As Long

    If Lo_Target.DataBodyRange Is Nothing Then
        FilasTgt = 0
    Else
        FilasTgt = Lo_Target.ListRows.Count
    End If
    StartRowAdd = FilCabecera + FilasTgt + 1        '- 1a fila libre de datos tras la tabla

    RangoACopiar.Copy Destination:=ws.Cells(StartRowAdd, ColIniTgt)

    '- Como copio un Rango, Lo_Target NO se expande sola: hay que redimensionarla para que
    '-  absorba las filas nuevas.
    Dim RowsACopiar     As Long:    RowsACopiar = RangoACopiar.Rows.Count
    Lo_Target.Resize ws.Range(ws.Cells(FilCabecera, ColIniTgt), _
                              ws.Cells(StartRowAdd + RowsACopiar - 1, ColFinTgt))
    
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
        .ListObjects(1).ShowTotals = False          ' Ocultar Fila de Totales
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
' Cuenta las filas VISIBLES del cuerpo de una tabla, leyendo por la columna indicada.
'   Sustituye al patron antiguo, repetido en 45 sitios del pipeline:
'       rowfind = Lo.Range.Columns(Col).SpecialCells(xlCellTypeVisible).Cells.Count - 1
'   que contaba sobre .Range (CON la cabecera) y compensaba con un -1 a mano. Ese -1 solo era
'   correcto con ShowTotals = False; con la fila de totales visible contaba 1 de mas.
'   Contando sobre DataBodyRange desaparecen las tres dependencias ocultas del patron viejo:
'     1) el -1 magico por la cabecera,
'     2) el estado de ShowTotals,
'     3) que la columna leida estuviera VISIBLE (el comentario "OJO, TIENE QUE ESTAR VISIBLE LA
'        COLUMNA..." que acompanaba a cada linea): si esta oculta, SpecialCells da error 1004;
'        aqui se captura y devuelve 0 en vez de propagar un conteo erroneo.
'   Devuelve 0 (sin error) si la tabla esta vacia, si la columna esta fuera de rango o si no
'   hay ninguna fila visible.
'   Columna va ByVal a proposito: todas las llamadas pasan constantes Integer (G04_*, BD_*) y
'   un parametro ByRef As Long no admite un Integer (error de compilacion). ByVal lo promociona.
Function Fnc_Lo_Contar_Visibles(Lo_Data As ListObject, ByVal Columna As Long) As Long
' --------------------------------------------------------------------------------------------------
    Dim Rg_Visibles     As Range
    Fnc_Lo_Contar_Visibles = 0
    If Lo_Data Is Nothing Then Exit Function
    If Lo_Data.DataBodyRange Is Nothing Then Exit Function      '- Tabla sin cuerpo
    If Columna < 1 Or Columna > Lo_Data.ListColumns.Count Then Exit Function
    On Error Resume Next        '- Sin filas visibles (o columna oculta) -> SpecialCells da error 1004
    Set Rg_Visibles = Lo_Data.DataBodyRange.Columns(Columna).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not Rg_Visibles Is Nothing Then Fnc_Lo_Contar_Visibles = Rg_Visibles.Cells.Count
End Function
' --------------------------------------------------------------------------------------------------


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
        RowsFind = Fnc_Lo_Contar_Visibles(Lo_Data, ColSearch1)
        If RowsFind > 0 Then .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete          '- Borrar Filas visibles
    End With
    Call Rut_Lo_Filtros_Quitar(Lo_Data)                '- Quitar filtros
    Lo_Data.ShowTotals = Sw_ShowTotals
Debug.Print "Rut_Lo_DataBodyRange_Filter_y_DEL: ColSearch1: " & ColSearch1 & ", Criterio1: _" & Criterio1 & ", Criterio2: _" & Criterio2 & "_ RowsFind: _" & RowsFind & " reg."
End Sub
'###################################################################################################




