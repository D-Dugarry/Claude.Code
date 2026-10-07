Attribute VB_Name = "Rut_Lo"
' Last Rev. 2026-10-07 10:49
Option Explicit

'- Hojas (por CodeName) cuyas tablas quedan con la fila de totales visible cuando NO tienen fila en Lo_ListObjAPP, ademas de todas
'- las de las hojas Prog_DefCol_*. La aplica Rut_Lo_Totales_Mostrar (Docs/Plan_ShowTotals.md, fases 4 y 6).
Private Const Lo_Totales_Hojas     As String = "Sht__BD,Sht__Inf_Recibos_TIO"

'###################################################################################################################################
Sub Rut_Lo_Columns_Show_Hide(WrkSht As Worksheet, WsDefCol As Worksheet, Col_HiddenSw As Integer, Optional Reset As Boolean = False)
Debug.Print "Rut_Lo_Columns_Show_Hide"
    Dim Lo_DefCol       As ListObject:      Set Lo_DefCol = WsDefCol.ListObjects(1)
    Dim Lo_Table        As ListObject:      Set Lo_Table = WrkSht.ListObjects(1)
    Dim Cont_Col        As Integer
    Dim ColEnBlco       As Integer:         ColEnBlco = Lo_Table.Range.Columns(1).Column - 1    '- Por si hay columnas en blanco a la derecha de la ListObject.
    Dim HiddenCol       As Boolean
    Dim SetHidden       As Variant
    Dim SW_Col_Hide     As Boolean:         SW_Col_Hide = Fnc_ColHide_Get(WrkSht.CodeName)
    'Rango completo de columnas afectadas
    Dim RngCols         As Range:           Set RngCols = WrkSht.Range(WrkSht.Columns(ColEnBlco + 1), WrkSht.Columns(ColEnBlco + Lo_Table.ListColumns.Count))

    If Reset Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        WrkSht.Columns.Hidden = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim Calc_Prev       As XlCalculation:   Calc_Prev = Application.Calculation
    Dim Events_Prev     As Boolean:         Events_Prev = Application.EnableEvents
    Dim Rng_Ocultar     As Range
    Dim Rng_Mostrar     As Range
    Dim Err_Num         As Long
    Dim Err_Desc        As String
    On Error GoTo Restaurar
    Application.Calculation = xlCalculationManual       '- Sin recalculos ni eventos mientras se cambia el ancho de las columnas
    Application.EnableEvents = False

    If SW_Col_Hide = True Then
        'Un solo disparo para mostrar todas
        RngCols.EntireColumn.Hidden = False
    Else
        'Leer de una vez la columna de definicion en una matriz
        SetHidden = Lo_DefCol.DataBodyRange.Columns(Col_HiddenSw).Value2
        'Agrupar las columnas a ocultar y las a mostrar, para aplicar Hidden dos veces en lugar de una por columna
        '- La DefCol puede tener menos filas que columnas tiene la tabla (BDatos: 54 filas, 64 columnas): las que no tienen fila no se tocan.
        For Cont_Col = 1 To Application.Min(Lo_Table.ListColumns.Count, UBound(SetHidden, 1))
            HiddenCol = SetHidden(Cont_Col, 1)
            If WrkSht.Columns(ColEnBlco + Cont_Col).Hidden = HiddenCol Then GoTo SigCol    '- Ya esta como debe: no se toca (mostrar es lo caro)
            If HiddenCol Then
                If Rng_Ocultar Is Nothing Then Set Rng_Ocultar = WrkSht.Columns(ColEnBlco + Cont_Col) Else Set Rng_Ocultar = Union(Rng_Ocultar, WrkSht.Columns(ColEnBlco + Cont_Col))
            Else
                If Rng_Mostrar Is Nothing Then Set Rng_Mostrar = WrkSht.Columns(ColEnBlco + Cont_Col) Else Set Rng_Mostrar = Union(Rng_Mostrar, WrkSht.Columns(ColEnBlco + Cont_Col))
            End If
SigCol:
        Next Cont_Col
        If Not Rng_Mostrar Is Nothing Then Rng_Mostrar.EntireColumn.Hidden = False
        If Not Rng_Ocultar Is Nothing Then Rng_Ocultar.EntireColumn.Hidden = True
    End If

    'Actualizar el switch solo una vez. El flag que se guardaba en la fila de totales de la DefCol no lo leia nadie (decide
    'el switch Sw_Col_Hide de Lo_ListObjAPP): se quito el 2026-10-04 (Docs/Plan_ShowTotals.md, fase 1).
    Lo_DefCol.ShowTotals = True
    Call Rut_ColHide_Set(WrkSht.CodeName, Not SW_Col_Hide)

Restaurar:
    Err_Num = Err.Number:   Err_Desc = Err.Description
    On Error Resume Next
    Application.EnableEvents = Events_Prev
    Application.Calculation = Calc_Prev
    On Error GoTo 0
    If Err_Num <> 0 Then Err.Raise Err_Num, "Rut_Lo_Columns_Show_Hide", Err_Desc

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
    
    '- Sin la fila de totales mientras se pega debajo de la tabla; al salir, como estaba (Docs/Plan_ShowTotals.md, fase 2).
    Dim Totales_Visibles    As Boolean:     Totales_Visibles = Fnc_Lo_Totales_Ocultar(Lo_Target)
    If DelFirstLoTarget And Not Lo_Target.DataBodyRange Is Nothing Then Lo_Target.DataBodyRange.Delete
    
    Dim RangoACopiar    As Range
    On Error Resume Next
    Set RangoACopiar = Lo_Source.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If RangoACopiar Is Nothing Then
        Call Rut_Lo_Totales_Restaurar(Lo_Target, Totales_Visibles)
        Exit Sub
    End If
    
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
    Call Rut_Lo_Totales_Restaurar(Lo_Target, Totales_Visibles)
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
    Dim Sw_ShowTotals    As Boolean:    Sw_ShowTotals = Fnc_Lo_Totales_Ocultar(Lo_Data)
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
        RowsFind = Fnc_Lo_Contar_Visibles(Lo_Data, ColSearch1)
        If RowsFind > 0 Then .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete          '- Borrar Filas visibles
    End With
    Call Rut_Lo_Filtros_Quitar(Lo_Data)                '- Quitar filtros
    Call Rut_Lo_Totales_Restaurar(Lo_Data, Sw_ShowTotals)
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
    Dim Sw_ShowTotals    As Boolean:    Sw_ShowTotals = Fnc_Lo_Totales_Ocultar(Lo_Data)
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
        RowsFind = Fnc_Lo_Contar_Visibles(Lo_Data, ColSource)
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
    Call Rut_Lo_Totales_Restaurar(Lo_Data, Sw_ShowTotals)
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
        Call Rut_ColHide_Set(WrkSht.CodeName, False)
        .Rows.EntireRow.Hidden = False              '-2º Mostrar todas las Filas
        Call Rut_Lo_Filtros_Quitar(.ListObjects(1)) '-3º Quitar Filtros
        If .ListObjects(1).ShowTotals Then .ListObjects(1).ShowTotals = False   '-4º Ocultar la Fila de Totales, solo si se ve (Docs/Plan_ShowTotals.md, fase 5)
'        .Protect allowFiltering:=True, DrawingObjects:=True, allowSorting:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
    End With
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
'###################################################################################################################################
' Fila de totales mientras se trabaja con una tabla (Docs/Plan_ShowTotals.md, fase 2; patrón de JyC tras el fallo de Lo_TPV):
'   Dim Totales_Visibles As Boolean:   Totales_Visibles = Fnc_Lo_Totales_Ocultar(Lo)
'   ... vaciar, pegar, borrar filas ...
'   Call Rut_Lo_Totales_Restaurar(Lo, Totales_Visibles)       '- en TODAS las salidas
' Pegar debajo de la cabecera con la fila de totales visible la pisa y deja los datos fuera de la tabla. Las dos solo cambian
' ShowTotals si hace falta: tras M_110/M_210, cambiarlo en Tb_INSS da -2147417848 (ver CLAUDE.md).
Function Fnc_Lo_Totales_Ocultar(ByVal Lo As ListObject) As Boolean
' ----------------------------------------------------------------------------------------------------------------------------------
    Fnc_Lo_Totales_Ocultar = Lo.ShowTotals
    If Lo.ShowTotals Then Lo.ShowTotals = False
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Totales_Restaurar(ByVal Lo As ListObject, ByVal Visibles As Boolean)
' ----------------------------------------------------------------------------------------------------------------------------------
    If Lo.ShowTotals <> Visibles Then Lo.ShowTotals = Visibles
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
' Cuenta las filas VISIBLES del cuerpo de una tabla, leyendo por la columna indicada (Docs/Plan_ShowTotals.md, fase 3).
'   Sustituye al patrón repetido en 57 sitios:   n = Lo.Range.Columns(Col).SpecialCells(xlCellTypeVisible).Cells.Count - 1
'   que restaba la cabecera a mano y solo acertaba con la fila de totales oculta (con ella visible contaba 1 de más; M_130
'   restaba 2 porque allí se ve). Aquí se descuenta la fila de totales si se ve, así que da igual cómo esté.
'   Sigue contando sobre .Range, como JyC (Count - 1 + .ShowTotals), y no sobre DataBodyRange como la versión de EP, que
'   devuelve 0 si SpecialCells falla. En PPub hay un fallo de Excel conocido que lo hace fallar (Tb_INSS, ver CLAUDE.md):
'   con 0, M_111 o M_311 se saltarían los borrados sin avisar. Aquí el error salta, igual que antes.
'   Tabla sin filas: 0 (el patrón viejo contaba 1, la fila de inserción). La columna tiene que estar VISIBLE, como antes.
'   Columna va ByVal: las llamadas pasan constantes Integer (BD_*, LS06_*...) y un ByRef As Long no las admite.
Function Fnc_Lo_Contar_Visibles(ByVal Lo_Data As ListObject, ByVal Columna As Long) As Long
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim N_Vis       As Long
    If Lo_Data.DataBodyRange Is Nothing Then Exit Function
    N_Vis = Lo_Data.Range.Columns(Columna).SpecialCells(xlCellTypeVisible).Cells.Count - 1      '- menos la cabecera
    If Lo_Data.ShowTotals Then
        If Not Lo_Data.TotalsRowRange.EntireRow.Hidden Then N_Vis = N_Vis - 1                   '- menos la fila de totales
    End If
    Fnc_Lo_Contar_Visibles = N_Vis
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
' Politica de la fila de totales (Docs/Plan_ShowTotals.md, fases 4 y 6): al terminar cada tarea y al abrir el libro, cada tabla queda
' como diga la columna Sw_TRow_Hide de Lo_ListObjAPP (hoja ListObjAPP, CodeName Prog__APP_ListObj; una fila por tabla):
'     False (o "No")  = totales VISIBLES      True (o "Si") = totales OCULTOS      vacio = LIBRE (no se toca)
' Las rutinas pueden ocultarla o mostrarla para trabajar; aqui se deja como esta configurado. Se llama donde ya esta la red de
' seguridad de Rut_Reset_State: Rut_Progreso_Cerrar (M_90), el final de tarea de Form_Menu y del menu dinamico (OnAction_Dynamic_Task,
' M___RibbonUI) y RuT_Al_Abrir_WorkBook (M_000_Ini_APP).
' Una tabla que NO tiene fila en Lo_ListObjAPP (o si la hoja no existe) se rige por la regla de antes de la fase 6: las de las hojas
' Prog_DefCol_* y las de Lo_Totales_Hojas, visibles; el resto, libres. Rut_Lo_ListObjAPP_Capturar da de alta las tablas que falten.
Public Sub Rut_Lo_Totales_Mostrar()
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim Ws      As Worksheet
    Dim Lo      As ListObject
    Dim Estado  As Long
    For Each Ws In ThisWorkbook.Worksheets
        For Each Lo In Ws.ListObjects
            Estado = Fnc_Lo_Totales_Estado(Lo)
            If Estado = 1 And Not Lo.ShowTotals Then
                Call Rut_Lo_Totales_Fijar_Lo(Lo, True)
            ElseIf Estado = 0 And Lo.ShowTotals Then
                Call Rut_Lo_Totales_Fijar_Lo(Lo, False)
            End If
        Next Lo
    Next Ws
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
'- Estado configurado de la fila de totales de una tabla: 1 = visible, 0 = oculta, -1 = libre.
Private Function Fnc_Lo_Totales_Estado(ByVal Lo As ListObject) As Long
    Dim Cfg     As ListObject
    Dim Fila    As Variant
    Dim V       As Variant
    Dim S       As String
    Fnc_Lo_Totales_Estado = Fnc_Lo_Totales_Estado_Regla(Lo.Parent)      '- por defecto, la regla de antes de la fase 6
    Set Cfg = Fnc_Lo_ListObjAPP()
    If Cfg Is Nothing Then Exit Function
    If Cfg.DataBodyRange Is Nothing Then Exit Function
    Fila = Application.Match(Lo.Name, Cfg.ListColumns("Nom_ListObj").DataBodyRange, 0)
    If IsError(Fila) Then Exit Function                                 '- tabla sin fila: vale la regla
    V = Cfg.ListColumns("Sw_TRow_Hide").DataBodyRange.Cells(Fila, 1).Value
    Fnc_Lo_Totales_Estado = -1                                          '- con fila, solo se impone lo que diga la celda
    If IsError(V) Or IsEmpty(V) Then Exit Function
    If VarType(V) = vbBoolean Then
        Fnc_Lo_Totales_Estado = IIf(V, 0, 1)
        Exit Function
    End If
    S = LCase$(Trim$(CStr(V)))
    Select Case S
        Case "true", "verdadero", "si", "s" & ChrW(237), "1", "-1": Fnc_Lo_Totales_Estado = 0
        Case "false", "falso", "no", "0":                           Fnc_Lo_Totales_Estado = 1
    End Select
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'- Regla de la fase 4 (antes de Lo_ListObjAPP): 1 = visible para las DefCol y las hojas de Lo_Totales_Hojas, -1 = libre el resto.
Private Function Fnc_Lo_Totales_Estado_Regla(ByVal Ws As Worksheet) As Long
    If (Ws.CodeName Like "Prog_DefCol_*") _
    Or (InStr(1, "," & Lo_Totales_Hojas & ",", "," & Ws.CodeName & ",", vbTextCompare) > 0) Then
        Fnc_Lo_Totales_Estado_Regla = 1
    Else
        Fnc_Lo_Totales_Estado_Regla = -1
    End If
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'- La tabla de configuracion (Lo_ListObjAPP) o Nothing si no existe. Se busca por el CodeName de la hoja, sin nombrarlo en el codigo,
'- para que el modulo compile aunque la hoja todavia no este en el libro.
Private Function Fnc_Lo_ListObjAPP() As ListObject
    Dim Ws      As Worksheet
    For Each Ws In ThisWorkbook.Worksheets
        If Ws.CodeName = "Prog__APP_ListObj" Then
            On Error Resume Next
            Set Fnc_Lo_ListObjAPP = Ws.ListObjects("Lo_ListObjAPP")
            On Error GoTo 0
            Exit Function
        End If
    Next Ws
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'- Muestra u oculta la fila de totales de una tabla, desprotegiendo su hoja si hace falta y dejandola con los mismos permisos. Un
'- fallo (hoja con contrasena, fila de debajo ocupada...) no corta la politica: se anota en Inmediato y sigue con las demas.
Private Sub Rut_Lo_Totales_Fijar_Lo(ByVal Lo As ListObject, ByVal Visible As Boolean)
    Dim Est     As T_Prot_Estado
    On Error Resume Next
    Call Rut_Prot_Save(Lo.Parent, Est)
    Lo.ShowTotals = Visible
    If Err.Number <> 0 Then Debug.Print "!!! Rut_Lo_Totales_Mostrar: " & Lo.Name & " (" & Lo.Parent.Name & "): " & Err.Description
    Err.Clear
    Call Rut_Prot_Restore(Lo.Parent, Est)
    On Error GoTo 0
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
' Captura de las tablas del libro en Lo_ListObjAPP (Docs/Plan_ShowTotals.md, fase 6). Se lanza a mano desde Inmediato:
'   Call Rut_Lo_ListObjAPP_Capturar
' Anade una fila por cada tabla del libro que todavia no este (Nom_ListObj = nombre de la tabla, CodeName_Sheet y Name_Sheet = CodeName y nombre de su hoja) y
' NO toca las filas que ya existen. En las nuevas propone Sw_TRow_Hide = False (visible) si la regla de la fase 4 la daba por visible,
' y lo deja vacio (libre) en el resto. Avisa en Inmediato de las filas cuya tabla ya no existe, sin borrarlas.
Public Sub Rut_Lo_ListObjAPP_Capturar()
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim Cfg     As ListObject:      Set Cfg = Fnc_Lo_ListObjAPP()
    Dim Est     As T_Prot_Estado
    Dim Ws      As Worksheet
    Dim Lo      As ListObject
    Dim Fila    As ListRow
    Dim V       As Variant
    Dim N_Alta  As Long
    Dim N_Total As Long
    Dim i       As Long
    If Cfg Is Nothing Then
        MsgBox "No se encuentra la tabla Lo_ListObjAPP en la hoja de CodeName Prog__APP_ListObj.", vbExclamation
        Exit Sub
    End If
    Call Rut_Prot_Save(Cfg.Parent, Est)
    For Each Ws In ThisWorkbook.Worksheets
        For Each Lo In Ws.ListObjects
            N_Total = N_Total + 1
            If Cfg.DataBodyRange Is Nothing Then
                V = CVErr(xlErrNA)
            Else
                V = Application.Match(Lo.Name, Cfg.ListColumns("Nom_ListObj").DataBodyRange, 0)
            End If
            If IsError(V) Then
                Set Fila = Nothing                                              '- primero se aprovecha una fila vacia (la de insercion)
                If Not Cfg.DataBodyRange Is Nothing Then
                    If Cfg.ListRows.Count = 1 Then
                        If Trim$(CStr(Cfg.ListColumns("Nom_ListObj").DataBodyRange.Cells(1, 1).Value)) = "" Then Set Fila = Cfg.ListRows(1)
                    End If
                End If
                If Fila Is Nothing Then Set Fila = Cfg.ListRows.Add
                Fila.Range.Cells(1, Cfg.ListColumns("Nom_ListObj").Index).Value = Lo.Name
                Fila.Range.Cells(1, Cfg.ListColumns("CodeName_Sheet").Index).Value = Ws.CodeName
                Fila.Range.Cells(1, Cfg.ListColumns("Name_Sheet").Index).Value = Ws.Name
                If Fnc_Lo_Totales_Estado_Regla(Ws) = 1 Then
                    Fila.Range.Cells(1, Cfg.ListColumns("Sw_TRow_Hide").Index).Value = False
                    Fila.Range.Cells(1, Cfg.ListColumns.Count).Value = "Propuesto por la regla de la fase 4: totales visibles"
                End If
                N_Alta = N_Alta + 1
            End If
        Next Lo
    Next Ws
    If Not Cfg.DataBodyRange Is Nothing Then                                    '- filas cuya tabla ya no existe
        For i = 1 To Cfg.ListRows.Count
            V = Cfg.ListColumns("Nom_ListObj").DataBodyRange.Cells(i, 1).Value
            If Not Fnc_Lo_Existe(CStr(V)) Then Debug.Print "!!! Rut_Lo_ListObjAPP_Capturar: la fila " & i & " (" & V & ") no corresponde a ninguna tabla del libro"
        Next i
    End If
    Call Rut_Prot_Restore(Cfg.Parent, Est, True)
    Debug.Print "Rut_Lo_ListObjAPP_Capturar: " & N_Total & " tablas en el libro, " & N_Alta & " dadas de alta"
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
Private Function Fnc_Lo_Existe(ByVal Nombre As String) As Boolean
    Dim Ws      As Worksheet
    Dim Lo      As ListObject
    For Each Ws In ThisWorkbook.Worksheets
        For Each Lo In Ws.ListObjects
            If StrComp(Lo.Name, Nombre, vbTextCompare) = 0 Then Fnc_Lo_Existe = True: Exit Function
        Next Lo
    Next Ws
End Function
' ----------------------------------------------------------------------------------------------------------------------------------

'###################################################################################################################################
Sub Rut_Lo_Sort(ByRef Lo_Tb As ListObject, Columna As Integer, VarOrden As String, Optional SW_Clear As Boolean = False, _
                Optional Aplicar As Boolean = True)
' ----------------------------------------------------------------------------------------------------------------------------------
'- Aplicar:=False solo añade la Col. a los criterios de ordenación, sin ordenar todavía. En una tanda de ordenaciones seguidas
'- de la misma tabla basta con ordenar en la última, que ya lleva todas las Col.: la tabla se ordena una vez en lugar de una
'- por Col., con el mismo resultado (la ordenación de Excel es estable).
    With Lo_Tb.Sort
        If SW_Clear Then .SortFields.Clear
        .SortFields.Add Key:=Lo_Tb.ListColumns(Columna).Range, SortOn:=xlSortOnValues, Order:=VarOrden, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        If Aplicar Then .Apply
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






' ----------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
' Switch "ocultar columnas" de cada hoja: columna Sw_Col_Hide de Lo_ListObjAPP (hoja ListObjAPP), en la fila de la hoja (CodeName_Sheet).
' Sustituye a los nombres Sw_Col_Hide_<CodeName> de Lo_SwitchsAPP. Reglas:
'     - Celda vacia (o la hoja sin fila) = la hoja NO tiene switch: Fnc_ColHide_Existe da False, Get da False y Set no escribe nada.
'     - Si una hoja tiene varias filas (varias tablas), vale la PRIMERA; en esas hojas no se usa el switch.
'     - Las filas las da de alta Rut_Lo_ListObjAPP_Capturar; Rut_ColHide_Migrar_Desde_Switchs copia los valores de Lo_SwitchsAPP.
Private Function Fnc_ColHide_Celda(ByVal CodeName As String) As Range
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim Cfg     As ListObject:      Set Cfg = Fnc_Lo_ListObjAPP()
    Dim Fila    As Variant
    If Cfg Is Nothing Then Exit Function
    If Cfg.DataBodyRange Is Nothing Then Exit Function
    Fila = Application.Match(CodeName, Cfg.ListColumns("CodeName_Sheet").DataBodyRange, 0)
    If IsError(Fila) Then Exit Function
    Set Fnc_ColHide_Celda = Cfg.ListColumns("Sw_Col_Hide").DataBodyRange.Cells(Fila, 1)
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'- Convierte el contenido de una celda de switch en Boolean (Boolean, o texto True/Verdadero/Si/1; vacio y error = False).
Private Function Fnc_Valor_Booleano(ByVal V As Variant) As Boolean
    If IsError(V) Or IsEmpty(V) Then Exit Function
    If VarType(V) = vbBoolean Then
        Fnc_Valor_Booleano = V
        Exit Function
    End If
    Select Case LCase$(Trim$(CStr(V)))
        Case "true", "verdadero", "si", "s" & ChrW(237), "1", "-1": Fnc_Valor_Booleano = True
    End Select
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'- True si la hoja tiene switch de ocultar columnas (fila en Lo_ListObjAPP con Sw_Col_Hide relleno).
Public Function Fnc_ColHide_Existe(ByVal CodeName As String) As Boolean
    Dim Celda   As Range:           Set Celda = Fnc_ColHide_Celda(CodeName)
    If Celda Is Nothing Then Exit Function
    If IsError(Celda.Value2) Then Exit Function
    Fnc_ColHide_Existe = (Trim$(CStr(Celda.Value2)) <> "")
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'- Valor del switch de la hoja (False si no lo tiene).
Public Function Fnc_ColHide_Get(ByVal CodeName As String) As Boolean
    If Not Fnc_ColHide_Existe(CodeName) Then Exit Function
    Fnc_ColHide_Get = Fnc_Valor_Booleano(Fnc_ColHide_Celda(CodeName).Value2)
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'- Fija el valor del switch de la hoja. Si la hoja no tiene switch, no hace nada. Escribe con Rut_Prot_Save/Restore (la hoja ListObjAPP
'- puede estar protegida) y no deja la proteccion quitada si la escritura falla.
Public Sub Rut_ColHide_Set(ByVal CodeName As String, ByVal Valor As Boolean)
    Dim Celda   As Range
    Dim Est     As T_Prot_Estado
    If Not Fnc_ColHide_Existe(CodeName) Then Exit Sub
    Set Celda = Fnc_ColHide_Celda(CodeName)
    Call Rut_Prot_Save(Celda.Worksheet, Est)
    On Error Resume Next
    Celda.Value = Valor
    If Err.Number <> 0 Then Debug.Print "!!! Rut_ColHide_Set: " & CodeName & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    Call Rut_Prot_Restore(Celda.Worksheet, Est, True)
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
' Migracion de un solo uso (se lanza a mano desde Inmediato, DESPUES de Rut_Lo_ListObjAPP_Capturar):
'   Call Rut_ColHide_Migrar_Desde_Switchs
' Copia el valor de cada fila Sw_Col_Hide_<CodeName> de Lo_SwitchsAPP a la columna Sw_Col_Hide de Lo_ListObjAPP (fila de esa hoja). No borra
' nada de Lo_SwitchsAPP. Avisa en Inmediato de las hojas que no existen y de las que no tienen fila en Lo_ListObjAPP.
Public Sub Rut_ColHide_Migrar_Desde_Switchs()
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim Cfg     As ListObject:      Set Cfg = Fnc_Lo_ListObjAPP()
    Dim Sw      As ListObject
    Dim Est     As T_Prot_Estado
    Dim Ws      As Worksheet
    Dim Hoja    As Worksheet
    Dim Fila    As Variant
    Dim V       As Variant
    Dim Nom     As String
    Dim Code    As String
    Dim i       As Long
    Dim N_Ok    As Long
    Dim N_Aviso As Long
    If Cfg Is Nothing Then
        MsgBox "No se encuentra la tabla Lo_ListObjAPP en la hoja de CodeName Prog__APP_ListObj.", vbExclamation
        Exit Sub
    End If
    If Cfg.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    Set Sw = Prog__APP_Switch.ListObjects("Lo_SwitchsAPP")
    On Error GoTo 0
    If Sw Is Nothing Then
        MsgBox "No se encuentra la tabla Lo_SwitchsAPP en la hoja de CodeName Prog__APP_Switch.", vbExclamation
        Exit Sub
    End If
    Call Rut_Prot_Save(Cfg.Parent, Est)
    For i = 1 To Sw.ListRows.Count
        Nom = Trim$(CStr(Sw.ListColumns("NombreRango").DataBodyRange.Cells(i, 1).Value))
        If LCase$(Left$(Nom, 12)) = "sw_col_hide_" Then
            Code = Mid$(Nom, 13)
            Set Hoja = Nothing
            For Each Ws In ThisWorkbook.Worksheets
                If StrComp(Ws.CodeName, Code, vbTextCompare) = 0 Then Set Hoja = Ws: Exit For
            Next Ws
            Fila = Application.Match(Code, Cfg.ListColumns("CodeName_Sheet").DataBodyRange, 0)
            If Hoja Is Nothing Then
                Debug.Print "!!! Rut_ColHide_Migrar: " & Nom & ": no hay ninguna hoja con CodeName " & Code & ", no se migra"
                N_Aviso = N_Aviso + 1
            ElseIf IsError(Fila) Then
                Debug.Print "!!! Rut_ColHide_Migrar: " & Nom & ": la hoja " & Code & " no tiene fila en Lo_ListObjAPP (lanza antes Rut_Lo_ListObjAPP_Capturar)"
                N_Aviso = N_Aviso + 1
            Else
                V = Sw.ListColumns("Switch").DataBodyRange.Cells(i, 1).Value2
                Cfg.ListColumns("Sw_Col_Hide").DataBodyRange.Cells(Fila, 1).Value = Fnc_Valor_Booleano(V)
                N_Ok = N_Ok + 1
            End If
        End If
    Next i
    Call Rut_Prot_Restore(Cfg.Parent, Est, True)
    Debug.Print "Rut_ColHide_Migrar_Desde_Switchs: " & N_Ok & " migrados, " & N_Aviso & " avisos"
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
'###################################################################################################################################
' Borra los nombres de ambito HOJA de una hoja que apuntan a #REF! (restos de copiar SwitchsAPP a ListObjAPP: hacian sombra a los nombres
' Sw_* buenos, de ambito libro, al leerlos desde esa hoja). No toca los nombres de ambito libro. Desde Inmediato:
'   Call Rut_Names_Borrar_REF_Hoja(Prog__APP_ListObj)
Public Sub Rut_Names_Borrar_REF_Hoja(ByVal Ws As Worksheet)
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim i       As Long
    Dim N       As Long
    For i = Ws.Names.Count To 1 Step -1
        If InStr(Ws.Names(i).Name, "!") > 0 Then                                '- los de ambito hoja salen como "Hoja!Nombre"
            If InStr(Ws.Names(i).RefersTo, "#REF!") > 0 Then
                Ws.Names(i).Delete
                N = N + 1
            End If
        End If
    Next i
    Debug.Print "Rut_Names_Borrar_REF_Hoja: " & N & " nombres borrados de " & Ws.Name
End Sub
