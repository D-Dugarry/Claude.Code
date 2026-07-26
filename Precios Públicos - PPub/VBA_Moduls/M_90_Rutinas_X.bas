Attribute VB_Name = "M_90_Rutinas_X"
Option Explicit

'' ==================================================================================================================================
'            '###################################################################################################################################
'                    Sub Rut_WrkSheet_LstObj_LiberarEspacio_xx()
''                        Application.Workbooks(ThisWorkbook.Name).Sheets(ActiveSheet.Name).ListObjects(1).DataBodyRange.Delete
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(ActiveSheet)
'                    End Sub
''###################################################################################################################################
'Sub Rut_WrkSheet_LstObj_LiberarEspacio(ByVal WrkSht As String)  '--- Borra TODO a la Derecha y Abajo de .ListObjects(1) ----------------------
'' ==================================================================================================================================
'        Application.Calculation = xlManual
'        Application.ScreenUpdating = False
'    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
'            .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
'            .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
'            If .FilterMode Then .ShowAllData            ' Deshacer Filtros
'        With .ListObjects(1).Range                      ' Borra la filas de abajo y columnas de la derecha del la Tabla .ListObjects(1)
'            Range(.Cells(.Rows.Count, .Columns.Count).Address).Select   ' Selecciona la última celda de la tabla
'            ActiveCell.Offset(1, 1).Select
'            Range(ActiveCell.Address & ":" & Cells(Rows.Count, 1).Address).EntireRow.Delete
'            Range(ActiveCell.Address & ":" & Cells(1, Columns.Count).Address).EntireColumn.Delete
'            ActiveSheet.UsedRange                       ' Para restablecer el rango de celdas en uso
'        End With
'    End With
'        Application.ScreenUpdating = True
'        Application.Calculation = xlAutomatic
'End Sub
'' -------------------------------------------------------------------------------------------------------------------------------<<<
' ==================================================================================================================================
Sub Rut_x_Filtro_LoTb(ByRef LoTb As ListObject, Columna As Integer, Criterio As String, Optional SW_Clear As Boolean = False)
' ----------------------------------------------------------------------------------------------------------------------------------
    If SW_Clear And Not LoTb.AutoFilter Is Nothing Then LoTb.AutoFilter.ShowAllData
    LoTb.Range.AutoFilter Field:=Columna, Criteria1:=Criterio
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

' ==================================================================================================================================
Sub Rut_x_Ocultar_Cuadro_Instrucciones()
' ----------------------------------------------------------------------------------------------------------------------------------
    ActiveSheet.Shapes("Cuadro_Instrucciones").Visible = False
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================












' ==================================================================================================================================
Sub Rut_Quita_Ascii_160(ByRef Lo_Tb As ListObject, Columna As Integer)
' ==================================================================================================================================
    Lo_Tb.ListColumns(6).DataBodyRange.Select
    Selection.Replace What:=Chr(160), Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'==================================================================================================================================
Sub Rut_Ajustar_V_H_Alignment()    ' Ajustar la Vertical/Horizontal-Alignment de todas las columnas  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================================================
Dim Col As Integer
With Lo_Aplic.DataBodyRange
    For Col = 1 To Lo_Prog_Colns.DataBodyRange.Columns.Count
        .Range(Cells(1, Col), Cells(.Rows.Count, Col)).HorizontalAlignment = Lo_Prog_Colns.DataBodyRange.Cells(5, Col).Value2
        .Range(Cells(1, Col), Cells(.Rows.Count, Col)).VerticalAlignment = Lo_Prog_Colns.DataBodyRange.Cells(11, Col).Value2
    Next Col
End With
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
Sub Rut_Filas_Ajustar_Alto()        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================================================
    Lo_Aplic.HeaderRowRange.RowHeight = 35
    Lo_Aplic.DataBodyRange.Rows.EntireRow.AutoFit
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
Sub Rut_Columnas_Ajustar_Ancho()      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================================================
    Dim Cont_Col As Integer
    Dim Ancho   As Integer
    For Cont_Col = 1 To LastCol_Tb_Solicitudes
        If Not Columns(Cont_Col).Hidden And Lo_Prog_Colns.DataBodyRange.Cells(7, Cont_Col) <> "Ocultar" Then
            Ancho = Lo_Prog_Colns.DataBodyRange.Cells(4, Cont_Col)      '.Value2
            Columns(Cont_Col).ColumnWidth = Ancho
        End If
    Next Cont_Col
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
Sub Rut_Filas_Mostrar()      ' Muestra Todas las Solicitudes  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Rows.EntireRow.Hidden = False
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
Sub Rut_Filas_Mostrar_WrkSht(WrkSht As String)      ' Muestra Todas las Solicitudes  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Sheets(WrkSht).Rows.EntireRow.Hidden = False
    Sheets(WrkSht).Rows.EntireRow.AutoFit
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
Sub Rut_Columnas_Mostrar_WrkSht(WrkSht As String)     ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================================================
    Sheets(WrkSht).Columns.EntireColumn.Hidden = False
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
Sub Rut_Columnas_Mostrar()      ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================================================
Dim Cont_Col As Integer
Application.ScreenUpdating = False
        For Cont_Col = 1 To LastCol_Tb_Solicitudes
            If Lo_Prog_Colns.DataBodyRange.Cells(7, Cont_Col) <> "Ocultar" Then Columns(Cont_Col).Hidden = False
        Next Cont_Col
Application.ScreenUpdating = True
Debug.Print "Columnas Ocultas Visibles"
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
' ==================================================================================================================================
Sub Rut_Select_File(Título As String, ByRef NomFich As String, TipoFich As String)  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================================================
Dim fDialog As Office.FileDialog
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = Título
        .InitialFileName = ThisWorkbook.Path & "\"
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        .ButtonName = "Seleccionar"
        .Filters.Clear
        .Filters.Add "Sólo Ficheros Excel", TipoFich, 1
        If .Show = False Then
'            MsgBox "Ha pulsado el botón <Cancelar>.", vbOKOnly, "Proceso: Seleccionar Archivo"
            MsgBx_Msg = "Ha pulsado el botón <Cancelar>."
            MsgBx_Title = "Proceso: Seleccionar Archivo."
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", 1, "Ask"): Form_MsgBox.Show     '- ([Font-Size]=16, [Red-Border]=False, [Buttons]="Ok", [Default-Button]=1, [Image]="Msg")
            NomFich = "Cancel"
        Else
            NomFich = .SelectedItems(1)
        End If
    End With
Set fDialog = Nothing
' ---------------------------------------------------------------------------------------------------------------------------<<<
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' ==================================================================================================================================
' =====================     RuT_Visible_Hidde_Tablas_Prog              ==============================================================
' ==================================================================================================================================
Sub Rut_Visible_Hidde_Tablas_Prog()
    Dim WkS As Worksheet
    For Each WkS In Worksheets
        If Left(WkS.Name, 5) = "Prog_" Then
            Select Case Visibilidad_Hoja
                Case Is = "Visible"
                WkS.Visible = xlSheetVisible
                Case Is = "Hidden"
                WkS.Visible = xlSheetHidden
                Case Is = "VeryHidden"
                WkS.Visible = xlSheetVeryHidden
            End Select  ' Case wks.Visible
        End If
    Next
End Sub     '      <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

' ==================================================================================================================================
'' ==================================================================================================================================
'' =====================     RuT_Ocultar_Col_SN     ==============================================================
'' ==================================================================================================================================
'' ==================================================================================================================================
'Sub Rut_Ocultar_Col_SN()
'Dim Cont    As Integer
'Dim Fila    As Integer
'Dim Clmn    As Integer
'Rut_Off_Functions
'    If SW_Col_Ocultas Then                    ' Mostrar Columnas
'        SW_Col_Ocultas = False
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(200, 200, 250)
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Ocultar Col's"
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN-2")).Visible = True
'        Columns.EntireColumn.Hidden = False
'    Else                                        ' Ocultar Columnas
'        SW_Col_Ocultas = True
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(150, 200, 100)
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Mostrar Col's"
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN-2")).Visible = False
'            Fila = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Row - 1
'            Clmn = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Column
'        For Cont = Clmn To ActiveSheet.ListObjects(1).Range.Columns.Count + Clmn
'            If IsEmpty(Cells(Fila, Cont)) And Not Columns(Cont).Hidden Then Columns(Cont).Hidden = True
'            If IsEmpty(Prog_DefCol.Cells(2, Cont)) And Not Columns(Cont).Hidden Then Columns(Cont).Hidden = True
'        Next Cont
'    End If
'Rut_On_Functions
'End Sub     ' RuT_Ocultar_Col_SN     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'' ==================================================================================================================================

' ==================================================================================================================================
' =====================  RuT_Unhide_All_Sheets            ====================================================
' ==================================================================================================================================
Sub Rut_Unhide_All_Sheets()     ' Para uso interno, hacer visible todas las hojas
    Dim WkS As Worksheet
 
    For Each WkS In ThisWorkbook.Worksheets
        WkS.Visible = xlSheetVisible
    Next WkS
End Sub
' ==================================================================================================================================
    
' ==================================================================================================================================
Sub Rut_x_Vaciar_WrkSht(WrkSht As Worksheet)      ' Borra la filas de abajo y columnas de la derecha del la Tabla  ----------------------------------
With WrkSht
        .Unprotect
        .Range("a1", .Cells(Rows.Count, 1)).EntireRow.Delete
        .Range("a1", .Cells(1, Columns.Count)).EntireColumn.Delete
'        .UsedRange
End With
End Sub


