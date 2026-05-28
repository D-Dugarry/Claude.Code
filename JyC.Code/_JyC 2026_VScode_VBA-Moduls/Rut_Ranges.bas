Attribute VB_Name = "Rut_Ranges"
Option Explicit
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Sub Ejemplo_Selección_Múltiple()
    Union(Cells(44 + 4, 4), Range(Cells(44 + 5, 3), Cells(44 + 3 + 5, 11))).Select
End Sub
' ==================================================================================================================================
Function Fnc_Range_Exist(RngName As String) As Boolean
    Dim Rng As Range
    On Error Resume Next
    Set Rng = Range(RngName)
    On Error GoTo 0
    Fnc_Range_Exist = Not Rng Is Nothing
End Function
' ==================================================================================================================================
Sub Rut_Ranges_List_ALL()
' ----------------------------------------------------------------
    Dim Ws As Worksheet
    Dim i       As Integer
    Dim nm      As Name
'    Set ws = Worksheets.Add
    Set Ws = ThisWorkbook.Sheets(Prog_List_Ranges.Name)
    If Not Ws.ListObjects(1).DataBodyRange Is Nothing Then Ws.ListObjects(1).DataBodyRange.Delete
    
    i = 1
    Ws.Cells(i, 1).Value2 = "Name"
    Ws.Cells(i, 2).Value2 = "RefersTo"
    Ws.Cells(i, 3).Value2 = "Comment"
    Ws.Cells(i, 4).Value2 = "Creator"
    Ws.Cells(i, 5).Value2 = "Index"
    Ws.Cells(i, 6).Value2 = "MacroType"
    Ws.Cells(i, 7).Value2 = "NameLocal"
    Ws.Cells(i, 8).Value2 = "RefersToLocal"
    Ws.Cells(i, 9).Value2 = "RefersToR1C1Local"
    Ws.Cells(i, 10).Value2 = "ValidWorkbookParameter"
    Ws.Cells(i, 11).Value2 = "Visible"
    Ws.Cells(i, 12).Value2 = "WorkbookParameter"
    Ws.Cells(i, 13).Value2 = "RefersToR1C1"
    Ws.Cells(i, 14).Value2 = "Value"
    i = 2
    For Each nm In ThisWorkbook.Names
        Ws.Cells(i, 1).Value2 = nm.Name
        Ws.Cells(i, 2).Value2 = "'" & nm.RefersTo
        Ws.Cells(i, 3).Value2 = nm.Comment
        Ws.Cells(i, 4).Value2 = nm.Creator
        Ws.Cells(i, 5).Value2 = nm.index
        Ws.Cells(i, 6).Value2 = nm.MacroType
        Ws.Cells(i, 7).Value2 = nm.NameLocal
        Ws.Cells(i, 8).Value2 = "'" & nm.RefersToLocal
        Ws.Cells(i, 9).Value2 = "'" & nm.RefersToR1C1Local
        Ws.Cells(i, 10).Value2 = nm.ValidWorkbookParameter
        Ws.Cells(i, 11).Value2 = nm.Visible
        Ws.Cells(i, 12).Value2 = nm.WorkbookParameter
        Ws.Cells(i, 13).Value2 = nm.RefersToR1C1
        Ws.Cells(i, 14).Value2 = nm.Value
        i = i + 1
    Next
'    ws.Columns("A:l").AutoFit
End Sub


' ==================================================================================================================================
' ============= Método 1.           Copia un rango especificado pero sólo las celdas visibles     ==================================
' ==================================================================================================================================
Sub RuT_UsedRange_Save_New_WorkBook_Liq_TPV_2()
    Dim Lo_Liq      As ListObject:      Set Lo_Liq = H_Liq_TPV.ListObjects(1)
    
Application.ScreenUpdating = False
    '   Copy and Save a Specific Range in a New WorkBook -----------------------------
    Dim Rng As Range
    On Error Resume Next
        Set Rng = ActiveSheet.Range(Cells(1, 1), Cells(Lo_Liq.TotalsRowRange.Row, Lo_Liq.Range.Columns(C_TPV_Liq_Obs).Column))
        Set Rng = Rng.SpecialCells(xlCellTypeVisible)   '- Copio sólo las celdas visibles ----------<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim WB_New      As Workbook:        Set WB_New = Workbooks.Add
    With Rng.Copy
        WB_New.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        WB_New.ActiveSheet.Paste
    End With
    Application.CutCopyMode = False
    WB_New.ActiveSheet.Range("A1").Select
    Dim IntialName As String
    Dim sFileSaveName As Variant
    Dim FPath       As String:          FPath = Fnc_NEXE_RutaAPP() & "\"
    IntialName = FPath & "Liquid_" & H_Liq_TPV.Range("Liq_Núm") & "_JyC_" & H_Liq_TPV.Range("Liq_Siglas")    ' & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo Restablecer_Valores
'            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
'            Application.DisplayAlerts = True
        End If
    If WB_New.ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
    WB_New.Close SaveChanges:=True    '= ActiveWorkbook.Close savechanges:=True
    MsgBox "Hecho"
    Application.ScreenUpdating = True
    Exit Sub
Restablecer_Valores:
    Application.ScreenUpdating = True
    WB_New.Close SaveChanges:=False
End Sub     ' RuT_UsedRange_Save_New_WorkBook_Liq_TPV

' ==================================================================================================================================
' ============= Método 2.           Copia un rango especificado pero sólo las celdas visibles     ==================================
' ==================================================================================================================================
Sub RuT_UsedRange2_Save_New_WorkBook_Liq_TPV()
    Dim Lo_Liq      As ListObject:      Set Lo_Liq = H_Liq_TPV.ListObjects(1)
    
Application.ScreenUpdating = False
    '   Copy and Save a Specific Range in a New WorkBook -----------------------------
    ActiveSheet.Range(Cells(1, 1), Cells(Lo_Liq.TotalsRowRange.Row, Lo_Liq.Range.Columns(C_TPV_Liq_Obs).Column)).Copy
        Dim WB_New      As Workbook:        Set WB_New = Workbooks.Add
        With WB_New.ActiveSheet.Range("A1")
            .PasteSpecial xlPasteValues
            .PasteSpecial xlPasteColumnWidths
            .PasteSpecial xlPasteFormats
        End With
        Application.CutCopyMode = False
    '- Borro líneas ocultas.  ---------------------<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim Lin     As Integer
    For Lin = 1 To Lo_Liq.TotalsRowRange.Row
        If H_Liq_TPV.Rows(Lin).EntireRow.Hidden Then WB_New.ActiveSheet.Rows(Lin).EntireRow.Hidden = True
    Next Lin
    
    WB_New.ActiveSheet.Range("A1").Select
    
    Dim IntialName As String
    Dim sFileSaveName As Variant
    Dim FPath       As String:          FPath = Fnc_NEXE_RutaAPP() & "\"
    IntialName = FPath & "Liquid_" & H_Liq_TPV.Range("Liq_Núm") & "_JyC_" & H_Liq_TPV.Range("Liq_Siglas")    ' & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo Restablecer_Valores
'            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
'            Application.DisplayAlerts = True
        End If
    If WB_New.ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
    WB_New.Close SaveChanges:=True    '= ActiveWorkbook.Close savechanges:=True
    MsgBox "Hecho"
    Application.ScreenUpdating = True
    Exit Sub
Restablecer_Valores:
    Application.ScreenUpdating = True
    WB_New.Close SaveChanges:=False
End Sub     ' RuT_UsedRange2_Save_New_WorkBook_Liq_TPV

' ==================================================================================================================================
' ============= Método 3.           Copia un rango que seleccionamos pero sólo las celdas visibles     ==================================
' ==================================================================================================================================
Sub RuT_Range_Save_New_WorkBook_Liq_TPV_V3()
    Dim FPath           As String:    FPath = Fnc_NEXE_RutaAPP() & "\"
'   Copy and Save a Selected Range in a New WorkBook -----------------------------
    Dim Rng As Range        '- Select a Range -------------<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    On Error Resume Next
        Set Rng = Application.InputBox(Title:="Select a Range", _
                            Prompt:="Select a Range to send in the e-mail Body.", Type:=8)
        Set Rng = Rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    Application.ScreenUpdating = False
    If Rng Is Nothing Then Exit Sub
    'Sub sbSaveExcelDialog()
    Dim WB_New As Workbook
    Set WB_New = Workbooks.Add
    With Rng.Copy
        WB_New.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ActiveSheet.Paste
    End With
    WB_New.ActiveSheet.Range("A1").Select
    
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "Liquid_" & H_Liq_TPV.Range("Liq_Núm") & "_JyC_" & H_Liq_TPV.Range("Liq_Siglas")    ' & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo Restablecer_Valores
'            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
'            Application.DisplayAlerts = True
        End If
    If WB_New.ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
    WB_New.Close SaveChanges:=True
    'ActiveWorkbook.Close savechanges:=True
    MsgBox "Hecho"
    Application.ScreenUpdating = True
    Exit Sub
Restablecer_Valores:
    Application.ScreenUpdating = True
    ActiveWorkbook.Close SaveChanges:=False
End Sub     ' RuT_Range_Save_New_WorkBook_Liq_TPV

' ==================================================================================================================================
' ============= Método 3.           Copia un rango que seleccionamos pero sólo las celdas visibles     ==================================
' ==================================================================================================================================
Sub RuT_Range1_Save_Liquidación_New_WorkBook()
    Dim FPath           As String

'    FPath = ActiveWorkbook.Path & "\"
    FPath = Fnc_NEXE_RutaAPP() & "\"
    
'   Copy and Save a Selected Range in a New WorkBook -----------------------------
    '- Select a Range -------------
    Dim Rng As Range
    On Error Resume Next
        Set Rng = Application.InputBox(Title:="Select a Range", _
                    Prompt:="Select a Range to send in the e-mail Body.", Type:=8)
        Set Rng = Rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    Application.ScreenUpdating = False
    If Rng Is Nothing Then Exit Sub
    Rng.Copy
    'Sub sbSaveExcelDialog()
    Dim WB_New As Workbook
    Set WB_New = Workbooks.Add
    WB_New.ActiveSheet.Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "Liquid_" & Range("e2") & "_JyC_" & Range("e1")    ' & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo Restablecer_Valores
'            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
'            Application.DisplayAlerts = True
        End If
Range("a1").Select
If ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
ActiveWorkbook.Close SaveChanges:=True
MsgBox "Hecho"
Application.ScreenUpdating = True
Exit Sub
Restablecer_Valores:
Application.ScreenUpdating = True
ActiveWorkbook.Close SaveChanges:=False
End Sub     ' RuT_Range_Save_Liquidación_New_WorkBook


