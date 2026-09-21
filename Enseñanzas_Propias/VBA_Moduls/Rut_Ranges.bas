Attribute VB_Name = "Rut_Ranges"
' Last Rev. 2026-09-21 12:12
Option Explicit
' ==================================================================================================
Function Fnc_Range_Exist(RngName As String) As Boolean
    Dim Rng As Range
    On Error Resume Next
    Set Rng = Range(RngName)
    On Error GoTo 0
    Fnc_Range_Exist = Not Rng Is Nothing
End Function
' ==================================================================================================
'Sub Rut_Ranges_List_ALL()
'' ----------------------------------------------------------------
'    Dim ws As Worksheet
'    Dim i       As Integer
'    Dim nm      As Name
''    Set ws = Worksheets.Add
'    Set ws = ThisWorkbook.Sheets(Prog_List_Ranges.Name)
'    If Not ws.ListObjects(1).DataBodyRange Is Nothing Then ws.ListObjects(1).DataBodyRange.Delete
'
'    i = 1
'    ws.Cells(i, 1).Value2 = "Name"
'    ws.Cells(i, 2).Value2 = "RefersTo"
'    ws.Cells(i, 3).Value2 = "Comment"
'    ws.Cells(i, 4).Value2 = "Creator"
'    ws.Cells(i, 5).Value2 = "Index"
'    ws.Cells(i, 6).Value2 = "MacroType"
'    ws.Cells(i, 7).Value2 = "NameLocal"
'    ws.Cells(i, 8).Value2 = "RefersToLocal"
'    ws.Cells(i, 9).Value2 = "RefersToR1C1Local"
'    ws.Cells(i, 10).Value2 = "ValidWorkbookParameter"
'    ws.Cells(i, 11).Value2 = "Visible"
'    ws.Cells(i, 12).Value2 = "WorkbookParameter"
'    ws.Cells(i, 13).Value2 = "RefersToR1C1"
'    ws.Cells(i, 14).Value2 = "Value"
'    i = 2
'    For Each nm In ThisWorkbook.Names
'        ws.Cells(i, 1).Value2 = nm.Name
'        ws.Cells(i, 2).Value2 = "'" & nm.RefersTo
'        ws.Cells(i, 3).Value2 = nm.Comment
'        ws.Cells(i, 4).Value2 = nm.Creator
'        ws.Cells(i, 5).Value2 = nm.Index
'        ws.Cells(i, 6).Value2 = nm.MacroType
'        ws.Cells(i, 7).Value2 = nm.NameLocal
'        ws.Cells(i, 8).Value2 = "'" & nm.RefersToLocal
'        ws.Cells(i, 9).Value2 = "'" & nm.RefersToR1C1Local
'        ws.Cells(i, 10).Value2 = nm.ValidWorkbookParameter
'        ws.Cells(i, 11).Value2 = nm.Visible
'        ws.Cells(i, 12).Value2 = nm.WorkbookParameter
'        ws.Cells(i, 13).Value2 = nm.RefersToR1C1
'        ws.Cells(i, 14).Value2 = nm.Value
'        i = i + 1
'    Next
''    ws.Columns("A:l").AutoFit
'End Sub


' ==================================================================================================
' ============ Método 1.           Copia un rango especificado pero sólo las celdas visibles     ===
' ==================================================================================================
'Sub RuT_UsedRange_Save_New_WorkBook_Liq_TPV()
'    Dim Lo_Liq      As ListObject:      Set Lo_Liq = H_Liq_TPV.ListObjects(1)
'
'Application.ScreenUpdating = False
'    '   Copy and Save a Specific Range in a New WorkBook -----------------------------
'    Dim Rng As Range
'    On Error Resume Next
'        Set Rng = ActiveSheet.Range(Cells(1, 1), Cells(Lo_Liq.TotalsRowRange.Row, Lo_Liq.Range.Columns(C_TPV_Liq_Obs).Column))
'        Set Rng = Rng.SpecialCells(xlCellTypeVisible)   '- Copio sólo las celdas visibles ---------
'    Dim Wb_new      As Workbook:        Set Wb_new = Workbooks.Add
'    With Rng.Copy
'        Wb_new.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'        Wb_new.ActiveSheet.Paste
'    End With
'    Application.CutCopyMode = False
'    Wb_new.ActiveSheet.Range("A1").Select
'    Dim IntialName As String
'    Dim sFileSaveName As Variant
'    Dim FPath       As String:          FPath = Fnc_NEXE_RutaAPP() & "\"
'    IntialName = FPath & "Liquid_" & H_Liq_TPV.Range("Liq_Núm") & "_JyC_" & H_Liq_TPV.Range("Liq_Siglas")    ' & ".xlsx"
'    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
'        If sFileSaveName <> False Then
'            On Error GoTo Restablecer_Valores
''            Application.DisplayAlerts = False
'            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
''            Application.DisplayAlerts = True
'        End If
'    If Wb_new.ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
'    Wb_new.Close SaveChanges:=True    '= ActiveWorkbook.Close savechanges:=True
'    MsgBox "Hecho"
'    Application.ScreenUpdating = True
'    Exit Sub
'Restablecer_Valores:
'    Application.ScreenUpdating = True
'    Wb_new.Close SaveChanges:=False
'End Sub     ' RuT_UsedRange_Save_New_WorkBook_Liq_TPV

' ==================================================================================================
' ============ Método 2.           Copia un rango especificado pero sólo las celdas visibles     ===
' ==================================================================================================
'Sub RuT_UsedRange2_Save_New_WorkBook_Liq_TPV()
'    Dim Lo_Liq      As ListObject:      Set Lo_Liq = H_Liq_TPV.ListObjects(1)
'
'Application.ScreenUpdating = False
'    '   Copy and Save a Specific Range in a New WorkBook -----------------------------
'    ActiveSheet.Range(Cells(1, 1), Cells(Lo_Liq.TotalsRowRange.Row, Lo_Liq.Range.Columns(C_TPV_Liq_Obs).Column)).Copy
'        Dim Wb_new      As Workbook:        Set Wb_new = Workbooks.Add
'        With Wb_new.ActiveSheet.Range("A1")
'            .PasteSpecial xlPasteValues
'            .PasteSpecial xlPasteColumnWidths
'            .PasteSpecial xlPasteFormats
'        End With
'        Application.CutCopyMode = False
'    '- Borro líneas ocultas.  ---------------------<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Dim Lin     As Integer
'    For Lin = 1 To Lo_Liq.TotalsRowRange.Row
'        If H_Liq_TPV.Rows(Lin).EntireRow.Hidden Then Wb_new.ActiveSheet.Rows(Lin).EntireRow.Hidden = True
'    Next Lin
'
'    Wb_new.ActiveSheet.Range("A1").Select
'
'    Dim IntialName As String
'    Dim sFileSaveName As Variant
'    Dim FPath       As String:          FPath = Fnc_NEXE_RutaAPP() & "\"
'    IntialName = FPath & "Liquid_" & H_Liq_TPV.Range("Liq_Núm") & "_JyC_" & H_Liq_TPV.Range("Liq_Siglas")    ' & ".xlsx"
'    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
'        If sFileSaveName <> False Then
'            On Error GoTo Restablecer_Valores
''            Application.DisplayAlerts = False
'            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
''            Application.DisplayAlerts = True
'        End If
'    If Wb_new.ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
'    Wb_new.Close SaveChanges:=True    '= ActiveWorkbook.Close savechanges:=True
'    MsgBox "Hecho"
'    Application.ScreenUpdating = True
'    Exit Sub
'Restablecer_Valores:
'    Application.ScreenUpdating = True
'    Wb_new.Close SaveChanges:=False
'End Sub     ' RuT_UsedRange2_Save_New_WorkBook_Liq_TPV

' ==================================================================================================
' ======= Método 3.           Copia un rango que seleccionamos pero sólo las celdas visibles     ===
' ==================================================================================================
'Sub RuT_Range_Save_New_WorkBook_Liq_TPV()
'    Dim FPath           As String:    FPath = Fnc_NEXE_RutaAPP() & "\"
''   Copy and Save a Selected Range in a New WorkBook -----------------------------
'    Dim Rng As Range        '- Select a Range -------------<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    On Error Resume Next
'        Set Rng = Application.InputBox(Title:="Select a Range", _
'                            Prompt:="Select a Range to send in the e-mail Body.", Type:=8)
'        Set Rng = Rng.SpecialCells(xlCellTypeVisible)
'    On Error GoTo 0
'    Application.ScreenUpdating = False
'    If Rng Is Nothing Then Exit Sub
'    'Sub sbSaveExcelDialog()
'    Dim Wb_new As Workbook
'    Set Wb_new = Workbooks.Add
'    With Rng.Copy
'        Wb_new.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'        ActiveSheet.Paste
'    End With
'    Wb_new.ActiveSheet.Range("A1").Select
'
'    Dim IntialName As String
'    Dim sFileSaveName As Variant
'    IntialName = FPath & "Liquid_" & H_Liq_TPV.Range("Liq_Núm") & "_JyC_" & H_Liq_TPV.Range("Liq_Siglas")    ' & ".xlsx"
'    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
'        If sFileSaveName <> False Then
'            On Error GoTo Restablecer_Valores
''            Application.DisplayAlerts = False
'            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
''            Application.DisplayAlerts = True
'        End If
'    If Wb_new.ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
'    Wb_new.Close SaveChanges:=True
'    'ActiveWorkbook.Close savechanges:=True
'    MsgBox "Hecho"
'    Application.ScreenUpdating = True
'    Exit Sub
'Restablecer_Valores:
'    Application.ScreenUpdating = True
'    ActiveWorkbook.Close SaveChanges:=False
'End Sub     ' RuT_Range_Save_New_WorkBook_Liq_TPV

' ==================================================================================================
' ======= Método 3.           Copia un rango que seleccionamos pero sólo las celdas visibles     ===
' ==================================================================================================
'Sub RuT_Range1_Save_Liquidación_New_WorkBook()
'    Dim FPath           As String
'
''    FPath = ActiveWorkbook.Path & "\"
'    FPath = Fnc_NEXE_RutaAPP() & "\"
'
''   Copy and Save a Selected Range in a New WorkBook -----------------------------
'    '- Select a Range -------------
'    Dim Rng As Range
'    On Error Resume Next
'        Set Rng = Application.InputBox(Title:="Select a Range", _
'                    Prompt:="Select a Range to send in the e-mail Body.", Type:=8)
'        Set Rng = Rng.SpecialCells(xlCellTypeVisible)
'    On Error GoTo 0
'    Application.ScreenUpdating = False
'    If Rng Is Nothing Then Exit Sub
'    Rng.Copy
'    'Sub sbSaveExcelDialog()
'    Dim Wb_new As Workbook
'    Set Wb_new = Workbooks.Add
'    Wb_new.ActiveSheet.Range("A1").Select
'    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'    ActiveSheet.Paste
'
'    Dim IntialName As String
'    Dim sFileSaveName As Variant
'    IntialName = FPath & "Liquid_" & Range("e2") & "_JyC_" & Range("e1")    ' & ".xlsx"
'    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
'        If sFileSaveName <> False Then
'            On Error GoTo Restablecer_Valores
''            Application.DisplayAlerts = False
'            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
''            Application.DisplayAlerts = True
'        End If
'Range("a1").Select
'If ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
'ActiveWorkbook.Close SaveChanges:=True
'MsgBox "Hecho"
'Application.ScreenUpdating = True
'Exit Sub
'Restablecer_Valores:
'Application.ScreenUpdating = True
'ActiveWorkbook.Close SaveChanges:=False
'End Sub     ' RuT_Range_Save_Liquidación_New_WorkBook


