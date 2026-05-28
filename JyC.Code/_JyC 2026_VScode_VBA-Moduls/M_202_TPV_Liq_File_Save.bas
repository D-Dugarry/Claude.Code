Attribute VB_Name = "M_202_TPV_Liq_File_Save"
'M_102_CTA_Liq_File_Save
Option Explicit

' ==================================================================================================================================
' ============= Método 1.           Copia un rango especificado pero sólo las celdas visibles     ==================================
' ==================================================================================================================================
Sub RuT_UsedRange_Save_New_WorkBook_Liq_TPV()
    Dim Lo_Liq      As ListObject:      Set Lo_Liq = H_Liq_TPV.ListObjects(1)
    
Application.ScreenUpdating = False
    '   Copy and Save a Specific Range in a New WorkBook -----------------------------
    Dim Rng As Range
    On Error Resume Next
        Set Rng = ActiveSheet.Range(Cells(1, 1), Cells(Lo_Liq.TotalsRowRange.Row, Lo_Liq.Range.Columns(C_TPV_Liq_Obs).Column))
        Set Rng = Rng.SpecialCells(xlCellTypeVisible)   '- Copio sólo las celdas visibles ----------<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim WB_New      As Workbook:        Set WB_New = Workbooks.Add
    With Rng.Copy
        With WB_New.ActiveSheet.Range("A1")
            .PasteSpecial xlPasteValues
            .PasteSpecial xlPasteColumnWidths
            .PasteSpecial xlPasteFormats
        End With
        WB_New.ActiveSheet.Paste
    End With
'    With Rng.Copy
'        WB_New.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'        WB_New.ActiveSheet.Paste
'    End With
    Application.CutCopyMode = False
    WB_New.ActiveSheet.Range("A1").Select
    WB_New.ActiveSheet.Rows(1).RowHeight = 40
    WB_New.ActiveSheet.Rows("2:9").RowHeight = 20
    Dim IntialName As String
    Dim sFileSaveName As Variant
    Dim FPath       As String
    FPath = Prog__APP.Range("APP_User_Unid_Red") & Prog__APP.Range("APP_RutaRed_Data") & "\"
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
Sub RuT_Range_Save_New_WorkBook_Liq_TPV_33()
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
' =============================     RuT_Sheet_Save_Liquidación_New_WorkBook     =======================================================================
' ==================================================================================================================================
Sub RuT_Sheet_Save_Liquidación_New_WorkBook()
    Dim FPath           As String

Application.ScreenUpdating = False
    FPath = Fnc_NEXE_RutaAPP() & "\"
'   Copy and Save a Sheet in a New WorkBook -----------------------------
    ActiveSheet.Copy
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
Range("a1").Select
If ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
ActiveWorkbook.Close SaveChanges:=True
MsgBox "Hecho"
Application.ScreenUpdating = True
Exit Sub
Restablecer_Valores:
Application.ScreenUpdating = True
ActiveWorkbook.Close SaveChanges:=False
End Sub     ' RuT_Sheet_Save_Liquidación_New_WorkBook



