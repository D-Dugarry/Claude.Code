Attribute VB_Name = "M90_Rut_Format_Colmns"
' Last Rev. 2026-09-23 18:56
'- M90_Rut_Format_Colmns - Modif: 2025-10-08
Option Explicit

        Sub Rut_X_Format_LoData_LoDefCol_ByHand()  '--- Solo hace falta cambiar las variables de la Rutina
            Call Rut_X_Format_LoData_LoDefCol(Prog_LsGes04.ListObjects(1), Prog_DefCol_G04.ListObjects(1))
'            Call Rut_X_Format_LoData_LoDefCol(Prog_BD.ListObjects(1), Prog_DefCol_BD.ListObjects(1))
        End Sub
' ==================================================================================================
Sub Rut_X_Format_LoData_LoDefCol(ByRef LoData As ListObject, ByRef LoDefCol As ListObject)  '#######
' ==================================================================================================
Dim Ccol    As Integer

    If LoData.DataBodyRange Is Nothing Then
        MsgBox "¡¡¡ Tabla SIN DATOS !!!", vbOKOnly + vbExclamation, "Proceso: Formatear Tabla ListObjects"
        Exit Sub
    End If
    LoData.DataBodyRange.Select
    With Selection.Font
        .Name = "Courier New"
        .Size = 11
        .ColorIndex = xlAutomatic
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    '--- Formatear las Columnas --------------------------------------------------------------------
    Dim MaxCol  As Integer
    MaxCol = Application.Min(LoData.Range.Columns.Count, LoDefCol.ListRows.Count)
    For Ccol = 1 To MaxCol
        If Not LoDefCol.DataBodyRange.Cells(Ccol, DefC_FormatCol) Then GoTo NextCol    '- Sólo si se desea formatear la Col.
'            '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------
'            CantFormatCol = CantFormatCol + 1
'            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "___________________ Formateando Col. " & Ccol & " (" & CantFormatCol & "ª.), de " & TotCantFormatCol & " Col.", 0, , , TxT_Prog, , , 2)
        Select Case LoDefCol.DataBodyRange.Cells(Ccol, DefC_TipVar)
            Case "T"
                    LoData.DataBodyRange.Columns(Ccol).Select
                With Selection
                    .NumberFormat = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Format)
                    .WrapText = LoDefCol.DataBodyRange.Cells(Ccol, DefC_WrapTxt).Value2
                End With
            Case "F"
               With LoData.DataBodyRange
                    .Columns(Ccol).Select
                    Selection.NumberFormat = "General"
                    If LoData.ListColumns.Count = Ccol Then
                        LoData.ListColumns.Add
                    Else
                        .Columns(Ccol + 1).Select
                        Selection.ListObject.ListColumns.Add Position:=Ccol + 1
                    End If
                    .Cells(1, Ccol + 1).Select
                         ActiveCell.FormulaR1C1 = "=RC[-1]*1"       '    ActiveCell.FormulaR1C1 = "=+[@FECHAEMISION]*1"
                    .Columns(Ccol + 1).Select
                    Selection.Copy
                    .Cells(1, Ccol).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    .Columns(Ccol + 1).Delete
                    .Columns(Ccol).Select
                    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                        ReplaceFormat:=False ', FormulaVersion:=xlReplaceFormula2    '- Quitado porque da error según el ord de 32b o 64b... _
                           "Error de compilación en el módulo oculto: <nombre del módulo>" Este error se produce normalmente cuando el código es incompatible con la versión o la arquitectura de esta aplicación (por ejemplo, el código de un documento está dirigido a aplicaciones de Microsoft Office de 32 bits pero se está intentando ejecutar en Office de 64 bits).
                    Selection.NumberFormat = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Format)
               End With
            Case "N"
               With LoData.DataBodyRange
                    .Columns(Ccol).Select
                    If WorksheetFunction.CountA(.Columns(Ccol)) > 0 Then
                        '- Convierto a números, texto con formato punto de millares y coma decimal.
                        Selection.TextToColumns Destination:=Range(.Cells(1, Ccol).Address(False, False)), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                            Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
                            FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
                    End If
'                    Dim Rc As Range '--- Si es un número muy grande lo muestra como 99999E+12, con el For-Next lo quitamos
'                    For Each Rc In .Columns(1)
'                        With Rc.Cells(1)
'                            If IsNumeric(.Value2) And .Text Like "*E+*" Then
'                                Rc.NumberFormat = "#"
'                            End If
'                        End With
'                    Next
                    With Selection
                        .NumberFormat = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Format)
                    End With
                End With
            Case Else
                MsgBox "Error en Tipo de Variable, NO es T,F ó N ???", vbExclamation, "Procedimiento: Formatear Tabla."
        End Select
NextCol:
'        Ws_Data.Columns(LoData.ListColumns(Ccol).Range.Column).ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht).Value
        LoData.Range.Columns(Ccol).ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht).Value
        LoData.Range.Columns(Ccol).HorizontalAlignment = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Align).Value
Next
        
Debug.Print "<<< Rut_Lo_Format_LoData_LoDefCol"
End Sub     ' Rut_X_Format_LoData_LoDefCol
' ==================================================================================================

