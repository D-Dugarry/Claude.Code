Attribute VB_Name = "Rut_Lo_Format_LoData_LoDefCol"
Option Explicit

        Sub Rut_Lo_Format_LoData_LoDefColData_ByHand()  '--- Solo hace falta cambiar las variables de la Rutina -----------------------
            Call Rut_Lo_Format_LoData_LoDefColData(Sht__BD.ListObjects(1), Prog_DefCol.ListObjects(1))
        End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Formatea una Tabla Listobject ----------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Format_LoData_LoDefColData(ByRef LoData As ListObject, _
                                      ByRef LoDefCol As ListObject)
'-----------------------------------------------------------------------------------------------------------------------------------
Debug.Print ">>> Rut_Lo_Format_LoData_LoDefColData"
    '- Setting ListObjects ------------------------------------
    LoDefCol.ListColumns(DefC_FormatCol).TotalsCalculation = xlTotalsCalculationSum
    Dim CantFormatCol       As Integer:     CantFormatCol = 0
    Dim TotCantFormatCol    As Integer:     TotCantFormatCol = LoData.Range.Columns.Count
    Dim TxT_Progreso        As String
    Dim TxT_Prog            As String
    Dim Ccol                As Integer
    Dim Ws_Data             As Worksheet:   Set Ws_Data = LoData.Parent
    Dim HiddenCol           As Boolean

    If LoData.DataBodyRange Is Nothing Then
        MsgBox "¡¡¡ Tabla SIN DATOS !!!", vbOKOnly, "Proceso: Formatear Tabla ListObjects"
        GoTo ExitSub
    End If
    '--- Quitar todo formato -----------------------------------
    LoData.DataBodyRange.Select
    With Selection
        ' Eliminar comentarios (opcional: descomenta si lo deseas)
        .Cells.ClearComments
        ' Eliminar toda la validación de datos
        On Error Resume Next ' Por si no hay validación
        .Cells.Validation.Delete
        On Error GoTo 0
        ' Eliminar formato condicional
        .Cells.FormatConditions.Delete
        ' Eliminar todo el formato (incluyendo fuentes, colores, bordes, alineación, etc.)
        .Cells.ClearFormats
        ' Restaurar ancho/alto predeterminado
        .Columns.AutoFit
        .Rows.AutoFit
        ' Opcional: restablecer formato numérico estándar
        .Cells.NumberFormat = "General"
    End With
    
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            TxT_Progreso = ActivForm.Controls("TBx_Informe")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Formateando el Excel. ", 0, , , TxT_Prog, , , 4)
            TxT_Prog = ActivForm.Controls("TBx_Informe")
    '--- Formatear las Columnas --------------------------------------------------------------------------------------
    For Ccol = 1 To Application.Min(LoData.Range.Columns.Count, LoDefCol.ListRows.Count)
        If Not LoDefCol.DataBodyRange.Cells(Ccol, DefC_FormatCol) Then GoTo NextCol    '- Sólo si se desea formatear la Col.
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            CantFormatCol = CantFormatCol + 1
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "___________________ Formateando Col. " & Ccol & " (" & CantFormatCol & "ª.), de " & TotCantFormatCol & " Col.", 0, , , TxT_Prog, , , 2)
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
                        '- Convierto a números, texto con formato punto de millares y coma decimal. -------------------------------
                        Selection.TextToColumns Destination:=Range(.Cells(1, Ccol).Address(False, False)), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                            Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
                            FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
                    End If
'                    Dim Rc As Range '--- Si es un número muy grande lo muestra como 99999E+12, con el For-Next lo quitamos ------
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
        
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Formateadas: ", LastTimeLap, CantFormatCol & " de " & TotCantFormatCol & " col.", , TxT_Progreso)
        
ExitSub:
Debug.Print "<<< Rut_Lo_Format_LoData_LoDefColData"
End Sub     ' Rut_Lo_Format_LoData_LoDefColData
' ==================================================================================================================================

Sub kk()
        Dim HiddenCol   As Boolean
        HiddenCol = True
        ActiveSheet.Columns(ActiveSheet.ListObjects(1).ListColumns(6).Range.Column).Hidden = HiddenCol '- Para Ocultar Col.
End Sub

'''' ==================================================================================================================================
'''Sub Rut_Format_LoData_LoDefCol(ByRef LoData As ListObject, _
'''                               ByRef LoDefCol As ListObject)
'''' ==================================================================================================================================
'''Dim Ccol    As Integer
'''
'''    If LoData.DataBodyRange Is Nothing Then
'''        MsgBox "¡¡¡ Tabla SIN DATOS !!!", vbOKOnly, "Proceso: Formatear Tabla ListObjects"
'''        GoTo ExitSub
'''    End If
'''
'''    '--- Quitar todo formato -----------------------------------
'''    LoData.DataBodyRange.Select
'''    With Selection
'''        ' Eliminar comentarios (opcional: descomenta si lo deseas)
'''        .Cells.ClearComments
'''        ' Eliminar toda la validación de datos
'''        On Error Resume Next ' Por si no hay validación
'''        .Cells.Validation.Delete
'''        On Error GoTo 0
'''        ' Eliminar formato condicional
'''        .Cells.FormatConditions.Delete
'''        ' Eliminar todo el formato (incluyendo fuentes, colores, bordes, alineación, etc.)
'''        .Cells.ClearFormats
'''        ' Restaurar ancho/alto predeterminado
'''        .Columns.AutoFit
'''        .Rows.AutoFit
'''        ' Opcional: restablecer formato numérico estándar
'''        .Cells.NumberFormat = "General"
'''    End With
'''
'''    '--- Formatear las Columnas --------------------------------------------------------------------------------------
'''    For Ccol = 1 To LoData.Range.Columns.Count
'''        If Not LoDefCol.DataBodyRange.Cells(Ccol, DefC_FormatCol) Then GoTo NextCol    '- Sólo si se desea formatear la Col.
'''        Select Case LoDefCol.DataBodyRange.Cells(Ccol, DefC_TipVar)
'''            Case "T"
'''                    LoData.DataBodyRange.Columns(Ccol).Select
'''                With Selection
'''                    .NumberFormat = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Format)
''''                    .ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht)
''''                    .HorizontalAlignment = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Align).Value2
'''                    .WrapText = LoDefCol.DataBodyRange.Cells(Ccol, DefC_WrapTxt).Value2
'''                End With
'''            Case "F"
'''               Call Rut_Format_Date_Col(LoData, LoDefCol, Ccol)
'''            Case "N"
'''               With LoData.DataBodyRange
'''                    .Columns(Ccol).Select
'''                    If WorksheetFunction.CountA(.Columns(Ccol)) > 0 Then
'''                        Selection.TextToColumns Destination:=Range(.Cells(1, Ccol).Address(False, False)), DataType:=xlDelimited, _
'''                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'''                            Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
'''                            FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
'''                    End If
''''                    Dim Rc As Range '--- Si es un número muy grande lo muestro como 99999E+12, con el For-Next lo quitamos ------
''''                    For Each Rc In .Columns(1)
''''                        With Rc.Cells(1)
''''                            If IsNumeric(.Value2) And .Text Like "*E+*" Then
''''                                Rc.NumberFormat = "#"
''''                            End If
''''                        End With
''''                    Next
'''                    With Selection
'''                        .NumberFormat = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Format)
''''                        .ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht)
''''                        .HorizontalAlignment = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Align).Value2
'''                    End With
'''                End With
'''            Case Else
'''                MsgBox "Error en Tipo de Variable, NO es T,F ó N ???", vbExclamation, "Procedimiento: Formatear Tabla."
'''        End Select
'''NextCol:
'''        LoData.Range.Columns(Ccol).Hidden = LoDefCol.DataBodyRange.Cells(Ccol, DefC_HiddenCol)  '- Para Ocultar Col.
'''        LoData.Range.ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht)
'''        LoData.Range.HorizontalAlignment = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Align).Value2
'''    Next
'''
'''ExitSub:
'''Debug.Print "<<< Rut_Lo_Format_LoData_LoDefColData"
'''End Sub     ' Rut_X_Format_LoData_LoBjConFig
'''' ----------------------------------------------------------------------------------------------------------------------------------


