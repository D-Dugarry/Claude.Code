Attribute VB_Name = "Rut_Lo_Format"
'2025-11-11
Option Explicit

'        Sub Rut_Format_LoData_LoDefCol_ByHand()  '--- Solo hace falta cambiar las variables de la Rutina -----------------------
'            Call Rut_Format_LoData_LoDefCol(Prog_TLSGES04.ListObjects(1), Prog_DefCol.ListObjects(1))
'            Call Rut_Format_LoData_LoDefCol(Prog_TitPH.ListObjects(1), Prog_TitPH_DefCol.ListObjects(1))
'        End Sub
'- Comentado: Prog_TitPH/Prog_TitPH_DefCol no existen en este proyecto (vestigio de otro libro de Titulos Propios),
'- y Rut_Format_LoData_LoDefCol solo sobrevive comentada en Rut_Lo_Format_LoData_LoDefCol.bas. No compilaria.
' ==================================================================================================================================

                    ' ==============================================================================
'                    Sub Rut_Format_Date_Col_ByHand()
'                        Call Rut_Format_Date_Col(ActiveSheet.ListObjects(1), 2)
'                    End Sub
'- Comentado: Rut_Format_Date_Col exige 3 argumentos (Lo_Data, Lo_DefCol, Ccol), aqui solo se pasan 2.
'- Ademas Rut_Format_Date_Col no tiene ningun otro llamador activo en el proyecto (solo esta prueba). No compilaria.
' ==================================================================================================================================
Sub Rut_Format_Date_Col(ByRef Lo_Data As ListObject, ByRef Lo_DefCol As ListObject, Ccol As Integer)  '####################################################
' ==================================================================================================================================

    With Lo_Data.DataBodyRange
        .Columns(Ccol).Select
        Selection.NumberFormat = "General"
        If Lo_Data.ListColumns.Count = Ccol Then
            Lo_Data.ListColumns.Add
        Else
            .Columns(Ccol + 1).Select
            Selection.ListObject.ListColumns.Add Position:=Ccol + 1
        End If
        .Cells(1, Ccol + 1).Select
             ActiveCell.FormulaR1C1 = "=RC[-1]*1"     '- ActiveCell.FormulaR1C1 = "=+[@FECHAEMISION]*1"
        .Columns(Ccol + 1).Select
        Selection.Copy
        .Cells(1, Ccol).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        .Columns(Ccol + 1).Delete
        Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False ', FormulaVersion:=xlReplaceFormula2    '- Quitado porque da error según el ord de 32b o 64b... _
               "Error de compilación en el módulo oculto: <nombre del módulo>" Este error se produce normalmente cuando el código es incompatible con la versión o la arquitectura de esta aplicación (por ejemplo, el código de un documento está dirigido a aplicaciones de Microsoft Office de 32 bits pero se está intentando ejecutar en Office de 64 bits).
        Selection.NumberFormat = Lo_DefCol.DataBodyRange.Cells(Ccol, DefC_Format)
'        Selection.ColumnWidth = Lo_DefCol.DataBodyRange.Cells(Ccol, DefC_Widht)
'        Selection.HorizontalAlignment = Lo_DefCol.DataBodyRange.Cells(Ccol, DefC_Align).Value2
    End With
        
End Sub     ' Rut_Format_Date_Col
' ----------------------------------------------------------------------------------------------------------------------------------

                    ' ==============================================================================
                    Sub Rut_Lo_Col_Format_Date_ByHand()
                        Call Rut_Lo_Col_Format_Date(ActiveSheet.ListObjects(1), 2)
                    End Sub
' ==================================================================================================================================
Sub Rut_Lo_Col_Format_Date(ByRef LoBjDatos As ListObject, Ccol As Integer)     '############## Convierte Número Texto en fecha ###################
    With LoBjDatos.DataBodyRange
        .Columns(Ccol).Select
        Selection.NumberFormat = "General"
        Dim Lin    As Long
        With .Columns(Ccol)
            For Lin = 1 To .Rows.Count
                    .Cells(Lin) = Left(.Cells(Lin), 10)
            Next Lin
        End With
        .Columns(Ccol).Select
        Selection.TextToColumns DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 4), TrailingMinusNumbers:=True
        Selection.NumberFormat = "dd-mm-yyyy"
    End With
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
Sub Rut_Format_Date()     '############## Convierte Número Texto en fecha ###################
With ActiveSheet.ListObjects(1).DataBodyRange
    Columns("A:A").Select
    Selection.NumberFormat = "General"
    Selection.TextToColumns Destination:=Range("A:A"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 4), TrailingMinusNumbers:=True
    Selection.NumberFormat = "dd-mm-yyyy"
    Selection.Value = Selection.Value
End With
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
Sub Rut_Format_Num_BIGs()     '############## Convierte Número Texto Grandes 2345E+10 en Números sin E+ ###################
With ActiveSheet.ListObjects(1).DataBodyRange
    .Columns(1).Select
    Selection.TextToColumns Destination:=Range("a2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

    Dim Rc As Range
    For Each Rc In .Columns(1)
        With Rc.Cells(1)
            If IsNumeric(.Value2) And .Text Like "*E+*" Then
                Rc.NumberFormat = "#"
            End If
        End With
    Next
    Selection.NumberFormat = "0000""_""00000000#"
End With
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
Sub Rut_Format_Num()     '############## Convierte Número Texto Grandes 2345E+10 en Números sin E+ ###################
With ActiveSheet.ListObjects(1).DataBodyRange
    .Columns(1).Select
    Selection.TextToColumns Destination:=Range("a2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.NumberFormat = "General" '- or "0"  od "0.00"
    Selection.Value = Selection.Value
End With
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
Sub Rut_Format_Datos_TextoEnColumnas()
Dim Cont_Col    As Integer
ActiveSheet.Unprotect
Cont_Col = 9
With Prog_H_Borrador.ListObjects(1).DataBodyRange
    .Columns(Cont_Col + 1).Select
    Selection.ListObject.ListColumns.Add Position:=Cont_Col + 1
    .Cells(1, Cont_Col + 1).Select
         ActiveCell.FormulaR1C1 = "=RC[-1]*1"
'    ActiveCell.FormulaR1C1 = "=+[@FECHAEMISION]*1"
    .Columns(Cont_Col + 1).Select
    Selection.Copy
    .Cells(1, Cont_Col).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    .Columns(Cont_Col + 1).Delete
    .Columns(Cont_Col).Select
    Selection.NumberFormat = "General"
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.NumberFormat = "m/d/yyyy"
End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Rut_Colors__Table_Code_HEX_RGB_____()
' https://htmlcolorcodes.com/es/
' https://htmlcolorcodes.com/color-names/
' PowerToysUserSetup-0.80.1-x64  -------- instalar para averiguar colores de la pantalla con el puntero......
End Sub


