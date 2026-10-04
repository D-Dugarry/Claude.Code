Attribute VB_Name = "Rut_Lo_Format_LoData_LoDefCol"
' Last Rev. 2026-10-04 11:58
Option Explicit

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
    Dim T_Ini               As Single:      T_Ini = Timer                   '- Tiempos del formateo, para el informe
    Dim T_Paso              As Single
    Dim T_Quitar            As Single                                       '- Quitar formatos y validaciones
    Dim T_Cols              As Single                                       '- Suma de las Col. formateadas
    Dim Tiempos             As Collection:  Set Tiempos = New Collection    '- Array(Texto, Segundos) de cada paso
    Dim Paso                As Variant

    If LoData.DataBodyRange Is Nothing Then
        MsgBox "¡¡¡ Tabla SIN DATOS !!!", vbOKOnly, "Proceso: Formatear Tabla ListObjects"
        GoTo ExitSub
    End If
    T_Paso = Timer
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
        ' Sin AutoFit: los anchos se fijan después desde DefCol (NextCol). El alto de las filas se fija al estándar de la hoja:
        ' con las filas en alto automático, las Col. N tardaban 5 seg. más (Excel recalcula el alto al escribir en ellas).
        .RowHeight = Ws_Data.StandardHeight
        ' Opcional: restablecer formato numérico estándar
        .Cells.NumberFormat = "General"
    End With
    T_Quitar = Timer - T_Paso
    Tiempos.Add Array("   Quitar formatos y validaciones, alto de filas", T_Quitar)
    
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
        T_Paso = Timer
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
                        '- Convierto a números los textos numéricos, celda a celda en RAM: la Col. puede mezclar números y textos.
                        '- Sustituye a TextToColumns, que con la config. española deja "443.73" como texto y lee "3867.9250" como 38.679.250.
                        Call Rut_Col_Textos_a_Numeros(.Columns(Ccol))
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
        T_Paso = Timer - T_Paso
        T_Cols = T_Cols + T_Paso
        Tiempos.Add Array("   Col. " & Ccol & " " & LoData.ListColumns(Ccol).Name & " (" & LoDefCol.DataBodyRange.Cells(Ccol, DefC_TipVar) & ")", T_Paso)
NextCol:
'        Ws_Data.Columns(LoData.ListColumns(Ccol).Range.Column).ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht).Value
        LoData.Range.Columns(Ccol).ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht).Value
        LoData.Range.Columns(Ccol).HorizontalAlignment = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Align).Value
     Next
        
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Formateadas: ", LastTimeLap, CantFormatCol & " de " & TotCantFormatCol & " col.", , TxT_Progreso)
            '- Tiempo de cada paso del formateo (suman el total que da el llamador) -----------------------------------
            Tiempos.Add Array("   Anchos, alineación y resto", (Timer - T_Ini) - T_Quitar - T_Cols)
            For Each Paso In Tiempos
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", CStr(Paso(0)), 0, Format(Paso(1), "0.00") & " seg.")
            Next Paso
        
ExitSub:
Debug.Print "<<< Rut_Lo_Format_LoData_LoDefColData"
End Sub     ' Rut_Lo_Format_LoData_LoDefColData
' ==================================================================================================================================

'- ----------------------------------------------------------------------------------------------------------------------------
'- Convierte en número los textos numéricos de una Col., celda a celda en RAM. La Col. puede mezclar números y textos: los
'- números se dejan como están. Entiende los dos formatos de texto que llegan:
'-      - el del Robot (Robot_PPub_Fusión): punto decimal y sin millares    -> "443.73", "3867.9250", "-300"
'-      - el español: coma decimal y punto de millares                      -> "1.234,56", "-300,00"
'- Si algún texto numérico de la Col. lleva coma, la coma es el decimal y el punto los millares; si ninguno la lleva, el punto
'- es el decimal. El signo "-" puede ir delante o detrás ("300-"), como con el TrailingMinusNumbers de TextToColumns, y los
'- textos vacíos quedan vacíos. Los textos que no son números ("FLY", un nombre de fichero...) se quedan como están.
'- ----------------------------------------------------------------------------------------------------------------------------
Private Sub Rut_Col_Textos_a_Numeros(Rng As Range)
    Dim Datos           As Variant
    Dim Fila            As Long
    Dim ComaDecimal     As Boolean
    Dim Num             As Double
    Dim Cambios         As Long

    If Rng.Cells.CountLarge = 1 Then
        ReDim Datos(1 To 1, 1 To 1)
        Datos(1, 1) = Rng.Value2
    Else
        Datos = Rng.Value2
    End If
    For Fila = 1 To UBound(Datos, 1)                        '- 1ª pasada: ¿algún texto numérico lleva coma decimal?
        If VarType(Datos(Fila, 1)) = vbString Then
            If InStr(Datos(Fila, 1), ",") > 0 Then
                If Fnc_Texto_a_Numero(Datos(Fila, 1), True, Num) Then ComaDecimal = True: Exit For
            End If
        End If
    Next Fila
    For Fila = 1 To UBound(Datos, 1)                        '- 2ª pasada: conversión
        If VarType(Datos(Fila, 1)) = vbString Then
            If Datos(Fila, 1) = "" Then
                Datos(Fila, 1) = Empty
                Cambios = Cambios + 1
            ElseIf Fnc_Texto_a_Numero(Datos(Fila, 1), ComaDecimal, Num) Then
                Datos(Fila, 1) = Num
                Cambios = Cambios + 1
            End If
        End If
    Next Fila
    If Cambios > 0 Then Rng.Value2 = Datos
End Sub     ' Rut_Col_Textos_a_Numeros
'- ----------------------------------------------------------------------------------------------------------------------------

'- True (y el valor en Num) si Txt es un número escrito como texto. ComaDecimal: True = "1.234,56" / False = "1234.56" -------
Private Function Fnc_Texto_a_Numero(ByVal Txt As String, ByVal ComaDecimal As Boolean, ByRef Num As Double) As Boolean
    Dim SepDec      As String
    Dim SepMil      As String
    Dim PosDec      As Long
    Dim Negativo    As Boolean

    Txt = Trim$(Txt)
    If Left$(Txt, 1) = "-" Then
        Negativo = True
        Txt = Mid$(Txt, 2)
    ElseIf Right$(Txt, 1) = "-" Then
        Negativo = True
        Txt = Left$(Txt, Len(Txt) - 1)
    End If
    If Not Txt Like "*#*" Then Exit Function                    '- Sin ninguna cifra
    If Txt Like "*[!0-9.,]*" Then Exit Function                 '- Algo que no es cifra ni separador
    If ComaDecimal Then
        SepDec = ",":   SepMil = "."
    Else
        SepDec = ".":   SepMil = ","
    End If
    PosDec = InStr(Txt, SepDec)
    If PosDec > 0 Then
        If InStr(PosDec + 1, Txt, SepDec) > 0 Then Exit Function    '- Dos separadores decimales
        If InStr(PosDec + 1, Txt, SepMil) > 0 Then Exit Function    '- Millares detrás del decimal
    End If
    Num = Val(Replace(Replace(Txt, SepMil, ""), SepDec, "."))       '- Val usa siempre el punto como decimal, sea cual sea la config.
    If Negativo Then Num = -Num
    Fnc_Texto_a_Numero = True
End Function    ' Fnc_Texto_a_Numero
' ==================================================================================================================================

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


