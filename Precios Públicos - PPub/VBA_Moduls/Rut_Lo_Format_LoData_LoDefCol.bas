Attribute VB_Name = "Rut_Lo_Format_LoData_LoDefCol"
' Last Rev. 2026-10-04 12:32
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
                    If WorksheetFunction.CountA(.Columns(Ccol)) > 0 Then
                        '- Convierto a fechas los textos de fecha, celda a celda en RAM: la Col. puede mezclar fechas y textos.
                        '- Sustituye a la Col. auxiliar "=RC[-1]*1" + pegar valores + Replace "0" (5,8 seg. las 3 Col.).
                        Call Rut_Col_Textos_a_Fechas(.Columns(Ccol))
                    End If
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

'- -------------------------------------------------------------------------------------------------
'- Convierte en fecha (nº de serie de Excel) los textos de fecha de una Col., celda a celda en RAM. Sustituye a la Col. auxiliar
'- "=RC[-1]*1" + pegar valores + Replace de "0" por "", con el mismo resultado salvo la hora, que se quita:
'-      - textos "dd/mm/aaaa hh:mm:ss" (los del Robot y los de siempre), con o sin hora y con "/" o "-"; también "aaaa-mm-dd"
'-        y los años de 2 cifras (00-29 -> 20xx, 30-99 -> 19xx, como Excel)
'-      - SIN LA HORA: no se usa, y con ella un recibo del 31/12 a las 10:00 cuenta como posterior al cierre del 31/12
'-        (M_111: F_Emi > cierre se borra, F_Cob > cierre se vacía)
'-      - celdas vacías, textos vacíos y ceros -> celda vacía (lo que hacía el Replace)
'-      - los números (fechas ya convertidas) se quedan como están, quitándoles la hora si la traen
'- Un texto que no es una fecha se queda como texto (con "*1" daba #¡VALOR!). Cada texto distinto se convierte una sola vez: en
'- un LSGES04 de 184.000 rec. hay unos 7.000 textos distintos en F_Emi y menos de 600 en F_Vto y en F_Cob.
'- -------------------------------------------------------------------------------------------------
Private Sub Rut_Col_Textos_a_Fechas(Rng As Range)
    Dim Datos           As Variant
    Dim Fila            As Long
    Dim Txt             As String
    Dim Serie           As Double
    Dim Cambios         As Long
    Dim DiccFechas      As Object:      Set DiccFechas = CreateObject("Scripting.Dictionary")     '- Texto -> fecha o Empty

    If Rng.Cells.CountLarge = 1 Then
        ReDim Datos(1 To 1, 1 To 1)
        Datos(1, 1) = Rng.Value2
    Else
        Datos = Rng.Value2
    End If
    For Fila = 1 To UBound(Datos, 1)
        Select Case VarType(Datos(Fila, 1))
            Case vbString
                Txt = Datos(Fila, 1)
                If Not DiccFechas.Exists(Txt) Then
                    If Fnc_Texto_a_Fecha(Txt, Serie) Then
                        If Serie = 0 Then DiccFechas.Add Txt, Empty Else DiccFechas.Add Txt, Serie
                    End If
                End If
                If DiccFechas.Exists(Txt) Then                  '- Si no está, no es una fecha: se queda como texto
                    Datos(Fila, 1) = DiccFechas.Item(Txt)
                    Cambios = Cambios + 1
                End If
            Case vbDouble
                If Datos(Fila, 1) = 0 Then
                    Datos(Fila, 1) = Empty
                    Cambios = Cambios + 1
                ElseIf Datos(Fila, 1) <> Int(Datos(Fila, 1)) Then  '- Fecha con hora: le quito la hora
                    Datos(Fila, 1) = Int(Datos(Fila, 1))
                    Cambios = Cambios + 1
                End If
        End Select
    Next Fila
    If Cambios > 0 Then Rng.Value2 = Datos
End Sub     ' Rut_Col_Textos_a_Fechas
'- -------------------------------------------------------------------------------------------------

'- True (y el nº de serie, sin hora, en Serie) si Txt es una fecha o un texto vacío (Serie = 0) ----
Private Function Fnc_Texto_a_Fecha(ByVal Txt As String, ByRef Serie As Double) As Boolean
    Dim Partes      As Variant
    Dim TxtDia      As String
    Dim TxtMes      As String
    Dim TxtAno      As String
    Dim TxtHora     As String
    Dim PosEsp      As Long
    Dim Dia         As Long
    Dim Mes         As Long
    Dim Ano         As Long
    Dim Fecha       As Date
    Dim i           As Long

    Txt = Trim$(Txt)
    If Txt = "" Then Serie = 0: Fnc_Texto_a_Fecha = True: Exit Function
    PosEsp = InStr(Txt, " ")                                    '- La hora va detrás de un espacio
    If PosEsp > 0 Then
        TxtHora = Trim$(Mid$(Txt, PosEsp + 1))
        Txt = Left$(Txt, PosEsp - 1)
    End If
    Partes = Split(Replace(Txt, "-", "/"), "/")                 '- Fecha: 3 partes separadas por "/" o "-"
    If UBound(Partes) <> 2 Then Exit Function
    If Len(Partes(0)) = 4 Then                                  '- aaaa-mm-dd
        TxtAno = Partes(0):     TxtMes = Partes(1):     TxtDia = Partes(2)
    Else                                                        '- dd/mm/aaaa o dd/mm/aa
        TxtDia = Partes(0):     TxtMes = Partes(1):     TxtAno = Partes(2)
    End If
    If Not (Fnc_Son_Cifras(TxtDia, 2) And Fnc_Son_Cifras(TxtMes, 2) And Fnc_Son_Cifras(TxtAno, 4)) Then Exit Function
    Dia = CLng(TxtDia):     Mes = CLng(TxtMes):     Ano = CLng(TxtAno)
    Select Case Len(TxtAno)
        Case 4
        Case 2:     Ano = Ano + IIf(Ano < 30, 2000, 1900)       '- Años de 2 cifras, como Excel
        Case Else:  Exit Function
    End Select
    If Ano < 1900 Or Mes < 1 Or Mes > 12 Or Dia < 1 Or Dia > 31 Then Exit Function
    Fecha = DateSerial(Ano, Mes, Dia)
    If Day(Fecha) <> Dia Then Exit Function                     '- Día que no existe (31/02)
    If TxtHora <> "" Then                                       '- Hora (hh:mm o hh:mm:ss): se comprueba, pero no se guarda
        Partes = Split(TxtHora, ":")
        If UBound(Partes) < 1 Or UBound(Partes) > 2 Then Exit Function
        For i = 0 To UBound(Partes)
            If Not Fnc_Son_Cifras(Partes(i), 2) Then Exit Function
            If CLng(Partes(i)) > IIf(i = 0, 23, 59) Then Exit Function
        Next i
    End If
    Serie = CDbl(Fecha)                                         '- Sin la hora
    If Serie < 61 Then Serie = Serie - 1                        '- Excel cuenta el 29/02/1900, que no existió
    Fnc_Texto_a_Fecha = True
End Function    ' Fnc_Texto_a_Fecha

'- True si Txt son solo cifras, de 1 a LenMax ------------------------------------------------------
Private Function Fnc_Son_Cifras(ByVal Txt As String, ByVal LenMax As Long) As Boolean
    Fnc_Son_Cifras = Len(Txt) >= 1 And Len(Txt) <= LenMax And Not Txt Like "*[!0-9]*"
End Function    ' Fnc_Son_Cifras
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


