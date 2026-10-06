Attribute VB_Name = "Rut_Lo_Format_LoData_LoDefCol"
' Last Rev. 2026-10-06 11:10
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Formatea una Tabla Listobject ----------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Format_LoData_LoDefColData(ByRef LoData As ListObject, _
                                      ByRef LoDefCol As ListObject, _
                                      Optional ByVal Convertir As Boolean = True)
'- Convertir = False (2026-10-06, fase 4 del paso a RAM): las Col. N y F ya vienen convertidas, porque la importación las
'- convirtió en RAM (Rut_Lo_Import_LoData_LoDefCol con Convertir_En_Ram): aquí solo se les pone el formato.
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
    Dim UltCol              As Integer:     UltCol = Application.Min(LoData.Range.Columns.Count, LoDefCol.ListRows.Count)
    Dim N_Hasta             As Integer                                      '- Última Col. del bloque de Col. N ya convertido
    Dim TxT_Bloque          As String                                       '- Para el informe de tiempos de las Col. N

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
    For Ccol = 1 To UltCol
        If Not LoDefCol.DataBodyRange.Cells(Ccol, DefC_FormatCol) Then GoTo NextCol    '- Sólo si se desea formatear la Col.
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            CantFormatCol = CantFormatCol + 1
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "___________________ Formateando Col. " & Ccol & " (" & CantFormatCol & "ª.), de " & TotCantFormatCol & " Col.", 0, , , TxT_Prog, , , 2)
        T_Paso = Timer
        TxT_Bloque = ""
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
                    If Convertir And WorksheetFunction.CountA(.Columns(Ccol)) > 0 Then
                        '- Convierto a fechas los textos de fecha, celda a celda en RAM: la Col. puede mezclar fechas y textos.
                        '- Sustituye a la Col. auxiliar "=RC[-1]*1" + pegar valores + Replace "0" (5,8 seg. las 3 Col.).
                        Call Rut_Col_Textos_a_Fechas(.Columns(Ccol))
                    End If
                    Selection.NumberFormat = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Format)
               End With
            Case "N"
               With LoData.DataBodyRange
                    '- Las Col. N seguidas se convierten de una vez, con una lectura y una escritura para todo el bloque: la 1ª Col.
                    '- del bloque las convierte todas y las demás solo ponen su formato. (Col. a Col., las 11 Col. N tardaban 7,5 seg.)
                    If Not Convertir Then
                        '- Ya convertida en RAM al importar: solo el formato
                    ElseIf Ccol > N_Hasta Then
                        N_Hasta = Ccol
                        Do While N_Hasta < UltCol
                            If Not LoDefCol.DataBodyRange.Cells(N_Hasta + 1, DefC_FormatCol) Then Exit Do
                            If LoDefCol.DataBodyRange.Cells(N_Hasta + 1, DefC_TipVar) <> "N" Then Exit Do
                            N_Hasta = N_Hasta + 1
                        Loop
                        If WorksheetFunction.CountA(.Columns(Ccol).Resize(, N_Hasta - Ccol + 1)) > 0 Then
                            '- Convierto a números los textos numéricos, celda a celda en RAM: la Col. puede mezclar números y textos.
                            '- Sustituye a TextToColumns, que con la config. española deja "443.73" como texto y lee "3867.9250" como 38.679.250.
                            Call Rut_Cols_Textos_a_Numeros(.Columns(Ccol).Resize(, N_Hasta - Ccol + 1))
                        End If
                        If N_Hasta > Ccol Then TxT_Bloque = ", bloque " & Ccol & "-" & N_Hasta
                    Else
                        TxT_Bloque = ", ya convertida en su bloque"
                    End If
'                    Dim Rc As Range '--- Si es un número muy grande lo muestra como 99999E+12, con el For-Next lo quitamos ------
'                    For Each Rc In .Columns(1)
'                        With Rc.Cells(1)
'                            If IsNumeric(.Value2) And .Text Like "*E+*" Then
'                                Rc.NumberFormat = "#"
'                            End If
'                        End With
'                    Next
                    .Columns(Ccol).NumberFormat = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Format)
                End With
            Case Else
                MsgBox "Error en Tipo de Variable, NO es T,F ó N ???", vbExclamation, "Procedimiento: Formatear Tabla."
        End Select
        T_Paso = Timer - T_Paso
        T_Cols = T_Cols + T_Paso
        Tiempos.Add Array("   Col. " & Ccol & " " & LoData.ListColumns(Ccol).Name & " (" & LoDefCol.DataBodyRange.Cells(Ccol, DefC_TipVar) & TxT_Bloque & ")", T_Paso)
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
'- Convierte en RAM, con las mismas reglas que el formateo, las Col. N y F (FormatCol = VERDADERO en LoDefCol) de Datos: las
'- filas de un fichero leídas con .Value2, antes de escribirlas en la tabla (Rut_Lo_Import_LoData_LoDefCol con Convertir_En_Ram).
'- Reescrita(Col) = True si la Col. ha cambiado: el formateo la habría vuelto a escribir en la hoja, y al escribirla Excel
'- interpreta los textos que queden en ella ("1/2" pasa a fecha); las Col. sin cambios no las tocaba. Escribiendo las Reescritas
'- tal cual y las demás con formato texto, la tabla queda igual que pegando valores y convirtiendo en la hoja (comprobado el
'- 2026-10-06 con el LSGES04 del Robot: 184.251 x 29 celdas iguales en valor y tipo).
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Ram_Textos_a_Numeros_y_Fechas(Datos As Variant, LoDefCol As ListObject, Reescrita() As Boolean)
    Dim Col             As Long
    Dim ConTextos       As Boolean

    ReDim Reescrita(1 To UBound(Datos, 2))
    For Col = 1 To Application.Min(UBound(Datos, 2), LoDefCol.ListRows.Count)
        If LoDefCol.DataBodyRange.Cells(Col, DefC_FormatCol) Then
            Select Case LoDefCol.DataBodyRange.Cells(Col, DefC_TipVar)
                Case "N":   Reescrita(Col) = (Fnc_Ram_Col_Textos_a_Numeros(Datos, Col, ConTextos) > 0)
                Case "F":   Reescrita(Col) = (Fnc_Ram_Col_Textos_a_Fechas(Datos, Col) > 0)
            End Select
        End If
    Next Col
End Sub     ' Rut_Ram_Textos_a_Numeros_y_Fechas
' ==================================================================================================================================

'- ----------------------------------------------------------------------------------------------------------------------------
'- Convierte en número los textos numéricos de un bloque de Col. seguidas, celda a celda en RAM, con una sola lectura y, casi
'- siempre, una sola escritura para todo el bloque. Cada Col. puede mezclar números y textos: los números se dejan como están.
'- Entiende los dos formatos de texto que llegan:
'-      - el del Robot (Robot_PPub_Fusión): punto decimal y sin millares    -> "443.73", "3867.9250", "-300"
'-      - el español: coma decimal y punto de millares                      -> "1.234,56", "-300,00"
'- Se decide en cada Col.: si algún texto numérico de la Col. lleva coma, la coma es el decimal y el punto los millares; si
'- ninguno la lleva, el punto es el decimal. El signo "-" puede ir delante o detrás ("300-"), como con el TrailingMinusNumbers de
'- TextToColumns, y los textos vacíos quedan vacíos. Los textos que no son números ("FLY", un nombre de fichero...) se quedan como
'- están.
'- Se escribe el bloque entero salvo que alguna Col. sin cambios tenga textos: entonces solo las Col. con cambios, porque al
'- reescribir un texto Excel lo interpreta como si se tecleara ("1/2" pasaría a fecha). Así queda igual que convirtiendo Col. a Col.
'- ----------------------------------------------------------------------------------------------------------------------------
Private Sub Rut_Cols_Textos_a_Numeros(Rng As Range)
    Dim Datos           As Variant
    Dim ColDatos        As Variant
    Dim Fila            As Long
    Dim Col             As Long
    Dim NumCols         As Long
    Dim Cambios()       As Long                             '- (1 To NumCols): celdas convertidas en cada Col.
    Dim ConTextos()     As Boolean                          '- (1 To NumCols): la Col. conserva textos que no son números
    Dim TotCambios      As Long
    Dim Bloque          As Boolean

    If Rng.Cells.CountLarge = 1 Then
        ReDim Datos(1 To 1, 1 To 1)
        Datos(1, 1) = Rng.Value2
    Else
        Datos = Rng.Value2
    End If
    NumCols = UBound(Datos, 2)
    ReDim Cambios(1 To NumCols)
    ReDim ConTextos(1 To NumCols)
    For Col = 1 To NumCols
        Cambios(Col) = Fnc_Ram_Col_Textos_a_Numeros(Datos, Col, ConTextos(Col))
        TotCambios = TotCambios + Cambios(Col)
    Next Col
    If TotCambios = 0 Then Exit Sub

    Bloque = True
    For Col = 1 To NumCols
        If Cambios(Col) = 0 And ConTextos(Col) Then Bloque = False
    Next Col
    If Bloque Then
        Rng.Value2 = Datos
    Else
        ReDim ColDatos(1 To UBound(Datos, 1), 1 To 1)
        For Col = 1 To NumCols
            If Cambios(Col) > 0 Then
                For Fila = 1 To UBound(Datos, 1)
                    ColDatos(Fila, 1) = Datos(Fila, Col)
                Next Fila
                Rng.Columns(Col).Value2 = ColDatos
            End If
        Next Col
    End If
End Sub     ' Rut_Cols_Textos_a_Numeros

'- Convierte en número los textos numéricos de la Col. Col de Datos, en RAM, con las reglas de Rut_Cols_Textos_a_Numeros. ------
'- Devuelve cuántas celdas ha cambiado, y en ConTextos si en la Col. quedan textos que no son números.
Private Function Fnc_Ram_Col_Textos_a_Numeros(Datos As Variant, ByVal Col As Long, ByRef ConTextos As Boolean) As Long
    Dim Fila            As Long
    Dim ComaDecimal     As Boolean
    Dim Num             As Double
    Dim Cambios         As Long

    ConTextos = False
    For Fila = 1 To UBound(Datos, 1)                        '- 1ª pasada: ¿algún texto numérico lleva coma decimal?
        If VarType(Datos(Fila, Col)) = vbString Then
            If InStr(Datos(Fila, Col), ",") > 0 Then
                If Fnc_Texto_a_Numero(Datos(Fila, Col), True, Num) Then ComaDecimal = True: Exit For
            End If
        End If
    Next Fila
    For Fila = 1 To UBound(Datos, 1)                        '- 2ª pasada: conversión
        If VarType(Datos(Fila, Col)) = vbString Then
            If Datos(Fila, Col) = "" Then
                Datos(Fila, Col) = Empty
                Cambios = Cambios + 1
            ElseIf Fnc_Texto_a_Numero(Datos(Fila, Col), ComaDecimal, Num) Then
                Datos(Fila, Col) = Num
                Cambios = Cambios + 1
            Else
                ConTextos = True
            End If
        End If
    Next Fila
    Fnc_Ram_Col_Textos_a_Numeros = Cambios
End Function    ' Fnc_Ram_Col_Textos_a_Numeros
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

    If Rng.Cells.CountLarge = 1 Then
        ReDim Datos(1 To 1, 1 To 1)
        Datos(1, 1) = Rng.Value2
    Else
        Datos = Rng.Value2
    End If
    If Fnc_Ram_Col_Textos_a_Fechas(Datos, 1) > 0 Then Rng.Value2 = Datos
End Sub     ' Rut_Col_Textos_a_Fechas

'- Convierte en fecha los textos de fecha de la Col. Col de Datos, en RAM, con las reglas de Rut_Col_Textos_a_Fechas. ----------
'- Devuelve cuántas celdas ha cambiado.
Private Function Fnc_Ram_Col_Textos_a_Fechas(Datos As Variant, ByVal Col As Long) As Long
    Dim Fila            As Long
    Dim Txt             As String
    Dim Serie           As Double
    Dim Cambios         As Long
    Dim DiccFechas      As Object:      Set DiccFechas = CreateObject("Scripting.Dictionary")     '- Texto -> fecha o Empty

    For Fila = 1 To UBound(Datos, 1)
        Select Case VarType(Datos(Fila, Col))
            Case vbString
                Txt = Datos(Fila, Col)
                If Not DiccFechas.Exists(Txt) Then
                    If Fnc_Texto_a_Fecha(Txt, Serie) Then
                        If Serie = 0 Then DiccFechas.Add Txt, Empty Else DiccFechas.Add Txt, Serie
                    End If
                End If
                If DiccFechas.Exists(Txt) Then                  '- Si no está, no es una fecha: se queda como texto
                    Datos(Fila, Col) = DiccFechas.Item(Txt)
                    Cambios = Cambios + 1
                End If
            Case vbDouble
                If Datos(Fila, Col) = 0 Then
                    Datos(Fila, Col) = Empty
                    Cambios = Cambios + 1
                ElseIf Datos(Fila, Col) <> Int(Datos(Fila, Col)) Then  '- Fecha con hora: le quito la hora
                    Datos(Fila, Col) = Int(Datos(Fila, Col))
                    Cambios = Cambios + 1
                End If
        End Select
    Next Fila
    Fnc_Ram_Col_Textos_a_Fechas = Cambios
End Function    ' Fnc_Ram_Col_Textos_a_Fechas
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


