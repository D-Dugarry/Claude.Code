Attribute VB_Name = "M_029_Import_N43_to_Tb_N43"
'- M_029_Import_N43_to_Tb_N43
Option Explicit

Const Lcab                      As Integer = 6     ' Línea de Cabecera de las Línea Resultante con datos
Dim Finalizar_Proceso           As Boolean
Dim Posición_Left_Botón         As Integer
Const ConsT_Posisción_Izq       As Integer = 610
Const ConsT_Posisción_Dcha      As Integer = 693

Sub Rut_Extraer_Norma43()

Dim Registro        As String           ' 1º y2º carácter, para identificar el tipo de línea de datos
Dim Indice          As String           ' 3º y 4º carácter, para identificar de las líneas nº23, el orden de los Reg. Complementarios del Movimiento
Dim LineaN43        As String           ' Toda la línea de datos
Dim Hoja_Activa     As String
Dim Anualidad       As String           ' Para Controlar el cambio de años el en número de orden que genero

Dim Ld              As Integer          ' Línea de Detalle
Dim Ultima_Ld       As Integer
Dim Lr              As Integer          ' Línea de Resultante
Const Cr            As Integer = 1      ' Primera Columna de Línea Resultante con datos

Dim Sw_Canceled     As Boolean

Dim Num_Orden       As Long
Dim Hora_Inicio     As Long             ' Para Saber el tiempo de proceso


Application.Calculation = xlCalculationManual   'Evita que se recalcule todo cada vez que se pegan o modifican datos
Application.EnableEvents = False                ' Eventos de pantalla ejecutados por macros en la hoja y libro son desactivados
Application.DisplayAlerts = False               'True si Microsoft Excel muestra alertas y mensajes determinados mientras se ejecuta una macro. Boolean de lectura y escritura.
Application.DisplayFormulaBar = False           'ocultar barra de fórmulas
Application.ScreenUpdating = False              'Para que no se vea en pantalla la apertura del libro (Tarda menos)
ActiveSheet.DisplayPageBreaks = False           ' Sirve para evitar algunos problemas de compatibilidad entre macros Excel 2003 vs. 2007/2010
On Error Resume Next
    Hora_Inicio = Timer                ' Para Saber el tiempo de proceso
    Rut_Off_Functions
    Prog_N43_TxT.Visible = xlSheetVisible
    Prog_N43_CTA.Visible = xlSheetVisible
    ' ---------------        Importar Fichero de Texto Norma43     -------------------------------'
    Call Rut_Importar_Fichero_Norma43(Sw_Canceled)    '- Si Sw_Canceled=True el proceso ha sido Cancelado
    If Sw_Canceled Then GoTo Restablecer_Valores
    
    Prog_N43_CTA.Visible = xlSheetVisible
    Prog_N43_CTA.Unprotect
    Prog_N43_CTA.Select
    
    RuT_Cells_Clear

Range("e1") = "Saldo Inic."
Range("e2") = "Ap.D.:"
Range("e3") = "Ap.H.:"
Range("e4") = "Saldo Final"

Cells(Lcab, 1) = "Orden"
Cells(Lcab, 2) = "Oficina"
Cells(Lcab, 3) = "F. Operación"
Cells(Lcab, 4) = "F.Valor"
Cells(Lcab, 5) = "_ Importe _"
Cells(Lcab, 6) = "_ Saldo _"

Cells(Lcab, 7) = "Documento"
Cells(Lcab, 8) = "Referencia 1"
Cells(Lcab, 9) = "Registro Complementario de Movimiento 01"
Cells(Lcab, 10) = "Registro Complementario de Movimiento 02"
Cells(Lcab, 11) = "Registro Complementario de Movimiento 03"
Cells(Lcab, 12) = "Registro Complementario de Movimiento 04"
Cells(Lcab, 13) = "Registro Complementario de Movimiento 05"

Anualidad = Mid(LineaN43, 21, 2)
Num_Orden = 0
Lr = Lcab
Ultima_Ld = Prog_N43_TxT.Cells(Rows.Count, 1).End(xlUp).Row

For Ld = 1 To Ultima_Ld ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    LineaN43 = Prog_N43_TxT.Cells(Ld, 1)
    Registro = Left(LineaN43, 2)
    Select Case Registro    ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Case 11
        Range("h1") = "Bco.:  " & Mid(LineaN43, 3, 4) & ",     Entidad:  " & Mid(LineaN43, 8, 4)      'Bco. y Entidad
        Range("h2") = "Cta.:  " & Mid(LineaN43, 11, 10)         'Cta.
        Range("h3") = "F.Inicio:  " & DateSerial(Mid(LineaN43, 21, 2), Mid(LineaN43, 23, 2), Mid(LineaN43, 25, 2))   'F.Ini.
        Range("h4") = "F.Inicio:  " & DateSerial(Mid(LineaN43, 27, 2), Mid(LineaN43, 29, 2), Mid(LineaN43, 31, 2))    'F.Fin
        If Mid(LineaN43, 33, 1) = 1 Then
            Range("f1") = Val(Mid(LineaN43, 34, 14)) * -0.01 'Saldo Inic.
        Else
            Range("f1") = Val(Mid(LineaN43, 34, 14)) * 0.01    'Saldo Inic.
        End If
        Range("i2") = Mid(LineaN43, 52, 50)  'Empresa C.C.
    Case 22
        If Anualidad <> Mid(LineaN43, 17, 2) Then Num_Orden = 0: Anualidad = Mid(LineaN43, 17, 2)
        Lr = Lr + 1
        Num_Orden = Num_Orden + 1
        Cells(Lr, Cr) = "20" & Mid(LineaN43, 17, 2) & Right("00000" & Num_Orden, 6)
'        Cells(Lr, Cr) = "20" & Mid(LineaN43, 17, 2) & Right("00000" & Lr - Lcab, 6)
        Cells(Lr, Cr + 1) = Mid(LineaN43, 7, 4)       'Oficina
        Cells(Lr, Cr + 2) = DateSerial(Mid(LineaN43, 11, 2), Mid(LineaN43, 13, 2), Mid(LineaN43, 15, 2))        'F.Op.
        Cells(Lr, Cr + 3) = DateSerial(Mid(LineaN43, 17, 2), Mid(LineaN43, 19, 2), Mid(LineaN43, 21, 2))        'F.Valor
        If Mid(LineaN43, 28, 1) = 1 Then
            Cells(Lr, Cr + 4) = Val(Mid(LineaN43, 29, 14)) * -0.01          'Importe
        Else
            Cells(Lr, Cr + 4) = Val(Mid(LineaN43, 29, 14)) * 0.01       'Importe
        End If
        Cells(Lr, Cr + 4).NumberFormat = "#,##0.00;[Red]-#,##0.00;0"
        If Lr - 1 = Lcab Then
            Cells(Lr, Cr + 5) = Range("f1") + Cells(Lr, Cr + 4)
        Else
            Cells(Lr, Cr + 5) = Cells(Lr - 1, Cr + 5) + Cells(Lr, Cr + 4)
        End If
        Cells(Lr, Cr + 5).NumberFormat = "#,##0.00;[Red]-#,##0.00;0"
        Cells(Lr, Cr + 6) = Mid(LineaN43, 43, 10)     'Documento
        Cells(Lr, Cr + 7) = Mid(LineaN43, 53, 50)     'Referencia 1
    Case 23
        Indice = Mid(LineaN43, 3, 2)
        Cells(Lr, Cr + 7 + Indice) = Mid(LineaN43, 5, 75)      'Reg. Complementarios de Mov.
    Case 33
        Range("e2") = Range("e2") & Mid(LineaN43, 21, 5)   'Nº apuntes Debe
        Range("e3") = Range("e3") & Mid(LineaN43, 40, 5)   'Nº apuntes Haber
        Range("f2") = Val(Mid(LineaN43, 26, 14)) * -0.01               ' Total importes Debe
        Range("f3") = Val(Mid(LineaN43, 45, 14)) * 0.01                ' Total importes Haber
        If Mid(LineaN43, 59, 1) = 1 Then
            Range("f4") = Val(Mid(LineaN43, 60, 14)) * -0.01 'Saldo Final
        Else
            Range("f4") = Val(Mid(LineaN43, 60, 14)) * 0.01  'Saldo Final
        End If
    Case Else
    End Select  ' Registro = Left(LineaN43, 2)     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Next Ld ' Ld = 1 To Ultima_Ld     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

With Range("I" & Lcab & ":M" & Ld)  '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    .Replace What:="ORDENANTE DE LA TRANSFERENCIA :", Replacement:="Ordenante:", _
                  LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                  SearchFormat:=False, ReplaceFormat:=False
    .Replace What:="NOMBRE DEL ORDENANTE", Replacement:="Ordenante: ", _
                  LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                  SearchFormat:=False, ReplaceFormat:=False
    .Replace What:="BENEFICIARIO DE LA TRANSFERENCIA :", Replacement:="Beneficiario: ", _
                  LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                  SearchFormat:=False, ReplaceFormat:=False
    .Replace What:="BENEFICIARIO ", Replacement:="Beneficiario: ", _
                  LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                  SearchFormat:=False, ReplaceFormat:=False
    .Replace What:="OBSERVACIONES :", Replacement:="Observaciones: ", _
                  LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                  SearchFormat:=False, ReplaceFormat:=False
    .Replace What:="CONCEPTO/OBSERVACIONES", Replacement:="Concepto/Observaciones: ", _
                  LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                  SearchFormat:=False, ReplaceFormat:=False
    .Replace What:="PAIS DEL BANCO ORDENANTE", Replacement:="País del Banco ordenante", _
                  LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                  SearchFormat:=False, ReplaceFormat:=False
    .Replace What:="BANCO ORDENANTE", Replacement:="Banco ordenante", _
                  LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                  SearchFormat:=False, ReplaceFormat:=False
    .Font.Size = 10
    .Font.Name = "Arial"
    .VerticalAlignment = xlCenter
    End With    ' Range("I" & Lcab & ":M" & Ld)   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    Dim celda As Range      ' Quita los espacios en blanco repetidos >>>>>>>>>>>>>>>>>>>>>>>
        For Each celda In Range("i7:m" & Lr)
            celda.Value = WorksheetFunction.Trim(celda.Value)
        Next        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    Range("f1:f4").NumberFormat = "#,##0.00;[Red]-#,##0.00;0"
    With Range("A" & Lcab & ":M" & Lcab)    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        .EntireRow.RowHeight = 30
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlCenter
        .Font.Size = 12
        .Font.Name = "Arial"
        .Font.ThemeColor = xlThemeColorLight1
        .Font.ThemeColor = xlThemeColorAccent5
        .Font.Bold = True
        .Interior.ThemeColor = xlThemeColorAccent1
        .Font.TintAndShade = 0.799981688894314
    End With    ' Range("A" & Lcab & ":M" & Lcab)   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Columns("a:d").EntireColumn.HorizontalAlignment = xlCenter
    Columns("h").EntireColumn.HorizontalAlignment = xlCenter
         Range("h1:h4").HorizontalAlignment = xlCenter
    With Range("h1:h4").Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .Bold = True
    End With    ' Range("G1:G4").Font
    Rows(5).EntireRow.Hidden = True
    ActiveWindow.FreezePanes = False    ''''
    Range("h7").Select                  ''''
    ActiveWindow.FreezePanes = True     ''''
    If Not ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    Range("A" & Lcab & ":M" & Lcab).EntireColumn.AutoFit
    Columns("g:g").EntireColumn.Hidden = True
    With Range("e1:h4").Interior
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    With Range("i2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
    With Range("i2").Interior
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
    End With
    With Range("i2").Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
        .Bold = True
        .Italic = True
    End With
    Range("i2").Select
    Posición_Left_Botón = ConsT_Posisción_Dcha
    RuT_Mostrar_Botones
    If ActiveSheet.DisplayPageBreaks = True Then ActiveSheet.DisplayPageBreaks = False

    MsgBox "He tardado: " & Timer - Hora_Inicio & "  segundos" & vbCrLf & vbCrLf & vbCrLf & _
           "He Convertido : " & Ultima_Ld & " Líneas de Norma43 en: " & Lr - Lcab & " Apuntes Bancarios.", _
            vbOKOnly, "Importación Fichero Bancario de Extracto en Formato Norma43"

Restablecer_Valores:

Application.StatusBar = False
Application.Calculation = xlCalculationAutomatic
ActiveSheet.DisplayPageBreaks = True
Application.EnableEvents = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.CutCopyMode = False
End Sub

' ==================================================================================================================================
' =====================     RuT_Seleccionar_Tabla     ==============================================================================
' ==================================================================================================================================
Sub RuT_Seleccionar_Tabla()
Dim Ultima_linea As Long
    Ultima_linea = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & Lcab & ":M" & Ultima_linea).Select
End Sub

' ==================================================================================================================================
' =====================  RuT_Quitar_Filtros   ======================================================================================
' ==================================================================================================================================
Sub RuT_Quitar_Filtros()
    ' Ordenar x Tipo_Subvencion(h:h) Descendente, Objetivo_Estrategico(i:i) Ascendente y Num_Formulario(d:d) Ascendente en la hoja Consulta
    Range("A" & Lcab).Sort Key1:=Range("C" & Lcab + 1), Order1:=xlAscending, Header:=xlYes
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Range("i2").Select
End Sub

' ==================================================================================================================================
' =====================  RuT_Ver_Ocultar_Saldo   ===================================================================================
' ==================================================================================================================================
Sub RuT_Ver_Ocultar_Saldo()
    If Columns("f").Hidden = True Then
        Range("d1:e4").Cut Destination:=Range("e1:f4")
        Columns("f").Hidden = False
        Posición_Left_Botón = ConsT_Posisción_Dcha
        RuT_Mostrar_Botones
        Else
        Range("E1:F4").Cut Destination:=Range("D1:E4")
'        Range("D1:E4").Select
        Columns("f").Hidden = True
        Posición_Left_Botón = ConsT_Posisción_Izq
        RuT_Mostrar_Botones
    End If
    Range("i2").Select
End Sub

' ==================================================================================================================================
' =====================  RuT_Reiniciar   ===========================================================================================
' ==================================================================================================================================
Sub RuT_Reiniciar()
    RuT_Cells_Clear
    RuT_Inicializar
End Sub

' ==================================================================================================================================
' =====================  RuT_Cells_Clear   =========================================================================================
' ==================================================================================================================================
Sub RuT_Cells_Clear()
    With Cells
        .Clear
        .Font.Name = "Arial"
    End With
End Sub

' ==================================================================================================================================
' =====================  RuT_Mostrar_Botones   =====================================================================================
' ==================================================================================================================================
Sub RuT_Mostrar_Botones()
'Place image next to the upper left corner
ActiveSheet.Shapes("Botón_Quitar_Filtros").Left = Posición_Left_Botón
ActiveSheet.Shapes("Botón_Quitar_Filtros").Top = 33
ActiveSheet.Shapes("Botón_Seleccionar_Tabla").Left = ActiveSheet.Shapes("Botón_Quitar_Filtros").Left _
                                                    + ActiveSheet.Shapes("Botón_Quitar_Filtros").Width + 10
ActiveSheet.Shapes("Botón_Seleccionar_Tabla").Top = 33
ActiveSheet.Shapes("Botón_Saldo").Left = ActiveSheet.Shapes("Botón_Quitar_Filtros").Left _
                                                    + ActiveSheet.Shapes("Botón_Quitar_Filtros").Width + 10 _
                                                    + ActiveSheet.Shapes("Botón_Seleccionar_Tabla").Width + 10
ActiveSheet.Shapes("Botón_Saldo").Top = 33
ActiveSheet.Shapes("Botón_Reiniciar").Left = ActiveSheet.Shapes("Botón_Quitar_Filtros").Left _
                                                    + ActiveSheet.Shapes("Botón_Quitar_Filtros").Width + 10 _
                                                    + ActiveSheet.Shapes("Botón_Seleccionar_Tabla").Width + 10 _
                                                    + ActiveSheet.Shapes("Botón_Saldo").Width + 90
ActiveSheet.Shapes("Botón_Reiniciar").Top = 3
'Show image
ActiveSheet.Shapes("Botón_Quitar_Filtros").Visible = True
ActiveSheet.Shapes("Botón_Seleccionar_Tabla").Visible = True
ActiveSheet.Shapes("Botón_Saldo").Visible = True
ActiveSheet.Shapes("Botón_Reiniciar").Visible = True
End Sub

' ==================================================================================================================================
' =====================  RuT_Ocultar_Botones   =====================================================================================
' ==================================================================================================================================
Sub RuT_Ocultar_Botones()
    ActiveSheet.Shapes("Botón_Quitar_Filtros").Visible = False
    ActiveSheet.Shapes("Botón_Seleccionar_Tabla").Visible = False
    ActiveSheet.Shapes("Botón_Saldo").Visible = False
    ActiveSheet.Shapes("Botón_Reiniciar").Visible = False
End Sub

' ==================================================================================================================================
' =====================  Rut_Quitar_Ribbon   =======================================================================================
' ==================================================================================================================================
Sub Rut_Quitar_Ribbon()
    If CommandBars("Ribbon").Controls(1).Height > 100 Then CommandBars.ExecuteMso ("MinimizeRibbon")
End Sub

' ==================================================================================================================================
' =====================  RuT_Inicializar   =========================================================================================
' ==================================================================================================================================
Sub RuT_Inicializar()
    Rut_Quitar_Ribbon
If IsEmpty(Range("B6")) Then
    RuT_Cells_Clear
    ActiveWindow.DisplayGridlines = False
    ActiveSheet.DisplayPageBreaks = False
    ActiveWindow.FreezePanes = False

    RuT_Ocultar_Botones
    With Range("B6")
        .Value = "Al pulsar el botón:"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = vbRed
        .Font.Underline = True
    End With
    With Range("B7")
        .Value = "1.- Abre el explorador para seleccionar un fichero de extracto bancario, con formato Norma43."
        .Characters(1, 2).Font.Bold = True
        .Characters(1, 2).Font.Color = vbRed
    End With
    With Range("B8")
        .Value = "2.- Lo importa en una hoja de cálulo que llamará ""Norma43""."
        .Characters(1, 2).Font.Bold = True
        .Characters(1, 2).Font.Color = vbRed
    End With
    With Range("B9")
        .Value = "3.- Extrae los datos en esta hoja ""Extracto"", datos generales (zona superior) y de movimientos (Tabla)."
        .Characters(1, 2).Font.Bold = True
        .Characters(1, 2).Font.Color = vbRed
    End With
    With Range("B10")
        .Value = "4.- Añade una columna con un número de orden que crea para cada movimiento bancario."
        .Characters(1, 2).Font.Bold = True
        .Characters(1, 2).Font.Color = vbRed
    End With
    With Range("B11")
        .Value = "5.- Formatea las columnas de fecha operación, fecha valor e importe."
        .Characters(1, 2).Font.Bold = True
        .Characters(1, 2).Font.Color = vbRed
    End With
    With Range("B12")
        .Value = "6.- Activa los filtros."
        .Characters(1, 2).Font.Bold = True
        .Characters(1, 2).Font.Color = vbRed
    End With
    With Range("B13")
        .Value = "7.- Habilita un botón para quitar filtros y otro para seleccionar toda la tabla."
        .Characters(1, 2).Font.Bold = True
        .Characters(1, 2).Font.Color = vbRed
    End With
    Range("B6:B13").EntireColumn.AutoFit
    Range("B6").Select

End If  'If IsEmpty(Cells(6, 2)) Then

End Sub     ' RuT_Inicializar



