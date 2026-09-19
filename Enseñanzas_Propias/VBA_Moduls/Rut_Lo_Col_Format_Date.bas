Attribute VB_Name = "Rut_Lo_Col_Format_Date"
' Módulo: FormateadorFechas
' Descripción: Formatea columnas de fechas en ListObjects con múltiples formatos

Option Explicit

Public Enum enuFormatoFecha
    Formato_DDMMAAAA = 1
    Formato_MMDDAAAA = 2
    Formato_AAAAMMDD = 3
    Formato_Texto = 4
End Enum

Sub kk()
'    Call FormatearColumnaFechas(Prog_LsGes04.ListObjects(1), "FECHAEMIS")
    Call FormatearColumnaFechas(Prog_LsGes04.ListObjects(1), Prog_LsGes04.ListObjects(1).Range(15))
End Sub


' Función principal para formatear columna de fechas
Public Sub FormatearColumnaFechas( _
    ByRef tbl As ListObject, _
    ByVal nombreColumna As String, _
    Optional ByVal formatoDestino As enuFormatoFecha = Formato_DDMMAAAA, _
    Optional ByVal separador As String = "/", _
    Optional ByVal incluirHora As Boolean = False)
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim columna As Range
    Dim Celda As Range
    Dim fechaConvertida As Date
    Dim formatoOrigen As enuFormatoFecha
    Dim estadisticas As Object ' Cambiado a Object
    
    ' Validaciones iniciales
    If tbl Is Nothing Then
        MsgBox "La tabla no es válida", vbExclamation
        Exit Sub
    End If
    
    If Not ColumnaExiste(tbl, nombreColumna) Then
        MsgBox "La columna '" & nombreColumna & "' no existe en la tabla", vbExclamation
        Exit Sub
    End If
    
    Set ws = tbl.Parent
    Set columna = tbl.ListColumns(nombreColumna).DataBodyRange
    Set estadisticas = CreateObject("Scripting.Dictionary") ' Cambiado aquí
    
    ' Deshabilitar actualización de pantalla y cálculos
    Application.ScreenUpdating = False
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Procesar cada celda
    For Each Celda In columna
        If Not IsEmpty(Celda.Value) Then
            formatoOrigen = DetectarFormatoFecha(Celda.Value)
            
            ' Registrar estadísticas
            If estadisticas.Exists(formatoOrigen) Then
                estadisticas(formatoOrigen) = estadisticas(formatoOrigen) + 1
            Else
                estadisticas.Add formatoOrigen, 1
            End If
            
            ' Convertir fecha
            If ConvertirFecha(Celda.Value, formatoOrigen, fechaConvertida) Then
                If incluirHora Then
                    Celda.Value = fechaConvertida
                Else
                    Celda.Value = DateValue(fechaConvertida)
                End If
            Else
                ' Marcar celdas problemáticas en amarillo
                Celda.Interior.Color = RGB(255, 255, 0)
                Celda.Value = "ERROR_FECHA"
            End If
        End If
    Next Celda
    
    ' Aplicar formato final
    AplicarFormatoColumna columna, formatoDestino, separador, incluirHora
    
    ' Mostrar reporte
    MostrarReporteEstadisticas estadisticas, tbl.Name, nombreColumna
    
Cleanup:
    ' Restaurar configuración de Excel
    Application.ScreenUpdating = True
    Application.Calculation = Sw_Calculation
    Call Rut_EnableEvents_Status_Reset
    Set estadisticas = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description & " (Línea: " & Erl & ")", vbCritical
    Resume Cleanup
End Sub

' Detectar formato de fecha automáticamente
Private Function DetectarFormatoFecha(ByVal valor As Variant) As enuFormatoFecha
    
    If IsDate(valor) Then
        DetectarFormatoFecha = Formato_Texto
        Exit Function
    End If
    
    Dim Texto As String
    Texto = CStr(valor)
    
    ' Remover espacios extra
    Texto = Trim(Texto)
    
    ' Patrones comunes
    Dim patronDDMMAAAA As Object
    Dim patronMMDDAAAA As Object
    Dim patronAAAAMMDD As Object
    
    Set patronDDMMAAAA = CreateObject("VBScript.RegExp")
    Set patronMMDDAAAA = CreateObject("VBScript.RegExp")
    Set patronAAAAMMDD = CreateObject("VBScript.RegExp")
    
    With patronDDMMAAAA
        .Pattern = "^(0[1-9]|[12][0-9]|3[01])[/\-\.](0[1-9]|1[0-2])[/\-\.](\d{4})$"
        .Global = False
        .IgnoreCase = True
    End With
    
    With patronMMDDAAAA
        .Pattern = "^(0[1-9]|1[0-2])[/\-\.](0[1-9]|[12][0-9]|3[01])[/\-\.](\d{4})$"
        .Global = False
        .IgnoreCase = True
    End With
    
    With patronAAAAMMDD
        .Pattern = "^(\d{4})[/\-\.](0[1-9]|1[0-2])[/\-\.](0[1-9]|[12][0-9]|3[01])$"
        .Global = False
        .IgnoreCase = True
    End With
    
    If patronDDMMAAAA.Test(Texto) Then
        DetectarFormatoFecha = Formato_DDMMAAAA
    ElseIf patronMMDDAAAA.Test(Texto) Then
        DetectarFormatoFecha = Formato_MMDDAAAA
    ElseIf patronAAAAMMDD.Test(Texto) Then
        DetectarFormatoFecha = Formato_AAAAMMDD
    Else
        DetectarFormatoFecha = Formato_Texto
    End If
    
End Function

' Convertir fecha al formato estándar
Private Function ConvertirFecha( _
    ByVal valor As Variant, _
    ByVal formatoOrigen As enuFormatoFecha, _
    ByRef fechaSalida As Date) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim Texto As String
    Texto = CStr(valor)
    
    Dim partes() As String
    Dim dia As Integer, mes As Integer, año As Integer
    
    ' Separar por delimitadores comunes
    Texto = Replace(Texto, "-", "/")
    Texto = Replace(Texto, ".", "/")
    partes = Split(Texto, "/")
    
    Select Case formatoOrigen
        Case Formato_DDMMAAAA
            dia = CInt(partes(0))
            mes = CInt(partes(1))
            año = CInt(partes(2))
            
        Case Formato_MMDDAAAA
            mes = CInt(partes(0))
            dia = CInt(partes(1))
            año = CInt(partes(2))
            
        Case Formato_AAAAMMDD
            año = CInt(partes(0))
            mes = CInt(partes(1))
            dia = CInt(partes(2))
            
        Case Formato_Texto
            If IsDate(valor) Then
                fechaSalida = CDate(valor)
                ConvertirFecha = True
                Exit Function
            Else
                ConvertirFecha = False
                Exit Function
            End If
    End Select
    
    ' Validar fecha
    If EsFechaValida(dia, mes, año) Then
        fechaSalida = DateSerial(año, mes, dia)
        ConvertirFecha = True
    Else
        ConvertirFecha = False
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertirFecha = False
End Function

' Validar si una fecha es válida
Private Function EsFechaValida(ByVal dia As Integer, ByVal mes As Integer, ByVal año As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    If año < 100 Then
        If año >= 0 And año <= 29 Then
            año = año + 2000
        ElseIf año >= 30 And año <= 99 Then
            año = año + 1900
        End If
    End If
    
    If año < 1900 Or año > 2100 Then
        EsFechaValida = False
        Exit Function
    End If
    
    If mes < 1 Or mes > 12 Then
        EsFechaValida = False
        Exit Function
    End If
    
    If dia < 1 Or dia > 31 Then
        EsFechaValida = False
        Exit Function
    End If
    
    ' Verificar días del mes
    Dim ultimoDia As Integer
    ultimoDia = Day(DateSerial(año, mes + 1, 0))
    
    EsFechaValida = (dia <= ultimoDia)
    
    Exit Function
    
ErrorHandler:
    EsFechaValida = False
End Function

' Aplicar formato a la columna
Private Sub AplicarFormatoColumna( _
    ByRef rango As Range, _
    ByVal formato As enuFormatoFecha, _
    ByVal separador As String, _
    ByVal incluirHora As Boolean)
    
    Dim formatoTexto As String
    
    Select Case formato
        Case Formato_DDMMAAAA
            If incluirHora Then
                formatoTexto = "dd" & separador & "mm" & separador & "yyyy hh:mm:ss"
            Else
                formatoTexto = "dd" & separador & "mm" & separador & "yyyy"
            End If
            
        Case Formato_MMDDAAAA
            If incluirHora Then
                formatoTexto = "mm" & separador & "dd" & separador & "yyyy hh:mm:ss"
            Else
                formatoTexto = "mm" & separador & "dd" & separador & "yyyy"
            End If
            
        Case Formato_AAAAMMDD
            If incluirHora Then
                formatoTexto = "yyyy" & separador & "mm" & separador & "dd hh:mm:ss"
            Else
                formatoTexto = "yyyy" & separador & "mm" & separador & "dd"
            End If
            
        Case Formato_Texto
            If incluirHora Then
                formatoTexto = "dd" & separador & "mm" & separador & "yyyy hh:mm:ss"
            Else
                formatoTexto = "dd" & separador & "mm" & separador & "yyyy"
            End If
    End Select
    
    rango.NumberFormat = formatoTexto
End Sub

' Verificar si la columna existe en la tabla
Private Function ColumnaExiste(ByRef tbl As ListObject, ByVal nombreColumna As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim Col As ListColumn
    Set Col = tbl.ListColumns(nombreColumna)
    ColumnaExiste = True
    Exit Function
    
ErrorHandler:
    ColumnaExiste = False
End Function

' Mostrar reporte de estadísticas
Private Sub MostrarReporteEstadisticas( _
    ByRef estadisticas As Scripting.Dictionary, _
    ByVal nombreTabla As String, _
    ByVal nombreColumna As String)
    
    Dim mensaje As String
    Dim clave As Variant
    Dim totalCeldas As Long
    Dim formatoTexto As String
    
    mensaje = "REPORTE DE CONVERSIÓN DE FECHAS" & vbCrLf & vbCrLf
    mensaje = mensaje & "Tabla: " & nombreTabla & vbCrLf
    mensaje = mensaje & "Columna: " & nombreColumna & vbCrLf & vbCrLf
    
    For Each clave In estadisticas.Keys
        Select Case clave
            Case Formato_DDMMAAAA
                formatoTexto = "DD/MM/AAAA"
            Case Formato_MMDDAAAA
                formatoTexto = "MM/DD/AAAA"
            Case Formato_AAAAMMDD
                formatoTexto = "AAAA/MM/DD"
            Case Formato_Texto
                formatoTexto = "Texto/Fecha Excel"
        End Select
        
        mensaje = mensaje & formatoTexto & ": " & estadisticas(clave) & " celdas" & vbCrLf
        totalCeldas = totalCeldas + estadisticas(clave)
    Next clave
    
    mensaje = mensaje & vbCrLf & "Total procesado: " & totalCeldas & " celdas"
    
    MsgBox mensaje, vbInformation, "Reporte de Formateo"
End Sub
