Attribute VB_Name = "M_815_Mail_HTML"
Option Explicit

'==================================================================================================
' M_815_Mail_HTML
'   Funciones de generacion de HTML para cuerpos de correo
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' Fnc_HTML_Tabla
'   Genera el HTML de una tabla de 2 columnas y N filas
'
'   aTabla : Array de Arrays, cada fila con 4 elementos:
'               aTabla(fila)(0) = Color columna 1  (string, vacio = negro por defecto)
'               aTabla(fila)(1) = Texto  columna 1
'               aTabla(fila)(2) = Color columna 2  (string, vacio = negro por defecto)
'               aTabla(fila)(3) = Texto  columna 2
'
'   EJEMPLO DE LLAMADA:
'       Dim aTabla As Variant
'       aTabla = Array( _
'           Array("",       "Liquidacion realizada:", "Crimson", Range("Liq_Solicitud")), _
'           Array("",       "Actividad:",             "blue",    Range("Liq_Nombre")), _
'           Array("",       "Importe:",               "red",     Format(Range("Liq_DI_Imp"), "#,##0.00") & " EUR") _
'       )
'       Mail_Body = Fnc_HTML_Tabla(aTabla)
'--------------------------------------------------------------------------------------------------
Public Function Fnc_HTML_Tabla(ByVal aTabla As Variant) As String

    Dim i       As Integer
    Dim html    As String
    Dim c1Col   As String
    Dim c1Txt   As String
    Dim c2Col   As String
    Dim c2Txt   As String

    html = "<table style='border-collapse:collapse; font-size:14px;' cellpadding='2'>"

    For i = 0 To UBound(aTabla)

        c1Col = CStr(aTabla(i)(0))
        c1Txt = CStr(aTabla(i)(1))
        c2Col = CStr(aTabla(i)(2))
        c2Txt = CStr(aTabla(i)(3))

        ' --- Columna 1 ---
        html = html & "<tr>" & _
               "<td style='white-space:nowrap; padding-right:15px; vertical-align:top;'><b>"
        If c1Col <> "" Then html = html & "<font color='" & c1Col & "'>"
        html = html & c1Txt
        If c1Col <> "" Then html = html & "</font>"
        html = html & "</b></td>"

        ' --- Columna 2 ---
        html = html & "<td><b>"
        If c2Col <> "" Then html = html & "<font color='" & c2Col & "'>"
        html = html & c2Txt
        If c2Col <> "" Then html = html & "</font>"
        html = html & "</b></td></tr>"

    Next i

    html = html & "</table>"
    Fnc_HTML_Tabla = html

End Function
