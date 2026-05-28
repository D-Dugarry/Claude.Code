Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub Rut_Email_Send_Formats(ByVal MailDestinatario As String, _
                         ByVal MailAsunto As String, _
                         Optional ByVal FicheroAdjunto As String, _
                         Optional ByVal MailCuerpo As String)
    
    '===================================================
    ' CONFIGURACIÓN SMTP
    '===================================================
    Const cdoServer As String = "smtp.gmail.com"
    Const cdoPort As Integer = 587
    Const cdoUser As String = "ingresos@gcloud.ua.es"
    Const cdoPassword As String = "Trng.1936-Bmb"
    Const cdoSSL As Boolean = True
    Const cdoAuth As Boolean = True
    Const cdoFromEmail As String = "ingresos@gcloud.ua.es"
    Const cdoFromName As String = "Ingresos"
    '===================================================
    
    Dim objCDO As Object, iConf As Object, Flds As Variant
    Dim wsLog As Worksheet, wsHTML As Worksheet, wsEML As Worksheet
    Dim strBody As String, strHTML As String, strEML As String
    Dim FilePath As String, FileNumber As Integer
    
    On Error GoTo ErrorHandler
    
    ' --- Crear hojas de registro ---
    Application.ScreenUpdating = False
    CreateLogSheets
    
    ' --- Configurar cuerpo del mensaje ---
    If MailCuerpo = "" Then MailCuerpo = "Mensaje enviado desde Excel."
    strBody = MailCuerpo
    
    ' --- Registrar en hoja Log ---
    Set wsLog = ThisWorkbook.Sheets("LogCorreos")
    With wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Offset(1, 0)
        .Value = Now()
        .Offset(0, 1).Value = MailDestinatario
        .Offset(0, 2).Value = MailAsunto
        .Offset(0, 3).Value = IIf(InStr(1, strBody, "<html>") > 0, "HTML", "Texto")
        .Offset(0, 4).Value = IIf(FicheroAdjunto = "", "No", "Sí")
    End With
    
    ' --- Exportar a HTML (hoja HTML) ---
    Set wsHTML = ThisWorkbook.Sheets("HTML_Export")
    strHTML = BuildHTMLContent(MailDestinatario, MailAsunto, strBody, FicheroAdjunto)
    With wsHTML.Cells(wsHTML.Rows.Count, "A").End(xlUp).Offset(1, 0)
        .Value = strHTML
        .WrapText = False
        .Columns("A").ColumnWidth = 100
        .Rows.AutoFit
    End With
    
    ' --- Exportar a EML (hoja EML) ---
    Set wsEML = ThisWorkbook.Sheets("EML_Export")
    strEML = BuildEMLContent(MailDestinatario, MailAsunto, strBody, cdoFromEmail, cdoFromName)
    With wsEML.Cells(wsEML.Rows.Count, "A").End(xlUp).Offset(1, 0)
        .Value = strEML
        .WrapText = True
        .Columns("A").ColumnWidth = 100
        .Rows.AutoFit
    End With
    
    ' --- Opción para enviar ---
    If MsgBox("¿Desea enviar el correo ahora?" & vbCrLf & _
              "Se ha guardado copia en las hojas HTML_Export y EML_Export", _
              vbQuestion + vbYesNo, "Confirmar envío") = vbYes Then
        
        ' Configurar CDO
        Set objCDO = CreateObject("CDO.Message")
        Set iConf = CreateObject("CDO.Configuration")
        Set Flds = iConf.Fields
        
        With Flds
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = cdoServer
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = cdoPort
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = IIf(cdoAuth, 1, 0)
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = cdoSSL
            If cdoAuth Then
                .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = cdoUser
                .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = cdoPassword
            End If
            .Update
        End With
        
        ' Configurar mensaje
        With objCDO
            Set .Configuration = iConf
            .From = """" & cdoFromName & """ <" & cdoFromEmail & ">"
            .To = MailDestinatario
            .Subject = MailAsunto
            
            If InStr(1, strBody, "<html>") > 0 Then
                .HTMLBody = strBody
            Else
                .TextBody = strBody
            End If
            
            If FicheroAdjunto <> "" Then
                If Dir(FicheroAdjunto) <> "" Then
                    .AddAttachment FicheroAdjunto
                Else
                    MsgBox "Archivo no encontrado: " & FicheroAdjunto, vbExclamation
                    Exit Sub
                End If
            End If
        End With
        
        ' Enviar
        On Error Resume Next
        objCDO.Send
        If Err.Number <> 0 Then
            MsgBox "Error al enviar: " & Err.Description, vbCritical
            wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Offset(0, 5).Value = "FALLÓ"
        Else
            wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Offset(0, 5).Value = "ENVIADO"
        End If
        On Error GoTo 0
    Else
        wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Offset(0, 5).Value = "GUARDADO"
    End If
    
    ' --- Ajustar formato de hojas ---
    FormatLogSheets
    
    MsgBox "Correo procesado correctamente." & vbCrLf & _
           "- HTML listo para pegar en hoja [HTML_Export]" & vbCrLf & _
           "- EML listo para pegar en hoja [EML_Export]", vbInformation
    
ExitSub:
    Application.ScreenUpdating = True
    Set objCDO = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en Rut_Email_Send: " & Err.Description & vbCrLf & _
           "Linea: " & Erl, vbCritical
    Resume ExitSub
End Sub

'-----------------------------------------------------------
' FUNCIÓN PARA CONSTRUIR CONTENIDO HTML
'-----------------------------------------------------------
Private Function BuildHTMLContent(ByVal Destinatario As String, _
                                ByVal Asunto As String, _
                                ByVal Cuerpo As String, _
                                Optional ByVal Adjunto As String) As String
    
    Dim strHTML As String
    Dim strAdjunto As String
    
    strAdjunto = IIf(Adjunto = "", "No hay archivos adjuntos", "Adjunto: " & Adjunto)
    
    ' Plantilla HTML completa
    strHTML = "<!DOCTYPE html>" & vbCrLf & _
              "<html>" & vbCrLf & _
              "<head>" & vbCrLf & _
              "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf & _
              "<title>" & Asunto & "</title>" & vbCrLf & _
              "</head>" & vbCrLf & _
              "<body>" & vbCrLf & _
              "<div style='font-family: Arial, sans-serif; line-height: 1.6;'>" & vbCrLf & _
              "  <p><strong>Para:</strong> " & Destinatario & "</p>" & vbCrLf & _
              "  <p><strong>Asunto:</strong> " & Asunto & "</p>" & vbCrLf & _
              "  <p><strong>Adjunto:</strong> " & strAdjunto & "</p>" & vbCrLf & _
              "  <hr>" & vbCrLf & _
              "  <div style='margin-top: 20px;'>"
    
    ' Insertar cuerpo (ya sea HTML o texto)
    If InStr(1, Cuerpo, "<html>") > 0 Then
        ' Extraer solo el body si es HTML completo
        strHTML = strHTML & ExtractBodyContent(Cuerpo)
    Else
        ' Convertir texto plano a HTML
        strHTML = strHTML & "<pre style='white-space: pre-wrap;'>" & Cuerpo & "</pre>"
    End If
    
    strHTML = strHTML & vbCrLf & _
              "  </div>" & vbCrLf & _
              "</div>" & vbCrLf & _
              "</body>" & vbCrLf & _
              "</html>"
    
    BuildHTMLContent = strHTML
End Function

'-----------------------------------------------------------
' FUNCIÓN PARA CONSTRUIR CONTENIDO EML
'-----------------------------------------------------------
Private Function BuildEMLContent(ByVal Destinatario As String, _
                               ByVal Asunto As String, _
                               ByVal Cuerpo As String, _
                               ByVal RemitenteEmail As String, _
                               ByVal RemitenteNombre As String) As String
    
    Dim strEML As String
    Dim strBoundary As String
    Dim strMessageID As String
    
    ' Generar valores únicos
    strBoundary = "----=_NextPart_" & Format(Now, "yyyymmddhhnnss") & "_" & Rnd() * 1000
    strMessageID = "<" & Format(Now, "yyyymmddhhnnss") & "." & Rnd() * 1000 & "@excel>"
    
    ' Cabeceras EML
    strEML = "From: " & """" & RemitenteNombre & """" & " <" & RemitenteEmail & ">" & vbCrLf & _
             "To: " & Destinatario & vbCrLf & _
             "Subject: " & Asunto & vbCrLf & _
             "Date: " & Format(Now, "ddd, dd mmm yyyy hh:mm:ss") & " " & _
             Replace(Format(Now, "zzz"), ":", "") & vbCrLf & _
             "Message-ID: " & strMessageID & vbCrLf & _
             "MIME-Version: 1.0" & vbCrLf & _
             "Content-Type: multipart/alternative; " & vbCrLf & _
             "    boundary=" & """" & strBoundary & """" & vbCrLf & _
             vbCrLf & _
             "This is a multi-part message in MIME format." & vbCrLf & _
             vbCrLf & _
             "--" & strBoundary & vbCrLf
    
    ' Contenido del mensaje
    If InStr(1, Cuerpo, "<html>") > 0 Then
        ' Parte HTML
        strEML = strEML & "Content-Type: text/html; charset=utf-8" & vbCrLf & _
                 "Content-Transfer-Encoding: 7bit" & vbCrLf & _
                 vbCrLf & _
                 Cuerpo & vbCrLf & _
                 vbCrLf & _
                 "--" & strBoundary & vbCrLf
        
        ' Parte de texto plano (alternativa)
        strEML = strEML & "Content-Type: text/plain; charset=utf-8" & vbCrLf & _
                 "Content-Transfer-Encoding: 7bit" & vbCrLf & _
                 vbCrLf & _
                 "Este mensaje contiene formato HTML. Su cliente de correo no lo muestra correctamente." & vbCrLf
    Else
        ' Solo texto plano
        strEML = strEML & "Content-Type: text/plain; charset=utf-8" & vbCrLf & _
                 "Content-Transfer-Encoding: 7bit" & vbCrLf & _
                 vbCrLf & _
                 Cuerpo & vbCrLf
    End If
    
    strEML = strEML & vbCrLf & _
             "--" & strBoundary & "--" & vbCrLf
    
    BuildEMLContent = strEML
End Function

'-----------------------------------------------------------
' FUNCIÓN PARA EXTRAER SOLO EL CONTENIDO DEL BODY DE HTML
'-----------------------------------------------------------
Private Function ExtractBodyContent(ByVal HTMLContent As String) As String
    Dim StartPos As Integer, EndPos As Integer
    
    StartPos = InStr(1, HTMLContent, "<body")
    If StartPos > 0 Then
        StartPos = InStr(StartPos, HTMLContent, ">") + 1
        EndPos = InStr(StartPos, HTMLContent, "</body>")
        
        If EndPos > StartPos Then
            ExtractBodyContent = Mid(HTMLContent, StartPos, EndPos - StartPos)
            Exit Function
        End If
    End If
    
    ' Si no encuentra body, devolver todo
    ExtractBodyContent = HTMLContent
End Function

'-----------------------------------------------------------
' PROCEDIMIENTO PARA CREAR HOJAS DE REGISTRO
'-----------------------------------------------------------
Private Sub CreateLogSheets()
    On Error Resume Next
    
    ' Hoja Log principal
    ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "LogCorreos"
    With ThisWorkbook.Sheets("LogCorreos")
        .Range("A1:F1").Value = Array("Fecha", "Destinatario", "Asunto", "Formato", "Adjunto", "Estado")
        .Rows(1).Font.Bold = True
        .Columns("A:F").AutoFit
    End With
    
    ' Hoja HTML Export
    ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "HTML_Export"
    With ThisWorkbook.Sheets("HTML_Export")
        .Range("A1").Value = "HTML listo para copiar/pegar"
        .Rows(1).Font.Bold = True
        .Columns("A").ColumnWidth = 100
    End With
    
    ' Hoja EML Export
    ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "EML_Export"
    With ThisWorkbook.Sheets("EML_Export")
        .Range("A1").Value = "EML listo para copiar/pegar"
        .Rows(1).Font.Bold = True
        .Columns("A").ColumnWidth = 100
    End With
    
    On Error GoTo 0
End Sub

'-----------------------------------------------------------
' PROCEDIMIENTO PARA FORMATEAR HOJAS DE REGISTRO
'-----------------------------------------------------------
Private Sub FormatLogSheets()
    On Error Resume Next
    
    ' Formatear hoja Log
    With ThisWorkbook.Sheets("LogCorreos")
        .UsedRange.Borders.LineStyle = xlContinuous
        .ListObjects.Add(xlSrcRange, .UsedRange, , xlYes).Name = "tblLogCorreos"
        .Columns("A:F").AutoFit
    End With
    
    ' Formatear hoja HTML
    With ThisWorkbook.Sheets("HTML_Export")
        .UsedRange.WrapText = False
        .UsedRange.Rows.AutoFit
    End With
    
    ' Formatear hoja EML
    With ThisWorkbook.Sheets("EML_Export")
        .UsedRange.WrapText = True
        .UsedRange.Rows.AutoFit
    End With
    
    On Error GoTo 0
End Sub

