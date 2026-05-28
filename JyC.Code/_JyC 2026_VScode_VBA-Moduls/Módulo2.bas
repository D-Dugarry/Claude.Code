Attribute VB_Name = "Módulo2"
Option Explicit

Public Sub Rut_Email_Send_Sheet(ByVal MailDestinatario As String, _
                         ByVal MailAsunto As String, _
                         Optional ByVal FicheroAdjunto As String, _
                         Optional ByVal MailCuerpo As String)
    
    '===================================================
    ' CONFIGURACIÓN SMTP (MODIFÍCALO)
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
    Dim wsLog As Worksheet
    Dim rngLog As Range
    Dim strBody As String
    Dim strLog As String
    
    On Error GoTo ErrorHandler
    
    ' --- Crear hoja de registro si no existe ---
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LogCorreos")
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "LogCorreos"
        ' Encabezados
        With wsLog
            .Range("A1:D1").Value = Array("Fecha", "Destinatario", "Asunto", "Cuerpo del Mensaje")
            .Rows(1).Font.Bold = True
            .Columns("A:D").AutoFit
        End With
    End If
    
    ' --- Registrar datos en hoja ---
    Set rngLog = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Offset(1, 0)
    
    ' Formatear cuerpo para registro (convertir HTML a texto si es necesario)
    If InStr(1, MailCuerpo, "<html>") > 0 Or InStr(1, MailCuerpo, "<p>") > 0 Then
        strBody = "HTML: " & Replace(Replace(MailCuerpo, vbCr, ""), vbLf, "")
    Else
        strBody = MailCuerpo
    End If
    
    ' Guardar registro
    With rngLog
        .Value = Now()
        .Offset(0, 1).Value = MailDestinatario
        .Offset(0, 2).Value = MailAsunto
        .Offset(0, 3).Value = strBody
        .Offset(0, 4).Value = IIf(FicheroAdjunto = "", "Sin adjunto", FicheroAdjunto)
    End With
    
    ' --- Configuración CDO para envío ---
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
    
    ' --- Configurar mensaje ---
    With objCDO
        Set .Configuration = iConf
        .From = """" & cdoFromName & """ <" & cdoFromEmail & ">"
        .To = MailDestinatario
        .Subject = MailAsunto
        
        If MailCuerpo = "" Then
            .TextBody = "Mensaje enviado desde Excel."
        ElseIf InStr(1, MailCuerpo, "<html>") > 0 Then
            .HTMLBody = MailCuerpo
        Else
            .TextBody = MailCuerpo
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
    
    ' --- Enviar (opcional) ---
    If MsgBox("¿Desea enviar el correo ahora?", vbQuestion + vbYesNo, "Confirmar envío") = vbYes Then
        On Error Resume Next
        objCDO.Send
        If Err.Number <> 0 Then
            MsgBox "Error al enviar: " & Err.Description, vbCritical
        Else
            rngLog.Offset(0, 5).Value = "ENVIADO"
            MsgBox "Correo enviado y registrado correctamente.", vbInformation
        End If
        On Error GoTo 0
    Else
        rngLog.Offset(0, 5).Value = "GUARDADO (NO ENVIADO)"
        MsgBox "Correo guardado en hoja 'LogCorreos' para enviar manualmente.", vbInformation
    End If
    
    ' --- Formatear hoja de registro ---
    With wsLog
        .Columns("A:F").AutoFit
        .ListObjects.Add(xlSrcRange, .UsedRange, , xlYes).Name = "tblLogCorreos"
        .Cells(1, 1).Activate
    End With
    
ExitSub:
    Set objCDO = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en Rut_Email_Send: " & Err.Description & vbCrLf & _
           "Linea: " & Erl, vbCritical
    Resume ExitSub
End Sub

