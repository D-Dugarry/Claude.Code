Attribute VB_Name = "M80_Mandar_Correo"
' Last Rev. 2026-09-21 12:12
'###################################################################################################
'   Nota Importante: Tenemos que Activar en este libro de Excel en Herramientas->Referencias->Microsoft CDO for Windows 2000 library
'###################################################################################################
Option Explicit

'===================================================================================================
Sub Rut_Email_Confirmar_Redistribucion_Hecha() ' ===================================================
Dim Respuesta   As Integer
Dim DirMail     As String

    DirMail = InputBox("Introducir el Correo Electrónico del Usuario", "Mandar E-Mail")
    
    If Len(DirMail) = 0 Then
        Form_Menu.TB_Informe = "Operación Cancelada"
        Exit Sub
    End If
    
    Respuesta = MsgBox("Mandar Email de Notificación de Redistribución a: " & vbCrLf & vbCrLf & "Email: " & DirMail, vbExclamation + vbOKCancel, "Rutinas de Correspondencia")
    If Respuesta = vbOK Then
        Dim Mail_Body        As String
        Dim Asunto          As String
        Asunto = Range("Liquid_Plan") & ", Redistribución realizada"
        Mail_Body = "<b>Buenos días <br><br><font size=3 color=""blue"">Os informamos que acabamos de redistribuir los fondos del Plan " _
                    & Range("Liquid_Plan") & ". <br>" _
                    & Left(Range("Liquid_Plan_Name"), 10) & Mid(Range("Liquid_Plan_Name"), 40) & ". </b></font><br><br>"
        Call Enviar_Email(DirMail, Asunto, , Mail_Body)
        Form_Menu.TB_Informe = "Email de notificación de solicitud enviada al Banco A: " & DirMail
    Else
        Form_Menu.TB_Informe = "Operación Cancelada"
    End If
        
End Sub     ' Rut_Email_Confirmar_Solicitud_Enviada_Bco --------------------------

'===================================================================================================
Sub Rut_Email_SinRDT_Motivo_Descripción() ' ========================================================
Dim Pos_Ini     As Long
Dim Pos_Fin     As Long
Dim Motivo      As String

    Motivo = Form_Menu.TBx_Descripción
    Pos_Ini = InStr(Motivo, "[[")
    Pos_Fin = InStr(Motivo, "]]")
    If Pos_Ini * Pos_Fin = 0 Then   '---Controla que existe marca de inicio y fin y que hay algun dato entre marcas
        Form_Menu.TB_Informe = "Error: No hay lista de Columnas a visualizar, o no empieza por [[, o no acaba por ]]." & vbCrLf & Now
        Exit Sub
    End If

    Motivo = Mid(Motivo, Pos_Ini + 2, Pos_Fin - Pos_Ini - 2)    '---extraigo la Motivo

    Form_Menu.TB_Informe = "Motivo de la Anulació: [[" & Motivo & "]]" & vbCrLf & Now

Dim Respuesta   As Integer
Dim DirMail     As String

    DirMail = InputBox("Introducir el Correo Electrónico del Usuario", "Mandar E-Mail")
    
    If Len(DirMail) = 0 Then
        Form_Menu.TB_Informe = "Operación Cancelada"
        Exit Sub
    End If
    
    Respuesta = MsgBox("Mandar Email de Notificación de Redistribución a: " & vbCrLf & vbCrLf & "Email: " & DirMail, vbExclamation + vbOKCancel, "Rutinas de Correspondencia")
    If Respuesta = vbOK Then
        Dim Mail_Body        As String
        Dim Asunto          As String
        Asunto = Range("Liquid_Plan") & ", Redistribución NO realizada"
        Mail_Body = "<b>Buenos días <br><br><font size=3 color=""blue"">Os informamos que NO podemos redistribuir los fondos del Plan " _
                    & Range("Liquid_Plan") & ". <br>" _
                    & Left(Range("Liquid_Plan_Name"), 10) & Mid(Range("Liquid_Plan_Name"), 40) & ". </font><br><br>" _
                    & "<font size=3 color=""red"">" & Motivo & " </b></font><br><br>"
        Call Enviar_Email(DirMail, Asunto, , Mail_Body)
        Form_Menu.TB_Informe = "Email de notificación de solicitud enviada al Banco A: " & DirMail
    Else
        Form_Menu.TB_Informe = "Operación Cancelada"
    End If
    
End Sub     ' Rut_Email_Confirmar_Solicitud_Enviada_Bco --------------------------


'===================================================================================================
'===================================================================================================
Sub Enviar_Email(ByVal MailDestinatario As String, ByVal MailAsunto As String, Optional ByVal FicheroAdjunto As String, Optional ByVal MailCuerpo As String)
'===================================================================================================
Dim Correo                  As CDO.Message
Dim Configuracion_Correo    As CDO.Configuration
Dim Campos                  As Variant
Dim FirmaCorreo             As String
    FirmaCorreo = Replace(Range("APP_MailFirm"), "ext. ____", "ext. " & Range("Usuario_Ext"))    '- Pongo la Extención correspondiente

    Set Correo = New CDO.Message
    Set Configuracion_Correo = New CDO.Configuration
    
    Configuracion_Correo.Load -1
    
    Set Campos = Configuracion_Correo.Fields
    
    With Correo
        .From = Range("APP_MailCta")               ' Emisor
        .Subject = MailAsunto               ' Asunto del correo
            If Prog__APP_Switch.Range("Sw_Probando") Then                 ' Destinatario  ----------
                .To = "dugarry@gcloud.ua.es"        ' Destinatario  En Prueba ----------------------
            Else                                    ' Destinatario  --------------------------------
                .To = MailDestinatario              ' Destinatario  --------------------------------
            End If                                  ' Destinatario  --------------------------------
        .Cc = ""
        .BCC = ""
        .HTMLBody = MailCuerpo & vbCrLf & "<b>Saludos, " & Range("Usuario_Name") & "</b><br><br>" & FirmaCorreo

                            
        If FicheroAdjunto <> "" Then .AddAttachment FicheroAdjunto
    End With    ' Correo
    
    Dim msConfigURL             As String
        msConfigURL = "http://schemas.microsoft.com/cdo/configuration"
    
    With Campos
        .Item(msConfigURL & "/smtpusessl") = True
        .Item(msConfigURL & "/smtpauthenticate") = 1
        
        .Item(msConfigURL & "/smtpserver") = "smtp.gmail.com"
        .Item(msConfigURL & "/smtpserverport") = 465
        .Item(msConfigURL & "/sendusing") = 2
        
        .Item(msConfigURL & "/sendusername") = Range("APP_MailCta")
        .Item(msConfigURL & "/sendpassword") = Range("APP_MailClau")
        .Update
    
    End With    ' Campos
    
    Correo.Configuration = Configuracion_Correo
    Correo.Send
    
'// Release object memory
Set Correo = Nothing
Set Configuracion_Correo = Nothing

End Sub     ' Enviar_Email -------------------------------------------------------------------------
'===================================================================================================

'===================================================================================================
Sub Rut_Cambiar_Contraseña_Email()
Dim NuevaClau   As String
    NuevaClau = InputBox("La Contraseña Actual es:   " & Range("APP_MailClau") & vbCrLf & vbCrLf & "Introduce la nueva Contraseña: ", "Actualizar Contraseña E-mail")
    If Len(NuevaClau) > 0 Then
        Range("APP_MailClau") = NuevaClau
        Form_Menu.TB_Informe = "Operación Realizada:  " & Now() & vbCrLf & "La nueva Contraseña es:  " & NuevaClau
    Else
        Form_Menu.TB_Informe = "La Contraseña Actual es:   " & Range("APP_MailClau") & "   - No se ha cambiado." & vbCrLf & vbCrLf & "Operación fue Anulada."
    End If
End Sub     ' Rut_Cambiar_Contraseña_Email    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<




