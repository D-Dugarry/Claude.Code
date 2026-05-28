Attribute VB_Name = "M_800_Mail_Send_New"
Option Explicit

'==================================================================================================
' M_800_Mail_Send_New
'
' REQUISITO: Necesita referencia a "Microsoft CDO for Windows 2000 Library"
'            Herramientas -> Referencias -> Microsoft CDO for Windows 2000 Library
'
' CAUSA DEL FALLO DEL MODULO ANTERIOR:
'   Google desactivó la autenticación básica (contraseña normal) para SMTP en mayo 2022.
'   SOLUCION: Generar una "Contraseña de Aplicación" (App Password) de 16 caracteres
'   en la cuenta de Google y guardarla en el rango con nombre APP_MailClau.
'
' PASOS PARA CREAR LA CONTRASEÑA DE APLICACION:
'   1. Ir a myaccount.google.com
'   2. Seguridad -> Verificación en dos pasos (debe estar activada)
'   3. Al final de esa página: "Contraseñas de aplicaciones"
'   4. Crear una nueva -> Otro (nombre personalizado) -> "Excel JyC"
'   5. Copiar la contraseña de 16 caracteres generada
'   6. Guardarla en la celda/rango APP_MailClau del libro (sin espacios)
'
' RANGOS EN Prog__APP (Config_APP):
'   APP_MailCta    -> Cuenta Gmail de autenticación SMTP (p.ej. ingresos@gcloud.ua.es)
'   APP_MailFrom   -> Dirección remitente visible (p.ej. ingresos@ua.es). Si vacío usa APP_MailCta.
'                     Debe estar verificada como alias en la cuenta de Gmail.
'   APP_MailClau   -> Contraseña de Aplicación de Google (16 chars, sin espacios)
'   APP_MailFirm   -> Firma de correo (HTML, con placeholder "ext. ____")
'   APP_User_Ext   -> Extensión telefónica para la firma
'   SW_Test        -> TRUE = modo prueba (redirige a DIR_PRUEBA)
'
' INTERFAZ PRINCIPAL:
'   Call Rut_Email_Send(Destinatario, Asunto, [Adjunto], [CuerpoHTML])
'==================================================================================================

Private Const SMTP_SERVIDOR  As String = "smtp.gmail.com"
Private Const SMTP_PUERTO    As Long = 465      ' CDO solo soporta SSL (465), no soporta STARTTLS (587)
Private Const SMTP_SSL       As Boolean = True   ' True obligatorio para puerto 465
Private Const DIR_PRUEBA     As String = "dugarry@gcloud.ua.es"
Private Const SCHEMA         As String = "http://schemas.microsoft.com/cdo/configuration/"


'--------------------------------------------------------------------------------------------------
' Rut_Email_Send
'   Destinatario  : dirección(es) de destino, separadas por ";"
'   Asunto        : asunto del correo
'   Adjunto       : ruta completa al archivo adjunto (opcional, "" si no hay)
'   CuerpoHTML    : cuerpo del mensaje en HTML (opcional)
'--------------------------------------------------------------------------------------------------
Public Sub Rut_Email_Send(ByVal Destinatario As String, _
                          ByVal Asunto As String, _
                          Optional ByVal Adjunto As String = "", _
                          Optional ByVal CuerpoHTML As String = "")

    Dim MailCta     As String
    Dim MailFrom    As String
    Dim MailClau    As String
    Dim MailFirma   As String
    Dim MailDest    As String
    Dim MailExt     As String
    Dim AppWeb_es   As String
    Dim APPWeb_va   As String
    Dim AppUnidad   As String
    Dim AppServicio As String
    Dim MailRubrica As String
    Dim SWTest      As Boolean
    Dim Correo      As Object
    Dim Configuracion As Object

    On Error GoTo Err_Email

    ' --- Leer configuración desde hoja Prog__APP ---
    MailCta = Trim(CStr(Prog__APP.Range("APP_MailCta").Value))
    MailFrom = Trim(CStr(Prog__APP.Range("APP_MailFrom").Value))
    If MailFrom = "" Then MailFrom = MailCta    ' Fallback: usar cuenta de auth si no hay alias
    MailClau = Trim(CStr(Prog__APP.Range("APP_MailClau").Value))
    MailExt = Trim(CStr(Prog__APP.Range("APP_User_Ext").Value))
    AppWeb_es = Trim(CStr(Prog__APP.Range("APP_Web_es").Value))
    APPWeb_va = Trim(CStr(Prog__APP.Range("APP_Web_va").Value))
    AppUnidad = Trim(CStr(Prog__APP.Range("APP_Unidad").Value))
    AppServicio = Trim(CStr(Prog__APP.Range("APP_Servicio").Value))
    MailRubrica = vbLf & Trim(CStr(Prog__APP.Range("APP_User_Rubrica").Value)) & vbLf
    SWTest = CBool(Prog__APP.Range("SW_Test").Value)

    ' Firma
    MailFirma = Prog__APP.Range("APP_MailFirm").Value
    MailFirma = Replace(MailFirma, "ext. ____", "ext. " & MailExt)
    MailFirma = Replace(MailFirma, "https://sc.ua.es/es/_____", AppWeb_es)
    MailFirma = Replace(MailFirma, "https://sc.ua.es/va/_____", APPWeb_va)
    MailFirma = Replace(MailFirma, "Servei _____", AppServicio)
    MailFirma = Replace(MailFirma, "Unitat _____", AppUnidad)

    ' Validaciones básicas
    If MailCta = "" Then
        MsgBox "No hay cuenta de correo configurada en APP_MailCta.", vbCritical, "Error Email"
        Exit Sub
    End If
    If MailClau = "" Then
        MsgBox "No hay contraseña en APP_MailClau." & vbNewLine & vbNewLine & _
               "Necesitas una Contraseña de Aplicación de Google." & vbNewLine & _
               "Ver comentarios en el inicio del módulo para instrucciones.", _
               vbCritical, "Error Email"
        Exit Sub
    End If
    If Not Fnc_Valid_Email_Multi(Destinatario) Then Exit Sub

    ' --- Modo prueba: redirigir destinatario ---
    If SWTest Then
        MailDest = DIR_PRUEBA
        Asunto = "[PRUEBA] " & Asunto
    Else
        MailDest = Destinatario
    End If

    ' --- Crear objeto CDO.Message ---
    Set Correo = CreateObject("CDO.Message")

    With Correo
        .From = MailFrom    ' Alias visible (ingresos@ua.es); auth sigue siendo MailCta
        .To = MailDest
        .Subject = Asunto

        ' Cuerpo HTML con firma
        If CuerpoHTML <> "" Then
            .HTMLBody = CuerpoHTML & "<Font color=""black""><b>" & MailRubrica & "<b/><br><br>" & MailFirma
        Else
            .HTMLBody = "<html><body><Font color=""black""><b>" & MailRubrica & "</b><br><br>" & MailFirma & "</body></html>"
        End If

        ' Adjunto
        If Adjunto <> "" Then
            If Dir(Adjunto) <> "" Then
                .AddAttachment Adjunto
            Else
                MsgBox "No se encontró el archivo adjunto:" & vbNewLine & Adjunto, vbExclamation, "Aviso"
            End If
        End If

        ' --- Configuración SMTP ---
        Set Configuracion = .Configuration
        With Configuracion.Fields
            .Item(SCHEMA & "smtpserver") = SMTP_SERVIDOR
            .Item(SCHEMA & "smtpserverport") = SMTP_PUERTO
            .Item(SCHEMA & "sendusing") = 2                    ' cdoSendUsingPort
            .Item(SCHEMA & "smtpauthenticate") = 1             ' cdoBasic
            .Item(SCHEMA & "sendusername") = MailCta
            .Item(SCHEMA & "sendpassword") = MailClau
            .Item(SCHEMA & "smtpusessl") = SMTP_SSL
            .Item(SCHEMA & "smtpconnectiontimeout") = 60
            .Update
        End With

        .Send

    End With

    Set Correo = Nothing
    Set Configuracion = Nothing

    Exit Sub

Err_Email:
    Dim MsgErr As String
    MsgErr = "Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine

    Select Case Err.Number
        Case -2147220973  ' 0x80040213 - Sin conexión
            MsgErr = MsgErr & "Sin conexión a Internet o servidor no accesible." & vbNewLine & _
                              "Verifica la conexión y que smtp.gmail.com:465 no esté bloqueado."
        Case -2147220977  ' 0x8004020F - Credenciales incorrectas
            MsgErr = MsgErr & "Credenciales incorrectas." & vbNewLine & vbNewLine & _
                              "IMPORTANTE: Google ya NO acepta la contraseña normal." & vbNewLine & _
                              "Necesitas una 'Contraseña de Aplicación' (App Password):" & vbNewLine & _
                              "  1. myaccount.google.com -> Seguridad" & vbNewLine & _
                              "  2. Verificación en dos pasos -> Contraseñas de aplicaciones" & vbNewLine & _
                              "  3. Crear una nueva y guardarla en APP_MailClau"
        Case Else
            MsgErr = MsgErr & "Cuenta: " & MailCta & vbNewLine & _
                              "Servidor: " & SMTP_SERVIDOR & ":" & SMTP_PUERTO
    End Select

    MsgBox MsgErr, vbCritical, "Error al enviar correo"
    Set Correo = Nothing
    Set Configuracion = Nothing
End Sub


'--------------------------------------------------------------------------------------------------
' Rut_Cambiar_Contrasena_Email  -  Actualiza APP_MailClau con la nueva App Password
'--------------------------------------------------------------------------------------------------
Public Sub Rut_Cambiar_Contrasena_Email()
    Dim NuevaClau  As String
    Dim ClauActual As String

    On Error Resume Next
    ClauActual = Trim(CStr(Prog__APP.Range("APP_MailClau").Value))
    On Error GoTo 0

    NuevaClau = InputBox("La Contraseña Actual es:   " & Prog__APP.Range("APP_MailClau") & vbCrLf & vbCrLf & "Introduce la nueva Contraseña: ", "Actualizar Contraseña E-mail")

    If NuevaClau = "" Then
        MsgBox "Operación anulada.", vbInformation
        Exit Sub
    End If

    NuevaClau = Replace(NuevaClau, " ", "")  ' Quitar espacios que Google añade al mostrarla

    Prog__APP.Range("APP_MailClau").Value = NuevaClau
    MsgBox "Contraseña actualizada correctamente." & vbNewLine & _
           "Longitud: " & Len(NuevaClau) & " caracteres.", vbInformation, "Contraseña Actualizada"
End Sub
