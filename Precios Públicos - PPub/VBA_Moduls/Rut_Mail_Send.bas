Attribute VB_Name = "Rut_Mail_Send"
' Last Rev. 2026-10-13 12:28
Option Explicit

'===================================================================================================
' Rut_Mail_Send
'
' Motor de envío de correo: CDO con enlace tardío (CreateObject, sin marcar ninguna referencia en el
' VBE) por smtp.gmail.com:465 con SSL, autenticado con una contraseña de aplicación de Google.
' Portado de Jornadas_y_Congresos (M_801: configuración en la hoja ConfigMail y firma con 5 huecos)
' con la forma de función de Certificados_de_Pago (M_800: devuelve Boolean y la causa en ErrMsg).
'
' NO MUESTRA NINGÚN MsgBox. Si algo falla, devuelve False y la causa en ErrMsg; el llamador decide
' cómo avisar. Motivo: los envíos se lanzan con Form_Running_Rut abierto a pantalla completa, y un
' MsgBox normal se abre DETRÁS del formulario: Excel parece colgado (pasó el 2026-10-11). Para avisar
' con el formulario abierto: MsgBox ... + vbMsgBoxSetForeground + vbSystemModal (Func_MsgBox_vbYesNo).
'
' Configuración (se lee siempre cualificando con la hoja donde vive cada nombre):
'   Prog__APP_Mail (hoja ConfigMail):
'     App_MailCta    Cuenta de Gmail con la que se autentica (ingresos@gcloud.ua.es)
'     App_MailClau   Contraseña de aplicación de Google (16 caracteres, sin espacios)
'     App_MailFrom   Remitente visible (ingresos@ua.es, alias de esa cuenta). Vacío = App_MailCta
'     App_MailFirm   Firma HTML con 5 huecos: "ext. ____", "Servei _____", "Unitat _____",
'                    "https://sc.ua.es/es/_____" y "https://sc.ua.es/va/_____"
'     App_Servicio, App_Unidad, App_Web_es y App_Web_va: lo que se pone en esos huecos
'   Prog__APP (hoja ConfigAPP):
'     App_User_Ext   Extensión del usuario activo (hueco "ext. ____")
'     App_User_Name  Rúbrica: el nombre que va encima de la firma
'     App_User_Mail  Destinatario en modo prueba
'   Prog__APP_Switch (hoja SwitchsAPP):
'     Sw_Test        VERDADERO = modo prueba: el correo va a App_User_Mail y el asunto lleva "[PRUEBA] "
'
' Contraseña de aplicación (desde mayo de 2022 Google no acepta la contraseña normal por SMTP):
'   myaccount.google.com -> Seguridad -> Verificación en dos pasos (activada) -> Contraseñas de
'   aplicaciones -> crear una -> guardarla con Rut_Cambiar_Contrasena_Email (le quita los espacios).
'
' Subs y funciones públicas:
'   Fnc_Email_Send(Dest, Asunto, [Adjunto], [CuerpoHTML], [ErrMsg], [DestReal]) - Envía un correo
'   Fnc_Valid_Email(Email)                    - Una dirección bien escrita (sin avisos)
'   Fnc_Valid_Email_Lista(Lista, [ErrMsg])    - Limpia una lista de direcciones y dice cuáles están mal
'   Fnc_HTML_Texto(Texto)                     - Texto de la hoja preparado para el cuerpo HTML
'   Rut_Cambiar_Contrasena_Email              - Guarda una contraseña de aplicación nueva
'   Rut_Mail_Probar([ConAdjunto])             - Prueba el motor (Inmediato: Call Rut_Mail_Probar)
'   Fnc_Mail_Probar_Sin_Enviar                - La parte de la prueba que no envía nada
'===================================================================================================

Private Const SMTP_SERVIDOR     As String = "smtp.gmail.com"
Private Const SMTP_PUERTO       As Long = 465           ' CDO solo admite SSL (465), no STARTTLS (587)
Private Const SMTP_SSL          As Boolean = True       ' obligatorio con el puerto 465
Private Const SCHEMA            As String = "http://schemas.microsoft.com/cdo/configuration/"

'===================================================================================================
'- Envía un correo. Devuelve True si ha salido; si no, False y la causa en ErrMsg ------------------
'-   Destinatario : una o varias direcciones, separadas por ";" (o ","). Se validan SIEMPRE, también
'-                  en modo prueba, para que la prueba destape una lista mal escrita.
'-   Asunto       : en modo prueba se le antepone "[PRUEBA] ".
'-   Adjunto      : ruta completa de un fichero (vacío = sin adjunto). Si no existe, NO se envía.
'-   CuerpoHTML   : fragmento HTML, sin <html> ni <body>; el motor añade la rúbrica y la firma.
'-                  Los textos de la hoja, pasados antes por Fnc_HTML_Texto.
'-   ErrMsg       : (salida) la causa, si no se ha enviado.
'-   DestReal     : (salida) a quién va el correo de verdad (en modo prueba, App_User_Mail); vacío si
'-                  falla antes de saberlo.
Function Fnc_Email_Send(ByVal Destinatario As String, _
                        ByVal Asunto As String, _
                        Optional ByVal Adjunto As String = "", _
                        Optional ByVal CuerpoHTML As String = "", _
                        Optional ByRef ErrMsg As String = "", _
                        Optional ByRef DestReal As String = "") As Boolean
    Dim MailCta         As String
    Dim MailClau        As String
    Dim MailFrom        As String
    Dim MailFirma       As String
    Dim Rubrica         As String
    Dim SwTest          As Boolean
    Dim Correo          As Object
    Dim Paso            As String

    ErrMsg = ""
    DestReal = ""
    On Error GoTo Err_Email

    Paso = "leer la configuración"
    MailCta = Trim$(CStr(Prog__APP_Mail.Range("App_MailCta").Value))
    MailClau = Trim$(CStr(Prog__APP_Mail.Range("App_MailClau").Value))
    MailFrom = Trim$(CStr(Prog__APP_Mail.Range("App_MailFrom").Value))
    If MailFrom = "" Then MailFrom = MailCta
    Rubrica = Trim$(CStr(Prog__APP.Range("App_User_Name").Value))
    SwTest = CBool(Prog__APP_Switch.Range("Sw_Test").Value)
    MailFirma = Fnc_Mail_Firma()

    Paso = "comprobar los datos"
    If MailCta = "" Then
        ErrMsg = "No hay cuenta de correo en App_MailCta (hoja ConfigMail)."
        GoTo Salir
    End If
    If MailClau = "" Then
        ErrMsg = "No hay contraseña en App_MailClau (hoja ConfigMail)." & vbLf & _
                 "Hace falta una contraseña de aplicación de Google: ver Rut_Cambiar_Contrasena_Email."
        GoTo Salir
    End If
    If Not Fnc_Valid_Email_Lista(Destinatario, ErrMsg) Then GoTo Salir
    If SwTest Then
        DestReal = Trim$(CStr(Prog__APP.Range("App_User_Mail").Value))
        If Not Fnc_Valid_Email_Lista(DestReal, ErrMsg) Then
            ErrMsg = "Modo prueba (Sw_Test): App_User_Mail no vale como destinatario." & vbLf & ErrMsg
            DestReal = ""
            GoTo Salir
        End If
        Asunto = "[PRUEBA] " & Asunto
    Else
        DestReal = Destinatario
    End If
    If Adjunto <> "" Then
        If Not Fnc_Mail_Fichero_Existe(Adjunto) Then
            ErrMsg = "No se encuentra el adjunto; no se ha enviado nada:" & vbLf & Adjunto
            If LCase$(Left$(Adjunto, 4)) = "http" Then
                ErrMsg = ErrMsg & vbLf & "Las rutas URL (https://...) no valen: usar la letra de unidad (p. ej. Y:\...)."
            End If
            GoTo Salir
        End If
    End If

    Paso = "preparar el mensaje"
    Set Correo = CreateObject("CDO.Message")
    With Correo
        '- UTF-8 antes del asunto y del cuerpo. Sin esto CDO usa iso-8859-1, que no tiene el euro: en el
        '- asunto llegaba "?" (2026-10-13). Comprobado en el MIME que genera CDO (GetStream): asunto,
        '- texto y HTML salen en utf-8 y se leen bien; antes el HTML iba sin juego de caracteres.
        .BodyPart.Charset = "utf-8"
        .From = MailFrom                                ' remitente visible; se autentica con MailCta
        .To = DestReal
        .Subject = Asunto
        .HTMLBody = "<html><body>" & CuerpoHTML & _
                    "<br><font color=""black""><b>" & Fnc_HTML_Texto(Rubrica) & "</b></font><br><br>" & _
                    MailFirma & "</body></html>"
        .HTMLBodyPart.Charset = "utf-8"
        If Adjunto <> "" Then .AddAttachment Adjunto
        With .Configuration.Fields
            .Item(SCHEMA & "smtpserver") = SMTP_SERVIDOR
            .Item(SCHEMA & "smtpserverport") = SMTP_PUERTO
            .Item(SCHEMA & "sendusing") = 2                 ' cdoSendUsingPort
            .Item(SCHEMA & "smtpauthenticate") = 1          ' cdoBasic
            .Item(SCHEMA & "sendusername") = MailCta
            .Item(SCHEMA & "sendpassword") = MailClau
            .Item(SCHEMA & "smtpusessl") = SMTP_SSL
            .Item(SCHEMA & "smtpconnectiontimeout") = 60
            .Update
        End With
        Paso = "enviar"
        .Send
    End With
    Fnc_Email_Send = True

Salir:
    Set Correo = Nothing
    Exit Function

Err_Email:
    ErrMsg = Fnc_Mail_Texto_Error(Err.Number, Err.Description, Paso, MailCta, Adjunto)
    Resume Salir
End Function    ' Fnc_Email_Send
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Una dirección de correo bien escrita (la expresión regular de JyC M_810), sin avisos ------------
Function Fnc_Valid_Email(ByVal Email As String) As Boolean
    Dim RegExp      As Object

    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.IgnoreCase = True
    RegExp.Pattern = "^[a-z0-9_.+\-]+@[a-z0-9.\-]{2,}\.[a-z]{2,6}$"
    Fnc_Valid_Email = RegExp.Test(Trim$(Email))
    Set RegExp = Nothing
End Function    ' Fnc_Valid_Email
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Limpia una lista de direcciones y dice cuáles están mal escritas --------------------------------
'-   Lista  : direcciones separadas por ";" o ",". Si todas están bien, vuelve limpia: sin espacios,
'-            sin separadores sobrantes, sin repetidas y separadas por ";". Si no, no se toca.
'-   ErrMsg : (salida) las direcciones mal escritas, o que no hay ninguna.
'- A diferencia de JyC M_810, no pregunta si quitar las malas: las listas de PPub son de configuración
'- y se corrigen en su celda.
Function Fnc_Valid_Email_Lista(ByRef Lista As String, Optional ByRef ErrMsg As String = "") As Boolean
    Dim Parte       As Variant
    Dim Direccion   As String
    Dim Limpia      As String
    Dim Malas       As String

    ErrMsg = ""
    For Each Parte In Split(Replace(Lista, ",", ";"), ";")
        Direccion = Replace(Replace(Replace(Replace(CStr(Parte), Chr(160), " "), vbTab, " "), vbCr, " "), vbLf, " ")
        Direccion = Trim$(Direccion)
        If Direccion <> "" Then
            If Not Fnc_Valid_Email(Direccion) Then
                Malas = Malas & vbLf & "   " & Direccion
            ElseIf InStr(1, ";" & Limpia & ";", ";" & Direccion & ";", vbTextCompare) = 0 Then
                Limpia = Limpia & IIf(Limpia = "", "", ";") & Direccion
            End If
        End If
    Next Parte

    If Malas <> "" Then
        ErrMsg = "Direcciones de correo mal escritas:" & Malas
    ElseIf Limpia = "" Then
        ErrMsg = "No hay ninguna dirección de correo."
    Else
        Lista = Limpia
        Fnc_Valid_Email_Lista = True
    End If
End Function    ' Fnc_Valid_Email_Lista
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Prepara un texto de la hoja para el cuerpo HTML (de JyC M_815, más el escape de & < >) ----------
'-   & < >  : se escapan, para que el texto salga tal cual y no se tome por HTML.
'-   Euro   : el símbolo literal le llega al destinatario como "?"; la entidad &euro; es ASCII y
'-            llega bien (comprobado en JyC el 2026-09-16).
'-   Saltos : en HTML un vbLf no pinta nada; pasan a <br>.
Function Fnc_HTML_Texto(ByVal Texto As String) As String
    Dim S           As String

    S = Replace(Texto, "&", "&amp;")                    ' el primero: los demás añaden "&"
    S = Replace(S, "<", "&lt;")
    S = Replace(S, ">", "&gt;")
    S = Replace(S, ChrW(8364), "&euro;")
    S = Replace(S, vbCrLf, "<br>")
    S = Replace(S, vbLf, "<br>")
    S = Replace(S, vbCr, "<br>")
    Fnc_HTML_Texto = S
End Function    ' Fnc_HTML_Texto
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Guarda en App_MailClau una contraseña de aplicación nueva, sin los espacios que pone Google -----
'- Tarea sin botón (Tb_Tareas, para Boss). Pregunta con InputBox: no lanzarla con Form_Running_Rut
'- abierto.
Sub Rut_Cambiar_Contrasena_Email()
    Dim ClauActual  As String
    Dim NuevaClau   As String

    ClauActual = Trim$(CStr(Prog__APP_Mail.Range("App_MailClau").Value))
    NuevaClau = InputBox("La contraseña actual es:   " & ClauActual & vbLf & vbLf & _
                         "Introduce la nueva contraseña de aplicación de Google:", "Contraseña del correo")
    NuevaClau = Replace(Replace(NuevaClau, " ", ""), Chr(160), "")
    If NuevaClau = "" Then
        MsgBox "No se ha cambiado nada.", vbInformation + vbMsgBoxSetForeground, "Contraseña del correo"
        Exit Sub
    End If

    On Error GoTo Err_Guardar
    Prog__APP_Mail.Range("App_MailClau").Value = NuevaClau
    On Error GoTo 0
    MsgBox "Contraseña guardada (" & Len(NuevaClau) & " caracteres)." & _
           IIf(Len(NuevaClau) <> 16, vbLf & vbLf & "Ojo: las de Google tienen 16.", ""), _
           vbInformation + vbMsgBoxSetForeground, "Contraseña del correo"
    Exit Sub

Err_Guardar:
    MsgBox "No se ha podido guardar la contraseña en App_MailClau (hoja ConfigMail):" & vbLf & _
           Err.Description, vbCritical + vbMsgBoxSetForeground, "Contraseña del correo"
End Sub     ' Rut_Cambiar_Contrasena_Email
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Prueba del motor. Desde Inmediato: Call Rut_Mail_Probar  (con adjunto: Call Rut_Mail_Probar(True))
'-   1) Sin enviar nada: Fnc_Mail_Probar_Sin_Enviar. Cada fallo sale en Inmediato.
'-   2) Un correo al usuario activo (App_User_Mail) con tildes, eñe y euro en el asunto y en el cuerpo
'-      (con Sw_Test, el asunto lleva además "[PRUEBA] "). Hay que mirar en el correo que llegan bien.
Sub Rut_Mail_Probar(Optional ByVal ConAdjunto As Boolean = False)
    Dim Fallos      As String
    Dim ErrMsg      As String
    Dim DestReal    As String
    Dim Adjunto     As String
    Dim Cuerpo      As String
    Dim Msg         As String
    Dim Ok          As Boolean
    Dim NumFich     As Integer

    '- 1) Sin enviar nada --------------------------------------------------------------------------
    Fallos = Fnc_Mail_Probar_Sin_Enviar()
    If Fallos <> "" Then Debug.Print Fallos
    Msg = "1) Sin enviar nada: " & IIf(Fallos = "", "todo bien.", "hay fallos, ver Inmediato.")
    Debug.Print Msg

    '- 2) Un correo de prueba al usuario activo ----------------------------------------------------
    If ConAdjunto Then
        Adjunto = Environ$("TEMP") & "\PPub_Prueba_Correo.txt"
        NumFich = FreeFile
        Open Adjunto For Output As #NumFich
        Print #NumFich, "Adjunto de prueba del motor de correo de PPub (" & Format$(Now, "dd/mm/yyyy hh:nn:ss") & ")."
        Close #NumFich
    End If
    Cuerpo = "<p>Correo de prueba del motor de correo de PPub (Rut_Mail_Send).</p><p>" & _
             Fnc_HTML_Texto("Tildes y eñe: año, cálculo, Unitat d'Ingressos." & vbLf & _
                            "Euro: 1.234,56 €" & vbLf & "Signos: <, > y &") & "</p>"
    Ok = Fnc_Email_Send(Trim$(CStr(Prog__APP.Range("App_User_Mail").Value)), _
                        "Prueba del correo de PPub: año, cálculo, 1.234,56 €", Adjunto, Cuerpo, ErrMsg, DestReal)
    If Adjunto <> "" Then
        On Error Resume Next
        Kill Adjunto
        On Error GoTo 0
    End If

    If Ok Then
        Msg = Msg & vbLf & vbLf & "2) Correo enviado a " & DestReal & "." & vbLf & _
              "Mirar en el correo que el asunto y el cuerpo traen bien las tildes, la eñe y el euro" & _
              IIf(ConAdjunto, ", y que llega el adjunto.", ".")
    Else
        Msg = Msg & vbLf & vbLf & "2) No se ha enviado:" & vbLf & ErrMsg
    End If
    Debug.Print Mid$(Msg, InStr(Msg, "2)"))
    MsgBox Msg, IIf(Ok And Fallos = "", vbInformation, vbExclamation) + vbMsgBoxSetForeground, "Prueba del correo"
End Sub     ' Rut_Mail_Probar
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Prueba del motor sin enviar nada. Devuelve los fallos, uno por línea (vacío = todo bien) --------
'- La validación de direcciones, Fnc_HTML_Texto y que Fnc_Email_Send no envíe sin destinatarios ni
'- con un adjunto que no existe (los dos casos se paran antes de conectar con el servidor).
Function Fnc_Mail_Probar_Sin_Enviar() As String
    Dim Casos       As Variant
    Dim Caso        As Variant
    Dim Lista       As String
    Dim ErrMsg      As String
    Dim Ok          As Boolean
    Dim Fallos      As String

    '- Cada caso: lista de entrada, si es válida y cómo debe quedar
    Casos = Array(Array("a@ua.es", True, "a@ua.es"), _
                  Array(" a@ua.es ; ;b.c@gcloud.ua.es; ", True, "a@ua.es;b.c@gcloud.ua.es"), _
                  Array("a@ua.es, b@ua.es", True, "a@ua.es;b@ua.es"), _
                  Array("a@ua.es;A@UA.ES", True, "a@ua.es"), _
                  Array("", False, ""), _
                  Array(" ; ", False, ""), _
                  Array("a@ua.es;sin_arroba.ua.es", False, ""), _
                  Array("a@ua", False, ""))
    For Each Caso In Casos
        Lista = Caso(0)
        Ok = Fnc_Valid_Email_Lista(Lista, ErrMsg)
        If Ok <> Caso(1) Or (Ok And Lista <> Caso(2)) Then
            Fallos = Fallos & "FALLO al validar """ & Caso(0) & """: " & Ok & ", """ & Lista & """, " & ErrMsg & vbLf
        End If
    Next Caso
    If Fnc_HTML_Texto("a & b <c>" & vbLf & "1 €") <> "a &amp; b &lt;c&gt;<br>1 &euro;" Then
        Fallos = Fallos & "FALLO en Fnc_HTML_Texto: " & Fnc_HTML_Texto("a & b <c>" & vbLf & "1 €") & vbLf
    End If
    If Fnc_Email_Send("", "Prueba", , , ErrMsg) Or InStr(ErrMsg, "ninguna dirección") = 0 Then
        Fallos = Fallos & "FALLO sin destinatarios: " & ErrMsg & vbLf
    End If
    If Fnc_Email_Send("a@ua.es", "Prueba", Environ$("TEMP") & "\PPub_no_existe.xlsx", , ErrMsg) _
       Or InStr(ErrMsg, "No se encuentra el adjunto") = 0 Then
        Fallos = Fallos & "FALLO con un adjunto que no existe: " & ErrMsg & vbLf
    End If
    Fnc_Mail_Probar_Sin_Enviar = Fallos
End Function    ' Fnc_Mail_Probar_Sin_Enviar
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Firma de ConfigMail con sus 5 huecos rellenos (como JyC M_801) ----------------------------------
Private Function Fnc_Mail_Firma() As String
    Dim Firma       As String

    Firma = CStr(Prog__APP_Mail.Range("App_MailFirm").Value)
    Firma = Replace(Firma, "ext. ____", "ext. " & Trim$(CStr(Prog__APP.Range("App_User_Ext").Value)))
    Firma = Replace(Firma, "https://sc.ua.es/es/_____", Trim$(CStr(Prog__APP_Mail.Range("App_Web_es").Value)))
    Firma = Replace(Firma, "https://sc.ua.es/va/_____", Trim$(CStr(Prog__APP_Mail.Range("App_Web_va").Value)))
    Firma = Replace(Firma, "Servei _____", Trim$(CStr(Prog__APP_Mail.Range("App_Servicio").Value)))
    Firma = Replace(Firma, "Unitat _____", Trim$(CStr(Prog__APP_Mail.Range("App_Unidad").Value)))
    Fnc_Mail_Firma = Firma
End Function    ' Fnc_Mail_Firma
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Si existe un fichero (no una carpeta). Sin Dir$, que deja a medias otro Dir$ en curso -----------
Private Function Fnc_Mail_Fichero_Existe(ByVal Ruta As String) As Boolean
    On Error Resume Next
    Fnc_Mail_Fichero_Existe = ((GetAttr(Ruta) And vbDirectory) = 0)
    If Err.Number <> 0 Then Fnc_Mail_Fichero_Existe = False
    On Error GoTo 0
End Function    ' Fnc_Mail_Fichero_Existe
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Texto de un error de envío, con la ayuda de JyC/Certificados según el código --------------------
Private Function Fnc_Mail_Texto_Error(ByVal Num As Long, ByVal Descr As String, ByVal Paso As String, _
                                      ByVal MailCta As String, ByVal Adjunto As String) As String
    Dim Texto       As String
    Dim Ayuda       As String

    Ayuda = "Google ya NO acepta la contraseña normal: hace falta una contraseña de aplicación." & vbLf & _
            "  1. myaccount.google.com -> Seguridad" & vbLf & _
            "  2. Verificación en dos pasos -> Contraseñas de aplicaciones" & vbLf & _
            "  3. Crear una nueva y guardarla con Rut_Cambiar_Contrasena_Email (App_MailClau)"

    Descr = Trim$(Replace(Replace(Descr, vbCr, " "), vbLf, " "))   ' la de CDO acaba en salto de línea
    Texto = "Error " & Num & " al " & Paso & ": " & Descr & vbLf & vbLf
    Select Case Num
        Case -2147220973                                ' 0x80040213: no conecta con el servidor
            Texto = Texto & "Sin conexión a Internet o servidor no accesible." & vbLf & _
                            "Comprobar la conexión y que " & SMTP_SERVIDOR & ":" & SMTP_PUERTO & " no esté bloqueado."
        Case -2147220977                                ' 0x8004020F: el servidor rechaza el envío
            Texto = Texto & "El servidor ha rechazado el envío; con Gmail casi siempre son las credenciales " & _
                            "(cuenta " & MailCta & ")." & vbLf & Ayuda
        Case -2147220975                                ' 0x80040211: no se ha podido entregar al servidor
            If InStr(1, Descr, "80040217", vbTextCompare) > 0 Then      ' 0x80040217: no inicia sesión
                Texto = Texto & "Credenciales incorrectas (cuenta " & MailCta & ")." & vbLf & Ayuda
            Else
                Texto = Texto & "El servidor no ha aceptado el correo." & vbLf & _
                                "Cuenta: " & MailCta & vbLf & "Servidor: " & SMTP_SERVIDOR & ":" & SMTP_PUERTO
            End If
        Case 52, 53, 76                                 ' ruta del adjunto no válida o inaccesible
            Texto = Texto & "Ruta del adjunto no válida o inaccesible:" & vbLf & Adjunto & vbLf & _
                            "Las rutas URL (https://...) no valen: usar la letra de unidad (p. ej. Y:\...)."
        Case Else
            Texto = Texto & "Cuenta: " & MailCta & vbLf & "Servidor: " & SMTP_SERVIDOR & ":" & SMTP_PUERTO
    End Select
    Fnc_Mail_Texto_Error = Texto
End Function    ' Fnc_Mail_Texto_Error
'---------------------------------------------------------------------------------------------------
