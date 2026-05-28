Attribute VB_Name = "M_110_CTA_Mail_Do_RDT"
Option Explicit

'==================================================================================================
' M_110_CTA_Mail_Do_RDT
'
' Envia un e-mail solicitando que se realice la Redistribucion (RDT) de la liquidacion CTA.
'
' DESTINATARIO : ingresos@gcloud.ua.es (cuenta de gestion de ingresos de la UA)
' ADJUNTO      : ninguno
' TABLA HTML   : cabecera de la liquidacion (H_Liq_CTA, rango a1:h7)
'
' REQUIERE:
'   - M_800_Mail_Send_New : Rut_Email_Send
'   - M_815_Mail_HTML     : Fnc_HTML_Tabla
'   - M_820_Range_TO_HTML : Fnc_RangeToHTML
'
' NOTA HISTORICA:
'   Hasta 2024 el envio fallaba por el 2FA de Google. Se generaba el correo en la hoja
'   "CTAMailDoRDT" para copiarlo manualmente. Resuelto en 2026 con App Password (CDO).
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' Rut_Email_CTA_Do_RDT
'   Construye y envia el e-mail de solicitud de RDT para una liquidacion CTA
'--------------------------------------------------------------------------------------------------
Public Sub Rut_Email_CTA_Do_RDT()

    Dim DirMail     As String
    Dim Asunto      As String
    Dim Mail_Body   As String
    Dim Rng         As Range
    Dim aTabla      As Variant

    ' --- Validaciones previas ---
    If Not Fnc_Validar_Campos_RDT() Then Exit Sub

    ' --- Rango de cabecera de la liquidacion ---
    Set Rng = H_Liq_CTA.Range("a1:h7")

    ' --- Destinatario y asunto ---
    DirMail = "ingresos@gcloud.ua.es"
    Asunto = "JyC RDT, Sol. " & UCase(Range("Liq_Solicitud")) & _
              ", Liq. nº " & Range("Liq_Núm") & _
              ", JyC " & Range("Liq_Siglas")

    '- Este trozo de código, genera el contenido del Mail en una Sheet,
    '- para poder copiar y pegar, cuando falla la progración VBA de envío de correo,
    '- En 2024 con el 2fa falló y se arregló en 2026, mientras tanto se usó esta opción.
        'Const SheetMail     As String = "CTAMailDoRDT"
        'Sheets(SheetMail).Range("a10:h16").ClearContents
        'With Sheets(SheetMail)
        '    .Range("b1") = Asunto
        '    .Range("e4") = "JyC Solicitud " & UCase(Range("Liq_Solicitud")) & ", Liquidación nº " & Range("Liq_Núm") & ", JyC " & Range("Liq_Siglas")
        '    .Range("c5") = "Liquid_" & Range("Liq_Núm") & "_JyC_" & Range("Liq_Siglas")
        '    .Range("c6") = "\_Unidad_Ingresos\4_JORNADAS Y CONGRESOS\_Liquidaciones_JyC " & Range("APP_AñoCont") & "\" & Range("Liq_Siglas") _
        '                                    & "\Liquid_" & Range("Liq_Núm") & "_" & Range("APP_AñoCont") & "_" & Range("Liq_Solicitud")
        '    .Range("b8") = "Liq_JyC_" & Format(Date, "yyyy") & "__" & Range("Liq_Siglas") & "_" & Range("Liq_Núm") & "__" & Range("Liq_Nombre")
        '    .Range("a4:e8").WrapText = False
        '    ' Copio la Tabla de Datos
        '    Rng.Copy (.Range("a10"))
        'End With
    
    ' --- Tabla de datos del cuerpo del correo ---
    aTabla = Array( _
        Array("", "Tenemos que Contabilizar en UXXI y hacer la RDT de:", _
              "Crimson", "JyC Solicitud " & UCase(Range("Liq_Solicitud")) & _
                      ", Liquidación nº " & Range("Liq_Núm") & _
                      ", JyC " & Range("Liq_Siglas")), _
        Array("", "Fichero Datos:", _
              "blue", "Liquid_" & Range("Liq_Núm") & "_JyC_" & Range("Liq_Siglas") & ".xlsx"), _
        Array("", "Ruta NEXE:", _
              "blue", "\_Unidad_Ingresos\4_JORNADAS Y CONGRESOS\_Liquidaciones_JyC " & _
                      Prog__APP.Range("APP_AñoCont") & "\" & Range("Liq_Siglas") & _
                      "\Liquid_" & Range("Liq_Núm") & "_" & _
                      Prog__APP.Range("APP_AñoCont") & "_" & Range("Liq_Solicitud")), _
        Array("Crimson", "Nota importante:", _
              "Crimson", "El ingreso ha sido en la cuenta de JyC, " & _
                         "por lo que NO hay que hacer retención " & _
                         "en la ficha de datos económicos del JI."), _
        Array("", "Concepto JI:", _
              "blue", "Liq_JyC_" & Format(Date, "yyyy") & "__" & Range("Liq_Siglas") & "_" & _
                      Format(Range("Liq_Núm"), "00") & "__" & Range("Liq_Nombre")) _
    )

    ' --- Cuerpo del correo ---
    Mail_Body = "<div style='line-height:1.8; font-family:Arial, sans-serif; font-size:14px;'>" & _
                "<b>Hola</b><br><br>" & _
                Fnc_HTML_Tabla(aTabla) & "<br>" & _
                Fnc_RangeToHTML(Rng) & ".<br><br>" & _
                "</div>"

    ' --- Enviar ---
    Call Rut_Email_Send(DirMail, Asunto, , Mail_Body)

    ' --- Confirmacion ---
    MsgBx_Title = "Proceso: Enviar E-mail para realizar la Redistribución solicitada."
    If CBool(Prog__APP.Range("SW_Test").Value) Then
        MsgBx_Msg = "Email enviado en modo prueba!"
    Else
        MsgBx_Msg = "Email enviado!"
    End If
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, , "OK", , "MailSent"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")

End Sub     '- Rut_Email_CTA_Do_RDT ---


'--------------------------------------------------------------------------------------------------
' Fnc_Validar_Campos_RDT  (privada)
'   Verifica que los campos obligatorios de la liquidacion esten informados
'   Devuelve TRUE si todo es correcto, FALSE si falta algun campo
'--------------------------------------------------------------------------------------------------
Private Function Fnc_Validar_Campos_RDT() As Boolean

    Dim aValidaciones   As Variant
    Dim i               As Integer
    Dim TITULO          As String:    TITULO = "Proceso: Enviar E-mail para realizar la Redistribución solicitada."

    aValidaciones = Array( _
        Array("Liq_Siglas", "No hay Siglas de la Jornada o Congreso"), _
        Array("Liq_Núm", "No hay número de Liquidación"), _
        Array("Liq_Solicitud", "No hay Referencia de Solicitud") _
    )

    For i = 0 To UBound(aValidaciones)
        If Range(aValidaciones(i)(0)) = "" Then
            MsgBx_Msg = aValidaciones(i)(1)
            MsgBx_Title = TITULO
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, , "OK", , "stop"): Form_MsgBox.Show
            Fnc_Validar_Campos_RDT = False
            Exit Function
        End If
    Next i

    Fnc_Validar_Campos_RDT = True

End Function     '- Fnc_Validar_Campos_RDT ---
