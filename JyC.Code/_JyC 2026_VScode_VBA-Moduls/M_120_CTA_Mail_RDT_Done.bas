Attribute VB_Name = "M_120_CTA_Mail_RDT_Done"
Option Explicit

'==================================================================================================
' M_120_CTA_Mail_RDT_Done
'
' Envia un e-mail de confirmacion al organizador indicando que la RDT ha sido realizada.
'
' DESTINATARIO : Range("Liq_Email") & ";facturacion@ua.es"  (organizador + facturacion UA)
' ADJUNTO      : ninguno
' TABLA HTML   : tabla completa de liquidacion CTA (Lo_Liq redimensionada a C_Cta_Liq_Obs cols)
'
' DIFERENCIA con M_110_CTA_Mail_Do_RDT:
'   M_110 = solicita que se haga la RDT (destinatario: ingresos)
'   M_120 = confirma que la RDT ya fue realizada (destinatario: organizador + facturacion)
'
' REQUIERE:
'   - M_800_Mail_Send_New : Rut_Email_Send
'   - M_815_Mail_HTML     : Fnc_HTML_Tabla
'   - M_820_Range_TO_HTML : Fnc_RangeToHTML
'
' NOTA HISTORICA:
'   Hasta 2024 el envio fallaba por el 2FA de Google. Se generaba el correo en la hoja
'   "CTAMailRDTDone" para copiarlo manualmente. Resuelto en 2026 con App Password (CDO).
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' Rut_Email_CTA_RDT_Done
'   Construye y envia el e-mail de confirmacion de RDT realizada para una liquidacion CTA
'--------------------------------------------------------------------------------------------------
Public Sub Rut_Email_CTA_RDT_Done()

    Dim DirMail     As String
    Dim Asunto      As String
    Dim Mail_Body   As String
    Dim Lin_JyC     As Long
    Dim Rng         As Range
    Dim aTabla      As Variant
    Dim Lo_Liq      As ListObject:  Set Lo_Liq = H_Liq_CTA.ListObjects(1)

    ' --- Validacion previa ---
    If Range("Liq_Solicitud") = "" Then
        MsgBx_Msg = "No hay Referencia de Solicitud!"
        MsgBx_Title = "Proceso: Enviar E-mail de Confirmación de Redistribución realizada."
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, , "OK", , "stop"): Form_MsgBox.Show
        Exit Sub
    End If

    Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    Call Rut_LstObj_Buscar(Lo_Lst, Range("Liq_Siglas"), 1, Lin_JyC)
    If Lin_JyC = 0 Then MsgBox "Siglas JyC, not find": Exit Sub
    Application.ScreenUpdating = False

    ' --- Rango de la tabla de liquidacion (todas las columnas hasta Obs) ---
    Set Rng = Lo_Liq.ListColumns(1).Range.Resize(, C_Cta_Liq_Obs)

    ' --- Destinatario y asunto ---
    DirMail = Range("Liq_Email") & ";facturacion@ua.es"
    Asunto = "JyC RDT_Done, Solicitud " & UCase(Range("Liq_Solicitud")) & _
              ", Liquidación nº " & Range("Liq_Núm") & _
              ", " & Range("Liq_Siglas")

    '- Genera Mail en una Sheet, para poder copiar y pegar cuando falla la programacion VBA.
    '- En 2024 con el 2FA fallo y se arreglo en 2026, mientras tanto se uso esta opcion.
    'With Sheets("CTAMailRDTDone")  ...  'End With

    ' --- Tabla de datos del cuerpo del correo ---
    aTabla = Array( _
        Array("", "Liquidación realizada:", _
              "Crimson", "JyC Solicitud " & UCase(Range("Liq_Solicitud")) & _
                         ", Liquidación nº " & Range("Liq_Núm") & _
                         ", JyC " & Range("Liq_Siglas")), _
        Array("", "Actividad:", _
              "blue", Range("Liq_Nombre")), _
        Array("", "Orgánica de destino:", _
              "blue", Range("Liq_Orgánica")), _
        Array("", "Liquidación número:", _
              "blue", Range("Liq_Núm")), _
        Array("", "Importe redistribuido:", _
              "red", Format(Lo_Liq.TotalsRowRange(C_Cta_Liq_Imp), "#,##0.00") & " &euro;") _
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
    MsgBx_Title = "Proceso: Enviar E-mail de Confirmación de Redistribución realizada."
    If CBool(Prog__APP.Range("SW_Test").Value) Then
        MsgBx_Msg = "Email enviado en modo prueba!"
    Else
        MsgBx_Msg = "Email enviado!"
    End If
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, , "OK", , "MailSent"): Form_MsgBox.Show

    Application.ScreenUpdating = True

End Sub     '- Rut_Email_CTA_RDT_Done ---
