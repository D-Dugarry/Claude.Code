Attribute VB_Name = "M_225_TPV_Mail_DEV_Done"
Option Explicit

'==================================================================================================
' M_225_TPV_Mail_DEV_Done
'
' Envia un e-mail de confirmacion al organizador indicando que la DEV ha sido realizada.
'
' DESTINATARIO : Range("Liq_Email") & ";facturacion@ua.es"  (organizador + facturacion UA)
' ADJUNTO      : ninguno
' TABLA HTML   : tabla completa de liquidacion TPV (Lo_Liq redimensionada a C_TPV_Liq_Obs cols)
'
' DIFERENCIA con M_220_TPV_Mail_RDT_Done:
'   M_220 = confirmacion de RDT (Redistribucion de ingreso)
'   M_225 = confirmacion de DEV (Devolucion de ingreso, liquidacion con sufijo -DEV)
'
' DIFERENCIA con M_125_CTA_Mail_DEV_Done:
'   M_225 = liquidacion TPV (ingreso por TPV/Bizum)
'   M_125 = liquidacion CTA (ingreso por transferencia bancaria)
'
' REQUIERE:
'   - M_800_Mail_Send_New : Rut_Email_Send
'   - M_815_Mail_HTML     : Fnc_HTML_Tabla
'   - M_820_Range_TO_HTML : Fnc_RangeToHTML
'
' NOTA HISTORICA:
'   Hasta 2024 el envio fallaba por el 2FA de Google. Se generaba el correo en la hoja
'   "TPVMailDevDone" para copiarlo manualmente. Resuelto en 2026 con App Password (CDO).
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' Rut_Email_TPV_DEV_Done
'   Construye y envia el e-mail de confirmacion de DEV realizada para una liquidacion TPV
'--------------------------------------------------------------------------------------------------
Public Sub Rut_Email_TPV_DEV_Done()

    Dim DirMail     As String
    Dim Asunto      As String
    Dim Mail_Body   As String
    Dim Lin_JyC     As Long
    Dim Rng         As Range
    Dim aTabla      As Variant
    Dim Lo_Liq      As ListObject:  Set Lo_Liq = H_Liq_TPV.ListObjects(1)

    ' --- Validaciones previas ---
    If Not Fnc_Validar_Campos_DEV() Then Exit Sub

    Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    Call Rut_LstObj_Buscar(Lo_Lst, Range("Liq_Siglas"), 1, Lin_JyC)
    If Lin_JyC = 0 Then MsgBox "Siglas JyC, not find": Exit Sub
    Application.ScreenUpdating = False

    ' --- Rango de la tabla de liquidacion (todas las columnas hasta Obs) ---
    Set Rng = Lo_Liq.ListColumns(1).Range.Resize(, C_TPV_Liq_Obs)

    ' --- Destinatario y asunto ---
    DirMail = Range("Liq_Email") & ";facturacion@ua.es"
    Asunto = "JyC DEV_Done, Solicitud " & UCase(Range("Liq_Solicitud")) & _
             ", Liq. nº " & Range("Liq_Núm") & "-DEV, " & Range("Liq_Siglas")

    '- Este trozo de codigo genera el contenido del Mail en una Sheet,
    '- para poder copiar y pegar cuando falla la programacion VBA de envio de correo.
    '- En 2024 con el 2FA fallo y se arreglo en 2026, mientras tanto se uso esta opcion.
        'Const SheetMail     As String = "TPVMailDevDone"
        'With Sheets(SheetMail)  ...  'End With

    ' --- Tabla de datos del cuerpo del correo ---
    aTabla = Array( _
        Array("", "Devolución realizada:", _
              "Crimson", "JyC Solicitud " & UCase(Range("Liq_Solicitud")) & _
                      ", Liquidación nº " & Range("Liq_Núm") & _
                      "-DEV, JyC " & Range("Liq_Siglas")), _
        Array("", "Actividad:", _
              "blue", Range("Liq_Nombre")), _
        Array("", "Liquidación número:", _
              "blue", "nº " & Range("Liq_Núm") & " - DEV"), _
        Array("", "Importe Devolución:", _
              "blue", Format(Range("Liq_RDT_DI_Imp"), "#,##0.00") & " &euro;") _
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
    MsgBx_Title = "Proceso: Enviar E-mail de Confirmación de Devolución realizada."
    If CBool(Prog__APP.Range("SW_Test").Value) Then
        MsgBx_Msg = "Email enviado en modo prueba!"
    Else
        MsgBx_Msg = "Email enviado!"
    End If
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, , "OK", , "MailSent"): Form_MsgBox.Show

    Application.ScreenUpdating = True

End Sub     '- Rut_Email_TPV_DEV_Done ---


'--------------------------------------------------------------------------------------------------
' Fnc_Validar_Campos_DEV  (privada)
'   Verifica que los campos obligatorios de la liquidacion esten informados
'   Devuelve TRUE si todo es correcto, FALSE si falta algun campo
'--------------------------------------------------------------------------------------------------
Private Function Fnc_Validar_Campos_DEV() As Boolean

    Dim aValidaciones   As Variant
    Dim i               As Integer
    Dim TITULO          As String:  TITULO = "Proceso: Enviar E-mail de Confirmación de Devolución realizada."

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
            Fnc_Validar_Campos_DEV = False
            Exit Function
        End If
    Next i

    Fnc_Validar_Campos_DEV = True

End Function     '- Fnc_Validar_Campos_DEV ---
