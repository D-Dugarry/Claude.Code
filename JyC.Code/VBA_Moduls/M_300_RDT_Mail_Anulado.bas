Attribute VB_Name = "M_300_RDT_Mail_Anulado"
'M_225_TPV_Mail_DEV_Done
'###################################################################################################################################
'   Nota Importante: Tenemos que Activar en este libro de Excel en Herramientas->Referencias->Microsoft CDO for Windows 2000 library
'###################################################################################################################################
Option Explicit

'===================================================================================================================================
Sub Rut_Email_RDT_ANULADA()
    Const SheetMail     As String = "MailRDTAnulada"
    Dim DirMail         As String
    Dim Mail_Body       As String
    Dim Asunto          As String
    Dim Lin_JyC         As Long
    Dim Rng             As Range
    Dim Lo_Liq      As ListObject:      Set Lo_Liq = H_Liq_TPV.ListObjects(1)
    
    If Range("Liq_Solicitud") = "" Then
        MsgBx_Msg = "¡No hay Referencia de Solicitud!"
        MsgBx_Title = "Proceso: Enviar E-mail de Confirmación de Redistribución realizada."
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, , "OK", , "stop"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
        Exit Sub
    End If
    Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    Call Rut_LstObj_Buscar(Lo_Lst, Range("Liq_Siglas"), 1, Lin_JyC)
    If Lin_JyC = 0 Then MsgBox "Siglas JyC, not find": Exit Sub
    Application.ScreenUpdating = False
'    Prog__APP.Range("SW_Test") = True

    Set Rng = Lo_Liq.ListColumns(1).Range.Resize(, C_TPV_Liq_Obs) '.Select  '- Resize(,xxx) es para ampliar a xxx columnas
'    Set Rng = Rng.SpecialCells(xlCellTypeVisible)
    
    DirMail = Range("Liq_Email") ' & ";facturacion@ua.es"      '- Si más de uno "ingresos@gmail.ua.es;d.dugarry@gmail.com"
    
    Asunto = "JyC-RDT-ANULADA, Solicitud " & UCase(Range("Liq_Solicitud")) & ", Liquidación nº " & Range("Liq_Núm") & ", " & Range("Liq_Siglas") & " - ¡Anulada!"

    With Sheets(SheetMail)
        .Range("b1") = Asunto
        .Range("h2") = Range("Liq_Email")
        .Range("c7").WrapText = True
        .Range("c7") = InputBox("Motivo Anulación:", "Configuración email de anulación de solicitud.")
        .Range("c9") = "JyC Solicitud " & UCase(Range("Liq_Solicitud")) & ", Liquidación nº " & Range("Liq_Núm") & ", JyC " & Range("Liq_Siglas")
        .Range("c10") = Range("Liq_Nombre")
        .Range("c11") = "Liquidación nº " & Range("Liq_Núm")
        .Range("c9:c11").WrapText = False
    End With
    
    Mail_Body = "<font size=3 color=""Black""><b>Hola </b><br><br><font color=""IndianRed"">" & _
                "Procedemos a ANULAR su solicitud por defecto de forma o falta de datos y le rogamos nos vuelva a hacer una nueva solicitud.<br>" & _
                "Por favor, para cualquier duda, póngase en contacto con nosotros.<br>" _
                & "<b>Motivo anulación: </b>El ingreso de referencia 9924002532257 a nombre de Raquel Pinilla Gómez está repetido. (y el PDF de solicitud no está firmado digitalmente).<br><br>" _
                & "<font color=""teal""><b>Actividad:</b>   " & Range("Liq_Nombre") & " <br>" _
                & "<b>Liquidación: </b>   núm." & Range("Liq_Núm") & " <font color=""Black""><br><br>"
                
    'Call Rut_Email_Send(DirMail, Asunto, , Mail_Body)
    MsgBx_Title = "Proceso: Enviar E-mail para realizar la Redistribución solicitada."
    
    MsgBx_Msg = "¡Email enviado!"
    MsgBx_Title = "Proceso: Enviar E-mail de Confirmación de Redistribución realizada."
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, , "OK", , "MailSent"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
    
    Application.ScreenUpdating = True
    
End Sub     ' Rut_Email_TPV_DEV_Done --------------------------

