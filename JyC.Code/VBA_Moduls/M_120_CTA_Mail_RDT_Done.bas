Attribute VB_Name = "M_120_CTA_Mail_RDT_Done"
'M_111_CTA_Mail_RDT_Done
'###################################################################################################################################
'   Nota Importante: Tenemos que Activar en este libro de Excel en Herramientas->Referencias->Microsoft CDO for Windows 2000 library
'###################################################################################################################################
Option Explicit

'===================================================================================================================================
Sub Rut_Email_CTA_RDT_Done()
    Const SheetMail     As String = "CTAMailRDTDone"
    Dim DirMail         As String
    Dim Mail_Body       As String
    Dim Asunto          As String
    Dim Lin_JyC         As Long
    Dim Lo_Liq      As ListObject:      Set Lo_Liq = H_Liq_CTA.ListObjects(1)
    
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
    
    '- Rango de la tabla de liquidación  ---
    Dim Rng             As Range
    Set Rng = Lo_Liq.ListColumns(1).Range.Resize(, C_Cta_Liq_Obs) '.Select  '- Resize(,xxx) es para ampliar a xxx columnas
    
    DirMail = Range("Liq_Email") & ";facturacion@ua.es"      '- Si más de uno "ingresos@gmail.ua.es;d.dugarry@gmail.com"
    
    Asunto = "JyC RDT_Done, Solicitud " & UCase(Range("Liq_Solicitud")) & ", Liquidación nº " & Range("Liq_Núm") & ", " & Range("Liq_Siglas")
    
    '- Genera Mail en una Sheet, para poder copiar y pegar, cuando falla la progración VBA de envío de correo, en 2024 con el 2fa falló y se arregló en 2026
    'With Sheets(SheetMail)
    '    .Range("b1") = Asunto
    '    .Range("h2") = Range("Liq_Email") & ", facturacion@ua.es"
    '    .Range("c4") = "JyC Solicitud " & UCase(Range("Liq_Solicitud")) & ", Liquidación nº " & Range("Liq_Núm") & ", JyC " & Range("Liq_Siglas")
    '    .Range("c5") = Range("Liq_Nombre")
    '    .Range("c6") = Range("Liq_Orgánica")
    '    .Range("c7") = "Liquidación nº " & Range("Liq_Núm")
    '    .Range("c8") = Format(Lo_Liq.TotalsRowRange(C_Cta_Liq_Imp), "#,##0.00 €")     '- Format(Range("Liq_RDT_DI_Imp"), "#,##0.00 €")
    '    .Range("c4:c7").WrapText = False
    '    .Range("a10:h" & .Cells(.Rows.Count, "A").End(xlUp).Row).EntireRow.Delete     '- ClearContents
    '    Rng.Copy (.Range("a10"))
    '    .Range("a" & .Cells(.Rows.Count, "A").End(xlUp).Row + 2) = "Saludos Dugarry,"
    '    .Range("a3").Copy
    '    .Range("a" & .Cells(.Rows.Count, "A").End(xlUp).Row).PasteSpecial Paste:=xlPasteFormats
    'End With
    
    Mail_Body = "<div style='line-height:1.8; font-family:Arial, sans-serif; font-size:14px;'>" & _
                "<b>Hola</b><br><br>" & _
                "<table style='border-collapse:collapse; font-size:14px;' cellpadding='2'>" & _
                "<tr>" & _
                    "<td style='white-space:nowrap;padding-right:15px;'><b>Liquidación realizada:</b></td>" & _
                    "<td><b><font color='Crimson'>JyC Solicitud " & UCase(Range("Liq_Solicitud")) & _
                        ", Liquidación nº " & Range("Liq_Núm") & _
                        ", JyC " & Range("Liq_Siglas") & "</font></b></td>" & _
                "</tr>" & "<tr>" & _
                    "<td style='white-space:nowrap;'><b>Actividad:</b></td>" & _
                    "<td><b><font color='blue'>" & Range("Liq_Nombre") & "</font></b></td>" & _
                "</tr>" & "<tr>" & _
                    "<td style='white-space:nowrap;'><b>Orgánica de destino:</b></td>" & _
                    "<td><b><font color='blue'>" & Range("Liq_Orgánica") & "</font></b></td>" & _
                "</tr>" & "<tr>" & _
                    "<td style='white-space:nowrap; vertical-align:top;'><b>Liquidación número:</b></td>" & _
                    "<td><b><font color='blue'>" & Range("Liq_Núm") & "</font></b></td>" & _
                "</tr>" & "<tr>" & _
                    "<td style='white-space:nowrap;'><b>Importe redistribuido:</b></td>" & _
                    "<td><b><font color='red'>" & Format(Lo_Liq.TotalsRowRange(C_Cta_Liq_Imp), "#,##0.00") & " &euro;" & "</font></b></td>" & _
                "</tr>" & _
                "</table><br>" & _
                Fnc_RangeToHTML(Rng) & ".<br><br>" & _
                "</div>"
    
    Call Rut_Email_Send(DirMail, Asunto, , Mail_Body)
    
    MsgBx_Title = "Proceso: Enviar E-mail de Confirmación de Redistribución realizada."
    If CBool(Prog__APP.Range("SW_Test").Value) Then MsgBx_Msg = "¡Email enviado en modo prueba!" Else MsgBx_Msg = "¡Email enviado!"
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, , "OK", , "MailSent"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
    
    Application.ScreenUpdating = True
    
End Sub     ' Rut_Email_RDT_Done --------------------------




