Attribute VB_Name = "Rut_Wb_CopSegTimed_USB_HD"
' Last Rev. 2026-09-30 18:25
Option Explicit

' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
'===================================================================================================================================
Sub Rut_WrkBook_CopSegTimed_USB(Optional Tipo As String = "")  '- Guarda Copia de Este Excel con marca de tiempo en el nombre del archivo. (Tipo="Data"/"VBA", etc.) ----------
Debug.Print "Rut_WrkBook_CopSegTimed_USB,   Tipo: " & Tipo
    Dim Answer          As VbMsgBoxResult
    Dim FichNom         As String
        FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    Dim FichExt         As String
        FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
    Dim fichPath        As String
        'fichPath = ThisWorkbook.Path & "\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
        If Not Fnc_Range_Exist("APP_CopSeg_Usb_Path") Then
            MsgBox "¡¡¡ Falta crear el Range('APP_CopSeg_Usb_Path') !!!", vbExclamation, "Procedimiento: Copia de Seguridad"
            fichPath = "F:\__CopSeg Versiones Programas\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
        Else
            fichPath = Range("APP_CopSeg_Usb_Path") & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
        End If
    Dim FichSelect      As Variant
        FichSelect = Application.GetSaveAsFilename(fichPath, "Excel Files (*" & FichExt & "), *" & FichExt)
        
        If FichSelect <> False Then
            On Error GoTo Finalizar
            Application.DisplayAlerts = False
            ThisWorkbook.SaveCopyAs Filename:=FichSelect 'ConflictResolution:=True     '??? no se lo que hace, habría que investigar
            Range("APP_CopSeg_Usb") = Now()
            Application.DisplayAlerts = True
            On Error GoTo 0
        Else
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & "Proceso Abortado: " & Format(Now(), "dd-mmm-yy hh:mm")
            GoTo Finalizar
        End If
    If Fnc_Range_Exist("APP_CopSeg_Usb_Path") Then
        If Range("APP_CopSeg_Usb_Path") <> Left(FichSelect, InStrRev(FichSelect, "\")) Then
            ' Pedir confirmación
            Dim Mensage     As String
            Mensage = "¿ Cambiamos esta ruta: " & Range("APP_CopSeg_Usb_Path") & vbLf & _
                      " por esta ? " & Left(FichSelect, InStrRev(FichSelect, "\")) & vbCrLf & vbCrLf
            Answer = MsgBox(Mensage, vbExclamation + vbYesNo + vbDefaultButton2, "Proceso: Cambio de Ruta para las Copias de Seguridad.")
            If Answer = vbYes Then
                Range("APP_CopSeg_Usb_Path") = Left(FichSelect, InStrRev(FichSelect, "\"))
            End If
        End If
    End If
    Debug.Print Left(FichSelect, InStrRev(FichSelect, "\"))
    Prog__APP.Range("APP_Task_Inf") = "Copia Realizada en la Carpeta del USB: " & FichSelect & vbCrLf & String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm")

'    ThisWorkbook.Close savechanges:=True
Finalizar:
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------









