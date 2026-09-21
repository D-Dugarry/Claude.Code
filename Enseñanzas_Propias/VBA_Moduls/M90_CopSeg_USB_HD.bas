Attribute VB_Name = "M90_CopSeg_USB_HD"
' Last Rev. 2026-09-21 12:12
Option Explicit

'' =================================================================================================
'Sub CopSegTimed_WorckBook_USB()  '- Guarda Copia de Este Excel con marca de tiempo en el nombre del archivo.
'
'    Dim FichNom                             As String
'        FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
'    Dim FichExt                              As String
'        FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
'    Dim fichPath                               As String
'        fichPath = ThisWorkbook.Path & "\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & FichExt
'        fichPath = "E:\__CopSeg Versiones Programas\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & FichExt
'
'    Dim FichSelect As Variant
'        FichSelect = Application.GetSaveAsFilename(fichPath, "Excel Files (*" & FichExt & "), *" & FichExt)
'
'        If FichSelect <> False Then
'            On Error GoTo Finalizar
'            Application.DisplayAlerts = False
'            ThisWorkbook.SaveCopyAs Filename:=FichSelect
'            Application.DisplayAlerts = True
'            On Error GoTo 0
'        End If
'
'    Form_Menu.TB_Informe = "Copia Realizada en la Carpeta del USB: " & FichSelect & vbCrLf & Now
'
''    ThisWorkbook.Close savechanges:=True
'Finalizar:
'End Sub
'' =================================================================================================
'
' --------------------------------------------------------------------------------------------------



