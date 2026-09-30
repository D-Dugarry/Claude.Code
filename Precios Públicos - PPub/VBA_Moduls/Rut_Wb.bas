Attribute VB_Name = "Rut_Wb"
' Last Rev. 2026-09-30 18:25
Option Explicit
'
'' ==================================================================================================================================
'Sub Rut_WrkBook_CopSegTimed()  '- Guarda Copia de Este Excel con marca de tiempo en el nombre del archivo. ----------
'
'    Dim FichNom                             As String
'        FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
'    Dim FichExt                              As String
'        FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
'    Dim fichPath                               As String
'        fichPath = ThisWorkbook.Path & "\" & FichNom & " " & Format(Now, "ddmmmyy_hhmm") & FichExt
'
'    Dim FichSelect As Variant
'        FichSelect = Application.GetSaveAsFilename(fichPath, "Excel Files (*" & FichExt & "), *" & FichExt)
'
'        If FichSelect <> False Then
'            On Error GoTo Finalizar
'            Application.DisplayAlerts = False
'            Debug.Print FichSelect
'            ThisWorkbook.SaveCopyAs Filename:=FichSelect 'ConflictResolution:=True     '??? no se lo que hace, habría que investigar
'            Application.DisplayAlerts = True
'            On Error GoTo 0
'            MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Copia de Seguridad"
'        End If
'
''    ThisWorkbook.Close savechanges:=True
'Finalizar:
'End Sub
'' ==================================================================================================================================

'E:\__CopSeg Versiones Programas

' ==================================================================================================================================
'- Eliminado: Workbook_Open en un modulo estandar NUNCA se dispara (ese evento solo funciona en ThisWorkbook.cls).
'- El arranque real ya vive, correcto, en ThisWorkbook.cls, que llama a Rut_WrkBook_MinimizeAllExcelExceptThisWB (definida abajo).
' ==================================================================================================================================
Sub Rut_WrkBook_MinimizeAllExcelExceptThisWB()  '- It's working ok.
Debug.Print "Rut_WrkBook_MinimizeAllExcelExceptThisWB"
    Dim wb As Workbook
    Application.ScreenUpdating = False
    For Each wb In Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            wb.Activate
            wb.Windows.Application.WindowState = xlMinimized
        End If
    Next wb
    ThisWorkbook.Activate
    ActiveWindow.WindowState = xlMaximized
'    Application.ScreenUpdating = True
End Sub
' ==================================================================================================================================






