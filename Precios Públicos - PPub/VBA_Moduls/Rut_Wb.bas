Attribute VB_Name = "Rut_Wb"
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
Sub Rut_WrkBook_MinimizeAllWindowsExceptThisExcel()     '- Minimiza Todas las ventanas de cualquier programa...
Debug.Print "Rut_WrkBook_MinimizeAllWindowsExceptThisExcel"
Dim objShell As Object
    Set objShell = CreateObject("Shell.Application")
    objShell.MinimizeAll
    'Optional, if your system is slow to minimize windows,
    'or you typically have a lot of windows open.
''    Application.Wait (Now + TimeValue("0:00:01"))
    Application.WindowState = xlMaximized
    Set objShell = Nothing
End Sub
' ==================================================================================================================================
'- Eliminado: Workbook_Open en un modulo estandar NUNCA se dispara (ese evento solo funciona en ThisWorkbook.cls).
'- El arranque real ya vive, correcto, en ThisWorkbook.cls, que llama a Rut_WrkBook_MinimizeAllExcelExceptThisWB (definida abajo).
' ==================================================================================================================================
Sub Rut_WrkBook_MinimizeOnlyThisWB()    '- It's working ok.
Debug.Print "Rut_WrkBook_MinimizeOnlyThisWB"
Dim wb As Workbook
    Application.ScreenUpdating = False
    For Each wb In Workbooks
        If wb.Name = ThisWorkbook.Name Then
            wb.Activate
            ActiveWindow.WindowState = xlMinimized
            Exit For
        End If
    Next wb
    Application.ScreenUpdating = True
End Sub
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
Sub Rut_WrkBook_BringWindowToFront()
    Dim wb As Workbook
    Dim wbName As String
    ' Set the name of the workbook you want to bring to the front
    wbName = ThisWorkbook.Name ' Change this to the name of your workbook
    For Each wb In Workbooks
        If wb.Name = wbName Then
            wb.Activate     '- Activate the workbook to bring it to the front
'            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWindow.WindowState = xlMaximized
            Exit Sub
        End If
    Next wb
    ' If the workbook is not found, display a message
    MsgBox "Workbook '" & wbName & "' not found!"
End Sub
' ==================================================================================================================================





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''RUTINAS PARA HACER GENÉRICAS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''FALTA ADAPTARLAS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ==================================================================================================================================
Sub Rut_WrkBook_Exportar_La_Liquidación()   '- Copia una Sheet concreta
' ==================================================================================================================================
Rut_Off_Functions
Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    Hora_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Copio la Sheet entera y esto es lo que voy a grabar. -----------------------------------------
    Wk_TitP_Liquid.Copy
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "Liquid_" & Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo GestError
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True, FileFormat:=51
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    '- Quito los Elementos: Botones (Shapes), Comentarios de Celdas y Borro la Fila de Filtrado (la de arriba de los títulos de la Tabla ------------
        ActiveWorkbook.ActiveSheet.Unprotect
        ActiveWorkbook.ActiveSheet.Shapes.SelectAll:   Selection.Delete
        ActiveWorkbook.ActiveSheet.UsedRange.ClearComments
        'Application.EnableEvents = False
        ActiveWorkbook.ActiveSheet.Rows(ActiveWorkbook.ActiveSheet.ListObjects(1).Range.Rows(1).Row - 1).ClearContents
    '- Grabo los cambios y Cierro el Archivo ---------------------------
    ActiveWorkbook.Close SaveChanges:=True
        MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Archivar Liquidación"
    Prog__APP.Range("APP_Task_Inf") = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Format(Timer - Hora_Inicio, "0.00") & _
            " seg.  -.-  " & Now() & vbCrLf & vbCrLf & "Exportado el Resumen de la Liquidación de:   " & Wk_TitP_Liquid.Range("Liquid_Plan_Name") & _
            vbCrLf & vbCrLf & "En el Archivo:   " & sFileSaveName
    GoTo Salir_Sub
GestError:
    Debug.Print "Error Rut_Exportar_La_Liquidación ", Err.Number, Err.Description, Err.Source
    Debug.Print sFileSaveName
    MsgBox "Rut_Exportar_La_Liquidación " & "Filename:=" & vbCrLf & sFileSaveName, vbExclamation + vbOKOnly, "Rutina de Remesado"
    Prog__APP.Range("APP_Task_Inf") = "Rut_Exportar_La_Liquidación " & vbCrLf & "Filename:=" & sFileSaveName
Salir_Sub:
Rut_On_Functions
End Sub     ' Rut_WrkBook_Exportar_La_Liquidación
' ==================================================================================================================================





