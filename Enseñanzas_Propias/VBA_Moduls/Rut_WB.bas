Attribute VB_Name = "Rut_WB"
Option Explicit

' ==================================================================================================================================
Sub Rut_Gen_Excel_EFP_CFC_WorckBook()  '- Guarda Copia de Este Excel con marca de tiempo en el nombre del archivo. ----------
' ==================================================================================================================================
    Dim FichNom                             As String
        FichNom = Range("APP_EFP_o_CFC") & "_" & Range("APP_CursAcad") & "_BaseDatos-Liq_V2.2(3)"
    Dim FichExt                              As String
        FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
    Dim FichPath                               As String
        FichPath = ThisWorkbook.Path & "\" & FichNom & FichExt
'        FichPath = "E:\__CopSeg Versiones Programas\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & FichExt
    Dim Response As VbMsgBoxResult
    
    ' Verificar si el archivo ya existe -----------
    If Func_ArchivoExiste(FichPath) Then
        ' Preguntar si desea reemplazar
        Response = MsgBox("El Archivo ya Existe. ¿desea reemplazarlo?", vbYesNo + vbQuestion, "Generar Fichero Proceso.")
        
        If Response = vbNo Then
            MsgBox "Operación cancelada por el usuario.", vbInformation, "Generar Fichero Proceso."
            Exit Sub
        End If
    End If
'    ' Guardar Copia Segura ----------
'        ' Obtener ruta temporal única
'    Dim TempFilePath As String
'    Dim FolderPath As String
'    FolderPath = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\"))
'    TempFilePath = FolderPath & "temp_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm"
'        ' Guardar el archivo actual como copia
'    ThisWorkbook.SaveCopyAs TempFilePath
'        ' Abrir la copia temporal y guardarla con el nombre final
'    Dim wbCopia As Workbook
'    Set wbCopia = Workbooks.Open(TempFilePath)
'        ' Guardar con el nombre final (esto reemplazará si existe)
'    Application.DisplayAlerts = False ' Desactivar alertas de reemplazo
'    wbCopia.SaveAs fichPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    Application.DisplayAlerts = True
'        ' Cerrar la copia sin guardar cambios
'    wbCopia.Close SaveChanges:=False
'        ' Eliminar el archivo temporal
'    On Error Resume Next
'    Kill TempFilePath
'    On Error GoTo 0
'    Form_Menu.TB_Informe = "Copia Realizada en la Carpeta actual: " & fichPath & vbCrLf & Now

    
    ' Generar la Copia
    Dim FichSelect As Variant
        FichSelect = Application.GetSaveAsFilename(FichPath, "Excel Files (*" & FichExt & "), *" & FichExt)

        If FichSelect <> False Then
            On Error GoTo Finalizar
            Application.DisplayAlerts = False
            ThisWorkbook.SaveCopyAs Filename:=FichSelect ', ConflictResolution:=True     '??? no se lo que hace, habría que investigar
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    Form_Menu.TB_Informe = "Copia Realizada en la Carpeta actual: " & FichSelect & vbCrLf & Now
    

'    ThisWorkbook.Close savechanges:=True
Finalizar:
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
' Función para obtener ruta temporal única
Private Function Func_ObtenerRutaTemporal() As String
' ==================================================================================================================================
    Dim FolderPath As String
    FolderPath = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\"))
    
    Func_ObtenerRutaTemporal = FolderPath & "temp_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm"
End Function
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
' Función para verificar si un archivo existe
Private Function Func_ArchivoExiste(FichPath As String) As Boolean
' ==================================================================================================================================
    On Error Resume Next
    Func_ArchivoExiste = (Dir(FichPath) <> "")
    On Error GoTo 0
End Function
' ----------------------------------------------------------------------------------------------------------------------------------

