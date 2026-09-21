Attribute VB_Name = "M15_Export_Liquid_to_xlsx"
' Last Rev. 2026-09-21 12:12
Option Explicit

' ==================================================================================================
Sub Rut_Exportar_La_Liquidación()   '- Copia una Sheet concreta
' ==================================================================================================
Rut_Off_Functions
Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Copio la Sheet entera y esto es lo que voy a grabar. ----------------------------------------
    Wk_TitP_Liquid.Copy
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "LIQ-TitProp_" & Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo GestError
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True, FileFormat:=51
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    '- Quito los Elementos: Botones (Shapes), Comentarios de Celdas y Borro la Fila de Filtrado (la de arriba de los títulos de la Tabla
        ActiveWorkbook.ActiveSheet.Unprotect
        ActiveWorkbook.ActiveSheet.Shapes.SelectAll:   Selection.Delete
        ActiveWorkbook.ActiveSheet.UsedRange.ClearComments
        Application.EnableEvents = False
        ActiveWorkbook.ActiveSheet.Rows(ActiveWorkbook.ActiveSheet.ListObjects(1).Range.Rows(1).Row - 1).Clear
    '- Grabo los cambios y Cierro el Archivo ---------------------------
    ActiveWorkbook.Close SaveChanges:=True
        MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Archivar Liquidación"
    Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & _
            "  -.-  " & Now() & vbCrLf & vbCrLf & "Exportado el Resumen de la Liquidación de:   " & Wk_TitP_Liquid.Range("Liquid_Plan_Name") & _
            vbCrLf & vbCrLf & "En el Archivo:   " & sFileSaveName
    GoTo Salir_Sub
GestError:
    Debug.Print "Error Rut_Exportar_La_Liquidación ", Err.Number, Err.Description, Err.Source
    Debug.Print sFileSaveName
    MsgBox "Rut_Exportar_La_Liquidación " & "Filename:=" & vbCrLf & sFileSaveName, vbExclamation + vbOKOnly, "Rutina de Remesado"
    Form_Menu.TB_Informe = "Rut_Exportar_La_Liquidación " & vbCrLf & "Filename:=" & sFileSaveName
Salir_Sub:
Call Rut_EnableEvents_Status_Reset
Rut_On_Functions
End Sub     ' Rut_Exportar_La_Liquidación
'---------------------------------------------------------------------------------------------------



