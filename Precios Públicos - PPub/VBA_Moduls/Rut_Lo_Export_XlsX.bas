Attribute VB_Name = "Rut_Lo_Export_XlsX"
'2025-12-20
Option Explicit


'- ----------------------------------------------------------------------------------------------------------------------------
    Sub Rut_Lo_Export_to_New_WB_ByHand()
        Rut_Lo_Export_to_New_WB (ActiveSheet.ListObjects(1))
    End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Exportar ListObject a New_WB
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Export_to_New_WB(ByVal Lo_Data As ListObject, _
                            Optional FichName As String)
    Dim WbNew       As Workbook
    Dim WsNew       As Worksheet
    Dim SaveNomArch       As Variant
    Dim RutaINI   As String
    Dim NomArch   As String
    Dim ArchFullName   As String
    Dim calcMode    As XlCalculation
    
    ' Proponer ruta = del libro actual
    If Len(ThisWorkbook.Path) > 0 Then
        RutaINI = ThisWorkbook.Path & Application.PathSeparator
    Else
        RutaINI = CurDir$ & Application.PathSeparator
    End If
    
    ' Proponer nombre = nombre de la tabla + .xlsx
    NomArch = Lo_Data.Name & ".xlsx"
    
    ' Cuadro Guardar como, con ruta y nombre sugeridos
    SaveNomArch = Application.GetSaveAsFilename( _
                    InitialFileName:=RutaINI & NomArch, _
                    FileFilter:="Excel (*.xlsx), *.xlsx")
    If SaveNomArch = False Then Exit Sub          ' Usuario cancela
    
    ArchFullName = CStr(SaveNomArch)
    
    ' Optimizar
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        calcMode = .Calculation
        .Calculation = xlCalculationManual
    End With
    
    On Error GoTo ErrHandler
    
    ' Crear nuevo libro y copiar solo la tabla
    Set WbNew = Workbooks.Add(xlWBATWorksheet)   ' Libro con 1 hoja
    Set WsNew = WbNew.Worksheets(1)
    
    ' Copiar rango completo de la tabla (incluye cabecera y datos)
    Lo_Data.Range.Copy Destination:=WsNew.Range("A1")
    
    ' Opcional: ajustar ancho de columnas
    WsNew.Columns.AutoFit
    
    ' Guardar como .xlsx
    WbNew.SaveAs Filename:=ArchFullName, FileFormat:=xlOpenXMLWorkbook
    Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(ArchFullName, FichName)
    
Fin:
    On Error Resume Next
    ' Cerrar libro nuevo (si quieres dejarlo abierto, comenta la siguiente línea)
    WbNew.Close SaveChanges:=False
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calcMode
    End With
    Exit Sub

ErrHandler:
    Resume Fin
End Sub









Sub Rut_Lo_Export_KKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKK(WrkSht As Worksheet, _
                      Optional SheetNom As String = "", _
                      Optional AskDelRegs As Boolean = True)
Debug.Print ">>> Rut_Lo_Export"
Rut_Off_Functions
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    '- Indicar Nombre del Archivo y Ruta para almacenar ----------------------------------------------------------------------------
    Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    Dim IntialName      As String
    Dim FichSelect      As Variant
    If SheetNom = "" Then
        IntialName = WrkSht.Name & "_" & Format(Now, "(yymmdd_hhmm)") & ".xlsx"
    Else
        IntialName = SheetNom & ".xlsx"
    End If
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Seleccionar el fichero y la ruta, para exportar la Hoja: " & IntialName, 0)
    '- Selecciono la Ruta y el Nombre, y grabo del Excel. ----------------------------------------------------------------------------------
    FichSelect = Application.GetSaveAsFilename(FPath & IntialName, "Excel Files (*.xlsx), *.xlsx")
        If FichSelect <> False Then
            On Error GoTo GestError
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Seleccionado el fichero y la ruta, estamos exportando la Hoja: " & _
                                            WrkSht.CodeName & " (" & WrkSht.Name & ")" & vbLf & "Filename:=" & FichSelect, 0)
            '- Copio la Sheet entera. ------------------------------------------------------------------------------------------------------
            WrkSht.Copy     ' This creates a new workbook with the copied sheet
            Application.DisplayAlerts = False
'            ActiveWorkbook.SaveAs FichSelect, FileFormat:=51, ConflictResolution:=True
            ActiveWorkbook.SaveAs FichSelect, FileFormat:=51
            ActiveWorkbook.Close SaveChanges:=False     ' Close the new workbook
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
        
        '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Exportada la Hoja: " & WrkSht.CodeName & " (" & WrkSht.Name & ")" & vbLf & "Filename:=" & FichSelect & vbLf & Now, LastTimeLap, , , , True)
    
    Dim GuardarReg As VbMsgBoxResult
    If AskDelRegs Then
        GuardarReg = MsgBox("Ya hemos Exportado la Hoja: " & WrkSht.CodeName & " (" & WrkSht.Name & ")" & vbLf & "Filename:=" & FichSelect & vbLf & vbLf & _
                            "¿Quieres Borrar los Registros de la Tabla de " & WrkSht.CodeName & " (" & WrkSht.Name & ")?", _
                            vbYesNo + vbQuestion + vbDefaultButton2, "Proceso: Exportar una Hoja a Excel.")
        If GuardarReg = vbYes Then
            If Not Sht__BD.ListObjects(1).DataBodyRange Is Nothing Then Sht__BD.ListObjects(1).DataBodyRange.Delete
            Call Rut_WrkSheet_LstObj_LiberarEspacio(WrkSht)
                '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Y se han borrado los datos de la tabla origen.", 0)
        Else
                '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Y se han mantenido los datos de la tabla origen.", 0)
        End If
    End If
        
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)
    GoTo Salir_Sub
GestError:
    Debug.Print "Error Rut_Lo_Export ", Err.Number, Err.Description, Err.Source
    Debug.Print "Exportada la Hoja: " & WrkSht & vbCrLf & "Filename:=" & FichSelect & vbCrLf & Now
Salir_Sub:
Rut_On_Functions
End Sub     ' Rut_WrkBook_Exportar_La_Liquidación
' ==================================================================================================================================




