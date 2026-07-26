Attribute VB_Name = "Rut_Ws_Export_XLSM"
Option Explicit
'- ----------------------------------------------------------------------------------------------------------------------------
'- Seleccionar Excel y Exportarlo ---------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Export_WS_XlsM(WrkSht As Worksheet, _
                      Optional SheetNom As String = "", _
                      Optional Formato As Integer = 52, _
                      Optional Sw_DelShapes As Boolean = False, _
                      Optional AskDelRegs As Boolean = True)    '- Formato: 51=Xls, 52=Xlsm
Debug.Print ">>> Rut_Lo_Export"
Rut_Off_Functions
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    '- Indicar Nombre del Archivo y Ruta para almacenar ----------------------------------------------------------------------------
    Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    Dim IntialName      As String
    Dim FichSelect      As Variant
    If SheetNom = "" Then
        IntialName = WrkSht.Name & "_" & Format(Now, "(yymmdd_hhmm)") & ".xlsm"
    Else
        IntialName = SheetNom & ".xlsm"
    End If
'''        '- Visualizo el progreso --------
'''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Seleccionar el fichero y la ruta, para exportar la Hoja: " & IntialName, 0)
    '- Selecciono la Ruta y el Nombre, y grabo del Excel. ----------------------------------------------------------------------------------
    FichSelect = Application.GetSaveAsFilename(FPath & IntialName, "Excel Files (*.xlsm), *.xlsm")
        If FichSelect <> False Then
            On Error GoTo GestError
            '- Visualizo el progreso --------
'''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Seleccionado el fichero y la ruta, estamos exportando la Hoja: " & _
                                            WrkSht.CodeName & " (" & WrkSht.Name & ")" & vbLf & "Filename:=" & FichSelect, 0)
            '- Copio la Sheet entera. ------------------------------------------------------------------------------------------------------
            WrkSht.Copy     ' This creates a new workbook with the copied sheet
                '- Si Requerido, Borro Shapes -----------
                If Sw_DelShapes Then
                    Dim shp As Shape
                    For Each shp In ActiveWorkbook.Sheets(1).Shapes
                        shp.Delete
                    Next shp
                End If
            On Error GoTo 0
            Application.DisplayAlerts = False
            'ActiveWorkbook.SaveAs FichSelect, FileFormat:=Formato, ¿¿¿¿¿ConflictResolution:=True?????
            ActiveWorkbook.SaveAs FichSelect, FileFormat:=Formato    '- Formato: Optional [ 51=Xls ], 52=Xlsm
            ActiveWorkbook.Close SaveChanges:=False     ' Close the new workbook
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
        
'''        '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
'''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Exportada la Hoja: " & WrkSht.CodeName & " (" & WrkSht.Name & ")" & vbLf & "Filename:=" & FichSelect & vbLf & Now, LastTimeLap, , , , True)
    
    Dim DelRegs As VbMsgBoxResult
    If AskDelRegs Then
        DelRegs = MsgBox("Ya hemos Exportado la Hoja: " & WrkSht.CodeName & " (" & WrkSht.Name & ")" & vbLf & "Filename:=" & FichSelect & vbLf & vbLf & _
                            "¿Quieres Borrar los Registros de la Tabla de " & WrkSht.CodeName & " (" & WrkSht.Name & ")?", _
                            vbYesNo + vbQuestion + vbDefaultButton2, "Proceso: Exportar una Hoja a Excel.")
        If DelRegs = vbYes Then
            If Not WrkSht.ListObjects(1).DataBodyRange Is Nothing Then WrkSht.ListObjects(1).DataBodyRange.Delete
            Call Rut_WrkSheet_ReducirPeso(WrkSht)
'''                '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
'''                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Y se han borrado los datos de la tabla origen.", 0)
        Else
'''                '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
'''                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Y se han mantenido los datos de la tabla origen.", 0)
        End If
    End If
        
'''    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)
    GoTo Salir_Sub
GestError:
    MsgBox "Error Rut_Lo_Export ", Err.Number, Err.Description, Err.Source
    Debug.Print "Error Rut_Lo_Export ", Err.Number, Err.Description, Err.Source
    Prog__APP.Range("APP_Last_Export_WS") = "Error Rut_Lo_Export:    Err.Number= " & Err.Number & _
                                             ",   Err.Description= " & Err.Description & ",  Err.Source= " & Err.Source & _
                                            "la Hoja: " & WrkSht.Name & vbCrLf & "Filename:=" & FichSelect & vbCrLf & Now
    Debug.Print "Exportada la Hoja: " & WrkSht.Name & vbCrLf & "Filename:=" & FichSelect & vbCrLf & Now
Salir_Sub:
Rut_On_Functions
End Sub     ' Rut_WrkBook_Exportar_La_Liquidación
' ==================================================================================================================================






