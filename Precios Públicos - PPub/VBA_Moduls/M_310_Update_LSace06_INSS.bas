Attribute VB_Name = "M_310_Update_LSace06_INSS"
'Rev.: 2026-01-22
'- M_310_Update_LSace06_INSS -----------------------------------------------------------------------------------------------------------

Option Explicit

' ==================================================================================================================================
            Sub Call_RuT_Update_LSace06_CAcad_ImpAdm_INSS()
                Debug.Print "================== >>> Call_RuT_Update_LSace06_CAcad_ImpAdm_INSS"
                Dim ActivSheet  As String:  ActivSheet = ThisWorkbook.ActiveSheet.Name
                Prog__APP.Range("APP_Task_Rut") = "RuT_Update_LSace06_CAcad_ImpAdm_INSS"
            On Error GoTo ManejoError
                DoEvents ' Permite que Excel procese eventos pendientes
                Form_Running_Rut.Show
            On Error GoTo 0
                ThisWorkbook.Sheets(ActivSheet).Select
                Debug.Print "================== <<< RuT_Update_LSace06_CAcad_ImpAdm_INSS"
            Exit Sub
ManejoError:
                Static Intentos As Integer
                If Err.Number = -2147417848 And Intentos < 5 Then
                    Intentos = Intentos + 1
                    Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
                    Resume ' Reintenta la línea que falló
                Else
                    MsgBox "Error: " & Err.Description & vbCrLf & "Intentos: " & Intentos, vbCritical
                    Intentos = 0
                End If
                MsgBox "<<< Err_Rut Form_Running_Rut >>>"
            End Sub

'==================================================================================================================================
Sub RuT_Update_LSace06_CAcad_ImpAdm_INSS()  '- Importar LSace06 por Curso_Acad, para Identificar los recibos con Tasa Adm. del seguro obligatorio del INSS
'==================================================================================================================================
Debug.Print "------------------------- >>> RuT_Update_LSace06_CAcad_ImpAdm_INSS()"
    Dim TimeLap2        As Single
    Dim NomFichLSace06  As String
    Dim RutaFichLsace06 As String
    Dim TxT_Progreso    As String
    Dim Arch_New_Name   As String
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
    Application.ScreenUpdating = False
    
    '- Setting ListObjects ------------------------------------
    Dim Lo_INSS             As ListObject:      Set Lo_INSS = Sht__BD_INSS.ListObjects(1)
    Dim Lo_DefCol_LSace06   As ListObject:      Set Lo_DefCol_LSace06 = Prog_DefCol_LSace06.ListObjects(1)
    
    '- Setting Sheets ------------------------------------
    Sht__BD_INSS.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_INSS)
    Sht__BD_INSS.Unprotect
    Lo_INSS.ShowTotals = False
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
If Not Func_MsgBox_vbYesNo("¿ Importamos LSace06 Del Curso_Acad " & C_Acad_Ant & " o " & C_Acad_Pos & " ?" & vbLf & vbLf & _
                           "¡¡¡ O sólo copiamos los datos de BD_INSS a BDatos.  !!!") Then GoTo Rut_Copy_ImpINSS_en_BDatos
        
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Importar LSace06 Del Curso_Acad " & C_Acad_Ant & " o " & C_Acad_Pos & vbLf & _
                    "- Identificar Recibos con concepto Eco. Adm. del seguro obligatorio del INSS," & vbLf & _
                    "- Y añadir el dato a la tabla BDatos." & vbLf & _
                    Sht__BD_INSS.Range("c2") & vbLf & Sht__BD_INSS.Range("c3") & vbLf, 0)

    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import LSace06 del Curso_Acad_Ant ------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Arch_New_Name = "LSACE06_" & C_Acad_Ant
    Call Rut_Lo_Import_LoData_LoDefCol_LSace06(Lo_INSS, Lo_DefCol_LSace06, DefC_TitColGenInf, Arch_New_Name)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
        Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NomFichLSace06, RutaFichLsace06)

    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import LSace06 del Curso_Acad_Pos ------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Arch_New_Name = "LSACE06_" & C_Acad_Pos
    Call Rut_Lo_Import_LoData_LoDefCol_LSace06(Lo_INSS, Lo_DefCol_LSace06, DefC_TitColGenInf, Arch_New_Name)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
        Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NomFichLSace06, RutaFichLsace06)

Rut_Copy_ImpINSS_en_BDatos:
If Not Func_MsgBox_vbYesNo("¿ Trasladar el ImpINSS del C_Acad_Ant y C_Acad_Pos a BDatos ?" & vbLf & vbLf & _
                           "¡¡¡ Tienen que estar todos los Recibos que deben de estar, como los AE4.  !!!") Then GoTo Terminar
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_315_Copy_INSS_a_BD, Trasladar el ImpAdm, ImpAcad y ImpDto del C_Acad_Ant a BDatos --------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call Rut_Copy_ImpINSS_en_BDatos

Terminar:
    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)

    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe")
    Lo_INSS.ShowTotals = True
    Sht__BD_INSS.Calculate

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_INSS.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_INSS.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_INSS.Visible = xlSheetVeryHidden
'Sht__BD_INSS.Visible = xlSheetVeryHidden
Rut_On_Functions
    Application.Calculation = Sw_Calculation
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
'    Set ActivForm = Nothing
Debug.Print "------------------------- <<< Sub RuT_Update_LSGES04_IAdm_CAcadAnt()"
End Sub     ' RuT_Update_LSGES04_IAdm_CAcadAnt   --------------------------------------------------------------------------------------------
'===================================================================================================================================






