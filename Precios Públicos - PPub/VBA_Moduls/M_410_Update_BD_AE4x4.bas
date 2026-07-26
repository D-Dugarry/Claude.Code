Attribute VB_Name = "M_410_Update_BD_AE4x4"
'2026-01-017
'- M_410_Update_Lo_AE4 -----------------------------------------------------------------------------------------------------------

Option Explicit

' ==================================================================================================================================
            Sub Call_RuT_Update_AE4x4()
                Debug.Print "================== >>> Call_RuT_Update_AE4x4"
                Dim ActivSheet  As String:  ActivSheet = ThisWorkbook.ActiveSheet.Name
                Prog__APP.Range("APP_Task_Rut") = "RuT_Update_AE4x4"
            On Error GoTo ManejoError
                DoEvents ' Permite que Excel procese eventos pendientes
                Form_Running_Rut.Show
            On Error GoTo 0
                ThisWorkbook.Sheets(ActivSheet).Select
                Debug.Print "================== <<< Call_RuT_Update_AE4x4"
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
Sub RuT_Update_AE4x4()  '- Importar los 4 WB: EFP y CFC de AñoCon_Ant/Pos
'==================================================================================================================================
Debug.Print "------------------------- >>> RuT_Update_AE4x4()"
    Dim TimeLap2        As Single
    Dim ContRecibos     As Long
    Dim rowfind         As Variant
    Dim TxT_Progreso    As String
    Dim Arch_New_Name   As String
    Dim NameFile        As String
    Dim PathFile        As String
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim Rng_Informe     As Range:       Set Rng_Informe = Sht__BD_AE4x4.Range("g5")
    
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
    Application.ScreenUpdating = False
    
    '- Setting ListObjects ------------------------------------
    Dim Lo_AE4          As ListObject:      Set Lo_AE4 = Sht__BD_AE4x4.ListObjects(1)
    Dim Lo_DefCol_BD    As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
    
    '- Setting Sheets ------------------------------------
    Sht__BD_AE4x4.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_AE4x4)
    Sht__BD_AE4x4.Unprotect
    Sht__BD_AE4x4.Select
    Lo_AE4.ShowTotals = False
    Sht__BD_AE4x4.Range("f2:g5").ClearContents
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Importar Rec. AE4 - EFPyAFC Del Curso_Acad " & C_Acad_Ant & " y " & C_Acad_Pos & vbLf & _
                    "- Sólo los Recibos del Año Contable: " & AñoCont & " (Del Acont_Cob = AñoCont-1), y añadir Rec. a tabla BDatos.", 0)

If Not Func_MsgBox_vbYesNo("¿ Importamos WBs_AE4 ?" & vbLf & vbLf & _
                            "EFP_" & C_Acad_Ant & " y EFP_" & C_Acad_Pos & vbLf & vbLf & _
                            "CFCyAFC_" & C_Acad_Ant & " y CFCyAFC_" & C_Acad_Pos) Then GoTo Rut_Copy_AE4x4_en_BDatos
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import WB_AE4: EFP y CFCyAFC por Curso_Acad_Ant y Curso_Acad_Pos -----------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Arch_New_Name = "EFP_" & C_Acad_Ant
        '- Visualizo el progreso -------- "EFP_" & C_Acad_Ant
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Importarmos 1º: " & Arch_New_Name, 0)
    Call Rut_Lo_Import_LoData_LoDefCol_AE4x1(Lo_AE4, Lo_DefCol_BD, DefC_TitColLstObj, Arch_New_Name, NameFile, PathFile, "BDatos", True)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    Sht__BD_AE4x4.Range("d2") = "Úlitma Importación AE4: EFP_" & C_Acad_Ant & " - el " & Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD_AE4x4.Range("b2") = C_Acad_Ant
    Call Rut_Lo_Filtros_Quitar(Lo_AE4)
    ContRecibos = Lo_AE4.ListRows.Count
    Sht__BD_AE4x4.Range("f2") = ContRecibos
    Sht__BD_AE4x4.Range("g2") = Rng_Informe
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copiados los nuevos recibos:", LastTimeLap)
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "en BD_AE4x4 Rec. de " & NameFile, 0, _
                    Format(ContRecibos, "#,##0") & " reg", "Total: " & Format(Lo_AE4.ListRows.Count, "#,##0") & " reg")

    Arch_New_Name = "EFP_" & C_Acad_Pos
        '- Visualizo el progreso -------- "EFP_" & C_Acad_Pos
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Importarmos 2º: " & Arch_New_Name, 0)
    Call Rut_Lo_Import_LoData_LoDefCol_AE4x1(Lo_AE4, Lo_DefCol_BD, DefC_TitColLstObj, Arch_New_Name, NameFile, PathFile, "BDatos", False)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    Sht__BD_AE4x4.Range("d3") = "Úlitma Importación AE4: EFP_" & C_Acad_Pos & " - el " & Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD_AE4x4.Range("b3") = C_Acad_Pos
    Call Rut_Lo_Filtros_Quitar(Lo_AE4)
    ContRecibos = Lo_AE4.ListRows.Count - ContRecibos
    Sht__BD_AE4x4.Range("f3") = ContRecibos
    Sht__BD_AE4x4.Range("g3") = Rng_Informe
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copiados los nuevos recibos:", LastTimeLap)
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "en BD_AE4x4 Rec. de " & NameFile, 0, _
                    Format(ContRecibos, "#,##0") & " reg", "Total: " & Format(Lo_AE4.ListRows.Count, "#,##0") & " reg")
    ContRecibos = Lo_AE4.ListRows.Count

    Arch_New_Name = "CFCyAFC_" & C_Acad_Ant
        '- Visualizo el progreso -------- "CFCyAFC_" & C_Acad_Ant
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Importarmos 3º: " & Arch_New_Name, 0)
    Call Rut_Lo_Import_LoData_LoDefCol_AE4x1(Lo_AE4, Lo_DefCol_BD, DefC_TitColLstObj, Arch_New_Name, NameFile, PathFile, "BDatos", False)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    Sht__BD_AE4x4.Range("d4") = "Úlitma Importación AE4: CFCyAFC_" & C_Acad_Ant & " - el " & Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD_AE4x4.Range("b4") = C_Acad_Ant
    Call Rut_Lo_Filtros_Quitar(Lo_AE4)
    ContRecibos = Lo_AE4.ListRows.Count - ContRecibos
    Sht__BD_AE4x4.Range("f4") = ContRecibos
    Sht__BD_AE4x4.Range("g4") = Rng_Informe
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copiados los nuevos recibos:", LastTimeLap)
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "en BD_AE4x4 Rec. de " & NameFile, 0, _
                    Format(ContRecibos, "#,##0") & " reg", "Total: " & Format(Lo_AE4.ListRows.Count, "#,##0") & " reg")
    ContRecibos = Lo_AE4.ListRows.Count

    Arch_New_Name = "CFCyAFC_" & C_Acad_Pos
        '- Visualizo el progreso -------- "CFCyAFC_" & C_Acad_Pos
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Importarmos 4º: " & Arch_New_Name, 0)
    Call Rut_Lo_Import_LoData_LoDefCol_AE4x1(Lo_AE4, Lo_DefCol_BD, DefC_TitColLstObj, Arch_New_Name, NameFile, PathFile, "BDatos", False)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    Sht__BD_AE4x4.Range("d5") = "Úlitma Importación AE4: CFCyAFC_" & C_Acad_Pos & " - el " & Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD_AE4x4.Range("b5") = C_Acad_Pos
    Call Rut_Lo_Filtros_Quitar(Lo_AE4)
    ContRecibos = Lo_AE4.ListRows.Count - ContRecibos
    Sht__BD_AE4x4.Range("f5") = ContRecibos
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copiados los nuevos recibos:", LastTimeLap)
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "en BD_AE4x4 Rec. de " & NameFile, 0, _
                    Format(ContRecibos, "#,##0") & " reg", "Total: " & Format(Lo_AE4.ListRows.Count, "#,##0") & " reg")
    ContRecibos = Lo_AE4.ListRows.Count

        
Rut_Copy_AE4x4_en_BDatos:
If Not Func_MsgBox_vbYesNo("¿ Copiamos en BDatos los nuevos Rec. AE4 ?" & vbLf & vbLf & _
                            "EFP_" & C_Acad_Ant & " y EFP_" & C_Acad_Pos & vbLf & vbLf & _
                            "CFCyAFC_" & C_Acad_Ant & " y CFCyAFC_" & C_Acad_Pos) Then GoTo Terminar
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_415_Copy_AE4_a_BD, Copia los Rec. AE4 del C_Acad Ant y Pos a BDatos ------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call Rut_Copy_AE4x4_en_BDatos

Terminar:
    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe")
    
    Prog__APP.Range("APP_Last_Import_AE4") = Format(Now(), "dd-mmm-yy hh:mm")
    Lo_AE4.ShowTotals = True
    Sht__BD_AE4x4.Calculate

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_AE4x4.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_AE4x4.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_AE4x4.Visible = xlSheetVeryHidden
'Sht__BD_AE4x4.Visible = xlSheetVeryHidden
Lo_AE4.ShowTotals = True
Rut_On_Functions
    Application.Calculation = Sw_Calculation
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
'    Set ActivForm = Nothing
Application.Speech.Speak "Proceso completado."
Debug.Print "------------------------- <<< Sub RuT_Update_LSGES04_IAdm_CAcadAnt()"
End Sub     ' RuT_Update_AE4x4   --------------------------------------------------------------------------------------------
'===================================================================================================================================








