Attribute VB_Name = "M50_Inf_Cont_AE4x4"
'2026-02-16
'- M_410_Update_Lo_AE4 -----------------------------------------------------------------------------------------------------------

Option Explicit

'==================================================================================================================================
Sub RuT_Inf_Contable_Recibos_AE4x4()  '- Importar los 4 WB: EFP y CFC de AñoCon_Ant/Pos
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
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_CursAcad")
    Dim Rng_Informe     As Range:       Set Rng_Informe = Sht__BD_AE4x4.Range("g5")
    Dim Lo_DefCol_BD    As ListObject:  Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
    
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
    Application.ScreenUpdating = False
    
    Dim WsBuffer As Worksheet: Set WsBuffer = ThisWorkbook.Worksheets("Sheet_Buffer")  'hoja fija/oculta
    ' 1)Limpiar anterior tabla en Buffer
    If WsBuffer.ListObjects.Count > 0 Then WsBuffer.ListObjects(1).Delete
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Importar Rec. AE4 - EFPyAFC Del Curso_Acad " & C_Acad_Ant & " y " & C_Acad_Pos & vbLf & _
                    "- Sólo los Recibos del Año Contable: " & AñoCont & " (Del Acont_Cob = AñoCont-1), y añadir Rec. a tabla BDatos.", 0)

    '- Visualizo el progreso -------- "EFP_" & C_Acad_Ant
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Importarmos 1º: " & Arch_New_Name, 0)
    Arch_New_Name = "EFP_" & C_Acad_Ant
    Call Rut_Lo_Import_AE4x1(Sht__Buffer, Arch_New_Name, "BDatos", NameFile, PathFile, False)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores

    '- Visualizo el progreso -------- "EFP_" & C_Acad_Pos
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Importarmos 2º: " & Arch_New_Name, 0)
    Arch_New_Name = "EFP_" & C_Acad_Pos
    Call Rut_Lo_Import_AE4x1(Sht__Buffer, Arch_New_Name, "BDatos", NameFile, PathFile, False)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores

    '- Visualizo el progreso -------- "CFCyAFC_" & C_Acad_Ant
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Importarmos 3º: " & Arch_New_Name, 0)
    Arch_New_Name = "CFCyAFC_" & C_Acad_Ant
    Call Rut_Lo_Import_AE4x1(Sht__Buffer, Arch_New_Name, "BDatos", NameFile, PathFile, False)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores

    '- Visualizo el progreso -------- "CFCyAFC_" & C_Acad_Pos
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Importarmos 4º: " & Arch_New_Name, 0)
    Arch_New_Name = "CFCyAFC_" & C_Acad_Pos
    Call Rut_Lo_Import_AE4x1(Sht__Buffer, Arch_New_Name, "BDatos", NameFile, PathFile, False)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    
    ContRecibos = Lo_AE4.ListRows.Count

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








