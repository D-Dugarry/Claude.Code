Attribute VB_Name = "M_210_Update_LSGES04_C_Acad_Ant"
' Last Rev. 2026-10-04 13:09
'2026-01-23
'- M_210_Update_LSGES04_C_Acad -----------------------------------------------------------------------------------------------------------

Option Explicit

'    '- Importar LSGES04_GE por Curso_Acad_Ant, para hallar Imp.Acad. Imp.TAdm. e Imp.Dto del CAcad.Ant pagadas este AnoCont.
'
'    - Rut_Lo_Import_LoData_LoDefCol, Import LSGES04 del Curso_Acad_Ant
'    - Rut_Lo_ListColumns_ClearContents_DefC_ProtectData, Borrar por protección de Datos, Información sensible y no necesarias, según DefCol
'    - Rut_Lo_Format_LoData_LoDefColData, Format Sht__BD_IAdm_CAcadAnt
'    - M_211, Filtrar y Borrar Registros NO deseados:
'        - Borrar Recibos de Movimiento
'        - Borrar Recibos BD_C_Acad <> C_Acad_Ant
'        - Borrar Recibos con Imp.Rec. < 0
'        - Borrar Recibos NO Martrícula
'        - Borrar Recibos ANULADOS
'        - Borrar Recibos INVALIDADOS
'        - Borrar Incongruencias de Fechas
'        - Borrar Recibos de Matrículas de coste CERO
'    - M_212, Gestionar Duplicados
'    - M_214_Find_IAdm_CAcad_Ant, Identificar Reg. de ImpAcad. ImpAdm. e ImpDto - Y - Borrar Reg SIN esos Datos

'- (2026-10-04) Call_RuT_Update_LSGES04_IAdm_CAcadAnt eliminado: el boton del Ribbon abre el formulario de progreso
'  (Rut_Progreso_Abrir/Cerrar, M_90_Rutinas_Menu_Aux) y llama a la rutina directamente.
'==================================================================================================================================
Sub RuT_Update_LSGES04_IAdm_CAcadAnt()  '- Importar LSGES04_GE por Curso-Acad., para hallar Imp.Acad. Imp.TAdm. e Imp.Dto del CAcad.Ant pagadas este AnoCont.
'==================================================================================================================================
Debug.Print "------------------------- >>> RuT_Update_LSGES04_IAdm_CAcadAnt()"
    Dim TimeLap2        As Single
    Dim TxT_Progreso    As String
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad          As String:      C_Acad = Prog__APP.Range("APP_CursAcad")
    
    Set ActivForm = Form_Running_Rut                      '- lo abre el boton del Ribbon (Rut_Progreso_Abrir, M_90)
    Application.ScreenUpdating = False
    
    '- Setting ListObjects ------------------------------------
    Dim Lo_ImpAdmCAcadAnt   As ListObject:      Set Lo_ImpAdmCAcadAnt = Sht__BD_IAdm_CAcadAnt.ListObjects(1)
    Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
    
    '- Setting Sheets ------------------------------------
    Sht__BD_IAdm_CAcadAnt.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_IAdm_CAcadAnt)
'    Lo_DefCol_BD.TotalsRowRange(DefC_HiddenCol) = False
    Lo_DefCol_BD.ShowTotals = True
    Lo_DefCol_BD.TotalsRowRange(DefC_HiddenCol) = False
    Prog__APP_Switch.Range("Sw_Col_Hide_Sht__BD_IAdm_CAcadAnt") = False
    Sht__BD_IAdm_CAcadAnt.Unprotect
    Lo_ImpAdmCAcadAnt.ShowTotals = False
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Importar LSGES04 Del C_Acad_Ant " & C_Acad_Ant & vbLf & _
                    " Para extraer Recibos Académicos de TIO y EP. con Imp.Adm del Curso " & C_Acad_Ant & vbLf & _
                    " Y añadir el Imp.Acad. e Imp.Adm. a los Rec. de BDatos.", 0)

If Not Func_MsgBox_vbYesNo("¿ Importamos LSGES04 C_Acad_" & C_Acad_Ant & " ?" & vbLf & vbLf & _
                           "¡¡¡ NO tiene que estar formateado !!!" & vbLf & vbLf & _
                           "¡¡¡ se formatea en el proceso !!!") Then GoTo SalaAqui
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import LSGES04 del Curso_Acad_Ant ------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim Arch_New_Name         As String:    Arch_New_Name = "LSGES04_GE_SinDtos_Curso_" & C_Acad_Ant & "_BD" & "|LsGes04_C_Acad_" & C_Acad_Ant     '- El de siempre o el de Robot_PPub_Fusión
    Rut_Off_Functions   '- Antes lo hacía (sin cerrarlo) la rutina de importación; el On está en Restablecer_Valores
    Call Rut_Lo_Import_LoData_LoDefCol(Lo_ImpAdmCAcadAnt, Lo_DefCol_BD, DefC_TitColGenInf, Arch_New_Name)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    Prog__APP.Range("APP_Last_Import_CAcad") = Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD_IAdm_CAcadAnt.Name = "BD_ImpAdm_" & C_Acad_Ant

Rut_Lo_ListColumns_ClearContents_DefC_ProtectData:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Rut_Lo: Borrar por protección de Datos, Información sensible y no necesarias -------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    If Prog__APP_Switch.Range("Sw_ProteccionDatosActivado") Then
        Call Rut_Lo_ListColumns_ClearContents_DefC_ProtectData(Lo_ImpAdmCAcadAnt, Lo_DefCol_BD, DefC_ProtectData)
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Eliminada Información sensible.", LastTimeLap)
    Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "¡ Información sensible SIN Eliminar !", LastTimeLap)
    End If

Rut_Lo_Format:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Rut_Lo_Format Sht__BD_IAdm_CAcadAnt ---------------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Sht__BD_IAdm_CAcadAnt.Select
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            TxT_Progreso = ActivForm.Controls("TBx_Informe")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Formateando el Excel. (+-30s)", LastTimeLap)
            TimeLap2 = LastTimeLap
    Call Rut_Lo_Format_LoData_LoDefColData(Lo_ImpAdmCAcadAnt, Lo_DefCol_BD)
            LastTimeLap = TimeLap2
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Formateado el Excel, con  ", LastTimeLap, , Format(Lo_ImpAdmCAcadAnt.ListRows.Count, "#,##0") & " reg.", TxT_Progreso, , , , 2)
    
RuT_Remove_Null_Reg:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_211_Remove_Null_Reg_CAcad: Filtrar y Borrar Registros NO deseados --------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
            TimeLap2 = LastTimeLap
    Call RuT_Remove_Null_Reg_CAcad(Lo_ImpAdmCAcadAnt)
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_IAdm_CAcadAnt)
            LastTimeLap = TimeLap2
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "---- Eliminados los Recibos no necesarios,", LastTimeLap, , "quedan: " & Format(Lo_ImpAdmCAcadAnt.ListRows.Count, "#,##0") & " reg.", , , , , 2)
    If Lo_ImpAdmCAcadAnt.ListRows.Count = 0 Then GoTo Restablecer_Valores
    
RuT_Duplicates_Search:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_212_Manage_Duplicates_CAcad ----------------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call RuT_Duplicates_Search_and_Del(Lo_ImpAdmCAcadAnt, BD_Ref)
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_IAdm_CAcadAnt)
        
RuT_Find_Imp_AdmAcad_C_Acad:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_214_Find_IAdm_CAcad_Ant, Identificar Reg. de ImpAcad. ImpAdm. e ImpDto - Y - Borrar Reg SIN esos Datos -------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call RuT_Find_Imp_AdmAcad_C_Acad(Lo_ImpAdmCAcadAnt)
                
'GoTo Terminar
SalaAqui:

    
Rut_Copy_ImpAdm_CAcadAnt_a_BDatos:
    Application.Speech.Speak "¿Transladamos el importe administrativo a la base de datos?", True
If Not Func_MsgBox_vbYesNo("¿ Trasladar el ImpAdm, ImpAcad y ImpDto del C_Acad_Ant a BDatos ?") Then GoTo Terminar
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_215_Copy_IAdmCAcad_Ant_a_BD, Trasladar el ImpAdm, ImpAcad y ImpDto del C_Acad_Ant a BDatos --------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call Rut_Copy_ImpAdm_CAcadAnt_a_BDatos

Terminar:
    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)

    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe").Text
    Prog__APP.Range("APP_Last_Import_CAcad") = Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD_IAdm_CAcadAnt.Range("d2") = "Úlitma Importación: " & Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD.Range("b3") = " Tabla Rec. Acad. TIO y EP  +  Imp.Adm Curso " & Prog__APP.Range("APP_C_Acad_Pos") & _
                          " y Curso " & C_Acad & "."
    Call Rut_Lo_Filtros_Quitar(Lo_ImpAdmCAcadAnt)
    Lo_ImpAdmCAcadAnt.ShowTotals = True
    Sht__BD_IAdm_CAcadAnt.Calculate
    
Restablecer_Valores:
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_IAdm_CAcadAnt.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_IAdm_CAcadAnt.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_IAdm_CAcadAnt.Visible = xlSheetVeryHidden
'Sht__BD_IAdm_CAcadAnt.Visible = xlSheetVeryHidden
Rut_On_Functions
    Application.Calculation = Sw_Calculation
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
'    Set ActivForm = Nothing
Debug.Print "------------------------- <<< Sub RuT_Update_LSGES04_IAdm_CAcadAnt()"
End Sub     ' RuT_Update_LSGES04_IAdm_CAcadAnt   --------------------------------------------------------------------------------------------
'===================================================================================================================================




