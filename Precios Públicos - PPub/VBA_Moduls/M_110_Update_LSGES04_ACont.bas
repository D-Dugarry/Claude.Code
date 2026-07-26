Attribute VB_Name = "M_110_Update_LSGES04_ACont"
'Rev.: 2026-01-22
'- M_110_Load_LSGES04_ACont -----------------------------------------------------------------------------------------------------------
Option Explicit

'    Importar Última Consulta de LSGES04_GE, Actualizar registros existentes y Añadir Nuevos.
'
'    - Rut_Lo_Import_LoData_LoDefCol, Import LSGES04 por Año_Contable en Sht_BD
'    - Rut_Lo_ListColumns_ClearContents_DefC_ProtectData, Borrar por protección de Datos, Información sensible y no necesarias, según DefCol
'    - Rut_Lo_Format_LoData_LoDefColData, Format Sht__BD
'    - M_111, Filtrar y Borrar Registros NO deseados:
'        - Borrar Recibos AE4 Enseñanzas Propias
'        - Borrar Recibos de Matrículas de coste CERO
'        - Borrar Recibos con Imp.Rec. < 0
'        - Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"
'        - Borrar Recibos ANULADOS
'        - Borrar Recibos NO Martrícula
'        - Borrar Recibos INVALIDADOS
'        - Borrar Incongruencias de Fechas
'        - Borrar Recibos con fechas FUERA DEL PERÍODO CONTABLE, ¡¡ o Borrar Datos del cobro !!
'    - M_112, Gestionar Duplicados
'    - M_113, Asignar Código Concepto-Eco del Rec Y Rellenar Col Cta_Ingreso con el nº de Cta. correspondiente y Año de Vencimiento en ACont_Vto
'    - M_114, Identificar y Asignar al primer registro de la matrícula el Importe Académico y el Administrativo
'    - M_116, Clasificar Recibos en Emitidos, Aplazaados, EjeAnt, ADxAplz, Añejos
'    - M_215_Copy_IAdmCAcad_Ant_a_BD, Trasladar el ImpAdm, ImpAcad y ImpDto del C_Acad_Ant a BDatos
'    - Import BDatos de una versión anterior y Copiar en BDatos_Ant
'        - Actualizar BDatos con BDatos_Ant

' ==================================================================================================================================
            Sub Call_RuT_Update_LSGES04_ACont()
                Debug.Print "================== >>> Call_RuT_Update_LSGES04_ACont"
                Dim ActivSheet  As String:  ActivSheet = ThisWorkbook.ActiveSheet.Name
                Prog__APP.Range("APP_Task_Rut") = "RuT_Update_LSGES04_ACont"
            On Error GoTo ManejoError
                DoEvents ' Permite que Excel procese eventos pendientes
                Form_Running_Rut.Show
            On Error GoTo 0
                ThisWorkbook.Sheets(ActivSheet).Select
                Debug.Print "================== <<< Call_RuT_Update_LSGES04_ACont"
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
Function Func_MsgBox_vbYesNo(Optional Msg As String = "¿ Seguimos ?") As Boolean
    Dim Respuesta    As VbMsgBoxResult
    Func_MsgBox_vbYesNo = True
    Respuesta = MsgBox(Msg, vbYesNo + vbQuestion, "procedimiento")
    If Respuesta = vbNo Then Func_MsgBox_vbYesNo = False
End Function
'==================================================================================================================================
Sub Rut_AskDate_CierreContable(RngFechaCierre As Range)
    Dim FechaCierre     As Variant
    Dim Mensaje         As String
    
    Mensaje = "La fecha de cierre contable vigente es: " & RngFechaCierre.Value & vbCrLf & _
              "Si deseas cambiarla, escribe la nueva fecha y pulsa Intro." & vbCrLf & _
              "Si la fecha es correcta, simplemente pulsa Intro."
    
    FechaCierre = InputBox(Mensaje, "Fecha de Cierre Contable", RngFechaCierre.Value)
    
    If FechaCierre <> "" And FechaCierre <> RngFechaCierre.Value Then
        If IsDate(FechaCierre) Then
            RngFechaCierre.Value = FechaCierre
            MsgBox "La nueva fecha de cierre contable es: " & FechaCierre
        Else
            MsgBox "La fecha introducida no es válida. Por favor, introduce una fecha correcta."
        End If
    Else
        MsgBox "La fecha de cierre contable no se ha modificado."
    End If
End Sub
'==================================================================================================================================

'==================================================================================================================================
'- Importar Última Consulta de LSGES04_GE, Actualizar registros existentes y Añadir Nuevos.
'==================================================================================================================================
Sub RuT_Update_LSGES04_ACont()
'==================================================================================================================================
Debug.Print "------------------------- >>> RuT_Update_LSGES04_ACont()"
    Dim TimeLap2        As Single
    Dim TxT_Progreso    As String
    Dim Arch_New_Name   As String
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
    Application.ScreenUpdating = False
    
    '- Setting ListObjects ------------------------------------
    Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
    
    Prog_DefCol_BD.Visible = xlSheetVisible
    Sht__BD.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
'    Lo_DefCol_BD.TotalsRowRange(DefC_HiddenCol) = False
    Prog__APP.Range("SW_Col_Hide_Sht__BD") = False
    Sht__BD.Unprotect
    Lo_BD.ShowTotals = False
    Lo_DefCol_BD.ShowTotals = False
    
    Dim Lo_BD_Ant           As ListObject:      Set Lo_BD_Ant = Sht__BD_Ant.ListObjects(1)
    Sht__BD_Ant.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_Ant)
    Prog__APP.Range("SW_Col_Hide_Sht__BD_Ant") = Not Prog__APP.Range("SW_Col_Hide_Sht__BD_Ant")
    Sht__BD_Ant.Unprotect
    Lo_BD_Ant.ShowTotals = False
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    Rut_Off_Functions
    Application.DisplayAlerts = False
    
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", _
                                        "Proceso: Importar LSGES04 para extraer Recibos Académicos de TIO y EP. Imp.Adm del Curso " & _
                                        Prog__APP.Range("APP_C_Acad_Pos") & ".", 0)

        '- Ajuste de la Fecha del Cierro Contable --------------
'''        Call Rut_AskDate_CierreContable(Range("APP_FechCierreCont"))
        
'GoTo SaltoAquí

    If Not Func_MsgBox_vbYesNo("¿ Importamos LSGES04 ?") Then GoTo SaltoAquí
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import LSGES04 por Año Contable --------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Arch_New_Name = "LSGES04_GE_SinDtos_Año_" & AñoCont
    Call Rut_Lo_Import_LoData_LoDefCol(Lo_BD, Lo_DefCol_BD, DefC_TitColGenInf, Arch_New_Name)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    Prog__APP.Range("APP_Last_Import") = Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD.Range("d2") = "Última Importación: " & Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD.Range("b3") = " Tabla Recibos Académicos de TIO y EP. y de Imp.Adm del Año_Contable " & Prog__APP.Range("APP_AñoCont") & _
                            ", y Cursos: " & Prog__APP.Range("APP_C_Acad_Ant") & " y " & Prog__APP.Range("APP_C_Acad_Pos") & "."
    
Rut_Lo_ListColumns_ClearContents_DefC_ProtectData:
    'If Not Func_MsgBox_vbYesNo("¿ Borrar Datos por protección de Datos ?") Then GoTo Rut_Lo_Format
    '    '- ----------------------------------------------------------------------------------------------------------------------------
    '    '- Rut_Lo: Borrar por protección de Datos, Información sensible y no necesarias -------------------------------------------------------
    '    '- ----------------------------------------------------------------------------------------------------------------------------
    '    If Prog__APP.Range("SW_ProtecciónDatosActivado") Then
    '        Call Rut_Lo_ListColumns_ClearContents_DefC_ProtectData(Lo_BD, Lo_DefCol_BD, DefC_ProtectData)    '- Borrar por protección de Datos, Información sensible y no necesarios
    '            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
    '            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Eliminada Información sensible.", LastTimeLap)
    '    Else
    '            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "¡ Información sensible SIN Eliminar !", LastTimeLap)
    '    End If
    
Rut_Lo_Format:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Format Sht__BD ---------------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Sht__BD.Select
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            TxT_Progreso = ActivForm.Controls("TBx_Informe")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Formateando el Excel.", LastTimeLap)
            TimeLap2 = LastTimeLap
    Call Rut_Lo_Format_LoData_LoDefColData(Lo_BD, Lo_DefCol_BD)
            LastTimeLap = TimeLap2
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Formateado el Excel.", LastTimeLap, , , , , , , 2)

RuT_Remove_Reg_No_Valid:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_111, Filtrar y Borrar Registros NO deseados: -----------------------------------------------------------------------------
        '-  Recibos AE4 Enseñanzas Propias
        '-  Recibos de Matrículas de coste CERO
        '-  Recibos con Imp.Rec. < 0
        '-  Recibos con fechas FUERA DEL PERÍODO CONTABLE
        '-  Recibos con Incongruencias de Fechas
        '-  Recibos NO Martrícula
        '-  Recibos ANULADOS
        '-  Recibos INVALIDADOS
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim Lo_BD_ErrDate       As ListObject:      Set Lo_BD_ErrDate = Sht__BD_ErrDate.ListObjects(1)
            Sht__BD_ErrDate.Visible = xlSheetVisible
            Call Rut_Lo_WrkSht_Preparar(Sht__BD_ErrDate)
            Prog__APP.Range("SW_Col_Hide_Sht__BD_ErrDate") = False
            Sht__BD_ErrDate.Unprotect
            Lo_BD_ErrDate.ShowTotals = False
            TimeLap2 = LastTimeLap
    Call RuT_Remove_Reg_No_Valid(Lo_BD, Lo_BD_ErrDate)
            Set Lo_BD_ErrDate = Nothing
            LastTimeLap = TimeLap2
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Recibos no requeridos para el procedimiento.", LastTimeLap, , , , , , , 2)
    
RuT_Duplicates_Search:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_112, Gestionar Duplicados -------------------------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim Lo_BD_Dpl           As ListObject:      Set Lo_BD_Dpl = Sht__BD_Dupl.ListObjects(1)
            Sht__BD_Dupl.Visible = xlSheetVisible
            Call Rut_Lo_WrkSht_Preparar(Sht__BD_Dupl)
            Prog__APP.Range("SW_Col_Hide_Sht__BD_Dupl") = False
            Sht__BD_Dupl.Unprotect
            Lo_BD_Dpl.ShowTotals = False
    Call RuT_Duplicates_Search(Lo_BD, Lo_DefCol_BD, Lo_BD_Dpl, BD_Ref, BD_Incidencias, BD_H_Incidencias)
            Set Lo_BD_Dpl = Nothing
        
'''Rut_Copy_AE4x4_en_BDatos:
'''    If Not Func_MsgBox_vbYesNo("¿ Trasladar los recibos de AE4x4 a BDatos ?") Then GoTo RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio
'''    '- ----------------------------------------------------------------------------------------------------------------------------
'''    '- M_315_Copy_INSS_a_BD, Trasladar el importe INSS a los recibos de BDatos ----------------------------------------------------
'''    '- ----------------------------------------------------------------------------------------------------------------------------
'''    Call Rut_Copy_AE4x4_en_BDatos   '¡¡¡ Quizás preguntar si no hay que actualizar antes con el módúlo Update_LSGES04_C_Acad_Ant !!!

RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_113, Asignar Código Concepto-Eco del Rec Y Rellenar Col Cta_Ingreso con el nº de Cta. correspondiente y Año de Vencimiento en ACont_Vto-----
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Determinar Fecha de Vencimiento
    '- Determinar Cta-CCC Ingreso
    '- Determinar Concepto Económico
    '- Determinar Tipo de Enseñanza TIO-EP
    Call RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio(Lo_BD)

RuT_Clasif_Recibos:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_114, Clasificar Recibos en Emitidos, Aplazados, EjeAnt, ADxAplz, Añejos ----------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call RuT_Clasif_Recibos

Rut_Assign_Imp_AdmAcad_C_Acad_Pos:
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_115, Identificar y Asignar al primer registro de la matrícula el Importe Académico y el Administrativo --------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call Rut_Assign_Imp_AdmAcad_C_Acad_Pos(Lo_BD)


Rut_Copy_ImpAdm_CAcadAnt_a_BDatos:
If Not Func_MsgBox_vbYesNo("¿ Trasladar el ImpAdm C_Acad_Ant a BDatos ?") Then GoTo Rut_Copy_ImpINSS_en_BDatos
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_215_Copy_IAdmCAcad_Ant_a_BD, Trasladar el ImpAdm, ImpAcad y ImpDto del C_Acad_Ant a BDatos -------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call Rut_Copy_ImpAdm_CAcadAnt_a_BDatos   '¡¡¡ Quizás preguntar si no hay que actualizar antes con el módúlo Update_LSGES04_C_Acad_Ant !!!
        
Rut_Copy_ImpINSS_en_BDatos:
If Not Func_MsgBox_vbYesNo("¿ Trasladar el Imp_INSS del C_Acad_Ant/Pos a BDatos ?") Then GoTo Rut_Lo_Import_WorkSheet
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_315_Copy_INSS_a_BD, Trasladar el importe INSS a los recibos de BDatos ----------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call Rut_Copy_ImpINSS_en_BDatos   '¡¡¡ Quizás preguntar si no hay que actualizar antes con el módúlo Update_LSGES04_C_Acad_Ant !!!
        

SaltoAquí:

Rut_Lo_Import_WorkSheet:
If Not Func_MsgBox_vbYesNo("¿ Importar Datos del Mes Anterior ?") Then GoTo Rut_Actualizar_LoBDatos_con_LoBD_Ant
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_118_Actualiz_BD_con_BD_Ant -------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim Ws_BD_Ant   As Worksheet:   Set Ws_BD_Ant = Sht__BD_Ant
    Arch_New_Name = "PPub_BDatos_Prog-RibbonX V-"
    Call Rut_Lo_Import_WorkSheet(Ws_BD_Ant, Arch_New_Name, Sht__BD.Name)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
        
Rut_Actualizar_LoBDatos_con_LoBD_Ant:
If Not Func_MsgBox_vbYesNo("¿ Actualizams BDatos con BDatos del Mes Anterior ?") Then GoTo Terminar
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_118_Actualiz_BD_con_BD_Ant -------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim Ws_BD_Ant   As Worksheet:   Set Ws_BD_Ant = Sht__BD_Ant
    Arch_New_Name = "PPub_BDatos_Prog-RibbonX V-"
    Call Rut_Lo_Import_WorkSheet(Ws_BD_Ant, Arch_New_Name, Sht__BD.Name)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    'Call Rut_Actualizar_LoBDatos_con_LoBD_Ant
        
        
Terminar:
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)
    
Final:
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe").Text
    Sht__BD.Select
    Lo_BD.ShowTotals = True
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD)
    
Restablecer_Valores:
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD.Visible = xlSheetVeryHidden
'Sht__BD.Visible = xlSheetVeryHidden
Rut_On_Functions
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set ActivForm = Nothing
Debug.Print "------------------------- <<< Sub RuT_Update_LSGES04_ACont()"
End Sub ' RuT_Importar_LSGES04_GE   --------------------------------------------------------------------------------------------
'===================================================================================================================================

