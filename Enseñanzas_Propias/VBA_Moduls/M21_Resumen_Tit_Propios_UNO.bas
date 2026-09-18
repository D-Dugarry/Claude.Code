Attribute VB_Name = "M21_Resumen_Tit_Propios_UNO"
' ==================================================================================================================
' *** MODULO COMPLETO DESACTIVADO (comentado) el 2026-09-19 00:27 ***
'
' Motivo: no compila. 'Cod_Plan' (el filtro por plan de la linea 106) no se
' declara en ningun sitio. Ver B19 del Informe_Bugs: el valor deberia salir
' del rango TP_Cod_Plan (celda F1 de la hoja Tit_Propio_UNO / Wk_TitP_UNO),
' pero la rutina VACIA ese rango en la linea 94, antes de usarlo, asi que no
' bastaba con declarar la variable: habia que decidir antes de donde sale el
' plan a resumir. Decision aplazada por el usuario.
'
' OJO - a diferencia de M50/M51, esta rutina SI estaba viva: es la tarea 31
' de la tabla Tb_Tareas del menu auxiliar, '2_ Exportar Tabla Resumen de UN
' Tit.Propio'. Mientras el modulo este comentado, esa entrada del menu fallara
' al pulsarla (Form_Menu la lanza con Application.Run y no encontrara la macro).
' Conviene quitar o marcar esa fila en Tb_Tareas mientras tanto.
' ==================================================================================================================
'Option Explicit

' ==================================================================================================================================
'Sub Rut_Resumen_Tab_TitPropios_UNO()    ' ===============================================================================================
' ==================================================================================================================================
'--- Tabla Wk_TitP_UNO  Tit.Propios-Resumen ----------------------
' Const CtTP_Cod_Plan          As Integer = 1
' Const CtTP_Curso_Acad        As Integer = 2
' Const CtTP_Año_Emi           As Integer = 3
' Const CtTP_Plan_Curss        As Integer = 4
' Const CtTP_NomPlan           As Integer = 5
' Const CtTP_Orgánica          As Integer = 6
' Const CtTP_Ref_JI            As Integer = 7
' Const CtTP_Concepto          As Integer = 8
' Const CtTP_Incidencia        As Integer = 9
' Const CtTP_Cant_Reg          As Integer = 10
' Const CtTP_Ret_VRI           As Integer = 11
' Const CtTP_Emi_Total         As Integer = 12
' Const CtTP_Emi_Org           As Integer = 13
' Const CtTP_Emi_VRI           As Integer = 14
' Const CtTP_Emi_Tadm          As Integer = 15
' Const CtTP_Emi_SinOrg        As Integer = 16
' Const CtTP_Cob_Total         As Integer = 17
' Const CtTP_Cob_Org           As Integer = 18
' Const CtTP_Cob_VRI           As Integer = 19
' Const CtTP_Cob_Tadm          As Integer = 20
' Const CtTP_Cob_SinOrg        As Integer = 21
 
' Const CtTP_RDT_Total         As Integer = 22
' Const CtTP_RDT_Org           As Integer = 23
' Const CtTP_RDT_VRI           As Integer = 24
' Const CtTP_RDT_Tadm          As Integer = 25
' Const CtTP_RDT_SinOrg        As Integer = 26
' Const CtTP_SLD_Total         As Integer = 27
' Const CtTP_SLD_Org           As Integer = 28
' Const CtTP_SLD_VRI           As Integer = 29
' Const CtTP_SLD_Tadm          As Integer = 30
' Const CtTP_SLD_SinOrg        As Integer = 31
' Const CtTP_AD_Total         As Integer = 32
' Const CtTP_AD_Org           As Integer = 33
' Const CtTP_AD_VRI           As Integer = 34
' Const CtTP_AD_Tadm          As Integer = 35
' Const CtTP_AD_SinOrg        As Integer = 36
 
' Const CtTP_Pdte_Cob          As Integer = 37
' Const CtTP_Descripción       As Integer = 38


'Dim Fila_DR             As Long
'Dim F_Plan              As Long
'Dim F_Ant               As Long

'Dim Plan_Ant        As String
'Dim Curso_Acad_Ant  As String
'Dim Año_Emi_Ant     As String

'Dim Cont_Tot_Reg        As Long
'Dim Cont_Reg_Proc       As Long

'Dim R_Emi_Acad          As Currency
'Dim R_Cob_Acad          As Currency
'Dim R_Rdt_Acad          As Currency

'Dim ProgresoTarea       As String

'Dim Lo_TPH          As ListObject:       Set Lo_TPH = Prog_BD.ListObjects(1)
'Dim Lo_TPResum      As ListObject:       Set Lo_TPResum = Wk_TitP_UNO.ListObjects(1)

'Rut_Off_Functions
    
'    Form_Menu.TB_Informe = "Generando la Tabla:      " & Format(Now, "hh:mm:ss")
'    ProgresoTarea = Form_Menu.TB_Informe & vbCrLf
'    Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False

'    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Ordenar por PLAN y DNI ==================
'    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
'    Call Rut_Lo_Sort(Lo_TPH, BD_Plan, xlAscending, True)
'    Call Rut_Lo_Sort(Lo_TPH, BD_C_Acad, xlAscending)
'    Call Rut_Lo_Sort(Lo_TPH, BD_FEmi, xlAscending)
    '- Preparar Tabla de Wk_TitP_UNO ==================
'    Wk_TitP_UNO.Visible = xlSheetVisible
'    Wk_TitP_UNO.Select
'    Call Rut_Lo_WrkSht_Preparar(Wk_TitP_UNO)
'    Wk_TitP_UNO.Unprotect
    '- Vacío la Tabla de Tit.Prop.  =====================================
'    Lo_TPResum.AutoFilter.ShowAllData
'    If Not Lo_TPResum.DataBodyRange Is Nothing Then Lo_TPResum.DataBodyRange.Delete
'GoTo Restablecer_Valores
    ' ==================================================================================================================================
    ' ###############################  Genero la Tabla de Planes de DR  #####################################
'    Plan_Ant = "":   Curso_Acad_Ant = "":   Año_Emi_Ant = 0: Cont_Tot_Reg = 0:
'    Range("TP_Cod_Plan") = ""
'    Range("TP_Cod_Plan").Offset(0, 1) = Range("APP_CursAcad")
    ' ##################################################################################################################
    ' =============  Recorrer todos los Registros filtrados y Crear la Tabla de Planes de DR  =====================================
'Dim RwPH        As ListRow
'Dim RwTP        As ListRow

'    For Fila_DR = 1 To 5000 'Lo_TPH.ListRows.Count
    
'        Set RwPH = Nothing
'        Set RwPH = Lo_TPH.ListRows(Fila_DR)
        
'        If RwPH.Range(BD_Plan) <> Cod_Plan Then GoTo Siguiente_Fila
'        If Len(RwPH.Range(BD_FEmi)) = 0 Then GoTo Siguiente_Fila
'        If RwPH.Range(BD_ImpRec) <= 0 Then GoTo Siguiente_Fila
'        If RwPH.Range(BD_Anul) = "S" Then GoTo Siguiente_Fila
'        If RwPH.Range(BD_Matricula) = "N" Then GoTo Siguiente_Fila
        
'        Cont_Reg_Proc = Cont_Reg_Proc + 1
'        Cont_Tot_Reg = Cont_Tot_Reg + 1
'''GoTo Siguiente_Fila
        ' --------------=============  Tratamiento de los Datos  ==================
'        If Plan_Ant & Curso_Acad_Ant & Año_Emi_Ant = RwPH.Range(BD_Plan) & RwPH.Range(BD_C_Acad) & Format(RwPH.Range(BD_FEmi), "yyyy") Then    ' ------- Control cambio de Plan de estudio y de Curso Académico ---------------
            ' Añadir JI si no está -------------------
'            If RwPH.Range(BD_JI_Emi_Acad) <> "" Then
'                If InStr(RwTP.Range(CtTP_Ref_JI), RwPH.Range(BD_JI_Emi_Acad)) = 0 Then
'                    RwTP.Range(CtTP_Ref_JI) = RwTP.Range(CtTP_Ref_JI) & " - " & RwPH.Range(BD_JI_Emi_Acad)
'                End If
'            End If
            ' Añadir Orgánica si no está ------------
'            If RwPH.Range(BD_Orgánica) <> "" Then
'                If InStr(RwTP.Range(CtTP_Orgánica), RwPH.Range(BD_Orgánica)) = 0 Then
'                    RwTP.Range(CtTP_Orgánica) = RwTP.Range(CtTP_Orgánica) & " - " & RwPH.Range(BD_Orgánica)
'                End If
'            End If
            '--- Importe Emitido --------
'            If RwPH.Range(BD_ImpRec) > 0 Then RwTP.Range(CtTP_Emi_Total) = RwTP.Range(CtTP_Emi_Total) + RwPH.Range(BD_ImpRec)
'            RwTP.Range(CtTP_Emi_Tadm) = RwTP.Range(CtTP_Emi_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
            '--- Importe Cobrado --------
'            If RwPH.Range(BD_ImpCob) > 0 Then
'                RwTP.Range(CtTP_Cob_Total) = RwTP.Range(CtTP_Cob_Total) + RwPH.Range(BD_ImpCob)
'                RwTP.Range(CtTP_Cob_Tadm) = RwTP.Range(CtTP_Cob_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
                '--- Importes Redistribuidos --------=====================================================================================
'                If RwPH.Range(BD_Liquidado) <> "" Then
'                    RwTP.Range(CtTP_RDT_Total) = RwTP.Range(CtTP_RDT_Total) + RwPH.Range(BD_ImpCob)
'                    RwTP.Range(CtTP_RDT_Tadm) = RwTP.Range(CtTP_RDT_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
'                End If
'            End If
            
'            RwTP.Range(CtTP_Cant_Reg) = RwTP.Range(CtTP_Cant_Reg) + 1

'        Else    '------ Es el primero de una serie y tengo que introducir los datos comunes ---------------------------------------------------------
        
'            Set RwTP = Nothing
'            Set RwTP = Lo_TPResum.ListRows.Add
'            F_Plan = Lo_TPResum.ListRows.Count
           
'            Plan_Ant = RwPH.Range(BD_Plan)
'            Curso_Acad_Ant = RwPH.Range(BD_C_Acad)
'            Año_Emi_Ant = Format(RwPH.Range(BD_FEmi), "yyyy")
            
'            RwTP.Range(CtTP_Cod_Plan) = RwPH.Range(BD_Plan)
'            RwTP.Range(CtTP_Curso_Acad) = RwPH.Range(BD_C_Acad)
'            RwTP.Range(CtTP_Año_Emi) = Format(RwPH.Range(BD_FEmi), "yyyy")
'            RwTP.Range(CtTP_Plan_Curss) = RwPH.Range(BD_Plan) & "_" & RwPH.Range(BD_C_Acad)
'            RwTP.Range(CtTP_NomPlan) = RwPH.Range(BD_NomPlan)
            
'            RwTP.Range(CtTP_Ref_JI) = RwPH.Range(BD_JI_Emi_Acad)
'            RwTP.Range(CtTP_Orgánica) = RwPH.Range(BD_Orgánica)
            
'            If IsNumeric(Left(RwPH.Range(BD_Plan), 1)) Then RwTP.Range(CtTP_Concepto) = 1311 Else RwTP.Range(CtTP_Concepto) = 1311.03

'            RwTP.Range(CtTP_Cant_Reg) = 1
'            RwTP.Range(CtTP_Ret_VRI) = RwPH.Range(BD_Coef_VRI)
            
'            If RwPH.Range(BD_ImpRec) > 0 Then RwTP.Range(CtTP_Emi_Total) = RwPH.Range(BD_ImpRec)
'            RwTP.Range(CtTP_Emi_Tadm) = RwPH.Range(BD_Rec_Imp_Adm)
            
'            If RwPH.Range(BD_ImpCob) > 0 Then
'                RwTP.Range(CtTP_Cob_Total) = RwPH.Range(BD_ImpCob)
'                RwTP.Range(CtTP_Cob_Tadm) = RwPH.Range(BD_Rec_Imp_Adm)
                '--- Importes Redistribuidos --------=====================================================================================
'                If RwPH.Range(BD_Liquidado) <> "" Then
'                    RwTP.Range(CtTP_RDT_Total) = RwPH.Range(BD_ImpCob)
'                    RwTP.Range(CtTP_RDT_Tadm) = RwPH.Range(BD_Rec_Imp_Adm)
'                End If
'            End If
            
            ' =====================================================================================
            ' ---------------=============  Cálculos Redistribución de la Fila Anterior  ==================
'            If F_Plan > 1 Then
'                    F_Ant = F_Plan - 1
                
'                With Lo_TPResum.DataBodyRange
                    
                    '----- Reparto de lo Emitido -------------------------------------------
'                    R_Emi_Acad = .Cells(F_Ant, CtTP_Emi_Total) - .Cells(F_Ant, CtTP_Emi_Tadm)
'                    .Cells(F_Ant, CtTP_Emi_VRI) = Application.WorksheetFunction.Round(R_Emi_Acad * .Cells(F_Ant, CtTP_Ret_VRI) / 100, 2)
'                    .Cells(F_Ant, CtTP_Emi_Org) = R_Emi_Acad - .Cells(F_Ant, CtTP_Emi_VRI)
'                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Emi_SinOrg) = .Cells(F_Ant, CtTP_Emi_Org) + .Cells(F_Ant, CtTP_Emi_Tadm)
                    '----- Reparto de lo Cobrado -------------------------------------------
'                    R_Cob_Acad = .Cells(F_Ant, CtTP_Cob_Total) - .Cells(F_Ant, CtTP_Cob_Tadm)
'                    .Cells(F_Ant, CtTP_Cob_VRI) = Application.WorksheetFunction.Round(R_Cob_Acad * .Cells(F_Ant, CtTP_Ret_VRI) / 100, 2)
'                    .Cells(F_Ant, CtTP_Cob_Org) = R_Cob_Acad - .Cells(F_Ant, CtTP_Cob_VRI)
'                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Cob_SinOrg) = .Cells(F_Ant, CtTP_Cob_Org) + .Cells(F_Ant, CtTP_Cob_Tadm)
                    '--- Reparto Importes Redistribuidos -----------------------------------
'                    R_Rdt_Acad = .Cells(F_Ant, CtTP_RDT_Total) - .Cells(F_Ant, CtTP_RDT_Tadm)
'                    .Cells(F_Ant, CtTP_RDT_VRI) = Application.WorksheetFunction.Round(R_Rdt_Acad * .Cells(F_Ant, CtTP_Ret_VRI) / 100, 2)
'                    .Cells(F_Ant, CtTP_RDT_Org) = R_Rdt_Acad - .Cells(F_Ant, CtTP_RDT_VRI)
'                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_RDT_SinOrg) = .Cells(F_Ant, CtTP_RDT_Org) + .Cells(F_Ant, CtTP_RDT_Tadm)
                    '----- Reparto Saldos --------------------------------------------------
'                    .Cells(F_Ant, CtTP_SLD_Total) = .Cells(F_Ant, CtTP_Cob_Total) - .Cells(F_Ant, CtTP_RDT_Total)
'                    .Cells(F_Ant, CtTP_SLD_Org) = .Cells(F_Ant, CtTP_Cob_Org) - .Cells(F_Ant, CtTP_RDT_Org)
'                    .Cells(F_Ant, CtTP_SLD_VRI) = .Cells(F_Ant, CtTP_Cob_VRI) - .Cells(F_Ant, CtTP_RDT_VRI)
'                    .Cells(F_Ant, CtTP_SLD_Tadm) = .Cells(F_Ant, CtTP_Cob_Tadm) - .Cells(F_Ant, CtTP_RDT_Tadm)
'                    .Cells(F_Ant, CtTP_SLD_SinOrg) = .Cells(F_Ant, CtTP_Cob_SinOrg) - .Cells(F_Ant, CtTP_RDT_SinOrg)
                    '----- Pendiente de Cobro AD -------------------------------------------
'                    .Cells(F_Ant, CtTP_AD_Total) = .Cells(F_Ant, CtTP_Emi_Total) - .Cells(F_Ant, CtTP_Cob_Total)
'                    .Cells(F_Ant, CtTP_AD_Org) = .Cells(F_Ant, CtTP_Emi_Org) - .Cells(F_Ant, CtTP_Cob_Org)
'                    .Cells(F_Ant, CtTP_AD_VRI) = .Cells(F_Ant, CtTP_Emi_VRI) - .Cells(F_Ant, CtTP_Cob_VRI)
'                    .Cells(F_Ant, CtTP_AD_Tadm) = .Cells(F_Ant, CtTP_Emi_Tadm) - .Cells(F_Ant, CtTP_Cob_Tadm)
'                    .Cells(F_Ant, CtTP_AD_SinOrg) = .Cells(F_Ant, CtTP_Emi_SinOrg) - .Cells(F_Ant, CtTP_Cob_SinOrg)
'                End With
'            End If
            ' =====================================================================================
'        End If
'Siguiente_Fila:

'        Set RwPH = Nothing

'        If Fila_DR Mod 500 = 0 Then
'            Debug.Print "Analizando Prog_BD:  " & Format(Fila_DR, "#,##0") & " de " & Format(Lo_TPH.ListRows.Count, "#,##0")
'            Form_Menu.TB_Informe = ProgresoTarea & vbCrLf & "Tiempo transcurrido:  " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & "Analizando Prog_BD:  " & Format(Fila_DR, "#,##0") & " de " & _
'                                             Format(Lo_TPH.ListRows.Count, "#,##0") & " reg.       Procesados:  " & Cont_Reg_Proc & " reg."
'            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
'        End If

'    Next Fila_DR
             
            ' ---------------------------=============  Cálculos Redistribución  de la última Fila ==================
'            If F_Plan > 1 Then
'                    F_Ant = F_Plan

'                With Lo_TPResum.DataBodyRange
                    
                    '----- Reparto de lo Emitido -------------------------------------------
'                    R_Emi_Acad = .Cells(F_Ant, CtTP_Emi_Total) - .Cells(F_Ant, CtTP_Emi_Tadm)
'                    .Cells(F_Ant, CtTP_Emi_VRI) = Application.WorksheetFunction.Round(R_Emi_Acad * .Cells(F_Ant, CtTP_Ret_VRI) / 100, 2)
'                    .Cells(F_Ant, CtTP_Emi_Org) = R_Emi_Acad - .Cells(F_Ant, CtTP_Emi_VRI)
'                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Emi_SinOrg) = .Cells(F_Ant, CtTP_Emi_Org) + .Cells(F_Ant, CtTP_Emi_Tadm)
                    '----- Reparto de lo Cobrado -------------------------------------------
'                    R_Cob_Acad = .Cells(F_Ant, CtTP_Cob_Total) - .Cells(F_Ant, CtTP_Cob_Tadm)
'                    .Cells(F_Ant, CtTP_Cob_VRI) = Application.WorksheetFunction.Round(R_Cob_Acad * .Cells(F_Ant, CtTP_Ret_VRI) / 100, 2)
'                    .Cells(F_Ant, CtTP_Cob_Org) = R_Cob_Acad - .Cells(F_Ant, CtTP_Cob_VRI)
'                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Cob_SinOrg) = .Cells(F_Ant, CtTP_Cob_Org) + .Cells(F_Ant, CtTP_Cob_Tadm)
                    '--- Reparto Importes Redistribuidos -----------------------------------
'                    R_Rdt_Acad = .Cells(F_Ant, CtTP_RDT_Total) - .Cells(F_Ant, CtTP_RDT_Tadm)
'                    .Cells(F_Ant, CtTP_RDT_VRI) = Application.WorksheetFunction.Round(R_Rdt_Acad * .Cells(F_Ant, CtTP_Ret_VRI) / 100, 2)
'                    .Cells(F_Ant, CtTP_RDT_Org) = R_Rdt_Acad - .Cells(F_Ant, CtTP_RDT_VRI)
'                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_RDT_SinOrg) = .Cells(F_Ant, CtTP_RDT_Org) + .Cells(F_Ant, CtTP_RDT_Tadm)
                    '----- Reparto Saldos ----------------------------------------------
'                    .Cells(F_Ant, CtTP_SLD_Total) = .Cells(F_Ant, CtTP_Cob_Total) - .Cells(F_Ant, CtTP_RDT_Total)
'                    .Cells(F_Ant, CtTP_SLD_Org) = .Cells(F_Ant, CtTP_Cob_Org) - .Cells(F_Ant, CtTP_RDT_Org)
'                    .Cells(F_Ant, CtTP_SLD_VRI) = .Cells(F_Ant, CtTP_Cob_VRI) - .Cells(F_Ant, CtTP_RDT_VRI)
'                    .Cells(F_Ant, CtTP_SLD_Tadm) = .Cells(F_Ant, CtTP_Cob_Tadm) - .Cells(F_Ant, CtTP_RDT_Tadm)
'                    .Cells(F_Ant, CtTP_SLD_SinOrg) = .Cells(F_Ant, CtTP_Cob_SinOrg) - .Cells(F_Ant, CtTP_RDT_SinOrg)
                    '----- Pendiente de Cobro AD -------------------------------------------
'                    .Cells(F_Ant, CtTP_AD_Total) = .Cells(F_Ant, CtTP_Emi_Total) - .Cells(F_Ant, CtTP_Cob_Total)
'                    .Cells(F_Ant, CtTP_AD_Org) = .Cells(F_Ant, CtTP_Emi_Org) - .Cells(F_Ant, CtTP_Cob_Org)
'                    .Cells(F_Ant, CtTP_AD_VRI) = .Cells(F_Ant, CtTP_Emi_VRI) - .Cells(F_Ant, CtTP_Cob_VRI)
'                    .Cells(F_Ant, CtTP_AD_Tadm) = .Cells(F_Ant, CtTP_Emi_Tadm) - .Cells(F_Ant, CtTP_Cob_Tadm)
'                    .Cells(F_Ant, CtTP_AD_SinOrg) = .Cells(F_Ant, CtTP_Emi_SinOrg) - .Cells(F_Ant, CtTP_Cob_SinOrg)

'                End With
'            End If
   
'Restablecer_Valores:
'Rut_EnableEvents_Status_Reset
'MsgBox "Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & "    Registros: " & Cont_Tot_Reg & "    Planes: " & F_Plan

'Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
'        "En la Nueva Consulta hay:  " & "  -  Tot.Reg. " & Cont_Tot_Reg & "    Planes: " & F_Plan & vbCrLf & Now()
        
'Rut_On_Functions
'End Sub     ' Rut_Resumen_Tab_TitPropios_UNO  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
            





