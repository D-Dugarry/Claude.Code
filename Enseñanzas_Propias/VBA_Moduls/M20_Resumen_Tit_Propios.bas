Attribute VB_Name = "M20_Resumen_Tit_Propios"
' Last Rev. 2026-09-19 08:28
Option Explicit

' ==================================================================================================================================
Sub Rut_Resumen_Tab_TitPropios()    ' ===============================================================================================
' ==================================================================================================================================
'--- Tabla Sht__Inf_EPs_UXXI  Tit.Propios-Resumen ----------------------
 Dim AñoCont                As String:          AñoCont = Prog__APP.Range("APP_AñoCont")
 Dim CursAcad               As String:          CursAcad = Prog__APP.Range("APP_CursAcad")
 Dim AñoContAnt            As Integer:         AñoContAnt = Left(CursAcad, 4)      '- El 1º Año del Curso Académico
 Const CtTP_Cod_Plan          As Integer = 1
 Const CtTP_Curso_Acad        As Integer = 2
 Const CtTP_Año_Emi           As Integer = 3
 Const CtTP_Plan_Curso        As Integer = 4
 Const CtTP_NomPlan           As Integer = 5
 Const CtTP_Orgánica          As Integer = 6
 Const CtTP_Ref_JI            As Integer = 7
 Const CtTP_Concepto          As Integer = 8
 Const CtTP_Incidencia        As Integer = 9
 Const CtTP_Cant_Reg          As Integer = 10
 Const CtTP_Coef_VRI           As Integer = 11
 
 Const CtTP_Emi_Total         As Integer = 12
 Const CtTP_Emi_Org           As Integer = 13
 Const CtTP_Emi_VRI           As Integer = 14
 Const CtTP_Emi_Tadm          As Integer = 15
 Const CtTP_Emi_SinOrg        As Integer = 16
 
 Const CtTP_Cob_Acont         As Integer = 17
 Const CtTP_Cob_C_ACad        As Integer = 18
 Const CtTP_Cob_Org           As Integer = 19
 Const CtTP_Cob_VRI           As Integer = 20
 Const CtTP_Cob_Tadm          As Integer = 21
 Const CtTP_Cob_SinOrg        As Integer = 22

 Const CtTP_RDT_Total         As Integer = 23
 Const CtTP_RDT_Org           As Integer = 24
 Const CtTP_RDT_VRI           As Integer = 25
 Const CtTP_RDT_Tadm          As Integer = 26
 
 Const CtTP_PdteRDT_Total         As Integer = 27
 Const CtTP_PdteRDT_Org           As Integer = 28
 Const CtTP_PdteRDT_VRI           As Integer = 29
 Const CtTP_PdteRDT_Tadm          As Integer = 30
 Const CtTP_PdteRDT_SinOrg        As Integer = 31
 
 Const CtTP_ADx_Total         As Integer = 32
 Const CtTP_ADx_Org           As Integer = 33
 Const CtTP_ADx_VRI           As Integer = 34
 Const CtTP_ADx_Tadm          As Integer = 35
 Const CtTP_ADx_SinOrg        As Integer = 36
 
 Const CtTP_Aplz_Total         As Integer = 37
 Const CtTP_Aplz_Org           As Integer = 38
 Const CtTP_Aplz_VRI           As Integer = 39
 Const CtTP_Aplz_Tadm          As Integer = 40
 Const CtTP_Aplz_SinOrg        As Integer = 41
 
 Const CtTP_Pdt_Total         As Integer = 42
 Const CtTP_Pdt_Org           As Integer = 43
 Const CtTP_Pdt_VRI           As Integer = 44
 Const CtTP_Pdt_Tadm          As Integer = 45
 Const CtTP_Pdt_SinOrg        As Integer = 46
 
 Const CtTP_Descripción       As Integer = 47


Dim Fila_DR             As Long
Dim F_Plan              As Long
Dim F_Ant               As Long

Dim Plan_Ant        As String
Dim Curso_Acad_Ant  As String
Dim Año_Emi_Ant     As String

Dim Cont_Tot_Reg        As Long
Dim Cont_Reg_Proc       As Long

Dim R_Emi_Acad          As Currency
Dim R_Cob_Acad          As Currency
Dim R_Rdt_Acad          As Currency
Dim R_ADx_Acad          As Currency
Dim R_Pdt_Acad          As Currency

Dim ProgresoTarea       As String

Dim Lo_BD          As ListObject:       Set Lo_BD = Prog_BD.ListObjects(1)
Dim Lo_TPResum      As ListObject:       Set Lo_TPResum = Sht__Inf_EPs_UXXI.ListObjects(1)

Rut_Off_Functions
Application.EnableEvents = False
    
    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TB_Informe = "Generando la Tabla:      " & Format(Now, "hh:mm:ss")
    ProgresoTarea = Form_Menu.TB_Informe & vbCrLf
    Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False

    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Ordenar por PLAN y DNI ==================
    Prog_BD.Unprotect
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending)
    Call Rut_Lo_Sort(Lo_BD, BD_FEmi, xlAscending)
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending)
    '- Preparar Tabla de Sht__Inf_EPs_UXXI ==================
    Sht__Inf_EPs_UXXI.Visible = xlSheetVisible
    Sht__Inf_EPs_UXXI.Select
    Call Rut_Lo_WrkSht_Preparar(Sht__Inf_EPs_UXXI)
    Sht__Inf_EPs_UXXI.Unprotect
    '- Vacío la Tabla de Tit.Prop.  =====================================
    Lo_TPResum.AutoFilter.ShowAllData
    If Not Lo_TPResum.DataBodyRange Is Nothing Then Lo_TPResum.DataBodyRange.Delete
    ' ==================================================================================================================================
    ' ###############################  Genero la Tabla de Planes de BDatos  #####################################
    Plan_Ant = "":   Curso_Acad_Ant = "":   Año_Emi_Ant = 0: Cont_Tot_Reg = 0:
    Range("TP_Cod_Plan") = ""
    Range("TP_Cod_Plan").Select
    ActiveCell.Offset(0, 1) = Range("APP_CursAcad")
    ActiveCell.Offset(0, -1) = "Enseñanzas de:  " & Range("APP_EFP_o_CFC")
    ActiveCell.Offset(0, 2) = " Última actualización: " & Now()
    ActiveCell.Offset(0, 1).Select
    ' ##################################################################################################################
    ' =============  Recorrer todos los Registros filtrados y Crear la Tabla de Planes de DR  =====================================
Dim RwPH        As ListRow
Dim RwTP        As ListRow

    For Fila_DR = 1 To Lo_BD.ListRows.Count
    
        Set RwPH = Nothing
        Set RwPH = Lo_BD.ListRows(Fila_DR)
        
        If Len(RwPH.Range(BD_FEmi)) = 0 Then GoTo Siguiente_Fila
        If RwPH.Range(BD_Matricula) = "N" Then GoTo Siguiente_Fila
        If RwPH.Range(BD_ImpRec) < 0 Then GoTo Siguiente_Fila
        If RwPH.Range(BD_ImpAdm) < 0 Then GoTo Siguiente_Fila
        'If RwPH.Range(BD_Anul) = "S" Then GoTo Siguiente_Fila
        
        Cont_Reg_Proc = Cont_Reg_Proc + 1
        Cont_Tot_Reg = Cont_Tot_Reg + 1
        ' --------------=============  Tratamiento de los Datos  ==================
        If Plan_Ant & Curso_Acad_Ant & Año_Emi_Ant = RwPH.Range(BD_Plan) & RwPH.Range(BD_C_Acad) & Format(RwPH.Range(BD_FEmi), "yyyy") Then    ' ------- Control cambio de Plan de estudio y de Curso Académico ---------------
            ' Añadir JI si no está -------------------
            If RwPH.Range(BD_JI_Emi_Acad) <> "" Then
                If InStr(RwTP.Range(CtTP_Ref_JI), RwPH.Range(BD_JI_Emi_Acad)) = 0 Then
                    RwTP.Range(CtTP_Ref_JI) = RwTP.Range(CtTP_Ref_JI) & " - " & RwPH.Range(BD_JI_Emi_Acad)
                End If
            End If
            ' Añadir Orgánica si no está ------------
            If RwPH.Range(BD_Orgánica) <> "" Then
                If InStr(RwTP.Range(CtTP_Orgánica), RwPH.Range(BD_Orgánica)) = 0 Then
                    RwTP.Range(CtTP_Orgánica) = RwTP.Range(CtTP_Orgánica) & " - " & RwPH.Range(BD_Orgánica)
                End If
            End If
            
            '--- Importe Emitido --------
            If RwPH.Range(BD_ImpRec) > 0 Then RwTP.Range(CtTP_Emi_Total) = RwTP.Range(CtTP_Emi_Total) + RwPH.Range(BD_ImpRec)
            RwTP.Range(CtTP_Emi_Tadm) = RwTP.Range(CtTP_Emi_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
            
            '--- Importe Cobrado Acont --------
            If RwPH.Range(BD_ImpCob) > 0 And RwPH.Range(BD_ACont_Cob) = AñoCont Then
                RwTP.Range(CtTP_Cob_Acont) = RwTP.Range(CtTP_Cob_Acont) + RwPH.Range(BD_ImpCob)
                RwTP.Range(CtTP_Cob_Tadm) = RwTP.Range(CtTP_Cob_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Importe Cobrado Curso Acad --------
            If RwPH.Range(BD_ImpCob) > 0 Then
                RwTP.Range(CtTP_Cob_C_ACad) = RwTP.Range(CtTP_Cob_C_ACad) + RwPH.Range(BD_ImpCob)
                RwTP.Range(CtTP_Cob_Tadm) = RwTP.Range(CtTP_Cob_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
                '--- Imp RDT --------
                If RwPH.Range(BD_RDT) <> "" Then
                    RwTP.Range(CtTP_RDT_Total) = RwTP.Range(CtTP_RDT_Total) + RwPH.Range(BD_ImpCob)
                    RwTP.Range(CtTP_RDT_Tadm) = RwTP.Range(CtTP_RDT_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
                End If
            End If
            
            '--- Importe ADxAplz --------
            If RwPH.Range(BD_ImpCob) = 0 And RwPH.Range(BD_ACont_Vto) > AñoCont Or _
               RwPH.Range(BD_ImpCob) > 0 And RwPH.Range(BD_ACont_Cob) = AñoCont + 1 And RwPH.Range(BD_ACont_Vto) = AñoCont + 1 Then
                    RwTP.Range(CtTP_ADx_Total) = RwTP.Range(CtTP_ADx_Total) + RwPH.Range(BD_ImpRec)
                    RwTP.Range(CtTP_ADx_Tadm) = RwTP.Range(CtTP_ADx_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Importe Aplz Aplazado --------
            If RwPH.Range(BD_ImpCob) > 0 And RwPH.Range(BD_ACont_Vto) > AñoContAnt Or _
               RwPH.Range(BD_ImpCob) > 0 And RwPH.Range(BD_ACont_Cob) = AñoCont + 1 And RwPH.Range(BD_ACont_Vto) = AñoCont + 1 Then
                    RwTP.Range(CtTP_Aplz_Total) = RwTP.Range(CtTP_Aplz_Total) + RwPH.Range(BD_ImpRec)
                    RwTP.Range(CtTP_Aplz_Tadm) = RwTP.Range(CtTP_Aplz_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Importe Pdte_Cobro --------
            If RwPH.Range(BD_ACont_Cob) <> AñoCont And RwPH.Range(BD_ACont_Emi) = AñoCont And RwPH.Range(BD_ACont_Vto) = AñoCont Then
                    RwTP.Range(CtTP_Pdt_Total) = RwTP.Range(CtTP_Pdt_Total) + RwPH.Range(BD_ImpRec)
                    RwTP.Range(CtTP_Pdt_Tadm) = RwTP.Range(CtTP_Pdt_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
            End If
            
            RwTP.Range(CtTP_Cant_Reg) = RwTP.Range(CtTP_Cant_Reg) + 1

        Else    '------ Es el primero de una serie y tengo que introducir los datos comunes ---------------------------------------------------------

            Set RwTP = Nothing
            Set RwTP = Lo_TPResum.ListRows.Add
            F_Plan = Lo_TPResum.ListRows.Count
           
            Plan_Ant = RwPH.Range(BD_Plan)
            Curso_Acad_Ant = RwPH.Range(BD_C_Acad)
            Año_Emi_Ant = Format(RwPH.Range(BD_FEmi), "yyyy")
            
            RwTP.Range(CtTP_Cod_Plan) = RwPH.Range(BD_Plan)
            RwTP.Range(CtTP_Curso_Acad) = RwPH.Range(BD_C_Acad)
            RwTP.Range(CtTP_Año_Emi) = Format(RwPH.Range(BD_FEmi), "yyyy")
            RwTP.Range(CtTP_Plan_Curso) = RwPH.Range(BD_Plan) & "_" & RwPH.Range(BD_C_Acad) & "_Cont" & Right(RwPH.Range(BD_FEmi), 2)
            
            RwTP.Range(CtTP_NomPlan) = RwPH.Range(BD_NomPlan)
            
            RwTP.Range(CtTP_Ref_JI) = RwPH.Range(BD_JI_Emi_Acad)
            RwTP.Range(CtTP_Orgánica) = RwPH.Range(BD_Orgánica)
            
            If IsNumeric(Left(RwPH.Range(BD_Plan), 1)) Then RwTP.Range(CtTP_Concepto) = 1311 Else RwTP.Range(CtTP_Concepto) = 1311.03

            RwTP.Range(CtTP_Cant_Reg) = 1
            RwTP.Range(CtTP_Coef_VRI) = RwPH.Range(BD_Coef_VRI)
            
            '--- Importe Emitido --------
            If RwPH.Range(BD_ImpRec) > 0 Then RwTP.Range(CtTP_Emi_Total) = RwPH.Range(BD_ImpRec)
            RwTP.Range(CtTP_Emi_Tadm) = RwPH.Range(BD_Rec_Imp_Adm)
            
            '--- Importe Cobrado Acont --------
            If RwPH.Range(BD_ImpCob) > 0 And RwPH.Range(BD_ACont_Cob) = AñoCont Then
                RwTP.Range(CtTP_Cob_Acont) = RwTP.Range(CtTP_Cob_Acont) + RwPH.Range(BD_ImpCob)
                RwTP.Range(CtTP_Cob_Tadm) = RwTP.Range(CtTP_Cob_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Importe Cobrado Curso Acad --------
            If RwPH.Range(BD_ImpCob) > 0 Then
                RwTP.Range(CtTP_Cob_C_ACad) = RwTP.Range(CtTP_Cob_C_ACad) + RwPH.Range(BD_ImpCob)
                RwTP.Range(CtTP_Cob_Tadm) = RwTP.Range(CtTP_Cob_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
                '--- Imp RDT --------
                If RwPH.Range(BD_RDT) <> "" Then
                    RwTP.Range(CtTP_RDT_Total) = RwTP.Range(CtTP_RDT_Total) + RwPH.Range(BD_ImpCob)
                    RwTP.Range(CtTP_RDT_Tadm) = RwTP.Range(CtTP_RDT_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
                End If
            End If
            
            '--- Importe ADxAplz --------
            If RwPH.Range(BD_ImpCob) = 0 And RwPH.Range(BD_ACont_Vto) > AñoCont Or _
               RwPH.Range(BD_ImpCob) > 0 And RwPH.Range(BD_ACont_Cob) = AñoCont + 1 And RwPH.Range(BD_ACont_Vto) = AñoCont + 1 Then
                    RwTP.Range(CtTP_ADx_Total) = RwPH.Range(BD_ImpRec)
                    RwTP.Range(CtTP_ADx_Tadm) = RwPH.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Importe Pdte_Cobro --------
            If RwPH.Range(BD_ImpCob) = 0 And RwPH.Range(BD_ACont_Emi) = AñoCont And RwPH.Range(BD_ACont_Vto) = AñoCont Then
                    RwTP.Range(CtTP_Pdt_Total) = RwTP.Range(CtTP_Pdt_Total) + RwPH.Range(BD_ImpRec)
                    RwTP.Range(CtTP_Pdt_Tadm) = RwTP.Range(CtTP_Pdt_Tadm) + RwPH.Range(BD_Rec_Imp_Adm)
            End If
            
            ' -------=============  Genero Texto de Descripción del JI  ==================---------
            '- Generar el campo Descripción. -----------
            If Range("APP_PlanMicroCred") <> "" Then
                RwTP.Range(CtTP_Descripción) = "LIQ-TitProp_" & RwPH.Range(BD_Plan) & _
                                               "-N Curso_" & CursAcad & _
                                               " AñoCont_" & AñoCont & _
                                               " - MicCred_" & RwPH.Range("APP_PlanMicroCred") & _
                                               " - " & RwPH.Range(BD_NomPlan)
            Else
                RwTP.Range(CtTP_Descripción) = "LIQ-TitProp_" & RwPH.Range(BD_Plan) & _
                                               "-N Curso_" & CursAcad & _
                                               " AñoCont_" & AñoCont & _
                                               " - " & RwPH.Range(BD_NomPlan)
            End If
            
            ' =====================================================================================
            ' ---------------=============  Cálculos Redistribución de la Fila Anterior  ==================
            If F_Plan > 1 Then
                    F_Ant = F_Plan - 1
                
                With Lo_TPResum.DataBodyRange
                    
                    '----- Reparto de lo Emitido -------------------------------------------
                    R_Emi_Acad = .Cells(F_Ant, CtTP_Emi_Total) - .Cells(F_Ant, CtTP_Emi_Tadm)
                    .Cells(F_Ant, CtTP_Emi_VRI) = Application.Round(R_Emi_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_Emi_Org) = R_Emi_Acad - .Cells(F_Ant, CtTP_Emi_VRI)
                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Emi_SinOrg) = .Cells(F_Ant, CtTP_Emi_Org) + .Cells(F_Ant, CtTP_Emi_Tadm)
                    '----- Reparto de lo Cobrado -------------------------------------------
                    R_Cob_Acad = .Cells(F_Ant, CtTP_Cob_Acont) - .Cells(F_Ant, CtTP_Cob_Tadm)
                    .Cells(F_Ant, CtTP_Cob_VRI) = Application.Round(R_Cob_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_Cob_Org) = R_Cob_Acad - .Cells(F_Ant, CtTP_Cob_VRI)
                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Cob_SinOrg) = .Cells(F_Ant, CtTP_Cob_Org) + .Cells(F_Ant, CtTP_Cob_Tadm)
                    '--- Reparto Importes Redistribuidos -----------------------------------
                    R_Rdt_Acad = .Cells(F_Ant, CtTP_RDT_Total) - .Cells(F_Ant, CtTP_RDT_Tadm)
                    .Cells(F_Ant, CtTP_RDT_VRI) = Application.Round(R_Rdt_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_RDT_Org) = R_Rdt_Acad - .Cells(F_Ant, CtTP_RDT_VRI)
'''                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_RDT_SinOrg) = .Cells(F_Ant, CtTP_RDT_Org) + .Cells(F_Ant, CtTP_RDT_Tadm)
                    '--- Reparto Importes ADXAplz -----------------------------------------
                    R_ADx_Acad = .Cells(F_Ant, CtTP_ADx_Total) - .Cells(F_Ant, CtTP_ADx_Tadm)
                    .Cells(F_Ant, CtTP_ADx_VRI) = Application.Round(R_ADx_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_ADx_Org) = R_ADx_Acad - .Cells(F_Ant, CtTP_ADx_VRI)
                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_ADx_SinOrg) = .Cells(F_Ant, CtTP_ADx_Org) + .Cells(F_Ant, CtTP_ADx_Tadm)
                    '--- Reparto Importes Pendiente de Cobro -------------------------------
                    R_Pdt_Acad = .Cells(F_Ant, CtTP_Pdt_Total) - .Cells(F_Ant, CtTP_Pdt_Tadm)
                    .Cells(F_Ant, CtTP_Pdt_VRI) = Application.Round(R_Pdt_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_Pdt_Org) = R_Pdt_Acad - .Cells(F_Ant, CtTP_Pdt_VRI)
                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Pdt_SinOrg) = .Cells(F_Ant, CtTP_Pdt_Org) + .Cells(F_Ant, CtTP_Pdt_Tadm)
                    '----- Reparto Importes SIN RDT ----------------------------------------
                    .Cells(F_Ant, CtTP_PdteRDT_Total) = .Cells(F_Ant, CtTP_Cob_Acont) - .Cells(F_Ant, CtTP_RDT_Total)
                    .Cells(F_Ant, CtTP_PdteRDT_Org) = .Cells(F_Ant, CtTP_Cob_Org) - .Cells(F_Ant, CtTP_RDT_Org)
                    .Cells(F_Ant, CtTP_PdteRDT_VRI) = .Cells(F_Ant, CtTP_Cob_VRI) - .Cells(F_Ant, CtTP_RDT_VRI)
                    .Cells(F_Ant, CtTP_PdteRDT_Tadm) = .Cells(F_Ant, CtTP_Cob_Tadm) - .Cells(F_Ant, CtTP_RDT_Tadm)
                    .Cells(F_Ant, CtTP_PdteRDT_SinOrg) = .Cells(F_Ant, CtTP_Cob_SinOrg) '''- .Cells(F_Ant, CtTP_RDT_SinOrg)
                End With
            End If
            ' =====================================================================================
        End If
Siguiente_Fila:

        Set RwPH = Nothing

        If Fila_DR Mod 500 = 0 Then
'            Debug.Print "Analizando Prog_BD:  " & Format(Fila_DR, "#,##0") & " de " & Format(Lo_BD.ListRows.Count, "#,##0")
            Form_Menu.TB_Informe = ProgresoTarea & vbCrLf & "Tiempo transcurrido:  " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & "Analizando Prog_BD:  " & Format(Fila_DR, "#,##0") & " de " & _
                                             Format(Lo_BD.ListRows.Count, "#,##0") & " reg.       Procesados:  " & Cont_Reg_Proc & " reg."
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
        End If

    Next Fila_DR
             
            ' ---------------------------=============  Cálculos Redistribución  de la última Fila ==================
            If F_Plan > 1 Then
                    F_Ant = F_Plan

                With Lo_TPResum.DataBodyRange
                    
                    '----- Reparto de lo Emitido -------------------------------------------
                    R_Emi_Acad = .Cells(F_Ant, CtTP_Emi_Total) - .Cells(F_Ant, CtTP_Emi_Tadm)
                    .Cells(F_Ant, CtTP_Emi_VRI) = Application.Round(R_Emi_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_Emi_Org) = R_Emi_Acad - .Cells(F_Ant, CtTP_Emi_VRI)
                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Emi_SinOrg) = .Cells(F_Ant, CtTP_Emi_Org) + .Cells(F_Ant, CtTP_Emi_Tadm)
                    '----- Reparto de lo Cobrado -------------------------------------------
                    R_Cob_Acad = .Cells(F_Ant, CtTP_Cob_Acont) - .Cells(F_Ant, CtTP_Cob_Tadm)
                    .Cells(F_Ant, CtTP_Cob_VRI) = Application.Round(R_Cob_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_Cob_Org) = R_Cob_Acad - .Cells(F_Ant, CtTP_Cob_VRI)
                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_Cob_SinOrg) = .Cells(F_Ant, CtTP_Cob_Org) + .Cells(F_Ant, CtTP_Cob_Tadm)
                    '--- Reparto Importes Redistribuidos -----------------------------------
                    R_Rdt_Acad = .Cells(F_Ant, CtTP_RDT_Total) - .Cells(F_Ant, CtTP_RDT_Tadm)
                    .Cells(F_Ant, CtTP_RDT_VRI) = Application.Round(R_Rdt_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_RDT_Org) = R_Rdt_Acad - .Cells(F_Ant, CtTP_RDT_VRI)
'''                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_RDT_SinOrg) = .Cells(F_Ant, CtTP_RDT_Org) + .Cells(F_Ant, CtTP_RDT_Tadm)
                    '--- Reparto Importes ADXAplz -----------------------------------
                    R_ADx_Acad = .Cells(F_Ant, CtTP_ADx_Total) - .Cells(F_Ant, CtTP_ADx_Tadm)
                    .Cells(F_Ant, CtTP_ADx_VRI) = Application.Round(R_ADx_Acad * .Cells(F_Ant, CtTP_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, CtTP_ADx_Org) = R_ADx_Acad - .Cells(F_Ant, CtTP_ADx_VRI)
                    If .Cells(F_Ant, CtTP_Orgánica) = "" Then .Cells(F_Ant, CtTP_ADx_SinOrg) = .Cells(F_Ant, CtTP_ADx_Org) + .Cells(F_Ant, CtTP_ADx_Tadm)
                    '----- Reparto Saldos ----------------------------------------------
                    .Cells(F_Ant, CtTP_PdteRDT_Total) = .Cells(F_Ant, CtTP_Cob_Acont) - .Cells(F_Ant, CtTP_RDT_Total)
                    .Cells(F_Ant, CtTP_PdteRDT_Org) = .Cells(F_Ant, CtTP_Cob_Org) - .Cells(F_Ant, CtTP_RDT_Org)
                    .Cells(F_Ant, CtTP_PdteRDT_VRI) = .Cells(F_Ant, CtTP_Cob_VRI) - .Cells(F_Ant, CtTP_RDT_VRI)
                    .Cells(F_Ant, CtTP_PdteRDT_Tadm) = .Cells(F_Ant, CtTP_Cob_Tadm) - .Cells(F_Ant, CtTP_RDT_Tadm)
                    .Cells(F_Ant, CtTP_PdteRDT_SinOrg) = .Cells(F_Ant, CtTP_Cob_SinOrg) ''' - .Cells(F_Ant, CtTP_RDT_SinOrg)
'                    '----- Pendiente de Cobro AD -------------------------------------------
'                    If RwPH.Range(BD_ACont_Vto) > AñoCont Then
'                        .Cells(F_Ant, CtTP_ADx_Total) = .Cells(F_Ant, CtTP_Emi_Total) - .Cells(F_Ant, CtTP_Cob_Acont)
'                        .Cells(F_Ant, CtTP_ADx_Org) = .Cells(F_Ant, CtTP_Emi_Org) - .Cells(F_Ant, CtTP_Cob_Org)
'                        .Cells(F_Ant, CtTP_ADx_VRI) = .Cells(F_Ant, CtTP_Emi_VRI) - .Cells(F_Ant, CtTP_Cob_VRI)
'                        .Cells(F_Ant, CtTP_ADx_Tadm) = .Cells(F_Ant, CtTP_Emi_Tadm) - .Cells(F_Ant, CtTP_Cob_Tadm)
'                        .Cells(F_Ant, CtTP_ADx_SinOrg) = .Cells(F_Ant, CtTP_Emi_SinOrg) - .Cells(F_Ant, CtTP_Cob_SinOrg)
'                    End If
                End With
            End If
    Range("h1") = " Última actualización: " & Now()
   
Restablecer_Valores:
Call Rut_EnableEvents_Status_Reset

'MsgBox "Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & "    Registros: " & Cont_Tot_Reg & "    Planes: " & F_Plan

Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
        "En la Nueva Consulta hay:  " & "  -  Tot.Reg. " & Cont_Tot_Reg & "    Planes: " & F_Plan & vbCrLf & Now()
        
Rut_On_Functions
    Call Rut_EnableEvents_Status_Reset
End Sub     ' Rut_Resumen_Tab_TitPropios  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
            

