Attribute VB_Name = "M22_Inf_EPs_para_UXXI"
Option Explicit

' ==================================================================================================================================
Sub Rut_Informe_EPs_para_UXXI_OLD()    ' ===============================================================================================
' ==================================================================================================================================
'--- Tabla Sht__Inf_EPs_UXXI  EPs-Resumen ----------------------
    Dim CursoAcad       As String:     CursoAcad = Prog__APP.Range("APP_CursAcad") '- Curso Académico
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")    '- Año Contable
    Dim ACont           As Integer:     ACont = Right(AñoCont, 2)
    Dim AñoContAnt      As Integer:     AñoContAnt = Left(CursoAcad, 4)             '- El 1º Año de Curso
    Dim AContAnt        As Integer:     AContAnt = Right(AñoContAnt, 2)             '- El 1º Año de Curso Corto
    Dim AñoContPos      As Integer:     AñoContPos = "20" & Right(CursoAcad, 2)     '- El 2º Año de Curso
    Dim AContPos        As Integer:     AContPos = Right(AñoContPos, 2)             '- El 2º Año de Curso Corto

Const BdUx_Cod_Plan                   As Integer = 1    ' col: a
Const BdUx_Curso_Acad                 As Integer = 2    ' col: b
Const BdUx_Año_Emi                    As Integer = 3    ' col: c
Const BdUx_Plan_Curso                 As Integer = 4    ' col: d
Const BdUx_NomPlan                    As Integer = 5    ' col: e
Const BdUx_Orgánica                   As Integer = 6    ' col: f
Const BdUx_Ref_JI                     As Integer = 7    ' col: g
Const BdUx_Concepto                   As Integer = 8    ' col: h
Const BdUx_Incidencia                 As Integer = 9    ' col: i
Const BdUx_Cant_Reg                   As Integer = 10   ' col: j
Const BdUx_Coef_VRI                   As Integer = 11   ' col: k
Const BdUx_Emi_Total                  As Integer = 12   ' col: l
Const BdUx_Emi_Org                    As Integer = 13   ' col: m
Const BdUx_Emi_VRI                    As Integer = 14   ' col: n
Const BdUx_Emi_Tadm                   As Integer = 15   ' col: o
Const BdUx_Emi_SinOrg                 As Integer = 16   ' col: p
Const BdUx_Cob_C_ACad                 As Integer = 17   ' col: q
Const BdUx_Cob_C_Acad_Adm             As Integer = 18   ' col: r
Const BdUx_Cob_ACont                  As Integer = 19   ' col: s
Const BdUx_Cob_Org                    As Integer = 20   ' col: t
Const BdUx_Cob_VRI                    As Integer = 21   ' col: u
Const BdUx_Cob_Tadm                   As Integer = 22   ' col: v
Const BdUx_Cob_SinOrg                 As Integer = 23   ' col: w
Const BdUx_RDT_Total                  As Integer = 24   ' col: x
Const BdUx_RDT_Org                    As Integer = 25   ' col: y
Const BdUx_RDT_VRI                    As Integer = 26   ' col: z
Const BdUx_RDT_Tadm                   As Integer = 27   ' col: aa
Const BdUx_PdteRDT_Total              As Integer = 28   ' col: ab
Const BdUx_PdteRDT_Org                As Integer = 29   ' col: ac
Const BdUx_PdteRDT_VRI                As Integer = 30   ' col: ad
Const BdUx_PdteRDT_TAdm               As Integer = 31   ' col: ae
Const BdUx_PdteRDT_SinOrg             As Integer = 32   ' col: af
Const BdUx_ADx_Total                  As Integer = 33   ' col: ag
Const BdUx_ADx_Org                    As Integer = 34   ' col: ah
Const BdUx_ADx_VRI                    As Integer = 35   ' col: ai
Const BdUx_ADx_Tadm                   As Integer = 36   ' col: aj
Const BdUx_ADx_SinOrg                 As Integer = 37   ' col: ak
Const BdUx_Aplz_Total                 As Integer = 38   ' col: al
Const BdUx_Aplz_Org                   As Integer = 39   ' col: am
Const BdUx_Aplz_VRI                   As Integer = 40   ' col: an
Const BdUx_Aplz_Tadm                  As Integer = 41   ' col: ao
Const BdUx_Aplz_SinOrg                As Integer = 42   ' col: ap
Const BdUx_Pdt_Total                  As Integer = 43   ' col: aq
Const BdUx_Pdt_Org                    As Integer = 44   ' col: ar
Const BdUx_Pdt_VRI                    As Integer = 45   ' col: as
Const BdUx_Pdt_Tadm                   As Integer = 46   ' col: at
Const BdUx_Pdt_SinOrg                 As Integer = 47   ' col: au
Const BdUx_Descripción                As Integer = 48   ' col: av

Dim Fila_DR             As Long
Dim F_Plan              As Long
Dim F_Ant               As Long

Dim Plan_Ant        As String
Dim Curso_Acad_Ant  As String
Dim Año_Emi_Ant     As String

Dim Cont_Tot_Reg        As Long
Dim Cont_Reg_Proc       As Long

Dim R_Emi_Acad          As Currency
Dim R_Cob_C_Acad_Adm    As Currency
Dim R_Cob_Acont         As Currency
Dim R_Rdt_Acad          As Currency
Dim R_ADx_Acad          As Currency

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
    ActiveCell.Offset(0, 1) = CursoAcad
    ActiveCell.Offset(0, -1) = "Enseñanzas de:  " & Range("APP_EFP_o_CFC")
    ActiveCell.Offset(0, 2) = " Última actualización: " & Now()
    ActiveCell.Offset(0, 1).Select
    ' ##################################################################################################################
    ' =============  Recorrer todos los Registros filtrados y Crear la Tabla de Planes de DR  =====================================
Dim RwBD        As ListRow
Dim RwUxi        As ListRow

    For Fila_DR = 1 To Lo_BD.ListRows.Count
    
        Set RwBD = Nothing
        Set RwBD = Lo_BD.ListRows(Fila_DR)
        
        If Len(RwBD.Range(BD_FEmi)) = 0 Then GoTo Siguiente_Fila
        If RwBD.Range(BD_Matricula) = "N" Then GoTo Siguiente_Fila
        If RwBD.Range(BD_ImpRec) < 0 Then GoTo Siguiente_Fila
        If RwBD.Range(BD_ImpAdm) < 0 Then GoTo Siguiente_Fila
        'If RwBD.Range(BD_Anul) = "S" Then GoTo Siguiente_Fila
        
        Cont_Reg_Proc = Cont_Reg_Proc + 1
        Cont_Tot_Reg = Cont_Tot_Reg + 1
        ' --------------=============  Tratamiento de los Datos  ==================
        If Plan_Ant & Curso_Acad_Ant & Año_Emi_Ant = RwBD.Range(BD_Plan) & RwBD.Range(BD_C_Acad) & Format(RwBD.Range(BD_FEmi), "yyyy") Then    ' ------- Control cambio de Plan de estudio y de Curso Académico ---------------
            ' Añadir JI si no está -------------------
            If RwBD.Range(BD_JI_Emi_Acad) <> "" Then
                If InStr(RwUxi.Range(BdUx_Ref_JI), RwBD.Range(BD_JI_Emi_Acad)) = 0 Then
                    RwUxi.Range(BdUx_Ref_JI) = RwUxi.Range(BdUx_Ref_JI) & " - " & RwBD.Range(BD_JI_Emi_Acad)
                End If
            End If
            ' Añadir Orgánica si no está ------------
            If RwBD.Range(BD_Orgánica) <> "" Then
                If InStr(RwUxi.Range(BdUx_Orgánica), RwBD.Range(BD_Orgánica)) = 0 Then
                    RwUxi.Range(BdUx_Orgánica) = RwUxi.Range(BdUx_Orgánica) & " - " & RwBD.Range(BD_Orgánica)
                End If
            End If
            
            '--- Imp Emitido --------
            If RwBD.Range(BD_ImpRec) > 0 Then
                RwUxi.Range(BdUx_Emi_Total) = RwUxi.Range(BdUx_Emi_Total) + RwBD.Range(BD_ImpRec)
            End If
                RwUxi.Range(BdUx_Emi_Tadm) = RwUxi.Range(BdUx_Emi_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
                        
            '--- Imp Cobrado Acont --------
            If RwBD.Range(BD_ImpCob) > 0 And RwBD.Range(BD_ACont_Cob) = AñoCont Then
                RwUxi.Range(BdUx_Cob_ACont) = RwUxi.Range(BdUx_Cob_ACont) + RwBD.Range(BD_ImpCob)
                RwUxi.Range(BdUx_Cob_Tadm) = RwUxi.Range(BdUx_Cob_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Imp Cobrado C_Acad --------
            If RwBD.Range(BD_ImpCob) > 0 Then
                RwUxi.Range(BdUx_Cob_C_ACad) = RwUxi.Range(BdUx_Cob_C_ACad) + RwBD.Range(BD_ImpCob)
                RwUxi.Range(BdUx_Cob_C_Acad_Adm) = RwUxi.Range(BdUx_Cob_C_Acad_Adm) + RwBD.Range(BD_Rec_Imp_Adm)
                '--- Imp RDT --------
                If RwBD.Range(BD_RDT) <> "" Then
                    RwUxi.Range(BdUx_RDT_Total) = RwUxi.Range(BdUx_RDT_Total) + RwBD.Range(BD_ImpCob)
                    RwUxi.Range(BdUx_RDT_Tadm) = RwUxi.Range(BdUx_RDT_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
                End If
            End If
            
            '--- Imp ADxAplz --------
            If RwBD.Range(BD_ImpCob) = 0 And RwBD.Range(BD_ACont_Vto) > AñoContPos Or _
               RwBD.Range(BD_ImpCob) > 0 And RwBD.Range(BD_ACont_Cob) = AñoContPos And RwBD.Range(BD_ACont_Vto) = AñoContPos Then
                    RwUxi.Range(BdUx_ADx_Total) = RwUxi.Range(BdUx_ADx_Total) + RwBD.Range(BD_ImpRec)
                    RwUxi.Range(BdUx_ADx_Tadm) = RwUxi.Range(BdUx_ADx_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Imp Aplazado --------
            If RwBD.Range(BD_ACont_Cob) = AñoContPos And RwBD.Range(BD_ACont_Vto) >= AñoContPos Then
                    RwUxi.Range(BdUx_Aplz_Total) = RwUxi.Range(BdUx_Aplz_Total) + RwBD.Range(BD_ImpRec)
                    RwUxi.Range(BdUx_Aplz_Tadm) = RwUxi.Range(BdUx_Aplz_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Imp Pdte_Cobro --------
            If RwBD.Range(BD_ImpCob) = 0 And RwBD.Range(BD_ACont_Emi) = AñoCont And RwBD.Range(BD_ACont_Vto) = AñoCont Then
                    RwUxi.Range(BdUx_Pdt_Total) = RwUxi.Range(BdUx_Pdt_Total) + RwBD.Range(BD_ImpRec)
                    RwUxi.Range(BdUx_Pdt_Tadm) = RwUxi.Range(BdUx_Pdt_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
            End If
            
            RwUxi.Range(BdUx_Cant_Reg) = RwUxi.Range(BdUx_Cant_Reg) + 1

        Else    '------ Es el primero de una serie y tengo que introducir los datos comunes ---------------------------------------------------------

            Set RwUxi = Nothing
            Set RwUxi = Lo_TPResum.ListRows.Add
            F_Plan = Lo_TPResum.ListRows.Count
           
            Plan_Ant = RwBD.Range(BD_Plan)
            Curso_Acad_Ant = RwBD.Range(BD_C_Acad)
            Año_Emi_Ant = Format(RwBD.Range(BD_FEmi), "yyyy")
            
            RwUxi.Range(BdUx_Cod_Plan) = RwBD.Range(BD_Plan)
            RwUxi.Range(BdUx_Curso_Acad) = RwBD.Range(BD_C_Acad)
            RwUxi.Range(BdUx_Año_Emi) = Format(RwBD.Range(BD_FEmi), "yyyy")
            RwUxi.Range(BdUx_Plan_Curso) = RwBD.Range(BD_Plan) & "_" & RwBD.Range(BD_C_Acad) & "_Cont" & Right(RwBD.Range(BD_FEmi), 2)
            
            RwUxi.Range(BdUx_NomPlan) = RwBD.Range(BD_NomPlan)
            
            RwUxi.Range(BdUx_Ref_JI) = RwBD.Range(BD_JI_Emi_Acad)
            RwUxi.Range(BdUx_Orgánica) = RwBD.Range(BD_Orgánica)
            
            If IsNumeric(Left(RwBD.Range(BD_Plan), 1)) Then RwUxi.Range(BdUx_Concepto) = 1311 Else RwUxi.Range(BdUx_Concepto) = 1311.03

            RwUxi.Range(BdUx_Cant_Reg) = 1
            RwUxi.Range(BdUx_Coef_VRI) = RwBD.Range(BD_Coef_VRI)
            
            
            '--- Imp Emitido --------
            If RwBD.Range(BD_ImpRec) > 0 Then
                RwUxi.Range(BdUx_Emi_Total) = RwUxi.Range(BdUx_Emi_Total) + RwBD.Range(BD_ImpRec)
            End If
                RwUxi.Range(BdUx_Emi_Tadm) = RwUxi.Range(BdUx_Emi_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
                        
            '--- Imp Cobrado Acont --------
            If RwBD.Range(BD_ImpCob) > 0 And RwBD.Range(BD_ACont_Cob) = AñoCont Then
                RwUxi.Range(BdUx_Cob_ACont) = RwUxi.Range(BdUx_Cob_ACont) + RwBD.Range(BD_ImpCob)
                RwUxi.Range(BdUx_Cob_Tadm) = RwUxi.Range(BdUx_Cob_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Imp Cobrado C_Acad --------
            If RwBD.Range(BD_ImpCob) > 0 Then
'                RwUxi.Range(BdUx_Cob_C_ACad) = RwUxi.Range(BdUx_Cob_C_ACad) + RwBD.Range(BD_ImpCob)
                RwUxi.Range(BdUx_Cob_C_Acad_Adm) = RwUxi.Range(BdUx_Cob_C_Acad_Adm) + RwBD.Range(BD_Rec_Imp_Adm)
                '--- Imp RDT --------
                If RwBD.Range(BD_RDT) <> "" Then
                    RwUxi.Range(BdUx_RDT_Total) = RwUxi.Range(BdUx_RDT_Total) + RwBD.Range(BD_ImpCob)
                    RwUxi.Range(BdUx_RDT_Tadm) = RwUxi.Range(BdUx_RDT_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
                End If
            End If
            
            '--- Imp ADxAplz --------
            If RwBD.Range(BD_ImpCob) = 0 And RwBD.Range(BD_ACont_Vto) > AñoContPos Or _
               RwBD.Range(BD_ImpCob) > 0 And RwBD.Range(BD_ACont_Cob) = AñoContPos And RwBD.Range(BD_ACont_Vto) = AñoContPos Then
                    RwUxi.Range(BdUx_ADx_Total) = RwUxi.Range(BdUx_ADx_Total) + RwBD.Range(BD_ImpRec)
                    RwUxi.Range(BdUx_ADx_Tadm) = RwUxi.Range(BdUx_ADx_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Imp Aplazado --------
            If RwBD.Range(BD_ACont_Cob) = AñoContPos And RwBD.Range(BD_ACont_Vto) > AñoContPos Then
                    RwUxi.Range(BdUx_Aplz_Total) = RwUxi.Range(BdUx_Aplz_Total) + RwBD.Range(BD_ImpRec)
                    RwUxi.Range(BdUx_Aplz_Tadm) = RwUxi.Range(BdUx_Aplz_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
            End If
            
            '--- Imp Pdte_Cobro --------
            If RwBD.Range(BD_ImpCob) = 0 And RwBD.Range(BD_ACont_Emi) = AñoCont And RwBD.Range(BD_ACont_Vto) = AñoCont Then
                    RwUxi.Range(BdUx_Pdt_Total) = RwUxi.Range(BdUx_Pdt_Total) + RwBD.Range(BD_ImpRec)
                    RwUxi.Range(BdUx_Pdt_Tadm) = RwUxi.Range(BdUx_Pdt_Tadm) + RwBD.Range(BD_Rec_Imp_Adm)
            End If
            
            ' -------=============  Genero Texto de Descripción del JI  ==================---------
            '- Generar el campo Descripción. -----------
            If Range("APP_PlanMicroCred") <> "" Then
                RwUxi.Range(BdUx_Descripción) = "LIQ-TitProp_" & RwBD.Range(BD_Plan) & _
                                               "-N Curso_" & CursoAcad & _
                                               " AñoCont_" & AñoCont & _
                                               " - MicCred_" & RwBD.Range("APP_PlanMicroCred") & _
                                               " - " & RwBD.Range(BD_NomPlan)
            Else
                RwUxi.Range(BdUx_Descripción) = "LIQ-TitProp_" & RwBD.Range(BD_Plan) & _
                                               "-N Curso_" & CursoAcad & _
                                               " AñoCont_" & AñoCont & _
                                               " - " & RwBD.Range(BD_NomPlan)
            End If
            
            ' =====================================================================================
            ' ---------------=============  Cálculos Redistribución de la Fila Anterior  ==================
            If F_Plan > 1 Then
                    F_Ant = F_Plan - 1
                
                With Lo_TPResum.DataBodyRange
                    
                    '----- Reparto Emitido -------------------------------------------
                    R_Emi_Acad = .Cells(F_Ant, BdUx_Emi_Total) - .Cells(F_Ant, BdUx_Emi_Tadm)
                    .Cells(F_Ant, BdUx_Emi_VRI) = Application.Round(R_Emi_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_Emi_Org) = R_Emi_Acad - .Cells(F_Ant, BdUx_Emi_VRI)
                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_Emi_SinOrg) = .Cells(F_Ant, BdUx_Emi_Org) + .Cells(F_Ant, BdUx_Emi_Tadm)
                    '----- Reparto Cobrado -------------------------------------------
                    R_Cob_Acont = .Cells(F_Ant, BdUx_Cob_ACont) - .Cells(F_Ant, BdUx_Cob_Tadm)
                    .Cells(F_Ant, BdUx_Cob_VRI) = Application.Round(R_Cob_Acont * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_Cob_Org) = R_Cob_Acont - .Cells(F_Ant, BdUx_Cob_VRI)
                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_Cob_SinOrg) = .Cells(F_Ant, BdUx_Cob_Org) + .Cells(F_Ant, BdUx_Cob_Tadm)
                    '--- Reparto RDT -----------------------------------
                    R_Rdt_Acad = .Cells(F_Ant, BdUx_RDT_Total) - .Cells(F_Ant, BdUx_RDT_Tadm)
                    .Cells(F_Ant, BdUx_RDT_VRI) = Application.Round(R_Rdt_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_RDT_Org) = R_Rdt_Acad - .Cells(F_Ant, BdUx_RDT_VRI)
'''                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_RDT_SinOrg) = .Cells(F_Ant, BdUx_RDT_Org) + .Cells(F_Ant, BdUx_RDT_Tadm)
                    '--- Reparto ADXAplz -----------------------------------------
                    R_ADx_Acad = .Cells(F_Ant, BdUx_ADx_Total) - .Cells(F_Ant, BdUx_ADx_Tadm)
                    .Cells(F_Ant, BdUx_ADx_VRI) = Application.Round(R_ADx_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_ADx_Org) = R_ADx_Acad - .Cells(F_Ant, BdUx_ADx_VRI)
                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_ADx_SinOrg) = .Cells(F_Ant, BdUx_ADx_Org) + .Cells(F_Ant, BdUx_ADx_Tadm)
                    '--- Reparto Aplazado -----------------------------------------
                    R_Aplz_Acad = .Cells(F_Ant, BdUx_Aplz_Total) - R_Cob_C_Acad_Adm
                    .Cells(F_Ant, BdUx_Aplz_VRI) = Application.Round(R_Aplz_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_Aplz_Org) = R_Aplz_Acad - .Cells(F_Ant, BdUx_Aplz_VRI)
                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_ADx_SinOrg) = .Cells(F_Ant, BdUx_ADx_Org) + .Cells(F_Ant, BdUx_ADx_Tadm)
'''                    '--- Reparto Importes Pendiente de Cobro -------------------------------
'''                    R_Pdt_Acad = .Cells(F_Ant, BdUx_Pdt_Total) - .Cells(F_Ant, BdUx_Pdt_Tadm)
'''                    .Cells(F_Ant, BdUx_Pdt_VRI) = Application.Round(R_Pdt_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
'''                    .Cells(F_Ant, BdUx_Pdt_Org) = R_Pdt_Acad - .Cells(F_Ant, BdUx_Pdt_VRI)
'''                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_Pdt_SinOrg) = .Cells(F_Ant, BdUx_Pdt_Org) + .Cells(F_Ant, BdUx_Pdt_Tadm)
                    '----- Reparto Importes SIN RDT ----------------------------------------
                    .Cells(F_Ant, BdUx_PdteRDT_Total) = .Cells(F_Ant, BdUx_Cob_C_ACad) - .Cells(F_Ant, BdUx_RDT_Total)
                    .Cells(F_Ant, BdUx_PdteRDT_TAdm) = .Cells(F_Ant, BdUx_Cob_C_Acad_Adm) - .Cells(F_Ant, BdUx_RDT_Tadm)
                    .Cells(F_Ant, BdUx_PdteRDT_Org) = .Cells(F_Ant, BdUx_Cob_Org) - .Cells(F_Ant, BdUx_RDT_Org)
                    .Cells(F_Ant, BdUx_PdteRDT_VRI) = .Cells(F_Ant, BdUx_Cob_VRI) - .Cells(F_Ant, BdUx_RDT_VRI)
                    .Cells(F_Ant, BdUx_PdteRDT_SinOrg) = .Cells(F_Ant, BdUx_Cob_SinOrg) '''- .Cells(F_Ant, BdUx_RDT_SinOrg)
                End With
            End If
            ' =====================================================================================
        End If
Siguiente_Fila:

        Set RwBD = Nothing

'        If Fila_DR Mod 500 = 0 Then
''            Debug.Print "Analizando Prog_BD:  " & Format(Fila_DR, "#,##0") & " de " & Format(Lo_BD.ListRows.Count, "#,##0")
'            Form_Menu.TB_Informe = ProgresoTarea & vbCrLf & "Tiempo transcurrido:  " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & "Analizando Prog_BD:  " & Format(Fila_DR, "#,##0") & " de " & _
'                                             Format(Lo_BD.ListRows.Count, "#,##0") & " reg.       Procesados:  " & Cont_Reg_Proc & " reg."
'            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
'        End If

    Next Fila_DR
             
            ' ---------------------------=============  Cálculos Redistribución  de la última Fila ==================
            If F_Plan > 1 Then
                    F_Ant = F_Plan

                With Lo_TPResum.DataBodyRange
                    
                    '----- Reparto Emitido -------------------------------------------
                    R_Emi_Acad = .Cells(F_Ant, BdUx_Emi_Total) - .Cells(F_Ant, BdUx_Emi_Tadm)
                    .Cells(F_Ant, BdUx_Emi_VRI) = Application.Round(R_Emi_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_Emi_Org) = R_Emi_Acad - .Cells(F_Ant, BdUx_Emi_VRI)
                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_Emi_SinOrg) = .Cells(F_Ant, BdUx_Emi_Org) + .Cells(F_Ant, BdUx_Emi_Tadm)
                    '----- Reparto Cobrado -------------------------------------------
                    R_Cob_Acont = .Cells(F_Ant, BdUx_Cob_ACont) - .Cells(F_Ant, BdUx_Cob_Tadm)
                    .Cells(F_Ant, BdUx_Cob_VRI) = Application.Round(R_Cob_Acont * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_Cob_Org) = R_Cob_Acont - .Cells(F_Ant, BdUx_Cob_VRI)
                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_Cob_SinOrg) = .Cells(F_Ant, BdUx_Cob_Org) + .Cells(F_Ant, BdUx_Cob_Tadm)
                    '--- Reparto RDT -----------------------------------
                    R_Rdt_Acad = .Cells(F_Ant, BdUx_RDT_Total) - .Cells(F_Ant, BdUx_RDT_Tadm)
                    .Cells(F_Ant, BdUx_RDT_VRI) = Application.Round(R_Rdt_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_RDT_Org) = R_Rdt_Acad - .Cells(F_Ant, BdUx_RDT_VRI)
'''                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_RDT_SinOrg) = .Cells(F_Ant, BdUx_RDT_Org) + .Cells(F_Ant, BdUx_RDT_Tadm)
                    '--- Reparto ADXAplz -----------------------------------------
                    R_ADx_Acad = .Cells(F_Ant, BdUx_ADx_Total) - .Cells(F_Ant, BdUx_ADx_Tadm)
                    .Cells(F_Ant, BdUx_ADx_VRI) = Application.Round(R_ADx_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
                    .Cells(F_Ant, BdUx_ADx_Org) = R_ADx_Acad - .Cells(F_Ant, BdUx_ADx_VRI)
                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_ADx_SinOrg) = .Cells(F_Ant, BdUx_ADx_Org) + .Cells(F_Ant, BdUx_ADx_Tadm)
'''                    '--- Reparto Importes Pendiente de Cobro -------------------------------
'''                    R_Pdt_Acad = .Cells(F_Ant, BdUx_Pdt_Total) - .Cells(F_Ant, BdUx_Pdt_Tadm)
'''                    .Cells(F_Ant, BdUx_Pdt_VRI) = Application.Round(R_Pdt_Acad * .Cells(F_Ant, BdUx_Coef_VRI) / 100, 2)
'''                    .Cells(F_Ant, BdUx_Pdt_Org) = R_Pdt_Acad - .Cells(F_Ant, BdUx_Pdt_VRI)
'''                    If .Cells(F_Ant, BdUx_Orgánica) = "" Then .Cells(F_Ant, BdUx_Pdt_SinOrg) = .Cells(F_Ant, BdUx_Pdt_Org) + .Cells(F_Ant, BdUx_Pdt_Tadm)
                    '----- Reparto Importes SIN RDT ----------------------------------------
                    .Cells(F_Ant, BdUx_PdteRDT_Total) = .Cells(F_Ant, BdUx_Cob_C_ACad) - .Cells(F_Ant, BdUx_RDT_Total)
                    .Cells(F_Ant, BdUx_PdteRDT_TAdm) = .Cells(F_Ant, BdUx_Cob_C_Acad_Adm) - .Cells(F_Ant, BdUx_RDT_Tadm)
                    .Cells(F_Ant, BdUx_PdteRDT_VRI) = .Cells(F_Ant, BdUx_Cob_VRI) - .Cells(F_Ant, BdUx_RDT_VRI)
                    .Cells(F_Ant, BdUx_PdteRDT_TAdm) = .Cells(F_Ant, BdUx_Cob_Tadm) - .Cells(F_Ant, BdUx_RDT_Tadm)
                    .Cells(F_Ant, BdUx_PdteRDT_SinOrg) = .Cells(F_Ant, BdUx_Cob_SinOrg) '''- .Cells(F_Ant, BdUx_RDT_SinOrg)
                End With
            End If
    Range("h1") = " Última actualización: " & Now()
   
Restablecer_Valores:
Call Rut_EnableEvents_Status_Reset

'MsgBox "Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & "    Registros: " & Cont_Tot_Reg & "    Planes: " & F_Plan

Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
        "En la Nueva Consulta hay:  " & "  -  Tot.Reg. " & Cont_Tot_Reg & "    Planes: " & F_Plan & vbCrLf & Now()
        
Rut_On_Functions
    Application.ScreenUpdating = True
    Call Rut_EnableEvents_Status_Reset
    Application.Speech.Speak "Proceso completado."
End Sub     ' Rut_Resumen_Tab_TitPropios  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
            





