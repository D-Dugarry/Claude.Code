Attribute VB_Name = "M_600_Informe_EP_RSm"
Option Explicit

' ==================================================================================================================================
Sub Rut_Recalcular_Tabla_Inf_RSm()    ' ===============================================================================================
' ==================================================================================================================================
    On Error GoTo ErrorHandler
    Dim Fila_EPs        As Long
    Dim F_Plan          As Long
    Dim F_Ant           As Long
    Dim Plan            As String
    Dim CAcad           As String
    Dim AñoEmi          As String
    Dim AñoCob          As String
    Dim TIO_EP          As String
    Dim Anexo           As String
    Dim Cont_Tot_Reg        As Long
    Dim Cont_Reg_Proc       As Long
    
    Dim T_RDT_Tot          As Currency
    Dim T_RDT_Adm          As Currency
    Dim T_Emi_Tot          As Currency
    Dim T_Emi_Adm          As Currency
    Dim T_Cob_Tot          As Currency
    Dim T_Cob_Adm          As Currency
    Dim CountRec           As Long
    
    Dim R_Emi_Acad          As Currency
    Dim R_Cob_Acad          As Currency
    Dim R_Rdt_Acad          As Currency
    Dim ProgresoTarea       As String
    Dim RetenVRI            As Double
    Dim IndxTipRec          As Integer
    
    Dim RowFound        As Variant
    Dim Tipo_Rec        As Variant:    Tipo_Rec = Array("Emitido", "EjeAnt", "Añejo", "ADxAplz", "Aplazado")

    Dim Lo_EPs          As ListObject:       Set Lo_EPs = Sht__BD.ListObjects(1)
    Dim Lo_InfRSm       As ListObject:       Set Lo_InfRSm = Wk_Inf_Rsm.ListObjects(1)
    Dim Lo_RetVRI       As ListObject:       Set Lo_RetVRI = Prog_RetVRI.ListObjects(1)
    
    Call Rut_Off_Functions

    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Ordenar por PLAN y DNI ==================
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Call Rut_Lo_Sort(Lo_EPs, BD_TIO_EP, xlAscending, True)
    Call Rut_Lo_Sort(Lo_EPs, BD_Plan, xlAscending)
    Call Rut_Lo_Sort(Lo_EPs, BD_C_Acad, xlAscending)
    Call Rut_Lo_Sort(Lo_EPs, BD_ACont_Emi, xlAscending)
'    Call Rut_Lo_Sort(Lo_EPs, BD_Ref, xlAscending)
    '- Preparar Tabla de Wk_Inf_Rsm ==================
    Wk_Inf_Rsm.Visible = xlSheetVisible
    Wk_Inf_Rsm.Select
    Call Rut_Lo_WrkSht_Preparar(Wk_Inf_Rsm)
    Wk_Inf_Rsm.Unprotect
    '- Vacío la Tabla de Tit.Prop.  =====================================
    Call Rut_Lo_Filtros_Quitar(Lo_InfRSm)
    If Not Lo_InfRSm.DataBodyRange Is Nothing Then Lo_InfRSm.DataBodyRange.Delete
    ' ==================================================================================================================================
    ' ###############################  Genero la Tabla de Planes de DR  #####################################
    Plan = "":   CAcad = "":   AñoEmi = 0: Cont_Tot_Reg = 0:
    Range("TP_Cod_Plan") = ""
    Range("TP_Cod_Plan").Select
'    ActiveCell.Offset(0, 1) = Prog__APP.Range("APP_CursAcad")
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Offset(0, 1) = " Última actualización: " & Now()
    ActiveCell.Offset(0, 1).Select
    ' ##################################################################################################################
    ' =============  Recorrer todos los Registros filtrados y Crear la Tabla de Planes de DR  =====================================
    Dim RwEPs        As ListRow
    Dim RwInf        As ListRow

    For Fila_EPs = 1 To Lo_EPs.ListRows.Count
        
        With Lo_EPs.DataBodyRange
'            If .Cells(Fila_EPs, BD_TIO_EP) = "AFC" Then GoTo Siguiente_Fila
'            If .Cells(Fila_EPs, BD_TIO_EP) = "CFC" Then GoTo Siguiente_Fila
            If .Cells(Fila_EPs, BD_TIO_EP) = "Doctorado" Then GoTo Siguiente_Fila
'            If .Cells(Fila_EPs, BD_TIO_EP) = "EFP" Then GoTo Siguiente_Fila
            If .Cells(Fila_EPs, BD_TIO_EP) = "Grado" Then GoTo Siguiente_Fila
            If .Cells(Fila_EPs, BD_TIO_EP) = "Master" Then GoTo Siguiente_Fila
            If .Cells(Fila_EPs, BD_TIO_EP) = "Rec_Adm" Then GoTo Siguiente_Fila
            If .Cells(Fila_EPs, BD_TIO_EP) = "TNCT" Then GoTo Siguiente_Fila
            
'            If .Cells(Fila_EPs, BD_Plan) <> "9361" Then GoTo Siguiente_Fila
            
        End With
        
        Set RwEPs = Nothing
        Set RwEPs = Lo_EPs.ListRows(Fila_EPs)
        
        Cont_Reg_Proc = Cont_Reg_Proc + 1
        Cont_Tot_Reg = Cont_Tot_Reg + 1
        ' --------------=============  Tratamiento de los Datos  ==================
        If Plan & CAcad & AñoEmi = RwEPs.Range(BD_Plan) & RwEPs.Range(BD_C_Acad) & RwEPs.Range(BD_ACont_Emi) Then GoTo Siguiente_Fila   ' ------- Control cambio de Plan de estudio y de Curso Académico ---------------
            
        Plan = RwEPs.Range(BD_Plan)
        CAcad = RwEPs.Range(BD_C_Acad)
        AñoEmi = RwEPs.Range(BD_ACont_Emi)
        AñoCob = RwEPs.Range(BD_ACont_Cob)
        TIO_EP = RwEPs.Range(BD_TIO_EP)
        
        '- Por cada tipo de Recibos: Emitido, AjeAnt, Añejo, ADxAplz, Aplazado
        For IndxTipRec = LBound(Tipo_Rec) To UBound(Tipo_Rec)
            
            With Lo_EPs.DataBodyRange
            
                '- Acumulo Importes Rec. Emitidos ---------
                T_Emi_Tot = Application.SumIfs(.Columns(BD_ImpRec), _
                                                            .Columns(BD_Plan), "=" & Plan, _
                                                            .Columns(BD_C_Acad), "=" & CAcad, _
                                                            .Columns(BD_ACont_Emi), "=" & AñoEmi, _
                                                            .Columns(BD_ImpRec), ">0", _
                                                            .Columns(BD_Tipo_Rec), Tipo_Rec(IndxTipRec), _
                                                            .Columns(BD_TIO_EP), TIO_EP)
                If T_Emi_Tot = 0 Then GoTo Siguiente_Tipo_Rec
                
                '- Acumulo Importes Rec. Cobrados -------
                T_Cob_Tot = Application.SumIfs(.Columns(BD_ImpCob), _
                                                            .Columns(BD_Plan), "=" & Plan, _
                                                            .Columns(BD_C_Acad), "=" & CAcad, _
                                                            .Columns(BD_ACont_Emi), "=" & AñoEmi, _
                                                            .Columns(BD_ImpRec), ">0", _
                                                            .Columns(BD_Tipo_Rec), Tipo_Rec(IndxTipRec), _
                                                            .Columns(BD_TIO_EP), TIO_EP)
                '- Acumulo Importes Adm. Emitidos -------
                T_Emi_Adm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                            .Columns(BD_Plan), "=" & Plan, _
                                                            .Columns(BD_C_Acad), "=" & CAcad, _
                                                            .Columns(BD_ACont_Emi), "=" & AñoEmi, _
                                                            .Columns(BD_ImpRec), ">0", _
                                                            .Columns(BD_Tipo_Rec), Tipo_Rec(IndxTipRec), _
                                                            .Columns(BD_TIO_EP), TIO_EP)
                '- Acumulo Importes Adm. Cobrados ------
                T_Cob_Adm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                            .Columns(BD_Plan), "=" & Plan, _
                                                            .Columns(BD_C_Acad), "=" & CAcad, _
                                                            .Columns(BD_ACont_Emi), "=" & AñoEmi, _
                                                            .Columns(BD_ImpCob), ">0", _
                                                            .Columns(BD_Tipo_Rec), Tipo_Rec(IndxTipRec), _
                                                            .Columns(BD_TIO_EP), TIO_EP)
                
                '- Cuento Cantidad de Rec. ----------------
                CountRec = Application.CountIfs( _
                                                            .Columns(BD_Plan), "=" & Plan, _
                                                            .Columns(BD_C_Acad), "=" & CAcad, _
                                                            .Columns(BD_ACont_Emi), "=" & AñoEmi, _
                                                            .Columns(BD_ImpRec), ">0", _
                                                            .Columns(BD_Tipo_Rec), Tipo_Rec(IndxTipRec), _
                                                            .Columns(BD_TIO_EP), TIO_EP)
            
                '- Acumulo Importes Rec. RDT Redistribuidos -------
                T_RDT_Tot = Application.SumIfs(.Columns(BD_ImpCob), _
                                                            .Columns(BD_RDT), "<>", _
                                                            .Columns(BD_Plan), "=" & Plan, _
                                                            .Columns(BD_C_Acad), "=" & CAcad, _
                                                            .Columns(BD_ACont_Emi), "=" & AñoEmi, _
                                                            .Columns(BD_ImpRec), ">0", _
                                                            .Columns(BD_Tipo_Rec), Tipo_Rec(IndxTipRec), _
                                                            .Columns(BD_TIO_EP), TIO_EP)
                
                '- Acumulo Importes Adm. Cobrados y RDT Redistribuidos ------
                T_RDT_Adm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                            .Columns(BD_RDT), "<>", _
                                                            .Columns(BD_Plan), "=" & Plan, _
                                                            .Columns(BD_C_Acad), "=" & CAcad, _
                                                            .Columns(BD_ACont_Emi), "=" & AñoEmi, _
                                                            .Columns(BD_ImpCob), ">0", _
                                                            .Columns(BD_Tipo_Rec), Tipo_Rec(IndxTipRec), _
                                                            .Columns(BD_TIO_EP), TIO_EP)
                
            End With    '- Lo_EPs.DataBodyRange
            
            '------ Es el primero de una serie y tengo que introducir los datos comunes ---------------------------------------------------------
            Set RwInf = Nothing
            Set RwInf = Lo_InfRSm.ListRows.Add
            F_Plan = Lo_InfRSm.ListRows.Count
            
            ' Buscar Orgánica -------------------
            RowFound = Application.Match(RwEPs.Range(BD_Plan) & "_" & RwEPs.Range(BD_C_Acad), Lo_RetVRI.DataBodyRange.Columns(3), 0)
            If IsError(RowFound) Then
                RwInf.Range(IRs_Orgánica) = "Not Foud"
                                 RetenVRI = RwEPs.Range(BD_Dtos)
            Else
                RwInf.Range(IRs_Orgánica) = Lo_RetVRI.DataBodyRange.Columns(5).Cells(RowFound)
                                 RetenVRI = Lo_RetVRI.DataBodyRange.Columns(4).Cells(RowFound)
            End If
            
            RwInf.Range(IRs_Ret_VRI) = RetenVRI
            
            '- Completar importes ----------------
            RwInf.Range(IRs_Emi_Tot) = T_Emi_Tot
            RwInf.Range(IRs_Emi_Adm) = T_Emi_Adm
            RwInf.Range(IRs_Emi_VRI) = Application.Round((T_Emi_Tot - T_Emi_Adm) * RetenVRI / 100, 2)
            RwInf.Range(IRs_Emi_Org) = (T_Emi_Tot - T_Emi_Adm) - RwInf.Range(IRs_Emi_VRI)
'            RwInf.Range(IRs_Emi_SinOrg) = RwInf.Range(IRs_Emi_Org) + T_Emi_Adm
            
            RwInf.Range(IRs_RDT_Tot) = T_RDT_Tot
            RwInf.Range(IRs_RDT_Adm) = T_RDT_Adm
            RwInf.Range(IRs_RDT_Pte_Tot) = T_Cob_Tot - T_RDT_Tot
            
            If Tipo_Rec(IndxTipRec) <> "ADxAplz" Then
                RwInf.Range(IRs_Cob_Tot) = T_Cob_Tot
                RwInf.Range(IRs_Cob_Adm) = T_Cob_Adm
                RwInf.Range(IRs_Cob_VRI) = Application.Round((T_Cob_Tot - T_Cob_Adm) * RetenVRI / 100, 2)
                RwInf.Range(IRs_Cob_Org) = (T_Cob_Tot - T_Cob_Adm) - RwInf.Range(IRs_Cob_VRI)
'                RwInf.Range(IRs_Cob_SinOrg) = RwInf.Range(IRs_Cob_Org) + T_Cob_Adm
            Else
                RwInf.Range(IRs_AD_Tot) = RwInf.Range(IRs_Emi_Tot)
                RwInf.Range(IRs_AD_Adm) = RwInf.Range(IRs_Emi_Adm)
                RwInf.Range(IRs_AD_VRI) = RwInf.Range(IRs_Emi_VRI)
                RwInf.Range(IRs_AD_Org) = RwInf.Range(IRs_Emi_Org)
'                RwInf.Range(IRs_AD_SinOrg) = RwInf.Range(IRs_Emi_SinOrg)
            End If
            
            RwInf.Range(IRs_Cant_Reg) = CountRec
            
            Select Case Tipo_Rec(IndxTipRec)
                
                Case "Añejo"
                        RwInf.Range(IRs_Incidencia) = "Añejo_" & RwEPs.Range(BD_C_Acad)
                        Anexo = "_Cont" & Right(Prog__APP.Range("APP_AñoCont"), 2) & " Añejo"
                Case "EjeAnt"
                        RwInf.Range(IRs_Incidencia) = "Pdte.Cob." & Prog__APP.Range("APP_AñoCont") - 1
                        Anexo = "_Cont" & Right(RwEPs.Range(BD_ACont_Emi), 2) & " Em_" & Right(RwEPs.Range(BD_C_Acad), 5)
                Case "Emitido"
                        RwInf.Range(IRs_Incidencia) = "Emi Curso_" & RwEPs.Range(BD_C_Acad)
                        Anexo = "_Cont" & Right(RwEPs.Range(BD_ACont_Emi), 2) & " Em_" & Right(RwEPs.Range(BD_C_Acad), 5)
                Case "ADxAplz"
                        RwInf.Range(IRs_Incidencia) = "ADxAplz_" & Prog__APP.Range("APP_AñoCont")
                        Anexo = "_Cont" & Right(RwEPs.Range(BD_ACont_Emi), 2) & " AD_" & Right(RwEPs.Range(BD_C_Acad), 5)
                Case "Aplazado"
                        RwInf.Range(IRs_Incidencia) = "ADxAplz_" & Prog__APP.Range("APP_AñoCont") - 1
                        Anexo = "_Cont" & Right(RwEPs.Range(BD_ACont_Emi), 2) + 1 & " Rm_" & Right(RwEPs.Range(BD_C_Acad), 5)
            End Select
            
            RwInf.Range(IRs_Cod_Plan) = RwEPs.Range(BD_Plan)
            RwInf.Range(IRs_Curso_Acad) = RwEPs.Range(BD_C_Acad)
            RwInf.Range(IRs_Año_Emi) = RwEPs.Range(BD_ACont_Emi)
            RwInf.Range(IRs_Plan_Curso) = RwEPs.Range(BD_Plan) & "_" & RwEPs.Range(BD_C_Acad) & Anexo
            RwInf.Range(IRs_NomPlan) = RwEPs.Range(BD_NomPlan)
            RwInf.Range(IRs_Ref_JI) = RwEPs.Range(BD_JI_Emi_Acad)
            RwInf.Range(IRs_Concepto) = RwEPs.Range(BD_Concepto)

            If RwEPs.Range(BD_Tipo_Rec) = "Aplazado" Then
                RwInf.Range(IRs_Descripc) = "LIQ-EP_Plan_" & RwEPs.Range(BD_Plan) & _
                                            "-N Curso_" & RwEPs.Range(BD_C_Acad) & _
                                            "_AñoCont_" & Prog__APP.Range("APP_AñoCont") & _
                                            " TipoRec_Rms443" & _
                                            " - " & RwEPs.Range(BD_NomPlan)
            Else
                RwInf.Range(IRs_Descripc) = "LIQ-EP_Plan_" & RwEPs.Range(BD_Plan) & _
                                            "-N Curso_" & RwEPs.Range(BD_C_Acad) & _
                                            "_AñoCont_" & Prog__APP.Range("APP_AñoCont") & _
                                            " TipoRec_" & RwEPs.Range(BD_Tipo_Rec) & _
                                            " - " & RwEPs.Range(BD_NomPlan)
            End If

''                    '----- Reparto de lo Emitido -------------------------------------------
''                    R_Emi_Acad = .Cells(F_Ant, IRs_Emi_Tot) - .Cells(F_Ant, IRs_Emi_Adm)
''                    .Cells(F_Ant, IRs_Emi_VRI) = Application.Round(R_Emi_Acad * .Cells(F_Ant, IRs_Ret_VRI) / 100, 2)
''                    .Cells(F_Ant, IRs_Emi_Org) = R_Emi_Acad - .Cells(F_Ant, IRs_Emi_VRI)
''                    If .Cells(F_Ant, IRs_Orgánica) = "" Then .Cells(F_Ant, IRs_Emi_SinOrg) = .Cells(F_Ant, IRs_Emi_Org) + .Cells(F_Ant, IRs_Emi_Adm)
''                    '----- Reparto de lo Cobrado -------------------------------------------
''                    R_Cob_Acad = .Cells(F_Ant, IRs_Cob_Tot) - .Cells(F_Ant, IRs_Cob_Adm)
''                    .Cells(F_Ant, IRs_Cob_VRI) = Application.Round(R_Cob_Acad * .Cells(F_Ant, IRs_Ret_VRI) / 100, 2)
''                    .Cells(F_Ant, IRs_Cob_Org) = R_Cob_Acad - .Cells(F_Ant, IRs_Cob_VRI)
''                    If .Cells(F_Ant, IRs_Orgánica) = "" Then .Cells(F_Ant, IRs_Cob_SinOrg) = .Cells(F_Ant, IRs_Cob_Org) + .Cells(F_Ant, IRs_Cob_Adm)
''                    '--- Reparto Importes Redistribuidos -----------------------------------
''                    R_Rdt_Acad = .Cells(F_Ant, IRs_RDT_Tot) - .Cells(F_Ant, IRs_RDT_Adm)
''                    .Cells(F_Ant, IRs_RDT_VRI) = Application.Round(R_Rdt_Acad * .Cells(F_Ant, IRs_Ret_VRI) / 100, 2)
''                    .Cells(F_Ant, IRs_RDT_Org) = R_Rdt_Acad - .Cells(F_Ant, IRs_RDT_VRI)
''                    If .Cells(F_Ant, IRs_Orgánica) = "" Then .Cells(F_Ant, IRs_RDT_SinOrg) = .Cells(F_Ant, IRs_RDT_Org) + .Cells(F_Ant, IRs_RDT_Adm)
''                    '----- Reparto Saldos --------------------------------------------------
''                    .Cells(F_Ant, IRs_RDT_Pte_Tot) = .Cells(F_Ant, IRs_Cob_Tot) - .Cells(F_Ant, IRs_RDT_Tot)
''                    .Cells(F_Ant, IRs_RDT_Pte_Org) = .Cells(F_Ant, IRs_Cob_Org) - .Cells(F_Ant, IRs_RDT_Org)
''                    .Cells(F_Ant, IRs_RDT_Pte_VRI) = .Cells(F_Ant, IRs_Cob_VRI) - .Cells(F_Ant, IRs_RDT_VRI)
''                    .Cells(F_Ant, IRs_RDT_Pte_Adm) = .Cells(F_Ant, IRs_Cob_Adm) - .Cells(F_Ant, IRs_RDT_Adm)
''                    .Cells(F_Ant, IRs_RDT_Pte_SinOrg) = .Cells(F_Ant, IRs_Cob_SinOrg) - .Cells(F_Ant, IRs_RDT_SinOrg)
''                    '----- Pendiente de Cobro AD -------------------------------------------
''                    .Cells(F_Ant, IRs_AD_Tot) = .Cells(F_Ant, IRs_Emi_Tot) - .Cells(F_Ant, IRs_Cob_Tot)
''                    .Cells(F_Ant, IRs_AD_Org) = .Cells(F_Ant, IRs_Emi_Org) - .Cells(F_Ant, IRs_Cob_Org)
''                    .Cells(F_Ant, IRs_AD_VRI) = .Cells(F_Ant, IRs_Emi_VRI) - .Cells(F_Ant, IRs_Cob_VRI)
''                    .Cells(F_Ant, IRs_AD_Adm) = .Cells(F_Ant, IRs_Emi_Adm) - .Cells(F_Ant, IRs_Cob_Adm)
''                    .Cells(F_Ant, IRs_AD_SinOrg) = .Cells(F_Ant, IRs_Emi_SinOrg) - .Cells(F_Ant, IRs_Cob_SinOrg)
Siguiente_Tipo_Rec:
        Next IndxTipRec
Siguiente_Fila:

        If Fila_EPs Mod 1000 = 0 Then
            Debug.Print "Analizando Prog_TitPH:  " & Format(Fila_EPs, "#,##0") & " de " & Format(Lo_EPs.ListRows.Count, "#,##0")
'            Form_Menu.Tbx_Informe = ProgresoTarea & vbCrLf & "Tiempo transcurrido:  " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & "Analizando Prog_TitPH:  " & Format(Fila_EPs, "#,##0") & " de " & _
'                                             Format(Lo_EPs.ListRows.Count, "#,##0") & " reg.       Procesados:  " & Cont_Reg_Proc & " reg."
'            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
        End If

    Next Fila_EPs
             
    Range("a1") = " Última actualiz. " & Now()
   
Restablecer_Valores:
Rut_Events_Status_Choose ("ON")

MsgBox "Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & "    Registros: " & Cont_Tot_Reg & "    Planes: " & F_Plan

'Prog__APP.Range("APP_Task_Inf") = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
        "En la Nueva Consulta hay:  " & "  -  Tot.Reg. " & Cont_Tot_Reg & "    Planes: " & F_Plan & vbCrLf & Now()
        
Rut_On_Functions
Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbLf & _
           "Procesando la fila " & Fila_EPs & " de " & Lo_EPs.ListRows.Count, vbCritical
    Rut_Events_Status_Choose ("ON")
    Rut_On_Functions
End Sub     ' Rut_Resumen_Tab_TitPropios  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
            







