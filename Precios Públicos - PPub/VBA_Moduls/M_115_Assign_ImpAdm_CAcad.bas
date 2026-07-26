Attribute VB_Name = "M_115_Assign_ImpAdm_CAcad"
'Rev.: 2026-01-26
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Identificar del C_Acad_Pos, los 1º Rec. de c/matrícula para obtener la T-Adm ---------------------------------------------
'-      ClearContents de BD_Rec_Imp_Adm, de Recibos con Imp.Adm. < 0 -->> Son ajustes de matrículas
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Assign_Imp_AdmAcad_C_Acad_Pos(LoBDatos As ListObject)

Debug.Print ">>> Rut_Assign_Imp_AdmAcad_C_Acad_Pos"
    Dim CantImpAcad     As Long
    Dim CantImpAdm      As Long
    Dim CantImpDto      As Long
    Dim ImpTAcad        As Currency
    Dim ImpTAdm         As Currency
    Dim ImpTAdmNeg      As Currency
    Dim ImpTDto         As Currency
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    Dim Fila            As Long
    Dim TF_LoBDatos     As Long:        TF_LoBDatos = LoBDatos.ListRows.Count
    Dim RowsDel         As Long
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad          As String:      C_Acad = Prog__APP.Range("APP_CursAcad")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim PlanDNI_Ant     As String:      PlanDNI_Ant = ""
    Dim PlanDNI_New     As String:
    Dim TxT_ProgIni     As String
    Dim TxT_Progreso    As String
    Dim RowData         As ListRow
    Dim RowsFind        As Variant
    Dim Sh_Data         As Worksheet:   Set Sh_Data = LoBDatos.Parent

    '- Visualizo el progreso ----------------------------------------------------------------------------------------
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Identificar ImpAcad ImpAdm e ImpDto del Curso: " & _
                         C_Acad & ", en " & Format(LoBDatos.ListRows.Count, "#,##0") & " reg.", 0, , , , , , 2)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    
    '-ClearContents de Rec. BD_C_Acad = C_Acad -------------------------------------
    Call Rut_Lo_Filtros_Quitar(LoBDatos)
    With LoBDatos
        .Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & C_Acad
        RowsFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RowsFind > 0 Then
            .DataBodyRange.Columns(BD_Rec_Imp_Acad).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_Rec_Imp_Adm).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_Rec_Imp_Dto).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_Obs_Conta).SpecialCells(xlCellTypeVisible).ClearContents
        End If
    End With
    
    Call Rut_Lo_Filtros_Quitar(LoBDatos)
    ' Ordenar por columnas  ------------------------------
        Call Rut_Lo_Sort(LoBDatos, BD_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_ActivEco, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_Plan, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_DNI, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_NumRec, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_Ref, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    
    '- Recorro toda la tabla LoBDatos -----------------------------------------------------------------
    For Fila = 1 To TF_LoBDatos   '--- Bucle para recorrer todas la filas de LoBDatos
        Set RowData = LoBDatos.ListRows(Fila)
        If RowData.Range(BD_C_Acad) <> C_Acad Then GoTo Sig_Reg         '- NO tenemos en cuenta loas Recibos de Otros C_Acad, porque NO tenemos toda la información
        If RowData.Range(BD_ActivEco) = 4 Then GoTo Sig_Reg             '- NO tenemos en cuenta los Recibos de AE4,  ya vienen con sus ImpAdm de AE4x4
        If RowData.Range(BD_ActivEco) > 6 Then                          '- NO tenemos en cuenta los Recibos de Movimiento
                                                Fila = TF_LoBDatos
                                                GoTo Fin_Bucle
                                                End If
        PlanDNI_New = RowData.Range(BD_Plan) & "_" & RowData.Range(BD_DNI)
        If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
           PlanDNI_Ant = PlanDNI_New
                    If RowData.Range(BD_Anul) <> "S" Then
                            RowData.Range(BD_Rec_Imp_Acad) = RowData.Range(BD_ImpAcad)
                            RowData.Range(BD_Rec_Imp_Adm) = RowData.Range(BD_ImpAdm)
                            RowData.Range(BD_Rec_Imp_Dto) = RowData.Range(BD_ImpDto)
                            CantImpAdm = CantImpAdm + 1
                End If
        End If
Sig_Reg:
        If Fila Mod 4000 = 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Indentificados " & CantImpAdm & " Rec. de AE_2/5y6 y de " & AñoCont & ", con Imp.Adm. en: ", 0, _
                                Format(Fila, "#,##0") & "reg.", "de " & Format(TF_LoBDatos, "#,##0") & "reg.", TxT_Progreso, True, , 2)
        End If
Fin_Bucle:
    Next
    
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Adm. a " & CantImpAdm & " Rec. de AE_2/5y6 y de " & AñoCont & ", de un total de: ", 0, _
                        Format(TF_LoBDatos, "#,##0") & "reg.", , TxT_ProgIni, True, , 2)
    
'''    '- ------------------------------------------------------------------------------------------------------------------
'''    '- Plan="M013", Seminario Mayores 25a
'''    '- Imp.Adm. ==a==>>Imp.Rec.Adm. ---------------------------------------------------------------------------------------
'''        Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(LoBDatos, BD_ImpAcad, BD_Rec_Imp_Acad, BD_Plan, "=M013B")
'''        Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(LoBDatos, BD_ImpAdm, BD_Rec_Imp_Adm, BD_Plan, "=M013B")
'''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Adm. ==a==>>Imp.Rec.Adm., a Todos los Rec. con Plan=M013B, Seminario Mayores 25a", 0)
'''
'''    '- ------------------------------------------------------------------------------------------------------------------
'''    '- Plan="UPUA", Cursos universidad permanente
'''    '- Imp.Adm. ==a==>>Imp.Rec.Adm. ---------------------------------------------------------------------------------------
'''        Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(LoBDatos, BD_ImpRec, BD_Rec_Imp_Acad, BD_Plan, "=UPUA")
'''        Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(LoBDatos, BD_ImpAdm, BD_Rec_Imp_Adm, BD_Plan, "=UPUA")
'''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Adm. ==a==>>Imp.Rec.Adm., a Todos los Rec. con Plan=UPUA, Cursos universidad permanente", 0)
'''
'''    '- ------------------------------------------------------------------------------------------------------------------
'''    '- Plan="PNB1", Pruebas Competencias Idiomas
'''    '- Imp.Rec. ==a==>>Imp.Rec.Adm. (porque puede tener Dto's) ------------------------------------------------------------
'''        Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(LoBDatos, BD_ImpRec, BD_Rec_Imp_Adm, BD_Plan, "=PNB1")
'''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Rec. ==a==>>Imp.Rec.Adm., a Todos los Rec. con Plan=PNB1, Pruebas Competencias Idiomas", 0)
    
'- ------------------------------------------------------------------------------------------------------------------
'- AE=80, Pruebas Acceso UA
'- Imp.Rec. ==a==>> Imp.Rec.Adm. (porque puede tener Dto's) ------------------------------------------------------------
    Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(LoBDatos, BD_ImpRec, BD_Rec_Imp_Adm, BD_ActivEco, "=80")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Rec. ==a==>>Imp.Rec.Adm., a Todos los Rec. con AE=80, Pruebas Acceso UA", 0)
    
'- ------------------------------------------------------------------------------------------------------------------
'- AE=>6 y <>80, Rec.Mov. (PUDI)
'- Imp.Rec. ==a==>> Imp.Rec.Adm. (porque puede tener Dto's) ------------------------------------------------------------
    Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(LoBDatos, BD_ImpRec, BD_Rec_Imp_Adm, BD_ActivEco, ">6", BD_ActivEco, "<>80")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Rec. ==a==>>Imp.Rec.Adm., a Todos los Rec.Mov. menos AE=80, Recibos de Movimiento (PUDI)", 0)
    
    
    
'''        Call Rut_Lo_Filtros_Quitar(LoBDatos)
'''        Call Rut_Lo_Sort(LoBDatos, BD_Plan, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
'''        LoBDatos.ShowTotals = False
'''    With LoBDatos
'''        .Range.AutoFilter Field:=BD_Plan, Criteria1:="=M013B"
'''        RowsFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
'''        If RowsFind > 0 Then
'''            Dim rngSrc As Range
'''            Dim rngDest As Range
'''            Set rngSrc = LoBDatos.ListColumns(BD_ImpAdm).DataBodyRange
'''            Set rngSrc = rngSrc.SpecialCells(xlCellTypeVisible)
'''            Set rngDest = LoBDatos.ListColumns(BD_Rec_Imp_Adm).DataBodyRange
'''            Set rngDest = rngDest.SpecialCells(xlCellTypeVisible)
'''            rngSrc.Copy
'''            rngDest.PasteSpecial xlPasteValues
'''            Application.CutCopyMode = False
'''        End If
'''    End With
    
'- ------------------------------------------------------------------------------------------------------------------
'- Calcular de Recibos con Imp.Adm. < 0 -->> Imp.Adm. = 0 ------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(LoBDatos)
        LoBDatos.ShowTotals = False
    With LoBDatos
        Call Rut_Lo_Sort(LoBDatos, BD_Rec_Imp_Adm, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_ImpAdm, Criteria1:="<0"
        .Range.AutoFilter Field:=BD_Anul, Criteria1:="=S"
        RowsFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RowsFind > 0 Then
            ImpTAdmNeg = Application.Sum(.DataBodyRange.Columns(BD_ImpAdm).SpecialCells(xlCellTypeVisible))
            '.DataBodyRange.Columns(BD_Rec_Imp_Adm).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No se tendrán en cuenta los Rec. Anulados, con Imp.Adm. <0 ", 0, _
                                        Format(RowsFind, "#,##0") & " reg.", Format(ImpTAdmNeg, "#,##0.00€"), , , , 2)
    End With
    
    '- Visualizo el progreso ----------------------------------------------------------------------------------------
    With LoBDatos
        .AutoFilter.ShowAllData            ' Elimina los filtros
        '- Visualizo el progreso ----------------------------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Resultado de Identificar Importes del AñoCont_" & AñoCont & _
                             ", en: " & Format(TF_LoBDatos, "#,##0") & "reg.", TimeLapSub, , , , , , 2)
        '- Sumatorios ---------
        With .DataBodyRange
            ImpTAcad = Application.Sum(.Columns(BD_Rec_Imp_Acad))
            ImpTAdm = Application.Sum(.Columns(BD_Rec_Imp_Adm))
            ImpTDto = Application.Sum(.Columns(BD_Rec_Imp_Dto))
            CantImpAcad = Application.Count(.Columns(BD_Rec_Imp_Acad))
            CantImpAdm = Application.Count(.Columns(BD_Rec_Imp_Adm))
            CantImpDto = Application.Count(.Columns(BD_Rec_Imp_Dto))
        End With
    End With
    
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- PPub Acad. de Matrícula por un importe de:    ", 0, _
                            Format(ImpTAcad, "#,##0.00€"), "en " & Format(CantImpAcad, "#,##0 reg."))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- Descuentos de Matrícula por un importe de:    ", 0, _
                            Format(ImpTDto, "#,##0.00€"), "en " & Format(CantImpDto, "#,##0 reg."))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- PPub Adm.  de Matrícula por un importe de:    ", 0, _
                            Format(ImpTAdm, "#,##0.00€"), "en " & Format(CantImpAdm, "#,##0 reg."))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- Importes  Acad. + Adm. - Dto.  Totalizado:    ", 0, _
                            Format(ImpTAcad + ImpTAdm + ImpTDto, "#,##0.00€"))
        
Call Rut_Lo_Filtros_Quitar(LoBDatos)
Debug.Print "<<< Rut_Assign_Imp_AdmAcad_C_Acad_Pos" & C_Acad
End Sub     ' -------------------------------------------------------------------------------------------------------------------------<<<
' ========================================================================================================================================

Sub kk()
'- ------------------------------------------------------------------------------------------------------------------
'- Asignar en Recibos Plan="M013": Imp.Rec.Adm. = Imp.Adm. ----------------------------------------------------------
    Dim Fila        As Long
    Dim PlanDNI_Ant As String
    Dim PlanDNI_New As String
    Dim RowData     As ListRow
    Dim RowDataCab  As ListRow
    Dim NumRecAnt   As Integer
    Dim ImpADMAnt   As Currency
    Dim LoBDatos    As ListObject
    Set LoBDatos = Sht__BD_Pruebas_VBA.ListObjects(1)
    Sht__BD_Pruebas_VBA.Select
        Call Rut_Lo_Filtros_Quitar(LoBDatos)
        LoBDatos.ShowTotals = False
    
    Call Rut_Lo_Filtros_Quitar(LoBDatos)
    LoBDatos.DataBodyRange.Columns(BD_Matricula).ClearContents
    LoBDatos.DataBodyRange.Columns(BD_Anul).ClearContents
    LoBDatos.DataBodyRange.Columns(BD_RegMov).ClearContents
    ' Ordenar por columnas  ------------------------------
        Call Rut_Lo_Sort(LoBDatos, BD_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_ActivEco, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_Plan, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_DNI, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_NumRec, xlDescending, False)    '- Ordenar primero accelera un montón el borrado -----
    
    '- Recorro toda la tabla LoBDatos -----------------------------------------------------------------
    Set RowDataCab = LoBDatos.ListRows(1)
    ImpADMAnt = RowDataCab.Range(BD_ImpAdm)
    For Fila = 1 To LoBDatos.ListRows.Count  '--- Bucle para recorrer todas la filas de LoBDatos
        Set RowData = LoBDatos.ListRows(Fila)
        If RowData.Range(BD_ActivEco) < 6 Then GoTo Sig_Reg
        If RowData.Range(BD_ActivEco) > 6 Then Fila = LoBDatos.ListRows.Count: GoTo Sig_Reg
        PlanDNI_New = RowData.Range(BD_Plan) & "_" & RowData.Range(BD_DNI)
        If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
           PlanDNI_Ant = PlanDNI_New
                RowData.Range(BD_Matricula) = RowData.Range(BD_NumRec)      '- Núm.Rec. Mayor
                RowDataCab.Range(BD_Anul) = NumRecAnt                   '- Núm.Rec. Menor
                Set RowDataCab = LoBDatos.ListRows(Fila)
                ImpADMAnt = RowDataCab.Range(BD_ImpAdm)
        End If
        If ImpADMAnt <> RowData.Range(BD_ImpAdm) Then RowDataCab.Range(BD_RegMov) = RowData.Range(BD_ImpAdm)
        NumRecAnt = RowData.Range(BD_NumRec)
Sig_Reg:
    Next
    
        LoBDatos.ShowTotals = True
    
 Application.Speech.Speak "Proceso completado."
        Call Rut_Lo_Filtros_Quitar(LoBDatos)
End Sub


Sub kk3kkl()
Dim Lo_ClsBk As ListObject: Set Lo_ClsBk = Sht__BD.ListObjects(1)
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
    ' Ordenar por columnas  ------------------------------
    Call Rut_Lo_Sort(Lo_ClsBk, BD_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, BD_ActivEco, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, BD_Plan, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, BD_DNI, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, BD_NumRec, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, BD_Ref, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----


End Sub


