Attribute VB_Name = "M_413_Assign_ImpAdm_AE4x1"
'Rev.: 2026-01-26
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Identificar del C_Acad_Pos, los 1º Rec. de c/matrícula para obtener la T-Adm ---------------------------------------------
'-      ClearContents de BD_Rec_Imp_Adm, de Recibos con Imp.Adm. < 0 -->> Son ajustes de matrículas
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Assign_Imp_AdmAcad_AE4x1(Lo_ClsBk_AE4x1 As ListObject, _
                                 TipoEP As String)

Debug.Print ">>> Rut_Assign_Imp_AdmAcad_C_Acad_Pos"
    Dim CantImpAcad     As Long
    Dim CantImpAdm      As Long
    Dim CantImpAdmAntig As Long
    Dim CantImpDto      As Long
    Dim ImpTAcad        As Currency
    Dim ImpTAdm         As Currency
    Dim ImpTAdmNeg      As Currency
    Dim ImpTDto         As Currency
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    Dim Fila            As Long
    Dim TF_Lo_ClsBk_AE4x1     As Long:        TF_Lo_ClsBk_AE4x1 = Lo_ClsBk_AE4x1.ListRows.Count
    Dim RowsDel         As Long
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad          As String:      C_Acad = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim PlanDNI_Ant     As String:      PlanDNI_Ant = ""
    Dim PlanDNI_New     As String:
    Dim TxT_ProgIni     As String
    Dim TxT_Progreso    As String
    Dim RowData         As ListRow
    Dim RowsFind        As Variant
    Dim Sh_Data         As Worksheet:   Set Sh_Data = Lo_ClsBk_AE4x1.Parent

    '- Visualizo el progreso ----------------------------------------------------------------------------------------
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Identificar ImpAcad ImpAdm e ImpDto del Curso: " & _
                         C_Acad & ", en " & Format(Lo_ClsBk_AE4x1.ListRows.Count, "#,##0") & " reg.", 0, , , , , , 2)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_AE4x1)
    ' Ordenar por columnas  ------------------------------
        Call Rut_Lo_Sort(Lo_ClsBk_AE4x1, BD_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_ClsBk_AE4x1, BD_ActivEco, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_ClsBk_AE4x1, BD_Plan, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_ClsBk_AE4x1, BD_DNI, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_ClsBk_AE4x1, BD_NumRec, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_ClsBk_AE4x1, BD_Ref, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    
'- Asignar Imp.Adm. que no estén en el resto de la tabla Lo_ClsBk_AE4x1 -----------------------------------------------------------------
    TxT_Progreso = ActivForm.Controls("TBx_Informe")
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_AE4x1)
    For Fila = 1 To TF_Lo_ClsBk_AE4x1   '--- Bucle para recorrer todas la filas de Lo_ClsBk_AE4x1
        Set RowData = Lo_ClsBk_AE4x1.ListRows(Fila)
        If RowData.Range(BD_Plan) = "M013" Then GoTo Sig_Reg            '- Tratamiento expecial...
        If RowData.Range(BD_Plan) = "PNB1" Then GoTo Sig_Reg            '- Tratamiento expecial...
        If RowData.Range(BD_Plan) = "UPUA" Then GoTo Sig_Reg            '- Tratamiento expecial...
        PlanDNI_New = RowData.Range(BD_Plan) & "_" & RowData.Range(BD_DNI)
        If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
           PlanDNI_Ant = PlanDNI_New
                    If RowData.Range(BD_Anul) <> "S" Then
                            RowData.Range(BD_Rec_Imp_Acad) = RowData.Range(BD_ImpAcad)
                            RowData.Range(BD_Rec_Imp_Dto) = RowData.Range(BD_ImpDto)
                        If RowData.Range(BD_Rec_Imp_Adm) <> "" Then                     '- Respeta si en la APP de AE4 ya se introdujo un ImpAdm manual: NO se sobreescribe.
                            CantImpAdmAntig = CantImpAdmAntig + 1
                        Else
                            RowData.Range(BD_Rec_Imp_Adm) = RowData.Range(BD_ImpAdm)
                            CantImpAdm = CantImpAdm + 1
                        End If
                End If
        End If
Sig_Reg:
        If Fila Mod 4000 = 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Indentificados " & CantImpAdm & " new, + " & CantImpAdmAntig & " antiguos, de AE_4 y de " & AñoCont & ", con Imp.Adm. en: ", 0, _
                                Format(Fila, "#,##0") & "reg.", "de " & Format(TF_Lo_ClsBk_AE4x1, "#,##0") & "reg.", TxT_Progreso, True, , 2)
        End If
Fin_Bucle:
    Next
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Adm. a " & CantImpAdm & " reg., + " & CantImpAdmAntig & " antiguos, de AE_4 y de " & AñoCont & _
                                            ", de un total de: ", 0, Format(TF_Lo_ClsBk_AE4x1, "#,##0") & "reg.", , TxT_Progreso, True, , 2)
    
'- ------------------------------------------------------------------------------------------------------------------
'- Plan="M013", Seminario Mayores 25a
'- Imp.Adm. ==a==>>Imp.Rec.Adm. ---------------------------------------------------------------------------------------
    Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(Lo_ClsBk_AE4x1, BD_ImpAcad, BD_Rec_Imp_Acad, BD_Plan, "=M013B")
    Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(Lo_ClsBk_AE4x1, BD_ImpAdm, BD_Rec_Imp_Adm, BD_Plan, "=M013B")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Adm. ==a==>>Imp.Rec.Adm., a Todos los Rec. con Plan=M013B, Seminario Mayores 25a", 0)
    
'- ------------------------------------------------------------------------------------------------------------------
'- Plan="UPUA", Cursos universidad permanente
'- Imp.Adm. ==a==>>Imp.Rec.Adm. ---------------------------------------------------------------------------------------
    Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(Lo_ClsBk_AE4x1, BD_ImpRec, BD_Rec_Imp_Acad, BD_Plan, "=UPUA")
    Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(Lo_ClsBk_AE4x1, BD_ImpAdm, BD_Rec_Imp_Adm, BD_Plan, "=UPUA")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Adm. ==a==>>Imp.Rec.Adm., a Todos los Rec. con Plan=UPUA, Cursos universidad permanente", 0)
    
'- ------------------------------------------------------------------------------------------------------------------
'- Plan="PNB1", Pruebas Competencias Idiomas
'- Imp.Rec. ==a==>>Imp.Rec.Adm. (porque puede tener Dto's) ------------------------------------------------------------
    Call Rut_Lo_DataBodyRange_Filter_x2Crit_Copy_ColSource_to_ColTarget(Lo_ClsBk_AE4x1, BD_ImpRec, BD_Rec_Imp_Adm, BD_Plan, "=PNB1")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Rec. ==a==>>Imp.Rec.Adm., a Todos los Rec. con Plan=PNB1, Pruebas Competencias Idiomas", 0)
    
'- ------------------------------------------------------------------------------------------------------------------
'- Calcular de Recibos con Imp.Adm. < 0 -->> Imp.Adm. = 0 ------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_AE4x1)
        Lo_ClsBk_AE4x1.ShowTotals = False
    With Lo_ClsBk_AE4x1
        Call Rut_Lo_Sort(Lo_ClsBk_AE4x1, BD_Rec_Imp_Adm, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
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
    
'    '- Visualizo el progreso ----------------------------------------------------------------------------------------
'    With Lo_ClsBk_AE4x1
'        Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_AE4x1)
'        '- Visualizo el progreso ----------------------------------------------------------------------------------------
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Resultado de Identificar Importes del AñoCont_" & AñoCont & _
'                             ", en: " & Format(TF_Lo_ClsBk_AE4x1, "#,##0") & "reg.", TimeLapSub, , , , , , 2)
'        '- Sumatorios ---------
'        With .DataBodyRange
'            ImpTAcad = Application.Sum(.Columns(BD_Rec_Imp_Acad))
'            ImpTAdm = Application.Sum(.Columns(BD_Rec_Imp_Adm))
'            ImpTDto = Application.Sum(.Columns(BD_Rec_Imp_Dto))
'            CantImpAcad = Application.Count(.Columns(BD_Rec_Imp_Acad))
'            CantImpAdm = Application.Count(.Columns(BD_Rec_Imp_Adm))
'            CantImpDto = Application.Count(.Columns(BD_Rec_Imp_Dto))
'        End With
'    End With
'
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- PPub Acad. de Matrícula por un importe de:    ", 0, _
'                            Format(ImpTAcad, "#,##0.00€"), "en " & Format(CantImpAcad, "#,##0 reg."))
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- Descuentos de Matrícula por un importe de:    ", 0, _
'                            Format(ImpTDto, "#,##0.00€"), "en " & Format(CantImpDto, "#,##0 reg."))
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- PPub Adm.  de Matrícula por un importe de:    ", 0, _
'                            Format(ImpTAdm, "#,##0.00€"), "en " & Format(CantImpAdm, "#,##0 reg."))
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- Importes  Acad. + Adm. - Dto.  Totalizado:    ", 0, _
'                            Format(ImpTAcad + ImpTAdm + ImpTDto, "#,##0.00€"))
        
Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_AE4x1)
Debug.Print "<<< Rut_Assign_Imp_AdmAcad_C_Acad_Pos" & C_Acad
End Sub     ' -------------------------------------------------------------------------------------------------------------------------<<<
' ========================================================================================================================================





