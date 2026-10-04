Attribute VB_Name = "M_115_Assign_ImpAdm_CAcad"
' Last Rev. 2026-10-04 17:08
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
    Dim AnoCont         As String:      AnoCont = Prog__APP.Range("APP_AnoCont")
    '- (2026-10-04) Curso Pos, como dice el nombre de la rutina. Antes leía el selector APP_CursAcad: con el selector en el
    '- curso Ant calculaba el curso que M_215 borra justo después, y el curso Pos se quedaba sin Rec_Imp_*.
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim PlanDNI_Ant     As String:      PlanDNI_Ant = ""
    Dim PlanDNI_New     As String:
    Dim TxT_ProgIni     As String
    Dim TxT_Progreso    As String
    Dim RowsFind        As Variant
    Dim Sh_Data         As Worksheet:   Set Sh_Data = LoBDatos.Parent

    '- Visualizo el progreso ----------------------------------------------------------------------------------------
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Identificar ImpAcad ImpAdm e ImpDto del Curso: " & _
                         C_Acad_Pos & ", en " & Format(LoBDatos.ListRows.Count, "#,##0") & " reg.", 0, , , , , , 2)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    
    '-ClearContents de Rec. BD_C_Acad = C_Acad_Pos -------------------------------------
    Call Rut_Lo_Filtros_Quitar(LoBDatos)
    With LoBDatos
        .Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & C_Acad_Pos
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
        Call Rut_Lo_Sort(LoBDatos, BD_C_Acad, xlAscending, True, Aplicar:=False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_ActivEco, xlAscending, False, Aplicar:=False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_Plan, xlAscending, False, Aplicar:=False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_DNI, xlAscending, False, Aplicar:=False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_NumRec, xlAscending, False, Aplicar:=False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(LoBDatos, BD_Ref, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    
    '- Recorro toda la tabla LoBDatos, en RAM (Rut_Lo_TablaRam): celda a celda con ListRows(Fila) tardaba ~1 min. --------
    '- Se carga con .Value2: las Col. Rec_Imp_* se devuelven enteras a la hoja y así las filas que no se tocan quedan idénticas.
    Dim T_BD            As T_TablaRam
    Call Rut_TablaRam_Cargar(T_BD, LoBDatos, Array(BD_C_Acad, BD_ActivEco, BD_Plan, BD_DNI, BD_Anul, _
                                                  BD_ImpAcad, BD_ImpAdm, BD_ImpDto, _
                                                  BD_Rec_Imp_Acad, BD_Rec_Imp_Adm, BD_Rec_Imp_Dto), True)
    For Fila = 1 To TF_LoBDatos   '--- Bucle para recorrer todas la filas de LoBDatos
        If T_BD.Datos(Fila, BD_C_Acad) <> C_Acad_Pos Then GoTo Sig_Reg         '- NO tenemos en cuenta loas Recibos de Otros C_Acad, porque NO tenemos toda la información
        If T_BD.Datos(Fila, BD_ActivEco) = 4 Then GoTo Sig_Reg             '- NO tenemos en cuenta los Recibos de AE4,  ya vienen con sus ImpAdm de AE4x4
        If T_BD.Datos(Fila, BD_ActivEco) > 6 Then Exit For                 '- NO tenemos en cuenta los Recibos de Movimiento (ordenada por AE, ya no quedan más)
        PlanDNI_New = T_BD.Datos(Fila, BD_Plan) & "_" & T_BD.Datos(Fila, BD_DNI)
        If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
           PlanDNI_Ant = PlanDNI_New
                    If T_BD.Datos(Fila, BD_Anul) <> "S" Then
                            T_BD.Datos(Fila, BD_Rec_Imp_Acad) = T_BD.Datos(Fila, BD_ImpAcad)
                            T_BD.Datos(Fila, BD_Rec_Imp_Adm) = T_BD.Datos(Fila, BD_ImpAdm)
                            T_BD.Datos(Fila, BD_Rec_Imp_Dto) = T_BD.Datos(Fila, BD_ImpDto)
                            CantImpAdm = CantImpAdm + 1
                End If
        End If
Sig_Reg:
        If Fila Mod 4000 = 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Indentificados " & CantImpAdm & " Rec. de AE_2/5y6 y de " & AnoCont & ", con Imp.Adm. en: ", 0, _
                                Format(Fila, "#,##0") & "reg.", "de " & Format(TF_LoBDatos, "#,##0") & "reg.", TxT_Progreso, True, , 2)
        End If
    Next
    T_BD.Modificada(BD_Rec_Imp_Acad) = True
    T_BD.Modificada(BD_Rec_Imp_Adm) = True
    T_BD.Modificada(BD_Rec_Imp_Dto) = True
    Call Rut_TablaRam_Volcar(T_BD, LoBDatos)                            '- Devuelvo a la hoja las 3 Col. Rec_Imp_*
    Erase T_BD.Datos                                                    '- Libero la RAM
    
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Asignado Imp.Adm. a " & CantImpAdm & " Rec. de AE_2/5y6 y de " & AnoCont & ", de un total de: ", 0, _
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
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Resultado de Identificar Importes del AñoCont_" & AnoCont & _
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
Debug.Print "<<< Rut_Assign_Imp_AdmAcad_C_Acad_Pos" & C_Acad_Pos
End Sub     ' -------------------------------------------------------------------------------------------------------------------------<<<
' ========================================================================================================================================



