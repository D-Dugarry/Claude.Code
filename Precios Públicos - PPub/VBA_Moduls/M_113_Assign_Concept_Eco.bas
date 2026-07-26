Attribute VB_Name = "M_113_Assign_Concept_Eco"
'Rev.: 2026-01-22
'                           ¡¡¡  OJO HE MIDIFICADO CONCEPTO ECO. por 1303.00 Y NO 1303 = 1030,00   !!!
Option Explicit

    '- Determinar Fecha de Vencimiento
    '- Determinar Cta-CCC Ingreso
    '- Determinar Concepto Económico
    '- Determinar Tipo de Enseñanza TIO-EP

            Sub RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio_ByHand()
                Sht__BD.Unprotect
                Sht__BD.Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
                Call RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio(Sht__BD.ListObjects(1))
'                Call RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio(Sht__BD.ListObjects(1), Prog_ClasifEco.ListObjects(1))
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
    '- Determinar Fecha de Vencimiento --------------------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio(Lo_Data As ListObject)
Debug.Print ">>> RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio"
    Dim TimeLapSub          As Single:      TimeLapSub = LastTimeLap
    Dim TF_Sin_Concepto     As Integer
    Dim TF_ActivNotFind     As Integer
    Dim AñoCont             As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim Cont_Fail           As Long
    Dim ContErrFVto         As Long
    Dim i                   As Long
    Dim F_BD                As Long:
    Dim TF_BD               As Long:        TF_BD = Lo_Data.ListRows.Count
    Dim rowfind             As Variant
    Dim TxT_Prog            As String

    Lo_Data.ShowTotals = False
        
    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '- Determinar Fecha de Vencimiento --------------------------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------------------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación Fecha de Vto.", 0, , , , , , 2)
    With Lo_Data
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_FVto, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .DataBodyRange.Columns(BD_ACont_Vto).ClearContents   '- Se supone que está vacía...
        
    '- Año_Vto = AñoCont -------------------------------------------------------------------------------------------------------
        '- Cuando F_Vto anterior al 1-Ene del AñoCont-1, Es decir que es un Rec. Añejo Pongo F_Vto = AñoCont
                .Range.AutoFilter Field:=BD_FVto, Criteria1:="<01/01/" & AñoCont - 1
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_ACont_Vto).SpecialCells(xlCellTypeVisible).Cells.Value = AñoCont
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Año_Vto = " & AñoCont & " ( F_Vto < " & AñoCont - 1 & ")", 0, Format(rowfind, "#,##0"))
        
    '- Año_Vto = AñoCont -------------------------------------------------------------------------------------------------------
        '- Cuando F_Vto corresponde al AñoCont, Es decir que es un Rec. Emitido o Aplazado, pongo Año_Vto = AñoCont
        .AutoFilter.ShowAllData            ' Elimina los filtros
                .Range.AutoFilter Field:=BD_FVto, Criteria1:=">=01/01/" & AñoCont, _
                                 Operator:=xlAnd, Criteria2:="<01/01/" & AñoCont + 1
                                 'Operator:=xlAnd, Criteria2:="<=12/31/" & AñoCont   ¡¡¡ OJO !!! así no funciona seguramente porque habría que poner algo como por ejemplo 12/31/2025 23:59:59
        '.Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_FVto_Act")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_ACont_Vto).SpecialCells(xlCellTypeVisible).Cells.Value = AñoCont
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Año_Vto = " & AñoCont & " ( F_Vto = " & AñoCont & ")", 0, Format(rowfind, "#,##0"))
    
    '- Año_Vto = AñoCont - 1 -------------------------------------------------------------------------------------------------------
        '- Cuando F_Vto corresponde al AñoCont-1, Es decir que es un Rec. EjeAnt, pongo Año_Vto = AñoCont - 1
        .AutoFilter.ShowAllData            ' Elimina los filtros
                .Range.AutoFilter Field:=BD_FVto, Criteria1:=">=01/01/" & AñoCont - 1, _
                                 Operator:=xlAnd, Criteria2:="<01/01/" & AñoCont
'        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_FVto_Ant")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_ACont_Vto).SpecialCells(xlCellTypeVisible).Cells.Value = AñoCont - 1
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Año_Vto = " & AñoCont - 1 & " ( F_Vto = " & AñoCont - 1 & ")", 0, Format(rowfind, "#,##0"))
        
    '- Año_Vto = AñoCont +1 -------------------------------------------------------------------------------------------------------
        '- Cuando F_Vto corresponde al AñoCont + 1, Es decir que es un Rec. ADxAplz, pongo Año_Vto = AñoCont + 1 ------------
        .AutoFilter.ShowAllData            ' Elimina los filtros
                .Range.AutoFilter Field:=BD_FVto, Criteria1:=">=01/01/" & AñoCont + 1
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_ACont_Vto).SpecialCells(xlCellTypeVisible).Cells.Value = AñoCont + 1
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Año_Vto = " & AñoCont + 1 & " ( F_Vto = " & AñoCont + 1 & ")", 0, Format(rowfind, "#,##0"))
        
        '.ShowAutoFilter = True          '- El AdvancedFilter con Rango de Criterio desactiva el "ShowFilterMarck"
        
    '- Detectar errores: AñoEmi>AñoVto
        For i = 1 To TF_BD
            With .DataBodyRange
                If .Cells(i, BD_ACont_Emi) > .Cells(i, BD_ACont_Vto) Then
                    .Cells(i, BD_ACont_Vto) = .Cells(i, BD_ACont_Emi)
                    ContErrFVto = ContErrFVto + 1
                End If
            End With
        Next i
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Errores de ACont_Vto < ACont_Emi, Cambiado Fecha: ACont_Vto = ACont_Emi", 0, , Format(ContErrFVto, "#,##0"))
    
        '- Sumatorios Revisión AñoVto
            Dim CantAConAnt As Long
            Dim CantACon As Long
            Dim CantAConPos As Long
            With .DataBodyRange
                CantAConAnt = Application.CountIfs(.Columns(BD_ACont_Vto), AñoCont - 1)
                CantACon = Application.CountIfs(.Columns(BD_ACont_Vto), AñoCont)
                CantAConPos = Application.CountIfs(.Columns(BD_ACont_Vto), AñoCont + 1)
'                ImpTAdm = Application.SumIfs(.Columns(BD_ImpAdm), .Columns(BD_Obs_Conta), "RecCab")
'                ImpTDto = Application.Sum(.Columns(BD_ImpDto))
            End With
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificado Rec. AñoVto " & AñoCont - 1, 0, _
                                                        Format(CantAConAnt, "#,##0"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificado Rec. AñoVto " & AñoCont, 0, _
                                                        Format(CantACon, "#,##0"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificado Rec. AñoVto " & AñoCont + 1, 0, _
                                                        Format(CantAConPos, "#,##"))
        
        
        '- Visualizo el progreso --------
        Cont_Fail = Application.CountIf(Lo_Data.DataBodyRange.Columns(BD_ACont_Vto), "")
        
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Rec. SIN F_Vto. Asignada.", 0, Format(Cont_Fail, "#,##0"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Rec. en BDatos", 0, Format(.ListRows.Count, "#,##0"))
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación Fecha de Vto.", TimeLapSub)
    
    End With    '-  Lo_Data

'GoTo Restablecer_Valores

    '''    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '''    '- ------------------------ FUNCIONA PERO LO DESACTIVO POR NO NECESITARLO AHORA -----------------------------------------------------------
    '''    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '''    '- ---------------------------------- SI HICIERA FALTA REACTIVARLO, HABRÍA QUE COMPROBAR SU EFICACIA. -------------------------------------
    '''    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '''    '- Determinar Cta-CCC Ingreso -------------------------------------------------------------------------------------------------------------
    '''    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '''    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación de Cta-CCC de Ingreso.", 0, , , , , , 2)
    '''    With Lo_Data
    '''        Call Rut_Lo_Filtros_Quitar(Lo_Data)
    '''        .DataBodyRange.Columns(BD_Cta_Ing).ClearContents   '- Se supone que está vacía...
    '''
    '''        '-Copy Col BD_CtaPag en Col BD_Cta_Ing ---------------------------------------------------------------------------------
    '''        .ListColumns(BD_CtaPag).DataBodyRange.Copy
    '''        .ListColumns(BD_Cta_Ing).DataBodyRange.PasteSpecial Paste:=xlPasteValues
    '''
    '''
    '''        '-Filtra Recibos "Imp_Rec <=0"  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData            ' Elimina los filtros
    '''        Call Rut_Lo_Sort(Lo_Data, BD_ImpRec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    '''        .Range.AutoFilter Field:=BD_ImpRec, Criteria1:="<0"
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "Imp_Rec <0"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "Imp_Rec <=0", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos "No Cobrado"  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        Call Rut_Lo_Sort(Lo_Data, BD_ACont_Cob, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    '''        .Range.AutoFilter Field:=BD_ACont_Cob, Criteria1:="="
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "No Cobrado"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "No Cobrado", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos BD_CtaPag = "FLY WIRE    "  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        Call Rut_Lo_Sort(Lo_Data, BD_CtaPag, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    '''        .Range.AutoFilter Field:=BD_CtaPag, Criteria1:="=FLY WIRE*"
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416175503"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " Cta_CCC " & "FLY WIRE", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos BD_CtaPag Inf-Regularizado= "FLY"  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        Call Rut_Lo_Sort(Lo_Data, BD_InfRegulariz, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    '''        .Range.AutoFilter Field:=BD_InfRegulariz, Criteria1:="=FLY*"
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416175503"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
    '''                                            " Cta_CCC " & "FLY Regularizado", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos BD_CtaPag Inf-Regularizado= "0049 "  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        .Range.AutoFilter Field:=BD_InfRegulariz, Criteria1:="=0049 "
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416175503"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
    '''                                            " Cta_CCC " & "Regularizado G.Acad", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos BD_CtaPag Inf-Regularizado= "6659072416125620 "  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        .Range.AutoFilter Field:=BD_InfRegulariz, Criteria1:="=6659072416125620*"
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416125620"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
    '''                                            " Cta_CCC " & "Regularizado G.Acad ???", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos BD_CtaPag Inf-Regularizado= "(0049)"  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        .Range.AutoFilter Field:=BD_InfRegulariz, Criteria1:="=*(0049)"
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "0049 6659 07 2416125620"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
    '''                                            " Cta_CCC " & "Regularizado S.Inf.", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos BD_CtaPag Inf-Regularizado= "(2100)"  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        .Range.AutoFilter Field:=BD_InfRegulariz, Criteria1:="=*(2100)"
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "2100 8984 16 0200003529"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
    '''                                            " Cta_CCC " & "Regularizado S.Inf.", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos BD_CtaPag Inf-Regularizado= "(0081)"  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        .Range.AutoFilter Field:=BD_InfRegulariz, Criteria1:="=*(0081)"
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "0081 3191 42 0001068211"
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
    '''                                            " Cta_CCC " & "Regularizado S.Inf.", 0)
    '''        End If
    '''
    '''        '-Filtra Recibos BD_CtaPag Inf-Regularizado= "(0014)"  ---------------------------------------------------------------------------------
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''        .Range.AutoFilter Field:=BD_InfRegulariz, Criteria1:="=*(0014)"
    '''        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    '''        If RowFind > 0 Then
    '''            .DataBodyRange.Columns(BD_Cta_Ing).SpecialCells(xlCellTypeVisible).Cells.Value = "9000 0005 00 0260000014)"   '- Bco.Esp.
    '''            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
    '''                                            " Cta_CCC " & "Regularizado S.Inf.", 0)
    '''        End If
    '''
    '''        .AutoFilter.ShowAllData         ' Elimina los filtros
    '''
    '''        '- Visualizo el progreso --------
    '''        Cont_Fail = Application.CountIf(Lo_Data.DataBodyRange.Columns(BD_Cta_Ing), "")
    '''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(Cont_Fail, "#,##0"), 8) & _
    '''                                        " Rec. SIN Cta-CCC de Ingreso Asignados.", 0)
    '''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & "x.xxx", 8) & _
    '''                                        " Rec. CON Cta-CCC de Ingreso Asignados previamente.", 0)
    '''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "=") & Format(.ListRows.Count, "#,##0"), 8) & " Rec. en BDatos", 0)
    '''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación de Cta-CCC de Ingreso.", TimeLapSub)
    '''
    '''    End With    '- Lo_Data.
    
    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '- Determinar Concepto Económico y Tipo de Enseñanza TIO-EP -------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------------------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación Concepto Económico y Tipo Ensañanza.", 0, , , , , , 2)
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_Lo_Sort(Lo_Data, BD_ActivEco, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_Data, BD_TipoCurso, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_Data, BD_Plan, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    
    With Lo_Data
        .DataBodyRange.Columns(BD_Concepto).ClearContents   '- Se supone que está vacía...
        .DataBodyRange.Columns(BD_TIO_EP).ClearContents   '- Se supone que está vacía...
        
    Call Rut_Lo_Sort(Lo_Data, BD_ActivEco, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_Data, BD_TipoCurso, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        
'-1º Recibos Cod_Activ = 6 - Grado -----------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_AE6")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1310.00"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "Grado"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1310.00'    Reg. AE6 Grado", 0)
        
'-2º Recibos Cod_Activ = 5 - Master -----------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AutoFilter Field:=BD_ActivEco, Criteria1:=5
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1310.01"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "Master"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1310.01'    Reg. AE5 Master", 0)
        
'-3º Recibos Cod_Activ = 2 - Doctorado -----------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AutoFilter Field:=BD_ActivEco, Criteria1:=2
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1310.02"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "Doctorado"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1310.02'    Reg. AE2 Doctorado", 0)
        
'-4º Recibos - EFP - Estudios de Formación Permanente: Máster, Especialista, Experto. -------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_EFP")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1311.00"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "EFP"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                " '1311.00'    Reg. AE4 EFP: Estudios de Formación Permanente: Master, Especialista y Experto.", 0)

'-5º Recibos - CFC - Cursos de Formación Contínua -------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_CFC")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1311.03"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "CFC"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                " '1311.03'    Reg. AE4 CFC: Cursos de Formación Contínua", 0)

'-5º Recibos - AFC - Actividades de Formación Complementaria -------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_AFC")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1311.03"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "AFC"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                " '1311.03'    Reg. AE4 AFC: Actividades de Formación Complementaria", 0)

'-6º Recibos Cod_Activ = 80 - Pruebas de aptitud para acceso a la Universidad -----------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AutoFilter Field:=BD_ActivEco, Criteria1:=80
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1315.00"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "PruebasAccesoUni"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1315.00'    Reg. AE80 Pruebas Acceso Univ.", 0)
        
'-7º Recibos - TNCT-M013 AE300 - Cursos NO Contabilizables como EFP (M013) sino como Rec.Mov. por Secretaría de Acceso ---------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_TNCT_M013")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1312.00"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "TNCT_M013"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " '1312.00'    Reg. AE300 TNCT_M013: Seminario Orientación Pruebas > 25 años", 0)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(27, " ") & "(Secretaría de Acceso AE4/300/710)", 0)
        
'-8º _UPUA_ Recibos - CFC-TUP - _UPUA_ Universidad Permanente, Cursos de Formación Contínua (TUP) -------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_TUP")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1312.02"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "CFC_UPUA"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                " '1312.02'    Reg. AE4 CFC_UPUA (TUP): Programa Univ. para Mayores UA. (Univ. Permanente)", 0)

'-9º Recibos - TNCT-PNB1 - Cursos NO Contabilizables como Títulos Propios Universidad (PNB1) -------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_TNCT_PNB1")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1303.01"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "TNCT_PNB1"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " '1303.01'    Reg. AE4 TNCT_PNB1: Prueba de competencias idioma extranjero.", 0)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(27, " ") & "(Centro Sup. Idiomas AE4/21)", 0)
        
'-10º Recibos de Movimiento - Rec_Adm - Recibos de Actividad Administrativa, SIN relación con Matrícula Académica ---
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_RecAdm")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1303.00"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "Rec_Adm"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                            " '1303.00'    Reg. Rec_Adm: Recibos de una actividad púramente Administrativa.", 0)

'-11º Recibos EURLE PLAN = C404 (Escuela Univ. Relaciones Laborales -ELDA-) ----------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_Plan, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_EURLE")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "EURLElda"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "EURLElda"
            .DataBodyRange.Columns(BD_Obs_Conta).SpecialCells(xlCellTypeVisible).Cells.Value = "EURLElda"
            '.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                            " 'EURLElda'   Reg. AE6-Plan_C404 EURLElda: Esc. Univ. Relaciones Laborales Elda", 0)

'- Recibos de Matrículas de coste CERO - ImpMatCero - Recibos Matrícula de Actividad Académica a Coste CERO. ------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_ImpMatCero")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "NoContab"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "ImpMatCero"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                            " 'No Contab.' Reg. ImpMatCero: Matrícula de Actividad Académica a Coste CERO.", 0)

'- Recibos Sin Concepto o Tipo ---------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AutoFilter Field:=BD_TIO_EP, Criteria1:="="
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "NoContab"
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "NoContab"
            MsgBox "¡¡¡ Recibos SIN identificar Concepto-Eco o Tipo  !!!" & vbLf & vbLf & Format(rowfind, "#,##0") & _
                            " reg.", vbOKOnly + vbExclamation, "Proceso: Asignar Concepto Eco. y Tipo de Enseñanza"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                            " 'No Contab.' Reg. Recibos SIN identificar Concepto-Eco o Tipo.", 0)
        .ShowAutoFilter = True          '- El AdvancedFilter con Rango de Criterio desactiva el "ShowFilterMarck"
    
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "=") & Format(.ListRows.Count, "#,##0"), 8) & " Reg. en BDatos", 0)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso Finalizado: Asignación Concepto Económico y Tipo Ensañanza.", TimeLapSub)
    
    End With    '-  Lo_Data

    Lo_Data.ShowTotals = True
        
Restablecer_Valores:
Call Rut_Lo_Filtros_Quitar(Lo_Data)
Debug.Print "<<< RuT_Assign_AñoVto_CtaCCC_ConcepEco_y_TipoEstudio"
End Sub

