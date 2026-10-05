Attribute VB_Name = "M_113_Assign_Concept_Eco"
' Last Rev. 2026-10-05 23:11
'Rev.: 2026-01-22
'                           ¡¡¡  OJO HE MIDIFICADO CONCEPTO ECO. por 1303.00 Y NO 1303 = 1030,00   !!!
Option Explicit

    '- Determinar Fecha de Vencimiento
    '- Determinar Cta-CCC Ingreso
    '- Determinar Concepto Económico
    '- Determinar Tipo de Enseñanza TIO-EP

            Sub RuT_Assign_AnoVto_CtaCCC_ConcepEco_y_TipoEstudio_ByHand()
                Sht__BD.Unprotect
                Sht__BD.Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
                Call RuT_Assign_AnoVto_CtaCCC_ConcepEco_y_TipoEstudio(Sht__BD.ListObjects(1))
'                Call RuT_Assign_AnoVto_CtaCCC_ConcepEco_y_TipoEstudio(Sht__BD.ListObjects(1), Prog_ClasifEco.ListObjects(1))
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
    '- Determinar Fecha de Vencimiento --------------------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Assign_AnoVto_CtaCCC_ConcepEco_y_TipoEstudio(Lo_Data As ListObject)
'- Desde el 2026-10-05 trabaja en RAM (fase 3 del paso a RAM): los filtros se evalúan sobre la copia en memoria de la
'- tabla (Rut_Lo_CriT_Ram, con las mismas reglas que los filtros de Excel) y se asigna en RAM. Después, en la hoja: se
'- devuelven ACont_Vto, Concepto, TIO_EP y Obs_Conta, y UNA ordenación con el mismo resultado que las cuatro de antes
'- (M_115 da el importe al primer recibo de cada matrícula, así que el orden importa).
'- Antes se filtraba y escribía en la hoja en cada paso, con un bucle celda a celda para ACont_Vto (14 s).
Debug.Print ">>> RuT_Assign_AnoVto_CtaCCC_ConcepEco_y_TipoEstudio"
    Dim TimeLapSub          As Single:      TimeLapSub = LastTimeLap
    Dim AnoCont             As Integer:     AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim Cont_Fail           As Long
    Dim ContErrFVto         As Long
    Dim rowfind             As Long
    Dim Fila                As Long
    Dim T                   As T_TablaRam
    Dim Cumple()            As Boolean
    Dim Ene_Ant             As Double:      Ene_Ant = CDbl(DateSerial(AnoCont - 1, 1, 1))  '- 1 de enero de AnoCont-1,
    Dim Ene_Act             As Double:      Ene_Act = CDbl(DateSerial(AnoCont, 1, 1))      '-   de AnoCont
    Dim Ene_Pos             As Double:      Ene_Pos = CDbl(DateSerial(AnoCont + 1, 1, 1))  '-   y de AnoCont+1

    Lo_Data.ShowTotals = False
    If Lo_Data.DataBodyRange Is Nothing Then GoTo Restablecer_Valores
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_TablaRam_Cargar(T, Lo_Data, Array(BD_FVto, BD_ACont_Emi, BD_ACont_Vto, BD_ActivEco, BD_TipoCurso, BD_Plan, _
                             BD_Concepto, BD_TIO_EP, BD_Obs_Conta, BD_ImpRec, BD_ImpDto), True)

    '- ---------------------------------------------------------------------------------------------
    '- Determinar Fecha de Vencimiento -------------------------------------------------------------
    '- ---------------------------------------------------------------------------------------------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación Fecha de Vto.", 0, , , , , , 2)
    For Fila = 1 To T.NumFilas
        T.Datos(Fila, BD_ACont_Vto) = Empty                             '- Se supone que está vacía...
    Next Fila
    T.Modificada(BD_ACont_Vto) = True

    '- Año_Vto = AnoCont ---------------------------------------------------------------------------
        '- Cuando F_Vto anterior al 1-Ene del AnoCont-1, Es decir que es un Rec. Añejo Pongo F_Vto = AnoCont
        Cumple = Fnc_Filtro_Filas(T, BD_FVto, "<", Ene_Ant)
        rowfind = Fnc_Asignar(T, Cumple, BD_ACont_Vto, AnoCont)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Año_Vto = " & AnoCont & " ( F_Vto < " & AnoCont - 1 & ")", 0, Format(rowfind, "#,##0"))

    '- Año_Vto = AnoCont ---------------------------------------------------------------------------
        '- Cuando F_Vto corresponde al AnoCont, Es decir que es un Rec. Emitido o Aplazado, pongo Año_Vto = AnoCont
        Cumple = Fnc_Filtro_Filas(T, BD_FVto, "<", Ene_Pos, Fnc_Filtro_Filas(T, BD_FVto, ">=", Ene_Act))
        rowfind = Fnc_Asignar(T, Cumple, BD_ACont_Vto, AnoCont)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Año_Vto = " & AnoCont & " ( F_Vto = " & AnoCont & ")", 0, Format(rowfind, "#,##0"))

    '- Año_Vto = AnoCont - 1 -----------------------------------------------------------------------
        '- Cuando F_Vto corresponde al AnoCont-1, Es decir que es un Rec. EjeAnt, pongo Año_Vto = AnoCont - 1
        Cumple = Fnc_Filtro_Filas(T, BD_FVto, "<", Ene_Act, Fnc_Filtro_Filas(T, BD_FVto, ">=", Ene_Ant))
        rowfind = Fnc_Asignar(T, Cumple, BD_ACont_Vto, AnoCont - 1)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Año_Vto = " & AnoCont - 1 & " ( F_Vto = " & AnoCont - 1 & ")", 0, Format(rowfind, "#,##0"))

    '- Año_Vto = AnoCont +1 ------------------------------------------------------------------------
        '- Cuando F_Vto corresponde al AnoCont + 1, Es decir que es un Rec. ADxAplz, pongo Año_Vto = AnoCont + 1 ------------
        Cumple = Fnc_Filtro_Filas(T, BD_FVto, ">=", Ene_Pos)
        rowfind = Fnc_Asignar(T, Cumple, BD_ACont_Vto, AnoCont + 1)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Año_Vto = " & AnoCont + 1 & " ( F_Vto = " & AnoCont + 1 & ")", 0, Format(rowfind, "#,##0"))

    '- Detectar errores: AnoEmi>AñoVto
        For Fila = 1 To T.NumFilas
            If T.Datos(Fila, BD_ACont_Emi) > T.Datos(Fila, BD_ACont_Vto) Then
                T.Datos(Fila, BD_ACont_Vto) = T.Datos(Fila, BD_ACont_Emi)
                ContErrFVto = ContErrFVto + 1
            End If
        Next Fila
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Errores de ACont_Vto < ACont_Emi, Cambiado Fecha: ACont_Vto = ACont_Emi", 0, , Format(ContErrFVto, "#,##0"))

        '- Sumatorios Revisión AñoVto
            Dim CantAConAnt As Long
            Dim CantACon As Long
            Dim CantAConPos As Long
            CantAConAnt = Fnc_Contar_Igual(T, BD_ACont_Vto, AnoCont - 1)
            CantACon = Fnc_Contar_Igual(T, BD_ACont_Vto, AnoCont)
            CantAConPos = Fnc_Contar_Igual(T, BD_ACont_Vto, AnoCont + 1)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificado Rec. AñoVto " & AnoCont - 1, 0, _
                                                        Format(CantAConAnt, "#,##0"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificado Rec. AñoVto " & AnoCont, 0, _
                                                        Format(CantACon, "#,##0"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificado Rec. AñoVto " & AnoCont + 1, 0, _
                                                        Format(CantAConPos, "#,##"))


        '- Visualizo el progreso --------
        For Fila = 1 To T.NumFilas
            If IsEmpty(T.Datos(Fila, BD_ACont_Vto)) Then
                Cont_Fail = Cont_Fail + 1
            ElseIf VarType(T.Datos(Fila, BD_ACont_Vto)) = vbString Then
                If Len(T.Datos(Fila, BD_ACont_Vto)) = 0 Then Cont_Fail = Cont_Fail + 1
            End If
        Next Fila

        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Rec. SIN F_Vto. Asignada.", 0, Format(Cont_Fail, "#,##0"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " Rec. en BDatos", 0, Format(T.NumFilas, "#,##0"))
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación Fecha de Vto.", TimeLapSub)

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

    '- ---------------------------------------------------------------------------------------------
    '- Determinar Concepto Económico y Tipo de Enseñanza TIO-EP ------------------------------------
    '- ---------------------------------------------------------------------------------------------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación Concepto Económico y Tipo Ensañanza.", 0, , , , , , 2)
    For Fila = 1 To T.NumFilas
        T.Datos(Fila, BD_Concepto) = Empty                              '- Se supone que está vacía...
        T.Datos(Fila, BD_TIO_EP) = Empty                                '- Se supone que está vacía...
    Next Fila
    '- Los códigos van como texto ("1310.00"), igual que antes: al volcarlos, Excel los convierte en número (1310).

'-1º Recibos Cod_Activ = 6 - Grado -----------------------------------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_Reg_AE6")
        rowfind = Fnc_Asignar2(T, Cumple, "1310.00", "Grado")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1310.00'    Reg. AE6 Grado", 0)

'-2º Recibos Cod_Activ = 5 - Master ----------------------------------------------------------------
        Cumple = Fnc_Filtro_Filas(T, BD_ActivEco, "=", 5)
        rowfind = Fnc_Asignar2(T, Cumple, "1310.01", "Master")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1310.01'    Reg. AE5 Master", 0)

'-3º Recibos Cod_Activ = 2 - Doctorado -------------------------------------------------------------
        Cumple = Fnc_Filtro_Filas(T, BD_ActivEco, "=", 2)
        rowfind = Fnc_Asignar2(T, Cumple, "1310.02", "Doctorado")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1310.02'    Reg. AE2 Doctorado", 0)

'-4º Recibos - EFP - Estudios de Formación Permanente: Máster, Especialista, Experto. --------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_EFP")
        rowfind = Fnc_Asignar2(T, Cumple, "1311.00", "EFP")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                " '1311.00'    Reg. AE4 EFP: Estudios de Formación Permanente: Master, Especialista y Experto.", 0)

'-5º Recibos - CFC - Cursos de Formación Contínua --------------------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_CFC")
        rowfind = Fnc_Asignar2(T, Cumple, "1311.03", "CFC")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                " '1311.03'    Reg. AE4 CFC: Cursos de Formación Contínua", 0)

'-5º Recibos - AFC - Actividades de Formación Complementaria ---------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_AFC")
        rowfind = Fnc_Asignar2(T, Cumple, "1311.03", "AFC")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                " '1311.03'    Reg. AE4 AFC: Actividades de Formación Complementaria", 0)

'-6º Recibos Cod_Activ = 80 - Pruebas de aptitud para acceso a la Universidad ----------------------
        Cumple = Fnc_Filtro_Filas(T, BD_ActivEco, "=", 80)
        rowfind = Fnc_Asignar2(T, Cumple, "1315.00", "PruebasAccesoUni")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1315.00'    Reg. AE80 Pruebas Acceso Univ.", 0)

'-7º Recibos - TNCT-M013 AE300 - Cursos NO Contabilizables como EFP (M013) sino como Rec.Mov. por Secretaría de Acceso ---------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_TNCT_M013")
        rowfind = Fnc_Asignar2(T, Cumple, "1312.00", "TNCT_M013")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " '1312.00'    Reg. AE300 TNCT_M013: Seminario Orientación Pruebas > 25 años", 0)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(27, " ") & "(Secretaría de Acceso AE4/300/710)", 0)

'-8º _UPUA_ Recibos - CFC-TUP - _UPUA_ Universidad Permanente, Cursos de Formación Contínua (TUP) -------------------------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_TUP")
        rowfind = Fnc_Asignar2(T, Cumple, "1312.02", "CFC_UPUA")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                " '1312.02'    Reg. AE4 CFC_UPUA (TUP): Programa Univ. para Mayores UA. (Univ. Permanente)", 0)

'-9º Recibos - TNCT-PNB1 - Cursos NO Contabilizables como Títulos Propios Universidad (PNB1) -------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_TNCT_PNB1")
        rowfind = Fnc_Asignar2(T, Cumple, "1303.01", "TNCT_PNB1")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " '1303.01'    Reg. AE4 TNCT_PNB1: Prueba de competencias idioma extranjero.", 0)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(27, " ") & "(Centro Sup. Idiomas AE4/21)", 0)

'-10º Recibos de Movimiento - Rec_Adm - Recibos de Actividad Administrativa, SIN relación con Matrícula Académica ---
'-     (Tb_CriT_RecAdm pide Concepto y TIO_EP vacíos: los que no ha cogido ninguno de los pasos de arriba)
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_RecAdm")
        rowfind = Fnc_Asignar2(T, Cumple, "1303.00", "Rec_Adm")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                            " '1303.00'    Reg. Rec_Adm: Recibos de una actividad púramente Administrativa.", 0)

'-11º Recibos EURLE PLAN = C404 (Escuela Univ. Relaciones Laborales -ELDA-) ------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_Reg_EURLE")
        rowfind = Fnc_Asignar2(T, Cumple, "EURLElda", "EURLElda")
        Call Fnc_Asignar(T, Cumple, BD_Obs_Conta, "EURLElda")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                            " 'EURLElda'   Reg. AE6-Plan_C404 EURLElda: Esc. Univ. Relaciones Laborales Elda", 0)

'- Recibos de Matrículas de coste CERO - ImpMatCero - Recibos Matrícula de Actividad Académica a Coste CERO. ------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_ImpMatCero")
        rowfind = Fnc_Asignar2(T, Cumple, "NoContab", "ImpMatCero")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                            " 'No Contab.' Reg. ImpMatCero: Matrícula de Actividad Académica a Coste CERO.", 0)

'- Recibos Sin Concepto o Tipo ---------------------------------------------------------------------
        Cumple = Fnc_Filtro_Filas(T, BD_TIO_EP, "=", "")
        rowfind = Fnc_Asignar2(T, Cumple, "NoContab", "NoContab")
        If rowfind > 0 Then
            MsgBox "¡¡¡ Recibos SIN identificar Concepto-Eco o Tipo  !!!" & vbLf & vbLf & Format(rowfind, "#,##0") & _
                            " reg.", vbOKOnly + vbExclamation, "Proceso: Asignar Concepto Eco. y Tipo de Enseñanza"
        End If
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                                            " 'No Contab.' Reg. Recibos SIN identificar Concepto-Eco o Tipo.", 0)

    '- En la hoja: las columnas asignadas y una ordenación con el resultado de las de antes: F_Vto; Activ_Eco, TipoCurso
    '- y Plan; Activ_Eco y TipoCurso; Plan (antes del paso de EURLE).
    Call Rut_TablaRam_Volcar(T, Lo_Data)                                '- ACont_Vto, Concepto, TIO_EP y Obs_Conta
    Erase T.Datos
    Call Rut_TablaRam_Ordenar_Tandas(Lo_Data, Array(Array(BD_FVto), Array(BD_ActivEco, BD_TipoCurso, BD_Plan), _
                                                    Array(BD_ActivEco, BD_TipoCurso), Array(BD_Plan)))

        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "=") & Format(T.NumFilas, "#,##0"), 8) & " Reg. en BDatos", 0)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso Finalizado: Asignación Concepto Económico y Tipo Ensañanza.", TimeLapSub)

Restablecer_Valores:
    Lo_Data.ShowTotals = True
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
Debug.Print "<<< RuT_Assign_AnoVto_CtaCCC_ConcepEco_y_TipoEstudio"
End Sub
'- -------------------------------------------------------------------------------------------------

'- Pone Valor en la columna Col de las filas que cumplen y devuelve cuántas son -------------------
Private Function Fnc_Asignar(T As T_TablaRam, Cumple() As Boolean, ByVal Col As Long, ByVal Valor As Variant) As Long
    Dim Fila    As Long
    For Fila = 1 To T.NumFilas
        If Cumple(Fila) Then
            T.Datos(Fila, Col) = Valor
            Fnc_Asignar = Fnc_Asignar + 1
        End If
    Next Fila
    T.Modificada(Col) = True
End Function

'- Pone el Concepto y el TIO_EP de las filas que cumplen y devuelve cuántas son -------------------
Private Function Fnc_Asignar2(T As T_TablaRam, Cumple() As Boolean, ByVal Cod_Concepto As String, ByVal Tipo_Ens As String) As Long
    Call Fnc_Asignar(T, Cumple, BD_TIO_EP, Tipo_Ens)
    Fnc_Asignar2 = Fnc_Asignar(T, Cumple, BD_Concepto, Cod_Concepto)
End Function

'- Cuántas filas valen ese número en la columna (como CountIfs con un número: también el texto "2026") -
Private Function Fnc_Contar_Igual(T As T_TablaRam, ByVal Col As Long, ByVal Valor As Double) As Long
    Dim Fila    As Long
    Dim V       As Variant
    For Fila = 1 To T.NumFilas
        V = T.Datos(Fila, Col)
        If Not IsEmpty(V) Then
            If IsNumeric(V) Then
                If CDbl(V) = Valor Then Fnc_Contar_Igual = Fnc_Contar_Igual + 1
            End If
        End If
    Next Fila
End Function
'- -------------------------------------------------------------------------------------------------
