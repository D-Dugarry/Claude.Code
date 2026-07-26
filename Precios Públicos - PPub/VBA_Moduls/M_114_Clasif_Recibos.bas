Attribute VB_Name = "M_114_Clasif_Recibos"
'Rev.: 2026-01-22
Option Explicit

            Sub RuT_Clasif_Recibos_ByHand()
                Sht__BD.Select
                Sht__BD.Unprotect
                Sht__BD.Columns(Sht__BD.ListObjects(1).Range.Columns(BD_Ref).Column).Hidden = False
                Call RuT_Clasif_Recibos
                Sht__BD.Unprotect
                MsgBox "FIN"
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Clasificar Recibos en Emitidos, Remesados, EjeAnt, ADxAplz, Añejas ---------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Clasif_Recibos()
Debug.Print ">>> RuT_Clasif_Recibos"
    Dim RegsEmitido         As Long
    Dim RegsEjeAnt          As Long
    Dim RegsAñejo           As Long
    Dim RegsAplazado        As Long
    Dim RegsADxAplz         As Long
    Dim RegsADxAplzEPCurs   As Long
    Dim RegsAplazadoEPCurs  As Long
    Dim RegsSinTipo         As Long
    Dim RegsErrDate         As Long
    Dim RegsAnulado         As Long
    Dim RegsContabAnt       As Long
    Dim RegsClasifs         As Long
    Dim RegsDevolucion      As Long
    Dim CantDto             As Long
    Dim CantDEV             As Long
    Dim RegsCanTot          As Long
    Dim RegsNOCUADRA        As Long
    
    Dim rowfind             As Variant
    Dim APP_AñoCont         As String:      APP_AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim TxT_Resumen         As String
    Dim TimeLapSub          As Single:      TimeLapSub = LastTimeLap

    Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
    
    Sht__BD.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
'    Lo_DefCol_BD.TotalsRowRange(DefC_HiddenCol) = False
    Prog__APP.Range("SW_Col_Hide_Sht__BD") = False
    Sht__BD.Unprotect
    Lo_BD.ShowTotals = False
    
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificación de Recibos, Estadística:", 0, , , , , , 2)
    With Lo_BD
    
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        Call Rut_Lo_Sort(Lo_BD, BD_ACont_Emi, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_BD, BD_ACont_Vto, xlAscending, False)
        Call Rut_Lo_Sort(Lo_BD, BD_ACont_Cob, xlAscending, False)
                
        '-ClearContents de Rec. BD_C_Acad = C_Acad -------------------------------------
'        Call Rut_Lo_Filtros_Quitar(Lo_BD)
'        .Range.AutoFilter Field:=BD_ActivEco, Criteria1:="<>4"  '- Los AE4 ya los tienen
'        RowsFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
'        If RowsFind > 0 Then
'            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).ClearContents
'        End If
    
        .DataBodyRange.Columns(BD_Tipo_Rec).ClearContents
        .DataBodyRange.Columns(BD_CriT_Emi).Resize(, 56).ClearContents
        '.DataBodyRange.Columns(BD_CriT_Emi).Resize(, BD_CriT_ErrDate - BD_CriT_Emi + 1).ClearContents
        '.DataBodyRange.Columns(BD_CriT_Emi).Resize(, BD_CriT_ErrDate - BD_CriT_Emi + 2).ClearContents

    '-Rec. Emitidos -------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Emitido")
        RegsEmitido = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref

        If RegsEmitido > 0 Then
            .DataBodyRange.Columns(BD_CriT_Emi).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
        End If
            
    '-Rec. EjeAnt -------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_EjeAnt")
        RegsEjeAnt = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RegsEjeAnt > 0 Then
            .DataBodyRange.Columns(BD_CriT_EjeAnt).SpecialCells(xlCellTypeVisible).Cells.Value = "EjeAnt"
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "EjeAnt"
        End If
    
    '-Rec. Añejos -------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Añeja")
        RegsAñejo = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RegsAñejo > 0 Then
            .DataBodyRange.Columns(BD_CriT_Añejo).SpecialCells(xlCellTypeVisible).Cells.Value = "Añejo"
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Añejo"
        End If
    
    '-Rec. Aplazado -------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Aplazado")
        RegsAplazado = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RegsAplazado > 0 Then
            .DataBodyRange.Columns(BD_CriT_Aplazado).SpecialCells(xlCellTypeVisible).Cells.Value = "Aplazado"
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Aplazado"
        End If
    
    '-Rec. ADxAplz -------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_ADxAplz")
        RegsADxAplz = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RegsADxAplz > 0 Then
            .DataBodyRange.Columns(BD_CriT_ADxAplz).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz"
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz"
        End If
              
    '-----------------------------------------------------------------------------------------------------------------
    '------------ A partir de ahora, las marcas pueden sobreescribir algún valor anterior ----------------------------
    '-----------------------------------------------------------------------------------------------------------------
    '-_Dev_EP_--------------------------------------------------------------------------------------------------------
        '-Filtra Cobradas en Años anteriores al de Emisión -------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        .Range.AutoFilter Field:=BD_ImpRec, Criteria1:="<0"
        RegsDevolucion = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RegsDevolucion > 0 Then
            .DataBodyRange.Columns(BD_CriT_DevEP).SpecialCells(xlCellTypeVisible).Cells.Value = "_Dev_EP_"
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_Dev_EP_"
        End If
        
    '-Filtra Recibos Anulados, NO Matrícula o Invalidados -------------------------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_Anul")
        RegsAnulado = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RegsAnulado > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_Reg_Anul_"
        End If
    
    '-_Contab_Ant_--------------------------------------------------------------------------------------------------------
        '-Filtra Cobradas en Años anteriores al de Emisión -------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        .Range.AutoFilter Field:=BD_ACont_Cob, Criteria1:="<" & APP_AñoCont
        RegsContabAnt = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RegsContabAnt > 0 Then
            .DataBodyRange.Columns(BD_CriT_ContabAnt).SpecialCells(xlCellTypeVisible).Cells.Value = "_Contab_Ant_"
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_Contab_Ant_"
        End If
        
        rowfind = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=Emitido")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. Emididos: ", 0, Format(rowfind, "#,##0") & " reg.")
        
        rowfind = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=EjeAnt")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. EjeAnt: ", 0, Format(rowfind, "#,##0") & " reg.")
        
        rowfind = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=Añejo")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. Añejos: ", 0, Format(rowfind, "#,##0") & " reg.")
        
        rowfind = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=Aplazado")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. Aplazados: ", 0, Format(rowfind, "#,##0") & " reg.")
        
        rowfind = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=ADxAplz")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. ADxAplz: ", 0, Format(rowfind, "#,##0") & " reg.")
        
        rowfind = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=_Dev_EP_")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. de EP=AE4x4 _Devolución_EP_: ", 0, Format(rowfind, "#,##0") & " reg.")
        
        rowfind = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=_Reg_Anul_")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. _Reg_Anul_: ", 0, Format(rowfind, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Registros: Anulados, NO Matrícula o Invalidados: ", 0)
        
        rowfind = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=_Contab_Ant_")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. _Contab_Ant_: ", 0, Format(rowfind, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Registros, Cobrados anteriormente y por lo tanto, ya Contabilizados. ", 0)
        
        '---------------------------------------------------------------------------------------------------------
        RegsClasifs = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "<>")
        RegsNOCUADRA = Application.CountIfs(.DataBodyRange.Columns(BD_Tipo_Rec), "=")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Sumatorio Recibos clasificados (de " & Format(Lo_BD.ListRows.Count, "#,##0") & " reg.)", 0, _
                                        Format(RegsClasifs, "#,##0") & " reg.", "faltan = " & Format(RegsNOCUADRA, "#,##0") & " reg.", , , , 2, 2)
        
        '---------------------------------------------------------------------------------------------------------
        '-Filtra Recibos ErrDate - Cobradas en Años anteriores al de Emisión -------------------------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_Err")
        RegsErrDate = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RegsErrDate > 0 Then
            .DataBodyRange.Columns(BD_Incidencias).SpecialCells(xlCellTypeVisible).Cells.Value = "_ERR_Date_"
        End If
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. _ERR_Date_: ", 0, _
                                        Format(RegsErrDate, "#,##0") & " reg.")
    
        '--------------------------------------------------------------------------------------------------------
        '-Filtra Recibos Sin Tipo ---------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------------------------
        .ShowAutoFilter = True          '- El AdvancedFilter con Rango de Criterio desactiva el "ShowFilterMarck"
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Total Recibos BDatos: ", 0, _
                                        Format(Lo_BD.ListRows.Count, "#,##0") & " reg.", , , , , 2)
        '-Filtra Recibos Sin Tipo ---------------------------------------------------------------------------------
        Call Rut_Lo_Sort(Lo_BD, BD_Tipo_Rec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="="
        RegsSinTipo = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
Debug.Print "RegsSinTipo", RegsSinTipo
        RegsCanTot = .ListRows.Count
Debug.Print "RegsCanTot", , Format(RegsCanTot, "#,##0")
        RegsNOCUADRA = RegsCanTot - (RegsEmitido + RegsEjeAnt + RegsAñejo + RegsAplazado + RegsADxAplz + RegsSinTipo)
Debug.Print "RegsNOCUADRA", RegsNOCUADRA
        'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Left("Reg. Sin Clasificar: " & String(35, "_"), 35) & _
                Right(String(15, "_") & Format(RegsNOCUADRA, "#,##0") & " reg.", 15), LastTimeLap)

        .ShowAutoFilter = True          '- El AdvancedFilter con Rango de Criterio desactiva el "ShowFilterMarck"

    Call Rut_Lo_Filtros_Quitar(Lo_BD)
       
    End With        '- Lo_BD
    
    '- Visualizo el progreso -
    'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificación de registros finalizada.", TimeLapSub)

    Lo_BD.ShowTotals = True

Debug.Print "<<< RuT_Clasif_Recibos"

End Sub


