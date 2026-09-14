Attribute VB_Name = "Módulo3"
    Dim RegsEmis            As Long     ' regs Emitidos
    Dim Imp_Emis            As Currency
    Dim TImpEmis            As Currency
    Dim Imp_EmisAdm         As Currency
    Dim TImpEmisAdm         As Currency
    Dim Imp_EmisAcad        As Currency
    Dim TImpEmisAcad        As Currency
    
    Dim RegsEmisAnt         As Long     ' regs Emitidos
    Dim Imp_EmisAnt         As Currency
    Dim TImpEmisAnt         As Currency
    Dim Imp_EmisAntAdm      As Currency
    Dim TImpEmisAntAdm      As Currency
    Dim Imp_EmisAntAcad     As Currency
    Dim TImpEmisAntAcad     As Currency
    
    Dim RegsEmisPos         As Long     ' regs Emitidos
    Dim Imp_EmisPos         As Currency
    Dim TImpEmisPos         As Currency
    Dim Imp_EmisPosAdm     As Currency
    Dim TImpEmisPosAdm     As Currency
    Dim Imp_EmisPosAcad    As Currency
    Dim TImpEmisPosAcad    As Currency

    Dim RegsAnul            As Long     ' regs Emitidos
    Dim Imp_Anul            As Currency
    Dim TImpAnul            As Currency
    Dim Imp_AnulAdm         As Currency
    Dim TImpAnulAdm         As Currency
    Dim Imp_AnulAcad        As Currency
    Dim TImpAnulAcad        As Currency
    
    Dim Imp_Cobr            As Currency
    Dim TImpCobr            As Currency
    Dim Imp_CobrAdm         As Currency
    Dim TImpCobrAdm         As Currency
    Dim Imp_CobrAcad        As Currency
    Dim TImpCobrAcad        As Currency
    
    Dim Imp__RDT            As Currency
    Dim TImp_RDT            As Currency
    Dim Imp__RDTAdm         As Currency
    Dim TImp_RDTAdm         As Currency
    Dim Imp__RDTAcad        As Currency
    Dim TImp_RDTAcad        As Currency
    
    Dim Imp_SRDT            As Currency
    Dim TImpSRDT            As Currency
    Dim Imp_SRDTAdm         As Currency
    Dim TImpSRDTAdm         As Currency
    Dim Imp_SRDTAcad        As Currency
    Dim TImpSRDTAcad        As Currency
    
    Dim Imp__ADx            As Currency
    Dim TImp_ADx            As Currency
    Dim Imp__ADxAdm         As Currency
    Dim TImp_ADxAdm         As Currency
    Dim Imp__ADxAcad        As Currency
    Dim TImp_ADxAcad        As Currency

    Dim Imp_Aplz            As Currency
    Dim TImpAplz            As Currency
    Dim Imp_AplzAdm         As Currency
    Dim TImpAplzAdm         As Currency
    Dim Imp_AplzAcad        As Currency
    Dim TImpAplzAcad        As Currency

    Dim Imp_PdtCob          As Currency
    Dim TimpPdtCob          As Currency
    Dim Imp_PdtCobAdm       As Currency
    Dim TimpPdtCobAdm       As Currency
    Dim Imp_PdtCobAcad      As Currency
    Dim TimpPdtCobAcad      As Currency

    Dim RregsEjeAnt         As Long     ' regs Pendiente de pago
    Dim Imp__EjeAnt         As Currency
    Dim TImp_EjeAnt         As Currency
    Dim Imp__EjeAntAdm      As Currency
    Dim TImp_EjeAntAdm      As Currency
    Dim Imp__EjeAntAcad     As Currency
    Dim TImp_EjeAntAcad     As Currency
    
    Const Lst_Orden                 As Integer = 1
    Const Lst_Plan                  As Integer = 2
    Const Lst_RegsEmis              As Integer = 3
    Const Lst_Imp_Emis              As Integer = 4
    Const Lst_Imp_EmisAdm           As Integer = 5
    Const Lst_Imp_EmisAcad          As Integer = 6
    Const Lst_Imp_EmisAnt           As Integer = 7
    Const Lst_Imp_EmisAntAdm        As Integer = 8
    Const Lst_Imp_EmisAntAcad       As Integer = 9
    Const Lst_Imp_EmisPos           As Integer = 10
    Const Lst_Imp_EmisPosAdm        As Integer = 11
    Const Lst_Imp_EmisPosAcad       As Integer = 12
    Const Lst_Imp_Cobr              As Integer = 13
    Const Lst_Imp_CobrAdm           As Integer = 14
    Const Lst_Imp_CobrAcad          As Integer = 15
    Const Lst_Imp_CobrAnt           As Integer = 16
    Const Lst_Imp_CobrAntAdm        As Integer = 17
    Const Lst_Imp_CobrAntAcad       As Integer = 18
    Const Lst_Imp_CobrPos           As Integer = 19
    Const Lst_Imp_CobrPosAdm        As Integer = 20
    Const Lst_Imp_CobrPosAcad       As Integer = 21
    Const Lst_Imp_Anul              As Integer = 22
    Const Lst_Imp_AnulAdm           As Integer = 23
    Const Lst_Imp_AnulAcad          As Integer = 24
    Const Lst_Imp__RDT              As Integer = 25
    Const Lst_Imp__RDTAdm           As Integer = 26
    Const Lst_Imp__RDTAcad          As Integer = 27
    Const Lst_Imp_RDTpdt              As Integer = 28
    Const Lst_Imp_RDTpdtAdm           As Integer = 29
    Const Lst_Imp_RDTpdtAcad          As Integer = 30
    Const Lst_Imp_ADx               As Integer = 31
    Const Lst_Imp_ADxAdm            As Integer = 32
    Const Lst_Imp_ADxAcad           As Integer = 33
    Const Lst_Imp_Aplz              As Integer = 34
    Const Lst_Imp_AplzAdm           As Integer = 35
    Const Lst_Imp_AplzAcad          As Integer = 36
    Const Lst_Imp_EjeAnt            As Integer = 37
    Const Lst_Imp_EjeAntAdm         As Integer = 38
    Const Lst_Imp_EjeAntAcad        As Integer = 39
    Const Lst_Imp_PdtCob            As Integer = 40
    Const Lst_Imp_PdtCobAdm         As Integer = 41
    Const Lst_Imp_PdtCobAcad        As Integer = 42





'==================================================================================================================================
Sub RuT_Estadística_Contable_Planes_CAcad_Pos(Lo_BD As ListObject, _
                                         CursoAcad As String, _
                                         TipoCurso As String, _
                                         Ws_Lista As Worksheet)
'==================================================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim Cont            As Long
    Dim ContIni         As Long:        ContIni = 1
    Dim Txt_Cabecera    As String
    Dim Cont_Plan       As Integer
    Dim PlanesSinCob    As Integer:     PlanesSinCob = 0

    Dim RegsEmis        As Long     ' regs Emitidos
    Dim Imp_Emis        As Currency
    Dim TImpEmis        As Currency

    Dim RegsEmisAdm     As Long     ' regs Emitidos
    Dim Imp_EmisAdm     As Currency
    Dim TImpEmisAdm     As Currency

    Dim RegsEmisAcad    As Long     ' regs Emitidos
    Dim Imp_EmisAcad    As Currency
    Dim TImpEmisAcad    As Currency

    Dim Imp_Cobr            As Currency
    Dim TImpCobr            As Currency
    Dim Imp_CobrAdm         As Currency
    Dim TImpCobrAdm         As Currency
    Dim Imp_CobrAcad        As Currency
    Dim TImpCobrAcad        As Currency

    Dim Imp__RDT            As Currency
    Dim TImp_RDT            As Currency
    Dim Imp__RDTAdm         As Currency
    Dim TImp_RDTAdm         As Currency
    Dim Imp__RDTAcad        As Currency
    Dim TImp_RDTAcad        As Currency

    Dim Imp_SRDT            As Currency
    Dim TImpSRDT            As Currency
    Dim Imp_SRDTAdm         As Currency
    Dim TImpSRDTAdm         As Currency
    Dim Imp_SRDTAcad        As Currency
    Dim TImpSRDTAcad        As Currency

    Dim Imp__ADx            As Currency
    Dim TImp_ADx            As Currency

    Dim RegsADxAdm     As Long     ' regs Pendiente de pago
    Dim Imp_ADxAdm     As Currency
    Dim TImpADxAdm     As Currency

    Dim RegsADxAcad    As Long     ' regs Pendiente de pago
    Dim Imp_ADxAcad    As Currency
    Dim TImpADxAcad    As Currency

    Dim Imp_PdtCob        As Currency
    Dim TimpPdtCob        As Currency
    Dim Imp_PdtCobAdm     As Currency
    Dim TimpPdtCobAdm     As Currency
    Dim Imp_PdtCobAcad    As Currency
    Dim TimpPdtCobAcad    As Currency

    Const Lst_Orden             As Integer = 1
    Const Lst_Plan              As Integer = 2
    Const Lst_RegsEmis          As Integer = 3
    Const Lst_Imp_Emis          As Integer = 4
    Const Lst_Imp_EmisAdm       As Integer = 5
    Const Lst_Imp_EmisAcad      As Integer = 6
    Const Lst_Imp_Cobr          As Integer = 7
    Const Lst_Imp__RDT          As Integer = 8
    Const Lst_Imp__RDTAdm       As Integer = 9
    Const Lst_Imp__RDTAcad      As Integer = 10
    Const Lst_Imp_RDTpdt         As Integer = 11
    Const Lst_Imp_RDTpdtAdm       As Integer = 12
    Const Lst_Imp_RDTpdtAcad      As Integer = 13
    Const Lst_RegsSRDT          As Integer = 14
    Const Lst_Imp__ADx          As Integer = 15
    Const Lst_Imp__ADxAdm       As Integer = 16
    Const Lst_Imp__ADxAcad      As Integer = 17
    Const Lst_Imp_PdtCob        As Integer = 18
    Const Lst_Imp_PdtCobAdm     As Integer = 19
    Const Lst_Imp_PdtCobAcad    As Integer = 20

    Dim Regs_Dev            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula
    Dim Imp_Devo            As Currency
    Dim RDev_Pag            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, Pagados
    Dim IDev_Pag            As Currency
    Dim RDevPdte            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, No Pagados, pero Ajuste Matrícula NO deben ser pagados
    Dim IDevPdte            As Currency

'''    Dim Ws_Lista    As Worksheet:       Set Ws_Lista = Sheets("Cierre_Planes_CAcad_Pos_" & TipoCurso)
    Dim Lo_Lst      As ListObject:      Set Lo_Lst = Ws_Lista.ListObjects(1)
    Dim RowNew      As ListRow
    Dim Row_BD      As ListRow

    Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso

    If Not Lo_Lst.DataBodyRange Is Nothing Then Lo_Lst.DataBodyRange.Delete

    '- Identificar Todos los Planes, Crea Dictionary para valores únicos (eficiente para grandes datos)
    Dim Collection_Planes      As New Collection
    Dim RngPlanes       As Range
    Set RngPlanes = Lo_BD.ListColumns(BD_Plan).DataBodyRange.SpecialCells(xlCellTypeVisible)
    Dim Celda As Range
    Dim CodPlan As String
    On Error Resume Next
    For Each Celda In RngPlanes.Cells
        If Not IsEmpty(Celda.Value) Then
            CodPlan = CStr(Celda.Value)
            Collection_Planes.Add Celda.Value, CodPlan   ' Solo agrega si no existe la CodPlan
        End If
    Next Celda
    On Error GoTo 0

    Call Rut_Lo_Sort(Lo_BD, BD_Tipo_Rec, xlAscending, False)
    ' Inicio Listado en Tabla excel
    Txt_Cabecera = "Resumen del Año Contable " & AñoCont & ", de los Planes de " & TipoCurso & "_" & CursoAcad & String(10, " ") & Now
    Ws_Lista.Range("b2") = Txt_Cabecera
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Emis).Value = "Importe" & vbLf & "Emitido" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisAdm).Value = "Emitido" & vbLf & "1303.00 " & vbLf & "Adm. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisAcad).Value = "Emitido" & vbLf & "1311.00 " & vbLf & "Acad. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Cobr).Value = "Importe" & vbLf & "Cobrado" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_CobrAdm).Value = "Cobrado" & vbLf & "1303.00 " & vbLf & "Adm. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_CobrAcad).Value = "Cobrado" & vbLf & "1311.00 " & vbLf & "Acad. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDT).Value = "Imp_RDT" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDTAdm).Value = "Imp_RDT" & vbLf & "1303.00 " & vbLf & "Adm. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDTAcad).Value = "Imp_RDT" & vbLf & "1311.00 " & vbLf & "Acad. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_RDTpdt).Value = "Pdte_RDT"
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_RDTpdtAdm).Value = "Pdte_RDT" & vbLf & "1303.00 " & vbLf & "Adm. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_RDTpdtAcad).Value = "Pdte_RDT" & vbLf & "1311.00 " & vbLf & "Acad. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADx).Value = "Emi'" & Right(AñoCont, 2) & " Vto'" & Right(AñoCont + 1, 2) & vbLf & "ADxAplz" & vbLf & AñoCont + 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADxAdm).Value = "Imp.ADxAplz" & vbLf & "1303.00 " & vbLf & "Adm. " & AñoCont + 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADxAcad).Value = "Imp.ADxAplz" & vbLf & "1311.00 " & vbLf & "Acad. " & AñoCont + 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCob).Value = "Emi'" & Right(AñoCont, 2) & " Vto'" & Right(AñoCont, 2) & vbLf & "Pdte_Cob" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCobAdm).Value = "Pdte_Cob" & vbLf & "1303.00 " & vbLf & "Adm. " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCobAcad).Value = "Pdte_Cob" & vbLf & "1311.00 " & vbLf & "Acad. " & AñoCont

    '- Visualizo el progreso ---------------------------------------------------------------------------------------
    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TB_Informe = Txt_Cabecera & vbCrLf & vbCrLf
    Txt_Cabecera = String(27, " ") & "Cob-" & AñoCont & String(18, " ") & "Cob-" & AñoCont & String(6, " ") & "ADxAplz" & String(6, " ") & "ADxAplz" & String(7, " ") & AñoCont & vbLf & _
                   "    Plan     Imp_Rec       Imp_Cob       Imp_RDT     Pdte.RDT     Adm." & AñoCont + 1 & "     Acad." & AñoCont + 1 & "     Pde_Cob"
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera & vbLf

    
    '- Recorro la Collection con todos los Planes
         Dim Cod_Plan As Variant
    With Lo_BD.DataBodyRange
    For Each Cod_Plan In Collection_Planes

        Cont_Plan = Cont_Plan + 1
        Set RowNew = Lo_Lst.ListRows.Add
            RowNew.Range(Lst_Orden) = Cont_Plan
            RowNew.Range(Lst_Plan) = Cod_Plan

'- Importe Recibos Emi_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_Emis = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            If Imp_Emis = 0 Then GoTo Siguiente_Plan    '- Si todo está cobrado en año anterior
            RowNew.Range(Lst_Imp_Emis) = Imp_Emis
            TImpEmis = TImpEmis + Imp_Emis
            RegsEmis = Application.CountIfs( _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_RegsEmis) = RegsEmis

            '- Importe Recibos Emi_Adm -----------------------------------------------------------------------------------------------------------------
            Imp_EmisAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_EmisAdm) = Imp_EmisAdm
            TImpEmisAdm = TImpEmisAdm + Imp_EmisAdm
            
            '- Importe Recibos Emi_Acad -----------------------------------------------------------------------------------------------------------------
            Imp_EmisAcad = Imp_Emis - Imp_EmisAdm
            RowNew.Range(Lst_Imp_EmisAcad) = Imp_EmisAcad
            TImpEmisAcad = TImpEmisAcad + Imp_EmisAcad
            
'- Importe Cob_AñoCont, Emi_AñoCont --------------------------------------------------------------------------------------------------------------------
            Imp_Cobr = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_Cobr) = Imp_Cobr
            TImpCobr = TImpCobr + Imp_Cobr
            
'- Importe RDT Cob_AñoCont, Emi_AñoCont --------------------------------------------------------------------------------------------------------------------
            Imp__RDT = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "<>", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp__RDT) = Imp__RDT
            TImp_RDT = TImp_RDT + Imp__RDT

            '- Importe Adm RDT Cob_AñoCont, Emi_AñoCont --------------------------------------------------------------------------------------------------------------------
            Imp__RDTAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "<>", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp__RDTAdm) = Imp__RDTAdm
            TImp_RDTAdm = TImp_RDTAdm + Imp__RDTAdm
            '- Importe Adm RDT Cob_AñoCont, Emi_AñoCont --------------------------------------------------------------------------------------------------------------------
            Imp__RDTAcad = Imp__RDT - Imp__RDTAdm
            RowNew.Range(Lst_Imp__RDTAcad) = Imp__RDTAcad
            TImpCobrAcad = TImpCobrAcad + Imp__RDTAcad

'- Importe Pdte_RDT --------------------------------------------------------------------------------------------------------------------
            Imp_SRDT = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "=", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_RDTpdt) = Imp_SRDT
            TImpSRDT = TImpSRDT + Imp_SRDT

            '- Importe Adm Pdte_RDT --------------------------------------------------------------------------------------------------------------------
            Imp_SRDTAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "=", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_RDTpdtAdm) = Imp_SRDTAdm
            TImpSRDTAdm = TImpSRDTAdm + Imp_SRDTAdm
            '- Importe Acad Pdte_RDT --------------------------------------------------------------------------------------------------------------------
            Imp_SRDTAcad = Imp_SRDT - Imp_SRDTAdm
            RowNew.Range(Lst_Imp_RDTpdtAcad) = Imp_SRDTAcad
            TImpSRDTAcad = TImpSRDTAcad + Imp_SRDTAcad

'- Importe ADxADxAplz --------------------------------------------------------------------------------------------------------------------
            Imp__ADx = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont + 1, _
                                          .Columns(BD_ACont_Cob), "<>" & AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp__ADx) = Imp__ADx
            TImp_ADx = TImp_ADx + Imp__ADx

            '- Importe ADxAplzAdm --------------------------------------------------------------------------------------------------------------------
            Imp_ADxAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont + 1, _
                                          .Columns(BD_ACont_Cob), "<>" & AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp__ADxAdm) = Imp_ADxAdm
            TImpADxAdm = TImpADxAdm + Imp_ADxAdm

            '- Importe ADxAplzAcad --------------------------------------------------------------------------------------------------------------------
            Imp_ADxAcad = Imp__ADx - Imp_ADxAdm
            RowNew.Range(Lst_Imp__ADxAcad) = Imp_ADxAcad
            TImpADxAcad = TImpADxAcad + Imp_ADxAcad


'- Importe Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            Imp_PdtCob = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont, _
                                          .Columns(BD_ACont_Cob), "=", _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_PdtCob) = Imp_PdtCob
            TimpPdtCob = TimpPdtCob + Imp_PdtCob

            '- Importe Adm Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            Imp_PdtCobAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont, _
                                          .Columns(BD_ACont_Cob), "=", _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_PdtCobAdm) = Imp_PdtCobAdm
            TimpPdtCobAdm = TimpPdtCobAdm + Imp_PdtCobAdm
            
            '- Importe Acad Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            Imp_PdtCobAcad = Imp_PdtCob - Imp_PdtCobAdm
            RowNew.Range(Lst_Imp_PdtCobAcad) = Imp_PdtCobAcad
            TimpPdtCobAcad = TimpPdtCobAcad + Imp_PdtCobAcad
            


            '-- Informe --------------------------------------------------------------------------------------------------------------------------
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Cont_Plan, "00") & "º " & Cod_Plan & " " & _
                Right(String(12, " ") & Format(Imp_Emis, "#,##0.00"), 13) & _
                IIf(Imp_Cobr > 0, " " & Right(String(12, "·") & Format(Imp_Cobr, "#,##0.00"), 13), "  ¡Sin Cobros!") & _
                IIf(Imp__RDT > 0, " " & Right(String(12, "·") & Format(Imp__RDT, "#,##0.00"), 13), " " & String(13, "·")) & _
                IIf(Imp_SRDT > 0, " " & Right(String(12, "·") & Format(Imp_SRDT, "#,##0.00"), 12), " " & String(12, "·")) & _
                IIf(Imp_ADxAdm > 0, " " & Right(String(12, "·") & Format(Imp_ADxAdm, "#,##0.00"), 11), " " & String(11, "·")) & _
                IIf(Imp_ADxAcad > 0, " " & Right(String(12, "·") & Format(Imp_ADxAcad, "#,##0.00"), 13), " " & String(13, "·")) & _
                IIf(Imp_PdtCob > 0, " " & Right(String(12, "·") & Format(Imp_PdtCob, "#,##0.00"), 12), " " & String(12, "·")) & vbLf

            ' Intercalo Cabecera Columnas
            If Cont_Plan Mod 20 = 0 Then Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Txt_Cabecera & vbLf

            If Imp_Cobr = 0 Then PlanesSinCob = PlanesSinCob + 1
Siguiente_Plan:
    Next Cod_Plan
    End With    ' Lo_BD.DataBodyRange

            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Txt_Cabecera
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & " Totales "
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                Right(String(12, " ") & Format(TImpEmis, "#,##0.00"), 13) & _
                Right(String(12, " ") & Format(TImpCobr, "#,##0.00"), 14) & _
                Right(String(12, " ") & Format(TImp_RDT, "#,##0.00"), 14) & _
                Right(String(12, " ") & Format(TImpSRDT, "#,##0.00"), 13) & _
                Right(String(12, " ") & Format(TImpADxAdm, "#,##0.00"), 12) & _
                Right(String(12, " ") & Format(TImpADxAcad, "#,##0.00"), 14) & _
                Right(String(12, " ") & Format(TimpPdtCob, "#,##0.00"), 13) & vbLf


            If PlanesSinCob > 1 Then
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡¡ " & PlanesSinCob & " Planes sin Cobros !!"
            ElseIf PlanesSinCob = 1 Then
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡¡ 1 Plan sin Cobros !!"
            Else
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡ No hay planes sin Cobros !"
            End If

'- Visualizo el progreso ---------------------------------------------------------------------------------------
Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & _
                                                "¡¡¡ Proceso concluido !!! día: " & Now() & _
                                                " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg."
'End With    '- Lo_BD.DataBodyRange

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
''''Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Listar_Planes   --------------------------------------------------------------------------------------------
'===================================================================================================================================


