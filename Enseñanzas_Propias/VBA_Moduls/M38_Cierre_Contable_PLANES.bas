Attribute VB_Name = "M38_Cierre_Contable_PLANES"
'2026-01-27
'- M31_Cierre_Contable_PLANES
Option Explicit

'==================================================================================================================================
Sub Rut_Cierre_Contable_AñoCont()
'==================================================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim CursoAcad       As String:      CursoAcad = Prog__APP.Range("APP_CursAcad")
    Dim CursoAcadAnt    As String:      CursoAcadAnt = Prog__APP.Range("APP_C_Acad_Ant")
    Dim TipoCurso       As String:      TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")

    '- Verifica que la ActiveSheet es coherente con el CursoAcad y AñoCont, es decir Ant o Pos.
    If CursoAcad = CursoAcadAnt Then
        If ActiveSheet.Name <> Sht__Inf_Cierre_ACont_CAcad_Ant.Name Then
                MsgBox "Error, estamos en el AñoCont " & AñoCont & vbLf & vbLf & _
                       " Y No coincide el CursoAcad " & CursoAcad & vbLf & vbLf & _
                       " Con la hoja de Informe " & Sht__Inf_Cierre_ACont_CAcad_Ant.Name & vbLf & vbLf & _
                       " Cambiar a la hoja de Informe " & Sht__Inf_Cierre_ACont_CAcad_Pos.Name, _
                       vbOKOnly, "Acutalizar Informe Cierre AñoCont"
                GoTo Salir
        End If
    Else
        If ActiveSheet.Name <> Sht__Inf_Cierre_ACont_CAcad_Pos.Name Then
                MsgBox "Error, estamos en el AñoCont " & AñoCont & vbLf & vbLf & _
                       " Y No coincide el CursoAcad " & CursoAcad & vbLf & vbLf & _
                       " Con la hoja de Informe " & Sht__Inf_Cierre_ACont_CAcad_Pos.Name & vbLf & vbLf & _
                       " Cambiar a la hoja de Informe " & Sht__Inf_Cierre_ACont_CAcad_Ant.Name, _
                       vbOKOnly, "Acutalizar Informe Cierre AñoCont"
                GoTo Salir
        End If
    End If
    
    Dim RngVisible  As Range
    Dim Lo_BD           As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
        Prog_BD.Unprotect:     Lo_BD.ShowTotals = False
        Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Dim Lo_BD_Filtrada  As ListObject
                
    Dim WsBuffer As Worksheet: Set WsBuffer = ThisWorkbook.Worksheets("Sheet_Buffer")  'hoja fija/oculta
    ' 1)Limpiar anterior tabla en Buffer
    If WsBuffer.ListObjects.Count > 0 Then WsBuffer.ListObjects(1).Delete
    ' 2) Ordenaciones
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_ACont_Emi, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    ' 3) Filtros
    Lo_BD.Range.AutoFilter Field:=BD_ImpAdm, Criteria1:=">=0"        '- Si <0 son Ajustes de Matrícula
    Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="<>_ERR_Date_"
    ' 4) Rango visible filtrado (incluye cabeceras)
    On Error Resume Next
        Set RngVisible = Lo_BD.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo SalirLimpiando
        If RngVisible Is Nothing Then GoTo SalirLimpiando
    ' Limpia cualquier tabla previa en Buffer
    If WsBuffer.ListObjects.Count > 0 Then
        WsBuffer.ListObjects(1).Delete  '  .Unlist o .Delete si prefieres quitar formato tabla
    End If
    WsBuffer.Cells.Clear
    ' 6) Copiar solo lo visible
    RngVisible.Copy Destination:=WsBuffer.Range("A1")
    ' 7) Crear nuevo ListObject con los datos filtrados
    Set Lo_BD_Filtrada = WsBuffer.ListObjects.Add( _
                            SourceType:=xlSrcRange, _
                            Source:=WsBuffer.Range("A1").CurrentRegion, _
                            XlListObjectHasHeaders:=xlYes)
    ' 8) Llamar a tu rutina que trabaja con ListObject
    If CursoAcad = CursoAcadAnt Then
        Call RuT_Cierre_Contable_Planes_CAcad_Ant(Lo_BD_Filtrada, CursoAcad, TipoCurso, Sht__Inf_Cierre_ACont_CAcad_Ant)
    Else
        Call RuT_Cierre_Contable_Planes_CAcad_Pos(Lo_BD_Filtrada, CursoAcad, TipoCurso, Sht__Inf_Cierre_ACont_CAcad_Pos)
    End If
    GoTo Salir

    
SalirLimpiando:
    ' Opcional: quitar tabla en Buffer para dejar solo datos o la hoja limpia
    On Error Resume Next
    If Not WsBuffer Is Nothing Then
        If WsBuffer.ListObjects.Count > 0 Then WsBuffer.ListObjects(1).Unlist
        'Si quieres dejar la hoja totalmente vacía:
        WsBuffer.Cells.Clear
    End If
    On Error GoTo 0
    
Salir:
    ' Restaurar estado de Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Call Rut_EnableEvents_Status_Reset
    Application.Speech.Speak "Proceso Terminado"
End Sub '------------------------------------------------------------------------------------------------------------------------




'==================================================================================================================================
Sub RuT_Cierre_Contable_Planes_CAcad_Pos(Lo_BD As ListObject, _
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
    Dim Imp_EmisAdm     As Currency
    Dim TImpEmisAdm     As Currency
    Dim Imp_EmisAcad    As Currency
    Dim TImpEmisAcad    As Currency

    Dim RegsAnul        As Long     ' regs Emitidos
    Dim Imp_Anul        As Currency
    Dim TImpAnul        As Currency
    Dim Imp_AnulAdm     As Currency
    Dim TImpAnulAdm     As Currency
    Dim Imp_AnulAcad    As Currency
    Dim TImpAnulAcad    As Currency

    Dim Imp_Cobr            As Currency
    Dim TImpCobr            As Currency
    Dim Imp_Emi_VS_Cobr     As Currency
    Dim TImpEmi_VS_Cobr     As Currency
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
    Const Lst_Imp_Anul          As Integer = 7
    Const Lst_Imp_AnulAdm       As Integer = 8
    Const Lst_Imp_AnulAcad      As Integer = 9
    Const Lst_Imp_Cobr          As Integer = 10
    Const Lst_Imp_CobrAdm       As Integer = 11
    Const Lst_Imp_CobrAcad      As Integer = 12
    Const Lst_Imp__RDT          As Integer = 13
    Const Lst_Imp__RDTAdm       As Integer = 14
    Const Lst_Imp__RDTAcad      As Integer = 15
    Const Lst_Imp_RDTpdt          As Integer = 16
    Const Lst_Imp_RDTpdtAdm       As Integer = 17
    Const Lst_Imp_RDTpdtAcad      As Integer = 18
    Const Lst_Imp__ADx          As Integer = 19
    Const Lst_Imp__ADxAdm       As Integer = 20
    Const Lst_Imp__ADxAcad      As Integer = 21
    Const Lst_Imp_PdtCob        As Integer = 22
    Const Lst_Imp_PdtCobAdm     As Integer = 23
    Const Lst_Imp_PdtCobAcad    As Integer = 24
    Const Lst_Imp_Emi_VS_Cobr   As Integer = 25

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
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisAdm).Value = "Emitido" & vbLf & "1303.00 " & vbLf & "Adm.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisAcad).Value = "Emitido" & vbLf & "1311.00 " & vbLf & "Acad.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Anul).Value = "Importe" & vbLf & "Anulado" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_AnulAdm).Value = "Anulado" & vbLf & "1303.00 " & vbLf & "Adm.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_AnulAcad).Value = "Anulado" & vbLf & "1311.00 " & vbLf & "Acad.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Cobr).Value = "Importe" & vbLf & "Cobrado" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDT).Value = "Imp_RDT" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDTAdm).Value = "Imp_RDT" & vbLf & "1303.00 " & vbLf & "Adm.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDTAcad).Value = "Imp_RDT" & vbLf & "1311.00 " & vbLf & "Acad.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_RDTpdt).Value = "Pdte_RDT"
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_RDTpdtAdm).Value = "Pdte_RDT" & vbLf & "1303.00 " & vbLf & "Adm.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_RDTpdtAcad).Value = "Pdte_RDT" & vbLf & "1311.00 " & vbLf & "Acad.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADx).Value = "Emi'" & Right(AñoCont, 2) & " Vto'" & Right(AñoCont + 1, 2) & vbLf & "¡ Sin Cobrar !" & vbLf & "ADxAplz'" & Right(AñoCont + 1, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADxAdm).Value = "Imp.ADxAplz" & vbLf & "1303.00 " & vbLf & "Adm.'" & Right(AñoCont + 1, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADxAcad).Value = "Imp.ADxAplz" & vbLf & "1311.00 " & vbLf & "Acad.'" & Right(AñoCont + 1, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCob).Value = "Emi'" & Right(AñoCont, 2) & " Vto'" & Right(AñoCont, 2) & vbLf & "Pdte_Cob" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCobAdm).Value = "Pdte_Cob" & vbLf & "1303.00 " & vbLf & "Adm.'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCobAcad).Value = "Pdte_Cob" & vbLf & "1311.00 " & vbLf & "Acad.'" & Right(AñoCont, 2)

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


'- Importe Recibos Emi_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_Emis = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            If Imp_Emis = 0 Then GoTo Siguiente_Plan    '- Si todo está cobrado en año anterior
            '- Inicio una nueva línea de datos !!! --->
            Cont_Plan = Cont_Plan + 1
            Set RowNew = Lo_Lst.ListRows.Add
                RowNew.Range(Lst_Orden) = Cont_Plan
                RowNew.Range(Lst_Plan) = Cod_Plan
            '-----------------------------------------<
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
            
'- Importe Anul --------------------------------------------------------------------------------------------------------------------
            Imp_Anul = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), "<0", _
                                          .Columns(BD_Plan), Cod_Plan)
            RowNew.Range(Lst_Imp_Anul) = Imp_Anul
            TImpAnul = TImpAnul + Imp_Anul
            If Imp_Anul <> 0 Then
                '- Importe-Adm ADxAplz --------------------------------------------------------------------------------------------------------------------
                Imp_AnulAdm = 0
                RowNew.Range(Lst_Imp_AnulAdm) = "?"
                TImpAnulAdm = TImpAnulAdm + Imp_AnulAdm
    
                '- Importe Acad ADxAplz -----------------------------------------------------------------------------------------------------------------
                Imp_AnulAcad = 0
                RowNew.Range(Lst_Imp_AnulAcad) = "?"
                TImpAnulAcad = TImpAnulAcad + Imp_AnulAcad
            End If
            
'- Importe Cob_AñoCont, Emi_AñoCont --------------------------------------------------------------------------------------------------------------------
            Imp_Cobr = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_Cobr) = Imp_Cobr
            TImpCobr = TImpCobr + Imp_Cobr
            
            '- Importe Emitido y Cob_AñoCont, Emi_AñoCont -------- ¡¡¡ Para detectar Diferencias entre Emitido y Cobrado ------------------------------------------------------------------------------------------------------------
            Imp_Emi_VS_Cobr = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_Emi_VS_Cobr) = Imp_Emi_VS_Cobr - Imp_Cobr
            TImpEmi_VS_Cobr = TImpEmi_VS_Cobr + Imp_Emi_VS_Cobr - Imp_Cobr
            
            '- Importe Adm Cob_AñoCont, Emi_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_CobrAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_CobrAdm) = Imp_CobrAdm
            TImpCobrAdm = TImpCobrAdm + Imp_CobrAdm
            
            '- Importe Acad Cob_AñoCont, Emi_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_CobrAcad = Imp_Cobr - Imp_CobrAdm
            RowNew.Range(Lst_Imp_CobrAcad) = Imp_CobrAcad
            TImpCobrAcad = TImpCobrAcad + Imp_CobrAcad
            
            
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
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "=", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_RDTpdt) = Imp_SRDT
            TImpSRDT = TImpSRDT + Imp_SRDT

            '- Importe Adm Pdte_RDT --------------------------------------------------------------------------------------------------------------------
            Imp_SRDTAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
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
                                          .Columns(BD_ACont_Cob), "=", _
                                          .Columns(BD_ACont_Emi), AñoCont) _
                     + Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont + 1, _
                                          .Columns(BD_ACont_Cob), "=" & AñoCont + 1, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp__ADx) = Imp__ADx
            TImp_ADx = TImp_ADx + Imp__ADx

            '- Importe ADxAplzAdm --------------------------------------------------------------------------------------------------------------------
            Imp_ADxAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont + 1, _
                                          .Columns(BD_ACont_Cob), "=", _
                                          .Columns(BD_ACont_Emi), AñoCont) _
                     + Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont + 1, _
                                          .Columns(BD_ACont_Cob), "=" & AñoCont + 1, _
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
                                          .Columns(BD_ACont_Emi), AñoCont) _
                     + Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont, _
                                          .Columns(BD_ACont_Cob), "=" & AñoCont + 1, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_PdtCob) = Imp_PdtCob
            TimpPdtCob = TimpPdtCob + Imp_PdtCob

            '- Importe Adm Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            Imp_PdtCobAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont, _
                                          .Columns(BD_ACont_Cob), "=", _
                                          .Columns(BD_ACont_Emi), AñoCont) _
                     + Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont, _
                                          .Columns(BD_ACont_Cob), "=" & AñoCont + 1, _
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

'==================================================================================================================================
Sub RuT_Cierre_Contable_Planes_CAcad_Ant(Lo_BD As ListObject, _
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
    
    Dim RegsEmis            As Long     ' regs Emitidos
    Dim Imp_Emis            As Currency
    Dim TImpEmis            As Currency
    Dim Imp_Emi_Y_Cobr     As Currency
    Dim TImpEmi_VS_Cobr     As Currency
    
    Dim RegsEmisAnt     As Long     ' regs Emitidos
    Dim Imp_EmisAnt     As Currency
    Dim TImpEmisAnt     As Currency
    
    Dim RegsEmisPos     As Long     ' regs Emitidos
    Dim Imp_EmisPos     As Currency
    Dim TImpEmisPos     As Currency
    Dim Imp_EmisPosAdm     As Currency
    Dim TImpEmisPosAdm     As Currency
    Dim Imp_EmisPosAcad    As Currency
    Dim TImpEmisPosAcad    As Currency

    Dim Imp_Anul        As Currency
    Dim TImpAnul        As Currency
    Dim Imp_AnulAdm     As Currency
    Dim TImpAnulAdm     As Currency
    Dim Imp_AnulAcad    As Currency
    Dim TImpAnulAcad    As Currency
    
    Dim Imp_CobrAnt        As Currency
    Dim TImpCobrAnt        As Currency
    
    Dim Imp_CobrPos        As Currency
    Dim TImpCobrPos        As Currency
    Dim Imp_CobrPosAdm     As Currency
    Dim TImpCobrPosAdm     As Currency
    Dim Imp_CobrPosAcad    As Currency
    Dim TImpCobrPosAcad    As Currency
    
    Dim Imp_Cob_C_Acad As Currency
    Dim TImpCob_C_Acad As Currency
    
    Dim Imp__RDT        As Currency
    Dim TImp_RDT        As Currency
    
    Dim Imp_SRDT        As Currency
    Dim TImpSRDT        As Currency
    
    Dim Imp__ADx        As Currency
    Dim TImp_ADx        As Currency
    Dim Imp__ADxAdm     As Currency
    Dim TImp_ADxAdm     As Currency
    Dim Imp__ADxAcad    As Currency
    Dim TImp_ADxAcad    As Currency

    Dim Imp_Aplz        As Currency
    Dim TImpAplz        As Currency

    Dim Imp_PdtCob        As Currency
    Dim TimpPdtCob        As Currency
    Dim Imp_PdtCobAdm     As Currency
    Dim TimpPdtCobAdm     As Currency
    Dim Imp_PdtCobAcad    As Currency
    Dim TimpPdtCobAcad    As Currency

    Dim R_EjeAnt        As Long     ' regs Pendiente de pago
    Dim I_EjeAnt        As Currency
    Dim T_EjeAnt        As Currency

    Const Lst_Orden             As Integer = 1
    Const Lst_Plan              As Integer = 2
    Const Lst_RegsEmis          As Integer = 3
    Const Lst_Imp_Emis          As Integer = 4
    Const Lst_Imp_EmisAnt       As Integer = 5
    Const Lst_ImpAnul           As Integer = 6
    Const Lst_ImpAnulAdm        As Integer = 7
    Const Lst_ImpAnulAcad       As Integer = 8
    Const Lst_Imp_ADx           As Integer = 9
    Const Lst_Imp_ADxAdm        As Integer = 10
    Const Lst_Imp_ADxAcad       As Integer = 11
    Const Lst_Imp_Aplz          As Integer = 12
    Const Lst_ImpEjeAnt         As Integer = 13
    Const Lst_Imp_EmisPos       As Integer = 14
    Const Lst_Imp_EmisPosAdm    As Integer = 15
    Const Lst_Imp_EmisPosAcad   As Integer = 16
    Const Lst_Imp_CobrAnt       As Integer = 17
    Const Lst_Imp_CobrPos       As Integer = 18
    Const Lst_Imp_CobrPosAdm    As Integer = 19
    Const Lst_Imp_CobrPosAcad   As Integer = 20
    Const Lst_Imp__RDT          As Integer = 21
    Const Lst_Imp_RDTpdt          As Integer = 22
    Const Lst_Imp_PdtCob        As Integer = 23
    Const Lst_Imp_PdtCobAdm     As Integer = 24
    Const Lst_Imp_PdtCobAcad    As Integer = 25
    Const Lst_Imp_Cob_C_Acad    As Integer = 26
    Const Lst_Imp_Emi_y_Cob     As Integer = 27
    Const Lst_Imp_Emi_VS_Cobr   As Integer = 28
    
    Dim Regs_Dev            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula
    Dim Imp_Devo            As Currency
    Dim RDev_Pag            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, Pagados
    Dim IDev_Pag            As Currency
    Dim RDevPdte            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, No Pagados, pero Ajuste Matrícula NO deben ser pagados
    Dim IDevPdte            As Currency
    
'''    Dim Ws_Lista    As Worksheet:       Set Ws_Lista = Sheets("Cierre_Planes_CAcad_Ant_" & TipoCurso)
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
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Emis).Value = "Contab." & AñoCont & vbLf & "Imp_Emi" & vbLf & CursoAcad
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisAnt).Value = "Imp_Emi" & vbLf & "en " & AñoCont - 1 & vbLf & "Conta'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_ImpAnul).Value = "Importes" & vbLf & "Anulados" & vbLf & CursoAcad
    Lo_Lst.HeaderRowRange.Cells(Lst_ImpAnulAdm).Value = "Imp.Anul" & vbLf & "Adm." & vbLf & "???"
    Lo_Lst.HeaderRowRange.Cells(Lst_ImpAnulAcad).Value = "Imp.Anul" & vbLf & "Acad." & vbLf & "???"
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_ADx).Value = "Imp_Emi'" & Right(AñoCont - 1, 2) & vbLf & "y ADxAplz" & vbLf & "en " & AñoCont - 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_ADxAdm).Value = "Imp_ADxAplz" & vbLf & "Adm. en" & vbLf & AñoCont - 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_ADxAcad).Value = "Imp_ADxAplz" & vbLf & "Acad. en" & vbLf & AñoCont - 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Aplz).Value = "ADxAplz'" & Right(AñoCont - 1, 2) & vbLf & "Cob " & AñoCont & vbLf & "Aplazado'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_ImpEjeAnt).Value = "Emi." & AñoCont - 1 & vbLf & "Rec. EjeAnt" & vbLf & "Cob " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisPos).Value = "Imp_Emi" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisPosAdm).Value = "Imp_Emi" & vbLf & "Adm. en" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisPosAcad).Value = "Imp_Emi" & vbLf & "Acad. en" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_CobrAnt).Value = "Imp.Cob. " & vbLf & AñoCont - 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_CobrPos).Value = "Imp_Cob" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_CobrPosAdm).Value = "Imp_Cob" & vbLf & "Adm." & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_CobrPosAcad).Value = "Imp_Cob" & vbLf & "Acad." & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDT).Value = "Imp_RDT" & vbLf & CursoAcad
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_RDTpdt).Value = "Pdte. RDT"
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCob).Value = "Pdte.Cob. " & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCobAdm).Value = "Pdte.Cob. " & vbLf & AñoCont & vbLf & "Adm."
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCobAcad).Value = "Pdte.Cob. " & vbLf & AñoCont & vbLf & "Acad."
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Cob_C_Acad).Value = "Imp.Cob. " & vbLf & "Total" & vbLf & CursoAcad
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Emi_y_Cob).Value = "Imp.Rec.Emi" & vbLf & "y Cob." & vbLf & CursoAcad


    '- Visualizo el progreso ---------------------------------------------------------------------------------------
    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TB_Informe = Txt_Cabecera & vbCrLf & vbCrLf
    Txt_Cabecera = String(16, " ") & AñoCont - 1 & String(11, " ") & AñoCont & String(11, " ") & AñoCont & "         Cob." & AñoCont & "       Cob." & AñoCont & vbLf & _
                   "    Plan       Imp_Emi        Imp_Emi        ADxAplz       Aplazado        Eje_Ant"
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera & vbLf
    
    
    '- Recorro la Collection con todos los Planes
         Dim Cod_Plan As Variant
    With Lo_BD.DataBodyRange
    For Each Cod_Plan In Collection_Planes
        
        Cont_Plan = Cont_Plan + 1
        Set RowNew = Lo_Lst.ListRows.Add
            RowNew.Range(Lst_Orden) = Cont_Plan
            RowNew.Range(Lst_Plan) = Cod_Plan
    
'- Importe Emi -----------------------------------------------------------------------------------------------------------------
            Imp_Emis = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan)
            RowNew.Range(Lst_Imp_Emis) = Imp_Emis
            TImpEmis = TImpEmis + Imp_Emis
            RegsEmis = Application.CountIfs( _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan)
            RowNew.Range(Lst_RegsEmis) = RegsEmis
    
'- Importe EmiAnt_AñoCont-1 -----------------------------------------------------------------------------------------------------------------
            Imp_EmisAnt = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Emi), AñoCont - 1)
            RowNew.Range(Lst_Imp_EmisAnt) = Imp_EmisAnt
            TImpEmisAnt = TImpEmisAnt + Imp_EmisAnt
    
'- Importe Anul --------------------------------------------------------------------------------------------------------------------
            Imp_Anul = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), "<0", _
                                          .Columns(BD_Plan), Cod_Plan)
            RowNew.Range(Lst_ImpAnul) = Imp_Anul
            TImpAnul = TImpAnul + Imp_Anul
            If Imp_Anul <> 0 Then
                '- Importe-Adm ADxAplz --------------------------------------------------------------------------------------------------------------------
                Imp_AnulAdm = 0
                RowNew.Range(Lst_ImpAnulAdm) = "?"
                TImpAnulAdm = TImpAnulAdm + Imp_AnulAdm
    
                '- Importe Acad ADxAplz -----------------------------------------------------------------------------------------------------------------
                Imp_AnulAcad = 0
                RowNew.Range(Lst_ImpAnulAcad) = "?"
                TImpAnulAcad = TImpAnulAcad + Imp_AnulAcad
            End If
            
'- Importe ADxAplz --------------------------------------------------------------------------------------------------------------------
            Imp__ADx = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont, _
                                          .Columns(BD_ACont_Cob), "<>" & AñoCont - 1, _
                                          .Columns(BD_ACont_Emi), AñoCont - 1)
            RowNew.Range(Lst_Imp_ADx) = Imp__ADx
            TImp_ADx = TImp_ADx + Imp__ADx

            '- Importe-Adm ADxAplz --------------------------------------------------------------------------------------------------------------------
            Imp__ADxAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont, _
                                          .Columns(BD_ACont_Cob), "<>" & AñoCont - 1, _
                                          .Columns(BD_ACont_Emi), AñoCont - 1)
            RowNew.Range(Lst_Imp_ADxAdm) = Imp__ADxAdm
            TImp_ADxAdm = TImp_ADxAdm + Imp__ADxAdm

            '- Importe Acad ADxAplz -----------------------------------------------------------------------------------------------------------------
            Imp__ADxAcad = Imp__ADx - Imp__ADxAdm
            RowNew.Range(Lst_Imp_ADxAcad) = Imp__ADxAcad
            TImp_ADxAcad = TImp_ADxAcad + Imp__ADxAcad
            
'- Importe Aplazado (ADxAplz el AñoCont -1 ----------------------------------------------------------------------------------------------------------------
            Imp_Aplz = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont, _
                                          .Columns(BD_ACont_Cob), "<>" & AñoCont - 1, _
                                          .Columns(BD_ACont_Emi), AñoCont - 1)
            RowNew.Range(Lst_Imp_Aplz) = Imp_Aplz
            TImpAplz = TImpAplz + Imp_Aplz

'- Importe EjeAnt ( Emitido el AñoCont-1, Vto AñoCont-1 y Cob Añocont ----------------------------------------------------------------------------------------------------------------
            I_EjeAnt = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Vto), AñoCont - 1, _
                                          .Columns(BD_ACont_Cob), "=" & AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont - 1)
            RowNew.Range(Lst_ImpEjeAnt) = I_EjeAnt
            T_EjeAnt = T_EjeAnt + I_EjeAnt

'- Importe EmiPos_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_EmisPos = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_EmisPos) = Imp_EmisPos
            TImpEmisPos = TImpEmisPos + Imp_EmisPos
    
            '- Importe  EmiPosAdm_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_EmisPosAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_EmisPosAdm) = Imp_EmisPosAdm
            TImpEmisPosAdm = TImpEmisPosAdm + Imp_EmisPosAdm
    
            '- Importe  EmiPosAdm_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_EmisPosAcad = Imp_EmisPos - Imp_EmisPosAdm
            RowNew.Range(Lst_Imp_EmisPosAcad) = Imp_EmisPosAcad
            TImpEmisPosAcad = TImpEmisPosAcad + Imp_EmisPosAcad

'- Importe CobrAnt Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_CobrAnt = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont - 1)
            RowNew.Range(Lst_Imp_CobrAnt) = Imp_CobrAnt
            TImpCobrAnt = TImpCobrAnt + Imp_CobrAnt
    
'- Importe CobrPos Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_CobrPos = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont)
            RowNew.Range(Lst_Imp_CobrPos) = Imp_CobrPos
            TImpCobrPos = TImpCobrPos + Imp_CobrPos
    
            '- Importe Adm Recibos Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_CobrPosAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont)
            RowNew.Range(Lst_Imp_CobrPosAdm) = Imp_CobrPosAdm
            TImpCobrPosAdm = TImpCobrPosAdm + Imp_CobrPosAdm
    
            '- Importe Acad Recibos Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_CobrPosAcad = Imp_CobrPos - Imp_CobrPosAdm
            RowNew.Range(Lst_Imp_CobrPosAcad) = Imp_CobrPosAcad
            TImpCobrPosAcad = TImpCobrPosAcad + Imp_CobrPosAcad
            
'- Importe Recibos RDT -----------------------------------------------------------------------------------------------------------------
            Imp__RDT = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "<>")
            RowNew.Range(Lst_Imp__RDT) = Imp__RDT
            TImp_RDT = TImp_RDT + Imp__RDT

'- Importe Sin_RDT --------------------------------------------------------------------------------------------------------------------
            Imp_SRDT = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "=")
            RowNew.Range(Lst_Imp_RDTpdt) = Imp_SRDT
            TImpSRDT = TImpSRDT + Imp_SRDT

'- Importe Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            Imp_PdtCob = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), "=")
            RowNew.Range(Lst_Imp_PdtCob) = Imp_PdtCob
            TimpPdtCob = TimpPdtCob + Imp_PdtCob
    
            '- Importe Adm Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            Imp_PdtCobAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), "=")
            RowNew.Range(Lst_Imp_PdtCobAdm) = Imp_PdtCobAdm
            TimpPdtCobAdm = TimpPdtCobAdm + Imp_PdtCobAdm
            '- Importe Acad Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            Imp_PdtCobAcad = Imp_PdtCob - Imp_PdtCobAdm
            RowNew.Range(Lst_Imp_PdtCobAcad) = Imp_PdtCobAcad
            TimpPdtCobAcad = TimpPdtCobAcad + Imp_PdtCobAcad
    
'- Importe Cob_C_Acad -----------------------------------------------------------------------------------------------------------------
            Imp_Cob_C_Acad = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan)
            RowNew.Range(Lst_Imp_Cob_C_Acad) = Imp_Cob_C_Acad
            TImpCob_C_Acad = TImpCob_C_Acad + Imp_Cob_C_Acad
    
            '- Importe Emitido y Cob_AñoCont, Emi_AñoCont -------- ¡¡¡ Para detectar Diferencias entre Emitido y Cobrado ------------------------------------------------------------------------------------------------------------
            Imp_Emi_Y_Cobr = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan)
            RowNew.Range(Lst_Imp_Emi_y_Cob) = Imp_Emi_Y_Cobr
            RowNew.Range(Lst_Imp_Emi_VS_Cobr) = Imp_Emi_Y_Cobr - Imp_Cob_C_Acad
            TImpEmi_VS_Cobr = TImpEmi_VS_Cobr + Imp_Emi_Y_Cobr - Imp_Cob_C_Acad
            
''''- Informe --------------------------------------------------------------------------------------------------------------------
'''            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Cont_Plan, "00") & "º " & Cod_Plan & " " & _
'''                Right(String(12, " ") & Format(Imp_EmisAnt, "#,##0.00"), 14) & _
'''                IIf(Imp_Cobr > 0, " " & Right(String(12, "·") & Format(Imp_EmisPos, "#,##0.00"), 14), "  No hay Cobros") & _
'''                IIf(Imp__RDT > 0, " " & Right(String(12, "·") & Format(Imp__ADx, "#,##0.00"), 14), " " & String(14, "·")) & _
'''                IIf(Imp_SRDT > 0, " " & Right(String(12, "·") & Format(Imp_Aplz, "#,##0.00"), 14), " " & String(14, "·")) & _
'''                IIf(Imp_PdtCob > 0, " " & Right(String(12, "·") & Format(I_EjeAnt, "#,##0.00"), 14), " " & String(14, "·")) & vbLf
'''
'''            ' Intercalo Cabecera Columnas
'''            If Cont_Plan Mod 20 = 0 Then Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Txt_Cabecera & vbLf
'''
'''            If Imp_Cobr = 0 Then PlanesSinCob = PlanesSinCob + 1
Siguiente_Plan:
    Next Cod_Plan
    End With    ' Lo_BD.DataBodyRange
    
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Txt_Cabecera
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & " Totales "
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                Right(String(12, " ") & Format(TImpEmisAnt, "#,##0.00"), 14) & _
                Right(String(12, " ") & Format(TImpEmisPos, "#,##0.00"), 15) & _
                Right(String(12, " ") & Format(TImp_ADx, "#,##0.00"), 15) & _
                Right(String(12, " ") & Format(TImpAplz, "#,##0.00"), 15) & _
                Right(String(12, " ") & Format(T_EjeAnt, "#,##0.00"), 15) & vbLf


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
'''Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Listar_Planes   --------------------------------------------------------------------------------------------
'===================================================================================================================================
