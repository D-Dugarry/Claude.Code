Attribute VB_Name = "M22_Inf_EPs_para_UXXI_NEW"
'2026-02-05
Option Explicit

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

'==================================================================================================================================
Sub Rut_Informe_EPs_para_UXXI()
'==================================================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim CursoAcad       As String:      CursoAcad = Prog__APP.Range("APP_CursAcad")
    Dim CursoAcadAnt    As String:      CursoAcadAnt = Prog__APP.Range("APP_C_Acad_Ant")
    Dim TipoCurso       As String:      TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")

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
    Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="<>_ERR_Date_"
    Lo_BD.Range.AutoFilter Field:=BD_ACont_Emi, Criteria1:="<=" & AñoCont   '- No contabilizamos Recibos emitidos Futuros...
    Lo_BD.Range.AutoFilter Field:=BD_ImpAdm, Criteria1:=">=0"                '- Ajustes de matrícula que distorcionan la contabilidad
    Lo_BD.Range.AutoFilter Field:=BD_ImpRec, Criteria1:=">0"                '- No se requieren los recibos de devoluciones de tasas
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
    On Error GoTo 0
        Call Rut_Informe_EPs_para_UXXI_Rellena_Tabla(Lo_BD_Filtrada, CursoAcad, TipoCurso, Sht__Inf_EPs_UXXI)
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

' ==================================================================================================================================
Sub Rut_Informe_EPs_para_UXXI_Rellena_Tabla(Lo_BD As ListObject, _
                                            CursoAcad As String, _
                                            TipoCurso As String, _
                                            Ws_Lista As Worksheet)
' ==================================================================================================================================
'--- Tabla Sht__Inf_EPs_UXXI  EPs-Resumen ----------------------
    'Dim CursoAcad       As String:     CursoAcad = Prog__APP.Range("APP_CursAcad") '- Curso Académico
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")    '- Año Contable
    Dim ACont           As Integer:     ACont = Right(AñoCont, 2)
    Dim AñoContAnt      As Integer:     AñoContAnt = Left(CursoAcad, 4)             '- El 1º Año de Curso
    Dim AContAnt        As Integer:     AContAnt = Right(AñoContAnt, 2)             '- El 1º Año de Curso Corto
    Dim AñoContPos      As Integer:     AñoContPos = "20" & Right(CursoAcad, 2)     '- El 2º Año de Curso
    Dim AContPos        As Integer:     AContPos = Right(AñoContPos, 2)             '- El 2º Año de Curso Corto

    
    Dim Fila_DR             As Long
    Dim F_Plan              As Long
    Dim F_Ant               As Long
    
    Dim Cod_Plan        As String
    Dim Curso_Acad_Ant  As String
    Dim Año_Emi     As String
    
    Dim Cont_Tot_Reg        As Long
    Dim Cont_Reg_Proc       As Long
    
    Dim R_Emi_Acad          As Currency
    Dim R_Cob_C_Acad_Adm    As Currency
    Dim R_Cob_Acont         As Currency
    Dim R_Rdt_Acad          As Currency
    Dim R_ADx_Acad          As Currency
    
    Dim ProgresoTarea       As String
    
    Dim Lo_Lst      As ListObject:       Set Lo_Lst = Ws_Lista.ListObjects(1)
    Dim RowNew      As ListRow
    Dim Row_BD      As ListRow

    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    Rut_Off_Functions
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Ws_Lista.Unprotect
    If Not Lo_Lst.DataBodyRange Is Nothing Then Lo_Lst.DataBodyRange.Delete

    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
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

    '- Ordenar por PLAN y DNI ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending)
    Call Rut_Lo_Sort(Lo_BD, BD_FEmi, xlAscending)
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending)
    '- Preparar Tabla de Ws_Lista ==================
    Ws_Lista.Visible = xlSheetVisible
    Ws_Lista.Select
    Call Rut_Lo_WrkSht_Preparar(Ws_Lista)
    Ws_Lista.Unprotect
    '- Vacío la Tabla de Tit.Prop.  =====================================
    Lo_Lst.AutoFilter.ShowAllData
    If Not Lo_Lst.DataBodyRange Is Nothing Then Lo_Lst.DataBodyRange.Delete
    ' ==================================================================================================================================
    ' ###############################  Genero la Tabla de Planes de BDatos  #####################################
    Cod_Plan = "":   Curso_Acad_Ant = "":   Año_Emi = 0: Cont_Tot_Reg = 0:
    Range("TP_Cod_Plan") = ""
    Range("TP_Cod_Plan").Select
    ActiveCell.Offset(0, 1) = CursoAcad
    ActiveCell.Offset(0, -1) = "Enseñanzas de:  " & Range("APP_EFP_o_CFC")
    ActiveCell.Offset(0, 2) = " Última actualización: " & Now()
    ActiveCell.Offset(0, 1).Select
    
    '- Datos Cabecera Tabla ------------------------------
    Lo_Lst.HeaderRowRange.Cells(BdUx_Emi_Total).Offset(-1, 0).Value = "Año Contable: " & AñoContAnt & " / " & AñoContPos
    Lo_Lst.HeaderRowRange.Cells(BdUx_Cob_C_ACad).Offset(-1, 0).Value = "Año Cont: " & AñoContAnt & " / " & AñoContPos
    Lo_Lst.HeaderRowRange.Cells(BdUx_Cob_ACont).Offset(-1, 0).Value = "Año Contable: " & AñoContAnt & " / " & AñoContPos
    Lo_Lst.HeaderRowRange.Cells(BdUx_RDT_Total).Offset(-1, 0).Value = "Curso Acad: " & CursoAcad
    Lo_Lst.HeaderRowRange.Cells(BdUx_PdteRDT_Total).Offset(-1, 0).Value = "Curso Acad: " & CursoAcad
    Lo_Lst.HeaderRowRange.Cells(BdUx_ADx_Total).Offset(-1, 0).Value = "Emitido'" & AContAnt & ", Vencimiento'" & AContPos & " y NO Cobrado'" & AContAnt
    Lo_Lst.HeaderRowRange.Cells(BdUx_Aplz_Total).Offset(-1, 0).Value = "Emitido'" & AContAnt & ", ADxAplz'" & AContAnt & " y Cobrado'" & AContPos
    Lo_Lst.HeaderRowRange.Cells(BdUx_Pdt_Total).Offset(-1, 0).Value = "Pendiente de Cobro Curso Acad: " & CursoAcad
            
            
    Debug.Print ""
    ' ##################################################################################################################
    ' =============  Recorrer todos los Registros filtrados y Crear la Tabla de Planes de DR  =====================================
Dim RwBD        As ListRow
Dim RwUxi        As ListRow

    With Lo_BD.DataBodyRange
    For Fila_DR = 1 To Lo_BD.ListRows.Count
    
        Set RwBD = Nothing
        Set RwBD = Lo_BD.ListRows(Fila_DR)
        
        Cont_Reg_Proc = Cont_Reg_Proc + 1
        Cont_Tot_Reg = Cont_Tot_Reg + 1
        ' --------------=============  Tratamiento de los Datos  ==================
        'If Cod_Plan & Curso_Acad_Ant & Año_Emi = RwBD.Range(BD_Plan) & RwBD.Range(BD_C_Acad) & Format(RwBD.Range(BD_FEmi), "yyyy") Then    ' ------- Control cambio de Plan de estudio y de Curso Académico ---------------
        If Cod_Plan & Año_Emi = RwBD.Range(BD_Plan) & RwBD.Range(BD_ACont_Emi) Then GoTo Siguiente_Fila    ' ------- Control cambio de Plan de estudio y de Curso Académico ---------------
           
        '------ Es el primero de una serie y tengo que introducir los datos comunes ---------------------------------------------------------

        Set RwUxi = Nothing
        Set RwUxi = Lo_Lst.ListRows.Add
        F_Plan = Lo_Lst.ListRows.Count
       
        Cod_Plan = RwBD.Range(BD_Plan)
        Año_Emi = RwBD.Range(BD_ACont_Emi)
        
        RwUxi.Range(BdUx_Cod_Plan) = RwBD.Range(BD_Plan)
        RwUxi.Range(BdUx_Curso_Acad) = RwBD.Range(BD_C_Acad)
        RwUxi.Range(BdUx_Año_Emi) = RwBD.Range(BD_ACont_Emi)
        RwUxi.Range(BdUx_Plan_Curso) = RwBD.Range(BD_Plan) & "_" & RwBD.Range(BD_C_Acad) & "_Cont" & Right(RwBD.Range(BD_FEmi), 2)
        
        RwUxi.Range(BdUx_NomPlan) = RwBD.Range(BD_NomPlan)
        
        RwUxi.Range(BdUx_Ref_JI) = RwBD.Range(BD_JI_Emi_Acad)
        RwUxi.Range(BdUx_Orgánica) = RwBD.Range(BD_Orgánica)
        
        If IsNumeric(Left(RwBD.Range(BD_Plan), 1)) Then RwUxi.Range(BdUx_Concepto) = 1311 Else RwUxi.Range(BdUx_Concepto) = 1311.03

        RwUxi.Range(BdUx_Cant_Reg) = 1
        RwUxi.Range(BdUx_Coef_VRI) = RwBD.Range(BD_Coef_VRI)
        
Dim Imp_Acad        As Currency

'- Importe Emis C_Acad -----------------------------------------------------------------------------------------------------------------
        RwUxi.Range(BdUx_Emi_Total) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                         .Columns(BD_ACont_Emi), "=" & Año_Emi, _
                                                         .Columns(BD_Plan), Cod_Plan)
        '- Importe-EmisAdm --------------------------------------------------------------------------------------------------------------------
        RwUxi.Range(BdUx_Emi_Tadm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                        .Columns(BD_ACont_Emi), "=" & Año_Emi, _
                                                        .Columns(BD_Plan), Cod_Plan)
        '----- Reparto Emitido -------------------------------------------
        Imp_Acad = RwUxi.Range(BdUx_Emi_Total) - RwUxi.Range(BdUx_Emi_Tadm)
        RwUxi.Range(BdUx_Emi_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
        RwUxi.Range(BdUx_Emi_Org) = Imp_Acad - RwUxi.Range(BdUx_Emi_VRI)
        If RwUxi.Range(BdUx_Orgánica) = "" Then RwUxi.Range(BdUx_Emi_SinOrg) = RwUxi.Range(BdUx_Emi_Total) + RwUxi.Range(BdUx_Emi_Tadm)
                    
'- Importe Cobr C_Acad -----------------------------------------------------------------------------------------------------------------
        RwUxi.Range(BdUx_Cob_C_ACad) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                          .Columns(BD_ACont_Cob), "<>", _
                                                          .Columns(BD_Plan), Cod_Plan)
        '- Importe Adm Recibos Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
        RwUxi.Range(BdUx_Cob_C_Acad_Adm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                          .Columns(BD_ACont_Cob), "<>", _
                                                          .Columns(BD_Plan), Cod_Plan)

'- Importe Cobr AñoCont -----------------------------------------------------------------------------------------------------------------
        If Año_Emi = AñoContPos Then
            RwUxi.Range(BdUx_Cob_ACont) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                              .Columns(BD_ACont_Cob), "=" & AñoContPos, _
                                                              .Columns(BD_Plan), Cod_Plan)
            '- Importe Adm Recibos Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
            RwUxi.Range(BdUx_Cob_Tadm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                              .Columns(BD_ACont_Cob), "=" & AñoContPos, _
                                                              .Columns(BD_Plan), Cod_Plan)
            '----- Reparto Cobr Año_Cont -------------------------------------------
            Imp_Acad = RwUxi.Range(BdUx_Cob_ACont) - RwUxi.Range(BdUx_Cob_Tadm)
            RwUxi.Range(BdUx_Cob_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
            RwUxi.Range(BdUx_Cob_Org) = Imp_Acad - RwUxi.Range(BdUx_Cob_VRI)
            If RwUxi.Range(BdUx_Orgánica) = "" Then RwUxi.Range(BdUx_Cob_SinOrg) = RwUxi.Range(BdUx_Cob_ACont) + RwUxi.Range(BdUx_Cob_Tadm)
        End If
        
        If Año_Emi = AñoContAnt Then
            RwUxi.Range(BdUx_Cob_ACont) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                              .Columns(BD_ACont_Cob), "=" & AñoContAnt, _
                                                              .Columns(BD_Plan), Cod_Plan)
            '- Importe Adm Recibos Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
            RwUxi.Range(BdUx_Cob_Tadm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                              .Columns(BD_ACont_Cob), "=" & AñoContAnt, _
                                                              .Columns(BD_Plan), Cod_Plan)
            '----- Reparto Cobr Año_Cont -------------------------------------------
            Imp_Acad = RwUxi.Range(BdUx_Cob_ACont) - RwUxi.Range(BdUx_Cob_Tadm)
            RwUxi.Range(BdUx_Cob_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
            RwUxi.Range(BdUx_Cob_Org) = Imp_Acad - RwUxi.Range(BdUx_Cob_VRI)
            If RwUxi.Range(BdUx_Orgánica) = "" Then RwUxi.Range(BdUx_Cob_SinOrg) = RwUxi.Range(BdUx_Cob_ACont) + RwUxi.Range(BdUx_Cob_Tadm)
        End If
    
'- Importe RDT -----------------------------------------------------------------------------------------------------------------
        If AñoCont = AñoContPos And Año_Emi = AñoContPos Or AñoCont = AñoContAnt And Año_Emi = AñoContAnt Then
            RwUxi.Range(BdUx_RDT_Total) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                             .Columns(BD_RDT), "<>", _
                                                             .Columns(BD_Plan), Cod_Plan)
            '- Importe Adm RDT -----------------------------------------------------------------------------------------------------------------
            RwUxi.Range(BdUx_RDT_Tadm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                              .Columns(BD_RDT), "<>", _
                                                              .Columns(BD_Plan), Cod_Plan)
            '----- Reparto RDT -------------------------------------------
            Imp_Acad = RwUxi.Range(BdUx_RDT_Total) - RwUxi.Range(BdUx_RDT_Tadm)
            RwUxi.Range(BdUx_RDT_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
            RwUxi.Range(BdUx_RDT_Org) = Imp_Acad - RwUxi.Range(BdUx_RDT_VRI)
        End If
    
'- Importe Pdte RDT -----------------------------------------------------------------------------------------------------------------
        If AñoCont = AñoContPos And Año_Emi = AñoContPos Or AñoCont = AñoContAnt And Año_Emi = AñoContAnt Then
            RwUxi.Range(BdUx_PdteRDT_Total) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                                 .Columns(BD_ACont_Cob), "<>", _
                                                                 .Columns(BD_Plan), Cod_Plan)
            RwUxi.Range(BdUx_PdteRDT_Total) = RwUxi.Range(BdUx_PdteRDT_Total) - RwUxi.Range(BdUx_RDT_Total)
            '- Importe Adm Pdte_RDT -----------------------------------------------------------------------------------------------------------------
            RwUxi.Range(BdUx_PdteRDT_TAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                .Columns(BD_ACont_Cob), "<>", _
                                                                .Columns(BD_Plan), Cod_Plan)
            RwUxi.Range(BdUx_PdteRDT_TAdm) = RwUxi.Range(BdUx_PdteRDT_TAdm) - RwUxi.Range(BdUx_RDT_Tadm)
            '----- Reparto Pdte_RDT -------------------------------------------
            Imp_Acad = RwUxi.Range(BdUx_PdteRDT_Total) - RwUxi.Range(BdUx_PdteRDT_TAdm)
            RwUxi.Range(BdUx_PdteRDT_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
            RwUxi.Range(BdUx_PdteRDT_Org) = Imp_Acad - RwUxi.Range(BdUx_PdteRDT_VRI)
        End If
'- Importe ADxAplz -----------------------------------------------------------------------------------------------------------------
        If Año_Emi = AñoContAnt Then
            RwUxi.Range(BdUx_ADx_Total) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                             .Columns(BD_ACont_Cob), "<>" & AñoContAnt, _
                                                             .Columns(BD_ACont_Emi), "=" & AñoContAnt, _
                                                             .Columns(BD_ACont_Vto), ">" & AñoContAnt, _
                                                             .Columns(BD_Plan), Cod_Plan)
            '- Importe Adm ADxAplz -----------------------------------------------------------------------------------------------------------------
            RwUxi.Range(BdUx_ADx_Tadm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                             .Columns(BD_ACont_Cob), "<>" & AñoContAnt, _
                                                             .Columns(BD_ACont_Emi), "=" & AñoContAnt, _
                                                             .Columns(BD_ACont_Vto), ">" & AñoContAnt, _
                                                            .Columns(BD_Plan), Cod_Plan)
            '----- Reparto ADxAplz -------------------------------------------
            If RwUxi.Range(BdUx_ADx_Tadm) < 0 Then
                Imp_Acad = RwUxi.Range(BdUx_ADx_Total)
            Else
                Imp_Acad = RwUxi.Range(BdUx_ADx_Total) - RwUxi.Range(BdUx_ADx_Tadm)
            End If
            RwUxi.Range(BdUx_ADx_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
            RwUxi.Range(BdUx_ADx_Org) = Imp_Acad - RwUxi.Range(BdUx_ADx_VRI)
        End If
    
''- Importe ADxAplz -----------------------------------------------------------------------------------------------------------------
'            RwUxi.Range(BdUx_ADx_Total) = Application.SumIfs(.Columns(BD_ImpRec), _
'                                                             .Columns(BD_ACont_Cob), "<>" & AñoContAnt, _
'                                                             .Columns(BD_ACont_Emi), "=" & Año_Emi, _
'                                                             .Columns(BD_ACont_Vto), ">" & Año_Emi, _
'                                                             .Columns(BD_Plan), Cod_Plan)
'            '- Importe Adm ADxAplz -----------------------------------------------------------------------------------------------------------------
'            RwUxi.Range(BdUx_ADx_Tadm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
'                                                            .Columns(BD_ACont_Cob), "<>" & AñoContAnt, _
'                                                            .Columns(BD_ACont_Emi), "=" & Año_Emi, _
'                                                            .Columns(BD_ACont_Vto), ">" & Año_Emi, _
'                                                            .Columns(BD_Plan), Cod_Plan)
'            '----- Reparto ADxAplz -------------------------------------------
'            Imp_Acad = RwUxi.Range(BdUx_ADx_Total) - RwUxi.Range(BdUx_ADx_Tadm)
'            RwUxi.Range(BdUx_ADx_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
'            RwUxi.Range(BdUx_ADx_Org) = Imp_Acad - RwUxi.Range(BdUx_ADx_VRI)

'- Importe Aplazado -----------------------------------------------------------------------------------------------------------------
        If Año_Emi = AñoContPos Then
           RwUxi.Range(BdUx_Aplz_Total) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                              .Columns(BD_ACont_Cob), "=" & AñoContPos, _
                                                              .Columns(BD_ACont_Emi), "=" & AñoContAnt, _
                                                              .Columns(BD_ACont_Vto), ">" & AñoContAnt, _
                                                              .Columns(BD_Plan), Cod_Plan)
            '- Importe Adm Aplazado -----------------------------------------------------------------------------------------------------------------
            RwUxi.Range(BdUx_Aplz_Tadm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                              .Columns(BD_ACont_Cob), "=" & AñoContPos, _
                                                              .Columns(BD_ACont_Emi), "=" & AñoContAnt, _
                                                              .Columns(BD_ACont_Vto), ">" & AñoContAnt, _
                                                              .Columns(BD_Plan), Cod_Plan)
            '----- Reparto Aplazado -------------------------------------------
            Imp_Acad = RwUxi.Range(BdUx_Aplz_Total) - RwUxi.Range(BdUx_Aplz_Tadm)
            RwUxi.Range(BdUx_Aplz_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
            RwUxi.Range(BdUx_Aplz_Org) = Imp_Acad - RwUxi.Range(BdUx_Aplz_VRI)
        End If
    
'- Importe Pdte_Cob C_Acad -----------------------------------------------------------------------------------------------------------------
'        If Año_Emi = AñoContPos Then
        If AñoCont = AñoContPos And Año_Emi = AñoContPos Or AñoCont = AñoContAnt And Año_Emi = AñoContAnt Then
            RwUxi.Range(BdUx_Pdt_Total) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                             .Columns(BD_Plan), Cod_Plan)
            RwUxi.Range(BdUx_Pdt_Total) = RwUxi.Range(BdUx_Pdt_Total) - Application.SumIfs(.Columns(BD_ImpCob), _
                                                                                           .Columns(BD_ACont_Cob), "<>", _
                                                                                           .Columns(BD_Plan), Cod_Plan)
            '- Importe Adm Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            RwUxi.Range(BdUx_Pdt_Tadm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                            .Columns(BD_Plan), Cod_Plan)
            RwUxi.Range(BdUx_Pdt_Tadm) = RwUxi.Range(BdUx_Pdt_Tadm) - Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                                         .Columns(BD_ACont_Cob), "<>", _
                                                                                         .Columns(BD_Plan), Cod_Plan)
            '----- Reparto Emitido -------------------------------------------
            Imp_Acad = RwUxi.Range(BdUx_Pdt_Total) - RwUxi.Range(BdUx_Pdt_Tadm)
            RwUxi.Range(BdUx_Pdt_VRI) = Application.Round(Imp_Acad * RwUxi.Range(BdUx_Coef_VRI) / 100, 2)
            RwUxi.Range(BdUx_Pdt_Org) = Imp_Acad - RwUxi.Range(BdUx_Pdt_VRI)
            If RwUxi.Range(BdUx_Orgánica) = "" Then RwUxi.Range(BdUx_Pdt_SinOrg) = RwUxi.Range(BdUx_Pdt_Total) + RwUxi.Range(BdUx_Pdt_Tadm)
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
            
Siguiente_Fila:

        Set RwBD = Nothing

    Next Fila_DR
    End With
    Range("h1") = " Última actualización: " & Now()
   
Restablecer_Valores:
Call Rut_EnableEvents_Status_Reset

'MsgBox "Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & "    Registros: " & Cont_Tot_Reg & "    Planes: " & F_Plan

Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
        "En la Nueva Consulta hay:  " & "  -  Tot.Reg. " & Cont_Tot_Reg & "    Planes: " & F_Plan & vbCrLf & Now()
        
Rut_On_Functions
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Call Rut_EnableEvents_Status_Reset
'    Application.Speech.Speak "Proceso completado."
End Sub     ' Rut_Resumen_Tab_TitPropios  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
            

'2026-02-02
'===================================================================================================================================
'- Guarda Copia de la ActiveSheet ==================================================================================================
'===================================================================================================================================
Sub Rut_WrkSht_Export_Inf_para_UXXI()
    Debug.Print "Rut_WrkSht_Export_Cierre_Contable_AñoCont"
    Dim TipoEP          As String
    Dim FichName        As String
    
    '--- Elegir prefijo según APP ---
    If Prog__APP.Range("APP_EFP_o_CFC") = "EFP" Then
        TipoEP = "EFP_"
    Else
        TipoEP = "CFCyAFC_"
    End If
    
    ' Nombre sugerido: Archiv + resto
    FichName = "InfpaUXXI_" & TipoEP & Prog__APP.Range("APP_CursAcad") _
               & "_Cierre_2025 " & Format(Now, "(yyyy-mm-dd_hhmm)") & ".xlsx"

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    Call Rut_WrkSheet_Export_To_xlsx(ActiveSheet, FichName)
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    
    
'    Call Rut_EnableEvents_Status_Reset
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------










