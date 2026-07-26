Attribute VB_Name = "M_602_Cierre_Contable_PLANES"
'2026-01-27
'- M31_Cierre_Contable_PLANES
Option Explicit

Sub EjemploInputBox4Opciones()
    Dim opcion As String
    opcion = InputBox("Elija 1, 2, 3 o 4:", "Selección", "1")
    
    Select Case Val(opcion)
        Case 1: MsgBox "Opción A"
        Case 2: MsgBox "Opción B"
        Case 3: MsgBox "Opción C"
        Case 4: MsgBox "Opción D"
        Case Else: MsgBox "Opción inválida"
    End Select
End Sub



'==================================================================================================================================
Sub RuT_Cierre_Contable_Planes_CAcad_Ant_y_Pos()
'==================================================================================================================================
    Dim CursoAcad       As String:      CursoAcad = Prog__APP.Range("APP_C_Acad_Pos")
    Dim CursoAcadAnt    As String:      CursoAcadAnt = Prog__APP.Range("APP_C_Acad_Ant")

    Dim RngVisible  As Range
    Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_BD_Filtrada  As ListObject
                
    Dim WsBuffer As Worksheet: Set WsBuffer = ThisWorkbook.Worksheets("Sheet_Buffer")  'hoja fija/oculta
    ' Limpiar anterior tabla en Buffer
    If WsBuffer.ListObjects.Count > 0 Then WsBuffer.ListObjects(1).Delete

    Application.ScreenUpdating = False
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo SalirLimpiando
    
    Dim opcion As String
    opcion = InputBox("Elija:" & vbLf & _
                                "1 EFP 2024-25" & vbLf & _
                                "2 EFP 2025-26" & vbLf & _
                                "3 CFCyAFC 2024-25" & vbLf & _
                                "4 CFCyAFC 2025-26", _
                                "Selección Informe Cierre")
    Select Case Val(opcion)
        Case 1: GoTo Ciclo_1
        Case 2: GoTo Ciclo_2
        Case 3: GoTo Ciclo_3
        Case 4: GoTo Ciclo_4
        Case Else: MsgBox "Opción inválida"
    End Select
    GoTo Salir

Ciclo_1:
    ' 1) =============  IIº Ciclo EFP 2024-25 =================================================================
    Sht__BD.Unprotect:     Lo_BD.ShowTotals = False
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    ' 2) Ordenaciones
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_ACont_Emi, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    
    ' 3) Filtros
    Lo_BD.Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & CursoAcadAnt
    Lo_BD.Range.AutoFilter Field:=BD_Concepto, Criteria1:="=1311"
    Lo_BD.Range.AutoFilter Field:=BD_Anul, Criteria1:="<>S"
    Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="<>_Devol_", Operator:=xlAnd, Criteria2:="<>_ERR_Date_"
    ' 4) Rango visible filtrado (incluye cabeceras)
    On Error Resume Next
    Set RngVisible = Lo_BD.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo SalirLimpiando
    If RngVisible Is Nothing Then GoTo SalirLimpiando
    ' Limpia cualquier tabla previa en Buffer
    If WsBuffer.ListObjects.Count > 0 Then
        WsBuffer.ListObjects(1).Unlist   'o .Delete si prefieres quitar formato tabla
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
    Call RuT_Cierre_Contable_Planes_CAcad_Ant(Lo_BD_Filtrada, CursoAcadAnt, "EFP")
    GoTo Salir
    
Ciclo_2:
    ' 1) =============  Iº Ciclo EFP 2025-26 =================================================================
    Sht__BD.Unprotect:     Lo_BD.ShowTotals = False
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    ' 2) Ordenaciones
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_ACont_Emi, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    
    ' 3) Filtros
    Lo_BD.Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & CursoAcad
    Lo_BD.Range.AutoFilter Field:=BD_Concepto, Criteria1:="=1311"
    Lo_BD.Range.AutoFilter Field:=BD_Anul, Criteria1:="<>S"        ' se debería haber puesto y se escapó!!!
    Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="<>_Devol_", Operator:=xlAnd, Criteria2:="<>_ERR_Date_"
    ' 4) Rango visible filtrado (incluye cabeceras)
    On Error Resume Next
    Set RngVisible = Lo_BD.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo SalirLimpiando
    If RngVisible Is Nothing Then GoTo SalirLimpiando
    ' Limpia cualquier tabla previa en Buffer
    If WsBuffer.ListObjects.Count > 0 Then
        WsBuffer.ListObjects(1).Unlist   'o .Delete si prefieres quitar formato tabla
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
    Call RuT_Cierre_Contable_Planes_CAcad_Pos(Lo_BD_Filtrada, CursoAcad, "EFP")
    GoTo Salir
    
Ciclo_3:
    ' 1) =============  IIº Ciclo CFCyAFC 2024-25 =================================================================
    Sht__BD.Unprotect:     Lo_BD.ShowTotals = False
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    ' 2) Ordenaciones
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_ACont_Emi, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    
    ' 3) Filtros
    Lo_BD.Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & CursoAcadAnt
    Lo_BD.Range.AutoFilter Field:=BD_Concepto, Criteria1:="=1311,03"
    Lo_BD.Range.AutoFilter Field:=BD_Anul, Criteria1:="<>S"
    Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="<>_Devol_", Operator:=xlAnd, Criteria2:="<>_ERR_Date_"
    ' 4) Rango visible filtrado (incluye cabeceras)
    On Error Resume Next
    Set RngVisible = Lo_BD.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo SalirLimpiando
    If RngVisible Is Nothing Then GoTo SalirLimpiando
    ' Limpia cualquier tabla previa en Buffer
    If WsBuffer.ListObjects.Count > 0 Then
        WsBuffer.ListObjects(1).Unlist   'o .Delete si prefieres quitar formato tabla
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
    Call RuT_Cierre_Contable_Planes_CAcad_Ant(Lo_BD_Filtrada, CursoAcadAnt, "CFCyAFC")
    GoTo Salir
    
Ciclo_4:
    ' 1) =============  Iº Ciclo CFCyAFC 2025-26 =================================================================
    Sht__BD.Unprotect:     Lo_BD.ShowTotals = False
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    ' 2) Ordenaciones
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_ACont_Emi, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    
    ' 3) Filtros
    Lo_BD.Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & CursoAcad
    Lo_BD.Range.AutoFilter Field:=BD_Concepto, Criteria1:="=1311,03"
    Lo_BD.Range.AutoFilter Field:=BD_Anul, Criteria1:="<>S"        ' se debería haber puesto y se escapó!!!
    Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="<>_Devol_", Operator:=xlAnd, Criteria2:="<>_ERR_Date_"
    ' 4) Rango visible filtrado (incluye cabeceras)
    On Error Resume Next
    Set RngVisible = Lo_BD.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo SalirLimpiando
    If RngVisible Is Nothing Then GoTo SalirLimpiando
    ' Limpia cualquier tabla previa en Buffer
    If WsBuffer.ListObjects.Count > 0 Then
        WsBuffer.ListObjects(1).Unlist   'o .Delete si prefieres quitar formato tabla
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
    Call RuT_Cierre_Contable_Planes_CAcad_Pos(Lo_BD_Filtrada, CursoAcad, "CFCyAFC")
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
    If Not Lo_BD Is Nothing Then
        Call Rut_Lo_Filtros_Quitar(Lo_BD)     '- Retira el AutoFilter dejado por el Ciclo_N ejecutado
        Lo_BD.ShowTotals = True
        Sht__BD.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
    End If
    Application.ScreenUpdating = True
    Application.Calculation = Sw_Calculation
    Application.EnableEvents = True
    Application.Speech.Speak "Proceso Terminado"
End Sub '------------------------------------------------------------------------------------------------------------------------



'==================================================================================================================================
Sub RuT_Cierre_Contable_Planes_CAcad_Ant(Lo_BD As ListObject, _
                                         CursoAcad As String, _
                                         TipoCurso As String)
'==================================================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim Cont            As Long
    Dim ContIni         As Long:        ContIni = 1
    Dim Txt_Cabecera    As String
    Dim Cont_Plan       As Integer
    Dim PlanesSinCob    As Integer:     PlanesSinCob = 0
    
    Dim RegsEmis        As Long     ' regs Emitidos
    Dim Imp_Emis        As Currency
    Dim TimpEmis        As Currency
    
    Dim RegsEmisAnt     As Long     ' regs Emitidos
    Dim Imp_EmisAnt     As Currency
    Dim TImpEmisAnt     As Currency
    
    Dim RegsEmisPos     As Long     ' regs Emitidos
    Dim Imp_EmisPos     As Currency
    Dim TImpEmisPos     As Currency
    Dim Imp_EmisPosAdm     As Currency
    Dim TimpEmisPosAdm     As Currency
    Dim Imp_EmisPosAcad    As Currency
    Dim TimpEmisPosAcad    As Currency

    Dim Imp_Cobr        As Currency
    Dim TImpCobr        As Currency
    
    Dim Imp_CobrAdm     As Currency
    Dim TImpCobradm     As Currency
    
    Dim Imp_CobrAcad    As Currency
    Dim TImpCobrAcad    As Currency
    
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

    Dim RegsPdtCob        As Long     ' regs Pendiente de pago
    Dim Imp_PdtCob        As Currency
    Dim TimpPdtCob        As Currency

    Dim R_EjeAnt        As Long     ' regs Pendiente de pago
    Dim I_EjeAnt        As Currency
    Dim T_EjeAnt        As Currency

    Const Lst_Orden         As Integer = 1
    Const Lst_Plan          As Integer = 2
    Const Lst_RegsEmis      As Integer = 3
    Const Lst_Imp_Emis      As Integer = 4
    Const Lst_Imp_EmisAnt   As Integer = 5
    Const Lst_Imp_ADx       As Integer = 6
    Const Lst_Imp_ADxAdm    As Integer = 7
    Const Lst_Imp_ADxAcad   As Integer = 8
    Const Lst_Imp_Aplz      As Integer = 9
    Const Lst_ImpEjeAnt     As Integer = 10
    Const Lst_Imp_EmisPos   As Integer = 11
    Const Lst_Imp_EmisPosAdm    As Integer = 12
    Const Lst_Imp_EmisPosAcad   As Integer = 13
    Const Lst_Imp_Cobr      As Integer = 14
    Const Lst_Imp_CobrAdm   As Integer = 15
    Const Lst_Imp_CobrAcad  As Integer = 16
    Const Lst_Imp__RDT      As Integer = 17
    Const Lst_Imp_SRDT      As Integer = 18
    Const Lst_Imp_PdtCob    As Integer = 19
    
    Dim Regs_Dev            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula
    Dim Imp_Devo            As Currency
    Dim RDev_Pag            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, Pagados
    Dim IDev_Pag            As Currency
    Dim RDevPdte            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, No Pagados, pero Ajuste Matrícula NO deben ser pagados
    Dim IDevPdte            As Currency
    
    Dim Ws_Lista    As Worksheet:       Set Ws_Lista = Sheets("Cierre_Planes_CAcad_Ant_" & TipoCurso)
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
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_ADx).Value = "Imp_Emi'" & Right(AñoCont - 1, 2) & vbLf & "y ADxAplz" & vbLf & "en " & AñoCont - 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_ADxAdm).Value = "Imp_ADxAplz" & vbLf & "Adm. en" & vbLf & AñoCont - 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_ADxAcad).Value = "Imp_ADxAplz" & vbLf & "Acad. en" & vbLf & AñoCont - 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Aplz).Value = "ADxAplz'" & Right(AñoCont - 1, 2) & vbLf & "Cob " & AñoCont & vbLf & "Aplazado'" & Right(AñoCont, 2)
    Lo_Lst.HeaderRowRange.Cells(Lst_ImpEjeAnt).Value = "Emi." & AñoCont - 1 & vbLf & "Rec. EjeAnt" & vbLf & "Cob " & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisPos).Value = "Imp_Emi" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisPosAdm).Value = "Imp_Emi" & vbLf & "Adm. en" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisPosAcad).Value = "Imp_Emi" & vbLf & "Acad. en" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Cobr).Value = "Imp_Cob" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_CobrAdm).Value = "Imp_Cob" & vbLf & "Adm." & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_CobrAcad).Value = "Imp_Cob" & vbLf & "Acad." & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDT).Value = "Imp_RDT" & vbLf & CursoAcad
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_SRDT).Value = "Pdte. RDT"
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCob).Value = "Pdte.Cob. " & vbLf & AñoCont
    
    
'    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_ADx).Value = "Imp_ADxAplz" & vbLf & AñoCont + 1
'    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCob).Value = "Pdte_Cob" & vbLf & AñoCont

    '- Visualizo el progreso ---------------------------------------------------------------------------------------
    Form_Menu.Lb_Tít_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TBx_Informe = Txt_Cabecera & vbCrLf & vbCrLf
    Txt_Cabecera = String(16, " ") & AñoCont - 1 & String(11, " ") & AñoCont & String(11, " ") & AñoCont & "         Cob." & AñoCont & "       Cob." & AñoCont & vbLf & _
                   "    Plan       Imp_Emi        Imp_Emi        ADxAplz       Aplazado        Eje_Ant"
    Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & Txt_Cabecera & vbLf
    
    
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
            TimpEmis = TimpEmis + Imp_Emis
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
            TimpEmisPosAdm = TimpEmisPosAdm + Imp_EmisPosAdm
    
            '- Importe  EmiPosAdm_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_EmisPosAcad = Imp_EmisPos - Imp_EmisPosAdm
            RowNew.Range(Lst_Imp_EmisPosAcad) = Imp_EmisPosAcad
            TimpEmisPosAcad = TimpEmisPosAcad + Imp_EmisPosAcad
    
'- Importe Recibos Cob -----------------------------------------------------------------------------------------------------------------
            Imp_Cobr = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan)
            RowNew.Range(Lst_Imp_Cobr) = Imp_Cobr
            TImpCobr = TImpCobr + Imp_Cobr
    
            '- Importe Adm Recibos Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_CobrAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont)
            RowNew.Range(Lst_Imp_CobrAdm) = Imp_CobrAdm
            TImpCobradm = TImpCobradm + Imp_CobrAdm
    
            '- Importe Acad Recibos Cob_AñoCont -----------------------------------------------------------------------------------------------------------------
            Imp_CobrAcad = Imp_Cobr - Imp_CobrAdm
            RowNew.Range(Lst_Imp_CobrAcad) = Imp_CobrAcad
            TImpCobrAcad = TImpCobrAcad + Imp_CobrAcad
            
'- Importe Recibos RDT -----------------------------------------------------------------------------------------------------------------
            Imp__RDT = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpCob), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "<>")
            RowNew.Range(Lst_Imp__RDT) = Imp__RDT
            TImp_RDT = TImp_RDT + Imp__RDT

'- Importe Sin_RDT --------------------------------------------------------------------------------------------------------------------
            Imp_SRDT = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "=", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_SRDT) = Imp_SRDT
            TImpSRDT = TImpSRDT + Imp_SRDT

'- Importe Pdte_Cob --------------------------------------------------------------------------------------------------------------------
            Imp_PdtCob = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), "=")
            RowNew.Range(Lst_Imp_PdtCob) = Imp_PdtCob
            TimpPdtCob = TimpPdtCob + Imp_PdtCob
    
'- Informe --------------------------------------------------------------------------------------------------------------------
            Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & Format(Cont_Plan, "00") & "º " & Cod_Plan & " " & _
                Right(String(12, " ") & Format(Imp_EmisAnt, "#,##0.00"), 14) & _
                IIf(Imp_EmisPos > 0, " " & Right(String(12, "·") & Format(Imp_EmisPos, "#,##0.00"), 14), "  No hay Cobros") & _
                IIf(Imp__ADx > 0, " " & Right(String(12, "·") & Format(Imp__ADx, "#,##0.00"), 14), " " & String(14, "·")) & _
                IIf(Imp_Aplz > 0, " " & Right(String(12, "·") & Format(Imp_Aplz, "#,##0.00"), 14), " " & String(14, "·")) & _
                IIf(I_EjeAnt > 0, " " & Right(String(12, "·") & Format(I_EjeAnt, "#,##0.00"), 14), " " & String(14, "·")) & vbLf

            ' Intercalo Cabecera Columnas
            If Cont_Plan Mod 20 = 0 Then Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & Txt_Cabecera & vbLf
    
            If Imp_Cobr = 0 Then PlanesSinCob = PlanesSinCob + 1
Siguiente_Plan:
    Next Cod_Plan
    End With    ' Lo_BD.DataBodyRange
    
            Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & Txt_Cabecera
            Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & " Totales "
            Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & _
                Right(String(12, " ") & Format(TImpEmisAnt, "#,##0.00"), 14) & _
                Right(String(12, " ") & Format(TImpEmisPos, "#,##0.00"), 15) & _
                Right(String(12, " ") & Format(TImp_ADx, "#,##0.00"), 15) & _
                Right(String(12, " ") & Format(TImpAplz, "#,##0.00"), 15) & _
                Right(String(12, " ") & Format(T_EjeAnt, "#,##0.00"), 15) & vbLf


            If PlanesSinCob > 1 Then
                Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & "¡¡ " & PlanesSinCob & " Planes sin Cobros !!"
            ElseIf PlanesSinCob = 1 Then
                Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & "¡¡ 1 Plan sin Cobros !!"
            Else
                Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & "¡ No hay planes sin Cobros !"
            End If
    
'- Visualizo el progreso ---------------------------------------------------------------------------------------
Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & vbLf & _
                                                "¡¡¡ Proceso concluido !!! día: " & Now() & _
                                                " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg."
'End With    '- Lo_BD.DataBodyRange

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Sht__BD.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'''Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Listar_Planes   --------------------------------------------------------------------------------------------
'===================================================================================================================================


'==================================================================================================================================
Sub RuT_Cierre_Contable_Planes_CAcad_Pos(Lo_BD As ListObject, _
                                         CursoAcad As String, _
                                         TipoCurso As String)
'==================================================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim Cont            As Long
    Dim ContIni         As Long:        ContIni = 1
    Dim Txt_Cabecera    As String
    Dim Cont_Plan       As Integer
    Dim PlanesSinCob    As Integer:     PlanesSinCob = 0

    Dim RegsEmis        As Long     ' regs Emitidos
    Dim Imp_Emis        As Currency
    Dim TimpEmis        As Currency

    Dim RegsEmisAdm     As Long     ' regs Emitidos
    Dim Imp_EmisAdm     As Currency
    Dim TimpEmisAdm     As Currency

    Dim RegsEmisAcad    As Long     ' regs Emitidos
    Dim Imp_EmisAcad    As Currency
    Dim TimpEmisAcad    As Currency

    Dim RegsCobr        As Long     ' regs Cobrados
    Dim Imp_Cobr        As Currency
    Dim TImpCobr        As Currency

    Dim Regs_RDT        As Long     ' regs Cobrados
    Dim Imp__RDT        As Currency
    Dim TImp_RDT        As Currency

    Dim RegsSRDT        As Long     ' regs Pendiente de pago
    Dim Imp_SRDT        As Currency
    Dim TImpSRDT        As Currency

    Dim Regs_ADx        As Long     ' regs Pendiente de pago
    Dim Imp__ADx        As Currency
    Dim TImp_ADx        As Currency

    Dim RegsADxAdm     As Long     ' regs Pendiente de pago
    Dim Imp_ADxAdm     As Currency
    Dim TImpADxAdm     As Currency

    Dim RegsADxAcad    As Long     ' regs Pendiente de pago
    Dim Imp_ADxAcad    As Currency
    Dim TImpADxAcad    As Currency

    Dim RegsPdtCob        As Long     ' regs Pendiente de pago
    Dim Imp_PdtCob        As Currency
    Dim TimpPdtCob        As Currency

    Const Lst_Orden         As Integer = 1
    Const Lst_Plan          As Integer = 2
    Const Lst_RegsEmis      As Integer = 3
    Const Lst_Imp_Emis      As Integer = 4
    Const Lst_Imp_EmisAdm   As Integer = 5
    Const Lst_Imp_EmisAcad  As Integer = 6
    Const Lst_Imp_Cobr      As Integer = 7
'    Const Lst_RegsCobr     As Integer = 6
    Const Lst_Imp__RDT      As Integer = 8
'    Const Lst_Regs_RDT      As Integer = 8
    Const Lst_Imp_SRDT      As Integer = 9
    Const Lst_RegsSRDT      As Integer = 10
    Const Lst_Imp__ADx      As Integer = 11
    Const Lst_Imp__ADxAdm   As Integer = 12
    Const Lst_Imp__ADxAcad  As Integer = 13
    Const Lst_Imp_PdtCob    As Integer = 14

    Dim Regs_Dev            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula
    Dim Imp_Devo            As Currency
    Dim RDev_Pag            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, Pagados
    Dim IDev_Pag            As Currency
    Dim RDevPdte            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, No Pagados, pero Ajuste Matrícula NO deben ser pagados
    Dim IDevPdte            As Currency

    Dim Ws_Lista    As Worksheet:       Set Ws_Lista = Sheets("Cierre_Planes_CAcad_Pos_" & TipoCurso)
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
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Emis).Value = "Imp_Emi" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisAdm).Value = "Imp_Emi" & vbLf & AñoCont & vbLf & "Adm."
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_EmisAcad).Value = "Imp_Emi" & vbLf & AñoCont & vbLf & "Acad."
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_Cobr).Value = "Imp_Cob" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__RDT).Value = "Imp_RDT" & vbLf & AñoCont
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_SRDT).Value = "Pdte_RDT"
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADx).Value = "Imp_ADxAplz" & vbLf & AñoCont + 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADxAdm).Value = "Imp_ADxAplz" & vbLf & "1303.00 " & vbLf & "Adm. " & AñoCont + 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp__ADxAcad).Value = "Imp_ADxAplz" & vbLf & "1311.00 " & vbLf & "Acad " & AñoCont + 1
    Lo_Lst.HeaderRowRange.Cells(Lst_Imp_PdtCob).Value = "Pdte_Cob" & vbLf & AñoCont

    '- Visualizo el progreso ---------------------------------------------------------------------------------------
    Form_Menu.Lb_Tít_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TBx_Informe = Txt_Cabecera & vbCrLf & vbCrLf
    Txt_Cabecera = String(27, " ") & "Cob-" & AñoCont & String(18, " ") & "Cob-" & AñoCont & String(6, " ") & "ADxAplz" & String(6, " ") & "ADxAplz" & String(7, " ") & AñoCont & vbLf & _
                   "    Plan     Imp_Rec       Imp_Cob       Imp_RDT     Pdte.RDT     Adm." & AñoCont + 1 & "     Acad." & AñoCont + 1 & "     Pde_Cob"
    Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & Txt_Cabecera & vbLf




    '- Recorro la Collection con todos los Planes
         Dim Cod_Plan As Variant
    With Lo_BD.DataBodyRange
    For Each Cod_Plan In Collection_Planes

        Cont_Plan = Cont_Plan + 1
        Set RowNew = Lo_Lst.ListRows.Add
            RowNew.Range(Lst_Orden) = Cont_Plan
            RowNew.Range(Lst_Plan) = Cod_Plan

            '- Importe Recibos Emi_2025 -----------------------------------------------------------------------------------------------------------------
            Imp_Emis = Application.SumIfs(.Columns(BD_ImpRec), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            If Imp_Emis = 0 Then GoTo Siguiente_Plan    '- Si todo está cobrado en año anterior
            RowNew.Range(Lst_Imp_Emis) = Imp_Emis
            TimpEmis = TimpEmis + Imp_Emis
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
            TimpEmisAdm = TimpEmisAdm + Imp_EmisAdm
            
            '- Importe Recibos Emi_Acad -----------------------------------------------------------------------------------------------------------------
            Imp_EmisAcad = Imp_Emis - Imp_EmisAdm
            RowNew.Range(Lst_Imp_EmisAcad) = Imp_EmisAcad
            TimpEmisAcad = TimpEmisAcad + Imp_EmisAcad
            
            '- Importe Cob_2025, Emi_2025 --------------------------------------------------------------------------------------------------------------------
            Imp_Cobr = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_Cobr) = Imp_Cobr
            TImpCobr = TImpCobr + Imp_Cobr
            
            '- Importe RDT --------------------------------------------------------------------------------------------------------------------
            Imp__RDT = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_ImpRec), ">0", _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "<>", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp__RDT) = Imp__RDT
            TImp_RDT = TImp_RDT + Imp__RDT

            '- Importe Pdte_RDT --------------------------------------------------------------------------------------------------------------------
            Imp_SRDT = Application.SumIfs(.Columns(BD_ImpCob), _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "=", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_Imp_SRDT) = Imp_SRDT
            TImpSRDT = TImpSRDT + Imp_SRDT
            RegsSRDT = Application.CountIfs( _
                                          .Columns(BD_Plan), Cod_Plan, _
                                          .Columns(BD_RDT), "=", _
                                          .Columns(BD_ACont_Cob), AñoCont, _
                                          .Columns(BD_ACont_Emi), AñoCont)
            RowNew.Range(Lst_RegsSRDT) = RegsSRDT

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


'''            '- Importe Recibos Emi_2025 -----------------------------------------------------------------------------------------------------------------
'''            Imp_Emis = Application.SumIfs(.Columns(BD_ImpRec), _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_ACont_Emi), AñoCont)
''''            If Imp_Emis = 0 Then GoTo Siguiente_Plan    '- Si todo está cobrado en año anterior
'''            RowNew.Range(Lst_Imp_Emis) = Imp_Emis
'''            TimpEmis = TimpEmis + Imp_Emis
'''            RegsEmis = Application.CountIfs( _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_RegsEmis) = RegsEmis
'''
'''            '- Importe Cob_2025, Emi_2025 --------------------------------------------------------------------------------------------------------------------
'''            Imp_Cobr = Application.SumIfs(.Columns(BD_ImpCob), _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_ACont_Cob), AñoCont, _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_Imp_Cobr) = Imp_Cobr
'''            TImpCobr = TImpCobr + Imp_Cobr
'''            RegsCobr = Application.CountIfs( _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_ACont_Cob), AñoCont, _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_RegsCobr) = RegsCobr
'''
'''            '- Importe RDT --------------------------------------------------------------------------------------------------------------------
'''            Imp__RDT = Application.SumIfs(.Columns(BD_ImpCob), _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_RDT), "<>", _
'''                                          .Columns(BD_ACont_Cob), AñoCont, _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_Imp__RDT) = Imp__RDT
'''            TImp_RDT = TImp_RDT + Imp__RDT
'''            Regs_RDT = Application.CountIfs( _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_RDT), "<>", _
'''                                          .Columns(BD_ACont_Cob), AñoCont, _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_Regs_RDT) = Regs_RDT
'''
'''            '- Importe Pdte_RDT --------------------------------------------------------------------------------------------------------------------
'''            Imp_SRDT = Application.SumIfs(.Columns(BD_ImpCob), _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_RDT), "=", _
'''                                          .Columns(BD_ACont_Cob), AñoCont, _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_Imp_SRDT) = Imp_SRDT
'''            TImpSRDT = TImpSRDT + Imp_SRDT
'''            RegsSRDT = Application.CountIfs( _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_RDT), "=", _
'''                                          .Columns(BD_ACont_Cob), AñoCont, _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_RegsSRDT) = RegsSRDT
'''
'''            '- Importe ADxAplz --------------------------------------------------------------------------------------------------------------------
'''            Imp__ADx = Application.SumIfs(.Columns(BD_ImpRec), _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_ACont_Vto), AñoCont + 1, _
'''                                          .Columns(BD_ACont_Cob), "<>" & AñoCont, _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_Imp__ADx) = Imp__ADx
'''            TImp_ADx = TImp_ADx + Imp__ADx
'''
'''            '- Importe ADxAplzAdm --------------------------------------------------------------------------------------------------------------------
'''            Imp_ADxAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_ACont_Vto), AñoCont + 1, _
'''                                          .Columns(BD_ACont_Cob), "<>" & AñoCont, _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_Imp__ADxAdm) = Imp_ADxAdm
'''            TImpADxAdm = TImpADxAdm + Imp_ADxAdm
'''
'''            '- Importe ADxAplzAcad --------------------------------------------------------------------------------------------------------------------
'''            Imp_ADxAcad = Imp__ADx - Imp_ADxAdm
'''            RowNew.Range(Lst_Imp__ADxAcad) = Imp_ADxAcad
'''            TImpADxAcad = TImpADxAcad + Imp_ADxAcad
'''
'''
'''            '- Importe Pdte_Cob --------------------------------------------------------------------------------------------------------------------
'''            Imp_PdtCob = Application.SumIfs(.Columns(BD_ImpRec), _
'''                                          .Columns(BD_ImpRec), ">0", _
'''                                          .Columns(BD_Plan), Cod_Plan, _
'''                                          .Columns(BD_Anul), "<>S", _
'''
'''
'''                                          .Columns(BD_ACont_Vto), AñoCont, _
'''                                          .Columns(BD_ACont_Cob), "=", _
'''                                          .Columns(BD_ACont_Emi), AñoCont)
'''            RowNew.Range(Lst_Imp_PdtCob) = Imp_PdtCob
'''            TimpPdtCob = TimpPdtCob + Imp_PdtCob
'''

            '-- Informe --------------------------------------------------------------------------------------------------------------------------
            Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & Format(Cont_Plan, "00") & "º " & Cod_Plan & " " & _
                Right(String(12, " ") & Format(Imp_Emis, "#,##0.00"), 13) & _
                IIf(Imp_Cobr > 0, " " & Right(String(12, "·") & Format(Imp_Cobr, "#,##0.00"), 13), "  ¡Sin Cobros!") & _
                IIf(Imp__RDT > 0, " " & Right(String(12, "·") & Format(Imp__RDT, "#,##0.00"), 13), " " & String(13, "·")) & _
                IIf(Imp_SRDT > 0, " " & Right(String(12, "·") & Format(Imp_SRDT, "#,##0.00"), 12), " " & String(12, "·")) & _
                IIf(Imp_ADxAdm > 0, " " & Right(String(12, "·") & Format(Imp_ADxAdm, "#,##0.00"), 11), " " & String(11, "·")) & _
                IIf(Imp_ADxAcad > 0, " " & Right(String(12, "·") & Format(Imp_ADxAcad, "#,##0.00"), 13), " " & String(13, "·")) & _
                IIf(Imp_PdtCob > 0, " " & Right(String(12, "·") & Format(Imp_PdtCob, "#,##0.00"), 12), " " & String(12, "·")) & vbLf

            ' Intercalo Cabecera Columnas
            If Cont_Plan Mod 20 = 0 Then Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & Txt_Cabecera & vbLf

            If Imp_Cobr = 0 Then PlanesSinCob = PlanesSinCob + 1
Siguiente_Plan:
    Next Cod_Plan
    End With    ' Lo_BD.DataBodyRange

            Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & Txt_Cabecera
            Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & " Totales "
            Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & _
                Right(String(12, " ") & Format(TimpEmis, "#,##0.00"), 13) & _
                Right(String(12, " ") & Format(TImpCobr, "#,##0.00"), 14) & _
                Right(String(12, " ") & Format(TImp_RDT, "#,##0.00"), 14) & _
                Right(String(12, " ") & Format(TImpSRDT, "#,##0.00"), 13) & _
                Right(String(12, " ") & Format(TImpADxAdm, "#,##0.00"), 12) & _
                Right(String(12, " ") & Format(TImpADxAcad, "#,##0.00"), 14) & _
                Right(String(12, " ") & Format(TimpPdtCob, "#,##0.00"), 13) & vbLf


            If PlanesSinCob > 1 Then
                Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & "¡¡ " & PlanesSinCob & " Planes sin Cobros !!"
            ElseIf PlanesSinCob = 1 Then
                Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & "¡¡ 1 Plan sin Cobros !!"
            Else
                Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & "¡ No hay planes sin Cobros !"
            End If

'- Visualizo el progreso ---------------------------------------------------------------------------------------
Form_Menu.TBx_Informe = Form_Menu.TBx_Informe & vbLf & vbLf & _
                                                "¡¡¡ Proceso concluido !!! día: " & Now() & _
                                                " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg."
'End With    '- Lo_BD.DataBodyRange

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Sht__BD.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
''''Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Listar_Planes   --------------------------------------------------------------------------------------------
'===================================================================================================================================
