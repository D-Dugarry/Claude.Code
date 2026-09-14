Attribute VB_Name = "M32_List_Anul_y_Devoluciones"
'2026-01-31
'- M31_Cierre_Contable_PLANES
Option Explicit

'==================================================================================================================================
Sub RuT_Inf_Anulaciones_y_Devoluciones()
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
    Call Rut_Lo_Sort(Lo_BD, BD_Anul, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_DNI, xlAscending, False)
    ' 3) Filtros
    Lo_BD.Range.AutoFilter Field:=BD_Anul, Criteria1:="=S"
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
    Call RuT_Lista_Anulaciones(Lo_BD_Filtrada, CursoAcad, TipoCurso, Wk_Inf_Anulados)
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
Sub RuT_Lista_Anulaciones(Lo_BD As ListObject, _
                            CursoAcad As String, _
                            TipoCurso As String, _
                            Ws_Lista As Worksheet)
'==================================================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim AñoContAnt      As Integer:     AñoContAnt = Left(CursoAcad, 4)
    Dim AñoContPos      As Integer:     AñoContPos = "20" & Right(CursoAcad, 2)
    Dim ACont           As Integer:     ACont = Right(Prog__APP.Range("APP_AñoCont"), 2)
    Dim AContAnt        As String:     AContAnt = Mid$(CursoAcad, 3, 2)
    Dim AContPos        As String:     AContPos = Right(CursoAcad, 2)

    Dim Cont                As Long
    Dim ContIni             As Long:        ContIni = 1
    Dim Cod_Plan            As String
    Dim ClaveNew            As String
    Dim ClaveAnt            As String
    Dim Txt_Cabecera        As String
    Dim Cont_Anul           As Integer
    Dim PlanesSinCob        As Integer:     PlanesSinCob = 0
    
    Dim RegsEmis            As Long     ' regs Emitidos
    Dim Imp_Emis            As Currency
    Dim RegsCobr            As Long     ' regs Cobrados
    Dim Imp_Cobr            As Currency
    Dim RegsPdts            As Long     ' regs Pendiente de pago
    Dim Imp_Pdte            As Currency
    
    Dim Regs_Dev            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula
    Dim Imp_Devo            As Currency
    Dim RDev_Pag            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, Pagados
    Dim IDev_Pag            As Currency
    Dim RDevPdte            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, No Pagados, pero Ajuste Matrícula NO deben ser pagados
    Dim IDevPdte            As Currency
    
    Dim Lo_Lst      As ListObject:      Set Lo_Lst = Ws_Lista.ListObjects(1)
    Dim RowNew      As ListRow
    
    Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    
    Call Rut_Lo_Filtros_Quitar(Lo_Lst)
    If Not Lo_Lst.DataBodyRange Is Nothing Then Lo_Lst.DataBodyRange.Delete

    Const Lst_Orden         As Integer = 1
    Const Lst_Plan          As Integer = 2
    Const Lst_DNI           As Integer = 3
    Const Lst_Nombre        As Integer = 4
    Const Lst_FEmiNeg       As Integer = 5
    Const Lst_ImpNeg        As Integer = 6
    Const Lst_RegsNeg       As Integer = 7
    Const Lst_FPagoNeg      As Integer = 8
    Const Lst_ImpNegPago    As Integer = 9
    Const Lst_RegsNegPag    As Integer = 10
    Const Lst_FEmiPos       As Integer = 11
    Const Lst_ImpPos        As Integer = 12
    Const Lst_RegsPos       As Integer = 13
    Const Lst_FCob          As Integer = 14
    Const Lst_ImpCob        As Integer = 15
    Const Lst_RegsCob       As Integer = 16
    Const Lst_ImpAcad       As Integer = 17
    Const Lst_ImpAdm        As Integer = 18
    Const Lst_Obs           As Integer = 19
    
    Dim T_ImpEmi        As Currency
    Dim T_ImpCob        As Currency
    Dim T_ImpPdt        As Currency


With Lo_BD.DataBodyRange

    ' Inicio Listado en Tabla excel
    Txt_Cabecera = "Planes de " & Range("APP_EFP_o_CFC") & "_" & Range("APP_CursAcad") & String(10, " ") & Now
    Ws_Lista.Range("c2") = Txt_Cabecera
    
    '- Visualizo el progreso ---------------------------------------------------------------------------------------
    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TB_Informe = Txt_Cabecera
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera & vbLf & vbLf
    Txt_Cabecera = "    Plan  Nombre  (dni)"
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera
        
    ' Recorro toda la Tabla ---------------------------------------------------------------------------------------
    For Cont = 1 To Lo_BD.ListRows.Count
        'If .Cells(Cont, BD_Tipo_Rec) = "Deleted" Then GoTo Reg_Siguiente
        ClaveNew = .Cells(Cont, BD_Plan) & "_" & .Cells(Cont, BD_DNI)
        '- <<<<<  PLAN NUEVO  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        If ClaveAnt <> ClaveNew Then
            
            If Cont_Anul = 0 Then GoTo Nueva_Fila
            '- Cumplimento Observaciones, resumen de la fila
            '- Imp_Adm < 0
            If RowNew.Range(Lst_ImpAdm) < 0 Then
                RowNew.Range(Lst_Obs) = "Anulado parte matrícula no emitida." '-  (Debería menguar el Emitido de Matrícula en la contabilidad)
            End If
            '- ImpNeg < 0
            If RowNew.Range(Lst_ImpNeg) < 0 Then
                If RowNew.Range(Lst_ImpNegPago) < 0 Then
                    RowNew.Range(Lst_Obs) = "Devolución Total/Parcial de recibo Remesada."
                Else
                    RowNew.Range(Lst_Obs) = "Devolución Total/Parcial de recibo, sin Remesa."
                End If
            End If
            ' Visualizo Nuevo Cod_Plan
            'Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Format(Cont_Anul, "00") & "º " & Cod_Plan & " " & Left(RowNew.Range(Lst_DNI) & String(15, " "), 16) & _
                                   Left(.Cells(Cont, BD_Nom) & String(15, " "), 30) & " " & RowNew.Range(Lst_Obs) & vbLf
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Cont_Anul, "00") & "º " & Cod_Plan & " " & _
                                   .Cells(Cont, BD_Nom) & "  ( " & RowNew.Range(Lst_DNI) & " )" & vbLf & _
                                   String(15, " ") & "Obs.: "" " & RowNew.Range(Lst_Obs) & " """ & vbLf
            '- Detalle del Informe ---------------------------------------------------------------------------
            If RowNew.Range(Lst_ImpNeg) < 0 Then
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & String(15, " ") & _
                    Left("Emitido Recibo Negativo de devolución " & String(45, "·"), 45) & Right(String(15, " ") & _
                    Format(RowNew.Range(Lst_ImpNeg), "#,##0.00 €;-#,##0.00 €;"), 15) & _
                    "   (" & Format(RowNew.Range(Lst_FEmiNeg), "dd/mm/yyyy") & ")" & vbLf
                If RowNew.Range(Lst_ImpNegPago) < 0 Then
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & String(15, " ") & _
                    Left("Tranferencia hecha " & String(45, "·"), 45) & _
                    Right(String(15, " ") & Format(RowNew.Range(Lst_ImpNegPago) * 1, "#,##0.00 €;-#,##0.00 €;0.00 €"), 15) & _
                    "   (" & Format(RowNew.Range(Lst_FPagoNeg), "dd/mm/yyyy") & ")" & vbLf
                Else
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & String(15, " ") & _
                    Left("Tranferencia pendiente" & String(45, "·"), 45) & _
                    Right(String(15, " ") & Format(RowNew.Range(Lst_ImpNegPago) * 1, "#,##0.00 €;-#,##0.00 €;0.00 €"), 15) & vbLf
                End If
            End If
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
            String(15, " ") & Left("Emitido Recibo de matrícula de alumno " & String(45, "·"), 45) & _
                    Right(String(15, " ") & Format(RowNew.Range(Lst_ImpPos), "#,##0.00 €;-#,##0.00 €;"), 15) & _
                    "   (" & Format(RowNew.Range(Lst_FEmiPos), "dd/mm/yyyy") & ")" & vbLf
            If RowNew.Range(Lst_ImpCob) > 0 Then
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                String(15, " ") & Left("Recibo de matrícula de alumno cobrado " & String(45, "·"), 45)
            Else
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                String(15, " ") & Left("Recibo de matrícula de alumno NO cobrado " & String(45, "·"), 45)
            End If
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                Right(String(15, " ") & Format(RowNew.Range(Lst_ImpCob) * 1, "#,##0.00 €;-#,##0.00 €;"), 15) & _
                "   (" & Format(RowNew.Range(Lst_FCob), "dd/mm/yyyy") & ")" & vbLf
                
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                String(85, " ") & " Imp.Acad" & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpAcad), "#,##0.00;-#,##0.00;"), 9) & " " & vbLf & _
                String(85, " ") & " Imp.Adm" & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpAdm), "#,##0.00;-#,##0.00;"), 10) & " " & vbLf

'            ¡¡¡¡¡  Programación mucho más corta, pero enfarragosa de leer !!!!!! ------------------------------------------------------------------------------------------
'            - Detalle del Informe -----------------------------------------------------------------------------------------------------------------------------------------
'            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
'                IIf(RowNew.Range(Lst_ImpNeg) < 0, String(15, " ") & Left("Emitido Recibo Negativo de devolución " & String(45, "·"), 45) & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpNeg), "#,##0.00 €;-#,##0.00 €;"), 15) & vbLf & _
'                String(15, " ") & IIf(RowNew.Range(Lst_ImpNegPago) < 0, Left("Tranferencia hecha " & String(45, "·"), 45), Left("Tranferencia pendiente" & String(45, "·"), 45)) & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpNegPago) * 1, "#,##0.00 €;-#,##0.00 €;0.00 €"), 15) & vbLf, "") & _
'                String(15, " ") & Left("Emitido Recibo de matrícula de alumno " & String(45, "·"), 45) & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpPos), "#,##0.00 €;-#,##0.00 €;"), 15) & vbLf & _
'                String(15, " ") & IIf(RowNew.Range(Lst_ImpCob) > 0, Left("Recibo de matrícula de alumno cobrado " & String(45, "·"), 45) & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpCob), "#,##0.00 €;-#,##0.00 €;"), 15), _
'                                                                    Left("Recibo de matrícula de alumno NO cobrado " & String(45, "·"), 45) & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpCob) * 1, "#,##0.00 €;-#,##0.00 €;"), 15)) & vbLf & _
'                String(15, " ") & " Imp.Acad" & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpAcad), "#,##0.00;-#,##0.00;"), 9) & " " & vbLf & _
'                String(15, " ") & " Imp.Adm" & Right(String(15, " ") & Format(RowNew.Range(Lst_ImpAdm), "#,##0.00;-#,##0.00;"), 10) & " " & vbLf

Nueva_Fila:
            ClaveAnt = ClaveNew
            Cont_Anul = Cont_Anul + 1
            Set RowNew = Lo_Lst.ListRows.Add
            RowNew.Range(Lst_Orden) = Cont_Anul
            Cod_Plan = .Cells(Cont, BD_Plan)
            RowNew.Range(Lst_Plan) = Cod_Plan
            RowNew.Range(Lst_DNI) = .Cells(Cont, BD_DNI)
            RowNew.Range(Lst_Nombre) = .Cells(Cont, BD_Nom)
            RowNew.Range(Lst_ImpAcad) = .Cells(Cont, BD_ImpAcad)
            RowNew.Range(Lst_ImpAdm) = .Cells(Cont, BD_ImpAdm)
            
        End If
        
        '- <<<<<  Acumulado en PLAN  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos de Devolución (a pagar) / Ajuste Matrícula
                                        RowNew.Range(Lst_RegsNeg) = RowNew.Range(Lst_RegsNeg) + 1
                                        RowNew.Range(Lst_ImpNeg) = RowNew.Range(Lst_ImpNeg) + .Cells(Cont, BD_ImpRec)
                                        RowNew.Range(Lst_FEmiNeg) = .Cells(Cont, BD_FEmi)
                If .Cells(Cont, BD_ImpCob) < 0 Then
                                        RowNew.Range(Lst_RegsNegPag) = RowNew.Range(Lst_RegsNegPag) + 1
                                        RowNew.Range(Lst_ImpNegPago) = RowNew.Range(Lst_ImpNegPago) + .Cells(Cont, BD_ImpRec)
                                        RowNew.Range(Lst_FPagoNeg) = .Cells(Cont, BD_FCob)
                End If
            Else
                                        RowNew.Range(Lst_Obs) = "Recibo Anulado..."
                                        RowNew.Range(Lst_RegsPos) = RowNew.Range(Lst_RegsPos) + 1
                                        RowNew.Range(Lst_ImpPos) = RowNew.Range(Lst_ImpPos) + .Cells(Cont, BD_ImpRec)
                                        RowNew.Range(Lst_FEmiPos) = .Cells(Cont, BD_FEmi)
                If .Cells(Cont, BD_ImpCob) > 0 Then
                                        RowNew.Range(Lst_RegsCob) = RowNew.Range(Lst_RegsCob) + 1
                                        RowNew.Range(Lst_ImpCob) = RowNew.Range(Lst_ImpCob) + .Cells(Cont, BD_ImpRec)
                                        RowNew.Range(Lst_FCob) = .Cells(Cont, BD_FCob)
                End If
        End If

Reg_Siguiente:
    Next
    
            If Cont_Anul > 0 Then       '- Cumplimento Observaciones
                '- Imp_Adm < 0
                If RowNew.Range(Lst_ImpAdm) < 0 Then
                    RowNew.Range(Lst_Obs) = "Anulado parte matrícula no emitida." '-  (Debería menguar el Emitido de Matrícula en la contabilidad)
                End If
                '- ImpNeg < 0
                If RowNew.Range(Lst_ImpNeg) < 0 Then
                    If RowNew.Range(Lst_ImpNegPago) < 0 Then
                        RowNew.Range(Lst_Obs) = "Devolución Total/Parcial de recibo Remesada."
                    Else
                        RowNew.Range(Lst_Obs) = "Devolución Total/Parcial de recibo, sin Remesa."
                    End If
                End If
            End If
            
    
'- Visualizo el progreso ---------------------------------------------------------------------------------------

Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & _
                "¡¡¡ Los Recibos con Anulado de parte matrícula NO emitida, deberían menguar el monto Emitido de Matrícula en la contabilidad !!!"

Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & _
                                                "¡¡¡ Proceso concluido !!! día: " & Now() & _
                                                " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg."
End With    '- Lo_BD.DataBodyRange

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Listar_Planes   --------------------------------------------------------------------------------------------
'===================================================================================================================================




