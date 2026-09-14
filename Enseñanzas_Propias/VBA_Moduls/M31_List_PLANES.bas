Attribute VB_Name = "M31_List_PLANES"
'2026-01-12
'- M31_Listar_PLANES
Option Explicit
'==================================================================================================================================
Sub RuT_Listar_Planes()
'==================================================================================================================================
    Dim Cont                As Long
    Dim ContIni             As Long:        ContIni = 1
    Dim Cod_Plan            As String
    Dim Txt_Cabecera        As String
    Dim Cont_Plan           As Integer:     Cont_Plan = 1
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
    
    Dim Ws_Lista    As Worksheet:       Set Ws_Lista = Wk_Lista_Panes
    Dim Lo_BD       As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
    Dim Lo_Lst      As ListObject:      Set Lo_Lst = Ws_Lista.ListObjects(1)
    Dim RowNew      As ListRow
    
    Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    
    If Not Lo_Lst.DataBodyRange Is Nothing Then Lo_Lst.DataBodyRange.Delete

    Const Lst_Orden         As Integer = 1
    Const Lst_Plan          As Integer = 2
    Const Lst_ImpEmi        As Integer = 3
    Const Lst_ImpEmiRgs     As Integer = 4
    Const Lst_ImpCob        As Integer = 5
    Const Lst_ImpCobRgs     As Integer = 6
    Const Lst_ImpPdt        As Integer = 7
    Const Lst_ImpPdtRgs     As Integer = 8
    
    Dim T_ImpEmi        As Currency
    Dim T_ImpCob        As Currency
    Dim T_ImpPdt        As Currency

    ' =============  Preparar Tabla de TitPH ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, False)

With Lo_BD.DataBodyRange
    '- Tratando el 1º Recibo
    Do While .Cells(ContIni, BD_Tipo_Rec) = "Deleted" Or .Cells(ContIni, BD_ImpAdm) < 0
        ContIni = ContIni + 1
    Loop
    Cont = ContIni
    Cod_Plan = .Cells(Cont, BD_Plan)
    If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos de Devolución (a pagar) / Ajuste Matrícula
                                                Regs_Dev = 1:       Imp_Devo = .Cells(Cont, BD_ImpRec)
        If .Cells(Cont, BD_ImpCob) < 0 Then                         ' recibos Negativos de Devolución Pagados
                                                RDev_Pag = 1:       IDev_Pag = .Cells(Cont, BD_ImpCob)
        End If
    Else                                                            ' recibos Emitidos, Cobrados o NO
                                                RegsEmis = 1:       Imp_Emis = .Cells(Cont, BD_ImpRec)
        If .Cells(Cont, BD_ImpCob) > 0 Then                         ' recibos Emitidos Cobrados
                                                RegsCobr = 1:       Imp_Cobr = .Cells(Cont, BD_ImpCob)
        End If
    End If
    ' Inicio Listado en Tabla excel
    Txt_Cabecera = "Planes de " & Range("APP_EFP_o_CFC") & "_" & Range("APP_CursAcad") & String(10, " ") & Now & vbCrLf & vbCrLf
    Ws_Lista.Range("c2") = Txt_Cabecera
    Set RowNew = Lo_Lst.ListRows.Add:       RowNew.Range(Lst_Orden) = Cont_Plan:        RowNew.Range(Lst_Plan) = Cod_Plan
    
    '- Visualizo el progreso ---------------------------------------------------------------------------------------
    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TB_Informe = Txt_Cabecera
    Txt_Cabecera = "    Plan    Imp_Rec   r.  -  Imp_Cob   r. =  Pdte.Cob   r."
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera & vbLf
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Cont_Plan, "00") & "º " & Cod_Plan & " "
        
    ' Recorro toda la Tabla ---------------------------------------------------------------------------------------
    For Cont = ContIni + 1 To Lo_BD.ListRows.Count
        If .Cells(Cont, BD_Tipo_Rec) = "Deleted" Then GoTo Reg_Siguiente
        If .Cells(Cont, BD_ImpAdm) < 0 Then GoTo Reg_Siguiente              '- Ajustes de Matrícula que distorcionan la Contabilidad !!!
        '- <<<<<  PLAN NUEVO  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        If .Cells(Cont, BD_Plan) <> Cod_Plan Then
            Cod_Plan = .Cells(Cont, BD_Plan)
            Cont_Plan = Cont_Plan + 1
            ' Relleno Tabla ListObject del Informe, con Acumulados Plan Anterior
                T_ImpEmi = T_ImpEmi + Imp_Emis:         T_ImpCob = T_ImpCob + Imp_Cobr:         T_ImpPdt = T_ImpPdt + Imp_Pdte
                RowNew.Range(Lst_ImpEmi) = Imp_Emis:        RowNew.Range(Lst_ImpEmiRgs) = RegsEmis
                RowNew.Range(Lst_ImpCob) = Imp_Cobr:        RowNew.Range(Lst_ImpCobRgs) = RegsCobr
                RowNew.Range(Lst_ImpPdt) = Imp_Pdte:        RowNew.Range(Lst_ImpPdtRgs) = RegsPdts
            ' Añado Row ListObject con el nuevo PLAN
            Set RowNew = Lo_Lst.ListRows.Add:       RowNew.Range(Lst_Orden) = Cont_Plan:        RowNew.Range(Lst_Plan) = Cod_Plan
            
            ' Visualizo Acumulados Plan Anterior
                RegsPdts = RegsEmis - RegsCobr:             Imp_Pdte = Imp_Emis - Imp_Cobr
                RDevPdte = Regs_Dev - RDev_Pag:             IDevPdte = Imp_Devo - IDev_Pag
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                Right(String(8, " ") & Format(Imp_Emis, "#,##0.00"), 11) & Right("   (" & RegsEmis, 4) & ")" & _
                IIf(Imp_Cobr > 0, Right(String(8, " ") & Format(Imp_Cobr, "#,##0.00"), 12) & Right("   (" & RegsCobr, 4) & ")", " " & String(16, "·")) & _
                IIf(Imp_Pdte > 0, Right(String(8, " ") & Format(Imp_Pdte, "#,##0.00"), 12) & Right("   (" & RegsPdts, 4) & ")", " " & String(16, " "))
            If Imp_Cobr = 0 Then    ' Visualizo, "No se ha pagado NADA"
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & "   " & ChrW(8709) & "  Plan Sin Cobros."
                PlanesSinCob = PlanesSinCob + 1
            End If
            ' Intercalo Cabecera Columnas
            If Cont_Plan Mod 20 = 1 Then Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & Txt_Cabecera
            ' Visualizo Nuevo Cod_Plan
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Format(Cont_Plan, "00") & "º " & Cod_Plan & " "
            ' Inicializo variables
                Regs_Dev = 0:       Imp_Devo = 0:       RDev_Pag = 0:       IDev_Pag = 0
                RegsEmis = 0:       Imp_Emis = 0:       RegsCobr = 0:       Imp_Cobr = 0
        End If
        '- <<<<<  Acumulado en PLAN  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos de Devolución (a pagar) / Ajuste Matrícula
                                                    Regs_Dev = Regs_Dev + 1:     Imp_Devo = Imp_Devo + .Cells(Cont, BD_ImpRec)
            If .Cells(Cont, BD_ImpCob) < 0 Then                         ' recibos Negativos de Devolución Pagados
                                                    RDev_Pag = RDev_Pag + 1:     IDev_Pag = IDev_Pag + .Cells(Cont, BD_ImpCob)
            End If
        Else                                                            ' recibos Emitidos, Cobrados o NO
                                                    RegsEmis = RegsEmis + 1:     Imp_Emis = Imp_Emis + .Cells(Cont, BD_ImpRec)
            If .Cells(Cont, BD_ImpCob) > 0 Then                         ' recibos Emitidos Cobrados
                                                    RegsCobr = RegsCobr + 1:     Imp_Cobr = Imp_Cobr + .Cells(Cont, BD_ImpCob)
            End If
        End If

Reg_Siguiente:
    Next
                    RegsPdts = RegsEmis - RegsCobr:             Imp_Pdte = Imp_Emis - Imp_Cobr
                    RDevPdte = Regs_Dev - RDev_Pag:             IDevPdte = Imp_Devo - IDev_Pag
                T_ImpEmi = T_ImpEmi + Imp_Emis:         T_ImpCob = T_ImpCob + Imp_Cobr:         T_ImpPdt = T_ImpPdt + Imp_Pdte
                RowNew.Range(Lst_ImpEmi) = Imp_Emis:        RowNew.Range(Lst_ImpEmiRgs) = RegsEmis
                RowNew.Range(Lst_ImpCob) = Imp_Cobr:        RowNew.Range(Lst_ImpCobRgs) = RegsCobr
                RowNew.Range(Lst_ImpPdt) = Imp_Pdte:        RowNew.Range(Lst_ImpPdtRgs) = RegsPdts
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                Right(String(8, " ") & Format(Imp_Emis, "#,##0.00"), 11) & Right("   (" & RegsEmis, 4) & ")" & _
                IIf(Imp_Cobr > 0, Right(String(8, " ") & Format(Imp_Cobr, "#,##0.00"), 12) & Right("   (" & RegsCobr, 4) & ")", " " & String(16, "·")) & _
                IIf(Imp_Pdte > 0, Right(String(8, " ") & Format(Imp_Pdte, "#,##0.00"), 12) & Right("   (" & RegsPdts, 4) & ")", " " & String(16, " "))
            If Imp_Cobr = 0 Then    ' No se ha pagado NADA
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & "   " & ChrW(8709) & "  Plan sin Cobros."
            End If

            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & Txt_Cabecera
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & " Totales "
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                Right(String(8, " ") & Format(T_ImpEmi, "#,##0.00"), 15) & _
                Right(String(8, " ") & Format(T_ImpCob, "#,##0.00"), 17) & _
                Right(String(8, " ") & Format(T_ImpPdt, "#,##0.00"), 17)
            
            If PlanesSinCob > 1 Then
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & "¡¡ " & PlanesSinCob & " Planes sin Cobros !!"
            ElseIf PlanesSinCob = 1 Then
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & "¡¡ 1 Plan sin Cobros !!"
            Else
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & "¡ No hay planes sin Cobros !"
            End If
    
'- Visualizo el progreso ---------------------------------------------------------------------------------------
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


