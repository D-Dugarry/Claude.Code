Attribute VB_Name = "M32_List_PLANES_Devoluc"
' Last Rev. 2026-09-21 12:12
'- M32_Listar_PLANES_Devoluc
Option Explicit
'===================================================================================================
Sub RuT_Listar_Planes_Devoluciones()
'===================================================================================================
    Dim Cont                As Long
    Dim ContIni             As Long:        ContIni = 1
    Dim Cod_Plan            As String
    Dim Txt_Cabecera        As String
    Dim Cont_Plan           As Integer:     Cont_Plan = 1
    Dim PlanesConDevol      As Integer:     PlanesConDevol = 0
    Dim SW_PlanConDevol     As Boolean:     SW_PlanConDevol = False
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
    
    Dim Lo_TitPH    As ListObject:      Set Lo_TitPH = Prog_BD.ListObjects(1)
    
    Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso

    ' =============  Preparar Tabla de TitPH ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_Sort(Lo_TitPH, BD_Plan, xlAscending, True)

With Lo_TitPH.DataBodyRange
    '- Tratando el 1º Recibo
    Do While ContIni <= Lo_TitPH.ListRows.Count And (.Cells(ContIni, BD_Tipo_Rec) = "Deleted" Or .Cells(ContIni, BD_ImpAdm) < 0)
        ContIni = ContIni + 1
    Loop
    Cont = ContIni
    Cod_Plan = .Cells(Cont, BD_Plan)
    If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos de Devolución (a pagar) / Ajuste Matrícula
        SW_PlanConDevol = True
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
    '- Visualizo el progreso -----------------------------------------------------------------------
    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TB_Informe = "Planes de " & Range("APP_EFP_o_CFC") & "_" & Range("APP_CursAcad") & " que tienen recibos de Devolución o de Ajuste de Matrícula.   " & Now & vbCrLf & vbCrLf
    Txt_Cabecera = "    Plan   Imp_Rec   r.  - Imp_Cob   r. = Pdte.Cob   r.    Imp_Dev  r.    Pagado       A pagar/Ajuste"
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera & vbLf
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Cont_Plan, "00") & "º " & Cod_Plan & " "
    
    ' Recorro toda la Tabla ------------------------------------------------------------------------
    For Cont = 2 To Lo_TitPH.ListRows.Count
        If .Cells(Cont, BD_Tipo_Rec) = "Deleted" Then GoTo Reg_Siguiente
        If .Cells(Cont, BD_ImpAdm) < 0 Then GoTo Reg_Siguiente              '- Ajustes de Matrícula que distorcionan la Contabilidad !!!
        If .Cells(Cont, BD_Plan) <> Cod_Plan Then
            Cod_Plan = .Cells(Cont, BD_Plan)
            Cont_Plan = Cont_Plan + 1
                    RegsPdts = RegsEmis - RegsCobr:             Imp_Pdte = Imp_Emis - Imp_Cobr
                    RDevPdte = Regs_Dev - RDev_Pag:             IDevPdte = Imp_Devo - IDev_Pag
            If Imp_Devo = 0 Then    ' Dejo  ' Plan sin Devoluciones
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                    Right(String(8, " ") & Format(Imp_Emis, "#,##0.00"), 10) & Right("   (" & RegsEmis, 4) & ")" & _
                    IIf(Imp_Cobr > 0, Right(String(8, " ") & Format(Imp_Cobr, "#,##0.00"), 11) & Right("   (" & RegsCobr, 4) & ")", " " & String(15, "·")) & _
                    IIf(Imp_Pdte > 0, Right(String(8, " ") & Format(Imp_Pdte, "#,##0.00"), 11) & Right("   (" & RegsPdts, 4) & ")", " " & String(15, " ")) & _
                    String(8, " ") & ChrW(8709) & "   - SIN Devoluciones -"
                If Cont_Plan Mod 20 = 1 Then Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & Txt_Cabecera
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Format(Cont_Plan, "00") & "º " & Cod_Plan & " "
            Else
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                    Right(String(8, " ") & Format(Imp_Emis, "#,##0.00"), 10) & Right("   (" & RegsEmis, 4) & ")" & _
                    IIf(Imp_Cobr > 0, Right(String(8, " ") & Format(Imp_Cobr, "#,##0.00"), 11) & Right("   (" & RegsCobr, 4) & ")", " " & String(15, "·")) & _
                    IIf(Imp_Pdte > 0, Right(String(8, " ") & Format(Imp_Pdte, "#,##0.00"), 11) & Right("   (" & RegsPdts, 4) & ")", " " & String(15, " ")) & _
                    IIf(Imp_Devo < 0, Right(String(8, " ") & Format(Imp_Devo, "#,##0.00"), 11) & Right("   (" & Regs_Dev, 3) & ")", " " & String(14, " ")) & _
                    IIf(IDev_Pag < 0, Right(String(8, " ") & Format(IDev_Pag, "#,##0.00"), 11) & Right("   (" & RDev_Pag, 3) & ")", " " & String(14, " ")) & _
                    IIf(IDevPdte < 0, Right(String(8, " ") & Format(IDevPdte, "#,##0.00"), 11) & Right("   (" & RDevPdte, 3) & ")", " ")
                If Cont_Plan Mod 20 = 1 Then Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & Txt_Cabecera
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Format(Cont_Plan, "00") & "º " & Cod_Plan & " "
            End If
            
            If SW_PlanConDevol Then PlanesConDevol = PlanesConDevol + 1
                SW_PlanConDevol = False
                                                        Regs_Dev = 0:       Imp_Devo = 0
                                                        RDev_Pag = 0:       IDev_Pag = 0
                                                        RegsEmis = 0:       Imp_Emis = 0
                                                        RegsCobr = 0:       Imp_Cobr = 0
            If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos de Devolución (a pagar) / Ajuste Matrícula
                SW_PlanConDevol = True
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
        Else
            If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos de Devolución (a pagar) / Ajuste Matrícula
                SW_PlanConDevol = True
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
        End If
Reg_Siguiente:
    Next
                    RegsPdts = RegsEmis - RegsCobr:             Imp_Pdte = Imp_Emis - Imp_Cobr
                    RDevPdte = Regs_Dev - RDev_Pag:             IDevPdte = Imp_Devo - IDev_Pag
            If Imp_Devo = 0 Then    ' Dejo  ' Plan sin Devoluciones
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                    Right(String(8, " ") & Format(Imp_Emis, "#,##0.00"), 10) & Right("   (" & RegsEmis, 4) & ")" & _
                    IIf(Imp_Cobr > 0, Right(String(8, " ") & Format(Imp_Cobr, "#,##0.00"), 11) & Right("   (" & RegsCobr, 4) & ")", " " & String(15, "·")) & _
                    IIf(Imp_Pdte > 0, Right(String(8, " ") & Format(Imp_Pdte, "#,##0.00"), 11) & Right("   (" & RegsPdts, 4) & ")", " " & String(15, " ")) & _
                    String(8, " ") & ChrW(8709) & "   - SIN Devoluciones -"
            Else
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                    Right(String(8, " ") & Format(Imp_Emis, "#,##0.00"), 10) & Right("   (" & RegsEmis, 4) & ")" & _
                    IIf(Imp_Cobr > 0, Right(String(8, " ") & Format(Imp_Cobr, "#,##0.00"), 11) & Right("   (" & RegsCobr, 4) & ")", " " & String(15, "·")) & _
                    IIf(Imp_Pdte > 0, Right(String(8, " ") & Format(Imp_Pdte, "#,##0.00"), 11) & Right("   (" & RegsPdts, 4) & ")", " " & String(15, " ")) & _
                    IIf(Imp_Devo < 0, Right(String(8, " ") & Format(Imp_Devo, "#,##0.00"), 11) & Right("   (" & Regs_Dev, 3) & ")", " " & String(14, " ")) & _
                    IIf(IDev_Pag < 0, Right(String(8, " ") & Format(IDev_Pag, "#,##0.00"), 11) & Right("   (" & RDev_Pag, 3) & ")", " " & String(14, " ")) & _
                    IIf(IDevPdte < 0, Right(String(8, " ") & Format(IDevPdte, "#,##0.00"), 11) & Right("   (" & RDevPdte, 3) & ")", " ")
            End If
            
            If SW_PlanConDevol Then PlanesConDevol = PlanesConDevol + 1
            If PlanesConDevol > 0 Then
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & "¡¡ " & PlanesConDevol & " Planes con Devoluciones / Ajustes de Matrícula !!"
            Else
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & "¡ No hay planes con Devoluciones / Ajustes de Matrícula !"
            End If
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "     (Pueden ser Devoluciones de Matrícula o Ajustes de Matrícula)"
    
'- Visualizo el progreso ---------------------------------------------------------------------------
Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & _
                                                "¡¡¡ Proceso concluido !!! día: " & Now() & _
                                                " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg."
End With    '- Lo_TitPH.DataBodyRange

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA
Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Listar_Planes   ------------------------------------------------------------------
'===================================================================================================


