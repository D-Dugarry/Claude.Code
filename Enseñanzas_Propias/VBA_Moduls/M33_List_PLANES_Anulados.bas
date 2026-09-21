Attribute VB_Name = "M33_List_PLANES_Anulados"
' Last Rev. 2026-09-21 12:12
'- M33_Listar_PLANES_Anulados ----------------------------------------------------------------------
Option Explicit
'===================================================================================================
Sub RuT_Listar_Planes_Anul()
'===================================================================================================
    Dim Cont                As Long
    Dim ContIni             As Long:        ContIni = 1
    Dim Cod_Plan            As String
    Dim Txt_Cabecera        As String
    Dim Cont_Plan           As Integer:     Cont_Plan = 1
    Dim PlanesConAnulados   As Integer:     PlanesConAnulados = 0
    
    Dim RegsEmis            As Long     ' regs Emitidos
    Dim Imp_Emis            As Currency
    Dim RegsCobr            As Long     ' regs Cobrados
    Dim Imp_Cobr            As Currency
    Dim RegsPdts            As Long     ' regs Pdtes de Cobro
    Dim Imp_Pdte            As Currency
    
    Dim RegsAnul            As Long     ' regs Anulados
    Dim Imp_Anul            As Currency
    Dim RgAnuCob            As Long     ' regs Anulados Cobrados
    Dim IpAnuCob            As Currency
    Dim RgAnuPdt            As Long     ' regs Anulados Pendientes de Pago
    Dim IpAnuPdt            As Currency
    
    Dim RgAnuDev            As Long     ' regs Anulados Devolución Ordenada
    Dim IpAnuDev            As Currency
    Dim RAnDePag            As Long     ' regs Anulados Devolución Ordenada y Pagada
    Dim IAnDePag            As Currency
    Dim RAnDePdt            As Long     ' regs Anulados Devolución Ordenada y Pendiente de pago
    Dim IAnDePdt            As Currency
    
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
    If .Cells(Cont, BD_Anul) = "S" Then             ' Recibos Anulados Negativos para Devolución pagados o NO y también para ¡RESTAR A LA MATRÍCULA!
        PlanesConAnulados = PlanesConAnulados + 1
        If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos Devolución (a pagar) / Ajuste Matrícula
                                                    RgAnuDev = 1:       IpAnuDev = .Cells(Cont, BD_ImpRec)
            If .Cells(Cont, BD_ImpCob) < 0 Then                         ' recibos Negativos de Devolución de importes Pagados
                                                    RAnDePag = 1:       IAnDePag = .Cells(Cont, BD_ImpCob)
            End If
        Else                                        ' Recibos Anulados Positivos, Cobrados o NO, los Cobrados, pueden ser objeto de Devolución
                                                    RegsAnul = 1:       Imp_Anul = .Cells(Cont, BD_ImpRec)
            If .Cells(Cont, BD_ImpCob) > 0 Then     ' Recibos Anulados Positivos Cobrados, pueden ser objeto de Devolución
                                                    RgAnuCob = 1:       IpAnuCob = .Cells(Cont, BD_ImpCob)
            End If
        End If
    End If
    '- Visualizo el progreso -----------------------------------------------------------------------
    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TB_Informe = "Planes de " & Range("APP_EFP_o_CFC") & "_" & Range("APP_CursAcad") & " que tienen recibos Anulados. (Dev. Matrícula)       " & Now & vbCrLf & vbCrLf
    Txt_Cabecera = "    Plan    Imp_Rec  r. -  Imp_Cob  r.  =   Pdte.  r.    Imp_Dev  r.     Pagado  r.   Pde. Pago r."
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera & vbLf
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Cont_Plan, "00") & "º " & Cod_Plan & " "
    
    ' Recorro toda la Tabla ------------------------------------------------------------------------
    For Cont = 2 To Lo_TitPH.ListRows.Count
        If .Cells(Cont, BD_Tipo_Rec) = "Deleted" Then GoTo Reg_Siguiente
        If .Cells(Cont, BD_ImpAdm) < 0 Then GoTo Reg_Siguiente
        If .Cells(Cont, BD_Plan) <> Cod_Plan Then
            Cod_Plan = .Cells(Cont, BD_Plan)
            Cont_Plan = Cont_Plan + 1
            If RgAnuDev = 0 And RegsEmis = 0 Then   ' Dejo la línea en blanco
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & ChrW(8709)
                If Cont_Plan Mod 20 = 1 Then Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & Txt_Cabecera
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Format(Cont_Plan, "00") & "º " & Cod_Plan & " "
            Else
                        RgAnuPdt = RegsAnul - RgAnuCob:             IpAnuPdt = Imp_Anul - IpAnuCob
                        RAnDePdt = RgAnuDev - RAnDePag:             IAnDePdt = IpAnuDev - IAnDePag
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                    Right(String(8, " ") & Format(Imp_Anul, "#,##0.00"), 11) & Right("   (" & RegsAnul, 3) & ")" & _
                    IIf(IpAnuCob > 0, Right(String(8, " ") & Format(IpAnuCob, "#,##0.00"), 11) & Right("   (" & RgAnuCob, 3) & ")", " " & String(14, "·")) & _
                    IIf(IpAnuPdt > 0, Right(String(8, " ") & Format(IpAnuPdt, "#,##0.00"), 11) & Right("   (" & RgAnuPdt, 3) & ")", " " & String(14, " ")) & _
                    IIf(IpAnuDev < 0, Right(String(8, " ") & Format(IpAnuDev, "#,##0.00"), 11) & Right("   (" & RgAnuDev, 3) & ")", " " & String(14, "·")) & _
                    IIf(IAnDePag < 0, Right(String(8, " ") & Format(IAnDePag, "#,##0.00"), 11) & Right("   (" & RAnDePag, 3) & ")", " " & String(14, "·")) & _
                    IIf(IAnDePdt < 0, Right(String(8, " ") & Format(IAnDePdt, "#,##0.00"), 11) & Right("   (" & RAnDePdt, 3) & ")", "")
                If Cont_Plan Mod 20 = 1 Then Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & Txt_Cabecera
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Format(Cont_Plan, "00") & "º " & Cod_Plan & " "
            End If
                
                                                            RgAnuDev = 0:       IpAnuDev = 0
                                                            RAnDePag = 0:       IAnDePag = 0
                                                            RegsAnul = 0:       Imp_Anul = 0
                                                            RgAnuCob = 0:       IpAnuCob = 0
            
            If .Cells(Cont, BD_Anul) = "S" Then             ' Recibos Anulados Negativos
                PlanesConAnulados = PlanesConAnulados + 1
                If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos de Devolución de importe
                                                            RgAnuDev = 1:       IpAnuDev = .Cells(Cont, BD_ImpRec)
                    If .Cells(Cont, BD_ImpCob) < 0 Then                         ' recibos Negativos de Devolución de importes Pagados
                                                            RAnDePag = 1:       IAnDePag = .Cells(Cont, BD_ImpCob)
                    End If
                Else                                        ' Recibos Anulados Positivos
                                                            RegsAnul = 1:       Imp_Anul = .Cells(Cont, BD_ImpRec)
                    If .Cells(Cont, BD_ImpCob) > 0 Then     ' Recibos Anulados Positivos Cobrados
                                                            RgAnuCob = 1:       IpAnuCob = .Cells(Cont, BD_ImpCob)
                    End If
                End If
            End If
        Else
            If .Cells(Cont, BD_Anul) = "S" Then             ' Recibos Anulados Negativos
                PlanesConAnulados = PlanesConAnulados + 1
                If .Cells(Cont, BD_ImpRec) < 0 Then                             ' recibos Negativos de Devolución de importe
                                                            RgAnuDev = RgAnuDev + 1:     IpAnuDev = IpAnuDev + .Cells(Cont, BD_ImpRec)
                    If .Cells(Cont, BD_ImpCob) < 0 Then                         ' recibos Negativos de Devolución de importes Pagados
                                                            RAnDePag = RAnDePag + 1:     IAnDePag = IAnDePag + .Cells(Cont, BD_ImpCob)
                    End If
                Else                                        ' Recibos Anulados Positivos
                                                            RegsAnul = RegsAnul + 1:     Imp_Anul = Imp_Anul + .Cells(Cont, BD_ImpRec)
                    If .Cells(Cont, BD_ImpCob) > 0 Then     ' Recibos Anulados Positivos Cobrados
                                                            RgAnuCob = RgAnuCob + 1:     IpAnuCob = IpAnuCob + .Cells(Cont, BD_ImpCob)
                    End If
                End If
            End If
        End If
Reg_Siguiente:
    Next
        If RgAnuDev = 0 And RegsEmis = 0 Then   ' Dejo la línea en blanco
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & ChrW(8709)
        Else
                        RgAnuPdt = RegsAnul - RgAnuCob:             IpAnuPdt = Imp_Anul - IpAnuCob
                        RAnDePdt = RgAnuDev - RAnDePag:             IAnDePdt = IpAnuDev - IAnDePag
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                    Right(String(8, " ") & Format(Imp_Anul, "#,##0.00"), 11) & Right("   (" & RegsAnul, 3) & ")" & _
                    IIf(IpAnuCob > 0, Right(String(8, " ") & Format(IpAnuCob, "#,##0.00"), 11) & Right("   (" & RgAnuCob, 3) & ")", " " & String(14, "·")) & _
                    IIf(IpAnuPdt > 0, Right(String(8, " ") & Format(IpAnuPdt, "#,##0.00"), 11) & Right("   (" & RgAnuPdt, 3) & ")", " " & String(14, " ")) & _
                    IIf(IpAnuDev < 0, Right(String(8, " ") & Format(IpAnuDev, "#,##0.00"), 11) & Right("   (" & RgAnuDev, 3) & ")", " " & String(14, "·")) & _
                    IIf(IAnDePag < 0, Right(String(8, " ") & Format(IAnDePag, "#,##0.00"), 11) & Right("   (" & RAnDePag, 3) & ")", " " & String(14, "·")) & _
                    IIf(IAnDePdt < 0, Right(String(8, " ") & Format(IAnDePdt, "#,##0.00"), 11) & Right("   (" & RAnDePdt, 3) & ")", "")
        End If
            
        If PlanesConAnulados > 0 Then
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & "¡¡ " & PlanesConAnulados & " Planes con Recibos Anulados !!" & vbLf & _
                                  "    (A priori, son Devoluciones de Matrícula.)"
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & "¡ No hay planes anulados !"
        End If
    
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

