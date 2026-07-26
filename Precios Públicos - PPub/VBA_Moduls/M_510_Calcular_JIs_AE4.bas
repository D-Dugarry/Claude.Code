Attribute VB_Name = "M_510_Calcular_JIs_AE4"
'- M_510_Generar_JIs_AE4

Option Explicit

            ' ==================================================================================================================================
            Sub Call_Rut_Recalcular_Tabla_JIs_AE4()
                Debug.Print "================== >>> Call_Rut_Recalcular_Tabla_JIs_AE4"
                Dim ActivSheet  As String:  ActivSheet = ThisWorkbook.ActiveSheet.Name
                Prog__APP.Range("APP_Task_Rut") = "Rut_Recalcular_Tabla_JIs_AE4"
            On Error GoTo ManejoError
                DoEvents ' Permite que Excel procese eventos pendientes
                Form_Running_Rut.Show
            On Error GoTo 0
                ThisWorkbook.Sheets(ActivSheet).Select
                Debug.Print "================== <<< Call_Rut_Recalcular_Tabla_JIs_AE4"
            Exit Sub
ManejoError:
                Static Intentos As Integer
                If Err.Number = -2147417848 And Intentos < 5 Then
                    Intentos = Intentos + 1
                    Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
                    Resume ' Reintenta la línea que falló
                Else
                    MsgBox "Error: " & Err.Description & vbCrLf & "Intentos: " & Intentos, vbCritical
                    Intentos = 0
                End If
                MsgBox "<<< Err_Rut Form_Running_Rut >>>"
            End Sub

' ==================================================================================================================================
Sub Rut_Recalcular_Tabla_JIs_AE4()
' ==================================================================================================================================
Debug.Print "Rut_Recalcular_Tabla_JIs_AE4"

    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    'Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

Dim Concept         As Variant
Dim Concept2        As String
Dim Tp_Rec          As String
Dim TipRec_Cncpt    As String:      TipRec_Cncpt = ""
Dim DescripciónA    As String
Dim DescripciónB    As String
Dim MenúAux_Msg     As String
Dim F_Concept       As Integer

Dim Clv             As String
Dim ClvAnt          As String:          ClvAnt = ""
Dim CAcadAnt        As String:          CAcadAnt = Prog__APP.Range("APP_C_Acad_Ant")
Dim CAcadPos        As String:          CAcadPos = Prog__APP.Range("APP_C_Acad_Pos")
Dim PlanAnt         As String:          PlanAnt = ""
Dim KolorPlan       As Integer:         KolorPlan = 36
Dim KolorCacad      As Integer:         KolorCacad = 19

Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
Dim Lo_PlanAE4      As ListObject:      Set Lo_PlanAE4 = Sht__BD_JIs_AE4.ListObjects(1)
Dim Lo_Concept      As ListObject:      Set Lo_Concept = Prog_Concept.ListObjects("Tb_Conceptos")
Dim RwPln           As ListRow
Dim F_BD            As Long
Dim TF_BD           As Long:        TF_BD = Lo_BD.ListRows.Count

    Sht__BD.Visible = xlSheetVisible
    Sht__BD.Unprotect
    Lo_BD.ShowTotals = False
    Sht__BD_JIs_AE4.Visible = xlSheetVisible
    
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '------------ Preparo Sht__BD y Ordeno  -----------------------------------
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_JIs_AE4)
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Call Rut_Lo_Sort(Lo_BD, BD_ActivEco, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_Tipo_Rec, xlAscending, False)
    
    '-------------- Relleno Tabla para JI's de Tasas por Tipo y por Concepto -----------------------------------
    If Not Lo_PlanAE4.DataBodyRange Is Nothing Then Lo_PlanAE4.DataBodyRange.Delete
    Sht__BD_JIs_AE4.Select
    Sht__BD_JIs_AE4.Unprotect
    Lo_PlanAE4.ShowTotals = False
'    With Lo_PlanAE4.HeaderRowRange
'        .Cells(AE4_JI_Ant) = "JI-" & Prog__APP.Range("APP_AñoCont") - 1
'        .Cells(AE4_AD_Ant) = "AD-" & Prog__APP.Range("APP_AñoCont") - 1
'        .Cells(AE4_JI_Actual) = "JI-" & Prog__APP.Range("APP_AñoCont")
'        .Cells(AE4_AD_Actual) = "AD-" & Prog__APP.Range("APP_AñoCont")
'    End With
    With Lo_BD.DataBodyRange
    For F_BD = 1 To TF_BD
        If .Cells(F_BD, BD_ActivEco) < 4 Then GoTo SigRec
        If .Cells(F_BD, BD_ActivEco) > 4 Then Exit For
        Clv = .Cells(F_BD, BD_Plan) & .Cells(F_BD, BD_C_Acad) & .Cells(F_BD, BD_Tipo_Rec)
        If Clv <> ClvAnt Then
            ClvAnt = Clv
            Set RwPln = Lo_PlanAE4.ListRows.Add
            RwPln.Range(AE4_TipRec) = .Cells(F_BD, BD_Tipo_Rec)
            RwPln.Range(AE4_Plan) = .Cells(F_BD, BD_Plan)
            RwPln.Range(AE4_PlanNom) = .Cells(F_BD, BD_NomPlan)
            RwPln.Range(AE4_C_Acad) = .Cells(F_BD, BD_C_Acad)
'            RwPln.Range(AE4_ConcptEco) = Format(.Cells(F_BD, BD_Concepto), "000.00")
            RwPln.Range(AE4_ConcptEco) = Format(.Cells(F_BD, BD_Concepto), "000.00")
            '- Importe Recibos --------------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Tot_Emi) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                    .Columns(BD_Concepto), .Cells(F_BD, BD_Concepto), _
                                                    .Columns(BD_Plan), RwPln.Range(AE4_Plan), _
                                                    .Columns(BD_C_Acad), RwPln.Range(AE4_C_Acad), _
                                                    .Columns(BD_Tipo_Rec), RwPln.Range(AE4_TipRec))
            '- Importe Cobrado --------------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Tot_Cob) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                    .Columns(BD_Concepto), .Cells(F_BD, BD_Concepto), _
                                                    .Columns(BD_Plan), RwPln.Range(AE4_Plan), _
                                                    .Columns(BD_C_Acad), RwPln.Range(AE4_C_Acad), _
                                                    .Columns(BD_Tipo_Rec), RwPln.Range(AE4_TipRec))
            '- Saldo Contable ---------------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Tot_Pdte) = RwPln.Range(AE4_Tot_Emi) - RwPln.Range(AE4_Tot_Cob)

            '- Importe Adm ------------------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Adm_Emi) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_Concepto), .Cells(F_BD, BD_Concepto), _
                                                    .Columns(BD_Plan), RwPln.Range(AE4_Plan), _
                                                    .Columns(BD_C_Acad), RwPln.Range(AE4_C_Acad), _
                                                    .Columns(BD_Tipo_Rec), RwPln.Range(AE4_TipRec))
            '- Importe Adm. Cobrado ---------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Adm_Cob) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_Concepto), .Cells(F_BD, BD_Concepto), _
                                                    .Columns(BD_ImpCob), ">0", _
                                                    .Columns(BD_Plan), RwPln.Range(AE4_Plan), _
                                                    .Columns(BD_C_Acad), RwPln.Range(AE4_C_Acad), _
                                                    .Columns(BD_Tipo_Rec), RwPln.Range(AE4_TipRec))
            '- Saldo Imp. Adm. --------------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Adm_Pdte) = RwPln.Range(AE4_Adm_Emi) - RwPln.Range(AE4_Adm_Cob)

            '- Imp. Acad. -------------------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Acad_Emi) = RwPln.Range(AE4_Tot_Emi) - RwPln.Range(AE4_Adm_Emi)
            '- Imp. Acad. Cobrado ------------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Acad_Cob) = RwPln.Range(AE4_Tot_Cob) - RwPln.Range(AE4_Adm_Cob)
            '- Saldo Imp. Acad. ------------------------------------------------------------------------------------------------------------------------
            RwPln.Range(AE4_Acad_Pdte) = RwPln.Range(AE4_Acad_Emi) - RwPln.Range(AE4_Acad_Cob)
            
'''            Dim RowFind      As Variant
'''            RowFind = Application.Match(.Cells(F_BD, BD_Concepto), Lo_Concept.DataBodyRange.Columns(1), 0)
'''            If IsError(RowFind) Then     ' ¡¡¡ NO Existe !!! ------------------------
'''                MsgBox "¡ No Existe El Concepto Eco. !"
'''            Else                            ' ¡¡¡ Existe la Rutina !!! ------------------------
'''                DescripciónB = Lo_Concept.DataBodyRange.Cells(RowFind, 2)
'''            End If
                    
            RwPln.Range(AE4_Descrip_JI) = "Liq.PPub_" & Right(RwPln.Range(AE4_ConcptEco), 6) & "__" & _
                    RwPln.Range(AE4_TipRec) & "_" & Prog__APP.Range("APP_AñoCont") & "__Plan_" & _
                    RwPln.Range(AE4_Plan) & "_C_Acad_" & RwPln.Range(AE4_C_Acad) & "_" & RwPln.Range(AE4_PlanNom)
            
            If PlanAnt <> RwPln.Range(AE4_Plan) Then
                PlanAnt = RwPln.Range(AE4_Plan)
                If KolorPlan = 36 Then KolorPlan = 37 Else KolorPlan = 36
            End If
'            If CAcadAnt <> RwPln.Range(AE4_C_Acad) Then
'                CAcadAnt = RwPln.Range(AE4_C_Acad)
'                If KolorCacad = 19 Then KolorCacad = 20 Else KolorCacad = 19
'            End If
            If CAcadAnt = RwPln.Range(AE4_C_Acad) Then
                KolorCacad = 19
            ElseIf CAcadPos = RwPln.Range(AE4_C_Acad) Then
                KolorCacad = 20
            Else
                KolorCacad = 24
            End If
        
            RwPln.Range.RowHeight = 25
            RwPln.Range.VerticalAlignment = xlCenter
            RwPln.Range.Interior.ColorIndex = KolorPlan
            RwPln.Range(AE4_C_Acad).Interior.ColorIndex = KolorCacad
            RwPln.Range(AE4_TipRec).Interior.ColorIndex = KolorCacad
        
        End If
            
SigRec:
        If F_BD Mod 1000 = 0 Then
            Debug.Print Format(Now, "hh:mm:ss") & " Revizado BDatos:  " & Format(F_BD, "#,##0") & " de " & Format(TF_BD, "#,##0") & " reg."
'            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
        End If
'            RwPln.Range(AE4_Descrip_JI) = "Liq.PPub_" & Concept2 & "__" & Tp_Rec & "_" & Prog__APP.Range("APP_AñoCont") & "__" & DescripciónB & ".  "
'
    Next F_BD
    End With        '- Lo_BD.DataBodyRange
    
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_JIs_AE4)
    
    Lo_PlanAE4.DataBodyRange.Columns(AE4_Descrip_JI).Select
    Selection.InsertIndent 1
    Lo_PlanAE4.ShowTotals = True
    
    Sht__BD_JIs_AE4.Calculate
    Sht__BD_JIs_AE4.Range("d2") = "Último Cálculo: " & Format(Now, "dd mmmm yyyy - hh:mm")
    Sht__BD_JIs_AE4.Range("d2").Select
    
    Prog__APP.Range("APP_Last_Calc_JIs_AE4") = Format(Now(), "dd-mmm-yy hh:mm")
    
    '- Visualizo el progreso ---------------------------------------------------------------------------------------
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MenúAux_Msg = Format(Now, "hh:mm:ss") & "  Tabla generada." & vbCrLf & _
        vbCrLf & Format(Now, "hh:mm:ss") & "  Realizado el: " & Date & "  " & "-   Tiempo transcurrido: " & Round(Timer - H_Inicio, 2) & " seg."
    MsgBox MenúAux_Msg
    Application.Calculation = Sw_Calculation
    Application.Speech.Speak "Proceso completado puede verificar el resultado.", True
End Sub
' ==================================================================================================================================


