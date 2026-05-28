Attribute VB_Name = "M_800_INICIO_Inf_Liq_Summary"
Option Explicit

' ------------------------------------------------------------------------------------------------------
Sub Rut_Inicio_Lo_Liq_Summary()
' ------------------------------------------------------------------------------------------------------
    Dim Color       As Integer:     Color = 19
    Dim LinX        As Long
    Dim LinY        As Long
    Dim RefAnt      As String:      RefAnt = ""
    Dim SiglasAnt   As String:      SiglasAnt = ""
    Dim TPV_Visibl  As Boolean:     TPV_Visibl = Prog_TPV_Tb.Visible
    Dim Cta_Visibl  As Boolean:     Cta_Visibl = Prog_CTA_Tb.Visible
    Dim INI_Visibl  As Boolean:     INI_Visibl = H_INICI.Visible
    Dim RowTPV          As ListRow
    Dim RowCta          As ListRow
    Dim RowINI          As ListRow
    Dim Lo_TPV      As ListObject:      Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Dim Lo_Cta      As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    Dim Lo_INI      As ListObject:      Set Lo_INI = H_INICI.ListObjects(1)
    Dim Lo_JCL      As ListObject:      Set Lo_JCL = Prog_JyC_List.ListObjects(1)
    Dim VisRng          As Range
    Dim RngTPV          As Range
    Dim RngCta          As Range
    Dim Lin_Reg     As Variant

    Application.ScreenUpdating = False
    H_INICI.Visible = xlSheetVisible
    H_INICI.Select
    H_INICI.Unprotect
    Call Rut_LstObj_Filtros_Quitar(Lo_INI)
    If Not Lo_INI.DataBodyRange Is Nothing Then Lo_INI.DataBodyRange.Delete

'- Genero la lista de TPV ---------------
    Prog_TPV_Tb.Visible = xlSheetVisible
    Call Rut_LstObj_WrkSht_Preparar(Prog_TPV_Tb)
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_Siglas, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_N_Liq, xlAscending, False)
    For LinX = 1 To Lo_TPV.ListRows.Count
        Set RowTPV = Lo_TPV.ListRows(LinX)
        If RowTPV.Range(C_TPV_Siglas) = "" Then GoTo SigLinX
        If RowTPV.Range(C_TPV_F_Pago) = "" Then GoTo SigLinX
        If Val(Right(RowTPV.Range(C_TPV_F_Pago), 4)) <> Prog__APP.Range("App_AñoCont") Then GoTo SigLinX
        If RefAnt <> RowTPV.Range(C_TPV_Siglas) & RowTPV.Range(C_TPV_N_Liq) Then
            RefAnt = RowTPV.Range(C_TPV_Siglas) & RowTPV.Range(C_TPV_N_Liq)
            Set RowINI = Lo_INI.ListRows.Add
            RowINI.Range(C_INI_Siglas) = RowTPV.Range(C_TPV_Siglas)
            RowINI.Range(C_INI_N_Liq) = RowTPV.Range(C_TPV_N_Liq)
            RowINI.Range(C_INI_Org) = RowTPV.Range(C_TPV_Org)
            RowINI.Range(C_INI_CtaTPV) = "TPV"
        End If
SigLinX:
    Next LinX
    Prog_TPV_Tb.Visible = TPV_Visibl
'- Genero la lista de Cta ---------------
    Prog_CTA_Tb.Visible = xlSheetHidden
    Call Rut_LstObj_WrkSht_Preparar(Prog_CTA_Tb)
    Call Rut_LstObj_Columns_Show(Prog_CTA_Tb, Sheets("DefCol_" & Prog_CTA_Tb.Name))
    Call Rut_LstObj_Sort(Lo_Cta, C_Cta_Siglas, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_Cta, C_Cta_N_Liq, xlAscending, False)
    For LinX = 1 To Lo_Cta.ListRows.Count
        Set RowCta = Lo_Cta.ListRows(LinX)
        If RowCta.Range(C_Cta_N_Liq) = "" Then GoTo SigLinX2
    '    If UCase(RowCta.Range(C_Cta_N_Liq)) = "X" Then GoTo SigLinX2
        If RowCta.Range(C_Cta_N_Liq) = "--" Then GoTo SigLinX2
        If RowCta.Range(C_Cta_Siglas) = "" Then GoTo SigLinX2
        If RowCta.Range(C_Cta_Siglas) Like "_*" Then GoTo SigLinX2
        If Val(Right(RowCta.Range(C_Cta_F_VAL), 4)) <> Prog__APP.Range("App_AñoCont") Then GoTo SigLinX2
        If RefAnt <> RowCta.Range(C_Cta_Siglas) & RowCta.Range(C_Cta_N_Liq) Then
            RefAnt = RowCta.Range(C_Cta_Siglas) & RowCta.Range(C_Cta_N_Liq)
            Set RowINI = Lo_INI.ListRows.Add
            RowINI.Range(C_INI_Siglas) = RowCta.Range(C_Cta_Siglas)
            RowINI.Range(C_INI_N_Liq) = RowCta.Range(C_Cta_N_Liq)
            RowINI.Range(C_INI_Org) = RowCta.Range(C_Cta_Org)
            RowINI.Range(C_INI_CtaTPV) = "Cta"
        End If
SigLinX2:
    Next LinX
    
'- Recorro todas la filas de Lo_INI para rellenar las columnas de acumulados de importes  -----------------------
    Prog_CTA_Tb.Visible = Prog_CTA_Tb.Visible
    Call Rut_LstObj_Sort(Lo_INI, C_INI_N_Liq, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_INI, C_INI_Siglas, xlAscending, True)
    For LinX = 1 To Lo_INI.ListRows.Count
        Set RowINI = Lo_INI.ListRows(LinX)
        '-- TPV -----------------------------------------------------------------------------------------------------
        '-- TPV -----------------------------------------------------------------------------------------------------
        '-- TPV -----------------------------------------------------------------------------------------------------
        If RowINI.Range(C_INI_CtaTPV) = "TPV" Then
            Call Rut_LstObj_Filtro(Lo_TPV, C_TPV_Siglas, RowINI.Range(C_INI_Siglas), True)
            Call Rut_LstObj_Filtro(Lo_TPV, C_TPV_Ref, "=99" & Mid(Prog__APP.Range("App_AñoCont"), 3, 2) & "*", False)
            Call Rut_LstObj_Filtro(Lo_TPV, C_TPV_N_Liq, "=" & RowINI.Range(C_INI_N_Liq) & "*", False)
            Set VisRng = Lo_TPV.DataBodyRange.SpecialCells(xlCellTypeVisible)
            For LinY = 1 To VisRng.Rows.Count
                Set RngTPV = VisRng.Rows(LinY)
                If LinY = 1 Then
                    '- Acumular Núm de Ref-Sol ----
                    RowINI.Range(C_INI_Ref_Sol) = LCase(RngTPV.Cells(C_TPV_Ref_Sol))
                    '- Acumular JI ----
                    RowINI.Range(C_INI_JI) = RngTPV.Cells(C_TPV_JI)
                    '- Acumular Orgánica --------------------------------
                    RowINI.Range(C_INI_Org) = RngTPV.Cells(C_TPV_Org)
                    '- Acumular JI ----
                    RowINI.Range(C_INI_JI) = RngTPV.Cells(C_TPV_JI)
                    '- Acumular ExpAdm ----
                    RowINI.Range(C_INI_ExpAdm) = RngTPV.Cells(C_TPV_ExpAdm)
                    '- Acumular RDT -------------------------------------
                    RowINI.Range(C_INI_N_RDT) = RngTPV.Cells(C_TPV_RDT)
                    '- Acumular DI ----
                    RowINI.Range(C_INI_DI) = RngTPV.Cells(C_TPV_DI)
                    '- Acumular PMP ----
                    RowINI.Range(C_INI_PMP) = RngTPV.Cells(C_TPV_PMP)
                    
                    RowINI.Range(C_INI_F_Fin) = RngTPV.Cells(C_TPV_F_Fin)
'''                    If RngTPV.Cells(C_TPV_Obs) <> "" Then RowINI.Range(C_INI_Obs) = "# " & RngTPV.Cells(C_TPV_Obs)
                End If
                    '- Acumular Importes ----
                        RowINI.Range(C_INI_ImpDI) = RowINI.Range(C_INI_ImpDI) + RngTPV.Cells(C_TPV_DI_Imp)
                    If RngTPV.Cells(C_TPV_Neto) > 0 And RngTPV.Cells(C_TPV_Imp) = "" Then
'                        RowINI.Range(C_INI_TPV_ImpRec) = RowINI.Range(C_INI_TPV_ImpRec) + RngTPV.Cells(C_TPV_Neto)
'''                        RowINI.Range(C_INI_TPV_RecEmi) = RowINI.Range(C_INI_TPV_RecEmi) + RngTPV.Cells(C_TPV_Neto)
                    Else
                        RowINI.Range(C_INI_TPV_ImpRec) = RowINI.Range(C_INI_TPV_ImpRec) + RngTPV.Cells(C_TPV_Imp)
'''                        RowINI.Range(C_INI_TPV_RecEmi) = RowINI.Range(C_INI_TPV_RecEmi) + RngTPV.Cells(C_TPV_Imp)
                    End If
                    RowINI.Range(C_INI_TPV_ComBco) = RowINI.Range(C_INI_TPV_ComBco) + RngTPV.Cells(C_TPV_ComBco)
                    RowINI.Range(C_INI_TPV_Neto) = RowINI.Range(C_INI_TPV_Neto) + Val(RngTPV.Cells(C_TPV_Neto))
                    If RngTPV.Cells(C_TPV_DI) <> "" Then RowINI.Range(C_INI_TPV_Dev) = RowINI.Range(C_INI_TPV_Dev) + RngTPV.Cells(C_TPV_DI_Imp)
                    '- Acumular Observaciones ----
                    If RngTPV.Cells(C_TPV_Obs) <> "" And InStr(RowINI.Range(C_INI_Obs), RngTPV.Cells(C_TPV_Obs)) = 0 Then
                        If RowINI.Range(C_INI_Obs) <> "" Then
                            RowINI.Range(C_INI_Obs) = RowINI.Range(C_INI_Obs) & vbLf & "# " & RngTPV.Cells(C_TPV_Obs)
                        Else
                            RowINI.Range(C_INI_Obs) = "# " & RngTPV.Cells(C_TPV_Obs)
                        End If
                    End If
                    '- Acumular Núm de Ref-Sol ----
                    If InStr(UCase(RowINI.Range(C_INI_Ref_Sol)), UCase(RngTPV.Cells(C_TPV_Ref_Sol))) = 0 And RowINI.Range(C_TPV_Ref_Sol) <> "" _
                       Then RowINI.Range(C_INI_Ref_Sol) = RowINI.Range(C_INI_Ref_Sol) & vbLf & LCase(RngTPV.Cells(C_TPV_Ref_Sol))
                    '- Acumular Orgánica --------------------------------
                    If InStr(UCase(RowINI.Range(C_INI_Org)), UCase(RngTPV.Cells(C_TPV_Org))) = 0 And RowINI.Range(C_TPV_Org) <> "" _
                       Then RowINI.Range(C_INI_Org) = RowINI.Range(C_INI_Org) & vbLf & RngTPV.Cells(C_TPV_Org)
                    '- Acumular JI ----
                    If InStr(UCase(RowINI.Range(C_INI_JI)), UCase(RngTPV.Cells(C_TPV_JI))) = 0 And RowINI.Range(C_TPV_JI) <> "" _
                       Then RowINI.Range(C_INI_JI) = RowINI.Range(C_INI_JI) & vbLf & RngTPV.Cells(C_TPV_JI)
                    '- Acumular ExpAdm ----
                    If InStr(UCase(RowINI.Range(C_INI_ExpAdm)), UCase(RngTPV.Cells(C_TPV_ExpAdm))) = 0 And RowINI.Range(C_TPV_ExpAdm) <> "" _
                       Then RowINI.Range(C_INI_ExpAdm) = RowINI.Range(C_INI_ExpAdm) & vbLf & RngTPV.Cells(C_TPV_ExpAdm)
                    '- Acumular RDT -------------------------------------
                    If InStr(UCase(RowINI.Range(C_INI_N_RDT)), UCase(RngTPV.Cells(C_TPV_RDT))) = 0 And RowINI.Range(C_TPV_RDT) <> "" _
                       Then RowINI.Range(C_INI_N_RDT) = RowINI.Range(C_INI_N_RDT) & vbLf & RngTPV.Cells(C_TPV_RDT)
                    '- Acumular DI ----
                    If InStr(UCase(RowINI.Range(C_INI_DI)), UCase(RngTPV.Cells(C_TPV_DI))) = 0 And RowINI.Range(C_TPV_DI) <> "" _
                       Then RowINI.Range(C_INI_DI) = RowINI.Range(C_INI_DI) & vbLf & RngTPV.Cells(C_TPV_DI)
                    '- Acumular PMP ----
                    If InStr(UCase(RowINI.Range(C_INI_PMP)), UCase(RngTPV.Cells(C_TPV_PMP))) = 0 And RowINI.Range(C_TPV_PMP) <> "" _
                       Then RowINI.Range(C_INI_PMP) = RowINI.Range(C_INI_PMP) & vbLf & RngTPV.Cells(C_TPV_PMP)
SigLinY:
            Next LinY
        
        Else
        '-- Cta -----------------------------------------------------------------------------------------------------
        '-- Cta -----------------------------------------------------------------------------------------------------
        '-- Cta -----------------------------------------------------------------------------------------------------
            Call Rut_LstObj_Filtro(Lo_Cta, C_Cta_Siglas, RowINI.Range(C_INI_Siglas), True)
            Call Rut_LstObj_Filtro(Lo_Cta, C_Cta_Ordinal, ">" & Prog__APP.Range("App_AñoCont") & "000000", False)
            Call Rut_LstObj_Filtro(Lo_Cta, C_Cta_N_Liq, "=" & RowINI.Range(C_INI_N_Liq) & "*", False)
            Set VisRng = Lo_Cta.DataBodyRange.SpecialCells(xlCellTypeVisible)
            For LinY = 1 To VisRng.Rows.Count
                Set RngCta = VisRng.Rows(LinY)
                If LinY = 1 Then
                    Lin_Reg = Application.Match(RngCta.Cells(C_Cta_Siglas), Lo_JCL.DataBodyRange.Columns(C_JCL_Siglas), 0)
                    If Not IsError(Lin_Reg) Then    ' DATO Encontrado ------------------------
                        RowINI.Range(C_INI_F_Fin) = Lo_JCL.DataBodyRange.Cells(Lin_Reg, C_JCL_F_Fin)
                    Else                            ' NO ENCONTRADO   ------------------------
                        RowINI.Range(C_INI_F_Fin) = "_Not Foud " & RngCta.Cells(C_Cta_Siglas)
                    End If
                    '- Acumular Núm de Ref-Sol ----
                    RowINI.Range(C_INI_Ref_Sol) = LCase(RngCta.Cells(C_Cta_Ref_Sol))
                    '- Acumular JI ----
                    RowINI.Range(C_INI_JI) = RngCta.Cells(C_Cta_JI)
                    '- Acumular Orgánica --------------------------------
                    RowINI.Range(C_INI_Org) = RngCta.Cells(C_Cta_Org)
                    '- Acumular JI ----
                    RowINI.Range(C_INI_JI) = RngCta.Cells(C_Cta_JI)
                    '- Acumular ExpAdm ----
                    RowINI.Range(C_INI_ExpAdm) = RngCta.Cells(C_Cta_ExpAdm)
                    '- Acumular RDT -------------------------------------
                    RowINI.Range(C_INI_N_RDT) = RngCta.Cells(C_Cta_RDT)
                    '- Acumular DI ----
                    RowINI.Range(C_INI_DI) = RngCta.Cells(C_Cta_DI)
                    '- Acumular PMP ----
                    RowINI.Range(C_INI_PMP) = RngCta.Cells(C_Cta_PMP)
                End If
'                    If UCase(RngCta.Cells(C_Cta_N_Liq)) = "X" Then RowINI.Range(C_INI_Cta_ImpRec) = RowINI.Range(C_INI_Cta_ImpRec) + RngCta.Cells(C_Cta_Imp)
                    '- Acumular Importes ----
                    If RngCta.Cells(C_Cta_Imp) > 0 Then
                        RowINI.Range(C_INI_Cta_ImpRec) = RowINI.Range(C_INI_Cta_ImpRec) + RngCta.Cells(C_Cta_Imp)
                        RowINI.Range(C_INI_Cta_Imp) = RowINI.Range(C_INI_Cta_Imp) + RngCta.Cells(C_Cta_Imp)
                        RowINI.Range(C_INI_ImpDI) = RowINI.Range(C_INI_ImpDI) + RngCta.Cells(C_Cta_DI_Imp)
                    Else
                        If Left(RngCta.Cells(C_Cta_Siglas), 1) = "_" Then
                            RowINI.Range(C_INI_Cta_ImpRec) = RowINI.Range(C_INI_Cta_ImpRec) + RngCta.Cells(C_Cta_Imp)
                        Else
                            RowINI.Range(C_INI_Cta_Transf) = RowINI.Range(C_INI_Cta_Transf) + RngCta.Cells(C_Cta_Imp)
                        End If
                    End If
                    '- Acumular Observaciones ----
                    If RngCta.Cells(C_Cta_Obs) <> "" And InStr(RowINI.Range(C_INI_Obs), RngCta.Cells(C_Cta_Obs)) = 0 Then
                        If RowINI.Range(C_INI_Obs) <> "" Then
                            RowINI.Range(C_INI_Obs) = RowINI.Range(C_INI_Obs) & vbLf & "# " & RngCta.Cells(C_Cta_Obs)
                        Else
                            RowINI.Range(C_INI_Obs) = "# " & RngCta.Cells(C_Cta_Obs)
                        End If
                    End If
                    '- Acumular Núm de Ref-Sol ----
                    If InStr(UCase(RowINI.Range(C_INI_Ref_Sol)), UCase(RngCta.Cells(C_Cta_Ref_Sol))) = 0 And RngCta.Cells(C_Cta_Ref_Sol) <> "" _
                       Then RowINI.Range(C_INI_Ref_Sol) = RowINI.Range(C_INI_Ref_Sol) & vbLf & LCase(RngCta.Cells(C_Cta_Ref_Sol))
                    '- Acumular Orgánica --------------------------------
                    If InStr(UCase(RowINI.Range(C_INI_Org)), UCase(RngCta.Cells(C_Cta_Org))) = 0 And RngCta.Cells(C_Cta_Org) <> "" _
                       Then RowINI.Range(C_INI_Org) = RowINI.Range(C_INI_Org) & vbLf & RngCta.Cells(C_Cta_Org)
                    '- Acumular JI ----
                    If InStr(UCase(RowINI.Range(C_INI_JI)), UCase(RngCta.Cells(C_Cta_JI))) = 0 And RngCta.Cells(C_Cta_JI) <> "" _
                       Then RowINI.Range(C_INI_JI) = RowINI.Range(C_INI_JI) & vbLf & RngCta.Cells(C_Cta_JI)
                    '- Acumular ExpAdm ----
                    If InStr(UCase(RowINI.Range(C_INI_ExpAdm)), UCase(RngCta.Cells(C_Cta_ExpAdm))) = 0 And RngCta.Cells(C_Cta_ExpAdm) <> "" _
                       Then RowINI.Range(C_INI_ExpAdm) = RowINI.Range(C_INI_ExpAdm) & vbLf & RngCta.Cells(C_Cta_ExpAdm)
                    '- Acumular RDT -------------------------------------
                    If InStr(UCase(RowINI.Range(C_INI_N_RDT)), UCase(RngCta.Cells(C_Cta_RDT))) = 0 And RngCta.Cells(C_Cta_RDT) <> "" _
                       Then RowINI.Range(C_INI_N_RDT) = RowINI.Range(C_INI_N_RDT) & vbLf & RngCta.Cells(C_Cta_RDT)
                    '- Acumular DI ----
                    If InStr(UCase(RowINI.Range(C_INI_DI)), UCase(RngCta.Cells(C_Cta_DI))) = 0 And RngCta.Cells(C_Cta_DI) <> "" _
                       Then RowINI.Range(C_INI_DI) = RowINI.Range(C_INI_DI) & vbLf & RngCta.Cells(C_Cta_DI)
                    '- Acumular PMP ----
                    If InStr(UCase(RowINI.Range(C_INI_PMP)), UCase(RngCta.Cells(C_Cta_PMP))) = 0 And RngCta.Cells(C_Cta_PMP) <> "" _
                       Then RowINI.Range(C_INI_PMP) = RowINI.Range(C_INI_PMP) & vbLf & RngCta.Cells(C_Cta_PMP)
SigLinY2:
            Next LinY
        End If
    Next LinX
    
    '- Reordeno la Lo_INI, diferencio Filas-Siglas x Interior.ColorIndex ----------
    Call Rut_LstObj_Sort(Lo_INI, C_INI_N_Liq, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_INI, C_INI_Siglas, xlAscending, True)
        Lo_INI.DataBodyRange.Rows.AutoFit
'    Lo_INI.DataBodyRange.EntireRow.AutoFit
    '- Asigna color de fonda a cada grupo de JyC -----
    For LinX = 1 To Lo_INI.ListRows.Count
        Set RowINI = Lo_INI.ListRows(LinX)
        If RowINI.Range(C_INI_N_Liq) = "x" Then
            RowINI.Range(C_INI_N_Liq).Font.Bold = True
            RowINI.Range(C_INI_Cta_ImpRec).Font.Bold = True
            RowINI.Range(C_INI_Cta_Imp) = "_Pdte. Liq."
            RowINI.Range(C_INI_Cta_Imp).Font.Bold = True
            RowINI.Range(C_INI_PdteLiq) = RowINI.Range(C_INI_Cta_ImpRec)
        
            If RowINI.Range(C_INI_F_Fin) < Date Then
                RowINI.Range(C_INI_F_Fin).Font.ColorIndex = 9
                If RowINI.Range(C_INI_PdteLiq) > 0 Then
                    RowINI.Range(C_INI_PdteLiq).Font.Bold = True
                    RowINI.Range(C_INI_PdteLiq).Font.ColorIndex = 1
                End If
            Else
                RowINI.Range(C_INI_F_Fin).Font.Bold = True
            End If
        
        
        End If
        If SiglasAnt <> RowINI.Range(C_INI_Siglas) Then
            SiglasAnt = RowINI.Range(C_INI_Siglas)
            If Color = 19 Then Color = 20 Else Color = 19
        End If
        RowINI.Range.Interior.ColorIndex = Color
    Next LinX
    Lo_INI.TotalsRowRange.Interior.ColorIndex = 44
    Lo_INI.TotalsRowRange.Cells(1) = Format(Lo_INI.DataBodyRange.Rows.Count, "  0 reg.")
    
    H_INICI.Columns(Lo_INI.ListColumns(C_INI_N_Liq).Range.Column).Hidden = False
    H_INICI.Columns(Lo_INI.ListColumns(C_INI_Ref_Sol).Range.Column).Hidden = False
    
    With ActiveSheet.Shapes.Range(Array("Tit_Lo_Inicio"))
        .TextFrame.Characters.Text = "Tabla Resumen de Pagos por Transferencia o TPV/Bizum de Jornadas y Congresos, y sus Liquidaciones."
        .Fill.ForeColor.RGB = RGB(220, 210, 240)
    End With

    H_INICI.Visible = INI_Visibl
'    Application.ScreenUpdating = True
'    Application.Speech.Speak "Proceso completado puede verificar el resultado."
End Sub     '- Rut_Inicio_LstObj_Refresh
' ------------------------------------------------------------------------------------------------------

Sub kk()
    With ActiveSheet.Shapes.Range(Array("Tit_Lo_Inicio"))
        .TextFrame.Characters.Text = "Tabla Resumen de Pagos por Transferencia o TPV/Bizum de Jornadas y Congresos."
        .Fill.ForeColor.RGB = RGB(180, 200, 230)
    End With
    
'    With ActiveSheet.Shapes.Range(Array("Tit_Lo_Inicio"))
'        .TextFrame.Characters.Text = "Tabla Resumen de Pagos por Transferencia o TPV/Bizum de JyC y sus Liquidaciones."
'        .Fill.ForeColor.RGB = RGB(220, 230, 200)
'    End With

    With ActiveSheet.Shapes.Range(Array("Tit_Lo_Inicio"))
        .TextFrame.Characters.Text = "Tabla Resumen de Pagos por Transferencia o TPV/Bizum de Jornadas y Congresos, y sus Liquidaciones."
        .Fill.ForeColor.RGB = RGB(220, 210, 240)
    End With

End Sub
