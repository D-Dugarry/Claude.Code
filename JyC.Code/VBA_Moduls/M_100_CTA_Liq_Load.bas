Attribute VB_Name = "M_100_CTA_Liq_Load"
'M_100_CTA_Liquid
Option Explicit

' ==================================================================================================================================
' ==================================================================================================================================
' =====================     RuT_Generar_Tabla_Liquidación_Cta      ==============================================================
' ==================================================================================================================================
' ==================================================================================================================================
Sub RuT_Generar_Tabla_Liquidación_Cta()
    Dim RHeight     As Integer
    Dim SW_BcoDiff  As Boolean:     SW_BcoDiff = False
    Dim F_Cta       As Long
    Dim Lin_Lst     As Long
    Dim Imp_DI      As Currency
    Dim Imp_RDT     As Currency
    Dim Fecha_Ini   As Date
    Dim RowLiq      As ListRow
    Dim RowCta      As ListRow

    Application.Calculation = xlCalculationManual:      Application.EnableEvents = False:   Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Set Lo_Lst = Prog_JyC_List.ListObjects(1)
        Call Rut_LstObj_Buscar(Lo_Lst, H_Liq_CTA.Range("Liq_Siglas"), 1, Lin_Lst)
        If Lin_Lst = 0 Then MsgBox "Siglas JyC, not find": Exit Sub
    Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    Set Lo_Liq = H_Liq_CTA.ListObjects(1)
    Prog_CTA_Tb.Protect allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
    Call Rut_LstObj_WrkSht_Preparar(Application.Workbooks(ThisWorkbook.Name).Sheets(Prog_CTA_Tb.Name))
    Call Rut_LstObj_Sort(Lo_Cta, 1, xlAscending, True)
    
    H_Liq_CTA.Select
    Call Rut_LstObj_WrkSht_Preparar(Application.Workbooks(ThisWorkbook.Name).Sheets(H_Liq_CTA.Name))
    Call Rut_LstObj_Columns_Ajustar_Ancho(H_Liq_CTA, Prog_DefColCtaLiq)
    Rows(1).RowHeight = 40
    Rows("2:9").RowHeight = 20
    Lo_Liq.ShowTotals = False
    If Not Lo_Liq.DataBodyRange Is Nothing Then Lo_Liq.DataBodyRange.Delete
    
        Range("b1") = Now()
        Range("b2").ClearContents
        Range("e3").ClearContents
        Range("e5").ClearContents
        Range("e6").ClearContents
        Range("e7").ClearContents
        Range("e8").ClearContents
        Range("e9").ClearContents
    
'    Fecha_Ini = "01/01/" & Format(Now(), "yyyy")
    Fecha_Ini = "01/01/" & Prog__APP.Range("APP_AñoCont")
    H_Liq_CTA.Range("Liq_Nombre") = Lo_Lst.DataBodyRange.Cells(Lin_Lst, C_JCL_Nom)
    H_Liq_CTA.Range("Liq_Email") = Lo_Lst.DataBodyRange.Cells(Lin_Lst, C_JCL_Email)
    H_Liq_CTA.Range("Liq_Orgánica") = Lo_Lst.DataBodyRange.Cells(Lin_Lst, C_JCL_Orgánica)
    H_Liq_CTA.Range("Liq_ExpAdm") = Lo_Lst.DataBodyRange.Cells(Lin_Lst, C_JCL_ExpAdm)
    
    With H_Liq_CTA
        If Mid(.Range("Liq_Orgánica"), 7, 1) = 6 Then
            .Range("Liq_Económica") = "1345.01.01"
            .Range("Liq_Económica").Interior.Color = RGB(220, 230, 241)
            .Range("Liq_Económica").Font.ColorIndex = 3    'amarillo
            .Range("Liq_Orgánica").Offset(0, -1) = "¡ La Orgánica es de Investigación !"
            .Range("Liq_Orgánica").Offset(0, -1).Interior.ColorIndex = 2
            .Range("Liq_Orgánica").Offset(0, -1).Font.ColorIndex = 1     'negro
        ElseIf Mid(.Range("Liq_Orgánica"), 7, 1) = 2 Then
            .Range("Liq_Económica") = "1311.03"
            .Range("Liq_Económica").Interior.Color = 192
            .Range("Liq_Económica").Font.ColorIndex = 6    'amarillo
            .Range("Liq_Orgánica").Offset(0, -1) = "¡ La Orgánica NO es de Investigación !"
            .Range("Liq_Orgánica").Offset(0, -1).Interior.Color = 192
            .Range("Liq_Orgánica").Offset(0, -1).Font.ColorIndex = 6     'amarillo
        Else
            .Range("Liq_Económica") = "¡¡ Ojo !!"
            .Range("Liq_Económica").Interior.Color = 1
            .Range("Liq_Económica").Font.ColorIndex = 6     'amarillo
            .Range("Liq_Orgánica").Offset(0, -1) = "¡¡ La Orgánica No es cap.2 ó 6 !!"
            .Range("Liq_Orgánica").Offset(0, -1).Interior.Color = 1
            .Range("Liq_Orgánica").Offset(0, -1).Font.ColorIndex = 6     'amarillo
        End If
    End With    '- H_Liq_CTA
    
With Lo_Cta.DataBodyRange
    ' -------------------------------------------------------
    For F_Cta = 1 To Lo_Cta.ListRows.Count ' recorro cada fila  de la tabla
    
        Set RowCta = Lo_Cta.ListRows(F_Cta)
    
        ' --------------------------------=============  Filtrar Fech_Valor ==================
        If RowCta.Range(C_Cta_F_VAL) < Fecha_Ini Then GoTo Siguiente_Fila
        ' --------------------------------=============  Filtrar H_Liq_CTA.Range("Liq_Siglas") ==================
        If RowCta.Range(C_Cta_Siglas) <> H_Liq_CTA.Range("Liq_Siglas") Then GoTo Siguiente_Fila
        ' --------------------------------=============  Filtrar H_Liq_CTA.Range("Liq_Núm") ==================
        If Len(RowCta.Range(C_Cta_N_Liq)) = 0 Then GoTo Siguiente_Fila
        ' --------------------------------=============  Filtrar H_Liq_CTA.Range("Liq_Núm") ==================
        If InStr(RowCta.Range(C_Cta_N_Liq), H_Liq_CTA.Range("Liq_Núm")) = 0 Then GoTo Siguiente_Fila
        ' --------------------------------=============  Filtrar Importe Negativos ==================
        If RowCta.Range(C_Cta_Imp) < 0 And RowCta.Range(C_Cta_DI) = "" Then
                GoTo Siguiente_Fila
        End If
        
        ' #############################==================  Tratamos los datos ==================###############################
        Set RowLiq = Lo_Liq.ListRows.Add
        
        RowLiq.Range(C_Cta_Liq_Ordinal) = RowCta.Range(C_Cta_Ordinal)
        RowLiq.Range(C_Cta_Liq_F_OPE) = RowCta.Range(C_Cta_F_OPE)
        RowLiq.Range(C_Cta_Liq_Imp) = RowCta.Range(C_Cta_Imp)
            Imp_RDT = Imp_RDT + Val(RowCta.Range(C_Cta_Imp))
            
        RowLiq.Range(C_Cta_Liq_Reg_Mov1) = RowCta.Range(C_Cta_Reg_Mov1)
        RowLiq.Range(C_Cta_Liq_Reg_Mov3) = RowCta.Range(C_Cta_Reg_Mov3)
        RowLiq.Range(C_Cta_Liq_Reg_Mov4) = RowCta.Range(C_Cta_Reg_Mov4)
        RowLiq.Range(C_Cta_Liq_Inscrito) = RowCta.Range(C_Cta_Inscrito)
        RowLiq.Range(C_Cta_Liq_Obs) = RowCta.Range(C_Cta_Obs)
        
        RowLiq.Range(C_Cta_Liq_N_Liq) = RowCta.Range(C_Cta_N_Liq)
'        H_Liq_CTA.Range("Liq_Núm") = RowCta.Range(C_Cta_N_Liq)
        RowLiq.Range(C_Cta_Liq_Siglas) = RowCta.Range(C_Cta_Siglas)
        RowLiq.Range(C_Cta_Liq_Ref_Sol) = RowCta.Range(C_Cta_Ref_Sol)
        RowLiq.Range(C_Cta_Liq_Ope) = RowCta.Range(C_Cta_Ope)
        
        RowLiq.Range(C_Cta_Liq_Org) = RowCta.Range(C_Cta_Org)
        RowLiq.Range(C_Cta_Liq_JI) = RowCta.Range(C_Cta_JI)
        RowLiq.Range(C_Cta_Liq_ExpAdm) = RowCta.Range(C_Cta_ExpAdm)
        RowLiq.Range(C_Cta_Liq_N_RDT) = RowCta.Range(C_Cta_RDT)
        RowLiq.Range(C_Cta_Liq_DI) = RowCta.Range(C_Cta_DI)
        RowLiq.Range(C_Cta_Liq_PMP) = RowCta.Range(C_Cta_PMP)
        
        RowLiq.Range(C_Cta_Liq_N_OT) = RowCta.Range(C_Cta_N_OT)
        
'        If RowCta.Range(C_Cta_DI_Imp) = "" Then
'            'RowLiq.Range(C_Cta_Liq_DI_Imp) = RowCta.Range(C_Cta_Imp)
'        Else
'            RowLiq.Range(C_Cta_Liq_DI_Imp) = RowCta.Range(C_Cta_DI_Imp)
'        End If  Imp_RDT

            RowLiq.Range(C_Cta_Liq_DI_Imp) = RowCta.Range(C_Cta_DI_Imp)
            Imp_DI = Imp_DI + Val(RowCta.Range(C_Cta_DI_Imp))
        
        If RowCta.Range(C_Cta_Bco) = "xxx" Then SW_BcoDiff = True

'''        H_Liq_CTA.Range("Liq_Núm") = RowCta.Range(C_Cta_N_Liq)
'''        H_Liq_CTA.Range("Liq_ExpAdm") = RowCta.Range(C_Cta_ExpAdm)
        '- Añado Referencias si no existen ya en la celda...
        If Len(RowCta.Range(C_Cta_Ref_Sol)) > 0 And InStr(1, H_Liq_CTA.Range("Liq_Solicitud"), RowCta.Range(C_Cta_Ref_Sol)) = 0 Then _
                        H_Liq_CTA.Range("Liq_Solicitud") = H_Liq_CTA.Range("Liq_Solicitud") & " - " & RowCta.Range(C_Cta_Ref_Sol)
        If Len(RowCta.Range(C_Cta_JI)) > 0 And InStr(1, H_Liq_CTA.Range("Liq_JI"), RowCta.Range(C_Cta_JI).Value) = 0 Then _
                        H_Liq_CTA.Range("Liq_JI") = H_Liq_CTA.Range("Liq_JI") & " - " & RowCta.Range(C_Cta_JI)
        If Len(RowCta.Range(C_Cta_DI)) > 0 And InStr(1, H_Liq_CTA.Range("Liq_DI"), RowCta.Range(C_Cta_DI).Value) = 0 Then _
                        H_Liq_CTA.Range("Liq_DI") = H_Liq_CTA.Range("Liq_DI") & " - " & RowCta.Range(C_Cta_DI)
        If Len(RowCta.Range(C_Cta_PMP)) > 0 And InStr(1, H_Liq_CTA.Range("Liq_PMP"), RowCta.Range(C_Cta_PMP).Value) = 0 Then _
                        H_Liq_CTA.Range("Liq_PMP") = H_Liq_CTA.Range("Liq_PMP") & " - " & RowCta.Range(C_Cta_PMP)
        If Len(RowCta.Range(C_Cta_RDT)) > 0 And InStr(1, H_Liq_CTA.Range("Liq_RDT"), RowCta.Range(C_Cta_RDT).Value) = 0 Then _
                        H_Liq_CTA.Range("Liq_RDT") = H_Liq_CTA.Range("Liq_RDT") & " - " & RowCta.Range(C_Cta_RDT)
        '- Ajusto el alto de la fila añadida...
        RowLiq.Range.EntireRow.AutoFit
        
Siguiente_Fila:
    Next F_Cta
    
    '- Quitar el último separador " - " si existe...
    If Left(H_Liq_CTA.Range("Liq_Solicitud"), 3) = " - " Then H_Liq_CTA.Range("Liq_Solicitud") = Mid(H_Liq_CTA.Range("Liq_Solicitud"), 4)
    If Left(H_Liq_CTA.Range("Liq_JI"), 3) = " - " Then H_Liq_CTA.Range("Liq_JI") = Mid(H_Liq_CTA.Range("Liq_JI"), 4)
    If Left(H_Liq_CTA.Range("Liq_DI"), 3) = " - " Then H_Liq_CTA.Range("Liq_DI") = Mid(H_Liq_CTA.Range("Liq_DI"), 4)
    If Left(H_Liq_CTA.Range("Liq_PMP"), 3) = " - " Then H_Liq_CTA.Range("Liq_PMP") = Mid(H_Liq_CTA.Range("Liq_PMP"), 4)
    If Left(H_Liq_CTA.Range("Liq_RDT"), 3) = " - " Then H_Liq_CTA.Range("Liq_RDT") = Mid(H_Liq_CTA.Range("Liq_RDT"), 4)
    
End With    ' Lo_CTA.DataBodyRange

    Lo_Liq.ShowTotals = True
    Call Rut_LstObj_Sort(Lo_Liq, 1, xlAscending, True)

    Range("Liq_Imp") = Lo_Liq.TotalsRowRange.Columns(C_Cta_Liq_Imp)
    Range("f4") = Lo_Liq.TotalsRowRange.Columns(C_Cta_Liq_Ordinal)
    
'    If Range("Liq_DI") = "" Then
            Rows("8:9").EntireRow.Hidden = True
            Rows(10).RowHeight = 40
            Range("Liq_RDT_DI_Imp").NumberFormat = """Importe RDT: "" #,##0.00"" euros"""
            Range("Liq_RDT_DI_Imp") = Lo_Liq.TotalsRowRange.Columns(C_Cta_Liq_Imp)
            Range("Liq_Nota").Interior.Color = -1
            Range("Liq_Nota") = ""
            If SW_BcoDiff Then
                Rows(10).RowHeight = 10
                Range("Liq_Nota").Interior.ColorIndex = 6
                Range("Liq_Nota").Value = RowLiq.Range(C_Cta_Liq_Obs)
                Range("Liq_Nota").Offset(0, 1) = Range("Liq_Nota")
                RHeight = Range("Liq_Nota").RowHeight
                RHeight = RHeight / 3 + 3
                If RHeight < 18 Then RHeight = 18
                Rows("5:7").RowHeight = RHeight
                Range("Liq_Nota").Offset(0, 1).ClearContents
            End If
                Rows("1:7").EntireRow.AutoFit
'    Else
    If Imp_DI > 0 Then  '- !! Hay Devolución !! -----------------
        If Imp_RDT = Imp_DI Then        '- !! Hay Devolución TOTAL !! -----------------
            Rows("8:9").EntireRow.Hidden = False
            Rows(10).RowHeight = 10
            Range("Liq_RDT_DI_Imp").NumberFormat = """Importe Devolución: "" #,##0.00"" euros"""
            Range("Liq_RDT_DI_Imp") = Imp_DI
            Range("Liq_Nota").Interior.ColorIndex = 6
            Range("Liq_Nota").Value = Lo_Liq.DataBodyRange.Cells(1, C_Cta_Liq_Obs)
            Range("Liq_Nota").Offset(0, 1) = Range("Liq_Nota")
            Rows("1:9").EntireRow.AutoFit
            RHeight = Range("Liq_Nota").RowHeight
            RHeight = RHeight / 3 + 3
            If RHeight < 18 Then RHeight = 18
            Rows("5:7").RowHeight = RHeight
            Range("Liq_Nota").Offset(0, 1).ClearContents
            Rows("1:9").EntireRow.AutoFit
        Else                            '- !! Hay Devolución Parcial !! -----------------
            Rows("8:9").EntireRow.Hidden = False
            Rows(10).RowHeight = 10
            Range("Liq_RDT_DI_Imp") = "Importe RDT: " & Format(Imp_RDT - Imp_DI, "#,##0.00") & " €,  " & _
                                      "Importe Devolución: " & Format(Imp_DI, "#,##0.00") & " €"
            Range("Liq_Nota").Interior.ColorIndex = 6
            Range("Liq_Nota").Value = Lo_Liq.DataBodyRange.Cells(1, C_Cta_Liq_Obs)
            Range("Liq_Nota").Offset(0, 1) = Range("Liq_Nota")
            Rows("1:9").EntireRow.AutoFit
            RHeight = Range("Liq_Nota").RowHeight
            RHeight = RHeight / 3 + 3
            If RHeight < 18 Then RHeight = 18
            Rows("5:7").RowHeight = RHeight
            Range("Liq_Nota").Offset(0, 1).ClearContents
            Rows("1:9").EntireRow.AutoFit
        End If
    End If
    
    If Rows(1).RowHeight < 40 Then Rows(1).RowHeight = 40
 
Restablecer_Valores:
    Application.Calculation = xlCalculationAutomatic: Application.EnableEvents = True: Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Lo_Liq.Range.Cells(1, 1).Select
    ActiveCell.Offset(1, 0).Select
End Sub     ' RuT_Generar_Tabla_Liquidación_Cta     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================


