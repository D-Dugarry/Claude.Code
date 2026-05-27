Attribute VB_Name = "M_200_TPV_Liq_Load"
'M_200_Liquid_TPV
Option Explicit

' ==================================================================================================================================
' ==================================================================================================================================
' =====================     RuT_Generar_Tabla_Liquidación_TPV      ==============================================================
' ==================================================================================================================================
' ==================================================================================================================================
Sub RuT_Generar_Tabla_Liquidación_TPV()
    Dim RHeight     As Integer
    Dim FindNumLiq  As String
    Dim F_TPV       As Long
    Dim Lin_Lst     As Long
    Dim Imp_DI      As Double
    Dim Imp_RDT     As Currency
    Dim Fecha_Ini   As Date
    Dim RowLiq      As ListRow
    Dim RowTPV      As ListRow

    Application.Calculation = xlCalculationManual:      Application.EnableEvents = False:   Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Set Lo_Lst = Prog_JyC_List.ListObjects(1)
        Call Rut_LstObj_Buscar(Lo_Lst, H_Liq_TPV.Range("Liq_Siglas"), 1, Lin_Lst)
        If Lin_Lst = 0 Then MsgBox "Siglas JyC, not find": Exit Sub
    Set Lo_Liq = H_Liq_TPV.ListObjects(1)
    Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Prog_TPV_Tb.Protect allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
    Call Rut_LstObj_WrkSht_Preparar(Application.Workbooks(ThisWorkbook.Name).Sheets(Prog_TPV_Tb.Name))
    Call Rut_LstObj_Sort(Lo_TPV, 1, xlAscending, True)
    
    H_Liq_TPV.Select
    Call Rut_LstObj_WrkSht_Preparar(Application.Workbooks(ThisWorkbook.Name).Sheets(H_Liq_TPV.Name))
    Call Rut_LstObj_Columns_Ajustar_Ancho(H_Liq_TPV, Prog_DefColTPVLiq)
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
        
    Fecha_Ini = "01/01/" & Format(Now(), "yyyy")
    H_Liq_TPV.Range("Liq_Nombre") = Lo_Lst.DataBodyRange.Cells(Lin_Lst, C_JCL_Nom)
    H_Liq_TPV.Range("Liq_Email") = Lo_Lst.DataBodyRange.Cells(Lin_Lst, C_JCL_Email)
    H_Liq_TPV.Range("Liq_Orgánica") = Lo_Lst.DataBodyRange.Cells(Lin_Lst, C_JCL_Orgánica)
    H_Liq_TPV.Range("Liq_ExpAdm") = Lo_Lst.DataBodyRange.Cells(Lin_Lst, C_JCL_ExpAdm)
    
'    If Mid(H_Liq_TPV.Range("Liq_Orgánica"), 7, 1) = 6 Then
'        H_Liq_TPV.Range("Liq_Económica") = "1345.01.01"
'    ElseIf Mid(H_Liq_TPV.Range("Liq_Orgánica"), 7, 1) = 2 Then
'        H_Liq_TPV.Range("Liq_Económica") = "1345.01.00"
'    Else
'        H_Liq_TPV.Range("Liq_Económica") = "¡¡ No es cap.2 ó 6 !!"
'    End If
    With H_Liq_TPV
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
    End With    '- H_Liq_TPV
        
    FindNumLiq = H_Liq_TPV.Range("Liq_Núm")
    
With Lo_TPV.DataBodyRange
    ' -------------------------------------------------------
    For F_TPV = 1 To Lo_TPV.ListRows.Count ' recorro cada fila  de la tabla
    
        Set RowTPV = Lo_TPV.ListRows(F_TPV)
    
        ' --------------------------------=============  Filtrar Fech_Valor ==================
        If RowTPV.Range(C_TPV_F_Pago) < Fecha_Ini Then GoTo Siguiente_Fila
        ' --------------------------------=============  Filtrar H_Liq_CTA.Range("Liq_Siglas") ==================
        If RowTPV.Range(C_TPV_Siglas) <> H_Liq_TPV.Range("Liq_Siglas") Then GoTo Siguiente_Fila
        ' --------------------------------=============  Filtrar H_Liq_CTA.Range("Liq_Núm") ==================
        If Len(RowTPV.Range(C_TPV_N_Liq)) = 0 Then GoTo Siguiente_Fila
        ' --------------------------------=============  Filtrar H_Liq_CTA.Range("Liq_Núm") ==================
        If InStr(1, RowTPV.Range(C_TPV_N_Liq).Value, FindNumLiq) = 0 Then GoTo Siguiente_Fila
        ' --------------------------------=============  Filtrar Importe Negativos ==================
'        If RowTPV.Range(C_TPV_Neto) < 0 And RowTPV.Range(C_Cta_DI) = "" Then
'                GoTo Siguiente_Fila
'        End If
        
        ' #############################==================  Tratamos los datos ==================###############################
        Set RowLiq = Lo_Liq.ListRows.Add
        
        RowLiq.Range(C_TPV_Liq_Ref) = RowTPV.Range(C_TPV_Ref)
        RowLiq.Range(C_TPV_Liq_F_Pag) = RowTPV.Range(C_TPV_F_Pago)
        RowLiq.Range(C_TPV_Liq_Form_Pag) = RowTPV.Range(C_TPV_Tipo_Pago)
        RowLiq.Range(C_TPV_Liq_Ordenante) = RowTPV.Range(C_TPV_NomApe)
        RowLiq.Range(C_TPV_Liq_Inscrito) = RowTPV.Range(C_TPV_Inscrito)
        RowLiq.Range(C_TPV_Liq_Imp) = RowTPV.Range(C_TPV_Imp)
            Imp_RDT = Imp_RDT + Val(RowTPV.Range(C_TPV_Imp))
        
        RowLiq.Range(C_TPV_Liq_ComBco) = RowTPV.Range(C_TPV_ComBco)
        RowLiq.Range(C_TPV_Liq_Neto) = RowTPV.Range(C_TPV_Neto)
        
        RowLiq.Range(C_TPV_Liq_Imp) = RowTPV.Range(C_TPV_Neto) + RowTPV.Range(C_TPV_ComBco)
        
        RowLiq.Range(C_TPV_Liq_N_Liq) = RowTPV.Range(C_TPV_N_Liq)
        H_Liq_TPV.Range("Liq_Núm") = RowTPV.Range(C_TPV_N_Liq)
        RowLiq.Range(C_TPV_Liq_Siglas) = RowTPV.Range(C_TPV_Siglas)
        RowLiq.Range(C_Cta_Liq_Ref_Sol) = RowTPV.Range(C_TPV_Ref_Sol)
        RowLiq.Range(C_TPV_Liq_Obs) = RowTPV.Range(C_TPV_Obs)
        
        RowLiq.Range(C_TPV_Liq_Org) = RowTPV.Range(C_TPV_Org)
        RowLiq.Range(C_TPV_Liq_JI) = RowTPV.Range(C_TPV_JI)
        RowLiq.Range(C_TPV_Liq_ExpAdm) = RowTPV.Range(C_TPV_ExpAdm)
        RowLiq.Range(C_TPV_Liq_N_RDT) = RowTPV.Range(C_TPV_RDT)
        RowLiq.Range(C_TPV_Liq_N_RDT_inv) = RowTPV.Range(C_TPV_RDT_inv)
        RowLiq.Range(C_TPV_Liq_DI) = RowTPV.Range(C_TPV_DI)
        If RowTPV.Range(C_TPV_DI_Imp) = "" Then
'            RowLiq.Range(C_TPV_Liq_DI_Imp) = RowTPV.Range(C_TPV_Imp)
'            Imp_DI = Imp_DI + RowTPV.Range(C_TPV_Imp)
        Else
            RowLiq.Range(C_TPV_Liq_DI_Imp) = RowTPV.Range(C_TPV_DI_Imp)
            Imp_DI = Imp_DI + RowTPV.Range(C_TPV_DI_Imp)
        End If
        
        RowLiq.Range(C_TPV_Liq_PMP) = RowTPV.Range(C_TPV_PMP)
        RowLiq.Range(C_TPV_Liq_Ope) = RowTPV.Range(C_TPV_Ope)
        
'''        H_Liq_TPV.Range("Liq_Núm") = RowTPV.Range(C_TPV_N_Liq)
'''        H_Liq_TPV.Range("Liq_ExpAdm") = RowTPV.Range(C_TPV_ExpAdm)
        '- Añado Referencias si no existen ya en la celda...
        If Len(RowTPV.Range(C_TPV_Ref_Sol)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_Solicitud"), RowTPV.Range(C_TPV_Ref_Sol)) = 0 Then _
                        H_Liq_TPV.Range("Liq_Solicitud") = H_Liq_TPV.Range("Liq_Solicitud") & " - " & RowTPV.Range(C_TPV_Ref_Sol)
        If Len(RowTPV.Range(C_TPV_JI)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_JI"), RowTPV.Range(C_TPV_JI).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_JI") = H_Liq_TPV.Range("Liq_JI") & " - " & RowTPV.Range(C_TPV_JI)
        If Len(RowTPV.Range(C_TPV_DI)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_DI"), RowTPV.Range(C_TPV_DI).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_DI") = H_Liq_TPV.Range("Liq_DI") & " - " & RowTPV.Range(C_TPV_DI)
        If Len(RowTPV.Range(C_TPV_PMP)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_PMP"), RowTPV.Range(C_TPV_PMP).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_PMP") = H_Liq_TPV.Range("Liq_PMP") & " - " & RowTPV.Range(C_TPV_PMP)
        If Len(RowTPV.Range(C_TPV_RDT)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_RDT"), RowTPV.Range(C_TPV_RDT).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_RDT") = H_Liq_TPV.Range("Liq_RDT") & " - " & RowTPV.Range(C_TPV_RDT)
        '- Ajusto el alto de la fila añadida...
        RowLiq.Range.EntireRow.AutoFit
        
Siguiente_Fila:
    Next F_TPV
    
    '- Quitar el último separador " - " si existe...
    If Left(H_Liq_TPV.Range("Liq_Solicitud"), 3) = " - " Then H_Liq_TPV.Range("Liq_Solicitud") = Mid(H_Liq_TPV.Range("Liq_Solicitud"), 4)
    If Left(H_Liq_TPV.Range("Liq_JI"), 3) = " - " Then H_Liq_TPV.Range("Liq_JI") = Mid(H_Liq_TPV.Range("Liq_JI"), 4)
    If Left(H_Liq_TPV.Range("Liq_DI"), 3) = " - " Then H_Liq_TPV.Range("Liq_DI") = Mid(H_Liq_TPV.Range("Liq_DI"), 4)
    If Left(H_Liq_TPV.Range("Liq_PMP"), 3) = " - " Then H_Liq_TPV.Range("Liq_PMP") = Mid(H_Liq_TPV.Range("Liq_PMP"), 4)
    If Left(H_Liq_TPV.Range("Liq_RDT"), 3) = " - " Then H_Liq_TPV.Range("Liq_RDT") = Mid(H_Liq_TPV.Range("Liq_RDT"), 4)
    
End With    ' Lo_TPV.DataBodyRange

'    Call Rut_LstObj_Sort(Prog_TPV_Tb.ListObjects(1), 1, xlAscending, True)
    Lo_Liq.ShowTotals = True
    Call Rut_LstObj_Sort(Lo_Liq, 1, xlAscending, True)

    Range("Liq_Imp") = Lo_Liq.TotalsRowRange.Columns(C_TPV_Liq_Imp)
    Range("f4") = Lo_Liq.TotalsRowRange.Columns(C_TPV_Liq_Ref)
    
'    If Range("Liq_DI") = "" Then
            Rows("8:9").EntireRow.Hidden = True
            Rows(10).RowHeight = 40
            Range("f3").NumberFormat = """Importe RDT: "" #,##0.00"" euros"""
            Range("f3") = Lo_Liq.TotalsRowRange.Columns(C_TPV_Liq_Imp)
            Range("Liq_Nota").Interior.Color = -1
            Range("Liq_Nota") = ""
            
            ' ??¿¿ no sé si falta algo
            
            Rows("1:7").EntireRow.AutoFit
'    Else
    If Imp_DI > 0 Then  '- !! Hay Devolución !! -----------------
        If Imp_RDT = Imp_DI Then        '- !! Hay Devolución TOTAL !! -----------------
            Rows("8:9").EntireRow.Hidden = False
            Rows(10).RowHeight = 10
            Range("Liq_RDT_DI_Imp").NumberFormat = """Importe Devolución: "" #,##0.00"" euros"""
            Range("Liq_RDT_DI_Imp") = Imp_DI
            Range("Liq_Nota").Interior.ColorIndex = 6
            Range("Liq_Nota").Value = Lo_Liq.DataBodyRange.Cells(1, C_TPV_Liq_Obs)
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
            Range("Liq_RDT_DI_Imp") = "Importe RDT: " & Format(Imp_RDT - Imp_DI, "#,##0.00") & " €, y " & vbLf & _
                                      "Importe Devolución: " & Format(Imp_DI, "#,##0.00") & " €"
            Rows("1:9").EntireRow.AutoFit
            Range("Liq_RDT_DI_Imp").RowHeight = 40
            Range("Liq_Nota").Interior.ColorIndex = 6
            Range("Liq_Nota").Value = Lo_Liq.DataBodyRange.Cells(1, C_TPV_Liq_Obs)
            Range("Liq_Nota").Offset(0, 1) = Range("Liq_Nota")
            RHeight = Range("Liq_Nota").RowHeight
            RHeight = RHeight / 3 + 3
            If RHeight < 18 Then RHeight = 18
            Rows("5:7").RowHeight = RHeight
            Range("Liq_Nota").Offset(0, 1).ClearContents
'            Rows("1:9").EntireRow.AutoFit
        End If
    End If
    If Rows(1).RowHeight < 40 Then Rows(1).RowHeight = 40
    
Restablecer_Valores:
    Application.Calculation = xlCalculationAutomatic: Application.EnableEvents = True: Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Lo_Liq.Range.Cells(1, 1).Select
    ActiveCell.Offset(1, 0).Select
End Sub     ' RuT_Generar_Tabla_Liquidación_TPV     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================


