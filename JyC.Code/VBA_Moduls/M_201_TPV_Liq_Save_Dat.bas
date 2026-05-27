Attribute VB_Name = "M_201_TPV_Liq_Save_Dat"
Option Explicit

' ==================================================================================================================================
' ==================================================================================================================================
' =====================     RuT_Grabar_Datos_TPV_Liq_a_TPV_(Tb)_y_Lst_(Tb)      ==============================================================
' ==================================================================================================================================
' ==================================================================================================================================
Sub RuT_Grabar_Datos_TPV_Liq_a_TPV_y_Lst()
    
    Dim F_TPV               As Long
    Dim F_Liq               As Long:        F_Liq = 1
    Dim RowLst          As Long
    Dim TotImpDI        As Currency
    Dim RowLiq          As ListRow
    Dim RowTPV          As ListRow
    
    Dim Lo_Liq          As ListObject:      Set Lo_Liq = H_Liq_TPV.ListObjects(1)
    Dim Lo_TPV          As ListObject:      Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Dim Lo_Lst          As ListObject:      Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    
    Call Rut_LstObj_Buscar(Lo_Lst, H_Liq_TPV.Range("Liq_Siglas"), 1, RowLst)
    If RowLst = 0 Then MsgBox "Siglas JyC, not find": Exit Sub
    
    Application.Calculation = xlCalculationManual:      Application.EnableEvents = False:   Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    With Application.Workbooks(ThisWorkbook.Name).Sheets(Prog_TPV_Tb.Name)
        .Protect allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                 UserInterfaceOnly:=True    '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  =======
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        Call Rut_LstObj_Filtros_Quitar(.ListObjects(1))
    End With
    Call RuT_DesHacer_Filtros
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_Ref, xlAscending, True)
    
    H_Liq_TPV.Select
    Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
    Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
    Call Rut_LstObj_Sort(Lo_Liq, 1, xlAscending, True)
    
    ' -------------------------------------------------------
    If H_Liq_TPV.Range("Liq_ExpAdm") <> "" Then Lo_Lst.DataBodyRange.Cells(RowLst, C_JCL_ExpAdm) = H_Liq_TPV.Range("Liq_ExpAdm")
    ' -------------------------------------------------------
    Set RowLiq = Lo_Liq.ListRows(1)
    For F_TPV = 1 To Lo_TPV.ListRows.Count ' recorro cada fila  de la tabla
    
        Set RowTPV = Lo_TPV.ListRows(F_TPV)
        '- Si coincide la Referencia del Recibo...
        If RowTPV.Range(C_TPV_Ref) = RowLiq.Range(C_TPV_Liq_Ref) Then
            
            RowTPV.Range(C_TPV_Ope) = Prog__APP.Range("APP_User_Name")
            
            RowTPV.Range(C_TPV_Obs) = RowLiq.Range(C_TPV_Liq_Obs)
            RowTPV.Range(C_TPV_Inscrito) = RowLiq.Range(C_TPV_Liq_Inscrito)
            RowTPV.Range(C_TPV_RDT_inv) = RowLiq.Range(C_TPV_Liq_N_RDT_inv)
            
            RowTPV.Range(C_TPV_DI_Imp) = RowLiq.Range(C_TPV_Liq_DI_Imp)
            TotImpDI = TotImpDI + RowLiq.Range(C_TPV_Liq_DI_Imp)
            
            RowTPV.Range(C_TPV_N_Liq) = H_Liq_TPV.Range("Liq_Núm")
            RowTPV.Range(C_TPV_Siglas) = H_Liq_TPV.Range("Liq_Siglas")
            RowTPV.Range(C_TPV_Ref_Sol) = H_Liq_TPV.Range("Liq_Solicitud")
            RowTPV.Range(C_TPV_Org) = H_Liq_TPV.Range("Liq_Orgánica")
            RowTPV.Range(C_TPV_JI) = H_Liq_TPV.Range("Liq_JI")
            RowTPV.Range(C_TPV_ExpAdm) = H_Liq_TPV.Range("Liq_ExpAdm")
            RowTPV.Range(C_TPV_RDT) = H_Liq_TPV.Range("Liq_RDT")
            RowTPV.Range(C_TPV_PMP) = H_Liq_TPV.Range("Liq_PMP")
            
'''                RowTPV.Range(C_TPV_DI) = H_Liq_TPV.Range("Liq_DI")
            
            'If H_Liq_TPV.Range("Liq_DI") <> "" Then
            If RowLiq.Range(C_TPV_Liq_DI_Imp) > 0 Then
                RowTPV.Range(C_TPV_DI) = H_Liq_TPV.Range("Liq_DI")
                RowTPV.Range(C_TPV_DI_Imp) = RowLiq.Range(C_TPV_Liq_DI_Imp)
'                H_Liq_TPV.Range("Liq_RDT_DI_Imp").NumberFormat = """Importe Devolución: "" #,##0.00"" euros"""
            Else
                RowTPV.Range(C_TPV_DI).ClearContents
                RowTPV.Range(C_TPV_DI_Imp).ClearContents
'                H_Liq_TPV.Range("Liq_RDT_DI_Imp").NumberFormat = """Importe RDT: "" #,##0.00"" euros"""
                Range("Liq_RDT_DI_Imp") = Lo_Liq.TotalsRowRange.Columns(C_TPV_Liq_Imp)
'                H_Liq_TPV.Range("Liq_RDT_DI_Imp").NumberFormat = """Importe RDT: "" #,##0.00"" euros"""
            End If
            
            If F_Liq = Lo_Liq.ListRows.Count Then Exit For
            F_Liq = F_Liq + 1
            Set RowLiq = Lo_Liq.ListRows(F_Liq)

        ElseIf RowTPV.Range(C_TPV_Ref) > RowLiq.Range(C_TPV_Liq_Ref) Then
        
            F_TPV = F_TPV - 1
            
            If F_Liq = Lo_Liq.ListRows.Count Then Exit For
            F_Liq = F_Liq + 1
            Set RowLiq = Lo_Liq.ListRows(F_Liq)
            
        End If
    
    Next F_TPV

End Sub     ' RuT_Grabar_Datos_TPV_Liq_a_TPV_y_Lst     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================





