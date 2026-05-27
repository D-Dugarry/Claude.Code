Attribute VB_Name = "M_101_CTA_Liq_Dat_Save"
'M_101_CTA_Liq_Dat_Save
Option Explicit

' ==================================================================================================================================
' ==================================================================================================================================
' =====================     RuT_Grabar_Datos_CTA_Liq_a_CTA_y_Lst      ==============================================================
' ==================================================================================================================================
' ==================================================================================================================================
Sub RuT_Grabar_Datos_CTA_Liq_a_CTA_y_Lst()
    
    Dim F_Cta           As Long
    Dim F_Liq           As Long:        F_Liq = 1
    Dim RowLst          As Long
    Dim Imp_DI          As Currency
    Dim Imp_RDT         As Currency
    Dim RowLiq          As ListRow
    Dim RowCta          As ListRow
    
    Dim Lo_Cta          As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    Dim Lo_Liq          As ListObject:      Set Lo_Liq = H_Liq_CTA.ListObjects(1)
    Dim Lo_Lst          As ListObject:      Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    
    Call Rut_LstObj_Buscar(Lo_Lst, H_Liq_CTA.Range("Liq_Siglas"), 1, RowLst)
    If RowLst = 0 Then MsgBox "Siglas JyC, not find": Exit Sub
    
    Application.Calculation = xlCalculationManual:      Application.EnableEvents = False:   Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    With Application.Workbooks(ThisWorkbook.Name).Sheets(Prog_CTA_Tb.Name)
        .Protect allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                 UserInterfaceOnly:=True    '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  =======
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        Call Rut_LstObj_Filtros_Quitar(.ListObjects(1))
    End With
    Call RuT_DesHacer_Filtros
    Call Rut_LstObj_Sort(Lo_Cta, 1, xlAscending, True)
    
    H_Liq_CTA.Select
    Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
    Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
    Call Rut_LstObj_Sort(Lo_Liq, C_Cta_Liq_Ordinal, xlAscending, True)
    
    ' -------------------------------------------------------
    Lo_Lst.DataBodyRange.Cells(RowLst, C_JCL_ExpAdm) = H_Liq_CTA.Range("Liq_ExpAdm")
    ' -------------------------------------------------------
    Set RowLiq = Lo_Liq.ListRows(1)
    For F_Cta = 1 To Lo_Cta.ListRows.Count ' recorro cada fila  de la tabla
    
        Set RowCta = Lo_Cta.ListRows(F_Cta)
        
        If RowCta.Range(C_Cta_Ordinal) = RowLiq.Range(C_Cta_Liq_Ordinal) Then
            
            RowCta.Range(C_Cta_Ope) = Prog__APP.Range("APP_User_Name")
            RowCta.Range(C_Cta_Ref_Sol) = H_Liq_CTA.Range("Liq_Solicitud")
            RowCta.Range(C_Cta_Inscrito) = RowLiq.Range(C_Cta_Liq_Inscrito)
            RowCta.Range(C_Cta_Obs) = RowLiq.Range(C_Cta_Liq_Obs)
            RowCta.Range(C_Cta_Org) = H_Liq_CTA.Range("Liq_Orgánica")
            RowCta.Range(C_Cta_JI) = H_Liq_CTA.Range("Liq_JI")
            RowCta.Range(C_Cta_ExpAdm) = H_Liq_CTA.Range("Liq_ExpAdm")
            RowCta.Range(C_Cta_RDT) = H_Liq_CTA.Range("Liq_RDT")
            RowCta.Range(C_Cta_PMP) = H_Liq_CTA.Range("Liq_PMP")
            
                RowCta.Range(C_Cta_DI) = H_Liq_CTA.Range("Liq_DI")
                RowCta.Range(C_Cta_DI_Imp) = RowLiq.Range(C_Cta_Liq_DI_Imp)
'                Imp_DI = Imp_DI + RowLiq.Range(C_Cta_Liq_DI_Imp)
            
'''            If H_Liq_CTA.Range("Liq_DI") <> "" Then
''''                RowCta.Range(C_Cta_DI) = H_Liq_CTA.Range("Liq_DI")
''''                RowCta.Range(C_Cta_DI_Imp) = H_Liq_CTA.Range("Liq_RDT_DI_Imp")
'''                H_Liq_CTA.Range("Liq_RDT_DI_Imp").NumberFormat = """Importe Devolución: "" #,##0.00"" euros"""
'''            Else
''''                RowCta.Range(C_Cta_DI).ClearContents
''''                RowCta.Range(C_Cta_DI_Imp).ClearContents
'''                H_Liq_CTA.Range("Liq_RDT_DI_Imp").NumberFormat = """Importe RDT: "" #,##0.00"" euros"""
'''                Range("Liq_RDT_DI_Imp") = Lo_Liq.TotalsRowRange.Columns(C_Cta_Liq_Imp)
'''            End If
            
            If F_Liq = Lo_Liq.ListRows.Count Then Exit For
            F_Liq = F_Liq + 1
            Set RowLiq = Lo_Liq.ListRows(F_Liq)

        ElseIf RowCta.Range(C_Cta_Ordinal) > RowLiq.Range(C_Cta_Liq_Ordinal) Then
        
            F_Cta = F_Cta - 1
            
            If F_Liq = Lo_Liq.ListRows.Count Then Exit For
            F_Liq = F_Liq + 1
            Set RowLiq = Lo_Liq.ListRows(F_Liq)
            
        End If
    
    Next F_Cta

End Sub     ' RuT_Grabar_Datos_CTA_Liq_a_CTA_y_Lst     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================





