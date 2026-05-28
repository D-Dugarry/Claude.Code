Attribute VB_Name = "M_051_Load_TPV_Solic"
'- M_051_Load_TPV_Solic
' Generador de informe en: https://cvnet.cpd.ua.es/uaGenInf/Home/Consulta/38051
Option Explicit

Dim Finalizar_Proceso           As Boolean
Dim AñoCont                     As String           ' Para Controlar el cambio de años el en número de orden que genero

' ==================================================================================================================================
Sub Load_TPV_Solic()    '= Load_Request
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim F_TPV           As Long
    Dim F_Sol           As Long:        F_Sol = 12
    Dim RowLiq          As ListRow
    Dim RowTPV          As ListRow
    Dim Reg_Ref         As String

Rut_Off_Functions
    
'    H_Liq_TPV.Visible = True
'    H_Liq_TPV.Select

    Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Set Lo_Liq = H_Liq_TPV.ListObjects(1)

    Application.ScreenUpdating = False
    
    Call Rut_LstObj_WrkSht_Preparar(Application.Workbooks(ThisWorkbook.Name).Sheets(Prog_TPV_Tb.Name))
    Call Rut_LstObj_Sort(Lo_TPV, 1, xlAscending, True)
    
'    H_Liq_TPV.Select
    Application.Calculation = xlCalculationManual:      Application.EnableEvents = False:   Application.DisplayAlerts = False
    Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
    Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
    Lo_Liq.ShowTotals = False
    If Not Lo_Liq.DataBodyRange Is Nothing Then Lo_Liq.DataBodyRange.Delete
    
        Range("b1") = Now()
        Range("b2").ClearContents
        Range("e1:e9").ClearContents
        Range("f3") = ""
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
        
Do Until Prog_TPV_Solic.Cells(F_Sol, C_Sol_Ref) = ""
    
'    Call Rut_LstObj_Buscar(Prog_TPV_Tb.ListObjects(1), 9924002323730#, C_TPV_Ref, F_TPV)
    Reg_Ref = Right(CStr(Prog_TPV_Solic.Cells(F_Sol, C_Sol_Ref)), 13)
    Debug.Print CStr(Prog_TPV_Solic.Cells(F_Sol, C_Sol_Ref))
    Debug.Print Reg_Ref
    Call Rut_LstObj_Buscar(Lo_TPV, Reg_Ref, C_TPV_Ref, F_TPV)
    If F_TPV = 0 Then
        MsgBx_Msg = "Referencia de Pagos-TPV no localizada." & vbLf & vbLf & _
                    "ref: _" & CStr(Prog_TPV_Solic.Cells(F_Sol, C_Sol_Ref)) & vbLf & vbLf & vbLf & _
                    "Ver de Actualizar la Tabla de Cobros por TPV/Bizum/Fly." & vbLf & vbLf & _
                    "Importando del Generador de Informes, el Informe: 'Pagos no académicos'."
        MsgBx_Title = "Proceso: Cargar Solicitud de Liquidación TPV"
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(, True, , , "Stop"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
        GoTo Restablecer_Valores
    End If
    Set RowTPV = Lo_TPV.ListRows(F_TPV)
    
        ' #############################==================  Tratamos los datos ==================###############################
        Set RowLiq = Lo_Liq.ListRows.Add

        RowLiq.Range(C_TPV_Liq_Ref) = RowTPV.Range(C_TPV_Ref)
        RowLiq.Range(C_TPV_Liq_F_Pag) = RowTPV.Range(C_TPV_F_Pago)
        RowLiq.Range(C_TPV_Liq_Form_Pag) = RowTPV.Range(C_TPV_Tipo_Pago)
        RowLiq.Range(C_TPV_Liq_Ordenante) = RowTPV.Range(C_TPV_NomApe)
        RowLiq.Range(C_TPV_Liq_Inscrito) = Prog_TPV_Solic.Cells(F_Sol, C_Sol_Inscrit)
        RowLiq.Range(C_TPV_Liq_Imp) = RowTPV.Range(C_TPV_Imp)
        RowLiq.Range(C_TPV_Liq_ComBco) = RowTPV.Range(C_TPV_ComBco)
        RowLiq.Range(C_TPV_Liq_Neto) = RowTPV.Range(C_TPV_Neto)
        
        If RowLiq.Range(C_TPV_Liq_Neto) > 0 And RowLiq.Range(C_TPV_Liq_Imp) = "" Then RowLiq.Range(C_TPV_Liq_Imp) = RowLiq.Range(C_TPV_Liq_Neto)
        
        RowLiq.Range(C_TPV_Liq_Obs) = RowTPV.Range(C_TPV_Obs)
        RowLiq.Range(C_TPV_Liq_JI) = RowTPV.Range(C_TPV_JI)
        RowLiq.Range(C_TPV_Liq_ExpAdm) = RowTPV.Range(C_TPV_ExpAdm)
        RowLiq.Range(C_TPV_Liq_N_RDT) = RowTPV.Range(C_TPV_RDT)
        RowLiq.Range(C_TPV_Liq_DI) = RowTPV.Range(C_TPV_DI)
        RowLiq.Range(C_TPV_Liq_PMP) = RowTPV.Range(C_TPV_PMP)

        RowLiq.Range(C_TPV_Liq_Siglas) = RowTPV.Range(C_TPV_Siglas)
        RowLiq.Range(C_Cta_Liq_Ref_Sol) = RowTPV.Range(C_TPV_Ref_Sol)
        RowLiq.Range(C_TPV_Liq_N_Liq) = RowTPV.Range(C_TPV_N_Liq)
        RowLiq.Range(C_TPV_Liq_Org) = RowTPV.Range(C_TPV_Org)
        
'        H_Liq_TPV.Range("Liq_Núm") = RowTPV.Range(C_TPV_N_Liq)
'        H_Liq_TPV.Range("Liq_ExpAdm") = RowTPV.Range(C_TPV_ExpAdm)
        
        '- Añado Referencias si no existen ya en la celda...
        If Len(RowTPV.Range(C_TPV_N_Liq)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_Núm"), RowTPV.Range(C_TPV_N_Liq)) = 0 Then _
                        H_Liq_TPV.Range("Liq_Núm") = H_Liq_TPV.Range("Liq_Núm") & " - " & RowTPV.Range(C_TPV_N_Liq)
                        
        If Len(RowTPV.Range(C_TPV_Ref_Sol)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_Solicitud"), RowTPV.Range(C_TPV_Ref_Sol)) = 0 Then _
                        H_Liq_TPV.Range("Liq_Solicitud") = H_Liq_TPV.Range("Liq_Solicitud") & " - " & RowTPV.Range(C_TPV_Ref_Sol)
                        
        If Len(RowTPV.Range(C_TPV_JI)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_JI"), RowTPV.Range(C_TPV_JI).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_JI") = H_Liq_TPV.Range("Liq_JI") & " - " & RowTPV.Range(C_TPV_JI)
                        
        If Len(RowTPV.Range(C_TPV_ExpAdm)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_ExpAdm"), RowTPV.Range(C_TPV_ExpAdm).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_ExpAdm") = H_Liq_TPV.Range("Liq_ExpAdm") & " - " & RowTPV.Range(C_TPV_ExpAdm)
                        
        If Len(RowTPV.Range(C_TPV_DI)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_DI"), RowTPV.Range(C_TPV_DI).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_DI") = H_Liq_TPV.Range("Liq_DI") & " - " & RowTPV.Range(C_TPV_DI)
        If Len(RowTPV.Range(C_TPV_PMP)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_PMP"), RowTPV.Range(C_TPV_PMP).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_PMP") = H_Liq_TPV.Range("Liq_PMP") & " - " & RowTPV.Range(C_TPV_PMP)
        If Len(RowTPV.Range(C_TPV_RDT)) > 0 And InStr(1, H_Liq_TPV.Range("Liq_RDT"), RowTPV.Range(C_TPV_RDT).Value) = 0 Then _
                        H_Liq_TPV.Range("Liq_RDT") = H_Liq_TPV.Range("Liq_RDT") & " - " & RowTPV.Range(C_TPV_RDT)
        '- Ajusto el alto de la fila añadida...
        RowLiq.Range.EntireRow.AutoFit
    F_Sol = F_Sol + 1
    
Loop
    
    Lo_Liq.ShowTotals = True
    
    Range("e3") = Lo_Liq.TotalsRowRange.Columns(C_TPV_Liq_Imp)
    Range("f3") = Lo_Liq.TotalsRowRange.Columns(C_TPV_Liq_Imp)
    Range("f4") = Lo_Liq.TotalsRowRange.Columns(C_TPV_Liq_Ref)
    
    If H_Liq_TPV.Range("Liq_Núm") <> "" And Not H_Liq_TPV.Range("Liq_Núm") <> "x" Then
        MsgBx_Msg = "Algún apunte ha sido liquidado en: " & H_Liq_TPV.Range("Liq_Núm") & vbCrLf & vbCrLf & _
                    "Mirar las columnas finales de datos de la tabla de liquidación."
        MsgBx_Title = "Proceso: Cargar los Datos de una solicitud de liquidación de JyC, TPV / Bizum."
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "stop"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
    End If
       
Restablecer_Valores:
    Lo_Liq.ShowTotals = True
    Rut_On_Functions
'    Prog_N43_TxT.Visible = xlSheetVeryHidden
'    Prog_N43_CTA.Visible = xlSheetVeryHidden
End Sub     '- Load_TPV_Solic
' ==================================================================================================================================

' ==================================================================================================================================
Sub Import_Charge_TPV_Solic_Range_SIN_USO()     '- OJO versión para seleccionar uno mismo el rango de referencias en la solicitud, Sin USO
' ----------------------------------------------------------------------------------------------------------------------------------
Rut_Off_Functions
    Prog_TPV_Solic.Select
    Prog_TPV_Solic.Visible = True
    Range("a1").Select
    
    '- Select a Range ---------------------------------------------------------------------------------
    Dim Rng As Range
        On Error Resume Next
              Set Rng = Application.InputBox(Title:="Select a Range", _
                          Prompt:="Selecciona el rango de Referencias de Pagos de la JyC.", Type:=8)
              Set Rng = Rng.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    
    If Rng Is Nothing Then
        MsgBx_Msg = "Rango vacio = Operación Anulada"
        MsgBx_Title = "Proceso: Importar Datos Solicitud de Liquidación."
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
        Prog__APP.Range("APP_Task_Inf") = "Proceso Cancelado: Importar Datos Solicitud de Liquidación. " & vbCrLf & Now()
        GoTo Restablecer_Valores
    End If
    
    
    Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Set Lo_Liq = H_Liq_TPV.ListObjects(1)
    
    
    
    Dim Cel     As Range
    For Each Cel In Rng
        Debug.Print Cel
    Next Cel
    
        
        Rng.Copy
        ThisWorkbook.Sheets(Prog_TPV_Solic.Name).Cells(2, 2).PasteSpecial xlPasteValues
    
    
'    Dim Cel     As Range
'    Do
'        Debug.Print Cel
'    loop until
'    H_Liq_TPV.Select
'    H_Liq_TPV.Visible = True
'    H_Liq_TPV.Unprotect
'
'    If Range("c11") <> "Referencia TPV" Then
'        MsgBx_Msg = "Algo pasa con el Excel, no tiene el formato correcto."
'        MsgBx_Title = "Proceso: Importar Solicitud de Liquidación de TPV"
'        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
'        prog__app.range("APP_Task_Inf") = "Proceso Cancelado: " & Now()
'        GoTo Restablecer_Valores
'    End If
    
Restablecer_Valores:
Rut_On_Functions
'    Prog_N43_TxT.Visible = xlSheetVeryHidden
'    Prog_N43_CTA.Visible = xlSheetVeryHidden
End Sub     '- Import_Charge_TPV_Solic_Range
' ==================================================================================================================================

'Sub kk2()
'    '- Select a Range ---------------------------------------------------------------------------------
'    Dim Rng As Range
'        On Error Resume Next
'              Set Rng = Application.InputBox(Title:="Select a Range", _
'                          Prompt:="Selecciona el rango de Referencias de Pagos de la JyC.", Type:=8)
'              Set Rng = Rng.SpecialCells(xlCellTypeVisible)
'        On Error GoTo 0
'
'    Dim Cel     As Range
'    If Rng Is Nothing Then
'        MsgBx_Msg = "Rango vacio = Operación Anulada"
'        MsgBx_Title = "Proceso: Importar Datos Liquidación"
'        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
'        prog__app.range("APP_Task_Inf") = "Proceso Cancelado: " & Now()
'        GoTo Restablecer_Valores
'    End If
'    For Each Cel In Rng
'        Debug.Print Cel
'    Next Cel
'Restablecer_Valores:
'End Sub
'
'Sub kk()
'Range(Cells(8900, 1), Cells(8900, 10)).Copy
'Cells(8914, 1).Select
''Range(Cells(8914, 1), Cells(8914, 1)).Select
''Range("a8914").Select
''Selection.Paste
'ActiveSheet.Paste
'End Sub




