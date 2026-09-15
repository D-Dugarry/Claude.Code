Attribute VB_Name = "M12_Genera_LIQx_PDF"
' Last Rev. 2026-09-15 14:10
Option Explicit

' ==================================================================================================================================
Sub Rut_Genero_LIQx_PDF_V2()    ' Genera el Listado de Todos los Recibos del Plan, hayan sido liquidados o NO.
' ==================================================================================================================================
Dim F_Liq           As Integer:     F_Liq = 1
Dim F_Inform        As Integer:     F_Inform = 5
Dim T_Alu_Rectif    As Double
Dim T_Alu_Err       As Double
Dim T_Inf_Err       As Double:      T_Inf_Err = 0
Dim T_Inf_Rectif    As Double:      T_Inf_Rectif = 0
Dim T_Adm_Accu      As Double:      T_Adm_Accu = 0
Dim Anulación_Mat   As Double:      Anulación_Mat = 0
Dim Rec_Emi         As Double
Dim Ant_DNI         As String
Dim ColorFondo      As Integer:     ColorFondo = 35

Dim LoT_TPLiquid        As ListObject
Set LoT_TPLiquid = Wk_TitP_Liquid.ListObjects(1)

Rut_Off_Functions

    Call Rut_WrkSheet_Vaciar(Wk_TitP_LIQx_PDF.Name)
    
    Wk_TitP_LIQx_PDF.Select
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.2)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
    End With
    Application.PrintCommunication = True
    '- Genero la cabecera -----------------------------------------------
    Columns(1).ColumnWidth = 9
    Columns(2).ColumnWidth = 32
    Columns(3).ColumnWidth = 14
    Columns(4).ColumnWidth = 10
    Columns(5).ColumnWidth = 4
    Columns(6).ColumnWidth = 12
    Columns(7).ColumnWidth = 12
    Columns(1).Font.Size = 10
    Columns(2).Font.Size = 11
    Columns(3).Font.Size = 10
    Columns(4).Font.Size = 10
    Columns(5).Font.Size = 11
    Columns(6).Font.Size = 11
    Columns(7).Font.Size = 11
    Columns(1).HorizontalAlignment = xlLeft
    Columns(3).HorizontalAlignment = xlCenter
    Columns(4).HorizontalAlignment = xlCenter
    Columns(3).NumberFormat = "00 000000000"
    Columns(4).NumberFormat = "dd-mm-yyyy"
    Columns(6).NumberFormat = "#,##0.00"
    Columns(7).NumberFormat = "#,##0.00"
    Columns(2).Font.Bold = False
    Rows(2).RowHeight = 10
    Rows(4).RowHeight = 10
    Range("a3:g3").Merge
    Range("a3:g3").WrapText = True
    Range("a3:g3").HorizontalAlignment = xlLeft
    Range("a3") = Wk_TitP_Liquid.Range("Liquid_Plan_Name")
    Range("a3").Font.Size = 11
    Rows(3).RowHeight = 40
    Rows(F_Inform).Font.Size = 10
    Rows(F_Inform).Font.Bold = True
    Cells(F_Inform, 1) = "DNI"
    Cells(F_Inform, 2) = "Nombre"
    Cells(F_Inform, 3) = "Referencia"
    Cells(F_Inform, 4) = "Cobro"
    Cells(F_Inform, 5) = "Plazo"
    Cells(F_Inform, 6) = "Importe"
    Cells(F_Inform, 7) = "  Imp. Rectif."
    Cells(F_Inform, 6).HorizontalAlignment = xlRight
    Cells(F_Inform, 7).HorizontalAlignment = xlRight
    Range("a1:g1").Merge
    Range("a1") = "Listado de recibos cobrados de tasas académicas"
    Range("a1").HorizontalAlignment = xlCenter
    Range("a1").Font.Bold = True
    Range("a1").Font.Italic = True
    Range("a1").Font.Size = 16
    Rows(1).RowHeight = 30
    Range("A1:g1").Interior.ColorIndex = 40
    Range("A1:g1").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
    F_Inform = F_Inform + 1
'    '- Salto los primero retistros que no están liquidado, porque sólo me interesan los marcados como Liquidados
'    Do While LoT_TPLiquid.ListRows(F_Liq).Range(CLiq_NumLiquid) = ""
'        If F_Liq < LoT_TPLiquid.ListRows.Count Then
'            F_Liq = F_Liq + 1
'        Else
'            MsgBox "No hay ninguna tasa liquidada", vbExclamation + vbOKOnly, "Módulo: Generar Listado PDF"
'            GoTo Restablecer_Valores
'        End If
'    Loop
    '- Imprimo Datos del 1º registro ------------------------------------------------
    With LoT_TPLiquid.ListRows(F_Liq)
        Rows(F_Inform).RowHeight = 22
        Cells(F_Inform, 1) = .Range(CLiq_DNI)
        Cells(F_Inform, 2) = .Range(CLiq_Nombre)
'        Rows(F_Inform).AutoFit
        Cells(F_Inform, 3) = .Range(CLiq_Ref)
        Cells(F_Inform, 4) = .Range(CLiq_F_Cobro)
        Cells(F_Inform, 5) = .Range(CLiq_Núm_Rec)
        Cells(F_Inform, 6) = .Range(CLiq_Imp_Cob)
        Range(Cells(F_Inform, 1), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
        '- Inicializo valores ------------------------
        Ant_DNI = .Range(CLiq_DNI)
        T_Alu_Rectif = .Range(CLiq_Imp_Cob)
            '- Si tiene Tasa Adm. -----------------------
            If .Range(CLiq_Imp_Adm) <> "" Then
                T_Adm_Accu = -.Range(CLiq_Imp_Adm)
                '- Si tiene Anulación de Matrícula, la Tasa Adm. es Negativa -----------------------
            End If
            If .Range(CLiq_T_Adm_Neg) <> "" Then
                Anulación_Mat = .Range(CLiq_T_Adm_Neg)
            End If
    End With
    '- Recorro toda la tabla del Curso a Liquidar ---------------------------------------------
    For F_Liq = F_Liq + 1 To LoT_TPLiquid.DataBodyRange.Rows.Count
        With LoT_TPLiquid.ListRows(F_Liq)
'            If .Range(CLiq_NumLiquid) = "" Then GoTo SiguienteFila    '- Si no está liquidado, lo saltamos
            F_Inform = F_Inform + 1
            
            '--- Gestiono el registro -----------------------------------------------
            If Ant_DNI = .Range(CLiq_DNI) Then
                '- Imprimo una línea de detalle ----------------------
                Cells(F_Inform, 3) = .Range(CLiq_Ref)
                Cells(F_Inform, 4) = .Range(CLiq_F_Cobro)
                Cells(F_Inform, 5) = .Range(CLiq_Núm_Rec)
                Cells(F_Inform, 6) = .Range(CLiq_Imp_Cob)
                Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
                T_Alu_Rectif = T_Alu_Rectif + .Range(CLiq_Imp_Cob)
                '- Gestiono Si hay Tasa Adm. ---------------------------------
                If .Range(CLiq_Imp_Adm) <> "" Then
                    T_Adm_Accu = T_Adm_Accu - .Range(CLiq_Imp_Adm)
                    T_Alu_Rectif = T_Alu_Rectif - .Range(CLiq_Imp_Adm)
                    '- Si tiene Anulación de Matrícula, la Tasa Adm. es Negativa -----------------------
                    If .Range(CLiq_T_Adm_Neg) <> "" Then
                        Anulación_Mat = Anulación_Mat + .Range(CLiq_T_Adm_Neg)
                    End If
                End If
                
            Else    '. Hay cambio de Alumno ------------------------
                '- Imprimo una línea de Tasa Adm. Si hay del Alumno Anterio ----------------------
                If Anulación_Mat <> 0 Then
                    Cells(F_Inform, 5) = "Importe Anulación Mat.:"
                    Cells(F_Inform, 5).HorizontalAlignment = xlRight
                    Cells(F_Inform, 5).Font.Color = vbRed
                    Cells(F_Inform, 6) = Anulación_Mat
                    Cells(F_Inform, 6).Font.Color = vbRed
                    Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
                    Cells(F_Inform, 7) = T_Adm_Accu
                    Cells(F_Inform, 7).Font.Color = vbRed
                    F_Inform = F_Inform + 1
                Else
                    '- Imprimo una línea de Tasa Adm. Si hay del Alumno Anterio ----------------------
                    If T_Adm_Accu <> 0 Then
                        Cells(F_Inform, 5) = "Importe Adm.:"
                        Cells(F_Inform, 5).HorizontalAlignment = xlRight
                        Cells(F_Inform, 6) = T_Adm_Accu
                        Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
                        F_Inform = F_Inform + 1
                    End If
                End If
                '- Imprimo una línea de Total del Alumno Anterior ----------------------
                T_Alu_Rectif = T_Alu_Rectif + T_Adm_Accu
                Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                Cells(F_Inform, 5) = "Total Alumno:"
                Cells(F_Inform, 5).HorizontalAlignment = xlRight
                If Anulación_Mat <> 0 Then
                    Cells(F_Inform, 6) = T_Alu_Rectif - T_Adm_Accu
                    Cells(F_Inform, 7).Font.Color = vbRed
                Else
                    Cells(F_Inform, 6) = T_Alu_Rectif
                End If
                Cells(F_Inform, 7) = T_Alu_Rectif
                T_Inf_Err = T_Inf_Err + Cells(F_Inform, 6)
                T_Inf_Rectif = T_Inf_Rectif + Cells(F_Inform, 7)
                
                '- Inicializo valores ------------------------
                Ant_DNI = .Range(CLiq_DNI)
                T_Alu_Rectif = .Range(CLiq_Imp_Cob)
                '- Si tiene Tasa Adm. -----------------------
                If .Range(CLiq_Imp_Adm) <> "" Then
                    T_Adm_Accu = -.Range(CLiq_Imp_Adm)
                    '- Si tiene Anulación de Matrícula, la Tasa Adm. es Negativa -----------------------
                End If
                If .Range(CLiq_T_Adm_Neg) <> "" Then
                    Anulación_Mat = .Range(CLiq_T_Adm_Neg)
                Else
                    Anulación_Mat = 0
                End If
                '- Imprimo una línea de Detalle del siguiente Alumno ----------------------
                F_Inform = F_Inform + 1
                Range(Cells(F_Inform, 1), Cells(F_Inform, 6)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                Cells(F_Inform, 1) = .Range(CLiq_DNI)
                Cells(F_Inform, 2) = .Range(CLiq_Nombre)
                Rows(F_Inform).AutoFit
                Rows(F_Inform).RowHeight = Application.WorksheetFunction.Max(22, Rows(F_Inform).RowHeight)
                Cells(F_Inform, 3) = .Range(CLiq_Ref)
                Cells(F_Inform, 4) = .Range(CLiq_F_Cobro)
                Cells(F_Inform, 5) = .Range(CLiq_Núm_Rec)
                Cells(F_Inform, 6) = .Range(CLiq_Imp_Cob)
                
                If ColorFondo = 35 Then ColorFondo = 37 Else ColorFondo = 35
                Range(Cells(F_Inform, 1), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
            End If

        End With
        
SiguienteFila:
    Next
    '- Imprimo una línea de Tasa Adm. Si hay del Alumno Anterio ----------------------
    If Anulación_Mat <> 0 Then
        F_Inform = F_Inform + 1
        Cells(F_Inform, 5) = "Importe Anulación Mat.:"
        Cells(F_Inform, 5).HorizontalAlignment = xlRight
        Cells(F_Inform, 6) = Anulación_Mat
        Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
        Cells(F_Inform, 7) = T_Adm_Accu
        Cells(F_Inform, 7).Font.Color = vbRed
    Else
        '- Imprimo una línea de Tasa Adm. Si hay del Alumno Anterio ----------------------
        If T_Adm_Accu <> 0 Then
            F_Inform = F_Inform + 1
            Cells(F_Inform, 5) = "Importe Adm.:"
            Cells(F_Inform, 5).HorizontalAlignment = xlRight
            Cells(F_Inform, 6) = T_Adm_Accu
            Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
        End If
    End If
    '- Imprimo una línea de Total del Alumno Anterior ----------------------
    F_Inform = F_Inform + 1
    
    '- Imprimo una línea de Total del Alumno Anterior ----------------------
    T_Alu_Rectif = T_Alu_Rectif + T_Adm_Accu
    Cells(F_Inform, 5) = "Total Alumno:"
    Cells(F_Inform, 5).HorizontalAlignment = xlRight
    If Anulación_Mat <> 0 Then
        Cells(F_Inform, 6) = T_Alu_Rectif - T_Adm_Accu
        Cells(F_Inform, 7).Font.Color = vbRed
    Else
        Cells(F_Inform, 6) = T_Alu_Rectif
    End If
    Cells(F_Inform, 7) = T_Alu_Rectif
    T_Inf_Err = T_Inf_Err + Cells(F_Inform, 6)
    T_Inf_Rectif = T_Inf_Rectif + Cells(F_Inform, 7)
    
    '- Imprimo una línea de Total del Informe ----------------------
    F_Inform = F_Inform + 1
    Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlDouble
    Rows(F_Inform).RowHeight = 22
    Cells(F_Inform, 5) = "Total Listado:"
        Cells(F_Inform, 5).HorizontalAlignment = xlRight
    Cells(F_Inform, 6) = T_Inf_Err
        Cells(F_Inform, 6).Font.Bold = True
    Cells(F_Inform, 7) = T_Inf_Rectif
        Cells(F_Inform, 7).Font.Bold = True
        Cells(F_Inform, 7).Font.Color = vbRed
        Cells(F_Inform, 7).Interior.Color = vbYellow
    F_Inform = F_Inform + 1
    Cells(F_Inform, 5) = "Desvío: "
    Cells(F_Inform, 5).HorizontalAlignment = xlRight
    Cells(F_Inform, 6) = T_Inf_Rectif - T_Inf_Err
        Cells(F_Inform, 6).Font.Color = vbRed
        Cells(F_Inform, 6).Font.Bold = True
        
Range("a1").Select
'Call Rut_WrkSheet_To_PDF(Wk_TitP_LIQx_PDF.Name)
Rut_Exportar_La_LIQx_PDF
    
Restablecer_Valores:
Rut_On_Functions
'    IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ------------------------------------------------------
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Wk_TitP_Liquid.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
End Sub     ' Rut_Genero_Liquid_PDF     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' ==================================================================================================================================
Sub Rut_Genero_LIQx_PDF()    ' Genera el Listado de Todos los Recibos del Plan que tienen Liquidación. "If .Range(CLiq_NumLiquid) = "" Then GoTo SiguienteFila"
' ==================================================================================================================================
Dim F_Liq           As Integer:     F_Liq = 1
Dim F_Inform        As Integer:     F_Inform = 5
Dim T_Alu_Rectif    As Double
Dim T_Alu_Err       As Double
Dim T_Inf_Err       As Double:      T_Inf_Err = 0
Dim T_Inf_Rectif    As Double:      T_Inf_Rectif = 0
Dim T_Adm_Accu      As Double:      T_Adm_Accu = 0
Dim Anulación_Mat   As Double:      Anulación_Mat = 0
Dim Rec_Emi         As Double
Dim Ant_DNI         As String
Dim ColorFondo      As Integer:     ColorFondo = 35

Dim LoT_TPLiquid        As ListObject
Set LoT_TPLiquid = Wk_TitP_Liquid.ListObjects(1)

Rut_Off_Functions

    Call Rut_WrkSheet_Vaciar(Wk_TitP_LIQx_PDF.Name)
    
    Wk_TitP_LIQx_PDF.Select
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.2)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
    End With
    Application.PrintCommunication = True
    '- Genero la cabecera -----------------------------------------------
    Columns(1).ColumnWidth = 9
    Columns(2).ColumnWidth = 32
    Columns(3).ColumnWidth = 14
    Columns(4).ColumnWidth = 10
    Columns(5).ColumnWidth = 4
    Columns(6).ColumnWidth = 12
    Columns(7).ColumnWidth = 12
    Columns(1).Font.Size = 10
    Columns(2).Font.Size = 11
    Columns(3).Font.Size = 10
    Columns(4).Font.Size = 10
    Columns(5).Font.Size = 11
    Columns(6).Font.Size = 11
    Columns(7).Font.Size = 11
    Columns(1).HorizontalAlignment = xlLeft
    Columns(3).HorizontalAlignment = xlCenter
    Columns(4).HorizontalAlignment = xlCenter
    Columns(3).NumberFormat = "00 000000000"
    Columns(4).NumberFormat = "dd-mm-yyyy"
    Columns(6).NumberFormat = "#,##0.00"
    Columns(7).NumberFormat = "#,##0.00"
    Columns(2).Font.Bold = False
    Rows(2).RowHeight = 10
    Rows(4).RowHeight = 10
    Range("a3:g3").Merge
    Range("a3:g3").WrapText = True
    Range("a3:g3").HorizontalAlignment = xlLeft
    Range("a3") = Wk_TitP_Liquid.Range("Liquid_Plan_Name")
    Range("a3").Font.Size = 11
    Rows(3).RowHeight = 40
    Rows(F_Inform).Font.Size = 10
    Rows(F_Inform).Font.Bold = True
    Cells(F_Inform, 1) = "DNI"
    Cells(F_Inform, 2) = "Nombre"
    Cells(F_Inform, 3) = "Referencia"
    Cells(F_Inform, 4) = "Cobro"
    Cells(F_Inform, 5) = "Plazo"
    Cells(F_Inform, 6) = "Importe"
    Cells(F_Inform, 7) = "  Imp. Rectif."
    Cells(F_Inform, 6).HorizontalAlignment = xlRight
    Cells(F_Inform, 7).HorizontalAlignment = xlRight
    Range("a1:g1").Merge
    Range("a1") = "Listado de recibos cobrados de tasas académicas"
    Range("a1").HorizontalAlignment = xlCenter
    Range("a1").Font.Bold = True
    Range("a1").Font.Italic = True
    Range("a1").Font.Size = 16
    Rows(1).RowHeight = 30
    Range("A1:g1").Interior.ColorIndex = 40
    Range("A1:g1").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
    F_Inform = F_Inform + 1
    '- Salto los primero retistros que no están liquidado, porque sólo me interesan los marcados como Liquidados
    Do While LoT_TPLiquid.ListRows(F_Liq).Range(CLiq_NumLiquid) = ""
        If F_Liq < LoT_TPLiquid.ListRows.Count Then
            F_Liq = F_Liq + 1
        Else
            MsgBox "No hay ninguna tasa liquidada", vbExclamation + vbOKOnly, "Módulo: Generar Listado PDF"
            GoTo Restablecer_Valores
        End If
    Loop
    '- Imprimo Datos del 1º registro ------------------------------------------------
    With LoT_TPLiquid.ListRows(F_Liq)
        Rows(F_Inform).RowHeight = 22
        Cells(F_Inform, 1) = .Range(CLiq_DNI)
        Cells(F_Inform, 2) = .Range(CLiq_Nombre)
'        Rows(F_Inform).AutoFit
        Cells(F_Inform, 3) = .Range(CLiq_Ref)
        Cells(F_Inform, 4) = .Range(CLiq_F_Cobro)
        Cells(F_Inform, 5) = .Range(CLiq_Núm_Rec)
        Cells(F_Inform, 6) = .Range(CLiq_Imp_Cob)
        Range(Cells(F_Inform, 1), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
        '- Inicializo valores ------------------------
        Ant_DNI = .Range(CLiq_DNI)
        T_Alu_Rectif = .Range(CLiq_Imp_Cob)
            '- Si tiene Tasa Adm. -----------------------
            If .Range(CLiq_Imp_Adm) <> "" Then
                T_Adm_Accu = -.Range(CLiq_Imp_Adm)
                '- Si tiene Anulación de Matrícula, la Tasa Adm. es Negativa -----------------------
            End If
            If .Range(CLiq_T_Adm_Neg) <> "" Then
                Anulación_Mat = .Range(CLiq_T_Adm_Neg)
            End If
    End With
    '- Recorro toda la tabla del Curso a Liquidar ---------------------------------------------
    For F_Liq = F_Liq + 1 To LoT_TPLiquid.DataBodyRange.Rows.Count
        With LoT_TPLiquid.ListRows(F_Liq)
            If .Range(CLiq_NumLiquid) = "" Then GoTo SiguienteFila    '- Si no está liquidado, lo saltamos
            F_Inform = F_Inform + 1
            
            '--- Gestiono el registro -----------------------------------------------
            If Ant_DNI = .Range(CLiq_DNI) Then
                '- Imprimo una línea de detalle ----------------------
                Cells(F_Inform, 3) = .Range(CLiq_Ref)
                Cells(F_Inform, 4) = .Range(CLiq_F_Cobro)
                Cells(F_Inform, 5) = .Range(CLiq_Núm_Rec)
                Cells(F_Inform, 6) = .Range(CLiq_Imp_Cob)
                Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
                T_Alu_Rectif = T_Alu_Rectif + .Range(CLiq_Imp_Cob)
                '- Gestiono Si hay Tasa Adm. ---------------------------------
                If .Range(CLiq_Imp_Adm) <> "" Then
                    T_Adm_Accu = T_Adm_Accu - .Range(CLiq_Imp_Adm)
                    T_Alu_Rectif = T_Alu_Rectif - .Range(CLiq_Imp_Adm)
                    '- Si tiene Anulación de Matrícula, la Tasa Adm. es Negativa -----------------------
                    If .Range(CLiq_T_Adm_Neg) <> "" Then
                        Anulación_Mat = Anulación_Mat + .Range(CLiq_T_Adm_Neg)
                    End If
                End If
                
            Else    '. Hay cambio de Alumno ------------------------
                '- Imprimo una línea de Tasa Adm. Si hay del Alumno Anterio ----------------------
                If Anulación_Mat <> 0 Then
                    Cells(F_Inform, 5) = "Importe Anulación Mat.:"
                    Cells(F_Inform, 5).HorizontalAlignment = xlRight
                    Cells(F_Inform, 5).Font.Color = vbRed
                    Cells(F_Inform, 6) = Anulación_Mat
                    Cells(F_Inform, 6).Font.Color = vbRed
                    Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
                    Cells(F_Inform, 7) = T_Adm_Accu
                    Cells(F_Inform, 7).Font.Color = vbRed
                    F_Inform = F_Inform + 1
                Else
                    '- Imprimo una línea de Tasa Adm. Si hay del Alumno Anterio ----------------------
                    If T_Adm_Accu <> 0 Then
                        Cells(F_Inform, 5) = "Importe Adm.:"
                        Cells(F_Inform, 5).HorizontalAlignment = xlRight
                        Cells(F_Inform, 6) = T_Adm_Accu
                        Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
                        F_Inform = F_Inform + 1
                    End If
                End If
                '- Imprimo una línea de Total del Alumno Anterior ----------------------
                T_Alu_Rectif = T_Alu_Rectif + T_Adm_Accu
                Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                Cells(F_Inform, 5) = "Total Alumno:"
                Cells(F_Inform, 5).HorizontalAlignment = xlRight
                If Anulación_Mat <> 0 Then
                    Cells(F_Inform, 6) = T_Alu_Rectif - T_Adm_Accu
                    Cells(F_Inform, 7).Font.Color = vbRed
                Else
                    Cells(F_Inform, 6) = T_Alu_Rectif
                End If
                Cells(F_Inform, 7) = T_Alu_Rectif
                T_Inf_Err = T_Inf_Err + Cells(F_Inform, 6)
                T_Inf_Rectif = T_Inf_Rectif + Cells(F_Inform, 7)
                
                '- Inicializo valores ------------------------
                Ant_DNI = .Range(CLiq_DNI)
                T_Alu_Rectif = .Range(CLiq_Imp_Cob)
                '- Si tiene Tasa Adm. -----------------------
                If .Range(CLiq_Imp_Adm) <> "" Then
                    T_Adm_Accu = -.Range(CLiq_Imp_Adm)
                    '- Si tiene Anulación de Matrícula, la Tasa Adm. es Negativa -----------------------
                End If
                If .Range(CLiq_T_Adm_Neg) <> "" Then
                    Anulación_Mat = .Range(CLiq_T_Adm_Neg)
                Else
                    Anulación_Mat = 0
                End If
                '- Imprimo una línea de Detalle del siguiente Alumno ----------------------
                F_Inform = F_Inform + 1
                Range(Cells(F_Inform, 1), Cells(F_Inform, 6)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
                Cells(F_Inform, 1) = .Range(CLiq_DNI)
                Cells(F_Inform, 2) = .Range(CLiq_Nombre)
                Rows(F_Inform).AutoFit
                Rows(F_Inform).RowHeight = Application.WorksheetFunction.Max(22, Rows(F_Inform).RowHeight)
                Cells(F_Inform, 3) = .Range(CLiq_Ref)
                Cells(F_Inform, 4) = .Range(CLiq_F_Cobro)
                Cells(F_Inform, 5) = .Range(CLiq_Núm_Rec)
                Cells(F_Inform, 6) = .Range(CLiq_Imp_Cob)
                
                If ColorFondo = 35 Then ColorFondo = 37 Else ColorFondo = 35
                Range(Cells(F_Inform, 1), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
            End If

        End With
        
SiguienteFila:
    Next
    '- Imprimo una línea de Tasa Adm. Si hay del Alumno Anterio ----------------------
    If Anulación_Mat <> 0 Then
        F_Inform = F_Inform + 1
        Cells(F_Inform, 5) = "Importe Anulación Mat.:"
        Cells(F_Inform, 5).HorizontalAlignment = xlRight
        Cells(F_Inform, 6) = Anulación_Mat
        Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
        Cells(F_Inform, 7) = T_Adm_Accu
        Cells(F_Inform, 7).Font.Color = vbRed
    Else
        '- Imprimo una línea de Tasa Adm. Si hay del Alumno Anterio ----------------------
        If T_Adm_Accu <> 0 Then
            F_Inform = F_Inform + 1
            Cells(F_Inform, 5) = "Importe Adm.:"
            Cells(F_Inform, 5).HorizontalAlignment = xlRight
            Cells(F_Inform, 6) = T_Adm_Accu
            Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Interior.ColorIndex = ColorFondo
        End If
    End If
    '- Imprimo una línea de Total del Alumno Anterior ----------------------
    F_Inform = F_Inform + 1
    
    '- Imprimo una línea de Total del Alumno Anterior ----------------------
    T_Alu_Rectif = T_Alu_Rectif + T_Adm_Accu
    Cells(F_Inform, 5) = "Total Alumno:"
    Cells(F_Inform, 5).HorizontalAlignment = xlRight
    If Anulación_Mat <> 0 Then
        Cells(F_Inform, 6) = T_Alu_Rectif - T_Adm_Accu
        Cells(F_Inform, 7).Font.Color = vbRed
    Else
        Cells(F_Inform, 6) = T_Alu_Rectif
    End If
    Cells(F_Inform, 7) = T_Alu_Rectif
    T_Inf_Err = T_Inf_Err + Cells(F_Inform, 6)
    T_Inf_Rectif = T_Inf_Rectif + Cells(F_Inform, 7)
    
    '- Imprimo una línea de Total del Informe ----------------------
    F_Inform = F_Inform + 1
    Range(Cells(F_Inform, 3), Cells(F_Inform, 6)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlDouble
    Rows(F_Inform).RowHeight = 22
    Cells(F_Inform, 5) = "Total Listado:"
        Cells(F_Inform, 5).HorizontalAlignment = xlRight
    Cells(F_Inform, 6) = T_Inf_Err
        Cells(F_Inform, 6).Font.Bold = True
    Cells(F_Inform, 7) = T_Inf_Rectif
        Cells(F_Inform, 7).Font.Bold = True
        Cells(F_Inform, 7).Font.Color = vbRed
        Cells(F_Inform, 7).Interior.Color = vbYellow
    F_Inform = F_Inform + 1
    Cells(F_Inform, 5) = "Desvío: "
    Cells(F_Inform, 5).HorizontalAlignment = xlRight
    Cells(F_Inform, 6) = T_Inf_Rectif - T_Inf_Err
        Cells(F_Inform, 6).Font.Color = vbRed
        Cells(F_Inform, 6).Font.Bold = True
        
Range("a1").Select
'Call Rut_WrkSheet_To_PDF(Wk_TitP_LIQx_PDF.Name)
Rut_Exportar_La_LIQx_PDF
    
Restablecer_Valores:
Rut_On_Functions
'    IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ------------------------------------------------------
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Wk_TitP_Liquid.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
End Sub     ' Rut_Genero_Liquid_PDF     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
' ==================================================================================================================================
Sub Rut_Exportar_La_LIQx_PDF()   '- Copia una Sheet concreta
' ==================================================================================================================================
Rut_Off_Functions
Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Copio la Sheet entera y esto es lo que voy a grabar. -----------------------------------------
    Wk_TitP_LIQx_PDF.Copy
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName  As String
    Dim FullName    As String
    Dim sFileSaveName As Variant
    IntialName = "LIQxPDF_" & Wk_TitP_Liquid.Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
    FullName = FPath & "LIQxPDF_" & Wk_TitP_Liquid.Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(FullName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo GestError
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True, FileFormat:=51
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    '- Quito los Elementos: Botones (Shapes), Comentarios de Celdas y Borro la Fila de Filtrado (la de arriba de los títulos de la Tabla ------------
        ActiveWorkbook.ActiveSheet.Unprotect
'        ActiveWorkbook.ActiveSheet.Shapes.SelectAll:   Selection.Delete
        ActiveWorkbook.ActiveSheet.UsedRange.ClearComments
        Application.EnableEvents = False
    '- Grabo los cambios y Cierro el Archivo ---------------------------
    ActiveWorkbook.Close SaveChanges:=True
    
        MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Archivar Liquidación"
    Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & _
            "  -.-  " & Now() & vbCrLf & vbCrLf & "Exportado el Listado ePDF de la Liquidación de:   " & Wk_TitP_Liquid.Range("Liquid_Plan_Name") & _
            vbCrLf & vbCrLf & "En el Archivo:   " & sFileSaveName
    GoTo Salir_Sub
GestError:
    Debug.Print "Error Rut_Exportar_La_Liquidación ", Err.Number, Err.Description, Err.Source
    Debug.Print sFileSaveName
    MsgBox "Rut_Exportar_La_Liquidación " & "Filename:=" & vbCrLf & sFileSaveName, vbExclamation + vbOKOnly, "Rutina de Remesado"
    Form_Menu.TB_Informe = "Rut_Exportar_La_Liquidación " & vbCrLf & "Filename:=" & sFileSaveName
Salir_Sub:
Call Rut_EnableEvents_Status_Reset
Rut_On_Functions
End Sub     ' Rut_Exportar_La_Liquidación
'-----------------------------------------------------------------------------------------------------------------------------------



' ==================================================================================================================================
Sub Rut_Exportar_La_LIQxn_PDF()   '- Copia una Sheet concreta
' ==================================================================================================================================
Rut_Off_Functions
Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    
    '- Solicitar Número de Liquidación      ---------------------------------------------------
    Dim NumLIQ          As String
    NumLIQ = InputBox("Introducir el número de Liquidación:", "Exportar Liquidación X a PDF")
    If Len(NumLIQ) = 0 Then
        Form_Menu.TB_Informe = "Operación Cancelada"
        Exit Sub
    End If
    '- Filtrar la Liquidación Número X     ---------------------------------------------------
    Application.DisplayAlerts = False
    On Error Resume Next
    Wk_TitP_Liquid.ListObjects(1).AutoFilter.ShowAllData
    Wk_TitP_Liquid.ListObjects(1).Range.AutoFilter Field:=CLiq_NumLiquid, Criteria1:="=" & NumLIQ
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    '- Copio la Sheet entera y esto es lo que voy a grabar. -----------------------------------------
    Wk_TitP_Liquid.Copy
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "LIQ-TitProp_" & Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo GestError
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True, FileFormat:=51
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    '- Quito los Elementos: Botones (Shapes), Comentarios de Celdas y Borro la Fila de Filtrado (la de arriba de los títulos de la Tabla ------------
        ActiveWorkbook.ActiveSheet.Unprotect
        ActiveWorkbook.ActiveSheet.Shapes.SelectAll:   Selection.Delete
        ActiveWorkbook.ActiveSheet.UsedRange.ClearComments
        Application.EnableEvents = False
        ActiveWorkbook.ActiveSheet.Rows(ActiveWorkbook.ActiveSheet.ListObjects(1).Range.Rows(1).Row - 1).Clear
    '- Grabo los cambios y Cierro el Archivo ---------------------------
    ActiveWorkbook.Close SaveChanges:=True
        MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Archivar Liquidación"
    Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & _
            "  -.-  " & Now() & vbCrLf & vbCrLf & "Exportado el Resumen de la Liquidación de:   " & Wk_TitP_Liquid.Range("Liquid_Plan_Name") & _
            vbCrLf & vbCrLf & "En el Archivo:   " & sFileSaveName
    GoTo Salir_Sub
GestError:
    Debug.Print "Error Rut_Exportar_La_Liquidación ", Err.Number, Err.Description, Err.Source
    Debug.Print sFileSaveName
    MsgBox "Rut_Exportar_La_Liquidación " & "Filename:=" & vbCrLf & sFileSaveName, vbExclamation + vbOKOnly, "Rutina de Remesado"
    Form_Menu.TB_Informe = "Rut_Exportar_La_Liquidación " & vbCrLf & "Filename:=" & sFileSaveName
Salir_Sub:
With Wk_TitP_Liquid.ListObjects(1)
    .AutoFilter.ShowAllData
End With
Call Rut_EnableEvents_Status_Reset
Rut_On_Functions
End Sub     ' Rut_Exportar_La_LIQxn_PDF
'-----------------------------------------------------------------------------------------------------------------------------------







'' ==================================================================================================================================
'Sub Rut_Exportar_La_LIQx_PDF()   '- Copia una Sheet concreta
'' ==================================================================================================================================
'Rut_Off_Functions
'Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
'    H_Inicio = timer                ' Para Saber el tiempo de proceso
'    '- Copio la Sheet entera y esto es lo que voy a grabar. -----------------------------------------
'    Wk_TitP_Liquid.Copy
'    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
'    Dim IntialName As String
'    Dim sFileSaveName As Variant
'    IntialName = FPath & "LIQ-TitProp_" & Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
'    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
'        If sFileSaveName <> False Then
'            On Error GoTo GestError
'            Application.DisplayAlerts = False
'            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True, FileFormat:=51
'            Application.DisplayAlerts = True
'            On Error GoTo 0
'        End If
'    '- Quito los Elementos: Botones (Shapes), Comentarios de Celdas y Borro la Fila de Filtrado (la de arriba de los títulos de la Tabla ------------
'        ActiveWorkbook.ActiveSheet.Unprotect
'        ActiveWorkbook.ActiveSheet.Shapes.SelectAll:   Selection.Delete
'        ActiveWorkbook.ActiveSheet.UsedRange.ClearComments
'        Application.EnableEvents = False
'        ActiveWorkbook.ActiveSheet.Rows(ActiveWorkbook.ActiveSheet.ListObjects(1).Range.Rows(1).Row - 1).Clear
'    '- Grabo los cambios y Cierro el Archivo ---------------------------
'    ActiveWorkbook.Close savechanges:=True
'        MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Archivar Liquidación"
'    Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Format(Now() - H_Inicio, "hh.mm.ss") & _
'            " seg.  -.-  " & Now() & vbCrLf & vbCrLf & "Exportado el Resumen de la Liquidación de:   " & Wk_TitP_Liquid.Range("Liquid_Plan_Name") & _
'            vbCrLf & vbCrLf & "En el Archivo:   " & sFileSaveName
'    GoTo Salir_Sub
'GestError:
'    Debug.Print "Error Rut_Exportar_La_Liquidación ", Err.Number, Err.Description, Err.Source
'    Debug.Print sFileSaveName
'    MsgBox "Rut_Exportar_La_Liquidación " & "Filename:=" & vbCrLf & sFileSaveName, vbExclamation + vbOKOnly, "Rutina de Remesado"
'    Form_Menu.TB_Informe = "Rut_Exportar_La_Liquidación " & vbCrLf & "Filename:=" & sFileSaveName
'Salir_Sub:
'Rut_On_Functions
'End Sub     ' Rut_Exportar_La_LIQx_PDF
''-----------------------------------------------------------------------------------------------------------------------------------
