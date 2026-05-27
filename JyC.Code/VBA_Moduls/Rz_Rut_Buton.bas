Attribute VB_Name = "Rz_Rut_Buton"
Option Explicit
'###################################################################################################################################
Sub RuT_Mostrar_Liquidado_T_S()
    If Liquidado_T_S Then
        Liquidado_T_S = False
        Liquidado_T_Bco = False
        Liquidado_T_N = False
        Range("A" & Lcab).AutoFilter C_Cta_Ref_Sol
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_S")).Fill.ForeColor.RGB = RGB(200, 230, 230)
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_Bco")).Fill.ForeColor.RGB = RGB(166, 166, 166)
'        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_N")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Else
        Liquidado_T_S = True
        Liquidado_T_Bco = False
        Liquidado_T_N = False
        Range("A" & Lcab).AutoFilter C_Cta_Ref_Sol, Criteria1:="<>"
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_S")).Fill.ForeColor.RGB = RGB(146, 208, 80)
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_Bco")).Fill.ForeColor.RGB = RGB(255, 70, 70)
'        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_N")).Fill.ForeColor.RGB = RGB(255, 70, 70)
    End If
End Sub     ' RuT_Mostrar_Liquidado_T
'###################################################################################################################################
Sub RuT_Mostrar_Liquidado_T_Bco()
    If Liquidado_T_Bco Then
        Liquidado_T_S = False
        Liquidado_T_Bco = False
        Liquidado_T_N = False
        Range("A" & Lcab).AutoFilter C_Cta_Ref_Sol
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_S")).Fill.ForeColor.RGB = RGB(200, 230, 230)
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_Bco")).Fill.ForeColor.RGB = RGB(166, 166, 166)
'        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_N")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Else
        Liquidado_T_S = False
        Liquidado_T_Bco = True
        Liquidado_T_N = False
        Range("A" & Lcab).AutoFilter C_Cta_Ref_Sol, Criteria1:="="
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_S")).Fill.ForeColor.RGB = RGB(255, 70, 70)
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_Bco")).Fill.ForeColor.RGB = RGB(146, 208, 80)
'        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_N")).Fill.ForeColor.RGB = RGB(255, 70, 70)
    End If
End Sub     ' RuT_Mostrar_Liquidado_T
'###################################################################################################################################
Sub RuT_Mostrar_Liquidado_T_N()
    If Liquidado_T_N Then
        Liquidado_T_S = False
        Liquidado_T_Bco = False
        Liquidado_T_N = False
        Range("A" & Lcab).AutoFilter C_Cta_Ref_Sol
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_S")).Fill.ForeColor.RGB = RGB(200, 230, 230)
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_Bco")).Fill.ForeColor.RGB = RGB(166, 166, 166)
'        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_N")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Else
        Liquidado_T_S = False
        Liquidado_T_Bco = False
        Liquidado_T_N = True
        Range("A" & Lcab).AutoFilter C_Cta_Ref_Sol, Criteria1:="=N"
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_S")).Fill.ForeColor.RGB = RGB(255, 70, 70)
        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_Bco")).Fill.ForeColor.RGB = RGB(255, 70, 70)
'        ActiveSheet.Shapes.Range(Array("Botón_Liquidado_T_N")).Fill.ForeColor.RGB = RGB(146, 208, 80)
    End If
End Sub     ' RuT_Mostrar_Liquidado_T
'###################################################################################################################################
Sub RuT_Mostrar_Pagos()
    If Mostrar_Pagos Then
        Mostrar_Pagos = False
        Mostrar_Cobros = False
        Range("A" & Lcab).AutoFilter C_Cta_Imp
        ActiveSheet.Shapes.Range(Array("Botón_Menor_q_0")).Fill.ForeColor.RGB = RGB(200, 230, 230)
        ActiveSheet.Shapes.Range(Array("Botón_Mayor_q_0")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Else
        Mostrar_Pagos = True
        Mostrar_Cobros = False
    Range("A" & Lcab).AutoFilter C_Cta_Imp, "< 0"
    ActiveSheet.Shapes.Range(Array("Botón_Menor_q_0")).Fill.ForeColor.RGB = RGB(146, 208, 80)
    ActiveSheet.Shapes.Range(Array("Botón_Mayor_q_0")).Fill.ForeColor.RGB = RGB(255, 70, 70)
    End If
End Sub     ' RuT_Mostrar_Pagos
'###################################################################################################################################
Sub RuT_Mostrar_Cobros()
    If Mostrar_Cobros Then
        Mostrar_Pagos = False
        Mostrar_Cobros = False
        Range("A" & Lcab).AutoFilter C_Cta_Imp
        ActiveSheet.Shapes.Range(Array("Botón_Menor_q_0")).Fill.ForeColor.RGB = RGB(200, 230, 230)
        ActiveSheet.Shapes.Range(Array("Botón_Mayor_q_0")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Else
        Mostrar_Cobros = True
        Mostrar_Pagos = False
    Range("A" & Lcab).AutoFilter C_Cta_Imp, "> 0"
    ActiveSheet.Shapes.Range(Array("Botón_Mayor_q_0")).Fill.ForeColor.RGB = RGB(146, 208, 80)
    ActiveSheet.Shapes.Range(Array("Botón_Menor_q_0")).Fill.ForeColor.RGB = RGB(255, 70, 70)
    End If
End Sub     ' RuT_Mostrar_Cobros
'###################################################################################################################################
Sub RuT_Ocultar_Apuntes_Superfluos()
    If Ocultar_Superfluos Then
        Ocultar_Superfluos = False
        Range("A" & Lcab).AutoFilter C_Cta_Siglas
        ActiveSheet.Shapes.Range(Array("Botón_Ocultar_Ap_Superfluos")).Fill.ForeColor.RGB = RGB(200, 230, 230)
        ActiveSheet.Shapes.Range(Array("Botón_Ocultar_Ap_Superfluos")).TextFrame.Characters.Text = "Oculta Apuntes Superfluos"
    Else
        Ocultar_Superfluos = True
        Range("A" & Lcab).AutoFilter C_Cta_Siglas, Criteria1:="<>__*"
        ActiveSheet.Shapes.Range(Array("Botón_Ocultar_Ap_Superfluos")).Fill.ForeColor.RGB = RGB(146, 208, 80)
        ActiveSheet.Shapes.Range(Array("Botón_Ocultar_Ap_Superfluos")).TextFrame.Characters.Text = "Muestra " & ChrW(11375) & " los Apuntes"
'        Range("f1") = AscW(Range("f2"))
    End If
End Sub     ' RuT_Ocultar_Apuntes_Superfluos
'###################################################################################################################################
Sub RuT_Mostrar_Apuntes_Superfluos()
        Range("A" & Lcab).AutoFilter C_Cta_Siglas, Criteria1:="_*"
End Sub     ' RuT_Ocultar_Apuntes_Superfluos

'###################################################################################################################################
Sub RuT_Ocultar_O_Mostrar_Col_Marcadas()
On Error Resume Next
    Ultima_Col = Cells(Lcab, Columns.Count).End(xlToLeft).Column    ' Última Columna con datos
    For Cont_Col = 1 To Ultima_Col
        If IsEmpty(Cells(Lcab - 2, Cont_Col)) Then
            If Columns(Cont_Col).Hidden Then Columns(Cont_Col).Hidden = False
        Else
            If BooL_Col_Ocultas Then
                Columns(Cont_Col).Hidden = False
            Else
                Columns(Cont_Col).Hidden = True
            End If
        End If
    Next Cont_Col
    If BooL_Col_Ocultas Then
        BooL_Col_Ocultas = False
        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(200, 230, 230)
        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Oculta Col's"
    Else
        BooL_Col_Ocultas = True
        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(146, 208, 80)
        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Muestra Col's"
    End If
'''    Range("f" & Lcab).Select
'''    ActiveCell.Offset(1, 1).Select
End Sub     ' RuT_Ocultar_O_Mostrar_Col_Marcadas

'###################################################################################################################################
Sub RuT_DesHacer_Filtros()
On Error Resume Next
    Mostrar_2023 = False
    Mostrar_2024 = False
    Mostrar_Cobros = False
    Mostrar_Pagos = False
    Liquidado_T_S = False
    Liquidado_T_Bco = False
    Liquidado_T_N = False
    ' -------------------------------------------------------
    '   Quitar color botones
    ' -------------------------------------------------------
'    Prog_CTA_Tb.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Prog_CTA_Tb.Shapes.Range(Array("Botón_Ver_2024")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Prog_CTA_Tb.Shapes.Range(Array("Botón_Ver_2023")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Prog_CTA_Tb.Shapes.Range(Array("Botón_Liquidado_T_S")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Prog_CTA_Tb.Shapes.Range(Array("Botón_Liquidado_T_Bco")).Fill.ForeColor.RGB = RGB(166, 166, 166)
    Prog_CTA_Tb.Shapes.Range(Array("Botón_Liquidado_T_N")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Prog_CTA_Tb.Shapes.Range(Array("Botón_Mayor_q_0")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    Prog_CTA_Tb.Shapes.Range(Array("Botón_Menor_q_0")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    ' -------------------------------------------------------
    '   quitar la ordenación por fecha.
    ' -------------------------------------------------------
    Prog_CTA_Tb.Cells(Lcab + 1, 1).Select    ' si no te colocas dentro de la tabla no la reconoce
    If Prog_CTA_Tb.FilterMode Then Prog_CTA_Tb.ShowAllData
    ' -------------------------------------------------------
    '   restablecer la ordenación del Ordinal en ascendente
    ' -------------------------------------------------------
    Call Rut_LstObj_Sort(Prog_CTA_Tb.ListObjects(1), C_Cta_Liq_Ordinal, xlAscending, True)
    Prog_CTA_Tb.ListObjects(1).DataBodyRange(1).Select
End Sub     ' RuT_DesHacer_Filtros
'###################################################################################################################################
Sub RuT_Quitar_Sombreados()
On Error Resume Next
    Ultima_Fila = Cells(Rows.Count, 1).End(xlUp).Row                ' Última Fila con datos
    Ultima_Col = Cells(Lcab, Columns.Count).End(xlToLeft).Column    ' Última Columna con datos
    Range(Cells(Lcab + 1, 1), Cells(Ultima_Fila, Ultima_Col)).Interior.Color = -1
End Sub     ' RuT_Quitar_Sombreados
'###################################################################################################################################
Sub RuT_Mostrar_Todas_Col()
On Error Resume Next
    Ultima_Col = Cells(Lcab, Columns.Count).End(xlToLeft).Column    ' Última Columna con datos
    Columns(1).Resize(, Ultima_Col).Hidden = False
    ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Oculta Col's"
    ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(200, 230, 230)
    BooL_Col_Ocultas = False
End Sub     ' RuT_Mostrar_Todas_Col

