Attribute VB_Name = "M08_Actualizar_Tb_Coef_VRI"
' Last Rev. 2026-09-18 19:19
'2026-02-06
'- M31_Listar_PLANES
Option Explicit
'- Actualiza la Tabla de Referencia de los Coeficientes de Retención para el VRI
'==================================================================================================================================
Sub RuT_Lo_Coef_VRI_Actualizar()
'==================================================================================================================================
    Dim Cont                As Long
    Dim ContIni             As Long:        ContIni = 1
    Dim Coef_VRI            As Integer
    Dim Cod_Plan            As String
    Dim Orgánica            As String
    Dim Txt_Cabecera        As String
    Dim Cont_Plan           As Integer:     Cont_Plan = 1
    Dim PlanesSinCob        As Integer:     PlanesSinCob = 0
    
    Dim RegsEmis            As Long     ' regs Emitidos
    Dim Imp_Emis            As Currency
    Dim RegsCobr            As Long     ' regs Cobrados
    Dim Imp_Cobr            As Currency
    Dim RegsPdts            As Long     ' regs Pendiente de pago
    Dim Imp_Pdte            As Currency
    
    Dim Regs_Dev            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula
    Dim Imp_Devo            As Currency
    Dim RDev_Pag            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, Pagados
    Dim IDev_Pag            As Currency
    Dim RDevPdte            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, No Pagados, pero Ajuste Matrícula NO deben ser pagados
    Dim IDevPdte            As Currency
    
    Dim Lo_BD       As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
    Dim Lo_RetVRI   As ListObject:      Set Lo_RetVRI = Prog_Coef_Ret_VRI.ListObjects(1)
    Dim RowNew      As ListRow
    Dim rowfind     As Variant
    
    Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    
    Dim T_ImpEmi        As Currency
    Dim T_ImpCob        As Currency
    Dim T_ImpPdt        As Currency

    ' =============  Preparar Hojas de Tablas ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, True)
    Call Rut_Lo_WrkSht_Preparar(Prog_Coef_Ret_VRI)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)


With Lo_BD.DataBodyRange
    ' Recorro toda la Tabla ---------------------------------------------------------------------------------------
    For Cont = 1 To Lo_BD.ListRows.Count
        If .Cells(Cont, BD_Plan) <> Cod_Plan Then       '- <<<<<  Cambio PLAN BDatos NUEVO  >>>>
            Cod_Plan = .Cells(Cont, BD_Plan)
            Coef_VRI = .Cells(Cont, BD_Coef_VRI)
            Orgánica = .Cells(Cont, BD_Orgánica)
            Cont_Plan = Cont_Plan + 1
            '   Buscar Plan en Lo_RetVRI
            rowfind = Application.Match(Cod_Plan, Lo_RetVRI.DataBodyRange.Columns(CoefVRI_Plan), 0)
            If Not IsError(rowfind) Then    ' Plan Encontrado en Lo_RetVRI ==>>
                Set RowNew = Lo_RetVRI.ListRows(rowfind)
                RowNew.Range(CoefVRI_Concepto) = .Cells(Cont, BD_Concepto)
                RowNew.Range(CoefVRI_TipCurs) = .Cells(Cont, BD_Tipo_EP)
                If InStr(RowNew.Range(CoefVRI_Orgánica), Orgánica) = 0 And Orgánica <> "" Then
                    If RowNew.Range(CoefVRI_Orgánica) = "" Then
                        RowNew.Range(CoefVRI_Orgánica) = Orgánica & "_(" & .Cells(Cont, BD_C_Acad) & ")"
                    Else
                        RowNew.Range(CoefVRI_Orgánica) = RowNew.Range(CoefVRI_Orgánica) & vbLf & Orgánica & "_(" & .Cells(Cont, BD_C_Acad) & ")"
                    End If
                End If
            Else                            ' Plan No Existe, lo CREO en Lo_RetVRI==>>
                Set RowNew = Lo_RetVRI.ListRows.Add
                RowNew.Range(CoefVRI_Plan) = Cod_Plan
                RowNew.Range(CoefVRI_CoefVRI) = Coef_VRI
                If Orgánica <> "" Then RowNew.Range(CoefVRI_Orgánica) = Orgánica & "_(" & .Cells(Cont, BD_C_Acad) & ")"
                RowNew.Range(CoefVRI_Concepto) = .Cells(Cont, BD_Concepto)
                RowNew.Range(CoefVRI_TipCurs) = .Cells(Cont, BD_Tipo_EP)
                RowNew.Range(CoefVRI_ORden) = "x"
            End If
            
        End If
        If RowNew.Range(CoefVRI_ORden) = "x" Then
               RowNew.Range(CoefVRI_CoefVRI) = Application.Max(RowNew.Range(CoefVRI_CoefVRI), Coef_VRI)
            If RowNew.Range(CoefVRI_CoefVRI) = 0 Then
                If Range("APP_EFP_o_CFC") = "EFP" Then
                    RowNew.Range(CoefVRI_CoefVRI) = Lo_RetVRI.DataBodyRange.Cells(1, CoefVRI_CoefVRI)
                Else
                    RowNew.Range(CoefVRI_CoefVRI) = Lo_RetVRI.DataBodyRange.Cells(2, CoefVRI_CoefVRI)
                End If
            End If
        End If
    Next
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)
    
    '- Visualizo el progreso
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                        " Actualizada la Tabla de Referencia de los Coef_VRI." & vbCrLf
End With    '- Lo_BD.DataBodyRange

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Lo_Coef_VRI_Actualizar
'===================================================================================================================================





