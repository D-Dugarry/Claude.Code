Attribute VB_Name = "M31_List_PLANES_Orgánicas"
' Last Rev. 2026-09-21 12:12
'2026-02-06
'- M31_Listar_PLANES
Option Explicit
'===================================================================================================
Sub RuT_Listar_Planes_y_sus_Orgánicas()
'===================================================================================================
    Dim Cont                As Integer
    Dim i                   As Long
    Dim Txt_Cabecera        As String
    Dim Txt_Orgánicas       As String
    Dim Txt_detalle         As String
    Dim TipoCurso       As String:      TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")
    
    Dim Lo_RetVRI   As ListObject:      Set Lo_RetVRI = Prog_Coef_Ret_VRI.ListObjects(1)
    
    Const CoefVRI_Plan          As Integer = 1
    Const CoefVRI_CoefVRI       As Integer = 2
    Const CoefVRI_Concepto      As Integer = 3
    Const CoefVRI_TipCurs       As Integer = 4
    Const CoefVRI_Orgánica      As Integer = 5
    Const CoefVRI_Obs           As Integer = 6
    Const CoefVRI_ORden         As Integer = 7

    ' =============  Preparar Tabla ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_Coef_Ret_VRI)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)

    ' Inicio Listado en Tabla excel
    Txt_Cabecera = "Planes de " & Range("APP_EFP_o_CFC") & "_" & Range("APP_CursAcad") & String(10, " ") & Now & vbCrLf & vbCrLf
    
    '- Visualizo el progreso -----------------------------------------------------------------------
    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
    Form_Menu.TB_Informe = Txt_Cabecera
    Txt_Cabecera = "    Plan  Coef.Ret.VRI y Orgánicas"
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera & vbLf
        
    ' Recorro toda la Tabla ------------------------------------------------------------------------
    With Lo_RetVRI.DataBodyRange
    For i = 3 To Lo_RetVRI.ListRows.Count
        Txt_Orgánicas = Replace(.Cells(i, CoefVRI_Orgánica), vbLf, ", ")
'        If Txt_Orgánicas <> "" Then Txt_Orgánicas = ",  " & Txt_Orgánicas & "."
        Txt_detalle = .Cells(i, CoefVRI_Plan) & _
                      Right("     " & .Cells(i, CoefVRI_CoefVRI), 5) & _
                      " %   " & Txt_Orgánicas & vbLf
        If .Cells(i, CoefVRI_ORden) <> "x" Then
            Cont = Cont + 1
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                        Right("   " & Cont, 3) & "º " & Txt_detalle
        End If
        If TipoCurso = "EFP" And .Cells(i, CoefVRI_TipCurs) = TipoCurso Then
            Cont = Cont + 1
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                        Right("   " & Cont, 3) & "º " & Txt_detalle
        ElseIf .Cells(i, CoefVRI_TipCurs) = TipoCurso Then
            Cont = Cont + 1
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
                        Right("   " & Cont, 3) & "º " & Txt_detalle
        End If
    Next
    End With
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & "Fin listado."

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
End Sub     ' RuT_Listar_Planes   ------------------------------------------------------------------
'===================================================================================================




