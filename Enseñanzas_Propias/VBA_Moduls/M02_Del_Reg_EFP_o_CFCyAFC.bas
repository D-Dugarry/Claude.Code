Attribute VB_Name = "M02_Del_Reg_EFP_o_CFCyAFC"
'20265-01-11
Option Explicit

' ==================================================================================================================================
'- -----------------------------------------------------------------------------------------------------------------------------
'- Borrar Tipo de Enseñanza EFP o CFCyAFC (EFP, CFC, AFC, TNCT) No Deseados ---------------------------------
'- -----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Borrar_Rec_EFP_o_CFCyAFC(Lo_Data As ListObject)
Debug.Print ">>> Rut_Borrar_Rec_EFP_o_CFCyAFC"
    
    Dim Curso_Acad      As String:      Curso_Acad = Prog__APP.Range("APP_CursAcad")
    Dim TipoCurso       As String:      TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")
    Dim rowfind             As Variant
    Dim TxT_Progreso        As String
    Dim TxtMsg1  As String, TxtMsg2  As String, TxtMsg3  As String
    Dim Cont_Fail           As Long

    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_Lo_Sort(Lo_Data, BD_TipoCurso, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----

    Application.DisplayAlerts = False
    With Lo_Data
        .ListColumns.Add
        .ShowTotals = False
        .Range(BD_TipoCurso) = "TipoCurso"  ' para q funcionen los criterios de filtro deben tener ese nombre en la cabecera de la Col.
        .Range(BD_ActivEco) = "Activ_Eco"  ' para q funcionen los criterios de filtro deben tener ese nombre en la cabecera de la Col.
        '-Filtra Recibos - EFP - Estudios de Formación Permanente: Máster, Especialista, Experto. -------------------
        .Range.AdvancedFilter xlFilterInPlace, Prog_Filtros_Concepto.Range("Tb_CriT_EFP")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then     '- hay rec. de EFP
            If TipoCurso = "EFP" Then
                .DataBodyRange.Columns(.ListColumns.Count).SpecialCells(xlCellTypeVisible).Cells.Value = "EFP"
                '- Borra los que NO son "Tít. Propios" EFP ----
                Call Rut_Lo_Filtros_Quitar(Lo_Data)
                .Range.AutoFilter Field:=.ListColumns.Count, Criteria1:="<>EFP"     '- Selecciono los que NO mson EFP
                rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
                If rowfind > .ListColumns.Count Then
                    .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                        TxtMsg1 = "Borrados Recibos de estudios NO EFP_" & Curso_Acad
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
                End If
                GoTo FinRut
            Else
                .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                        TxtMsg1 = "Borrados Recibos de estudios EFP_" & Curso_Acad
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            End If
        Else    '- No hay Rec. EFP
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe("NO hay Rec. de estudios de EFP_", , " quedan: " & Format(.ListRows.Count, "#,##0") & "reg.")
        End If
FinRut:
        .ListColumns(.ListColumns.Count).Delete
    End With    '-  Lo_Data
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Application.DisplayAlerts = True
End Sub     '- Rut_Incorporar_Concept_Eco_y_Tipo_Enseñanza -------------------------------------------------------------------------




