Attribute VB_Name = "M72_Del_Reg_No_Válidos_BDatos"
' Last Rev. 2026-09-21 12:12
'2026-01-11
Option Explicit

'        - Borrar Recibos NO Pertinentes:  otros C_Acad, Matrícula=N, AE<>4
'                - Borrar Recibos de otro Curso_Acad
'                - Borrar Recibos que BD_Matricula = "N"
'                - Borrar Recibos que BD_ActivEco <> 4 (Enseñanzas Propias)
'                - Borrar Recibos de Matrículas de coste CERO - ImpMatCero - Recibos Matrícula de Actividad Académica a Coste CERO.
'                - Borrar Recibos Importe CERO - Subvencionado- Imp_Acad <= Imp_Dto

            Sub RuT_Del_Reg_NO_Válidos_Bdatos_ByHand()
                Prog_BD.Select
                Prog_BD.Unprotect
                Call RuT_Del_Reg_NO_Válidos_Bdatos(Prog_BD.ListObjects(1))
            End Sub
'- -------------------------------------------------------------------------------------------------
'- Borrar Rec. NO Válidos: otros C_Acad, Matrícula=N, AE<>4, Matrícula Coste Cero--------
'- -------------------------------------------------------------------------------------------------
Sub RuT_Del_Reg_NO_Válidos_Bdatos(Lo_Data As ListObject)
Debug.Print ">>> RuT_Del_Reg_NO_Válidos"
    Dim rowfind         As Variant
    Dim Cont_Fail       As Long
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim Curso_Acad      As String:      Curso_Acad = Prog__APP.Range("APP_CursAcad")
    Dim TxtMsg1  As String, TxtMsg2  As String, TxtMsg3  As String
    Lo_Data.ShowTotals = False
    Application.DisplayAlerts = False
    With Lo_Data
        .ShowTotals = False
        
        '- Borra los que son de otro Curso_Acad ----------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_C_Acad, Criteria1:="<>" & Curso_Acad   '- Elimino los que son de otro curso
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                        TxtMsg1 = "Del Recibos de Curso-Acad " & ChrW(&H2260) & " " & Curso_Acad
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            rowfind = .ListRows.Count
            If rowfind = 0 Then
                MsgBox "No hay Recibos de Curso-Acad = " & Curso_Acad, vbOKOnly + vbExclamation, "Proceso: Importar LsGes04."
                GoTo Restablecer_Valores
            End If
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & "No hay Recibos de Curso-Acad " & ChrW(&H2260) & " " & Curso_Acad
        End If
        
        '- Borra los que BD_Matricula = "N" --------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_Matricula, xlAscending, True)    '- Ordenar primero accelera un montón el borrado --
        .Range.AutoFilter Field:=BD_Matricula, Criteria1:="=N"                          '- Elimino las NO Martrículas
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                TxtMsg1 = "Del Recibos con Matrícula = N "
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & "No hay Recibos con Matrícula = N. "
        End If
        
        '- Borra los que BD_ActivEco <> 4 (Enseñanzas Propias) -------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .DataBodyRange.Columns(BD_ActivEco).Select     '- Datos - Texto en Columnas - Finalizar - PARA NÚMEROS
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
        Call Rut_Lo_Sort(Lo_Data, BD_ActivEco, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---
        .Range.AutoFilter Field:=BD_ActivEco, Criteria1:="<>4", Operator:=xlAnd, Criteria2:="<>300"        '- Elimino las NO Títulos Propios
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                TxtMsg1 = "Del Recibos con AE " & ChrW(&H2260) & " 4 "
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "No hay Recibos con AE " & ChrW(&H2260) & " 4."
        End If
        
        '-Filtra y Borrar Recibos de Matrículas de coste CERO - ImpMatCero - Recibos Matrícula de Actividad Académica a Coste CERO.
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_ImpDto, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_Data, BD_ImpRec, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AdvancedFilter xlFilterInPlace, Prog_Filtros_TipRec.Range("Tb_CriT_ImpMatCero")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1   '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                TxtMsg1 = "Del Recibos Matrícula de Actividad Acad. a Coste CERO"
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "No hay Recibos de Matrícula de Actividad Académica a Coste CERO."
        End If
        
        '-Filtra y Borrar Recibos Importe CERO - Subvencionado- Imp_Acad <= Imp_Dto ----------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_ImpRec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_ImpRec, Criteria1:="=0,00"
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                TxtMsg1 = "Del Rec Subvencionados 100% (ImpDto " & ChrW(&H2265) & " ImpAcad " & ChrW(&H21D2) & " Rec=0€)"
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            rowfind = .ListRows.Count
            If rowfind = 0 Then
                MsgBox "No hay Rec. Subvencionados ImpDto " & ChrW(&H2265) & " ImpAcad " & ChrW(&H21D2) & " ImpRec = 0", vbOKOnly + vbExclamation, "Proceso: Importar LsGes04."
                GoTo Restablecer_Valores
            End If
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & "No hay Rec. Subvencionados ImpDto " & ChrW(&H2265) & " ImpAcad " & ChrW(&H21D2) & " ImpRec = 0"
        End If
        
    End With    ' Lo_Data
        
Restablecer_Valores:
    Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False     '- Visualizo el progreso
    Lo_Data.ShowTotals = True
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Application.DisplayAlerts = True
End Sub
'- -------------------------------------------------------------------------------------------------





