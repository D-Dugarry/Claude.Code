Attribute VB_Name = "M_211_Del_Null_Reg_CAcadAnt"
'2026-01-23
'- M_211_Remove_Null_Reg_CAcad
Option Explicit

'    - M_211, Filtrar y Borrar Registros NO deseados:
'        - Borrar Recibos AE4 Enseñanzas Propias
'        - Borrar Recibos de Movimiento Menos AE=80 'Pruebas Acceso UA'
'        - Borrar Recibos con Imp.Rec. < 0
'        - Borrar Recibos de Matrículas de coste CERO
'        - Borrar Borrar Recibos - Subvencionado-, Imp_Rec =0 porque Imp_Dto >0
'        - Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"
'        - Borrar Recibos BD_C_Acad <> C_Acad_Ant
'        - Borrar Recibos ANULADOS
'        - Borrar Recibos NO Martrícula
'        - Borrar Recibos INVALIDADOS

            Sub RuT_Remove_Null_Reg_ByHand()
                
                Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
                Dim Lo_BD_Ant           As ListObject:      Set Lo_BD_Ant = Sht__BD_Ant.ListObjects(1)
                Dim Lo_BD_Dpl           As ListObject:      Set Lo_BD_Dpl = Sht__BD_Dupl.ListObjects(1)
                Dim Lo_BD_ErrDate           As ListObject:      Set Lo_BD_ErrDate = Sht__BD_ErrDate.ListObjects(1)
                Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
                
                Dim Lo_DrWrk            As ListObject:      Set Lo_DrWrk = Prog_DrWrk.ListObjects(1)
                
                Dim Lo_BdM013           As ListObject:      Set Lo_BdM013 = Sht__BD_M013.ListObjects(1)
                Dim Lo_BdPNB1           As ListObject:      Set Lo_BdPNB1 = Sht__BD_PNB1.ListObjects(1)
                Dim Lo_BdAdmP           As ListObject:      Set Lo_BdAdmP = Sht__BD_AdmP.ListObjects(1)
                
                Call RuT_Remove_Null_Reg_CAcad(ActiveSheet.ListObjects(1))
                
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Filtrar y Borrar Registros NO deseados -------------------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Remove_Null_Reg_CAcad(Lo_G04 As ListObject)
                        
Debug.Print ">>> RuT_Remove_Null_Reg_CAcad"
    Dim rowfind         As Variant
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")

    Call Rut_Lo_Filtros_Quitar(Lo_G04)                '- Quitar filtros
    '- ----------------------------------------------------------------------------------------------------------------------------
    With Lo_G04
        .ShowTotals = False
                
'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos AE4 Enseñanzas Propias -----------------------------------------------------------------------------------
        If Prog__APP.Range("SW_DelRegAE4") Then
            rowfind = .ListRows.Count
            Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_G04, BD_ActivEco, "=4")
            rowfind = rowfind - .ListRows.Count
            If rowfind > 0 Then
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. AE4 ", 0, _
                                                                Format(rowfind, " #,##0") & " reg. ", _
                                                                " de " & Format(.ListRows.Count, "#,##0") & " reg.")
                If Lo_G04.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
            Else
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. AE4 ", 0)
            End If
        Else
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Se ha decidido mantener los Rec. AE4", 0)
        End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos de Movimiento Menos AE=80 'Pruebas Acceso UA' -----------------------------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_G04, BD_RegMov, "=S", BD_ActivEco, "<>80")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec.Mov. Menos AE=80 'Pruebas Acceso UA'", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            If Lo_G04.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de Movimiento. ", 0)
        End If
                
'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos con Imp.Rec. < 0 -----------------------------------------------------------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_G04, BD_ImpRec, "<0")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. con Imp.Rec. < 0. ", 0, _
                                                            Format(rowfind, " #,##0") & " reg. ", _
                                                            " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            If Lo_G04.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Negativos. ", 0)
        End If
                
'- ---------------------------------------------------------------------------------------------------------------
'- Borrar Recibos de Matrículas de coste CERO --------------------------------------------------------------------
'--------- El importe del recibo es Cero, el importe Académico es Cero, y el importe Administrativo es Cero. -----
'--------- Es decir que la Matrícula del estudio es gratuita, no tiene coste alguno. -----------------------------
'- ---------------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Sort(Lo_G04, BD_ImpRec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---------------------
        Call Rut_Lo_Sort(Lo_G04, BD_ImpDto, xlAscending, False)    '- Ordenar primero accelera un montón el borrado ---------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_ImpMatCero")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec. de Matrícula_Cero ", 0, _
                     Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de Matrícula_Cero", 0)
        End If
        .ShowAutoFilter = True          '- El AdvancedFilter con Rango de Criterio desactiva el "ShowFilterMarck"
        .AutoFilter.ShowAllData         ' Elimina los filtros

'- ---------------------------------------------------------------------------------------------------------------
'- Borrar Borrar Recibos - Subvencionado-, Imp_Rec =0 porque Imp_Dto >0 ------------------------------
'- ---------------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, BD_ImpRec, xlAscending, True)
        .Range.AutoFilter Field:=BD_ImpRec, Criteria1:="=0"
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            If Prog__APP.Range("SW_DelRegMatrículaCero") Then
                .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec. de Matrícula_Cero Subvencionada", 0, _
                     Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            Else
                .DataBodyRange.Columns(BD_Obs_Conta).SpecialCells(xlCellTypeVisible).Cells.Value = "ImpMatCero"
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Marcados Rec. de Matrícula_Cero Subvencionada", 0, _
                     Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            End If
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de Matrícula_Cero Subvencionada", 0)
        End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"  --------------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_G04, BD_DNI, "=1")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Ficticios, DNI=1", 0, _
                                                            Format(rowfind, " #,##0") & " reg. ", _
                                                            " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            If Lo_G04.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Ficticios, DNI=1", 0)
        End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos BD_C_Acad <> C_Acad_Ant ---------------------------------------------------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_G04, BD_C_Acad, "<>" & C_Acad_Ant)
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. de C_Acad. <> " & C_Acad_Ant, 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            '- Comprobar que a Lo_G04 le quedan datos ---------------------
            If Lo_G04.ListRows.Count = 0 Then
Proceso_Finalizado_por_quedarse_sin_Registros:
                MsgBx_Title = "Proceso: Importar LsGES04 C_Acad_Ant para ImpAdm"
                MsgBx_Msg = "¡ A la tabla Lo_G04 no le quedan Recibos procesables !"
                MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), 0)
                Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe") & vbLf & vbLf & MsgBx_Msg & Now()
                Exit Sub
            End If
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de C_Acad. <> " & C_Acad_Ant, 0)
        End If
                
'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos ANULADOS ---------------------------------------------------------------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_G04, BD_Anul, "=S")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Anulados ", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            If Lo_G04.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Anulados ", 0)
        End If
                
'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos NO Martrícula ----------------------------------------------------------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_G04, BD_Matricula, "=N")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. NO Matrícula ", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            If Lo_G04.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. NO Matrícula ", 0)
        End If
                
'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos INVALIDADOS ------------------------------------------------------------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_G04, BD_Hinvalid, "=S")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Invalidados ", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            If Lo_G04.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Invalidados ", 0)
        End If
                
        '- -----------------------------------------------------------------------------
    End With    '- Lo_G04

Debug.Print "<<< RuT_Remove_Null_Reg_CAcad"

End Sub



