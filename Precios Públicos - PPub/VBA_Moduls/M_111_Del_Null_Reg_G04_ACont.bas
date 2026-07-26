Attribute VB_Name = "M_111_Del_Null_Reg_G04_ACont"
'Rev.: 2026-01-22
Option Explicit

'    - M_111, Filtrar y Borrar Registros NO deseados:
'        - Borrar Recibos AE4 Enseñanzas Propias
'        - Borrar Recibos con Imp.Rec. < 0
'        - Borrar Recibos de Matrículas de coste CERO ==> Imp_Rec = Imp_Dto = 0 ó Vacío (Tb_CriT_ImpMatCero)
'        - Borrar Recibos Importe CERO - Subvencionado- Imp_Rec =0 porque Imp_Dto >0
'        - Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"
'        - Borrar Recibos con AE=300 (M013 mal matriculado en Gestión Académica) ==>> Seminario orientación para preparación pruebas para mayores de 25 años.
'        - Borrar Recibos ANULADOS
'        - Borrar Recibos NO Martrícula
'        - Borrar Recibos INVALIDADOS
'        - Filtra Recibos con F_Cob > APP_FechCierreCont y Limpia-Clear las Columnas BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag y BD_HTipCob
'        - Borrar Recibos con F_Emi > APP_FechCierreCont "FUERA DEL PERÍODO CONTABLE"
'        - Borrar Incongruencias de Fechas, Recibos Cobrados en Años Anteriores o Posteriores a AñoCont

            Sub RuT_Remove_Null_Reg_ByHand()
                
                Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
                Dim Lo_BD_Ant           As ListObject:      Set Lo_BD_Ant = Sht__BD_Ant.ListObjects(1)
                Dim Lo_BD_Dpl           As ListObject:      Set Lo_BD_Dpl = Sht__BD_Dupl.ListObjects(1)
                Dim Lo_BD_ErrDate       As ListObject:      Set Lo_BD_ErrDate = Sht__BD_ErrDate.ListObjects(1)
                Dim Lo_BD_RegAnul       As ListObject:      Set Lo_BD_RegAnul = Sht__BD_RegAnul.ListObjects(1)
                Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
                
                Dim Lo_BdM013           As ListObject:      Set Lo_BdM013 = Sht__BD_M013.ListObjects(1)
                Dim Lo_BdPNB1           As ListObject:      Set Lo_BdPNB1 = Sht__BD_PNB1.ListObjects(1)
                Dim Lo_BdAdmP           As ListObject:      Set Lo_BdAdmP = Sht__BD_AdmP.ListObjects(1)
                
                Call RuT_Remove_Reg_No_Valid(Lo_BD, Lo_BD_ErrDate)
                
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Filtrar y Borrar Registros NO deseados: ------------------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Remove_Reg_No_Valid(Lo_Data As ListObject, _
                        Optional Lo_BD_ErrDate As ListObject)
                        
Debug.Print ">>> RuT_Remove_Reg_No_Valid"
    Dim rowfind     As Variant
    Dim AñoCont         As String:  AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim FechCierreCont  As String:  FechCierreCont = Prog__APP.Range("APP_FechCierreCont")
    Dim FechCierreCriteria As String:  FechCierreCriteria = Format(CDate(FechCierreCont), "mm\/dd\/yyyy")   '- Criteria robusto ante config. regional (no depende de recortar posiciones de texto)
    Dim Sh_Data     As Worksheet:   Set Sh_Data = Lo_Data.Parent

    '- ----------------------------------------------------------------------------------------------------------------------------
    With Lo_Data
        .ShowTotals = False
                
'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos AE4 Enseñanzas Propias -----------------------------------------------------------------------------------
        If Prog__APP.Range("SW_DelRegAE4") Then
            Call Rut_Lo_Filtros_Quitar(Lo_Data)
            rowfind = .ListRows.Count
            Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_ActivEco, "=4")
            rowfind = rowfind - .ListRows.Count
            If rowfind > 0 Then
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. AE4 ", 0, _
                                                                Format(rowfind, " #,##0") & " reg. ", _
                                                                " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            Else
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. AE4 ", 0)
            End If
        Else
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Se ha decidido mantener los Rec. AE4", 0)
        End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos con Imp.Rec. < 0 -----------------------------------------------------------------------------------
        If Prog__APP.Range("SW_DelRegNeg") Then
            Call Rut_Lo_Filtros_Quitar(Lo_Data)
            rowfind = .ListRows.Count
            Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_ImpRec, "<0")
            rowfind = rowfind - .ListRows.Count
            If rowfind > 0 Then
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Negativos. ", 0, _
                                                                Format(rowfind, " #,##0") & " reg. ", _
                                                                " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            Else
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Negativos. ", 0)
            End If
        Else
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Se ha decidido mantener los Rec. Negativos.", 0)
        End If

'- ---------------------------------------------------------------------------------------------------------------
'- Borrar Recibos de Matrículas de Coste CERO ==> Imp_Rec = Imp_Dto = 0 ó Vacío (Tb_CriT_ImpMatCero) -------------
'--------- El importe del recibo es Cero, el importe Académico NO es Cero, o el importe Administrativo NO es Cero. -----
'--------- Es decir que la Matrícula del estudio es gratuita, aunque tiene coste. -----------------------------
'- ---------------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_ImpRec, xlAscending, True)
        Call Rut_Lo_Sort(Lo_Data, BD_ImpDto, xlAscending, False)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_ImpMatCero")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            If Prog__APP.Range("SW_DelRegMatrículaCero") Then
                .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec. de Matrícula_Cero A Coste Cero", 0, _
                     Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            Else
                .DataBodyRange.Columns(BD_Obs_Conta).SpecialCells(xlCellTypeVisible).Cells.Value = "ImpMatCero"
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Marcados Rec. de Matrícula_Cero A Coste Cero", 0, _
                     Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            End If
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de Matrícula_Cero A Coste Cero", 0)
        End If

'- ---------------------------------------------------------------------------------------------------------------
'- Borrar Borrar Recibos - Subvencionado-, Imp_Rec =0 porque Imp_Dto >0 ------------------------------
'- ---------------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_ImpRec, xlAscending, True)
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
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_DNI, "=1")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Ficticio, DNI=1", 0, _
                                                            Format(rowfind, " #,##0") & " reg. ", _
                                                            " de " & Format(.ListRows.Count, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Ficticio, DNI=1", 0)
        End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos con AE=300 (M013 mal matriculado en Gestión Académica) ==>> "SEMINARI D'ORIENTACIÓ PER A PREPARACIÓ DE PROVES PER A MAJORS DE 25 ANYS"  --------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_ActivEco, "=300")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. AE=300 que debería ser AE=4 y Plan=M013.", 0, _
                                                            Format(rowfind, " #,##0") & " reg. ", _
                                                            " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " - Son Rec. EFP, Plan=M013 'Seminario orientación pruebas > 65', mal matriculado x Secretaría de Acceso", 0)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " - Mal matrículados con Rec.Mov. con AE=300, los borro porque vendrán en el AE4x4 ¡¡rectificados a mano!!", 0)
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Ficticio, DNI=1", 0)
        End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos ANULADOS ---------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_Anul, "=S")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Anulados ", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Anulados ", 0)
        End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos NO Martrícula ----------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_Matricula, "=N")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. NO Matrícula ", 0, _
                                                            Format(rowfind, " #,##0") & " reg. ", _
                                                            " de " & Format(.ListRows.Count, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. NO Matrícula ", 0)
        End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos INVALIDADOS ------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_Hinvalid, "=S")
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Invalidados ", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Invalidados ", 0)
        End If
                
'- ------------------------------------------------------------------------------------------------------------------
'-Filtra Recibos con F_Cob > APP_FechCierreCont y Limpia-Clear las Columnas BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag y BD_HTipCob -------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AutoFilter Field:=BD_FCob, Criteria1:=">" & FechCierreCriteria  '- ¡¡¡ EL FORMATO DEBE SER MM/DD/YYYY !!!)
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_FCob).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_ImpCob).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_FormPag).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_CtaPag).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_HTipCob).SpecialCells(xlCellTypeVisible).ClearContents
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clear Data Rec. F_Cob > " & Prog__APP.Range("APP_FechCierreCont"), 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. F_Cob > " & Prog__APP.Range("APP_FechCierreCont"), 0)
        End If
        
'- ------------------------------------------------------------------------------------------------------------------
'- ----------- Borrar Recibos con fechas FUERA DEL PERÍODO CONTABLE -------------------------------------------------
'- ------------------------------------------------------------------------------------------------------------------
'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos Emitidos en Años Posteriores a AñoCont ---------------------------------------------------------
'-Filtra Recibos con F_Emi > APP_FechCierreCont y los Borra  ----------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_FEmi, ">" & FechCierreCriteria)  '- ¡¡¡ EL FORMATO DEBE SER MM/DD/YYYY !!!
        rowfind = rowfind - .ListRows.Count
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. F_Emi > " & FechCierreCont, 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. F_Emi > " & FechCierreCont, 0)
        End If

'- ---------------------------------------------------------------------------------------------------------------
'- Borrar Incongruencias de Fechas -------------------------------------------------------------------------------
'- ---------------------------------------------------------------------------------------------------------------
'- Borrar Recibos Cobrados en Años Anteriores o Posteriores a AñoCont --------------------------------------------
'- ---------------------------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_Err")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Obs_Conta).SpecialCells(xlCellTypeVisible).Cells.Value = "Tb_CriT_Err_Date"
'            ' Preguntar si se desea guardar los registros considerados erróneos y borrados, antes de borralos. ------------------
'            Dim GuardarReg As VbMsgBoxResult
'            GuardarReg = MsgBox("¿Quieres guardar los registros con fechas incongruentes, antes de borralos?", vbYesNo + vbQuestion, "Hemos encontrado Reg. que consideramos Erróneos.")
'            If GuardarReg = vbYes Then
'                Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Data, Lo_BD_ErrDate, True)
'            End If
            Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Data, Lo_BD_ErrDate, True)
            If Prog__APP.Range("SW_DelRegErrDate") Then
                .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec. con Fechas Incongruentes. ", 0, _
                     Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            Else
                '- Visualizo el progreso --------
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Marcados Rec. con Fechas Incongruentes. ", 0, _
                     Format(rowfind, " #,##0") & " reg. ", " de " & Format(.ListRows.Count, "#,##0") & " reg.")
            End If
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. con Fechas Incongruentes.", 0)
        End If

        '- -----------------------------------------------------------------------------
    End With    '- Lo_Data

Debug.Print "<<< RuT_Remove_Null_Reg"
    Lo_Data.ShowTotals = True
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sh_Data)
End Sub     ' RuT_Remove_Reg_No_Valid

