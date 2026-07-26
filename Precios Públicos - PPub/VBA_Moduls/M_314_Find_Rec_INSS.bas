Attribute VB_Name = "M_314_Find_Rec_INSS"
'Rev.: 2026-01-22
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Identificar de un C_Acad, los 1º Rec. con Seguro obligatorio INSS ----------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Find_Rec_INSS_C_Acad(Lo_ClsBk As ListObject, _
                             NomFichLSace06 As String, _
                             RutaFichLsace06 As String)
                                  
Debug.Print ">>> RuT_Find_Rec_INSS_C_Acad" & Prog__APP.Range("APP_CursAcad")
    Dim ImpINSS         As Currency
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    Dim Fila            As Long
    Dim T_Filas         As Long:        T_Filas = Lo_ClsBk.ListRows.Count
    Dim RowsDel         As Long
    Dim CantRecINSS     As Long
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim PlanDNI_Ant     As String:      PlanDNI_Ant = ""
    Dim PlanDNI_New     As String:
    Dim TxT_ProgIni     As String:      TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Dim TxT_Progreso    As String
    Dim FichName        As String
    Dim RowData         As ListRow
    Dim rowfind         As Variant
    Dim Sh_ClsBk        As Worksheet:   Set Sh_ClsBk = Lo_ClsBk.Parent

    '- Visualizo el progreso
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Proceso: Identificar Recibos INSS: " & _
                         C_Acad_Ant & ", en ", 0, , Format(T_Filas, "#,##0") & " reg.", , , , 2)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
    Lo_ClsBk.DataBodyRange.Columns(LS06_RecINSS).ClearContents
    ' Ordenar por columnas  ------------------------------
    Call Rut_Lo_Sort(Lo_ClsBk, LS06_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, LS06_ActivEco, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, LS06_Plan, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, LS06_DNI, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, LS06_NumRec, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_ClsBk, LS06_Ref, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----

    '- Recorro toda la tabla Lo_ClsBk -----------------------------------------------------------------
    For Fila = 1 To T_Filas   '--- Bucle para recorrer todas la filas de la Consulta Prog_TitPH
        Set RowData = Lo_ClsBk.ListRows(Fila)
        PlanDNI_New = RowData.Range(LS06_Plan) & "_" & RowData.Range(LS06_DNI)
        If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
                                            RowData.Range(LS06_RecINSS) = "RecCab"    '- para quitar marca de borrar Rec.
                                            PlanDNI_Ant = PlanDNI_New
                                            CantRecINSS = CantRecINSS + 1
        End If
        If Fila Mod 2000 = 0 Then
            '- Visualizo el progreso
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Identificados " & Format(CantRecINSS, "#,##0") & " Recibos INSS, entre: ", 0, _
                                Format(Fila, "#,##0") & " rec.", "de " & Format(T_Filas, "#,##0"), TxT_Progreso, True, , 2)
        End If
    Next
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Identificados " & Format(CantRecINSS, "#,##0") & " Recibos INSS, entre ", 0, Format(T_Filas, "#,##0"), , TxT_ProgIni)
    
    '- Borrar Recibos NO Cabecera. -----------------------------------------------
    With Lo_ClsBk
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk, LS06_RecINSS, "=")
        RowsDel = rowfind - .ListRows.Count
        If RowsDel > 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Del de LSace06 Rec. SIN INSS: ", 0, _
                                                            Format(RowsDel, "#,##0") & " reg", _
                                                            "quedan " & Format(.ListRows.Count, "#,##0") & " reg", TxT_ProgIni)
        Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "No hay Rec. Con Seguro INSS. ", 0, , , TxT_ProgIni)
        End If
    End With

    '- Exporto??? Sólo los Recibos INSS en un Excel, para otros Procesos !!! --------------------------------------------------
    Dim GuardarReg As VbMsgBoxResult
    GuardarReg = MsgBox("¿ Quieres Exportar los Recibos INSS del Curso_Acad en un excel-INSS ?" & vbLf & vbLf & _
                        "¡ Para otros procesos de estadísticas de recibos INSS para Felipe !", _
                        vbYesNo + vbQuestion + vbDefaultButton2, "Proceso: Exportar una Hoja a Excel.")
    If GuardarReg = vbYes Then
            FichName = Left(NomFichLSace06, InStrRev(NomFichLSace06, ".") - 1) & "_INSS"
            Call Rut_Lo_Export_to_New_WB(Lo_ClsBk, FichName)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Exportados los Recibos INSS, en :" & FichName, 0)
    Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No se han exportados los Recibos INSS.", 0)
    End If

    '- Borrar Recibos de C_Acad_Ant y Cobrados en Año_Cont_Ant. -----------------
    With Lo_ClsBk
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk, LS06_ACont_Cob, "=" & AñoCont - 1, _
                                                LS06_C_Acad, "=" & C_Acad_Ant)
        RowsDel = rowfind - .ListRows.Count
        If RowsDel > 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Del de LSace06 Rec.INSS F_Cob=" & AñoCont - 1, 0, _
                                                            Format(RowsDel, "#,##0") & " reg", _
                                                            "quedan " & Format(.ListRows.Count, "#,##0") & " reg")
        Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "No hay Rec. Rec.INSS F_Cob=" & AñoCont - 1, 0)
        End If
    End With
        
    '- Sumatorios  ----------------------------------------------------------------------
    With Lo_ClsBk
        ImpINSS = Application.Sum(.DataBodyRange.Columns(LS06_Concept_Imp))
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Identificados " & Format(.ListRows.Count, "#,##0") & _
                                                       " Recibos INSS por un importe de: ", TimeLapSub, , Format(ImpINSS, "#,##0.00€"))
    End With


Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
Call Rut_WrkSheet_LstObj_LiberarEspacio(Sh_ClsBk)
Debug.Print "<<< RuT_Find_Rec_INSS_C_Acad" & C_Acad_Pos
End Sub     ' -------------------------------------------------------------------------------------------------------------------------<<<
' ========================================================================================================================================

