Attribute VB_Name = "M_214_Find_ImpAdm_CAcadAnt"
'2026-01-25
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Identificar de un C_Acad, los 1º Rec. de c/matrícula para obt. Imp_Acad, Imp_Adm, Imp_Dto ---------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Find_Imp_AdmAcad_C_Acad(Lo_AE4x1 As ListObject)
                                  
Debug.Print ">>> RuT_Find_Imp_AdmAcad_C_Acad" & Prog__APP.Range("APP_CursAcad")
    Dim ImpTAcad        As Currency
    Dim ImpTAdm         As Currency
    Dim ImpNegTAdm      As Currency
    Dim ImpTDto         As Currency
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    Dim Fila            As Long
    Dim TF_Bdata        As Long:        TF_Bdata = Lo_AE4x1.ListRows.Count
    Dim RowsDel         As Long
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad          As String:      C_Acad = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim PlanDNI_Ant     As String:      PlanDNI_Ant = ""
    Dim PlanDNI_New     As String:
    Dim TxT_ProgIni     As String
    Dim TxT_Progreso    As String
    Dim RowData         As ListRow
    Dim rowfind         As Variant
    Dim Sh_Data         As Worksheet:   Set Sh_Data = Lo_AE4x1.Parent

    '- Visualizo el progreso ----------------------------------------------------------------------------------------
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Identificar ImpAcad ImpAdm e ImpDto del Curso: " & _
                         C_Acad_Ant & ", en " & Format(TF_Bdata, "#,##0") & " reg.", 0, , , , , , 2)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    Call Rut_Lo_Filtros_Quitar(Lo_AE4x1)
    Lo_AE4x1.DataBodyRange.Columns(BD_Rec_Imp_Adm).ClearContents
    Lo_AE4x1.DataBodyRange.Columns(BD_Rec_Imp_Acad).ClearContents
    Lo_AE4x1.DataBodyRange.Columns(BD_Rec_Imp_Dto).ClearContents
    Lo_AE4x1.DataBodyRange.Columns(BD_Obs_Conta).ClearContents
    
    ' Ordenar por columnas  ------------------------------
        Call Rut_Lo_Sort(Lo_AE4x1, BD_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_AE4x1, BD_ActivEco, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_AE4x1, BD_Plan, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_AE4x1, BD_DNI, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_AE4x1, BD_NumRec, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_AE4x1, BD_Ref, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    
    '- Recorro toda la tabla Lo_AE4x1 -----------------------------------------------------------------
    For Fila = 1 To TF_Bdata   '--- Bucle para recorrer todas la filas de la Consulta Prog_TitPH
        Set RowData = Lo_AE4x1.ListRows(Fila)
        If RowData.Range(BD_ActivEco) = 80 Then RowData.Range(BD_Obs_Conta) = "RecCab": GoTo Sig_Reg
        PlanDNI_New = RowData.Range(BD_Plan) & "_" & RowData.Range(BD_DNI)
        If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
                    If RowData.Range(BD_Anul) = "S" Then
                                            RowData.Range(BD_Obs_Conta) = "RecNegCab"    '- para quitar marca de borrar Rec.
                                    Else
                                            RowData.Range(BD_Obs_Conta) = "RecCab"    '- para quitar marca de borrar Rec.
                    End If
                                            PlanDNI_Ant = PlanDNI_New
        End If
Sig_Reg:
        If Fila Mod 10000 = 0 Then
            '- Visualizo el progreso ----------------------------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Identificados, ImpAcad, ImpAdm e ImpDto de Matrícula, en: " & _
                                Format(Fila, "#,##0") & "reg.", 0, , , TxT_Progreso, True, , 2)
        End If
    Next
    
    With Lo_AE4x1
        Call Rut_Lo_Filtros_Quitar(Lo_AE4x1)
        '- Visualizo el progreso ----------------------------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Resultado de Identificar Importes Adm. de todo el Curso " & C_Acad_Ant, 0, , _
                                                        "en " & Format(TF_Bdata, "#,##0") & "reg.", TxT_ProgIni, , , 2)
        '- Borrar Recibos Sin ImpAdm -----------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_AE4x1, BD_Obs_Conta, "=")
        RowsDel = rowfind - .ListRows.Count
        If RowsDel > 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Borrados Rec. SIN Imp.Adm. ", 0, _
                                            Format(RowsDel, " #,##0") & " reg.", _
                                            ",  quedan " & Format(TF_Bdata, "#,##0") & " reg.")
        Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. SIN Imp.Adm. ", 0, , , , , , 2)
        End If
        '- Sumatorios ---------
            With .DataBodyRange
                ImpTAcad = Application.SumIfs(.Columns(BD_ImpAcad), .Columns(BD_Obs_Conta), "=RecCab")
                ImpTAdm = Application.SumIfs(.Columns(BD_ImpAdm), .Columns(BD_Obs_Conta), "=RecCab")
                ImpTDto = Application.SumIfs(.Columns(BD_ImpDto), .Columns(BD_Obs_Conta), "RecCab")
            End With
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificas, Tasas Acad. de Matrícula por un importe de:    ", 0, , _
                                                        Format(ImpTAcad, "#,##0.00€"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificas, Descuentos  de Matrícula por un importe de:    ", 0, , _
                                                        Format(ImpTDto, "#,##0.00€"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificas, Tasas Adm.  de Matrícula por un importe de:    ", 0, , _
                                                        Format(ImpTAdm, "#,##0.00€"))
        On Error Resume Next
                ImpTAdm = Application.SumIfs(.DataBodyRange.Columns(BD_ImpAdm), .DataBodyRange.Columns(BD_ImpAdm), "<0")
                ImpNegTAdm = Application.CountIfs(.DataBodyRange.Columns(BD_ImpAdm), .DataBodyRange.Columns(BD_ImpAdm), "<0")
        On Error GoTo 0
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Sin Computar Rec. Anulados e Imp.Adm.<=0, de Matrícula por un importe de:    ", 0, _
                                                        Format(ImpTAdm, "#,##0.00€"), "en " & Format(ImpNegTAdm, "#,##0.00€"))

    
        '- Borrar Recibos Cobrados en Año_Cont_Ant -----------------------------------
        rowfind = .ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_AE4x1, BD_ACont_Cob, "=" & AñoCont - 1)
        TF_Bdata = .ListRows.Count
        RowsDel = rowfind - TF_Bdata
        If RowsDel > 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "- Borrados Rec. Cobrados en Año_Cont_Ant " & AñoCont - 1, 0, _
                                Format(RowsDel, " #,##0") & " reg.", "quedan " & Format(TF_Bdata, "#,##0") & " reg.", , , , 2)
        Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Cobrados en Año_Cont_Ant " & AñoCont - 1, 0, , , , , , 2)
        End If
        '- Visualizo el progreso ----------------------------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Resultado de Identificar Importes Adm. del Curso " & C_Acad_Ant & ", cobrado en " & AñoCont, 0, , _
                                                        "en " & Format(TF_Bdata, "#,##0") & "reg.", , , , 2)
        '- Sumatorios ---------
            With .DataBodyRange
                ImpTAcad = Application.SumIfs(.Columns(BD_ImpAcad), .Columns(BD_Obs_Conta), "=RecCab")
                ImpTAdm = Application.SumIfs(.Columns(BD_ImpAdm), .Columns(BD_Obs_Conta), "=RecCab")
                ImpTDto = Application.SumIfs(.Columns(BD_ImpDto), .Columns(BD_Obs_Conta), "RecCab")
            End With
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificas, Tasas Acad. de Matrícula por un importe de:    ", 0, , _
                                                        Format(ImpTAcad, "#,##0.00€"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificas, Descuentos  de Matrícula por un importe de:    ", 0, , _
                                                        Format(ImpTDto, "#,##0.00€"))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Identificas, Tasas Adm.  de Matrícula por un importe de:    ", 0, , _
                                                        Format(ImpTAdm, "#,##0.00€"))
        On Error Resume Next
                ImpTAdm = Application.SumIfs(.DataBodyRange.Columns(BD_ImpAdm), .DataBodyRange.Columns(BD_ImpAdm), "<0")
                ImpNegTAdm = Application.CountIfs(.DataBodyRange.Columns(BD_ImpAdm), .DataBodyRange.Columns(BD_ImpAdm), "<0")
        On Error GoTo 0
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Sin Computar Rec. Anulados e Imp.Adm.<=0, de Matrícula por un importe de:    ", 0, _
                                                        Format(ImpTAdm, "#,##0.00€"), "en " & Format(ImpNegTAdm, "#,##0.00€"))

    End With

Restablecer_Valores:
Call Rut_Lo_Filtros_Quitar(Lo_AE4x1)
Call Rut_WrkSheet_LstObj_LiberarEspacio(Sh_Data)
Debug.Print "<<< RuT_Find_Imp_AdmAcad_C_Acad" & C_Acad
End Sub     ' -------------------------------------------------------------------------------------------------------------------------<<<
' ========================================================================================================================================



