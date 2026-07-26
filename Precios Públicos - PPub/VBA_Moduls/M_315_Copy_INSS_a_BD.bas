Attribute VB_Name = "M_315_Copy_INSS_a_BD"
'Rev.: 2026-01-22
Option Explicit


            Sub Rut_Copy_ImpINSS_en_BDatos_ByHand()
                
                Call Rut_Copy_ImpAdm_CAcadAnt_a_BDatos
                MsgBox "FIN"
'                Application.Speech.Speak "Proceso completado."
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- M_315_Copy_INSS_a_BD, Trasladar el importe INSS a los recibos de BDatos ----------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Copy_ImpINSS_en_BDatos()
Debug.Print ">>> Rut_Copy_ImpINSS_en_BDatos"

    Dim CantImpINSS      As Long:        CantImpINSS = 0
    Dim ImpTINSS         As Currency:      ImpTINSS = 0
    Dim F_BD            As Long
    Dim F_BDINSS        As Long
    Dim Found           As Long:        Found = 0
    Dim NotFound        As Long:        NotFound = 0
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    Dim TxT_Progreso    As String
    Dim TxT_ProgIni     As String
    Dim rowfind         As Variant
    
    Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_Bd_INSS      As ListObject:      Set Lo_Bd_INSS = Sht__BD_INSS.ListObjects(1)
    Dim TF_BD           As Long:            TF_BD = Lo_BD.ListRows.Count
    Dim TF_BbINSS       As Long:            TF_BbINSS = Lo_Bd_INSS.ListRows.Count
    
    Sht__BD.Visible = xlSheetVisible:           Sht__BD.Activate:         Sht__BD.Unprotect:          Lo_BD.ShowTotals = False
    Sht__BD_INSS.Visible = xlSheetVisible:      Sht__BD_INSS.Activate:    Sht__BD_INSS.Unprotect:      Lo_Bd_INSS.ShowTotals = False
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    Call Rut_Lo_Filtros_Quitar(Lo_Bd_INSS)
    
    '- Visualizo el progreso
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Procedimiento: Incorporar Imp. INSS, Cursos " & C_Acad_Ant & " y " & C_Acad_Pos, 0, , , , , , 4)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")

    '- --------------------------------------------------------------------------------------------------------
    '- ClearContents en BDatos los Imp.INSS de los recibos de C_Acad_Ant y C_Acad_Pos --------------------------------
    With Lo_BD
        Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)                           '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & C_Acad_Ant, Operator:=xlOr, Criteria2:="=" & C_Acad_Pos   '- Filtro los Recibos del Curso-Acad-Ant/Pos
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Rec_Imp_INSS).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        .AutoFilter.ShowAllData            ' Elimina los filtros
    End With
    
   '- -----------------------------------------------------------------------------------------------------------------------
   '- - Actualizo BDatos con Lo_Bd_INSS ------------------------------------------------------------------------------------------------
   '- -----------------------------------------------------------------------------------------------------------------------
    Lo_BD.DataBodyRange.Columns(BD_Rec_Imp_INSS).ClearContents
'    Lo_Bd_INSS.DataBodyRange.Columns(BD_Incidencias).ClearContents
    With Lo_Bd_INSS.DataBodyRange
    
        Call Rut_Lo_Sort(Lo_Bd_INSS, LS06_Ref, xlAscending, True)   '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)          '- Ordenar primero accelera un montón el borrado -----
        F_BD = 1
        For F_BDINSS = 1 To TF_BbINSS
            Select Case Lo_BD.DataBodyRange.Cells(F_BD, BD_Ref)
                Case Is < .Cells(F_BDINSS, LS06_Ref) '- Ref_BD  NO-EXISTE-EN  Sht__BD_Adm
                    If F_BD < TF_BD Then
                        F_BD = F_BD + 1
                        F_BDINSS = F_BDINSS - 1
                    Else
                        '- BD agotada: este y todos los INSS restantes quedan sin encontrar --------
                        NotFound = NotFound + 1
                        .Cells(F_BDINSS, LS06_RecFound) = "Not Found en BD. (BD agotada)"
                    End If
                Case Is = .Cells(F_BDINSS, LS06_Ref)  '- Ref_BD  SÍ-EXISTE-EN  Sht__BD_Adm
                    Lo_BD.DataBodyRange.Cells(F_BD, BD_Rec_Imp_INSS) = .Cells(F_BDINSS, LS06_Concept_Imp)
                    Found = Found + 1
                    .Cells(F_BDINSS, LS06_RecFound) = "Found en BD."
                    If F_BD < TF_BD Then F_BD = F_BD + 1
                Case Is > .Cells(F_BDINSS, LS06_Ref)  '- Ref_BD_Adm  NO-EXISTE-EN  Sht__BD
                    NotFound = NotFound + 1
                    .Cells(F_BDINSS, LS06_RecFound) = "Not Found en BD."
            End Select
            If (F_BDINSS Mod 4000 = 0 And F_BDINSS <> 0) Or F_BD Mod 4000 = 0 Then
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporados datos de: " & Format(Found, "#,##0") & _
                            " reg. de " & Format(F_BDINSS, "#,##0") & "/" & Format(TF_BbINSS, "#,##0") & "reg.,  entre " & _
                            Format(F_BD, "#,##0") & "/" & Format(TF_BD, "#,##0") & "reg.", 0, , , TxT_Progreso, True, , 2)
            End If
        Next F_BDINSS
    End With        '-  Lo_BD_INSS.DataBodyRange
    
        
        '- Sumatorios ---------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        With Lo_BD.DataBodyRange
                ImpTINSS = Application.Sum(.Columns(BD_Rec_Imp_INSS))
                CantImpINSS = Application.Count(.Columns(BD_Rec_Imp_INSS))
        End With
        '- Visualizo el progreso ----------------------------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Procedimiento: Incorporar Imp.INSS. de los C_Acad " & C_Acad_Ant & " y " & C_Acad_Pos, 0, , , TxT_ProgIni, , , 2)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Recibos NO encontrados: ", TimeLapSub, _
                            Format(NotFound, "#,##0") & " reg.", " de " & Format(TF_BbINSS, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "Incorporados datos de: ", 0, _
                            Format(Found, "#,##0") & " reg.", " de " & Format(TF_BbINSS, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "Total Importe Seguro Obl. INSS", 0, _
                            Format(ImpTINSS, "#,##0.00€     "), " de " & Format(CantImpINSS, "#,##0") & " reg.")
Lo_BD.ShowTotals = True
Lo_Bd_INSS.ShowTotals = True
Sht__BD_INSS.Range("a1").Select

Restablecer_Valores:
    Application.Speech.Speak "Proceso completado."
End Sub     ' Rut_Copy_ImpINSS_en_BDatos     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

