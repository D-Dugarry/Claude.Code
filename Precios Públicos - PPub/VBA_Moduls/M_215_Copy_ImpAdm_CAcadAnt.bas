Attribute VB_Name = "M_215_Copy_ImpAdm_CAcadAnt"
'2026-01-23
Option Explicit


            Sub Rut_Copy_ImpAdm_CAcadAnt_a_BDatos_ByHand()
                
                Call Rut_Copy_ImpAdm_CAcadAnt_a_BDatos
                MsgBox "FIN"
'                Application.Speech.Speak "Proceso completado."
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Importar BDatos_Ant ------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Copy_ImpAdm_CAcadAnt_a_BDatos()
Debug.Print ">>> Rut_Copy_ImpAdm_CAcadAnt_a_BDatos"

    Dim CantImpAcad     As Long:        CantImpAcad = 0
    Dim CantImpAdm      As Long:        CantImpAdm = 0
    Dim CantImpDto      As Long:        CantImpDto = 0
    Dim ImpTAcad        As Currency:      ImpTAcad = 0
    Dim ImpTAdm         As Currency:      ImpTAdm = 0
    Dim ImpTDto         As Currency:      ImpTDto = 0
    Dim ImpTAdmNeg      As Currency
    Dim F_BD            As Long
    Dim F_BDAdm         As Long
    Dim Found           As Long:        Found = 0
    Dim NotFound        As Long:        NotFound = 0
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad          As String:      C_Acad = Prog__APP.Range("APP_CursAcad")
    Dim TimeLapSub      As Single:      TimeLapSub = Time
    Dim TxT_Progreso    As String
    Dim TxT_ProgIni     As String
    Dim rowfind         As Variant
    
    Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_BD_CAcad_Ant     As ListObject:      Set Lo_BD_CAcad_Ant = Sht__BD_IAdm_CAcadAnt.ListObjects(1)
    
    Sht__BD.Visible = xlSheetVisible:                   Sht__BD.Unprotect:                  Lo_BD.ShowTotals = False
    Sht__BD_IAdm_CAcadAnt.Visible = xlSheetVisible:     Sht__BD_IAdm_CAcadAnt.Unprotect:    Lo_BD_CAcad_Ant.ShowTotals = False
    
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_IAdm_CAcadAnt)
    
    Dim TF_BD           As Long:        TF_BD = Lo_BD.ListRows.Count
    Dim TF_BDAdm        As Long:        TF_BDAdm = Lo_BD_CAcad_Ant.ListRows.Count
    If TF_BD = 0 Or TF_BDAdm = 0 Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay registros que cruzar (BD o BD_IAdm_CAcadAnt vacías).", 0)
        GoTo Restablecer_Valores
    End If
    
    '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Procedimiento: Incorporar Imp. Acad. y Adm., Curso_Ant: " & C_Acad_Ant & ", a BDatos de " & AñoCont, 0, , , , , , 4)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")

    '- --------------------------------------------------------------------------------------------------------
    '- Borrar Recibos "RecNegCab", de Lo_BD_CAcad_Ant (Sht__BD_IAdm_CAcadAnt) con ImpAdm <= 0 --------------------------------
    With Lo_BD_CAcad_Ant
        Call Rut_Lo_Sort(Lo_BD_CAcad_Ant, BD_C_Acad, xlAscending, True)                     '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_Obs_Conta, Criteria1:="RecNegCab"                       '- Filtro los Recibos con ImpAdm <= 0
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            ImpTAdmNeg = Application.Sum(.DataBodyRange.Columns(BD_Rec_Imp_Adm).SpecialCells(xlCellTypeVisible))
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. 'RecNegCab' Imp.Adm. <0 ", 0, _
                                            Format(rowfind, "#,##0") & " reg.", Format(ImpTAdmNeg, "#,##0.00€"), TxT_ProgIni, , , 2)
        End If
    End With
    
    '- --------------------------------------------------------------------------------------------------------
    '- Inicializo los Recibos de BDatos de C_Acad_Ant y Borro Cols Imp.Adm. Acad. y Dto. de los Recibos de BD de C_Acad_Ant --------------------------------
    With Lo_BD
        Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)                           '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & C_Acad_Ant                     '- Filtro los Recibos del Curso-Acad-Ant
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Rec_Imp_Acad).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_Rec_Imp_Adm).SpecialCells(xlCellTypeVisible).ClearContents
            .DataBodyRange.Columns(BD_Rec_Imp_Dto).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        .AutoFilter.ShowAllData            ' Elimina los filtros
    End With
    
   '- -----------------------------------------------------------------------------------------------------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_BD_CAcad_Ant)
    Lo_BD_CAcad_Ant.DataBodyRange.Columns(BD_Incidencias).ClearContents
    With Lo_BD_CAcad_Ant.DataBodyRange
    
        Call Rut_Lo_Sort(Lo_BD_CAcad_Ant, BD_Ref, xlAscending, True)   '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)          '- Ordenar primero accelera un montón el borrado -----
        F_BD = 1
        For F_BDAdm = 1 To TF_BDAdm
            Select Case Lo_BD.DataBodyRange.Cells(F_BD, BD_Ref)
                Case Is < .Cells(F_BDAdm, BD_Ref) '- Ref_BD  NO-EXISTE-EN  Sht__BD_Adm
                    If F_BD < TF_BD Then F_BD = F_BD + 1 Else Exit For
                    F_BDAdm = F_BDAdm - 1
                Case Is = .Cells(F_BDAdm, BD_Ref)  '- Ref_BD  SÍ-EXISTE-EN  Sht__BD_Adm
                    Lo_BD.DataBodyRange.Cells(F_BD, BD_Rec_Imp_Acad) = .Cells(F_BDAdm, BD_ImpAcad)
                    Lo_BD.DataBodyRange.Cells(F_BD, BD_Rec_Imp_Adm) = .Cells(F_BDAdm, BD_ImpAdm)
                    Lo_BD.DataBodyRange.Cells(F_BD, BD_Rec_Imp_Dto) = .Cells(F_BDAdm, BD_ImpDto)
                    Found = Found + 1
                    .Cells(F_BDAdm, BD_Incidencias) = "Found en BD."
                    If F_BD < TF_BD Then F_BD = F_BD + 1 Else Exit For
                Case Is > .Cells(F_BDAdm, BD_Ref)  '- Ref_BD_Adm  NO-EXISTE-EN  Sht__BD
                    NotFound = NotFound + 1
                    .Cells(F_BDAdm, BD_Incidencias) = "Not Found en BD."
            End Select
        
            If (F_BDAdm Mod 5000 = 0 And F_BDAdm <> 0) Or F_BD Mod 5000 = 0 Then
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporados datos de: " & _
                            Format(Found, "#,##0") & " reg. de " & Format(F_BDAdm, "#,##0") & "/" & Format(TF_BDAdm, "#,##0") & "reg.,  entre " & _
                            Format(F_BD, "#,##0") & "/" & Format(TF_BD, "#,##0") & "reg.", 0, , , TxT_Progreso, True, , 2)
            End If
        Next F_BDAdm
        '- Visualizo el progreso ----------------------------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporado Imp.Adm. de C_Acad_Ant_" & C_Acad_Ant & _
                                                       ", Cobrado en " & AñoCont & " a BDatos.", 0, , , TxT_ProgIni, , , 2)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporados datos a: " & _
                            Format(Found, "#,##0") & " reg. de BDatos, de un total de " & Format(TF_BDAdm, "#,##0") & "reg. de " & Sht__BD_IAdm_CAcadAnt.Name, 0)
    End With        '-  Lo_BD_CAcad_Ant.DataBodyRange
    
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    With Lo_BD
        .AutoFilter.ShowAllData            ' Elimina los filtros
        
        '- Sumatorios ---------
        With .DataBodyRange
            ImpTAcad = Application.Sum(.Columns(BD_Rec_Imp_Acad))
            ImpTAdm = Application.Sum(.Columns(BD_Rec_Imp_Adm))
            ImpTDto = Application.Sum(.Columns(BD_Rec_Imp_Dto))
            CantImpAcad = Application.Count(.Columns(BD_Rec_Imp_Acad))
            CantImpAdm = Application.Count(.Columns(BD_Rec_Imp_Adm))
            CantImpDto = Application.Count(.Columns(BD_Rec_Imp_Dto))
        End With
    End With
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Tasas Acad. de Matrícula por un importe de:    ", 0, _
                                        Format(ImpTAcad, "#,##0.00€"), "en " & Format(CantImpAcad, "#,##0 reg."))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Descuentos  de Matrícula por un importe de:    ", 0, _
                                        Format(ImpTDto, "#,##0.00€"), "en " & Format(CantImpDto, "#,##0 reg."))
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Tasas Acad. de Matrícula por un importe de:    ", 0, _
                                        Format(ImpTAdm, "#,##0.00€"), "en " & Format(CantImpAdm, "#,##0 reg."))

Restablecer_Valores:
    Lo_BD.ShowTotals = True
    Lo_BD_CAcad_Ant.ShowTotals = True
    Call Rut_Lo_Filtros_Quitar(Lo_BD_CAcad_Ant)
    Application.Speech.Speak "Proceso completado."
End Sub     ' Rut_Añadir_Campos_DR_Anteriores     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================



