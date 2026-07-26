Attribute VB_Name = "M_118_Actualiz_BD_con_BD_Ant24"
'Rev.: 2026-01-22
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Importar BDatos_Ant ------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Actualizar_LoBDatos_con_LoBD_Ant()
Debug.Print ">>> Rut_Actualizar_LoBDatos_con_LoBD_Ant"

    Dim Found           As Long:        Found = 0
    Dim NotFound        As Long:        NotFound = 0
    Dim TxT_Progreso    As String
    Dim TxT_ProgIni     As String
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    
    Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_BDant        As ListObject:      Set Lo_BDant = Sht__BD_Ant.ListObjects(1)
    Dim F_BD            As Long
    Dim F_BDant         As Long
    Dim TF_BD           As Long:            TF_BD = Lo_BD.ListRows.Count
    Dim TF_BDant        As Long:            TF_BDant = Lo_BDant.ListRows.Count
    If TF_BD = 0 Or TF_BDant = 0 Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay registros que cruzar (BD o BD_Ant vacías).", 0)
        GoTo Restablecer_Valores
    End If
    Sht__BD.Visible = xlSheetVisible:       Sht__BD.Unprotect:          Lo_BD.ShowTotals = False
    Sht__BD_Ant.Visible = xlSheetVisible: Sht__BD_Ant.Unprotect:    Lo_BDant.ShowTotals = False
    
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    Call Rut_Lo_Filtros_Quitar(Lo_BDant)
    
    '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Procedimiento: Incorporar Datos de BDatos_Ant.", 0, , , , , , 4)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")

   '- -----------------------------------------------------------------------------------------------------------------------
'    Lo_BD.DataBodyRange.Columns(BD_Incidencias).ClearContents
    Lo_BD.DataBodyRange.Columns(BD_Obs_Conta).ClearContents
    Lo_BD.DataBodyRange.Columns(BD_JI_443_Adm).ClearContents
    Lo_BD.DataBodyRange.Columns(BD_JI_Emi_Acad).ClearContents
    Lo_BD.DataBodyRange.Columns(BD_AD_0010).ClearContents
    
    '- Todos los Rec. "Emitidos" del AñoCont_Anterior pasan a "EjeAnt" en el AñoCont corriente
    '- Todos los Rec. "ADxAplz" del AñoCont Anterior pasan a "Aplazado" en el AñoCont corriente
    Call Rut_Lo_Sort(Lo_BDant, BD_Tipo_Rec, xlAscending, True)   '- Ordenar primero accelera un montón el borrado -----
    With Lo_BDant.ListColumns(BD_Tipo_Rec).DataBodyRange
        .Replace What:="Emitido", Replacement:="EjeAnt", _
                 LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False
        .Replace What:="ADxAplz", Replacement:="Aplazado", _
                 LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False
    End With
    
    With Lo_BDant.DataBodyRange
    
        Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)          '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_BDant, BD_Ref, xlAscending, True)   '- Ordenar primero accelera un montón el borrado -----
        F_BD = 1
        For F_BDant = 1 To TF_BDant
            Select Case Lo_BD.DataBodyRange.Cells(F_BD, BD_Ref)
                Case Is < .Cells(F_BDant, BD_Ref) '- Ref_BD  NO-EXISTE-EN  Sht__BD_Adm
                    If F_BD < TF_BD Then F_BD = F_BD + 1 Else Exit For
                    F_BDant = F_BDant - 1
                Case Is = .Cells(F_BDant, BD_Ref)  '- Ref_BD  SÍ-EXISTE-EN  Sht__BD_Adm
                    Lo_BD.DataBodyRange.Cells(F_BD, BD_Rec_Imp_Adm) = .Cells(F_BDant, BD_Rec_Imp_Adm)      '- Sólo para comparar...
                    Lo_BD.DataBodyRange.Cells(F_BD, BD_Tipo_Rec) = .Cells(F_BDant, BD_Tipo_Rec)
                    Lo_BD.DataBodyRange.Cells(F_BD, BD_JI_Emi_Acad) = .Cells(F_BDant, BD_JI_Emi_Acad)
                    Lo_BD.DataBodyRange.Cells(F_BD, BD_AD_0010) = .Cells(F_BDant, BD_AD_0010)
'                    Lo_BD.DataBodyRange.Cells(F_BD, 52) = .Cells(F_BDant, BD_AD_0010)
                    Found = Found + 1
                    If F_BD < TF_BD Then F_BD = F_BD + 1 Else Exit For
                Case Is > .Cells(F_BDant, BD_Ref)  '- Ref_BD_Adm  NO-EXISTE-EN  Sht__BD
                    NotFound = NotFound + 1
                    .Cells(F_BDant, BD_Incidencias) = "Not Found en BD."
            End Select
        
            If (F_BDant Mod 200 = 0 And F_BDant <> 0) Or F_BD Mod 500 = 0 Then
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporados datos de: " & _
                            Format(Found, "#,##0") & " reg. de " & Format(TF_BDant, "#,##0") & "reg.,  entre " & _
                            Format(F_BD, "#,##0") & "reg." & " de " & Format(TF_BD, "#,##0") & "reg.", 0, , , TxT_Progreso, True, , 2)
            End If
        Next F_BDant
        '- Visualizo el progreso ----------------------------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Procedimiento: Incorporar Datos de BDatos_Ant.", 0, , , TxT_ProgIni, , , 2)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporados datos de: " & _
                            Format(Found, "#,##0") & " reg. en Bdatos, de un total de " & Format(TF_BDant, "#,##0") & "reg. en BD_Ant", TimeLapSub)
    End With        '-  Lo_BDant.DataBodyRange
    
Lo_BD.ShowTotals = True
Lo_BDant.ShowTotals = True
Restablecer_Valores:
    Application.Speech.Speak "Proceso completado."
End Sub     ' Rut_Añadir_Campos_DR_Anteriores     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
