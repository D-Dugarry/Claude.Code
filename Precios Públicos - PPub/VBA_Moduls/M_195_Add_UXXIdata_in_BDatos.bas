Attribute VB_Name = "M_195_Add_UXXIdata_in_BDatos"
'- M_195_Import_a_BD_Ant -----------------------------------------------------------------------------------------------------------
Option Explicit

'==================================================================================================================================
Sub RuT_Add_UXXIdata_in_BDatos()  '- Importar LSGES04_GE  Anterior.
'==================================================================================================================================
Debug.Print "------------------------- >>> RuT_Add_UXXIdata_in_BDatos()"
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    Dim TxT_Progreso    As String
    Dim TxT_ProgIni     As String
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad          As String:      C_Acad = Prog__APP.Range("APP_CursAcad")
    
    '- Setting ListObjects ------------------------------------
    Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_BD_Ant           As ListObject:      Set Lo_BD_Ant = Sht__BD_Ant.ListObjects(1)
    Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
    Dim Lo_DefCol_BD_Ant    As ListObject:      Set Lo_DefCol_BD_Ant = Prog_DefCol_BD_Ant.ListObjects(1)
    
    '- Setting Sheets ------------------------------------
    Sht__BD.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Lo_DefCol_BD.TotalsRowRange(DefC_HiddenCol) = False
    Prog__APP.Range("SW_Col_Hide_Sht__BD") = False
    Lo_BD.ShowTotals = False
    
    Sht__BD_Ant.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_Ant)
    Lo_DefCol_BD_Ant.TotalsRowRange(DefC_HiddenCol) = False
    Prog__APP.Range("SW_Col_Hide_Sht__BD_Ant") = False
    Lo_BD_Ant.ShowTotals = False
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    Rut_Off_Functions
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
    
    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Importar, si necesario, BD_Ant e Incorporar datos UXXI a BDatos", 0)
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    TxT_Progreso = ActivForm.Controls("TBx_Informe")

If Not Func_MsgBox_vbYesNo("¿ Importamos el Fichero ?") Then GoTo Rut_Copy_DatosUXXI_a_BDatos
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import LSGES04 por Curso_Acad_Ant ------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim Arch_New_Name         As String ':    Arch_New_Name = "LSGES04_GE_SinDtos_Curso_" & C_Acad_Ant
    Call Rut_Lo_Import_LoData_LoDefCol(Lo_BD_Ant, Lo_DefCol_BD_Ant, DefC_TitColLstObj, Arch_New_Name)
        If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    Prog__APP.Range("APP_Last_Import_CAcad") = Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD_Ant.Range("d2") = "Úlitma Importación: " & Format(Now(), "dd-mmm-yy hh:mm")
    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importado BD_Ant: " & Arch_New_Name, TimeLapSub, , , TxT_ProgIni, True)
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
Rut_Copy_DatosUXXI_a_BDatos:
If Not Func_MsgBox_vbYesNo("¿ Trasladar Datos Contables UXXI de BD_Ant a BDatos ?") Then GoTo Terminar
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Trasladar Datos Contables UXXI de BD_Ant a BDatos --------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim F_BD            As Long
    Dim F_BDant         As Long
    Dim TF_BD           As Long:        TF_BD = Lo_BD.ListRows.Count
    Dim TF_BDant        As Long:        TF_BDant = Lo_BD_Ant.ListRows.Count
    If TF_BD = 0 Or TF_BDant = 0 Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay registros que cruzar (BD o BD_Ant vacías).", 0)
        GoTo Terminar
    End If
    Dim Found           As Long:        Found = 0
    Dim NotFound        As Long:        NotFound = 0
    Dim NoData          As Long:        NoData = 0
    Dim RngOrigen       As Range
    Dim RngDestino      As Range
    
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)          '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_BD_Ant, BD_Ref, xlAscending, True)   '- Ordenar primero accelera un montón el borrado -----
                
    With Lo_BD_Ant.DataBodyRange
    
        F_BD = 1
        For F_BDant = 1 To TF_BDAnt
            Select Case Lo_BD.DataBodyRange.Cells(F_BD, BD_Ref)
                Case Is < .Cells(F_BDant, BD_Ref) '- Ref_BD  NO-EXISTE-EN  Sht__BD_Ant
                    If F_BD < TF_BD Then F_BD = F_BD + 1 Else Exit For
                    F_BDant = F_BDant - 1
                    NoData = NoData + 1
                Case Is = .Cells(F_BDant, BD_Ref)  '- Ref_BD  SÍ-EXISTE-EN  Sht__BD_Ant
                    Set RngOrigen = Lo_BD_Ant.ListColumns(BD_Coef_VRI).DataBodyRange.Rows(F_BD).Resize(, BD_JI_443_Adm - BD_Coef_VRI + 1)
                    Set RngDestino = Lo_BD.ListColumns(BD_Coef_VRI).DataBodyRange.Rows(F_BD).Resize(, BD_JI_443_Adm - BD_Coef_VRI + 1)
'                    Set RngOrigen = .Range(.Cells(F_BD, BD_Coef_VRI), .Cells(F_BD, BD_JI_443_Adm))
'                    Set RngDestino = Lo_BD.DataBodyRange.Range(Lo_BD.DataBodyRange.Cells(F_BD, BD_Coef_VRI), Lo_BD.DataBodyRange.Cells(F_BD, BD_JI_443_Adm))
                    ' Copia los valores (sin formato)
                    RngDestino.Value = RngOrigen.Value
                    Found = Found + 1
                    If F_BD < TF_BD Then F_BD = F_BD + 1 Else Exit For
                Case Is > .Cells(F_BDant, BD_Ref)  '- Ref_BD_Ant  NO-EXISTE-EN  Sht__BD
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
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Incorporados Datos UXXI en BDatos: ", 0, , , TxT_ProgIni)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporados datos de: " & _
                            Format(Found, "#,##0") & " reg. de un total de " & Format(F_BDant, "#,##0") & "reg.", TimeLapSub)
    '- Visualizo el progreso ----------------------------------------------------------------------------------------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "  Reg. Sin Añadir datos", 0, _
                        " NotFound: " & Format(NotFound, "#,##0") & " REG.")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "  Reg. Añadidos datos", 0, _
                        " Found: " & Format(Found, "#,##0") & " REG.")
'''    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "  Reg. Añadidos datos del curso " & C_Acad_Ant & ", emitidos en " & _
                        AñoCont & Right(String(10, " .") & " Found: " & Format(Found, "#,##0") & " REG.", 30), 0)
    End With        '-  Lo_BD.DataBodyRange
    
    Lo_BD.ShowTotals = True
    Lo_BD_Ant.ShowTotals = True

Terminar:
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe")
    Prog__APP.Range("APP_Last_Import") = Format(Now(), "dd-mmm-yy hh:mm")
    Sht__BD_Ant.Range("d1") = "Traspasados Datos UXXI a BDatos: " & Format(Now(), "dd-mmm-yy hh:mm")
    Lo_BD_Ant.ShowTotals = True
    Sht__BD.Calculate

    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BDxxx.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================

Rut_On_Functions
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set ActivForm = Nothing
Debug.Print "------------------------- <<< Sub RuT_Update_LSGES04_IAdm_CAcadAnt()"
End Sub     ' RuT_Update_LSGES04_IAdm_CAcadAnt   --------------------------------------------------------------------------------------------
'===================================================================================================================================






