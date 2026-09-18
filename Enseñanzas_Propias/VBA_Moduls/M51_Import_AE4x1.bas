Attribute VB_Name = "M51_Import_AE4x1"
' Last Rev. 2026-09-18 23:49
' ==================================================================================================================
' *** MODULO COMPLETO DESACTIVADO (comentado) el 2026-09-18 23:49 ***
'
' Motivo: el flujo AE4x4/AE4x1 no compila. Bugs pendientes del Informe_Bugs:
'   - B5/B6: 'Lo_AE4x1' no se declara en ningun sitio Y no se le hace Set en
'            ninguna rama (la rama If solo copia el rango; el Else ya lo usa).
'            Ademas las 4 llamadas de M50 pasan siempre False, asi que nunca
'            se entraria por la rama de inicializacion.
'
' Verificado antes de desactivarlo: ninguna invocacion viva por las 4 vias
' (codigo VBA, macros asignadas a shapes, tabla Tb_Tareas del menu auxiliar,
' resto del XML del libro). Reactivar exige resolver antes B5/B6.
' ==================================================================================================================
'2026-01-25
'Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Seleccionar Excel pero no lo importa, sólo lo abre -------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
'Sub Rut_Lo_Import_AE4x1(Ws_AE4x1 As Worksheet, _
'                                  Arch_New_Name As String, _
'                                  SheetNom As String, _
'                                  Optional NameFileAE4 As String, _
'                                  Optional PathFileAE4 As String, _
'                                  Optional SW_Inicilizar_Ws As Boolean = False)
                             
'Debug.Print ">>> Rut_Lo_Import_LoData_LoDefCol_AE4x1"
'    Dim Ccol            As Integer
'    Dim RowsFind        As Variant
'    Dim ArchRequest     As String:      ArchRequest = Arch_New_Name
'    Dim TxT_ProgIni     As String:      TxT_ProgIni = ActivForm.Controls("TBx_Informe")
'    Dim TxT_Progreso    As String
'    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
'    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
'    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
'    Dim FechCierreCont  As String:      FechCierreCont = Prog__APP.Range("APP_FechCierreCont")
'    Dim Ws_AE4          As Worksheet:   Set Ws_AE4 = Lo_AE4x1.Parent
'    Dim Rng_Informe     As Range:       Set Rng_Informe = Ws_AE4.Range("g5")

'    Dim SwScrUp     As Boolean:             SwScrUp = Application.ScreenUpdating:       Application.ScreenUpdating = False

'    H_Inicio = Timer                '- Para saber el tiempo de proceso
'    LastTimeLap = Timer             '- Para saber tiempos intermedios
        
    '- Select File -------------------------------------------------------------------------------------
'    Call Rut_File_Select_V2("Seleccionar el Fichero Excel " & Arch_New_Name & ": ", Arch_New_Name, "Excel", "*.xlsm")
'    Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NameFileAE4, PathFileAE4)
'    If Arch_New_Name = "Cancel" Then
'        MsgBx_Title = "Proceso: Importar " & ArchRequest
'        MsgBx_Msg = "¡ Cancelado a petición del Usuario !    "
'        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
'        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        '- Visualizo el progreso  <<<<>>>>
'            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
'        Arch_New_Name = "Cancel"
'        Exit Sub
'    End If
    '- Comprueba que se ha seleccionado el nombre adecuado de Excel. --------------------------
'    If Left(NameFileAE4, Len(ArchRequest)) <> ArchRequest Then
'        MsgBx_Title = "Proceso: Importar " & ArchRequest
'        MsgBx_Msg = "¡ Cancelado ¡" & vbLf & "El fichero debe ser:" & vbLf & ArchRequest & vbLf & vbLf & _
'                    "  y se ha seleccionado:" & vbLf & NameFileAE4
'        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
'        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & vbLf & Now()
        '- Visualizo el progreso  <<<<>>>>
'            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), 0)
'        Arch_New_Name = "Cancel"
'        Exit Sub
'    End If
        '- Visualizo el progreso  <<<<>>>>
'        TxT_Progreso = ActivForm.Controls("TBx_Informe")
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Excel Seleccionado: " & NameFileAE4, LastTimeLap, , , TxT_Progreso)
    
    '- --------------------------------------------------------------------------------------------------------------
'    Prog__APP.Range("SW_WB_Deactivate") = False     '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ---->>>
'    Dim ClosedBook      As Workbook:        Set ClosedBook = Workbooks.Open(Arch_New_Name, ReadOnly:=True)
'    Dim Ws_ClsBk        As Worksheet:       Set Ws_ClsBk = ClosedBook.Sheets(SheetNom)
'    Dim Lo_ClsBk        As ListObject:      Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
    
'    Ws_ClsBk.Unprotect
'    Lo_ClsBk.ShowTotals = False
    
    '- Comprobar que la Tabla Lo_ClsBk No está vacía ---------------------
'    If Lo_ClsBk.ListRows.Count = 0 Then
'        MsgBx_Title = "Proceso: Importar " & ArchRequest
'        MsgBx_Msg = "¡ La tabla no contiene datos !    "
'        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
'        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        '- Visualizo el progreso  <<<<>>>>
'            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
'        GoTo Cancel_Rut
'    End If
            '- Visualizo el progreso  <<<<>>>>
'            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Excel Seleccionado: " & NameFileAE4, 0, , _
'                                                            Format(Lo_ClsBk.ListRows.Count, "#,##0") & " reg", TxT_Progreso)

    '- Copy ClosedBook:
'    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
'    Call Rut_WrkSheet_Preparar(Ws_AE4x1)
    '-          Si SW_Del_LoData=true Borrar Lo_AE4x1
'    If SW_Inicilizar_Ws Then       '- Es el 1º, sólo copiar Lo_ClsBk en Ws_AE4x1
'        Call Rut_WrkSheet_Vaciar(Ws_AE4x1)
'        Lo_ClsBk.Range.Copy Destination:=WsBuffer.Range("A1")
'    Else
'        Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_AE4x1, False)
    
'    End If
'    ClosedBook.Close SaveChanges:=False
'    Set ClosedBook = Nothing
'    Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<

    ' Restaurar entorno
'    Application.ScreenUpdating = SwScrUp
'Exit Sub

'Cancel_Rut:
'    Arch_New_Name = "Cancel"
'    ClosedBook.Close SaveChanges:=False
'    Set ClosedBook = Nothing
'    Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------


