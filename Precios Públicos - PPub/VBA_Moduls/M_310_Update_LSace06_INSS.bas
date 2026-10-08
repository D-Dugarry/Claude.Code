Attribute VB_Name = "M_310_Update_LSace06_INSS"
' Last Rev. 2026-10-08 13:46
'Rev.: 2026-01-22
'- M_310_Update_LSace06_INSS -----------------------------------------------------------------------------------------------------------

Option Explicit

'- (2026-10-04) Call_RuT_Update_LSace06_CAcad_ImpAdm_INSS eliminado: el boton del Ribbon abre el formulario de progreso
'  (Rut_Progreso_Abrir/Cerrar, M_90_Rutinas_Menu_Aux) y llama a la rutina directamente.

'==================================================================================================================================
Sub RuT_Update_LSace06_CAcad_ImpAdm_INSS()  '- Importar LSace06 por Curso_Acad, para Identificar los recibos con Tasa Adm. del seguro obligatorio del INSS
'==================================================================================================================================
Debug.Print "------------------------- >>> RuT_Update_LSace06_CAcad_ImpAdm_INSS()"
    Dim TimeLap2        As Single
    Dim NomFichLSace06  As String
    Dim RutaFichLsace06 As String
    Dim TxT_Progreso    As String
    Dim Arch_New_Name   As String
    Dim AnoCont         As String:      AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    
    Set ActivForm = Form_Running_Rut                      '- lo abre el boton del Ribbon (Rut_Progreso_Abrir, M_90)
    Application.ScreenUpdating = False
    
    '- Setting ListObjects ------------------------------------
    Dim Lo_INSS             As ListObject:      Set Lo_INSS = Sht__BD_INSS.ListObjects(1)
    Dim Lo_DefCol_LSace06   As ListObject:      Set Lo_DefCol_LSace06 = Prog_DefCol_LSace06.ListObjects(1)
    
    '- Setting Sheets ------------------------------------
    Sht__BD_INSS.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_INSS)
    Sht__BD_INSS.Unprotect
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
If Not Func_MsgBox_vbYesNo("¿ Importamos LSace06 Del Curso_Acad " & C_Acad_Ant & " o " & C_Acad_Pos & " ?" & vbLf & vbLf & _
                           "¡¡¡ O sólo copiamos los datos de BD_INSS a BDatos.  !!!") Then GoTo Rut_Copy_ImpINSS_en_BDatos
        
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Importar LSace06 Del Curso_Acad " & C_Acad_Ant & " o " & C_Acad_Pos & vbLf & vbLf & _
                    "- Importar Recibos con concepto Eco. Adm. del seguro obligatorio del INSS," & vbLf & _
                    "- Y añadir el dato a la tabla BDatos." & vbLf & _
                    Sht__BD_INSS.Range("c2") & vbLf & Sht__BD_INSS.Range("c3") & vbLf, 0)

    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import LSace06 del Curso_Acad_Ant ------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Constancia de la existencia de los DOS ficheros antes de tocar nada. Si falta alguno (y no se elige a mano), se pregunta:
    '- abortar sin tocar BD_INSS, o vaciarla y quedarse con los datos de un solo curso.
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim Ruta_Ant        As String:      Ruta_Ant = Fnc_LSace06_Elegir_Fichero(C_Acad_Ant)
    Dim Ruta_Pos        As String:      Ruta_Pos = Fnc_LSace06_Elegir_Fichero(C_Acad_Pos)
    If Ruta_Ant = "Cancel" And Ruta_Pos = "Cancel" Then
        MsgBox "No hay ningún fichero LSace06 INSS que importar (ni del " & C_Acad_Ant & " ni del " & C_Acad_Pos & ")." & vbLf & vbLf & _
               "No se ha tocado BD_INSS.", vbExclamation, "Proceso: Importar LSace06"
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Abortado: no hay ningún LSace06 INSS que importar. No se ha tocado BD_INSS.", 0)
        GoTo Terminar
    ElseIf Ruta_Ant = "Cancel" Or Ruta_Pos = "Cancel" Then
        Dim C_Acad_Falta As String:     C_Acad_Falta = IIf(Ruta_Ant = "Cancel", C_Acad_Ant, C_Acad_Pos)
        Dim C_Acad_Queda As String:     C_Acad_Queda = IIf(Ruta_Ant = "Cancel", C_Acad_Pos, C_Acad_Ant)
        If MsgBox("No hay LSace06 INSS del curso " & C_Acad_Falta & "." & vbLf & vbLf & _
                  "Sí = vaciar BD_INSS y quedarnos SÓLO con los datos del curso " & C_Acad_Queda & vbLf & _
                  "No = abortar, sin tocar BD_INSS", vbYesNo + vbExclamation + vbDefaultButton2, "Proceso: Importar LSace06") = vbNo Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Abortado: falta el LSace06 INSS del curso " & C_Acad_Falta & ". No se ha tocado BD_INSS.", 0)
            GoTo Terminar
        End If
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Sin LSace06 INSS del curso " & C_Acad_Falta & ": BD_INSS se queda sólo con el curso " & C_Acad_Queda & ".", 0)
    End If
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Vaciar Lo_INSS ANTES de importar: se importan los dos cursos enteros, así que no queda nada que conservar. Se hace aquí, antes de abrir
    '- ningún fichero grande (ver el DoEvents de M_311), y el informe dice cuántos recibos de cada curso había. ---------------------------------
    Dim DiccCurso   As Object:      Set DiccCurso = CreateObject("Scripting.Dictionary")
    Dim VCurso      As Variant
    Dim KCurso      As Variant
    Dim ICurso      As Long
    Call Rut_Lo_Filtros_Quitar(Lo_INSS)
    DiccCurso.Add C_Acad_Ant, 0
    DiccCurso.Add C_Acad_Pos, 0
    If Not Lo_INSS.DataBodyRange Is Nothing Then
        VCurso = Lo_INSS.ListColumns(LS06_C_Acad).DataBodyRange.Value2
        If IsArray(VCurso) Then
            For ICurso = 1 To UBound(VCurso, 1)
                KCurso = CStr(VCurso(ICurso, 1))
                If DiccCurso.Exists(KCurso) Then DiccCurso(KCurso) = DiccCurso(KCurso) + 1 Else DiccCurso.Add KCurso, 1
            Next ICurso
        Else
            KCurso = CStr(VCurso)
            If DiccCurso.Exists(KCurso) Then DiccCurso(KCurso) = DiccCurso(KCurso) + 1 Else DiccCurso.Add KCurso, 1
        End If
        VCurso = Empty
        Lo_INSS.DataBodyRange.Delete
        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_INSS)
    End If
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Vaciada la tabla BD_INSS, antes de importar. Recibos que tenía:", 0)
    For Each KCurso In DiccCurso.Keys
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "del C_Acad " & KCurso & ": ", 0, Format(DiccCurso(KCurso), "#,##0") & " reg")
    Next KCurso
    Set DiccCurso = Nothing
    Rut_Off_Functions   '- Antes lo hacía (sin cerrarlo) la rutina de importación; el On está en Restablecer_Valores
    If Ruta_Ant <> "Cancel" Then
        Arch_New_Name = Ruta_Ant
        Call Rut_Lo_Import_LoData_LoDefCol_LSace06(Lo_INSS, Lo_DefCol_LSace06, DefC_TitColGenInf, C_Acad_Ant, Arch_New_Name)
            If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
            Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NomFichLSace06, RutaFichLsace06)
    End If

    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import LSace06 del Curso_Acad_Pos ------------------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    If Ruta_Pos <> "Cancel" Then
        Arch_New_Name = Ruta_Pos
        Call Rut_Lo_Import_LoData_LoDefCol_LSace06(Lo_INSS, Lo_DefCol_LSace06, DefC_TitColGenInf, C_Acad_Pos, Arch_New_Name)
            If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
            Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NomFichLSace06, RutaFichLsace06)
    End If

Rut_Copy_ImpINSS_en_BDatos:
If Not Func_MsgBox_vbYesNo("¿ Trasladar el ImpINSS del C_Acad_Ant y C_Acad_Pos a BDatos ?" & vbLf & vbLf & _
                           "¡¡¡ Tienen que estar todos los Recibos que deben de estar, como los AE4.  !!!") Then GoTo Terminar
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_315_Copy_INSS_a_BD, Trasladar el ImpAdm, ImpAcad y ImpDto del C_Acad_Ant a BDatos --------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call Rut_Copy_ImpINSS_en_BDatos

Terminar:
    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)

    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe").Text
    '- Sin 'Lo_INSS.ShowTotals = True' (2026-10-04, decisión del usuario): sus totales no se requieren, y volver a mostrarlos obligaba a ocultarlos en la siguiente ejecución, que es lo que falla tras M_110/M_210 (ver CLAUDE.md).
    Sht__BD_INSS.Calculate

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_INSS.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_INSS.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD_INSS.Visible = xlSheetVeryHidden
'Sht__BD_INSS.Visible = xlSheetVeryHidden
Rut_On_Functions
    Application.Calculation = Sw_Calculation
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
'    Set ActivForm = Nothing
Debug.Print "------------------------- <<< Sub RuT_Update_LSGES04_IAdm_CAcadAnt()"
End Sub     ' RuT_Update_LSGES04_IAdm_CAcadAnt   --------------------------------------------------------------------------------------------
'===================================================================================================================================






