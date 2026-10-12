Attribute VB_Name = "M_120_Update_LSace06_ImpAdm"
' Last Rev. 2026-10-11 17:45
'- M_120_Update_LSace06_ImpAdm ---------------------------------------------------------------------------------------------------------

Option Explicit

'   Importar el LSace06 "IAdm" del Robot (LSace06_C_Acad_<curso>_IAdm_(<fecha>).xlsx) de los cursos C_Acad_Ant y C_Acad_Pos en la
'   hoja BD_ImpAdm (Sht__BD_ImpAdm, tabla Lo_ImpAdm). Modelo: M_310 (INSS) + M_311, con los mismos criterios:
'
'    - El fichero se localiza solo en la carpeta del Robot (APP_Ruta_Robot_LSc06_Cacad + curso + carpeta de fecha mas reciente);
'      si falla, se avisa y se ofrece la fecha anterior, el explorador o no importar ese curso (Fnc_LSace06_Elegir_Fichero, M_311,
'      con Tipo = "IAdm").
'    - Se localizan los DOS ficheros antes de tocar nada. Si falta uno, se pregunta: abortar o vaciar BD_ImpAdm y quedarse con un curso.
'    - Se vacia Lo_ImpAdm entera (informando de los recibos de cada curso) y se importan los dos cursos a continuacion.
'    - Cada fichero trae UN solo curso: solo se comprueba que sea el pedido. A diferencia del INSS NO se borran los importes
'      negativos: en IAdm el descuento (DTO_ADM) puede ser negativo y es un dato valido.
'    - Solo se copian los recibos emitidos o cobrados en el Año Contable (APP_AnoCont).
'    - NO se copian los datos a BDatos (otra fase). La col. "Rec found" queda vacia hasta entonces.

'==================================================================================================================================
Sub RuT_Update_LSace06_CAcad_ImpAdm()  '- Importar LSace06 IAdm (Imp. INSS, Imp. Adm. y Dto. Adm.) de C_Acad_Ant y C_Acad_Pos en BD_ImpAdm
'==================================================================================================================================
Debug.Print "------------------------- >>> RuT_Update_LSace06_CAcad_ImpAdm()"
    Dim Arch_New_Name   As String
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")

    Set ActivForm = Form_Running_Rut                      '- lo abre el boton del Ribbon (Rut_Progreso_Abrir, M_90)

    '- Setting ListObjects ------------------------------------
    Dim Lo_ImpAdm           As ListObject:      Set Lo_ImpAdm = Sht__BD_ImpAdm.ListObjects(1)
    Dim Lo_DefCol_ImpAdm    As ListObject:      Set Lo_DefCol_ImpAdm = Prog_DefCol_BD_ImpAdm.ListObjects(1)

    '- Setting Sheets ------------------------------------
    Sht__BD_ImpAdm.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_ImpAdm)
    Sht__BD_ImpAdm.Unprotect

    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios

    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Rut_Off_Functions   '- El On esta en Restablecer_Valores

    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Importar LSace06 IAdm de los Curso_Acad " & C_Acad_Ant & " y " & C_Acad_Pos & vbLf & vbLf & _
                "- Importar Imp. INSS, Imp. Adm. y Dto. Adm. de los Recibos, en la tabla BD_ImpAdm." & vbLf & _
                Sht__BD_ImpAdm.Range("c2") & vbLf & Sht__BD_ImpAdm.Range("c3") & vbLf, 0)

    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Constancia de la existencia de los DOS ficheros antes de tocar nada. Si falta alguno (y no se elige a mano), se pregunta:
    '- abortar sin tocar BD_ImpAdm, o vaciarla y quedarse con los datos de un solo curso.
    '- ----------------------------------------------------------------------------------------------------------------------------
    Dim Ruta_Ant        As String:      Ruta_Ant = Fnc_LSace06_Elegir_Fichero(C_Acad_Ant, "IAdm")
    Dim Ruta_Pos        As String:      Ruta_Pos = Fnc_LSace06_Elegir_Fichero(C_Acad_Pos, "IAdm")
    If Ruta_Ant = "Cancel" And Ruta_Pos = "Cancel" Then
        MsgBox "No hay ningún fichero LSace06 IAdm que importar (ni del " & C_Acad_Ant & " ni del " & C_Acad_Pos & ")." & vbLf & vbLf & _
               "No se ha tocado BD_ImpAdm.", vbExclamation, "Proceso: Importar LSace06 IAdm"
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Abortado: no hay ningún LSace06 IAdm que importar. No se ha tocado BD_ImpAdm.", 0)
        GoTo Terminar
    ElseIf Ruta_Ant = "Cancel" Or Ruta_Pos = "Cancel" Then
        Dim C_Acad_Falta As String:     C_Acad_Falta = IIf(Ruta_Ant = "Cancel", C_Acad_Ant, C_Acad_Pos)
        Dim C_Acad_Queda As String:     C_Acad_Queda = IIf(Ruta_Ant = "Cancel", C_Acad_Pos, C_Acad_Ant)
        If MsgBox("No hay LSace06 IAdm del curso " & C_Acad_Falta & "." & vbLf & vbLf & _
                  "Sí = vaciar BD_ImpAdm y quedarnos SÓLO con los datos del curso " & C_Acad_Queda & vbLf & _
                  "No = abortar, sin tocar BD_ImpAdm", vbYesNo + vbExclamation + vbDefaultButton2, "Proceso: Importar LSace06 IAdm") = vbNo Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Abortado: falta el LSace06 IAdm del curso " & C_Acad_Falta & ". No se ha tocado BD_ImpAdm.", 0)
            GoTo Terminar
        End If
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Sin LSace06 IAdm del curso " & C_Acad_Falta & ": BD_ImpAdm se queda sólo con el curso " & C_Acad_Queda & ".", 0)
    End If

    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Vaciar Lo_ImpAdm ANTES de importar: se importan los dos cursos enteros, así que no queda nada que conservar. Se hace aquí,
    '- antes de abrir ningún fichero grande, y el informe dice cuántos recibos de cada curso había. -----------------------------------
    Dim DiccCurso   As Object:      Set DiccCurso = CreateObject("Scripting.Dictionary")
    Dim VCurso      As Variant
    Dim KCurso      As Variant
    Dim ICurso      As Long
    Call Rut_Lo_Filtros_Quitar(Lo_ImpAdm)
    DiccCurso.Add C_Acad_Ant, 0
    DiccCurso.Add C_Acad_Pos, 0
    If Not Lo_ImpAdm.DataBodyRange Is Nothing Then
        VCurso = Lo_ImpAdm.ListColumns(IAdm_C_Acad).DataBodyRange.Value2
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
        Lo_ImpAdm.DataBodyRange.Delete
        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_ImpAdm)
    End If
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Vaciada la tabla BD_ImpAdm, antes de importar. Recibos que tenía:", 0)
    For Each KCurso In DiccCurso.Keys
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "del C_Acad " & KCurso & ": ", 0, Format(DiccCurso(KCurso), "#,##0") & " reg")
    Next KCurso
    Set DiccCurso = Nothing

    '- ----------------------------------------------------------------------------------------------------------------------------
    '- Import LSace06 IAdm del Curso_Acad_Ant y del Curso_Acad_Pos -----------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------
    If Ruta_Ant <> "Cancel" Then
        Arch_New_Name = Ruta_Ant
        Call Rut_Lo_Import_LSace06_ImpAdm(Lo_ImpAdm, Lo_DefCol_ImpAdm, DefC_TitColGenInf, C_Acad_Ant, Arch_New_Name)
            If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    End If
    If Ruta_Pos <> "Cancel" Then
        Arch_New_Name = Ruta_Pos
        Call Rut_Lo_Import_LSace06_ImpAdm(Lo_ImpAdm, Lo_DefCol_ImpAdm, DefC_TitColGenInf, C_Acad_Pos, Arch_New_Name)
            If Arch_New_Name = "Cancel" Then GoTo Restablecer_Valores
    End If

Rut_Copy_IAdm_a_BDatos:
If Not Func_MsgBox_vbYesNo("¿ Trasladar Imp. INSS, Imp. Adm. y Dto. Adm. del C_Acad_Ant y C_Acad_Pos a BDatos ?" & vbLf & vbLf & _
                           "¡¡¡ Tienen que estar todos los Recibos que deben de estar, como los AE4.  !!!") Then GoTo Terminar
    '- ----------------------------------------------------------------------------------------------------------------------------
    '- M_122_Copy_IAdm_a_BD, Trasladar Imp. INSS, Imp. Adm., Dto. Adm. y Tot. Dto. a BDatos (como M_310 con M_315)
    '- ----------------------------------------------------------------------------------------------------------------------------
    Call Rut_Copy_IAdm_a_BDatos

Terminar:
    '- Visualizo el progreso --------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)

    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe").Text
    Sht__BD_ImpAdm.Calculate

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Rut_On_Functions
    Application.Calculation = Sw_Calculation
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
Debug.Print "------------------------- <<< Sub RuT_Update_LSace06_CAcad_ImpAdm()"
End Sub     ' RuT_Update_LSace06_CAcad_ImpAdm   ----------------------------------------------------------------------------------------
'===================================================================================================================================

'- ----------------------------------------------------------------------------------------------------------------------------
'- Abre el LSace06 IAdm de UN curso (Arch_New_Name, ya localizado por M_120), lo formatea, comprueba que sea del curso pedido y lo
'- copia al final de Lo_ImpAdm (que llega vacía, o con el otro curso). Arch_New_Name es SALIDA: "Cancel" si se ha abortado.
'- Equivale a Rut_Lo_Import_LoData_LoDefCol_LSace06 (M_311) para INSS.
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Import_LSace06_ImpAdm(Lo_ImpAdm As ListObject, _
                                 LoDefCol As ListObject, _
                                 Col_Header As Integer, _
                                 C_Acad_Imp As String, _
                                 Arch_New_Name As String)
Debug.Print ">>> Rut_Lo_Import_LSace06_ImpAdm"
Rut_Off_Functions

    Dim ArchRequest     As String:      ArchRequest = "LSace06 IAdm del curso " & C_Acad_Imp
    Dim Text            As String:      Text = "C_Acad " & C_Acad_Imp
    Dim NomFichLSace06  As String
    Dim RutaFichLsace06 As String
    Dim Ccol            As Integer
    Dim rowfind         As Variant
    Dim TxT_Progreso    As String
    Dim AnoCont         As String:      AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim Wh_ImpAdm       As Worksheet:   Set Wh_ImpAdm = Lo_ImpAdm.Parent

    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios

    Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NomFichLSace06, RutaFichLsace06)
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Excel Seleccionado: " & NomFichLSace06, 0, , , TxT_Progreso)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Ruta: " & RutaFichLsace06 & vbLf, 0)
        TxT_Progreso = ActivForm.Controls("TBx_Informe")

    '- --------------------------------------------------------------------------------------------------------------
    '- Compruebo ClsBk:
    '-            1º Si no existe ListObject lo creo
    '-            2º Comprobar que la Tabla Lo_ClsBk_LSace06 tiene datos
    '-            3º Comprobar que la cabecera de la Tabla Lo_ClsBk_LSace06 corresponde con la establecida en la LoDefCol
    '- --------------------------------------------------------------------------------------------------------------
    Prog__APP_Switch.Range("Sw_WB_Deactivate") = False     '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ---->>>
    Dim ClosedBook          As Workbook
    Dim Ws_ClsBk_LSace06    As Worksheet
    Dim Lo_ClsBk_LSace06    As ListObject
    Set ClosedBook = Workbooks.Open(Arch_New_Name, ReadOnly:=True)
    Set Ws_ClsBk_LSace06 = ClosedBook.Sheets(1)
    Ws_ClsBk_LSace06.Name = "Ws_LSace06"   '¡¡ Tengo que cambiar el nombre pq sino coinciden entre los dos ficheros !!
    '- Crear ListObject Lo_ClsBk_LSace06 si no existe. ---------------------------------
    If Ws_ClsBk_LSace06.ListObjects.Count = 0 Then
            Ws_ClsBk_LSace06.UsedRange.Cells(1, 1).Select  'Posicionar cursor
        Set Lo_ClsBk_LSace06 = Ws_ClsBk_LSace06.ListObjects.Add(xlSrcRange, Ws_ClsBk_LSace06.UsedRange, , xlYes)
    Else
        Set Lo_ClsBk_LSace06 = Ws_ClsBk_LSace06.ListObjects(1)
        Lo_ClsBk_LSace06.ShowTotals = False
    End If
    Lo_ClsBk_LSace06.Name = "Lo_LSace06"   '¡¡ Tengo que cambiar el nombre pq sino coinciden entre los dos ficheros !!
    '- Comprobar que la Tabla Lo_ClsBk_LSace06 tiene datos ---------------------
    If Lo_ClsBk_LSace06.ListRows.Count = 0 Then
        MsgBx_Msg = "¡ La tabla no contiene datos !    " & vbLf & vbLf & NomFichLSace06
        GoTo Abortar_con_Aviso
    End If
    '- Comprobar que la cabecera de la Tabla Lo_ClsBk_LSace06 corresponde con la establecida en la LoDefCol ---------------------
    If Not Func_LstObj_ListColumns_DefCol_Check_OK(Lo_ClsBk_LSace06, LoDefCol, Col_Header) Then
        MsgBx_Msg = "¡ La cabecera de la tabla no coincide !    " & vbLf & "Puede haber un cambio en la estructura de la Tabla"
        GoTo Abortar_con_Aviso
    End If

    '- Robot_PPub_Fusión añade al final de sus ficheros una Col. ORIGEN: no es un dato del informe, se quita antes de procesar. -----
    Dim Lc_Origen       As ListColumn
    On Error Resume Next
    Set Lc_Origen = Lo_ClsBk_LSace06.ListColumns("ORIGEN")
    On Error GoTo 0
    If Not Lc_Origen Is Nothing Then
        Lc_Origen.Delete
        Set Lc_Origen = Nothing
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Quitada la Col. ORIGEN que añade Robot_PPub_Fusión.", 0)
    End If

    '- Formateo. ----------------------------------------------------
    Call Rut_Lo_Format_LoData_LoDefColData(Lo_ClsBk_LSace06, LoDefCol)

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- El fichero del Robot es de UN solo curso: sólo se comprueba el curso. Si trae recibos de otro curso, se aborta antes de tocar BD_ImpAdm.
    Dim VCurso      As Variant
    Dim OtroCurso   As Long
    VCurso = Lo_ClsBk_LSace06.ListColumns(IAdm_C_Acad).DataBodyRange.Value2
    If IsArray(VCurso) Then
        For rowfind = 1 To UBound(VCurso, 1)
            If CStr(VCurso(rowfind, 1)) <> C_Acad_Imp Then OtroCurso = OtroCurso + 1
        Next rowfind
    ElseIf CStr(VCurso) <> C_Acad_Imp Then
        OtroCurso = 1
    End If
    VCurso = Empty
    If OtroCurso > 0 Then
        MsgBx_Msg = "¡ El fichero tiene " & Format(OtroCurso, "#,##0") & " recibos de un curso distinto de " & C_Acad_Imp & " !" & vbLf & vbLf & NomFichLSace06
        GoTo Abortar_con_Aviso
    End If
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "El LSace06 IAdm trae Rec. de " & Text & ": ", 0, _
                                                        Format(Lo_ClsBk_LSace06.ListRows.Count, "#,##0") & " reg")

    '- -------------------------------------------------------------------------------------------------------------------------------------
    '- -------------------------------------------------------------------------------------------------------------------------------------
    '- Solo se copian los recibos EMITIDOS o COBRADOS en el Año Contable (ACont_Emi = AnoCont o ACont_Cob = AnoCont): se borran del ClsBk
    '- los que no cumplen ninguna de las dos (ACont_Emi <> AnoCont Y ACont_Cob <> AnoCont; un cobro vacío cuenta como distinto).
    '- -------------------------------------------------------------------------------------------------------------------------------------
    rowfind = Lo_ClsBk_LSace06.ListRows.Count
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk_LSace06, IAdm_ACont_Emi, "<>" & AnoCont, IAdm_ACont_Cob, "<>" & AnoCont)
    Dim NoAnoCont   As Long:    NoAnoCont = rowfind - Lo_ClsBk_LSace06.ListRows.Count
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Borrados Rec. NO emitidos ni cobrados en " & AnoCont & ": ", 0, _
                                                        Format(NoAnoCont, "#,##0") & " reg", "quedan " & Format(Lo_ClsBk_LSace06.ListRows.Count, "#,##0") & " reg")

    '- Copiar Lo_ClsBk_LSace06 en Lo_ImpAdm (que llega vacía, o con el otro curso: lo vacía RuT_Update_LSace06_CAcad_ImpAdm).
    '- Antes hay que darle un respiro a Excel (DoEvents): tras procesar un fichero grande del Robot, tocar la estructura de una tabla
    '- sin él puede fallar con -2147417848 (80010108), ver M_311. ------------------------------------------------------------------------
    DoEvents
    Call Rut_Lo_WrkSht_Preparar(Wh_ImpAdm)
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_LSace06)
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk_LSace06, Lo_ImpAdm)
    rowfind = Lo_ClsBk_LSace06.ListRows.Count
    ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing

    '- Anchos y columnas ocultas de Lo_ImpAdm según la DefCol -----------------
    Dim HiddenCol   As Boolean
    For Ccol = 1 To LoDefCol.DataBodyRange.Rows.Count
        Wh_ImpAdm.Columns(Lo_ImpAdm.ListColumns(Ccol).Range.Column).ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht).Value
        HiddenCol = LoDefCol.DataBodyRange.Cells(Ccol, DefC_HiddenCol)
        Wh_ImpAdm.Columns(Lo_ImpAdm.ListColumns(Ccol).Range.Column).Hidden = HiddenCol
    Next Ccol

    Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        '- Visualizo el progreso  ---------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copiados los nuevos recibos:", LastTimeLap)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "en BD_ImpAdm Rec. de " & Text, 0, _
                        Format(rowfind, "#,##0") & " reg", "Total: " & Format(Lo_ImpAdm.ListRows.Count, "#,##0") & " reg")

    '- Fecha de la última importación de cada curso (celdas C2 / C3 de la hoja; propias de BD_ImpAdm, no tocan las de INSS) ------
    If C_Acad_Imp = C_Acad_Pos Then
        Wh_ImpAdm.Range("c3") = "Última Importación LSace06 IAdm C_Acad_Pos - " & C_Acad_Pos & " - el " & Format(Now(), "dd-mmm-yy hh:mm")
    Else
        Wh_ImpAdm.Range("c2") = "Última Importación LSace06 IAdm C_Acad_Ant - " & C_Acad_Ant & " - el " & Format(Now(), "dd-mmm-yy hh:mm")
    End If
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe").Text

Rut_On_Functions
Debug.Print "<<< Rut_Lo_Import_LSace06_ImpAdm"
Exit Sub

'- Salida de aborto: cierra el ClsBk (si estaba abierto), devuelve "Cancel" y deja el informe cerrado. ------------------------------------
Abortar_con_Aviso:
    MsgBox MsgBx_Msg, vbExclamation, "Proceso: Importar " & ArchRequest
    Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & vbLf & Now()
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
    Arch_New_Name = "Cancel"
    If Not ClosedBook Is Nothing Then ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing
    Set Lo_ClsBk_LSace06 = Nothing
    Set Ws_ClsBk_LSace06 = Nothing
    Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
    Rut_On_Functions
    Debug.Print "<<< Rut_Lo_Import_LSace06_ImpAdm (abortado)"
End Sub
