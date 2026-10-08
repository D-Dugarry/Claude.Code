Attribute VB_Name = "M_311_Import_LSace06_CAcad"
' Last Rev. 2026-10-08 13:46
'Rev.: 2026-01-22
Option Explicit

'   EN ESTE MÓDULO PROCESO TODO LO QUE PUEDO EN EL ClsBk PARA TRABAJAR AL MÁXIMO EN LA RAM

' Rut_Lo_Import_LoData_LoDefCol_LSace06
'
'    Importar LSsace06 de Curso_Acad_Ant o Curso_Acad_Pos
'           Para Identificar Rec. Matrículas con seguro obligatorio INSS en Tasa Adm.
'           Tengo que tener los dos Cursos Acad actualizados.
'
'    - Select File
'    - Con el ClsBk: (RAM)
'    - Compruebo ClsBk:
'    _            1º Si he dado una SheetNom, ver si Existe
'    -            2º Si no Existe LisObject la creo
'    -            3º Comprobar que la Tabla Lo_ClsBk_LSace06 tiene datos
'    -            4º Comprobar que la cabecera de la Tabla Lo_ClsBk_LSace06 corresponde con la establecida en la LoDefCol
'    - --------------------------------------------------------------------------------------------------------------
'    - Proceso ClsBk:
'    -            Formateo.
'    -            Determinar si en el LSace06 Hay UNO o DOS Cursos Académicos, Crea Dictionary para valores únicos (eficiente para grandes datos)
'    -            Comprobar que todos los Recibos son del C_Acad pedido (el Robot ya trae un solo curso).
'    -            (Ya no se borran los "<>INSS" ni se queda 1 Rec. por Plan/DNI: el LSace06 INSS del Robot ya viene filtrado.)
'    -            Borrar Recibos con importe < 0 (el Robot no los excluye).
'    -            M_314_Find_Rec_INSS RETIRADO (2026-10-08): el LSace06 INSS del Robot ya viene filtrado.
'    -  ?????          Borrar Recibos de C_Acad_Ant y Cobrados en Año_Cont_Ant.
'    -            Borro los datos de las columnas NO necesarias.
'    - --------------------------------------------------------------------------------------------------------------
'    - Copy ClsBk:
'    -            1º Borrar los Recibos de Lo_INSS (Tabla DB_INSS) con el/los C_Acad del Nuevo LSace06
'                       Puesto que voy a copiar los nuevos registros, tengo que eliminar los viejos.
'                       Dependiendo de si en Lo_Source hay 1 ó 2 Cursos Académicos, filtro por 1 o 2 Cursos.
'    -            2º Copiar Lo_ClsBk_LSace06 en Lo_INSS.
'    -            3º Oculto columnas en Lo_INSS de BD_INSS
'    - --------------------------------------------------------------------------------------------------------------

'- ----------------------------------------------------------------------------------------------------------------------------
'- Seleccionar Excel pero no lo importa, sólo lo abre -------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Import_LoData_LoDefCol_LSace06(Lo_INSS As ListObject, _
                                        LoDefCol As ListObject, _
                                        Col_Header As Integer, _
                                        C_Acad_Imp As String, _
                                        Arch_New_Name As String, _
                                        Optional SheetNom As String = "")
'- C_Acad_Imp:    curso academico a importar (p.ej. "2026-27").
'- Arch_New_Name: SALIDA. Ruta completa del fichero que se ha leido, o "Cancel" si se ha abortado.
'- El fichero se busca solo (Fnc_LSace06_Robot_Localizar); si falla se avisa y se ofrece otra fecha, el explorador o abortar.

Debug.Print ">>> Rut_Lo_Import_LoData_LoDefCol_LSace06"
Rut_Off_Functions

    Dim rowfind         As Variant
    Dim NombresValidos  As String:      NombresValidos = Fnc_LSace06_Prefijo(C_Acad_Imp)     '- "LSace06_C_Acad_<curso>_INSS_"
    Dim ArchRequest     As String:      ArchRequest = "LSace06 INSS del curso " & C_Acad_Imp
    Dim Text            As String:      Text = "C_Acad " & C_Acad_Imp
    Dim TxT_ProgIni     As String:      TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Dim TxT_Progreso    As String
    Dim NomFichLSace06  As String
    Dim RutaFichLsace06 As String
    Dim Ccol            As Integer
    Dim FichNameINSS    As String
    Dim AnoCont         As String:      AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim Wh_INSS         As Worksheet:   Set Wh_INSS = Lo_INSS.Parent
    Dim RutaRobot       As String
    Dim RutaAnt         As String
    Dim Fallo           As String
    Dim CarpetaOk       As String
    Dim Resp            As VbMsgBoxResult

    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios

    '- Localizar el fichero: lo normal es que M_310 ya lo haya localizado y llegue en Arch_New_Name (ruta completa);
    '- si no llega, se busca aquí (ruta del Robot + curso + fecha más reciente, o aviso y alternativas).
    If Arch_New_Name = "" Or Arch_New_Name = "Cancel" Then Arch_New_Name = Fnc_LSace06_Elegir_Fichero(C_Acad_Imp)
    If Arch_New_Name = "Cancel" Then GoTo Cancelado_Usuario
    Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NomFichLSace06, RutaFichLsace06)
            '- Visualizo el progreso
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Excel Seleccionado: " & NomFichLSace06, 0, , , TxT_Progreso)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Ruta: " & RutaFichLsace06 & vbLf, 0)
            TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Compruebo ClsBk:
    '-            1º Si he dado una SheetNom si Existe
    '-            2º Si no Existe LisObject la creo
    '-            3º Comprobar que la Tabla Lo_ClsBk_LSace06 tiene datos
    '-            4º Comprobar que la cabecera de la Tabla Lo_ClsBk_LSace06 corresponde con la establecida en la LoDefCol
    '- --------------------------------------------------------------------------------------------------------------
    Prog__APP_Switch.Range("Sw_WB_Deactivate") = False     '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ---->>>
    Dim SheetIndx       As Integer:         SheetIndx = 1
    Dim ClosedBook      As Workbook
    Dim Ws_ClsBk_LSace06    As Worksheet
    Dim Lo_ClsBk_LSace06    As ListObject
    Set ClosedBook = Workbooks.Open(Arch_New_Name, ReadOnly:=True)
    '- Comprobar que la Hoja Existe SI hemos solicitado una hoja concreta para copiar --------
    If SheetNom <> "" Then
        SheetIndx = 0
        For Each Ws_ClsBk_LSace06 In ClosedBook.Sheets
            If Ws_ClsBk_LSace06.Name = SheetNom Then
                SheetIndx = Ws_ClsBk_LSace06.Index
                Exit For
            End If
        Next Ws_ClsBk_LSace06
        If SheetIndx = 0 Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ La Sheet NO existe !    "
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
            '- Visualizo el progreso
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
            Arch_New_Name = "Cancel"
            ClosedBook.Close SaveChanges:=False
            Set ClosedBook = Nothing
            Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
            Rut_On_Functions
            Exit Sub
        End If
    End If
    Set Ws_ClsBk_LSace06 = ClosedBook.Sheets(SheetIndx)
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
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ La tabla no contiene datos !    "
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        '- Visualizo el progreso
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
        Arch_New_Name = "Cancel"
        ClosedBook.Close SaveChanges:=False
        Set ClosedBook = Nothing
        Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        Rut_On_Functions
        Exit Sub
    End If
    '- Comprobar que la cabecera de la Tabla Lo_ClsBk_LSace06 corresponde con la establecida en la LoDefCol ---------------------
    If Not Func_LstObj_ListColumns_DefCol_Check_OK(Lo_ClsBk_LSace06, LoDefCol, Col_Header) Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ La cabedera de la tabla no coincide !    " & vbLf & "Puede haber un cambio en la estructura de la Tabla"
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        '- Visualizo el progreso
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
        Arch_New_Name = "Cancel"
        ClosedBook.Close SaveChanges:=False
        Set ClosedBook = Nothing
        Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        Rut_On_Functions
        Exit Sub
    End If
    
    '- Robot_PPub_Fusión añade al final de sus ficheros una Col. ORIGEN (el fichero del Robot de donde sale cada fila).
    '- No es un dato del informe: se quita antes de procesar, para que no llegue a Lo_INSS. -------------------------------
    Dim Lc_Origen       As ListColumn
    On Error Resume Next
    Set Lc_Origen = Lo_ClsBk_LSace06.ListColumns("ORIGEN")
    On Error GoTo 0
    If Not Lc_Origen Is Nothing Then
        Lc_Origen.Delete
        Set Lc_Origen = Nothing
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Quitada la Col. ORIGEN que añade Robot_PPub_Fusión.", 0)
    End If
    '- --------------------------------------------------------------------------------------------------------------
    '- Proceso ClsBk:
    '-            Formateo.
    '-            Determinar si en el LSace06 Hay UNO o DOS Cursos Académicos, Crea Dictionary para valores únicos (eficiente para grandes datos)
    '-            Comprobar que todos los Recibos son del C_Acad pedido (el Robot ya trae un solo curso).
    '-            (Ya no se borran los "<>INSS" ni se queda 1 Rec. por Plan/DNI: el LSace06 INSS del Robot ya viene filtrado.)
    '-            Borrar Recibos con importe < 0 (el Robot no los excluye).
    '-            M_314_Find_Rec_INSS RETIRADO (2026-10-08): el LSace06 INSS del Robot ya viene filtrado.
    '-   ?????         Borrar Recibos de C_Acad_Ant y Cobrados en Año_Cont_Ant.
    '-            Borro los datos de las columnas NO necesarias.
    '- --------------------------------------------------------------------------------------------------------------
    
    '- Formateo. ----------------------------------------------------
    Call Rut_Lo_Format_LoData_LoDefColData(Lo_ClsBk_LSace06, LoDefCol)
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- El fichero del Robot es de UN solo curso y ya viene filtrado (concepto INSS, un recibo por Plan/DNI): sólo se comprueba el curso.
    '- Si trae recibos de otro curso, se aborta antes de tocar BD_INSS.
    Dim VCurso      As Variant
    Dim OtroCurso   As Long
    VCurso = Lo_ClsBk_LSace06.ListColumns(LS06_C_Acad).DataBodyRange.Value2
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
    If C_Acad_Imp = C_Acad_Pos Then
        Prog__APP.Range("APP_Last_LSace06_CAcadPos") = Format(Now(), "dd-mmm-yy hh:mm")
    Else
        Prog__APP.Range("APP_Last_LSace06_CAcadAnt") = Format(Now(), "dd-mmm-yy hh:mm")
    End If
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "El LSace06 trae Rec. INSS de " & Text & ": ", 0, _
                                                        Format(Lo_ClsBk_LSace06.ListRows.Count, "#,##0") & " reg")

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- Borrar Recibos con importe < 0 (el Robot no los excluye; se borran por seguridad). ----------------------------------------------------
    rowfind = Lo_ClsBk_LSace06.ListRows.Count
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk_LSace06, LS06_Concept_Imp, "<0")
    If Lo_ClsBk_LSace06.ListRows.Count > 0 Then Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk_LSace06, LS06_Concept_TImp, "<0")
    rowfind = rowfind - Lo_ClsBk_LSace06.ListRows.Count
    If rowfind > 0 Then
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Borrados Rec. con importe Negativo. ", 0, _
                                                        Format(rowfind, " #,##0") & " reg", _
                                                        "quedan " & Format(Lo_ClsBk_LSace06.ListRows.Count, "#,##0") & " reg")
        '- Comprobar que a la Tabla Lo_ClsBk_LSace06 le quedan datos ---------------------
        If Lo_ClsBk_LSace06.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "No hay Rec. con importe Negativo. ", 0)
    End If
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- Borro los datos de las columnas NO necesarias. ------------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_LSace06)
    Call Rut_Lo_ListColumns_ClearContents_DefC_ProtectData(Lo_ClsBk_LSace06, LoDefCol, DefC_ProtectData)
    Ccol = Application.CountIf(LoDefCol.ListColumns(DefC_ProtectData).DataBodyRange, True)
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Eliminados Datos de " & Ccol & " Columnas NO necesarias de un total de " & Lo_ClsBk_LSace06.ListColumns.Count & ".", LastTimeLap)
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Copy ClsBk:
    '-            1º Borrar los Recibos de Lo_INSS (Tabla DB_INSS) con el/los C_Acad del Nuevo LSace06
    '                   Puesto que voy a copiar los nuevos registros, tengo que eliminar los viejos.
    '                   Dependiendo de si en Lo_Source hay 1 ó 2 Cursos Académicos, filtro por 1 o 2 Cursos.
    '-            2º Copiar Lo_ClsBk_LSace06 en Lo_INSS.
    '-            3º Oculto columnas en Lo_INSS de BD_INSS
    '- --------------------------------------------------------------------------------------------------------------
    
    '- -------------------------------------------------------------------------------------------------------------------------------------
    '- Borrar los Recibos de Lo_INSS (Tabla DB_INSS) con el/los C_Acad del Nuevo LSace06
    '       Puesto que voy a copiar los nuevos registros, tengo que eliminar los viejos.
    '       Dependiendo de si en Lo_Source hay 1 ó 2 Cursos Académicos, filtro por 1 o 2 Cursos.
    '- -------------------------------------------------------------------------------------------------------------------------------------
    '- Antes de tocar Lo_INSS hay que darle un respiro a Excel (DoEvents). Tras procesar el ClsBk (en el LSace06 del Robot,
    '- 343.000 filas, de las que se borran casi todas), borrar filas de Lo_INSS fallaba con -2147417848 (80010108) y dejaba Excel
    '- inservible hasta cerrarlo (fallaba hasta leer .Hidden). No es ningún paso concreto: con una parada antes (punto de
    '- interrupción) funcionaba, y con DoEvents también (2026-10-05, probado en el trabajo, con el mismo Office que en casa).
    DoEvents
    Call Rut_Lo_WrkSht_Preparar(Wh_INSS)
    '- -------------------------------------------------------------------------------------------------------------------------------------
    '- Lo_INSS ya llega vacía: la vacía M_310 antes de importar el primer curso, y los dos cursos se añaden a continuación.
        
    '- Copiar Lo_ClsBk_LSace06 en Lo_INSS. ---------------------
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk_LSace06, Lo_INSS)
    rowfind = Lo_ClsBk_LSace06.ListRows.Count
    ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing
    
    '- Oculto columnas en Lo_INSS de BD_INSS -----------------
    Dim HiddenCol   As Boolean
    For Ccol = 1 To LoDefCol.DataBodyRange.Rows.Count
'        LoData.Range.Columns(Ccol).ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht).Value
        Wh_INSS.Columns(Lo_INSS.ListColumns(Ccol).Range.Column).ColumnWidth = LoDefCol.DataBodyRange.Cells(Ccol, DefC_Widht).Value
        HiddenCol = LoDefCol.DataBodyRange.Cells(Ccol, DefC_HiddenCol)
        Wh_INSS.Columns(Lo_INSS.ListColumns(Ccol).Range.Column).Hidden = HiddenCol
    Next Ccol
    
    
    Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
            Dim TimeLap2              As Single
            TimeLap2 = LastTimeLap
''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importado: " & Arch_New_Name, 0)
''            LastTimeLap = TimeLap2
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copiados los nuevos recibos:", LastTimeLap)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "en BD_INSS Rec. de " & Text, 0, _
                            Format(rowfind, "#,##0") & " reg", "Total: " & Format(Lo_INSS.ListRows.Count, "#,##0") & " reg")
            
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe").Text
    Sht__BD_INSS.Range("c2") = "Úlitma Importación LSace06 C_Acad_Ant - " & C_Acad_Ant & " - el " & Prog__APP.Range("APP_Last_LSace06_CAcadAnt")
    Sht__BD_INSS.Range("c3") = "Úlitma Importación LSace06 C_Acad_Pos - " & C_Acad_Pos & " - el " & Prog__APP.Range("APP_Last_LSace06_CAcadPos")

Rut_On_Functions
Debug.Print "<<< Rut_Lo_Import_LoData_LoDefCol_LSace06"
Exit Sub

'- Salidas de aborto: cierran el ClsBk (si estaba abierto), devuelven "Cancel" y dejan el informe cerrado. ------------------------------------
Proceso_Finalizado_por_quedarse_sin_Registros:
    MsgBx_Msg = "¡ A la tabla Lo_ClsBk_LSace06 no le quedan Recibos procesables !"
Abortar_con_Aviso:
    MsgBox MsgBx_Msg, vbExclamation, "Proceso: Importar " & ArchRequest
    Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & vbLf & Now()
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
    GoTo Abortar_Cierre
Cancelado_Usuario:
    MsgBx_Msg = "¡ Cancelado a petición del Usuario !    "
    Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
Abortar_Cierre:
    Arch_New_Name = "Cancel"
    If Not ClosedBook Is Nothing Then ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing
    Set Lo_ClsBk_LSace06 = Nothing
    Set Ws_ClsBk_LSace06 = Nothing
    Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
    Rut_On_Functions
    Debug.Print "<<< Rut_Lo_Import_LoData_LoDefCol_LSace06 (abortado)"
End Sub

'- ----------------------------------------------------------------------------------------------------------------------------
'- Nombre (sin fecha) de los ficheros INSS del Robot de un curso:  LSace06_C_Acad_2026-27_INSS_
'- ----------------------------------------------------------------------------------------------------------------------------
Function Fnc_LSace06_Prefijo(ByVal C_Acad As String) As String
    Fnc_LSace06_Prefijo = "LSace06_C_Acad_" & C_Acad & "_INSS_"
End Function

'- ----------------------------------------------------------------------------------------------------------------------------
'- Localiza el LSace06 INSS de un curso en la carpeta del Robot:
'-      <APP_Ruta_Robot_LSc06_Cacad>\<curso>\<carpeta de fecha mas reciente>\LSace06_C_Acad_<curso>_INSS_(<fecha>).xlsx
'- Devuelve True y Ruta (completa) si lo encuentra. Si no: False y
'-      Fallo     = que ha fallado (para el aviso),
'-      RutaAnt   = fichero de la fecha anterior mas cercana que SI lo tiene (o "" si no hay),
'-      CarpetaOk = la ultima carpeta que existe de la ruta, para abrir el explorador ahi.
'- ----------------------------------------------------------------------------------------------------------------------------
Function Fnc_LSace06_Robot_Localizar(ByVal C_Acad As String, ByRef Ruta As String, ByRef Fallo As String, _
                                     ByRef RutaAnt As String, ByRef CarpetaOk As String) As Boolean
    Dim Base        As String
    Dim CarpCurso   As String
    Dim Nom         As String
    Dim Nombres()   As String
    Dim Fechas()    As Date
    Dim N           As Long
    Dim i           As Long
    Dim iMax        As Long
    Dim Usado()     As Boolean
    Dim Pos         As Long
    Dim Fich        As String
    Dim Primera     As Boolean:     Primera = True
    Ruta = "":  Fallo = "":  RutaAnt = "":  CarpetaOk = ""

    '- 1) Ruta configurada ------------------------------------------------------------------------
    On Error Resume Next
    Base = Trim$(CStr(Prog__APP.Range("APP_Ruta_Robot_LSc06_Cacad").Value))
    On Error GoTo 0
    If Base = "" Then
        Fallo = "No está configurada la ruta del Robot LSace06 (nombre APP_Ruta_Robot_LSc06_Cacad, hoja ConfigAPP)."
        Exit Function
    End If
    Base = Fnc_Format_Ruta(Base)                                    '- Si es una URL de la Red Nexe, la pasa a letra de unidad
    Do While Right$(Base, 1) = "\" Or Right$(Base, 1) = "/"
        Base = Left$(Base, Len(Base) - 1)
    Loop
    If LCase$(Left$(Base, 4)) = "http" Then
        Fallo = "La ruta del Robot LSace06 es una URL que no se ha podido pasar a unidad de disco:" & vbLf & Base
        Exit Function
    End If
    If Not Fnc_LSc06_Carpeta_Existe(Base) Then
        Fallo = "No existe (o no está accesible) la ruta del Robot LSace06:" & vbLf & Base
        Exit Function
    End If
    CarpetaOk = Base

    '- 2) Carpeta del curso -----------------------------------------------------------------------
    CarpCurso = Base & "\" & C_Acad
    If Not Fnc_LSc06_Carpeta_Existe(CarpCurso) Then
        Fallo = "No existe la carpeta del curso " & C_Acad & ":" & vbLf & CarpCurso
        Exit Function
    End If
    CarpetaOk = CarpCurso

    '- 3) Carpetas de fecha (AAAA-MM-DD, con o sin sufijo) ------------------------------------------
    ReDim Nombres(1 To 1):  ReDim Fechas(1 To 1)
    On Error Resume Next
    Nom = Dir$(CarpCurso & "\*", vbDirectory)
    Do While Nom <> ""
        If Nom <> "." And Nom <> ".." Then
            If Nom Like "####-##-##*" Then
                If (GetAttr(CarpCurso & "\" & Nom) And vbDirectory) <> 0 Then
                    Dim Fx  As Date
                    Fx = 0
                    Fx = DateSerial(CInt(Left$(Nom, 4)), CInt(Mid$(Nom, 6, 2)), CInt(Mid$(Nom, 9, 2)))
                    If Fx > 0 And Format$(Fx, "yyyy-mm-dd") = Left$(Nom, 10) Then       '- fecha valida (descarta 2026-13-45)
                        N = N + 1
                        ReDim Preserve Nombres(1 To N):  ReDim Preserve Fechas(1 To N)
                        Nombres(N) = Nom:  Fechas(N) = Fx
                    End If
                End If
            End If
        End If
        Nom = Dir$
    Loop
    On Error GoTo 0
    If N = 0 Then
        Fallo = "No hay carpetas de fecha (AAAA-MM-DD) en:" & vbLf & CarpCurso
        Exit Function
    End If

    '- 4) De la mas reciente hacia atras: la 1ª con fichero es la buena; si la mas reciente no lo tiene, la siguiente es RutaAnt
    ReDim Usado(1 To N)
    Do
        iMax = 0
        For i = 1 To N
            If Not Usado(i) Then
                If iMax = 0 Then
                    iMax = i
                ElseIf Fechas(i) > Fechas(iMax) Then
                    iMax = i
                End If
            End If
        Next i
        If iMax = 0 Then Exit Do
        Usado(iMax) = True
        Fich = CarpCurso & "\" & Nombres(iMax) & "\" & Fnc_LSace06_Prefijo(C_Acad) & "(" & Left$(Nombres(iMax), 10) & ").xlsx"
        If Fnc_LSc06_Archivo_Existe(Fich) Then
            If Primera Then
                Ruta = Fich
                Fnc_LSace06_Robot_Localizar = True
            Else
                RutaAnt = Fich
            End If
            Exit Do
        End If
        If Primera Then
            Fallo = "En la carpeta más reciente (" & Nombres(iMax) & ") no está el fichero:" & vbLf & _
                    Fnc_LSace06_Prefijo(C_Acad) & "(" & Left$(Nombres(iMax), 10) & ").xlsx"
            CarpetaOk = CarpCurso & "\" & Nombres(iMax)
            Primera = False
        End If
    Loop
    If Not Fnc_LSace06_Robot_Localizar And Fallo = "" Then Fallo = "No se ha encontrado ningún fichero " & Fnc_LSace06_Prefijo(C_Acad) & "* en:" & vbLf & CarpCurso
End Function

'- Dir$ da error 52 con rutas URL: aqui se devuelve False en vez de reventar. --------------------------------------------------
Private Function Fnc_LSc06_Carpeta_Existe(ByVal Ruta As String) As Boolean
    On Error Resume Next
    Fnc_LSc06_Carpeta_Existe = ((GetAttr(Ruta) And vbDirectory) <> 0)
    If Err.Number <> 0 Then Fnc_LSc06_Carpeta_Existe = False
End Function
Private Function Fnc_LSc06_Archivo_Existe(ByVal Ruta As String) As Boolean
    On Error Resume Next
    Fnc_LSc06_Archivo_Existe = ((GetAttr(Ruta) And vbDirectory) = 0)
    If Err.Number <> 0 Then Fnc_LSc06_Archivo_Existe = False
End Function




'- ----------------------------------------------------------------------------------------------------------------------------
'- Devuelve la ruta completa del LSace06 INSS de un curso, o "Cancel". Lo busca en la carpeta del Robot (Fnc_LSace06_Robot_Localizar);
'- si falla, avisa del fallo y ofrece la fecha anterior (si la hay), el explorador o abortar. Lo elegido a mano tiene que llamarse
'- LSace06_C_Acad_<curso>_INSS_*. No abre el fichero ni toca nada: M_310 la usa para tener constancia de los dos antes de vaciar BD_INSS.
'- ----------------------------------------------------------------------------------------------------------------------------
Function Fnc_LSace06_Elegir_Fichero(ByVal C_Acad_Imp As String) As String
    Dim NombresValidos  As String:      NombresValidos = Fnc_LSace06_Prefijo(C_Acad_Imp)     '- "LSace06_C_Acad_<curso>_INSS_"
    Dim ArchRequest     As String:      ArchRequest = "LSace06 INSS del curso " & C_Acad_Imp
    Dim RutaRobot       As String
    Dim RutaAnt         As String
    Dim Fallo           As String
    Dim CarpetaOk       As String
    Dim Arch            As String
    Dim NomFich         As String
    Dim RutaFich        As String
    Dim Resp            As VbMsgBoxResult
    Fnc_LSace06_Elegir_Fichero = "Cancel"

    Do
        If Fnc_LSace06_Robot_Localizar(C_Acad_Imp, RutaRobot, Fallo, RutaAnt, CarpetaOk) Then
            Fnc_LSace06_Elegir_Fichero = RutaRobot
            Exit Function
        End If
        '- Ha fallado: aviso con el fallo concreto y las salidas posibles --------------------------------
        If RutaAnt <> "" Then
            Resp = MsgBox(Fallo & vbLf & vbLf & _
                          "Hay un fichero de una fecha anterior:" & vbLf & "     " & RutaAnt & vbLf & vbLf & _
                          "Sí = usar el de la fecha anterior" & vbLf & _
                          "No = elegir el fichero a mano (explorador)" & vbLf & _
                          "Cancelar = no importar este curso", _
                          vbYesNoCancel + vbExclamation + vbDefaultButton3, "Proceso: Importar " & ArchRequest)
            If Resp = vbYes Then
                Fnc_LSace06_Elegir_Fichero = RutaAnt
                Exit Function
            End If
        Else
            Resp = MsgBox(Fallo & vbLf & vbLf & _
                          "Sí = elegir el fichero a mano (explorador)" & vbLf & _
                          "No = no importar este curso", _
                          vbYesNo + vbExclamation + vbDefaultButton2, "Proceso: Importar " & ArchRequest)
            If Resp = vbYes Then Resp = vbNo Else Resp = vbCancel
        End If
        If Resp = vbCancel Then Exit Function
        '- A mano: tiene que llamarse LSace06_C_Acad_<curso>_INSS_... ------------------------------------
        Arch = NombresValidos
        Call Rut_File_Select_V2("Seleccionar el Fichero Excel " & NombresValidos & "*: ", Arch, "Excel", "*.xls?", CarpetaOk)
        If Arch = "Cancel" Then Exit Function
        Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch, NomFich, RutaFich)
        If Fnc_Nombre_Fichero_Valido(NomFich, NombresValidos) Then
            Fnc_LSace06_Elegir_Fichero = Arch
            Exit Function
        End If
        MsgBox "¡ El fichero debe llamarse  " & NombresValidos & "*  !" & vbLf & vbLf & _
               "y se ha seleccionado:  " & NomFich, vbExclamation, "Proceso: Importar " & ArchRequest
    Loop
End Function