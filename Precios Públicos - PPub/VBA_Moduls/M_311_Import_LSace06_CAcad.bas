Attribute VB_Name = "M_311_Import_LSace06_CAcad"
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
'    -            Borrar de Lo_ClsBk_LSace06 Recibos de LSave06 con C_Acad <> C_Acad_Ant y C_Acad_Pos, para aligerar el peso de la Tabla.
'    -            Borrar de Lo_ClsBk_LSace06 Recibos "<>INSS" en Nom_Concepto, para aligerar el peso de la Tabla.
'    -            Borrar Recibos con Imp.Rec. < 0.
'    -            M_314_Find_Rec_INSS, Identificar de un C_Acad, los 1º Rec. con Seguro obligatorio INSS, borrar los que no son.
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
                                        Arch_New_Name As String, _
                                        Optional Arch_Path As String, _
                                        Optional SheetNom As String = "")
                             
Debug.Print ">>> Rut_Lo_Import_LoData_LoDefCol_LSace06"
Rut_Off_Functions

    Dim rowfind         As Variant
    Dim ArchRequest     As String:      ArchRequest = Arch_New_Name
    Dim TxT_ProgIni     As String:      TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Dim TxT_Progreso    As String
    Dim NomFichLSace06  As String
    Dim RutaFichLsace06 As String
    Dim Ccol            As Integer
    Dim FichNameINSS    As String
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim Wh_INSS         As Worksheet:   Set Wh_INSS = Lo_INSS.Parent

    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
        
    '- Select File -------------------------------------------------------------------------------------
    Call Rut_File_Select_V2("Seleccionar el Nuevo Fichero Excel LSaces06 " & Arch_New_Name & ": ", Arch_New_Name, "Excel", "*.xls?")
        If Arch_New_Name = "Cancel" Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ Cancelado a petición del Usuario !    "
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
            '- Visualizo el progreso  <<<<>>>>
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
            Exit Sub
        End If
        Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NomFichLSace06, RutaFichLsace06)
        '- Comprueba que se ha seleccionado el nombre adecuado de Excel. --------------------------
        If Left(NomFichLSace06, Len(ArchRequest)) <> ArchRequest Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ Cancelado, el fichero debe ser un " & ArchRequest & " * !" & vbLf & vbLf & "  y se ha seleccionado:  " & Arch_New_Name
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & vbLf & Now()
            '- Visualizo el progreso
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
            Arch_New_Name = "Cancel"
            Exit Sub
        End If
            '- Visualizo el progreso
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Excel Seleccionado: " & NomFichLSace06, 0, , , TxT_Progreso)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Ruta: " & RutaFichLsace06, 0)
            TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Compruebo ClsBk:
    '-            1º Si he dado una SheetNom si Existe
    '-            2º Si no Existe LisObject la creo
    '-            3º Comprobar que la Tabla Lo_ClsBk_LSace06 tiene datos
    '-            4º Comprobar que la cabecera de la Tabla Lo_ClsBk_LSace06 corresponde con la establecida en la LoDefCol
    '- --------------------------------------------------------------------------------------------------------------
    Prog__APP.Range("SW_WB_Deactivate") = False     '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ---->>>
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
            Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
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
        Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
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
        Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        Exit Sub
    End If
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Proceso ClsBk:
    '-            Formateo.
    '-            Determinar si en el LSace06 Hay UNO o DOS Cursos Académicos, Crea Dictionary para valores únicos (eficiente para grandes datos)
    '-            Borrar de Lo_ClsBk_LSace06 Recibos de LSave06 con C_Acad <> C_Acad_Ant y C_Acad_Pos, para aligerar el peso de la Tabla.
    '-            Borrar de Lo_ClsBk_LSace06 Recibos "<>INSS" en Nom_Concepto, para aligerar el peso de la Tabla.
    '-            Borrar Recibos con Imp.Rec. < 0.
    '-            M_314_Find_Rec_INSS, Identificar de un C_Acad, los 1º Rec. con Seguro obligatorio INSS, borrar los que no son.
    '-   ?????         Borrar Recibos de C_Acad_Ant y Cobrados en Año_Cont_Ant.
    '-            Borro los datos de las columnas NO necesarias.
    '- --------------------------------------------------------------------------------------------------------------
    
    '- Formateo. ----------------------------------------------------
    Call Rut_Lo_Format_LoData_LoDefColData(Lo_ClsBk_LSace06, LoDefCol)
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- Determinar si en el LSace06 Hay UNO o DOS Cursos Académicos, Crea Dictionary para valores únicos (eficiente para grandes datos)
    Dim DiccCAcad As Object
    Set DiccCAcad = CreateObject("Scripting.Dictionary")
    Dim Cell    As Range
    Dim Valor   As String
    For Each Cell In Lo_ClsBk_LSace06.ListColumns(LS06_C_Acad).DataBodyRange     ' Recorre solo los datos (sin encabezado)
        Valor = CStr(Cell.Value)
        If Valor <> "" And Not DiccCAcad.Exists(Valor) Then DiccCAcad.Add Valor, Nothing
    Next Cell
    ' Convierte Keys en array indexado
    Dim CantCAcad   As Integer
    CantCAcad = DiccCAcad.Count
    Dim CursoArray As Variant
    CursoArray = DiccCAcad.Keys     '- para extraer datos: CStr(CursoArray(índice)), el índice empieza por CERO
    Dim Text    As String
    If CantCAcad = 2 Then
        Text = "C_Acad " & CStr(CursoArray(0)) & " y " & CStr(CursoArray(1))
        Prog__APP.Range("APP_Last_LSace06_CAcadAnt") = Format(Now(), "dd-mmm-yy hh:mm")
        Prog__APP.Range("APP_Last_LSace06_CAcadPos") = Format(Now(), "dd-mmm-yy hh:mm")
    Else
        Text = "C_Acad " & CStr(CursoArray(0))
        If CStr(CursoArray(0)) = C_Acad_Pos Then
            Prog__APP.Range("APP_Last_LSace06_CAcadPos") = Format(Now(), "dd-mmm-yy hh:mm")
        Else
            Prog__APP.Range("APP_Last_LSace06_CAcadAnt") = Format(Now(), "dd-mmm-yy hh:mm")
        End If
    End If
    Set DiccCAcad = Nothing
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- Borrar de Lo_ClsBk_LSace06 Recibos de LSave06 con C_Acad <> C_Acad_Ant y C_Acad_Pos, para aligerar el peso de la Tabla. ---------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_LSace06)
    Call Rut_Lo_Sort(Lo_ClsBk_LSace06, LS06_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---------------------
    Lo_ClsBk_LSace06.Range.AutoFilter Field:=LS06_C_Acad, Criteria1:="<>" & C_Acad_Ant, Operator:=xlAnd, Criteria2:="<>" & C_Acad_Pos
    rowfind = Lo_ClsBk_LSace06.Range.Columns(LS06_C_Acad).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    If rowfind > 0 Then
        Lo_ClsBk_LSace06.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        '- Comprobar que a la Tabla Lo_ClsBk_LSace06 le quedan datos ---------------------
        If Lo_ClsBk_LSace06.ListRows.Count = 0 Then
Proceso_Finalizado_por_quedarse_sin_Registros:
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ A la tabla Lo_ClsBk_LSace06 no le quedan Recibos procesables !"
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
            Arch_New_Name = "Cancel"
            ClosedBook.Close SaveChanges:=False
            Set ClosedBook = Nothing
            Set Lo_ClsBk_LSace06 = Nothing
            Set Ws_ClsBk_LSace06 = Nothing
            Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
            Exit Sub
        End If
    End If
            '- Visualizo el progreso
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Del de LSace06 Rec. <> " & Text, 0, _
                                                            Format(rowfind, "#,##0") & " reg", _
                                                            "quedan " & Format(Lo_ClsBk_LSace06.ListRows.Count, "#,##0") & " reg")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- Borrar de Lo_ClsBk_LSace06 Recibos "<>INSS" en Nom_Concepto -----------------------------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_LSace06)
    Call Rut_Lo_Sort(Lo_ClsBk_LSace06, LS06_Concept_Nom, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---------------------
    Lo_ClsBk_LSace06.Range.AutoFilter Field:=LS06_Concept_Nom, Criteria1:="<>*INSS*"    ', Operator:=xlAnd, Criteria2:="<>"
    rowfind = Lo_ClsBk_LSace06.Range.Columns(LS06_C_Acad).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    If rowfind > 0 Then
        Lo_ClsBk_LSace06.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        '- Comprobar que a la Tabla Lo_ClsBk_LSace06 le quedan datos ---------------------
        If Lo_ClsBk_LSace06.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    End If
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Del de LSace06 Matrículas SIN Seguro INSS: ", 0, _
                                                        Format(rowfind, "#,##0") & " reg", _
                                                        "quedan " & Format(Lo_ClsBk_LSace06.ListRows.Count, "#,##0") & " reg")
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- Borrar Recibos con Imp.Rec. < 0 -----------------------------------------------------------------------------------
    rowfind = Lo_ClsBk_LSace06.ListRows.Count
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk_LSace06, LS06_ImpRec, "<0")
    rowfind = rowfind - Lo_ClsBk_LSace06.ListRows.Count
    If rowfind > 0 Then
        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "Borrados Rec. Negativos. ", 0, _
                                                        Format(rowfind, " #,##0") & " reg", _
                                                        "de " & Format(Lo_ClsBk_LSace06.ListRows.Count, "#,##0") & " reg")
        '- Comprobar que a la Tabla Lo_ClsBk_LSace06 le quedan datos ---------------------
        If Lo_ClsBk_LSace06.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(11, " ") & "No hay Rec. Negativos. ", 0)
    End If
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- M_314_Find_Rec_INSS, Identificar de un C_Acad, los 1º Rec. con Seguro obligatorio INSS, borrar los que no son. -------------------
    Call RuT_Find_Rec_INSS_C_Acad(Lo_ClsBk_LSace06, NomFichLSace06, RutaFichLsace06)
        '- Comprobar que a la Tabla Lo_ClsBk_LSace06 le quedan datos ---------------------
        If Lo_ClsBk_LSace06.ListRows.Count = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '- Borro los datos de las columnas NO necesarias. ------------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk_LSace06)
    Call Rut_Lo_ListColumns_ClearContents_DefC_ProtectData(Lo_ClsBk_LSace06, LoDefCol, DefC_ProtectData)
    Ccol = Application.CountIf(LoDefCol.ListColumns(DefC_HiddenCol).DataBodyRange, True)
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
    Call Rut_Lo_WrkSht_Preparar(Wh_INSS)
    '- -------------------------------------------------------------------------------------------------------------------------------------
    If Not Lo_INSS.DataBodyRange Is Nothing Then
'        Call Rut_Lo_Filtros_Quitar(Lo_INSS)
        Call Rut_Lo_Sort(Lo_INSS, LS06_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---------------------
        If CantCAcad = 2 Then
            Lo_INSS.Range.AutoFilter Field:=LS06_C_Acad, Criteria1:="=" & CStr(CursoArray(0)), Operator:=xlOr, Criteria2:="=" & CStr(CursoArray(1))
        Else
            Lo_INSS.Range.AutoFilter Field:=LS06_C_Acad, Criteria1:="=" & CStr(CursoArray(0))
        End If
        rowfind = Lo_INSS.Range.Columns(LS06_C_Acad).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then Lo_INSS.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                '- Visualizo el progreso
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del de BD_INSS Rec. con " & Text, LastTimeLap, Format(rowfind, "#,##0") & " reg", _
                                                               "quedan " & Format(Lo_INSS.ListRows.Count, "#,##0") & " reg")
        Call Rut_Lo_Filtros_Quitar(Lo_INSS)
        Call Rut_WrkSheet_LstObj_LiberarEspacio(Wh_INSS)
    End If
        
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
    
    
    Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
            Dim TimeLap2              As Single
            TimeLap2 = LastTimeLap
''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importado: " & Arch_New_Name, 0)
''            LastTimeLap = TimeLap2
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copiados los nuevos recibos:", LastTimeLap)
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "en BD_INSS Rec. de " & Text, 0, _
                            Format(rowfind, "#,##0") & " reg", "Total: " & Format(Lo_INSS.ListRows.Count, "#,##0") & " reg")
            
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe")
    Sht__BD_INSS.Range("c2") = "Úlitma Importación LSace06 C_Acad_Ant - " & C_Acad_Ant & " - el " & Prog__APP.Range("APP_Last_LSace06_CAcadAnt")
    Sht__BD_INSS.Range("c3") = "Úlitma Importación LSace06 C_Acad_Pos - " & C_Acad_Pos & " - el " & Prog__APP.Range("APP_Last_LSace06_CAcadPos")

Debug.Print "<<< Rut_Lo_Import_LoData_LoDefCol"
End Sub


