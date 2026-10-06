Attribute VB_Name = "Rut_Lo_Import"
' Last Rev. 2026-10-06 11:10
'2026-01-18
Option Explicit

'- ============================================================================================================================
'- Seleccionar Excel e importarlo ----------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Import_WorkSheet(WrkSht As Worksheet, _
                            Arch_New_Name As String, _
                            Optional SheetNom As String = "")
                             
Debug.Print ">>> Rut_Lo_Import_LoData_LoDefCol"
Rut_Off_Functions
    Dim rowfind             As Variant
    Dim SheetIndx           As Integer:         SheetIndx = 1
    Dim ArchRequest         As String:          ArchRequest = Arch_New_Name
    Dim TxT_Progreso        As String
    Dim NomArch             As String
    Dim AnoCont             As String:          AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim Lo_Data             As ListObject:      Set Lo_Data = WrkSht.ListObjects(1)

    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
        
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Seleccionar el fichero y la ruta, para importar: " & Arch_New_Name, 0)
    '- Select File -------------------------------------------------------------------------------------
    Call Rut_File_Select("Seleccionar el Nuevo Fichero Excel " & Arch_New_Name & ": ", Arch_New_Name, "Excel", "*.xls?")
        If Arch_New_Name = "Cancel" Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ Cancelado a petición del Usuario !    "
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
'            With ActivForm.Controls("TBx_Informe")
'                .Value = MsgBx_Msg & vbLf & Now()
'                ActivForm.Repaint
'            End With
            Rut_On_Functions
            Exit Sub
        End If
        NomArch = Dir(Arch_New_Name)
        '- Comprueba que se ha seleccionado el nombre adecuado de Excel. --------------------------
        If Left(NomArch, Len(ArchRequest)) <> ArchRequest Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ Cancelado, el fichero debe ser un " & ArchRequest & " * !" & vbLf & vbLf & "  y se ha seleccionado:  " & NomArch
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & vbLf & Now()
'            With ActivForm.Controls("TBx_Informe")
'                .Value = MsgBx_Msg & vbLf & Now()
'                ActivForm.Repaint
'            End With
            Arch_New_Name = "Cancel"
            Rut_On_Functions
            Exit Sub
        End If
            
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
''''            ActivForm.Controls("Lb_Tit_Informe").Caption = "Progreso Tarea: Importar Informe: " & NomArch
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Excel Seleccionado: " & NomArch, LastTimeLap)
            TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    '- --------------------------------------------------------------------------------------------------------------
    '   Borrar el contenido de la Tabla -----------------------------------------------------------------------------
    '- Sin la fila de totales: con ella visible, el PasteSpecial de más abajo (Lo_Data.Range.Offset(1, 0)) la pisaría y dejaría
    '- los datos fuera de la tabla (el fallo de Lo_TPV en JyC). Se deja como estaba en cada salida (Docs/Plan_ShowTotals.md, fase 2).
    Dim Totales_Visibles    As Boolean:     Totales_Visibles = Fnc_Lo_Totales_Ocultar(Lo_Data)
    If Not Lo_Data.DataBodyRange Is Nothing Then Lo_Data.DataBodyRange.Delete
    Call Rut_WrkSheet_LstObj_LiberarEspacio(WrkSht)
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Base de Datos vaciada.", LastTimeLap)
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Copy File Without Opening it  ------------------------------------------------------------------
    '- --------------------------------------------------------------------------------------------------------------
            '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importando: " & NomArch, 0, , , , , , 4)
    Prog__APP_Switch.Range("Sw_WB_Deactivate") = False     '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ---->>>
    Dim Ws              As Worksheet
    Dim ClosedBook      As Workbook
    On Error Resume Next
    Set ClosedBook = Workbooks.Open(Arch_New_Name, ReadOnly:=True)
    On Error GoTo 0
    If ClosedBook Is Nothing Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡No se ha podido abrir el fichero!" & vbLf & _
                    "Puede que ya está abierto un libro con el mismo nombre en esta sesión de Excel," & vbLf & _
                    "o que la ruta no sea accesible: " & vbLf & Arch_New_Name
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        Arch_New_Name = "Cancel"
        Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
        Rut_On_Functions
        Exit Sub
    End If
        '- Comprobar que la Hoja Existe SI hemos solicitado una hoja concreta para copiar --------
        If SheetNom <> "" Then
            SheetIndx = 0
            For Each Ws In ClosedBook.Sheets
                If Ws.Name = SheetNom Then
                    SheetIndx = Ws.Index
                    Exit For
                End If
            Next Ws
            If SheetIndx = 0 Then
                MsgBx_Title = "Proceso: Importar " & ArchRequest
                MsgBx_Msg = "¡ La Sheet NO existe !    "
                MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
                Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
'                With ActivForm.Controls("TBx_Informe")
'                    .Value = MsgBx_Msg & vbLf & Now()
'                    ActivForm.Repaint
'                End With
                Arch_New_Name = "Cancel"
                ClosedBook.Close SaveChanges:=False
                Set ClosedBook = Nothing
                Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
                Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
                Rut_On_Functions
                Exit Sub
            End If
            
        End If
        '-  Crear si NO Existe ListObject Lo_ClsBk ---------------------------------
        Dim Ws_ClsBk    As Worksheet
        Set Ws_ClsBk = ClosedBook.Sheets(SheetIndx)
        Dim Lo_ClsBk    As ListObject
        If Ws_ClsBk.ListObjects.Count = 0 Then
                Ws_ClsBk.UsedRange.Cells(1, 1).Select  'Posicionar cursor
            Set Lo_ClsBk = Ws_ClsBk.ListObjects.Add(xlSrcRange, Ws_ClsBk.UsedRange, , xlYes)
        Else
            Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
        End If
        
        With Ws_ClsBk
            .Unprotect
            .Columns.EntireColumn.Hidden = False        '-1º Mostrar todas las Columnas
            .Rows.EntireRow.Hidden = False              '-2º Mostrar todas las Filas
            Call Rut_Lo_Filtros_Quitar(.ListObjects(1)) '-3º Quitar Filtros
        End With
        Lo_ClsBk.ShowTotals = False
        '- Comprobar que la Tabla tiene datos ---------------------
        If Lo_ClsBk.ListRows.Count = 0 Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ La tabla no contiene datos !    "
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
'            With ActivForm.Controls("TBx_Informe")
'                .Value = MsgBx_Msg & vbLf & Now()
'                ActivForm.Repaint
'            End With
            Arch_New_Name = "Cancel"
            ClosedBook.Close SaveChanges:=False
            Set ClosedBook = Nothing
            Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
            Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
            Rut_On_Functions
            Exit Sub
        End If
        '- Comprobar que la cabecera de la Tabla corresponde con la establecida en la LoDefCol ---------------------
        If Not Func_LstObj_HeaderRow_Check_2Lo_OK(Lo_Data, Lo_ClsBk) Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ La cabedera de la tabla no coincide !    " & vbLf & "Puede haber un cambio en la estructura de la Tabla"
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
'            With ActivForm.Controls("TBx_Informe")
'                .Value = MsgBx_Msg & vbLf & Now()
'                ActivForm.Repaint
'            End With
            Arch_New_Name = "Cancel"
            ClosedBook.Close SaveChanges:=False
                    Set ClosedBook = Nothing
                    Set Ws_ClsBk = Nothing
                    Set Lo_ClsBk = Nothing
            Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
            Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
            Rut_On_Functions
            Exit Sub
        End If
        
        '- La Tabla Lo_Data SÍ tiene datos -Y- las Columnas coinciden. ---------------------
        ClosedBook.Sheets(SheetIndx).ListObjects(1).DataBodyRange.Copy
        Lo_Data.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues    'xlPasteAll    xlPasteValues
        Application.CutCopyMode = False
        ClosedBook.Close SaveChanges:=False
                Set ClosedBook = Nothing
                Set Ws_ClsBk = Nothing
                Set Lo_ClsBk = Nothing
        Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
            '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
                Dim TimeLap2              As Single
                TimeLap2 = LastTimeLap
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importado: " & NomArch, 0, , , TxT_Progreso)
                LastTimeLap = TimeLap2
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importados nuevos datos: ", LastTimeLap, " ", Format(Lo_Data.ListRows.Count, "#,##0") & " reg.")
            
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe").Text

    Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
Rut_On_Functions
Debug.Print "<<< Rut_Lo_Import_LoData_LoDefCol"
End Sub







'- ----------------------------------------------------------------------------------------------------------------------------
'- Seleccionar Excel e importarlo ----------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Import_LoData_LoDefCol(Lo_Data As ListObject, _
                                  LoDefCol As ListObject, _
                                  Col_Header As Integer, _
                                  Arch_New_Name As String, _
                                  Optional SheetNom As String = "", _
                                  Optional ByVal Convertir_En_Ram As Boolean = False)
'- Convertir_En_Ram (2026-10-06, fase 4 del paso a RAM): en vez de copiar y pegar valores (14,6 de los 28 seg. de importar y
'- formatear el LSGES04 del Robot), lee el fichero en RAM, convierte allí las Col. N y F como el formateo
'- (Rut_Ram_Textos_a_Numeros_y_Fechas) y escribe la tabla de una vez (Rut_Lo_Escribir_Importados). El llamador formatea
'- después con Convertir:=False. Mismo resultado celda a celda, en valor y tipo. Sin Convertir_En_Ram, se copia y pega como siempre.
                             
Debug.Print ">>> Rut_Lo_Import_LoData_LoDefCol"
Rut_Off_Functions
    Dim rowfind             As Variant
    Dim SheetIndx           As Integer:         SheetIndx = 1
    Dim NombresValidos      As String:          NombresValidos = Arch_New_Name                      '- Nombre(s) admitidos, separados por "|"
    Dim ArchRequest         As String:          ArchRequest = Replace(Arch_New_Name, "|", " o ")    '- Para mostrarlos en los avisos
    Dim TxT_Progreso        As String
    Dim NomArch             As String
    Dim AnoCont             As String:          AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim WrkSht              As Worksheet:       Set WrkSht = Lo_Data.Parent
    Dim Datos               As Variant                                  '- Convertir_En_Ram: las filas del fichero
    Dim Reescrita()         As Boolean                                  '- Convertir_En_Ram: Col. que ha cambiado la conversión
    Dim T_Paso              As Single                                   '- Tiempos de cada paso, para el informe
    Dim T_Abrir             As Single
    Dim T_Leer              As Single
    Dim T_Convertir         As Single
    Dim T_Escribir          As Single

    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
        
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Seleccionar el fichero y la ruta, para importar: " & ArchRequest, 0)
    '- Select File -------------------------------------------------------------------------------------
    Arch_New_Name = Fnc_Nombres_Prefijo_Comun(NombresValidos)                  '- El diálogo filtra por la parte común de los nombres
    Call Rut_File_Select("Seleccionar el Nuevo Fichero Excel " & ArchRequest & ": ", Arch_New_Name, "Excel", "*.xls?")
        If Arch_New_Name = "Cancel" Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ Cancelado a petición del Usuario !    "
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
'            With ActivForm.Controls("TBx_Informe")
'                .Value = MsgBx_Msg & vbLf & Now()
'                ActivForm.Repaint
'            End With
            Rut_On_Functions
            Exit Sub
        End If
        NomArch = Dir(Arch_New_Name)
        '- Comprueba que se ha seleccionado el nombre adecuado de Excel. --------------------------
        If Not Fnc_Nombre_Fichero_Valido(NomArch, NombresValidos) Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ Cancelado, el fichero debe ser un " & ArchRequest & " * !" & vbLf & vbLf & "  y se ha seleccionado:  " & NomArch
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & vbLf & Now()
'            With ActivForm.Controls("TBx_Informe")
'                .Value = MsgBx_Msg & vbLf & Now()
'                ActivForm.Repaint
'            End With
            Arch_New_Name = "Cancel"
            Rut_On_Functions
            Exit Sub
        End If
            
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
''''            ActivForm.Controls("Lb_Tit_Informe").Caption = "Progreso Tarea: Importar Informe: " & NomArch
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Excel Seleccionado: " & NomArch, LastTimeLap)
            TxT_Progreso = ActivForm.Controls("TBx_Informe")
    
    '- --------------------------------------------------------------------------------------------------------------
    '   Borrar el contenido de la Tabla -----------------------------------------------------------------------------
    '- Sin la fila de totales: con ella visible, el PasteSpecial de más abajo (Lo_Data.Range.Offset(1, 0)) la pisaría y dejaría
    '- los datos fuera de la tabla (el fallo de Lo_TPV en JyC). Se deja como estaba en cada salida (Docs/Plan_ShowTotals.md, fase 2).
    Dim Totales_Visibles    As Boolean:     Totales_Visibles = Fnc_Lo_Totales_Ocultar(Lo_Data)
    If Not Lo_Data.DataBodyRange Is Nothing Then Lo_Data.DataBodyRange.Delete
    Call Rut_WrkSheet_LstObj_LiberarEspacio(WrkSht)
            '- Visualizo el progreso  <<<<>>>>  -----------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Base de Datos vaciada.", LastTimeLap)
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Copy File Without Opening it  ------------------------------------------------------------------
    '- --------------------------------------------------------------------------------------------------------------
            '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importando: " & NomArch, 0, , , , , , 4)
    Prog__APP_Switch.Range("Sw_WB_Deactivate") = False     '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ---->>>
    Dim Ws              As Worksheet
    Dim ClosedBook      As Workbook
    T_Paso = Timer
    On Error Resume Next
    Set ClosedBook = Workbooks.Open(Arch_New_Name, ReadOnly:=True)
    On Error GoTo 0
    T_Abrir = Timer - T_Paso
    If ClosedBook Is Nothing Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡No se ha podido abrir el fichero!" & vbLf & _
                    "Puede que ya está abierto un libro con el mismo nombre en esta sesión de Excel," & vbLf & _
                    "o que la ruta no sea accesible: " & vbLf & Arch_New_Name
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        Arch_New_Name = "Cancel"
        Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
        Rut_On_Functions
        Exit Sub
    End If
        '- Comprobar que la Hoja Existe SI hemos solicitado una hoja concreta para copiar --------
        If SheetNom <> "" Then
            SheetIndx = 0
            For Each Ws In ClosedBook.Sheets
                If Ws.Name = SheetNom Then
                    SheetIndx = Ws.Index
                    Exit For
                End If
            Next Ws
            If SheetIndx = 0 Then
                MsgBx_Title = "Proceso: Importar " & ArchRequest
                MsgBx_Msg = "¡ La Sheet NO existe !    "
                MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
                Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
'                With ActivForm.Controls("TBx_Informe")
'                    .Value = MsgBx_Msg & vbLf & Now()
'                    ActivForm.Repaint
'                End With
                Arch_New_Name = "Cancel"
                ClosedBook.Close SaveChanges:=False
                Set ClosedBook = Nothing
                Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
                Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
                Rut_On_Functions
                Exit Sub
            End If
            
        End If
        '-  Crear si NO Existe ListObject Lo_ClsBk ---------------------------------
        Dim Ws_ClsBk    As Worksheet
        Set Ws_ClsBk = ClosedBook.Sheets(SheetIndx)
        Dim Lo_ClsBk    As ListObject
        If Ws_ClsBk.ListObjects.Count = 0 Then
                Ws_ClsBk.UsedRange.Cells(1, 1).Select  'Posicionar cursor
            Set Lo_ClsBk = Ws_ClsBk.ListObjects.Add(xlSrcRange, Ws_ClsBk.UsedRange, , xlYes)
        Else
            Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
        End If
        Ws_ClsBk.Unprotect
        Ws_ClsBk.Columns.EntireColumn.Hidden = False
        Ws_ClsBk.Rows.EntireRow.Hidden = False
        Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
        Lo_ClsBk.ShowTotals = False
        '- Comprobar que la Tabla tiene datos ---------------------
        If Lo_ClsBk.ListRows.Count = 0 Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ La tabla no contiene datos !    "
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
'            With ActivForm.Controls("TBx_Informe")
'                .Value = MsgBx_Msg & vbLf & Now()
'                ActivForm.Repaint
'            End With
            Arch_New_Name = "Cancel"
            ClosedBook.Close SaveChanges:=False
            Set ClosedBook = Nothing
            Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
            Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
            Rut_On_Functions
            Exit Sub
        End If
        '- Comprobar que la cabecera de la Tabla corresponde con la establecida en la LoDefCol ---------------------
        If Not Func_LstObj_ListColumns_DefCol_Check_OK(Lo_ClsBk, LoDefCol, Col_Header) Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ La cabedera de la tabla no coincide !    " & vbLf & "Puede haber un cambio en la estructura de la Tabla"
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
'            With ActivForm.Controls("TBx_Informe")
'                .Value = MsgBx_Msg & vbLf & Now()
'                ActivForm.Repaint
'            End With
            Arch_New_Name = "Cancel"
            ClosedBook.Close SaveChanges:=False
                    Set ClosedBook = Nothing
                    Set Ws_ClsBk = Nothing
                    Set Lo_ClsBk = Nothing
            Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
            Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
            Rut_On_Functions
            Exit Sub
        End If
        
        '- Robot_PPub_Fusión añade al final de sus ficheros una Col. ORIGEN (el fichero del Robot de donde sale cada fila).
        '- No es un dato del informe: si se copiara, caería encima de la primera Col. calculada de la tabla de destino.
        Dim Lc_Origen       As ListColumn
        Dim Quitada_Origen  As Boolean
        On Error Resume Next
        Set Lc_Origen = Lo_ClsBk.ListColumns("ORIGEN")
        On Error GoTo 0
        If Not Lc_Origen Is Nothing Then
            Lc_Origen.Delete
            Set Lc_Origen = Nothing
            Quitada_Origen = True
        End If

        '- La Tabla Lo_Data SÍ tiene datos -Y- las Columnas coinciden. ---------------------
        If Convertir_En_Ram Then
            T_Paso = Timer
            Datos = Lo_ClsBk.DataBodyRange.Value2                       '- Ya sin la Col. ORIGEN
            T_Leer = Timer - T_Paso
        Else
            ClosedBook.Sheets(SheetIndx).ListObjects(1).DataBodyRange.Copy
            Lo_Data.Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues    'xlPasteAll    xlPasteValues
            Application.CutCopyMode = False
        End If
        ClosedBook.Close SaveChanges:=False
                Set ClosedBook = Nothing
                Set Ws_ClsBk = Nothing
                Set Lo_ClsBk = Nothing
        Prog__APP_Switch.Range("Sw_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        If Convertir_En_Ram Then
            T_Paso = Timer
            Call Rut_Ram_Textos_a_Numeros_y_Fechas(Datos, LoDefCol, Reescrita)
            T_Convertir = Timer - T_Paso
            T_Paso = Timer
            Call Rut_Lo_Escribir_Importados(Lo_Data, Datos, Reescrita)
            T_Escribir = Timer - T_Paso
            Datos = Empty                                               '- Libera la memoria
        End If
            '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
                Dim TimeLap2              As Single
                TimeLap2 = LastTimeLap
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importado: " & NomArch, 0, , , TxT_Progreso)
                LastTimeLap = TimeLap2
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Importados nuevos datos: ", LastTimeLap, " ", Format(Lo_Data.ListRows.Count, "#,##0") & " reg.")
            If Quitada_Origen Then Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Quitada la Col. ORIGEN que añade Robot_PPub_Fusión.", 0)
            If Convertir_En_Ram Then
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "   Importar: abrir el fichero", 0, Format(T_Abrir, "0.00") & " seg.")
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "   Importar: leer las filas en RAM", 0, Format(T_Leer, "0.00") & " seg.")
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "   Importar: convertir en RAM las Col. N y F", 0, Format(T_Convertir, "0.00") & " seg.")
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "   Importar: escribir en la tabla", 0, Format(T_Escribir, "0.00") & " seg.")
            End If
            
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe").Text

    Call Rut_Lo_Totales_Restaurar(Lo_Data, Totales_Visibles)
Rut_On_Functions
Debug.Print "<<< Rut_Lo_Import_LoData_LoDefCol"
End Sub
'- ----------------------------------------------------------------------------------------------------------------------------

'- ----------------------------------------------------------------------------------------------------------------------------
'- Escribe en Lo_Data las filas leídas de un fichero (Datos, ya convertidas en RAM), con el mismo resultado que pegar valores y
'- convertir después en la hoja (Rut_Lo_Format_LoData_LoDefColData):
'-      - las Col. Reescritas (cambiadas por la conversión) se escriben con formato General, como las reescribía el formateo:
'-        Excel interpreta los textos que queden en ellas
'-      - las demás, con formato texto ("@") mientras se escriben: así Excel guarda los textos tal cual ("00123", "1/2", "=1+1"),
'-        como el pegado, y los números siguen siendo números (comprobado el 2026-10-06). El formateo de después les quita el "@".
'- Cada grupo de Col. seguidas con el mismo trato va en una sola escritura.
'- ----------------------------------------------------------------------------------------------------------------------------
Private Sub Rut_Lo_Escribir_Importados(Lo_Data As ListObject, Datos As Variant, Reescrita() As Boolean)
    Dim WrkSht      As Worksheet:   Set WrkSht = Lo_Data.Parent
    Dim NumFilas    As Long:        NumFilas = UBound(Datos, 1)
    Dim NumCols     As Long:        NumCols = Application.Min(UBound(Datos, 2), Lo_Data.ListColumns.Count)
    Dim Bloque()    As Variant
    Dim Rng         As Range
    Dim C           As Long
    Dim C2          As Long
    Dim K           As Long
    Dim Fila        As Long

    With Lo_Data.HeaderRowRange                             '- La tabla, con tantas filas como el fichero
        Lo_Data.Resize WrkSht.Range(.Cells(1, 1), .Cells(1, .Columns.Count).Offset(NumFilas, 0))
    End With
    Lo_Data.DataBodyRange.Resize(, NumCols).NumberFormat = "General"
    C = 1
    Do While C <= NumCols
        C2 = C                                              '- Hasta dónde llegan las Col. seguidas con el mismo trato
        Do While C2 < NumCols
            If Reescrita(C2 + 1) <> Reescrita(C) Then Exit Do
            C2 = C2 + 1
        Loop
        ReDim Bloque(1 To NumFilas, 1 To C2 - C + 1)
        For K = C To C2
            For Fila = 1 To NumFilas
                Bloque(Fila, K - C + 1) = Datos(Fila, K)
            Next Fila
        Next K
        Set Rng = Lo_Data.DataBodyRange.Columns(C).Resize(, C2 - C + 1)
        If Not Reescrita(C) Then Rng.NumberFormat = "@"
        Rng.Value2 = Bloque
        C = C2 + 1
    Loop
End Sub     ' Rut_Lo_Escribir_Importados
'- ----------------------------------------------------------------------------------------------------------------------------

'- ----------------------------------------------------------------------------------------------------------------------------
'- Nombres de fichero que admite una importación: uno o varios prefijos separados por "|", p.ej. el nombre de siempre y el de
'- Robot_PPub_Fusión ("LSGES04_GE_SinDtos_Año_2026|LsGes04_AñoCont_2026"). Se comparan sin distinguir mayúsculas, como Windows.
'- ----------------------------------------------------------------------------------------------------------------------------
Function Fnc_Nombre_Fichero_Valido(ByVal NomArch As String, ByVal NombresValidos As String) As Boolean
    Dim Nombre      As Variant
    For Each Nombre In Split(NombresValidos, "|")
        If Nombre <> "" Then
            If LCase$(Left$(NomArch, Len(Nombre))) = LCase$(Nombre) Then
                Fnc_Nombre_Fichero_Valido = True
                Exit Function
            End If
        End If
    Next Nombre
End Function    ' Fnc_Nombre_Fichero_Valido

'- Parte inicial común de los nombres admitidos (sin distinguir mayúsculas), para el filtro del diálogo de selección: --------
'- "LSGES04_GE_SinDtos_Año_2026|LsGes04_AñoCont_2026" -> "LSGES04_", que en el diálogo encuentra los dos.
Function Fnc_Nombres_Prefijo_Comun(ByVal NombresValidos As String) As String
    Dim Nombres     As Variant
    Dim Prefijo     As String
    Dim i           As Long
    If NombresValidos = "" Then Exit Function
    Nombres = Split(NombresValidos, "|")
    Prefijo = Nombres(0)
    For i = 1 To UBound(Nombres)
        Do While LCase$(Left$(Nombres(i), Len(Prefijo))) <> LCase$(Prefijo)
            Prefijo = Left$(Prefijo, Len(Prefijo) - 1)
        Loop
    Next i
    Fnc_Nombres_Prefijo_Comun = Prefijo
End Function    ' Fnc_Nombres_Prefijo_Comun
'- ----------------------------------------------------------------------------------------------------------------------------
