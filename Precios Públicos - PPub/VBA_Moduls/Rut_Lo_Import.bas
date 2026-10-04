Attribute VB_Name = "Rut_Lo_Import"
' Last Rev. 2026-10-03 23:51
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
            
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe")

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
                                  Optional SheetNom As String = "")
                             
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
            If Quitada_Origen Then Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Quitada la Col. ORIGEN que añade Robot_PPub_Fusión.", 0)
            
    Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe")

Rut_On_Functions
Debug.Print "<<< Rut_Lo_Import_LoData_LoDefCol"
End Sub
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
