Attribute VB_Name = "Rut_WS"
Option Explicit

' ==================================================================================================================================
Function Fnc_WrkSheet_Exist(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    Fnc_WrkSheet_Exist = Not ws Is Nothing
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
            '###################################################################################################################################
                    Sub Rut_WrkSheet_ReducirPeso_xx()
'                        Application.Workbooks(ThisWorkbook.Name).Sheets(ActiveSheet.Name).ListObjects(1).DataBodyRange.Delete
                        Call Rut_WrkSheet_ReducirPeso(ActiveSheet.Name)
                        Debug.Print ActiveSheet.Name
                    End Sub
'###################################################################################################################################
Sub Rut_WrkSheet_ReducirPeso(ByVal WrkSht As String, Optional Sw_Del_DataBodyRange As Boolean = False) '--- Borra TODO a la Derecha y Abajo de .ListObjects(1) ----------------------
' ==================================================================================================================================
        Dim ws      As Worksheet:       Set ws = Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
        Dim Lo      As ListObject:      Set Lo = ws.ListObjects(1)
        Dim Sw_Calculation  As Boolean:  Sw_Calculation = Application.Calculation:   Application.Calculation = xlManual
    With ws
        .Columns.EntireColumn.Hidden = False    ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False          ' Mostrar todas las Filas
        Call Rut_Lo_Filtros_Quitar(Lo)          ' Deshacer Filtros
        If Sw_Del_DataBodyRange And Not Lo.DataBodyRange Is Nothing Then Lo.DataBodyRange.Delete
        With Lo.Range                           ' Borra la filas de abajo y columnas de la derecha del la Tabla .ListObjects(1)
            Range(.Cells(.Rows.Count, .Columns.Count).Address).Select   ' Selecciona la última celda de la tabla
            ActiveCell.Offset(1, 1).Select
        End With
        .Range(ActiveCell.Address & ":" & Cells(Rows.Count, 1).Address).EntireRow.Delete
        .Range(ActiveCell.Address & ":" & Cells(1, Columns.Count).Address).EntireColumn.Delete
        ActiveSheet.UsedRange                       ' Para restablecer el rango de celdas en uso
    End With
    Application.Calculation = Sw_Calculation
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
'###################################################################################################################################
Sub Rut_WrkSheet_Preparar(WrkSht As Worksheet)  '- Mostrar todas las Filas y Columnas, y Quitar Filtros.
' ----------------------------------------------------------------------------------------------------------------------------------
    With WrkSht
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        Call Rut_Lo_Filtros_Quitar(.ListObjects(1)) ' Deshacer Filtros
    End With
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<


            '###################################################################################################################################
                    Sub Rut_WrkSheet_Vaciar_xx()
                        Call Rut_WrkSheet_Vaciar(ActiveSheet.Name)
                    End Sub
' ==================================================================================================================================
Sub Rut_WrkSheet_Vaciar(ByVal WrkSht As String)      '--- Borra Toda la Hoja incluso los objetos (Shapes)  -------------------------------
' ==================================================================================================================================
    Dim WrkSht_Activa    As String:     WrkSht_Activa = ActiveSheet.Name
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual

        Application.ScreenUpdating = False
        
    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
            Dim Visual_Status   As Variant:  Visual_Status = .Visible   '--- para dejar la hoja en el mismo estado de Visibilidad ---
            Dim Protect_Status   As Boolean:  Protect_Status = .ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
        .Visible = xlHidden
        .Unprotect
        Call Rut_WrkSheet_Preparar(WrkSht)
        ' Limpiar filtros de todas las tablas
        Dim tbl As ListObject
        For Each tbl In .ListObjects
            If .ListObjects(1).ShowAutoFilter Then .ListObjects(1).AutoFilter.ShowAllData   '- Quitar filtro Tabla
        Next tbl
            .Columns.Delete     ' --- con esto se borran hasta los "Shapes"
            .UsedRange
        .Visible = Visual_Status
        If Protect_Status Then .Protect
        Sheets(WrkSht_Activa).Select
        
    End With
        Application.Calculation = Sw_Calculation
        Application.ScreenUpdating = True
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<


' ==================================================================================================================================
    Sub RuT_AllSheets_Crear_Lista()    ' Manejo interno, Hace la lista de todas las hojas de este libro       =============================
' ==================================================================================================================================
Dim WrkSht      As Worksheet
Dim NomHoja     As String
Dim NewRow      As ListRow
Application.ScreenUpdating = False
With Prog_HojasName
    .Unprotect
        ' --------------------------------=============  Preparar Tabla de DR_Unificada ==================
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        If .FilterMode Then .ShowAllData            ' Deshacer Filtros
        If Not .ListObjects("Tb_Sheets_Config").DataBodyRange Is Nothing Then .ListObjects("Tb_Sheets_Config").DataBodyRange.Delete
    With .ListObjects("Tb_Sheets_State")
        If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
        For Each WrkSht In Worksheets
            Set NewRow = .ListRows.Add
            With NewRow
                .Range(1) = WrkSht.CodeName
                .Range(2) = WrkSht.Name
                .Range(3) = WrkSht.Visible
                Select Case .Range(3)
                    Case -1
                        .Range(4) = "Visible"
                    Case 0
                        .Range(4) = "Hidden"
                    Case 2
                        .Range(4) = "VeryHidden"
                End Select
            End With ' NewRow
        Next
        .Range.Sort key1:=.ListColumns(1), order1:=xlAscending, _
            Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
'        .DataBodyRange.Copy
'        Prog_HojasName.ListObjects("Tb_Sheets_Config").Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
'        Application.CutCopyMode = False
    End With    ' .ListObjects("Tb_Sheets_State")
    Call Rut_Lo_DataBodyRange_Copy(Prog_HojasName.ListObjects("Tb_Sheets_State"), Prog_HojasName.ListObjects("Tb_Sheets_Config"), True)
    .Protect
End With ' Prog_HojasName
Application.ScreenUpdating = True
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
    Sub RuT_AllSheets_Visible_OrNot()    ' Estable la visibilidad establecida de cada Sheet en la Tabla(2) de Prog_Hojas_Names
' ==================================================================================================================================
Dim Cont        As Integer
Dim NomHoja     As String
On Error Resume Next    ' por si ha desaparecido una hoja ----
With Prog_HojasName.ListObjects("Tb_Sheets_Config").DataBodyRange
    For Cont = 1 To .Rows.Count
        NomHoja = .Cells(Cont, 2)
        Sheets(NomHoja).Visible = Val(.Cells(Cont, 3))
    Next
End With
On Error GoTo 0
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

            ' ======================================================================================================================
            Sub RuT_Sort_Sheets_ByHand()
                RuT_Sort_Sheets
            End Sub
' ==================================================================================================================================
Sub RuT_Sort_Sheets(Optional Sort_Ascending As Boolean = True)
' ==================================================================================================================================
Dim Pos_Outer  As Integer
Dim Pos_Inner  As Integer
Dim Cant_Sheets     As Integer:   Cant_Sheets = Sheets.Count

    For Pos_Outer = 1 To Cant_Sheets
        For Pos_Inner = 1 To Pos_Outer
            If Sort_Ascending Then
                If UCase(Sheets(Pos_Outer).Name) < UCase(Sheets(Pos_Inner).Name) Then
                        Sheets(Pos_Outer).Move Before:=Sheets(Pos_Inner)
                End If
            Else
                If UCase(Sheets(Pos_Outer).Name) > UCase(Sheets(Pos_Inner).Name) Then
                        Sheets(Pos_Outer).Move Before:=Sheets(Pos_Inner)
                End If
            End If
        Next Pos_Inner
    Next Pos_Outer
    
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''RUTINAS PARA HACER GENÉRICAS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''FALTA ADAPTARLAS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' ==================================================================================================================================
Sub Rutxxxx_Exportar_La_Liquidación()   '- Copia una Sheet concreta
' ==================================================================================================================================
Rut_Off_Functions
Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Copio la Sheet entera y esto es lo que voy a grabar. -----------------------------------------
    Wk_TitP_Liquid.Copy
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "Liquid_" & Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo GestError
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True, FileFormat:=51
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    '- Quito los Elementos: Botones (Shapes), Comentarios de Celdas y Borro la Fila de Filtrado (la de arriba de los títulos de la Tabla ------------
        ActiveWorkbook.ActiveSheet.Unprotect
        ActiveWorkbook.ActiveSheet.Shapes.SelectAll:   Selection.Delete
        ActiveWorkbook.ActiveSheet.UsedRange.ClearComments
        Application.EnableEvents = False
        ActiveWorkbook.ActiveSheet.Rows(ActiveWorkbook.ActiveSheet.ListObjects(1).Range.Rows(1).Row - 1).Clear
    '- Grabo los cambios y Cierro el Archivo ---------------------------
    ActiveWorkbook.Close SaveChanges:=True
        MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Archivar Liquidación"
    Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & _
            "  -.-  " & Now() & vbCrLf & vbCrLf & "Exportado el Resumen de la Liquidación de:   " & Wk_TitP_Liquid.Range("Liquid_Plan_Name") & _
            vbCrLf & vbCrLf & "En el Archivo:   " & sFileSaveName
    GoTo Salir_Sub
GestError:
    Debug.Print "Error Rut_Exportar_La_Liquidación ", Err.Number, Err.Description, Err.Source
    Debug.Print sFileSaveName
    MsgBox "Rut_Exportar_La_Liquidación " & "Filename:=" & vbCrLf & sFileSaveName, vbExclamation + vbOKOnly, "Rutina de Remesado"
    Form_Menu.TB_Informe = "Rut_Exportar_La_Liquidación " & vbCrLf & "Filename:=" & sFileSaveName
Salir_Sub:
Call Rut_EnableEvents_Status_Reset
Rut_On_Functions
End Sub     ' Rut_Exportar_La_Liquidación
'-----------------------------------------------------------------------------------------------------------------------------------
      
        ' ==========================================================================================================================
        Sub Rut_WrkSheet_To_PDF_ByHand()
            Rut_WrkSheet_To_PDF (ActiveSheet.Name)
        End Sub
' ==================================================================================================================================
Sub Rut_WrkSheet_To_PDF(ByVal WrkSht As String)     '   Exportar WorkSheet a PDF  --------------------------------------------------
' ==================================================================================================================================
Rut_Off_Functions
Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Copio la Sheet entera -----------------------------------------
    Wk_TitP_Liquid.Copy
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "LIQ-TitProp_" & Wk_TitP_Liquid.Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".pdf"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.pdf), *.pdf")
        If sFileSaveName <> False Then
            On Error GoTo Restablecer_Valores
            Application.DisplayAlerts = False   '- Para Reemplazar si ya existe el Excel, sin preguntar.
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sFileSaveName
'            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True, FileFormat:=51
'                                                           FileFormat:=xlOpenXMLWorkbook               ' ???
'                                                           FileFormat:=xlOpenXMLWorkbookMacroEnabled   ' ???
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    
Restablecer_Valores:
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------


'===================================================================================================================================
'- Exporta Copia de Sheet en Excel.xlsx ============================================================================================
'===================================================================================================================================
Sub Rut_WrkSheet_Export_To_xlsx(SheetToCopy As Worksheet, FichName As String)
    Debug.Print "Rut_WrkSheet_Export_To_xlsx"
    
    Dim WbOrig          As Workbook
    Dim Wb_new          As Workbook
    Dim FichPath        As String
    Dim shp             As Shape
    Dim ws              As Worksheet
    Dim ExportFichName  As String
    
    Set WbOrig = ThisWorkbook
    
    '--- Copiar hoja activa a libro nuevo ---
    SheetToCopy.Copy
    Set Wb_new = ActiveWorkbook
    
    '--- Quitar shapes y comentarios del nuevo libro ---
    For Each ws In Wb_new.Worksheets
                On Error Resume Next
            ws.Unprotect
                On Error GoTo 0
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
        ws.UsedRange.ClearComments
    Next ws
    
    '--- Ruta base: la del libro actual, o la actual si está sin guardar ---
    If Len(WbOrig.Path) > 0 Then
        FichPath = WbOrig.Path & Application.PathSeparator
    Else
        FichPath = CurDir$ & Application.PathSeparator
    End If
    
    '--- Cuadro Guardar como ---
    Dim Filedlg         As FileDialog
    Set Filedlg = Application.FileDialog(msoFileDialogSaveAs)
    With Filedlg
        .InitialFileName = FichPath & FichName
        .Title = "Exportar Cierre Contable 2025 (sin macros)"
        .FilterIndex = 1   ' Excel Workbook (*.xlsx)
        If .Show <> -1 Then
            ' Usuario canceló
            Wb_new.Close SaveChanges:=False
            Exit Sub
        End If
        ExportFichName = .SelectedItems(1)
    End With
    
    ' Guardar como .xlsx (sin macros)
    Wb_new.SaveAs Filename:=ExportFichName, FileFormat:=xlOpenXMLWorkbook
    
    ' Actualizar variables de ruta y nombre según tu rutina auxiliar
    Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(ExportFichName, FichName, FichPath)
    
    ' Dejamos el libro nuevo ABIERTO y NO tocamos WbOrig
    ' Si quisieras cerrar solo el nuevo, usarías: Wb_new.Close SaveChanges:=True
    Wb_new.Close SaveChanges:=True
    
    MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, _
           "Proceso: Exportar Hoja Informe de Cierre Contable."
    
    Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & _
        Round(Timer - H_Inicio, 2) & " seg." & "  -.-  " & Now() & vbCrLf & vbCrLf & _
        "Exportada Hoja Informe de Cierre Contable:   " & vbCrLf & vbCrLf & _
        "En el Archivo:   " & FichName & vbCrLf & _
        "Ruta: " & FichPath
    
Finalizar:
End Sub
'===================================================================================================================================



