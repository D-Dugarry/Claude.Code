Attribute VB_Name = "Rut_WS"
' Last Rev. 2026-09-21 19:05
Option Explicit

' ==================================================================================================
Function Fnc_WrkSheet_Exist(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    Fnc_WrkSheet_Exist = Not ws Is Nothing
End Function
'###################################################################################################
Sub Rut_WrkSheet_ReducirPeso(ByVal WrkSht As String, Optional Sw_Del_DataBodyRange As Boolean = False) '--- Borra TODO a la Derecha y Abajo de .ListObjects(1)
' ==================================================================================================
        Dim ws      As Worksheet:       Set ws = Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
        Dim Lo      As ListObject:      Set Lo = ws.ListObjects(1)
        Dim FilIni  As Long             '- Primera FILA a borrar (la siguiente a la Tabla)
        Dim ColIni  As Long             '- Primera COLUMNA a borrar (la siguiente a la Tabla)
        '- Ojo: se guardan como numero, NO como Range. Un objeto Range que apunte a la primera
        '-  fila tras la tabla queda INVALIDADO al borrar esa fila, y leer su .Address da
        '-  Error 424. Pasa cuando la tabla esta casi vacia (Sw_Del_DataBodyRange:=True).
        Dim Dummy_UsedRange As String   '- Solo para forzar la lectura de UsedRange (ver mas abajo)
        Dim Sw_Calculation  As XlCalculation:  Sw_Calculation = Application.Calculation:   Application.Calculation = xlManual
    With ws
        .Columns.EntireColumn.Hidden = False    ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False          ' Mostrar todas las Filas
        Call Rut_Lo_Filtros_Quitar(Lo)          ' Deshacer Filtros
        If Sw_Del_DataBodyRange And Not Lo.DataBodyRange Is Nothing Then Lo.DataBodyRange.Delete
        With Lo.Range                           ' Borra la filas de abajo y columnas de la derecha del la Tabla .ListObjects(1)
            FilIni = .Row + .Rows.Count         ' Primera fila tras la ultima de la tabla
            ColIni = .Column + .Columns.Count   ' Primera columna tras la ultima de la tabla
        End With
        If FilIni <= .Rows.Count Then _
            .Range(.Cells(FilIni, 1), .Cells(.Rows.Count, 1)).EntireRow.Delete
        If ColIni <= .Columns.Count Then _
            .Range(.Cells(1, ColIni), .Cells(1, .Columns.Count)).EntireColumn.Delete
        Dummy_UsedRange = .UsedRange.Address        ' Para restablecer el rango de celdas en uso (hay que LEER la propiedad para que surta efecto)
    End With
    Application.Calculation = Sw_Calculation
End Sub
' --------------------------------------------------------------------------------------------------
'###################################################################################################
Sub Rut_WrkSheet_Preparar(WrkSht As Worksheet)  '- Mostrar todas las Filas y Columnas, y Quitar Filtros.
' --------------------------------------------------------------------------------------------------
    With WrkSht
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        Call Rut_Lo_Filtros_Quitar(.ListObjects(1)) ' Deshacer Filtros
    End With
End Sub
' ==================================================================================================
Sub Rut_WrkSheet_Vaciar(WrkSht As Worksheet)      '--- Borra Toda la Hoja incluso los objetos (Shapes)
' ==================================================================================================
    Dim WrkSht_Activa    As String:     WrkSht_Activa = ActiveSheet.Name
    Dim Dummy_UsedRange  As String      '- Solo para forzar la lectura de UsedRange (ver mas abajo)
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual

        Application.ScreenUpdating = False
        
    With WrkSht
            Dim Visual_Status   As Variant:  Visual_Status = .Visible   '--- para dejar la hoja en el mismo estado de Visibilidad ---
            Dim Protect_Status   As Boolean:  Protect_Status = .ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
        .Visible = xlHidden
        .Unprotect
        Call Rut_WrkSheet_Preparar(WrkSht)
        ' Limpiar filtros de todas las tablas
        Dim tbl As ListObject
        For Each tbl In .ListObjects
            If tbl.ShowAutoFilter Then tbl.AutoFilter.ShowAllData   '- Quitar filtro Tabla
        Next tbl
            .Columns.Delete     ' --- con esto se borran hasta los "Shapes"
            Dummy_UsedRange = .UsedRange.Address   ' Para restablecer el rango de celdas en uso (hay que LEER la propiedad para que surta efecto)
        .Visible = Visual_Status
        If Protect_Status Then .Protect
        Sheets(WrkSht_Activa).Select
        
    End With
        Application.Calculation = Sw_Calculation
        Application.ScreenUpdating = True
End Sub
' --------------------------------------------------------------------------------------------------


' ==================================================================================================
    Sub RuT_AllSheets_Crear_Lista()    ' Manejo interno, Hace la lista de todas las hojas de este libro
' ==================================================================================================
Dim WrkSht      As Worksheet
Dim NomHoja     As String
Dim NewRow      As ListRow
Application.ScreenUpdating = False
With Prog_HojasName
    .Unprotect
        ' --------------------------------=============  Preparar Tabla de DR_Unificada ============
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
' --------------------------------------------------------------------------------------------------

' ==================================================================================================
    Sub RuT_AllSheets_Visible_OrNot()    ' Estable la visibilidad establecida de cada Sheet en la Tabla(2) de Prog_Hojas_Names
' ==================================================================================================
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
' --------------------------------------------------------------------------------------------------

            ' ======================================================================================
            Sub RuT_Sort_Sheets_ByHand()
                RuT_Sort_Sheets
            End Sub
' ==================================================================================================
Sub RuT_Sort_Sheets(Optional Sort_Ascending As Boolean = True)
' ==================================================================================================
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
' --------------------------------------------------------------------------------------------------



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


'---------------------------------------------------------------------------------------------------
      
        ' ==========================================================================================
        Sub Rut_WrkSheet_To_PDF_ByHand()
            Rut_WrkSheet_To_PDF (ActiveSheet.Name)
        End Sub
' ==================================================================================================
Sub Rut_WrkSheet_To_PDF(ByVal WrkSht As String)     '   Exportar WorkSheet a PDF  ------------------
' ==================================================================================================
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
'---------------------------------------------------------------------------------------------------


'===================================================================================================
'- Exporta Copia de Sheet en Excel.xlsx ============================================================
'===================================================================================================
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
'===================================================================================================



