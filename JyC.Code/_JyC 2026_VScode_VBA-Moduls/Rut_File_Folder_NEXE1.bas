Attribute VB_Name = "Rut_File_Folder_NEXE1"
''''Option Explicit
'''''###################################################################################################################################
''''Function Fnc_NEXE_RutaAPP() As String    ' Formatea la Ruta del Excel Actual, dependiendo de la Red Nexe o del disco Local -------
''''' ==================================================================================================================================
''''    Fnc_NEXE_RutaAPP = Application.Workbooks(ThisWorkbook.Name).Path
''''    Debug.Print "Fnc_NEXE_RutaAPP: " & Fnc_NEXE_RutaAPP
''''    If Left(Fnc_NEXE_RutaAPP, 18) = "https://nexe.ua.es" Then
''''        Fnc_NEXE_RutaAPP = Replace(Fnc_NEXE_RutaAPP, "/", "\")                                        '- Cambio / por \
''''        Fnc_NEXE_RutaAPP = Mid(Fnc_NEXE_RutaAPP, InStr(1, Fnc_NEXE_RutaAPP, Prog__APP.Range("App_MailUsu")) + Len(Prog__APP.Range("App_MailUsu")))
''''        Fnc_NEXE_RutaAPP = Prog__APP.Range("App_LetraUnidRed") & ":" & Fnc_NEXE_RutaAPP                     '- Añado la letra de la Unidad
''''    End If
''''End Function        ' Fnc_Format_Referencia
'''''###################################################################################################################################
''''   'call Rut_File_Select ("Seleccionar el fichero de...", FichSelect, "Excel", "*.xls?")      '- FichSelect = "Cancel"
''''Sub Rut_File_Select(Título As String, ByRef FichSelect As String, Optional TipoFich_txto As String = "Cualquier Fichero", Optional TipoFich As String = "*.*")
''''' ==================================================================================================================================
''''    With Application.FileDialog(msoFileDialogFilePicker)
''''            .Title = Título
''''            .InitialFileName = Fnc_NEXE_RutaAPP & "\"
''''            .InitialView = msoFileDialogViewDetails
''''            .AllowMultiSelect = False
''''            .ButtonName = "Seleccionar" ' o "Aceptar" o ...
''''            .Filters.Clear
''''            .Filters.Add TipoFich_txto, TipoFich, 1
''''        If .Show = True Then
''''            FichSelect = .SelectedItems(1)     '- Nombre Arch CON Ruta
''''           'NomArch = dir(FichSelect)          '- Nombre Arch SIN Ruta
''''           'RutaArch = Left(FichSelect, Len(FichSelect) - Len(NomArch)) ' Extrae solo la ruta del directorio
''''        Else
''''            MsgBox "Ha pulsado el botón <Cancelar>.", vbOKOnly, "Proceso: " & Título
''''            FichSelect = "Cancel"
''''        End If
''''    End With
''''End Sub
'''''###################################################################################################################################
''''Sub Rut_Folder_Select(Título As String, ByRef Directorio As String)     ' Seleccionar una ruta (carpeta) del explorador
''''' ==================================================================================================================================
''''    With Application.FileDialog(msoFileDialogFolderPicker)
''''            .Title = Título
''''            .InitialFileName = Application.Workbooks(ThisWorkbook.Name).Path & "\"
''''            .InitialView = msoFileDialogViewDetails
''''            .AllowMultiSelect = False
''''            .ButtonName = "Seleccionar Carpeta" ' o "Aceptar" o ...
''''            .Filters.Clear
''''        If .Show = True Then
''''            Directorio = .SelectedItems(1)
''''        Else
''''            MsgBox "Ha pulsado el botón <Cancelar>.", vbOKOnly, "Proceso: " & Título
''''            Directorio = "Cancel"
''''        End If
''''    End With
''''End Sub
''''' -------------------------------------------------------------------------------------------------------------------------------<<<
''''
'''''==================================================================================================================================
''''Sub RuT_WorkBook_Sheet_1_Import_IN(WrkSht As String)   '- Importar sheets(1) from workbook selected in ThisWorBook.Sheets(WrkSht).
''''    '   Seleccionar fichero     ---------------------------------------------------------------------------------------------------
''''        Dim NomFich     As String
''''    Call Rut_File_Select("Seleccionar el Fichero Excel de la última consulta de pagos por TPV del Generador de Informes: ", NomFich, "Excel", "*.xls?")       '- NomFich = "Cancel"
''''    If NomFich = "Cancel" Then
''''        MsgBox "Operación Cancelada"
''''        Exit Sub
''''    End If
''''    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
''''        '   Si está VeryHidden....
''''        Dim Visual_Status   As Variant:  Visual_Status = .Visible   '--- para dejar la hoja en el mismo estado de Visibilidad ---
''''        .Visible = xlHidden
''''        If .ProtectContents Then .Unprotect
''''        .Visible = Visual_Status
''''        '   Borrar el contenido de la hoja que recibe los datos    ------------------------------------------------------------
''''        Call Rut_WrkSheet_Vaciar(WrkSht)
''''        '   Copiar WorkBook_Selected.Sheet(1) in ThisWorkBook.sheets(WrkSht)  ----------------------------------------------
''''        Dim closedBook          As Variant
''''        Set closedBook = Workbooks.Open(NomFich)
''''        closedBook.Sheets(1).UsedRange.Copy .Range("A1")
''''        Application.CutCopyMode = False
''''        closedBook.Close SaveChanges:=False
''''        Set closedBook = Nothing
''''        '   Si no viene con Tabla la Creo       ------------------------------------------------------------------------
''''        If .ListObjects.Count = 0 Then
''''           .ListObjects.Add(xlSrcRange, .UsedRange, , xlYes).Name = "Tb_xxx"
''''        End If
''''    End With
''''End Sub
'''''==================================================================================================================================
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''
''''' ==================================================================================================================================
''''Sub Rut_Import_All_Files_Folder()   '- Importa todos los ficheros de un directorio -------------------------
''''' ==================================================================================================================================
''''Dim FichSistOjct As New FileSystemObject
''''Dim Fichero As File
''''Dim FichDialog As FileDialog
''''Dim FichRuta As String
''''Dim wb As Workbook
''''Dim Ws As Worksheet
''''Dim aWS As Worksheet
''''
''''    Set aWS = ActiveSheet
''''
''''    Set FichDialog = Application.FileDialog(msoFileDialogFolderPicker)
''''    With FichDialog
''''         .Title = "Choose Folder where you have excel files"
''''         .ButtonName = "Choose"
''''         If .Show = True Then
''''            If .SelectedItems.Count > 0 Then
''''               FichRuta = .SelectedItems(1)
''''            End If
''''         End If
''''    End With
''''
''''    If FichRuta <> "" Then
''''       For Each Fichero In FichSistOjct.GetFolder(FichRuta).Files
''''           If Fichero.Name Like "*.xl??" Or Fichero.Name Like "*.xl?" Then
''''           Set wb = Workbooks.Open(Fichero.Path, False)
''''           Set Ws = wb.Sheets(1)
''''
''''           Lr = Ws.Range("A" & Rows.Count).End(xlUp).Row
''''           Ws.Range("A2:C" & Lr).Copy
''''           Lr = aWS.Range("A" & Rows.Count).End(xlUp).Row + 1
''''           aWS.Range("A" & Lr).PasteSpecial xlPasteAll
''''           Application.CutCopyMode = False
''''           wb.Close False
''''           End If
''''       Next Fichero
''''       MsgBox "All Files Imported successfully!", vbInformation
''''    End If
''''End Sub
''''
''''
''''' ==================================================================================================================================
''''Sub Rut_WrkBook_Sheets_Select_Inport()      '- Añadir la hoja 1 de cada Excel seleccionado -------------
''''' ==================================================================================================================================
''''Dim Última_Fila         As Long
''''Dim Cont_Col            As Long
''''Dim LibroActual     As String
''''Dim LibroNuevo      As String
''''
''''Dim HojaOrigen              As Worksheet
''''Dim HojaDestino             As Worksheet
''''Dim ResultadoExplorador         As FileDialog
''''Dim LibroXlsSelec           As Variant
''''
''''H_Resumen.Select
''''H_Resumen.Unprotect
''''Application.ScreenUpdating = False
''''
''''LibroActual = ActiveWorkbook.Name
''''Set HojaDestino = Workbooks(LibroActual).Sheets(1)
''''    Set ResultadoExplorador = Application.FileDialog(msoFileDialogFilePicker)
''''    With ResultadoExplorador
''''        .InitialFileName = ActiveWorkbook.Path & "\*.xls*"
''''        If .Show = -1 Then
''''            For Each LibroXlsSelec In .SelectedItems
''''                '- Abro el WorkBook y copio la Hoja(1) un una Nueva Hoja al final del Libro actual ------------------
''''                Workbooks.OpenXML FileName:=LibroXlsSelec, LoadOption:=xlXmlLoadImportToList
''''                LibroNuevo = ActiveWorkbook.Name
''''                Set HojaOrigen = ActiveWorkbook.Worksheets(1)
''''                '- Si Existe la hoja, la Borro para reemplazarla ----------------
''''                If Workbooks(LibroActual).Worksheets(HojaOrigen.Name).Name = HojaOrigen.Name Then
''''                    Application.DisplayAlerts = False
''''                    Workbooks(LibroActual).Worksheets(HojaOrigen.Name).Delete
''''                    Application.DisplayAlerts = True
''''                End If
''''                '- Copiar la Hoja ---------------------
''''                HojaOrigen.Copy After:=Workbooks(LibroActual).Sheets(Workbooks(LibroActual).Sheets.Count)
''''                Workbooks(LibroNuevo).Close SaveChanges:=False
''''                Application.CutCopyMode = False
''''                '- Creo la Tabla ------------------------------------------------------------------------
''''                ActiveSheet.ListObjects.Add(xlSrcRange, ActiveSheet.Cells(1, 1).CurrentRegion, , xlYes).Name = "Tb_TPV"
''''                '- Añado fila de totales a la Tabla ----------
''''                ActiveSheet.ListObjects(1).ShowTotals = True
''''                    ActiveSheet.ListObjects(1).ListColumns("Importe").TotalsCalculation = xlTotalsCalculationSum
''''                    ActiveSheet.ListObjects(1).ListColumns("Comision").TotalsCalculation = xlTotalsCalculationSum
''''                    ActiveSheet.ListObjects(1).ListColumns("Cobrado").TotalsCalculation = xlTotalsCalculationSum
''''                '- Añado un Hipervínculo a la hoja de Resumen ----------------------------------
''''                ActiveSheet.Hyperlinks.Add Anchor:=Range("j1"), Address:="", SubAddress:=H_Resumen.Name & "!A7", TextToDisplay:="Inicio"
''''                '- Configuro los Datos de Fecha-Final y Descripción ------------------
''''                ActiveWindow.Zoom = 85
''''                With Cells.Font
''''                    .Name = "Arial"
''''                    .Size = 12
''''                End With
''''                With Rows("1:1")
''''                    .RowHeight = 24
''''                    .VerticalAlignment = xlCenter
''''                End With
''''                Columns("K:K").ColumnWidth = 13
''''                Range("k1") = "Fecha Fin"
''''                With Range("K1")
''''                    .HorizontalAlignment = xlRight
''''                    .Font.Bold = True
''''                End With
''''                Columns("L:L").ColumnWidth = 15
''''                Range("L1").NumberFormat = "m/d/yyyy"
''''                Columns("M:M").ColumnWidth = 17
''''                Range("m1") = "Concepto"
''''                With Range("m1")
''''                    .HorizontalAlignment = xlRight
''''                    .Font.Bold = True
''''                End With
''''                With Range("j1")
''''                    .HorizontalAlignment = xlCenter
''''                    .Interior.ColorIndex = 1
''''                    .Font.ColorIndex = 2
''''                    .Font.Bold = True
''''                End With
''''            Next LibroXlsSelec
''''        End If
''''    End With
''''    Set ResultadoExplorador = Nothing
''''End Sub     ' RuT_WrkBook_Sheets_Select_Inport
