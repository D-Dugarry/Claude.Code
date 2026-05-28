Attribute VB_Name = "Rut_WorkSheets"
Option Explicit


Sub Rut_WrkSheet_Active_Protect()
    
    ActiveSheet.Protect allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
    
'    ActiveSheet.Protect AllowFiltering:=True, _        Permite Filtrar
'                        AllowSorting:=True, _          Permite Ordenar
'                        DrawingObjects:=True, _        Impide que se seleccionen las Shapes
'                        UserInterfaceOnly:=True        Permite que las Macros puedan modificar la Sheet

End Sub

' ==================================================================================================================================
Function Fnc_WrkSheet_Exist(SheetName As String) As Boolean
    Dim Ws As Worksheet
    On Error Resume Next
    Set Ws = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    Fnc_WrkSheet_Exist = Not Ws Is Nothing
End Function
    
    Sub Rut_WrkSheet_All_Crear_Lista()    ' Manejo interno, Hace la lista de todas las hojas de este libro       =============================
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
        .Range.Sort Key1:=.ListColumns(1), Order1:=xlAscending, _
            Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    End With    ' .ListObjects("Tb_Sheets_State")
    Call Rut_LstObj_DataBodyRange_Copy(Prog_HojasName.ListObjects("Tb_Sheets_State"), Prog_HojasName.ListObjects("Tb_Sheets_Config"), True)
    .Protect
    End With ' Prog_HojasName
    Application.ScreenUpdating = True
    End Sub
'###################################################################################################################################
Sub Rut_WrkSheet_Columns_ALL_ShowHide()
 
    .Columns.EntireColumn.Hidden = False
    End Sub
'###################################################################################################################################
Sub Rut_WrkSheet_Rows_ALL_ShowHide()      ' Muestra Todas las Solicitudes  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    Rows.EntireRow.Hidden = False
    End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'###################################################################################################################################
Sub Rut_WrkSheet_Columns_Show_All2(Optional ByVal WrkSht As String)
    Dim WrkSht_Activa    As String:     WrkSht_Activa = ActiveSheet.Name
    Application.ScreenUpdating = False
    If WrkSht = "" Then WrkSht = WrkSht_Activa
    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
        Dim Visual_Status   As Variant:  Visual_Status = .Visible   '--- para dejar la hoja en el mismo estado de Visibilidad ---
        Dim Protect_Status   As Boolean:  Protect_Status = .ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
        .Visible = xlHidden
        .Unprotect
        .Columns.EntireColumn.Hidden = False
        .Visible = Visual_Status
        If Protect_Status Then .Protect
        Sheets(WrkSht_Activa).Select
    End With
        Application.ScreenUpdating = True
    Lo_Menu.DataBodyRange.Cells(TaskIndice, 5) = "All columns are visible in sheet: '" & Sheets(WrkSht).CodeName & "'  -  " & Now
    End Sub
            
            '###################################################################################################################################
                    Sub Rut_WrkSheet_LstObj_LiberarEspacio_ByHand()
                        '''Application.Workbooks(ThisWorkbook.Name).Sheets(ActiveSheet.Name).ListObjects(1).DataBodyRange.Delete
                        Call Rut_WrkSheet_LstObj_LiberarEspacio(ActiveSheet.Name)
                    End Sub
'###################################################################################################################################
Sub Rut_WrkSheet_LstObj_LiberarEspacio(ByVal WrkSht As String)  '--- Borra TODO a la Derecha y Abajo de .ListObjects(1) ----------------------
        
        Application.Calculation = xlManual
        Application.ScreenUpdating = False
    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
            .Protect , allowFiltering:=True, UserInterfaceOnly:=True    '- Mantiene protegida la hoja pero permite modificar con VBA
            .Columns.EntireColumn.Hidden = False                        ' Mostrar todas las Columnas
            .Rows.EntireRow.Hidden = False                              ' Mostrar todas las Filas
            If .FilterMode Then .ShowAllData                            ' Deshacer Filtros
        With .ListObjects(1).Range                                      ' Borra filas de abajo y columnas a la derecha del la Tabla .ListObjects(1)
            Cells(.Cells(.Rows.Count, 1).Row, .Cells(1, .Columns.Count).Column).Select  '- Última celda de la Tabla, independientemente de donde comience.
            ActiveCell.Offset(1, 1).Select
        End With
        .Range(ActiveCell.Address & ":" & Cells(Rows.Count, 1).Address).EntireRow.Delete
        .Range(ActiveCell.Address & ":" & Cells(1, Columns.Count).Address).EntireColumn.Delete
        ActiveSheet.UsedRange                       ' Para restablecer el rango de celdas en uso
    End With
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
' ==================================================================================================================================
Sub Rut_WrkSheet_LstObj_LiberarEspacio2()      ' Borra la filas de abajo y columnas de la derecha del la Tabla  ----------------------------------
With ActiveSheet
        .Unprotect
        .ListObjects(1).Range.Cells(.ListObjects(1).Range.Rows.Count, .ListObjects(1).Range.Columns.Count).Select
        ActiveCell.Offset(1, 1).Select
        .Range(ActiveCell.Address, .Cells(Rows.Count, 1)).EntireRow.Delete
        .Range(ActiveCell.Address, .Cells(1, Columns.Count)).EntireColumn.Delete
        .UsedRange
End With
End Sub

            '###################################################################################################################################
                    Sub Rut_WrkSheet_Vaciar_byhand()
                        Call Rut_WrkSheet_Vaciar(ActiveSheet.Name)
                    End Sub
' ==================================================================================================================================
Sub Rut_WrkSheet_Vaciar(ByVal WrkSht As String)      '--- Borra Toda la Hoja incluso los objetos (Shapes)  -------------------------------
' ==================================================================================================================================
Dim WrkSht_Activa    As String:     WrkSht_Activa = ActiveSheet.Name

        Application.Calculation = xlManual
        
    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
    
            Dim Visual_Status           As Variant:  Visual_Status = .Visible   '--- para dejar la hoja en el mismo estado de Visibilidad ---
            Dim Protect_Status          As Boolean:  Protect_Status = .ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
        .Visible = xlHidden
        .Unprotect
        
        .Columns.EntireColumn.Hidden = False
        .Rows.EntireRow.Hidden = False
        If .FilterMode Then .ShowAllData
        
        .Columns.Delete     ' --- con esto se borran hasta los "Shapes"
        .UsedRange
        .Visible = Visual_Status
        If Protect_Status Then .Protect
        Sheets(WrkSht_Activa).Select
        
    End With
        Application.Calculation = xlAutomatic
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
    Sub Rut_WrkSheet_All_Visible_OrNot()    ' Estable la visibilidad establecida de cada Sheet en la Tabla(2) de Prog_Hojas_Names
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
' ==================================================================================================================================
Sub Rut_WrkSheet_All_hide()
' ==================================================================================================================================
    Dim WrkSht          As Worksheet
    For Each WrkSht In Worksheets
        If Left(WrkSht.CodeName, 5) = "Prog_" Then WrkSht.Visible = xlSheetVeryHidden   'xlSheetVisible   '
    Next
    Lo_Menu.DataBodyRange.Cells(TaskIndice, 5) = "All Prog_sheets Hide  -  " & Now
End Sub
' ==================================================================================================================================
Sub Rut_WrkSheet_All_Visible()    ' Estable la visibilidad establecida de cada Sheet en la Tabla(2) de Prog_Hojas_Names
' ==================================================================================================================================
    Dim WrkSht      As Worksheet
    For Each WrkSht In Worksheets
            WrkSht.Visible = True
    Next
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
' ==================================================================================================================================
    Sub Rut_WrkSheet_Del()    '   =============================
' ==================================================================================================================================
Dim WrkSht      As Worksheet
    For Each WrkSht In Worksheets
        If Left(WrkSht.Name, 6) = "Pagos_" Then
            Application.DisplayAlerts = False
            WrkSht.Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
' ==================================================================================================================================
    Sub Rut_WrkSheet_Add()    '   Crea una nueva Sheet =============================
' ==================================================================================================================================
    Dim NewSheet As Worksheet
    Set NewSheet = Worksheets.Add
        NewSheet.Cells(i, 1).Value2 = "Bla Bla Bla..."
        NewSheet.Cells(i, 2).Value2 = "RefersTo"
        NewSheet.Columns("A:B").AutoFit
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

            ' ======================================================================================================================
            Sub Rut_WrkSheet_Sort_ByHand()
                RuT_Sort_Sheets
            End Sub
' ==================================================================================================================================
Sub Rut_WrkSheet_Sort(Optional Sort_Ascending As Boolean = True)
' ==================================================================================================================================
Dim Pos_Outer  As Integer
Dim Pos_Inner  As Integer
Dim Cant_Sheets     As Integer:   Cant_Sheets = Sheets.Count

    For Pos_Outer = 1 To Cant_Sheets
        For Pos_Inner = 1 To Pos_Outer
            If Sort_Ascending Then
                If UCase(Sheets(Pos_Outer).Name) < UCase(Sheets(Pos_Inner).Name) Then
                        Sheets(Pos_Outer).Move before:=Sheets(Pos_Inner)
                End If
            Else
                If UCase(Sheets(Pos_Outer).Name) > UCase(Sheets(Pos_Inner).Name) Then
                        Sheets(Pos_Outer).Move before:=Sheets(Pos_Inner)
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
Sub Rut_WrkSheet_Exportar_La_Liquidación()   '- Copia una Sheet concreta
' ==================================================================================================================================
Rut_Off_Functions
Dim FPath           As String:          FPath = ThisWorkbook.Path & "\"
    Hora_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Copio la Sheet entera y esto es lo que voy a grabar. -----------------------------------------
    Work_TitPropios_Liquid.Copy
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
    Lo_Menu.DataBodyRange.Cells(TaskIndice, 5) = "¡¡¡ Proceso concluido con éxito !!!   He tardado: " & Format(Timer - Hora_Inicio, "0.00") & _
            " seg.  -.-  " & Now() & vbCrLf & vbCrLf & "Exportado el Resumen de la Liquidación de:   " & Work_TitPropios_Liquid.Range("Liquid_Plan_Name") & _
            vbCrLf & vbCrLf & "En el Archivo:   " & sFileSaveName
    GoTo Salir_Sub
GestError:
    Debug.Print "Error Rut_Exportar_La_Liquidación ", Err.Number, Err.Description, Err.Source
    Debug.Print sFileSaveName
    MsgBox "Rut_Exportar_La_Liquidación " & "Filename:=" & vbCrLf & sFileSaveName, vbExclamation + vbOKOnly, "Rutina de Remesado"
    Lo_Menu.DataBodyRange.Cells(TaskIndice, 5) = "Rut_Exportar_La_Liquidación " & vbCrLf & "Filename:=" & sFileSaveName
Salir_Sub:
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
    Hora_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Copio la Sheet entera -----------------------------------------
    Work_TitPropios_Liquid.Copy
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "Liquid_" & Work_TitPropios_Liquid.Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".pdf"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.pdf), *.pdf")
        If sFileSaveName <> False Then
            On Error GoTo Restablecer_Valores
            Application.DisplayAlerts = False
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sFileSaveName
'            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True, FileFormat:=51
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    
Restablecer_Valores:
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------



' ==================================================================================================================================
' =============================     RuT_Guardar_Liquidación     =======================================================================
' ==================================================================================================================================
Sub RuT_WrkSheet_Save_in_New_WorkBook()
    Dim FPath           As String

Application.ScreenUpdating = False
'    FPath = ActiveWorkbook.Path & "\"
    FPath = ActiveWorkbook.Path
    FPath = Replace(App_RutaAPP, "/", "\")
    FPath = Mid(App_RutaAPP, InStr(1, App_RutaAPP, Prog__APP.Range("APP_User_Mail")) + Len(Prog__APP.Range("APP_User_Mail")))
    FPath = Prog__APP.Range("App_User_Unid_Red") & ":" & App_RutaAPP & "\"
    
'   Copy and Save a Sheet in a New WorkBook -----------------------------
    H_Liquid_TPV.Copy
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & "Liquid_" & Range("e2") & "_JyC_" & Range("e1")    ' & ".xlsx"
    sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
        If sFileSaveName <> False Then
            On Error GoTo Restablecer_Valores
'            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
'            Application.DisplayAlerts = True
        End If
Range("a1").Select
If ActiveSheet.Shapes.Count > 0 Then ActiveSheet.Shapes.SelectAll:  Selection.Delete
ActiveWorkbook.Close SaveChanges:=True
MsgBox "Hecho"
Application.ScreenUpdating = True
Exit Sub
Restablecer_Valores:
Application.ScreenUpdating = True
ActiveWorkbook.Close SaveChanges:=False
End Sub     ' RuT_Sheet_Save_Liquidación_New_WorkBook

'    ' ============================  Para mostrar u ocultar Columnas según Lista [[lista]]   ============================================
'    ' ==================================================================================================================================
    Sub Rut_WrkSheet_Cols_Show_Hide_IF()       ' CuadroTextoVistaColmns    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Dim Pos_Ini     As Integer
'    Dim Pos_Fin     As Integer
'    Dim Cont            As Integer
'    Dim Pos_Caract      As Integer
'    Dim Letra_Col   As String
'    Dim Cadena      As String
'    Dim Caract      As String
'
'        Cadena = Form_Menú.TBx_Descripción
'        Pos_Ini = InStr(Cadena, "[[")
'        Pos_Fin = InStr(Cadena, "]]")
'        If Pos_Ini * Pos_Fin = 0 Then   '---Controla que existe marca de inicio y fin y que hay algun dato entre marcas
'            APP_MenúAux_Msg = "Error: No hay lista de Columnas a visualizar, o no empieza por [[, o no acaba por ]]." & vbCrLf & Now
'            Exit Sub
'        End If
'
'        Cadena = Mid(Cadena, Pos_Ini + 2, Pos_Fin - Pos_Ini - 2)    '---extraigo la Cadena
'        Cadena = Func_Normalizar_Lista_Columnas(Cadena)                           '---si hay espacios, los quito
'        If Left(Cadena, 5) = "Error" Then
'            APP_MenúAux_Msg = Cadena & vbCrLf & Now
'            Exit Sub
'        End If
'
'        APP_MenúAux_Msg = "Lista de Columnas a visualizar: [[" & Cadena & "]]" & vbCrLf & Now
'        Form_Menú.TBx_Descripción = Left(Form_Menú.TBx_Descripción, Pos_Ini + 1) & Cadena & Mid(Form_Menú.TBx_Descripción, Pos_Fin)
'
'        Range(Columns(2), Columns(LastCol_Tb_Solicitudes)).Hidden = True    '---Oculta de la Columna 2 a la última
'
'        Pos_Ini = 1
'        Do
'            If InStr(",:", Mid(Cadena, Pos_Ini + 1, 1)) > 0 Then        ' Controlar si columna de 1 o 2 letras
'                Letra_Col = Mid(Cadena, Pos_Ini, 1)
'                Pos_Ini = Pos_Ini + 1
'            Else
'                Letra_Col = Mid(Cadena, Pos_Ini, 2)
'                Pos_Ini = Pos_Ini + 2
'            End If
'            If Mid(Cadena, Pos_Ini, 1) = ":" Then                       ' Controlo si ":"
'                Letra_Col = Letra_Col & ":"
'                Pos_Ini = Pos_Ini + 1
'
'                If InStr(",:", Mid(Cadena, Pos_Ini + 1, 1)) > 0 Then    ' Controlar si columna de 1 o 2 letras
'                    Letra_Col = Letra_Col & Mid(Cadena, Pos_Ini, 1)
'                    Pos_Ini = Pos_Ini + 1
'                Else
'                    Letra_Col = Letra_Col & Mid(Cadena, Pos_Ini, 2)
'                    Pos_Ini = Pos_Ini + 2
'                End If
'            Else
'            End If
'            Pos_Ini = Pos_Ini + 1
'            Columns(Letra_Col).Hidden = False
'        Loop While Pos_Ini < Len(Cadena)
'
'        LO_Tb_LS_CAS.DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1).Select        '  Me sitúo en la segunda celda del rango DataBodyRange de las celdas visibles que será la segunda celda de las primera fila visible.
'
    End Sub     ' Rut_WrkSheet_Cols_Show_Hide_IF     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    '==================================================================================================================================
'

