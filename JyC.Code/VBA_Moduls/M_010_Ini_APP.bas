Attribute VB_Name = "M_010_Ini_APP"
Option Explicit

' ==================================================================================================================================
Sub RuT_Al_Abrir_WorkBook()
    Rut_ConfigExcel_Establecer
    Rut_Off_Functions
    Prog__APP.Range("SW_Events") = True
    Application.EnableEvents = False
    Rut_Menú_HideAll
    Prog_CTA_Tb.Visible = xlSheetVisible
    Prog_TPV_Tb.Visible = xlSheetVisible
    Prog_CTA_Tb.Protect allowFiltering:=True, DrawingObjects:=True, allowSorting:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
    Prog_TPV_Tb.Protect allowFiltering:=True, DrawingObjects:=True, allowSorting:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
    H_INICI.Visible = xlSheetVisible
    Workbooks(ThisWorkbook.Name).Sheets(H_INICI.Name).Select
    H_INICI.Protect allowFiltering:=True, DrawingObjects:=True, allowSorting:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
    Application.ScreenUpdating = True      ' Lo repito más arriba para que funcione bien...
    Call RefreshRibbon
    Rut_On_Functions
End Sub     ' RuT_Al_Abrir_WorkBook     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' -------------------------------------------------------------------------------------------------------------------------------<<<
' ==================================================================================================================================
Sub Rut_ConfigExcel_Establecer()
' ==================================================================================================================================
On Error Resume Next
    ActiveWindow.DisplayHorizontalScrollBar = True                                      ' Show the Horizontal Scroll Bar
    ActiveWindow.DisplayVerticalScrollBar = True                                        ' Show the Vertical Scroll Bar
'    Application.EnableCancelKey = False                                                 ' Permito o NO utilizar Ctrl+Pausa
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    Application.DisplayFullScreen = False                                                'Ves pantalla completa
    Application.DisplayFormulaBar = False                                               'Muestra/Oculta la barra de formulas
    Application.DisplayStatusBar = False                                                'Muestra/Oculta la barra de estado
    ActiveWindow.DisplayHeadings = False                                                'Muestra/Oculta títulos de filas y columnas
    If ActiveWindow.DisplayWorkbookTabs Then ActiveWindow.DisplayWorkbookTabs = False   'Muestra/Oculta las pestañas de las hojas
    If ActiveWindow.DisplayGridlines Then ActiveWindow.DisplayGridlines = False         'Muestra/Oculta las lineas de la cuadricula
    If ActiveSheet.DisplayPageBreaks Then ActiveSheet.DisplayPageBreaks = False         'Muestra/Oculta las líneas de Salto de página

On Error GoTo 0
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
'###################################################################################################################################
Sub Rut_ConfigExcel_Restablecer()
' ==================================================================================================================================
    Application.CommandBars("Ply").Enabled = True   '- Permite visualizar el Context-Menú Right-ClicK
    Application.CommandBars("Cell").Reset
    Application.CommandBars("List Range Popup").Reset
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True                              'Muestra/Oculta la barra de estado
    ActiveWindow.DisplayHeadings = True                              'Muestra/Oculta títulos de filas y columnas
    ActiveWindow.DisplayWorkbookTabs = True                          'Muestra/Oculta las pestañas de las hohjs
    ActiveWindow.DisplayGridlines = True                             'Muestra/Oculta las lineas de la cuadricula
    Application.DisplayFormulaBar = True                             'Muestra/Oculta la barra de formulas
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    If Application.Toolbars("Ribbon").Visible = False Then Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
'###################################################################################################################################
Sub Rut_Off_Functions()
' ==================================================================================================================================
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    If Prog__APP.Range("SW_Events") Then Application.EnableEvents = False:     Prog__APP.Range("SW_Events") = False               ' DesHABILITA LOS EVENTOS
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
'###################################################################################################################################
Sub Rut_On_Functions()
' ==================================================================================================================================
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Prog__APP.Range("SW_Events") Then
        Application.EnableEvents = True
    Else
        Application.EnableEvents = True             ' HABILITA LOS EVENTOS
        Prog__APP.Range("SW_Events") = True
    End If
End Sub
' ==================================================================================================================================
Sub Rut_Menú_HideAll()
    H_INICI.Shapes("Traffic-Light").Visible = True
            SW_TrafficLight1 = True
            SW_TrafficLight2 = False
            SW_TrafficLight3 = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayStatusBar = False
    Application.DisplayFormulaBar = False
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
End Sub
' ==================================================================================================================================
Sub Rut_Menú_ShowAll()
    ' Exit Full Screen
    ' - If this was used to show full screen after other display settings
    '   were changed, then put it before those settings are changed back.
    '   Full-screen mode and normal mode maintain separate settings for these.
    '
    '   Si se utilizó para mostrar la pantalla completa después de cambiar otros ajustes
    '   de visualización, póngalo antes de volver a cambiar esos ajustes.
    '   El modo de pantalla completa y el modo normal mantienen configuraciones separadas para estos.
    
    H_INICI.Shapes("Traffic-Light").Visible = False
            SW_TrafficLight1 = False
            SW_TrafficLight2 = False
            SW_TrafficLight3 = True
'    Application.DisplayFullScreen = False
'    Application.DisplayFormulaBar = True
'    ActiveWindow.DisplayHeadings = True
'    Application.Toolbars("Ribbon").Visible = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayStatusBar = True

    ' Show the Ribbon Menu and Quick Access Toolbar
    If Application.Toolbars("Ribbon").Visible = False Then Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"
    If CommandBars("Ribbon").Height < 150 Then
        CommandBars.ExecuteMso ("MinimizeRibbon")
        DoEvents 'Important!
    End If
End Sub
' ==================================================================================================================================
Sub Rut_Menú_ShowAll_Short()
    ' Exit Full Screen
    ' - If this was used to show full screen after other display settings
    '   were changed, then put it before those settings are changed back.
    '   Full-screen mode and normal mode maintain separate settings for these.
    '
    '   Si se utilizó para mostrar la pantalla completa después de cambiar otros ajustes
    '   de visualización, póngalo antes de volver a cambiar esos ajustes.
    '   El modo de pantalla completa y el modo normal mantienen configuraciones separadas para estos.
    H_INICI.Shapes("Traffic-Light").Visible = False
            SW_TrafficLight1 = False
            SW_TrafficLight2 = True
            SW_TrafficLight3 = False
'    Application.DisplayFullScreen = False
'    ActiveWindow.DisplayHeadings = True
'    Application.DisplayFormulaBar = True
'    Application.Toolbars("Ribbon").Visible = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayStatusBar = True

    ' Show the Ribbon Menu and Quick Access Toolbar
    If Application.Toolbars("Ribbon").Visible = False Then Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"
    If CommandBars("Ribbon").Height > 150 Then
        CommandBars.ExecuteMso ("MinimizeRibbon")
        DoEvents 'Important!
    End If
End Sub
' ==================================================================================================================================
Sub Rut_Context_Buttons_Hide()  '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
Dim Item_Bar           As Integer
Dim Count              As Integer
Dim CommBarItem        As CommandBarControl
    ' And for each command bar, iterate through all the available controls
    Item_Bar = 38 '- (38)=("Cell") ----------------------------------------------------------
    For Count = 1 To Application.CommandBars(Item_Bar).Controls.Count
        Set CommBarItem = Application.CommandBars(Item_Bar).Controls(Count)
        If CommBarItem.Tag <> "My_Ctxt_Menu" Then
'            CommBarItem.Visible = False
        Else
            Debug.Print Item_Bar & "-" & Count & ". - " & CommBarItem.id & " - " & CommBarItem.Caption & " - " & CommBarItem.Enabled & " - " & CommBarItem.Visible

        End If
    Next Count
    Item_Bar = 74 '- (74)=("List Range Popup") ---------------------------------------------
    For Count = 1 To Application.CommandBars(Item_Bar).Controls.Count
        Set CommBarItem = Application.CommandBars(Item_Bar).Controls(Count)
        If CommBarItem.Tag <> "My_Ctxt_Menu" Then
            CommBarItem.Visible = False
        Else
            Debug.Print Item_Bar & "-" & Count & ". - " & CommBarItem.id & " - " & CommBarItem.Caption & " - " & CommBarItem.Enabled & " - " & CommBarItem.Visible

        End If
    Next Count
End Sub
' ==================================================================================================================================
Sub Rut_Context_Buttons_Restore()  '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
    Application.CommandBars("Ply").Enabled = True   '- Permite visualizar el Context-Menú Right-ClicK
    Application.CommandBars("Cell").Reset
    Application.CommandBars("List Range Popup").Reset
End Sub
' ==================================================================================================================================

