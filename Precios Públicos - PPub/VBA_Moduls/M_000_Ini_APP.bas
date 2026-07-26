Attribute VB_Name = "M_000_Ini_APP"
'- M00_Ini_APP --------
Option Explicit    ' Para obligar a definir todas las variable.  'lo he quitado porque me genera muchos errores.

' ==================================================================================================================================
Sub RuT_Al_Abrir_WorkBook()
Debug.Print "------------------------------------------->>> RuT_Al_Abrir_WorkBook"
    Rut_Off_Functions
    Rut_ConfigExcel_Establecer
    
    'Application.EnableEvents = True
    Sht__BD.Visible = xlSheetVisible
    Sht__BD.Select
    Sht__BD.Unprotect
    Sht__BD.ListObjects(1).ShowTotals = True
        Sht__BD.Protect , allowFiltering:=True, DrawingObjects:=False, Scenarios:=True, Contents:=True, UserInterfaceOnly:=True        '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
    ActiveWindow.DisplayWorkbookTabs = False
    Rut_On_Functions
Debug.Print "-------------------------------------------<<< RuT_Al_Abrir_WorkBook"
End Sub     ' RuT_Al_Abrir_WorkBook
' ==================================================================================================================================
' ==================================================================================================================================
Sub Rut_ConfigExcel_Establecer()
' ----------------------------------------------------------------------------------------------------------------------------------
Debug.Print "Rut_ConfigExcel_Establecer"
On Error Resume Next
    Call Rut_Context_Buttons_Hide   '- Oculta las opciones genéricas del Context-Menú / Right-ClicK para dejar sólo visible las opciones Custom
    ActiveWindow.DisplayHorizontalScrollBar = True                                      ' Show the Horizontal Scroll Bar
    ActiveWindow.DisplayVerticalScrollBar = True                                        ' Show the Vertical Scroll Bar
'    Application.EnableCancelKey = False                                                 ' Permito o NO utilizar Ctrl+Pausa
    Application.EnableEvents = True:        Prog__APP.Range("SW_Events") = True
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
' ==================================================================================================================================
Sub Rut_ConfigExcel_RESTABLECER()
Debug.Print "Rut_ConfigExcel_RESTABLECER"
    Call Rut_Context_Buttons_Restore    '- Restaura las opciones genéricas del Context-Menú / Right-ClicK para dejar sólo visible las opciones Custom
    Application.EnableEvents = True:        Prog__APP.Range("SW_Events") = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True                             'Muestra/Oculta la barra de estado
    ActiveWindow.DisplayHeadings = True                             'Muestra/Oculta títulos de filas y columnas
    ActiveWindow.DisplayWorkbookTabs = True                         'Muestra/Oculta las pestañas de las hohjs
    ActiveWindow.DisplayGridlines = True                             'Muestra/Oculta las lineas de la cuadricula
    Application.DisplayFormulaBar = True                             'Muestra/Oculta la barra de formulas
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    If Application.Toolbars("Ribbon").Visible = False Then Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"
End Sub
' ==================================================================================================================================
Sub Rut_Off_Functions()
Debug.Print "Rut_Off_Functions"
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False:     Prog__APP.Range("SW_Events") = False               ' DesHABILITA LOS EVENTOS (SIEMPRE, sin depender del switch)
End Sub
' ==================================================================================================================================
Sub Rut_On_Functions()
    If Prog__APP.Range("SW_App_Calculation") Then
        Application.Calculation = xlCalculationAutomatic
    Else
         Application.Calculation = xlCalculationManual
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If Prog__APP.Range("SW_Events") Then
        Application.EnableEvents = True
    Else
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
        Prog__APP.Range("SW_Events") = True
    End If
Debug.Print "Rut_On_Functions"
End Sub
' ==================================================================================================================================
Sub Rut_Menú_HideAll()
Debug.Print "Rut_Menú_HideAll"
'''    ActiveSheet.Shapes("Traffic-Light").Visible = True
'''            SW_TrafficLight1 = True
'''            SW_TrafficLight2 = False
'''            SW_TrafficLight3 = False
'''        MyRibbon.InvalidateControl "RibbViewNone"    '- Actualiza solo este botón
'''        MyRibbon.InvalidateControl "RibbViewMin"    '- Actualiza solo este botón
'''        MyRibbon.InvalidateControl "RibbViewMax"    '- Actualiza solo este botón
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayStatusBar = False
    Application.DisplayFormulaBar = False
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
Debug.Print "RibbonX Ocultado  -  " & Now
End Sub
' ==================================================================================================================================
Sub Rut_Menú_ShowAll()
Debug.Print "Rut_Menú_ShowAll"
    ' Exit Full Screen
    ' - If this was used to show full screen after other display settings
    '   were changed, then put it before those settings are changed back.
    '   Full-screen mode and normal mode maintain separate settings for these.
    '
    '   Si se utilizó para mostrar la pantalla completa después de cambiar otros ajustes
    '   de visualización, póngalo antes de volver a cambiar esos ajustes.
    '   El modo de pantalla completa y el modo normal mantienen configuraciones separadas para estos.
    
'''    ActiveSheet.Shapes("Traffic-Light").Visible = False
'''            SW_TrafficLight1 = False
'''            SW_TrafficLight2 = False
'''            SW_TrafficLight3 = True
'''        MyRibbon.InvalidateControl "RibbViewNone"    '- Actualiza solo este botón
'''        MyRibbon.InvalidateControl "RibbViewMin"    '- Actualiza solo este botón
'''        MyRibbon.InvalidateControl "RibbViewMax"    '- Actualiza solo este botón
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
'    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayStatusBar = True

    ' Show the Ribbon Menu and Quick Access Toolbar
    If Application.Toolbars("Ribbon").Visible = False Then Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"
    If CommandBars("Ribbon").Height < 150 Then
        CommandBars.ExecuteMso ("MinimizeRibbon")
        DoEvents 'Important!
    End If
Debug.Print "RibbonX Visible y Expandido  -  " & Now
End Sub
' ==================================================================================================================================
Sub Rut_Menú_ShowAll_Short()
Debug.Print "Rut_Menú_ShowAll_Short"
    ' Exit Full Screen
    ' - If this was used to show full screen after other display settings
    '   were changed, then put it before those settings are changed back.
    '   Full-screen mode and normal mode maintain separate settings for these.
    '
    '   Si se utilizó para mostrar la pantalla completa después de cambiar otros ajustes
    '   de visualización, póngalo antes de volver a cambiar esos ajustes.
    '   El modo de pantalla completa y el modo normal mantienen configuraciones separadas para estos.
    
'''    ActiveSheet.Shapes("Traffic-Light").Visible = False
'''            SW_TrafficLight1 = False
'''            SW_TrafficLight2 = True
'''            SW_TrafficLight3 = False
'''        MyRibbon.InvalidateControl "RibbViewNone"    '- Actualiza solo este botón
'''        MyRibbon.InvalidateControl "RibbViewMin"    '- Actualiza solo este botón
'''        MyRibbon.InvalidateControl "RibbViewMax"    '- Actualiza solo este botón
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
'    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayStatusBar = True

    ' Show the Ribbon Menu and Quick Access Toolbar
    If Application.Toolbars("Ribbon").Visible = False Then Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"
    If CommandBars("Ribbon").Height > 150 Then
        CommandBars.ExecuteMso ("MinimizeRibbon")
        DoEvents 'Important!
    End If
Debug.Print "RibbonX Visible sin Expandir  -  " & Now
End Sub
' ==================================================================================================================================
Sub Rut_Context_Buttons_Hide()  '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
Debug.Print "Rut_Context_Buttons_Hide"
Dim CmdBarNames        As Variant:         CmdBarNames = Array("Cell", "List Range Popup")
Dim ib                 As Integer
Dim Count              As Integer
Dim CommBarItem        As CommandBarControl
    ' And for each command bar, iterate through all the available controls
    For ib = LBound(CmdBarNames) To UBound(CmdBarNames)
        For Count = 1 To Application.CommandBars(CmdBarNames(ib)).Controls.Count
            Set CommBarItem = Application.CommandBars(CmdBarNames(ib)).Controls(Count)
            If CommBarItem.Tag <> "My_Ctxt_Menu" Then
                CommBarItem.Visible = False
            Else
                Debug.Print CmdBarNames(ib) & "-" & Count & ". - " & CommBarItem.ID & " - " & CommBarItem.Caption & " - " & CommBarItem.Enabled & " - " & CommBarItem.Visible

            End If
        Next Count
    Next ib
    Prog__APP.Range("SW_RightClickMenú_Restricted") = True
End Sub
' ==================================================================================================================================
Sub Rut_Context_Buttons_Restore()  '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
Debug.Print "Rut_Context_Buttons_Restore"
    Application.CommandBars("Ply").Enabled = True   '- Permite visualizar o NO, el Context-Menú / Right-ClicK
    Application.CommandBars("Cell").Reset           '- Restablece el menú contextual de las celdas a su estado original
    'Esta línea restablece la barra de comandos llamada "Cell" a su configuración original.
    'La barra de comandos "Cell" es la que aparece cuando haces clic derecho en una celda en Excel. _
     Al restablecerla, se eliminan cualquier personalización o cambio que se haya hecho a esta barra de comandos, y se vuelve a su estado predeterminado.
    Application.CommandBars("List Range Popup").Reset   '- Restablece el menú contextual de una tabla o en un rango con formato de tabla
    'Esta línea restablece la barra de comandos llamada "List Range Popup" a su configuración original.
    'La barra de comandos "List Range Popup" es la que aparece cuando haces clic derecho en un rango de celdas que forma parte de una lista o tabla en Excel. _
     Al restablecerla, se eliminan cualquier personalización o cambio que se haya hecho a esta barra de comandos, y se vuelve a su estado predeterminado.
    
    Prog__APP.Range("SW_RightClickMenú_Restricted") = False
End Sub
' ==================================================================================================================================
'===================================================================================================================================
Sub Rut_OnOff_SW_WB_Deactivate()
    If Prog__APP.Range("SW_WB_Deactivate") Then
        Prog__APP.Range("SW_WB_Deactivate") = False
        Form_Menu.Lb_SW_WB_Deactivate.Visible = False
        Prog__APP.Range("APP_Task_Inf") = "SW_WB_Deactivate - DesActivado  -  " & Now
    Else
        Prog__APP.Range("SW_WB_Deactivate") = True
        Form_Menu.Lb_SW_WB_Deactivate.Visible = True
        Prog__APP.Range("APP_Task_Inf") = "SW_WB_Deactivate - Activado  -  " & Now
    End If
End Sub
'===================================================================================================================================
Sub Rut_Right_Click_Control_KK()
    If Prog__APP.Range("SW_WB_Deactivate") Then
        Prog__APP.Range("SW_WB_Deactivate") = False
        Form_Menu.Lb_SW_WB_Deactivate.Visible = False
        Prog__APP.Range("APP_Task_Inf") = "SW_WB_Deactivate - DesActivado  -  " & Now
    Else
        Prog__APP.Range("SW_WB_Deactivate") = True
        Form_Menu.Lb_SW_WB_Deactivate.Visible = True
        Prog__APP.Range("APP_Task_Inf") = "SW_WB_Deactivate - Activado  -  " & Now
    End If
End Sub


