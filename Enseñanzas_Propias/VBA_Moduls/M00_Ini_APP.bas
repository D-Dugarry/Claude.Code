Attribute VB_Name = "M00_Ini_APP"
' Last Rev. 2026-09-19 21:08
'2026-01-09
'- M00_Ini_APP

Option Explicit    ' Para obligar a definir todas las variable.  'lo he quitado porque me genera muchos errores.

' ==================================================================================================================================
' ==================================================================================================================================
Sub RuT_Al_Abrir_WorkBook()
Dim WrkSht  As Worksheet

Rut_ConfigExcel_Establecer
Rut_Off_Functions
    
    
    Prog__APP_Switch.Range("Sw_EnableEvents") = True

    If Left(ActiveWorkbook.Path, 18) = "https://nexe.ua.es" And Range("Usuario_ID") = "" Then Form_Usuario.Show
'        Rut_Usuario_Chg
'    End If

    Wk_TitP_Liquid.Visible = xlSheetVisible
    Wk_TitP_Liquid.Select
    Wk_TitP_Liquid.Protect allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================

    Sht__Inf_EPs_UXXI.Visible = xlSheetHidden
    '- Ocultar todas las hojas de programación. ------
    For Each WrkSht In Worksheets
        If WrkSht.CodeName <> "Wk_TitP_Liquid" Then WrkSht.Visible = xlSheetVeryHidden   'xlSheetVisible   '
    Next
    '- Hago una copia de seguridad, si han pasado más de 7 días. ----------
    If Date - Range("APP_CopSeg_HD_Date") > 7 Then Call Rut_WrkBooK_CopSegTimed_WB_HD
    
    Debug.Print "Rut_00_Liquid_TitProp(UCase(Range('Liquid_Plan')), Range('Liquid_Curso_Acad'))", Range("Liquid_Plan")
    Call Rut_00_Liquid_TitProp(UCase(Range("Liquid_Plan")), Range("Liquid_Curso_Acad"))
    If Prog__APP_Switch.Range("Sw_VerRecNeg") Then Range("Liquid_Sw_VerRecNeg") = "Hide Rec.Neg." Else Range("Liquid_Sw_VerRecNeg") = "Ver Rec.Neg."

    ' Fijar Filas y Columnas ---------------------------
    With ActiveWindow
        .SplitColumn = 2
        .SplitRow = Wk_TitP_Liquid.ListObjects(1).Range.Rows(1).Row
        .FreezePanes = True
    End With
    Application.EnableEvents = False
    Application.ScreenUpdating = True       ' Lo repito más abajo para que funcione bien...
    Range("Liquid_Plan").Offset(0, -4).Select
        Do
            ActiveCell.Offset(0, 1).Select
        Loop Until ActiveCell.EntireColumn.Hidden = False
    Application.ScreenUpdating = True      ' Lo repito más arriba para que funcione bien...
Rut_On_Functions
    Call Rut_EnableEvents_Status_Reset
End Sub     ' RuT_Al_Abrir_WorkBook     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================


' ==================================================================================================================================
Sub Rut_ConfigExcel_Establecer()
' ==================================================================================================================================
On Error Resume Next
    With Application
        '.EnableCancelKey = False                                       ' Permito o NO utilizar Ctrl+Pausa
        .ScreenUpdating = False
        .DisplayFullScreen = True                                       'Ves pantalla completa
        .DisplayFormulaBar = False                                      'Muestra/Oculta la barra de formulas
        .DisplayStatusBar = False                                       'Muestra/Oculta la barra de estado
        .EnableEvents = Prog__APP_Switch.Range("Sw_EnableEvents")
        .ExecuteExcel4Macro "show.toolbar(""Ribbon"",false)"            'Muestra/Oculta la cinta de botones
    End With
    With ActiveWindow
        .DisplayHeadings = False                                        'Muestra/Oculta títulos de filas y columnas
        If .DisplayWorkbookTabs Then .DisplayWorkbookTabs = False       'Muestra/Oculta las pestañas de las hojas
        If .DisplayGridlines Then .DisplayGridlines = False             'Muestra/Oculta las lineas de la cuadricula
        If .DisplayPageBreaks Then .DisplayPageBreaks = False           'Muestra/Oculta las líneas de Salto de página
        '.DisplayHorizontalScrollBar = True                             ' Show the Horizontal Scroll Bar
        '.DisplayVerticalScrollBar = True                               ' Show the Vertical Scroll Bar
    End With
'    If CommandBars("Ribbon").Controls(1).Height > 100 Then CommandBars.ExecuteMso ("MinimizeRibbon")
On Error GoTo 0
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================
' ==================================================================================================================================
Sub Rut_ConfigExcel_RESTABLECER()        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================================================
    Application.CommandBars("Cell").Reset           '- Restablece el menú contextual de las celdas a su estado original
    'Esta línea restablece la barra de comandos llamada "Cell" a su configuración original.
    'La barra de comandos "Cell" es la que aparece cuando haces clic derecho en una celda en Excel. _
    ' Al restablecerla, se eliminan cualquier personalización o cambio que se haya hecho a esta barra de comandos, y se vuelve a su estado predeterminado.
    Application.CommandBars("List Range Popup").Reset   '- Restablece el menú contextual de una tabla o en un rango con formato de tabla
    'Esta línea restablece la barra de comandos llamada "List Range Popup" a su configuración original.
    'La barra de comandos "List Range Popup" es la que aparece cuando haces clic derecho en un rango de celdas que forma parte de una lista o tabla en Excel. _
    ' Al restablecerla, se eliminan cualquier personalización o cambio que se haya hecho a esta barra de comandos, y se vuelve a su estado predeterminado.
    
    
    On Error GoTo ErrorHandler
    Rut_Enable_Events_Status_Choose ("ON")
    With Application
        .ScreenUpdating = True                      ' Actualiza la pantalla cada vez
        .EnableEvents = True                        ' Permite Eventos al realizar acciones como seleccionar una celda o modificarla...
        .Calculation = xlCalculationAutomatic       ' Recalcula las fórmulas
        .DisplayAlerts = True                       ' muestra mensajes de alerta
        .DisplayFullScreen = True                   ' Ves pantalla completa
        .DisplayFormulaBar = True                   ' Muestra/Oculta la barra de formulas
        .DisplayStatusBar = True                    ' Muestra/Oculta la barra de estado
'        .CommandBars("Cell").Reset                  ' Restaura menú contextual botón derecho en celdas
        .ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"   'Muestra/Oculta la cinta de botones
    End With
    With ActiveWindow
        .DisplayHeadings = True                     ' Muestra/Oculta títulos de filas y columnas
        .DisplayWorkbookTabs = True                 ' Muestra/Oculta las pestañas de las hojas
        .DisplayGridlines = True                    ' Muestra/Oculta las lineas de la cuadricula
        '.DisplayHorizontalScrollBar = True         ' Show the Horizontal Scroll Bar
        '.DisplayVerticalScrollBar = True           ' Show the Vertical Scroll Bar
    End With
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Call Rut_EnableEvents_Status_Reset
    
    Application.CommandBars("Cell").Reset  ' Asegura reset en error
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================


' ==================================================================================================================================
Sub Rut_Off_Functions()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    If Prog__APP_Switch.Range("Sw_EnableEvents") Then Application.EnableEvents = False:     Prog__APP_Switch.Range("Sw_EnableEvents") = False               ' DesHABILITA LOS EVENTOS
End Sub
' ==================================================================================================================================
Sub Rut_On_Functions()
    If Prog__APP_Switch.Range("Sw_Calculation") Then
        Application.Calculation = xlCalculationAutomatic
    Else
         Application.Calculation = xlCalculationManual
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Prog__APP_Switch.Range("Sw_EnableEvents") Then
        Application.EnableEvents = True
    Else
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
        Prog__APP_Switch.Range("Sw_EnableEvents") = True
    End If

End Sub
' ==================================================================================================================================

' ======================================================================================================
Sub Rut_Reset_ToolsBar()       '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    With Application
        .DisplayFullScreen = True                        'Ves pantalla completa
        .DisplayFormulaBar = True                        'Oculta la barra de formulas
        .DisplayStatusBar = True                         'Oculta la barra de estado
    End With
    ActiveWindow.DisplayHeadings = True                        'Oculta títulos de filas y columnas
    ActiveWindow.DisplayWorkbookTabs = True                     'Oculta las fichas de las hohas
    ActiveWindow.DisplayGridlines = True                        'Oculta las lineas de la cuadricula
    ExecuteExcel4Macro ("show.toolbar(""ribbon"",1)")           'Oculta la cinta de botones

'    Form_Menu.TB_Informe = "ToolsBar Reset"
End Sub     ' Rut_Reset_ToolsBar    >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

' ==================================================================================================================================
Sub Rut_Switch_List_Status()
    Dim Lo_SW       As ListObject:      Set Lo_SW = Prog__APP.ListObjects("Tb_SW")
    Dim i           As Integer
    With Lo_SW.DataBodyRange
    For i = 1 To Lo_SW.ListRows.Count
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & Left(.Cells(i, 2) & String(10, " "), 10) & " - " & .Cells(i, 1) & vbLf & _
                               " - " & .Cells(i, 3) & vbLf & vbLf
    Next
    End With
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
'''    Sub Rut_Enable_Events_Status_Change()        ' Para permitir las rutinas que se activan cuando ocurre un evento
'''
'''        Call Rut_Enable_Events_Status_Choose("CHANGE")
'''
'''    End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ----------------------------------------------------------------------------------------------------------------------------------
' ==================================================================================================================================
    Sub Rut_Enable_Events_Status_Choose_ByHand()
        Call Rut_Enable_Events_Status_Choose("ON")
    End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
' ==================================================================================================================================
    Sub Rut_EnableEvents_Status_Reset()       '- Deja el Status_Events según el Switch ----
        Application.EnableEvents = Prog__APP_Switch.Range("Sw_EnableEvents")
    End Sub
' ----------------------------------------------------------------------------------------------------------------------------------
' ==================================================================================================================================
Sub Rut_Enable_Events_Status_Choose(Optional Choose As String = "CHANGE")      ' Para permitir las rutinas que se activan cuando ocurre un evento
    Select Case UCase(Choose)
        Case "CHANGE"
                            Prog__APP_Switch.Range("Sw_EnableEvents") = Not Prog__APP_Switch.Range("Sw_EnableEvents")
                            Application.EnableEvents = Prog__APP_Switch.Range("Sw_EnableEvents")
        Case "ON"
                            Application.EnableEvents = True
                            Prog__APP_Switch.Range("Sw_EnableEvents") = True
        Case "OFF"
'                        '- Verifica Application.Ready antes: Usa If Application.Ready Then Application.EnableEvents = False para evitar el error.
'                        If Application.Ready Then Application.EnableEvents = False                        ' INHABILITA LOS EVENTOS
'                        Prog__APP_Switch.Range("Sw_EnableEvents") = False
                        
                            Application.EnableEvents = False
                            Prog__APP_Switch.Range("Sw_EnableEvents") = False
    End Select
    Form_Menu.TB_Informe = "Enable Events Status now is: " & Prog__APP_Switch.Range("Sw_EnableEvents")
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

