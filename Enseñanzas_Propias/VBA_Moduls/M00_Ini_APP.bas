Attribute VB_Name = "M00_Ini_APP"
' Last Rev. 2026-09-21 20:40
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M00_Ini_APP - Arranque de la aplicacion y estado del entorno Excel
' =================================================================================================
'
' PROPOSITO
'  Punto de entrada del libro y control del entorno: deja Excel en modo
'  'aplicacion' (pantalla completa, sin Ribbon ni barras) y gobierna el
'  switch de eventos Sw_EnableEvents de la hoja SwitchsAPP.
'  Lo llama ThisWorkbook.Workbook_Open.
'
' INDICE DE RUTINAS Y FUNCIONES
'  RuT_Al_Abrir_WorkBook ............. Arranque completo del libro.
'  Rut_ConfigExcel_Establecer ........ Modo aplicacion (oculta interfaz de Excel).
'  Rut_ConfigExcel_RESTABLECER ....... Devuelve Excel a su estado normal.
'  Rut_Reset_ToolsBar ................ Restaura solo barras/Ribbon (sin tocar eventos).
'  Rut_Switch_List_Status ............ Vuelca la tabla Tb_SW al informe de Form_Menu.
'  Rut_Enable_Events_Status_Choose_ByHand  Atajo manual: fuerza eventos a ON.
'  Rut_EnableEvents_Status_Reset ..... Realinea Application.EnableEvents con el switch.
'  Rut_Enable_Events_Status_Choose ... CHANGE / ON / OFF sobre eventos + switch.
'
'  OJO: Rut_Off_Functions y Rut_On_Functions YA NO estan aqui; viven en
'  Rut_Wb_State_Manager.bas (skill excel-state-manager, con reentrancia).
'
' TRAMOS DE PROGRAMACION
'  RuT_Al_Abrir_WorkBook, en orden:
'    1. Config de entorno + Off_Functions y Sw_EnableEvents = True.
'    2. Si el libro vive en NEXE y no hay Usuario_ID, pide login (Form_Usuario).
'    3. Deja visible SOLO Wk_TitP_Liquid y la protege con UserInterfaceOnly:=True
'       (protegida para el usuario, escribible desde VBA); el resto de hojas a
'       xlSheetVeryHidden.
'    4. Copia de seguridad si han pasado mas de 7 dias desde APP_CopSeg_HD_Date.
'    5. Recalcula la liquidacion activa: Rut_00_Liquid_TitProp(Liquid_Plan, Liquid_Curso_Acad).
'    6. Rotula el boton Liquid_Sw_VerRecNeg segun el switch Sw_VerRecNeg.
'    7. Inmoviliza paneles (2 columnas + fila de cabecera del ListObject) y deja
'       el cursor en la primera columna visible a partir de Liquid_Plan.
'
'  Rut_ConfigExcel_Establecer / _RESTABLECER: pares simetricos. El Ribbon se
'  oculta/muestra con ExecuteExcel4Macro show.toolbar (este libro NO tiene
'  customUI14.xml). _RESTABLECER ademas resetea los menus contextuales Cell y
'  List Range Popup, y tiene ErrorHandler que reactiva pantalla y eventos.
'
'  Grupo Enable_Events: el estado 'oficial' de los eventos es la celda
'  Sw_EnableEvents (hoja SwitchsAPP / Prog__APP_Switch), no Application.
'  _Reset copia switch -> Application; _Choose escribe en ambos a la vez.
'
' NOTAS
'  Los switches viven en Prog__APP_Switch (hoja SwitchsAPP), NO en Prog__APP.
'  Rut_Switch_List_Status lee Tb_SW de Prog__APP: es otra tabla, solo informativa.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2026-01-09
'- M00_Ini_APP

Option Explicit    ' Para obligar a definir todas las variable.  'lo he quitado porque me genera muchos errores.

' ==================================================================================================
' ==================================================================================================
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
    Wk_TitP_Liquid.Protect allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA

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
End Sub     ' RuT_Al_Abrir_WorkBook     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================


' ==================================================================================================
Sub Rut_ConfigExcel_Establecer()
' ==================================================================================================
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
        '.DisplayHorizontalScrollBar = True                             ' Show the Horizontal Scroll Bar
        '.DisplayVerticalScrollBar = True                               ' Show the Vertical Scroll Bar
    End With
    If ActiveSheet.DisplayPageBreaks Then ActiveSheet.DisplayPageBreaks = False 'Muestra/Oculta las líneas de Salto de página
'    If CommandBars("Ribbon").Controls(1).Height > 100 Then CommandBars.ExecuteMso ("MinimizeRibbon")
On Error GoTo 0
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
' ==================================================================================================
Sub Rut_ConfigExcel_RESTABLECER()        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================
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
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================


' ==================================================================================================
' Rut_Off_Functions / Rut_On_Functions -> movidas a Rut_Wb_State_Manager.bas (excel-state-manager)
'   Mismo nombre publico, misma firma (sin argumentos): las llamadas existentes no cambian.
'   Anadido soporte de reentrancia (contador de anidamiento) que esta implementacion no tenia.
' ==================================================================================================

' ==================================================================================================
Sub Rut_Reset_ToolsBar()       '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
End Sub     ' Rut_Reset_ToolsBar    >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

' ==================================================================================================
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
' --------------------------------------------------------------------------------------------------

' ==================================================================================================
'''    Sub Rut_Enable_Events_Status_Change()        ' Para permitir las rutinas que se activan cuando ocurre un evento
'''
'''        Call Rut_Enable_Events_Status_Choose("CHANGE")
'''
'''    End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' --------------------------------------------------------------------------------------------------
' ==================================================================================================
    Sub Rut_Enable_Events_Status_Choose_ByHand()
        Call Rut_Enable_Events_Status_Choose("ON")
    End Sub
' --------------------------------------------------------------------------------------------------
' ==================================================================================================
    Sub Rut_EnableEvents_Status_Reset()       '- Deja el Status_Events según el Switch ----
        Application.EnableEvents = Prog__APP_Switch.Range("Sw_EnableEvents")
    End Sub
' --------------------------------------------------------------------------------------------------
' ==================================================================================================
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
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================

