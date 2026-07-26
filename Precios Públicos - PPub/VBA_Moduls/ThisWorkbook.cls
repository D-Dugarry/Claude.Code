VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private myRibUI             As IRibbonUI

Public Property Let ribbonUI(irib As IRibbonUI)
    Set myRibUI = irib
End Property
Public Property Get ribbonUI() As IRibbonUI
    Set ribbonUI = myRibUI
End Property

'Private Sub Workbook_Activate()
'Debug.Print "Sub Workbook_Activate"
''    Application.ErrorCheckingOptions.NumberAsText = False ' Desactiva el indicador de error para números almacenados como texto
'    Debug.Print Prog__APP.Range("SW_Events")
'    Call RefreshRibbon
'        Call Rut_ConfigExcel_Establecer
'End Sub

Private Sub Workbook_Deactivate()
Debug.Print "Workbook_Deactivate"
''    Application.ErrorCheckingOptions.NumberAsText = True
    If Prog__APP.Range("SW_WB_Deactivate") Then
        Debug.Print "Sub Workbook_Deactivate"
        Call Rut_ConfigExcel_RESTABLECER
    End If
End Sub

' ==================================================================================================================================
Sub Workbook_Open()
Debug.Print "Workbook_Open --------------------------->>> ThisWorkbook "
'    Call Rut_WrkBook_MinimizeAllWindowsExceptThisExcel
    Call Rut_WrkBook_MinimizeAllExcelExceptThisWB
    Application.Visible = False
    Application.ScreenUpdating = False
    Application.DisplayFullScreen = True
    Application.EnableEvents = True:        Prog__APP.Range("SW_Events") = True
    Prog__APP.Range("SW_Test") = False
    Prog__APP.Range("SW_RightClickMenú_Visible") = True
    Prog__APP.Range("SW_WB_Deactivate") = True     '- Cambiarlo a True para los usuarios
    
'    Application.ErrorCheckingOptions.NumberAsText = False
    Form_Usuario.Show
    Call RuT_Load_Task_Data("Rut_Chg_Usuario")
    Application.CommandBars("Ply").Enabled = True   '- Permite visualizar el Context-Menú Right-ClicK
    Call Rut_Context_Buttons_Hide   '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"
    
    Sht__BD.Visible = xlSheetVisible
    Sht__BD.Select
    Call Rut_Lo_Filtros_Quitar(Sht__BD.ListObjects(1))
    Prog__APP.Range("SW_Col_Hide_Sht__BD") = False
    
        Dim WrkSht      As Worksheet
    For Each WrkSht In Worksheets
        WrkSht.Visible = xlSheetVisible
        'If WrkSht.CodeName <> ActiveSheet.CodeName Then Sheets(WrkSht.Name).Visible = xlSheetVeryHidden
    Next
    
    Call Rut_Lo_FreezePanes_InmovilizaFxCx(2)
    Call RuT_Al_Abrir_WorkBook
    Call Rut_Menú_ShowAll
Debug.Print "Workbook_Open ---------------------------<<< ThisWorkbook "
End Sub
' ==================================================================================================================================
' ==================================================================================================================================
Sub Workbook_BeforeClose(Cancel As Boolean)
Debug.Print "Workbook_BeforeClose ------------------------- ThisWorkbook --- FIN "
    Application.ScreenUpdating = False
    Sht__BD.Visible = xlSheetVisible
    Sht__BD.Select
    Call Rut_Context_Buttons_Restore    '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
    
    Application.EnableEvents = True:        Prog__APP.Range("SW_Events") = True
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
' ==================================================================================================================================
'Private Sub Workbook_WindowResize(ByVal Wn As Window)
'    Wn.WindowState = xlMaximized
'End Sub
' ==================================================================================================================================

'    --------------------------------------------- excel vba reset right-click menu ------------------
'    There are several right-click menus.
'    On the assumption you are talking about the Cell menu,
'    open the VB Editor (Alt+f11), then the Immediate Window (Ctrl+G),  type the Code:
'                        Application.CommandBars("Cell").Reset
'                        Application.CommandBars("List Range Popup").Reset
' ==================================================================================================================================
'- Please, copy the following event code in ThisWorkbook code module:
' ==================================================================================================================================
Sub Workbook_SheetBeforeRightClick(ByVal SheetX As Object, ByVal Target As Range, Hide_RightClickMenú As Boolean)
    Hide_RightClickMenú = Not Prog__APP.Range("SW_RightClickMenú_Visible")  '- Si Falso: NO hay Right-Click...
Debug.Print "Workbook_SheetBeforeRightClick:   Hide_RightClickMenú = ", Hide_RightClickMenú
' ==================================================================================================================================
'- Para una hoja en concreto, introducir este código en la misma hoja...
' ==================================================================================================================================
'''Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'''        Cancel = Not Prog__APP.Range("SW_RightClick")     '- Si Cancela: NO hay Right-Click...
'''End Sub
' ==================================================================================================================================
End Sub


