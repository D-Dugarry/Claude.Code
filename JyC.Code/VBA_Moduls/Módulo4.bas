Attribute VB_Name = "Módulo4"
Option Explicit

Private myRibUI             As IRibbonUI
Private v_edit_1 As String

Public Property Let ribbonUI(irib As IRibbonUI)
    Set myRibUI = irib
End Property
Public Property Get ribbonUI() As IRibbonUI
    Set ribbonUI = myRibUI
End Property

Public Property Let xl_edit1(xval As String)
    v_edit_1 = xval
End Property
Public Property Get xl_edit1() As String
    xl_edit1 = v_edit_1
End Property

'Private Sub Workbook_Activate()
'Debug.Print "Sub Workbook_Activate"
'    Debug.Print prog__app.range("SW_Events")
'    Call RefreshRibbon
'End Sub
'
Private Sub Workbook_Deactivate()
    If Prog__APP.Range("SW_WB_Deactivate") Then
        Debug.Print "Sub Workbook_Deactivate"
        Call Rut_ConfigExcel_Restablecer
    End If
End Sub

' ==================================================================================================================================
Sub Workbook_Open()
Debug.Print "Workbook_Open --------------------------->>> ThisWorkbook "
'    Call Rut_WrkBook_MinimizeAllWindowsExceptThisExcel
    Call Rut_WrkBook_MinimizeAllExcelExceptThisWB
    Application.Visible = False
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    Application.DisplayFullScreen = True
'    Form_Usuario.Show
    Prog__APP.Range("SW_Events") = True
    Prog__APP.Range("SW_Test") = False
    Prog__APP.Range("SW_RightClick") = True
    Prog__APP.Range("SW_WB_Deactivate") = True
    
    Application.CommandBars("Ply").Enabled = True   '- Permite visualizar el Context-Menú Right-ClicK
    Call Rut_Context_Buttons_Hide   '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"
    
'    Rut_Menú_HideAll
    
        Dim WrkSht      As Worksheet
        H_INICI.Visible = xlSheetVisible
    For Each WrkSht In Worksheets
        If WrkSht.CodeName <> "H_INICI" Then Sheets(WrkSht.Name).Visible = xlSheetVeryHidden
    Next
    
    Call RuT_Al_Abrir_WorkBook
    Call Rut_Menú_ShowAll
Debug.Print "Workbook_Open ---------------------------<<< ThisWorkbook "
End Sub

' ==================================================================================================================================
Sub Workbook_BeforeClose(Cancel As Boolean)
Debug.Print "Workbook_BeforeClose ------------------------- ThisWorkbook --- FIN "
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
    H_INICI.Visible = xlSheetVisible
    Workbooks(ThisWorkbook.Name).Sheets(H_INICI.Name).Select
End Sub
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
Sub Workbook_SheetBeforeRightClick(ByVal SheetX As Object, ByVal Target As Range, Cancel As Boolean)
Debug.Print "Workbook_SheetBeforeRightClick"
    Cancel = Not Prog__APP.Range("SW_RightClick")  '- Si Cancela: NO hay Right-Click...
' ==================================================================================================================================
'- Para una hoja en concreto, introducir este código en la misma hoja...
' ==================================================================================================================================
'''Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'''        Cancel = Not prog__app.range("SW_RightClick")     '- Si Cancela: NO hay Right-Click...
'''End Sub
' ==================================================================================================================================
End Sub



