Attribute VB_Name = "Rut__Right_Click_VBA"
' Last Rev. 2026-09-19 08:28
 '### code for the ThisWorkbook code sheet ###
    '''Option Explicit
 
    '''Private Sub Workbook_BeforeClose(Cancel As Boolean)
    '''
    '''     'remove our custom menu before we leave
    '''    Run ("NewMenúRightClickCellCell")
    '''
    '''End Sub
 
    '''Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    '''
    '''    Run ("NewMenúRightClickCellCell") '- Elimina otros posible Menús
    '''    Run ("NewMenúRightClickCell") '- Genera el nuevo Menú
    '''
    '''End Sub
 '### code for the ThisWorkbook code sheet - END
 
 '### code for a new module ###
 
'    --------------------------------------------- excel vba reset right-click menu ------------------
'    There are several right-click menus.
'    On the assumption you are talking about the Cell menu,
'    open the VB Editor (Alt+f11),
'    then the Immediate Window (Ctrl+G),
'    type: Code:
'                        Application.CommandBars("Cell").Reset
'    and press enter.
 
Option Explicit
 
' ==================================================================================================================================
Sub Rut_Context_Buttons_Hide()  '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
Debug.Print "Rut_Context_Buttons_Hide"
Dim Item_Bar           As Integer
Dim Count              As Integer
Dim CommBarItem        As CommandBarControl
    ' And for each command bar, iterate through all the available controls
    Item_Bar = 38 '- (38)=("Cell") ----------------------------------------------------------
    For Count = 1 To Application.CommandBars(Item_Bar).Controls.Count
        Set CommBarItem = Application.CommandBars(Item_Bar).Controls(Count)
        If CommBarItem.Tag <> "My_Ctxt_Menu" Then
            CommBarItem.Visible = False
        Else
            Debug.Print Item_Bar & "-" & Count & ". - " & CommBarItem.ID & " - " & CommBarItem.Caption & " - " & CommBarItem.Enabled & " - " & CommBarItem.Visible

        End If
    Next Count
    Item_Bar = 74 '- (74)=("List Range Popup") ---------------------------------------------
    For Count = 1 To Application.CommandBars(Item_Bar).Controls.Count
        Set CommBarItem = Application.CommandBars(Item_Bar).Controls(Count)
        If CommBarItem.Tag <> "My_Ctxt_Menu" Then
            CommBarItem.Visible = False
        Else
            Debug.Print Item_Bar & "-" & Count & ". - " & CommBarItem.ID & " - " & CommBarItem.Caption & " - " & CommBarItem.Enabled & " - " & CommBarItem.Visible

        End If
    Next Count
    Prog__APP.Range("SW_RightClickMenú_Restricted") = True
End Sub
' ==================================================================================================================================
 
' ==================================================================================================================================
Sub Rut_Context_Buttons_Restore()  '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
Debug.Print "Rut_Context_Buttons_Restore"
    Application.CommandBars("Ply").Enabled = True   '- Permite visualizar o NO, el Context-Menú / Right-ClicK
    Application.CommandBars("Cell").Reset           '- Restablece el menú contextual de las celdas a su estado original
    'Esta línea restablece la barra de comandos llamada "Cell" a su configuración original.
    'La barra de comandos "Cell" es la que aparece cuando haces clic derecho en una celda en Excel. _
    ' Al restablecerla, se eliminan cualquier personalización o cambio que se haya hecho a esta barra de comandos, y se vuelve a su estado predeterminado.
    Application.CommandBars("List Range Popup").Reset   '- Restablece el menú contextual de una tabla o en un rango con formato de tabla
    'Esta línea restablece la barra de comandos llamada "List Range Popup" a su configuración original.
    'La barra de comandos "List Range Popup" es la que aparece cuando haces clic derecho en un rango de celdas que forma parte de una lista o tabla en Excel. _
    ' Al restablecerla, se eliminan cualquier personalización o cambio que se haya hecho a esta barra de comandos, y se vuelve a su estado predeterminado.
    
    Prog__APP.Range("SW_RightClickMenú_Restricted") = False
End Sub
' ==================================================================================================================================
' ==================================================================================================================================
' =====================  RuT_Antes_de_Cerrar_WorkBook   ============================================================================
' ==================================================================================================================================
Sub RuT_Antes_de_Cerrar_WorkBook()

    Run ("DelMenúRightClickCell") '- Elimina otros posible Menús XML
    Run ("DelMenúRightClickList") '- Elimina otros posible Menús XML
    
    Application.CommandBars("Cell").Reset
    Application.CommandBars("List Range Popup").Reset
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayScrollBars = True
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayGridlines = True
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
    Application.ScreenUpdating = True
    Application.DisplayFullScreen = True
    
    
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

Private Sub NewMenúRightClickCell()
     
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call DelMenúRightClickCell
    
    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")
    

    '------------ Si definimos el ContextMenu como "Cell" el Custom del Right-Click NO funcionará en Tablas "List Range Popup",
    '---------- y Si definimos el ContextMenu como "List Range Popup" el Right-Click No funcionará en las otras "Cell".
    '
    '    Excel uses a separate right-click menu for Tables.
    '    I only speak VBA, so you will have to translate...
    '    When right-clicking in a "normal" cell the CommandBars("Cell") menu is used.
    '    When right-clicking in a Table the CommandBars("List Range Popup") menu is used.
    
    ' Set ContextMenu to the "List Range Popu"p context menu.
    ' Set ContextMenu = Application.CommandBars("List Range Popup")
    
    ' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, Before:=1)

    With MySubMenu
        .Caption = "Insertar Rectángulo - 1"
        .Tag = "My_Cell_Control_Tag"
       
         'add the submenus
        Dim i As Integer
        For i = 50 To 250 Step 50 'add a few menu items
            Set MySubMenu = ContextMenu.Controls.Add
            With .Controls.Add(Type:=msoControlButton)
    '            .Style = msoButtonIconAndCaption
                .Caption = i & " x " & (i / 2) 'give them a name
                .FaceId = 1311
    '            .TooltipText = "Do Something"
                .Tag = i 'we'll use the tag property to hold a value
                .OnAction = "InsertShape" 'the routine called by the control
            End With
        Next
        
    End With

    ' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, Before:=1)

    With MySubMenu
        .Caption = "Case Menu"
        .Tag = "My_Cell_Control_Tag"

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "UpperMacro"
            .FaceId = 100
            .Caption = "Upper Case"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
            .FaceId = 91
            .Caption = "Lower Case"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "ProperMacro"
            .FaceId = 95
            .Caption = "Proper Case"
        End With
    End With

    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=1)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "ToggleCaseMacro"
        .FaceId = 59
        .Caption = "Toggle Case Upper/Lower/Proper"
        .Tag = "My_Cell_Control_Tag"
    End With

    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=1)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "CeldaColorRojo"
        .FaceId = 184
        .Caption = "Fill Red"
        .Tag = "My_Cell_Control_Tag"
    End With

'    ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, Before:=1
    

    ' Add a separator to the Cell context menu.
    ContextMenu.Controls(5).BeginGroup = True
    
End Sub
 
Private Sub DelMenúRightClickCell()
     
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each MySubMenu In ContextMenu.Controls
        If MySubMenu.Tag = "My_Cell_Control_Tag" Then MySubMenu.Delete
    Next MySubMenu

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
     
End Sub
 
 
Sub NewMenúRightClickList()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call DelMenúRightClickList
    
    ' Set ContextMenu to the Cell context menu.
    ' Set ContextMenu = Application.CommandBars("Cell")

    '------------ Si definimos el ContextMenu como "Cell" el Custom del Right-Click NO funcionará en Tablas "List Range Popup",
    '---------- y Si definimos el ContextMenu como "List Range Popup" el Right-Click No funcionará en las otras "Cell".
    '
    '    Excel uses a separate right-click menu for Tables.
    '    I only speak VBA, so you will have to translate...
    '    When right-clicking in a "normal" cell the CommandBars("Cell") menu is used.
    '    When right-clicking in a Table the CommandBars("List Range Popup") menu is used.
    
    ' Set ContextMenu to the List Range Popup context menu.
    Set ContextMenu = Application.CommandBars("List Range Popup")
    
'    ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, Before:=1
    
    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=2)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "ToggleCaseMacro"
        .FaceId = 59
        .Caption = "Toggle Case Upper/Lower/Proper"
        .Tag = "My_Cell_Control_Tag"
    End With

    ' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, Before:=3)

    With MySubMenu
        .Caption = "Case Menu"
        .Tag = "My_Cell_Control_Tag"

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "UpperMacro"
            .FaceId = 100
            .Caption = "Upper Case"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
            .FaceId = 91
            .Caption = "Lower Case"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "ProperMacro"
            .FaceId = 95
            .Caption = "Proper Case"
        End With
    End With

    ' Add a separator to the Cell context menu.
    ContextMenu.Controls(4).BeginGroup = True
End Sub

Sub DelMenúRightClickList()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    ' Set ContextMenu = Application.CommandBars("Cell")

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("List Range Popup")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each MySubMenu In ContextMenu.Controls
        If MySubMenu.Tag = "My_Cell_Control_Tag" Then MySubMenu.Delete
    Next MySubMenu

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub

 
 
 
'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------
'--------- Rutinas de los Botones del menú del Right-Click ------------------------------
'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------
 
Sub InsertShape()
     
    Dim t As Long
    Dim shp As Shape
     
     'get the tag property of the clicked control
    t = CLng(Application.CommandBars.ActionControl.Tag)
     
     'use the value of t and the active cell as size and position parameters
     'for adding a rectangle to the worksheet
    Set shp = ActiveSheet.Shapes.AddShape _
    (msoShapeRectangle, ActiveCell.Left, ActiveCell.Top, t, t / 2)
     'do something with our shape
    Randomize 'make it a random color from the workbook
    shp.Fill.ForeColor.SchemeColor = Int((56 - 1 + 1) * Rnd + 1)
     
End Sub

Sub CeldaColorRojo()
    Dim cell As Range
    For Each cell In Selection
        cell.Interior.Color = vbRed
    Next cell
End Sub

Sub ToggleCaseMacro()
    Dim CaseRange As Range
    Dim calcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        calcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        Select Case cell.Value
        Case UCase(cell.Value): cell.Value = LCase(cell.Value)
        Case LCase(cell.Value): cell.Value = StrConv(cell.Value, vbProperCase)
        Case Else: cell.Value = UCase(cell.Value)
        End Select
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calcMode
    End With
End Sub

Sub UpperMacro()
    Dim CaseRange As Range
    Dim calcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        calcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.Value = UCase(cell.Value)
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calcMode
    End With
End Sub

Sub LowerMacro()
    Dim CaseRange As Range
    Dim calcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        calcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.Value = LCase(cell.Value)
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calcMode
    End With
End Sub

Sub ProperMacro()
    Dim CaseRange As Range
    Dim calcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Intersect(Selection, Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        calcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.Value = StrConv(cell.Value, vbProperCase)
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calcMode
    End With
End Sub



