Attribute VB_Name = "M_000_RibbonUI"
'- M___RibbonUI ---------
Option Explicit

Public MyRibbon                 As IRibbonUI
Public MyTag                    As String
'---------------------------------------------
Public SaveUSB                  As String           '- Controla si se ha hecho una copia
'Public Filter_APP_AñoCont       As String
Public Filter_Siglas            As String
Public Filter_Texto             As String
Public Filter_Importe           As String
Public SW_Tag_Visible           As Boolean          '- Controla visualización de Tags
Public SW_TabExcelVisible       As Boolean          '- Para visualizar los Tags de Excel
Public SW_TrafficLight1         As Boolean
Public SW_TrafficLight2         As Boolean
Public SW_TrafficLight3         As Boolean
Public BtnGrabarSiVisible       As Boolean
Public BtnExportSiVisible       As Boolean

'------------------------------------------------------------------------------------------
Sub OnLoad_MyRibbon(Ribbon As IRibbonUI)
Debug.Print "OnLoad_MyRibbon"
    ThisWorkbook.ribbonUI = Ribbon
    Set MyRibbon = Ribbon
    
    SaveUSB = "No Copy"
    SW_TabExcelVisible = False
    
    BtnGrabarSiVisible = False
    BtnExportSiVisible = False
    Rut_Filtrar_Tareas (Prog__APP.Range("APP_User_ID"))
    MyRibbon.ActivateTab ("TabUserMenu")
'    MyRibbon.ActivateTab ("TabProgMenu")
'    MyRibbon.InvalidateControl control.id
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
Sub RefreshRibbon()
Debug.Print "RefreshRibbon"
    ' Guardas: no intentar si la app esta oculta o el ribbon no esta cargado
    If Not Application.Visible Then Exit Sub
    If MyRibbon Is Nothing Then Exit Sub
    On Error GoTo RestartExcel
        MyRibbon.Invalidate
    On Error GoTo 0
    Exit Sub
RestartExcel:
    MsgBx_Msg = "Please restart Excel for Ribbon UI changes to take effect"
    MsgBx_Title = "Ribbon UI Refresh Failed"
Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Stop"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
End Sub
'------------------------------------------------------------------------------------------
Sub GetVisibleTabExcel(control As IRibbonControl, ByRef Visible)
Debug.Print "GetVisibleTabExcel"
    Visible = SW_TabExcelVisible
End Sub
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetVsbl_UserButtons(control As IRibbonControl, ByRef Visible)
Debug.Print "GetVsbl_UserButtons"
    Visible = Func_Button_View(control.Tag)
End Sub
'------------------------------------------------------------------------------------------
Function Func_Button_View(Ctrl As String)
    Dim Cont_Row    As Integer
    With Prog__MnAux.ListObjects(1).DataBodyRange
        For Cont_Row = 1 To .Rows.Count     '- Recorre la Tabla para ver los que deben verse.
            MyTag = .Cells(Cont_Row, Task_Uribbon_Tags)
            SW_Tag_Visible = .Cells(Cont_Row, Task_Visible)
            If Ctrl Like MyTag Then Func_Button_View = SW_Tag_Visible
            If Ctrl Like MyTag & "_Separ" Then Func_Button_View = SW_Tag_Visible
        Next Cont_Row
    End With
End Function

'------------------------------------------------------------------------------------------
Sub getStip_CtrlTab(control As IRibbonControl, ByRef Supertip)
Debug.Print "getStip_CtrlTab"
    Supertip = Func_SuperTip_Value(control.Tag)
End Sub
'------------------------------------------------------------------------------------------
Function Func_SuperTip_Value(CtrlTag As String)
    Dim Lin_Lst     As Variant
    With Prog__MnAux.ListObjects(1).DataBodyRange
        Lin_Lst = Application.Match(CtrlTag, .Columns(Task_Uribbon_Tags), 0)
        If Not IsError(Lin_Lst) Then
            Func_SuperTip_Value = .Cells(Lin_Lst, Task_Descripción)
            If Prog__APP.Range("SW_Boss") Then
                Func_SuperTip_Value = Func_SuperTip_Value & vbLf & vbLf & _
                                     "BOSS_Rut: " & .Cells(Lin_Lst, Task_Nombre_Rut) & vbLf & _
                                     "Informe:" & vbLf & .Cells(Lin_Lst, Task_Rut_Informe)
            End If
        Else
            Func_SuperTip_Value = "¡ Control.Tag, NO encontrado !"
        End If
    End With
End Function

' ==================================================================================================================================
' ======= TabUserMenu ==============================================================================================================
' ==================================================================================================================================

'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'    "User" -------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_User(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_User"
    LabelVal = Prog__APP.Range("APP_User_Name")
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_CCtxtUser(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_CCtxtUser"
    LabelVal = "Cambiar el Usuario (" & Prog__APP.Range("APP_User_Name") & ")"
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ChangeUser(control As IRibbonControl)
    Call RuT_Ejecutar_Rut("Rut_Chg_Usuario")
Debug.Print "OnAct_ChangeUser"
'    Form_Usuario.Show
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'    "Reset" -------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_BtnReset(control As IRibbonControl)
    Call RuT_Ejecutar_Rut("Rut_Reset_App")
Debug.Print "OnAct_BtnReset"
End Sub
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'   "CTxtRibbonXVisibility" -------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_BtnRibbView(control As IRibbonControl)
Debug.Print "OnAct_BtnRibbView"
    Select Case control.Tag
        Case "RibbViewNone"
            Call Rut_Menú_HideAll
        Case "RibbViewMin"
            Call Rut_Menú_ShowAll_Short
        Case "RibbViewMax"
            Call Rut_Menú_ShowAll
    End Select
    RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
Sub GetVsbl_RibbView(control As IRibbonControl, ByRef Visible)
    Select Case control.Tag
        Case "RibbViewNone"
            Visible = Not SW_TrafficLight1
        Case "RibbViewMin"
            Visible = Not SW_TrafficLight2
        Case "RibbViewMax"
            Visible = Not SW_TrafficLight3
    End Select
Debug.Print "GetVsbl_RibbView", control.id
End Sub
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'   "for dynamic GoTo=Sheet getContent" -------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetContent_H_Sheet_List(control As IRibbonControl, ByRef returnedVal)
Debug.Print "GetContent_H_Sheet_List"
    returnedVal = Func_xmlGen_H_Sheet_List
End Sub
'------------------------------------------------------------------------------------------
Function Func_xmlGen_H_Sheet_List() As String
        Dim WrkSht      As Worksheet
        Dim Sheet_n     As Integer:     Sheet_n = 1
        Dim xml         As String
        Dim ActivSheet  As String:  ActivSheet = ActiveSheet.CodeName
    For Each WrkSht In Worksheets
        If Left(WrkSht.CodeName, 2) = "H_" And WrkSht.CodeName <> ActivSheet Then
            xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='" & WrkSht.Name & "' tag='" & WrkSht.Name & _
                        "' imageMso='GoToNextAppointment' onAction='OnAct_Goto_H_Sheet'/>"
            Sheet_n = Sheet_n + 1
        End If
    Next
        Func_xmlGen_H_Sheet_List = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui'>" & xml & "</menu>"
End Function
'------------------------------------------------------------------------------------------
Sub OnAct_Goto_H_Sheet(control As IRibbonControl)
Debug.Print "OnAct_Goto_H_Sheet"
    Dim ActivSheet  As String
    Application.ScreenUpdating = False
    ActivSheet = ActiveSheet.Name
    Sheets(control.Tag).Visible = xlSheetVisible
    Sheets(control.Tag).Select
    Sheets(ActivSheet).Visible = xlVeryHidden
    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_GoToSheet_LiqCta(control As IRibbonControl)
Debug.Print "OnAct_GoToSheet_LiqCta"
    Dim ActivSheet  As String
    Application.ScreenUpdating = False
    ActivSheet = ActiveSheet.Name
    H_Liq_CTA.Visible = xlSheetVisible
    H_Liq_CTA.Select
    Sheets(ActivSheet).Visible = xlVeryHidden
    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_GoToSheet_LiqTPV(control As IRibbonControl)
Debug.Print "OnAct_GoToSheet_LiqTPV"
    Dim ActivSheet  As String
    Application.ScreenUpdating = False
    ActivSheet = ActiveSheet.Name
    H_Liq_TPV.Visible = xlSheetVisible
    H_Liq_TPV.Select
    Sheets(ActivSheet).Visible = xlVeryHidden
    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'    "Import N43 or TPV" -------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_ImportN43(control As IRibbonControl)
Debug.Print "OnAct_ImportN43"
    Dim ActivSheet  As String:  ActivSheet = ActiveSheet.Name

        Prog__APP.Range("APP_Task_Rut") = "Rut_Import_Cta_N43"
        Form_Running_Rut.Show

'    Dim Lo_MnAux    As ListObject:  Set Lo_MnAux = Workbooks(ThisWorkbook.Name).Sheets(Prog__MnAux.Name).ListObjects(1)
'    TaskIndice = Application.Match("Import_Cta_Norma43", Lo_MnAux.DataBodyRange.Columns(Task_Nombre_Rut), 0)
''    Call Call_Rut_Import_Cta_N43
'    Lo_MnAux.DataBodyRange.Cells(TaskIndice, Task_Rut_Informe) = prog__app.range("APP_Task_Inf")
    Sheets(ActivSheet).Select
'    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ImportTPV(control As IRibbonControl)
Debug.Print "OnAct_ImportTPV"
    Dim ActivSheet  As String:  ActivSheet = ActiveSheet.Name
    Call Call_Rut_Import_TPV_Pagos
    Sheets(ActivSheet).Select
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_ImportDataCTA(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_ImportDataCTA"
    LabelVal = "Last Imp." & vbLf & Format(Prog__APP.Range("APP_Date_Imp_Cta"), "dd-mmm-yy")
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_ImportDataTPV(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_ImportDataTPV"
    LabelVal = "Last Imp." & vbLf & Format(Prog__APP.Range("APP_Date_Imp_TPV"), "dd-mmm-yy")
End Sub

'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'    "Group Show-Hide_Lo_Cols" -------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetVsbl_Group_ShowHide_Lo_Cols(control As IRibbonControl, ByRef Visible)
    Dim SheetName As String:    SheetName = "DefCol_" & ActiveSheet.Name
    If Fnc_WrkSheet_Exist(SheetName) Then
        Visible = True
    Else
        Visible = False
    End If
Debug.Print "GetVsbl_Group_ShowHide_Lo_Cols,     Visible= " & Visible
End Sub
'-----------------------------
Sub GetLbl_Group_ShowHide_Lo_Cols(control As IRibbonControl, ByRef LabelVal)
    LabelVal = "Table " & ActiveSheet.Name
Debug.Print "GetLbl_Group_ShowHide_Lo_Cols"
End Sub
'-----------------------------
Sub GetLbl_Btn_ShowHide_Lo_Cols(control As IRibbonControl, ByRef LabelVal)
    Dim SheetName As String:    SheetName = "DefCol_" & ActiveSheet.Name
    If Fnc_WrkSheet_Exist(SheetName) Then
        Application.ScreenUpdating = False
        Dim Lo_DefCol       As ListObject:  Set Lo_DefCol = Workbooks(ThisWorkbook.Name).Sheets("DefCol_" & ActiveSheet.Name).ListObjects(1)
        Dim Col_HiddenSw    As Integer:     Col_HiddenSw = Lo_DefCol.ListColumns("HiddenCol").Range.Column
        If Lo_DefCol.TotalsRowRange(Col_HiddenSw) Then
            LabelVal = "Show Cols"
            Call Rut_LstObj_Columns_Show_Hide(Sheets(ActiveSheet.Name), Sheets("DefCol_" & ActiveSheet.Name), True)
        Else
            LabelVal = "Hide Cols"
        End If
        '- Para Ajustar un grupo de columnas al ancho de la window.
        Select Case ActiveSheet.Name
        Case Prog_CTA_Tb.Name
            Sheets(ActiveSheet.Name).ListObjects(1).HeaderRowRange(1).Resize(, C_Cta_Siglas).Select
            ActiveWindow.Zoom = True
        End Select
        Sheets(ActiveSheet.Name).ListObjects(1).Range(1).Select
        Application.ScreenUpdating = True
    End If
Debug.Print "GetLbl_Btn_ShowHide_Lo_Cols,  LabelVal=", LabelVal
End Sub
'-----------------------------
Sub OnAct_Btn_ShowHide_Lo_Cols(control As IRibbonControl)
    Application.ScreenUpdating = False
    Call Rut_LstObj_Columns_Show_Hide(Sheets(ActiveSheet.Name), Sheets("DefCol_" & ActiveSheet.Name))
    '- Para Ajustar un grupo de columnas al ancho de la window.
    Select Case ActiveSheet.Name
    Case Prog_CTA_Tb.Name
        Sheets(ActiveSheet.Name).ListObjects(1).HeaderRowRange(1).Resize(, C_Cta_Siglas).Select
        ActiveWindow.Zoom = True
    End Select
    Sheets(ActiveSheet.Name).ListObjects(1).Range(1).Select
    RefreshRibbon
    Application.ScreenUpdating = True
Debug.Print "OnAct_ShowHide_Lo_Cols"
End Sub

'===========================================================================================
'===========================================================================================
'======== Sheet INICIO =====================================================================
'===========================================================================================
Sub GetLbl_Sheet_XXX(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_Sheet_XXX"
    LabelVal = "Sheet " & ActiveSheet.Name
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_Lo_Sheet_XXX(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_Lo_Sheet_XXX"
    LabelVal = "Sheet " & ActiveSheet.Name
End Sub

'------------------------------------------------------------------------------------------
'----------  Sheet INICIO  ----------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_Show_Lo_JyC_Summary(control As IRibbonControl)
'    Application.ScreenUpdating = False
    Call Rut_Inicio_Lo_JyC_Summary
    Application.ScreenUpdating = True
    MsgBx_Msg = "¡ Proceso completado !"
    MsgBx_Title = "Proceso: Actualizar Tabla Resumen de JyC."
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "Msg"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
Debug.Print "OnAct_Refresh_Lo_Ininio", control.id
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_Show_Lo_Liq_Summary(control As IRibbonControl)
'    Application.ScreenUpdating = False
    Call Rut_Inicio_Lo_Liq_Summary
    Application.ScreenUpdating = True
    MsgBx_Msg = "¡ Proceso completado !"
    MsgBx_Title = "Proceso: Actualizar Tabla Resumen de Liquidaciones de JyC."
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "Msg"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
Debug.Print "OnAct_Refresh_Lo_Ininio", control.id
End Sub

'------------------------------------------------------------------------------------------
'----------  Sheet Cta_Tb  -----------------------------------------------------------------
'------------------------------------------------------------------------------------------
'==========================================================================================
Sub GetLbl_Filter_Sheet_XXX(control As IRibbonControl, ByRef LabelVal)
    LabelVal = "Filters for the Sheet: " & ActiveSheet.Name
Debug.Print "GetLbl_Filter_Sheet_XXX"
End Sub
'============ Filtro Año Cont ====================================================================
Sub GetText_EditBoxAñoCont(control As IRibbonControl, ByRef returnedVal)
    'If Filter_APP_AñoCont = "" Then Filter_APP_AñoCont = Prog__APP.Range("APP_AñoCont")
    returnedVal = Prog__APP.Range("APP_AñoCont")
'    Call Rut_Filter_AñoCont(Prog__APP.Range("APP_AñoCont"))
Debug.Print "GetText_EditBoxAñoCont, returnedVal= ", returnedVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnChange_EditBoxAñoCont(control As IRibbonControl, FiltroAño As String)
Debug.Print "OnChange_EditBoxAñoCont "
    Prog__APP.Range("APP_AñoCont") = FiltroAño
    Call Rut_Filter_AñoCont(FiltroAño)
End Sub
'------------------------------------------------------------------------------------------
Sub Rut_Filter_AñoCont(FiltroAño As String)
    Dim nFilas As Variant
    Dim Lo_Cta      As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    Call Rut_LstObj_Filtros_Quitar(Lo_Cta)
    Call Rut_LstObj_Sort(Lo_Cta, C_Cta_Ordinal, xlAscending, True)
    With Lo_Cta
        If FiltroAño <> "" Then
            .Range.AutoFilter Field:=C_Cta_Ordinal, Criteria1:=">=" & FiltroAño & "000000", Operator:=xlAnd, _
                                            Criteria2:="<" & FiltroAño + 1 & "000000" ', Operator:=xlFilterValues   '- Filtrar
            nFilas = .Range.Columns(C_Cta_Ordinal).SpecialCells(xlCellTypeVisible).Cells.Count - 1 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
            If nFilas > 0 Then
                .DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
            Else
                MsgBx_Msg = "¡¡ No hay Apuntes de: [  " & FiltroAño & "  ] !!" & vbLf & vbLf & "¡ Vuelve a intentarlo !"
                MsgBx_Title = "Proceso: Buscar un pago por transferencia a la cuenta de JyC."
                Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "Exclam"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
                Call Rut_LstObj_Filtros_Quitar(Lo_Cta)
            End If
        Else
            .Range.AutoFilter Field:=C_Cta_Ordinal
        End If
    End With
    Prog__APP.Range("APP_AñoCont") = FiltroAño
Debug.Print "Rut_Filter_AñoCont"
End Sub

'==========================================================================================
'============= dynamic=DropDown Select Siglas_JyC =========================================
'==========================================================================================
Sub GetDDItemCount_DropDown_Siglas(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Prog_JyC_List.ListObjects(1).ListRows.Count
Debug.Print "GetDDItemCount_DropDown_Siglas"
End Sub
'------------------------------------------------------------------------------------------
Sub GetDDItemLabel_DropDown_Siglas(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = Prog_JyC_List.ListObjects(1).DataBodyRange.Cells(index, 1).Value
Debug.Print "GetDDItemLabel_DropDown_Siglas"
End Sub
'------------------------------------------------------------------------------------------
Sub GetDDItemImage_DropDown_Siglas(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = "GoToNextAppointment"
Debug.Print "GetDDItemImage_DropDown_Siglas"
End Sub
'------------------------------------------------------------------------------------------
Sub DDClicked_DropDown_Siglas(control As IRibbonControl, id As String, index As Integer)
    Dim FiltroSigla As String: FiltroSigla = Prog_JyC_List.ListObjects(1).DataBodyRange.Cells(index + 0, 1).Value
    Dim nFilas      As Variant
    Dim Lo_Cta      As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    With Lo_Cta
        Call Rut_LstObj_Filtros_Quitar(Lo_Cta)
        If index = 0 Then
            .Range.AutoFilter Field:=C_Cta_Siglas
            Filter_Siglas = "*"
        Else
            .Range.AutoFilter Field:=C_Cta_Siglas, Criteria1:=FiltroSigla  ', Operator:=xlFilterValues    '- Filtrar
        End If
        nFilas = .Range.Columns(C_Cta_Siglas).SpecialCells(xlCellTypeVisible).Cells.Count - 1 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
        If nFilas > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
            Filter_Siglas = FiltroSigla
        Else
            MsgBx_Msg = "¡¡ No hay Apuntes de: [  " & FiltroSigla & "  ] !!" & vbLf & vbLf & "¡ Vuelve a intentarlo !"
            MsgBx_Title = "Proceso: Buscar un pago por transferencia a la cuenta de JyC."
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "Exclam"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
            Call Rut_LstObj_Filtros_Quitar(Lo_Cta)
        End If
    End With
'    MyRibbon.InvalidateControl control.id
Debug.Print "DDClicked_DropDown_Siglas"
End Sub
'------------------------------------------------------------------------------------------
Sub GetSelectItem_DropDown_Siglas(control As IRibbonControl, ByRef returnedVal)
    returnedVal = 0
Debug.Print "GetSelectItem_DropDown_Siglas"
End Sub
'==========================================================================================
'=========== Filtro Texto a Buscar ========================================================
Sub GetText_EditBoxBuscar(control As IRibbonControl, ByRef returnedVal)
    Filter_Texto = returnedVal
Debug.Print "GetText_EditBoxBuscar"
End Sub
'------------------------------------------------------------------------------------------
Sub OnChange_EditBoxBuscar(control As IRibbonControl, FiltroTxt As String)
    Dim nFilas As Variant
    Dim Lo_Cta      As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    Prog_CTA_Tb.Unprotect
    Lo_Cta.ShowTotals = False
'    Lo_Cta.Range.AutoFilter Field:=C_Cta_Ref1
'    Lo_Cta.Range.AutoFilter Field:=C_Cta_Reg_Mov1        ' NO PONGO: ".Range.AutoFilter.ShowAllData"
'    Lo_Cta.Range.AutoFilter Field:=C_Cta_Reg_Mov2       ' Porque quiero mantener filtros en otras columnas.
'    Lo_Cta.Range.AutoFilter Field:=C_Cta_Reg_Mov3
'    Lo_Cta.Range.AutoFilter Field:=C_Cta_Reg_Mov4
'    Lo_Cta.Range.AutoFilter Field:=C_Cta_Reg_Mov5
    Lo_Cta.Range.AutoFilter Field:=C_Cta_Fusión
    Lo_Cta.Range.AutoFilter Field:=C_Cta_Siglas
    Lo_Cta.Range.AutoFilter Field:=C_Cta_Siglas, Criteria1:="_Desconocido", Operator:=xlOr, _
                                                 Criteria2:=Filter_Siglas
    Lo_Cta.Range.AutoFilter Field:=C_Cta_Fusión, Criteria1:="*" & FiltroTxt & "*" ', Operator:=xlFilterValues
        On Error Resume Next
    nFilas = Lo_Cta.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
        On Error GoTo 0
    If nFilas = 0 Then
        MsgBx_Msg = "¡¡ No he encontrado coincidencia !!" & vbLf & vbLf & "¡ Vuelve a intentarlo !"
        MsgBx_Title = "Proceso: Buscar un pago por transferencia a la cuenta de JyC."
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "Exclam"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
        Lo_Cta.Range.AutoFilter Field:=C_Cta_Fusión
    Else
        Lo_Cta.DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1, C_Cta_N_Liq).Select
    End If
'    Call RefreshRibbon
Debug.Print "OnChange_EditBoxBuscar"
End Sub
'------------------------------------------------------------------------------------------
'=========== Filter_Reapply ====================================================================
Sub OnAct_Filter_Reapplay(control As IRibbonControl)
    Call Rut_LstObj_Filtros_Quitar(ActiveSheet.ListObjects(1))
    Dim nFilas As Variant
    Dim Lo_Cta      As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    
    Lo_Cta.Range.AutoFilter Field:=C_Cta_Fusión, Criteria1:="*" & FiltroTxt & "*", Operator:=xlFilterValues
        On Error Resume Next
    nFilas = Lo_Cta.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
        On Error GoTo 0
    If nFilas = 0 Then
        MsgBx_Msg = "¡¡ No he encontrado coincidencia !!" & vbLf & vbLf & "¡ Vuelve a intentarlo !"
        MsgBx_Title = "Proceso: Buscar un pago por transferencia a la cuenta de JyC."
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "Exclam"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
        Lo_Cta.Range.AutoFilter Field:=C_Cta_Fusión
    Else
        Lo_Cta.DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1, C_Cta_N_Liq).Select
    End If
Debug.Print "OnAct_Filter_Reapplay"
Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
'============ Filter_ShowAlData ===============================================================
Sub OnAct_Filter_ShowAlData(control As IRibbonControl)
    Call Rut_LstObj_Filtros_Quitar(ActiveSheet.ListObjects(1))
'    Filter_APP_AñoCont = Prog__APP.Range("APP_AñoCont")
    Filter_Siglas = ""
    RefreshRibbon
Debug.Print "OnAct_Filter_ShowAlData"
    Call Rut_Filter_AñoCont(Prog__APP.Range("APP_AñoCont"))
End Sub
'============ Filtro Importe =====================================================================
Sub GetText_EB_FindImp(control As IRibbonControl, ByRef returnedVal)
    Filter_Importe = returnedVal
Debug.Print "GetText_EB_FindImp", control.id, returnedVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnChange_EB_FindImp(control As IRibbonControl, FiltroImporte As String)
    Dim nFilas As Variant
    Dim Lo_Cta      As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    With Lo_Cta
        .Range.AutoFilter Field:=C_Cta_Imp   ' NO PORGO: ".Range.AutoFilter.ShowAllData" Porque quiero mantener filtros en otras columnas.
        If FiltroImporte <> "" Then
            If IsNumeric(Left(FiltroImporte, 1)) Then
            End If
            .Range.AutoFilter Field:=C_Cta_Imp, Criteria1:=FiltroImporte ' & "*" ', Operator:=xlFilterValues      '- Filtrar
            
            nFilas = .Range.Columns(C_Cta_Liq_Ordinal).SpecialCells(xlCellTypeVisible).Cells.Count - 1 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
            If nFilas > 0 Then
                .DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
            Else
                MsgBx_Msg = "¡¡ No hay Apuntes de importe: [  " & FiltroImporte & "  ] !!" & vbLf & vbLf & "¡ Vuelve a intentarlo !"
                MsgBx_Title = "Proceso: Buscar un pago por transferencia a la cuenta de JyC."
                Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "Exclam"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
                Call Rut_LstObj_Filtros_Quitar(Lo_Cta)
            End If
        End If
    End With
Debug.Print "OnChange_EB_FindImp", control.id
'    Call RefreshRibbon
End Sub





'==========================================================================================
'==========  TPV  =========================================================================
'==========================================================================================
Sub OnAct_ImportTPVSolic(control As IRibbonControl)
Debug.Print "OnAct_ImportTPVSolic"
    Call Import_TPV_Solic
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_CargarTPVSolic(control As IRibbonControl)
Debug.Print "OnAct_CargarTPVSolic"
    Call Load_TPV_Solic
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_GrabarTPVSolic(control As IRibbonControl)
Debug.Print "OnAct_GrabarTPVSolic"
    Call RuT_Grabar_Datos_TPV_Liq_a_TPV_y_Lst
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_CargarTPVLiq(control As IRibbonControl)
Debug.Print "OnAct_CargarTPVLiq"
    Call RuT_Generar_Tabla_Liquidación_TPV
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ExportTPVLiq(control As IRibbonControl)
Debug.Print "OnAct_ExportTPVLiq"
    Call RuT_UsedRange_Save_New_WorkBook_Liq_TPV
End Sub
'==========================================================================================
'==========  TPV=Mail  ====================================================================
'==========================================================================================
Sub OnAct_BtnMail_DoTPVRDT(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_DoTPVRDT"
    Call Rut_Email_TPV_Do_RDT
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnMail_RDTTPVdone(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_RDTTPVdone"
    Call Rut_Email_TPV_RDT_Done
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnMail_DoTPVDev(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_DoTPVDev"
    Call Rut_Email_TPV_do_Devol
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnMail_DevTPVdone(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_DevTPVdone"
    Call Rut_Email_TPV_DEV_Done
End Sub
'==========================================================================================
'==========  Cta  =========================================================================
'==========================================================================================
Sub OnAct_CargarCtaLiq(control As IRibbonControl)
Debug.Print "OnAct_CargarCtaLiq"
    Call RuT_Generar_Tabla_Liquidación_Cta
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_GrabarCtaSolic(control As IRibbonControl)
Debug.Print "OnAct_GrabarCtaSolic"
    Call RuT_Grabar_Datos_CTA_Liq_a_CTA_y_Lst
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ExportCtaLiq(control As IRibbonControl)
Debug.Print "OnAct_ExportCtaLiq"
    Call RuT_UsedRange_Save_New_WorkBook_Liq_CTA
End Sub
'==========================================================================================
'==========  Cta=Mail  ====================================================================
'==========================================================================================
Sub OnAct_BtnMail_DoCtaRDT(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_DoCtaRDT"
    Call Rut_Email_CTA_Do_RDT
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnMail_RDTCtadone(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_RDTCtadone"
    Call Rut_Email_CTA_RDT_Done
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnMail_DoCtaDev(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_DoCtaDev"
    Call Rut_Email_CTA_do_Devol
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnMail_DevCtadone(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_DevCtadone"
    Call Rut_Email_CTA_DEV_Done
End Sub
'------------------------------------------------------------------------------------------
'==========================================================================================
'==========  Mail, Solicitud Anulada  =====================================================
'==========================================================================================
Sub OnAct_BtnMail_SolAnulada(control As IRibbonControl)
Debug.Print "OnAct_BtnMail_SolAnulada"
    Call Rut_Email_RDT_ANULADA
End Sub
'------------------------------------------------------------------------------------------


'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'   Programation's Buttons =======================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================
'==========================================================================================

Sub GetVsbl_TabProgMenu(control As IRibbonControl, ByRef Visible)
    If Prog__APP.Range("APP_User_ID") = "Boss" Then Visible = True Else Visible = False
Debug.Print "GetVsbl_TabProgMenu", control.id
End Sub

'------------------------------------------------------------------------------------------
'   "GroupMenuAUX" ------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_BtnMenuAux(control As IRibbonControl)
Debug.Print "OnAct_BtnMenuAux"
'    Rut_Btn_Menú_Aux
    Form_Menu.Show
    Application.ScreenUpdating = True
    DoEvents
End Sub
'------------------------------------------------------------------------------------------
Sub GetVsbl_ProgButtons(control As IRibbonControl, ByRef Visible)
Debug.Print "GetVsbl_ProgButtons"
    If Prog__APP.Range("APP_User_ID") = "Boss" Then Visible = True Else Visible = False
    'Debug.Print "GetVsbl_ProgButtons", control.id
End Sub

'------------------------------------------------------------------------------------------
'- GroupExcelTabsVisible ------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_GroupExcelTabsVisible(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_GroupExcelTabsVisible"
    If SW_TabExcelVisible Then
        LabelVal = "Tabs Show"
    Else
        LabelVal = "Tabs Hide"
    End If
End Sub
Sub GetLbl_CCtxtBtnExcelTabsVisible(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_CCtxtBtnExcelTabsVisible"
    If SW_TabExcelVisible Then
        LabelVal = "Excel Tabs Shown, Press to Hide them"
    Else
        LabelVal = "Excel Tabs Hidden, Press to Show them"
    End If
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnExcelTabsVisible(control As IRibbonControl)
Debug.Print "OnAct_BtnExcelTabsVisible"
    SW_TabExcelVisible = Not SW_TabExcelVisible
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
'- GroupResetExcelConfig ------------------------------------------------------------------------
Sub OnAct_BtnResetExcelConfig(control As IRibbonControl)
Debug.Print "OnAct_BtnResetExcelConfig"
    Call Rut_ConfigExcel_Restablecer
End Sub
'------------------------------------------------------------------------------------------
'- GroupSW_Prueba ------------------------------------------------------------------------
Sub GetLbl_GroupSW_Prueba(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_GroupSW_Prueba"
    If Prog__APP.Range("SW_Test") Then
        LabelVal = "ON"
    Else
        LabelVal = "OFF"
    End If
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_CCtxtBtnSW_PruebaONOFF(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_CCtxtBtnSW_PruebaONOFF"
    If Prog__APP.Range("SW_Test") Then
        LabelVal = "Sistema en Prueba ON, Press to OFF"
    Else
        LabelVal = "Sistema en Prueba OFF, Press to ON"
    End If
End Sub
'------------------------------------------------------------------------------------------
Sub GetVsbl_BtnSW_PruebaONOFF(control As IRibbonControl, ByRef Visible)
Debug.Print "GetVsbl_BtnSW_PruebaONOFF"
    If control.id = "BtnSW_PruebaON" Then
        Visible = Not Prog__APP.Range("SW_Test")
    Else
        Visible = Prog__APP.Range("SW_Test")
    End If
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnSW_PruebaONOFF(control As IRibbonControl)
Debug.Print "OnAct_BtnSW_PruebaONOFF"
    Prog__APP.Range("SW_Test") = Not Prog__APP.Range("SW_Test")
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
'- GroupRightClickMenú ------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_GroupRightClickMenú(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_GroupRightClickMenú"
    If Prog__APP.Range("SW_RightClick") Then
        LabelVal = "Allowed"
    Else
        LabelVal = "Restricted"
    End If
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_CCtxtBtnRightClickMenú(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_CCtxtBtnRightClickMenú"
    If Prog__APP.Range("SW_RightClick") Then
        LabelVal = "Context Menú on Right-Click is Allowed, Press to Restrict"
    Else
        LabelVal = "Context Menú on Right-Click is Restricted, Press to Allow"
    End If
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_RightClickMenú(control As IRibbonControl)
    Prog__APP.Range("SW_RightClick") = Not Prog__APP.Range("SW_RightClick")
    Call RefreshRibbon
     Application.CommandBars("Ply").Enabled = Prog__APP.Range("SW_RightClick")
Debug.Print "Right Click", Prog__APP.Range("SW_RightClick")
End Sub
'------------------------------------------------------------------------------------------
'======= SaveTimer_USB ========================================================================
Sub GetLbl_GroupSaveTimer_USB(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_GroupSaveTimer_USB"
LabelVal = SaveUSB
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_CCtxtBtnSaveTimer_USB(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_CCtxtBtnSaveTimer_USB"
    If SaveUSB = "No Copy" Then
        LabelVal = "Not still Copied, Press to Copy"
    Else
        LabelVal = "Last Copy: " & SaveUSB & ", Press to Copy again"
    End If
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnSaveData_Timer_USB(control As IRibbonControl)
Debug.Print "OnAct_BtnSaveData_Timer_USB"
    Call Rut_WrkBook_CopSegTimed_USB("_Data")
    SaveUSB = Format(Now(), "dd-mmm-yy hh:mm")
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_BtnSaveVBA_Timer_USB(control As IRibbonControl)
Debug.Print "OnAct_BtnSaveVBA_Timer_USB"
    Call Rut_WrkBook_CopSegTimed_USB("_VBA")
    SaveUSB = Format(Now(), "dd-mmm-yy hh:mm")
    Call RefreshRibbon
End Sub

'------------------------------------------------------------------------------------------
'- GroupExportVBA -------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_BtnExportVBA(control As IRibbonControl)
Debug.Print "OnAct_BtnExportVBA"
    Call Rut_VBA_Export_Moduls
End Sub





'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'--------- dynamic-Menú GoTo=Sheet --------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetContent_Sheet_List(control As IRibbonControl, ByRef returnedVal)
Debug.Print "GetContent_Sheet_List"
    returnedVal = Func_xmlGen_Sheet_List
End Sub
'------------------------------------------------------------------------------------------
Function Func_xmlGen_Sheet_List() As String
        Dim WrkSht      As Worksheet
        Dim Sheet_n     As Integer:     Sheet_n = 1
        Dim xml         As String
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Show All' tag='Show All'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Hide All' tag='Hide All'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Show Prog_*' tag='Prog_#'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Show DefCol_*' tag='DefCol_#'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Show H_Sheets*' tag='H_#'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<menuSeparator  id='Separ_XX' />"
    For Each WrkSht In Worksheets
        xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='" & WrkSht.Name & "' tag='" & WrkSht.Name & _
                    "' imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet'/>"
        Sheet_n = Sheet_n + 1
    Next
        Func_xmlGen_Sheet_List = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui'>" & xml & "</menu>"
Debug.Print "Func_xmlGen_Sheet_List"
End Function
'------------------------------------------------------------------------------------------
Sub OnAct_Goto_Sheet(control As IRibbonControl)
Debug.Print "OnAct_Goto_Sheet"
    Dim ActivSheet  As String:     ActivSheet = ActiveSheet.Name
    Dim WrkSht      As Worksheet
    Application.ScreenUpdating = False
    Select Case control.Tag
        Case "Show All"
            For Each WrkSht In Worksheets
                WrkSht.Visible = xlSheetVisible
            Next
        Case "Hide All"
            For Each WrkSht In Worksheets
                If WrkSht.Name <> ActivSheet Then WrkSht.Visible = xlSheetVeryHidden
            Next
        Case "Prog_#"
            For Each WrkSht In Worksheets                      '- 0 = xlSheetHidden, 2 = xlSheetVeryHidden, -1 = xlSheetVisible
                If WrkSht.CodeName Like "Prog_*" Then If WrkSht.Visible < 0 Then WrkSht.Visible = 2 Else WrkSht.Visible = -1
            Next
        Case "DefCol_#"
            For Each WrkSht In Worksheets
                If WrkSht.CodeName Like "*DefCol*" Then If WrkSht.Visible < 0 Then WrkSht.Visible = 2 Else WrkSht.Visible = -1
            Next
        Case "H_#"
            For Each WrkSht In Worksheets
                If Left(WrkSht.CodeName, 2) = "H_" Then If WrkSht.Visible < 0 Then WrkSht.Visible = 2 Else WrkSht.Visible = -1
            Next
        Case Else
            If Sheets(control.Tag).Visible = xlSheetVisible Then
                H_INICI.Visible = xlSheetVisible
                H_INICI.Select
                Sheets(control.Tag).Visible = xlVeryHidden
            Else
                Sheets(control.Tag).Visible = xlSheetVisible
                Sheets(control.Tag).Select
            End If
    End Select
    Application.ScreenUpdating = True
End Sub


'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

''------------------------------------------------------------------------------------------
''------------------------------------------------------------------------------------------
''------------------------------------------------------------------------------------------
''
'' There are several right-click menus. On the assumption you are talking about the Cell menu,
'' open the VB Editor (Alt+f11), then the Immediate Window (Ctrl+G), type the Code:
''
''              Application.CommandBars("Cell").Reset
''              Application.CommandBars("List Range Popup").Reset
''------------------------------------------------------------------------------------------
''------------------------------------------------------------------------------------------
''------------------------------------------------------------------------------------------




''------------------------------------------------------------------------------------------
'Sub GetVsbl_Buttons(control As IRibbonControl, ByRef visible)
'    Select Case control.Tag
'        Case "BtnMailRDT_Plan"
'            visible = Func_TagVisible("BtnMailRDT_Plan")
'        Case "BtnMailRDT_Null"
'            visible = Func_TagVisible("BtnMailRDT_Null")
'        Case "BtnMailRDT_kkk"
'            visible = Func_TagVisible("BtnMailRDT_kkk")
'    End Select
'Debug.Print "GetVsbl_Buttons", control.ID, "_" & MyTag & "_", SW_Tag_Visible
'End Sub
''------------------------------------------------------------------------------------------
'    Function Func_TagVisible(Tag As String)
'        Dim Cont_Row    As Integer
'        With Prog__MnAux.ListObjects(1).DataBodyRange
'            For Cont_Row = 1 To .Rows.Count
'                MyTag = .Cells(Cont_Row, 7)
'                SW_Tag_Visible = .Cells(Cont_Row, 8)
'                If Tag Like MyTag Then Func_TagVisible = SW_Tag_Visible
'            Next Cont_Row
'        End With
'    End Function
''------------------------------------------------------------------------------------------


'########## dynamicMenu id="dynamicShortRuts" ##################################################>>>>>
Sub getContent_dynamicShortRuts(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Func_xmlGen
Debug.Print "Rut: GetMenuContent1: " & Func_xmlGen
End Sub
'--------------------------------------------------------------------------------------------------
    Function Func_xmlGen() As String
        Dim Cont_Row        As Integer
        Dim Num_Tarea       As Integer:     Num_Tarea = 1
        Dim Usuario_ID      As String:      Usuario_ID = Prog__APP.Range("APP_User_ID")
        Dim xml As String
        '- Recorremos toda la tabla de Tareas y seleccionamos las que empiezan por "9_ "
        With Prog__MnAux.ListObjects("Tb_Tareas").DataBodyRange
            For Cont_Row = 1 To .Rows.Count
                If (InStr((.Cells(Cont_Row, 2)), Usuario_ID) > 0 And Len(Usuario_ID) > 3 Or .Cells(Cont_Row, 2) = "") And Left(.Cells(Cont_Row, 1), 3) = "9_ " Then
                    If Num_Tarea > 1 Then xml = xml & "<menuSeparator  id='Separ" & Num_Tarea & "' />"
                    xml = xml & "<button id='button" & Num_Tarea & "' tag='Task-Row_" & Format(Cont_Row, "00") & "' label='" & Mid(.Cells(Cont_Row, 1), 4) & "' imageMso='GoToNextAppointment'" & " onAction='OnAction_Dynamic_Task'/>"
                    Num_Tarea = Num_Tarea + 1
                End If
            Next Cont_Row
        End With
        Func_xmlGen = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui'>" & xml & "</menu>"
    End Function
'---------------------------------------------------------------------------------------------------
'--- Extrae del Tag el núm de row de la Tarea para obtener el nombre de la Rutina. -----------------
'---------------------------------------------------------------------------------------------------
Sub OnAction_Dynamic_Task(control As IRibbonControl)
Debug.Print "OnAction_Dynamic_Task"
    Dim Pos_Delimitador     As Integer
    Dim Rutinas_Name        As String
    Dim Rut_Name        As String
    TaskIndice = Val(Right(control.Tag, 2))
    Rutinas_Name = Prog__MnAux.ListObjects(1).DataBodyRange.Cells(TaskIndice, 3)
    Do       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Pos_Delimitador = InStr(Rutinas_Name, " + ")
        If Pos_Delimitador < 1 Then ' ------------------ Última o Única Rutina ---------------------
            Pos_Delimitador = Len(Rutinas_Name) + 1
        End If
        Rut_Name = Left(Rutinas_Name, Pos_Delimitador - 1)
        Rutinas_Name = Mid(Rutinas_Name, Pos_Delimitador + 3)
        Application.Run Rut_Name
        Prog__MnAux.ListObjects(1).DataBodyRange.Cells(TaskIndice, 5) = Prog__APP.Range("APP_Task_Inf")
    Loop While Len(Rutinas_Name) > 0
    Application.ScreenUpdating = True
    DoEvents
End Sub
'########## Dynamic-Menú #######################################################################<<<<<








