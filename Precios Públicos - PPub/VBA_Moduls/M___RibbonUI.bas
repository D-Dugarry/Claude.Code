Attribute VB_Name = "M___RibbonUI"
'2026-01-01
'- M___RibbonUI ---------
Option Explicit

' API CopyMemory 64 bits
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" _
                            Alias "RtlMoveMemory" ( _
                                ByRef Destination As Any, _
                                ByRef Source As Any, _
                                ByVal Length As LongPtr)

' Referencia global al Ribbon
Public gRibbon                  As IRibbonUI

Public MyRibbon                 As IRibbonUI
Public MyTag                    As String
'---------------------------------------------
Public SaveUSB                  As String           '- Controla si se ha hecho una copia
Public SW_Tag_Visible           As Boolean          '- Controla visualización de Tags
Public SW_TabExcelVisible       As Boolean          '- Para visualizar los Tags de Excel
Public SW_TrafficLight1         As Boolean
Public SW_TrafficLight2         As Boolean
Public SW_TrafficLight3         As Boolean
Public BtnGrabarSiVisible       As Boolean
Public BtnExportSiVisible       As Boolean

Public Filter_APP_AñoCont       As String
Public Filter_Siglas            As String
Public Filter_Texto             As String
Public Filter_Importe           As String

Public PlanImport               As String

'------------------------------------------------------------------------------------------
Sub OnLoad_MyRibbon(Ribbon As IRibbonUI)
Debug.Print "--------------------------------------------------- >>> OnLoad_MyRibbon"
    Dim lPtr As LongPtr
    
    ThisWorkbook.ribbonUI = Ribbon
    Set MyRibbon = Ribbon
    ' Guardar el puntero en un Name del libro
    lPtr = ObjPtr(Ribbon)
    ThisWorkbook.Names.Add Name:="RibbonPtr", RefersTo:="=" & CStr(lPtr)
    'ThisWorkbook.Names.Add Name:="RibbonID", RefersTo:="=" & ObjPtr(Ribbon)
    
    SaveUSB = "No Copy"
    SW_TabExcelVisible = False
    
    '--- Activo los valores de los Cursos Académicos ------------------------------------>>>>
    If Prog__APP.Range("APP_CursAcad") = Prog__APP.Range("APP_C_Acad_Ant") Then
            SW_C_Acad_Ant = True
            SW_C_Acad_Pos = False
        MyTag = "checkBoxAcad1"
    Else
            SW_C_Acad_Pos = True
            SW_C_Acad_Ant = False
        MyTag = "checkBoxAcad2"
    End If   '---------------------------------------------------------------------------<<<<
    If ActiveSheet.Comments.Count = 0 Then
        Prog__APP.Range("APP_Help_State") = "'Desactivados'"
    Else
        Prog__APP.Range("APP_Help_State") = "'Activados'"
    End If
    
    BtnGrabarSiVisible = False
    BtnExportSiVisible = False
    Call Rut_Filtrar_Tareas
    MyRibbon.ActivateTab ("TabUserMenu")        '- Para mostrar este Tab -------------
    'MyRibbon.ActivateTab ("TabProgMenu")       '- Para mostrar este Tab -------------
    'MyRibbon.InvalidateControl "Control_ID"    '- Actualiza solo este Control_ID
    Call RefreshRibbon
Debug.Print "--------------------------------------------------- <<< OnLoad_MyRibbon"
End Sub
'------------------------------------------------------------------------------------------
Sub RefreshRibbon()
Debug.Print "---------------------------------------->>> RefreshRibbon ---"
    On Error GoTo RestartExcel
    
'        If MyRibbon Is Nothing Then
'            Dim Lptr As LongPtr
'            Lptr = CLng(Replace(ThisWorkbook.Names("RibbonID").RefersTo, "=", ""))
'            Set MyRibbon = GetRibbon(Lptr)   ' o RetrieveObjRef
'        End If
'
'        MyRibbon.Invalidate
        
''    ' Si se ha perdido la referencia, reintentar con el puntero
''    If MyRibbon Is Nothing Then
''
''        Dim lPtr As LongPtr
''
''        Dim nm As Name
''        Dim s As String
''
''        ' Obtener el Name con el puntero
''        For Each nm In ThisWorkbook.Names
''            If nm.Name Like "*RibbonPtr" Then
''                s = nm.RefersTo
''                Exit For
''            End If
''        Next nm
''
''        If s <> vbNullString Then
''            ' RefersTo viene como "=123456"
''            lPtr = CLng(Replace(s, "=", ""))
''            Set MyRibbon = GetRibbon(lPtr)
''        End If
''    End If
''
''    ' Si ya tenemos Ribbon, invalidar
''    If Not MyRibbon Is Nothing Then
''        If ControlID = vbNullString Then
''            MyRibbon.Invalidate
''        Else
''            MyRibbon.InvalidateControl ControlID
''        End If
''    End If
''
''    Exit Sub
        
    If Not MyRibbon Is Nothing Then
        MyRibbon.Invalidate     ' invalida toda la cinta
    End If
        
        
    On Error GoTo 0
Debug.Print "----------------------------------------<<< RefreshRibbon ---"
    Exit Sub
RestartExcel:
    MsgBx_Msg = "Please restart Excel for Ribbon UI changes to take effect"
    MsgBx_Title = "Proceso: Ribbon UI Refresh Failed"
    MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
Debug.Print "--------------------------------------------------------------<<< ERR - RefreshRibbon >>>"
End Sub
' Función para reconstruir la referencia (GetRibbon/RetrieveObjRef)
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As IRibbonUI
    Dim objRibbon As Object
    
    If lRibbonPointer <> 0 Then
        CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
        Set GetRibbon = objRibbon
        Set objRibbon = Nothing
    End If
End Function
'------------------------------------------------------------------------------------------
Sub GetVisibleTabExcel(control As IRibbonControl, ByRef Visible)
Debug.Print "GetVisibleTabExcel", "ID: " & control.ID, "Tag: " & control.Tag
    Visible = SW_TabExcelVisible
End Sub
'------------------------------------------------------------------------------------------
Sub GetVsbl_CtextMenuCell(control As IRibbonControl, ByRef Visible)
    Visible = Func_CtrlTab_View(control.Tag)
Debug.Print "GetVsbl_CtextMenuCell", control.ID, Visible
End Sub

'------------------------------------------------------------------------------------------
Sub GetVsbl_CtrlTab(control As IRibbonControl, ByRef Visible)
    Visible = Func_CtrlTab_View(control.Tag)
End Sub
'------------------------------------------------------------------------------------------
Function Func_CtrlTab_View(Ctrl As String)
    Dim Cont_Row    As Integer
    With Prog__Menú_Aux.ListObjects(1).DataBodyRange
        For Cont_Row = 1 To .Rows.Count     '- Recorre la Tabla para ver los que deben verse.
            MyTag = .Cells(Cont_Row, Task_Uribbon_Tags)
            If MyTag <> "" Then MyTag = MyTag & "*"         '- única diff ----
            SW_Tag_Visible = .Cells(Cont_Row, Task_Visible)
            If Ctrl Like MyTag Then
                Func_CtrlTab_View = SW_Tag_Visible
                Debug.Print "GetVsbl_UserCtrlTabs", "Func_CtrlTab_View        Tag: " & MyTag, Ctrl, SW_Tag_Visible
                Exit Function
            End If
            If Ctrl Like MyTag & "_Separ" Then
                Func_CtrlTab_View = SW_Tag_Visible
                Debug.Print "GetVsbl_UserCtrlTabs", "Func_CtrlTab_View     _Separ: ", Ctrl, SW_Tag_Visible
                Exit Function
            End If
        Next Cont_Row
    End With
End Function
'------------------------------------------------------------------------------------------
Sub getStip_CtrlTab(control As IRibbonControl, ByRef Supertip)
    Supertip = Func_STip_CtrlTab_Value(control.Tag)
Debug.Print "getStip_CtrlTab", control.Tag, "Largo Supertip: " & Len(Supertip)
End Sub
'------------------------------------------------------------------------------------------
Function Func_STip_CtrlTab_Value(CtrlTag As String)
    Dim Lin_Lst     As Variant
    With Prog__Menú_Aux.ListObjects(1).DataBodyRange
        Lin_Lst = Application.Match(CtrlTag, .Columns(Task_Uribbon_Tags), 0)
        If Not IsError(Lin_Lst) Then
            Func_STip_CtrlTab_Value = .Cells(Lin_Lst, Task_Descripción)
            If Prog__APP.Range("SW_Boss") Then
                Func_STip_CtrlTab_Value = Func_STip_CtrlTab_Value & vbLf & vbLf & _
                                     "BOSS_Rut: " & vbLf & .Cells(Lin_Lst, Task_Nombre_Rut) & vbLf & vbLf
                If (Len(.Cells(Lin_Lst, Task_Rut_Informe)) + Len(Func_STip_CtrlTab_Value)) > 1000 Then
'                    Func_STip_CtrlTab_Value = Left(Func_STip_CtrlTab_Value, 980) & "Informe:" & vbLf & "¡¡ Informe demasiado extenso !!"
                    Func_STip_CtrlTab_Value = Left(Func_STip_CtrlTab_Value, 980) & "Informe:" & vbLf & "¡¡ Informe demasiado extenso !!" & vbLf & _
                            "Informe (parte final):" & vbLf & "..." & vbLf & Right(.Cells(Lin_Lst, Task_Rut_Informe), 950 - Len(Func_STip_CtrlTab_Value))
                Else
                    Func_STip_CtrlTab_Value = Func_STip_CtrlTab_Value & "Informe:" & vbLf & .Cells(Lin_Lst, Task_Rut_Informe)
                End If
                'Debug.Print CtrlTag, "____________________________BOSS_Rut: " & .Cells(Lin_Lst, Task_Nombre_Rut)
            End If
        Else
            Func_STip_CtrlTab_Value = "¡ Control.Tag, NO encontrado !"
        End If
    End With
'Debug.Print "Largo Func_STip_CtrlTab_Value: " & CtrlTag, Len(Func_STip_CtrlTab_Value)
End Function

' ==================================================================================================================================
' ==================================================================================================================================
' =======             ==============================================================================================================
' ======= TabUserMenu ==============================================================================================================
' =======             ==============================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
'__________________________________________________________________________________________
'
'    Change User --------------------------------------------------------------------------
'__________________________________________________________________________________________
Sub GetLbl_User(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Prog__APP.Range("APP_User_Name")
Debug.Print "GetLbl_User", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_CCtxtUser(control As IRibbonControl, ByRef LabelVal)
    LabelVal = "Cambiar el Usuario (" & Prog__APP.Range("APP_User_Name") & ")"
Debug.Print "GetLbl_CCtxtUser", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ChangeUser(control As IRibbonControl)
Debug.Print "OnAct_ChangeUser"
'    Rut_Chg_Usuario
    Call RuT_Ejecutar_Rut("Rut_Chg_Usuario")
'    Form_Usuario.Show
    Call RefreshRibbon
End Sub
'__________________________________________________________________________________________
'
'    Reset App ------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_Reset(control As IRibbonControl)
'    Call RuT_Ejecutar_Rut("Rut_Reset_App")
    Call Rut_Reset_App
Debug.Print "OnAct_Reset"
End Sub
'__________________________________________________________________________________________
'
'   Change RibbonX Visibility -------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_RibbView(control As IRibbonControl)
    Select Case control.Tag
        Case "RibbViewNone"
            Call RuT_Ejecutar_Rut("Rut_Menú_HideAll", False)
        Case "RibbViewMin"
            Call RuT_Ejecutar_Rut("Rut_Menú_ShowAll_Short", False)
        Case "RibbViewMax"
            Call RuT_Ejecutar_Rut("Rut_Menú_ShowAll", False)
    End Select
Debug.Print "OnAct_RibbView"
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
Debug.Print "GetVsbl_RibbView", control.ID, Visible
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   OnOff Help (sólo en alguna hoja) ------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_Help_Group(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Prog__APP.Range("APP_Help_State")
Debug.Print "GetLbl_HelpGroup", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_HelpComments(control As IRibbonControl, ByRef LabelVal)
    If ActiveSheet.Comments.Count = 0 Then
        LabelVal = "Activar Comentarios"
    Else
        LabelVal = "Desactivar Comentarios"
    End If
Debug.Print "GetLbl_HelpComments", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_CCtxtHelpComments(control As IRibbonControl, ByRef LabelVal)
    If ActiveSheet.Comments.Count = 0 Then
        LabelVal = "Activar Comentarios de las celdas"
    Else
        LabelVal = "Desactivar Comentarios de las celdas"
    End If
Debug.Print "GetLbl_CCtxtHelpComments", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_HelpComments(control As IRibbonControl)
        Call Rut_x_Help_ShowHide
Debug.Print "OnAct_HelpComments"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Import/Export BDatos ------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'   <group id="ExImportBDatos_Group"   tag="ExImportBDatos_Group"  label="Ex/Import BDatos"
'__________________________________________________________________________________________
'
'- Export BDatos --------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_ExportBDatos(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Format(Prog__APP.Range("APP_Last_BD_Export"), "dd-mmm-yy hh:mm") & " Export"
Debug.Print "GetLbl_ExportBDatos", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ExportBDatos(control As IRibbonControl)
Debug.Print "================== >>> OnAct_ExportBDatos"
        Prog__APP.Range("APP_Task_Rut") = "Rut_LstObj_Export_Bdatos"
        On Error GoTo ManejoError
                DoEvents ' Permite que Excel procese eventos pendientes
                Form_Running_Rut.Show
        On Error GoTo 0
Debug.Print "================== <<< OnAct_ExportBDatos"
Exit Sub
ManejoError:
    If Err.Number = -2147417848 Then
        Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
        Resume ' Reintenta la línea que falló
    Else
        MsgBox "Error: " & Err.Description
    End If
    MsgBox "<<< Err_Rut Form_Running_Rut >>>"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'- Import BDatos --------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_ImportBDatos(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Format(Prog__APP.Range("APP_Last_BD_Import"), "dd-mmm-yy hh:mm") & " Import"
Debug.Print "GetLbl_ImportBDatos", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ImportBDatos(control As IRibbonControl)
Debug.Print "================== >>> OnAct_ExportBDatos"
        Prog__APP.Range("APP_Task_Rut") = "Rut_LstObj_Import_Bdatos"
        On Error GoTo ManejoError
                DoEvents ' Permite que Excel procese eventos pendientes
                Form_Running_Rut.Show
        On Error GoTo 0
Debug.Print "================== <<< OnAct_ExportBDatos"
Exit Sub

ManejoError:
    If Err.Number = -2147417848 Then
        Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
        Resume ' Reintenta la línea que falló
    Else
        MsgBox "Error: " & Err.Description
    End If
    MsgBox "<<< Err_Rut Form_Running_Rut >>>"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'- Restituir BDatos -----------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_RestoreBDatos(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Format(Prog__APP.Range("APP_Last_BD_Restore"), "dd-mmm-yy hh:mm") & " Restore"
Debug.Print "GetLbl_RestoreBDatos", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_RestoreBDatos(control As IRibbonControl)
Debug.Print "================== >>> OnAct_RestoreBDatos"
        Prog__APP.Range("APP_Task_Rut") = "RuT_LstObj_Restore_BD"
        On Error GoTo ManejoError
                DoEvents ' Permite que Excel procese eventos pendientes
                Form_Running_Rut.Show
        On Error GoTo 0
Debug.Print "================== <<< OnAct_RestoreBDatos"
Exit Sub

ManejoError:
    If Err.Number = -2147417848 Then
        Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
        Resume ' Reintenta la línea que falló
    Else
        MsgBox "Error: " & Err.Description
    End If
    MsgBox "<<< Err_Rut Form_Running_Rut >>>"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Import File LsGes04 -------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="ImportLsGes04_Group"     tag="ImportLsGes04_Group"   label="Import LsGes04">
'__________________________________________________________________________________________
'
'   Importar LsGes04 x AñoCont a BDatos ---------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_Import_G04_ACont(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Format(Prog__APP.Range("APP_Last_Import"), "dd-mmm hh:mm") & " A.Cont " & Prog__APP.Range("APP_AñoCont")
Debug.Print "GetLbl_Import_G04_ACont", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_Import_G04_ACont(control As IRibbonControl)
Debug.Print "OnAct_Import_G04_ACont"
    Call Call_RuT_Update_LSGES04_ACont
'    MyRibbon.InvalidateControl "Import_G04_ACont"    '- Actualiza solo este Control_ID
End Sub
'__________________________________________________________________________________________
'
'   Importar LSGES04 x CursoAcad, para identificar Recibos de Imp.Adm ---------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_Import_G04_CAcadAnt(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Format(Prog__APP.Range("APP_Last_Import_CAcad"), "dd-mmm hh:mm") & " I.Adm." & vbLf & " C.Acad_Ant " & Prog__APP.Range("APP_C_Acad_Ant")
Debug.Print "GetLbl_Import_G04_CAcadAnt", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_Import_G04_CAcadAnt(control As IRibbonControl)
Debug.Print "OnAct_Import_G04_CAcadAnt"
    Call Call_RuT_Update_LSGES04_IAdm_CAcadAnt
'    MyRibbon.InvalidateControl "BtnImport_G04_CAcadAnt"    '- Actualiza solo este Control_ID
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Import File LSace06 x CursoAcad y de Conceptos Administrativos ------------------------
'------------------------------------------------------------------------------------------
'  <group id="ImportLSace06_Group"     tag="ImportLSace06_Group"   label="Import LSace06">
'__________________________________________________________________________________________
'
'   Importar Informe de LSace06 para las cuotas del INSS ----------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_Import_LSace06(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Prog__APP.Range("APP_C_Acad_Ant") & " " & Format(Prog__APP.Range("APP_Last_LSace06_CAcadAnt"), "dd-mmm hh:mm") & vbLf & _
               Prog__APP.Range("APP_C_Acad_Pos") & " " & Format(Prog__APP.Range("APP_Last_LSace06_CAcadPos"), "dd-mmm hh:mm")
Debug.Print "GetLbl_Import_LSace06", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_Import_LSace06(control As IRibbonControl)
Debug.Print "OnAct_Import_LSace06"
    Call Call_RuT_Update_LSace06_CAcad_ImpAdm_INSS
'    MyRibbon.InvalidateControl "BtnImport_LSace06"    '- Actualiza solo este Control_ID
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Import File los 4 de Wb_AE4 -----------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="ImportAE4x4_Group"     tag="ImportAE4x4_Group"   label="Import BD_AE4x4">
'__________________________________________________________________________________________
'
'   Importar Importar los 4 WB: EFP y CFC de AñoCon_Ant/Pos ------------
'------------------------------------------------------------------------------------------
Sub GetLbl_Import_AE4(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Format(Prog__APP.Range("APP_Last_Import_AE4"), "dd-mmm hh:mm") & vbLf & " AE4x4"
Debug.Print "GetLbl_Import_AE4", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_Import_AE4(control As IRibbonControl)
Debug.Print "OnAct_Import_AE4"
    Call Call_RuT_Update_AE4x4
'    MyRibbon.InvalidateControl "BtnImport_AE4"    '- Actualiza solo este Control_ID
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Dynamic Menu-Sheets_Sht to Goto -------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="Dynamic_SheetList_GoTo_Group" getLabel="GetLbl_Dynamic_SheetList_GoTo_Group">
'__________________________________________________________________________________________
'
'      dynamic-Menú Show/Hide-Sheet_Sht
'------------------------------------------------------------------------------------------
Sub GetLbl_Dynamic_SheetList_GoTo_Group(control As IRibbonControl, ByRef LabelVal)
    LabelVal = ActiveSheet.Name
Debug.Print "GetLbl_Dynamic_SheetList_GoToGroup"
End Sub
Sub GetContent_Dynamic_SheetList_GoTo(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Func_xmlGen_Sheet_Sht_List
Debug.Print "GetContent_Dynamic_SheetList_GoTo", returnedVal
End Sub
'------------------------------------------------------------------------------------------
Function Func_xmlGen_Sheet_Sht_List() As String
        Dim WrkSht      As Worksheet
        Dim Sheet_n     As Integer:     Sheet_n = 1
        Dim xml         As String
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Show All' tag='Show All'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet_Sht'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Show Only BD' tag='Show Only BD'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet_Sht'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Hide All' tag='Hide All'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet_Sht'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Hide All Else' tag='Hide All Else'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet_Sht'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<menuSeparator  id='Separ_XX' />"
    For Each WrkSht In Worksheets
        If WrkSht.CodeName Like "Sht__BD*" Then
            xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='" & WrkSht.Name & "' tag='" & WrkSht.Name & _
                        "' imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet_Sht'/>"
            Sheet_n = Sheet_n + 1
        End If
        If WrkSht.CodeName Like "Sht__Inf*" Then
            xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='" & WrkSht.Name & "' tag='" & WrkSht.Name & _
                        "' imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet_Sht'/>"
            Sheet_n = Sheet_n + 1
        End If
    Next
        Func_xmlGen_Sheet_Sht_List = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui'>" & xml & "</menu>"
Debug.Print "Func_xmlGen_Sheet_Sht_List"
End Function
'------------------------------------------------------------------------------------------
Sub OnAct_Goto_Sheet_Sht(control As IRibbonControl)
    Dim NomActivSheet   As String:     NomActivSheet = ActiveSheet.Name
    Dim WrkSht          As Worksheet
    Application.ScreenUpdating = False
    Select Case control.Tag
        Case "Show Only BD"
            Sht__BD.Visible = -1
            Sht__BD.Select
            For Each WrkSht In Worksheets
                If WrkSht.Name <> ActiveSheet.Name Then
                    WrkSht.Visible = xlSheetVeryHidden
                End If
            Next
        Case "Show All"
            For Each WrkSht In Worksheets
                If WrkSht.CodeName Like "Sht__BD*" Or WrkSht.CodeName Like "Sht__Inf*" Then
                    WrkSht.Visible = xlSheetVisible
                End If
            Next
        Case "Hide All"
            For Each WrkSht In Worksheets
                If WrkSht.Name <> NomActivSheet Then WrkSht.Visible = xlSheetVeryHidden
            Next
            Sht__BD.Visible = -1
            Sht__BD.Select
        Case "Hide All Else"
            For Each WrkSht In Worksheets
                If WrkSht.Name <> NomActivSheet Then WrkSht.Visible = xlSheetVeryHidden
            Next
        Case Else
            If control.Tag <> NomActivSheet Then
                Sheets(control.Tag).Visible = xlSheetVisible
                Sheets(control.Tag).Select
'                Sheets(NomActivSheet).Visible = xlVeryHidden
''''                MyRibbon.InvalidateControl "Dynamic_Sheets_Group"    '- Actualiza solo este Control_ID ¡¡¡¡ LO QUITO PORQUE DA ERROR Y NO SÉ PORQUÉ !!!!
            End If
    End Select
    Application.ScreenUpdating = True
Debug.Print "OnAct_Goto_Sheet_Sht", control.Tag
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Calculate JI's  -----------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="ReCalculate_ListObj_Sht_Group"        tag="ReCalculate_ListObj_Sht_Group"
'__________________________________________________________________________________________
'
'  ReCalcula el Contenido de las ListObjects de cada Sheet  -------------------------------
'- ----------------------------------------------------------------------------------------
Sub GetLbl_ReCalculate_ListObj_Sht_Group(control As IRibbonControl, ByRef LabelVal)
    LabelVal = "Calc_" & ActiveSheet.Name
Debug.Print "GetLbl_ReCalculate_ListObj_Sht_Group", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_ReCalculate_ListObj_Sht(control As IRibbonControl, ByRef LabelVal)
    If Fnc_Range_Exist("APP_Last_Calc_" & ActiveSheet.Name) Then
        LabelVal = Format(Prog__APP.Range("APP_Last_Calc_" & ActiveSheet.Name), "dd-mmm-yy hh:mm")
    Else
        LabelVal = "Calc_JIs"
    End If
Debug.Print "GetLbl_ReCalculate_ListObj_Sht", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ReCalculate_ListObj_Sht(control As IRibbonControl)
Debug.Print "OnAct_ReCalculate_ListObj_Sht"
    Application.Run "Rut_Recalcular_Tabla_" & ActiveSheet.Name
    MyRibbon.InvalidateControl "BtnCalculate_JIs"    '- Actualiza solo este Control_ID
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Exportar la WorkSheet  ----------------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="ExportWorkSheet_Group"                       tag="ExportWorkSheet_Group"
'__________________________________________________________________________________________
'
'- Export Active_WorkSheet ----------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_ExportWorkSheet_Group(control As IRibbonControl, ByRef LabelVal)
'    LabelVal = Format(Prog__APP.Range("APP_Last_Exp_JIsPPub"), "dd-mmm-yy hh:mm") & " Exp."
    If Fnc_Range_Exist("APP_Last_Exp_" & ActiveSheet.Name) Then
        LabelVal = "Export " & ActiveSheet.Name
    Else
        LabelVal = "Export WS" & " Exp."
    End If
Debug.Print "GetLbl_ExportWorkSheetGroup", LabelVal
End Sub
Sub GetLbl_ExportWorkSheet(control As IRibbonControl, ByRef LabelVal)
'    LabelVal = Format(Prog__APP.Range("APP_Last_Exp_JIsPPub"), "dd-mmm-yy hh:mm") & " Exp."
    If Fnc_Range_Exist("APP_Last_Exp_" & ActiveSheet.Name) Then
        LabelVal = Format(Prog__APP.Range("APP_Last_Exp_" & ActiveSheet.Name), "dd-mmm-yy hh:mm") & " Exp."
    Else
        LabelVal = "Export WS" & " Exp."
    End If
Debug.Print "GetLbl_ExportWorkSheet", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ExportWorkSheet(control As IRibbonControl)
Debug.Print "================== >>> OnAct_ExportWorkSheet"
    Prog__APP.Range("APP_Task_Rut") = "Rut_Lo_Export_WorkSheet"
    On Error GoTo ManejoError
            DoEvents ' Permite que Excel procese eventos pendientes
            Form_Running_Rut.Show
    On Error GoTo 0
    MyRibbon.InvalidateControl "BtnExportWorkSheet"    '- Actualiza solo este Control_ID
Debug.Print "================== <<< OnAct_ExportWorkSheet"
Exit Sub

ManejoError:
    If Err.Number = -2147417848 Then
        Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
        Resume ' Reintenta la línea que falló
    Else
        MsgBox "Error: " & Err.Description
    End If
    MsgBox "<<< Err_Rut Form_Running_Rut >>>"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Show-Hide ListObj_Cols_Active_Sheet  --------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="ShowHide_Lo_Cols_Group"   tag="ShowHide_Lo_Cols_Group"
'__________________________________________________________________________________________
'
'    "Group Show-Hide_Lo_Cols"
'------------------------------------------------------------------------------------------
Sub GetVsbl_ShowHide_Lo_Cols_group(control As IRibbonControl, ByRef Visible)
    Dim SheetName As String:     SheetName = ActiveSheet.Name
    If Fnc_WrkSheet_Exist(SheetName) Then
        Visible = True
    Else
        Visible = False
    End If
Debug.Print "GetVsbl_ShowHide_Lo_Cols_group,     Visible= " & Visible
End Sub
'-----------------------------
Sub GetLbl_ShowHide_Lo_Cols_Group(control As IRibbonControl, ByRef LabelVal)
    LabelVal = "Table: " & ActiveSheet.Name
    If Not Fnc_Range_Exist("SW_Col_Hide_" & ActiveSheet.CodeName) Then GoTo FinRut
    If Prog__APP.Range("SW_Col_Hide_" & ActiveSheet.CodeName) Then
        LabelVal = LabelVal & ", Cols Hidden"
    Else
        LabelVal = LabelVal & ", Cols Showed"
    End If
FinRut:
Debug.Print "GetLbl_ShowHide_Lo_Cols_Group", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_ShowHide_Lo_Cols(control As IRibbonControl, ByRef LabelVal)
    Application.ScreenUpdating = False
    If Not Fnc_Range_Exist("SW_Col_Hide_" & ActiveSheet.CodeName) Then
        LabelVal = "¿?":            GoTo FinRut
    End If
    If Prog__APP.Range("SW_Col_Hide_" & ActiveSheet.CodeName) Then
        LabelVal = "Show Cols"
    Else
        LabelVal = "Hide Cols"
    End If
    If control.ID = "ShowHide_Lo_Cols_Row1x" Then
        LabelVal = LabelVal & "_Row1"
    Else
        LabelVal = LabelVal & "_Def's"
    End If
    Sheets(ActiveSheet.Name).ListObjects(1).Range(1).Select
    Application.ScreenUpdating = True
FinRut:
Debug.Print "GetLbl_ShowHide_Lo_Cols,  LabelVal=", LabelVal
End Sub
'-----------------------------
Sub OnAct_ShowHide_Lo_Cols(control As IRibbonControl)
    Dim SheetDefCol As String:     SheetDefCol = "_DefCol_" & ActiveSheet.Name
    If Not Fnc_WrkSheet_Exist(SheetDefCol) Then GoTo FinRut
    Application.ScreenUpdating = False
    Call Rut_Lo_Columns_Show_Hide(Sheets(ActiveSheet.Name), Sheets(SheetDefCol), DefC_HiddenCol)
    '- Para Ajustar un grupo de columnas al ancho de la window.
'    Select Case ActiveSheet.Name
'    Case Prog_CTA_Tb.Name
'        Sheets(ActiveSheet.Name).ListObjects(1).HeaderRowRange(1).Resize(, C_Cta_Siglas).Select
'        ActiveWindow.Zoom = True
'    End Select
    Sheets(ActiveSheet.Name).ListObjects(1).Range(1).Select
    MyRibbon.InvalidateControl "ShowHide_Lo_Cols_Group"    '- Actualiza solo este Control_ID
    MyRibbon.InvalidateControl "ShowHide_Lo_Cols"    '- Actualiza solo este Control_ID
    MyRibbon.InvalidateControl "ShowHide_Lo_Cols_Row1x"    '- Actualiza solo este Control_ID
    Application.ScreenUpdating = True
FinRut:
Debug.Print "OnAct_BtnShowHide_Lo_Cols"
End Sub
'-----------------------------
Sub OnAct_ShowHide_Lo_Cols_Row1x(control As IRibbonControl)
    Application.ScreenUpdating = False
    If Fnc_Range_Exist("SW_Col_Hide_" & ActiveSheet.CodeName) Then
        Prog__APP.Range("SW_Col_Hide_" & ActiveSheet.CodeName) = Not Prog__APP.Range("SW_Col_Hide_" & ActiveSheet.CodeName)
        SW_ShowHide_Col = Prog__APP.Range("SW_Col_Hide_" & ActiveSheet.CodeName)
    Else
        SW_ShowHide_Col = Not SW_ShowHide_Col
    End If
    Call Rut_WrkSheet_Col_ShowHide_RowX1(SW_ShowHide_Col)   '- Esta rutina es independiente
    Sheets(ActiveSheet.Name).ListObjects(1).Range(1).Select
    MyRibbon.InvalidateControl "ShowHide_Lo_Cols_Group"    '- Actualiza solo este Control_ID
    MyRibbon.InvalidateControl "ShowHide_Lo_Cols"    '- Actualiza solo este Control_ID
    MyRibbon.InvalidateControl "ShowHide_Lo_Cols_Row1x"    '- Actualiza solo este Control_ID
    Application.ScreenUpdating = True
FinRut:
Debug.Print "OnAct_ShowHide_Lo_Cols_Row1x"
End Sub
'-----------------------------
Sub OnAct_Filter_ShowAlData(control As IRibbonControl)
    Call Rut_Lo_Filtros_Quitar(ActiveSheet.ListObjects(1))
Debug.Print "OnAct_Filter_ShowAlData"
End Sub












'__________________________________________________________________________________________
'
'   Añadir Tabla de Trabajo (BD_M013,BD_PNB1...) a la Tabla principal Bdatos ----------------------------------------------------------------------
'------------------------------------------------------------------------------------------


''__________________________________________________________________________________________
''
''   "GroupMails" -----------------------------------------------------------------------
''------------------------------------------------------------------------------------------
'Sub GetLbl_GroupMails(control As IRibbonControl, ByRef LabelVal)
'        LabelVal = "Mails para el Plan:  " & Range("Liquid_Plan")
'Debug.Print "GetLbl_GroupMails", LabelVal
'End Sub
''------------------------------------------------------------------------------------------
'    '''Sub GetVsbl_Buttons(control As IRibbonControl, ByRef Visible)
'    '''    Dim Cont_Row    As Integer
'    '''    With Prog__Menú_Aux.ListObjects(1).DataBodyRange
'    '''        For Cont_Row = 1 To .Rows.Count     '- Recorre la Tabla para ver los que deben verse.
'    '''            MyTag = .Cells(Cont_Row, 7)
'    '''            SW_Tag_Visible = .Cells(Cont_Row, 8)
'    '''            If control.ID Like MyTag Then Visible = SW_Tag_Visible
'    '''            If control.ID Like MyTag & "_Separ" Then Visible = SW_Tag_Visible
'    '''        Next Cont_Row
'    '''    End With
'    '''End Sub
''- BtnMailRDT_Plan ------------------------------------------------------------------------
'Sub OnAct_MailRDT_Plan(control As IRibbonControl)
'    Index_Mail = 1
'    Form_Mails.Show
'Debug.Print "OnAct_MailRDT_Plan"
'End Sub
''- BtnMails ------------------------------------------------------------------------
'Sub OnAct_Mails(control As IRibbonControl)
'    Form_Mails.Show
'Debug.Print "OnAct_Mails"
'End Sub
''- ClauMails ------------------------------------------------------------------------
'Sub OnAct_ClauMails(control As IRibbonControl)
'    Call Rut_Cambiar_Contraseña_Email
'Debug.Print "OnAct_ClauMails"
'End Sub


'__________________________________________________________________________________________
'
'   Importar Excel para Liquidar Curso < 200h  --------------------------------------------
'------------------------------------------------------------------------------------------
'- EditBoxPlanCurso -----------------------------------------------------------------------------------------
Sub GetText_EditBoxPlanCurso(control As IRibbonControl, ByRef PlanVal)
    PlanVal = ""
Debug.Print "GetText_EditBoxPlanCurso", PlanVal
End Sub
'- EditBoxPlanCurso -----------------------------------------------------------------------------------------
Sub OnChange_EditBoxPlanCurso(control As IRibbonControl, PlanVal As String)
     PlanImport = PlanVal
    'MyRibbon.InvalidateControl "GroupImpExcelLiqCurso"
    'MyRibbon.InvalidateControl "EditBoxPlanCurso"
Debug.Print "OnChange_EditBoxPlanCurso", PlanImport
End Sub




'------------------------------------------------------------------------------------------
'   "CTxtRibbonXVisibility" -------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_RibbonVisible(control As IRibbonControl)
    Select Case control.ID
        Case "CellCtxtRibbonXVisibilityNone"
            Call Rut_Menú_HideAll
        Case "CellCtxtRibbonXVisibilityMin"
            Call Rut_Menú_ShowAll_Short
        Case "CellCtxtRibbonXVisibilityMax"
            Call Rut_Menú_ShowAll
    End Select
'    Call RefreshRibbon
Debug.Print "OnChange_EditBoxPlanCurso"
End Sub





' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
' =======             ==============================================================================================================
' ======= TabProgMenu ==============================================================================================================
' =======             ==============================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================

Sub GetVsbl_ProgButtons(control As IRibbonControl, ByRef Visible)
    If Prog__APP.Range("APP_User_ID") = "Boss" Then Visible = True Else Visible = False
    Debug.Print "GetVsbl_ProgButtons", control.ID, Visible
End Sub
'------------------------------------------------------------------------------------------
Sub GetVsbl_TabProgMenu(control As IRibbonControl, ByRef Visible)
    If Prog__APP.Range("APP_User_ID") = "Boss" Then Visible = True Else Visible = False
Debug.Print "GetVsbl_TabProgMenu", control.ID
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Tools   Tools   Tools   Tools   -------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="Tools_Group"        label="Tools..."    >
'__________________________________________________________________________________________
'
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'- OnAct_MenuAux
'------------------------------------------------------------------------------------------
Sub OnAct_MenuAux(control As IRibbonControl)
Debug.Print "--------------------------------- >>> Sub OnAct_MenuAux <<< -----------------"
    On Error GoTo ManejoError
    DoEvents ' Permite que Excel procese eventos pendientes
    Form_Menu.Show
    Exit Sub
    
ManejoError:
    If Err.Number = -2147417848 Then
        Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
        Resume ' Reintenta la línea que falló
    Else
        MsgBox "Error: " & Err.Description
    End If
Debug.Print "--------------------------------- <<< Sub OnAct_MenuAux >>> -----------------"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'- OnAct_ModeProg
'------------------------------------------------------------------------------------------
Sub OnAct_ModeProg(control As IRibbonControl)
Debug.Print "-------------------------------- >>> Sub OnAct_ModeProg <<< -----------------"
    Call Rut_Activar_Programación
    Call RuT_Load_Task_Data("Rut_Activar_Programación")
Debug.Print "-------------------------------- <<< Sub OnAct_ModeProg >>> -----------------"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'- OnAct_ResetExcelConfig
'------------------------------------------------------------------------------------------
Sub OnAct_ResetExcelConfig(control As IRibbonControl)
'    Call RuT_Ejecutar_Rut("Rut_ConfigExcel_RESTABLECER")
    Call Rut_ConfigExcel_RESTABLECER
Debug.Print "Rut_ConfigExcel_RESTABLECER"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'- OnAct_RibbonRefresh
'------------------------------------------------------------------------------------------
Sub OnAct_RibbonRefresh(control As IRibbonControl)
    Call RuT_Ejecutar_Rut("Rut_RibbonRefresh", False)
Debug.Print "OnAct_RefreshRibbon"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'- OnAct_SearchVinculos
'------------------------------------------------------------------------------------------
Sub OnAct_SearchVinculos(control As IRibbonControl)
    Call RuT_Ejecutar_Rut("Rut_ListarHipervinculos", False)
Debug.Print "OnAct_SearchVinculos"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'- SW_WB_Deactivate_Group
'------------------------------------------------------------------------------------------
Sub GetLbl_SW_WB_Deactivate(control As IRibbonControl, ByRef LabelVal)
    If Prog__APP.Range("SW_WB_Deactivate") Then
        LabelVal = "Deactivate ON"
    Else
        LabelVal = "Deactivate OFF"
    End If
Debug.Print "GetLbl_SW_WB_Deactivate_Group", control.ID, LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_SW_WB_DeactivateOnOff(control As IRibbonControl)
Debug.Print "OnAct_SW_WB_DeactivateOnOff"
    Prog__APP.Range("SW_WB_Deactivate") = Not Prog__APP.Range("SW_WB_Deactivate")
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'- Protect / UnProtect ActiveSheet
'------------------------------------------------------------------------------------------
Sub GetLbl_SheetProtect(control As IRibbonControl, ByRef LabelVal)
    Dim Protect_Status   As Boolean:  Protect_Status = ActiveSheet.ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
    If Protect_Status Then
        LabelVal = "UnProtect Sheet"
    Else
        LabelVal = "protect Sheet"
    End If
Debug.Print "GetLbl_SheetProtect,  LabelVal=", LabelVal
End Sub
'------------------------
Sub GetVsbl_SheetProtect(control As IRibbonControl, ByRef Visible)
    Dim Protect_Status   As Boolean:  Protect_Status = ActiveSheet.ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
    If control.ID = "BtnSheetProtect" Then
        If Protect_Status Then Visible = False Else Visible = True
    Else
        If Protect_Status Then Visible = True Else Visible = False
    End If
Debug.Print "GetVsbl_SheetProtect,  control.ID=", control.ID, Visible
End Sub
'----------------------
Sub OnAct_SheetProtect(control As IRibbonControl)
    Dim Protect_Status   As Boolean:  Protect_Status = ActiveSheet.ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
    If Protect_Status Then
        ActiveSheet.Unprotect
    Else
        ActiveSheet.Protect
    End If
    MyRibbon.InvalidateControl "BtnSheetProtect"        '- Actualiza solo este Control_ID
    MyRibbon.InvalidateControl "BtnSheetUnProtect"      '- Actualiza solo este Control_ID
Debug.Print "OnAct_SheetProtect,  Protect_Status=", Protect_Status
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'- SW_Prueba_Group ------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetVsbl_SW_PruebaONOFF(control As IRibbonControl, ByRef Visible)
    If control.ID = "BtnSW_PruebaON" Then
        Visible = Not Prog__APP.Range("SW_Test")
    Else
        Visible = Prog__APP.Range("SW_Test")
    End If
Debug.Print "GetVsbl_BtnSW_PruebaONOFF", control.ID, Visible
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_SW_PruebaONOFF(control As IRibbonControl)
Debug.Print "OnAct_SW_PruebaONOFF"
    Prog__APP.Range("SW_Test") = Not Prog__APP.Range("SW_Test")
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Show/Hide Excels Tag's Buttons   ------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="ExcelTabsVisible_Group"
'__________________________________________________________________________________________
'
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'- GetLbl_ExcelTabsVisible_Group
'------------------------------------------------------------------------------------------
Sub GetLbl_ExcelTabsVisible_Group(control As IRibbonControl, ByRef LabelVal)
    If SW_TabExcelVisible Then
        LabelVal = "Tabs Show"
    Else
        LabelVal = "Tabs Hide"
    End If
Debug.Print "GetLbl_ExcelTabsVisible_Group", LabelVal
End Sub
Sub GetLbl_CCtxtBtnExcelTabsVisible(control As IRibbonControl, ByRef LabelVal)
    If SW_TabExcelVisible Then
        LabelVal = "Tabs Shown, Press to Hide them"
    Else
        LabelVal = "Tabs Hidden, Press to Show them"
    End If
Debug.Print "GetLbl_CCtxtBtnExcelTabsVisible", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_ExcelTabsVisible(control As IRibbonControl)
    SW_TabExcelVisible = Not SW_TabExcelVisible
    Call RefreshRibbon
Debug.Print "OnAct_ExcelTabsVisible"
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Dynamic Menu-Sheets to Goto   ---------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="Dynamic_Sheets_Group" label="Dynamic">
'__________________________________________________________________________________________
'
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'      dynamic-Menú Show/Hide-Sheet
'------------------------------------------------------------------------------------------
Sub GetContent_Sheet_List(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Func_xmlGen_Sheet_List
Debug.Print "GetContent_Sheet_List", returnedVal
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
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Show Sht__BD*' tag='Sht__BD_#'" & _
                " imageMso='GoToNextAppointment' onAction='OnAct_Goto_Sheet'/>"
                Sheet_n = Sheet_n + 1
    xml = xml & "<button id='sheet_" & Format(Sheet_n, "00") & "' label='Show Sht__Inf*' tag='Sht__Inf_#'" & _
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
    Dim BDatosSheet As Boolean
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
            Sht__BD.Visible = -1
            Sht__BD.Select
        Case "Prog_#"
            For Each WrkSht In Worksheets                      '- 0 = xlSheetHidden, 2 = xlSheetVeryHidden, -1 = xlSheetVisible
                If WrkSht.CodeName Like "Prog_*" Then If WrkSht.Visible < 0 Then WrkSht.Visible = 2 Else WrkSht.Visible = -1
            Next
        Case "DefCol_#"
            For Each WrkSht In Worksheets                      '- 0 = xlSheetHidden, 2 = xlSheetVeryHidden, -1 = xlSheetVisible
                If WrkSht.CodeName Like "*DefCol*" Then If WrkSht.Visible < 0 Then WrkSht.Visible = 2 Else WrkSht.Visible = -1
            Next
        Case "Sht__BD_#"
            Sht__BD.Visible = -1
            Sht__BD.Select
            For Each WrkSht In Worksheets                      '- 0 = xlSheetHidden, 2 = xlSheetVeryHidden, -1 = xlSheetVisible
                If WrkSht.Name <> ActivSheet Then
                    If WrkSht.CodeName Like "Sht__BD_*" Then If WrkSht.Visible < 0 Then WrkSht.Visible = 2 Else WrkSht.Visible = -1
                End If
            Next
        Case "Sht__Inf_#"
            Sht__BD.Visible = -1
            Sht__BD.Select
            For Each WrkSht In Worksheets                      '- 0 = xlSheetHidden, 2 = xlSheetVeryHidden, -1 = xlSheetVisible
                If WrkSht.Name <> ActivSheet Then
                    If WrkSht.CodeName Like "Sht__Inf_*" Then If WrkSht.Visible < 0 Then WrkSht.Visible = 2 Else WrkSht.Visible = -1
                End If
            Next
        Case Else
            If Sheets(control.Tag).Visible = xlSheetVisible Then
                Sheets(Prog__APP.Range("App_Ini_Sheet").Value).Visible = xlSheetVisible
                Sheets(Prog__APP.Range("App_Ini_Sheet").Value).Select
                If UCase(control.Tag) <> UCase(Prog__APP.Range("App_Ini_Sheet").Value) Then
                    Sheets(control.Tag).Visible = xlVeryHidden
                End If
            Else
                Sheets(control.Tag).Visible = xlSheetVisible
                Sheets(control.Tag).Select
            End If
    End Select
    Application.ScreenUpdating = True
Debug.Print "OnAct_Goto_Sheet", control.Tag
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Run Rut_Prueba   ----------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="ExcelTabsVisible_Group"
'__________________________________________________________________________________________
'
'- RunRutPrueba_Group ------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub OnAct_RunRutPrueba(control As IRibbonControl)
Debug.Print "OnAct_RunRutPrueba"
Debug.Print "================== >>> OnAct_RunRutPrueba"
        Prog__APP.Range("APP_Task_Rut") = Func_Rut_CtrlTab_Value(control.Tag)
        If Prog__APP.Range("APP_Task_Rut") = "NotFound" Then Exit Sub
        On Error GoTo ManejoError
                DoEvents ' Permite que Excel procese eventos pendientes
                Form_Running_Rut.Show
        On Error GoTo 0
Debug.Print "================== <<< OnAct_ExportBDatos"
Exit Sub

ManejoError:
    If Err.Number = -2147417848 Then
        Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
        Resume ' Reintenta la línea que falló
    Else
        MsgBox "Error: " & Err.Description
    End If
    MsgBox "<<< Err_Rut OnAct_RunRutPrueba >>>"
End Sub
'------------------------------------------------------------------------------------------
Function Func_Rut_CtrlTab_Value(CtrlTag As String)  '- Busca el nombre de la Rutina en Menú_Aux
    Dim Lin_Lst     As Variant
    With Prog__Menú_Aux.ListObjects(1).DataBodyRange
        Lin_Lst = Application.Match(CtrlTag, .Columns(Task_Uribbon_Tags), 0)
        If Not IsError(Lin_Lst) Then
            Func_Rut_CtrlTab_Value = .Cells(Lin_Lst, Task_Nombre_Rut)
        Else
            Func_Rut_CtrlTab_Value = "NotFound"
        End If
    End With
'Debug.Print "Largo Func_Rut_CtrlTab_Value: " & CtrlTag, Len(Func_Rut_CtrlTab_Value)
End Function
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Show or Hide Right Click Menú   -------------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="RightClickMenúVisible_Group"
'__________________________________________________________________________________________
'
'- RightClickMenúVisible_Group ------------------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_RightClickMenúVisible_Group(control As IRibbonControl, ByRef LabelVal)
    If Prog__APP.Range("SW_RightClickMenú_Visible") Then LabelVal = "Showed" Else LabelVal = "Hidden"
Debug.Print "GetLbl_RightClickMenúVisible_Group", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_CCtxtRightClickMenúVisible(control As IRibbonControl, ByRef LabelVal)
    If Prog__APP.Range("SW_RightClickMenú_Visible") Then LabelVal = "Context Menú on Right-Click is Showed, Press to Hide" _
                                                    Else LabelVal = "Context Menú on Right-Click is hidden, Press to Show"
Debug.Print "GetLbl_CCtxtRightClickMenúVisible", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_RightClickMenúVisible(control As IRibbonControl)
    Prog__APP.Range("SW_RightClickMenú_Visible") = Not Prog__APP.Range("SW_RightClickMenú_Visible")
    MyRibbon.InvalidateControl "RightClickMenúVisible_Group"    '- Actualiza solo este Control_ID
Debug.Print "OnAct_RightClickMenúVisible,  SW_RightClickMenú_Visible = ", Prog__APP.Range("SW_RightClickMenú_Visible")
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Restrict Right Click Menú Options   ---------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="RightClickMenúRestrictOptions_Group"
'__________________________________________________________________________________________
'
'- RightClickMenúRestrictOptions_Group ----------------------------------------------------
'------------------------------------------------------------------------------------------
Sub GetLbl_RightClickMenúRestrictOptions_Group(control As IRibbonControl, ByRef LabelVal)
    If Prog__APP.Range("SW_RightClickMenú_Restricted") Then LabelVal = "Restricted" Else LabelVal = "Not Restricted"
Debug.Print "GetLbl_RightClickMenúRestrictOptions_Group", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub GetLbl_CCtxtRightClickMenúRestrictOptions(control As IRibbonControl, ByRef LabelVal)
    If Prog__APP.Range("SW_RightClickMenú_Restricted") Then LabelVal = "Right-Click is Restricted, Press to Release" _
                                                       Else: LabelVal = "Right-Click is Non Restricted, Press to Restrict"
Debug.Print "GetLbl_CCtxtRightClickMenúRestrictOptions", LabelVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnAct_RightClickMenúRestrictOptions(control As IRibbonControl)
    If Prog__APP.Range("SW_RightClickMenú_Restricted") Then
        Call Rut_Context_Buttons_Restore    '- Restaura las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
    Else
        Call Rut_Context_Buttons_Hide   '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
    End If
    MyRibbon.InvalidateControl "RightClickMenúRestrictOptions_Group"    '- Actualiza solo este Control_ID
Debug.Print "OnAct_RightClickMenúRestrictOptions,  SW_RightClickMenú_Restricted = ", Prog__APP.Range("SW_RightClickMenú_Restricted")
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   Save Copy Timed in the USB  DATA/VBA   ------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="SaveTimer_USB_Group"
'__________________________________________________________________________________________
'
Sub GetLbl_SaveTimer_USB_Group(control As IRibbonControl, ByRef LabelVal)
Debug.Print "GetLbl_SaveTimer_USB_Group"
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
'- SaveData
'------------------------------------------------------------------------------------------
Sub OnAct_SaveData_Timer_USB(control As IRibbonControl)
Debug.Print "OnAct_BtnSaveData_Timer_USB"
    Call Rut_WrkBook_CopSegTimed_USB("_Data")
    SaveUSB = Format(Now(), "dd-mmm-yy hh:mm")
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
'- SaveVBA
'------------------------------------------------------------------------------------------
Sub OnAct_SaveVBA_Timer_USB(control As IRibbonControl)
Debug.Print "OnAct_BtnSaveVBA_Timer_USB"
    Call Rut_WrkBook_CopSegTimed_USB("_VBA")
    SaveUSB = Format(Now(), "dd-mmm-yy hh:mm")
    Call RefreshRibbon
End Sub
'------------------------------------------------------------------------------------------
'__________________________________________________________________________________________
'
'   "GroupParam" AñoCont y Curso Acad  ------------------------------------------------
'------------------------------------------------------------------------------------------
'  <group id="Parámetros_Group"     label="Parámetros APP">
'__________________________________________________________________________________________
'
'    Año-Cont
'------------------------------------------------------------------------------------------
Sub GetText_EditBoxACont(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Prog__APP.Range("APP_AñoCont")
Debug.Print "GetText_EditBoxACont", returnedVal
End Sub
'------------------------------------------------------------------------------------------
Sub OnChange_EditBoxACont(control As IRibbonControl, ACont As String)
    If Val(ACont) >= 2022 Then
        '--- Activo los valores de los Cursos Académicos ------------------------------------>>>>
        Prog__APP.Range("APP_C_Acad_Ant") = ACont - 1 & "-" & Right(ACont, 2)
        Prog__APP.Range("APP_C_Acad_Pos") = ACont & "-" & Right(ACont, 2) + 1
        Prog__APP.Range("APP_CursAcad") = Prog__APP.Range("APP_C_Acad_Ant")
        SW_C_Acad_Ant = True
        SW_C_Acad_Pos = False
        MyTag = "checkBoxAcad1"
        Prog__APP.Range("APP_AñoCont") = ACont
        '------------------------------------------------------------------------------------<<<<
    Else
        MsgBx_Msg = "¡¡¡ Año incorrecto, debe ser mayor que 2021 !!!"
        MsgBx_Title = "Proceso: Procedimiento de Liquidación"
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show  '- ([Font-Size]=16, [Red-Border]=False, [Buttons]="Ok", [Default-Button]=1, [Image]="Msg")
    End If
    Call RefreshRibbon
Debug.Print "OnChange_EditBoxACont,   value= ", ACont
End Sub
'     Select Curso Acad.
Sub OnAct_checkBoxAcad(control As IRibbonControl, pressed As Boolean)
    If MyTag <> control.Tag Then MyTag = control.Tag
    If MyTag = "checkBoxAcadAnt" Then
        SW_C_Acad_Ant = True
        SW_C_Acad_Pos = False
        Prog__APP.Range("APP_CursAcad") = Prog__APP.Range("APP_C_Acad_Ant")
    Else
        SW_C_Acad_Ant = False
        SW_C_Acad_Pos = True
        Prog__APP.Range("APP_CursAcad") = Prog__APP.Range("APP_C_Acad_Pos")
    End If
'    Wk_TitP_Liquid.Range("Liquid_Tipo_Plan") = Prog__APP.Range("APP_TitP_o_Curs") & " - " & Prog__APP.Range("APP_CursAcad")
    Call RefreshRibbon
Debug.Print "sub OnAct_checkBoxAcad", control.Tag, pressed
End Sub
'    SHOW -  Select Curso Acad.
Sub GetPressed_checkBoxAcad(control As IRibbonControl, ByRef ActiveBox)
    Select Case control.Tag
        Case "checkBoxAcadAnt"
            ActiveBox = SW_C_Acad_Ant
        Case "checkBoxAcadPos"
            ActiveBox = SW_C_Acad_Pos
    End Select
Debug.Print "sub GetPressed_checkBoxAcad", control.Tag, ActiveBox
End Sub
'---------------- BoxAcadAnt ------------------------------------------------------------------
Sub GetLbl_checkBoxAcadAnt(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Prog__APP.Range("APP_C_Acad_Ant")
Debug.Print "sub GetLbl_checkBoxAcadAnt", LabelVal
End Sub
'---------------- BoxAcadPos -------------------------------------------------------------------
Sub GetLbl_checkBoxAcadPos(control As IRibbonControl, ByRef LabelVal)
    LabelVal = Prog__APP.Range("APP_C_Acad_Pos")
Debug.Print "sub GetLbl_checkBoxAcadPos", LabelVal
End Sub
''''__________________________________________________________________________________________
''''    "GroupTipoEP" Identificar qué Enseñanzas Propias Gestiona esta APP -------------------
''''------------------------------------------------------------------------------------------
''''------------- Select Tipo de Enseñanza Propia --------------------------------------------
'''Sub OnAct_checkBoxTipoEP(control As IRibbonControl, pressed As Boolean)
'''    If MyTag <> control.Tag Then MyTag = control.Tag
'''    If MyTag = "checkBoxEPTP" Then
'''        Prog__APP.Range("APP_TitP_o_Curs") = "Tít. Propios"
'''        Wk_TitP_Liquid.Name = "Liquid_Tit_Propios"
'''    Else
'''        Prog__APP.Range("APP_TitP_o_Curs") = "Cursos<200h"
'''        Wk_TitP_Liquid.Name = "Liquid_Cursos200h"
'''    End If
''''    Wk_TitP_Liquid.Range("Liquid_Tipo_Plan") = Prog__APP.Range("APP_TitP_o_Curs") & " - " & Prog__APP.Range("APP_CursAcad")
'''    Call RefreshRibbon
'''Debug.Print "sub OnAct_checkBoxTipoEP", control.Tag, pressed, control.Tag, Prog__APP.Range("APP_TitP_o_Curs")
'''End Sub
''''------------- SHOW - Select Tipo de Enseñanza Propia --------------------------------------------
'''Sub GetPressed_checkBoxTipoEP(control As IRibbonControl, ByRef ActiveBox)
'''    Select Case control.Tag
'''        Case "checkBoxEPTP"
'''            ActiveBox = False
'''        Case "checkBoxEPCR"
'''            ActiveBox = False
'''    End Select
'''Debug.Print "sub GetPressed_checkBoxTipoEP", control.Tag, ActiveBox
'''End Sub





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
'        With Prog__Menú_Aux.ListObjects(1).DataBodyRange
'            For Cont_Row = 1 To .Rows.Count
'                MyTag = .Cells(Cont_Row, 7)
'                SW_Tag_Visible = .Cells(Cont_Row, 8)
'                If Tag Like MyTag Then Func_TagVisible = SW_Tag_Visible
'            Next Cont_Row
'        End With
'    End Function
''------------------------------------------------------------------------------------------

''------------------------------------------------------------------------------------------
'Sub GetVisible(control As IRibbonControl, ByRef Visible)
'    If control.Tag Like MyTag Then Visible = SW_Tag_Visible
''''    If MyTag = "ShowALL" Then
''''        visible = True
''''
''''    Else
''''        If control.Tag Like MyTag Then
''''            visible = True
''''        Else
''''            visible = False
''''        End If
''''    End If
'End Sub

'>>>>>##########################################################################################>>>>>
'########## dynamicMenu id="dynamicShortRuts" ##################################################>>>>>
Sub getContent_dynamicShortRuts(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Func_xmlGen
'    Debug.Print "Rut: GetMenuContent1: " & Func_xmlGen
Debug.Print "getContent_dynamicShortRuts", returnedVal
End Sub
'--------------------------------------------------------------------------------------------------
    Function Func_xmlGen() As String
        Dim Cont_Row        As Integer
        Dim Num_Tarea       As Integer:     Num_Tarea = 1
        Dim xml As String
        '- Recorremos toda la tabla de Tareas y seleccionamos las que empiezan por "9_ "
        With Prog__Menú_Aux.ListObjects(1).DataBodyRange
            For Cont_Row = 1 To .Rows.Count
                If (InStr((.Cells(Cont_Row, 2)), (Prog__APP.Range("APP_User_ID"))) > 0 And Len(Prog__APP.Range("APP_User_ID")) > 3 Or .Cells(Cont_Row, 2) = "") And Left(.Cells(Cont_Row, 1), 3) = "9_ " Then
                    If Num_Tarea > 1 Then xml = xml & "<menuSeparator  id='Separ" & Num_Tarea & "' />"
                    xml = xml & "<button id='button" & Num_Tarea & "' tag='Task-Row_" & Format(Cont_Row, "00") & _
                                "' label='" & Mid(.Cells(Cont_Row, 1), 4) & "' imageMso='GoToNextAppointment'" & _
                                " onAction='OnAction_Dynamic_Task'/>"
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
Debug.Print "------------------- >>> OnAction_Dynamic_Task"
    Dim Pos_Delimitador     As Integer
    Dim Rutinas_Name        As String
    Dim Rut_Name        As String
    Prog__APP.Range(APP_Task_Index) = Val(Right(control.Tag, 2))
    Rutinas_Name = Prog__Menú_Aux.ListObjects(1).DataBodyRange.Cells(Prog__APP.Range(APP_Task_Index), 3)
    Do       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Pos_Delimitador = InStr(Rutinas_Name, " + ")
        If Pos_Delimitador < 1 Then ' ------------------ Última o Única Rutina ---------------------
            Pos_Delimitador = Len(Rutinas_Name) + 1
        End If
        Rut_Name = Left(Rutinas_Name, Pos_Delimitador - 1)
        Rutinas_Name = Mid(Rutinas_Name, Pos_Delimitador + 3)
        Application.Run Rut_Name
        Prog__Menú_Aux.ListObjects(1).DataBodyRange.Cells(Prog__APP.Range(APP_Task_Index), 5) = Prog__APP.Range("APP_Task_Inf")
    Loop While Len(Rutinas_Name) > 0
    Application.ScreenUpdating = True
    DoEvents
Debug.Print "------------------- <<< OnAction_Dynamic_Task"
End Sub
'<<<<<##########################################################################################<<<<<
'########## Dynamic-Menú #######################################################################<<<<<
'<<<<<##########################################################################################<<<<<


