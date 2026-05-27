Attribute VB_Name = "M_900_Rutinas_Menú_Aux"
Option Explicit

'==================================================================================================================================
' ==================================================================================================================================
Sub RuT_Ejecutar_Rut(TaskRut As String)
    Dim Rutinas_Name    As String
    Dim TaskIndice      As Variant
    TaskIndice = Application.Match(TaskRut, Prog__MnAux.ListObjects(1).DataBodyRange.Columns(Task_Nombre_Rut), 0)
    If IsError(TaskIndice) Then     ' ¡¡¡ NO Existe la Rutina !!! ------------------------
        MsgBx_Msg = "¡ No Existe la Tarea o su nombre ha cambiado !"
    Else                            ' ¡¡¡ Existe la Rutina !!! ------------------------
        Prog__APP.Range("APP_Task_Rut") = TaskRut
        Prog__APP.Range("APP_Task_Index") = TaskIndice
        Application.Run Prog__APP.Range("APP_Task_Rut").Value
        MsgBx_Msg = Prog__APP.Range("APP_Task_Inf")
        Prog__MnAux.ListObjects(1).DataBodyRange.Cells(TaskIndice, Task_Rut_Informe) = Prog__APP.Range("APP_Task_Inf")
    End If
    MsgBx_Title = "Proceso de Ejecutar la Tarea:  " & TaskRut
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK"): Form_MsgBox.Show  '- ([Font-Size]=16, [Red-Border]=False, [Buttons]="Ok", [Default-Button]=1, [Image]="Msg")
End Sub
'==================================================================================================================================
' ==================================================================================================================================
Sub Rut_Chg_Usuario()
    Form_Usuario.Show
End Sub
' ==================================================================================================================================
Sub Rut_Activar_Programación()
    Call RuT_Al_Abrir_WorkBook
    Call Rut_Sheets_ShowAll
    Call Rut_ConfigExcel_Restablecer
    Prog__APP.Range("APP_Task_Inf") = "Estado de Programación Activado  -  " & Now
End Sub
' ==================================================================================================================================
Sub Rut_Reset_App()
    Call RuT_Al_Abrir_WorkBook
    Prog__APP.Range("APP_Task_Inf") = "App Reset " & Now
End Sub
' ==================================================================================================================================
Sub Rut_Btn_Menú_Aux()
    Form_Menu.Show
End Sub
' ==================================================================================================================================
Sub Rut_Columns_Show_All()
    Call Rut_Protect_Status_Save(ThisWorkbook.ActiveSheet.Name)
    ActiveSheet.Unprotect
    Columns.EntireColumn.Hidden = False
    Call Rut_Protect_Status_Restore(ThisWorkbook.ActiveSheet.Name)
    Prog__APP.Range("APP_Task_Inf") = "All columns are visible" & vbCrLf & Now
End Sub
' ==================================================================================================================================
Sub Rut_OnOff_SW_Test()
    If Prog__APP.Range("SW_Test") Then
        Prog__APP.Range("SW_Test") = False
        Form_Menu.Lb_SW_Test.Visible = False
        Prog__APP.Range("APP_Task_Inf") = "SW_Test - DesActivado  -  " & Now
    Else
        Prog__APP.Range("SW_Test") = True
        Form_Menu.Lb_SW_Test.Visible = True
        Prog__APP.Range("APP_Task_Inf") = "SW_Test - Activado  -  " & Now
    End If
End Sub
' ==================================================================================================================================
Sub Rut_Sheets_ShowAll()
Dim WrkSht          As Worksheet
    For Each WrkSht In Worksheets
                WrkSht.Visible = xlSheetVisible
    Next
    ActiveWindow.DisplayWorkbookTabs = True                     'Oculta las fichas de las hohas
    Prog__APP.Range("APP_Task_Inf") = "All sheets Visible " & Now
End Sub
' ==================================================================================================================================
Sub Rut_Sheets_Hide()
Dim WrkSht          As Worksheet
    H_INICI.Visible = xlSheetVisible
    H_INICI.Select
    For Each WrkSht In Worksheets
        If Left(WrkSht.CodeName, 5) = "Prog_" Then WrkSht.Visible = xlSheetVeryHidden   'xlSheetVisible   '
    Next
    ActiveWindow.DisplayWorkbookTabs = False                     'Oculta las fichas de las hohas
    Prog__APP.Range("APP_Task_Inf") = "All sheets Hide " & Now
End Sub
' ==================================================================================================================================
Sub Rut_OnOff_SW_WB_Deactivate()
    If Prog__APP.Range("SW_WB_Deactivate") Then
        Prog__APP.Range("SW_WB_Deactivate") = False
        Form_Menu.Lb_SW_Test.Visible = False
        Prog__APP.Range("APP_Task_Inf") = "SW_Test - DesActivado  -  " & Now
    Else
        Prog__APP.Range("SW_WB_Deactivate") = True
        Form_Menu.Lb_SW_Test.Visible = True
        Prog__APP.Range("APP_Task_Inf") = "SW_Test - Activado  -  " & Now
    End If
End Sub
' ==================================================================================================================================
Sub Rut_ProtectUnProtect_ActivSheet()
    If ActiveSheet.ProtectContents Then
        ActiveSheet.Unprotect
        Prog__APP.Range("APP_Task_Inf") = "ActiveSheet.UnProtect " & Now
    Else
        ActiveSheet.Protect
        Prog__APP.Range("APP_Task_Inf") = "ActiveSheet.Protect " & Now
    End If
End Sub
' ==================================================================================================================================
Sub Rut_Events_Status_Change()        ' Para permitir las rutinas que se activan cuando ocurre un evento
    If Application.EnableEvents Then
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
        Prog__APP.Range("SW_Events") = False
        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is OFF " & Now
    Else
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
        Prog__APP.Range("SW_Events") = True
        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is On " & Now
    End If
End Sub
' ==================================================================================================================================
    'Call Rut_Events_Status_Choose(Optional Choose As String = "CHANGE")    '- EnableEvents: ["CHANGE"], "ON", "OFF"
Sub Rut_Events_Status_Choose(Optional Choose As String = "CHANGE")      ' Para permitir las rutinas que se activan cuando ocurre un evento
    Select Case UCase(Choose)
        Case "CHOOSE"
                        If Application.EnableEvents Then
                            Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
                            Prog__APP.Range("SW_Events") = False
                            Prog__APP.Range("APP_Task_Inf") = "Events Status Change is OFF " & Now
                        Else
                            Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
                            Prog__APP.Range("SW_Events") = True
                            Prog__APP.Range("APP_Task_Inf") = "Events Status Change is On " & Now
                        End If
        Case "ON"
                        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
                        Prog__APP.Range("SW_Events") = True
                        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is On " & Now
        Case "OFF"
                        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
                        Prog__APP.Range("SW_Events") = False
                        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is OFF " & Now
    End Select
End Sub
