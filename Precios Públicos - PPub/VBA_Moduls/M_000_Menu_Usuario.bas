Attribute VB_Name = "M_000_Menu_Usuario"
' Last Rev. 2026-10-04 13:09
Option Explicit

' ==================================================================================================================================
'- Boton Cambiar Usuario del Ribbon (OnAct_ChangeUser). Antes pasaba por RuT_Ejecutar_Rut("Rut_Chg_Usuario"),
'  que buscaba la rutina en Tb_Tareas y mostraba un Form_MsgBox con APP_Task_Inf (2026-10-04).
Sub Rut_Usuario_Chg()
        Form_Usuario.Show
        Call Rut_RibbonUI_Guardar_Informe("ChangeUser", "Usuario activo: " & Prog__APP.Range("APP_User_Name") & "  -  " & Now)
        Call RefreshRibbon
        Application.ScreenUpdating = True
        DoEvents
End Sub

' ==================================================================================================================================
Sub Rut_Filtrar_Tareas()
    Dim Cont_Row            As Integer
    Dim Users               As String
    Dim Users_Allowed       As Boolean
    Dim SheetBtn            As String
    Dim SheetBtn_OK         As Boolean
    Dim Usuario_ID          As String:      Usuario_ID = UCase(Prog__APP.Range("APP_User_ID"))
    
    Call Rut_Lo_Sort(Prog__Menu_Aux.ListObjects(1), 1, xlAscending, True)
    
    With Prog__Menu_Aux.ListObjects(1).DataBodyRange
        For Cont_Row = 1 To .Rows.Count
            
            Users = UCase(.Cells(Cont_Row, Task_Usuario))
            Users_Allowed = (Users = "" Or InStr(Users, Usuario_ID) > 0)
            
            SheetBtn = UCase(.Cells(Cont_Row, Task_SheetsButton))
            SheetBtn_OK = (SheetBtn = "" Or InStr(SheetBtn, UCase(ActiveSheet.Name)) > 0)
            
            If Users_Allowed And SheetBtn_OK Then
                .Cells(Cont_Row, Task_Visible) = True       '- Button Visible
            Else
                .Cells(Cont_Row, Task_Visible) = False      '- Button NOT Visible
            End If
 
        Next Cont_Row
        Debug.Print "Rut_Filtrar_Tareas, .Rows.Count: " & .Rows.Count
    End With
End Sub
' ==================================================================================================================================

