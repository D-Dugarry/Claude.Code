Attribute VB_Name = "M_000_Menú_Usuario"
Option Explicit

' ==================================================================================================================================
Sub Rut_Usuario_Chg()
        Form_Usuario.Show
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
    
    Call Rut_Lo_Sort(Prog__Menú_Aux.ListObjects(1), 1, xlAscending, True)
    
    With Prog__Menú_Aux.ListObjects(1).DataBodyRange
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

