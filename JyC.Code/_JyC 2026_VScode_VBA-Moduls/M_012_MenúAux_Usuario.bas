Attribute VB_Name = "M_012_MenúAux_Usuario"
Option Explicit

' ==================================================================================================================================
Sub Rut_Menú_Aux_CALL()
    Form_Menu.Show
'    Application.ScreenUpdating = True
'    DoEvents
End Sub
'==================================================================================================================================
Sub Rut_Menú_Aux_Close()
    Unload Form_Menu
End Sub
' ==================================================================================================================================
Sub Rut_Usuario_Chg()
        Form_Usuario.Show
        Application.ScreenUpdating = True
        DoEvents
End Sub

            Sub Rut_Filtrar_Tareas_ByHand()
                Rut_Filtrar_Tareas ("Boss")
            End Sub
' ==================================================================================================================================
Sub Rut_Filtrar_Tareas(Usuario_ID As String)
    Dim Cont_Row    As Integer, Users_Allowed   As String, SheetBtn     As String
    
    With Prog__MnAux.ListObjects(1).DataBodyRange
        For Cont_Row = 1 To .Rows.Count
            Users_Allowed = .Cells(Cont_Row, Task_Usuario)
            SheetBtn = .Cells(Cont_Row, Task_SheetsButton)
            If (Users_Allowed = "" Or (Usuario_ID <> "" And InStr(Users_Allowed, Usuario_ID) > 0)) _
                And (SheetBtn = "" Or (SheetBtn <> "" And InStr(SheetBtn, ActiveSheet.Name) > 0)) Then
                .Cells(Cont_Row, Task_Visible) = True       '- Button Visible
            Else
                .Cells(Cont_Row, Task_Visible) = False      '- Button NOT Visible
            End If
        Next Cont_Row
        Debug.Print "Rut_Filtrar_Tareas, .Rows.Count: " & .Rows.Count
    End With
End Sub
' ==================================================================================================================================








