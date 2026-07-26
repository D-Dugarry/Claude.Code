Attribute VB_Name = "Rut_UserForms"
Option Explicit

' ==================================================================================================================================
Sub Rut_UserForm_Close_All()
    Dim frm As Object
        ' Loop through the VBComponents to find open forms
        For Each frm In VBA.UserForms
            Unload frm ' Close the form
        Next frm
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

Function Func_Form_IsAnyFormOpen() As Boolean
    Dim Formulario As Object
        Func_Form_IsAnyFormOpen = False
        ' Recorre la colección de UserForms abiertos
        For Each Formulario In VBA.UserForms
            Func_Form_IsAnyFormOpen = True
            Debug.Print "Formulario abierto: " & Formulario.Name
        Next Formulario
End Function

