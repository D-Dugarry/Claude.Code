Attribute VB_Name = "M_900_Func_CapsLock_State"
Option Explicit

'- Necesaria declaración para Detectar-CapsLock-State -------------------------------------------
Private Declare PtrSafe Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Function Func_CapsLock_State() As Boolean
    Func_CapsLock_State = (GetKeyState(vbKeyCapital) = 1)      '- Check if Caps Lock is on then return 'True' ---------
End Function

Sub Rut_CapsLock_State()
    If Func_CapsLock_State() Then
        Debug.Print "Mayúsculas Activadas"
    Else
        Debug.Print "Mayúsculas Desactivadas"
    End If
End Sub


