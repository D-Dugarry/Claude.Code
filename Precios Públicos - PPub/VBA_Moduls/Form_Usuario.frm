VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Usuario 
   Caption         =   "Identificación de Usuario"
   ClientHeight    =   7995
   ClientLeft      =   2120
   ClientTop       =   2460
   ClientWidth     =   17070
   OleObjectBlob   =   "Form_Usuario.frx":0000
End
Attribute VB_Name = "Form_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form_Usuario
Option Explicit

'- Necesaria declaración para Detectar-CapsLock-State -------------------------------------------
Private Declare PtrSafe Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

' ------------------------------------------------------------------------------------------------------
Function Func_CapsLock_State() As Boolean
Debug.Print "Func_CapsLock_State"
    Func_CapsLock_State = (GetKeyState(vbKeyCapital) = 1)      '- Check if Caps Lock is on then return 'True' ---------
End Function
' ------------------------------------------------------------------------------------------------------
Sub TBx_Usuario_ID_Change()
Debug.Print "TBx_Usuario_ID_Change"
    If Func_CapsLock_State() Then Lb_CapsLock_ON.Visible = True Else Lb_CapsLock_ON.Visible = False
End Sub
' ------------------------------------------------------------------------------------------------------
Sub Btn_Eixir_Click()
Debug.Print "Btn_Eixir_Click"
    Application.EnableEvents = True
    Application.Visible = True
    Unload Me
End Sub
' ------------------------------------------------------------------------------------------------------
Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Prog__APP.Range("APP_User_ID") = "" Then
        MsgBx_Msg = "¿Cerramos la aplicación?" & vbCrLf & "¿¿¿ Seguro que desea continuar SIN Grabar los Datos ???"
        MsgBx_Title = "Procedimiento: Identificación de Usuario."
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        
        If MsgBx_Answer = 1 Then
            MsgBx_Msg = "¡¡ Hasta luego Lukas... !!"
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Unload Me
            Application.Visible = True
            ThisWorkbook.Close False
        End If
        Cancel = 1
        Exit Sub
    End If
    Debug.Print "Nuevo Usuario: " & Prog__APP.Range("APP_User_Name") & vbLf & vbLf & "- Email: " & Prog__APP.Range("APP_User_Mail") & vbLf & "- Ext. " & Prog__APP.Range("APP_User_Ext")
Debug.Print "UserForm_QueryClose ------------------<<< Form Usuario "
End Sub
' ------------------------------------------------------------------------------------------------------
Sub UserForm_Initialize()       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Debug.Print "UserForm_Initialize ------------------>>> Form_Usuario "
    Form_Usuario.Caption = "   - Identificación de Usuario, es obligatorio para el correcto funcionamiento del programa."
    With Application
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
''    With Application
''        Me.Top = .Top + (.Height - Me.Height) / 2
''        Me.Left = .Left + (.Width - Me.Width) / 2
''    End With
End Sub
' ------------------------------------------------------------------------------------------------------
Sub UserForm_Activate()
Debug.Print "UserForm_Activate - Form_Usuario"
    Prog__APP.Range("APP_User_ID") = ""
    Prog__APP.Range("APP_User_Unid_Red") = ""
    Prog__APP.Range("APP_User_Mail") = ""
    Prog__APP.Range("APP_User_RutaAPP") = ""
    Prog__APP.Range("SW_Boss") = False
    Prog__APP.Range("SW_Test") = False
    App_RutaAPP = ""
    If Func_CapsLock_State() Then Lb_CapsLock_ON.Visible = True
    Me.TBx_Usuario_ID.SetFocus
End Sub
' ------------------------------------------------------------------------------------------------------
Sub Btn_Validar_Click()
Debug.Print "Btn_Validar_Click - Form_Usuario"
    Prog__APP.Range("APP_User_ID") = ""
    If Trim(TBx_Usuario_ID.Value) = "" Then      '- Sin Rellenar ----------------------
        Lbl_Obl_ID.Visible = True
        MsgBx_Msg = "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Sin Usuario especificado"
        MsgBx_Title = "Proceso: Identificación Usuario"
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        GoTo Volver_a_Preguntar
    End If
    With Prog__Usuarios.ListObjects(1).DataBodyRange
        Index_Usuario = Application.Match(TBx_Usuario_ID.Value, .Columns(1), 0)
        If IsError(Index_Usuario) Then        '- NO hay coincidencia ---------------
            MsgBx_Msg = "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Usuario Desconocido"
            MsgBx_Title = "Proceso: Identificación Usuario"
            MsgBox MsgBx_Msg, vbOKOnly + vbExclamation, MsgBx_Title
            GoTo Volver_a_Preguntar
        End If
        If TBx_Usuario_ID.Value <> .Cells(Index_Usuario, 1) Then      '- NO hay coincidencia Mays/Mins -----
            MsgBx_Msg = "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Usuario Desconocido"
            MsgBx_Title = "Proceso: Identificación Usuario"
            MsgBox MsgBx_Msg, vbOKOnly + vbExclamation, MsgBx_Title
            GoTo Volver_a_Preguntar  '- NO hay coincidencia Mays/Mins -----
        End If
        '- La coincidencia es TOTAL distingue mayúsculas de minúsculas -----------------------------------------------
        Prog__APP.Range("APP_User_ID") = .Cells(Index_Usuario, 1)
        Prog__APP.Range("APP_User_Name") = .Cells(Index_Usuario, 2)
        Prog__APP.Range("APP_User_Mail") = .Cells(Index_Usuario, 6)
        Prog__APP.Range("APP_User_Ext") = .Cells(Index_Usuario, 7)
        Tbx_UserName = Prog__APP.Range("APP_User_Name")
        Tbx_UserExt = Prog__APP.Range("APP_User_Ext")
        Tbx_UserMail = .Cells(Index_Usuario, 6)
        If CheckBox_Teletrabajo.Value = True Then
            Prog__APP.Range("APP_User_Unid_Red") = .Cells(Index_Usuario, 5)
        Else
            Prog__APP.Range("APP_User_Unid_Red") = .Cells(Index_Usuario, 4)
        End If
        Debug.Print "Nuevo Usuario: " & Prog__APP.Range("APP_User_Name") & vbLf & "Email: " & Prog__APP.Range("APP_User_Mail") & vbLf & "Etx.: " & Prog__APP.Range("APP_User_Ext")
        Prog__APP.Range("APP_User_RutaAPP") = Fnc_NEXE_RutaAPP
        Tbx_Unidad_NEXE = Prog__APP.Range("APP_User_Unid_Red")
        If Prog__APP.Range("APP_User_ID") = "Boss" Then Prog__APP.Range("SW_Boss") = True Else Prog__APP.Range("SW_Boss") = False
        Btn_Eixir.Visible = True
        Exit Sub
    End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
Volver_a_Preguntar:
    '- Usuario Erróneo ---------------
    Btn_Eixir.Visible = False
    Lbl_Obl_ID.Visible = False
    Prog__APP.Range("APP_User_ID") = ""
    Me.TBx_Usuario_ID = ""
    Me.TBx_Usuario_ID.SetFocus
    Application.SendKeys "{TAB}"       ' Getting setfocus back to the serial textbox
End Sub
' ------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserName_AfterUpdate()
    Prog__Usuarios.ListObjects(1).DataBodyRange.Cells(Index_Usuario, 2) = Me.Tbx_UserName
    Prog__APP.Range("APP_User_Name") = Me.Tbx_UserName
Debug.Print "Tbx_UserName_AfterUpdate: " & Prog__APP.Range("APP_User_Name")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserExt_AfterUpdate()
    Prog__Usuarios.ListObjects(1).DataBodyRange.Cells(Index_Usuario, 2) = Me.Tbx_UserExt
    Prog__APP.Range("APP_User_Ext") = Me.Tbx_UserExt
Debug.Print "Tbx_UserName_AfterUpdate: " & Prog__APP.Range("APP_User_Ext")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserMail_AfterUpdate()
    With Prog__Usuarios.ListObjects(1).DataBodyRange
        Prog__APP.Range("APP_User_Mail") = Me.Tbx_UserMail
        .Cells(Index_Usuario, 6) = Prog__APP.Range("APP_User_Mail")
        Prog__APP.Range("APP_User_RutaAPP") = Fnc_NEXE_RutaAPP
    End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
Debug.Print "Tbx_UserMail_AfterUpdate: " & Prog__APP.Range("APP_User_Mail"), Prog__APP.Range("APP_User_RutaAPP")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_Unidad_NEXE_AfterUpdate()
    With Prog__Usuarios.ListObjects(1).DataBodyRange
        Prog__APP.Range("APP_User_Unid_Red") = Me.Tbx_Unidad_NEXE
        If CheckBox_Teletrabajo.Value = True Then
            .Cells(Index_Usuario, 5) = Prog__APP.Range("APP_User_Unid_Red")
        Else
            .Cells(Index_Usuario, 4) = Prog__APP.Range("APP_User_Unid_Red")
        End If
        Prog__APP.Range("APP_User_RutaAPP") = Fnc_NEXE_RutaAPP
    End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
Debug.Print "CheckBox_Teletrabajo_AfterUpdate: " & Prog__APP.Range("APP_User_RutaAPP")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub CheckBox_Teletrabajo_AfterUpdate()
    With Prog__Usuarios.ListObjects(1).DataBodyRange
        If CheckBox_Teletrabajo.Value = True Then
            Prog__APP.Range("APP_User_Unid_Red") = .Cells(Index_Usuario, 5)
        Else
            Prog__APP.Range("APP_User_Unid_Red") = .Cells(Index_Usuario, 4)
        End If
            Prog__APP.Range("APP_User_RutaAPP") = ActiveWorkbook.Path
        Prog__APP.Range("APP_User_RutaAPP") = Fnc_NEXE_RutaAPP
    
    End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
    Me.Tbx_Unidad_NEXE = Prog__APP.Range("APP_User_Unid_Red")
Debug.Print "Tbx_Unidad_NEXE_AfterUpdate: " & Prog__APP.Range("APP_User_RutaAPP")
End Sub

