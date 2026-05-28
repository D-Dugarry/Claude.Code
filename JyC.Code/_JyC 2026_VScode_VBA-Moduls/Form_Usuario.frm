VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Usuario 
   Caption         =   "Identificación de Usuario"
   ClientHeight    =   8000
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

Dim Index_Usuario       As Variant

' ------------------------------------------------------------------------------------------------------
Function Func_CapsLock_State() As Boolean
    Func_CapsLock_State = (GetKeyState(vbKeyCapital) = 1)      '- Check if Caps Lock is on then return 'True' ---------
End Function
' ------------------------------------------------------------------------------------------------------
Private Sub TBx_Usuario_ID_Change()
    If Func_CapsLock_State() Then Lb_CapsLock_ON.Visible = True Else Lb_CapsLock_ON.Visible = False
End Sub
' ------------------------------------------------------------------------------------------------------
Sub Btn_Eixir_Click()
    Application.EnableEvents = True
    Application.Visible = True
    Unload Me
End Sub
' ------------------------------------------------------------------------------------------------------
Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Prog__APP.Range("APP_User_ID") = "" Then
        MsgBx_Msg = "¿Cerramos la aplicación?" & vbCrLf & "¿¿¿ Seguro que desea continuar SIN Grabar los Datos ???"
        MsgBx_Title = "Procedimiento: Identificación de Usuario."
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "Yes+No", 2, "Ask"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
        
        If MsgBx_Answer = 1 Then
            MsgBx_Msg = "¡¡ Hasta luego Lukas... !!"
            Load Form_MsgBox: Call Form_MsgBox.SetParameter: Form_MsgBox.Show   '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
            Unload Me
            Application.Visible = True
            ThisWorkbook.Close False
        End If
        Cancel = 1
        Exit Sub
    End If
    Prog__APP.Range("APP_Task_Inf") = "Nuevo Usuario: " & Prog__APP.Range("APP_User_Name") & ",  Email: " & Prog__APP.Range("APP_User_Mail") & ",  Ext. " & Prog__APP.Range("APP_User_Ext")
End Sub
' ------------------------------------------------------------------------------------------------------
 Sub UserForm_Initialize()       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
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
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
 Sub UserForm_Activate()
'    Dim Lin_Task     As Variant
'    With Prog__MnAux.ListObjects(1).DataBodyRange
'        Lin_Task = Application.Match(CtrlTag, .Columns(Task_Uribbon_Tags), 0)
'        If Not IsError(Lin_Task) Then
'            Func_SuperTip_Value = .Cells(Lin_Task, Task_Descripción)
'            If Prog__APP.Range("SW_Boss") Then
'                Func_SuperTip_Value = Func_SuperTip_Value & vbLf & vbLf & _
'                                     "BOSS_Rut: " & .Cells(Lin_Task, Task_Nombre_Rut) & vbLf & vbLf & _
'                                     "Informe:" & vbLf & .Cells(Lin_Task, Task_Rut_Informe)
'            End If
'        Else
'            Func_SuperTip_Value = "¡ Control.Tag, NO encontrado !"
'        End If
'    End With

 
    Prog__APP.Range("APP_User_ID") = ""
    Prog__APP.Range("App_User_Unid_Red") = ""
    Prog__APP.Range("APP_User_Mail") = ""
    Prog__APP.Range("APP_User_RutaAPP") = ""
    App_RutaAPP = ""
    Prog__APP.Range("SW_Boss") = False
    Prog__APP.Range("SW_Test") = False
    If Func_CapsLock_State() Then Lb_CapsLock_ON.Visible = True
    Me.TBx_Usuario_ID.SetFocus
End Sub
' ------------------------------------------------------------------------------------------------------
Sub Btn_Validar_Click()
    Prog__APP.Range("APP_User_ID") = ""
    If Trim(TBx_Usuario_ID.Value) = "" Then                                        '- Sin Rellenar ----------------------
        Lbl_Obl_ID.Visible = True
'        MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Sin Usuario especificado", _
'                            vbOKOnly + vbExclamation, "Proceso: Identificación Usuario"
        MsgBx_Msg = "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Sin Usuario especificado"
        MsgBx_Title = "Proceso: Identificación Usuario"
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show  '- ([Font-Size]=16, [Red-Border]=False, [Buttons]="Ok", [Default-Button]=1, [Image]="Msg")
        GoTo Volver_a_Preguntar
    End If
    With Prog__Users.ListObjects(1).DataBodyRange
        Index_Usuario = Application.Match(TBx_Usuario_ID.Value, .Columns(1), 0)
        If IsError(Index_Usuario) Then                                          '- NO hay coincidencia ---------------
'            MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Usuario Desconocido", _
'                                        vbOKOnly + vbExclamation, "Proceso: Identificación Usuario"
            MsgBx_Msg = "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Usuario Desconocido"
            MsgBx_Title = "Proceso: Identificación Usuario"
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show  '- ([Font-Size]=16, [Red-Border]=False, [Buttons]="Ok", [Default-Button]=1, [Image]="Msg")
            GoTo Volver_a_Preguntar
        End If
        If TBx_Usuario_ID.Value <> .Cells(Index_Usuario, 1) Then                   '- NO hay coincidencia Mays/Mins -----
'            MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Usuario Desconocido", _
'                                        vbOKOnly + vbExclamation, "Proceso: Identificación Usuario"
            MsgBx_Msg = "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Usuario Desconocido"
            MsgBx_Title = "Proceso: Identificación Usuario"
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show  '- ([Font-Size]=16, [Red-Border]=False, [Buttons]="Ok", [Default-Button]=1, [Image]="Msg")
            GoTo Volver_a_Preguntar  '- NO hay coincidencia Mays/Mins -----
        End If
        '- La coincidencia es TOTAL distingue mayúsculas de minúsculas -----------------------------------------------
        Prog__APP.Range("APP_User_ID") = .Cells(Index_Usuario, 1)
        Prog__APP.Range("APP_User_Name") = .Cells(Index_Usuario, 2)
        Prog__APP.Range("APP_User_Ext") = .Cells(Index_Usuario, 7)
        Prog__APP.Range("APP_User_Rubrica") = .Cells(Index_Usuario, 8)
        Tbx_UserName = Prog__APP.Range("APP_User_Name")
        Tbx_UserExt = Prog__APP.Range("APP_User_Ext")
        Tbx_UserMail = .Cells(Index_Usuario, 6)
        If CheckBox_Teletrabajo.Value = True Then
            Prog__APP.Range("App_User_Unid_Red") = .Cells(Index_Usuario, 5)
        Else
            Prog__APP.Range("App_User_Unid_Red") = .Cells(Index_Usuario, 4)
        End If
        Prog__APP.Range("APP_User_Mail") = .Cells(Index_Usuario, 6)
        Prog__APP.Range("APP_User_RutaAPP") = Fnc_NEXE_RutaAPP
        Tbx_UserNEXE = Prog__APP.Range("App_User_Unid_Red")
        If Prog__APP.Range("APP_User_ID") = "Boss" Then Prog__APP.Range("SW_Boss") = True Else Prog__APP.Range("SW_Boss") = False

Debug.Print "Form_Usuario, Btn_Validar_Click"

        Call Rut_Filtrar_Tareas(Prog__APP.Range("APP_User_ID"))
        Btn_Eixir.Visible = True

        Exit Sub
    End With    ' Prog__Users.ListObjects(1).DataBodyRange
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
    Prog__Users.ListObjects(1).DataBodyRange.Cells(Index_Usuario, 2) = Me.Tbx_UserName
    Prog__APP.Range("APP_User_Name") = Me.Tbx_UserName
'Debug.Print "Tbx_UserName_AfterUpdate: " & Prog__APP.Range("APP_User_Name")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserExt_AfterUpdate()
    Prog__Usuarios.ListObjects(1).DataBodyRange.Cells(Index_Usuario, 2) = Me.Tbx_UserExt
    Prog__APP.Range("APP_User_Ext") = Me.Tbx_UserExt
'Debug.Print "Tbx_UserName_AfterUpdate: " & Prog__APP.Range("APP_User_Ext")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserMail_AfterUpdate()
    With Prog__Usuarios.ListObjects(1).DataBodyRange
        Prog__APP.Range("APP_User_Mail") = Me.Tbx_UserMail
        .Cells(Index_Usuario, 6) = Prog__APP.Range("APP_User_Mail")
        Prog__APP.Range("APP_User_RutaAPP") = Fnc_NEXE_RutaAPP
    End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
'Debug.Print "Tbx_UserMail_AfterUpdate: " & Prog__APP.Range("APP_User_Mail"), Prog__APP.Range("APP_User_RutaAPP")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserNEXE_AfterUpdate()
    With Prog__Users.ListObjects(1).DataBodyRange
        Prog__APP.Range("App_User_Unid_Red") = Me.Tbx_UserNEXE
        If CheckBox_Teletrabajo.Value = True Then
            .Cells(Index_Usuario, 5) = Prog__APP.Range("App_User_Unid_Red")
        Else
            .Cells(Index_Usuario, 4) = Prog__APP.Range("App_User_Unid_Red")
        End If
        Prog__APP.Range("APP_User_RutaAPP") = Fnc_NEXE_RutaAPP
    End With    ' Prog__Users.ListObjects(1).DataBodyRange
'Debug.Print "CheckBox_Teletrabajo_AfterUpdate: " & App_RutaAPP
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub CheckBox_Teletrabajo_AfterUpdate()
    With Prog__Users.ListObjects(1).DataBodyRange
        If CheckBox_Teletrabajo.Value = True Then
            Prog__APP.Range("App_User_Unid_Red") = .Cells(Index_Usuario, 5)
        Else
            Prog__APP.Range("App_User_Unid_Red") = .Cells(Index_Usuario, 4)
        End If
        Prog__APP.Range("APP_User_RutaAPP") = Fnc_NEXE_RutaAPP
    End With    ' Prog__Users.ListObjects(1).DataBodyRange
    Me.Tbx_UserNEXE = Prog__APP.Range("App_User_Unid_Red")
'Debug.Print "Tbx_UserNEXE_AfterUpdate: " & App_RutaAPP
End Sub
