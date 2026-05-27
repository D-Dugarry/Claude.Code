VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_MsgBox 
   Caption         =   "UserForm1"
   ClientHeight    =   6600
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   15760
   OleObjectBlob   =   "Form_MsgBox.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Form_MsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'- Necesitamos poner estas 4 declaraciones y definir la constante GWL_Style, para poder quitar la barra de menú de la ventana de este UserForm. --------------------
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Const GWL_Style = (-16)
'Public MsgBx_Title       As String
'Public MsgBx_TitleBar    As Boolean
'Public MsgBx_Msg         As String
'Public MsgBx_Answer      As Integer
' ==================================================================================================================================
Private Sub UserForm_Initialize()
    '- Situa la ventana en el centro de la pantalla -----------
    With Application
        Me.Top = .Top + (Me.Height - .Height) / 2
        Me.Left = .Left + (Me.Width - .Width) / 2
    End With
    If MsgBx_TitleBar Then
        Form_MsgBox.Caption = MsgBx_Title
    Else
        '- Elimina la barra de título de la ventana ---------------
        Dim hwnd As LongPtr
        hwnd = FindWindow(vbNullString, Me.Caption)
        If hwnd <> 0 Then
          SetWindowLong hwnd, GWL_Style, 0
          DrawMenuBar hwnd
        End If  '--------------------------------------------------
    End If
    MsgBx_Answer = 0
End Sub
' ==================================================================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Cancel if the user press Alt-F4
    Cancel = CloseMode = vbFormControlMenu
End Sub
' ==================================================================================================================================
' ==================== Put this code on the module where you need use. =============================================================
    '''        MsgBx_Msg = "Hello Baby!!" & vbLf & "Esto es un ejemplo de como quedaría el mensaje con un salto de línea y texto más largo que una línea del cuadro de mensaje."
    '''        MsgBx_Title = "Título de la ventana de mensaje"
    '''        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, False, "OK+Cancel+Stop", 2, "Ask"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
    '''
    '''        Select Case MsgBx_Answer
    '''            Case 1
    '''              Debug.Print "botón1"
    '''            Case 2
    '''              Debug.Print "botón2"
    '''            Case Else
    '''              Debug.Print "botón3"
    '''        End Select
' ==================================================================================================================================
'- Esto permite utilizar el UserForm y pasar parámetros para definir su visualización ----------------------------------------------
Public Sub SetParameter(Optional TamanoFont As Integer = 16, _
                        Optional MarcoRojo As Boolean = False, _
                        Optional Btns As String = "Ok", _
                        Optional BtnSelect As Integer = 1, _
                        Optional Img As String = "Msg")
'-----------------------------------------------------------------------------------------------------------------------------------
Dim Cant_Btn    As Integer: Cant_Btn = 0
Dim Btn_N       As Integer: Btn_N = 0
Dim Btn_Width   As Integer: Btn_Width = 0
Dim Indice      As Integer
Dim Btn(2)      As String   '- Array de 3 buttons, del 0 al 2
Dim BtnName     As MSForms.control

    If Not MsgBx_TitleBar Then
        LB_TitleForm = MsgBx_Title
    Else
        LB_TitleForm.Visible = False
        Fondo_TitleForm.Visible = False
        Lb_Mensaje.Top = Lb_Mensaje.Top - Fondo_TitleForm.Height
        Fondo_Msg.Top = Fondo_Msg.Top - Fondo_TitleForm.Height
        Fondo_Msg_Rojo.Top = Fondo_Msg_Rojo.Top - Fondo_TitleForm.Height
    End If
    '- Configura el mensaje y ajusta su tamaño -------------------
    With Lb_Mensaje
        .Caption = MsgBx_Msg
        .Font.Size = TamanoFont
        .AutoSize = True
        .Height = .Height
        .AutoSize = False
        .Width = 780
        .WordWrap = True
    End With
    With Fondo_Msg
        .Height = Lb_Mensaje.Height + 20
    End With
    With Fondo_Msg_Rojo
        .Height = Lb_Mensaje.Height + 30
        .Visible = MarcoRojo
    End With
    '- Visualiza una imagen -------------------------------------
    If Img <> "" Then       '- Ask, Cancel, Exclam, Ok, Stop
        On Error Resume Next
        Set BtnName = Me.Controls("Lb_Img_" & Img)
        With BtnName
            .Visible = True
            .Top = Fondo_Msg_Rojo.Top + Fondo_Msg_Rojo.Height / 2 - .Height / 2
            .Left = 30
        End With
        On Error GoTo 0
    End If
    '- Posiciona los Buttons seleccionados ---------------------
    If Len(Btns) > 0 Then
        Do While InStr(Btns, "+") > 0
            Btn(Btn_N) = Left(Btns, InStr(Btns, "+") - 1)
            Btn_Width = Application.Max(Btn_Width, Len(Btn(Btn_N)) * 16)
            Btns = Mid(Btns, InStr(Btns, "+") + 1)
            Btn_N = Btn_N + 1
        Loop
        Btn(Btn_N) = Btns
        Btn_Width = Application.Max(Btn_Width, Len(Btn(Btn_N)) * 16)
        If Btn_Width < 50 Then Btn_Width = 50
        For Indice = 1 To Btn_N + 1
            Set BtnName = Me.Controls("Btn_" & Indice)
            BtnName.Caption = Btn(Indice - 1)
            With BtnName
                .Top = Fondo_Msg_Rojo.Top + Fondo_Msg_Rojo.Height + 30
                .Left = Fondo_Msg_Rojo.Left + Fondo_Msg_Rojo.Width - (Btn_Width * (Btn_N + 2 - Indice)) - (10 * (Btn_N + 2 - Indice) - 1)
                .Visible = True
                .Width = Btn_Width
            End With
        Next
            Set BtnName = Me.Controls("Btn_" & BtnSelect)
            BtnName.SetFocus

    End If
    Me.Height = Fondo_Msg_Rojo.Top + Fondo_Msg_Rojo.Height + 30 + Btn_1.Height + 20 + (-MsgBx_TitleBar * 25)
    '- --------------------------------------------------------
End Sub
' ==================================================================================================================================
Private Sub Btn_1_Click()
    MsgBx_Answer = 1
    Unload Me
End Sub
' ==================================================================================================================================
Private Sub Btn_2_Click()
    MsgBx_Answer = 2
    Unload Me
End Sub
' ==================================================================================================================================
Private Sub Btn_3_Click()
    MsgBx_Answer = 3
    Unload Me
End Sub
' ==================================================================================================================================

