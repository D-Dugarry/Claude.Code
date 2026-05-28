VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Running_Rut 
   Caption         =   "UserForm1"
   ClientHeight    =   10520
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   19470
   OleObjectBlob   =   "Form_Running_Rut.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Form_Running_Rut"
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
Dim TaskIndice             As Variant
' ==================================================================================================================================
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    'Cancel if the user press Alt-F4
'    Cancel = CloseMode = vbFormControlMenu
'End Sub
' ------------------------------------------------------------------------------------------------------
Sub UserForm_Initialize()       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Debug.Print "Sub UserForm_Initialize() - Form_Inf_Rut"
    With Application
        .WindowState = xlMaximized
        Zoom = Int(.Width / Me.Width * 100)
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
    '- Elimina la barra de título de la ventana ---------------
    Dim hwnd As LongPtr
    hwnd = FindWindow(vbNullString, Me.Caption)
    If hwnd <> 0 Then
      SetWindowLong hwnd, GWL_Style, 0
      DrawMenuBar hwnd
    End If  '--------------------------------------------------
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Sub UserForm_Activate()

Debug.Print "Sub UserForm_Activate() - Form_Inf_Rut"
    With Prog__MnAux.ListObjects(1).DataBodyRange    '-
        TaskIndice = Application.Match(Prog__APP.Range("APP_Task_Rut"), .Columns(Task_Nombre_Rut), 0)
        If Not IsError(TaskIndice) Then    ' ¡¡¡ Existe la Rutina !!! ------------------------
            Prog__APP.Range("APP_Task_Index") = TaskIndice
            Me.Lb_Tít_Informe.Caption = "Progreso de la Tarea: " & .Cells(TaskIndice, Task_Tarea)
        Else                            ' ¡¡¡ No Existe la Rutina !!!   ------------------------
            Me.Lb_Tít_Informe = "¡ Error: la Rutina NO Existe !"
            Me.TBx_Informe = "¡ Error: la Rutina NO Existe !"
            Btn_Eixir.Visible = True
            Me.Fnd_Tarea.BackColor = RGB(255, 145, 138)
            Me.Fnd_Tit.BackColor = RGB(255, 145, 138)
            Exit Sub
        End If
    End With
    Application.Run Prog__APP.Range("APP_Task_Rut").Value
    Btn_Eixir.Visible = True
        Me.Fnd_Tarea.BackColor = RGB(192, 255, 192)
        Me.Fnd_Tit.BackColor = RGB(192, 255, 192)
End Sub     ' UserForm_Activate    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Private Sub Btn_VerInf_Click()
Debug.Print "Sub Btn_VerInf_Click() - Form_Inf_Rut"
    Btn_VerInf.Visible = False
    Btn_Eixir.Visible = True
    With Prog__MnAux.ListObjects(1).DataBodyRange
        .Cells(Prog__APP.Range("APP_Task_Index"), Task_Rut_Informe) = Me.TBx_Informe
        Me.Lb_Tít_Informe.Caption = "Informe Tarea realizada: " & .Cells(Prog__APP.Range("APP_Task_Index"), Task_Tarea)
'        Me.TBx_Informe = prog__app.range("APP_Task_Inf")
    End With
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Btn_Eixir_Click()
    If IsError(TaskIndice) Then    ' ¡¡¡ No Existe la Rutina !!! ------------------------
        Prog__MnAux.ListObjects(1).DataBodyRange.Cells(TaskIndice, Task_Rut_Informe) = Me.TBx_Informe
        Prog__APP.Range("APP_Task_Inf") = Me.TBx_Informe
    End If
    Unload Me
End Sub

' ------------------------------------------------------------------------------------------------------
'Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Debug.Print "Sub UserForm_QueryClose() - Form_Inf_Rut"
'    Application.ScreenUpdating = True
'    DoEvents
'End Sub
' ------------------------------------------------------------------------------------------------------





