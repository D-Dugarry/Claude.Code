VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Running_Rut 
   Caption         =   "UserForm1"
   ClientHeight    =   10520
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   19470
   OleObjectBlob   =   "Form_Running_Rut.frx":0000
End
Attribute VB_Name = "Form_Running_Rut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Last Rev. 2026-10-04 13:09

Option Explicit

'- Necesitamos poner estas 4 declaraciones y definir la constante GWL_Style, para poder quitar la barra de menú de la ventana de este UserForm. --------------------
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Const GWL_Style = (-16)

'==================================================================================================
' Form_Running_Rut - formulario de progreso de las tareas largas de los botones del Ribbon
'
' USO (2026-10-04, patron de Jornadas y Congresos): el boton lo abre, llama a su rutina y lo cierra:
'      Call Rut_Progreso_Abrir("Titulo")          '- M_90_Rutinas_Menu_Aux: Load + titulo + Show vbModeless
'      Call RuT_Update_LSGES04_ACont             '- escribe su progreso en ActivForm.Controls("TBx_Informe")
'      Call Rut_Progreso_Cerrar(control.Tag)     '- informe a APP_Task_Inf y Lo_RibbonUI + Rut_Finalizada
'
' Antes el boton dejaba el nombre de la rutina en APP_Task_Rut y abria este formulario en modal; su
' UserForm_Activate la buscaba en Tb_Tareas (_Menu_Aux), ponia de titulo el nombre de la tarea y la
' lanzaba con Application.Run. Ese camino se ha retirado: el formulario ya no busca ni lanza nada.
'==================================================================================================

' ==================================================================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = (CloseMode = vbFormControlMenu)    'Cancel if the user press Alt-F4
Debug.Print "-################################- <<< Form_Running_Rut - UserForm_QueryClose - "
End Sub
' ------------------------------------------------------------------------------------------------------
Sub UserForm_Initialize()       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Debug.Print "............... >>> Form_Running_Rut - UserForm_Initialize() - "
    Btn_VerInf.Visible = False      '- se quedo sin codigo (2026-10-04): que no estorbe
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
Debug.Print "............... <<< Form_Running_Rut - UserForm_Initialize() - "
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Public Sub Rut_Finalizada()    '- remate visual: fondo verde + boton de salida (lo llama Rut_Progreso_Cerrar)
    Btn_VerInf.Visible = False
    Btn_Eixir.Visible = True
    Me.Fnd_Tarea.BackColor = RGB(192, 255, 192)
    Me.Fnd_Tit.BackColor = RGB(192, 255, 192)
    Me.Repaint
    DoEvents                                 '- que el remate verde se vea de inmediato
End Sub
' --------------------------------------------------------------------------------------------------
'- Btn_VerInf_Click se elimino el 2026-10-04: volcaba el informe a la fila de Tb_Tareas de la tarea,
'  que las rutinas con boton ya no tienen. El boton sigue en el .frx, oculto.
' --------------------------------------------------------------------------------------------------
Private Sub Btn_Eixir_Click()
    Unload Me
End Sub
