VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Menu 
   Caption         =   "Menú de tareas auxiliares"
   ClientHeight    =   13950
   ClientLeft      =   1170
   ClientTop       =   680
   ClientWidth     =   26450
   OleObjectBlob   =   "Form_Menu.frx":0000
End
Attribute VB_Name = "Form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TaskIndice   As Integer

Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

' ------------------------------------------------------------------------------------------------------
Private Sub Btn_Eixir_Click()
    Unload Me
End Sub

Private Sub Frame_C_Acad_Click()

End Sub

Private Sub Tbx_Informe_Change()

End Sub

' ------------------------------------------------------------------------------------------------------
Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Debug.Print "Sub UserForm_QueryClose() - Form_menu"
'    Application.ScreenUpdating = True
'    DoEvents
End Sub
' ------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------
 Sub UserForm_Initialize()       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
Debug.Print "Sub UserForm_Initialize() - Form_menu"
    With Application
        Me.Width = 1334
        Me.Height = 756
        Zoom = Int(.Width / Me.Width * 100)
'        Debug.Print ".Width", .Width
'        Debug.Print "Me.Width", Me.Width
'        Debug.Print "Int(.Width / Me.Width * 100)", Int(.Width / Me.Width * 100)
        'Zoom = 100
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
'    Dim fW As Single, fH As Single
'
'    ' Porcentaje de la ventana de Excel
'    fW = 1  '0.6   ' 60%
'    fH = 1  '0.6   ' 60%
'
'    With Application
'        Me.Width = .Width * fW
'        Me.Height = .Height * fH
'
'        ' Centrar en la ventana de Excel
'        Me.Left = .Left + (.Width - Me.Width) / 2
'        Me.Top = .Top + (.Height - Me.Height) / 2
'    End With
    
    'Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long  ' ¡ PONER AL PRINCIPIO DEL MÓDULO !
    Const SM_CXSCREEN As Long = 0   ' Ancho pantalla primaria en píxeles
    Const SM_CYSCREEN As Long = 1   ' Alto pantalla primaria en píxeles
    Dim BASE_WIDTH As Long, BASE_Height As Long
    BASE_WIDTH = GetSystemMetrics(SM_CXSCREEN)
    BASE_Height = GetSystemMetrics(SM_CYSCREEN)
    Debug.Print "Resolución pantalla primaria: " & BASE_WIDTH & " x " & BASE_Height & " píxeles"
'    Dim zoomFactor As Integer
'    zoomFactor = Int(Me.Width / BASE_WIDTH * 100)
'    If zoomFactor < 10 Then zoomFactor = 10
'    If zoomFactor > 400 Then zoomFactor = 400
'    Me.Zoom = zoomFactor
    
    Me.Caption = ThisWorkbook.Name & "        - Menú de tareas auxiliares."
    Prog__Menú_Aux.Unprotect
    If Lo_Tareas Is Nothing Then Set Lo_Tareas = Prog__Menú_Aux.ListObjects(1)
    Call Rut_Lo_Sort(Lo_Tareas, 1, xlAscending, True)

End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Sub UserForm_Activate()
Debug.Print "Sub UserForm_Activate() - Form_menu"
    Me.Tbx_AñoContable = Prog__APP.Range("APP_AñoCont")
    Me.FrOpBt_CAcadAnt.Caption = Prog__APP.Range("APP_C_Acad_Ant")
    Me.FrOpBt_CAcadPos.Caption = Prog__APP.Range("APP_C_Acad_Pos")
    
    Me.Lb_SW_Test.Visible = Prog__APP.Range("SW_Test")
    Me.Lb_SW_Boss.Visible = Prog__APP.Range("SW_Boss")
    Me.Lb_SW_WB_Deactivate.Visible = Prog__APP.Range("SW_WB_Deactivate")
    Me.Lb_SW_DelRegNeg.Visible = Prog__APP.Range("SW_DelRegNeg")
    
    If Prog__APP.Range("APP_CursAcad") = Prog__APP.Range("APP_C_Acad_Ant") Then
        Me.FrOpBt_CAcadAnt.Value = True
    Else
        Me.FrOpBt_CAcadPos.Value = True
    End If

'    If Prog__APP.Range("APP_TitP_o_Curs") = "Tít. Propios" Then
'        Me.FrOpBt_TitP.Value = True
'    Else
'        Me.FrOpBt_Curs.Value = True
'    End If

    If Prog__APP.Range("SW_Boss") Then
        Me.Lb_Tarea_Name.Visible = True
        Me.TBx_Tarea_Name.Visible = True
        Me.Lb_Rutina_Name.Visible = True
        Me.TBx_Rutina_Name.Visible = True
        Me.Frame_Select_Tipo_EP.Visible = True
    Else
        Me.Lb_Tarea_Name.Visible = False
        Me.TBx_Tarea_Name.Visible = False
        Me.Lb_Rutina_Name.Visible = False
        Me.TBx_Rutina_Name.Visible = False
        Me.Frame_Select_Tipo_EP.Visible = False
    End If
    
    Me.Tbx_UserName = Prog__APP.Range("APP_User_Name")
    Me.Tbx_UserExt = Prog__APP.Range("APP_User_Ext")
    Me.Tbx_Unidad_NEXE = Prog__APP.Range("APP_User_Unid_Red") & ":"
    
    Me.Lb_Fondo_Inf.BackColor = RGB(255, 222, 255)          '- Rosa Pink   RGB(255, 224, 192)
    Me.Lb_Fondo_Tit_Inf.BackColor = RGB(255, 224, 192)     '- Rosa Pink
    
    Call Mostrar_Tareas
        
End Sub     ' UserForm_Activate    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------
Sub Tbx_AñoContable_AfterUpdate()
    If Me.Tbx_AñoContable >= 2022 Then
        Prog__APP.Range("APP_AñoCont") = Me.Tbx_AñoContable
        Me.FrOpBt_CAcadAnt.Value = True
        Me.FrOpBt_CAcadAnt.Caption = Prog__APP.Range("APP_AñoCont") - 1 & "-" & Right(Prog__APP.Range("APP_AñoCont"), 2)
        Prog__APP.Range("APP_CursAcad") = Me.FrOpBt_CAcadAnt.Caption
        Prog__APP.Range("APP_C_Acad_Ant") = Me.FrOpBt_CAcadAnt.Caption
        Me.FrOpBt_CAcadPos.Caption = Prog__APP.Range("APP_AñoCont") & "-" & Right(Prog__APP.Range("APP_AñoCont"), 2) + 1
        Prog__APP.Range("APP_C_Acad_Pos") = Me.FrOpBt_CAcadPos.Caption
    Else
        MsgBox "¡¡¡ Año incorrecto, debe ser mayor que 2021 !!!", vbOKOnly, "Proceso: Procedimiento de Liquidación"
        Me.Tbx_AñoContable = Prog__APP.Range("APP_AñoCont")
        Me.Tbx_AñoContable.SetFocus
    End If
End Sub
' ------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------
Sub Mostrar_Tareas()
Dim Cont_Row            As Integer
Dim Cant_Tareas         As Integer:     Cant_Tareas = 0
Dim Ancho_Tareas        As Integer:     Ancho_Tareas = 0
Dim Nombre_Rut          As String
Dim Task_NotInterno     As Boolean
Dim Users               As String
Dim Users_Allowed       As Boolean
Dim SheetBtn            As String
Dim SheetBtn_OK         As Boolean
Dim Usuario_ID          As String:      Usuario_ID = UCase(Prog__APP.Range("APP_User_ID"))
Dim Rng_SW_Boss          As Range:      Set Rng_SW_Boss = Prog__APP.Range("SW_Boss")


    Call Rut_Lo_Sort(Lo_Tareas, 1, xlAscending, True)
    
    ' Cargo la lista desplegable de Tareas -------------------
    Me.LBx_Tareas.Clear
    With Lo_Tareas.DataBodyRange
    For Cont_Row = 1 To .Rows.Count
    
            Users = UCase(.Cells(Cont_Row, Task_Usuario))
            Users_Allowed = (Users = "" Or InStr(Users, Usuario_ID) > 0)
            
            SheetBtn = UCase(.Cells(Cont_Row, Task_SheetsButton))
            SheetBtn_OK = (SheetBtn = "" Or InStr(SheetBtn, UCase(ActiveSheet.Name)) > 0)
            
            Task_NotInterno = (InStr(UCase(.Cells(Cont_Row, Task_Tarea)), "INTERNO") = 0)
            
        If (Users_Allowed And SheetBtn_OK Or Rng_SW_Boss) And Task_NotInterno Then
            Nombre_Rut = .Cells(Cont_Row, Task_Tarea)
            Me.LBx_Tareas.AddItem (Nombre_Rut)
            Ancho_Tareas = Application.Max(Ancho_Tareas, Len(Nombre_Rut))
            Cant_Tareas = Cant_Tareas + 1
        End If
    Next Cont_Row
    End With
'    With LBx_Tareas
'        .Height = Application.Min(400, Cant_Tareas * 17.5)
'        .IntegralHeight = False
'        .Height = .Height
'        .IntegralHeight = True
'        .Width = Application.Min(Ancho_Tareas * 8, 400)
'        .Width = 400
'    End With
    Me.LBx_Tareas.Visible = True

End Sub     ' Mostrar_Tareas    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Sub Btn_Ejec_Tarea_Click()
    Dim Pos_Delimitador     As Integer
    Dim Rut_Name            As String
    Dim Rutinas_Name        As String
        Rutinas_Name = Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Nombre_Rut)
    Dim ImagenName          As String
        ImagenName = Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Imagen)
        
    Prog__APP.Range("APP_Task_Index") = TaskIndice
    Me.Lb_Fondo_Inf.BackColor = RGB(255, 222, 255)          '- Rosa Pink   RGB(255, 224, 192)
    Me.Lb_Fondo_Tit_Inf.BackColor = RGB(255, 224, 192)     '- Rosa Pink
    Me.Lb_Tít_Informe.Caption = "Tarea en Proceso: " & Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Tarea)
    Me.TBx_Informe = ""
    
    Do       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Pos_Delimitador = InStr(Rutinas_Name, " + ")
        If Pos_Delimitador < 1 Then ' ------------------ Última o Única Rutina ---------------------
            Pos_Delimitador = Len(Rutinas_Name) + 1
        End If
        Rut_Name = Left(Rutinas_Name, Pos_Delimitador - 1)
'        Prog__APP.Range("APP_Task_Index") = Index_RutAux
        Prog__APP.Range("APP_Task_Rut") = Rut_Name
        Rutinas_Name = Mid(Rutinas_Name, Pos_Delimitador + 3)
        If ImagenName <> "" Then
            Me.Controls(ImagenName).Visible = True
            Application.Run Rut_Name
            Me.Controls(ImagenName).Visible = False
        Else
            Application.Run Rut_Name
        End If
        Me.Lb_Tít_Informe.Caption = "Informe Tarea: " & Lo_Tareas.DataBodyRange.Cells(Prog__APP.Range("APP_Task_Index"), Task_Tarea)
        
        If Prog__APP.Range("APP_Task_Inf") <> "" Then Me.TBx_Informe = Prog__APP.Range("APP_Task_Inf")
        
        Lo_Tareas.DataBodyRange.Cells(Prog__APP.Range("APP_Task_Index"), Task_Rut_Informe) = Me.TBx_Informe.Text
        Prog__APP.Range("APP_Task_Inf") = Me.TBx_Informe.Text
        Me.Lb_Fondo_Inf.BackColor = RGB(192, 255, 192)
        Me.Lb_Fondo_Tit_Inf.BackColor = RGB(192, 255, 192)
        
    Loop While Len(Rutinas_Name) > 0
    
'    Me.Lb_Fondo_Inf.BackColor = RGB(192, 255, 192)
'    Me.Lb_Fondo_Tit_Inf.BackColor = RGB(192, 255, 192)
'    Application.ScreenUpdating = True
'    DoEvents

End Sub     ' Btn_Ejec_Tarea_Click   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Sub LBx_Tareas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Btn_Ejec_Tarea_Click
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Sub LBx_Tareas_Click()
    TaskIndice = Application.Match(Me.LBx_Tareas, Lo_Tareas.DataBodyRange.Columns(1), 0)
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
    Me.TBx_Descripción = Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Descripción)
    Me.TBx_Tarea_Name = Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Tarea)
    Me.TBx_Rutina_Name = Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Nombre_Rut)
    Me.Lb_Tít_Informe.Caption = "Informe última Tarea realizada: " & Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Tarea)
    Me.TBx_Informe = Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Rut_Informe)
    Me.Lb_Fondo_Inf.BackColor = RGB(224, 255, 255)         '- LightCyan
    Me.Lb_Fondo_Tit_Inf.BackColor = RGB(224, 255, 255)     '- LightCyan
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Sub TBx_Descripción_Change()
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
    Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Descripción) = Me.TBx_Descripción
    Me.Lb_Tít_Informe.Caption = "Tarea realizada: "
    Me.TBx_Informe = "Cambio de la Descripción de la Tarea."
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_Informe_AfterUpdate()
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
    Lo_Tareas.DataBodyRange.Cells(TaskIndice, 5) = Me.TBx_Informe.Text
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_Tarea_Name_AfterUpdate()
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
    Lo_Tareas.DataBodyRange.Cells(TaskIndice, 1) = Me.TBx_Tarea_Name
    Me.TBx_Tarea_Name = ""
    Call Mostrar_Tareas
    Me.Lb_Tít_Informe.Caption = "Tarea realizada: "
    Me.TBx_Informe = "Cambio del Nombre de la Tarea."
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub TBx_Rutina_Name_AfterUpdate()
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
    Lo_Tareas.DataBodyRange.Cells(TaskIndice, 3) = Me.TBx_Rutina_Name
    Me.TBx_Rutina_Name = ""
    Call Mostrar_Tareas
    Me.Lb_Tít_Informe.Caption = "Tarea realizada: "
    Me.TBx_Informe = "Cambio del Nombre de la Rutina de la Tarea."
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
End Sub
' ------------------------------------------------------------------------------------------------------
Sub FrOpBt_CAcadAnt_Click()
    Range("APP_CursAcad") = Me.FrOpBt_CAcadAnt.Caption
End Sub
' ------------------------------------------------------------------------------------------------------
Sub FrOpBt_CAcadPos_Click()
    Range("APP_CursAcad") = Me.FrOpBt_CAcadPos.Caption
End Sub
' ------------------------------------------------------------------------------------------------------
Sub FrOpBt_TitP_Click()
'    Range("APP_TitP_o_Curs") = Me.FrOpBt_TitP.Caption
'    Wk_TitP_Liquid.Name = "Liquid_Tit_Propios"
End Sub
' ------------------------------------------------------------------------------------------------------
Sub FrOpBt_Curs_Click()
'    Range("APP_TitP_o_Curs") = Me.FrOpBt_Curs.Caption
'    Wk_TitP_Liquid.Name = "Liquid_Cursos200h"
End Sub



