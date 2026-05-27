VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Menu 
   Caption         =   "Menú de tareas auxiliares"
   ClientHeight    =   15420
   ClientLeft      =   1110
   ClientTop       =   3180
   ClientWidth     =   29760
   OleObjectBlob   =   "Form_Menu.frx":0000
End
Attribute VB_Name = "Form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Lo_Tareas           As ListObject
Dim TaskIndice   As Integer

' ------------------------------------------------------------------------------------------------------
Private Sub Btn_Eixir_Click()
    Unload Me
End Sub
' ------------------------------------------------------------------------------------------------------
Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Debug.Print "Sub UserForm_QueryClose() - Form_menu"
    Set Lo_Tareas = Nothing
'    Application.ScreenUpdating = True
'    DoEvents
End Sub
' ------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------
Sub UserForm_Initialize()
Debug.Print "Sub UserForm_Initialize() - Form_menu"
    With Application
        Zoom = Int(.Width / Me.Width * 100)
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
    Prog__MnAux.Unprotect
    If Lo_Tareas Is Nothing Then Set Lo_Tareas = Prog__MnAux.ListObjects(1)
    Call Rut_LstObj_Sort(Lo_Tareas, 1, xlAscending, True)
End Sub
' ------------------------------------------------------------------------------------------------------
 Sub UserForm_Activate()
Debug.Print "Sub UserForm_Activate() - Form_menu"
    Me.Tbx_AñoContable = Prog__APP.Range("APP_AñoCont")
    
    Me.Lb_SW_Test.Visible = Prog__APP.Range("SW_Test")
    Me.Lb_SW_Boss.Visible = Prog__APP.Range("SW_Boss")
    Me.Lb_SW_WB_Deactivate.Visible = Prog__APP.Range("SW_WB_Deactivate")
       
    If Prog__APP.Range("SW_Boss") Then
        Me.Lb_Tarea_Name.Visible = True
        Me.TBx_Tarea_Name.Visible = True
        Me.Lb_Rutina_Name.Visible = True
        Me.TBx_Rutina_Name.Visible = True
    Else
        Me.Lb_Tarea_Name.Visible = False
        Me.TBx_Tarea_Name.Visible = False
        Me.Lb_Rutina_Name.Visible = False
        Me.TBx_Rutina_Name.Visible = False
    End If
    
    Me.Tbx_UserName = Prog__APP.Range("APP_User_Name")
    Me.Tbx_UserExt = Prog__APP.Range("APP_User_Ext")
    Me.Tbx_UserNEXE = Prog__APP.Range("App_User_Unid_Red") & ":"
    
    Me.Lb_Fondo_Inf.BackColor = RGB(255, 222, 255)          '- Rosa Pink   RGB(255, 224, 192)
    Me.Lb_Fondo_Tit_Inf.BackColor = RGB(255, 224, 192)     '- Rosa Pink
    
    Call Mostrar_Tareas
        
End Sub     ' UserForm_Activate    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------
Sub Tbx_AñoContable_AfterUpdate()
    If Me.Tbx_AñoContable >= 2022 Then
        Prog__APP.Range("APP_AñoCont") = Me.Tbx_AñoContable
        Me.Lb_Tít_Informe.Caption = "Cambiado Año Contable "
        Me.TBx_Informe = "Nuevo Año Contable: " & Prog__APP.Range("APP_AñoCont")
    Else
        MsgBox "¡¡¡ Año incorrecto, debe ser mayor que 2021 !!!", vbOKOnly, "Proceso: Procedimiento de Liquidación"
        Me.Tbx_AñoContable = Prog__APP.Range("APP_AñoCont")
        Me.Tbx_AñoContable.SetFocus
    End If
End Sub
' ------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------
Sub Mostrar_Tareas()
Dim Cont_Row        As Integer
Dim Cant_Tareas         As Integer:     Cant_Tareas = 0
Dim Ancho_Tareas        As Integer:     Ancho_Tareas = 0
Dim Nombre_Rut      As String

    Call Rut_LstObj_Sort(Lo_Tareas, 1, xlAscending, True)
    
    ' Cargo la lista desplegable de Tareas -------------------
    Me.LBx_Tareas.Clear
    With Lo_Tareas.DataBodyRange
        For Cont_Row = 1 To Lo_Tareas.ListRows.Count
            If (InStr(UCase(.Cells(Cont_Row, Task_Usuario)), UCase(Prog__APP.Range("APP_User_ID"))) > 0 Or .Cells(Cont_Row, Task_Usuario) = "") _
                And .Cells(Cont_Row, Task_Tarea) <> "INTERNO" Then
                Nombre_Rut = .Cells(Cont_Row, Task_Tarea)
                Me.LBx_Tareas.AddItem (Nombre_Rut)
                Ancho_Tareas = Application.WorksheetFunction.Max(Ancho_Tareas, Len(Nombre_Rut))
                Cant_Tareas = Cant_Tareas + 1
            End If
        Next Cont_Row
    End With '- Lo_Tareas.DataBodyRange
    With LBx_Tareas
        .Height = Application.Min(450, Cant_Tareas * 17.5)
        .IntegralHeight = False
        .Height = .Height
        .IntegralHeight = True
'        .Width = Application.WorksheetFunction.Min(Ancho_Tareas * 8, 426)
        .Width = 426
    End With
    Me.LBx_Tareas.Visible = True

End Sub     ' Mostrar_Tareas    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' ------------------------------------------------------------------------------------------------------
Sub Btn_Ejec_Tarea_Click()
Dim Pos_Delimitador     As Integer
Dim Rutinas_Name        As String
Dim Rut_Name        As String
    Rutinas_Name = Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Nombre_Rut)
Dim ImagenName        As String
    ImagenName = Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Imagen)
    
    Prog__APP.Range("APP_Task_Index") = TaskIndice
    Me.Lb_Fondo_Inf.BackColor = RGB(255, 222, 255)          '- Rosa Pink   RGB(255, 224, 192)
    Me.Lb_Fondo_Tit_Inf.BackColor = RGB(255, 224, 192)     '- Rosa Pink
    Me.Lb_Tít_Informe.Caption = "Progreso Tarea: " & Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Tarea)
    Me.TBx_Informe = ""
    
    Do       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Pos_Delimitador = InStr(Rutinas_Name, " + ")
        If Pos_Delimitador < 1 Then ' ------------------ Última o Única Rutina ---------------------
            Pos_Delimitador = Len(Rutinas_Name) + 1
        End If
        Rut_Name = Left(Rutinas_Name, Pos_Delimitador - 1)
        Prog__APP.Range("APP_Task_Index") = TaskIndice
        Prog__APP.Range("APP_Task_Rut") = Rut_Name
        Rutinas_Name = Mid(Rutinas_Name, Pos_Delimitador + 3)
        Prog__APP.Range("APP_Task_Inf") = ""
        If ImagenName <> "" Then
            Me.Controls(ImagenName).Visible = True
            Application.Run Rut_Name
            Me.Controls(ImagenName).Visible = False
        Else
            Application.Run Rut_Name
        End If
        Me.Lb_Tít_Informe.Caption = "Informe Tarea: " & Lo_Tareas.DataBodyRange.Cells(Prog__APP.Range("APP_Task_Index"), Task_Tarea)
        If Prog__APP.Range("APP_Task_Inf") <> "" Then Me.TBx_Informe = Prog__APP.Range("APP_Task_Inf")
        
        Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Rut_Informe) = Me.TBx_Informe
        Prog__APP.Range("APP_Task_Inf") = Me.TBx_Informe
        Me.Lb_Fondo_Tit_Inf.BackColor = RGB(192, 255, 192)
        Me.Lb_Fondo_Inf.BackColor = RGB(255, 222, 255)
        
    Loop While Len(Rutinas_Name) > 0
    
    Me.Lb_Fondo_Inf.BackColor = RGB(192, 255, 192)
    Me.Lb_Fondo_Tit_Inf.BackColor = RGB(192, 255, 192)
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
 Sub TBx_Tarea_Name_AfterUpdate()
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
    Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Tarea) = Me.TBx_Tarea_Name
    Me.TBx_Tarea_Name = ""
    Call Mostrar_Tareas
    Me.Lb_Tít_Informe.Caption = "Tarea realizada: "
    Me.TBx_Informe = "Cambio del Nombre de la Tarea."
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
End Sub
' ------------------------------------------------------------------------------------------------------
Sub TBx_Rutina_Name_AfterUpdate()
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
    Lo_Tareas.DataBodyRange.Cells(TaskIndice, Task_Nombre_Rut) = Me.TBx_Rutina_Name
    Me.TBx_Rutina_Name = ""
    Call Mostrar_Tareas
    Me.Lb_Tít_Informe.Caption = "Tarea realizada: "
    Me.TBx_Informe = "Cambio del Nombre de la Rutina de la Tarea."
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
End Sub




