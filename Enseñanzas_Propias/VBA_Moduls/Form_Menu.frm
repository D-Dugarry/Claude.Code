VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Menu 
   Caption         =   "Menú de tareas auxiliares"
   ClientHeight    =   35445
   ClientLeft      =   620
   ClientTop       =   1050
   ClientWidth     =   23760
   OleObjectBlob   =   "Form_Menu.frx":0000
End
Attribute VB_Name = "Form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Last Rev. 2026-09-21 12:12
'- M02_Importar_LSGES04_GE - Modif: 2025-10-08
Option Explicit

Private Ws_Saltar_Al_Activar    As Boolean
Dim Idx_Tarea                   As Integer

Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Sub Btn_Eixir_Click()
    Unload Me
End Sub

' --------------------------------------------------------------------------------------------------
Sub OpBtn_CAcadAnt_Click()
    Range("APP_CursAcad") = Me.OpBtn_CAcadAnt.Caption
End Sub
' --------------------------------------------------------------------------------------------------
Sub OpBtn_CAcadPos_Click()
    Range("APP_CursAcad") = Me.OpBtn_CAcadPos.Caption
End Sub
' --------------------------------------------------------------------------------------------------
Sub OpBt_TP_Click()
    Range("APP_EFP_o_CFC") = "EFP"
End Sub
' --------------------------------------------------------------------------------------------------
Sub OpBt_CR_Click()
    Range("APP_EFP_o_CFC") = "CFC"
End Sub
' --------------------------------------------------------------------------------------------------
Sub ChBx_Prueba_Click()
    If Ws_Saltar_Al_Activar Then Exit Sub
    Prog__APP_Switch.Range("Sw_Probando") = Not Prog__APP_Switch.Range("Sw_Probando")
End Sub
' --------------------------------------------------------------------------------------------------
Private Sub ChBx_DisplayAlerts_Click()
    If Ws_Saltar_Al_Activar Then Exit Sub
    Prog__APP_Switch.Range("Sw_DisplayAlerts") = Not Prog__APP_Switch.Range("Sw_DisplayAlerts")
    Application.DisplayAlerts = Prog__APP_Switch.Range("Sw_EnableEvents")
End Sub
' --------------------------------------------------------------------------------------------------
Private Sub ChBx_EnableEvents_Click()
    If Ws_Saltar_Al_Activar Then Exit Sub
    Prog__APP_Switch.Range("Sw_EnableEvents") = Not Prog__APP_Switch.Range("Sw_EnableEvents")
    Application.EnableEvents = Prog__APP_Switch.Range("Sw_EnableEvents")
End Sub
' --------------------------------------------------------------------------------------------------
Private Sub ChBx_WB_Deactivate_Click()
    If Ws_Saltar_Al_Activar Then Exit Sub
    Prog__APP_Switch.Range("Sw_WB_Deactivate") = Not Prog__APP_Switch.Range("Sw_WB_Deactivate")
End Sub
' --------------------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------------------
Sub Tbx_AñoContable_AfterUpdate()
    If Me.Tbx_AñoContable >= 2022 Then
        Range("APP_AñoCont") = Me.Tbx_AñoContable
        Me.OpBtn_CAcadAnt.Value = True
        Me.OpBtn_CAcadAnt.Caption = Range("APP_AñoCont") - 1 & "-" & Right(Range("APP_AñoCont"), 2)
        Range("APP_CursAcad") = Me.OpBtn_CAcadAnt.Caption
        Range("APP_C_Acad_Ant") = Me.OpBtn_CAcadAnt.Caption
        Me.OpBtn_CAcadPos.Caption = Range("APP_AñoCont") & "-" & Right(Range("APP_AñoCont"), 2) + 1
        Range("APP_C_Acad_Pos") = Me.OpBtn_CAcadPos.Caption
    Else
        MsgBox "¡¡¡ Año incorrecto, debe ser mayor que 2021 !!!", vbOKOnly + vbExclamation, "Proceso: Procedimiento de Liquidación"
        Me.Tbx_AñoContable = Range("APP_AñoCont")
        Me.Tbx_AñoContable.SetFocus
    End If
End Sub
' --------------------------------------------------------------------------------------------------
Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Wk_TitP_Liquid.Unprotect
        Wk_TitP_Liquid.Range("Liquid_Curso_Acad") = Range("APP_CursAcad")
    Wk_TitP_Liquid.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA
    Application.ScreenUpdating = True
    DoEvents
End Sub
' --------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------
 Sub UserForm_Initialize()       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
Debug.Print "Sub UserForm_Initialize() - Form_menu"
    With Application
        Me.Width = 1334
        Me.Height = 756
        Zoom = Int(.Width / Me.Width * 100)
'        Debug.Print ".Width", .Width
'        Debug.Print "Me.Width", Me.Width
'        Debug.Print "Int(.Width / Me.Width * 100)", Int(.Width / Me.Width * 100)
'        Zoom = 100
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

End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' --------------------------------------------------------------------------------------------------
 Sub UserForm_Activate()

    Me.Tbx_AñoContable = Prog__APP.Range("APP_AñoCont")
    Me.OpBtn_CAcadAnt.Caption = Prog__APP.Range("APP_C_Acad_Ant")
    Me.OpBtn_CAcadPos.Caption = Prog__APP.Range("APP_C_Acad_Pos")
    
    If Range("APP_CursAcad") = Prog__APP.Range("APP_C_Acad_Ant") Then
        Me.OpBtn_CAcadAnt.Value = True
    Else
        Me.OpBtn_CAcadPos.Value = True
    End If

    If Range("APP_EFP_o_CFC") = "EFP" Then
        Me.OpBt_TP.Value = True
    Else
        Me.OpBt_CR.Value = True
    End If

    Ws_Saltar_Al_Activar = True ' para que no salte el evento ChBx_xxx_Click() e invierta el valor
    Me.ChBx_Prueba.Value = Prog__APP_Switch.Range("Sw_Probando").Value
    Me.ChBx_DisplayAlerts.Value = Prog__APP_Switch.Range("Sw_DisplayAlerts").Value
    Me.ChBx_EnableEvents.Value = Prog__APP_Switch.Range("Sw_EnableEvents").Value
    Me.ChBx_WB_Deactivate.Value = Prog__APP_Switch.Range("Sw_WB_Deactivate")
    Ws_Saltar_Al_Activar = False
    
    If Prog__APP_Switch.Range("Sw_Boss") Then
        Me.Lb_Tarea_Name.Visible = True
        Me.TBx_Tarea_Name.Visible = True
        Me.Lb_Rutina_Name.Visible = True
        Me.TBx_Rutina_Name.Visible = True
        Me.Frm_TP_o_CR.Enabled = True
        Me.Frm_CursoAcad.Enabled = True
        Me.Frm_Switches.Visible = True
        Me.Frm_Switches.Enabled = True
    Else
        Me.Frm_CursoAcad.Enabled = False
        Me.Frm_TP_o_CR.Enabled = False
        Me.Frm_Switches.Visible = False
        Me.Frm_Switches.Enabled = False
    End If
    
    Me.Tbx_UserName = Range("Usuario_Name")
    Me.Tbx_UserExt = Range("Usuario_Ext")
    Me.Tbx_Unidad_NEXE = Range("App_LetraUnidRed") & ":"
    
    Me.TB_Informe.BackColor = RGB(255, 224, 192)        '- Rosa Pink
    Me.Fondo_Inf.BackColor = RGB(255, 224, 192)         '- Rosa Pink
    Me.Lb_Tit_Informe.BackColor = RGB(255, 224, 192)    '- Rosa Pink
    Me.Fondo_Tit_Inf.BackColor = RGB(255, 224, 192)     '- Rosa Pink
    
    Call Mostrar_Tareas(Range("Usuario_ID"))
        
End Sub     ' UserForm_Activate    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' --------------------------------------------------------------------------------------------------
Sub Mostrar_Tareas(Usuario As String)
Dim Cont_Row        As Integer
Dim Cant_Tot_Tareas     As Integer
'Dim Cant_Tareas     As Integer
'Dim Ancho_Tareas    As Integer
Dim Nombre_Rut      As String

'    Ancho_Tareas = 0
'    Cant_Tareas = 0
    Cant_Tot_Tareas = Lo_Tareas.DataBodyRange.Rows.Count
    
    ' Cargo la lista desplegable de Tareas -------------------
    Me.LBx_Tareas.Clear
    For Cont_Row = 1 To Cant_Tot_Tareas
        If InStr(UCase(Lo_Tareas.DataBodyRange.Cells(Cont_Row, 2)), UCase(Usuario)) > 0 And Len(Usuario) > 3 Then
            Nombre_Rut = (Lo_Tareas.DataBodyRange.Cells(Cont_Row, 1))
            Me.LBx_Tareas.AddItem (Nombre_Rut)
'            Ancho_Tareas = Application.WorksheetFunction.Max(Ancho_Tareas, Len(Nombre_Rut))
'            Cant_Tareas = Cant_Tareas + 1
        ElseIf Lo_Tareas.DataBodyRange.Cells(Cont_Row, 2) = "" Then
            Nombre_Rut = (Lo_Tareas.DataBodyRange.Cells(Cont_Row, 1))
            Me.LBx_Tareas.AddItem (Nombre_Rut)
'            Ancho_Tareas = Application.WorksheetFunction.Max(Ancho_Tareas, Len(Nombre_Rut))
'            Cant_Tareas = Cant_Tareas + 1
        End If
    Next Cont_Row
     
'    With LBx_Tareas
'        .Height = Application.Min(400, Cant_Tareas * 17.5)
'        .IntegralHeight = False
'        .Height = .Height
'        .IntegralHeight = True
'        .Width = Application.WorksheetFunction.Min(Ancho_Tareas * 8, 500)
'    End With
     
     Me.LBx_Tareas.Visible = True

End Sub     ' Mostrar_Tareas    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' --------------------------------------------------------------------------------------------------
Sub Btn_Ejec_Tarea_Click()
Dim Pos_Delimitador     As Integer
Dim Rutinas_Name        As String
If Idx_Tarea = 0 Then
    MsgBox "¡¡ Seleccionar una Tarea !!", vbOKOnly + vbExclamation, "Menú de Tareas."
    Exit Sub
End If
Dim Rut_Name        As String
    Rutinas_Name = Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 3)
Dim ImagenName        As String
    ImagenName = Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 6)
    
'    Me.TB_Informe.BackColor = RGB(224, 255, 255)       '- LightCyan
'    Me.Fondo_Inf.BackColor = RGB(224, 255, 255)         '- LightCyan
'    Me.Lb_Tit_Informe.BackColor = RGB(224, 255, 255)    '- LightCyan
'    Me.Fondo_Tit_Inf.BackColor = RGB(224, 255, 255)     '- LightCyan
    Me.Lb_Tit_Informe.Caption = "Tarea en Proceso: " & Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 1)
    Me.TB_Informe = ""
    Index_Tarea = Idx_Tarea
    Do       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Pos_Delimitador = InStr(Rutinas_Name, " + ")
        If Pos_Delimitador < 1 Then ' ------------------ Última o Única Rutina ---------------------
            Pos_Delimitador = Len(Rutinas_Name) + 1
        End If
        Rut_Name = Left(Rutinas_Name, Pos_Delimitador - 1)
        Rutinas_Name = Mid(Rutinas_Name, Pos_Delimitador + 3)
        If ImagenName <> "" Then
            If Prog__APP_Switch.Range("Sw_Boss") Then Me.Frm_Switches.Visible = False
            Me.Controls(ImagenName).Visible = True
            On Error GoTo Nombre_Rutina_NO_Encontrado
            Application.Run Rut_Name
            On Error GoTo 0
            Me.Controls(ImagenName).Visible = False
            If Prog__APP_Switch.Range("Sw_Boss") Then Me.Frm_Switches.Visible = True
        Else
            Application.Run Rut_Name
        End If
        Me.Lb_Tit_Informe.Caption = "Informe Tarea: " & Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 1)
        Lo_Tareas.DataBodyRange.Cells(Index_Tarea, 5).Value = Me.TB_Informe.Text
    Loop While Len(Rutinas_Name) > 0
    
    Me.TB_Informe.BackColor = RGB(192, 255, 192)        '- VerdeClaro
    Me.Fondo_Inf.BackColor = RGB(192, 255, 192)         '- VerdeClaro
    Me.Lb_Tit_Informe.BackColor = RGB(192, 255, 192)    '- VerdeClaro
    Me.Fondo_Tit_Inf.BackColor = RGB(192, 255, 192)     '- VerdeClaro
    
    Application.ScreenUpdating = True
    DoEvents
    Exit Sub
Nombre_Rutina_NO_Encontrado:
    MsgBox "¡¡ Nombre de rutina NO encontrado !!", vbOKOnly, "Proceso de Selección de Rutina."
    On Error GoTo 0
End Sub     ' Btn_Ejec_Tarea_Click   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' --------------------------------------------------------------------------------------------------
Sub LBx_Tareas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Index_Tarea = Idx_Tarea
    Btn_Ejec_Tarea_Click
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' --------------------------------------------------------------------------------------------------
Sub LBx_Tareas_Click()
'    Idx_Tarea = Application.Match(Me.LBx_Tareas, Lo_Tareas.DataBodyRange.Columns(1), 0)
    Idx_Tarea = Application.Match(Me.LBx_Tareas, Lo_Tareas.DataBodyRange.Columns(1), 0)
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
    Me.TB_Informe.BackColor = RGB(255, 224, 192)        '- Marrón Claro
    Me.Fondo_Inf.BackColor = RGB(255, 224, 192)         '- Marrón Claro
    Me.Lb_Tit_Informe.BackColor = RGB(255, 224, 192)    '- Marrón Claro
    Me.Fondo_Tit_Inf.BackColor = RGB(255, 224, 192)     '- Marrón Claro
    Me.TBx_Descripción = Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 4)
    Me.TBx_Tarea_Name = Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 1)
    Me.TBx_Rutina_Name = Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 3)
    Me.Lb_Tit_Informe.Caption = "Resultado Última Tarea realizada :" & Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 1)
    Me.TB_Informe = Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 5)
        Call Rut_EnableEvents_Status_Reset
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' --------------------------------------------------------------------------------------------------
Sub TBx_Descripción_Change()
            Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
        Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 4) = Me.TBx_Descripción
        Call Mostrar_Tareas(Range("Usuario_ID"))
        Call Rut_EnableEvents_Status_Reset
End Sub
' --------------------------------------------------------------------------------------------------
Private Sub TBx_Tarea_Name_AfterUpdate()
            Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
        Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 1) = Me.TBx_Tarea_Name
        Call Rut_Lo_Sort(Lo_Tareas, 1, xlAscending, True)
'        Me.TBx_Tarea_Name = ""
        Call Mostrar_Tareas(Range("Usuario_ID"))
            Call Rut_EnableEvents_Status_Reset
End Sub
' --------------------------------------------------------------------------------------------------
Private Sub TBx_Rutina_Name_AfterUpdate()
            Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
        Lo_Tareas.DataBodyRange.Cells(Idx_Tarea, 3) = Me.TBx_Rutina_Name
'        Me.TBx_Rutina_Name = ""
        Call Mostrar_Tareas(Range("Usuario_ID"))
            Call Rut_EnableEvents_Status_Reset
End Sub




