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
Option Explicit
' ------------------------------------------------------------------------------------------------------
Private Sub Btn_Eixir_Click()
            Call Rut_EnableEvents_Status_Reset                      ' INHABILITA LOS EVENTOS
    Unload Me
End Sub
' ------------------------------------------------------------------------------------------------------
 Sub UserForm_Initialize()       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    With Application
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
 Sub UserForm_Activate()
            Application.EnableEvents = False                      ' INHABILITA LOS EVENTOS
    Me.Tbx_UserName = ""
    Me.Tbx_UserMail = ""
    Me.Tbx_Unidad_NEXE = ""
End Sub     ' UserForm_Activate    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------
Private Sub TBx_Usuario_AfterUpdate()
    If TBx_Usuario.Value = "" Then
        Range("Usuario_ID") = ""
        Range("Usuario_Name") = ""
        Range("Usuario_Ext") = ""
        Range("App_LetraUnidRed") = ""
        Range("App_MailUsu") = ""
        Range("App_RutaAPP") = ""
        Prog__APP_Switch.Range("Sw_Boss") = False
        Prog__APP_Switch.Range("Sw_Probando") = False
        Me.Tbx_UserName = "Usuario sin Activar"
        Me.Tbx_UserMail = ""
        Me.Tbx_UserExt = ""
        Me.Tbx_Unidad_NEXE = ""
        Form_Menu.Tbx_UserName = "Usuario sin Activar"
        Form_Menu.Tbx_UserExt = ""
        Form_Menu.Tbx_Unidad_NEXE = ""
        Form_Menu.Lb_Tarea_Name.Visible = False
        Form_Menu.TBx_Tarea_Name.Visible = False
        Form_Menu.Lb_Rutina_Name.Visible = False
        Form_Menu.TBx_Rutina_Name.Visible = False
        Form_Menu.TB_Informe = "Sin Usuario. " & vbCrLf & "Usuario genérico." & vbCrLf & Now()
        
        If Left(ActiveWorkbook.Path, 18) = "https://nexe.ua.es" Then Me.Lbl_Obl_ID.Visible = True
'        Exit Sub
    Else
        Range("Usuario_ID") = TBx_Usuario.Value
        Range("Usuario_Name") = ""
        With Prog__Usuarios.ListObjects(1).DataBodyRange
            Index_Usuario = Application.Match(Range("Usuario_ID"), .Columns(1), 0) '------
            If IsError(Index_Usuario) Then GoTo Volver_a_Preguntar                  '- NO hay coincidencia ---------------
            If Range("Usuario_ID") <> .Cells(Index_Usuario, 1) Then GoTo Volver_a_Preguntar  '- NO hay coincidencia Mays/Mins -----
            '- La coincidencia es TOTAL distingue mayúsculas de minúsculas -----------------------------------------------
            Range("Usuario_ID") = .Cells(Index_Usuario, 1)
            Range("Usuario_Name") = .Cells(Index_Usuario, 2)
            Range("Usuario_Ext") = .Cells(Index_Usuario, 7)
            Me.Tbx_UserName = Range("Usuario_Name")
            Me.Tbx_UserExt = Range("Usuario_Ext")
            Me.Tbx_UserMail = .Cells(Index_Usuario, 6)
            Form_Menu.Tbx_UserName = Range("Usuario_Name")
            Form_Menu.Tbx_UserExt = Range("Usuario_Ext")
            If CheckBox_Teletrabajo.Value = True Then
                Range("App_LetraUnidRed") = .Cells(Index_Usuario, 5)
            Else
                Range("App_LetraUnidRed") = .Cells(Index_Usuario, 4)
            End If
            Range("App_MailUsu") = .Cells(Index_Usuario, 6)
                Range("App_RutaAPP") = ActiveWorkbook.Path
                Range("App_RutaAPP") = Replace(Range("App_RutaAPP"), "/", "\")
                Range("App_RutaAPP") = Mid(Range("App_RutaAPP"), InStr(1, Range("App_RutaAPP"), Range("App_MailUsu")) + Len(Range("App_MailUsu")))
            Range("App_RutaAPP") = Range("App_LetraUnidRed") & ":" & Range("App_RutaAPP") & "\"
            Me.Tbx_Unidad_NEXE = Range("App_LetraUnidRed") '& ":"
            Form_Menu.Tbx_Unidad_NEXE = Range("App_LetraUnidRed") & ":"
                        If Range("Usuario_ID") = "Boss" Then
                            Prog__APP_Switch.Range("Sw_Boss") = True
                            Form_Menu.Lb_Tarea_Name.Visible = True
                            Form_Menu.TBx_Tarea_Name.Visible = True
                            Form_Menu.Lb_Rutina_Name.Visible = True
                            Form_Menu.TBx_Rutina_Name.Visible = True
                        Else
                            Prog__APP_Switch.Range("Sw_Boss") = False
                            Form_Menu.Lb_Tarea_Name.Visible = False
                            Form_Menu.TBx_Tarea_Name.Visible = False
                            Form_Menu.Lb_Rutina_Name.Visible = False
                            Form_Menu.TBx_Rutina_Name.Visible = False
                        End If
            Exit Sub
        End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
    End If
Volver_a_Preguntar:
    '- NO hay coincidencia ---------------
    Range("Usuario_ID") = ""
    Me.TBx_Usuario = ""
    MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "Usuario Desconocido", _
                                vbOKOnly, "Proceso: Identificación Usuario"
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserName_AfterUpdate()
    Prog__Usuarios.ListObjects(1).DataBodyRange.Cells(Index_Usuario, 2) = Me.Tbx_UserName
    Range("Usuario_Name") = Me.Tbx_UserName
Debug.Print "Tbx_UserName_AfterUpdate: " & Range("Usuario_Name")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserExt_AfterUpdate()
    Prog__Usuarios.ListObjects(1).DataBodyRange.Cells(Index_Usuario, 7) = Me.Tbx_UserExt
    Range("Usuario_Ext") = Me.Tbx_UserExt
Debug.Print "Tbx_UserName_AfterUpdate: " & Range("Usuario_Ext")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_UserMail_AfterUpdate()
    With Prog__Usuarios.ListObjects(1).DataBodyRange
        Range("App_MailUsu") = Me.Tbx_UserMail
        .Cells(Index_Usuario, 6) = Range("App_MailUsu")
            Range("App_RutaAPP") = ActiveWorkbook.Path
            Range("App_RutaAPP") = Replace(Range("App_RutaAPP"), "/", "\")
            Range("App_RutaAPP") = Mid(Range("App_RutaAPP"), InStr(1, Range("App_RutaAPP"), Range("App_MailUsu")) + Len(Range("App_MailUsu")))
        Range("App_RutaAPP") = Range("App_LetraUnidRed") & ":" & Range("App_RutaAPP") & "\"
    End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
Debug.Print "Tbx_UserMail_AfterUpdate: " & Range("App_RutaAPP")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub Tbx_Unidad_NEXE_AfterUpdate()
    With Prog__Usuarios.ListObjects(1).DataBodyRange
        Range("App_LetraUnidRed") = Me.Tbx_Unidad_NEXE
        If CheckBox_Teletrabajo.Value = True Then
            .Cells(Index_Usuario, 5) = Range("App_LetraUnidRed")
        Else
            .Cells(Index_Usuario, 4) = Range("App_LetraUnidRed")
        End If
            Range("App_RutaAPP") = ActiveWorkbook.Path
            Range("App_RutaAPP") = Replace(Range("App_RutaAPP"), "/", "\")
            Range("App_RutaAPP") = Mid(Range("App_RutaAPP"), InStr(1, Range("App_RutaAPP"), Range("App_MailUsu")) + Len(Range("App_MailUsu")))
        Range("App_RutaAPP") = Range("App_LetraUnidRed") & ":" & Range("App_RutaAPP") & "\"
    End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
Debug.Print "CheckBox_Teletrabajo_AfterUpdate: " & Range("App_RutaAPP")
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub CheckBox_Teletrabajo_AfterUpdate()
    With Prog__Usuarios.ListObjects(1).DataBodyRange
        If CheckBox_Teletrabajo.Value = True Then
            Range("App_LetraUnidRed") = .Cells(Index_Usuario, 5)
        Else
            Range("App_LetraUnidRed") = .Cells(Index_Usuario, 4)
        End If
            Range("App_RutaAPP") = ActiveWorkbook.Path
            Range("App_RutaAPP") = Replace(Range("App_RutaAPP"), "/", "\")
            Range("App_RutaAPP") = Mid(Range("App_RutaAPP"), InStr(1, Range("App_RutaAPP"), Range("App_MailUsu")) + Len(Range("App_MailUsu")))
        Range("App_RutaAPP") = Range("App_LetraUnidRed") & ":" & Range("App_RutaAPP") & "\"
    End With    ' Prog__Usuarios.ListObjects(1).DataBodyRange
    Me.Tbx_Unidad_NEXE = Range("App_LetraUnidRed")
Debug.Print "Tbx_Unidad_NEXE_AfterUpdate: " & Range("App_RutaAPP")
End Sub

