VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_JyC_Lst 
   Caption         =   "Seleccionar una SIGLA de la Jornada o Congreso."
   ClientHeight    =   11450
   ClientLeft      =   580
   ClientTop       =   1030
   ClientWidth     =   22500
   OleObjectBlob   =   "Form_JyC_Lst.frx":0000
End
Attribute VB_Name = "Form_JyC_Lst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Lin_Lst     As Long

' ------------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()       ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ------------------------------------------------------------------------------------------------------
'    Me.Width = 600
'    Me.Height = 350
    Me.Zoom = 100
'    Me.Left = 50
'    Me.Top = 100
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ------------------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
' ------------------------------------------------------------------------------------------------------
    Dim Cont_Row        As Integer
    Dim Lo_Lst          As ListObject:      Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    ' Cargo la lista desplegable de Me.Estado -------------------
    With Lo_Lst.DataBodyRange
        For Cont_Row = 1 To Lo_Lst.ListRows.Count
            LBx_JyC_Lst.AddItem .Cells(Cont_Row, C_JCL_Siglas)
        Next Cont_Row
        LBx_JyC_Lst.Value = H_Liq_CTA.Range("Liq_Siglas")
    End With
End Sub
' ------------------------------------------------------------------------------------------------------

Private Sub Btn_Select_Click()
    If LBx_JyC_Lst.ListIndex <> -1 Then
        H_Liq_CTA.Range("Liq_Siglas").Value = Me.LBx_JyC_Lst
        Unload Me
    Else
        MsgBox "Please select an item!", vbExclamation
    End If
End Sub

Private Sub LBx_JyC_Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    H_Liq_CTA.Range("Liq_Siglas").Value = Trim(Left(Me.LBx_JyC_Lst, 25))
    Unload Me
End Sub

Private Sub LBx_JyC_Lst_Click()
    Lin_Lst = LBx_JyC_Lst.ListIndex + 1
    With Prog_JyC_List.ListObjects(1).DataBodyRange
        Tbx_Organiza = .Cells(Lin_Lst, C_JCL_Organiza)
        TBx_Nom = .Cells(Lin_Lst, C_JCL_Nom)
        TBx_Contact = .Cells(Lin_Lst, C_JCL_Contact)
        TBx_Email = .Cells(Lin_Lst, C_JCL_Email)
        TBx_Tlno = .Cells(Lin_Lst, C_JCL_Tlno)
        TBx_Obs = .Cells(Lin_Lst, C_JCL_Tipo)
'        TBx_Ingreso = .Cells(Lin_Lst, C_JCL_Ingreso)
'        TBx_Liq_Sol = .Cells(Lin_Lst, C_JCL_Liq_Sol)
'        TBx_Pdte_Liq = .Cells(Lin_Lst, C_JCL_Pdte_Liq)
    End With
End Sub

Private Sub Btn_Modif_Click()
    Lin_Lst = LBx_JyC_Lst.ListIndex + 1
    With Prog_JyC_List.ListObjects(1).DataBodyRange
        .Cells(Lin_Lst, C_JCL_Organiza) = Tbx_Organiza
        .Cells(Lin_Lst, C_JCL_Nom) = TBx_Nom
        .Cells(Lin_Lst, C_JCL_Contact) = TBx_Contact
        .Cells(Lin_Lst, C_JCL_Email) = TBx_Email
        .Cells(Lin_Lst, C_JCL_Tlno) = TBx_Tlno
        .Cells(Lin_Lst, C_JCL_Tipo) = TBx_Obs
'        .Cells(Lin_Lst, C_JCL_Ingreso) = TBx_Ingreso
'        .Cells(Lin_Lst, C_JCL_Liq_Sol) = TBx_Liq_Sol
'        .Cells(Lin_Lst, C_JCL_Pdte_Liq) = TBx_Pdte_Liq
    End With

End Sub

Private Sub Btn_Grabar_Click()
    With Prog_JyC_List.ListObjects(1)
        Lin_Lst = .ListRows.Count + 1
        .ListRows.Add
        With .DataBodyRange
        .Cells(Lin_Lst, C_JCL_Siglas) = TBx_Sigla
        .Cells(Lin_Lst, C_JCL_Organiza) = Tbx_Organiza
        .Cells(Lin_Lst, C_JCL_Nom) = TBx_Nom
        .Cells(Lin_Lst, C_JCL_Contact) = TBx_Contact
        .Cells(Lin_Lst, C_JCL_Email) = TBx_Email
        .Cells(Lin_Lst, C_JCL_Tlno) = TBx_Tlno
        .Cells(Lin_Lst, C_JCL_Tipo) = TBx_Obs
        End With
    End With
    Btn_Cancel_Click
    LBx_JyC_Lst.Clear
    UserForm_Activate
    LBx_JyC_Lst.Value = TBx_Sigla
    LBx_JyC_Lst.Enabled = True
End Sub

Private Sub Btn_Cancel_Click()
    Lb_Sigla.Visible = False
    TBx_Sigla.Visible = False
    Btn_Grabar.Visible = False
    Btn_Cancel.Visible = False
    LBx_JyC_Lst.Visible = True
    Lb_JyC_Lst.Visible = True
    Btn_Modif.Visible = True
    Btn_NewJyC.Visible = True
    Btn_Select.Visible = True
    LBx_JyC_Lst.Enabled = True
End Sub

Private Sub Btn_NewJyC_Click()
    LBx_JyC_Lst.Visible = False
    Lb_JyC_Lst.Visible = False
    Lb_IntroData.Visible = True
    Lb_Sigla.Visible = True
    TBx_Sigla.Visible = True
    Btn_Grabar.Visible = True
    Btn_Cancel.Visible = True
    Btn_Modif.Visible = False
    Btn_NewJyC.Visible = False
    Btn_Select.Visible = False
    Tbx_Organiza = ""
    TBx_Nom = ""
    TBx_Contact = ""
    TBx_Email = ""
    TBx_Tlno = ""
    TBx_Obs = ""
    LBx_JyC_Lst.Enabled = False
End Sub


