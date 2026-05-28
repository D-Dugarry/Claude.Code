VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_JyC_TPV_Refs 
   Caption         =   "UserForm1"
   ClientHeight    =   11420
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   22790
   OleObjectBlob   =   "Form_JyC_TPV_Refs.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Form_JyC_TPV_Refs"
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
    Dim EventName       As String
    Dim NewEventName    As String
    Dim Lo_TPV          As ListObject:      Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Prog_TPV_Tb.Protect allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
    Call Rut_LstObj_WrkSht_Preparar(Application.Workbooks(ThisWorkbook.Name).Sheets(Prog_TPV_Tb.Name))
    Call Rut_LstObj_Sort(Lo_TPV, C_Pag_EventoNom, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_TPV, C_Pag_EventoRef, xlAscending, False)
    
    ' Cargo la lista desplegable de Me.Estado -------------------
    With Lo_TPV.DataBodyRange
        For Cont_Row = 1 To Lo_TPV.ListRows.Count
            NewEventName = .Cells(Cont_Row, C_Pag_EventoRef) & " - " & .Cells(Cont_Row, C_Pag_EventoNom)
            If EventName <> NewEventName Then
                EventName = NewEventName
                LBx_Ref_Event.AddItem NewEventName
            End If
        Next Cont_Row
'        LBx_Ref_Event.Value = H_Liq_CTA.Range("Liq_Siglas")
    End With
    
    Call SortListbox(LBx_Ref_Event, 0, False)
    Application.ScreenUpdating = True

'    Dim Lo_Lst          As ListObject:      Set Lo_Lst = Prog_JyC_List.ListObjects(1)
'    ' Cargo la lista desplegable de Me.Estado -------------------
'    With Lo_Lst.DataBodyRange
'        For Cont_Row = 1 To Lo_Lst.ListRows.Count
'            LBx_Ref_Event.AddItem .Cells(Cont_Row, C_JCL_Siglas)
'        Next Cont_Row
'        LBx_Ref_Event.Value = H_Liq_CTA.Range("Liq_Siglas")
'    End With
End Sub
' ------------------------------------------------------------------------------------------------------
'==============================================================================
Sub SortListbox(LBX As MSForms.ListBox, col As Integer, Optional Asc As Boolean = True)
    Dim i As Long
    Dim j As Long
    Dim Temp As Variant
    Dim Y  As Integer
    With LBX
        If Asc Then
            For i = 0 To .ListCount - 2
                For j = i + 1 To .ListCount - 1
                    If .List(i, col) >= .List(j, col) Then
                        For Y = 0 To .ColumnCount - 1
                        Temp = .List(j, Y)
                        .List(j, Y) = .List(i, Y)
                        .List(i, Y) = Temp
                        Next Y
                    End If
                Next j
            Next i
        Else
            For i = 0 To .ListCount - 2
                For j = i + 1 To .ListCount - 1
                    If .List(i, col) <= .List(j, col) Then
                        For Y = 0 To .ColumnCount - 1
                        Temp = .List(j, Y)
                        .List(j, Y) = .List(i, Y)
                        .List(i, Y) = Temp
                        Next Y
                    End If
                Next j
            Next i
        End If
    End With
End Sub

Private Sub LBx_Ref_Event_Click()
    Dim Cont_Row        As Integer
    Dim RefEvent        As String:          RefEvent = Left(LBx_Ref_Event, 4)
    Dim Lo_Lst          As ListObject:      Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    '- Limpio los campos ---------------------------------------
    Tbx_Organiza = ""
    TBx_Nom = ""
    TBx_Contact = ""
    TBx_Email = ""
    TBx_Tlno = ""
    TBx_Obs = ""
    TBx_EventRefs = ""
    ' Visualizo datos si hay coincidencia -------------------
    With Lo_Lst.DataBodyRange
        For Cont_Row = 1 To Lo_Lst.ListRows.Count
            If InStr(.Cells(Cont_Row, C_JCL_TpvRefs), RefEvent) > 0 Then
                Lin_Lst = Cont_Row
                Tbx_Organiza = .Cells(Lin_Lst, C_JCL_Organiza)
                TBx_Nom = .Cells(Lin_Lst, C_JCL_Nom)
                TBx_Contact = .Cells(Lin_Lst, C_JCL_Contact)
                TBx_Email = .Cells(Lin_Lst, C_JCL_Email)
                TBx_Tlno = .Cells(Lin_Lst, C_JCL_Tlno)
                TBx_Obs = .Cells(Lin_Lst, C_JCL_Tipo)
                TBx_EventRefs = .Cells(Lin_Lst, C_JCL_TpvRefs)
                Exit For
            End If

'            LBx_Ref_Event.AddItem .Cells(Cont_Row, C_JCL_TpvRefs)
        Next Cont_Row
        If TBx_EventRefs = "" Then TBx_EventRefs = "- Sin Asignar -"
'        LBx_Ref_Event.Value = H_Liq_CTA.Range("Liq_Siglas")
    End With

End Sub

Private Sub Btn_Select_Click()
    If LBx_Ref_Event.ListIndex <> -1 Then
        H_Liq_CTA.Range("Liq_Siglas").Value = Me.LBx_Ref_Event
        Unload Me
    Else
        MsgBox "Please select an item!", vbExclamation
    End If
End Sub

Private Sub LBx_Ref_Event_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    H_Liq_CTA.Range("Liq_Siglas").Value = Trim(Left(Me.LBx_Ref_Event, 25))
    Unload Me
End Sub

Private Sub Btn_Modif_Click()
    Lin_Lst = LBx_Ref_Event.ListIndex + 1
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
    LBx_Ref_Event.Clear
    UserForm_Activate
    LBx_Ref_Event.Value = TBx_Sigla
    LBx_Ref_Event.Enabled = True
End Sub

Private Sub Btn_Cancel_Click()
    Lb_Sigla.Visible = False
    TBx_Sigla.Visible = False
    Btn_Grabar.Visible = False
    Btn_Cancel.Visible = False
    LBx_Ref_Event.Visible = True
    Lb_JyC_Lst.Visible = True
    Btn_Modif.Visible = True
    Btn_NewJyC.Visible = True
    Btn_Select.Visible = True
    LBx_Ref_Event.Enabled = True
End Sub

Private Sub Btn_NewJyC_Click()
    LBx_Ref_Event.Visible = False
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
    LBx_Ref_Event.Enabled = False
End Sub



