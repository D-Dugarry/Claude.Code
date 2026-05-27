VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_JyC_Select 
   Caption         =   "Seleccionar una Jornada o un Congreso."
   ClientHeight    =   11420
   ClientLeft      =   580
   ClientTop       =   1030
   ClientWidth     =   24980
   OleObjectBlob   =   "Form_JyC_Select.frx":0000
End
Attribute VB_Name = "Form_JyC_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    Dim Lo_INI          As ListObject:      Set Lo_INI = H_INICI.ListObjects(1)
    With Me.LBx_JyC_Lst
        .BackColor = RGB(230, 230, 230):    .ForeColor = RGB(30, 30, 90)
        .BoundColumn = 1        ' Número de la columna para devolver el valor, la primera columna es 0
        .ColumnCount = 3        ' Número de columnas del ComboBox
        .ColumnWidths = "170;40;60"
    End With    ' Me.ComboBox_Líneas
    ' Cargo la lista desplegable de Me.Estado -------------------
        '''    Me.LBx_JyC_Lst.List = Lo_INI.ListColumns(1).DataBodyRange.Resize(, 2).Value
    Rut_ListBox_Load
End Sub
' ------------------------------------------------------------------------------------------------------
'' Esta versión de la rutina, NO SIRVE porque la Lo_INICI NO separa las Liquid y sus DEV con RDT
''Sub Rut_ListBox_Load()
''    Dim LinX        As Long
''    Dim RowINI          As ListRow
''    Dim Lo_INI      As ListObject:      Set Lo_INI = H_INICI.ListObjects(1)
''    LBx_JyC_Lst.Clear
''    Application.ScreenUpdating = False
''    Rut_Inicio_LstObj_Refresh
''
''    For LinX = 1 To Lo_INI.ListRows.Count
''        Set RowINI = Lo_INI.ListRows(LinX)
''        LBx_JyC_Lst.AddItem
''        LBx_JyC_Lst.List(LinX - 1, 0) = RowINI.Range(C_INI_Siglas)
''        LBx_JyC_Lst.List(LinX - 1, 1) = RowINI.Range(C_INI_CtaTPV)
''        LBx_JyC_Lst.List(LinX - 1, 2) = RowINI.Range(C_INI_N_Liq)
''    Next LinX
''
''    Application.ScreenUpdating = True
''End Sub     '- Rut_ListBox_Load ----------
' ------------------------------------------------------------------------------------------------------
Sub Rut_ListBox_Load()
    Dim LinX        As Long
    Dim RefAnt      As String:      RefAnt = ""
    Dim RefLin      As Integer
    Dim TPV_Visibl  As Boolean:     TPV_Visibl = Prog_TPV_Tb.Visible
    Dim Cta_Visibl  As Boolean:     Cta_Visibl = Prog_CTA_Tb.Visible
    Dim RowTPV          As ListRow
    Dim RowCta          As ListRow
    Dim Lo_TPV      As ListObject:      Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Dim Lo_Cta      As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    LBx_JyC_Lst.Clear
    Application.ScreenUpdating = False
    '- Genero la lista de TPV ---------------
    Prog_TPV_Tb.Visible = xlSheetHidden
    Call Rut_LstObj_Filtros_Quitar(Lo_TPV)
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_N_Liq, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_Siglas, xlAscending, True)
    For LinX = 1 To Lo_TPV.ListRows.Count
        Set RowTPV = Lo_TPV.ListRows(LinX)
        If RowTPV.Range(C_TPV_Siglas) = "" Then GoTo SigLinX
        If Right(RowTPV.Range(C_TPV_F_Pago), 4) <> Prog__APP.Range("App_AñoCont") Then GoTo SigLinX
        If RefAnt <> RowTPV.Range(C_TPV_Siglas) & RowTPV.Range(C_TPV_N_Liq) Then
            RefAnt = RowTPV.Range(C_TPV_Siglas) & RowTPV.Range(C_TPV_N_Liq)
            LBx_JyC_Lst.AddItem
            LBx_JyC_Lst.List(RefLin, 0) = RowTPV.Range(C_TPV_Siglas)
            LBx_JyC_Lst.List(RefLin, 1) = "TPV"
            LBx_JyC_Lst.List(RefLin, 2) = RowTPV.Range(C_TPV_N_Liq)
            RefLin = RefLin + 1
        End If
SigLinX:
    Next LinX
    Prog_TPV_Tb.Visible = TPV_Visibl
    '- Genero la lista de Cta ---------------
    Prog_CTA_Tb.Visible = xlSheetHidden
    Call Rut_LstObj_Filtros_Quitar(Lo_Cta)
    Call Rut_LstObj_Sort(Lo_Cta, C_Cta_N_Liq, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_Cta, C_Cta_Siglas, xlAscending, True)
    For LinX = 1 To Lo_Cta.ListRows.Count
        Set RowCta = Lo_Cta.ListRows(LinX)
        If RowCta.Range(C_Cta_N_Liq) = "" Then GoTo SigLinX2
    '    If UCase(RowCta.Range(C_Cta_N_Liq)) = "X" Then GoTo SigLinX2
        If RowCta.Range(C_Cta_N_Liq) = "--" Then GoTo SigLinX2
        If RowCta.Range(C_Cta_Siglas) = "" Then GoTo SigLinX2
        If Right(RowCta.Range(C_Cta_F_VAL), 4) <> Prog__APP.Range("App_AñoCont") Then GoTo SigLinX2
        If RefAnt <> RowCta.Range(C_Cta_Siglas) & RowCta.Range(C_Cta_N_Liq) Then
            RefAnt = RowCta.Range(C_Cta_Siglas) & RowCta.Range(C_Cta_N_Liq)
            LBx_JyC_Lst.AddItem
            LBx_JyC_Lst.List(RefLin, 0) = RowCta.Range(C_Cta_Siglas)
            LBx_JyC_Lst.List(RefLin, 1) = "Cta"
            LBx_JyC_Lst.List(RefLin, 2) = RowCta.Range(C_Cta_N_Liq)
            RefLin = RefLin + 1
        End If
SigLinX2:
    Next LinX
    Prog_CTA_Tb.Visible = Prog_CTA_Tb.Visible
    Call SortListbox(LBx_JyC_Lst, 1)
    Call SortListbox(LBx_JyC_Lst, 0)
    Application.ScreenUpdating = True
End Sub     '- Rut_ListBox_Load ----------
'==============================================================================
Sub SortListbox(LBX As MSForms.ListBox, col As Integer)
    Dim i As Long
    Dim j As Long
    Dim Temp As Variant
    Dim Y  As Integer
    With LBX
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
    End With
End Sub
'==============================================================================
Private Sub Btn_Select_Click()
    If LBx_JyC_Lst.ListIndex <> -1 Then
        H_Liq_CTA.Range("Liq_Siglas").Value = LBx_JyC_Lst(LBx_JyC_Lst.ListIndex, 0)
        H_Liq_CTA.Range("Liq_Núm").Value = LBx_JyC_Lst(LBx_JyC_Lst.ListIndex, 2)
        Unload Me
    Else
        MsgBox "Please select an item!", vbExclamation
    End If
End Sub

Private Sub LBx_JyC_Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        H_Liq_CTA.Range("Liq_Siglas").Value = LBx_JyC_Lst(LBx_JyC_Lst.ListIndex, 0)
        H_Liq_CTA.Range("Liq_Núm").Value = LBx_JyC_Lst(LBx_JyC_Lst.ListIndex, 2)
    Unload Me
End Sub
          
Private Sub LBx_JyC_Lst_Click()
    Dim Lin_Lo_JyC_List     As Variant
    Dim Lin_LBx_JyC_Lst     As Integer:     Lin_LBx_JyC_Lst = LBx_JyC_Lst.ListIndex
    Dim Lin_Lo_INI          As Integer:     Lin_Lo_INI = Lin_LBx_JyC_Lst + 1
'    Dim Lin_Ini_INICIO      As Integer:     Lin_Ini_INICIO = 1
    Debug.Print LBx_JyC_Lst.List(Lin_LBx_JyC_Lst, 0), "Liq. " & LBx_JyC_Lst.List(Lin_LBx_JyC_Lst, 2)
    Lin_Lo_JyC_List = Application.Match(LBx_JyC_Lst.List(Lin_LBx_JyC_Lst, 0), Prog_JyC_List.ListObjects(1).DataBodyRange.Columns(1), 0)
    Debug.Print "Línea Lo_: " & Lin_Lo_JyC_List, "Línea LBx_: " & Lin_LBx_JyC_Lst
    With Prog_JyC_List.ListObjects(1).DataBodyRange
        Tbx_Organiza = .Cells(Lin_Lo_JyC_List, C_JCL_Organiza)
        Tbx_Orgánica = .Cells(Lin_Lo_JyC_List, C_JCL_Orgánica)
        TBx_Nom = .Cells(Lin_Lo_JyC_List, C_JCL_Nom)
        Tbx_F_Ini = .Cells(Lin_Lo_JyC_List, C_JCL_F_Ini)
        Tbx_F_Fin = .Cells(Lin_Lo_JyC_List, C_JCL_F_Fin)
        TBx_Contact = .Cells(Lin_Lo_JyC_List, C_JCL_Contact)
        TBx_Email = .Cells(Lin_Lo_JyC_List, C_JCL_Email)
        TBx_Tlno = .Cells(Lin_Lo_JyC_List, C_JCL_Tlno)
        TBx_Obs = .Cells(Lin_Lo_JyC_List, C_JCL_Tipo)
    End With
    With H_INICI.ListObjects(1).DataBodyRange
'        Do While Left(.Cells(Lin_LBx_JyC_Lst, C_INI_TPV_RecEmi), 1) = "_"
'            Lin_Ini_INICIO = Lin_Ini_INICIO + 1
'        Loop
        If .Cells(Lin_Lo_INI, C_INI_CtaTPV) = "TPV" Then
            TBx_FormPago = "Pago por TPV o Bizum."
            TBx_ImpEmi.Visible = True
            TBx_ImpEmi = .Cells(Lin_Lo_INI, C_INI_TPV_RecEmi)
            TBx_ImpCob.Visible = True
            TBx_ImpCob = .Cells(Lin_Lo_INI, C_INI_TPV_ImpRec)
            TBx_ComBco.Visible = True
            TBx_ComBco = .Cells(Lin_Lo_INI, C_INI_TPV_ComBco)
            TBx_ImpNeto.Visible = True
            TBx_ImpNeto = .Cells(Lin_Lo_INI, C_INI_TPV_Neto)
            TBx_ImpDEV.Visible = True
            TBx_ImpDEV = .Cells(Lin_Lo_INI, C_INI_TPV_Dev)
        Else
            TBx_FormPago = "Transf. a Cta.Bco."
            TBx_ImpEmi.Visible = True
            TBx_ImpEmi = .Cells(Lin_Lo_INI, C_INI_Cta_ImpRec)
            TBx_ImpCob.Visible = True
            TBx_ImpCob = .Cells(Lin_Lo_INI, C_INI_Cta_Imp)
            TBx_ComBco.Visible = False
            TBx_ImpNeto.Visible = False
            TBx_ImpDEV.Visible = True
            TBx_ImpDEV = .Cells(Lin_Lo_INI, C_INI_Cta_Dev)
        End If
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
        .Cells(Lin_Lst, C_JCL_Orgánica) = Tbx_Orgánica
        .Cells(Lin_Lst, C_JCL_Nom) = TBx_Nom
        .Cells(Lin_Lst, C_JCL_F_Ini) = Tbx_F_Ini
        .Cells(Lin_Lst, C_JCL_F_Fin) = Tbx_F_Fin
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
    Lb_JyC_Lst.Visible = True
    LBx_JyC_Lst.Visible = True
    Lb_Sigla.Visible = False
    TBx_Sigla.Visible = False
    Btn_Grabar.Visible = False
    Btn_Cancel.Visible = False
    Btn_Modif.Visible = True
    Btn_NewJyC.Visible = True
    Btn_Select.Visible = True
    Lb_Ingreso.Visible = True
    Lb_Liq_Sol.Visible = True
    LBx_JyC_Lst.Enabled = True
End Sub

Private Sub Btn_NewJyC_Click()
    Lb_JyC_Lst.Visible = False
    LBx_JyC_Lst.Visible = False
    Lb_Sigla.Visible = True
    TBx_Sigla.Visible = True
    Btn_Grabar.Visible = True
    Btn_Cancel.Visible = True
    Btn_Modif.Visible = False
    Btn_NewJyC.Visible = False
    Btn_Select.Visible = False
    Tbx_Organiza = ""
    Tbx_Orgánica = ""
    TBx_Nom = ""
    Tbx_F_Ini = ""
    Tbx_F_Fin = ""
    TBx_Contact = ""
    TBx_Email = ""
    TBx_Tlno = ""
    TBx_Obs = ""
    Lb_Ingreso.Visible = False
    Lb_Liq_Sol.Visible = False
    LBx_JyC_Lst.Enabled = False
    TBx_Sigla.SetFocus
End Sub


