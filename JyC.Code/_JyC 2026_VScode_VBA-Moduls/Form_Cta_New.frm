VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Cta_New 
   Caption         =   "Seleccionar una Jornada o un Congreso."
   ClientHeight    =   10460
   ClientLeft      =   580
   ClientTop       =   1030
   ClientWidth     =   22500
   OleObjectBlob   =   "Form_Cta_New.frx":0000
End
Attribute VB_Name = "Form_Cta_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Lin_Lst     As Long
Dim Lin_Reg     As Long

' ------------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
' ------------------------------------------------------------------------------------------------------
    With Application
        Zoom = Int(.Width / Me.Width * 100)
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
End Sub
' ------------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
' ------------------------------------------------------------------------------------------------------
    Dim Lo_Cta          As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
'    Dim Lo_CtaFind      As ListObject:      Set Lo_CtaFind = Prog_CTA_Find.ListObjects(1)
    Dim Lo_Lst          As ListObject:      Set Lo_Lst = Prog_JyC_List.ListObjects(1)
    Dim Cont_Row        As Long:            Cont_Row = Lo_Cta.ListRows.Count
    Dim SiglasLst       As Range
    Dim SiglasUnic       As Range
    Dim RngT            As Range
    Dim Rng1            As Range
    Dim Rng2            As Range
    
    With Me.LBx_JyC_Siglas
        .BackColor = RGB(230, 230, 230):    .ForeColor = RGB(30, 30, 90)
        .BoundColumn = 1        ' Número de la columna para devolver el valor, la primera columna es 0
        .ColumnCount = 1        ' Número de columnas del ComboBox
        .ColumnWidths = "150"
    End With
    LBx_JyC_Siglas.Clear
    Call Rut_LstObj_Filtro(Lo_Lst, C_JCL_Siglas, "<>_*", True)
    LBx_JyC_Siglas.List = Lo_Lst.DataBodyRange.SpecialCells(xlCellTypeVisible).Value

    With Me.LBx_Cta_Lst
        .BackColor = RGB(230, 230, 230):    .ForeColor = RGB(30, 30, 90)
        .BoundColumn = 1        ' Número de la columna para devolver el valor, la primera columna es 0
        .ColumnCount = 28        ' Número de columnas del ComboBox
'        .ColumnWidths = "50;0;0;50;50;0;0;0;250;250;250;250;50;50;20;50;0;0;0;0;0;0;0;0;0;0;0;0"
        .ColumnHeads = True
        .ColumnWidths = "60;0;0;50;30;0;0;0;0;0;0;0;0;0;20;80;0;0;0;0;0;0;0;0;0;0;0;850"
    End With    ' Me.ComboBox_Líneas
    
    LBx_Cta_Lst.Clear
    Call Rut_LstObj_Columns_Show(Sheets(Prog_CTA_Tb.Name), Sheets("DefCol_" & Prog_CTA_Tb.Name))
    If Not Lo_Cta.AutoFilter Is Nothing Then Lo_Cta.AutoFilter.ShowAllData
'    Lo_Cta.Range.AutoFilter Field:=C_Cta_Ordinal, Criteria1:=">2024000000"
'    Lo_Cta.Range.AutoFilter Field:=C_Cta_Siglas, Criteria1:="_Desconocido", Operator:=xlOr, Criteria2:="CRUE_Gerencias"
    Lo_Cta.Range.AutoFilter Field:=C_Cta_Siglas, Criteria1:="_Desconocido"
    Me.LBx_Cta_Lst.List = Lo_Cta.DataBodyRange.SpecialCells(xlCellTypeVisible).Value
    
    Debug.Print Lo_Cta.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
    
'    Call Rut_LstObj_DataBodyRange_Filtered_Copy(Lo_Cta, Lo_CtaFind, True)
'    Prog_CTA_Find.Select
'    Lo_CtaFind.ListColumns(C_Cta_Reg_Mov1).DataBodyRange.Resize(, 5).Select    '- Resize(,5) es para ampliar a 5 columnas desde "C_N43_Reg_Mov1"
'    With Selection
'        .Replace What:="Ordenante:", Replacement:="Ord. ", _
'                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
'                      SearchFormat:=False, ReplaceFormat:=False
'    End With
'
'    Me.LBx_Cta_Lst.List = Lo_CtaFind.DataBodyRange.SpecialCells(xlCellTypeVisible).Value





'    Lo_Cta.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=Rng1
'
'    With Rng1
'        .Replace What:="Observaciones:", Replacement:="", _
'                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
'                      SearchFormat:=False, ReplaceFormat:=False
'    End With
'    Me.LBx_Cta_Lst.List = Rng1.Value
'

'    Set SiglasLst = Lo_Lst.DataBodyRange.SpecialCells(xlCellTypeVisible)
'    Set SiglasLst = Lo_Lst.ListColumns(C_JCL_Siglas).DataBodyRange  ' Assuming unique values are in the first column
    
'    LBx_JyC_Siglas.List = WorksheetFunction.Unique(Lo_Cta.ListColumns(C_Cta_Siglas).DataBodyRange)
    
    ' Cargo la lista desplegable de Me.Estado -------------------
        '''    Me.LBx_Cta_Lst.List = Lo_INI.ListColumns(1).DataBodyRange.Resize(, 2).Value
'    Rut_ListBox_Load
    
'    Call Rut_LstObj_Filtro(Lo_TPV, C_TPV_Siglas, RowINI.Range(C_INI_Siglas), True)
'    Call Rut_LstObj_Filtro(Lo_TPV, C_TPV_N_Liq, "=" & RowINI.Range(C_INI_N_Liq) & "*", False)
'    Set VisRng = Lo_TPV.DataBodyRange.SpecialCells(xlCellTypeVisible)
    
'    With Lo_Cta.DataBodyRange
'        '- Localizo el último reg. reconocido ----------
'        Do While .Cells(1, 1) = "_Desconocido"
'            Cont_Row = Cont_Row - 1
'        Loop
'        Cont_Row = Cont_Row + 1
'        ' Cargo la lista desplegable de Me.Estado -------------------
'        For Lin_Reg = Cont_Row To Lo_Cta.ListRows.Count
'            LBx_Cta_Lst.AddItem .Cells(Cont_Row, C_JCL_Siglas)
'        Next Cont_Row
'        LBx_Cta_Lst.Value = H_Liq_CTA.Range("Liq_Siglas")
'    End With
End Sub
' ------------------------------------------------------------------------------------------------------

Private Sub Btn_Select_Click()
    If LBx_Cta_Lst.ListIndex <> -1 Then
        H_Liq_CTA.Range("Liq_Siglas").Value = Me.LBx_Cta_Lst
        Unload Me
    Else
        MsgBox "Please select an item!", vbExclamation
    End If
End Sub

Private Sub LBx_Cta_Lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    H_Liq_CTA.Range("Liq_Siglas").Value = Trim(Left(Me.LBx_Cta_Lst, 25))
    Unload Me
End Sub

Private Sub LBx_Cta_Lst_Click()
    Lin_Lst = LBx_Cta_Lst.ListIndex + 1
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
    Lin_Lst = LBx_Cta_Lst.ListIndex + 1
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
    LBx_Cta_Lst.Clear
    UserForm_Activate
    LBx_Cta_Lst.Value = TBx_Sigla
    LBx_Cta_Lst.Enabled = True
End Sub

Private Sub Btn_Cancel_Click()
    Lb_Sigla.Visible = False
    TBx_Sigla.Visible = False
    Btn_Grabar.Visible = False
    Btn_Cancel.Visible = False
    Btn_Modif.Visible = True
    Btn_NewJyC.Visible = True
    Btn_Select.Visible = True
    TBx_Ingreso.Visible = True
    TBx_Liq_Sol.Visible = True
    TBx_Pdte_Liq.Visible = True
    Lb_Ingreso.Visible = True
    Lb_Liq_Sol.Visible = True
    Lb_Pdte_Liq.Visible = True
    LBx_Cta_Lst.Enabled = True
End Sub

Private Sub Btn_NewJyC_Click()
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
    TBx_Ingreso.Visible = False
    TBx_Liq_Sol.Visible = False
    TBx_Pdte_Liq.Visible = False
    Lb_Ingreso.Visible = False
    Lb_Liq_Sol.Visible = False
    Lb_Pdte_Liq.Visible = False
    LBx_Cta_Lst.Enabled = False
End Sub


