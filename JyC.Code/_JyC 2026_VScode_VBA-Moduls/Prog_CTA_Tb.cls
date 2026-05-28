VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Prog_CTA_Tb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Call Rut_Filtrar_Tareas(Prog__APP.Range("APP_User_ID"))
    Call Rut_Context_Buttons_Hide
    RefreshRibbon
    Call Rut_Filter_AñoCont(Prog__APP.Range("APP_AñoCont"))
Debug.Print "Worksheet_Activate"
End Sub

Private Sub Worksheet_Deactivate()
    Application.CommandBars("Cell").Reset
    Application.CommandBars("List Range Popup").Reset
Debug.Print "Worksheet_Deactivate"
End Sub

'Private Sub Worksheet_Change(ByVal Target As Range)
'    ActiveWindow.Zoom = 100
'End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub Worksheet_Change(ByVal Target As Range)     '>>>>>>>>>>>>>>>>>>>>>>>> SI CAMBIA EL CONTENIDO DE LA CELDA SE ACTIVA (aunque acabe teniendo el mismo dato)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

' ============ Filtro en Columna Reg_Mov3...  ==========================================================================
    If Not Application.Intersect(Target, Range("i2")) Is Nothing Then
        Dim NumFilas As Long
        Dim Tb_JyC As Object
        Dim Rng As Range
        Set Tb_JyC = ThisWorkbook.ActiveSheet.ListObjects(1)

        Tb_JyC.Range.AutoFilter Field:=C_Cta_Ref1
        Tb_JyC.Range.AutoFilter Field:=C_Cta_Reg_Mov1        ' NO PONGO: ".Range.AutoFilter.ShowAllData"
        Tb_JyC.Range.AutoFilter Field:=C_Cta_Reg_Mov2       ' Porque quiero mantener filtros en otras columnas.
        Tb_JyC.Range.AutoFilter Field:=C_Cta_Reg_Mov3
        Tb_JyC.Range.AutoFilter Field:=C_Cta_Reg_Mov4
        Tb_JyC.Range.AutoFilter Field:=C_Cta_Reg_Mov5
        Tb_JyC.Range.AutoFilter Field:=C_Cta_Fusión
        Tb_JyC.Range.AutoFilter Field:=C_Cta_Fusión, Criteria1:="*" & Range("i2") & "*", Operator:=xlFilterValues
            On Error Resume Next
        NumFilas = Tb_JyC.DataBodyRange.Resize(, 1).SpecialCells(xlCellTypeVisible).Count
            On Error GoTo 0
        If NumFilas = 0 Then
            MsgBox "No he encontrado coincidencia"
            Tb_JyC.Range.AutoFilter Field:=C_Cta_Fusión
        Else
            Tb_JyC.DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1, C_Cta_N_Liq).Select
        End If
    End If
End Sub

Sub kk()
    ThisWorkbook.Activate events
End Sub
