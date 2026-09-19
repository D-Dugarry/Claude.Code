VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==================================================================================================================================
Sub Workbook_Open()
' ==================================================================================================================================
    Prog__APP_Switch.Range("Sw_EnableEvents") = True
    Call Rut_EnableEvents_Status_Reset
    Rut_ConfigExcel_Establecer
    Range("Usuario_ID") = ""
    Prog__APP_Switch.Range("Sw_Boss") = False
    Prog__APP_Switch.Range("Sw_Probando") = False
    Prog__APP_Switch.Range("Sw_EnableEvents") = True
    Set Lo_Tareas = Prog__Menú_Aux.ListObjects("Tb_Tareas")
    Call Rut_Usuario_Chg
    If Date - Range("APP_CopSeg_HD_Date") > 7 Then Call Rut_WrkBooK_CopSegTimed_WB_HD
    RuT_Al_Abrir_WorkBook
End Sub
' ==================================================================================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.ScreenUpdating = False
    Wk_TitP_Liquid.Select
    Rut_ConfigExcel_RESTABLECER
End Sub
' ==================================================================================================================================
Private Sub Workbook_Deactivate()
    Application.CommandBars("Cell").Reset  ' Restaura al cambiar de libro
End Sub
' ==================================================================================================================================
Private Sub Workbook_Activate()
'    Rut_ConfigExcel_Establecer
End Sub
' ==================================================================================================================================

