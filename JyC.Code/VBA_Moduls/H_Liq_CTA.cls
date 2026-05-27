VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "H_Liq_CTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ActiveWindow.Zoom = 100
End Sub

Private Sub worksheet_selectionchange(ByVal Target As Range)
    On Error GoTo LZoom
    Dim xZoom As Long
    xZoom = ActiveWindow.Zoom
    '- Si la celda seleccionada tiene una lista desplegable, entonces aumenta el zoom ----
    If Target.Validation.Type = xlValidateList Then xZoom = 130
LZoom:
    ActiveWindow.Zoom = xZoom
End Sub

Private Sub Worksheet_Activate()
    Call Rut_Filtrar_Tareas(Prog__APP.Range("APP_User_ID"))
    Call Rut_Context_Buttons_Hide
    RefreshRibbon
End Sub

Private Sub Worksheet_Deactivate()
    Application.CommandBars("Cell").Reset
    Application.CommandBars("List Range Popup").Reset
End Sub



