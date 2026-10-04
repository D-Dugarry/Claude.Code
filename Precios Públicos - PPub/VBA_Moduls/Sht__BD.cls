VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sht__BD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Last Rev. 2026-10-04 13:55
'Option Explicit

Private Sub Worksheet_Activate()
Debug.Print "Worksheet_Activate - " & ActiveSheet.Name
'        Call Rut_ConfigExcel_Establecer    '- Establece una Configuración Excel personalizada.
    Call RefreshRibbon
End Sub

Private Sub Worksheet_Deactivate()
Debug.Print "Worksheet_DeActivate - " & ActiveSheet.Name
    Call Rut_ConfigExcel_RESTABLECER    '- Restaura Configuración Excel.
End Sub




