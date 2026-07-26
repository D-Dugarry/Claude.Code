VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Prog__Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Option Explicit

Private Sub Worksheet_Activate()
Debug.Print "Worksheet_Activate - " & ActiveSheet.Name
    Call Rut_Filtrar_Tareas
        Call Rut_Context_Buttons_Hide   '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
    Call RefreshRibbon
End Sub

Private Sub Worksheet_Deactivate()
Debug.Print "Worksheet_Deactivate - " & ActiveSheet.Name
    Call Rut_Context_Buttons_Restore    '- Restaura las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
End Sub




