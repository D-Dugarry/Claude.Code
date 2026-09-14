Attribute VB_Name = "M79_Crear_WB_EP_CAcad"
'2026-02-01
Option Explicit

'- Guarda Copia de Este Excel Para convertirlo en Original =========================================================================
'===================================================================================================================================
Sub Rut_Crear_WB_EFP_o_CFCyAFC_de_CAcad_Ant_o_Pos(Optional FichPath As String = "")
    Debug.Print "Rut_Crear_WB_EFP_o_CFCyAFC_de_CAcad_Ant_o_Pos"
    Dim FichNom     As String
    Dim TipoEP      As String
    
    If Range("APP_EFP_o_CFC") = "EFP" Then TipoEP = "EFP_" Else TipoEP = "CFCyAFC_"
    FichNom = TipoEP & Range("APP_CursAcad") & "_BaseDatos_Liq_" & Range("App_VersiónApp") & ".xlsm"
    
    If FichPath = "" Then FichPath = Fnc_Format_Ruta(ThisWorkbook.Path) & "\" & FichNom
        
    Dim FichSelect      As Variant
    FichSelect = Application.GetSaveAsFilename(FichPath, "Excel Files (*.xlsm), *" & "xlsm")
    
    If FichSelect <> False Then
        On Error GoTo Finalizar
        Application.DisplayAlerts = False
        ThisWorkbook.SaveCopyAs Filename:=FichSelect 'ConflictResolution:=True     '??? no se lo que hace, habría que investigar
        Application.DisplayAlerts = True
        On Error GoTo 0
        Debug.Print Left(FichSelect, InStrRev(FichSelect, "\"))
        Form_Menu.TB_Informe = "Generado el Nuevo Excel de Enseñanzas Propias: " & Now & vbCrLf & FichSelect
    '    ThisWorkbook.Close savechanges:=True
    Else
        MsgBox "Proceso para generar un nuevo Excel de Enseñanzas Propias, cancelado. " & Now
    End If

Finalizar:
    Application.Speech.Speak "Proceso completado."
End Sub     '-  "Rut_Crear_WB_EFP_o_CFCyAFC_de_CAcad_Ant_o_Pos"
' ----------------------------------------------------------------------------------------------------------------------------------


