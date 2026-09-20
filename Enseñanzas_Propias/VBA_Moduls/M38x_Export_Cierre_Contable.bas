Attribute VB_Name = "M38x_Export_Cierre_Contable"
' Last Rev. 2026-09-20 14:39
Option Explicit

'===================================================================================================================================
'- Guarda Copia de la ActiveSheet ==================================================================================================
'===================================================================================================================================
Sub Rut_WrkSht_Export_Cierre_Contable_AñoCont()
    Debug.Print "Rut_WrkSht_Export_Cierre_Contable_AñoCont"
    Dim TipoEP          As String
    Dim FichName        As String
    
    '--- Elegir prefijo según APP ---
    If Prog__APP.Range("APP_EFP_o_CFC") = "EFP" Then
        TipoEP = "EFP_"
    Else
        TipoEP = "CFCyAFC_"
    End If
    
    ' Nombre sugerido: Archiv + resto
    FichName = "NUEVO_" & TipoEP & Prog__APP.Range("APP_CursAcad") _
               & "_Cierre_" & Prog__APP.Range("APP_AñoCont") & " " & Format(Now, "(yyyy-mm-dd_hhmm)") & ".xlsx"

    Call Rut_WrkSheet_Export_To_xlsx(ActiveSheet, FichName)

    Application.Speech.Speak "Proceso completado."
'    Call Rut_EnableEvents_Status_Reset
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------


