Attribute VB_Name = "M44_Export_Inf_Acont_CAcad"
'2026-02-02
Option Explicit

'===================================================================================================================================
'- Guarda Copia de la ActiveSheet ==================================================================================================
'===================================================================================================================================
Sub Rut_WrkSht_Export_Inf_ACon_CAcad()
    Debug.Print "Rut_WrkSht_Export_Inf_ACon_CAcad"
    Dim TipoEP          As String
    Dim FichName        As String
    
    TipoEP = Mid(ActiveSheet.Name, 5, 4)
    
    ' Nombre sugerido: Archiv + resto
    FichName = "Inf_Recibos_" & TipoEP & Mid(Range("e2"), 23, 7) & "__" & _
                Format(Prog__APP.Range("APP_FechCierreCont"), "yyyy-mmm-dd") & "__" & _
                Format(Now(), "(dd-mm-yy hh.mm)") & ".xlsx"

    Call Rut_WrkSheet_Export_To_xlsx(ActiveSheet, FichName)

End Sub
'-----------------------------------------------------------------------------------------------------------------------------------




