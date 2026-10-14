Attribute VB_Name = "M_142_Mail_Inf_Conta_TIO"
' Last Rev. 2026-10-14 09:39
Option Explicit

'===================================================================================================
' M_142_Mail_Inf_Conta_TIO
'
' Envía por correo el Informe Conta (hoja Inf_Conta_TIO, columnas B:W) como .xlsx (2026-10-14;
' Docs/Plan_Inf_Conta_TIO.md). Lo hace el motor de M_135 (Rut_Mail_Inf) con lo que dice
' Rut_Inf_Def_Conta (M_140): destinatarios en App_Mail_Inf_Conta_TIO, copia en la carpeta de red
' App_RutaRed_Inf_Conta_TIO\<año> y adjunto "Inf_Conta_TIO (datos del AAAA-MM-DD).xlsx".
'
' Subs públicas:
'   Rut_Mail_Inf_Conta                    - Botón «Enviar por correo» en Inf_Conta_TIO (lo lanza
'                                           OnAct_Mail_Inf_Recibos, M___RibbonUI)
'   Rut_Inf_Conta_Probar_Xlsx             - Prueba: genera el .xlsx sin enviarlo y lo señala en el
'                                           Explorador (Inmediato: Call Rut_Inf_Conta_Probar_Xlsx)
'===================================================================================================

Sub Rut_Mail_Inf_Conta()
    Dim Def         As T_Inf_Def
    Call Rut_Inf_Def_Conta(Def)                     '- M_140
    Call Rut_Mail_Inf(Def)                          '- M_135
End Sub     ' Rut_Mail_Inf_Conta
'---------------------------------------------------------------------------------------------------

Sub Rut_Inf_Conta_Probar_Xlsx()
    Dim Def         As T_Inf_Def
    Call Rut_Inf_Def_Conta(Def)                     '- M_140
    Call Rut_Inf_Probar_Xlsx(Def)                   '- M_135
End Sub     ' Rut_Inf_Conta_Probar_Xlsx
'---------------------------------------------------------------------------------------------------
