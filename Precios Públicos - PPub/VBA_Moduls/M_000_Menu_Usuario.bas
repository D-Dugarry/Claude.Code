Attribute VB_Name = "M_000_Menu_Usuario"
' Last Rev. 2026-10-04 13:55
Option Explicit

' ==================================================================================================================================
'- Boton Cambiar Usuario del Ribbon (OnAct_ChangeUser). Antes pasaba por RuT_Ejecutar_Rut("Rut_Chg_Usuario"),
'  que buscaba la rutina en Tb_Tareas y mostraba un Form_MsgBox con APP_Task_Inf (2026-10-04).
Sub Rut_Usuario_Chg()
        Form_Usuario.Show
        Call Rut_RibbonUI_Guardar_Informe("ChangeUser", "Usuario activo: " & Prog__APP.Range("APP_User_Name") & "  -  " & Now)
        Call RefreshRibbon
        Application.ScreenUpdating = True
        DoEvents
End Sub

' ==================================================================================================================================
'- Rut_Filtrar_Tareas se retiro el 2026-10-04: reescribia la columna Visible de Tb_Tareas en cada
'  activacion de hoja; la visibilidad de los botones sale ahora de Lo_RibbonUI (M___RibbonUI_Rules).
' ==================================================================================================================================

