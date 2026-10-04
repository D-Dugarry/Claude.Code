Attribute VB_Name = "Z_Migrar_RibbonUI"
' Last Rev. 2026-10-04 13:09
'==================================================================================================
' Z_Migrar_RibbonUI - modulo de UN SOLO USO (2026-10-04). Quitarlo del libro al terminar la fase 5.
'
' Z_Fase1_Rellenar_RibbonUI: vacia Lo_RibbonUI (traia los 61 registros de Jornadas y Congresos) y
' crea una fila por cada tag del Ribbon de PPub que lee la tabla (26). Usuario, hojas, rutina y
' grupo van fijados aqui; la descripcion y el informe se copian de la fila de Tb_Tareas de ese tag
' (columna Uribbon-Tags), para que viajen enteros, con sus saltos de linea y tildes.
' Ejecutar desde la ventana Inmediato:   Z_Fase1_Rellenar_RibbonUI
'==================================================================================================
Option Explicit

Private Lo_Rib         As ListObject
Private Lo_Tar         As ListObject

Public Sub Z_Fase1_Rellenar_RibbonUI()
    Set Lo_Rib = Prog__RibbonUI.ListObjects(1)
    Set Lo_Tar = Prog__Menu_Aux.ListObjects(1)
    If MsgBox("Se borran las " & Lo_Rib.ListRows.Count & " filas de " & Lo_Rib.Name & _
              " y se crean las 26 de PPub." & vbLf & vbLf & "¿Seguimos?", vbYesNo + vbQuestion, _
              "Fase 1: rellenar Lo_RibbonUI") <> vbYes Then Exit Sub
    On Error Resume Next
    Prog__RibbonUI.Unprotect                   '- las hojas de este libro no llevan contrasena
    On Error GoTo 0
    If Not Lo_Rib.DataBodyRange Is Nothing Then Lo_Rib.DataBodyRange.Delete

    '         Tag, Usuario, SheetsNames, Nom_Rut, Tag en Tb_Tareas, Group-Tag, [Descripcion fija]
    Call Z_Fila("ChangeUser", "", "", "Rut_Usuario_Chg", "ChangeUser", "ChangeUser_Group", "")
    Call Z_Fila("Reset", "", "", "Rut_Reset_App", "Reset", "Reset_Group", "")
    Call Z_Fila("RibbViewNone", "", "", "Rut_Menu_HideAll", "RibbViewNone", "Traffic_Light_Group", "")
    Call Z_Fila("RibbViewMin", "", "", "Rut_Menu_ShowAll_Short", "RibbViewMin", "Traffic_Light_Group", "")
    Call Z_Fila("RibbViewMax", "", "", "Rut_Menu_ShowAll", "RibbViewMax", "Traffic_Light_Group", "")
    Call Z_Fila("HelpComments", "Boss", "no", "Rut_x_Help_ShowHide", "HelpComments", "Help_Group", "")
    Call Z_Fila("ExImportBDatos_Group", "Boss", "BDatos", "", "ExImportBDatos_Group", "ExImportBDatos_Group", "")
    Call Z_Fila("ExportBDatos", "Boss,SusanaC,FannyR;RafaG", "", "Rut_Lo_Export_Hist_Bdatos", "ExportBDatos", "ExImportBDatos_Group", "")
    Call Z_Fila("ImportBDatos", "Boss,SusanaC,FannyR;RafaG", "", "Rut_Lo_Import_Hist_Bdatos", "ImportBDatos", "ExImportBDatos_Group", "")
    Call Z_Fila("RestoreBDatos", "Boss,SusanaC,FannyR;RafaG", "", "RuT_LstObj_Restore_BD", "RestoreBdatos", "ExImportBDatos_Group", "")
    Call Z_Fila("Import_G04_ACont", "Boss,SusanaC,FannyR;RafaG", "", "RuT_Update_LSGES04_ACont", "Import_G04_ACont", "ImportLsGes04_Group", "")
    Call Z_Fila("Import_G04_CAcadAnt", "Boss,SusanaC,FannyR;RafaG", "", "RuT_Update_LSGES04_IAdm_CAcadAnt", "Import_G04_CAcadAnt", "ImportLsGes04_Group", "")
    Call Z_Fila("Import_LSace06", "Boss,SusanaC,FannyR;RafaG", "", "RuT_Update_LSace06_CAcad_ImpAdm_INSS", "Import_LSace06", "ImportLSace06_Group", "")
    Call Z_Fila("Import_AE4", "Boss,SusanaC,FannyR;RafaG", "", "RuT_Update_AE4x4", "Import_AE4", "ImportAE4x4_Group", "")
    Call Z_Fila("ReCalculate_ListObj_Sht_Group", "Boss,SusanaC,FannyR;RafaG", "Inf_Recibos_TIO, Inf_RSm, JIs_AE4", "", "", "ReCalculate_ListObj_Sht_Group", "Recalcula la tabla de la hoja activa (Inf_Recibos_TIO, Inf_RSm o JIs_AE4).")
    Call Z_Fila("ReCalculate_ListObj_Sht", "Boss,SusanaC,FannyR;RafaG", "Inf_Recibos_TIO, Inf_RSm, JIs_AE4", "Rut_Recalcular_Tabla_Inf_Recibos, Rut_Recalcular_Tabla_Inf_RSm, Rut_Recalcular_Tabla_JIs_AE4", "ReCalculate_ListObj_Sht", "ReCalculate_ListObj_Sht_Group", "Genera la tabla de la hoja activa:" & vbLf & "- Inf_Recibos_TIO: la tabla para los JI's de los Recibos." & vbLf & "- Inf_RSm: la tabla de JI's de Precios Públicos." & vbLf & "- JIs_AE4: la tabla de JI's de ActivEco4.")
    Call Z_Fila("ExportWorkSheet_Group", "Boss,SusanaC,FannyR;RafaG", "_ConfigAPP, JIs_AE4, Inf_RSm, Inf_Felipe, Inf_Recibos_TIO", "", "ExportWorkSheet_Group", "ExportWorkSheet_Group", "")
    Call Z_Fila("ExportWorkSheet", "Boss", "BDatos, Inf_Recibos_TIO", "Rut_Lo_Export_WorkSheet", "ExportWorkSheet", "ExportWorkSheet_Group", "")
    Call Z_Fila("ShowHide_Lo_Cols_Group", "Boss", "BDatos, Inf_Recibos_TIO, BD_INSS", "", "ShowHide_Lo_Cols_Group", "ShowHide_Lo_Cols_Group", "")
    Call Z_Fila("ShowHide_Lo_Cols", "Boss", "BDatos, Inf_Recibos_TIO, BD_INSS", "", "ShowHide_Lo_Cols", "ShowHide_Lo_Cols_Group", "")
    Call Z_Fila("ShowHide_Lo_Cols_Row1x", "Boss", "BDatos, Inf_Recibos_TIO, BD_INSS", "", "ShowHide_Lo_Cols_Row1x", "ShowHide_Lo_Cols_Group", "")
    Call Z_Fila("ResetExcelConfig", "Boss", "", "Rut_ConfigExcel_Restablecer", "ResetExcelConfig", "Tools_Group", "")
    Call Z_Fila("RibbonRefresh", "Boss", "", "Rut_RibbonRefresh", "RibbonRefresh", "Tools_Group", "")
    Call Z_Fila("SearchVinculos", "Boss", "", "Rut_ListarHipervinculos", "SearchVinculos", "Tools_Group", "")
    Call Z_Fila("SW_WB_DeactivateOnOff", "Boss", "", "", "SW_WB_DeactivateOnOff", "Tools_Group", "")
    Call Z_Fila("RunRutPrueba", "Boss", "", "RuT_kkkk", "RunRutPrueba", "RunRutPrueba_Group", "Utilizado para probar la ejecución de rutinas: ejecuta la rutina de la columna Nom_Rut de esta " & "fila de Lo_RibbonUI (hoja RibbonUI). Cambiándola, se ejecutará la correspondiente.")

    MsgBox Lo_Rib.ListRows.Count & " filas creadas en " & Lo_Rib.Name & ".", vbInformation, "Fase 1"
    Call RefreshRibbon
End Sub
'==================================================================================================
Private Sub Z_Fila(Tag As String, Usuario As String, Hojas As String, NomRut As String, _
                   TagOrigen As String, Grupo As String, Optional Descripcion As String = "")
    Dim Fila        As ListRow
    Dim Pos         As Variant
    Dim Desc        As String:      Desc = Descripcion
    Dim Informe     As String
    If Len(TagOrigen) > 0 Then
        Pos = Application.Match(TagOrigen, Lo_Tar.ListColumns(Task_Uribbon_Tags).DataBodyRange, 0)
        If IsError(Pos) Then
            Debug.Print "Z_Fila: no hay fila en Tb_Tareas con Uribbon-Tags = " & TagOrigen
        Else
            If Len(Desc) = 0 Then Desc = CStr(Lo_Tar.DataBodyRange.Cells(Pos, Task_Descripcion).Value)
            Informe = CStr(Lo_Tar.DataBodyRange.Cells(Pos, Task_Rut_Informe).Value)
        End If
    End If
    If Desc Like "[=+@-]*" Then Desc = "'" & Desc            '- que Excel no lo tome por formula
    If Informe Like "[=+@-]*" Then Informe = "'" & Informe
    Set Fila = Lo_Rib.ListRows.Add
    With Fila.Range
        .Cells(1, Rib_Tag).Value = Tag
        .Cells(1, Rib_User).Value = Usuario
        .Cells(1, Rib_SheetsNames).Value = Hojas
        .Cells(1, Rib_Rut).Value = NomRut
        .Cells(1, Rib_Rut_Descrip).Value = Desc
        .Cells(1, Rib_Rut_Informe).Value = Informe
        .Cells(1, Rib_Group_Tag).Value = Grupo
    End With
End Sub
