Attribute VB_Name = "Z_Migrar_RibbonUI"
' Last Rev. 2026-10-04 13:55
'==================================================================================================
' Z_Migrar_RibbonUI - modulo de UN SOLO USO (2026-10-04). Quitarlo del libro al terminar la fase 5.
'
' Z_Fase1_Rellenar_RibbonUI: vacia Lo_RibbonUI (traia los 61 registros de Jornadas y Congresos) y
' crea una fila por cada tag del Ribbon de PPub que lee la tabla (26). Usuario, hojas, rutina y
' grupo van fijados aqui; la descripcion y el informe se copian de la fila de Tb_Tareas de ese tag
' (columna Uribbon-Tags), para que viajen enteros, con sus saltos de linea y tildes.
' Ejecutar desde la ventana Inmediato:   Z_Fase1_Rellenar_RibbonUI
'
' Z_Fase5_Podar_Tareas: borra de Tb_Tareas las filas de botones (las que tienen Uribbon-Tags), las
' tareas sin tag cuya rutina tiene boton (grupo C) y las que apuntan a rutinas que no existen
' (grupo D), 57 en total, y la columna Visible. Deben quedar 8 filas. Ejecutar DESPUES de
' reimportar la tanda 2 (M_000_Ini_Var_APP ya con Task_SheetsButton = 9).
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
'==================================================================================================
Public Sub Z_Fase5_Podar_Tareas()
    Dim Lo          As ListObject:      Set Lo = Prog__Menu_Aux.ListObjects(1)
    Dim Col         As ListColumn
    Dim C_Tarea     As Long
    Dim C_Tag       As Long
    Dim C_Rut       As Long
    Dim Fila        As Long
    Dim i           As Long
    Dim Rut         As String
    Dim Borrar      As New Collection
    Dim Grupo_C     As Variant
    Dim Grupo_D     As Variant

    Grupo_C = Array("Rut_Recalcular_Tabla_JIs_AE4", "Rut_Activar_Programacion", "Rut_WrkBook_CopSegTimed_USB", _
                    "Rut_RibbonX_ShowAll", "Rut_Context_Buttons_Hide", "Rut_Context_Buttons_Restore", _
                    "Rut_ProtectUnProtect_ActivSheet", "Rut_OnOff_SW_Probando", "Rut_Sheets_ShowAll")
    Grupo_D = Array("Rut_Cerrar_Menu", "Rut_LstObj_Export_WS_XlsM", "Rut_Lo_Export_WS_ByHand", _
                    "Rut_Recalcular_Tabla_JIs_303", "Rut_Recalcular_Tabla_JIs_PPb", "Rut_Genero_LIQx_PDF", _
                    "Rut_Genero_LIQxn_PDF", "Rut_Cambiar_Contrase*", "RuT_Listar_Planes", _
                    "RuT_Restituir_Tabla_Prog_TitPH", "Rut_Exportar_Saldo_Liquidacion", _
                    "Rut_Resumen_Tab_TitPropios_UNO", "Rut_Generar_Tabla_AD_TitPropios", _
                    "Rut_Generar_Tabla_Saldos_TitPropios", "RuT_Marcar_Plazos_AD", "Rut_Resumen_Tab_TitPropios", _
                    "Rut_Reset_ToolsBar", "Rut_Prueba_Rut_Progreso")

    C_Tarea = Lo.ListColumns("Tarea").Index
    C_Tag = Lo.ListColumns("Uribbon-Tags").Index
    C_Rut = Lo.ListColumns("Nombre_Rut").Index
    Debug.Print "Z_Fase5_Podar_Tareas - filas a borrar:"
    For Fila = 1 To Lo.ListRows.Count
        Rut = Trim$(CStr(Lo.DataBodyRange.Cells(Fila, C_Rut).Value))
        If Len(Trim$(CStr(Lo.DataBodyRange.Cells(Fila, C_Tag).Value))) > 0 _
           Or Z_En_Lista(Rut, Grupo_C) Or Z_En_Lista(Rut, Grupo_D) Then
            Borrar.Add Fila
            Debug.Print Fila, Lo.DataBodyRange.Cells(Fila, C_Tarea).Value, Rut
        End If
    Next Fila

    If MsgBox("Se borran " & Borrar.Count & " de las " & Lo.ListRows.Count & " filas de " & Lo.Name & _
              " (quedan " & Lo.ListRows.Count - Borrar.Count & "; se esperaban 57 y 8) y la columna Visible." & _
              vbLf & vbLf & "La lista está en la ventana Inmediato." & vbLf & vbLf & "¿Seguimos?", _
              vbYesNo + vbQuestion, "Fase 5: podar Tb_Tareas") <> vbYes Then Exit Sub
    On Error Resume Next
    Prog__Menu_Aux.Unprotect                   '- las hojas de este libro no llevan contrasena
    Set Col = Lo.ListColumns("Visible")
    On Error GoTo 0

    For i = Borrar.Count To 1 Step -1             '- de abajo arriba, para no mover las que faltan
        Lo.ListRows(Borrar(i)).Delete
    Next i
    If Not Col Is Nothing Then Col.Delete

    If Lo.ListColumns("SheetsButton").Index <> Task_SheetsButton Or Lo.ListColumns("Emails").Index <> Task_Emails Then
        MsgBox "¡ Las columnas de " & Lo.Name & " no casan con las constantes Task_* !" & vbLf & _
               "¿Se reimportó M_000_Ini_Var_APP de la tanda 2?", vbExclamation, "Fase 5"
    Else
        MsgBox "Hecho: quedan " & Lo.ListRows.Count & " filas en " & Lo.Name & "." & vbLf & vbLf & _
               "Ya puedes quitar el módulo Z_Migrar_RibbonUI del proyecto.", vbInformation, "Fase 5"
    End If
    Call RefreshRibbon
End Sub
'==================================================================================================
Private Function Z_En_Lista(ByVal Rut As String, ByVal Lista As Variant) As Boolean
    Dim Nom         As Variant
    If Len(Rut) = 0 Then Exit Function
    For Each Nom In Lista
        If LCase$(Rut) Like LCase$(CStr(Nom)) Then Z_En_Lista = True: Exit Function
    Next Nom
End Function
