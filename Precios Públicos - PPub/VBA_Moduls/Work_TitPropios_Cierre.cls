VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Work_TitPropios_Cierre"
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


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub Worksheet_Change(ByVal Target As Range) '>>>>>>>>>>>> SI CAMBIA EL CONTENIDO DE LA CELDA SE ACTIVA (aunque acabe teniendo el mismo dato)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


If Not Prog__APP.Range("SW_Events") Then Exit Sub

    ' ============ Cambiar Plan  ==========================================================================
    If Not Application.Intersect(Target, Range("TP_Cod_Plan")) Is Nothing Then
    Dim Criterio As String
        Application.ScreenUpdating = False
        ActiveSheet.Unprotect
        With ActiveSheet.ListObjects(1)
            .ShowAutoFilter = True
            .AutoFilter.ShowAllData
            If Range("TP_Cod_Plan") = "" Then
                .AutoFilter.ShowAllData
                Exit Sub
            End If
            Criterio = Range("TP_Cod_Plan")
            .Range.AutoFilter Field:=1, Criteria1:=Criterio
            If .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count < 2 Then
                MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "No hay Datos de este Plan:    """ & Criterio & """", _
                vbOKOnly, "Proceso: Filtrado de Planes de Títulos Propios" & vbCrLf & "Quito Filtros."
                .AutoFilter.ShowAllData
            End If
        End With
        Range("f2").Select
        Application.ScreenUpdating = True
        ActiveSheet.Protect
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

'    ' ============ Filtra DNI's en Tb_Tasas ==========================================================================
'
'    If Not Application.Intersect(Target, Range(Cells(1, C_Liquid_T_Tasas_ColIni + CPH_DNI - 1).Address)) Is Nothing Then
'        With Prog_LiquidTyC.ListObjects("Tb_Tasas")
'            Range(Cells(3, C_Liquid_T_Tasas_ColIni + CPH_DNI - 1).Address).Select
'            .AutoFilter.ShowAllData
'            .Range.AutoFilter Field:=CPH_DNI, Criteria1:="*" & Cells(1, C_Liquid_T_Tasas_ColIni + CPH_DNI - 1) & "*", Operator:=xlFilterValues
''            If Range(Cells(1, C_Liquid_T_Tasas_ColIni + CPH_DNI - 1).Address) = "" Then
''                Range(Cells(1, C_Liquid_T_Tasas_ColIni + CPH_DNI - 1).Address) = "¿ D.N.I. ?"
''                .AutoFilter.ShowAllData
''            End If
'            If Cells(1, C_Liquid_T_Tasas_ColIni + CPH_DNI - 1) = "" Then
'                Cells(1, C_Liquid_T_Tasas_ColIni + CPH_DNI - 1) = "¿ D.N.I. ?"
'                .AutoFilter.ShowAllData
'            End If
'        End With
'        Range(Cells(2, C_Liquid_T_Tasas_ColIni + CPH_DNI - 1).Address).Select
'        ActiveCell.Offset(1, 0).Select
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------

End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Sub Worksheet_SelectionChange(ByVal Target As Range)    '>>>>>>>>>>>>>>>>>>>>>>>> SÓLO CLICANDO EN LA CELDA SE ACTIVA
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'
'If not SW_Events  Then Exit Sub
'
'    ' ============ Form_Intro_Num_ExpAdm_JI_RDT  ==========================================================================
'    If Not Application.Intersect(Target, _
'            Range("a" & Plan_Lin_JI & ", a" & Plan_Lin_ExpAdm & ", a" & Plan_Lin_RDT & ", a" & Plan_Lin_IR & _
'                  ",b" & Plan_Lin_JI & ", b" & Plan_Lin_ExpAdm & ", b" & Plan_Lin_RDT & ", b" & Plan_Lin_IR)) Is Nothing Then
'        If IsEmpty(Range("a3,a9,a24")) Then MsgBox "Seleccionar primero una Liquidación", , "Módulo de Liquidaciones de Tasas": Exit Sub
'        Form_Intro_Num_ExpAdm_JI_RDT.Show
'        Range("f4").Select
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ RuT_Cargar_Datos_Plan_de_DR  ==========================================================================
'    If Not Application.Intersect(Target, Range("c1:d1,d" & Plan_Lin_JI)) Is Nothing Then
'        Range("f4").Select
'        Form_Select_Plan.Show
'        Range("f4").Select
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ RuT_Incorporar_Tasas_Adm_en_Tb_Tasas  ==========================================================================
'    If Not Application.Intersect(Target, Range("d" & Plan_Lin_T_Adm)) Is Nothing And IsNumeric(Left(Plan_Código, 1)) Then
'        ' ========== Si es un TÍTULO PROPIO - MASTER, ESPECIALISTA Y EXPERTO, el Plan empieza por Número.
'        Range("c" & Plan_Lin_T_Adm - 1).Select
'        RuT_Incorporar_Tasas_Adm_en_Tb_Tasas
'        RuT_Incorporar_Datos_Tb_Tasas_en_DR
'        Plan_Código = Range("d3")
'        RuT_Cargar_Datos_Plan_de_DR     ' Para Actualizar las Tablas
'        Range("c" & Plan_Lin_T_Adm - 1).Select
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ RuT_Incorporar_Datos_Tb_Tasas_en_DR  ==========================================================================
'    If Not Application.Intersect(Target, Range(Cells(1, C_Liquid_T_Tasas_ColIni + CPH_JI_Emi - 1).Address)) Is Nothing Then
'        ' ========== Si es un TÍTULO PROPIO - MASTER, ESPECIALISTA Y EXPERTO, el Plan empieza por Número.
'        Range(Cells(2, C_Liquid_T_Tasas_ColIni + CPH_JI_Emi - 1).Address).Select
'        RuT_Incorporar_Datos_Tb_Tasas_en_DR
'        Plan_Código = Range("d3")
'        RuT_Cargar_Datos_Plan_de_DR    ' Para Actualizar las Tablas
'        Range(Cells(2, C_Liquid_T_Tasas_ColIni + CPH_JI_Emi - 1).Address).Select
'        MsgBox "¡Hecho!", , "Módulo de Liquidaciones de Tasas"
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ RuT_Elegir_Liquidación  ==========================================================================
'    If Not Prog_LiquidTyC.ListObjects("Tb_Plan").DataBodyRange Is Nothing Then
'        With Prog_LiquidTyC.ListObjects("Tb_Plan").DataBodyRange
'            If Not Application.Intersect(Target, Range("f" & .Cells(1, 1).Row & ":f" & .Cells(0, 1).Row + .Rows.Count)) Is Nothing Then
'                    ' ---------------------- Cambiar Color Celda y Texto -------------------------------
'                    Range("f" & .Cells(1, 1).Row & ":f" & .Cells(0, 1).Row + .Rows.Count).Interior.Color = RGB(ColFar, ColFag, ColFab)
'                    Range("f" & .Cells(1, 1).Row & ":f" & .Cells(0, 1).Row + .Rows.Count).Font.Color = RGB(ColTar, ColTag, ColTab)
'                    ActiveCell.Interior.Color = RGB(ColFbr, ColFbg, ColFbb)
'                    ActiveCell.Font.Color = RGB(ColTbr, ColTbg, ColTbb)
'                    ' ----------------------------------------------------------------------------------
'                Plan_Fila_Num_Liquid = ActiveCell.Row - .Cells(0, 1).Row ' ------ Calcula el Número de Liquidación
'                Plan_TipoExp = ActiveCell.Offset(0, 1)
'                If Plan_TipoExp = """ Sin Liquidar """ Then Plan_Num_Liquid = 0 Else Plan_Num_Liquid = Plan_Fila_Num_Liquid
'                Application.ScreenUpdating = False
'                    RuT_Ver_Liquidación_Núm_X
'                Application.ScreenUpdating = True
'                Range("f21").Select
'                Exit Sub
'            End If
'        End With
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ RuT_Guardar_Liquidación  ==========================================================================
'    If Not Application.Intersect(Target, Range("f1:g1")) Is Nothing Then
'        Range("f4").Select
'        Application.ScreenUpdating = False
'            RuT_Guardar_Liquidación
'        Application.ScreenUpdating = True
'        Range("f4").Select
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ Incorporar_Nums_ExpAdm_JI_RDT_de_Liquidación_en_DR  ==========================================================================
'    If Not Application.Intersect(Target, Range("k1:l1")) Is Nothing Then
'        Range("f4").Select
'        RuT_Incorporar_Datos_Tb_Tasas_en_DR
'        Range("f4").Select
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ RuT_Mostrar_Ocultar_Col_Tabla_Tasas_Emitidas_Cobradas  ==========================================================================
'    If Not Application.Intersect(Target, Range(Cells(1, C_Liquid_Col_Tb_Recaudación - 1).Address)) Is Nothing Then
'        Application.ScreenUpdating = False
'            RuT_Mostrar_Ocultar_Col_Tabla_Tasas_Emitidas_Cobradas
'        Application.ScreenUpdating = True
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ RuT_Mostrar_Ocultar_Col_Tasas_Liquid  ==========================================================================
'    If Not Application.Intersect(Target, Range(Cells(1, C_Liquid_T_Tasas_ColIni).Address & "," & Cells(1, C_Liquid_T_Tasas_ColIni + CPH_ImpCob - 1).Address)) Is Nothing Then
'        Application.ScreenUpdating = False
'            RuT_Mostrar_Ocultar_Col_Tasas_Liquid
'        Application.ScreenUpdating = True
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'    ' ============ Quitar Filtros Tb_Tasas ==========================================================================
'    If Not Application.Intersect(Target, Range(Cells(1, C_Liquid_T_Tasas_ColIni + CPH_Liquidación - 1).Address)) Is Nothing Then
'        With Prog_LiquidTyC.ListObjects("Tb_Tasas")
'            Range(Cells(1, C_Liquid_T_Tasas_ColIni + CPH_Liquidación - 1).Address) = "¿ D.N.I. ?"
'            .AutoFilter.ShowAllData
'        End With
'        Range(Cells(2, C_Liquid_T_Tasas_ColIni + CPH_Liquidación - 1).Address).Select
'        ActiveCell.Offset(1, 0).Select
'        Exit Sub
'    End If  ' -----------------------------------------------------------------------------------------------
'
'
'End Sub     ' Worksheet_SelectionChange




