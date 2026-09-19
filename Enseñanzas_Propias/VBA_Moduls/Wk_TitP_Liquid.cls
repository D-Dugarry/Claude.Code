VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wk_TitP_Liquid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'2026-01-09
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   SÓLO CLICANDO EN LA CELDA SE ACTIVA   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub Worksheet_SelectionChange(ByVal Target As Range)    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

If Not Prog__APP_Switch.Range("Sw_EnableEvents") Then Exit Sub
Debug.Print "CantSelectionChange", CantSelectionChange, ActiveCell.Address
CantSelectionChange = CantSelectionChange + 1
    ' ============ Quitar Todos los filtros de la Liquidación  ===========================================================
    If Not Application.Intersect(Target, Range("Liquid_Del_Filtro")) Is Nothing Then
    
        With ActiveSheet.ListObjects(1)     '>>>>>>>>>>>>>>>>>>>>
            If .ShowAutoFilter Then                         '--- Compruebo que la Tabla tiene los Filtros Activos --------
                With .AutoFilter
                     If .FilterMode Then .ShowAllData
                End With
            Else                                            '--- Si NO tiene los Filtros Activos, los Activo -------------
                .ShowAutoFilter = True
            End If
        End With                            '<<<<<<<<<<<<<<<<<<<<
        
        ActiveCell.Offset(1, 0).Select
            Application.EnableEvents = False
                Range("Liquid_Filtro_N_Liquid") = ""
                Range("Liquid_Filtro_N_Recibo") = ""
                Range("Liquid_Filtro_Año_Emi") = ""
                Range("Liquid_Filtro_Año_Cob") = ""
                Range("Liquid_Filtro_Año_Vto") = ""
            Call Rut_EnableEvents_Status_Reset
        
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Mostrar u Ocultar los comentarios de toda la hoja  =======================================
    If Not Application.Intersect(Target, Range("Liquid_Help")) Is Nothing Then
    
    ThisWorkbook.ActiveSheet.Unprotect
'    ThisWorkbook.ActiveSheet.Protect UserInterfaceOnly:=True
            If ActiveSheet.Comments.Count = 0 Then
                Call Rut_x_Help_ShowHide(Prog__TitP_Liq_Help)
                ActiveSheet.Shapes("Cuadro_Instrucciones").Visible = True
                ActiveSheet.Shapes("Cuadro_Instrucciones").Top = 100
                ActiveSheet.Shapes("Cuadro_Instrucciones").Left = 100
            Else
                ActiveSheet.UsedRange.ClearComments
            End If
            
        Application.ScreenUpdating = True
        DoEvents
            
        ActiveCell.Offset(6, 0).Select
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Mostrar u Ocultar Recibos Negativos  =======================================
    If Not Application.Intersect(Target, Range("Liquid_Sw_VerRecNeg")) Is Nothing Then
        Call Rut_OnOff_SW_VerRecNeg
        Application.ScreenUpdating = True
        DoEvents
        ActiveCell.Offset(1, 0).Select
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Mostrar u Ocultar Recibos NO_Cobrados  =======================================
    If Not Application.Intersect(Target, Range("Liquid_Sw_VerRecNOCob")) Is Nothing Then
        Call Rut_OnOff_SW_VerRecNoCob
        Application.ScreenUpdating = True
        DoEvents
        ActiveCell.Offset(2, 0).Select
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ======= Alternar Vista de la Tabla RDT: Clasificado x Núm-RDT o x Núm-Liquidación  ====================
    If Not Application.Intersect(Target, Range("Liquid_RDT_Totales")) Is Nothing Then
    
        Select Case Range("Liquid_RDT_Totales")
        Case Is = "Tot:"
            Call Rut_02_Generar_Tabla_RDT_x_NumRDT(Wk_TitP_Liquid.ListObjects(1))
        Case Is = "Tot="
            Call Rut_01_Generar_Tabla_RDT_x_NumLiquid(Wk_TitP_Liquid.ListObjects(1))
        Case Is = "Tot-"
            Call Rut_03_Generar_Tabla_RDT_x_NumLiquid_Con_Devoluciones(Wk_TitP_Liquid.ListObjects(1))
        End Select
            
        ActiveCell.Offset(0, -1).Select
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------


End Sub     ' Worksheet_SelectionChange


'>>>>>>>>>> SI CAMBIA EL CONTENIDO DE LA CELDA SE ACTIVA (aunque acabe teniendo el mismo dato)   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub Worksheet_Change(ByVal Target As Range)     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

If Not Prog__APP_Switch.Range("Sw_EnableEvents") Then Exit Sub
Debug.Print "CantChanges", CantChanges, ActiveCell.Address
CantChanges = CantChanges + 1
    ' ============ Cambiar Liquidación  ==========================================================================
    If Not Application.Intersect(Target, Range("Liquid_Plan")) Is Nothing Then
        
        If Sw_Cmb Then
        
            Dim Answer As Integer
            Answer = MsgBox("Parece que has hecho cambios y no los has guardado." & vbCrLf & _
                    "¿¿¿ Seguro que desea continuar SIN Grabar los Datos ???", vbOKCancel + vbExclamation)
            If Answer = vbCancel Then
                Application.EnableEvents = False
                    Range("Liquid_Plan") = PlanAnt
                Call Rut_EnableEvents_Status_Reset
                Exit Sub
            End If
        End If
        
        Call Rut_00_Liquid_TitProp(UCase(Range("Liquid_Plan")), Prog__APP.Range("APP_CursAcad"))
        Range("Liquid_Plan").Offset(2, 0).Select
        Call Rut_EnableEvents_Status_Reset
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Añadir Núm JI a los Registros visibles de la Liquidación  ================================
    If Not Application.Intersect(Target, Range("Liquid_Num_JI_Emi")) Is Nothing Then
        Application.EnableEvents = False
        Call Rut_3_Añadir_Datos_Liquid_TitProp(CLiq_JI_Emi, Range("Liquid_Num_JI_Emi"))
            Range("Liquid_Num_JI_Emi") = ""
        Call Rut_EnableEvents_Status_Reset
        Sw_Cmb = True
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Añadir Núm AD a los Registros visibles de la Liquidación  ================================
    If Not Application.Intersect(Target, Range("Liquid_Num_AD_0010")) Is Nothing Then
        Application.EnableEvents = False
        Call Rut_3_Añadir_Datos_Liquid_TitProp(CLiq_AD_0010, Range("Liquid_Num_AD_0010"))
            Range("Liquid_Num_AD_0010") = ""
        Call Rut_EnableEvents_Status_Reset
        Sw_Cmb = True
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Añadir Núm JI a los Registros visibles de la Liquidación  ================================
    If Not Application.Intersect(Target, Range("Liquid_Num_JI_443")) Is Nothing Then
        Application.EnableEvents = False
        Call Rut_3_Añadir_Datos_Liquid_TitProp(CLiq_JI_443, Range("Liquid_Num_JI_443"))
            Range("Liquid_Num_JI_443") = ""
        Call Rut_EnableEvents_Status_Reset
        Sw_Cmb = True
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Añadir ExpAdm a los Registros visibles de la Liquidación  ================================
    If Not Application.Intersect(Target, Range("Liquid_Num_ExpAdm")) Is Nothing Then
        Application.EnableEvents = False
        Call Rut_3_Añadir_Datos_Liquid_TitProp(CLiq_ExpAdm, Range("Liquid_Num_ExpAdm"))
            Range("Liquid_Num_ExpAdm") = ""
        Call Rut_EnableEvents_Status_Reset
        Sw_Cmb = True
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Añadir RDT a los Registros visibles de la Liquidación  ================================
    If Not Application.Intersect(Target, Range("Liquid_Num_RDT")) Is Nothing Then
        Application.EnableEvents = False
        Call Rut_3_Añadir_Datos_Liquid_TitProp(CLiq_RDT, Range("Liquid_Num_RDT"))
            Range("Liquid_Num_RDT") = ""
        Call Rut_EnableEvents_Status_Reset
        Sw_Cmb = True
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Añadir Coef_VRI a los Registros visibles de la Liquidación  ================================
    If Not Application.Intersect(Target, Range("Liquid_Coef_VRI")) Is Nothing Then
        Application.EnableEvents = False
        Call Rut_3_Añadir_Datos_Liquid_TitProp(CLiq_Coef_VRI, Range("Liquid_Coef_VRI"))
            Range("Liquid_Coef_VRI") = ""
        Call Rut_EnableEvents_Status_Reset
        Sw_Cmb = True
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Añadir Orgánica a los Registros visibles de la Liquidación  ================================
    If Not Application.Intersect(Target, Range("Liquid_Orgánica")) Is Nothing Then
        Application.EnableEvents = False
        Call Rut_3_Añadir_Datos_Liquid_TitProp(CLiq_Orgánica, Range("Liquid_Orgánica"))
            Range("Liquid_Orgánica") = ""
        Call Rut_EnableEvents_Status_Reset
        Sw_Cmb = True
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Filtrar Número de Liquidación  ===========================================================
    If Not Application.Intersect(Target, Range("Liquid_Filtro_N_Liquid")) Is Nothing Then
        If Range("Liquid_Filtro_N_Liquid") = "" Then
            ActiveSheet.ListObjects(1).Range.AutoFilter Field:=CLiq_NumLiquid
        Else
            Call Rut_x_Filtro_LoTb(ActiveSheet.ListObjects(1), CLiq_NumLiquid, Range("Liquid_Filtro_N_Liquid"))
        End If
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Filtrar Número de Recibo (Plazo) =========================================================
    If Not Application.Intersect(Target, Range("Liquid_Filtro_N_Recibo")) Is Nothing Then
        If Range("Liquid_Filtro_N_Recibo") = "" Then
            ActiveSheet.ListObjects(1).Range.AutoFilter Field:=CLiq_Núm_Rec
        Else
            Call Rut_x_Filtro_LoTb(ActiveSheet.ListObjects(1), CLiq_Núm_Rec, Range("Liquid_Filtro_N_Recibo"))
        End If
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Filtrar Año Vto. =========================================================
    If Not Application.Intersect(Target, Range("Liquid_Filtro_Año_Vto")) Is Nothing Then
        If Range("Liquid_Filtro_Año_Vto") = "" Then
            ActiveSheet.ListObjects(1).Range.AutoFilter Field:=CLiq_AñoVto
        Else
            Call Rut_x_Filtro_LoTb(ActiveSheet.ListObjects(1), CLiq_AñoVto, Range("Liquid_Filtro_Año_Vto"))
        End If
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Filtrar Año de Emisión del Recibo  =======================================================
    If Not Application.Intersect(Target, Range("Liquid_Filtro_Año_Emi")) Is Nothing Then
        If Range("Liquid_Filtro_Año_Emi") = "" Then
            ActiveSheet.ListObjects(1).Range.AutoFilter Field:=CLiq_F_Emi
        Else
            Call Rut_x_Filtro_LoTb(ActiveSheet.ListObjects(1), CLiq_F_Emi, Range("Liquid_Filtro_Año_Emi") & "*")
        End If
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

    ' ============ Filtrar Año del Cobro del Recibo  ========================================================
    If Not Application.Intersect(Target, Range("Liquid_Filtro_Año_Cob")) Is Nothing Then
        If Range("Liquid_Filtro_Año_Cob") = "" Then
            ActiveSheet.ListObjects(1).Range.AutoFilter Field:=CLiq_F_Cobro
        Else
            Call Rut_x_Filtro_LoTb(ActiveSheet.ListObjects(1), CLiq_F_Cobro, Range("Liquid_Filtro_Año_Cob") & "*")
        End If
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------


End Sub



