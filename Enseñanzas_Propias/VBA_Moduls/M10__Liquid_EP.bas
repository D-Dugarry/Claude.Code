Attribute VB_Name = "M10__Liquid_EP"
'- M10_Liquid_EP - Modif: 2025-10-15
Option Explicit

Public Sw_Cmb        As Boolean
Public PlanAnt      As String
Public CrsAcadAnt   As String

Const ColRDT_Num_Liq    As Integer = 6
Const ColRDT_Imp_Cob    As Integer = 7
Const ColRDT_Imp_Adm    As Integer = 8
Const ColRDT_JI         As Integer = 9
Const ColRDT_ExpAdm     As Integer = 10
Const ColRDT_RDT        As Integer = 11
Const ColRDT_T_Acad     As Integer = 12
Const ColRDT_Org        As Integer = 13
Const ColRDT_Orgánica   As Integer = 14
Const ColRDT_VRI        As Integer = 15
Const ColRDT_Coef_VRI   As Integer = 16



' ==================================================================================================================================
Sub Rut_00_Liquid_TitProp(Liq_Plan As String, CursAcad As String)
' ==================================================================================================================================
Dim F_Liq               As Long
Dim Rec_Count           As Long:    Rec_Count = 0
Dim RowLiq          As ListRow
Dim RowDat          As ListRow
Dim rowfind         As Variant
Dim Pago_X_Alu      As Double
Dim TotEmi          As Double
Dim TotCob          As Double
Dim Ant_DNI         As String:      Ant_DNI = "InIcIo"
Dim ACont_1         As Integer:     ACont_1 = Left(CursAcad, 4)
Dim ACont_2         As Integer:     ACont_2 = "20" & Right(CursAcad, 2)

Dim Lo_BD           As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
Dim Lo_Liq          As ListObject:      Set Lo_Liq = Wk_TitP_Liquid.ListObjects(1)
Dim Lo_RetVRI       As ListObject:      Set Lo_RetVRI = Prog_Coef_Ret_VRI.ListObjects(1)
    
    Call Rut_Off_Functions
    Application.EnableEvents = False

H_Inicio = Timer                ' Para Saber el tiempo de proceso
Sw_Cmb = False:     PlanAnt = Liq_Plan:     CrsAcadAnt = CursAcad
    ' =============  Preparar Tabla ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_Coef_Ret_VRI)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)
    ' =============  Preparar Tabla  ==================
    Prog_BD.Unprotect
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    '--- Ordeno por Ref -----
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_DNI, xlAscending)
    Call Rut_Lo_Sort(Lo_BD, BD_NumRec, xlAscending)
    
    '--- Compruebo que hay datos en la Tabla --------------------------------------------
    If (Lo_BD.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1) < 1 Then
'        MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "No hay Tasas de este Plan: " & Liq_Plan, vbOKOnly,
        MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "No hay Tasas: " & Liq_Plan, vbOKOnly + vbExclamation, _
               "Proceso: Liquidación de Títulos Propios (EFP) o Cursos < 200h. (CFC, AFC, TNCT)"
        GoTo Restablecer_Valores
    End If
    
    '--- Detecto si es un Plan de Microcredenciales, Subvencionado ----------------------
    Call Rut_Detectar_Microcredencial(CursAcad, Liq_Plan)

    Wk_TitP_Liquid.Unprotect
    Wk_TitP_Liquid.Activate
    
    Application.EnableEvents = False
        Range("Liquid_Plan") = Liq_Plan
        Range("Liquid_Filtro_N_Liquid") = ""
        Range("Liquid_Filtro_N_Recibo") = ""
        Range("Liquid_Filtro_Año_Emi") = ""
        Range("Liquid_Filtro_Año_Cob") = ""
        '   Buscar Orgánica del Plan en Lo_RetVRI
        rowfind = Application.Match(Liq_Plan, Lo_RetVRI.DataBodyRange.Columns(CoefVRI_Plan), 0)
        If Not IsError(rowfind) Then    ' Plan Encontrado en Lo_RetVRI ==>>
            If Lo_RetVRI.DataBodyRange.Cells(rowfind, CoefVRI_Orgánica) = "" Then
                Range("Liquid_Plan_Name").Offset(0, 12) = "Orgánica desconocida, sorry!"
            Else
                Range("Liquid_Plan_Name").Offset(0, 12) = Lo_RetVRI.DataBodyRange.Cells(rowfind, CoefVRI_Orgánica)
            End If
            Range("Liquid_Plan_Name").Offset(0, 12).WrapText = False
        End If
    If Range("APP_EFP_o_CFC") = "EFP" Then
        Range("Liquid_Tipo_Plan").Offset(-1, 0) = "EFP, Eseñanza de Formación Permanente"
        Range("Liquid_Tipo_Plan") = "E.F.P. -  " & Range("Liquid_Curso_Acad")
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Name = "Arial Narrow"
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Size = 10
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Bold = False
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.ColorIndex = xlAutomatic
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Interior.ColorIndex = 19
    Else
        Range("Liquid_Tipo_Plan").Offset(-1, 0) = "CFC/AFC, Curso/Actividad de Formación Contínua, y otros."
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Name = "Arial Narrow"
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Size = 10
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Bold = False
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.ColorIndex = xlAutomatic
        Range("Liquid_Tipo_Plan").Offset(-1, 0).Interior.ColorIndex = 19
        If Range("APP_PlanMicroCred") <> "" Then
            Range("Liquid_Tipo_Plan").Offset(-1, 0) = "¡¡ Micro Credencial " & Range("APP_PlanMicroCred") & " !!"
            Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Name = "Arial"
            Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Bold = True
            Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.Size = 14
            Range("Liquid_Tipo_Plan").Offset(-1, 0).Font.ColorIndex = 6
            Range("Liquid_Tipo_Plan").Offset(-1, 0).Interior.ColorIndex = 3
        End If
        Range("Liquid_Tipo_Plan") = Lo_BD.Range.Columns(BD_Tipo_EP).SpecialCells(xlCellTypeVisible).Rows(2) & " - " & Range("Liquid_Curso_Acad")
    End If
    Range("Liquid_Plan_Name") = ""
    
    '--- Preparo Tabla Liquid. Tit.Prop., la dejo vacía ---------------------
    Call Rut_Lo_Filtros_Quitar(Lo_Liq)
    If Not Lo_Liq.DataBodyRange Is Nothing Then Lo_Liq.DataBodyRange.Delete
    If Lo_BD.DataBodyRange Is Nothing Then
            MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "No existen Tasas en el fich.", vbOKOnly + vbExclamation, _
                "Proceso: Liquidación de Títulos Propios (EFP) o Cursos < 200h. (CFC, AFC, TNCT)"
        GoTo Restablecer_Valores
    End If
        
    Application.EnableEvents = False
    For F_Liq = 1 To Lo_BD.DataBodyRange.Rows.Count

        Set RowDat = Lo_BD.ListRows(F_Liq)
        
        '--- Filtro los Registros -------------------------------------------------------------
        If RowDat.Range(BD_Plan) <> Liq_Plan Then GoTo Sig_Fila
        If RowDat.Range(BD_C_Acad) <> CursAcad Then GoTo Sig_Fila
        If RowDat.Range(BD_ImpAdm) < 0 Then GoTo Sig_Fila      '- Son Recibos de Ajuste de Matrícula, que distorcionan la Contabilidad.
        Rec_Count = Rec_Count + 1 '--- Existen Recibos pero no están pagados o el importe es <=0
        If Range("APP_PlanMicroCred") <> "" Then
            If InStr(RowDat.Range(BD_EP_GestReg), "Deleted") > 0 Then GoTo Sig_Fila

        Else
            If RowDat.Range(BD_ImpRec) < 0 And Not Range("Sw_VerRecNeg") Then GoTo Sig_Fila
            'If Len(RowDat.Range(BD_FCob)) = 0 And RowDat.Range(BD_ImpRec) >= 0 And Not Range("Sw_VerRecNoCob") Then GoTo Sig_Fila
            If RowDat.Range(BD_ImpCob) = 0 And Not Range("Sw_VerRecNoCob") And RowDat.Range(BD_ImpRec) > 0 Then GoTo Sig_Fila
        End If
        ' -------=============  Genero Texto de Descripción del JI  ==================---------
        If Range("Liquid_Plan_Name") = "" Then
            If Range("APP_PlanMicroCred") <> "" Then
                Range("Liquid_Plan_Name") = "LIQ-TitProp_" & Liq_Plan & _
                                            "-N Curso_" & Range("APP_CursAcad") & _
                                            " AñoCont_" & Range("APP_AñoCont") & _
                                            " - MicCred_" & Range("APP_PlanMicroCred") & _
                                            " - " & RowDat.Range(BD_NomPlan)
            Else
                Range("Liquid_Plan_Name") = "LIQ-TitProp_" & Liq_Plan & _
                                            "-N Curso_" & Range("APP_CursAcad") & _
                                            " AñoCont_" & Range("APP_AñoCont") & _
                                            " - " & RowDat.Range(BD_NomPlan)
            End If
        End If
        '--- Añado Registro -------------------------------------------------------------------
        Set RowLiq = Lo_Liq.ListRows.Add
        
        RowLiq.Range(CLiq_DNI) = RowDat.Range(BD_DNI)
        RowLiq.Range(CLiq_Nombre) = RowDat.Range(BD_Nom)
        RowLiq.Range(CLiq_NumLiquid) = RowDat.Range(BD_Liquidado)
        RowLiq.Range(CLiq_Núm_Rec) = RowDat.Range(BD_NumRec)
        RowLiq.Range(CLiq_AñoVto) = Right(RowDat.Range(BD_ACont_Vto), 2)
'            RowLiq.Range(CLiq_F_Emi) = Left(RowDat.Range(BD_FEmi), 4) & "-" & Right(RowDat.Range(BD_FEmi), 2)
'            RowLiq.Range(CLiq_F_Cobro) = Left(RowDat.Range(BD_FCob), 4) & "-" & Right(RowDat.Range(BD_FCob), 2)
        RowLiq.Range(CLiq_F_Emi) = RowDat.Range(BD_FEmi)
        RowLiq.Range(CLiq_F_Cobro) = RowDat.Range(BD_FCob)
        RowLiq.Range(CLiq_Imp_Rec) = RowDat.Range(BD_ImpRec)
        
        RowLiq.Range(CLiq_Imp_Cob) = RowDat.Range(BD_ImpCob)
        RowLiq.Range(CLiq_JI_Emi) = RowDat.Range(BD_JI_Emi_Acad)
        
'            If RowDat.Range(BD_ACont_Vto) > RowDat.Range(BD_ACont_Emi) And RowDat.Range(BD_AD_Emi_Acad) = "" Then
'                RowLiq.Range(CLiq_AD_0010) = "NO_AD"
'            Else
'            End If
        RowLiq.Range(CLiq_AD_0010) = RowDat.Range(BD_AD_Emi_Acad)
    
        RowLiq.Range(CLiq_JI_443) = RowDat.Range(BD_JI_443_Acad)
        RowLiq.Range(CLiq_ExpAdm) = RowDat.Range(BD_ExpAdm)
        RowLiq.Range(CLiq_RDT) = RowDat.Range(BD_RDT)
        RowLiq.Range(CLiq_Coef_VRI) = RowDat.Range(BD_Coef_VRI)
        RowLiq.Range(CLiq_Orgánica) = RowDat.Range(BD_Orgánica)
        RowLiq.Range(CLiq_Ref) = RowDat.Range(BD_Ref)
        RowLiq.Range(CLiq_Obs) = RowDat.Range(BD_Obs_Conta)
        
        RowLiq.Range(CLiq_ImpAcad) = RowDat.Range(BD_ImpAcad)
        RowLiq.Range(CLiq_ImpDto) = RowDat.Range(BD_ImpDto)
        RowLiq.Range(CLiq_Imp_Adm) = RowDat.Range(BD_Rec_Imp_Adm)
        If RowDat.Range(BD_Rec_Imp_Adm) <> RowDat.Range(BD_ImpAdm) And RowDat.Range(BD_Rec_Imp_Adm) <> "" _
            Then RowLiq.Range(CLiq_T_Adm_Neg) = RowDat.Range(BD_ImpAdm)
        
Sig_Fila:
    Next
    
    If Lo_Liq.DataBodyRange Is Nothing Then
        If Rec_Count > 0 Then
            MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & vbCrLf & "El plan: " & Liq_Plan & ", tiene Recibos Emitidos." & vbCrLf & vbCrLf & _
                    "Pero No hay Recibos pagados de este Plan." & vbCrLf & vbCrLf & "No hay nada para Liquidar." & vbCrLf & vbCrLf & _
                    vbCrLf & "¡ Puede probar con otro Plan !", vbOKOnly + vbExclamation, _
                   "Proceso: Liquidación de Títulos Propios (EFP) o Cursos < 200h. (CFC, AFC, TNCT)"
            GoTo Restablecer_Valores
        Else
            MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "No hay Tasas de este Plan: " & Liq_Plan, vbOKOnly + vbExclamation, _
                   "Proceso: Liquidación de Títulos Propios (EFP) o Cursos < 200h. (CFC, AFC, TNCT)"
            GoTo Restablecer_Valores
        End If
    End If
    
    Lo_Liq.ListColumns(CLiq_F_Emi).DataBodyRange.NumberFormat = "d-m-yyyy"
    Lo_Liq.ListColumns(CLiq_F_Cobro).DataBodyRange.NumberFormat = "d-m-yyyy"
    
'   Genera la Tabla de Resumen de las distintas Liquidaciones que ha tenido este Plan --------------------------------------------
    Call Rut_01_Generar_Tabla_RDT_x_NumLiquid(Lo_Liq)
    Call Rut_Calc_SubTotales
'    Call Rut_02_Generar_Tabla_RDT_x_NumRDT(Lo_Liq)   ' ==>>>> Esta opción se implementa como opción clicando en una celda de la hoja.

Restablecer_Valores:
    Call Rut_On_Functions
    Call Rut_EnableEvents_Status_Reset
'    IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ------------------------------------------------------
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Wk_TitP_Liquid.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
End Sub     ' RuT_00_Generar_Tabla_Rut_00_Liquid_TitProp     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

    
'==================================================================================================================================
'=============== Generar la Tabla de RDT
'==================================================================================================================================
Sub Rut_01_Generar_Tabla_RDT_x_NumLiquid(ByRef Lo_TPLiquid As ListObject)
'==================================================================================================================================

Dim F_Liq           As Integer
Dim F_Tb_RDT_N1     As Long:   F_Tb_RDT_N1 = 3
Dim F_Tb_RDT        As Long:   F_Tb_RDT = F_Tb_RDT_N1
Dim NumLiq_Ant         As String:  NumLiq_Ant = "xxxxx"

Dim Rg_Num_Liq    As Range
Dim Rg_Imp_Cob    As Range
Dim Rg_Imp_Adm    As Range
Dim Rg_Exp_Adm    As Range
Dim Rg_RDT        As Range
Dim Rg_T_Acad     As Range
Dim Rg_Org        As Range
Dim Rg_VRI        As Range
Dim Rg_Coef_VRI   As Range
Dim Rg_JI         As Range
Dim Rg_Orgánica   As Range
    
Dim RgT_Imp_Cob    As Range
Dim RgT_Imp_Adm    As Range
Dim RgT_T_Acad     As Range
Dim RgT_Org        As Range
Dim RgT_VRI        As Range
Dim RgT_Orgánica   As Range
    
Rut_Off_Functions
    
    Range("Liquid_RDT_Totales") = "Tot-"
    Cells(F_Tb_RDT - 1, 4) = "Tabla RDT x NºLiq sin Devoluciones"
    
    Lo_TPLiquid.AutoFilter.ShowAllData
    Call Rut_Lo_Sort(Lo_TPLiquid, CLiq_NumLiquid, xlAscending, True)

Do While Range("Liquid_RDT_Totales").Row > F_Tb_RDT
    Rows(F_Tb_RDT).Delete
Loop

    Set RgT_Imp_Cob = Cells(F_Tb_RDT, ColRDT_Imp_Cob)
    Set RgT_Imp_Adm = Cells(F_Tb_RDT, ColRDT_Imp_Adm)
    Set RgT_T_Acad = Cells(F_Tb_RDT, ColRDT_T_Acad)
    Set RgT_Org = Cells(F_Tb_RDT, ColRDT_Org)
    Set RgT_VRI = Cells(F_Tb_RDT, ColRDT_VRI)
    Set RgT_Orgánica = Cells(F_Tb_RDT, ColRDT_Orgánica)
RgT_Imp_Cob = 0
RgT_Imp_Adm = 0
RgT_T_Acad = 0
RgT_Org = 0
RgT_VRI = 0
RgT_Orgánica = ""

With Lo_TPLiquid.DataBodyRange

    For F_Liq = 1 To .Rows.Count   '--- Bucle para recorrer todas la filas de la Liquidación
        
        If NumLiq_Ant <> .Cells(F_Liq, CLiq_NumLiquid) Then
            '--- New RDT   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            NumLiq_Ant = .Cells(F_Liq, CLiq_NumLiquid)
            Rows(F_Tb_RDT).Insert
            
    Range("g4:o4").Select:          Selection.Copy:         Range("g3").Select  '- para copiar el formato.
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
            
            Set Rg_Num_Liq = Cells(F_Tb_RDT, ColRDT_Num_Liq)
            Set Rg_Imp_Cob = Cells(F_Tb_RDT, ColRDT_Imp_Cob)
            Set Rg_Imp_Adm = Cells(F_Tb_RDT, ColRDT_Imp_Adm)
            Set Rg_Exp_Adm = Cells(F_Tb_RDT, ColRDT_ExpAdm)
            Set Rg_RDT = Cells(F_Tb_RDT, ColRDT_RDT)
            Set Rg_T_Acad = Cells(F_Tb_RDT, ColRDT_T_Acad)
            Set Rg_Org = Cells(F_Tb_RDT, ColRDT_Org)
            Set Rg_VRI = Cells(F_Tb_RDT, ColRDT_VRI)
            Set Rg_Coef_VRI = Cells(F_Tb_RDT, ColRDT_Coef_VRI)
            Set Rg_JI = Cells(F_Tb_RDT, ColRDT_JI)
            Set Rg_Orgánica = Cells(F_Tb_RDT, ColRDT_Orgánica)
            
            If F_Tb_RDT > F_Tb_RDT_N1 Then
'                RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
'                RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
'                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob) > 0 Then RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob) > 0 Or Prog__APP.Range("Sw_VerRecNeg") Then RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm) > 0 Or Prog__APP.Range("Sw_VerRecNeg") Then RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
                RgT_T_Acad = RgT_T_Acad + Cells(F_Tb_RDT - 1, ColRDT_T_Acad)
                RgT_Org = RgT_Org + Cells(F_Tb_RDT - 1, ColRDT_Org)
                RgT_VRI = RgT_VRI + Cells(F_Tb_RDT - 1, ColRDT_VRI)
            End If
        
                F_Tb_RDT = F_Tb_RDT + 1
        
            If .Cells(F_Liq, CLiq_NumLiquid) = "" Then
                Rg_Num_Liq = "s/L"
            Else
                Rg_Num_Liq = .Cells(F_Liq, CLiq_NumLiquid)
            End If
            
            Rg_Exp_Adm = .Cells(F_Liq, CLiq_ExpAdm)
            Rg_JI = .Cells(F_Liq, CLiq_JI_Emi)
            Rg_Orgánica = .Cells(F_Liq, CLiq_Orgánica)
            Rg_RDT = .Cells(F_Liq, CLiq_RDT)
            Rg_Coef_VRI = Val(.Cells(F_Liq, CLiq_Coef_VRI))
            
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) And (.Cells(F_Liq, CLiq_Imp_Cob) > 0 Or Prog__APP.Range("Sw_VerRecNeg")) Then Rg_Imp_Cob = .Cells(F_Liq, CLiq_Imp_Cob) Else Rg_Imp_Cob = 0
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) And (.Cells(F_Liq, CLiq_Imp_Adm) > 0 Or Prog__APP.Range("Sw_VerRecNeg")) Then Rg_Imp_Adm = .Cells(F_Liq, CLiq_Imp_Adm) Else Rg_Imp_Adm = 0
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) Then Rg_Imp_Cob = .Cells(F_Liq, CLiq_Imp_Cob) Else Rg_Imp_Cob = 0
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) Then Rg_Imp_Adm = .Cells(F_Liq, CLiq_Imp_Adm) Else Rg_Imp_Adm = 0
            Rg_T_Acad = Rg_Imp_Cob - Rg_Imp_Adm
            Rg_VRI = Application.WorksheetFunction.Round(Rg_T_Acad * Rg_Coef_VRI / 100, 2)
            Rg_Org = Rg_T_Acad - Rg_VRI
        Else
            
            '--- ExpAdm/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Exp_Adm, .Cells(F_Liq, CLiq_ExpAdm), 1) = 0 Then ' Si tiene varias ExpAdm/s...
                Rg_Exp_Adm = Rg_Exp_Adm & "-" & .Cells(F_Liq, CLiq_ExpAdm)
            End If

            '--- Orgánica/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Orgánica, .Cells(F_Liq, CLiq_Orgánica), 1) = 0 Then ' Si tiene varias Orgánica/s...
                Rg_Orgánica = Rg_Orgánica & "-" & .Cells(F_Liq, CLiq_Orgánica)
            End If

            '--- RDT/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_RDT, .Cells(F_Liq, CLiq_RDT), 1) = 0 Then ' Si tiene varias RDT/s...
                Rg_RDT = Rg_RDT & "-" & .Cells(F_Liq, CLiq_RDT)
            End If

            '--- JI/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_JI, .Cells(F_Liq, CLiq_JI_Emi), 1) = 0 Then ' Si tiene varias JI/s...
                Rg_JI = Rg_JI & "-" & .Cells(F_Liq, CLiq_JI_Emi)
            End If

            '--- Acumula RDT   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Num_Liq, .Cells(F_Liq, CLiq_NumLiquid), 1) = 0 Then ' Si tiene varias liquidaciones...
                Rg_Num_Liq = Rg_Num_Liq & "-" & .Cells(F_Liq, CLiq_NumLiquid)
            End If
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) And (.Cells(F_Liq, CLiq_Imp_Cob) > 0 Or Prog__APP.Range("Sw_VerRecNeg")) Then Rg_Imp_Cob = Rg_Imp_Cob + .Cells(F_Liq, CLiq_Imp_Cob)
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) And (.Cells(F_Liq, CLiq_Imp_Adm) > 0 Or Prog__APP.Range("Sw_VerRecNeg")) Then Rg_Imp_Adm = Rg_Imp_Adm + .Cells(F_Liq, CLiq_Imp_Adm)
''            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) And .Cells(F_Liq, CLiq_Imp_Cob) > 0 Then Rg_Imp_Cob = Rg_Imp_Cob + .Cells(F_Liq, CLiq_Imp_Cob)
''            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) And .Cells(F_Liq, CLiq_Imp_Adm) > 0 Then Rg_Imp_Adm = Rg_Imp_Adm + .Cells(F_Liq, CLiq_Imp_Adm)
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) Then Rg_Imp_Cob = Rg_Imp_Cob + .Cells(F_Liq, CLiq_Imp_Cob)
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) Then Rg_Imp_Adm = Rg_Imp_Adm + .Cells(F_Liq, CLiq_Imp_Adm)
            Rg_T_Acad = Rg_Imp_Cob - Rg_Imp_Adm
            Rg_VRI = Application.WorksheetFunction.Round(Rg_T_Acad * Rg_Coef_VRI / 100, 2)
            Rg_Org = Rg_T_Acad - Rg_VRI
        End If
        
    Next F_Liq
End With

'                RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob) > 0 Then RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
                RgT_T_Acad = RgT_T_Acad + Cells(F_Tb_RDT - 1, ColRDT_T_Acad)
                RgT_Org = RgT_Org + Cells(F_Tb_RDT - 1, ColRDT_Org)
                RgT_VRI = RgT_VRI + Cells(F_Tb_RDT - 1, ColRDT_VRI)

    '--- Ordeno por Nombre -----
    Call Rut_Lo_Sort(Lo_TPLiquid, CLiq_Nombre, xlAscending, True)
    
Rut_On_Functions
End Sub     ' Rut_01_Generar_Tabla_RDT_x_NumLiquid     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================


'==================================================================================================================================
'=============== Generar la Tabla de RDT
' ==================================================================================================================================
Sub Rut_02_Generar_Tabla_RDT_x_NumRDT(ByRef Lo_TPLiquid As ListObject)

Dim F_Liq           As Integer
Dim F_Tb_RDT_N1     As Long:   F_Tb_RDT_N1 = 3
Dim F_Tb_RDT        As Long:   F_Tb_RDT = F_Tb_RDT_N1
Dim RDT_Ant         As String:  RDT_Ant = "xxxxx"

Dim Rg_Num_Liq    As Range
Dim Rg_Imp_Cob    As Range
Dim Rg_Imp_Adm    As Range
Dim Rg_Exp_Adm    As Range
Dim Rg_RDT        As Range
Dim Rg_T_Acad     As Range
Dim Rg_Org        As Range
Dim Rg_VRI        As Range
Dim Rg_Coef_VRI   As Range
Dim Rg_JI         As Range
Dim Rg_Orgánica   As Range
    
Dim RgT_Imp_Cob    As Range
Dim RgT_Imp_Adm    As Range
Dim RgT_T_Acad     As Range
Dim RgT_Org        As Range
Dim RgT_VRI        As Range
Dim RgT_Orgánica   As Range
    
Rut_Off_Functions
    
    Range("Liquid_RDT_Totales") = "Tot="
    Cells(F_Tb_RDT - 1, 4) = "Tabla RDT x NºRDT"
    
    Lo_TPLiquid.AutoFilter.ShowAllData
    Call Rut_Lo_Sort(Lo_TPLiquid, CLiq_RDT, xlAscending, True)

Do While Range("Liquid_RDT_Totales").Row > F_Tb_RDT
    Rows(F_Tb_RDT).Delete
Loop

    Set RgT_Imp_Cob = Cells(F_Tb_RDT, ColRDT_Imp_Cob)
    Set RgT_Imp_Adm = Cells(F_Tb_RDT, ColRDT_Imp_Adm)
    Set RgT_T_Acad = Cells(F_Tb_RDT, ColRDT_T_Acad)
    Set RgT_Org = Cells(F_Tb_RDT, ColRDT_Org)
    Set RgT_VRI = Cells(F_Tb_RDT, ColRDT_VRI)
    Set RgT_Orgánica = Cells(F_Tb_RDT, ColRDT_Orgánica)
RgT_Imp_Cob = 0
RgT_Imp_Adm = 0
RgT_T_Acad = 0
RgT_Org = 0
RgT_VRI = 0
RgT_Orgánica = ""

With Lo_TPLiquid.DataBodyRange

    For F_Liq = 1 To .Rows.Count   '--- Bucle para recorrer todas la filas de la Liquidación
        
        If RDT_Ant <> .Cells(F_Liq, CLiq_RDT) Then
            '--- New RDT   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            RDT_Ant = .Cells(F_Liq, CLiq_RDT)
            Rows(F_Tb_RDT).Insert
            
    Range("g4:o4").Select:          Selection.Copy:         Range("g3").Select  '- para copiar el formato.
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
            
            Set Rg_Num_Liq = Cells(F_Tb_RDT, ColRDT_Num_Liq)
            Set Rg_Imp_Cob = Cells(F_Tb_RDT, ColRDT_Imp_Cob)
            Set Rg_Imp_Adm = Cells(F_Tb_RDT, ColRDT_Imp_Adm)
            Set Rg_Exp_Adm = Cells(F_Tb_RDT, ColRDT_ExpAdm)
            Set Rg_RDT = Cells(F_Tb_RDT, ColRDT_RDT)
            Set Rg_T_Acad = Cells(F_Tb_RDT, ColRDT_T_Acad)
            Set Rg_Org = Cells(F_Tb_RDT, ColRDT_Org)
            Set Rg_VRI = Cells(F_Tb_RDT, ColRDT_VRI)
            Set Rg_Coef_VRI = Cells(F_Tb_RDT, ColRDT_Coef_VRI)
            Set Rg_JI = Cells(F_Tb_RDT, ColRDT_JI)
            Set Rg_Orgánica = Cells(F_Tb_RDT, ColRDT_Orgánica)
            

            If F_Tb_RDT > F_Tb_RDT_N1 Then
'                RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob) > 0 Then RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
                RgT_T_Acad = RgT_T_Acad + Cells(F_Tb_RDT - 1, ColRDT_T_Acad)
                RgT_Org = RgT_Org + Cells(F_Tb_RDT - 1, ColRDT_Org)
                RgT_VRI = RgT_VRI + Cells(F_Tb_RDT - 1, ColRDT_VRI)
            End If
        
                F_Tb_RDT = F_Tb_RDT + 1
        
            If .Cells(F_Liq, CLiq_NumLiquid) = "" Then
                Rg_Num_Liq = "s/L"
            Else
                Rg_Num_Liq = .Cells(F_Liq, CLiq_NumLiquid)
            End If
            
            Rg_Exp_Adm = .Cells(F_Liq, CLiq_ExpAdm)
            Rg_JI = .Cells(F_Liq, CLiq_JI_Emi)
            Rg_Orgánica = .Cells(F_Liq, CLiq_Orgánica)
            Rg_RDT = .Cells(F_Liq, CLiq_RDT)
            Rg_Coef_VRI = Val(.Cells(F_Liq, CLiq_Coef_VRI))
            
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) Then Rg_Imp_Cob = .Cells(F_Liq, CLiq_Imp_Cob) Else Rg_Imp_Cob = 0
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) Then Rg_Imp_Adm = .Cells(F_Liq, CLiq_Imp_Adm) Else Rg_Imp_Adm = 0
            Rg_T_Acad = Rg_Imp_Cob - Rg_Imp_Adm
            Rg_VRI = Application.WorksheetFunction.Round(Rg_T_Acad * Rg_Coef_VRI / 100, 2)
            Rg_Org = Rg_T_Acad - Rg_VRI
        Else
            
            '--- ExpAdm/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Exp_Adm, .Cells(F_Liq, CLiq_ExpAdm), 1) = 0 Then ' Si tiene varias ExpAdm/s...
                Rg_Exp_Adm = Rg_Exp_Adm & "-" & .Cells(F_Liq, CLiq_ExpAdm)
            End If

            '--- Orgánica/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Orgánica, .Cells(F_Liq, CLiq_Orgánica), 1) = 0 Then ' Si tiene varias Orgánica/s...
                Rg_Orgánica = Rg_Orgánica & "-" & .Cells(F_Liq, CLiq_Orgánica)
            End If

            '--- JI/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_JI, .Cells(F_Liq, CLiq_JI_Emi), 1) = 0 Then ' Si tiene varias JI/s...
                Rg_JI = Rg_JI & "-" & .Cells(F_Liq, CLiq_JI_Emi)
            End If

            '--- Acumula RDT   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Num_Liq, .Cells(F_Liq, CLiq_NumLiquid), 1) = 0 Then ' Si tiene varias Acumula RDT...
                Rg_Num_Liq = Rg_Num_Liq & "-" & .Cells(F_Liq, CLiq_NumLiquid)
            End If
            
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) And .Cells(F_Liq, CLiq_Imp_Cob) > 0 Then Rg_Imp_Cob = Rg_Imp_Cob + .Cells(F_Liq, CLiq_Imp_Cob)
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) And .Cells(F_Liq, CLiq_Imp_Adm) > 0 Then Rg_Imp_Adm = Rg_Imp_Adm + .Cells(F_Liq, CLiq_Imp_Adm)
            Rg_T_Acad = Rg_Imp_Cob - Rg_Imp_Adm
'''                Rg_Coef_VRI = .Cells(F_Liq, CLiq_Coef_VRI) '===>>> En teoría no hace falta pq todos deben tener el mismo......
            Rg_VRI = Application.WorksheetFunction.Round(Rg_T_Acad * Rg_Coef_VRI / 100, 2)
            Rg_Org = Rg_T_Acad - Rg_VRI
        End If
        
    Next F_Liq
End With

'                RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob) > 0 Then RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
                RgT_T_Acad = RgT_T_Acad + Cells(F_Tb_RDT - 1, ColRDT_T_Acad)
                RgT_Org = RgT_Org + Cells(F_Tb_RDT - 1, ColRDT_Org)
                RgT_VRI = RgT_VRI + Cells(F_Tb_RDT - 1, ColRDT_VRI)

    '--- Ordeno por Nombre -----
    Call Rut_Lo_Sort(Lo_TPLiquid, CLiq_Nombre, xlAscending, True)
    
Rut_On_Functions
End Sub     ' Rut_02_Generar_Tabla_RDT_x_NumRDT     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

'==================================================================================================================================
'=============== Generar la Tabla_RDT_x_NumLiquid_Con_Devoluciones
'==================================================================================================================================
Sub Rut_03_Generar_Tabla_RDT_x_NumLiquid_Con_Devoluciones(ByRef Lo_TPLiquid As ListObject)
'==================================================================================================================================

Dim F_Liq           As Integer
Dim F_Tb_RDT_N1     As Long:   F_Tb_RDT_N1 = 3
Dim F_Tb_RDT        As Long:   F_Tb_RDT = F_Tb_RDT_N1
Dim NumLiq_Ant         As String:  NumLiq_Ant = "xxxxx"

Dim Rg_Num_Liq    As Range
Dim Rg_Imp_Cob    As Range
Dim Rg_Imp_Adm    As Range
Dim Rg_Exp_Adm    As Range
Dim Rg_RDT        As Range
Dim Rg_T_Acad     As Range
Dim Rg_Org        As Range
Dim Rg_VRI        As Range
Dim Rg_Coef_VRI   As Range
Dim Rg_JI         As Range
Dim Rg_Orgánica   As Range
    
Dim RgT_Imp_Cob    As Range
Dim RgT_Imp_Adm    As Range
Dim RgT_T_Acad     As Range
Dim RgT_Org        As Range
Dim RgT_VRI        As Range
Dim RgT_Orgánica   As Range
    
Rut_Off_Functions
    
    Range("Liquid_RDT_Totales") = "Tot:"
    Cells(F_Tb_RDT - 1, 4) = "Tabla RDT x NºLiq con Devoluciones"
    
    Lo_TPLiquid.AutoFilter.ShowAllData
    Call Rut_Lo_Sort(Lo_TPLiquid, CLiq_NumLiquid, xlAscending, True)

Do While Range("Liquid_RDT_Totales").Row > F_Tb_RDT
    Rows(F_Tb_RDT).Delete
Loop

    Set RgT_Imp_Cob = Cells(F_Tb_RDT, ColRDT_Imp_Cob)
    Set RgT_Imp_Adm = Cells(F_Tb_RDT, ColRDT_Imp_Adm)
    Set RgT_T_Acad = Cells(F_Tb_RDT, ColRDT_T_Acad)
    Set RgT_Org = Cells(F_Tb_RDT, ColRDT_Org)
    Set RgT_VRI = Cells(F_Tb_RDT, ColRDT_VRI)
    Set RgT_Orgánica = Cells(F_Tb_RDT, ColRDT_Orgánica)
RgT_Imp_Cob = 0
RgT_Imp_Adm = 0
RgT_T_Acad = 0
RgT_Org = 0
RgT_VRI = 0
RgT_Orgánica = ""

With Lo_TPLiquid.DataBodyRange

    For F_Liq = 1 To .Rows.Count   '--- Bucle para recorrer todas la filas de la Liquidación
        
        If NumLiq_Ant <> .Cells(F_Liq, CLiq_NumLiquid) Then
            '--- New RDT   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            NumLiq_Ant = .Cells(F_Liq, CLiq_NumLiquid)
            Rows(F_Tb_RDT).Insert
            
    Range("g4:o4").Select:          Selection.Copy:         Range("g3").Select  '- para copiar el formato.
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
            
            Set Rg_Num_Liq = Cells(F_Tb_RDT, ColRDT_Num_Liq)
            Set Rg_Imp_Cob = Cells(F_Tb_RDT, ColRDT_Imp_Cob)
            Set Rg_Imp_Adm = Cells(F_Tb_RDT, ColRDT_Imp_Adm)
            Set Rg_Exp_Adm = Cells(F_Tb_RDT, ColRDT_ExpAdm)
            Set Rg_RDT = Cells(F_Tb_RDT, ColRDT_RDT)
            Set Rg_T_Acad = Cells(F_Tb_RDT, ColRDT_T_Acad)
            Set Rg_Org = Cells(F_Tb_RDT, ColRDT_Org)
            Set Rg_VRI = Cells(F_Tb_RDT, ColRDT_VRI)
            Set Rg_Coef_VRI = Cells(F_Tb_RDT, ColRDT_Coef_VRI)
            Set Rg_JI = Cells(F_Tb_RDT, ColRDT_JI)
            Set Rg_Orgánica = Cells(F_Tb_RDT, ColRDT_Orgánica)
            
            If F_Tb_RDT > F_Tb_RDT_N1 Then
'                RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
'                RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob) > 0 Or Range("Sw_VerRecNeg") Then RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm) > 0 Or Range("Sw_VerRecNeg") Then RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
                RgT_T_Acad = RgT_T_Acad + Cells(F_Tb_RDT - 1, ColRDT_T_Acad)
                RgT_Org = RgT_Org + Cells(F_Tb_RDT - 1, ColRDT_Org)
                RgT_VRI = RgT_VRI + Cells(F_Tb_RDT - 1, ColRDT_VRI)
            End If
        
                F_Tb_RDT = F_Tb_RDT + 1
        
            If .Cells(F_Liq, CLiq_NumLiquid) = "" Then
                Rg_Num_Liq = "s/L"
            Else
                Rg_Num_Liq = .Cells(F_Liq, CLiq_NumLiquid)
            End If
            
            Rg_Exp_Adm = .Cells(F_Liq, CLiq_ExpAdm)
            Rg_JI = .Cells(F_Liq, CLiq_JI_Emi)
            Rg_Orgánica = .Cells(F_Liq, CLiq_Orgánica)
            Rg_RDT = .Cells(F_Liq, CLiq_RDT)
            Rg_Coef_VRI = Val(.Cells(F_Liq, CLiq_Coef_VRI))
            
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) And (.Cells(F_Liq, CLiq_Imp_Cob) > 0 Or Range("Sw_VerRecNeg")) Then Rg_Imp_Cob = .Cells(F_Liq, CLiq_Imp_Cob) Else Rg_Imp_Cob = 0
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) And (.Cells(F_Liq, CLiq_Imp_Adm) > 0 Or Range("Sw_VerRecNeg")) Then Rg_Imp_Adm = .Cells(F_Liq, CLiq_Imp_Adm) Else Rg_Imp_Adm = 0
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) Then Rg_Imp_Cob = .Cells(F_Liq, CLiq_Imp_Cob) Else Rg_Imp_Cob = 0
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) Then Rg_Imp_Adm = .Cells(F_Liq, CLiq_Imp_Adm) Else Rg_Imp_Adm = 0
            Rg_T_Acad = Rg_Imp_Cob - Rg_Imp_Adm
            Rg_VRI = Application.WorksheetFunction.Round(Rg_T_Acad * Rg_Coef_VRI / 100, 2)
            Rg_Org = Rg_T_Acad - Rg_VRI
        Else
            
            '--- ExpAdm/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Exp_Adm, .Cells(F_Liq, CLiq_ExpAdm), 1) = 0 Then ' Si tiene varias ExpAdm/s...
                Rg_Exp_Adm = Rg_Exp_Adm & "-" & .Cells(F_Liq, CLiq_ExpAdm)
            End If

            '--- Orgánica/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Orgánica, .Cells(F_Liq, CLiq_Orgánica), 1) = 0 Then ' Si tiene varias Orgánica/s...
                Rg_Orgánica = Rg_Orgánica & "-" & .Cells(F_Liq, CLiq_Orgánica)
            End If

            '--- RDT/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_RDT, .Cells(F_Liq, CLiq_RDT), 1) = 0 Then ' Si tiene varias RDT/s...
                Rg_RDT = Rg_RDT & "-" & .Cells(F_Liq, CLiq_RDT)
            End If

            '--- JI/s   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_JI, .Cells(F_Liq, CLiq_JI_Emi), 1) = 0 Then ' Si tiene varias JI/s...
                Rg_JI = Rg_JI & "-" & .Cells(F_Liq, CLiq_JI_Emi)
            End If

            '--- Acumula RDT   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            If InStr(1, Rg_Num_Liq, .Cells(F_Liq, CLiq_NumLiquid), 1) = 0 Then ' Si tiene varias liquidaciones...
                Rg_Num_Liq = Rg_Num_Liq & "-" & .Cells(F_Liq, CLiq_NumLiquid)
            End If
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) And (.Cells(F_Liq, CLiq_Imp_Cob) > 0 Or Range("Sw_VerRecNeg")) Then Rg_Imp_Cob = Rg_Imp_Cob + .Cells(F_Liq, CLiq_Imp_Cob)
            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) And (.Cells(F_Liq, CLiq_Imp_Adm) > 0 Or Range("Sw_VerRecNeg")) Then Rg_Imp_Adm = Rg_Imp_Adm + .Cells(F_Liq, CLiq_Imp_Adm)
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) Then Rg_Imp_Cob = Rg_Imp_Cob + .Cells(F_Liq, CLiq_Imp_Cob)
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) Then Rg_Imp_Adm = Rg_Imp_Adm + .Cells(F_Liq, CLiq_Imp_Adm)
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Cob)) And .Cells(F_Liq, CLiq_Imp_Cob) > 0 Then Rg_Imp_Cob = Rg_Imp_Cob + .Cells(F_Liq, CLiq_Imp_Cob)
'            If IsNumeric(.Cells(F_Liq, CLiq_Imp_Adm)) And .Cells(F_Liq, CLiq_Imp_Adm) > 0 Then Rg_Imp_Adm = Rg_Imp_Adm + .Cells(F_Liq, CLiq_Imp_Adm)
            Rg_T_Acad = Rg_Imp_Cob - Rg_Imp_Adm
            Rg_VRI = Application.WorksheetFunction.Round(Rg_T_Acad * Rg_Coef_VRI / 100, 2)
            Rg_Org = Rg_T_Acad - Rg_VRI
        End If
        
    Next F_Liq
End With

'                RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
'                RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob) > 0 Or Range("Sw_VerRecNeg") Then RgT_Imp_Cob = RgT_Imp_Cob + Cells(F_Tb_RDT - 1, ColRDT_Imp_Cob)
                If Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm) > 0 Or Range("Sw_VerRecNeg") Then RgT_Imp_Adm = RgT_Imp_Adm + Cells(F_Tb_RDT - 1, ColRDT_Imp_Adm)
                
                RgT_T_Acad = RgT_T_Acad + Cells(F_Tb_RDT - 1, ColRDT_T_Acad)
                RgT_Org = RgT_Org + Cells(F_Tb_RDT - 1, ColRDT_Org)
                RgT_VRI = RgT_VRI + Cells(F_Tb_RDT - 1, ColRDT_VRI)

    '--- Ordeno por Nombre -----
    Call Rut_Lo_Sort(Lo_TPLiquid, CLiq_Nombre, xlAscending, True)
    
Rut_On_Functions
End Sub     ' Rut_03_Generar_Tabla_RDT_x_NumLiquid_Con_Devoluciones     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================


'-----   Actualiza los Datos de la Tabla de Histórico con los nuevos aportados en la tabla de Liquidación   ------------------------
' ==================================================================================================================================
Sub Rut_2_Actualizar_Dat_TitPropHist_con_Dat_Liquid()
' ==================================================================================================================================
Dim F_TitPH              As Long:   F_TitPH = 1
Dim F_Liquid               As Long
Rut_Off_Functions
    
    Dim Answer As Integer
    Answer = MsgBox("¿¿¿ Seguro que desea Actualizar los Datos de las Tasas con los datos aportados en esta Liquidación ???", vbOKCancel + vbExclamation)
    If Answer = vbCancel Then Exit Sub
    
Dim RowLiq As ListRow
Dim RowDat As ListRow
Dim Lo_TitPHist        As ListObject:           Set Lo_TitPHist = Prog_BD.ListObjects(1)
Dim Lo_TPLiquid        As ListObject:           Set Lo_TPLiquid = Wk_TitP_Liquid.ListObjects(1)
    
    Lo_TitPHist.AutoFilter.ShowAllData
    Call Rut_Lo_Sort(Lo_TitPHist, BD_Ref, xlAscending, True)
    Lo_TPLiquid.AutoFilter.ShowAllData
    Call Rut_Lo_Sort(Lo_TPLiquid, CLiq_Ref, xlAscending, True)

With Lo_TitPHist.DataBodyRange

    For F_Liquid = 1 To Lo_TPLiquid.DataBodyRange.Rows.Count   '--- Bucle para recorrer todas la filas de la Liquidación
        
        Select Case Lo_TPLiquid.DataBodyRange.Cells(F_Liquid, CLiq_Ref)
        
            '--- Saltar al siguiente Lo_TitPHist   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            Case Is > .Cells(F_TitPH, BD_Ref)
                F_Liquid = F_Liquid - 1
            '--- ACTUALIZAR CON Datos de la Liquidación <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            Case Is = .Cells(F_TitPH, BD_Ref)
            
                Set RowLiq = Lo_TPLiquid.ListRows(F_Liquid)
                Set RowDat = Lo_TitPHist.ListRows(F_TitPH)
                
                RowDat.Range(BD_Liquidado) = RowLiq.Range(CLiq_NumLiquid)
                RowDat.Range(BD_JI_Emi_Acad) = RowLiq.Range(CLiq_JI_Emi)
                RowDat.Range(BD_AD_Emi_Acad) = RowLiq.Range(CLiq_AD_0010)
                RowDat.Range(BD_JI_443_Acad) = RowLiq.Range(CLiq_JI_443)
                RowDat.Range(BD_ExpAdm) = RowLiq.Range(CLiq_ExpAdm)
                RowDat.Range(BD_RDT) = RowLiq.Range(CLiq_RDT)
                RowDat.Range(BD_Coef_VRI) = RowLiq.Range(CLiq_Coef_VRI)
                RowDat.Range(BD_Orgánica) = RowLiq.Range(CLiq_Orgánica)
                RowDat.Range(BD_Obs_Conta) = RowLiq.Range(CLiq_Obs)
            
                If IsEmpty(RowLiq.Range(CLiq_Ajst_Tadm)) Then
                    RowDat.Range(BD_Rec_Imp_Adm) = RowLiq.Range(CLiq_Imp_Adm)
                Else
                    RowDat.Range(BD_Rec_Imp_Adm) = RowLiq.Range(CLiq_Ajst_Tadm)
                End If
                
        End Select
        
        F_TitPH = F_TitPH + 1
        
    Next F_Liquid
End With
Sw_Cmb = False
Application.EnableEvents = False
    Range("Liquid_Filtro_N_Liquid") = ""
    Range("Liquid_Filtro_N_Recibo") = ""
    Range("Liquid_Filtro_Año_Emi") = ""
    Range("Liquid_Filtro_Año_Cob") = ""
Call Rut_EnableEvents_Status_Reset

    Call Rut_01_Generar_Tabla_RDT_x_NumLiquid(Lo_TPLiquid)
    '--- Ordeno por Nombre -----
    
    Call Rut_Lo_Sort(Lo_TitPHist, BD_Ref, xlAscending, True)
    Call Rut_Lo_Sort(Lo_TPLiquid, CLiq_Nombre, xlAscending, True)
MsgBox "¡¡¡ Hecho !!!" & vbCrLf & vbCrLf & "Ya están los datos de Tasa Adm. y Nº de Liquidación, JI, ExpAdm y RDT guardados, en: " & F_Liquid - 1 & " Reg.", _
                        vbOKOnly, "Proceso: Liquidación de Títulos Propios"
Rut_On_Functions
End Sub     ' Rut_2_Actualizar_Dat_TitPropHist_con_Dat_Liquid
' ==================================================================================================================================

'--------  Añade el Dato introducido en una celda a todas las líneas visibles (¡¡filtradas!!)  -----------------
' ==================================================================================================================================
Sub Rut_3_Añadir_Datos_Liquid_TitProp(CtLq_Col As Integer, ByRef valor As String)
' ==================================================================================================================================
Dim F_Liquid               As Long
'Rut_Off_Functions
Application.ScreenUpdating = False

    Dim Answer As Integer
    Answer = MsgBox("¿¿¿ Seguro que desea añadir el Dato a todos los registros selecionados (Visibles / Filtrados) ???", vbOKCancel + vbExclamation)
    If Answer = vbCancel Then GoTo Restablecer_Valores
    
Dim Lo_TPLiquid        As ListObject
Set Lo_TPLiquid = Wk_TitP_Liquid.ListObjects(1)
    
    For F_Liquid = 1 To Lo_TPLiquid.DataBodyRange.Rows.Count   '--- Bucle para recorrer todas la filas de la Liquidación
        
        If Lo_TPLiquid.DataBodyRange.Rows(F_Liquid).Hidden = False Then

                Lo_TPLiquid.DataBodyRange.Cells(F_Liquid, CtLq_Col) = valor
        End If
    Next F_Liquid
                        valor = ""
Restablecer_Valores:
'Rut_On_Functions
Application.ScreenUpdating = True
End Sub
' ==================================================================================================================================


'''    ' ==================================================================================================================================
'''    ' =============================     RuT_Guardar_Liquidación     =======================================================================
'''    ' ==================================================================================================================================
'''    Sub Rut_5_Guardar_Liquidación()
'''    Rut_Off_Functions
'''
'''    Dim FPath           As String
'''        FPath = ThisWorkbook.Path & "\"
'''
'''    '    Wk_TitP_Liquid.Select
'''        Wk_TitP_Liquid.Copy
'''    '    ActiveSheet.Copy
'''    Dim IntialName As String
'''    Dim sFileSaveName As Variant
'''        IntialName = FPath & "Liquid_" & Range("Liquid_Plan") & " - " & Format(Date, "dd-mmm-yyyy") & ".xlsx"   ' "_" & Format(Time, "hh-mm-ss") & ".xlsx"
'''        sFileSaveName = Application.GetSaveAsFilename(IntialName, "Excel Files (*.xlsx), *.xlsx")
'''            If sFileSaveName <> False Then
'''                On Error GoTo Restablecer_Valores
'''                Application.DisplayAlerts = False
'''                ActiveWorkbook.SaveAs sFileSaveName, ConflictResolution:=True
'''                Application.DisplayAlerts = True
'''                On Error GoTo 0
'''            End If
'''
'''    ActiveWorkbook.ActiveSheet.Shapes.SelectAll:   Selection.Delete
'''    ActiveWorkbook.ActiveSheet.UsedRange.ClearComments
'''    ActiveWorkbook.ActiveSheet.Rows(ActiveWorkbook.ActiveSheet.ListObjects(1).Range.Rows(1).Row - 1).Clear
'''    'ActiveWorkbook.ActiveSheet.Rows(ActiveWorkbook.ActiveSheet.ListObjects(1).Range.Rows(1).Row - 1).Clear
'''
'''    ActiveWorkbook.Close savechanges:=True
'''        MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Archivar Liquidación"
'''    Exit Sub
'''Restablecer_Valores:
'''    ActiveWorkbook.Close savechanges:=False
'''    Rut_On_Functions
'''    End Sub     ' RuT_Guardar_Liquidación
'''









' ==================================================================================================================================
Sub Rut_x_Help_ShowHide(ByRef WkS_Help As Worksheet)
' ==================================================================================================================================
Dim F_Help  As Integer
Dim Nombre  As String
Dim Rng     As Range

Dim Lo_TPLqHelp        As ListObject
Set Lo_TPLqHelp = WkS_Help.ListObjects(1)

    ActiveSheet.UsedRange.ClearComments
    
With Lo_TPLqHelp.DataBodyRange

    For F_Help = 1 To .Rows.Count
        Nombre = .Cells(F_Help, 1)
        Set Rng = Range(Nombre)
        ThisWorkbook.ActiveSheet.Range(Rng.Address).AddComment .Cells(F_Help, 2).Value2
        ThisWorkbook.ActiveSheet.Range(Rng.Address).Comment.Shape.TextFrame.AutoSize = True
    Next F_Help

End With

End Sub






' ==================================================================================================================================
Sub Rut_Calc_SubTotales()
' ==================================================================================================================================
Dim F_Liq           As Long
Dim Pago_X_Alu      As Double
Dim Rec_Emi         As Double
Dim Ant_DNI         As String
Dim ColorFondo      As Integer:     ColorFondo = 35
Dim ACont_1         As Integer:     ACont_1 = Left(Prog__APP.Range("APP_CursAcad"), 4)

Dim Lo_TPLiquid        As ListObject
Set Lo_TPLiquid = Wk_TitP_Liquid.ListObjects(1)
Dim Lo_TPLiquidPDF     As ListObject
Set Lo_TPLiquidPDF = Wk_TitP_LiqPDF.ListObjects(1)

Rut_Off_Functions
    
    Wk_TitP_Liquid.Unprotect
    
    Ant_DNI = Lo_TPLiquid.DataBodyRange.Cells(1, CLiq_DNI)
        
    For F_Liq = 1 To Lo_TPLiquid.DataBodyRange.Rows.Count

        With Lo_TPLiquid.ListRows(F_Liq)
            
            '--- Calculo el Total pagado por Alumno -----------------------------------------------
            If Ant_DNI = .Range(CLiq_DNI) Then
                Rec_Emi = Rec_Emi + .Range(CLiq_Imp_Rec)
                Pago_X_Alu = Pago_X_Alu + .Range(CLiq_Imp_Cob)
            Else
                Lo_TPLiquid.ListRows(F_Liq - 1).Range(CLiq_Rec_Emi) = Rec_Emi
                Lo_TPLiquid.ListRows(F_Liq - 1).Range(CLiq_Pag_X_Alu) = Pago_X_Alu
                If Pago_X_Alu Then Lo_TPLiquid.ListRows(F_Liq - 1).Range(CLiq_Diff) = Pago_X_Alu - Rec_Emi
                Rec_Emi = .Range(CLiq_Imp_Rec)
                Pago_X_Alu = .Range(CLiq_Imp_Cob)
                Ant_DNI = .Range(CLiq_DNI)
                If ColorFondo = 35 Then ColorFondo = 37 Else ColorFondo = 35
            End If
            .Range.Interior.ColorIndex = ColorFondo
            '- Resalta las celdas si es un recibo a Aplazar o Aplazado
            If Val("20" & .Range(CLiq_AñoVto)) > Format(.Range(CLiq_F_Emi), "yyyy") And _
               Format(.Range(CLiq_F_Cobro), "yyyy") <> ACont_1 Then
                .Range(CLiq_Núm_Rec).Interior.ColorIndex = 13
                .Range(CLiq_Núm_Rec).Font.ColorIndex = 6
                .Range(CLiq_JI_Emi).Interior.ColorIndex = 13
                .Range(CLiq_JI_Emi).Font.ColorIndex = 6
                .Range(CLiq_AD_0010).Interior.ColorIndex = 13
                .Range(CLiq_AD_0010).Font.ColorIndex = 6
                .Range(CLiq_JI_443).Interior.ColorIndex = 13
                .Range(CLiq_JI_443).Font.ColorIndex = 6
            End If
        
        End With
    
    Next
    Lo_TPLiquid.ListRows(F_Liq - 1).Range(CLiq_Rec_Emi) = Rec_Emi
    Lo_TPLiquid.ListRows(F_Liq - 1).Range(CLiq_Pag_X_Alu) = Pago_X_Alu
    
    
Restablecer_Valores:
Rut_On_Functions
'    IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ------------------------------------------------------
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Wk_TitP_Liquid.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
End Sub     ' Rut_Calc_SubTotales     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
Sub Rut_Copy_Liquid_PDF()
' ==================================================================================================================================
Dim Lo_TPLiquid        As ListObject
Set Lo_TPLiquid = Wk_TitP_Liquid.ListObjects(1)
Dim Lo_TPLiquidPDF     As ListObject
Set Lo_TPLiquidPDF = Wk_TitP_LiqPDF.ListObjects(1)

    Wk_TitP_Liquid.Unprotect
    Call Rut_Lo_DataBodyRange_Copy(Lo_TPLiquid, Lo_TPLiquidPDF, True)
    
    With Wk_TitP_LiqPDF
        .Range("b2") = Wk_TitP_Liquid.Range("Liquid_Plan_Name")
        
        .Columns().EntireColumn.Hidden = False
        .Columns("f").EntireColumn.Hidden = True
        .Columns("l:p").EntireColumn.Hidden = True
    End With
    
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------


' ==================================================================================================================================
Sub Rut_Detectar_Microcredencial(C_Acad As String, Plan As String)
' ==================================================================================================================================
Dim rowfind            As Variant
Dim Lo_Data        As ListObject
Set Lo_Data = Prog_MicroCred.ListObjects(1)
    ' -----------------=============  Buscar Tipo Plan  ==================--------------------------------------------------------------
    rowfind = Application.Match(C_Acad & " " & Plan, Lo_Data.DataBodyRange.Columns(1), 0)
    If Not IsError(rowfind) Then    ' Plan Encontrado ==>> Tendrá características ESPECIALES ------------------------
        Range("APP_PlanMicroCred") = Lo_Data.ListColumns("Cod_Dto").DataBodyRange(rowfind)
            MsgBx_Msg = "Plan: " & Plan & ", Curso Académico: " & C_Acad & vbLf & vbLf & _
                        Lo_Data.ListColumns("Nombre Dto").DataBodyRange(rowfind) & ", Cod.Dto.: " & Lo_Data.ListColumns("Cod_Dto").DataBodyRange(rowfind) & vbLf & vbLf & _
                        "Obs.: " & Lo_Data.ListColumns("Obs").DataBodyRange(rowfind)
            MsgBx_Title = "Proceso: Liquidar Plan de estudio."
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK"): Form_MsgBox.Show '- (Font-Size, Red-Border, Buttons, Default-Button, Image)
        'Wk_TitP_Liquid.Range(
    Else
        Range("APP_PlanMicroCred") = ""
    End If
End Sub     '      <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ----------------------------------------------------------------------------------------------------------------------------------


