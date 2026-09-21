Attribute VB_Name = "M04_Asign_Cnpto_Eco_y_Tip_Curso"
' Last Rev. 2026-09-21 12:12
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M04_Asign_Cnpto_Eco_y_Tip_Curso - Concepto economico y tipo de ensenanza
' =================================================================================================
'
' PROPOSITO
'  Clasifica cada recibo en su concepto economico presupuestario (1311.00,
'  1311.03, 1312.02, 1315.00, 1303.01) y en su tipo de ensenanza TIO-EP
'  (EFP, CFC, CFC_UPUA, AFC, TNCT_M013, TNCT_PNB1, PruebasAccesoUni).
'  De esta clasificacion dependen despues los informes contables.
'
' INDICE DE RUTINAS Y FUNCIONES
'  RuT_Determinar_Concepto_Eco_y_Tipo_Curso_ByHand ... Lanzadera manual.
'  RuT_Determinar_Concepto_Eco_y_Tipo_Curso(Lo_Data, Col_Ref, Col_Concepto,
'          Col_TIO_EP, Col_ActivEco, Col_TipoCurso, Col_Plan) ... Principal.
'
' TRAMOS DE PROGRAMACION
'    0. Escribe en el informe la leyenda de los tipos, limpia las dos columnas
'       destino y ordena por ActivEco, TipoCurso y Plan.
'
'    El cuerpo se bifurca segun APP_EFP_o_CFC:
'
'    CASO EFP (libro de Estudios de Formacion Permanente)
'       Un solo trazo: TODA la tabla es '1311.00' / 'EFP'. No hace falta
'       filtrar, porque M02 ya dejo unicamente recibos EFP.
'
'    CASO CFCyAFC (el resto), una pasada por familia:
'       ActivEco = 80                    -> 1315.00 / PruebasAccesoUni
'       AdvancedFilter Tb_CriT_CFC       -> 1311.03 / CFC
'       AdvancedFilter Tb_CriT_TUP       -> 1312.02 / CFC_UPUA (Univ. Permanente)
'       AdvancedFilter Tb_CriT_AFC       -> 1311.03 / AFC
'       AdvancedFilter Tb_CriT_TNCT_M013 -> 1311.03 / TNCT_M013 (acceso >25 anos)
'       AdvancedFilter Tb_CriT_TNCT_PNB1 -> 1303.01 / TNCT_PNB1 (idiomas)
'
'    Control final: cuenta los recibos que se quedaron SIN concepto y, si hay
'    alguno, lo avisa con MsgBox; ademas lo deja anotado en el informe.
'
'    Nota: CFC y AFC comparten concepto economico (1311.03) pero se distinguen
'    en el tipo TIO-EP, que es lo que luego separa los informes.
'
' NOTAS
'  Los rangos de criterios se invocan con Range('Tb_CriT_*') SIN cualificar la
'  hoja (a diferencia de M02, que usa Prog_Filtros_Concepto.Range). Funciona
'  porque la hoja activa es la correcta en ese momento; si se reordena el
'  pipeline puede dar error 1004.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2025-12-23  ¡¡¡  OJO HE MIDIFICADO CONCEPTO ECO. 1303.00 Y NO 1303 = 1030,00   !!!
Option Explicit

            Sub RuT_Determinar_Concepto_Eco_y_Tipo_Curso_ByHand()
                Prog_LsGes04.Unprotect
'                Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)          '- Quita filtros, filas y columnas ocultas
                Call RuT_Determinar_Concepto_Eco_y_Tipo_Curso(Prog_LsGes04.ListObjects(1), BD_Ref, BD_Concepto, BD_Tipo_EP, BD_ActivEco, BD_TipoCurso, BD_Plan)
            End Sub
'- -------------------------------------------------------------------------------------------------
'- Asignar Código Concepto-Eco y Tipo_Ensañanza ----------------------------------------------------
'- -------------------------------------------------------------------------------------------------
Sub RuT_Determinar_Concepto_Eco_y_Tipo_Curso(Lo_Data As ListObject, _
                                                Col_Ref As Integer, _
                                                Col_Concepto As Integer, _
                                                Col_TIO_EP As Integer, _
                                                Col_ActivEco As Integer, _
                                                Col_TipoCurso As Integer, _
                                                Col_Plan As Integer)
Debug.Print ">>> RuT_Determinar_Concepto_Eco_y_Tipo_Curso"
    Dim TxT_Progreso    As String
    Dim rowfind         As Variant
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim TipoCurso       As String:      TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")
    Lo_Data.ShowTotals = False
        
    '- Visualizo el progreso
    TxT_Progreso = Form_Menu.TB_Informe
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "Asignando Concepto Económico y Tipo de Enseñanza (EFP, CFC, AFC, TNCT)" & vbLf & _
                        " EFP: Estudios de Formación Permanente: Master, Especialista y Experto." & vbLf & _
                        " CFC_UPUA (TUP): Programa Univ. para Mayores de la UA. (Univ. Perm.)" & vbLf & _
                        " CFC: Cursos de Formación Contínua" & vbLf & _
                        " AFC: Actividades de Formación Complementaria" & vbLf & _
                        " TNCT_M013: Seminario Orientación Pruebas > 25 años (Secret. de Acceso)" & vbLf & _
                        " TNCT_PNB1: Prueba de competencias idioma extranjero. (C.Sup. Idiomas)" & vbLf

    '- ---------------------------------------------------------------------------------------------
    '- Determinar Concepto Económico y Tipo de Enseñanza TIO-EP ------------------------------------
    '- ---------------------------------------------------------------------------------------------
    'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: Asignación Concepto Económico y Tipo Ensañanza.", 0, , , , , , 2)
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_Lo_Sort(Lo_Data, Col_ActivEco, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_Data, Col_TipoCurso, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_Lo_Sort(Lo_Data, Col_Plan, xlAscending, False)    '- Ordenar primero accelera un montón el borrado -----
            
    
    With Lo_Data
        .DataBodyRange.Columns(Col_Concepto).ClearContents   '- Se supone que está vacía...
        .DataBodyRange.Columns(Col_TIO_EP).ClearContents   '- Se supone que está vacía...
        
        If TipoCurso = "EFP" Then
            '==================================================================
            '=== TipoCurso = "EFP" ============================================
            '==================================================================
            '-Estudios de Formación Permanente: Máster, Especialista, Experto. -------------------
            .DataBodyRange.Columns(Col_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1311.00"
            .DataBodyRange.Columns(Col_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "EFP"
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Right(String(8, "_") & Format(.ListRows.Count, "#,##0"), 8) & " '1311.00'    Reg. EFP: Estudios de Formación Permanente: Master, Especialista y Experto."
            'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " '1311.00'    Reg. EFP: Estudios de Formación Permanente: Master, Especialista y Experto.", 0)
        
        Else
            '==================================================================
            '=== TipoCurso = CFCyAFC ==========================================
            '==================================================================
            '-Filtra Recibos Cod_Activ = 80 - Pruebas de aptitud para acceso a la Universidad ------
            Call Rut_Lo_Filtros_Quitar(Lo_Data)
            .Range.AutoFilter Field:=Col_ActivEco, Criteria1:=80
            rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
            If rowfind > 0 Then
                .DataBodyRange.Columns(Col_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1315.00"
                .DataBodyRange.Columns(Col_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "PruebasAccesoUni"
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1315.00'    Reg. Pruebas Acceso Univ."
                'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " '1315.00'    Reg. Pruebas Acceso Univ.", 0)
            End If
            
            '-Filtra Recibos - CFC - Cursos de Formación Contínua ----------------------------------
            Call Rut_Lo_Filtros_Quitar(Lo_Data)
            .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_CFC")
            rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
            If rowfind > 0 Then
                .DataBodyRange.Columns(Col_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1311.03"
                .DataBodyRange.Columns(Col_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "CFC"
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1311.03'    Reg. CFC: Cursos de Formación Contínua"
                'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " '1311.03'    Reg. CFC: Cursos de Formación Contínua", 0)
            End If
    
            '-Filtra Recibos - CFC-TUP - Cursos de Formación Contínua (TUP) ------------------------
            Call Rut_Lo_Filtros_Quitar(Lo_Data)
            .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_TUP")
            rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
            If rowfind > 0 Then
                .DataBodyRange.Columns(Col_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1312.02"
                .DataBodyRange.Columns(Col_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "CFC_UPUA"
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1312.02'    Reg. CFC_UPUA (TUP): Programa Univ. para Mayores de la UA. (Univ. Permanente)"
                'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " '1312.02'    Reg. CFC_UPUA (TUP): Programa Univ. para Mayores de la UA. (Univ. Permanente)", 0)
            End If
    
            '-Filtra Recibos - AFC - Actividades de Formación Complementaria -----------------------
            Call Rut_Lo_Filtros_Quitar(Lo_Data)
            .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_AFC")
            rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
            If rowfind > 0 Then
                .DataBodyRange.Columns(Col_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1311.03"
                .DataBodyRange.Columns(Col_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "AFC"
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1311.03'    Reg. AFC: Actividades de Formación Complementaria"
                'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " '1311.03'    Reg. AFC: Actividades de Formación Complementaria", 0)
            End If
    
            '-Filtra Recibos - TNCT-M013 - Cursos NO Contabilizables como Títulos Propios Universidad (M013)
            Call Rut_Lo_Filtros_Quitar(Lo_Data)
            .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_TNCT_M013")
            rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
            If rowfind > 0 Then
                .DataBodyRange.Columns(Col_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1311.03"
                .DataBodyRange.Columns(Col_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "TNCT_M013"
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1311.03'    Reg. TNCT_M013: Seminario Orientación Pruebas > 25 años (Secretaría de Acceso)"
                'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " '1311.03'    Reg. TNCT_M013: Seminario Orientación Pruebas > 25 años (Secretaría de Acceso)", 0)
            End If
            
            '-Filtra Recibos - TNCT-PNB1 - Cursos NO Contabilizables como Títulos Propios Universidad (PNB1)
            Call Rut_Lo_Filtros_Quitar(Lo_Data)
            .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_TNCT_PNB1")
            rowfind = .Range.Columns(Col_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
            If rowfind > 0 Then
                .DataBodyRange.Columns(Col_Concepto).SpecialCells(xlCellTypeVisible).Cells.Value = "1303.01"
                .DataBodyRange.Columns(Col_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "TNCT_PNB1"
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " '1303.01'    Reg. TNCT_PNB1: Prueba de competencias idioma extranjero. (Centro Sup. Idiomas)"
                'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & " '1303.01'    Reg. TNCT_PNB1: Prueba de competencias idioma extranjero. (Centro Sup. Idiomas)", 0)
            End If
            
        End If  ' If TipoCurso = "EFP"

        '-------------------------------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------
        
        '-Filtra Recibos Sin Concepto o Tipo -------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        rowfind = Application.WorksheetFunction.CountIf(.DataBodyRange.Columns(Col_Concepto), "")
        If rowfind > 0 Then
            MsgBox "¡¡¡ Recibos SIN identificar Concepto-Eco o Tipo_Curso !!! " & rowfind & "reg.", vbOKOnly + vbExclamation
        End If
    
    End With    '-  Lo_Data

    With Lo_Data.DataBodyRange
        TxT_Progreso = _
            Right(String(8, "_") & Format(Application.WorksheetFunction.CountIf(.Columns(Col_Concepto), ""), "#,##0"), 8) & " 'Not Found'  Reg. SIN Tipo TIO-EP o Concepto Económico." & vbCrLf & _
            Right(String(8, "=") & " " & Format(Lo_Data.ListRows.Count, "#,##0"), 8) & " Reg. en BDatos."
    End With    '- Lo_Data.DataBodyRange
        
        '- Visualizo el progreso -------------------------------------------------------------------
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & TxT_Progreso
        'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", TxT_Progreso, 0)

    Lo_Data.ShowTotals = True
        
Call Rut_Lo_Filtros_Quitar(Lo_Data)
Debug.Print "<<< RuT_Determinar_Concepto_Eco_y_Tipo_Curso"
End Sub


