Attribute VB_Name = "M40_Inf_Contab_Recibos"
'Rev.: 2026-02-15
Option Explicit

'- Genera la Tabla Informe_Contable_de_Recibos
    '- Borrado Tabla explicativa de Tipos de Recibos: Emitido, ADxAplz, Aplazado, EjeAnt y Añejo.
    '- Vaciar Lo_Inf, Añadir Leyenda con AñoCont y Rellenar Lo_inf ----------------------------------------------
    '- Copio la Tabla explicativa de los Tipos de Recibos: Prog_TipoRec.ListObjects(1) ---
    '- Filas EURLE (Esc.Univ. Rel.Lab. Elda), las filtro, las copio al final del Informe y las borro de Lo_Inf.

'==================================================================================================================================
Sub Rut_Generar_Tabla_Inf_Contable_de_Recibos()
'==================================================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim CursoAcad       As String:      CursoAcad = Prog__APP.Range("APP_CursAcad")
    Dim AñoCont_1_CAcad As Integer:     AñoCont_1_CAcad = Left(CursoAcad, 4)
    Dim CursoAcadAnt    As String:      CursoAcadAnt = Prog__APP.Range("APP_C_Acad_Ant")
    Dim TipoCurso       As String:      TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")

    Dim rowfind     As Variant
    Dim RngVisible  As Range
    Dim Lo_BD           As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
        Prog_BD.Unprotect:     Lo_BD.ShowTotals = False
        Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Dim Lo_BD_Filtrada  As ListObject
                
    Dim WsBuffer As Worksheet: Set WsBuffer = ThisWorkbook.Worksheets("Sheet_Buffer")  'hoja fija/oculta
    ' 1)Limpiar anterior tabla en Buffer
    If WsBuffer.ListObjects.Count > 0 Then WsBuffer.ListObjects(1).Delete
    ' 2) Ordenaciones
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_ACont_Emi, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    ' 3) Filtros, Dependen del AñoCont
        Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="<>_ERR_Date_"
        Lo_BD.Range.AutoFilter Field:=BD_ImpAdm, Criteria1:=">=0"
        If AñoCont = AñoCont_1_CAcad Then   '- Sólo Recibos Emi del AñoCont
            Lo_BD.Range.AutoFilter Field:=BD_ACont_Emi, Criteria1:="=" & AñoCont
        End If
    ' 4) Rango visible filtrado (incluye cabeceras)
    On Error Resume Next
        Set RngVisible = Lo_BD.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo SalirLimpiando
        If RngVisible Is Nothing Then GoTo SalirLimpiando
    ' Limpia cualquier tabla previa en Buffer
    If WsBuffer.ListObjects.Count > 0 Then
        WsBuffer.ListObjects(1).Delete  '  .Unlist o .Delete si prefieres quitar formato tabla
    End If
    WsBuffer.Cells.Clear
    ' 6) Copiar solo lo visible
    RngVisible.Copy Destination:=WsBuffer.Range("A1")
    ' 7) Crear nuevo ListObject con los datos filtrados
    Set Lo_BD_Filtrada = WsBuffer.ListObjects.Add( _
                            SourceType:=xlSrcRange, _
                            Source:=WsBuffer.Range("A1").CurrentRegion, _
                            XlListObjectHasHeaders:=xlYes)
    ' 8) Ajustar el Tipo_Rec al AñoContable
    If AñoCont = AñoCont_1_CAcad Then   '- Tendrán Rec. ADxAplz y Sólo Recibos Emi del AñoCont
        Call Rut_Lo_Filtros_Quitar(Lo_BD_Filtrada)
        With Lo_BD_Filtrada
        .Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="=EjeAnt"
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
        End If
        End With
        Call Rut_Lo_Filtros_Quitar(Lo_BD_Filtrada)
    Else
        Call Rut_Lo_Filtros_Quitar(Lo_BD_Filtrada)
        With Lo_BD_Filtrada
        .Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="=ADxAplz"
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Aplazado"
        End If
        End With
        Call Rut_Lo_Filtros_Quitar(Lo_BD_Filtrada)
    End If
    
    ' 9) Llamar a tu rutina que trabaja con ListObject
    
    If AñoCont = AñoCont_1_CAcad Then       '- Estamos en el 1º AñoCont del C_Acad
        If TipoCurso = "EFP" Then
            Call Rut_Rellenar_Tabla_Inf_Contable_de_Recibos(Lo_BD_Filtrada, Sht__Inf_EFP_ACont1_CAcad)
        Else
            Call Rut_Rellenar_Tabla_Inf_Contable_de_Recibos(Lo_BD_Filtrada, Sht__Inf_CFC_ACont1_CAcad)
        End If
    Else                                    '- Estamos en el 2º AñoCont del C_Acad
        If TipoCurso = "EFP" Then
            Call Rut_Rellenar_Tabla_Inf_Contable_de_Recibos(Lo_BD_Filtrada, Sht__Inf_EFP_ACont2_CAcad)
        Else
            Call Rut_Rellenar_Tabla_Inf_Contable_de_Recibos(Lo_BD_Filtrada, Sht__Inf_CFC_ACont2_CAcad)
        End If
    End If
    
    GoTo Salir

    
SalirLimpiando:
    ' Opcional: quitar tabla en Buffer para dejar solo datos o la hoja limpia
    On Error Resume Next
    If Not WsBuffer Is Nothing Then
        If WsBuffer.ListObjects.Count > 0 Then WsBuffer.ListObjects(1).Unlist
        'Si quieres dejar la hoja totalmente vacía:
        WsBuffer.Cells.Clear
    End If
    On Error GoTo 0
    
Salir:
    ' Restaurar estado de Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Call Rut_EnableEvents_Status_Reset
    Application.Speech.Speak "Proceso Terminado"
End Sub '------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
Sub Rut_Rellenar_Tabla_Inf_Contable_de_Recibos(Lo_BD As ListObject, _
                                               Sht_Inf As Worksheet)
' ==================================================================================================================================
Debug.Print "Rut_Rellenar_Tabla_Inf_Contable_de_Recibos"

    Dim ImpEmiCursAnt           As Currency
    Dim ImpEmiCursPos           As Currency
    Dim ImpINSSCursAnt          As Currency
    Dim ImpINSSCursPos          As Currency
    Dim ImpAdmCursAnt           As Currency
    Dim ImpAdmCursPos           As Currency
    Dim ImpAcadCursAnt          As Currency
    Dim ImpAcadCursPos          As Currency
    Dim ImpCobCursAnt           As Currency
    Dim ImpCobCursPos           As Currency
    Dim ImpCobCursINSSAnt       As Currency
    Dim ImpCobCursINSSPos       As Currency
    Dim ImpCobCursAdmAnt        As Currency
    Dim ImpCobCursAdmPos        As Currency
    Dim ImpCobCursAcadAnt       As Currency
    Dim ImpCobCursAcadPos       As Currency

    Dim Concept         As Variant
    Dim Concept2        As String
    Dim Tp_Rec          As String
    Dim TipRec_Cncpt    As String:          TipRec_Cncpt = ""
    Dim DescripciónA    As String
    Dim DescripciónB    As String
    Dim MenúAux_Msg     As String
    Dim APP_AñoCont     As String:          APP_AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim APP_ACont       As String:          APP_ACont = Right(Prog__APP.Range("APP_AñoCont"), 2)
    Dim APP_AcadAnt     As String:          APP_AcadAnt = Prog__APP.Range("APP_C_Acad_Ant")
    Dim APP_AcadPos     As String:          APP_AcadPos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim TipoCurso       As String:          TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")
    Dim CursoAcad       As String:          CursoAcad = Prog__APP.Range("APP_CursAcad")
    Dim AñoCont_1_CAcad As Integer:         AñoCont_1_CAcad = Left(CursoAcad, 4)

    Dim FechCierreCont  As String:          FechCierreCont = "<" & DateAdd("d", 1, Prog__APP.Range("APP_FechCierreCont"))
    Dim F_Concept       As Integer
    Dim F_Tp_Rec        As Integer
    
    Dim Lo_Concept      As ListObject:      Set Lo_Concept = Prog_Concept.ListObjects("Tb_Conceptos")
    Dim Lo_Tipo_Rec     As ListObject:      Set Lo_Tipo_Rec = Prog_TipoRec.ListObjects(1)   '- "Tb_TipoRec"
    
    Dim Lo_Inf          As ListObject:      Set Lo_Inf = Sht_Inf.ListObjects(1)
    
    Dim RwInf            As ListRow
    
    Dim UltFila     As Integer:     UltFila = Lo_Inf.TotalsRowRange.Row + 2
    Dim TRows_LoTipoRec      As Integer:     TRows_LoTipoRec = Lo_Tipo_Rec.ListRows.Count + 2

    Application.ScreenUpdating = False
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    
    Dim Sht_BD      As Worksheet:       Set Sht_BD = Lo_BD.Parent
    Sht_BD.Visible = xlSheetVisible
    Sht_BD.Unprotect
    Lo_BD.ShowTotals = False
    Prog_Concept.Visible = xlSheetHidden
    Prog_TipoRec.Visible = xlSheetHidden
    Sht_Inf.Visible = xlSheetVisible
    Sht_Inf.Select
    '- Borrado Tabla explicativa de Tipos de Recibos: Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. --------------------------
    Sht_Inf.Range(Cells(UltFila, 1), Cells(UltFila + TRows_LoTipoRec + 20, 1)).EntireRow.Delete
    
    '------------ Preparo Sht_BD y Ordeno por Tipo_Tasa y Concepto_Económico -----------------------------------
'    Call Rut_Lo_Filtros_Quitar(Lo_Inf)
    Call Rut_Lo_WrkSht_Preparar(Sht_Inf)
    Call Rut_Lo_WrkSht_Preparar(Sht_BD)
    Call Rut_Lo_Sort(Lo_Concept, 1, xlAscending, True)          '- Ordeno por el número de Aplicación-Concepto 1303, 1310, 1311...
    Call Rut_Lo_Sort(Lo_BD, BD_Tipo_Rec, xlAscending, True)     '- Emitida, Aplazado, EjeAnt, ADxAplz, Añeja...
    Call Rut_Lo_Sort(Lo_BD, BD_Concepto, xlAscending, False)   '- 1303.00 1310.00 1311.00 Etc.
    
    Sht_Inf.Select
    Lo_Inf.ShowTotals = False
    
    '- Vaciar Lo_Inf, Añadir Leyenda con AñoCont y Rellenar Lo_inf ----------------------------------------------
    If Not Lo_Inf.DataBodyRange Is Nothing Then Lo_Inf.DataBodyRange.Delete

    '-------------- Relleno Tabla para JI's de Tasas por Tipo y por Concepto ------------------------------------
    With Lo_Inf.HeaderRowRange
        .Cells(InfRec_JI_Emi_Adm) = "JI-" & APP_AñoCont - 1 & " o JI-" & APP_AñoCont & vbLf & "1303.00" & vbLf & "Adm."
        .Cells(InfRec_JI_Emi_Acad) = "JI-" & APP_AñoCont - 1 & " o JI-" & APP_AñoCont & vbLf & "1310.xx 00/01/02" & vbLf & "Acad."
        .Cells(InfRec_AD_Emi_Adm) = "AD-" & APP_AñoCont & vbLf & "1303.00" & vbLf & "Adm."
        .Cells(InfRec_AD_Emi_Acad) = "AD-" & APP_AñoCont & vbLf & "1310.xx 00/01/02" & vbLf & "Acad."
        .Cells(InfRec_JI_443_Adm) = "JI-" & APP_AñoCont & vbLf & "1303.00" & vbLf & "Adm."
        .Cells(InfRec_JI_443_Acad) = "JI-" & APP_AñoCont & vbLf & "1310.xx 00/01/02" & vbLf & "Acad."
    
        .Cells(InfRec_ConcptEco2).Offset(-2, 0) = "Año Contable " & APP_AñoCont & ", Cursos Académicos " & APP_AcadAnt & " y " & APP_AcadPos & ", (importes académicos)"
        .Cells(InfRec_ImpAcad_Crs_Ant) = "Emi.'" & APP_ACont & vbLf & "Curso" & vbLf & APP_AcadAnt
        .Cells(InfRec_ImpAcad_Crs_Pos) = "Emi.'" & APP_ACont & vbLf & "Curso" & vbLf & APP_AcadPos
        
        .Cells(InfRec_ImpAcad_EmiAnt) = "Emi.'" & APP_ACont - 1 & vbLf & "No Cob.'" & APP_ACont - 1 & vbLf & "JI-" & APP_AñoCont - 1
        .Cells(InfRec_ImpAcad_Emi) = "Emi.'" & APP_ACont & vbLf & "No Cob.'" & APP_ACont & vbLf & "Vto.'" & APP_ACont + 1
        .Cells(InfRec_ImpAcad_Cob_AcadAnt) = "Cob.'" & APP_ACont & vbLf & "Curso" & vbLf & APP_AcadAnt
        .Cells(InfRec_ImpAcad_Cob_AcadPos) = "Cob.'" & APP_ACont & vbLf & "Curso" & vbLf & APP_AcadPos
        .Cells(InfRec_ImpAcad_Pdte) = vbLf & "Pdte. Cob."
        .Cells(InfRec_ADxAplz) = "Emi. y No Cob.'" & APP_ACont & vbLf & "Vto.'" & APP_ACont + 1 & vbLf & "AD-" & APP_AñoCont
        .Cells(InfRec_Aplazado) = "AD'" & APP_ACont - 1 & " Cob.'" & APP_ACont & vbLf & "JI-" & APP_AñoCont & vbLf & "A la 4430"
    End With
        
    '- Para cada Tipo de Recibos; Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. -----------------------------------
    For F_Tp_Rec = 1 To Lo_Tipo_Rec.ListRows.Count
        
        Tp_Rec = Lo_Tipo_Rec.DataBodyRange.Cells(F_Tp_Rec, 1)
        DescripciónA = Lo_Tipo_Rec.DataBodyRange.Cells(F_Tp_Rec, 2)
        
        '- Para cada Concepto Económico= Tipo de Enseñanza; 1310.00, 1310.01, 1311.00, 1303.00 .. (TIO, EFP, CFC/AF y ADM.) -----------------------------
        For F_Concept = 1 To Lo_Concept.ListRows.Count
            
            Concept = Lo_Concept.DataBodyRange.Cells(F_Concept, 1)
            If TipRec_Cncpt <> (Tp_Rec & Concept) Then
                Set RwInf = Lo_Inf.ListRows.Add
                TipRec_Cncpt = Tp_Rec & Concept
            End If
            Concept2 = Lo_Concept.DataBodyRange.Cells(F_Concept, 5)     '- 303.00, 303.01, 310.00, 310.01, 310.02, 311.00, 311.03, 312.00, 312.02, 315.00
            RwInf.Range(InfRec_TipRec) = Tp_Rec                             '- ADxAplz, Añejo, EjeAnt, Emitido, Aplazado
            RwInf.Range(InfRec_Enseñanza) = Lo_Concept.DataBodyRange.Cells(F_Concept, 4)    '- TIO, EFP, TNCT, CFC, ADM
            RwInf.Range(InfRec_ConcptEco) = Lo_Concept.DataBodyRange.Cells(F_Concept, 5)
            RwInf.Range(InfRec_ConcptEco2) = RwInf.Range(InfRec_ConcptEco)
            RwInf.Range(InfRec_Cta_Adm) = Lo_Concept.DataBodyRange.Cells(F_Concept, 6)      '- Cta. Adm
            DescripciónB = Lo_Concept.DataBodyRange.Cells(F_Concept, 2)
            RwInf.Range(InfRec_ConcptNom) = DescripciónB
            
            With Sht_BD.ListObjects(1).DataBodyRange
            '-----------------------------------------------------------------------------------------------------------------------
            '----- Primeras 12 columnas de la Tabla; Imp_Rec, Imp_INSS, Imp_Adm e Imp_Acad -----------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- de cada Tipo de Recibos; Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. ------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Recibos Emitidos: Acad. + Adm. ----------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe Recibos Emitidos --------------------------------------------------------------------------------------------
            RwInf.Range(InfRec_Tot_Emi) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Importe Recibos Emitidos Cobrado ------------------------------------------------------------------------------------
            RwInf.Range(InfRec_Tot_Cob) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                    .Columns(BD_ImpCob), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FCob), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Saldo Recibos Emitidos Pendiente ----------------------------------------------------------------------------------
            RwInf.Range(InfRec_Tot_Pdte) = RwInf.Range(InfRec_Tot_Emi) - RwInf.Range(InfRec_Tot_Cob)

            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Adm.: Seg.Obl. INSS ---------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe INSS ---------------------------------------------------------------------------------------------------------
            RwInf.Range(InfRec_Adm_INSS_Emi) = Application.SumIfs(.Columns(BD_Rec_Imp_INSS), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Importe INSS. Cobrado ------------------------------------------------------------------------------------------------
            RwInf.Range(InfRec_Adm_INSS_Cob) = Application.SumIfs(.Columns(BD_Rec_Imp_INSS), _
                                                    .Columns(BD_ImpCob), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FCob), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Saldo INSS. Pendiente ----------------------------------------------------------------------------------------------
            RwInf.Range(InfRec_Adm_INSS_Pdte) = RwInf.Range(InfRec_Adm_INSS_Emi) - RwInf.Range(InfRec_Adm_INSS_Cob)

            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Adm. SIN Imp.INSS -----------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe Adm. SIN INSS -----------------------------------------------------------------------------------------------
            RwInf.Range(InfRec_Adm_Emi) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            RwInf.Range(InfRec_Adm_Emi) = RwInf.Range(InfRec_Adm_Emi) - RwInf.Range(InfRec_Adm_INSS_Emi)
            '- Importe Adm. Cobrado ------------------------------------------------------------------------------------------------
            RwInf.Range(InfRec_Adm_Cob) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_ImpCob), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FCob), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            RwInf.Range(InfRec_Adm_Cob) = RwInf.Range(InfRec_Adm_Cob) - RwInf.Range(InfRec_Adm_INSS_Cob)
            '- Saldo Imp. Adm. SIN INSS --------------------------------------------------------------------------------------------
            RwInf.Range(InfRec_Adm_Pdte) = RwInf.Range(InfRec_Adm_Emi) - RwInf.Range(InfRec_Adm_Cob)
           
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Académicos SIN Adm. y SIN INSS ----------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            If RwInf.Range(InfRec_Enseñanza) <> "ADM" And RwInf.Range(InfRec_Enseñanza) <> "TNCT" Then
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwInf.Range(InfRec_Acad_Emi) = RwInf.Range(InfRec_Tot_Emi) - RwInf.Range(InfRec_Adm_Emi) - RwInf.Range(InfRec_Adm_INSS_Emi)
                '- Imp. Acad. Cobrado ----------------------------------------------------------------------------------------------
                RwInf.Range(InfRec_Acad_Cob) = RwInf.Range(InfRec_Tot_Cob) - RwInf.Range(InfRec_Adm_Cob) - RwInf.Range(InfRec_Adm_INSS_Cob)
                '- Saldo Imp. Acad. ------------------------------------------------------------------------------------------------
                RwInf.Range(InfRec_Acad_Pdte) = RwInf.Range(InfRec_Acad_Emi) - RwInf.Range(InfRec_Acad_Cob)
            Else
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwInf.Range(InfRec_Adm_Emi) = RwInf.Range(InfRec_Tot_Emi)
                '- Imp. Acad. Cobrado ----------------------------------------------------------------------------------------------
                RwInf.Range(InfRec_Adm_Cob) = RwInf.Range(InfRec_Tot_Cob)
                '- Saldo Imp. Acad. ------------------------------------------------------------------------------------------------
                RwInf.Range(InfRec_Adm_Pdte) = RwInf.Range(InfRec_Tot_Pdte)
                
'''                GoTo Sigiente_Concepto
                
            End If
            
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Académicos por Curso_Acad ---------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe Recibos Emitidos x Curso_Acad Ant / Pos ----------------------------------------------------------------------
            ImpEmiCursAnt = Application.SumIfs(.Columns(BD_ImpRec), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_C_Acad), "<>" & APP_AcadPos, _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            ImpEmiCursPos = Application.SumIfs(.Columns(BD_ImpRec), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_C_Acad), "=" & APP_AcadPos, _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Importe Adm. x Curso_Acad Ant / Pos -------------------------------------------------------------------------
            ImpAdmCursAnt = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_C_Acad), "<>" & APP_AcadPos, _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            ImpAdmCursPos = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_C_Acad), "=" & APP_AcadPos, _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            ImpAcadCursAnt = ImpEmiCursAnt - ImpAdmCursAnt
            ImpAcadCursPos = ImpEmiCursPos - ImpAdmCursPos
            
            '- Importe Recibos Cobrados x Curso_Acad Ant / Pos ----------------------------------------------------------------------
            ImpCobCursAnt = Application.SumIfs(.Columns(BD_ImpCob), _
                                                    .Columns(BD_ACont_Cob), "=" & APP_AñoCont, _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_C_Acad), "<>" & APP_AcadPos, _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            ImpCobCursPos = Application.SumIfs(.Columns(BD_ImpCob), _
                                                    .Columns(BD_ACont_Cob), "=" & APP_AñoCont, _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_C_Acad), "=" & APP_AcadPos, _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Importe Adm. Cobrados  x Curso_Acad Ant / Pos -------------------------------------------------------------------------
            ImpCobCursAdmAnt = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_ACont_Cob), "=" & APP_AñoCont, _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_C_Acad), "<>" & APP_AcadPos, _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            ImpCobCursAdmPos = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_ACont_Cob), "=" & APP_AñoCont, _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_C_Acad), "=" & APP_AcadPos, _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            ImpCobCursAcadAnt = ImpCobCursAnt - ImpCobCursAdmAnt
            ImpCobCursAcadPos = ImpCobCursPos - ImpCobCursAdmPos
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Académicos SIN Adm. y SIN INSS ----------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            If RwInf.Range(InfRec_Enseñanza) <> "ADM" And RwInf.Range(InfRec_Enseñanza) <> "TNCT" Then
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwInf.Range(InfRec_ImpAcad_Crs_Ant) = ImpAcadCursAnt
                RwInf.Range(InfRec_ImpAcad_Crs_Pos) = ImpAcadCursPos
                RwInf.Range(InfRec_ImpAcad_Cob_AcadAnt) = ImpCobCursAcadAnt
                RwInf.Range(InfRec_ImpAcad_Cob_AcadPos) = ImpCobCursAcadPos
                RwInf.Range(InfRec_ImpAcad_Pdte) = RwInf.Range(InfRec_Acad_Emi) - ImpCobCursAcadAnt - ImpCobCursAcadPos
            Else
'                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
'                RwInf.Range(InfRec_ImpAcad_Crs_Ant) = ImpEmiCursAnt
'                RwInf.Range(InfRec_ImpAcad_Crs_Pos) = ImpEmiCursPos
                
'''                GoTo Sigiente_Concepto
            
            End If
                
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '----- Resto columnas de la Tabla (Exclusivas de Inf_Peter) ------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            
            '- Importe Recibos Emitido ---------------------------------------------------------------------------------------------
            If RwInf.Range(InfRec_TipRec) = "Emitido" Then

                '- Opto por copiar de las Columnas anteriores para reducir el tiempo de ejecución -----------------------------
                RwInf.Range(InfRec_ImpAcad_Emi) = RwInf.Range(InfRec_Acad_Emi)
'                RwInf.Range(InfRec_ImpAcad_Cob_AcadAnt) = RwInf.Range(InfRec_Acad_Cob)
'                RwInf.Range(InfRec_ImpAcad_Pdte) = RwInf.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            
            End If

            '- Importe Recibos EjeAnt --------------------------------------------------------------------------------------------------------------
            If RwInf.Range(InfRec_TipRec) = "EjeAnt" Then
                RwInf.Range(InfRec_ImpAcad_EmiAnt) = RwInf.Range(InfRec_Acad_Emi)
'                RwInf.Range(InfRec_ImpAcad_Pdte) = RwInf.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            End If
            '- Importe Recibos Añejo --------------------------------------------------------------------------------------------------------------
            If RwInf.Range(InfRec_TipRec) = "Añejo" Then
                RwInf.Range(InfRec_ImpAcad_Emi) = RwInf.Range(InfRec_Acad_Emi)
'                RwInf.Range(InfRec_ImpAcad_Pdte) = RwInf.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            End If
            '- Importe Recibos Aplazado --------------------------------------------------------------------------------------------------------------
            If RwInf.Range(InfRec_TipRec) = "Aplazado" Then
                RwInf.Range(InfRec_ImpAcad_Emi) = RwInf.Range(InfRec_Acad_Emi)
'                RwInf.Range(InfRec_ImpAcad_Pdte) = RwInf.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
                RwInf.Range(InfRec_Aplazado) = RwInf.Range(InfRec_Acad_Cob)
            End If
            '- Importe Recibos ADxAplz --------------------------------------------------------------------------------------------------------------
            If RwInf.Range(InfRec_TipRec) = "ADxAplz" Then
                RwInf.Range(InfRec_ImpAcad_Emi) = RwInf.Range(InfRec_Acad_Emi)
                RwInf.Range(InfRec_ADxAplz) = RwInf.Range(InfRec_Acad_Emi)
            End If
                    

            End With
            
Sigiente_Concepto:

            Select Case RwInf.Range(InfRec_TipRec)
            Case "EjeAnt"
                RwInf.Range(InfRec_Descrip_Contab) = "Añadir línea de Cobro en JI-" & APP_AñoCont - 1 & " del Importe Cobrado. Concepto Eco. " _
                                             & Concept2 & "__" & DescripciónB & ".  "
            Case "Emitido", "Añejo"
                RwInf.Range(InfRec_Descrip_Contab) = "JI-" & APP_AñoCont & " del importe Emitido y línea de cobro del Importe Cobrado. Concepto Eco. " _
                                             & Concept2 & "__" & DescripciónB & ".  "
            
            Case "ADxAplz"
                RwInf.Range(InfRec_Descrip_Contab) = "JI-" & APP_AñoCont & " del importe Emitido y AD-" & APP_AñoCont & ", del mismo Importe. Concepto Eco. " _
                                             & Concept2 & "__" & DescripciónB & ".  "
            
            Case "Aplazado"
                RwInf.Range(InfRec_Descrip_Contab) = "JI-" & APP_AñoCont & " a la Cta. 4430, del importe Cobrado y línea de cobro del mismo Importe. Concepto Eco. " _
                                             & Concept2 & "__" & DescripciónB & ".  "
            End Select
            
            If RwInf.Range(InfRec_TipRec) = "EjeAnt" Then
                RwInf.Range(InfRec_Descrip_JI) = "Añadir línea de Cobro en JI-" & APP_AñoCont - 1 & _
                                             " del Importe Cobrado y de Concepto Eco. " & Concept2 & _
                                             "__" & RwInf.Range(InfRec_Enseñanza) & " - " & DescripciónB & ".  "
            Else
                RwInf.Range(InfRec_Descrip_JI) = "Liq.PPub_" & Concept2 & "__" & Tp_Rec & "_" & APP_AñoCont & _
                                             "__" & RwInf.Range(InfRec_Enseñanza) & " - " & DescripciónB & ".  "
            End If
            
            
            If WorksheetFunction.IsEven(F_Tp_Rec) Then             '- IsEven = Es Par
                RwInf.Range.Interior.ColorIndex = 19
            Else
                RwInf.Range.Interior.ColorIndex = 20
            End If
'            RwInf.Range.RowHeight = 30
            RwInf.Range.EntireRow.AutoFit
            RwInf.Range.VerticalAlignment = xlCenter
            
        Next F_Concept
    Next F_Tp_Rec
    
    '- Redondear todos los valores a 2 decimales ----------------------
    Dim Lin         As Integer, Col     As Integer
    Dim RngCol      As Range, Celda     As Range
    ' Recorrer columnas de la tabla
    For Col = InfRec_Tot_Emi To InfRec_Acad_Pdte
        If Col = InfRec_Cta_Adm Or Col = InfRec_ConcptEco Then GoTo SiguienteColumna
        Set RngCol = Lo_Inf.ListColumns(Col).DataBodyRange
        For Each Celda In RngCol
            If IsNumeric(Celda.Value) Then
                Celda.Value = Round(Celda.Value, 2)
            End If
        Next Celda
SiguienteColumna:
    Next Col
    
    '- Indento la columna de Descripción de JIs -------------
    Lo_Inf.ListColumns(InfRec_Descrip_JI).DataBodyRange.Select
    Selection.InsertIndent 1
    Lo_Inf.ListColumns(InfRec_Descrip_Contab).DataBodyRange.Select
    Selection.InsertIndent 1
    Lo_Inf.ListColumns(InfRec_ConcptNom).DataBodyRange.Select
    Selection.InsertIndent 1
    '- Filtro para mostrar la filas con datos ---------------
'    Lo_Inf.Range.AutoFilter Field:=4, Criteria1:="<>0"
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Inf, InfRec_Tot_Emi, "-")
    Call Rut_Lo_Sort(Lo_Inf, InfRec_TipRec, xlAscending, True)
    Call Rut_Lo_Sort(Lo_Inf, InfRec_ConcptEco, xlAscending, False)
    Lo_Inf.ShowTotals = True
    
    '- ------------------------------------------------------
    '- Copio la Tabla explicativa de los Tipos de Recibos: Prog_TipoRec.ListObjects(1) ---
    Dim ColIni      As Integer:     ColIni = InfRec_Descrip_JI
    Dim ColFin      As Integer:     ColFin = ColIni + 1
    UltFila = Lo_Inf.TotalsRowRange.Row + 3
    TRows_LoTipoRec = Prog_TipoRec.ListObjects(1).ListRows.Count
'    Prog_TipoRec.ListObjects(1).ListColumns(1).Range.Resize(, 2).Copy Destination:=Cells(UltFila , 3) ' Resize to include the second column
    Prog_TipoRec.ListObjects(1).ListColumns(1).Range.Resize(, 2).Copy            '- Resize to include the second column
    Cells(UltFila, ColIni).PasteSpecial Paste:=xlPasteFormats
    Cells(UltFila, ColIni).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    '- Añado Datos de AñoCont y Cursos Acad -----------------
    Cells(UltFila, ColIni + 1) = "Descripción para Recibos del Año Contable " & APP_AñoCont & vbLf & _
                            " Rec. que abarcan el Curso Académico " & APP_AcadAnt & " y " & APP_AcadPos & vbLf & _
                            " Y Son Rec. Emitidos en " & APP_AñoCont & " (Cobrados o NO)," & vbLf & _
                            " O Rec. Cobrados en " & APP_AñoCont & " (Emitidos en cualquier Año =< " & APP_AñoCont & ")"
    '- Cambio el alto de las filas y centro verticalmente ---
    Range(Cells(UltFila, ColIni), Cells(UltFila + TRows_LoTipoRec, ColIni + 1)).Select
            Selection.RowHeight = 80
            Selection.VerticalAlignment = xlCenter
    '- Combino Celdas para la Descripción -------------------
    Range(Cells(UltFila, ColIni + 1), Cells(UltFila + TRows_LoTipoRec, ColFin)).Select
    Selection.Merge True
    '- Bordes a toda la Tabla -------------------------------
    Range(Cells(UltFila, ColIni), Cells(UltFila + TRows_LoTipoRec, ColFin)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Union(Cells(UltFila, ColIni + 1), Range(Cells(UltFila + 1, ColIni), Cells(UltFila + TRows_LoTipoRec, ColFin))).Select
    With Selection.Font
        .Name = "Arial"
        .Size = 12
    End With
    
Finalizar:
    Call Rut_Lo_Filtros_Quitar(Lo_Inf)
    Sht_Inf.Calculate
    Sht_Inf.Range("e2") = "Enseñanzas " & TipoCurso & ", Curso " & CursoAcad & ", año contable " & APP_AñoCont
    Sht_Inf.Range("d3") = "Último Cálculo: " & Format(Now, "dd mmmm yyyy - hh:mm") & vbLf & "Fecha de cierro contable: " & Prog__APP.Range("APP_FechCierreCont")
    Sht_Inf.Range("d3").Select
    ''''Prog__APP.Range("APP_Last_Calc_" & Nom_Inf) = Format(Now(), "dd-mmm-yy hh:mm")
    Lo_BD.ShowTotals = True
    
    Application.ScreenUpdating = True
    Application.Calculation = Sw_Calculation
    Application.DisplayAlerts = True
'    Application.Speech.Speak "Proceso completado", True
    
''    '- Visualizo el progreso ---------------------------------------------------------------------------------------
''    MenúAux_Msg = Format(Now, "hh:mm:ss") & "  Tabla generada." & vbCrLf & _
''        vbCrLf & Format(Now, "hh:mm:ss") & "  Realizado el: " & Date & "  " & "-   Tiempo transcurrido: " & Round(Timer - H_Inicio, 2) & " seg."
''    MsgBox MenúAux_Msg
End Sub     '- Rut_Rellenar_Tabla_Inf_Contable_de_Recibos
' ==================================================================================================================================
'Sub Ejemplo_Selcción_Múltiple()
'    Union(Cells(44 + 4, 4), Range(Cells(44 + 5, 3), Cells(44 + 3 + 5, 11))).Select
'End Sub
'
'
