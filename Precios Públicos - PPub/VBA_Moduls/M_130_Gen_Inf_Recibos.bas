Attribute VB_Name = "M_130_Gen_Inf_Recibos"
'Rev.: 2026-02-15
Option Explicit

'- Genera la Tabla Informe_Recibos
    '- Borrado Tabla explicativa de Tipos de Recibos: Emitido, ADxAplz, Aplazado, EjeAnt y Añejo.
    '- Vaciar Lo_Inf, Añadir Leyenda con AñoCont y Rellenar Lo_inf ----------------------------------------------
    '- Copio la Tabla explicativa de los Tipos de Recibos: Prog_TipoRec.ListObjects(1) ---
    '- Filas EURLE (Esc.Univ. Rel.Lab. Elda), las filtro, las copio al final del Informe y las borro de Lo_Inf.

' ==================================================================================================================================
Sub Rut_Recalcular_Tabla_Inf_Recibos()
' ==================================================================================================================================
Debug.Print "Rut_Recalcular_Tabla_Inf_Recibos"

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

Dim FechCierreCont  As String:          FechCierreCont = "<" & DateAdd("d", 1, Prog__APP.Range("APP_FechCierreCont"))
Dim F_Concept       As Integer
Dim F_Tp_Rec        As Integer

Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
Dim Lo_Concept      As ListObject:      Set Lo_Concept = Prog_Concept.ListObjects("Tb_Conceptos")
Dim Lo_Tipo_Rec     As ListObject:      Set Lo_Tipo_Rec = Prog_TipoRec.ListObjects(1)   '- "Tb_TipoRec"

Dim Nom_Inf         As String:          Nom_Inf = "Inf_Recibos"
Dim Sht_Inf         As Worksheet:       Set Sht_Inf = Sht__Inf_Recibos_TIO
Dim Lo_Inf          As ListObject:      Set Lo_Inf = Sht_Inf.ListObjects(1)

Dim RwJI            As ListRow

Dim UltFila     As Integer:     UltFila = Lo_Inf.TotalsRowRange.Row + 2
Dim TRows_LoTipoRec      As Integer:     TRows_LoTipoRec = Lo_Tipo_Rec.ListRows.Count + 2

    Application.ScreenUpdating = False
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    
    Sht__BD.Visible = xlSheetVisible
    Sht__BD.Unprotect
    Lo_BD.ShowTotals = False
    Prog_Concept.Visible = xlSheetHidden
    Prog_TipoRec.Visible = xlSheetHidden
    Sht_Inf.Visible = xlSheetVisible
    Sht_Inf.Select
    '- Borrado Tabla explicativa de Tipos de Recibos: Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. --------------------------
    Sht_Inf.Range(Cells(UltFila, 1), Cells(UltFila + TRows_LoTipoRec + 20, 1)).EntireRow.Delete
    
    '------------ Preparo Sht__BD y Ordeno por Tipo_Tasa y Concepto_Económico -----------------------------------
'    Call Rut_Lo_Filtros_Quitar(Lo_Inf)
    Call Rut_Lo_WrkSht_Preparar(Sht_Inf)
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
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
            
            If Lo_Concept.DataBodyRange.Cells(F_Concept, 14) = 4 Then GoTo Sigiente_Concepto
            Concept = Lo_Concept.DataBodyRange.Cells(F_Concept, 1)
            If TipRec_Cncpt <> (Tp_Rec & Concept) Then
                Set RwJI = Lo_Inf.ListRows.Add
                TipRec_Cncpt = Tp_Rec & Concept
            End If
            Concept2 = Lo_Concept.DataBodyRange.Cells(F_Concept, 5)     '- 303.00, 303.01, 310.00, 310.01, 310.02, 311.00, 311.03, 312.00, 312.02, 315.00
            RwJI.Range(InfRec_TipRec) = Tp_Rec                             '- ADxAplz, Añejo, EjeAnt, Emitido, Aplazado
            RwJI.Range(InfRec_Enseñanza) = Lo_Concept.DataBodyRange.Cells(F_Concept, 4)    '- TIO, EFP, TNCT, CFC, ADM
            RwJI.Range(InfRec_ConcptEco) = Lo_Concept.DataBodyRange.Cells(F_Concept, 5)
            RwJI.Range(InfRec_ConcptEco2) = RwJI.Range(InfRec_ConcptEco)
            RwJI.Range(InfRec_Cta_Adm) = Lo_Concept.DataBodyRange.Cells(F_Concept, 6)      '- Cta. Adm
            DescripciónB = Lo_Concept.DataBodyRange.Cells(F_Concept, 2)
            RwJI.Range(InfRec_ConcptNom) = DescripciónB
            
            With Sht__BD.ListObjects(1).DataBodyRange
            '-----------------------------------------------------------------------------------------------------------------------
            '----- Primeras 12 columnas de la Tabla; Imp_Rec, Imp_INSS, Imp_Adm e Imp_Acad -----------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- de cada Tipo de Recibos; Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. ------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Recibos Emitidos: Acad. + Adm. ----------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe Recibos Emitidos --------------------------------------------------------------------------------------------
            RwJI.Range(InfRec_Tot_Emi) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Importe Recibos Emitidos Cobrado ------------------------------------------------------------------------------------
            RwJI.Range(InfRec_Tot_Cob) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                    .Columns(BD_ImpCob), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FCob), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Saldo Recibos Emitidos Pendiente ----------------------------------------------------------------------------------
            RwJI.Range(InfRec_Tot_Pdte) = RwJI.Range(InfRec_Tot_Emi) - RwJI.Range(InfRec_Tot_Cob)

            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Adm.: Seg.Obl. INSS ---------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe INSS ---------------------------------------------------------------------------------------------------------
            RwJI.Range(InfRec_Adm_INSS_Emi) = Application.SumIfs(.Columns(BD_Rec_Imp_INSS), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Importe INSS. Cobrado ------------------------------------------------------------------------------------------------
            RwJI.Range(InfRec_Adm_INSS_Cob) = Application.SumIfs(.Columns(BD_Rec_Imp_INSS), _
                                                    .Columns(BD_ImpCob), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FCob), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            '- Saldo INSS. Pendiente ----------------------------------------------------------------------------------------------
            RwJI.Range(InfRec_Adm_INSS_Pdte) = RwJI.Range(InfRec_Adm_INSS_Emi) - RwJI.Range(InfRec_Adm_INSS_Cob)

            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Adm. SIN Imp.INSS -----------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe Adm. SIN INSS -----------------------------------------------------------------------------------------------
            RwJI.Range(InfRec_Adm_Emi) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_ImpRec), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FEmi), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            RwJI.Range(InfRec_Adm_Emi) = RwJI.Range(InfRec_Adm_Emi) - RwJI.Range(InfRec_Adm_INSS_Emi)
            '- Importe Adm. Cobrado ------------------------------------------------------------------------------------------------
            RwJI.Range(InfRec_Adm_Cob) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                    .Columns(BD_ImpCob), ">0", _
                                                    .Columns(BD_Tipo_Rec), "*" & Tp_Rec & "*", _
                                                    .Columns(BD_FCob), FechCierreCont, _
                                                    .Columns(BD_Concepto), Concept)
            RwJI.Range(InfRec_Adm_Cob) = RwJI.Range(InfRec_Adm_Cob) - RwJI.Range(InfRec_Adm_INSS_Cob)
            '- Saldo Imp. Adm. SIN INSS --------------------------------------------------------------------------------------------
            RwJI.Range(InfRec_Adm_Pdte) = RwJI.Range(InfRec_Adm_Emi) - RwJI.Range(InfRec_Adm_Cob)
           
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Académicos SIN Adm. y SIN INSS ----------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            If RwJI.Range(InfRec_Enseñanza) <> "ADM" And RwJI.Range(InfRec_Enseñanza) <> "TNCT" Then
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwJI.Range(InfRec_Acad_Emi) = RwJI.Range(InfRec_Tot_Emi) - RwJI.Range(InfRec_Adm_Emi) - RwJI.Range(InfRec_Adm_INSS_Emi)
                '- Imp. Acad. Cobrado ----------------------------------------------------------------------------------------------
                RwJI.Range(InfRec_Acad_Cob) = RwJI.Range(InfRec_Tot_Cob) - RwJI.Range(InfRec_Adm_Cob) - RwJI.Range(InfRec_Adm_INSS_Cob)
                '- Saldo Imp. Acad. ------------------------------------------------------------------------------------------------
                RwJI.Range(InfRec_Acad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - RwJI.Range(InfRec_Acad_Cob)
            Else
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwJI.Range(InfRec_Adm_Emi) = RwJI.Range(InfRec_Tot_Emi)
                '- Imp. Acad. Cobrado ----------------------------------------------------------------------------------------------
                RwJI.Range(InfRec_Adm_Cob) = RwJI.Range(InfRec_Tot_Cob)
                '- Saldo Imp. Acad. ------------------------------------------------------------------------------------------------
                RwJI.Range(InfRec_Adm_Pdte) = RwJI.Range(InfRec_Tot_Pdte)
                
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
            If RwJI.Range(InfRec_Enseñanza) <> "ADM" And RwJI.Range(InfRec_Enseñanza) <> "TNCT" Then
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwJI.Range(InfRec_ImpAcad_Crs_Ant) = ImpAcadCursAnt
                RwJI.Range(InfRec_ImpAcad_Crs_Pos) = ImpAcadCursPos
                RwJI.Range(InfRec_ImpAcad_Cob_AcadAnt) = ImpCobCursAcadAnt
                RwJI.Range(InfRec_ImpAcad_Cob_AcadPos) = ImpCobCursAcadPos
                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAcadAnt - ImpCobCursAcadPos
            Else
'                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
'                RwJI.Range(InfRec_ImpAcad_Crs_Ant) = ImpEmiCursAnt
'                RwJI.Range(InfRec_ImpAcad_Crs_Pos) = ImpEmiCursPos
                
'''                GoTo Sigiente_Concepto
            
            End If
                
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '----- Resto columnas de la Tabla (Exclusivas de Inf_Peter) ------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            
            '- Importe Recibos Emitido ---------------------------------------------------------------------------------------------
            If RwJI.Range(InfRec_TipRec) = "Emitido" Then

                '- Opto por copiar de las Columnas anteriores para reducir el tiempo de ejecución -----------------------------
                RwJI.Range(InfRec_ImpAcad_Emi) = RwJI.Range(InfRec_Acad_Emi)
'                RwJI.Range(InfRec_ImpAcad_Cob_AcadAnt) = RwJI.Range(InfRec_Acad_Cob)
'                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            
            End If

            '- Importe Recibos EjeAnt --------------------------------------------------------------------------------------------------------------
            If RwJI.Range(InfRec_TipRec) = "EjeAnt" Then
                RwJI.Range(InfRec_ImpAcad_EmiAnt) = RwJI.Range(InfRec_Acad_Emi)
'                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            End If
            '- Importe Recibos Añejo --------------------------------------------------------------------------------------------------------------
            If RwJI.Range(InfRec_TipRec) = "Añejo" Then
                RwJI.Range(InfRec_ImpAcad_Emi) = RwJI.Range(InfRec_Acad_Emi)
'                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            End If
            '- Importe Recibos Aplazado --------------------------------------------------------------------------------------------------------------
            If RwJI.Range(InfRec_TipRec) = "Aplazado" Then
                RwJI.Range(InfRec_ImpAcad_Emi) = RwJI.Range(InfRec_Acad_Emi)
'                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
                RwJI.Range(InfRec_Aplazado) = RwJI.Range(InfRec_Acad_Cob)
            End If
            '- Importe Recibos ADxAplz --------------------------------------------------------------------------------------------------------------
            If RwJI.Range(InfRec_TipRec) = "ADxAplz" Then
                RwJI.Range(InfRec_ImpAcad_Emi) = RwJI.Range(InfRec_Acad_Emi)
                RwJI.Range(InfRec_ADxAplz) = RwJI.Range(InfRec_Acad_Emi)
            End If
                    

            End With
            
Sigiente_Concepto:

            Select Case RwJI.Range(InfRec_TipRec)
            Case "EjeAnt"
                RwJI.Range(InfRec_Descrip_Contab) = "Añadir línea de Cobro en JI-" & APP_AñoCont - 1 & " del Importe Cobrado. Concepto Eco. " _
                                             & Concept2 & "__" & DescripciónB & ".  "
            Case "Emitido", "Añejo"
                RwJI.Range(InfRec_Descrip_Contab) = "JI-" & APP_AñoCont & " del importe Emitido y línea de cobro del Importe Cobrado. Concepto Eco. " _
                                             & Concept2 & "__" & DescripciónB & ".  "
            
            Case "ADxAplz"
                RwJI.Range(InfRec_Descrip_Contab) = "JI-" & APP_AñoCont & " del importe Emitido y AD-" & APP_AñoCont & ", del mismo Importe. Concepto Eco. " _
                                             & Concept2 & "__" & DescripciónB & ".  "
            
            Case "Aplazado"
                RwJI.Range(InfRec_Descrip_Contab) = "JI-" & APP_AñoCont & " a la Cta. 4430, del importe Cobrado y línea de cobro del mismo Importe. Concepto Eco. " _
                                             & Concept2 & "__" & DescripciónB & ".  "
            End Select
            
            If RwJI.Range(JIs_TipRec) = "EjeAnt" Then
                RwJI.Range(InfRec_Descrip_JI) = "Añadir línea de Cobro en JI-" & APP_AñoCont - 1 & _
                                             " del Importe Cobrado y de Concepto Eco. " & Concept2 & _
                                             "__" & RwJI.Range(InfRec_Enseñanza) & " - " & DescripciónB & ".  "
            Else
                RwJI.Range(InfRec_Descrip_JI) = "Liq.PPub_" & Concept2 & "__" & Tp_Rec & "_" & APP_AñoCont & _
                                             "__" & RwJI.Range(InfRec_Enseñanza) & " - " & DescripciónB & ".  "
            End If
            
            
            If WorksheetFunction.IsEven(F_Tp_Rec) Then             '- IsEven = Es Par
                RwJI.Range.Interior.ColorIndex = 19
            Else
                RwJI.Range.Interior.ColorIndex = 20
            End If
'            RwJI.Range.RowHeight = 30
            RwJI.Range.EntireRow.AutoFit
            RwJI.Range.VerticalAlignment = xlCenter
            
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
    


'''''            '-----------------------------------------------------------------------------------------------------------------------
'''''            '- JIs del AñoCont_Ant
'''''            If .Range(BD_JI_Emi_Acad) <> "" Then
'''''                If InStr(RwJI.Range(InfRec_Adm_Pdte), .Columns(BD_JI_Emi_Acad)) Then
'''''                    RwJI.Range(InfRec_Adm_Pdte) = RwJI.Range(InfRec_Adm_Pdte) & "-" & .Columns(BD_JI_Emi_Acad)
'''''                End If
'''''            End If
            
            

'GoTo Finalizar
    
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
    
    '- Filas EURLE (Esc.Univ. Rel.Lab. Elda), las filtro, las copio al final del Informe y las borro de Lo_Inf.
    Dim RowsFind    As Variant
    Dim Rng         As Range
    With Lo_Inf
        .Range.AutoFilter Field:=InfRec_ConcptNom, Criteria1:="EURLE*"       '- Filtrar
        Sht_Inf.Calculate
        RowsFind = .Range.Columns(InfRec_ConcptNom).SpecialCells(xlCellTypeVisible).Cells.Count - 2 '2 + .ShowTotals  '- Si tiene TotalsRowRange .ShowTotals = -1 (True = -1, False = 0)
        If RowsFind > 0 Then
            Set Rng = Range("e2:r2")
            Rng.Copy Destination:=Cells(UltFila + TRows_LoTipoRec + 3, 5)
            Set Rng = .Range.Resize(.Range.Rows.Count, InfRec_Acad_Pdte).SpecialCells(xlCellTypeVisible)
            Rng.Copy Destination:=Cells(UltFila + TRows_LoTipoRec + 4, 2)
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        End If
    End With
    
Finalizar:
    Call Rut_Lo_Filtros_Quitar(Lo_Inf)
    Sht_Inf.Calculate
    Sht_Inf.Range("d2") = "Último Cálculo: " & Format(Now, "dd mmmm yyyy - hh:mm") & vbLf & "Fecha de cierro contable: " & Prog__APP.Range("APP_FechCierreCont")
    Sht_Inf.Range("d2").Select
    Prog__APP.Range("APP_Last_Calc_" & Nom_Inf) = Format(Now(), "dd-mmm-yy hh:mm")
    
''    '- Visualizo el progreso ---------------------------------------------------------------------------------------
''    MenúAux_Msg = Format(Now, "hh:mm:ss") & "  Tabla generada." & vbCrLf & _
''        vbCrLf & Format(Now, "hh:mm:ss") & "  Realizado el: " & Date & "  " & "-   Tiempo transcurrido: " & Round(Timer - H_Inicio, 2) & " seg."
''    MsgBox MenúAux_Msg
    Lo_BD.ShowTotals = True
    
    Application.ScreenUpdating = True
    Application.Calculation = Sw_Calculation
    Application.DisplayAlerts = True
    Application.Speech.Speak "Proceso completado", True
End Sub     '- Rut_Recalcular_Tabla_Inf_Recibos
' ==================================================================================================================================
'Sub Ejemplo_Selcción_Múltiple()
'    Union(Cells(44 + 4, 4), Range(Cells(44 + 5, 3), Cells(44 + 3 + 5, 11))).Select
'End Sub



