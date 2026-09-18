Attribute VB_Name = "M01_Importar_LsGes04_GE"
'2026-01-14
'- M02_Importar_LSGES04_GE
Option Explicit

'- Seleccionar fichero Excel LSGES04 a importar
'- Importar el Excel LSGES04 en un WorkBook_Close
'       - Si no viene con Tabla la Creo
'- Borrar Recibos NO Pertinentes:  Curso-Acad, Matrícula=N, AE<>4
'- Formatear la Tabla de Prog_LsGes04
'- Recorro toda la Tabla Prog_LsGes04 para actualizar Prog_BD (BD_Hist)
'       - Añado Todos los registros NUEVOS de TLSGES04 en BD_Hist
'   - Los registros que no están el TLGES04 y sí en BD_Hist, los marco como eliminados
'   - De los registros coindidentes:
'       - Informe de posibles cambios: ImpRec, ImpCob, ImpAmd, MatAnulada
'       - Copio los datos de TLSGES04 en BD_Hist
'- Actualiza en BD_Hist el Imp. Tasa Adm., Sólo si no tiene ya un Importe.
'- Incorporar Concepto Económico y Tipo de Enseñanza TIO-EP (EFP, CFC, AFC, TNCT)
'- Borrar Prog_LsGes04 menos los DUPLICADOS para su posible control

Function Func_Informe(Tx1 As String, Optional Tx2 As String = "", Optional Tx3 As String = "") As String
    Dim Texto As String
    Texto = Left(Tx1 & " " & String(56, "·"), 56)
    Texto = Texto & Right(String(25, "·") & " " & Tx2, 25)
    Func_Informe = Texto & Right(String(25, "·") & "·" & Tx3, 25)
End Function

'==================================================================================================================================
Sub Mod_Importar_LSGES04_GE()   '- Importar Última Consulta de LSGES04_GE, para Actualizar registros existentes y Añadir Nuevos.
'==================================================================================================================================
Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
Dim Curso_Acad      As String:      Curso_Acad = Prog__APP.Range("APP_CursAcad")
Dim TxT_Progreso        As String
Dim TxtMsg1  As String, TxtMsg2  As String, TxtMsg3  As String
Dim Ref_Ant             As String:      Ref_Ant = ""
Dim RwG4            As ListRow
Dim RwPH            As ListRow
Dim RngG4           As Range
Dim RngPH           As Range
Dim rowfind         As Variant

Dim Lo_Ges04            As ListObject:      Set Lo_Ges04 = Prog_LsGes04.ListObjects(1)
Dim Lo_Ges04_DefCol     As ListObject:      Set Lo_Ges04_DefCol = Prog_DefCol_BD.ListObjects(1)

Call Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)
    
    '- Seleccionar fichero Excel LSGES04 e importar en ClsBook (RAM) ------------------------------------------------------------------------
        Dim Arch__EP_New        As String
        Dim Nom_NewArch         As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = Application.Workbooks(ThisWorkbook.Name).Path & "\" & "LSGES04_GE_SinDtos_Curso_" & Curso_Acad & "*"
        .Title = "Seleccionar el Fichero Excel de la última consulta LSG4_GE de un PLAN ÚNICO del Generador de Informes: "
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        .ButtonName = "Seleccionar"
        .Filters.Clear
        .Filters.Add "Sólo Ficheros Excel", "*.xls?", 1
        If .Show <> -1 Then
            MsgBox "Cancelado", , "Rutinas"
            Form_Menu.TB_Informe = "Proceso Cancelado: " & Now()
            GoTo Restablecer_Valores
        Else
            Arch__EP_New = .SelectedItems(1)
'            Nom_NewArch = Dir(Arch__EP_New) '- Falla con NEXE
            Nom_NewArch = Right(Arch__EP_New, Len(Arch__EP_New) - InStrRev(Arch__EP_New, "/"))
            'Path_NewArch = Left(Arch__EP_New, InStrRev(Arch__EP_New, "\"))
        End If
    End With
    '- Importar en ClsBook (RAM)   --------------------------------------------------------------
    Dim ClsBk      As Workbook
    Set ClsBk = Workbooks.Open(Arch__EP_New, ReadOnly:=True)
    Dim Ws_ClsBk    As Worksheet
    Set Ws_ClsBk = ClsBk.Sheets(1)
    Dim Lo_ClsBk    As ListObject
    
    '   Si no viene con Tabla la Creo
    If Ws_ClsBk.ListObjects.Count = 0 Then
        Ws_ClsBk.UsedRange.Cells(1, 1).Select  'Posicionar cursor
        Set Lo_ClsBk = Ws_ClsBk.ListObjects.Add(xlSrcRange, Ws_ClsBk.UsedRange, , xlYes)
    Else
        Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
    End If
        rowfind = Lo_ClsBk.ListRows.Count
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & Nom_NewArch & vbLf
            TxtMsg1 = Now & " -  Importado Excel"
'            TxtMsg2 = ""
            TxtMsg3 = " tiene " & Format(rowfind, "#,##0") & "reg."
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, , TxtMsg3) & vbLf
    
    '- Borrar recibos de EFP o CFCyAFC según se requiera ----------------------
    Call Rut_Borrar_Rec_EFP_o_CFCyAFC(Lo_ClsBk)
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf

    '- Borrar Recibos NO Válidos:  otros C_Acad, Matrícula=N, AE<>4 ---------------------------------------------
    Call RuT_Del_Reg_NO_Válidos(Lo_ClsBk)
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf
    If Lo_ClsBk.DataBodyRange Is Nothing Then
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & "¡¡¡ El Excel seleccionado, NO tiene registros procesables AE=4 !!!"
        ClsBk.Close SaveChanges:=False
        Set ClsBk = Nothing
        GoTo Restablecer_Valores
    End If
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Copy ClsBk:
    '-            1º vaciar BD_LSGes04
    '-            2º Copiar Lo_ClsBk en Lo_Ges04.
    '- --------------------------------------------------------------------------------------------------------------
    Prog_LsGes04.Visible = xlSheetVisible
    Prog_LsGes04.Unprotect
    
    '- Borrar Lo_Ges04 y Copiar Lo_ClsBk en Lo_Ges04. ---------------------
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_Ges04, True)
    rowfind = Lo_Ges04.ListRows.Count
    ClsBk.Close SaveChanges:=False
    Set ClsBk = Nothing
            TxtMsg1 = Format(Now, "hh:mm:ss") & " Copiado Excel a Lo_Ges04:"
            TxtMsg2 = ""
            TxtMsg3 = Format(rowfind, "#,##0") & "reg."
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
    Prog_LsGes04.Select
        
    '- --------------------------------------------------------------------------------------------------------------
    '- Process Lo_LSGes04:
    '        - Formatear Lo_Ges04
    '        - M02_Manage_Duplicates
    '        - Asignar Año de Vencimiento Rec. en ACont_Vto
    '        - Asignar Col Cta_Ingreso con nº Cta. correspondiente
    '        - Asignar Código Concepto-Eco y Tipo_Ensañanza: 1310.00, 1311.03... EFP, CFC, TNCT, UPUA...
    '        - Asignar Tipo de Recibo: Emitido, EjeAnt, Añejo, ADxAplz o Aplazado
    '-
    '- --------------------------------------------------------------------------------------------------------------
    '- Formatear la Tabla de Lo_Ges04 ---------------------------------------------------------------------------------------------
    Call Rut_X_Format_LoData_LoDefCol(Lo_Ges04, Lo_Ges04_DefCol)
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Now, "hh:mm:ss") & " Formateado Lo_Ges04." & vbLf
            Form_Menu.TB_Informe.SelStart = Len(Form_Menu.TB_Informe)

    '- M02_Manage_Duplicates -------------------------------------------------------------------------------------------------------
    Dim Lo_BD_Dupl      As ListObject:      Set Lo_BD_Dupl = Prog_BD_Dupl.ListObjects(1)
    Dim Lo_DefCol_BD    As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
            Prog_BD_Dupl.Visible = xlSheetVisible
            Prog_BD_Dupl.Unprotect
            Call Rut_Lo_WrkSht_Preparar(Prog_BD_Dupl)
            Lo_BD_Dupl.ShowTotals = False
    Call RuT_Duplicates_Search(Lo_Ges04, Lo_DefCol_BD, Lo_BD_Dupl, BD_Ref, BD_Incidencias, BD_H_Incidencias)
        Set Lo_BD_Dupl = Nothing
    
    '- ¡¡¡ Modifico El ACont_Cob si ACont_Emi > ACont_Cob  ==>>  ACont_Cob = ACont_Emi !!! ----------------------------------------
    '- y Asignar Año de Vencimiento Rec. en ACont_Vto -----------------------------------------------------------------------------
    '----- Si el AñoVto no es correcto falla en Asignar Tipo de Recibo (Emitido, Aplazado...) -------------------------------------
    Dim i As Long
    Dim Cont    As Long
    With Lo_Ges04.DataBodyRange
        For i = 1 To Lo_Ges04.ListRows.Count
            .Cells(i, BD_ACont_Vto) = Format(.Cells(i, BD_FVto), "yyyy")
            If .Cells(i, BD_ACont_Vto) < .Cells(i, BD_ACont_Emi) Then .Cells(i, BD_ACont_Vto) = .Cells(i, BD_ACont_Emi)
            If .Cells(i, BD_ACont_Emi) > .Cells(i, BD_ACont_Cob) And Len(.Cells(i, BD_ACont_Cob)) > 0 Then
                .Cells(i, BD_ACont_Cob) = .Cells(i, BD_ACont_Emi)
                Cont = Cont + 1
            End If
        Next
    End With
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & vbCrLf & Format(Now, "hh:mm:ss") & " Añadido ACont_Vto." & vbCrLf
    If Cont > 0 Then
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Now, "hh:mm:ss") & " Modificado ACont_Cob = Acont_Emi, porque ¡¡¡  ACont_Cob < ACont_Emi !!!  en " & Cont & "reg." & vbCrLf
    End If
    
    '- Asignar Col Cta_Ingreso con nº Cta. correspondiente ----------------------------------------------------------------------
    Call RuT_Determinar_Cta_Ingreso(Lo_Ges04, BD_Ref, BD_CtaPag, BD_Cta_Ing)
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Now, "hh:mm:ss") & " Añadidas Cta. de ingreso." & vbCrLf
    
    '- Asignar Código Concepto-Eco y Tipo_Ensañanza: 1310.00, 1311.03... EFP, CFC, TNCT, UPUA... --------------------------------
    Call RuT_Determinar_Concepto_Eco_y_Tipo_Curso(Lo_Ges04, BD_Ref, BD_Concepto, BD_Tipo_EP, BD_ActivEco, BD_TipoCurso, BD_Plan)
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                                    " Añadido Concepto Económico y Tipo de Enseñanza." & vbCrLf
            Form_Menu.TB_Informe.SelStart = Len(Form_Menu.TB_Informe)
            Form_Menu.TB_Informe.SetFocus

    '- Asignar Tipo de Recibo: Emitido, EjeAnt, Añejo, ADxAplz o Aplazado -------------------------------------------------------
    Call RuT_Determinar_Tipo_Recibo
            Form_Menu.TB_Informe.SelStart = Len(Form_Menu.TB_Informe)
            Form_Menu.TB_Informe.SetFocus
    
    '- Identificar del C_Acad, los 1º Rec. de c/matrícula para obtener la T-Adm -------------------------------------------------
    Call Rut_Assign_Imp_AdmAcad_C_Acad
            Form_Menu.TB_Informe.SelStart = Len(Form_Menu.TB_Informe)
            Form_Menu.TB_Informe.SetFocus
    
saltar_Aqui:
    
    '- -------------------------------------------------------------------------------------------------------------------------
    '- Process Lo_BDatos:
    '- -------------------------------------------------------------------------------------------------------------------------
    '- Actualizar BDatos con LsGes04 -------------------------------------------------------------------------------------------
    Call RuT_Actualizar_BDatos_con_LsGes04
            Form_Menu.TB_Informe.SelStart = Len(Form_Menu.TB_Informe)
            Form_Menu.TB_Informe.SetFocus

    '- Actualiza la Tabla de Referencia de los Coeficientes de Retención para el VRI ---------------------------------------
    Call RuT_Lo_Coef_VRI_Actualizar
            Form_Menu.TB_Informe.SelStart = Len(Form_Menu.TB_Informe)
            Form_Menu.TB_Informe.SetFocus

            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False


Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf
GoTo Restablecer_Valores

'''    Dim Lo_BD               As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
'''    Dim Lo_Tb_Ret_VRI       As ListObject:      Set Lo_Tb_Ret_VRI = Prog_Coef_Ret_VRI.ListObjects(1)
'''
'''
'''
'''
'''    '- Recorro toda la Tabla Prog_LsGes04 para actualizar Prog_BD (BD_HIST) ----------------------------------------------------------
'''        '   Añado Todos los registros nuevos y marco los registros eliminados
'''    Dim F_Actualiz      As String:  F_Actualiz = Now()
'''    Dim Incidencia      As String
'''    Dim Cont            As Integer
'''    Dim F_PH            As Long:    F_PH = 1
'''    Dim F_G4            As Long:    F_G4 = 1
'''    Dim TRows_TitPH     As Long:    TRows_TitPH = Lo_BD.ListRows.Count
'''    Dim TRows_Ges04     As Long:    TRows_Ges04 = Lo_Ges04.ListRows.Count
'''    Dim Cont_Fail       As Long:    Cont_Fail = 0
'''    Dim Cont_Repes      As Long:    Cont_Repes = 0
'''    Dim Cont_Nuevo      As Long:    Cont_Nuevo = 0
'''    Dim Cont_Modif      As Long:    Cont_Modif = 0
'''    Dim Cont_Deleted    As Long:    Cont_Deleted = 0
'''    Dim Cont_Mat_Anul   As Long:    Cont_Mat_Anul = 0
'''    Dim Chg_ImpRec      As Long:    Chg_ImpRec = 0
'''    Dim Chg_ImpCob      As Long:    Chg_ImpCob = 0
'''    Dim Chg_ImpAdm      As Long:    Chg_ImpAdm = 0
'''
'''    '- Visualizo el progreso ---------------------------------------------------------------------------------------
'''    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & _
'''            "¡¡¡ Proceso concluido con éxito !!! día: " & Now() & " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
'''            "En la Anterior Consulta habían:  " & Right("__________" & TRows_TitPH, 8) & "  Reg." & vbCrLf & _
'''            "En la Nueva Consulta hay:        " & Right("__________" & TRows_Ges04, 8) & "  Reg.   Con: " & Cont_Repes & " Reg. Repetidos" & vbCrLf & _
'''            Right("__________" & Cont_Nuevo, 8) & "  Reg. nuevos." & vbCrLf & _
'''            Right("__________" & Cont_Modif, 8) & "  Reg. que existían y se han actualizado." & vbCrLf & _
'''            Right("__________" & Cont_Mat_Anul, 8) & "  Tasas Anuladas." & vbCrLf & _
'''            Right("__________" & Chg_ImpRec + Chg_ImpCob + Chg_ImpAdm, 8) & "  Reg. Actualizados con incidencias (Cambios destacables)." & vbCrLf & _
'''            Right("__________" & Chg_ImpRec, 8) & "  Reg. Cambio en Importe de Recibo." & vbCrLf & _
'''            Right("__________" & Chg_ImpCob, 8) & "  Reg. Cambio en Importe Cobrado." & vbCrLf & _
'''            Right("__________" & Chg_ImpAdm, 8) & "  Reg. Cambio en Importe Administrativo." & vbCrLf & vbCrLf & _
'''            "Hay activos ahora un Total de:  " & Lo_BD.DataBodyRange.Rows.Count & "  Reg."
'''
        
Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Wk_TitP_Liquid.Select
'    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    Call Rut_EnableEvents_Status_Reset
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Prog_BD.Protect , AllowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Prog_LsGes04.Protect , AllowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Prog_BD.Visible = xlSheetVeryHidden
'Prog_LsGes04.Visible = xlSheetVeryHidden
    Application.Speech.Speak "Proceso completado."

Rut_On_Functions
End Sub     ' RuT_Importar_LSGES04_GE   --------------------------------------------------------------------------------------------
'===================================================================================================================================

'''' ==================================================================================================================================
'''Sub Rut_Incorporar_Coef_VRI(ByRef Plan As String)
'''' ==================================================================================================================================
'''Dim RowFind            As Variant
'''Dim Lo_Tb_Ret_VRI        As ListObject
'''Set Lo_Tb_Ret_VRI = Prog_Coef_Ret_VRI.ListObjects(1)
'''    ' -----------------=============  Buscar Tipo Plan  ==================--------------------------------------------------------------
'''    RowFind = Application.Match(Plan, Lo_Tb_Ret_VRI.DataBodyRange.Columns(1), 0)
'''    If Not IsError(RowFind) Then    ' Plan Encontrado ==>> Tendrá características ESPECIALES ------------------------
'''        Coef_VRI = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(RowFind)
'''    Else                            ' NO ENCONTRADO
'''        If IsNumeric(Left(Plan, 1)) Then
'''            Coef_VRI = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(2)
'''        Else
'''            Coef_VRI = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(1)
'''        End If
'''    End If
'''    '---------------------------------------------------------------------------------------------------------------------------------------
'''End Sub     '      <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'''' ==================================================================================================================================


'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'---------------------------- Rut de trabajo interno, a eliminar ------------------------------------------------------------------
'---------------------------- Rut de trabajo interno, a eliminar ------------------------------------------------------------------
'---------------------------- Rut de trabajo interno, a eliminar ------------------------------------------------------------------
'---------------------------- Rut de trabajo interno, a eliminar ------------------------------------------------------------------
'==================================================================================================================================
'==================================================================================================================================
Sub RuT_Actualizar_Repetidos()   '- Aparecieron registros repetidos.
'==================================================================================================================================

Dim TxT_Progreso        As String
Dim Lo_TitPH            As ListObject
Dim Lo_Ges04            As ListObject
Dim Lo_Ges04_DefCol     As ListObject
Dim RwPHant            As ListRow
Dim RwPH            As ListRow

Dim rowfind             As Variant


Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso

Set Lo_TitPH = Prog_BD.ListObjects(1)

'   Recorro toda la Tabla Tip-Hist  ----------------------------------------------------------
Dim F_PH            As Long:    F_PH = 1
Dim ContFila        As Long:    ContFila = 0
Dim ContFilDatos    As Long:    ContFilDatos = 0
Dim ContFilDup      As Long:    ContFilDup = 0

Dim Cont            As Integer
Dim Ref_Ant      As String:     Ref_Ant = ""
Dim TRows_TitPH      As Long: TRows_TitPH = Lo_TitPH.ListRows.Count

    Sheets(Prog_BD.Name).Select
    Prog_BD.Unprotect
'    '- Copiar DataBodyRange ----------------------------------------------
''    Call Rut_Lo_DataBodyRange_Copy(Prog_BD1.ListObjects(1), Prog_BD.ListObjects(1), True)
'
    '- Ordenar por PLAN y DNI ==================
    Call Rut_Lo_Sort(Lo_TitPH, BD_Ref, xlAscending, True)
''    '- Datos --> Testo en columnas --> para convertir Todo a Texto -------------
''    Range("Tb_TitP[Plan]").Select
''    Selection.TextToColumns Destination:=Range("a4"), DataType:=xlDelimited, _
''        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
''        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
''        :=Array(1, 2), TrailingMinusNumbers:=True
'''    Range("Tb_TitP[NomPlan]").Select
'''    Selection.TextToColumns Destination:=Range("b4"), DataType:=xlDelimited, _
'''        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'''        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
'''        :=Array(1, 2), TrailingMinusNumbers:=True
''    Range("Tb_TitP[DNI]").Select
''    Selection.TextToColumns Destination:=Range("e4"), DataType:=xlDelimited, _
''        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
''        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
''        :=Array(1, 2), TrailingMinusNumbers:=True
''    Range("Tb_TitP[CTA. CCC]").Select
''    Selection.TextToColumns Destination:=Range("o4"), DataType:=xlDelimited, _
''        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
''        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
''        :=Array(1, 2), TrailingMinusNumbers:=True
''    '- Quitar las Marcas de Error #N/D  ----------------------------------------
''    Range("Tb_TitP[NomPlan]").Select
''    Selection.Replace "#N/A", "", xlWhole
''    Selection.Replace "#N/D", "", xlWhole
''    '- Corregir Num Cta.CCC  ----------------------------------------
''    Range("Tb_TitP[CTA. CCC]").Select
''    Selection.Replace "00496659072416175503", "0049 6659 07 2416175503", xlWhole
''    Selection.Replace "   ", "", xlWhole

'            ContFila = Prog_Dels_TPH.ListObjects(1).ListRows.Count
    '- Recorrer toda la tabla --------------------------
    For F_PH = 1 To Lo_TitPH.ListRows.Count   '--- Bucle para recorrer todas la filas de la Consulta Prog_BD

        Set RwPH = Lo_TitPH.ListRows(F_PH)

        If RwPH.Range(BD_Ref) = Ref_Ant Then

            Set RwPHant = Lo_TitPH.ListRows(F_PH - 1)

'            Lo_TitPH.ListRows(F_PH).Range.Select
                RwPHant.Range(BD_EP_Ctrl) = RwPHant.Range(BD_EP_Ctrl) & "_Duplicaty1"
                RwPH.Range(BD_EP_Ctrl) = RwPH.Range(BD_EP_Ctrl) & "_Duplicaty2"
                ContFilDup = ContFilDup + 1

            For Cont = 1 To Lo_TitPH.Range.Columns.Count
                If RwPH.Range(Cont) <> RwPHant.Range(Cont) Then
                    If Len(RwPHant.Range(Cont).Value) = 0 Then RwPHant.Range(Cont) = RwPH.Range(Cont)
                End If

            Next Cont

            If Len(RwPH.Range(BD_Orgánica) & RwPH.Range(BD_ExpAdm) & RwPH.Range(BD_Liquidado) & RwPH.Range(BD_RDT) & RwPH.Range(BD_JI_Emi_Acad)) > 0 Then
                '''    ContFila = ContFila + 1
                '''    Prog_Dels_TPH.ListObjects(1).ListRows.Add
                '''    RwPH.Range.Copy Prog_Dels_TPH.ListObjects(1).ListRows(ContFila).Range
                RwPHant.Range(BD_EP_Ctrl) = RwPHant.Range(BD_EP_Ctrl) & "_Datos"
                ContFilDatos = ContFilDatos + 1
'                Lo_TitPH.ListRows(F_PH).Delete
                F_PH = F_PH - 1
            End If
            
'            If Len(RwPH.Range(BD_Orgánica) & RwPH.Range(BD_ExpAdm) & RwPH.Range(BD_Liquidado) & RwPH.Range(BD_RDT) & RwPH.Range(BD_JI_Emi_Acad)) = 0 Then
'                ContFilDatos = ContFilDatos + 1
'                Lo_TitPH.ListRows(F_PH).Delete
'                F_PH = F_PH - 1
'            End If
        Else
            Ref_Ant = RwPH.Range(BD_Ref)
        End If


        If F_PH Mod 50 = 0 Then
            Debug.Print "Actualizando Tasas Adm.:" & _
                                        vbCrLf & Format(F_PH, "#,##0") & " de " & Format(TRows_TitPH, "#,##0")
        End If
    Next
Debug.Print ContFila

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Debug.Print ContFilDup, ContFilDatos
Rut_On_Functions
End Sub     ' RuT_Actualizar_Repetidos   --------------------------------------------------------------------------------------------
'===================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================

'==================================================================================================================================
Sub RuT_Marcar_Repetidos()   '- Aparecieron registros repetidos.
'==================================================================================================================================
Dim LstObj              As ListObject
Dim RowAnt              As ListRow
Dim RowNow              As ListRow
Dim F_Lo                As Long
Dim Cont                As Integer
Dim Ref_Ant             As String:      Ref_Ant = ""
Dim TRows_Lo            As Long

Rut_Off_Functions

Set LstObj = Prog_BD.ListObjects(1)
    TRows_Lo = LstObj.ListRows.Count

'   Recorro toda la Tabla Tip-Hist para localizar nombres de Usuarios ----------------------------------------------------------

    Sheets(Prog_BD.Name).Select
    Prog_BD.Unprotect

    '- Recorrer toda la tabla --------------------------
    For F_Lo = 1 To LstObj.ListRows.Count   '--- Bucle para recorrer todas la filas de la Consulta Prog_BD

        Set RowNow = LstObj.ListRows(F_Lo)

        If RowNow.Range(BD_Ref) = Ref_Ant Then

            RowNow.Range(BD_EP_Ctrl) = RowNow.Range(BD_EP_Ctrl) & "_Duplicaty"
            Set RowAnt = LstObj.ListRows(F_Lo - 1)
            RowAnt.Range(BD_EP_Ctrl) = RowAnt.Range(BD_EP_Ctrl) & "_Duplicaty"

        Else
            Ref_Ant = RowNow.Range(BD_Ref)
        End If

        If F_Lo Mod 100 = 0 Then
            Debug.Print "Actualizando Tasas Adm.:" & _
                                        vbCrLf & Format(F_Lo, "#,##0") & " de " & Format(TRows_Lo, "#,##0")
        End If
    Next


'    Application.DisplayAlerts = False
'    On Error Resume Next
'    With Prog_LsGes04.ListObjects(1)
'        .AutoFilter.ShowAllData
'        .Range.AutoFilter Field:=BD_EP_Ctrl, Criteria1:="<>*_Duplicaty*"
'            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
'        .AutoFilter.ShowAllData
'    End With
'    Application.DisplayAlerts = True
'    On Error GoTo 0

Restablecer_Valores:
Rut_On_Functions
End Sub     ' RuT_Marcar_Repetidos   --------------------------------------------------------------------------------------------
'===================================================================================================================================

''==================================================================================================================================
'Sub RuT_Añadir_AD_0010()   '- Voy a añadir manualmente los números de AD-0010.
''==================================================================================================================================
'Dim RowNow              As ListRow
'Dim F_Lo                As Long
'Dim Plan                As Integer:      Plan = 0
'Dim TRows_Lo            As Long
'Dim rowfind     As Variant
'
'Rut_Off_Functions
'
'Dim Lo_TitPH        As ListObject:      Set Lo_TitPH = Prog_BD.ListObjects(1)
'    TRows_Lo = Lo_TitPH.ListRows.Count
'Dim Lo_Plazos       As ListObject:      Set Lo_Plazos = Prog_TitP_Plazos.ListObjects(1)
'
''   Recorro toda la Tabla Tip-Hist  ----------------------------------------------------------
'
'    Prog_BD.Select
'    Prog_BD.Unprotect
'
'    Lo_TitPH.ShowTotals = False
'    If Lo_TitPH.Parent.FilterMode Then Lo_TitPH.Parent.ShowAllData
'    Call Rut_Lo_Sort(Lo_TitPH, BD_Plan, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
'    Call Rut_Lo_Sort(Lo_TitPH, BD_NumRec, xlAscending)    '- Ordenar primero accelera un montón el borrado -----
'    '- Recorrer toda la tabla --------------------------
'    For F_Lo = 1 To Lo_TitPH.ListRows.Count
'
'        Set RowNow = Lo_TitPH.ListRows(F_Lo)
'
'        If RowNow.Range(BD_Concepto) <> "1311.00" Then GoTo Sig_Fila
'        If RowNow.Range(BD_ACont_Emi) = "2024" Then GoTo Sig_Fila
'        If RowNow.Range(BD_ACont_Cob) = "2023" Then GoTo Sig_Fila
'        If Plan <> RowNow.Range(BD_Plan) Then
'            Plan = RowNow.Range(BD_Plan)
'             ' -----------------=============  Buscar Tipo Plan  ==================--------------------------------------------------------------
'            rowfind = Application.Match(Plan, Lo_Plazos.DataBodyRange.Columns(1), 0)
'            If IsError(rowfind) Then Debug.Print "Plan no encontrado: " & Plan:     GoTo Sig_Fila
'        Else
'            If IsError(rowfind) Then GoTo Sig_Fila
'        End If
'
'        RowNow.Range(BD_AD_Emi_Acad) = Lo_Plazos.DataBodyRange.Cells(rowfind, 2 + RowNow.Range(BD_NumRec).Value)
'
''        If F_Lo Mod 100 = 0 Then
''            Debug.Print "Actualizando Tasas Adm.:" & _
''                                        vbCrLf & Format(F_Lo, "#,##0") & " de " & Format(TRows_Lo, "#,##0")
''        End If
'Sig_Fila:
'    Next
'    '---------------------------------------------------------------------------------------------------------------------------------------
'Debug.Print "Finalizado"
'Restablecer_Valores:
'Rut_On_Functions
'End Sub     ' RuT_Añadir_AD_0010   --------------------------------------------------------------------------------------------
''===================================================================================================================================










