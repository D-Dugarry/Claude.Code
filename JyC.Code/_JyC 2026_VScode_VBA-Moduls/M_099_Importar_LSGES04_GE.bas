Attribute VB_Name = "M_099_Importar_LSGES04_GE"
'- M032_Importar_LSGES04_GE -----------------------------------------------------------------------------------------------------------
Option Explicit

'--- Tabla Prog_TPV_Pag  Tit.Propios y Cursos<200h-Histórico ----------------------
Const CG4_Plan              As Integer = 1       ' col: a
Const CG4_NomPlan           As Integer = 2       ' col: b
Const CG4_TipoCurso         As Integer = 3       ' col: c
Const CG4_C_Acad            As Integer = 4       ' col: d
Const CG4_Nom               As Integer = 5       ' col: e
Const CG4_DNI               As Integer = 6       ' col: f
Const CG4_Matricula         As Integer = 7       ' col: g
Const CG4_Anul              As Integer = 8       ' col: h
Const CG4_Ref               As Integer = 9       ' col: i
Const CG4_NumRec            As Integer = 10       ' col: j
Const CG4_ActivEco          As Integer = 11       ' col: k
Const CG4_FEmi              As Integer = 12       ' col: l
Const CG4_FCob              As Integer = 13       ' col: m
Const CG4_ImpRec            As Integer = 14       ' col: n
Const CG4_ImpCob            As Integer = 15       ' col: o
Const CG4_FormPag           As Integer = 16       ' col: p
Const CG4_CtaPag            As Integer = 17       ' col: q
Const CG4_RegMov            As Integer = 18       ' col: r
Const CG4_Grupo             As Integer = 19       ' col: s
Const CG4_ImpAcad           As Integer = 20       ' col: t
Const CG4_ImpAdm            As Integer = 21       ' col: u

'==================================================================================================================================
Sub RuT_Importar_LSGES04_GE()   '- Importar Última Consulta de LSGES04_GE, Actualizar registros existentes y Añadir Nuevos.
'==================================================================================================================================
Dim TxT_Progreso        As String
Dim Ref_Ant             As String:      Ref_Ant = ""
Dim Lo_TPV                As ListObject
Dim Lo_TPVpag                As ListObject
Dim Lo_TPVpag_DefCol         As ListObject
Dim RwG4            As ListRow
Dim RwPH            As ListRow
Dim RowFind     As Variant
Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '   Seleccionar fichero     ---------------------------------------------------------------------------------------------------
    Dim DatosFichDestino            As Variant
    Dim Arch__Pagos_TPV         As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = Application.Workbooks(ThisWorkbook.Name).Path & "\"
        .Title = "Seleccionar el Fichero Excel de la última consulta LSG4_GE de un PLAN ÚNICO del Generador de Informes: "
        .ButtonName = "Aceptar"
        .AllowMultiSelect = False
        .Filters.Add "Sólo Ficheros Excel", "*.xls?", 1
        If .Show <> -1 Then
'            MsgBox "Cancelado", , "Rutinas"
            MsgBx_Msg = "Cancelado"
            MsgBx_Title = "Proceso: Importar LSGES04_GE"
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
            Prog__APP.Range("APP_Task_Inf") = "Proceso Cancelado: " & Now()
            GoTo Restablecer_Valores
        Else
            Arch__Pagos_TPV = .SelectedItems(1)
        End If
    End With
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            Form_Menu.Lb_Tarea_Inform.Caption = "Progreso de la Tarea."
            Form_Menu.Lb_Tarea_Inform.Visible = True
            Form_Menu.TBx_Tarea_Inform.Visible = True
            Form_Menu.TBx_Tarea_Inform = "Importando Excel:      " & Format(Now, "hh:mm:ss")
            Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
    '   Borrar el contenido de la hoja Prog_TPV_Pag    ------------------------------------------------------------
    Prog_TPV_Pag.Visible = xlSheetVisible
    Call Rut_WrkSheet_Vaciar(Prog_TPV_Pag)
    '   Copy Sheet LSG4 Without Opening it       ------------------------------------------------------------------
        '''    Dim WB As Workbook       '- para averiguar si existe un Excel abierto --------
        '''    For Each WB In Workbooks
        '''        If WB.Name = mid(Arch__Pagos_TPV,instrRev(Arch__Pagos_TPV,"\")+1,99) Then
        '''            WB.Activate
        '''            MsgBox "Workbook Found!"
        '''            Exit Sub
        '''        End If
        '''    Next WB
    Dim closedBook          As Variant
    Set closedBook = Workbooks.Open(Arch__Pagos_TPV)
    closedBook.Sheets(1).ListObjects(1).Range.Copy ThisWorkbook.Sheets(Prog_TPV_Pag.Name).Range("A1")
    Application.CutCopyMode = False
    closedBook.Close SaveChanges:=False
    Set closedBook = Nothing
    '   Si no viene con Tabla la Creo       ------------------------------------------------------------------------
    If Prog_TPV_Pag.ListObjects.Count = 0 Then
       Prog_TPV_Pag.ListObjects.Add(xlSrcRange, Sheets(Prog_TPV_Pag.Name).UsedRange, , xlYes).Name = "Tb_LSG4"
    End If
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            Form_Menu.TBx_Tarea_Inform = Form_Menu.TBx_Tarea_Inform & vbCrLf & "Importado Excel:       " & Format(Now, "hh:mm:ss")
            Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
'----------------------------------------------------------
'- Setting ListObjects ------------------------------------
Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
Set Lo_TPVpag = Prog_TPV_Pag.ListObjects(1)
Set Lo_TPVpag_DefCol = Prog_DefCol.ListObjects(1)
'- Setting ListObjects ------------------------------------
'----------------------------------------------------------2023433453507
    Prog_TPV_Pag.Select
    Prog_TPV_Pag.Visible = True
    Prog_TPV_Pag.Unprotect
    'Filtrar el Curso Académico      -----------------------------------------------------------------------------------------------
    Application.DisplayAlerts = False
    On Error Resume Next
    With Prog_TPV_Pag.ListObjects(1)
        .ShowTotals = False
        
        '- -----------------------------------------------------------------------------
        If .Parent.FilterMode Then .Parent.ShowAllData
        .DataBodyRange.Columns(CPH_ActivEco).Select     '- Datos - Texto en Columnas - Finalizar - PARA NÚMEROS --------
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
        Call Rut_LstObj_Sort(Lo_TPVpag, CPH_ActivEco, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=CPH_ActivEco, Criteria1:="<>4", Operator:=xlAnd, Criteria2:="<>300"        '- Elimino las NO Títulos Propios
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        If .Parent.FilterMode Then .Parent.ShowAllData
        '- -----------------------------------------------------------------------------
        '- Según el nombre de Fichero identifico si gestiono tasas de Tit_Propios o de Cursos < 200h  ----------------------
        If .Parent.FilterMode Then .Parent.ShowAllData
        Lo_TPVpag.ListColumns(CPH_RegMov).DataBodyRange.ClearContents    '- ClearContents -----
        Dim Fila    As Long
        If Left(ThisWorkbook.Name, 11) = "Tit_Propios" Then
            Call Rut_LstObj_Sort(Lo_TPVpag, CPH_Plan, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
            '- Identificar Plan Tit_Prop / Curso <200h y según el WorkBook proceso o no el registro --------------------
            For Fila = 1 To Lo_TPVpag.ListRows.Count
                If Not IsNumeric(Left(Lo_TPVpag.DataBodyRange.Cells(Fila, CPH_Plan), 1)) Then Exit For
                Lo_TPVpag.DataBodyRange.Cells(Fila, CPH_RegMov) = "1311.00"
            Next
        Else
            Call Rut_LstObj_Sort(Lo_TPVpag, CPH_Plan, xlDescending, True)    '- Ordenar primero accelera un montón el borrado -----
            '- Identificar Plan Tit_Prop / Curso <200h y según el WorkBook proceso o no el registro --------------------
            For Fila = 1 To Lo_TPVpag.ListRows.Count
                If IsNumeric(Left(Lo_TPVpag.DataBodyRange.Cells(Fila, CPH_Plan), 1)) Then Exit For
                Lo_TPVpag.DataBodyRange.Cells(Fila, CPH_RegMov) = "1311.03"
            Next
        End If
        Call Rut_LstObj_Sort(Lo_TPVpag, CPH_RegMov, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=CPH_RegMov, Criteria1:="="        '- Elimino los registros sin Concepto
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        If .Parent.FilterMode Then .Parent.ShowAllData
        Lo_TPVpag.ListColumns(CPH_RegMov).DataBodyRange.ClearContents    '- ClearContents -----
        '- -----------------------------------------------------------------------------
    End With
    Application.DisplayAlerts = True
    On Error GoTo 0
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            Form_Menu.TBx_Tarea_Inform = Form_Menu.TBx_Tarea_Inform & vbCrLf & "Filtrado Excel:        " & Format(Now, "hh:mm:ss")
            Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
    'Formatear la Tabla de Prog_TPV_Pag ---------------------------------------------------------------------------------------------
    Call Rut_X_Format_LoBjDatos_LoBjConFig(ActiveSheet.ListObjects(1), Prog_DefCol.ListObjects(1))
'    Call Rut_X_Format_LoBjDatos_LoBjConFig(Sheets(Prog_TPV_Pag.Name).ListObjects(1), Sheets(Prog_DefCol.Name).ListObjects(1))
        '- Call Rut_X_Format_LoBjDatos_LoBjConFig(Lo_TPVpag, Lo_TPVpag_DefCol)  '- NO FUNCIONA DE ESTA MANERA -----------
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            Form_Menu.TBx_Tarea_Inform = Form_Menu.TBx_Tarea_Inform & vbCrLf & "Formateado Excel:      " & Format(Now, "hh:mm:ss")
            Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
'--- Recorro toda la Tabla Prog_TPV_Pag para actualizar Prog_TPV_Tb ----------------------------------------------------------
Dim F_Actualiz      As String:  F_Actualiz = Now()
Dim Incidencia      As String
Dim Cont            As Integer
Dim F_PH            As Long:    F_PH = 1
Dim F_G4            As Long:    F_G4 = 1
Dim TRows_TitPH     As Long:    TRows_TitPH = Lo_TPV.ListRows.Count
Dim TRows_Ges04     As Long:    TRows_Ges04 = Lo_TPVpag.ListRows.Count
Dim Cont_Repes      As Long:    Cont_Repes = 0
Dim Cont_Nuevo      As Long:    Cont_Nuevo = 0
Dim Cont_Modif      As Long:    Cont_Modif = 0
Dim Cont_Mat_Anul   As Long:    Cont_Mat_Anul = 0
Dim Chg_ImpRec      As Long:    Chg_ImpRec = 0
Dim Chg_ImpCob      As Long:    Chg_ImpCob = 0
Dim Chg_ImpAdm      As Long:    Chg_ImpAdm = 0
    '   Añado Todos los registros nuevos y marco los registros eliminados
    ' =============  Preparar Tabla de TitPH ==================
    Call Rut_LstObj_WrkSht_Preparar(Prog_TPV_Tb)
    Call Rut_LstObj_Sort(Lo_TPV, CPH_Ref, xlAscending, True)    ' Ordenar por una Columna
    ' =============  Preparar Tabla de TitPH ==================
    Call Rut_LstObj_WrkSht_Preparar(Prog_TPV_Pag)
    Call Rut_LstObj_Sort(Lo_TPVpag, CG4_Ref, xlAscending, True)    ' Ordenar por una Columna

    Lo_TPV.ListColumns(CPH_Incidencia).DataBodyRange.ClearContents    '- ClearContents -----
    Lo_TPV.ListColumns(CPH_Gest_Reg).DataBodyRange.ClearContents          '- ClearContents -----
    'Lo_TPV.ListColumns(CPH_Hist_Incid).DataBodyRange.ClearContents    '- ClearContents -----
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            TxT_Progreso = Form_Menu.TBx_Tarea_Inform & vbCrLf
            Form_Menu.TBx_Tarea_Inform = TxT_Progreso & "Incorporando LSGES04:  " & " 0 de " & Format(TRows_Ges04, "#,##0")
            Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
    '   Recorro toda la Tabla Prog_TPV_Pag para actualizar Prog_TPV_Tb ----------------------------------------------------------
    Rut_Off_Functions
    Prog_TPV_Tb.Unprotect
    Prog_TPV_Pag.Unprotect
    
    Do While F_G4 <= TRows_Ges04
        Set RwG4 = Lo_TPVpag.ListRows(F_G4)
            '- Identificar Referencias Duplicadas, la SALTO --------------------
            If RwG4.Range(CG4_Ref) = Ref_Ant Then
                RwG4.Range(CPH_RegMov) = RwG4.Range(CPH_RegMov) & "_Duplicaty"
                Lo_TPVpag.ListRows(F_G4 - 1).Range(CPH_RegMov) = Lo_TPVpag.ListRows(F_G4 - 1).Range(CPH_RegMov) & "_Duplicaty"
                Cont_Repes = Cont_Repes + 1
                F_G4 = F_G4 + 1
                GoTo Siguiente_Reg
            End If
        Set RwPH = Lo_TPV.ListRows(F_PH)
        '--- Referencias coincides <<<<  Ya existe y hay que ver de Actualizar si hay Cambios  <<<<<<<<<<<<<<<<
        '--- Referencias coincides <<<<  Ya existe y hay que ver de Actualizar si hay Cambios  <<<<<<<<<<<<<<<<
        '--- Referencias coincides <<<<  Ya existe y hay que ver de Actualizar si hay Cambios  <<<<<<<<<<<<<<<<
        '--- Referencias coincides <<<<  Ya existe y hay que ver de Actualizar si hay Cambios  <<<<<<<<<<<<<<<<
        If Val(RwPH.Range(CPH_Ref)) = Val(RwG4.Range(CG4_Ref)) Then   '- YA EXISTE, LO ACTUALIZO ----------------------------------------
            '- Compruebo posibles INCIDENCIAS -------------------------------------------------------------
            If RwPH.Range(CPH_ImpRec) <> RwG4.Range(CG4_ImpRec) * 1 Then                '- Cambio en el Imp. Recibo ----------------
                Incidencia = "Chg:PH_ImpRec=[" & RwPH.Range(CPH_ImpRec) & "]_#_"
                RwPH.Range(CPH_Incidencia) = RwPH.Range(CPH_Incidencia) & Incidencia
                RwPH.Range(CPH_Hist_Incid) = RwPH.Range(CPH_Hist_Incid) & Incidencia
                Chg_ImpRec = Chg_ImpRec + 1
                End If
            If RwPH.Range(CPH_ImpCob) > 0 And RwPH.Range(CPH_ImpCob) <> RwG4.Range(CG4_ImpCob) * 1 Then     '- Cambio en el Imp. Cobrado ----------------
                Incidencia = "Chg:PH_ImpCob=[" & RwPH.Range(CPH_ImpCob) & "]_#_"
                RwPH.Range(CPH_Incidencia) = RwPH.Range(CPH_Incidencia) & Incidencia
                RwPH.Range(CPH_Hist_Incid) = RwPH.Range(CPH_Hist_Incid) & Incidencia
                Chg_ImpCob = Chg_ImpCob + 1
                End If
            If RwPH.Range(CPH_ImpAdm) > 0 And RwPH.Range(CPH_ImpAdm) <> RwG4.Range(CG4_ImpAdm) * 1 Then     '- Cambio en el Imp. Adm. ----------------
                Incidencia = "Chg:PH_ImpAdm=[" & RwPH.Range(CPH_ImpAdm) & "]_#_"
                RwPH.Range(CPH_Incidencia) = RwPH.Range(CPH_Incidencia) & Incidencia
                RwPH.Range(CPH_Hist_Incid) = RwPH.Range(CPH_Hist_Incid) & Incidencia
                Chg_ImpAdm = Chg_ImpAdm + 1
                End If
            If RwG4.Range(CG4_Anul) = "S" Then     '- Tasa Anulada ----------------
                RwPH.Range(CPH_ObsConta) = "Mat.Anulada_"
                Cont_Mat_Anul = Cont_Mat_Anul + 1
                End If
            
            '- Actualizo Todos los datos Nuevos -----------------------------------------------------------
                'Lo_TPVpag.DataBodyRange.Rows(F_G4).Copy Lo_TPV.ListRows(F_PH).Range   '- OJO El rango a copiar lo repite si cabe el la fila destino
            For Cont = 1 To Lo_TPVpag.Range.Columns.Count
                RwPH.Range(Cont) = RwG4.Range(Cont)
            Next Cont
            '- Preparo Salto de registro ------------------------------------
            RwPH.Range(CPH_Gest_Reg) = " - Actualizado con G04, " & F_Actualiz
            Cont_Modif = Cont_Modif + 1
            If F_PH < TRows_TitPH Then F_PH = F_PH + 1
            F_G4 = F_G4 + 1
            Ref_Ant = RwG4.Range(CG4_Ref)
        '--- Referencia NUEVA NO EXISTE, ES UN NUEVO REGISTRO  <<<<  Ya existe y hay que ver de Actualizar si hay Cambios  <<<<<<<<<<<<<<<<
        '--- Referencia NUEVA NO EXISTE, ES UN NUEVO REGISTRO  <<<<  Ya existe y hay que ver de Actualizar si hay Cambios  <<<<<<<<<<<<<<<<
        '--- Referencia NUEVA NO EXISTE, ES UN NUEVO REGISTRO  <<<<  Ya existe y hay que ver de Actualizar si hay Cambios  <<<<<<<<<<<<<<<<
        '--- Referencia NUEVA NO EXISTE, ES UN NUEVO REGISTRO  <<<<  Ya existe y hay que ver de Actualizar si hay Cambios  <<<<<<<<<<<<<<<<
        ElseIf Val(RwPH.Range(CPH_Ref)) > Val(RwG4.Range(CG4_Ref)) Or F_PH >= TRows_TitPH Then        '- NO EXISTE, AÑADO REGISTRO --------------
            '--- Añado Registro -------------------------------------------------------------------
            Set RwPH = Lo_TPV.ListRows.Add
            '- Copio todas las celdas -----------------
            For Cont = 1 To Lo_TPVpag.Range.Columns.Count
                RwPH.Range(Cont) = RwG4.Range(Cont)
            Next Cont
            '--- Incorporar Campos Calculados Coef_VRI -----------------------------------------------
            Call Rut_Incorporar_Tipo_TitP_Curs(RwPH.Range(CPH_Plan))
                RwPH.Range(CPH_Coef_VRI) = Coef_VRI
            '- Preparo Salto de registro ------------------------------------
            RwPH.Range(CPH_Gest_Reg) = " - Nuev G04"
            RwPH.Range(CPH_AñEmi) = Format(RwPH.Range(CPH_FEmi), "yyyy")
            RwPH.Range(CPH_AñCob) = Format(RwPH.Range(CPH_FCob), "yyyy")
            Cont_Nuevo = Cont_Nuevo + 1
            F_G4 = F_G4 + 1
            Ref_Ant = RwG4.Range(CG4_Ref)
        '--- Referencia ANTIGUA NO EXISTE, ES UN REGISTRO ELIMINADO <<<< Lo marcamos para futuros controles <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '--- Referencia ANTIGUA NO EXISTE, ES UN REGISTRO ELIMINADO <<<< Lo marcamos para futuros controles <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '--- Referencia ANTIGUA NO EXISTE, ES UN REGISTRO ELIMINADO <<<< Lo marcamos para futuros controles <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '--- Referencia ANTIGUA NO EXISTE, ES UN REGISTRO ELIMINADO <<<< Lo marcamos para futuros controles <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        Else
            RwPH.Range(CPH_Gest_Reg) = " - Ref-TitPH Deleted"
            If F_PH < TRows_TitPH Then F_PH = F_PH + 1
        End If
                '- Visualizo el progreso ---------------------------------------------------------------------------------------
                If F_G4 Mod 100 = 0 Then
                    Form_Menu.TBx_Tarea_Inform = TxT_Progreso & "Incorporando LSGES04:  " & Format(F_G4, "#,##0") & " de " & Format(TRows_Ges04, "#,##0")
                    Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
                    Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
                End If
Siguiente_Reg:
    Loop
    '- Visualizo el progreso ---------------------------------------------------------------------------------------
    TxT_Progreso = Form_Menu.TBx_Tarea_Inform & vbCrLf
'=== IMPORTANTE, Actualizar El Importe de Tasa Adm  =====================================================================================
'=== IMPORTANTE, Actualizar El Importe de Tasa Adm  =====================================================================================
    ' =============  Preparar Tabla de TitP ==================
    Call Rut_LstObj_WrkSht_Preparar(Prog_TPV_Tb)                         '- Quita filtros, filas y columnas ocultas
    Call Rut_LstObj_Sort(Lo_TPV, CPH_C_Acad, xlAscending, True)       ' Ordenar por una Columna
    Call Rut_LstObj_Sort(Lo_TPV, CPH_Plan, xlAscending)
    Call Rut_LstObj_Sort(Lo_TPV, CPH_DNI, xlAscending)
    Call Rut_LstObj_Sort(Lo_TPV, CPH_Ref, xlAscending)
    '- Recorro toda la tabla Lo_TPV -----------------------------------------------------------------
    Dim DNI_Ant     As String:      DNI_Ant = ""
    For F_PH = 1 To Lo_TPV.ListRows.Count   '--- Bucle para recorrer todas la filas de la Consulta Prog_TPV_Tb
        Set RwPH = Lo_TPV.ListRows(F_PH)
        '--- Importe Tasa Adm --------
        If RwPH.Range(CPH_DNI) <> DNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
            If RwPH.Range(CPH_ImpAdm) > 0 Then  '--- Solo si el importe es positivo
                If Len(RwPH.Range(CPH_TasaAdm)) = 0 Then RwPH.Range(CPH_TasaAdm) = RwPH.Range(CPH_ImpAdm)
                DNI_Ant = RwPH.Range(CPH_DNI)
            End If
        End If
        '- Visualizo el progreso ---------------------------------------------------------------------------------------
        If F_PH Mod 100 = 0 Then
            Form_Menu.TBx_Tarea_Inform = TxT_Progreso & "Actualizando T.Adm.:   " & Format(TRows_Ges04, "#,##0") & " reg." & vbCrLf & "Actualizando Tasas Adm.:" & _
                                        vbCrLf & Format(F_PH, "#,##0") & " de " & Format(TRows_TitPH, "#,##0")
            Form_Tarea.TBx_Tarea = Form_Menu.TBx_Tarea_Inform
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
        End If
    Next
'=== Borrar Prog_TPV_Pag menos los DUPLICADOS para su posible control  ===================================================================
'=== Borrar Prog_TPV_Pag menos los DUPLICADOS para su posible control  ===================================================================
    Application.DisplayAlerts = False
    On Error Resume Next
    With Prog_TPV_Pag.ListObjects(1)
        .AutoFilter.ShowAllData
        .Range.AutoFilter Field:=CPH_RegMov, Criteria1:="<>*_Duplicaty*"    '........ Elimino los que no son duplicados
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        .AutoFilter.ShowAllData
    End With
    Application.DisplayAlerts = True
    On Error GoTo 0
'- Visualizo el progreso ---------------------------------------------------------------------------------------
Form_Menu.Lb_Tarea_Inform.Visible = False
Form_Menu.TBx_Tarea_Inform.Visible = False
Prog__APP.Range("APP_Task_Inf") = "¡¡¡ Proceso concluido con éxito !!! día: " & Now() & " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
        "En la Anterior Consulta habían:  " & Right("__________" & TRows_TitPH, 8) & "  Reg." & vbCrLf & _
        "En la Nueva Consulta hay:        " & Right("__________" & TRows_Ges04, 8) & "  Reg.   Con: " & Cont_Repes & " Reg. Repetidos" & vbCrLf & _
        Right("__________" & Cont_Nuevo, 8) & "  Reg. nuevos." & vbCrLf & _
        Right("__________" & Cont_Modif, 8) & "  Reg. que existían y se han actualizado." & vbCrLf & _
        Right("__________" & Cont_Mat_Anul, 8) & "  Tasas Anuladas." & vbCrLf & _
        Right("__________" & Chg_ImpRec + Chg_ImpCob + Chg_ImpAdm, 8) & "  Reg. Actualizados con incidencias (Cambios destacables)." & vbCrLf & _
        Right("__________" & Chg_ImpRec, 8) & "  Reg. Cambio en Importe de Recibo." & vbCrLf & _
        Right("__________" & Chg_ImpCob, 8) & "  Reg. Cambio en Importe Cobrado." & vbCrLf & _
        Right("__________" & Chg_ImpAdm, 8) & "  Reg. Cambio en Importe Administrativo." & vbCrLf & _
        "Hay activos ahora un Total de:  " & Lo_TPV.DataBodyRange.Rows.Count & "  Reg."
Prog__Config_APP.Range("APP_Last_Import") = Format(Now(), "dd-mmm-yy hh:mm")
Call RefreshRibbon
Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Prog_TPV_Tb.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Prog_TPV_Pag.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Prog_TPV_Tb.Visible = xlSheetVeryHidden
Prog_TPV_Pag.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Importar_LSGES04_GE   --------------------------------------------------------------------------------------------
'===================================================================================================================================

' ==================================================================================================================================
Sub Rut_Incorporar_Tipo_TitP_Curs(ByRef Plan As String)
' ==================================================================================================================================
Dim RowFind            As Variant
Dim Lo_Tb_Ret_VRI        As ListObject
Set Lo_Tb_Ret_VRI = Prog_Ret_VRI.ListObjects(1)
    ' -----------------=============  Buscar Tipo Plan  ==================--------------------------------------------------------------
    RowFind = Application.Match(Plan, Lo_Tb_Ret_VRI.DataBodyRange.Columns(1), 0)
    If Not IsError(RowFind) Then    ' Plan Encontrado ==>> Tendrá características ESPECIALES ------------------------
        Coef_VRI = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(RowFind)
        Concepto = Lo_Tb_Ret_VRI.ListColumns("Concepto").DataBodyRange(RowFind)
    Else                            ' NO ENCONTRADO
        If IsNumeric(Left(Plan, 1)) Then
            Coef_VRI = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(2)
            Concepto = Lo_Tb_Ret_VRI.ListColumns("Concepto").DataBodyRange(2)
        Else
            Coef_VRI = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(1)
            Concepto = Lo_Tb_Ret_VRI.ListColumns("Concepto").DataBodyRange(1)
        End If
    End If
    '---------------------------------------------------------------------------------------------------------------------------------------
End Sub     '      <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'---------------------------- Rut de trabajo interno, a eliminar ------------------------------------------------------------------
'==================================================================================================================================
'==================================================================================================================================
Sub RuT_Actualizar_Repetidos()   '- Aparecieron registros repetidos.
'==================================================================================================================================

Dim TxT_Progreso        As String
Dim Lo_TPV            As ListObject
Dim Lo_TPVpag            As ListObject
Dim Lo_TPVpag_DefCol     As ListObject
Dim RwPHant            As ListRow
Dim RwPH            As ListRow

Dim RowFind             As Variant


'Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso

Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)

'   Recorro toda la Tabla Tip-Hist  ----------------------------------------------------------
Dim F_PH            As Long:    F_PH = 1
Dim ContFila        As Long:    ContFila = 0
Dim ContFilDatos    As Long:    ContFilDatos = 0
Dim ContFilDup      As Long:    ContFilDup = 0

Dim Cont            As Integer
Dim Ref_Ant      As String:     Ref_Ant = ""
Dim TRows_TitPH      As Long: TRows_TitPH = Lo_TPV.ListRows.Count

    Sheets(Prog_TPV_Tb.Name).Select
    Prog_TPV_Tb.Unprotect
'    '- Copiar DataBodyRange ----------------------------------------------
''    Call Rut_LstObj_DataBodyRange_Copy(Prog_TPV1.ListObjects(1), Prog_TPV_Tb.ListObjects(1), True)
'
    '- Ordenar por PLAN y DNI ==================
    Call Rut_LstObj_Sort(Lo_TPV, CPH_Ref, xlAscending, True)
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
    For F_PH = 1 To Lo_TPV.ListRows.Count   '--- Bucle para recorrer todas la filas de la Consulta Prog_TPV_Tb

        Set RwPH = Lo_TPV.ListRows(F_PH)

        If RwPH.Range(CPH_Ref) = Ref_Ant Then

            Set RwPHant = Lo_TPV.ListRows(F_PH - 1)

'            Lo_TPV.ListRows(F_PH).Range.Select
                RwPHant.Range(CPH_RegMov) = RwPHant.Range(CPH_RegMov) & "_Duplicaty1"
                RwPH.Range(CPH_RegMov) = RwPH.Range(CPH_RegMov) & "_Duplicaty2"
                ContFilDup = ContFilDup + 1

            For Cont = 1 To Lo_TPV.Range.Columns.Count
                If RwPH.Range(Cont) <> RwPHant.Range(Cont) Then
                    If Len(RwPHant.Range(Cont).Value) = 0 Then RwPHant.Range(Cont) = RwPH.Range(Cont)
                End If

            Next Cont

            If Len(RwPH.Range(CPH_Orgánica) & RwPH.Range(CPH_ExpAdm) & RwPH.Range(CPH_Liquidado) & RwPH.Range(CPH_RDT) & RwPH.Range(CPH_JI_Emi)) > 0 Then
                '''    ContFila = ContFila + 1
                '''    Prog_Dels_TPH.ListObjects(1).ListRows.Add
                '''    RwPH.Range.Copy Prog_Dels_TPH.ListObjects(1).ListRows(ContFila).Range
                RwPHant.Range(CPH_RegMov) = RwPHant.Range(CPH_RegMov) & "_Datos"
                ContFilDatos = ContFilDatos + 1
'                Lo_TPV.ListRows(F_PH).Delete
                F_PH = F_PH - 1
            End If
            
'            If Len(RwPH.Range(CPH_Orgánica) & RwPH.Range(CPH_ExpAdm) & RwPH.Range(CPH_Liquidado) & RwPH.Range(CPH_RDT) & RwPH.Range(CPH_JI_Emi)) = 0 Then
'                ContFilDatos = ContFilDatos + 1
'                Lo_TPV.ListRows(F_PH).Delete
'                F_PH = F_PH - 1
'            End If
        Else
            Ref_Ant = RwPH.Range(CPH_Ref)
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

Set LstObj = Prog_TPV_Tb.ListObjects(1)
    TRows_Lo = LstObj.ListRows.Count

'   Recorro toda la Tabla Tip-Hist para localizar nombres de Usuarios ----------------------------------------------------------

    Sheets(Prog_TPV_Tb.Name).Select
    Prog_TPV_Tb.Unprotect

    '- Recorrer toda la tabla --------------------------
    For F_Lo = 1 To LstObj.ListRows.Count   '--- Bucle para recorrer todas la filas de la Consulta Prog_TPV_Tb

        Set RowNow = LstObj.ListRows(F_Lo)

        If RowNow.Range(CPH_Ref) = Ref_Ant Then

            RowNow.Range(CPH_RegMov) = RowNow.Range(CPH_RegMov) & "_Duplicaty"
            Set RowAnt = LstObj.ListRows(F_Lo - 1)
            RowAnt.Range(CPH_RegMov) = RowAnt.Range(CPH_RegMov) & "_Duplicaty"

        Else
            Ref_Ant = RowNow.Range(CPH_Ref)
        End If

        If F_Lo Mod 100 = 0 Then
            Debug.Print "Actualizando Tasas Adm.:" & _
                                        vbCrLf & Format(F_Lo, "#,##0") & " de " & Format(TRows_Lo, "#,##0")
        End If
    Next


'    Application.DisplayAlerts = False
'    On Error Resume Next
'    With Prog_TPV_Pag.ListObjects(1)
'        .AutoFilter.ShowAllData
'        .Range.AutoFilter Field:=CPH_RegMov, Criteria1:="<>*_Duplicaty*"
'            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
'        .AutoFilter.ShowAllData
'    End With
'    Application.DisplayAlerts = True
'    On Error GoTo 0

Restablecer_Valores:
Rut_On_Functions
End Sub     ' RuT_Actualizar_Concepto   --------------------------------------------------------------------------------------------
'===================================================================================================================================


Sub kk()
Prog_TPV_Tb.Select
Prog_TPV_Tb.Unprotect
    With Prog_TPV_Tb.ListObjects(1)
        .ShowTotals = False
        If .Parent.FilterMode Then .Parent.ShowAllData
        .Range.AutoFilter Field:=8, Criteria1:="=2023433453507"    '- Elimino Este que NO VALE y que cada vez que cargo la Consulta LSGES04-2023 me sale, es el del M013
        If .Range.Columns(8).SpecialCells(xlCellTypeVisible).Cells.Count - 1 > 0 Then Debug.Print "YES" Else Debug.Print "NO"
    End With

End Sub

'==================================================================================================================================
Sub RuT_Añadir_AD_0010()   '- Voy a añadir manualmente los números de AD-0010.
'==================================================================================================================================
Dim RowNow              As ListRow
Dim F_Lo                As Long
Dim Plan                As Integer:      Plan = 0
Dim TRows_Lo            As Long
Dim RowFind     As Variant

Rut_Off_Functions

Dim Lo_TPV        As ListObject:      Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    TRows_Lo = Lo_TPV.ListRows.Count
Dim Lo_Plazos       As ListObject:      Set Lo_Plazos = Prog_TitP_Plazos.ListObjects(1)

'   Recorro toda la Tabla Tip-Hist  ----------------------------------------------------------

    Prog_TPV_Tb.Select
    Prog_TPV_Tb.Unprotect
    
    Lo_TPV.ShowTotals = False
    If Lo_TPV.Parent.FilterMode Then Lo_TPV.Parent.ShowAllData
    Call Rut_LstObj_Sort(Lo_TPV, CPH_Plan, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_LstObj_Sort(Lo_TPV, CPH_NumRec, xlAscending)    '- Ordenar primero accelera un montón el borrado -----
    '- Recorrer toda la tabla --------------------------
    For F_Lo = 1 To Lo_TPV.ListRows.Count

        Set RowNow = Lo_TPV.ListRows(F_Lo)

        If RowNow.Range(CPH_Concept) <> "1311.00" Then GoTo Sig_Fila
        If RowNow.Range(CPH_AñEmi) = "2024" Then GoTo Sig_Fila
        If RowNow.Range(CPH_AñCob) = "2023" Then GoTo Sig_Fila
        If Plan <> RowNow.Range(CPH_Plan) Then
            Plan = RowNow.Range(CPH_Plan)
             ' -----------------=============  Buscar Tipo Plan  ==================--------------------------------------------------------------
            RowFind = Application.Match(Plan, Lo_Plazos.DataBodyRange.Columns(1), 0)
            If IsError(RowFind) Then Debug.Print "Plan no encontrado: " & Plan:     GoTo Sig_Fila
        Else
            If IsError(RowFind) Then GoTo Sig_Fila
        End If
        
        RowNow.Range(CPH_AD_0010) = Lo_Plazos.DataBodyRange.Cells(RowFind, 2 + RowNow.Range(CPH_NumRec).Value)
    
'        If F_Lo Mod 100 = 0 Then
'            Debug.Print "Actualizando Tasas Adm.:" & _
'                                        vbCrLf & Format(F_Lo, "#,##0") & " de " & Format(TRows_Lo, "#,##0")
'        End If
Sig_Fila:
    Next
    '---------------------------------------------------------------------------------------------------------------------------------------
Debug.Print "Finalizado"
Restablecer_Valores:
Rut_On_Functions
End Sub     ' RuT_Actualizar_Concepto   --------------------------------------------------------------------------------------------
'===================================================================================================================================


