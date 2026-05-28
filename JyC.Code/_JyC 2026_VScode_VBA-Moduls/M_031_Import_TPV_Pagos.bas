Attribute VB_Name = "M_031_Import_TPV_Pagos"
'- M_031_Import_TPV
' Generador de informe en: https://cvnet.cpd.ua.es/uaGenInf/Home/Consulta/38051
Option Explicit

'Dim AñoCont                     As String           ' Para Controlar el cambio de años el en número de orden que genero
' ==================================================================================================================================
            Sub Call_Rut_Import_TPV_Pagos()
                Prog__APP.Range("APP_Task_Rut") = "Rut_Import_TPV_Pagos"
                Form_Running_Rut.Show
            End Sub
' ==================================================================================================================================
Sub Rut_Import_TPV_Pagos()  '- Copio el Excel en TPV_Pagos, en TPV_Tb borro los Reg sin Liquidar, Incorporo los nuevos de TPV_Pagos en TPV_Tb y actualizo los de TPV_Tb si hay cambios.
' ==================================================================================================================================
    Dim ActivSheet          As String:          ActivSheet = ActiveSheet.Name
    Dim TBx_Informe_Hist    As String
    Dim Cant_Reg            As Long
    Dim Lo_TPV              As ListObject:      Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Dim Lo_Liqdatos         As ListObject:      Set Lo_Liqdatos = Prog_TPV_Liqdatos.ListObjects(1)
    Dim ActivForm           As Object           '- Identificamos qué Formulario está Activo.  ------------------------------
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)
    
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    Rut_Off_Functions
    Application.ScreenUpdating = False
    Call Rut_LstObj_WrkSht_Preparar(Prog_TPV_Liqdatos)  '- Quita Filtros y muestra todas la Columnas y Filas Ocultas

'- Seleccionar fichero Excel del Generador de Pagos-TPV: PAGOS.Xls
'- Borrar TPV_Pagos para traer el nuevo
'- Importo PAGOS.Xls (el Excel del Generador de Pagos-TPV) a TPV_Pagos
'- Borro de TPV_Pagos, los Pagos de Eventos que son Procesos Selectivos y NO JyC
'- Borro de TPV_Pagos, los Registros de Recibos de Eventos que NO están PAGADOS
'- Formatear Columnas de TPV_Pagos
'- Borro de TPV_Pagos, los Pagos de Eventos de Fecha_Fin anteriores al año contable
'- Añado a TPV_Pagos, Columnas para igualar las Tablas y poder copiar la tabla entera
'- Añadir las Siglas del Evento, avisar si falta por crear alguna nueva
'- Borrar los reg. SIN Liquidar de TPV_Tb (para guardar sólo los que tienen datos de liquidación)
'- Borrar los Registros de TPV_Tb, de Fecha_Fin anteriores al año contable
'- Guardar Reg. Liquidados de TPV_Tb en TPV-Liqdatos
'- Copia los nuevos datos TPV_Pag en TPV_Tb que está vacía
'- Copiar los Datos de TPV-Liqdatos al nuevo TPV_Tb

    '- Seleccionar fichero Excel del Generador de Pagos-TPV: PAGOS.Xls    -----------------------------------------------------
    Dim Arch__Pagos_TPV     As String
    Dim Título             As String
    Título = "Seleccionar el Fichero Excel de la última consulta de los pagos por TPV / Bizum del Generador de Informes: "
    Call Rut_File_Select(Título, Arch__Pagos_TPV, "Sólo Ficheros Excel", "*.xls?")       '- Arch__Pagos_TPV = "Cancel"
        If Arch__Pagos_TPV = "Cancel" Then
            MsgBx_Msg = "Cancelado: A petición Usuario."
            MsgBx_Title = "Proceso: Importar Pagos-TPV"
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = "Proceso Cancelado: " & Now()
            With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
                .Value = MsgBx_Msg
                ActivForm.Repaint
            End With
            GoTo Restablecer_Valores
        End If
        If VBA.UserForms.Count = 0 Then
            MsgBx_Msg = "Cancelado: ¡ Sin Formulario Activo !"
            MsgBx_Title = "Proceso: Importar Pagos-TPV"
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = "Proceso Cancelado: " & Now()
            With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
                .Value = MsgBx_Msg
                ActivForm.Repaint
            End With
            GoTo Restablecer_Valores
        End If

        '- Visualizo el progreso ---------------------------------------------------------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = "Guardo los datos de las Liquidaciones.  "
            ActivForm.Repaint
        End With
        
    '- Borrar TPV_Pagos para traer el nuevo   ------------------------------------------------------------
    Prog_TPV_Pag.Visible = xlSheetVisible
    Call Rut_WrkSheet_Vaciar(Prog_TPV_Pag.Name)
    
    '- Importo PAGOS.Xls (el Excel del Generador de Pagos-TPV) a TPV_Pagos   ------------------------------------------------
    Dim closedBook          As Variant
    Set closedBook = Workbooks.Open(Arch__Pagos_TPV)
'    closedBook.Sheets(1).ListObjects(1).Range.Copy ThisWorkbook.Sheets(Prog_TPV_Pag.Name).Range("A1")  '- Sólo si estamos seguros que viene la ListObject
    closedBook.Sheets(1).UsedRange.Copy ThisWorkbook.Sheets(Prog_TPV_Pag.Name).Range("A1")
    Application.CutCopyMode = False
    closedBook.Close SaveChanges:=False
    Set closedBook = Nothing
    '- Si no viene con Tabla la Creo       ------------------------------------------------------------------------
    If Prog_TPV_Pag.ListObjects.Count = 0 Then
       Prog_TPV_Pag.ListObjects.Add(xlSrcRange, Sheets(Prog_TPV_Pag.Name).UsedRange, , xlYes).Name = "Tb_TPV_Pagos"
    End If
    Dim Lo_TPVpag       As ListObject:      Set Lo_TPVpag = Prog_TPV_Pag.ListObjects(1)
    
    '- Borro de TPV_Pagos, los Pagos de Eventos que son Procesos Selectivos y NO JyC      --------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "Borro registros de Procesos Selectivos. (Ojo por si se cuela algún JyC)"
            ActivForm.Repaint
        End With
    Cant_Reg = Lo_TPVpag.ListRows.Count
    Call Rut_LstObj_Sort(Lo_TPVpag, C_Pag_EventoNom, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_LstObj_DataBodyRange_Filtered_DEL(Lo_TPVpag, C_Pag_EventoNom, "=REF. *")
    Call Rut_LstObj_DataBodyRange_Filtered_DEL(Lo_TPVpag, C_Pag_EventoNom, "=SOLICITUD DE PLAZAS*")
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "   - Habían: " & Format(Cant_Reg, "#,##0") & " reg. en el fichero de Pagos"
                    Cant_Reg = Lo_TPVpag.ListRows.Count
            .Value = .Value & ",   quedan: " & Format(Cant_Reg, "#,##0") & " reg. de JyC-TPV"
            ActivForm.Repaint
        End With
        
    '- Borro de TPV_Pagos, los Registros de Recibos de Eventos que NO están PAGADOS     --------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "Borro registros de Recibos NO Pagados."
            ActivForm.Repaint
        End With
    Cant_Reg = Lo_TPVpag.ListRows.Count
    Call Rut_LstObj_Sort(Lo_TPVpag, C_Pag_F_Pago, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_LstObj_DataBodyRange_Filtered_DEL(Lo_TPVpag, C_Pag_F_Pago, "=")
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "   - Habían: " & Format(Cant_Reg, "#,##0") & " reg. en el fichero de Pagos"
                    Cant_Reg = Lo_TPVpag.ListRows.Count
            .Value = .Value & ",   quedan: " & Format(Cant_Reg, "#,##0") & " reg. PAGADOS"
            ActivForm.Repaint
        End With
        
    '- Formatear Columnas de TPV_Pagos --------------------------------------------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "Formateo las columnas.  "
            ActivForm.Repaint
        End With
    Dim CantReg     As Long:        CantReg = Lo_TPVpag.ListRows.Count
    Prog_TPV_Pag.Visible = xlSheetVisible
    Prog_TPV_Pag.Select
    Call Rut_LstObj_Filtros_Quitar(Lo_TPVpag)
    With Lo_TPVpag
        .ShowTotals = False
        '- -----------------------------------------------------------------------------
        .DataBodyRange.Columns(C_Pag_F_Emi).Select     '- Datos - Texto en Columnas PARA Números --------
        Selection.TextToColumns Destination:=.DataBodyRange.Columns(C_Pag_F_Emi), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 4), TrailingMinusNumbers:=True

        Call Rut_LstObj_Col_Format_Date(Lo_TPVpag, C_Pag_F_Fin)
        Call Rut_LstObj_Col_Format_Date(Lo_TPVpag, C_Pag_F_Ini)

        .DataBodyRange.Columns(C_Pag_F_Vto).Select     '- Datos - Texto en Columnas PARA Números --------
        Selection.TextToColumns Destination:=.DataBodyRange.Columns(C_Pag_F_Vto), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 4), TrailingMinusNumbers:=True
            
        .DataBodyRange.Columns(C_Pag_F_Pago).Select     '- Datos - Texto en Columnas PARA Números --------
        Selection.TextToColumns Destination:=.DataBodyRange.Columns(C_Pag_F_Pago), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 4), TrailingMinusNumbers:=True
        .DataBodyRange.Columns(C_Pag_EventoRef).Select
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
        .DataBodyRange.Columns(C_Pag_Imp).Select
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
        .DataBodyRange.Columns(C_Pag_ComBco).Select
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
        .DataBodyRange.Columns(C_Pag_Neto).Select
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
        '- Borro datos sensibles DNI...     --------------------------------------------------------------
        .DataBodyRange.Columns(C_Pag_DNI).Select
        Selection.Resize(, 4).ClearContents
        .DataBodyRange.Columns(C_Pag_Tipo_Dto).ClearContents
    End With
    
    '- Borro de TPV_Pagos, los Pagos de Eventos de Fecha_Fin anteriores al año contable     --------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "Borro registros con Fecha Fin de Evento anterior al 1/1/" & Prog__APP.Range("APP_AñoCont") - 1
            ActivForm.Repaint
        End With
    Call Rut_LstObj_Sort(Lo_TPVpag, C_Pag_F_Fin, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_LstObj_DataBodyRange_Filtered_DEL(Lo_TPVpag, C_Pag_F_Fin, "<1/1/" & Prog__APP.Range("APP_AñoCont") - 1)
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
                    Cant_Reg = Lo_TPVpag.ListRows.Count
            .Value = .Value & ",   quedan: " & Format(Cant_Reg, "#,##0") & " reg. de JyC-TPV"
            ActivForm.Repaint
        End With
    
    '- Añado a TPV_Pagos, Columnas para igualar las Tablas y poder copiar la tabla entera   --------------------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "Añado Columnas para el Post-Precesado."
            ActivForm.Repaint
        End With
    Dim i       As Integer
    For i = 1 To Lo_TPV.ListColumns.Count - Lo_TPVpag.ListColumns.Count
        Lo_TPVpag.ListColumns.Add
    Next

    '- Añadir las Siglas del Evento, avisar si falta por crear alguna nueva   -----------------------------------------------------------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & vbLf & "Añado las Siglas de los Eventos."
            ActivForm.Repaint
        End With
    Dim LinDat      As Long:        LinDat = 1
    Dim RefAnt      As Integer:     RefAnt = 0
    Dim FirstCase   As Boolean:     FirstCase = True
    Dim Sigla       As String
    Dim N_Liq       As String
    Dim Lin_Lst     As Variant
    Dim Lo_Event      As ListObject:      Set Lo_Event = Prog_TPV_Events.ListObjects(1)
    Call Rut_LstObj_Sort(Lo_TPVpag, C_Pag_EventoRef, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_LstObj_Sort(Lo_Event, 1, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    With Lo_TPVpag.DataBodyRange
    For LinDat = 1 To Lo_TPVpag.ListRows.Count
        If RefAnt <> .Cells(LinDat, C_Pag_EventoRef) Then
            RefAnt = .Cells(LinDat, C_Pag_EventoRef)
'            If .Cells(LinDat, C_TPV_EventoNom) Like "REF. *" Or .Cells(LinDat, C_TPV_EventoNom) Like "SOLICITUD DE PLAZAS*" Then
'                Sigla = "_P.Select."
'                N_Liq = ""
'            Else
                N_Liq = "x"
                Lin_Lst = Application.Match(RefAnt, Lo_Event.DataBodyRange.Columns(1), 0)
                If Not IsError(Lin_Lst) Then    ' DATO Encontrado ------------------------
                    Sigla = Lo_Event.DataBodyRange.Cells(Lin_Lst, 2)
                Else                            ' NO ENCONTRADO   ------------------------
                    Sigla = "_Not Found " & RefAnt
                    With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
                        If FirstCase Then
                            .Value = .Value & vbLf & vbLf & "   - Falta Añadir las Siglas de algún Evento en la Tabla: Tb_TPV_Refs de la Hoja: TPV_Events." & _
                                                     vbLf & "          - Siglas del Evento con Referencia: " & RefAnt
                            FirstCase = False
                        Else
                            .Value = .Value & vbLf & "          - Siglas del Evento con Referencia: " & RefAnt
                        End If
                        ActivForm.Repaint
                    End With
                End If
'            End If
        End If
        .Cells(LinDat, C_TPV_Siglas) = Sigla
        .Cells(LinDat, C_TPV_N_Liq) = N_Liq
    Next LinDat
                    With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
                        If FirstCase Then
                            .Value = .Value & vbLf & "     ¡ Todas las Siglas añadidas sin incidencias !" & vbLf
                        Else
                            .Value = .Value & vbLf
                        End If
                        ActivForm.Repaint
                    End With
    End With    '- Lo_TPVpag.DataBodyRange

    '- Borrar los reg. SIN Liquidar de TPV_Tb (para guardar sólo los que tienen datos de liquidación) -------------------------
    Prog_TPV_Tb.Visible = xlSheetVisible
    Prog_TPV_Tb.Unprotect
    Prog_TPV_Tb.Select
    Call Rut_LstObj_DataBodyRange_Filtered_DEL(Lo_TPV, C_TPV_Siglas, "=")   '- Si NO tienen Sigla = SIN Liquidar
    
    ''- Borrar los Registros de TPV_Tb, de Fecha_Fin anteriores al año contable     --------------------------------------
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_F_Fin, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
    Call Rut_LstObj_DataBodyRange_Filtered_DEL(Lo_TPV, C_TPV_F_Fin, "<1/1/" & Prog__APP.Range("APP_AñoCont") - 1)

    '- Guardar Reg. Liquidados de TPV_Tb en TPV-Liqdatos     -------------------------------------------------------------------------------
    Call Rut_LstObj_DataBodyRange_Copy(Lo_TPV, Lo_Liqdatos, True)   ' (Lo_Source,Lo_Target, DataBodyRange.Delete)
        '- Visualizo el progreso -----
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "Importo el Fichero de Pagos.  "
            ActivForm.Repaint
        End With
        
    '- Copia los nuevos datos TPV_Pag en TPV_Tb que está vacía  ----------------------------------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "Copio TPV_Pagos a TPV_Tb. Ahora hay " & Format(Lo_TPVpag.ListRows.Count, "#,##0") & " reg."
            ActivForm.Repaint
        End With
    Call Rut_LstObj_DataBodyRange_Copy(Lo_TPVpag, Lo_TPV, True)   ' (Lo_Source,Lo_Target, DataBodyRange.Delete)
    Call Rut_WrkSheet_Vaciar(Prog_TPV_Pag.Name)
        '- Visualizo el progreso ---------------------------------------------------------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & "Añado la información a los Pagos Liquidados a TPV_Tb desde TPV_LiqDatos.  "
            ActivForm.Repaint
        End With
    
    '----------------------------------------------------------------------------------------------------------------------
    '- Copiar los Datos de TPV-Liqdatos al nuevo TPV_Tb     <<<<-----------------------------------------------------------
    Dim LastLinDat  As Long:        LastLinDat = Lo_Liqdatos.ListRows.Count
    Dim LastLinTPV  As Long:        LastLinTPV = Lo_TPV.ListRows.Count
    Dim LinTPV      As Long:        LinTPV = 1
    Dim col         As Integer
    Dim NotFound    As Integer:     NotFound = 0
    Dim Found       As Integer:     Found = 0
    Dim TextErr     As String

    Prog_TPV_Tb.Select
    Call Rut_LstObj_Sort(Lo_Liqdatos, C_TPV_Ref, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_Ref, xlAscending, True)
    Dim RowDat          As ListRow
    Dim RowTPV          As ListRow
    
    Lo_Liqdatos.ListColumns(C_TPV_Ctrl).DataBodyRange.ClearContents
    TBx_Informe_Hist = ActivForm.Controls("TBx_Informe")
    With Lo_TPV.DataBodyRange
    For LinDat = 1 To LastLinDat
        Set RowDat = Lo_Liqdatos.ListRows(LinDat)
        If RowDat.Range(C_TPV_Siglas) = "Not JyC" Then GoTo NextReg
        If LinTPV > LastLinTPV Then Exit For
        Set RowTPV = Lo_TPV.ListRows(LinTPV)
        Select Case RowDat.Range(C_TPV_Ref)
            '- Equal reg >>> Copy Data   -----------------------------------------------------
            Case Is = RowTPV.Range(C_TPV_Ref)
                If RowTPV.Range(C_TPV_Imp).Value <> RowDat.Range(C_TPV_Imp).Value Or RowTPV.Range(C_TPV_ComBco).Value <> RowDat.Range(C_TPV_ComBco).Value Or RowTPV.Range(C_TPV_Neto).Value <> RowDat.Range(C_TPV_Neto).Value Then
                    If RowDat.Range(C_TPV_Neto).Value <> "Esperando comisión" Then
                        MsgBx_Msg = "¡¡¡ Error, Cambio de Importes !!!" & vbCrLf & _
                                    "TPV_Imp:  '" & Format(RowTPV.Range(C_TPV_Imp), "#0.00") & "' -.- Dato Existente:  '" & RowDat.Range(C_TPV_Imp) & "'" & vbCrLf & _
                                    "TPV_Neto:  '" & Format(RowTPV.Range(C_TPV_Neto), "#0,00") & "' -.- Dato Existente:  '" & RowDat.Range(C_TPV_Neto) & "'" & vbCrLf & _
                                    "TPV_Ref:  " & RowTPV.Range(C_TPV_Ref) & "   -.- Num.Liq.: " & RowDat.Range(C_TPV_N_Liq)
                        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Stop"): Form_MsgBox.Show   '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
                        GoTo NextReg
                    End If
                End If
                If RowTPV.Range(C_TPV_Siglas) <> RowDat.Range(C_TPV_Siglas) Then
                    With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
                        TextErr = "   - Error; en la Sigla:  " & RowTPV.Range(C_TPV_Siglas) & ", tenía: " & RowDat.Range(C_TPV_Siglas) & ", la cambio (¡Verificar!)."
                        RowDat.Range(C_TPV_Siglas) = RowTPV.Range(C_TPV_Siglas)
                        If InStr(.Value, TextErr) = 0 Then
                            .Value = TBx_Informe_Hist & vbLf & TextErr
                            ActivForm.Repaint
                            TBx_Informe_Hist = .Value
                        End If
                    End With
                End If
                For col = C_TPV_Dias To C_TPV_PMP
                    RowTPV.Range(col) = RowDat.Range(col)
                Next
                RowTPV.Range(C_TPV_Ctrl) = "Found"
                RowDat.Range(C_TPV_Ctrl) = "Found"
                LinTPV = LinTPV + 1
                Found = Found + 1
                GoTo NextReg
            '- Not Find Data reg -----------------------------------------------------
            Case Is < RowTPV.Range(C_TPV_Ref)
                RowDat.Range(C_TPV_Ctrl) = "NotFound"
                NotFound = NotFound + 1
                GoTo NextReg
            '- No Data reg -----------------------------------------------------
            Case Is > RowTPV.Range(C_TPV_Ref)
                LinDat = LinDat - 1
                LinTPV = LinTPV + 1
        End Select
NextReg:    '-------------------------------------------------------------------
        '- Visualizo el progreso ---------------------------------------------------------------------------------------
        If LinTPV Mod 100 = 0 Then
            With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
                .Value = TBx_Informe_Hist & vbLf & vbLf & "   - Incorporando datos en Pagos-TPV a " & Format(LinTPV, "#,##0") & " de " & Format(CantReg, "#,##0") & " reg."
                ActivForm.Repaint
            End With
        End If

    Next LinDat
    '--------------------------------------------------------------------------------------------------------------------
    End With    ' Lo_TPV.DataBodyRange
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = TBx_Informe_Hist & vbLf & vbLf & _
                    "Incorporados datos en Pagos-TPV a " & Format(Found, "#,##0") & " de " & Format(CantReg, "#,##0") & " reg." & vbLf & _
                    "   - Not Found: " & Format(NotFound, "#,##0") & "  (Existían y han desaparecido)" & vbLf & _
                    "   - Found: " & Format(Found, "#,##0") & "  (Existían y se ha incorporado sus datos)"
            ActivForm.Repaint
        End With
    '   Seleccionar los reg. de Datos-TPV "Not Found" en el nuevo TPV y añadirlos en el nuevo TPV  -------------------------------------------
    If NotFound > 0 Then
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & vbLf & "Tenemos que añadir registros de datos anteriores al nuevo TPV." & vbCrLf & "Añadiremos: " & NotFound & "reg."
            ActivForm.Repaint
        End With
        Call Rut_LstObj_DataBodyRange_Filter_And_Copy(Lo_Liqdatos, Lo_TPV, C_TPV_Ctrl, "=NotFound", False)
        Call Rut_LstObj_Sort(Lo_TPV, C_TPV_Ref, xlAscending, True)
    End If
    RuT_Generar_TPV_List
        '- Visualizo el progreso ---------------------------------------------------------------------------------------
        With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
            .Value = .Value & vbLf & vbLf & "¡ Proceso concluido ! día: " & Format(Now(), "dd-mmm-yyyy ""a las"" hh:mm") & " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg."
            .SelStart = Len(.Text)
            .SetFocus
            ActivForm.Repaint
            Prog__APP.Range("APP_Task_Inf") = .Value
        End With

    Prog__APP.Range("APP_Date_Imp_TPV") = Date
   'Prog__APP.Range("APP_Date_Imp_TPV") = Format(Now(), "dd-mmm-yy")

'''    Call RefreshRibbon

Restablecer_Valores:
Rut_On_Functions
    Prog_TPV_Pag.Visible = xlSheetVeryHidden
    Prog_TPV_List.Visible = xlSheetVeryHidden
    Prog_TPV_Tb.Visible = xlSheetVeryHidden
    Sheets(ActivSheet).Visible = xlSheetVisible
    Sheets(ActivSheet).Select
'Application.Speech.Speak "Proceso completado puede verificar el resultado."
End Sub     '- Import_TPV_Pagos
' ==================================================================================================================================

' ==================================================================================================================================
Sub Rut_LstObj_Col_Format_Date(ByRef LoBjDatos As ListObject, Ccol As Integer)     '############## Convierte Número Texto en fecha ###################
    With LoBjDatos.DataBodyRange
        Dim Lin    As Long
        With .Columns(Ccol)
            For Lin = 1 To .Rows.Count
                    .Cells(Lin) = Left(.Cells(Lin), 10)
            Next Lin
        End With
        .Columns(Ccol).Select
        Selection.TextToColumns DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 4), TrailingMinusNumbers:=True
        Selection.NumberFormat = "dd-mm-yyyy"
    End With
End Sub
' ==================================================================================================================================


' ==================================================================================================================================
Sub RuT_Generar_TPV_List()      '- Añadir la hoja 1 de cada Excel seleccionado -------------
' ==================================================================================================================================
'Rut_Off_Functions

    Prog_TPV_Tb.Visible = xlSheetVisible
    Prog_TPV_Tb.Select

    Dim APP_AñoCont     As Integer:             APP_AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim Lo_TPV          As ListObject:          Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Dim Lo_Lst          As ListObject:          Set Lo_Lst = Prog_TPV_List.ListObjects(1)

    Dim RwTPV           As ListRow
    Dim RwLst           As ListRow
    Dim LinTPV          As Long:            LinTPV = 1
    Dim LinTPL          As Integer:         LinTPL = 1
    Dim RefAnt          As String
    Dim ImpRec          As Double
    Dim ComBCO          As Double

    ' --------------------------------=============  Preparar Tabla de Prog_TPV_List ==================
    Call Rut_LstObj_WrkSht_Preparar(Prog_TPV_List)  '- Quita Filtros y muestra todas la Columnas y Filas Ocultas
    Call Rut_LstObj_WrkSht_Preparar(Prog_TPV_List)  '- Quita Filtros y muestra todas la Columnas y Filas Ocultas
    Prog_TPV_List.Visible = xlSheetVisible
    Prog_TPV_List.Select
    Prog_TPV_List.Unprotect
    With Prog_TPV_List.ListObjects(1)
        .ShowTotals = False
        .ListColumns(C_TPL_Pdte_ComBco).DataBodyRange.ClearContents

        Call Rut_LstObj_Sort(Lo_TPV, C_TPV_EventoRef, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_LstObj_Sort(Lo_Lst, C_TPL_EventoRef, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Prog_TPV_List.Unprotect
        Set RwLst = .ListRows(LinTPL)
        For LinTPV = 1 To Lo_TPV.ListRows.Count
            Set RwTPV = Lo_TPV.ListRows(LinTPV)
            If RwTPV.Range(C_TPV_F_Fin) < CDate("1/7/" & APP_AñoCont - 1) Then GoTo NextReg
            If RwTPV.Range(C_TPV_EventoRef) < Lo_Lst.DataBodyRange.Cells(LinTPL, 1) Then GoTo NextReg
            If RwTPV.Range(C_TPV_EventoRef) > Lo_Lst.DataBodyRange.Cells(LinTPL, 1) Then
                LinTPL = LinTPL + 1
                If LinTPL > Lo_Lst.ListRows.Count Then
                    Set RwLst = .ListRows.Add
                Else
                    Set RwLst = .ListRows(LinTPL)
                    LinTPV = LinTPV - 1
                    GoTo NextReg
                End If
            End If

            If IsNumeric(RwTPV.Range(C_TPV_Neto)) Then
                ComBCO = RwTPV.Range(C_TPV_ComBco)
                ImpRec = RwTPV.Range(C_TPV_ComBco) + RwTPV.Range(C_TPV_Neto)
            Else
                ImpRec = 0
                ComBCO = 0
            End If
            If RwTPV.Range(C_TPV_EventoRef) = RefAnt Then
                With RwLst
                    .Range(C_TPL_ImpRec) = .Range(C_TPL_ImpRec) + ImpRec
                    .Range(C_TPL_ComBco) = .Range(C_TPL_ComBco) + ComBCO
                    .Range(C_TPL_ImpNeto) = .Range(C_TPL_ImpNeto) + ImpRec - ComBCO
                    If .Range(C_TPL_F_Vto) <> RwTPV.Range(C_TPV_F_Fin) Then Debug.Print RwTPV.Range(C_TPL_EventoRef) & " - ERRor Date: " & RwTPV.Range(C_TPV_F_Fin)
                End With
            Else
                RefAnt = RwTPV.Range(C_TPV_EventoRef)
'                Set RwLst = .ListRows.Add
                With RwLst
                    .Range(C_TPL_EventoRef) = RwTPV.Range(C_TPV_EventoRef)
                    .Range(C_TPL_EventoNom) = RwTPV.Range(C_TPV_EventoNom)
                    .Range(C_TPL_F_Vto) = RwTPV.Range(C_TPV_F_Fin)
                    .Range(C_TPL_ImpRec) = ImpRec
                    .Range(C_TPL_ComBco) = ComBCO
                    .Range(C_TPL_ImpNeto) = ImpRec - ComBCO
                End With
            End If

            '- Busco si algún cobro está pendiente de recibir importe de comisión ------
            If RwTPV.Range(C_TPV_Neto) = "Esperando comisión" Then
                With RwLst
                    .Range(C_TPL_Pdte_ComBco).Value = "Pdte. Comisión"
                End With
            End If
'    Debug.Print ImpRec, RwTPV.Range(C_TPV_ComBco), RwTPV.Range(C_TPV_Neto), RwLst.Range(C_TPL_ComBco), RwLst.Range(C_TPL_ImpNeto)
NextReg:
        Next LinTPV
        .ShowTotals = True
    End With

    Call Rut_LstObj_Sort(Lo_Lst, C_TPL_EventoRef, xlAscending, True)    ' "xlAscending"

    Range("d5") = Range("d5")
    Range("d4") = Now()
'    Prog__APP.Range("APP_Last_Inform") = Format(Now(), "dd-mmm-yy hh:mm")
'    Call RefreshRibbon

'''    MsgBx_Msg = "¡ Informe completado. !"
'''    MsgBx_Title = "Proceso: Generar Informe de Pagos de Pruebas Selectivas por TPV."
'''    Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, False, "OK", , "Ok"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
'''
'''Application.ScreenUpdating = True


Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Prog_TPV_Tb.Visible = xlSheetHidden
'Rut_On_Functions
End Sub     ' RuT_Generar_TPV_List   --------------------------------------------------------------------------------------------
'===================================================================================================================================


' ==================================================================================================================================
Sub Rut_Prueba_Form_Mensaje_Progreso()
    Dim LinDat  As Long
    Dim LinDat2     As Long
        MsgBx_Title = "Proceso: Progreso Importar Pagos-TPV"
'MsgBx_Msg = ""
    For LinDat = 1 To 100000
        If LinDat Mod 10000 = 0 Then
            For LinDat2 = 1 To 100000000: Next
            MsgBx_Msg = MsgBx_Msg & vbLf & "Incorporando Pagos-TPV:  " & Format(LinDat, "#,##0") & " de " & Format(100000, "#,##0")
            Load Form_Progreso: Call Form_Progreso.SetParameter(12, , "NO"): Form_Progreso.Show False '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
            With Form_Progreso.Lb_Mensaje
                .SelStart = Len(.Text)
                .SetFocus
            End With
            Form_Progreso.Repaint
        End If
    Next LinDat
'    Unload Form_Progreso
    Load Form_Progreso: Call Form_Progreso.SetParameter(12, True, "OK", , "Exclam"): Form_Progreso.Show False '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
    Form_Progreso.Repaint
End Sub

''' ==================================================================================================================================
''            Sub kk_Rut_Prueba_Rut_Progreso()
'''                MsgBx_Rut = "Rut_Prueba_Rut_Progreso"
''                Prog__APP.Range("APP_Task_Rut") = "Rut_Prueba_Rut_Progreso"
''                MsgBx_Title = "Proceso: Progreso Importar Pagos-TPV"
''                MsgBx_Msg = "Incorporando Pagos-TPV:  "
''                Form_Run_Rut.Show
''            End Sub
''' ==================================================================================================================================
''Sub Rut_Prueba_Rut_Progreso()
''    Dim LinDat  As Long
''    Dim LinDat2     As Long
''    Dim ActivForm    As Object
''    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)
''
'''MsgBx_Msg = ""
''    For LinDat = 1 To 100000
''        If LinDat Mod 3000 = 0 Then
''            For LinDat2 = 1 To 100000000: Next
''            With ActivForm.Controls("TBx_Informe")  '- Mostrar Progreso -----------
''                .Value = .Value & vbLf & "Incorporando Pagos-TPV:  " & Format(LinDat, "#,##0") & " de " & Format(100000, "#,##0")
''                .SelStart = Len(.Text)
''                .SetFocus
''            End With
''            Form_Run_Rut.Repaint
''        End If
''    Next LinDat
''End Sub





