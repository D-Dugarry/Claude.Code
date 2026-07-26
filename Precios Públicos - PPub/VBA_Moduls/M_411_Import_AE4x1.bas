Attribute VB_Name = "M_411_Import_AE4x1"
'2026-01-25
Option Explicit

'   EN ESTE MÓDULO PROCESO TODO LO QUE PUEDO EN EL ClsBk PARA TRABAJAR AL MÁXIMO EN LA RAM

' Rut_Lo_Import_LoData_LoDefCol_LSace06
'
'    Importar Última Consulta de LSsace06 de Curso_Acad_Ant o Curso_Acad_Pos
'
'        M_411_Import_CFCyAFC:
'        - Rut_Lo_Import_LoData_LoDefCol_AE4x1
'            - Rut_File_Select_V2 (4 veces; una por cada Fich.)
'            - Con el ClsBk: (RAM)
'            - Compruebo ClsBk:
'            -            1º Si he dado una SheetNom ver si Existe, sino ERROR, Cancelo
'            -            2º Comprobar que Existe Lo_Bdatos
'            -            3º Comprobar que la Tabla Lo_ClsBk tiene datos
'            -            4º Comprobar que la cabecera de la Tabla Lo_ClsBk corresponde con la establecida en la LoDefCol
'            - Proceso ClsBk:
'                        - Borrar Recibos ACont_Emi = AñoCont-1 y Acont_Cob <> AñoCont
'                        - Borrar Recibos ACont_Emi > AñoCont
'                        - Borrar Datos de Rec. con F_Cob > APP_FechCierreCont: Vaciar/Clear las Columnas BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag y BD_HTipCob
'                        - Borrar Recibos M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b --
'            - Copy ClsBk:
'            -          Si SW_Del_LoData=true Borrar Lo_AE4x1 (sólo la 1ª vez)
'            -          Añadir Lo_ClsBk al final de Lo_AE4x1.

'- ----------------------------------------------------------------------------------------------------------------------------
'- Seleccionar Excel pero no lo importa, sólo lo abre -------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Import_LoData_LoDefCol_AE4x1(Lo_AE4x1 As ListObject, _
                                  LoDefCol As ListObject, _
                                  Col_Header As Integer, _
                                  Arch_New_Name As String, _
                                  Optional NameFileAE4 As String, _
                                  Optional PathFileAE4 As String, _
                                  Optional SheetNom As String = "", _
                                  Optional SW_Del_LoData As Boolean = False)
                             
Debug.Print ">>> Rut_Lo_Import_LoData_LoDefCol_AE4x1"
    Dim Ccol            As Integer
    Dim RowsFind        As Variant
    Dim ArchRequest     As String:      ArchRequest = Arch_New_Name
    Dim TxT_ProgIni     As String:      TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Dim TxT_Progreso    As String
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim FechCierreCont  As String:      FechCierreCont = Prog__APP.Range("APP_FechCierreCont")
    Dim Ws_AE4          As Worksheet:   Set Ws_AE4 = Lo_AE4x1.Parent
    Dim Rng_Informe     As Range:       Set Rng_Informe = Ws_AE4.Range("g5")

    Dim SwScrUp     As Boolean:             SwScrUp = Application.ScreenUpdating:       Application.ScreenUpdating = False
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
        
    '- Select File -------------------------------------------------------------------------------------
    Call Rut_File_Select_V2("Seleccionar el Nuevo Fichero Excel " & Arch_New_Name & ": ", Arch_New_Name, "Excel", "*.xlsm")
    Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NameFileAE4, PathFileAE4)
    If Arch_New_Name = "Cancel" Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ Cancelado a petición del Usuario !    "
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        '- Visualizo el progreso  <<<<>>>>
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
        Arch_New_Name = "Cancel"
        Exit Sub
    End If
    '- Comprueba que se ha seleccionado el nombre adecuado de Excel. --------------------------
    If Left(NameFileAE4, Len(ArchRequest)) <> ArchRequest Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ Cancelado ¡" & vbLf & "El fichero debe ser:" & vbLf & ArchRequest & vbLf & vbLf & _
                    "  y se ha seleccionado:" & vbLf & NameFileAE4
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & vbLf & Now()
        '- Visualizo el progreso  <<<<>>>>
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), 0)
        Arch_New_Name = "Cancel"
        Exit Sub
    End If
        '- Visualizo el progreso  <<<<>>>>
        TxT_Progreso = ActivForm.Controls("TBx_Informe")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Excel Seleccionado: " & NameFileAE4, LastTimeLap, , , TxT_Progreso)
    
    '- --------------------------------------------------------------------------------------------------------------
    '- --------------------------------------------------------------------------------------------------------------
    '- Compruebo ClsBk:
    '            - Comprobar que la Hoja Existe, SI hemos solicitado una hoja concreta para copiar
    '            - Crear, si NO Existe ListObject Lo_ClsBk
    '            - Comprobar que la Tabla Lo_ClsBk No está vacía
    '            - Comprobar que la cabecera de la Tabla Lo_ClsBk corresponde con la establecida en la LoDefCol
    '- --------------------------------------------------------------------------------------------------------------
    Prog__APP.Range("SW_WB_Deactivate") = False     '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ---->>>
    Dim SheetIndx       As Integer:         SheetIndx = 1
    Dim ClosedBook      As Workbook
    Dim Ws_ClsBk    As Worksheet
    Dim Lo_ClsBk    As ListObject
    Set ClosedBook = Workbooks.Open(Arch_New_Name, ReadOnly:=True)
    '- Comprobar que la Hoja Existe SI hemos solicitado una hoja concreta para copiar ---------------------------------
    If SheetNom <> "" Then
        SheetIndx = 0
        For Each Ws_ClsBk In ClosedBook.Sheets
            If Ws_ClsBk.Name = SheetNom Then
                SheetIndx = Ws_ClsBk.Index
                Exit For
            End If
        Next Ws_ClsBk
        If SheetIndx = 0 Then
            MsgBx_Title = "Proceso: Importar " & ArchRequest
            MsgBx_Msg = "¡ La Sheet NO existe !    "
            MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
            '- Visualizo el progreso  <<<<>>>>
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
            GoTo Cancel_Rut
        End If
    End If
    '-  Crear si NO Existe ListObject Lo_ClsBk ---------------------------------
    Set Ws_ClsBk = ClosedBook.Sheets(SheetIndx)
    If Ws_ClsBk.ListObjects.Count = 0 Then
            Ws_ClsBk.UsedRange.Cells(1, 1).Select  'Posicionar cursor
        Set Lo_ClsBk = Ws_ClsBk.ListObjects.Add(xlSrcRange, Ws_ClsBk.UsedRange, , xlYes)
    Else
        Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
        Ws_ClsBk.Unprotect
        Lo_ClsBk.ShowTotals = False
    End If
    '- Comprobar que la Tabla Lo_ClsBk No está vacía ---------------------
    If Lo_ClsBk.ListRows.Count = 0 Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ La tabla no contiene datos !    "
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        '- Visualizo el progreso  <<<<>>>>
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
        GoTo Cancel_Rut
    End If
    '- Comprobar que la cabecera de la Tabla Lo_ClsBk corresponde con la establecida en la LoDefCol ---------------------
    If Not Func_LstObj_ListColumns_DefCol_Check_OK(Lo_ClsBk, LoDefCol, Col_Header) Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ La cabedera de la tabla no coincide !    " & vbLf & "Puede haber un cambio en la estructura de la Tabla"
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        '- Visualizo el progreso  <<<<>>>>
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), LastTimeLap)
        GoTo Cancel_Rut
    End If
            '- Visualizo el progreso  <<<<>>>>
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Excel Seleccionado: " & NameFileAE4, 0, , _
                                                            Format(Lo_ClsBk.ListRows.Count, "#,##0") & " reg", TxT_Progreso)
    '- --------------------------------------------------------------------------------------------------------------
    '- --------------------------------------------------------------------------------------------------------------
    '- Proceso ClosedBook: (para transferir sólo los recibos del AñoCont y no los del C_Acad)
    '        - Borrar Recibos ACont_Emi = AñoCont-1 y Acont_Cob <> AñoCont
    '        - Borrar Recibos ACont_Emi > AñoCont
    '        - Borrar Recibos Importe CERO - Subvencionado- Imp_Rec =0 porque Imp_Dto >0
    '        - Borrar Datos de Rec. con F_Cob > APP_FechCierreCont: Vaciar/Clear las Columnas BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag y BD_HTipCob
    '        - Borrar Recibos M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b --
    '- --------------------------------------------------------------------------------------------------------------
   
'- Borrar Recibos ACont_Emi = AñoCont-1 y Acont_Cob <> AñoCont -----------------------------------------------------------------------------------
    RowsFind = Lo_ClsBk.ListRows.Count
    Rng_Informe = " Fich. con " & Format(RowsFind, "#,##0") & " reg., "
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk, BD_ACont_Emi, "=" & AñoCont - 1, BD_ACont_Cob, "<>" & AñoCont)
    RowsFind = RowsFind - Lo_ClsBk.ListRows.Count
    Rng_Informe = Rng_Informe & " Del " & Format(RowsFind, "#,##0") & " reg. ACont_Emi=" & AñoCont - 1 & " y ACont_Cob" & ChrW(&H2260) & AñoCont
    If RowsFind > 0 Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "Borrados Rec. ACont_Emi = " & AñoCont - 1 & " y ACont_Cob " & ChrW(&H2260) & "  " & AñoCont, 0, _
                                                        Format(RowsFind, " #,##0") & " reg", _
                                                        "quedan " & Format(Lo_ClsBk.ListRows.Count, "#,##0") & " reg")
    Else
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "No hay Rec. ACont_Emi = " & AñoCont - 1 & " y ACont_Cob " & ChrW(&H2260) & "  " & AñoCont, 0)
    End If
    '- Comprobar que quedan registros
    If Lo_ClsBk.DataBodyRange Is Nothing Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbCrLf & String(10, " ") & "¡¡¡ El Excel seleccionado, NO tiene registros procesables AE=4 !!!", 0)
        GoTo Cancel_Rut
    End If

'- Borrar Recibos ACont_Emi > AñoCont -----------------------------------------------------------------------------------
    RowsFind = Lo_ClsBk.ListRows.Count
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk, BD_ACont_Emi, ">" & AñoCont)
    RowsFind = RowsFind - Lo_ClsBk.ListRows.Count
    Rng_Informe = Rng_Informe & ", Del " & Format(RowsFind, "#,##0") & " reg. ACont_Emi > " & AñoCont
    If RowsFind > 0 Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "Borrados Rec. ACont_Emi > " & AñoCont, 0, _
                                                        Format(RowsFind, " #,##0") & " reg", _
                                                        "quedan " & Format(Lo_ClsBk.ListRows.Count, "#,##0") & " reg")
    Else
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "No hay Rec. ACont_Emi > " & AñoCont, 0)
    End If
    '- Comprobar que quedan registros
    If Lo_ClsBk.DataBodyRange Is Nothing Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbCrLf & String(10, " ") & "¡¡¡ El Excel seleccionado, NO tiene registros procesables AE=4 !!!", 0)
        GoTo Cancel_Rut
    End If

'- Borrar Borrar Recibos - Subvencionado-, Imp_Rec =0 porque Imp_Dto >0 ------------------------------
    RowsFind = Lo_ClsBk.ListRows.Count
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk, BD_ImpRec, "=0,00")
    RowsFind = RowsFind - Lo_ClsBk.ListRows.Count
    Rng_Informe = Rng_Informe & ", Del " & Format(RowsFind, "#,##0") & " reg. ACont_Emi > " & AñoCont
    If RowsFind > 0 Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "Borrados Rec. de Matrícula_Cero Subvencionada, ImpRec = 0", 0, _
                                                        Format(RowsFind, " #,##0") & " reg", _
                                                        "quedan " & Format(Lo_ClsBk.ListRows.Count, "#,##0") & " reg")
    Else
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "No hay Rec. de Matrícula_Cero Subvencionada, ImpRec = 0", 0)
    End If
    '- Comprobar que quedan registros
    If Lo_ClsBk.DataBodyRange Is Nothing Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbCrLf & String(10, " ") & "¡¡¡ El Excel seleccionado, NO tiene registros procesables AE=4 !!!", 0)
        GoTo Cancel_Rut
    End If

'- Borrar Recibos M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b -- M013 -- NO M013b --
    With Lo_ClsBk
        RowsFind = Lo_ClsBk.ListRows.Count
        Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_ClsBk, BD_Plan, "=M013")
        RowsFind = RowsFind - Lo_ClsBk.ListRows.Count
        .Range.AutoFilter Field:=BD_Plan, Criteria1:="=M013"
        Rng_Informe = Rng_Informe & ", Del " & Format(RowsFind, "#,##0") & " reg. M013_Bad"
        If RowsFind > 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "Borrados Rec. ¡¡ M013 !!, los malos: ", 0, Format(RowsFind, " #,##0") & " reg. ")
        Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "No hay Rec. ¡¡ M013 !!,  de los malos: ", 0, _
                                                                         Format(RowsFind, " #,##0") & " reg. ")
        End If
    End With
    
'- Renombrar Recibos M013B como M013 --------------------------------------------------------------------------------------------
    With Lo_ClsBk
        Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
        Lo_ClsBk.ShowTotals = False
        Call Rut_Lo_Sort(Lo_ClsBk, BD_Plan, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_Plan, Criteria1:="=M013"
        RowsFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If RowsFind > 0 Then
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "M013"
        End If
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "Renombrar Rec. ¡¡ M013B como M013 !!, los buenos... ", 0, _
                                                                         "en " & Format(RowsFind, " #,##0") & " reg. ")
    End With
    
'- Marcar CFC_UPUA
    With Lo_ClsBk
        Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
        Lo_ClsBk.ShowTotals = False
        Call Rut_Lo_Sort(Lo_ClsBk, BD_Plan, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_Plan, Criteria1:="=UPUA"
        RowsFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        Rng_Informe = Rng_Informe & ", Assign 'CFC_UPUA' a Tipo_EP " & Format(RowsFind, "#,##0") & " reg."
        If RowsFind > 0 Then
            .DataBodyRange.Columns(BD_TIO_EP).SpecialCells(xlCellTypeVisible).Cells.Value = "CFC_UPUA"
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "Asignar 'CFC_UPUA' a los Rec. Plan = 'UPUA'", 0, _
                                                                             "en " & Format(RowsFind, " #,##0") & " reg. ")
        Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "No hay Rec. 'CFC_UPUA'", 0)
        End If
    End With    '- Lo_ClsBk

'- Borrar Datos de Rec. con F_Cob > APP_FechCierreCont: Vaciar/Clear las Columnas BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag y BD_HTipCob
    With Lo_ClsBk
        Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
        .Range.AutoFilter Field:=BD_FCob, Criteria1:=">" & Mid(FechCierreCont, 4, 3) & Left(FechCierreCont, 3) & Right(FechCierreCont, 4)  '- ¡¡¡ EL FORMATO DEBE SER MM/DD/YYYY !!!)
        RowsFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        Rng_Informe = Rng_Informe & ", Clear " & Format(RowsFind, "#,##0") & " reg. F_Cob > " & FechCierreCont
        If RowsFind > 0 Then
            .DataBodyRange.Columns(BD_FCob).SpecialCells(xlCellTypeVisible).Cells.ClearContents
            .DataBodyRange.Columns(BD_ImpCob).SpecialCells(xlCellTypeVisible).Cells.ClearContents
            .DataBodyRange.Columns(BD_FormPag).SpecialCells(xlCellTypeVisible).Cells.ClearContents
            .DataBodyRange.Columns(BD_CtaPag).SpecialCells(xlCellTypeVisible).Cells.ClearContents
            .DataBodyRange.Columns(BD_HTipCob).SpecialCells(xlCellTypeVisible).Cells.ClearContents
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "Clear Data en Rec. con F_Cob > " & FechCierreCont, 0, _
                 "en " & Format(RowsFind, " #,##0") & " reg. ")
        Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, " ") & "No hay Rec. con F_Cob > " & FechCierreCont, 0)
        End If
    End With    '- Lo_ClsBk


'- Completar Imp.Adm. -----------------------------------------------
    Dim TipoEP      As String:      TipoEP = ClosedBook.Sheets(Prog__APP.Name).Range("APP_TitP_o_Curs")
    Call Rut_Assign_Imp_AdmAcad_AE4x1(Lo_ClsBk, TipoEP)

    '- --------------------------------------------------------------------------------------------------------------
    '- --------------------------------------------------------------------------------------------------------------
    '- Copy ClosedBook:
    '            - Si SW_Del_LoData=true Borrar Lo_AE4x1 para iniciarlo
    '            - Añadir Lo_ClsBk al final de Lo_AE4x1.
    '- --------------------------------------------------------------------------------------------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
    Call Rut_Lo_Filtros_Quitar(Lo_AE4x1)

    '-          Si SW_Del_LoData=true Borrar Lo_AE4x1
    If SW_Del_LoData Then
        With Lo_AE4x1
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete
            End If
        End With
    End If
     
    '- Copiar Lo_ClsBk en Lo_AE4x1. ---------------------
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_AE4x1, False)
    RowsFind = Lo_ClsBk.ListRows.Count
    ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing
    Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
        '- Visualizo el progreso  <<<<>>>>  ---------------------------------------------------------------------
            Dim TimeLap2              As Single
            TimeLap2 = LastTimeLap

    ' Restaurar entorno
    Application.ScreenUpdating = SwScrUp
    Application.Calculation = Sw_Calculation
    Application.EnableEvents = True
Debug.Print "<<< Rut_Lo_Import_LoData_LoDefCol_AE4x1"
Exit Sub

Cancel_Rut:
    Arch_New_Name = "Cancel"
    ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing
    Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------


