Attribute VB_Name = "M_132_Añadir_Núm_JIs_al_Inf"
'2026-02-14
Option Explicit

'- Rellena la Tabla Informe_Recibos con todos los Números de JI's de BDatos. ------------------------

' ==================================================================================================================================
Sub Rut_Añadir_a_Inf_Recibos_JIs_de_BDatos()
' ==================================================================================================================================
Debug.Print "Rut_Añadir_a_Inf_Recibos_JIs_de_BDatos"

        Dim Nom_Inf         As String:          Nom_Inf = "Inf_Recibos"
        Dim Sht_Inf         As Worksheet:       Set Sht_Inf = Sht__Inf_Recibos_TIO
        Dim Lo_Inf          As ListObject:      Set Lo_Inf = Sht_Inf.ListObjects(1)
        Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)

Dim Tp_Rec          As String
Dim Concept         As String
Dim MenúAux_Msg     As String
Dim F_Inf           As Integer
Dim rowfind         As Variant
Dim RwJI            As ListRow
Dim RngJIs          As Range
Dim Celda           As Range

    Application.ScreenUpdating = False
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    
    Sht__BD.Unprotect
    Lo_BD.ShowTotals = False
    Sht_Inf.Visible = xlSheetVisible
    Sht_Inf.Select
    Call Rut_Lo_WrkSht_Preparar(Sht_Inf)
    '- Borrar Columnas de JI's
    Lo_Inf.DataBodyRange.Columns(InfRec_JI_Emi_Adm).ClearContents
    Lo_Inf.DataBodyRange.Columns(InfRec_JI_Emi_Acad).ClearContents
    Lo_Inf.DataBodyRange.Columns(InfRec_AD_Emi_Adm).ClearContents
    Lo_Inf.DataBodyRange.Columns(InfRec_AD_Emi_Acad).ClearContents
    Lo_Inf.DataBodyRange.Columns(InfRec_JI_443_Adm).ClearContents
    Lo_Inf.DataBodyRange.Columns(InfRec_JI_443_Acad).ClearContents
    
    '------------ Preparo Sht__BD y Ordeno por Tipo_Tasa y Concepto_Económico -----------------------------------
    Call Rut_Lo_WrkSht_Preparar(Sht_Inf)
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Lo_BD.ShowTotals = False
    Call Rut_Lo_Sort(Lo_BD, BD_Tipo_Rec, xlAscending, True)     '- Emitida, Aplazado, EjeAnt, ADxAplz, Añeja...
    Call Rut_Lo_Sort(Lo_BD, BD_Concepto, xlAscending, False)   '- 1303.00 1310.00 1311.00 Etc.
    Sht_Inf.Select
    Lo_Inf.ShowTotals = False
    Call Rut_Lo_Sort(Lo_Inf, InfRec_TipRec, xlAscending, True)     '- Emitida, Aplazado, EjeAnt, ADxAplz, Añeja...
    Call Rut_Lo_Sort(Lo_Inf, InfRec_ConcptEco, xlAscending, False)   '- 1303.00 1310.00 1311.00 Etc.
    
    '- Preguntar si quiero los nº de JI's de EFP y CFCyAFC ------------------------------------------------------------
    Dim Sw_VerJIsEP As Boolean
    Dim EFP_JIs     As VbMsgBoxResult
        EFP_JIs = MsgBox("¿ Quieres los Ji's de las Enseñanzas Propias: EFP y CFCyAFC ?", _
                            vbYesNo + vbQuestion + vbDefaultButton2, "Proceso: Buscar los nº de JI's en BDatos.")
        If EFP_JIs = vbYes Then Sw_VerJIsEP = True Else Sw_VerJIsEP = False
    
    
    '-------------- Relleno Tabla para JI's de Tasas por Tipo y por Concepto ------------------------------------
    '- Para cada Tipo de Recibos; Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. -----------------------------------
    For F_Inf = 1 To Lo_Inf.ListRows.Count
        Set RwJI = Lo_Inf.ListRows(F_Inf)
        Tp_Rec = RwJI.Range(InfRec_TipRec)
        Concept = "1" & RwJI.Range(InfRec_ConcptEco)
        If Not Sw_VerJIsEP And Left(Concept, 4) = "1311" Then
            RwJI.Range(InfRec_JI_Emi_Adm) = "No requerido"
            GoTo Sigiente_Concepto
        End If
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '- Filtrar Tipo_Rec y Concepto
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="=" & Tp_Rec
        Lo_BD.Range.AutoFilter Field:=BD_Concepto, Criteria1:="=" & Val(Concept)
        rowfind = Lo_BD.Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1   '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
                                        'Debug.Print Tp_Rec, Concept, RowFind & " reg."
        If rowfind > 0 Then
            Dim DiccJIs     As Object
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '- Identificar Todos los JI's de cada Tipo_Rec y Concepto, Crea Dictionary para valores únicos (eficiente para grandes datos)
            
            '1º-BD_JI_Emi_Adm-------------------- InfRec_JI_Emi_Adm
            Set RngJIs = Lo_BD.ListColumns(BD_JI_Emi_Adm).DataBodyRange.SpecialCells(xlCellTypeVisible)
            Set DiccJIs = CreateObject("Scripting.Dictionary")
            ' Recorrer SOLO celdas visibles (mucho más rápido que todo el rango)
            For Each Celda In RngJIs.Cells
                If Not IsEmpty(Celda.Value) Then
                    DiccJIs(Celda.Value) = True  ' Clave única automática
                End If
            Next Celda
                                        'Debug.Print Tp_Rec, Concept, RowFind & " reg.", DiccJIs.Count
            If DiccJIs.Count > 0 Then
                ' Colocar en celda destino
                RwJI.Range(InfRec_JI_Emi_Adm).Value = Join(DiccJIs.Keys, ", ")   ' Construir cadena separada por comas
            Else
                RwJI.Range(InfRec_JI_Emi_Adm).Value = "·-·"
            End If
            Set DiccJIs = Nothing
            Set RngJIs = Nothing
            
            '2º-BD_JI_Emi_Acad-------------------- InfRec_JI_Emi_Acad
            Set RngJIs = Lo_BD.ListColumns(BD_JI_Emi_Acad).DataBodyRange.SpecialCells(xlCellTypeVisible)
            Set DiccJIs = CreateObject("Scripting.Dictionary")
            ' Recorrer SOLO celdas visibles (mucho más rápido que todo el rango)
            For Each Celda In RngJIs.Cells
                If Not IsEmpty(Celda.Value) Then
                    DiccJIs(Celda.Value) = True  ' Clave única automática
                End If
            Next Celda
                                        'Debug.Print Tp_Rec, Concept, RowFind & " reg.", DiccJIs.Count
            If DiccJIs.Count > 0 Then
                ' Colocar en celda destino
                RwJI.Range(InfRec_JI_Emi_Acad).Value = Join(DiccJIs.Keys, ", ")   ' Construir cadena separada por comas
            Else
                RwJI.Range(InfRec_JI_Emi_Acad).Value = "·-·"
            End If
            Set DiccJIs = Nothing
            Set RngJIs = Nothing
            
            '3º-BD_AD_Emi_Adm-------------------- InfRec_AD_Emi_Adm
            Set RngJIs = Lo_BD.ListColumns(BD_AD_Emi_Adm).DataBodyRange.SpecialCells(xlCellTypeVisible)
            Set DiccJIs = CreateObject("Scripting.Dictionary")
            ' Recorrer SOLO celdas visibles (mucho más rápido que todo el rango)
            For Each Celda In RngJIs.Cells
                If Not IsEmpty(Celda.Value) Then
                    DiccJIs(Celda.Value) = True  ' Clave única automática
                End If
            Next Celda
                                        'Debug.Print Tp_Rec, Concept, RowFind & " reg.", DiccJIs.Count
            If DiccJIs.Count > 0 Then
                ' Colocar en celda destino
                RwJI.Range(InfRec_AD_Emi_Adm).Value = Join(DiccJIs.Keys, ", ")   ' Construir cadena separada por comas
            Else
                RwJI.Range(InfRec_AD_Emi_Adm).Value = "·-·"
            End If
            Set DiccJIs = Nothing
            Set RngJIs = Nothing
            
            '4º-BD_AD_Emi_Acad-------------------- InfRec_AD_Emi_Acad
            Set RngJIs = Lo_BD.ListColumns(BD_AD_Emi_Acad).DataBodyRange.SpecialCells(xlCellTypeVisible)
            Set DiccJIs = CreateObject("Scripting.Dictionary")
            ' Recorrer SOLO celdas visibles (mucho más rápido que todo el rango)
            For Each Celda In RngJIs.Cells
                If Not IsEmpty(Celda.Value) Then
                    DiccJIs(Celda.Value) = True  ' Clave única automática
                End If
            Next Celda
                                        'Debug.Print Tp_Rec, Concept, RowFind & " reg.", DiccJIs.Count
            If DiccJIs.Count > 0 Then
                ' Colocar en celda destino
                RwJI.Range(InfRec_AD_Emi_Acad).Value = Join(DiccJIs.Keys, ", ")   ' Construir cadena separada por comas
            Else
                RwJI.Range(InfRec_AD_Emi_Acad).Value = "·-·"
            End If
            Set DiccJIs = Nothing
            Set RngJIs = Nothing
            
            '5º-BD_JI_443_Adm-------------------- InfRec_JI_443_Adm
            Set RngJIs = Lo_BD.ListColumns(BD_JI_443_Adm).DataBodyRange.SpecialCells(xlCellTypeVisible)
            Set DiccJIs = CreateObject("Scripting.Dictionary")
            ' Recorrer SOLO celdas visibles (mucho más rápido que todo el rango)
            For Each Celda In RngJIs.Cells
                If Not IsEmpty(Celda.Value) Then
                    DiccJIs(Celda.Value) = True  ' Clave única automática
                End If
            Next Celda
                                        'Debug.Print Tp_Rec, Concept, RowFind & " reg.", DiccJIs.Count
            If DiccJIs.Count > 0 Then
                ' Colocar en celda destino
                RwJI.Range(InfRec_JI_443_Adm).Value = Join(DiccJIs.Keys, ", ")   ' Construir cadena separada por comas
            Else
                RwJI.Range(InfRec_JI_443_Adm).Value = "·-·"
            End If
            Set DiccJIs = Nothing
            Set RngJIs = Nothing
            
            '6º-BD_JI_443_Acad-------------------- InfRec_JI_443_Acad
            Set RngJIs = Lo_BD.ListColumns(BD_JI_443_Acad).DataBodyRange.SpecialCells(xlCellTypeVisible)
            Set DiccJIs = CreateObject("Scripting.Dictionary")
            ' Recorrer SOLO celdas visibles (mucho más rápido que todo el rango)
            For Each Celda In RngJIs.Cells
                If Not IsEmpty(Celda.Value) Then
                    DiccJIs(Celda.Value) = True  ' Clave única automática
                End If
            Next Celda
                                        'Debug.Print Tp_Rec, Concept, RowFind & " reg.", DiccJIs.Count
            If DiccJIs.Count > 0 Then
                ' Colocar en celda destino
                RwJI.Range(InfRec_JI_443_Acad).Value = Join(DiccJIs.Keys, ", ")   ' Construir cadena separada por comas
            Else
                RwJI.Range(InfRec_JI_443_Acad).Value = "·-·"
            End If
            Set DiccJIs = Nothing
            Set RngJIs = Nothing
            
        
        Else
'                RwJI.Range(InfRec_JI_Emi_Adm) = "Sin Recibos" & F_Inf
'                RwJI.Range(InfRec_AD_Emi_Adm) = "Sin Recibos" & F_Inf
            Debug.Print "Sin Recibos: ", Tp_Rec, Concept
        End If
        
Sigiente_Concepto:
            
    Next F_Inf
    
'GoTo Finalizar
    
Finalizar:
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    
'    '- Visualizo el progreso ---------------------------------------------------------------------------------------
'    MenúAux_Msg = Format(Now, "hh:mm:ss") & "  Tabla generada." & vbCrLf & _
'        vbCrLf & Format(Now, "hh:mm:ss") & "  Realizado el: " & Date & "  " & "-   Tiempo transcurrido: " & Round(Timer - H_Inicio, 2) & " seg."
'    MsgBox MenúAux_Msg
    
    Lo_BD.ShowTotals = True
    Lo_Inf.ShowTotals = True
    ' Debug.Print   Lo_Inf.ListColumns(InfRec_Tot_Emi).Range.Column === ¡¡Columna de comienzo!!
'    Sht_Inf.Columns(Lo_Inf.ListColumns(InfRec_Tot_Emi).Range.Column).Resize(, 10).EntireColumn.Hidden = True
'    Sht_Inf.Columns(Lo_Inf.ListColumns(InfRec_Acad_Emi).Range.Column).Resize(, 3).EntireColumn.Hidden = True
    
    Application.ScreenUpdating = True
    Application.Calculation = Sw_Calculation
    Application.DisplayAlerts = True
    Application.Speech.Speak "Proceso completado puede verificar el resultado.", True
End Sub






