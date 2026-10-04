Attribute VB_Name = "M_112_Manage_Duplicates"
' Last Rev. 2026-10-04 22:17
'Rev.: 2026-01-22
'M_112_Manage_Duplicates
Option Explicit

'=========================================================================================================================================
'- Gestionar Duplicados ------------------------------------------------------------------------------------------------------------------
'=========================================================================================================================================
Sub RuT_Duplicates_Search(Lo_Data As ListObject, _
                          Lo_DefCol As ListObject, _
                          Lo_Duplic As ListObject, _
                          Colref As Integer, _
                          ColIncidencia As Integer, _
                          Col_H_Incid As Integer)
                          
Debug.Print ">>> RuT_Duplicates_Search"
    Dim TimeLapSub          As Single:      TimeLapSub = LastTimeLap
    Dim DuplFind            As Integer:     DuplFind = 0    '- nº de Rec. que tienen repeticiones
    Dim CantRepe            As Integer:     CantRepe = 1
    Dim MaxNumRepe          As Integer:     MaxNumRepe = 0
    Dim RepsFind            As Integer:     RepsFind = 0
    Dim Cont                As Integer
    Dim FilaReg             As Long:        FilaReg = 1
    Dim TF_BD               As Long:        TF_BD = Lo_Data.ListRows.Count
    Dim Row_Find            As Variant
    Dim Sh_Data             As Worksheet:   Set Sh_Data = Lo_Data.Parent
    Dim Sh_Duplic           As Worksheet:   Set Sh_Duplic = Lo_Duplic.Parent
    
    Call Rut_Lo_WrkSht_Preparar(Sh_Data)
    Call Rut_Lo_WrkSht_Preparar(Sh_Duplic)
    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '- Gestionar Duplicados 1ª Parte: Los Identifica y Marca las Diferencias ------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------------------
    Call Rut_Lo_Sort(Lo_Data, Colref, xlAscending, True)    '- Ordenar primero accelera un montón el borrado
    '- El bucle trabaja en RAM (Rut_Lo_TablaRam): celda a celda tardaba ~17 seg. con 172.000 reg. ----------
    '- Se carga con .Value (no .Value2) porque las incidencias montan textos con las fechas, igual que antes.
    Dim T_Data              As T_TablaRam
    Dim NumCols             As Integer:     NumCols = Lo_Data.ListColumns.Count
    Dim Cont_Col            As Integer
    Dim Col_Comparar()      As Boolean:     ReDim Col_Comparar(1 To NumCols)
    Dim ColsRam()           As Variant:     ReDim ColsRam(1 To 3)
    Dim DiccColor           As Object:      Set DiccColor = CreateObject("Scripting.Dictionary")  '- Celda (Fila*1000+Col) -> ColorIndex
    ColsRam(1) = Colref:    ColsRam(2) = ColIncidencia:     ColsRam(3) = Col_H_Incid
    For Cont_Col = 2 To NumCols                                         '- Col. a comparar según Lo_DefCol (DefC_Compare)
        Col_Comparar(Cont_Col) = Lo_DefCol.ListColumns(DefC_Compare).DataBodyRange(Cont_Col)
        If Col_Comparar(Cont_Col) Then
            ReDim Preserve ColsRam(1 To UBound(ColsRam) + 1)
            ColsRam(UBound(ColsRam)) = Cont_Col
        End If
    Next Cont_Col
    Call Rut_TablaRam_Cargar(T_Data, Lo_Data, ColsRam)
    For FilaReg = 1 To TF_BD                                            '- ClearContents de las Col. de incidencias
        T_Data.Datos(FilaReg, ColIncidencia) = Empty
        T_Data.Datos(FilaReg, Col_H_Incid) = Empty
    Next FilaReg
    For FilaReg = 2 To TF_BD
        If T_Data.Datos(FilaReg, Colref) = T_Data.Datos(FilaReg - 1, Colref) Then   '- Existe "Dupla" coincidencia en las Referencias de Recibo
            If CantRepe = 1 Then                                        '- 1ª Repetición de esta Referencia
                DuplFind = DuplFind + 1                                 '- Cuento cuantos Reg. tienen repeticiones
                T_Data.Datos(FilaReg - 1, ColIncidencia) = "Repe01"     '- Marco Incidencia "Repe01" en el 1º Recibo de los 2 Repetidos
            End If                                                      '- Para esta repetición y las succesivas.. Para quedarme con el último de las repeticiones y saber cuantas repeticiones ha tenido este Recibo.
            T_Data.Datos(FilaReg - 1, ColIncidencia) = "Rp" & CantRepe  '- Al 1º de la Dupla le cambio la incidencia por una genérica "RP"
            CantRepe = CantRepe + 1                                     '- Acumulo contador de nº de repeticiones de esta Referencia
            T_Data.Datos(FilaReg, ColIncidencia) = "Repe" & CantRepe    '- Al 2º de la Dupla le pongo "Repe"+El némero de repetición por el que va de esta Referencia
            If MaxNumRepe < CantRepe Then MaxNumRepe = CantRepe         '- Registro cual es el máximo número de Repeticiones que hay de una misma Ref.
                                                                        'MaxNumRepe = WorksheetFunction.Max(MaxNumRepe, CantRepe)   '- Es más lento...
            Call RuT_Duplicates_Search_Mark_DIFF(T_Data, Col_Comparar, Col_H_Incid, FilaReg, CantRepe, DiccColor)
        Else                                                            '- Se ha acabado la serie de repeticiones, o no hay repetición
            CantRepe = 1                                                '- Inicializo contador de Repeticiones de una misma Ref.
        End If
    Next FilaReg
    T_Data.Modificada(ColIncidencia) = True
    T_Data.Modificada(Col_H_Incid) = True
    Call Rut_TablaRam_Volcar(T_Data, Lo_Data)                           '- Devuelvo a la hoja las 2 Col. de incidencias
    Call RuT_Duplicates_Aplicar_Colores(Lo_Data, DiccColor)             '- y pinto las celdas que pintaba el bucle celda a celda
    Erase T_Data.Datos                                                  '- Libero la RAM
        '- Visualizo el progreso --------
'''        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Find Ref. con Repeticiones: " & DuplFind & " reg.  y Máx nº Repeticiones, : " & MaxNumRepe & " veces).", LastTimeLap)
            With Lo_Data
                Call Rut_Lo_Sort(Lo_Data, ColIncidencia, xlAscending, True)
                For Cont = 2 To MaxNumRepe
                    .Range.AutoFilter Field:=ColIncidencia, Criteria1:="=Repe" & Cont & "*"    '- Todos los que se han quedados sin incidencias los borramos
                    Row_Find = Fnc_Lo_Contar_Visibles(Lo_Data, ColIncidencia)
                    If Row_Find > 0 Then
                        RepsFind = RepsFind + Row_Find
                        '- Visualizo el progreso --------
                        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Referencias Repes-" & Cont & " veces:", LastTimeLap, _
                             Format(Row_Find, " #,##0") & " reg. ", " de " & Format(Lo_Data.ListRows.Count, "#,##0") & " reg.")
                    End If
                    .AutoFilter.ShowAllData            ' Elimina los filtros
                Next Cont
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Find Ref. con Repeticiones: " & DuplFind & " reg.  y Máx nº Repeticiones, : " & MaxNumRepe & " veces).", TimeLapSub)
'''                        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(30, " ") & "Total Referencias con repeticiones:", 0, _
                             Format(RepsFind, " #,##0") & " reg. ", " de " & Format(Lo_Data.ListRows.Count, "#,##0") & " reg.")
            End With
        '- FIN, Visualizo el progreso --------

    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '- Gestionar Duplicados 2 Parte: Borra en Bdatos y deja en BD_Dpl los relevantes ----------------------------------------------------------
    '-      Borro de BDatos Rec. Repes NO Finalistas (el último de cada serie de repeticiones de referencia)
    '-      Copio los Repes Finalistas de BDatos a Lo_Duplic (Se añaden a los de otras ejecuciones)
    '-      Borro de Lo_Duplic, los Repes que ya estan repetidos "RpIdem" porque se juntan los repes de esta tanda con los de tandas anteriores
    '-      Borro de Lo_Duplic, los Repes que aun siendo repes, no han sufrido cambios "(en 0 Cols)"
    '- ----------------------------------------------------------------------------------------------------------------------------------------
'- Borrar Registros Repes en Lo_Data excepto el último --- OJO, PORQUE NO TENEMOS CRITERIO PARA SABER CUAL ES MEJOR QUEDARSE.
    With Lo_Data
        .Range.AutoFilter Field:=ColIncidencia, Criteria1:="=Rp*"
        Row_Find = Fnc_Lo_Contar_Visibles(Lo_Data, ColIncidencia)
        If Row_Find > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del en BDatos Rec. Repes NO finalistas", LastTimeLap, _
                 Format(Row_Find, " #,##0") & " reg. ", " quedan " & Format(Lo_Data.ListRows.Count, "#,##0") & " reg.")
        Else
             '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "¡ NO hay Referencias con Repeticiones !", LastTimeLap)
       End If
        .AutoFilter.ShowAllData            ' Elimina los filtros
        '- Copiar los Registros Repes Finalistas de Lo_Data (sólo los últimos de cada repetición) en Lo_Duplic
        Call Rut_Lo_Sort(Lo_Data, ColIncidencia, xlAscending, True)
        .Range.AutoFilter Field:=ColIncidencia, Criteria1:="=Repe*"
        Row_Find = Fnc_Lo_Contar_Visibles(Lo_Data, ColIncidencia)
        If Row_Find > 0 Then
            Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Data, Lo_Duplic, False)
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copiados Repes finalistas a Bd_Duplic", LastTimeLap, _
                 Format(Row_Find, " #,##0") & " reg. ", " tiene " & Format(Lo_Duplic.ListRows.Count, "#,##0") & " reg.")
        Else
             '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "¡ NO hay Referencias con Repeticiones !", LastTimeLap)
        End If
        .AutoFilter.ShowAllData            ' Elimina los filtros
    End With
'- Borrar Registros Repetidos entre ellos en Lo_Duplic
    With Lo_Duplic.DataBodyRange
        Call Rut_Lo_Sort(Lo_Duplic, Colref, xlAscending, True)
        .Columns(ColIncidencia).ClearContents   '- Se supone que está vacía...
        For FilaReg = 2 To Lo_Duplic.ListRows.Count
            If .Cells(FilaReg, Colref) = .Cells(FilaReg - 1, Colref) Then   '- Existe "Dupla" coincidencia en las Referencias de Recibo
                If .Cells(FilaReg, Col_H_Incid) = .Cells(FilaReg - 1, Col_H_Incid) Then   '- Es un Repetido que ya existe de antes e idéntico en incidencia
                    .Cells(FilaReg - 1, ColIncidencia) = "RpIdem"        '- Marco Incidencia "RpIdem" en el 1º Recibo
                End If
            End If
        Next FilaReg
    End With
    With Lo_Duplic
        Call Rut_Lo_Sort(Lo_Duplic, ColIncidencia, xlAscending, True)
        .Range.AutoFilter Field:=ColIncidencia, Criteria1:="=RpIdem"
        Row_Find = Fnc_Lo_Contar_Visibles(Lo_Duplic, ColIncidencia)
        If Row_Find > 0 Then .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Repes-X de repetidos RpIdem", LastTimeLap, _
                 Format(Row_Find, " #,##0") & " reg. ", " quedan " & Format(.ListRows.Count, "#,##0") & " reg.")
        Lo_Duplic.AutoFilter.ShowAllData            ' Elimina los filtros
    End With
    
    '- Borrar Registros en Lo_Duplic identificados como repetidos pero que no tienen ningún cambio en las columnas comparadas
    With Lo_Duplic
        Dim Celda As Range, RngCol_H_Incidencia As ListColumn:            Set RngCol_H_Incidencia = .ListColumns(Col_H_Incid)
        For Each Celda In RngCol_H_Incidencia.DataBodyRange    '- voy a borrar la referencias de duplicados a borrar por no tener cambios
            If InStr(Celda.Value, "_(en 0 Cols):") > 0 Then Celda.Value = "(en 0 Cols)"
        Next Celda
        Call Rut_Lo_Sort(Lo_Duplic, Col_H_Incid, xlAscending, True)
        .Range.AutoFilter Field:=Col_H_Incid, Criteria1:="=(en 0 Cols)"      '- Todos los que se han quedados sin incidencias los borramos
        Row_Find = Fnc_Lo_Contar_Visibles(Lo_Duplic, Col_H_Incid)
        If Row_Find > 0 Then .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Repes_Cambios(en 0 Cols)", LastTimeLap, _
                 Format(Row_Find, " #,##0") & " reg. ", " quedan " & Format(.ListRows.Count, "#,##0") & " reg.")
        .AutoFilter.ShowAllData            ' Elimina los filtros
    End With


        '- Visualizo el progreso --------
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(25, " ") & "Total Borrados: ", TimeLapSub, _
             Format(TF_BD - Lo_Data.ListRows.Count, " #,##0") & " reg. ", " de " & Format(Lo_Data.ListRows.Count, "#,##0") & " reg.")
        'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Encontrados: " & DuplFind & " reg. con repeticiones (Máx nº Repes: " & MaxNumRepe & " veces).", 0)
                
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD)

    If Lo_Data.AutoFilter.FilterMode Then Lo_Data.AutoFilter.ShowAllData            ' Elimina los filtros

Restaurar_Valores:
Debug.Print "<<< RuT_Duplicates_Search"
    Call Rut_Lo_WrkSht_Preparar(Sh_Data)
    Call Rut_Lo_Sort(Lo_Data, Colref, xlAscending, True)
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sh_Data)
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sh_Duplic)
    Lo_Data.ShowTotals = True
End Sub     ' RuT_Duplicates_Search
'-----------------------------------------------------------------------------------------------------------------------------------------

'=========================================================================================================================================
'- Marca los duplicados con "_Duplicati_1/2_" y añade en incidencia el valor del otro registro
'- Trabaja sobre la copia en RAM (T_Data); los colores no se pintan aquí: se anotan en DiccColor
'- (clave Fila*1000+Col, la última anotación de cada celda manda) y los pinta RuT_Duplicates_Aplicar_Colores.
Sub RuT_Duplicates_Search_Mark_DIFF(T_Data As T_TablaRam, _
                                    Col_Comparar() As Boolean, _
                                    ColHIncidencia As Integer, _
                                    Fila As Long, _
                                    CantRepe As Integer, _
                                    DiccColor As Object)
'Debug.Print ">>> RuT_Duplicates_Search_Mark_DIFF"
    Dim Cont_Col    As Integer
    Dim CAnt_Col    As Integer
    Dim F_Ant       As Long:    F_Ant = Fila - 1
    Dim Incidencia  As String

    With T_Data
        CAnt_Col = 0
        For Cont_Col = 2 To .NumCols
            If Col_Comparar(Cont_Col) Then
                If .Datos(F_Ant, Cont_Col) <> .Datos(Fila, Cont_Col) Then
                    Incidencia = " En Col. nº" & Cont_Col & " (" & .Titulos(Cont_Col) & _
                                 ")<>[" & .Datos(Fila, Cont_Col) & "] _-_-_ "
                    .Datos(F_Ant, ColHIncidencia) = .Datos(F_Ant, ColHIncidencia) & Incidencia
                    Incidencia = " En Col. nº" & Cont_Col & " (" & .Titulos(Cont_Col) & _
                                 ")<>[" & .Datos(F_Ant, Cont_Col) & "] _-_-_ "
                    .Datos(Fila, ColHIncidencia) = .Datos(Fila, ColHIncidencia) & Incidencia
                    CAnt_Col = CAnt_Col + 1
                    DiccColor(F_Ant * 1000& + Cont_Col) = 34
                    DiccColor(Fila * 1000& + Cont_Col) = 35
                End If

            End If
        Next
        '- Añadir la marca de duplicado ----
        Incidencia = "_Duplicati_" & CantRepe - 1 & "_(en " & CAnt_Col & " Cols):"
        If CantRepe = 2 Then .Datos(F_Ant, ColHIncidencia) = Incidencia & .Datos(F_Ant, ColHIncidencia)
        DiccColor(F_Ant * 1000& + ColHIncidencia) = 34
        DiccColor(F_Ant * 1000& + BD_Ref) = 34

        Incidencia = "_Duplicati_" & CantRepe & "_(en " & CAnt_Col & " Cols):"
        .Datos(Fila, ColHIncidencia) = Incidencia & .Datos(Fila, ColHIncidencia)
        DiccColor(Fila * 1000& + ColHIncidencia) = 35
        DiccColor(Fila * 1000& + BD_Ref) = 35
    End With

End Sub     ' RuT_Duplicates_Search_Mark_DIFF
'-----------------------------------------------------------------------------------------------------------------------------------------

'=========================================================================================================================================
'- Pinta las celdas anotadas en DiccColor (clave Fila*1000+Col de la tabla, valor ColorIndex). En vez de una llamada a Excel por celda,
'- junta las celdas del mismo color en direcciones "A5,C9,..." de hasta 250 caracteres y pinta cada grupo de una vez.
Private Sub RuT_Duplicates_Aplicar_Colores(Lo_Data As ListObject, DiccColor As Object)
    Dim Ws          As Worksheet:   Set Ws = Lo_Data.Parent
    Dim Fila1       As Long:        Fila1 = Lo_Data.DataBodyRange.Row
    Dim Col1        As Long:        Col1 = Lo_Data.DataBodyRange.Column
    Dim DiccDirecc  As Object:      Set DiccDirecc = CreateObject("Scripting.Dictionary")  '- ColorIndex -> direcciones pendientes de pintar
    Dim Letra()     As String
    Dim Clave       As Variant
    Dim ColorIdx    As Variant
    Dim Celda       As String
    Dim C           As Long

    If DiccColor.Count = 0 Then Exit Sub
    ReDim Letra(1 To Lo_Data.ListColumns.Count)                         '- Letra de columna de la hoja, de cada Col. de la tabla
    For C = 1 To Lo_Data.ListColumns.Count
        Letra(C) = Split(Ws.Cells(1, Col1 + C - 1).Address(True, False), "$")(0)
    Next C
    For Each Clave In DiccColor.Keys
        ColorIdx = DiccColor(Clave)
        Celda = Letra(Clave Mod 1000) & (Fila1 + Clave \ 1000 - 1)
        If Len(DiccDirecc(ColorIdx)) + Len(Celda) + 1 > 250 Then
            Ws.Range(DiccDirecc(ColorIdx)).Interior.ColorIndex = ColorIdx
            DiccDirecc(ColorIdx) = ""
        End If
        If DiccDirecc(ColorIdx) = "" Then
            DiccDirecc(ColorIdx) = Celda
        Else
            DiccDirecc(ColorIdx) = DiccDirecc(ColorIdx) & "," & Celda
        End If
    Next Clave
    For Each ColorIdx In DiccDirecc.Keys
        If DiccDirecc(ColorIdx) <> "" Then Ws.Range(DiccDirecc(ColorIdx)).Interior.ColorIndex = ColorIdx
    Next ColorIdx
End Sub     ' RuT_Duplicates_Aplicar_Colores
'-----------------------------------------------------------------------------------------------------------------------------------------




