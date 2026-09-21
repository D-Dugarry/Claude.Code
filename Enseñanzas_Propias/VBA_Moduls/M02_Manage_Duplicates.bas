Attribute VB_Name = "M02_Manage_Duplicates"
' Last Rev. 2026-09-21 12:12
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M02_Manage_Duplicates - Deteccion y gestion de referencias duplicadas
' =================================================================================================
'
' PROPOSITO
'  LSGES04 puede traer varias filas con la misma referencia de recibo (BD_Ref).
'  Este modulo las detecta, anota EN QUE columnas difieren, se queda con la
'  ultima de cada serie y archiva las relevantes en Prog_BD_Dupl para revision
'  manual. Lo llama M01 antes de clasificar.
'
' INDICE DE RUTINAS Y FUNCIONES
'  RuT_Duplicates_Search(Lo_Data, Lo_DefCol, Lo_Duplic, Colref,
'                        ColIncidencia, Col_H_Incid) ... Rutina principal.
'  RuT_Duplicates_Search_Mark_DIFF(Lo_Data, Lo_DefCol, ColHIncidencia,
'                        fila, CantRepe) ......... Auxiliar: compara dos filas
'                        consecutivas columna a columna y marca diferencias.
'
' TRAMOS DE PROGRAMACION
'  PARTE 1 - Identificar y marcar (RuT_Duplicates_Search)
'    Ordena por Colref y recorre la tabla comparando cada fila con la anterior.
'    Al encontrar referencias iguales:
'      - al PRIMERO de la serie lo marca 'Rp<n>'  (candidato a borrar),
'      - al ULTIMO lo marca  'Repe<n>'            (finalista, se conserva),
'      - lleva la cuenta del maximo numero de repeticiones de una misma ref.,
'      - llama a _Mark_DIFF para anotar las diferencias reales.
'    Luego informa, por numero de repeticiones, cuantas referencias hay.
'
'  PARTE 2 - Depurar (misma rutina, segundo bloque)
'    a) Borra de Lo_Data los 'Rp*' (repes NO finalistas). El comentario del
'       autor avisa: no hay criterio objetivo para elegir cual conservar, se
'       conserva el ultimo.
'    b) Copia los 'Repe*' (finalistas) a Lo_Duplic, ACUMULANDO con los de
'       ejecuciones anteriores.
'    c) En Lo_Duplic marca 'RpIdem' los que ya estaban de antes con identica
'       incidencia, y los borra (evita que se acumulen tandas repetidas).
'    d) Borra de Lo_Duplic los que, aun siendo duplicados, no cambian en
'       ninguna columna comparada: su historico contiene '_(en 0 Cols):'.
'
'  RuT_Duplicates_Search_Mark_DIFF
'    Recorre las columnas desde la 2 y compara solo las que tienen activo el
'    flag DefC_Compare en la tabla DefCol. Por cada diferencia escribe en el
'    historico de AMBAS filas el numero de columna, su titulo y el valor de la
'    otra fila, y colorea la celda (ColorIndex 34 la anterior, 35 la actual).
'    Al final antepone la marca '_Duplicati_<n>_(en <x> Cols):', que es
'    justamente lo que lee el paso (d) para descartar duplicados sin cambios.
'
' NOTAS
'  Dos columnas distintas de incidencia: ColIncidencia es la marca de trabajo
'  de esta pasada (se limpia al empezar) y Col_H_Incid es el HISTORICO que se
'  va acumulando y que sobrevive entre ejecuciones.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2026-01-11
'M_112_Manage_Duplicates
Option Explicit

'===================================================================================================
'- Gestionar Duplicados ----------------------------------------------------------------------------
'===================================================================================================
Sub RuT_Duplicates_Search(Lo_Data As ListObject, _
                          Lo_DefCol As ListObject, _
                          Lo_Duplic As ListObject, _
                          Colref As Integer, _
                          ColIncidencia As Integer, _
                          Col_H_Incid As Integer)
                          
Debug.Print ">>> RuT_Duplicates_Search"
    Dim TimeLapSub          As Single:      TimeLapSub = Timer
    Dim LastTimeLap         As Single:      LastTimeLap = Timer
    Dim DuplFind            As Integer:     DuplFind = 0    '- nº de Rec. que tienen repeticiones
    Dim CantRepe            As Integer:     CantRepe = 1
    Dim MaxNumRepe          As Integer:     MaxNumRepe = 0
    Dim RepsFind            As Integer:     RepsFind = 0
    Dim Cont                As Integer
    Dim FilaReg             As Long:        FilaReg = 1
    Dim TxtMsg1  As String, TxtMsg2  As String, TxtMsg3  As String
    Dim TF_BD               As Long:        TF_BD = Lo_Data.ListRows.Count
    Dim rowfind            As Variant
    Dim Sh_Data             As Worksheet:   Set Sh_Data = Lo_Data.Parent
    Dim Sh_Duplic           As Worksheet:   Set Sh_Duplic = Lo_Duplic.Parent
    
    Call Rut_Lo_WrkSht_Preparar(Sh_Data)
    Call Rut_Lo_WrkSht_Preparar(Sh_Duplic)
    Lo_Data.ShowTotals = False
    Lo_Duplic.ShowTotals = False
    
    Application.DisplayAlerts = False

    '- ---------------------------------------------------------------------------------------------
    '- Gestionar Duplicados 1ª Parte: Los Identifica y Marca las Diferencias -----------------------
    '- ---------------------------------------------------------------------------------------------
    Call Rut_Lo_Sort(Lo_Data, Colref, xlAscending, True)    '- Ordenar primero accelera un montón el borrado
    With Lo_Data.DataBodyRange
        .Columns(ColIncidencia).ClearContents   '- Se supone que está vacía...
        .Columns(Col_H_Incid).ClearContents   '- Se supone que está vacía...
        For FilaReg = 2 To TF_BD
            If .Cells(FilaReg, Colref) = .Cells(FilaReg - 1, Colref) Then   '- Existe "Dupla" coincidencia en las Referencias de Recibo
                If CantRepe = 1 Then                                        '- 1ª Repetición de esta Referencia
                    DuplFind = DuplFind + 1                                 '- Cuento cuantos Reg. tienen repeticiones
                    .Cells(FilaReg - 1, ColIncidencia) = "Repe01"        '- Marco Incidencia "Repe01" en el 1º Recibo de los 2 Repetidos
                End If                                                      '- Para esta repetición y las succesivas.. Para quedarme con el último de las repeticiones y saber cuantas repeticiones ha tenido este Recibo.
                .Cells(FilaReg - 1, ColIncidencia) = "Rp" & CantRepe     '- Al 1º de la Dupla le cambio la incidencia por una genérica "RP"
                CantRepe = CantRepe + 1                                     '- Acumulo contador de nº de repeticiones de esta Referencia
                .Cells(FilaReg, ColIncidencia) = "Repe" & CantRepe       '- Al 2º de la Dupla le pongo "Repe"+El némero de repetición por el que va de esta Referencia
                If MaxNumRepe < CantRepe Then MaxNumRepe = CantRepe         '- Registro cual es el máximo número de Repeticiones que hay de una misma Ref.
                                                                            'MaxNumRepe = WorksheetFunction.Max(MaxNumRepe, CantRepe)   '- Es más lento...
                Call RuT_Duplicates_Search_Mark_DIFF(Lo_Data, Lo_DefCol, Col_H_Incid, FilaReg, CantRepe)
            Else                                                            '- Se ha acabado la serie de repeticiones, o no hay repetición
                CantRepe = 1                                                '- Inicializo contador de Repeticiones de una misma Ref.
            End If
        Next FilaReg
    End With
    '- Visualizo el progreso --------
        With Lo_Data
            Call Rut_Lo_Sort(Lo_Data, ColIncidencia, xlAscending, True)
            For Cont = 2 To MaxNumRepe
                .Range.AutoFilter Field:=ColIncidencia, Criteria1:="=Repe" & Cont & "*"    '- Todos los que se han quedados sin incidencias los borramos
                rowfind = .Range.Columns(ColIncidencia).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                If rowfind > 0 Then
                    RepsFind = RepsFind + rowfind
                    '- Visualizo el progreso --------
                    TxtMsg1 = "Referencias Repes-" & Cont & " veces."
                    TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                    TxtMsg3 = " de " & Format(.ListRows.Count, "#,##0") & "reg."
                Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
                End If
                .AutoFilter.ShowAllData            ' Elimina los filtros
            Next Cont
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "Find Ref. con Repeticiones: " & DuplFind & _
                                           " ref.  y Máx nº Repeticiones: " & MaxNumRepe
        End With
    '- FIN, Visualizo el progreso --------

    '- ---------------------------------------------------------------------------------------------
    '- Gestionar Duplicados 2 Parte: Borra en Bdatos y deja en BD_Dpl los relevantes ---------------
    '-      Borro de BDatos Rec. Repes NO Finalistas (el último de cada serie de repeticiones de referencia)
    '-      Copio los Repes Finalistas de BDatos a Lo_Duplic (Se añaden a los de otras ejecuciones)
    '-      Borro de Lo_Duplic, los Repes que ya estan repetidos "RpIdem" porque se juntan los repes de esta tanda con los de tandas anteriores
    '-      Borro de Lo_Duplic, los Repes que aun siendo repes, no han sufrido cambios "(en 0 Cols)"
    '- ---------------------------------------------------------------------------------------------
    With Lo_Data
        '- Borrar Registros Repes en Lo_Data excepto el último --- OJO, PORQUE NO TENEMOS CRITERIO PARA SABER CUAL ES MEJOR QUEDARSE.
        .Range.AutoFilter Field:=ColIncidencia, Criteria1:="=Rp*"
        rowfind = .Range.Columns(ColIncidencia).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            '- Visualizo el progreso --------
                        TxtMsg1 = "Del en Lo_Ges04 Rec. Repes NO finalistas: "
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
        Else
             '- Visualizo el progreso --------
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡ NO hay Referencias con Repeticiones !"
       End If
        .AutoFilter.ShowAllData            ' Elimina los filtros
        '- Copiar los Registros Repes Finalistas de Lo_Data (sólo los últimos de cada repetición) en Lo_Duplic
        Call Rut_Lo_Sort(Lo_Data, ColIncidencia, xlAscending, True)
        .Range.AutoFilter Field:=ColIncidencia, Criteria1:="=Repe*"
        rowfind = .Range.Columns(ColIncidencia).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        If rowfind > 0 Then
            Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Data, Lo_Duplic, False)
            '- Visualizo el progreso --------
                        TxtMsg1 = "Copiados Repes finalistas a Bd_Duplic:"
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " tiene " & Format(Lo_Duplic.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
        Else
             '- Visualizo el progreso --------
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡ NO hay Referencias con Repeticiones !"
        End If
        .AutoFilter.ShowAllData            ' Elimina los filtros
    End With
    '- Borrar Registros Repetidos entre ellos en Lo_Duplic
    With Lo_Duplic.DataBodyRange
        Call Rut_Lo_Sort(Lo_Duplic, Colref, xlAscending, True)
        .Columns(BD_Incidencias).ClearContents   '- Se supone que está vacía...
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
        rowfind = .Range.Columns(ColIncidencia).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        If rowfind > 0 Then .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            '- Visualizo el progreso --------
                        TxtMsg1 = "Del Recibos Repes-X de repetidos RpIdem:"
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
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
        rowfind = .Range.Columns(Col_H_Incid).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        If rowfind > 0 Then .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
            '- Visualizo el progreso --------
                        TxtMsg1 = "Del Recibos Repes_Cambios(en 0 Cols):"
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
        .AutoFilter.ShowAllData            ' Elimina los filtros
    End With

Restaurar_Valores:
    Call Rut_Lo_WrkSht_Preparar(Sh_Data)
    Call Rut_Lo_Sort(Lo_Data, Colref, xlAscending, True)
    Lo_Data.ShowTotals = True
    Application.DisplayAlerts = True
Debug.Print "<<< RuT_Duplicates_Search"
End Sub     ' RuT_Duplicates_Search
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Marca los duplicados con "_Duplicati_1/2_" y añade en incidencia el valor del otro registro
Sub RuT_Duplicates_Search_Mark_DIFF(Lo_Data As ListObject, _
                                    Lo_DefCol As ListObject, _
                                    ColHIncidencia As Integer, _
                                    fila As Long, _
                                    CantRepe As Integer)
Debug.Print ">>> RuT_Duplicates_Search_Mark_DIFF"
    Dim Cont_Col    As Integer
    Dim CAnt_Col    As Integer
    Dim colot
    Dim F_Ant       As Long:    F_Ant = fila - 1
    Dim Incidencia  As String
    Dim SW_Compare  As Boolean
    Dim Sh_Data     As Worksheet:   Set Sh_Data = Lo_Data.Parent
    
    With Lo_Data.DataBodyRange
        CAnt_Col = 0
        For Cont_Col = 2 To Lo_Data.ListColumns.Count
            SW_Compare = Lo_DefCol.ListColumns(DefC_Compare).DataBodyRange(Cont_Col)
            If SW_Compare Then
                If .Cells(F_Ant, Cont_Col) <> .Cells(fila, Cont_Col) Then
                    Incidencia = " En Col. nº" & Cont_Col & " (" & .Cells(0, Cont_Col) & _
                                 ")<>[" & .Cells(fila, Cont_Col) & "] _-_-_ "
                    .Cells(F_Ant, ColHIncidencia) = .Cells(F_Ant, ColHIncidencia) & Incidencia
                    Incidencia = " En Col. nº" & Cont_Col & " (" & .Cells(0, Cont_Col) & _
                                 ")<>[" & .Cells(F_Ant, Cont_Col) & "] _-_-_ "
                    .Cells(fila, ColHIncidencia) = .Cells(fila, ColHIncidencia) & Incidencia
                    CAnt_Col = CAnt_Col + 1
                    .Cells(F_Ant, Cont_Col).Interior.ColorIndex = 34
                    .Cells(fila, Cont_Col).Interior.ColorIndex = 35
                End If
        
            End If
        Next
        '- Añadir la marca de duplicado ----
        Incidencia = "_Duplicati_" & CantRepe - 1 & "_(en " & CAnt_Col & " Cols):"
        If CantRepe = 2 Then .Cells(F_Ant, ColHIncidencia) = Incidencia & .Cells(F_Ant, ColHIncidencia)
        .Cells(F_Ant, ColHIncidencia).Interior.ColorIndex = 34
        .Cells(F_Ant, BD_Ref).Interior.ColorIndex = 34
        
        Incidencia = "_Duplicati_" & CantRepe & "_(en " & CAnt_Col & " Cols):"
        .Cells(fila, ColHIncidencia) = Incidencia & .Cells(fila, ColHIncidencia)
        .Cells(fila, ColHIncidencia).Interior.ColorIndex = 35
        .Cells(fila, BD_Ref).Interior.ColorIndex = 35
    End With
    
End Sub     ' RuT_Duplicates_Search_Mark_DIFF
'---------------------------------------------------------------------------------------------------




