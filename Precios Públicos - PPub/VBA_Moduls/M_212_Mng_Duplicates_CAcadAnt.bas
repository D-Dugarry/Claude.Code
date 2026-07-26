Attribute VB_Name = "M_212_Mng_Duplicates_CAcadAnt"
'Rev.: 2026-01-15
'M_212_Manage_Duplicates_CAcad
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Gestionar Duplicados -------------------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Duplicates_Search_and_Del(Lo_Data As ListObject, _
                                  Colref As Integer)
                                   
Debug.Print "RuT_Duplicates_Search_and_Del"
    Dim DuplFind            As Integer:     DuplFind = 0
    Dim CantRepe            As Integer:     CantRepe = 1
    Dim MaxNumRepe          As Integer:     MaxNumRepe = 0
    Dim AntCptoDto          As String:      AntCptoDto = ""
    Dim Fila                As Long
    Dim FilaReg             As Long
    Dim TF_BD               As Long:        TF_BD = Lo_Data.ListRows.Count
    Dim Sh_Data             As Worksheet:   Set Sh_Data = Lo_Data.Parent

        
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_Lo_Sort(Lo_Data, Colref, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------------------
    '- Gestionar Duplicados 1ª Parte: Los Identifica y Marca las Diferencias -------------------------------------------------------------------------
    '- ----------------------------------------------------------------------------------------------------------------------------------------
    With Lo_Data.DataBodyRange
        .Columns(BD_H_Incidencias).ClearContents   '- Se supone que está vacía...
        For FilaReg = 2 To TF_BD
            If .Cells(FilaReg, Colref) = .Cells(FilaReg - 1, Colref) Then   '- Existe "Dupla" coincidencia en las Referencias de Recibo
                If CantRepe = 1 Then                                        '- 1ª Repetición de esta Referencia
                    DuplFind = DuplFind + 1                                 '- Cuento cuantos Reg. tienen repeticiones
                    .Cells(FilaReg - 1, BD_H_Incidencias) = "Repe01"        '- Marco Incidencia "Repe01" en el 1º Recibo de los 2 Repetidos
                End If                                                      '- Para esta repetición y las succesivas.. Para quedarme con el último de las repeticiones y saber cuantas repeticiones ha tenido este Recibo.
                .Cells(FilaReg - 1, BD_H_Incidencias) = "Rp" & CantRepe     '- Al 1º de la Dupla le cambio la incidencia por una genérica "RP"
                CantRepe = CantRepe + 1                                     '- Acumulo contador de nº de repeticiones de esta Referencia
                .Cells(FilaReg, BD_H_Incidencias) = "Repe" & CantRepe       '- Al 2º de la Dupla le pongo "Repe"+El némero de repetición por el que va de esta Referencia
                If MaxNumRepe < CantRepe Then MaxNumRepe = CantRepe         '- Registro cual es el máximo número de Repeticiones que hay de una misma Ref.
                                                                                ' OJO, MaxNumRepe = WorksheetFunction.Max(MaxNumRepe, CantRepe)   '- Es más lento...
            Else                                                            '- Se ha acabado la serie de repeticiones, o no hay repetición
                CantRepe = 1                                                '- Inicializo contador de Repeticiones de una misma Ref.
            End If
        Next FilaReg
    End With    ' Lo_Data.DataBodyRange
        
    '- Borrar Registros Duplicados, EN ESTE CASO NO QUEDAMOS CON EL ÚLTIMO, así queda marcado cuantas repeticiones tenemos de esa Ref.  ---------------------------------
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Data, BD_H_Incidencias, "=Rp*")
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Find Ref. con Repeticiones: " & DuplFind & " y Máx nº Repeticiones, " & MaxNumRepe & " veces." & _
                vbLf & String(26, " ") & Left("Borrados " & String(35, "_"), 35) & Right(String(25, "_") & TF_BD - Lo_Data.ListRows.Count & " reg., Quedan " & Format(Lo_Data.ListRows.Count, "#,##0") & " reg.", 30), LastTimeLap)

    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sh_Data)

End Sub     ' RuT_Duplicates_Search_and_Del








