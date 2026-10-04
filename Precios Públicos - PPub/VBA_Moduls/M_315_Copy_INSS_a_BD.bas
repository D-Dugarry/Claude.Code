Attribute VB_Name = "M_315_Copy_INSS_a_BD"
' Last Rev. 2026-10-04 18:55
'Rev.: 2026-01-22
Option Explicit


            Sub Rut_Copy_ImpINSS_en_BDatos_ByHand()
                
                Call Rut_Copy_ImpAdm_CAcadAnt_a_BDatos
                MsgBox "FIN"
'                Application.Speech.Speak "Proceso completado."
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- M_315_Copy_INSS_a_BD, Trasladar el importe INSS a los recibos de BDatos ----------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Copy_ImpINSS_en_BDatos()
Debug.Print ">>> Rut_Copy_ImpINSS_en_BDatos"

    Dim CantImpINSS      As Long:        CantImpINSS = 0
    Dim ImpTINSS         As Currency:      ImpTINSS = 0
    Dim F_BD            As Long
    Dim F_BDINSS        As Long
    Dim Found           As Long:        Found = 0
    Dim NotFound        As Long:        NotFound = 0
    Dim AnoCont         As String:      AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    Dim TxT_Progreso    As String
    Dim TxT_ProgIni     As String
    Dim rowfind         As Variant
    
    Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_Bd_INSS      As ListObject:      Set Lo_Bd_INSS = Sht__BD_INSS.ListObjects(1)
    Dim TF_BD           As Long:            TF_BD = Lo_BD.ListRows.Count
    Dim TF_BbINSS       As Long:            TF_BbINSS = Lo_Bd_INSS.ListRows.Count
    
    Sht__BD.Visible = xlSheetVisible:           Sht__BD.Activate:         Sht__BD.Unprotect:          Lo_BD.ShowTotals = False
    Sht__BD_INSS.Visible = xlSheetVisible:      Sht__BD_INSS.Activate:    Sht__BD_INSS.Unprotect
    '- (2026-10-04) Sin quitar la fila de totales de Tb_INSS: llamada desde M_110, que llega con los totales puestos, el
    '- ShowTotals = False daba el error -2147417848 (80010108), también con el código antiguo. Desde M_310 ya llega sin
    '- ellos. El cruce solo usa las filas de datos, así que la fila de totales no estorba.
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    Call Rut_Lo_Filtros_Quitar(Lo_Bd_INSS)
    
    '- Visualizo el progreso
    TxT_ProgIni = ActivForm.Controls("TBx_Informe")
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Procedimiento: Incorporar Imp. INSS, Cursos " & C_Acad_Ant & " y " & C_Acad_Pos, 0, , , , , , 4)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")

    '- --------------------------------------------------------------------------------------------------------
    '- ClearContents en BDatos los Imp.INSS de los recibos de C_Acad_Ant y C_Acad_Pos --------------------------------
    With Lo_BD
        Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)                           '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & C_Acad_Ant, Operator:=xlOr, Criteria2:="=" & C_Acad_Pos   '- Filtro los Recibos del Curso-Acad-Ant/Pos
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Rec_Imp_INSS).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        .AutoFilter.ShowAllData            ' Elimina los filtros
    End With
    
   '- -----------------------------------------------------------------------------------------------------------------------
   '- - Actualizo BDatos con Lo_Bd_INSS ------------------------------------------------------------------------------------------------
   '- -----------------------------------------------------------------------------------------------------------------------
    Lo_BD.DataBodyRange.Columns(BD_Rec_Imp_INSS).ClearContents
'    Lo_Bd_INSS.DataBodyRange.Columns(BD_Incidencias).ClearContents
    Call Rut_Lo_Sort(Lo_Bd_INSS, LS06_Ref, xlAscending, True)   '- Las dos tablas por Ref, para cruzarlas en un solo recorrido
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)
    '- El cruce se hace en RAM (Rut_Lo_TablaRam), con el mismo recorrido de siempre: celda a celda en la hoja era lento.
    '- BD se carga con .Value2 y la tabla INSS con .Value, que es lo que se copiaba antes con "Celda = Celda". Rec_Imp_INSS no
    '- se carga: se acaba de vaciar, así que en RAM empieza vacía y se devuelve entera.
    Dim T_BD            As T_TablaRam
    Dim T_INSS          As T_TablaRam
    Call Rut_TablaRam_Cargar(T_BD, Lo_BD, Array(BD_Ref), True)
    Call Rut_TablaRam_Cargar(T_INSS, Lo_Bd_INSS, Array(LS06_Ref, LS06_Concept_Imp, LS06_RecFound))
    F_BD = 1
    For F_BDINSS = 1 To TF_BbINSS
        Select Case T_BD.Datos(F_BD, BD_Ref)
            Case Is < T_INSS.Datos(F_BDINSS, LS06_Ref)     '- Ref_BD  NO-EXISTE-EN  Sht__BD_Adm
                If F_BD < TF_BD Then
                    F_BD = F_BD + 1
                    F_BDINSS = F_BDINSS - 1
                Else
                    '- BD agotada: este y todos los INSS restantes quedan sin encontrar --------
                    NotFound = NotFound + 1
                    T_INSS.Datos(F_BDINSS, LS06_RecFound) = "Not Found en BD. (BD agotada)"
                End If
            Case Is = T_INSS.Datos(F_BDINSS, LS06_Ref)     '- Ref_BD  SÍ-EXISTE-EN  Sht__BD_Adm
                T_BD.Datos(F_BD, BD_Rec_Imp_INSS) = T_INSS.Datos(F_BDINSS, LS06_Concept_Imp)
                Found = Found + 1
                T_INSS.Datos(F_BDINSS, LS06_RecFound) = "Found en BD."
                If F_BD < TF_BD Then F_BD = F_BD + 1
            Case Is > T_INSS.Datos(F_BDINSS, LS06_Ref)     '- Ref_BD_Adm  NO-EXISTE-EN  Sht__BD
                NotFound = NotFound + 1
                T_INSS.Datos(F_BDINSS, LS06_RecFound) = "Not Found en BD."
        End Select
        If (F_BDINSS Mod 4000 = 0 And F_BDINSS <> 0) Or F_BD Mod 4000 = 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporados datos de: " & Format(Found, "#,##0") & _
                        " reg. de " & Format(F_BDINSS, "#,##0") & "/" & Format(TF_BbINSS, "#,##0") & "reg.,  entre " & _
                        Format(F_BD, "#,##0") & "/" & Format(TF_BD, "#,##0") & "reg.", 0, , , TxT_Progreso, True, , 2)
        End If
    Next F_BDINSS
    T_BD.Modificada(BD_Rec_Imp_INSS) = True
    T_INSS.Modificada(LS06_RecFound) = True
    Call Rut_TablaRam_Volcar(T_BD, Lo_BD)                               '- Devuelvo a la hoja la Col. Rec_Imp_INSS
    Call Rut_TablaRam_Volcar(T_INSS, Lo_Bd_INSS)                        '- y RecFound
    Erase T_BD.Datos                                                    '- Libero la RAM
    Erase T_INSS.Datos
    
        
        '- Sumatorios ---------
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        With Lo_BD.DataBodyRange
                ImpTINSS = Application.Sum(.Columns(BD_Rec_Imp_INSS))
                CantImpINSS = Application.Count(.Columns(BD_Rec_Imp_INSS))
        End With
        '- Visualizo el progreso ----------------------------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Procedimiento: Incorporar Imp.INSS. de los C_Acad " & C_Acad_Ant & " y " & C_Acad_Pos, 0, , , TxT_ProgIni, , , 2)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Recibos NO encontrados: ", TimeLapSub, _
                            Format(NotFound, "#,##0") & " reg.", " de " & Format(TF_BbINSS, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "Incorporados datos de: ", 0, _
                            Format(Found, "#,##0") & " reg.", " de " & Format(TF_BbINSS, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "Total Importe Seguro Obl. INSS", 0, _
                            Format(ImpTINSS, "#,##0.00€     "), " de " & Format(CantImpINSS, "#,##0") & " reg.")
Lo_BD.ShowTotals = True
Sht__BD_INSS.Range("a1").Select

Restablecer_Valores:
    Application.Speech.Speak "Proceso completado."
End Sub     ' Rut_Copy_ImpINSS_en_BDatos     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

