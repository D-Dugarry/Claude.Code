Attribute VB_Name = "M_315_Copy_INSS_a_BD"
' Last Rev. 2026-10-08 13:52
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
    Dim DFound          As Object:      Set DFound = CreateObject("Scripting.Dictionary")   '- Por Curso Acad: Recibos encontrados
    Dim DNot            As Object:      Set DNot = CreateObject("Scripting.Dictionary")     '- Por Curso Acad: Recibos NO encontrados
    Dim DImp            As Object:      Set DImp = CreateObject("Scripting.Dictionary")     '- Por Curso Acad: Importe INSS incorporado
    Dim K_CAcad         As String
    
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
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Procedimiento: Incorporar Imp. INSS, Cursos " & C_Acad_Ant & " y " & C_Acad_Pos, 0, , , , , , 4)
    TxT_Progreso = ActivForm.Controls("TBx_Informe")

    '- --------------------------------------------------------------------------------------------------------
    '- ClearContents en BDatos los Imp.INSS de los recibos de C_Acad_Ant y C_Acad_Pos --------------------------------
    With Lo_BD
        Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)                           '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & C_Acad_Ant, Operator:=xlOr, Criteria2:="=" & C_Acad_Pos   '- Filtro los Recibos del Curso-Acad-Ant/Pos
        rowfind = Fnc_Lo_Contar_Visibles(Lo_BD, BD_Ref)    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(BD_Rec_Imp_INSS).SpecialCells(xlCellTypeVisible).ClearContents
        End If
        .AutoFilter.ShowAllData            ' Elimina los filtros
    End With
    
   '- -----------------------------------------------------------------------------------------------------------------------
   '- - Actualizo BDatos con Lo_Bd_INSS ------------------------------------------------------------------------------------------------
   '- -----------------------------------------------------------------------------------------------------------------------
    '- (2026-10-08) Ya no se vacia la columna Rec_Imp_INSS entera: solo la de los recibos de C_Acad_Ant/Pos (arriba), para conservar la de otros cursos si BD_INSS se queda con uno solo.
'    Lo_Bd_INSS.DataBodyRange.Columns(BD_Incidencias).ClearContents
    Call Rut_Lo_Sort(Lo_Bd_INSS, LS06_Ref, xlAscending, True)   '- Las dos tablas por Ref, para cruzarlas en un solo recorrido
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)
    '- El cruce se hace en RAM (Rut_Lo_TablaRam), con el mismo recorrido de siempre: celda a celda en la hoja era lento.
    '- BD se carga con .Value2 y la tabla INSS con .Value, que es lo que se copiaba antes con "Celda = Celda". Rec_Imp_INSS SI se carga (ya no se vacia entera,
    '- para conservar la de otros cursos) y se devuelve entera, con sus valores.
    Dim T_BD            As T_TablaRam
    Dim T_INSS          As T_TablaRam
    Call Rut_TablaRam_Cargar(T_BD, Lo_BD, Array(BD_Ref, BD_Rec_Imp_INSS), True)
    Call Rut_TablaRam_Cargar(T_INSS, Lo_Bd_INSS, Array(LS06_Ref, LS06_C_Acad, LS06_Concept_Imp, LS06_RecFound))
    F_BD = 1
    For F_BDINSS = 1 To TF_BbINSS
        K_CAcad = CStr(T_INSS.Datos(F_BDINSS, LS06_C_Acad))
        Select Case T_BD.Datos(F_BD, BD_Ref)
            Case Is < T_INSS.Datos(F_BDINSS, LS06_Ref)     '- Ref_BD  NO-EXISTE-EN  Sht__BD_Adm
                If F_BD < TF_BD Then
                    F_BD = F_BD + 1
                    F_BDINSS = F_BDINSS - 1
                Else
                    '- BD agotada: este y todos los INSS restantes quedan sin encontrar --------
                    NotFound = NotFound + 1
                    T_INSS.Datos(F_BDINSS, LS06_RecFound) = "Not Found en BD. (BD agotada)"
                    Call Rut_INSS_Acum(DNot, K_CAcad, 1)
                End If
            Case Is = T_INSS.Datos(F_BDINSS, LS06_Ref)     '- Ref_BD  SÍ-EXISTE-EN  Sht__BD_Adm
                T_BD.Datos(F_BD, BD_Rec_Imp_INSS) = T_INSS.Datos(F_BDINSS, LS06_Concept_Imp)
                Found = Found + 1
                T_INSS.Datos(F_BDINSS, LS06_RecFound) = "Found en BD."
                Call Rut_INSS_Acum(DFound, K_CAcad, 1)
                Call Rut_INSS_Acum(DImp, K_CAcad, Fnc_INSS_Importe(T_INSS.Datos(F_BDINSS, LS06_Concept_Imp)))
                If F_BD < TF_BD Then F_BD = F_BD + 1
            Case Is > T_INSS.Datos(F_BDINSS, LS06_Ref)     '- Ref_BD_Adm  NO-EXISTE-EN  Sht__BD
                NotFound = NotFound + 1
                T_INSS.Datos(F_BDINSS, LS06_RecFound) = "Not Found en BD."
                Call Rut_INSS_Acum(DNot, K_CAcad, 1)
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
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Procedimiento: Incorporar Imp.INSS. de los C_Acad " & C_Acad_Ant & " y " & C_Acad_Pos, 0, , , TxT_ProgIni, , , 2)
        '- El resultado, por Curso Académico (Ant y Pos primero; cualquier otro curso que hubiera en BD_INSS, después) ---------------
        Dim DCursos     As Object:      Set DCursos = CreateObject("Scripting.Dictionary")
        Dim KCurso      As Variant
        Dim NFound      As Long
        Dim NNot        As Long
        Dim ImpCurso    As Currency
        Dim Primero     As Boolean:     Primero = True
        DCursos.Add C_Acad_Ant, 0
        DCursos.Add C_Acad_Pos, 0
        For Each KCurso In DFound.Keys
            If Not DCursos.Exists(KCurso) Then DCursos.Add KCurso, 0
        Next KCurso
        For Each KCurso In DNot.Keys
            If Not DCursos.Exists(KCurso) Then DCursos.Add KCurso, 0
        Next KCurso
        For Each KCurso In DCursos.Keys
            NFound = 0:     If DFound.Exists(KCurso) Then NFound = DFound(KCurso)
            NNot = 0:       If DNot.Exists(KCurso) Then NNot = DNot(KCurso)
            ImpCurso = 0:   If DImp.Exists(KCurso) Then ImpCurso = DImp(KCurso)
            If NFound + NNot = 0 Then
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "C_Acad " & KCurso & ":  sin recibos INSS en BD_INSS.", 0)
                Primero = False
            Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "C_Acad " & KCurso & ":  Recibos NO encontrados: ", 0, _
                                Format(NNot, "#,##0") & " reg.", " de " & Format(NFound + NNot, "#,##0") & " reg.")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "Incorporados datos de: ", 0, _
                                Format(NFound, "#,##0") & " reg.", " de " & Format(NFound + NNot, "#,##0") & " reg.")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "Total Importe Seguro Obl. INSS", 0, _
                                Format(ImpCurso, "#,##0.00€     "), " de " & Format(NFound, "#,##0") & " reg.")
            Primero = False
            End If
        Next KCurso
Lo_BD.ShowTotals = True
Sht__BD_INSS.Range("a1").Select

Restablecer_Valores:
    Application.Speech.Speak "Proceso completado."
End Sub     ' Rut_Copy_ImpINSS_en_BDatos     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

'- ----------------------------------------------------------------------------------------------------------------------------
'- Acumuladores por Curso Académico del informe de Rut_Copy_ImpINSS_en_BDatos (Dictionary: curso -> total) ----------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Private Sub Rut_INSS_Acum(ByVal Dic As Object, ByVal Clave As String, ByVal Valor As Currency)
    If Dic.Exists(Clave) Then Dic(Clave) = Dic(Clave) + Valor Else Dic.Add Clave, Valor
End Sub
Private Function Fnc_INSS_Importe(ByVal Valor As Variant) As Currency
    If IsNumeric(Valor) Then Fnc_INSS_Importe = CCur(Valor)       '- Vacío o texto = 0
End Function

