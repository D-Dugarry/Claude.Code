Attribute VB_Name = "M_122_Copy_IAdm_a_BD"
' Last Rev. 2026-10-13 12:28
'- M_122_Copy_IAdm_a_BD ---------------------------------------------------------------------------------------------------------------
Option Explicit

'   Trasladar a BDatos los importes del LSace06 IAdm (hoja BD_ImpAdm, tabla Lo_ImpAdm, importada por M_120): Imp. INSS, Imp. Adm.,
'   Dto. Adm. y Tot. Dto. de cada recibo de C_Acad_Ant y C_Acad_Pos, en las 4 columnas _IAdm_ de BDatos (BD_IAdm_*).
'   Modelo: M_315 (INSS), con el mismo cruce en RAM por Ref (las dos tablas ordenadas por Ref, un solo recorrido).
'
'    - Primero se vacian las 4 columnas en los recibos de C_Acad_Ant y C_Acad_Pos (las de otros cursos no se tocan, para conservarlas
'      si BD_ImpAdm se queda con un solo curso). Luego se rellenan con los importes de BD_ImpAdm.
'    - BD_ImpAdm trae un solo registro por recibo (Ref), con los importes ya calculados: se copian tal cual, sin sumar ni redondear.
'    - La col. "Rec found" de BD_ImpAdm se rellena con "Found en BD." / "Not Found en BD.".
'    - El informe sale por curso: recibos encontrados / no encontrados e importe incorporado de cada una de las 4 columnas.

'==================================================================================================================================
Sub Rut_Copy_IAdm_a_BDatos()  '- Copiar Imp. INSS, Imp. Adm., Dto. Adm. y Tot. Dto. de BD_ImpAdm a BDatos (C_Acad_Ant y C_Acad_Pos)
'==================================================================================================================================
Debug.Print ">>> Rut_Copy_IAdm_a_BDatos"

    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("App_C_Acad_Ant")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("App_C_Acad_Pos")
    Dim F_BD            As Long
    Dim F_IAdm          As Long
    Dim Found           As Long
    Dim NotFound        As Long
    Dim Borrados        As Long
    Dim K_CAcad         As String
    Dim C               As Long
    Dim TxT_Progreso    As String
    Dim TxT_ProgIni     As String
    Dim DFound          As Object:      Set DFound = CreateObject("Scripting.Dictionary")   '- Por Curso Acad: Recibos encontrados
    Dim DNot            As Object:      Set DNot = CreateObject("Scripting.Dictionary")     '- Por Curso Acad: Recibos NO encontrados
    Dim DImp(1 To 4)    As Object                                                           '- Por Curso Acad: Importe incorporado de cada columna
    Dim DCnt(1 To 4)    As Object                                                           '- Por Curso Acad: Recibos que SI traen dato en cada columna
    Dim Cols_BD         As Variant:     Cols_BD = Array(BD_IAdm_ImpINSS, BD_IAdm_ImpAdm, BD_IAdm_DtoAdm, BD_IAdm_TotDto)
    Dim Cols_IAdm       As Variant:     Cols_IAdm = Array(IAdm_Imp_INSS, IAdm_Imp_Adm, IAdm_Dto_Adm, IAdm_Tot_Dto)
    Dim Nom_Imp         As Variant:     Nom_Imp = Array("Imp. INSS", "Imp. Adm. ", "Dto. Adm. ", "Tot. Dto. ")
    For C = 1 To 4
        Set DImp(C) = CreateObject("Scripting.Dictionary")
        Set DCnt(C) = CreateObject("Scripting.Dictionary")
    Next C

    Set ActivForm = Form_Running_Rut                      '- lo abre el boton del Ribbon (Rut_Progreso_Abrir, M_90)

    Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_IAdm         As ListObject:      Set Lo_IAdm = Sht__BD_ImpAdm.ListObjects(1)
    Dim TF_BD           As Long:            TF_BD = Lo_BD.ListRows.Count
    Dim TF_IAdm         As Long:            TF_IAdm = Lo_IAdm.ListRows.Count

    Sht__BD.Visible = xlSheetVisible:           Sht__BD.Activate:           Sht__BD.Unprotect
    Sht__BD_ImpAdm.Visible = xlSheetVisible:    Sht__BD_ImpAdm.Unprotect
    Dim TotalesVisibles As Boolean:     TotalesVisibles = Fnc_Lo_Totales_Ocultar(Lo_BD)     '- se restaura en TODAS las salidas

    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    Rut_Off_Functions               '- El On esta en Restablecer_Valores
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    Call Rut_Lo_Filtros_Quitar(Lo_IAdm)

    '- Visualizo el progreso
    TxT_ProgIni = ActivForm.Controls("TBx_Informe").Text
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Procedimiento: Incorporar Imp. INSS, Imp. Adm. y Dto. Adm. (IAdm), Cursos " & _
                                                    C_Acad_Ant & " y " & C_Acad_Pos, 0, , , , , , 4)
    TxT_Progreso = ActivForm.Controls("TBx_Informe").Text

    If TF_IAdm = 0 Or TF_BD = 0 Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Abortado: " & IIf(TF_BD = 0, "BDatos", "BD_ImpAdm") & " no tiene registros. No se ha tocado BDatos.", 0)
        GoTo Restablecer_Valores
    End If

    '- --------------------------------------------------------------------------------------------------------
    '- Las dos tablas ordenadas por Ref, para cruzarlas en un solo recorrido -----------------------------------
    Call Rut_Lo_Sort(Lo_IAdm, IAdm_Ref, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)

    '- El cruce se hace en RAM (Rut_Lo_TablaRam): BD y BD_ImpAdm se cargan con .Value2, y de BD se devuelven solo las 4 columnas _IAdm_
    Dim T_BD            As T_TablaRam
    Dim T_IAdm          As T_TablaRam
    Call Rut_TablaRam_Cargar(T_BD, Lo_BD, Array(BD_Ref, BD_C_Acad, BD_IAdm_ImpINSS, BD_IAdm_ImpAdm, BD_IAdm_DtoAdm, BD_IAdm_TotDto), True)
    Call Rut_TablaRam_Cargar(T_IAdm, Lo_IAdm, Array(IAdm_Ref, IAdm_C_Acad, IAdm_Imp_INSS, IAdm_Imp_Adm, IAdm_Dto_Adm, IAdm_Tot_Dto, IAdm_RecFound), True)

    '- Vacio las 4 columnas en los recibos de C_Acad_Ant y C_Acad_Pos ------------------------------------------
    For F_BD = 1 To TF_BD
        K_CAcad = CStr(T_BD.Datos(F_BD, BD_C_Acad))
        If K_CAcad = C_Acad_Ant Or K_CAcad = C_Acad_Pos Then
            For C = 0 To 3
                If Not IsEmpty(T_BD.Datos(F_BD, Cols_BD(C))) Then
                    If CStr(T_BD.Datos(F_BD, Cols_BD(C))) <> "" Then Borrados = Borrados + 1
                    T_BD.Datos(F_BD, Cols_BD(C)) = Empty
                End If
            Next C
        End If
    Next F_BD

    '- --------------------------------------------------------------------------------------------------------
    '- Actualizo BDatos con BD_ImpAdm. Mismo recorrido que M_315 ---------------------------------------------
    F_BD = 1
    For F_IAdm = 1 To TF_IAdm
        K_CAcad = CStr(T_IAdm.Datos(F_IAdm, IAdm_C_Acad))
        Select Case T_BD.Datos(F_BD, BD_Ref)
            Case Is < T_IAdm.Datos(F_IAdm, IAdm_Ref)       '- Ref_BD  NO-EXISTE-EN  BD_ImpAdm
                If F_BD < TF_BD Then
                    F_BD = F_BD + 1
                    F_IAdm = F_IAdm - 1
                Else
                    '- BD agotada: este y todos los IAdm restantes quedan sin encontrar --------
                    NotFound = NotFound + 1
                    T_IAdm.Datos(F_IAdm, IAdm_RecFound) = "Not Found en BD. (BD agotada)"
                    Call Rut_IAdm_Acum(DNot, K_CAcad, 1)
                End If
            Case Is = T_IAdm.Datos(F_IAdm, IAdm_Ref)       '- Ref_BD  SI-EXISTE-EN  BD_ImpAdm
                For C = 0 To 3
                    T_BD.Datos(F_BD, Cols_BD(C)) = Fnc_IAdm_Valor(T_IAdm.Datos(F_IAdm, Cols_IAdm(C)))
                    Call Rut_IAdm_Acum(DImp(C + 1), K_CAcad, Fnc_IAdm_Importe(T_IAdm.Datos(F_IAdm, Cols_IAdm(C))))
                    If Not IsEmpty(T_BD.Datos(F_BD, Cols_BD(C))) Then Call Rut_IAdm_Acum(DCnt(C + 1), K_CAcad, 1)   '- solo los que traen dato
                Next C
                Found = Found + 1
                T_IAdm.Datos(F_IAdm, IAdm_RecFound) = "Found en BD."
                Call Rut_IAdm_Acum(DFound, K_CAcad, 1)
                If F_BD < TF_BD Then F_BD = F_BD + 1
            Case Is > T_IAdm.Datos(F_IAdm, IAdm_Ref)       '- Ref_BD_ImpAdm  NO-EXISTE-EN  BDatos
                NotFound = NotFound + 1
                T_IAdm.Datos(F_IAdm, IAdm_RecFound) = "Not Found en BD."
                Call Rut_IAdm_Acum(DNot, K_CAcad, 1)
        End Select
        If (F_IAdm Mod 4000 = 0 And F_IAdm <> 0) Or F_BD Mod 4000 = 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Incorporados datos de: " & Format(Found, "#,##0") & _
                        " reg. de " & Format(F_IAdm, "#,##0") & "/" & Format(TF_IAdm, "#,##0") & "reg.,  entre " & _
                        Format(F_BD, "#,##0") & "/" & Format(TF_BD, "#,##0") & "reg.", 0, , , TxT_Progreso, True, , 2)
        End If
    Next F_IAdm
    For C = 0 To 3
        T_BD.Modificada(Cols_BD(C)) = True
    Next C
    T_IAdm.Modificada(IAdm_RecFound) = True
    Call Rut_TablaRam_Volcar(T_BD, Lo_BD)                               '- Devuelvo a la hoja las 4 Col. _IAdm_
    Call Rut_TablaRam_Volcar(T_IAdm, Lo_IAdm)                           '- y RecFound
    Erase T_BD.Datos                                                    '- Libero la RAM
    Erase T_IAdm.Datos

    '- Visualizo el progreso ----------------------------------------------------------------------------------------
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", vbLf & "Procedimiento: Incorporar Imp. INSS, Imp. Adm. y Dto. Adm. (IAdm), C_Acad " & _
                                                    C_Acad_Ant & " y " & C_Acad_Pos, 0, , , TxT_ProgIni, , , 2)
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Vaciados antes de incorporar, en BDatos: ", 0, Format(Borrados, "#,##0") & " datos")

    '- El resultado, por Curso Academico (Ant y Pos primero; cualquier otro curso que hubiera en BD_ImpAdm, despues) ---------------
    Dim DCursos     As Object:      Set DCursos = CreateObject("Scripting.Dictionary")
    Dim KCurso      As Variant
    Dim NFound      As Long
    Dim NNot        As Long
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
        If NFound + NNot = 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "C_Acad " & KCurso & ":  sin recibos IAdm en BD_ImpAdm.", 0)
        Else
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "C_Acad " & KCurso & ":  Recibos NO encontrados: ", 0, _
                                Format(NNot, "#,##0") & " reg.", " de " & Format(NFound + NNot, "#,##0") & " reg.")
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "Incorporados datos de: ", 0, _
                                Format(NFound, "#,##0") & " reg.", " de " & Format(NFound + NNot, "#,##0") & " reg.")
            For C = 1 To 4
                Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(26, " ") & "Total " & Nom_Imp(C - 1), 0, _
                                Format(Fnc_IAdm_Total(DImp(C), KCurso), "#,##0.00 EUR     "), " en " & Format(Fnc_IAdm_Total(DCnt(C), KCurso), "#,##0") & " reg.")
            Next C
        End If
    Next KCurso

    Prog__APP.Range("App_Task_Inf") = ActivForm.Controls("TBx_Informe").Text

Restablecer_Valores:
    Call Rut_Lo_Totales_Restaurar(Lo_BD, TotalesVisibles)
    Rut_On_Functions
Debug.Print "<<< Rut_Copy_IAdm_a_BDatos"
End Sub     ' Rut_Copy_IAdm_a_BDatos   ----------------------------------------------------------------------------------------------
'===================================================================================================================================

'- ----------------------------------------------------------------------------------------------------------------------------
'- Auxiliares del cruce -------------------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Private Sub Rut_IAdm_Acum(ByVal Dic As Object, ByVal Clave As String, ByVal Valor As Currency)
    If Dic.Exists(Clave) Then Dic(Clave) = Dic(Clave) + Valor Else Dic.Add Clave, Valor
End Sub
Private Function Fnc_IAdm_Total(ByVal Dic As Object, ByVal Clave As Variant) As Currency
    If Dic.Exists(Clave) Then Fnc_IAdm_Total = Dic(Clave)
End Function
Private Function Fnc_IAdm_Importe(ByVal Valor As Variant) As Currency
    If IsError(Valor) Then Exit Function
    If IsNumeric(Valor) And Not IsEmpty(Valor) Then Fnc_IAdm_Importe = CCur(Valor)       '- Vacio o texto = 0
End Function
Private Function Fnc_IAdm_Valor(ByVal Valor As Variant) As Variant                       '- Lo que se escribe en BDatos: numero, o vacio
    If IsError(Valor) Then
        Fnc_IAdm_Valor = Empty
    ElseIf IsEmpty(Valor) Then
        Fnc_IAdm_Valor = Empty
    ElseIf VarType(Valor) = vbString Then
        If Valor <> "" And IsNumeric(Valor) Then Fnc_IAdm_Valor = CDbl(Valor) Else Fnc_IAdm_Valor = Empty
    Else
        Fnc_IAdm_Valor = Valor
    End If
End Function
