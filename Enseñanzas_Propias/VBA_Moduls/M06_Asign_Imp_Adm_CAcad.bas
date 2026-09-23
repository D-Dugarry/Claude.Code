Attribute VB_Name = "M06_Asign_Imp_Adm_CAcad"
' Last Rev. 2026-09-23 18:56
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M06_Asign_Imp_Adm_CAcad - Imputar la tasa administrativa al primer recibo
' =================================================================================================
'
' PROPOSITO
'  Una matricula puede pagarse en varios plazos, y LSGES04 repite el importe
'  administrativo, el academico y el descuento en TODOS los recibos de esa
'  matricula. Sumarlos tal cual multiplicaria la tasa.
'  Esta rutina identifica el PRIMER recibo de cada matricula (Plan + DNI) y
'  solo a el le copia esos importes en las columnas Rec_Imp_*, que son las que
'  luego suman los informes.
'
' INDICE DE RUTINAS Y FUNCIONES
'  Rut_Assign_Imp_AdmAcad_C_Acad ... Unica rutina. Sin argumentos: trabaja
'                                    siempre sobre Prog_LsGes04.
'
' TRAMOS DE PROGRAMACION
'    1. Ordena la tabla por Plan, DNI, NumRec y Ref. El orden es lo que define
'       quien es el 'primer' recibo de cada matricula: sin el, la rutina imputa
'       la tasa a un recibo cualquiera.
'
'    2. Recorre la tabla llevando la clave anterior en PlanDNI_Ant:
'         - Salta los recibos con BD_ImpRec < 0 (negativos/devoluciones).
'         - Al cambiar la clave Plan_DNI copia, solo en esa fila:
'             BD_ImpAcad -> BD_Rec_Imp_Acad
'             BD_ImpAdm  -> BD_Rec_Imp_Adm
'             BD_ImpDto  -> BD_Rec_Imp_Dto
'         - El resto de recibos de la misma matricula quedan con Rec_Imp_* vacio.
'
'    3. Totaliza y vuelca al informe los tres importes (academico, administrativo
'       y descuento) con su numero de registros. El total administrativo se suma
'       con SumIfs > 0 para excluir los ajustes negativos de matricula.
'
'    La linea que excluia los recibos anulados (BD_Anul = 'S') esta comentada:
'    hoy los anulados SI reciben la imputacion.
'
' NOTAS
'  Distinguir siempre BD_ImpAdm (lo que repite LSGES04 en cada recibo) de
'  BD_Rec_Imp_Adm (lo imputado una sola vez). Los informes deben sumar el
'  segundo.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2025-01-14
Option Explicit

'- -------------------------------------------------------------------------------------------------
'- Identificar del C_Acad, los 1º Rec. de c/matrícula para obtener la T-Adm ------------------------
'- -------------------------------------------------------------------------------------------------
Sub Rut_Assign_Imp_AdmAcad_C_Acad()

Debug.Print ">>> Rut_Assign_Imp_AdmAcad_C_Acad"
'    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
'    Dim RowsDel         As Long
''    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    
    Dim TxtMsg1  As String, TxtMsg2  As String, TxtMsg3  As String
    Dim TxT_ProgIni     As String
    Dim TxT_Progreso    As String
    Dim PlanDNI_New     As String:
    Dim PlanDNI_Ant     As String:      PlanDNI_Ant = ""
    Dim C_Acad          As String:      C_Acad = Prog__APP.Range("APP_CursAcad")
    '- FASE 2: ya no se recorre con ListRow (un acceso COM por celda). Se vuelca la
    '-  tabla a un array, se calcula en memoria y se escribe de una vez.

    Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)
    Dim Lo_G04          As ListObject:      Set Lo_G04 = Prog_LsGes04.ListObjects(1)
    Dim TRows_G04       As Long:            TRows_G04 = Lo_G04.ListRows.Count
    Dim fila            As Long

    '- Visualizo el progreso -----------------------------------------------------------------------
    TxT_ProgIni = Form_Menu.TB_Informe
    Form_Menu.TB_Informe = TxT_ProgIni & vbLf & "Proceso: Identificar 1º Rec. para ImpAdm del Curso: " & _
                         C_Acad & ", en " & Format(Lo_G04.ListRows.Count, "#,##0") & " reg." & vbLf & vbLf
    TxT_Progreso = Form_Menu.TB_Informe
    
    ' Ordenar por columnas  ------------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_G04)
    Call Rut_Lo_Sort(Lo_G04, G04_Plan, xlAscending, True)
    Call Rut_Lo_Sort(Lo_G04, G04_DNI, xlAscending, False)
    Call Rut_Lo_Sort(Lo_G04, G04_NumRec, xlAscending, False)
    Call Rut_Lo_Sort(Lo_G04, G04_Ref, xlAscending, False)

    '- Recorro toda la tabla Lo_G04 ----------------------------------------------------------------
    '-  FASE 2 (velocidad): antes se recorria con Lo_G04.ListRows(fila) y 6 accesos COM por fila
    '-  (~50.000 cruces VBA<->Excel con 8.400 reg.). Ahora se lee la tabla a un array, se decide
    '-  todo en memoria y se escriben las 3 columnas destino de UNA sola vez.
    '-  La logica de negocio es EXACTAMENTE la misma, y el recorrido mantiene el mismo orden
    '-  (imprescindible: "el primer recibo de cada matricula" depende del orden de la tabla).
    If TRows_G04 > 0 Then
        Dim aDatos      As Variant      '- Copia en RAM de toda la tabla (solo lectura)
        Dim aSalida()   As Variant      '- Lo que se escribira en Rec_Imp_Acad / _Adm / _Dto

        aDatos = Lo_G04.DataBodyRange.Value          '- 1 sola lectura de bloque
        ReDim aSalida(1 To TRows_G04, 1 To 3)        '- 3 col. CONTIGUAS: Acad(30), Adm(31), Dto(32)

        For fila = 1 To TRows_G04
            '- Arrastra el valor previo, para no borrar nada que no toque esta rutina
            aSalida(fila, 1) = aDatos(fila, G04_Rec_Imp_Acad)
            aSalida(fila, 2) = aDatos(fila, G04_Rec_Imp_Adm)
            aSalida(fila, 3) = aDatos(fila, G04_Rec_Imp_Dto)

            '- Comparacion IDENTICA a la del codigo original (que hacia RowData.Range(BD_ImpRec) < 0).
            '-  NO usar Val(): con decimales europeos Val("-12,50") devuelve -12 (corta en la coma).
            If Not (aDatos(fila, G04_ImpRec) < 0) Then     '- NO tenemos en cuenta los Recibos Negativos
                PlanDNI_New = aDatos(fila, G04_Plan) & "_" & aDatos(fila, G04_DNI)
                If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la 1a Tasa Adm (las demas la repiten)
                    PlanDNI_Ant = PlanDNI_New
                    aSalida(fila, 1) = aDatos(fila, G04_ImpAcad)
                    aSalida(fila, 2) = aDatos(fila, G04_ImpAdm)
                    aSalida(fila, 3) = aDatos(fila, G04_ImpDto)
                End If
            End If
        Next

        '- 1 sola escritura de bloque sobre las 3 columnas contiguas (no toca ninguna otra)
        Lo_G04.DataBodyRange.Cells(1, G04_Rec_Imp_Acad).Resize(TRows_G04, 3).Value = aSalida

        Erase aSalida
        aDatos = Empty
        fila = TRows_G04 + 1     '- El informe de abajo usa "fila - 1" como nº de reg. procesados
    End If
    '- Visualizo el progreso -----------------------------------------------------------------------
                    Dim CantImpAcad     As Long:        CantImpAcad = 0
                    Dim CantImpAdm      As Long:        CantImpAdm = 0
                    Dim CantImpDto      As Long:        CantImpDto = 0
                    Dim ImpTAcad        As Currency:      ImpTAcad = 0
                    Dim ImpTAdm         As Currency:      ImpTAdm = 0
                    Dim ImpTDto         As Currency:      ImpTDto = 0
    With Lo_G04
        '- Visualizo el progreso -------------------------------------------------------------------
        Form_Menu.TB_Informe = TxT_ProgIni & vbLf & "Actualizado Tasas Adm. en: " & Format(fila - 1, "#,##0") & " reg., de " & Format(TRows_G04, "#,##0") & " reg." & vbLf
        '- Sumatorios ---------
        With .DataBodyRange
            ImpTAcad = Application.Sum(.Columns(G04_Rec_Imp_Acad))
            ImpTAdm = Application.SumIfs(.Columns(G04_Rec_Imp_Adm), .Columns(G04_Rec_Imp_Adm), ">0")
            ImpTDto = Application.Sum(.Columns(G04_Rec_Imp_Dto))
            CantImpAcad = Application.Count(.Columns(G04_Rec_Imp_Acad))
            CantImpAdm = Application.CountIfs(.Columns(G04_Rec_Imp_Adm), ">0")
            CantImpDto = Application.Count(.Columns(G04_Rec_Imp_Dto))
        End With
    End With
            TxtMsg1 = "     Total Acad. de Matrícula por un importe de:"
            TxtMsg2 = Format(ImpTAcad, "#,##0.00 €")
            TxtMsg3 = "en " & Format(CantImpAcad, "#,##0") & " reg."
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3) & vbLf
            TxtMsg1 = "     Total Adm.  de Matrícula por un importe de:"
            TxtMsg2 = Format(ImpTAdm, "#,##0.00 €")
            TxtMsg3 = "en " & Format(CantImpAdm, "#,##0") & " reg."
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3) & vbLf
            TxtMsg1 = "     Total Dto.  de Matrícula por un importe de:"
            TxtMsg2 = Format(ImpTDto, "#,##0.00 €")
            TxtMsg3 = "en " & Format(CantImpDto, "#,##0") & " reg."
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3) & vbLf
        
Debug.Print "<<< Rut_Assign_Imp_AdmAcad_C_Acad" & C_Acad
End Sub     ' --------------------------------------------------------------------------------------
' ==================================================================================================


