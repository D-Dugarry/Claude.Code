Attribute VB_Name = "M06_Asign_Imp_Adm_CAcad"
' Last Rev. 2026-09-21 12:12
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
    Dim RowData         As ListRow

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
    Call Rut_Lo_Sort(Lo_G04, BD_Plan, xlAscending, True)
    Call Rut_Lo_Sort(Lo_G04, BD_DNI, xlAscending, False)
    Call Rut_Lo_Sort(Lo_G04, BD_NumRec, xlAscending, False)
    Call Rut_Lo_Sort(Lo_G04, BD_Ref, xlAscending, False)

    '- Recorro toda la tabla Lo_G04 ----------------------------------------------------------------
    For fila = 1 To TRows_G04
        Set RowData = Lo_G04.ListRows(fila)
        If RowData.Range(BD_ImpRec) < 0 Then GoTo Sig_Reg                 '- NO tenemos en cuenta loas Recibos Negativos
'        If RowData.Range(BD_Anul) = "S" Then GoTo Sig_Reg                 '- NO tenemos en cuenta loas Recibos Anulados
        PlanDNI_New = RowData.Range(BD_Plan) & "_" & RowData.Range(BD_DNI)
        If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
           PlanDNI_Ant = PlanDNI_New
                                RowData.Range(BD_Rec_Imp_Acad) = RowData.Range(BD_ImpAcad)
                                RowData.Range(BD_Rec_Imp_Adm) = RowData.Range(BD_ImpAdm)
                                RowData.Range(BD_Rec_Imp_Dto) = RowData.Range(BD_ImpDto)
        End If
Sig_Reg:
    Next
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
            ImpTAcad = Application.Sum(.Columns(BD_Rec_Imp_Acad))
            ImpTAdm = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), .Columns(BD_Rec_Imp_Adm), ">0")
            ImpTDto = Application.Sum(.Columns(BD_Rec_Imp_Dto))
            CantImpAcad = Application.Count(.Columns(BD_Rec_Imp_Acad))
            CantImpAdm = Application.CountIfs(.Columns(BD_Rec_Imp_Adm), ">0")
            CantImpDto = Application.Count(.Columns(BD_Rec_Imp_Dto))
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


