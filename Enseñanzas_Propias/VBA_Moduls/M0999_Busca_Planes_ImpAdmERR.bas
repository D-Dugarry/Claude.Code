Attribute VB_Name = "M0999_Busca_Planes_ImpAdmERR"
' Last Rev. 2026-09-20 23:21
' Modulo comentado en su totalidad (2026-09-20): C7 del Informe_Bugs - RowData.Range(BD_Rec_Imp_Adm) = -0.86
' escrito sobre Prog_BD (produccion) sin confirmacion ni idempotencia. Sub sin ninguna llamada en el proyecto
' (verificado con grep global); si se necesita reactivar, anadir antes un MsgBox de confirmacion y una
' comprobacion de 'ya aplicado' para no reescribir -0.86 sobre filas ya corregidas.
'2025-01-14
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Identificar Planes que tienen Más de un BD_Rec_Imp_Adm de un Plan de un mismo DNI
'- ----------------------------------------------------------------------------------------------------------------------------
'Sub Rut_Ajuste_BD_Rec_Imp_Adm()

'Debug.Print ">>> Rut_Assign_Imp_AdmAcad_C_Acad"
'    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
'    Dim RowsDel         As Long
''    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    
    'Dim TxT_Progreso    As String:      TxT_Progreso = "-Inicio-"
    'Dim Imp_Adm         As Currency
    'Dim Rec_Imp_Adm     As Currency
    'Dim Planes2Adm      As String:
    'Dim Plan_New     As String:
    'Dim Plan_Ant     As String:      Plan_Ant = ""
    'Dim PlanDNI_New     As String:
    'Dim PlanDNI_Ant     As String:      PlanDNI_Ant = ""
    'Dim C_Acad          As String:      C_Acad = Prog__APP.Range("APP_CursAcad")

    'Dim RowData         As ListRow
    'Dim Lo_BD          As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
    'Dim TRows_BD       As Long:            TRows_BD = Lo_BD.ListRows.Count
    'Dim fila            As Long
    'Dim contador        As Long
    'Dim Sw_Primer

    ' Ordenar por columnas  ------------------------------
    'Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    'Call Rut_Lo_Filtros_Quitar(Lo_BD)
    'Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, True)
    'Call Rut_Lo_Sort(Lo_BD, BD_DNI, xlAscending, False)
    'Call Rut_Lo_Sort(Lo_BD, BD_NumRec, xlAscending, False)
    'Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, False)

    '- Recorro toda la tabla Lo_BD -----------------------------------------------------------------
    'For fila = 1 To TRows_BD
        'Set RowData = Lo_BD.ListRows(fila)
        'If RowData.Range(BD_ImpRec) < 0 Then GoTo Sig_Reg                 '- NO tenemos en cuenta loas Recibos Negativos
'        If RowData.Range(BD_Anul) = "S" Then GoTo Sig_Reg                 '- NO tenemos en cuenta loas Recibos Anulados
        'PlanDNI_New = RowData.Range(BD_Plan) & "_" & RowData.Range(BD_DNI)
        'If PlanDNI_New <> PlanDNI_Ant Then   '--- Solo la primera Tasa Adm (es decir solo una tasa, porque las demás las repite)
            'PlanDNI_Ant = PlanDNI_New
                        'Debug.Print TxT_Progreso
                                'Imp_Adm = RowData.Range(BD_ImpAdm)
                                'Rec_Imp_Adm = RowData.Range(BD_Rec_Imp_Adm)
                                'TxT_Progreso = PlanDNI_New & "Imp_Adm=" & Imp_Adm & ", y Rec_Imp_Adm=" & Rec_Imp_Adm
                                'RowData.Range(BD_Rec_Imp_INSS) = RowData.Range(BD_Rec_Imp_Adm)
        'Else
                'If RowData.Range(BD_Rec_Imp_Adm) <> "" And RowData.Range(BD_Rec_Imp_Adm) <> 0 Then
                    'contador = contador + 1
                    'RowData.Range(BD_Rec_Imp_INSS) = RowData.Range(BD_Rec_Imp_Adm)
                    'RowData.Range(BD_Rec_Imp_Adm) = -0.86
                    
                    'If InStr(Planes2Adm, RowData.Range(BD_Plan)) = 0 Then
                        'Planes2Adm = Planes2Adm & " - " & RowData.Range(BD_Plan)
                        'Debug.Print RowData.Range(BD_Plan)
                    'End If
                    'If RowData.Range(BD_Rec_Imp_Adm) = Rec_Imp_Adm Then
                        'TxT_Progreso = TxT_Progreso & " -Repe-" & RowData.Range(BD_Rec_Imp_Adm)
                        
                    'Else
                        'TxT_Progreso = TxT_Progreso & " -Diff=" & RowData.Range(BD_Rec_Imp_Adm)
                    'End If
                    'Debug.Print TxT_Progreso
                'End If
        'End If
'Sig_Reg:
    'Next
        'Debug.Print "-FIN-   Repes y Diff = " & contador
        'Debug.Print "Planes: " & Planes2Adm
'End Sub     ' -------------------------------------------------------------------------------------------------------------------------<<<
' ========================================================================================================================================




