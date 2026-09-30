Attribute VB_Name = "M_520_Calcular_JIs_PPub_1303"
' Last Rev. 2026-09-30 21:12
Option Explicit

            ' ==================================================================================================================================
'- PENDIENTE: M_520 comentado entero (Sht__BD_AdmP y Sht__BD_JIs_303 no existen); resolver mas adelante.
            'Sub Call_Rut_Recalcular_Tabla_JIs_303()
                'Debug.Print "================== >>> Call_Rut_Recalcular_Tabla_JIs_303"
                'Dim ActivSheet  As String:  ActivSheet = ThisWorkbook.ActiveSheet.Name
                'Prog__APP.Range("APP_Task_Rut") = "Rut_Recalcular_Tabla_JIs_303"
            'On Error GoTo ManejoError
                        'DoEvents ' Permite que Excel procese eventos pendientes
                        'Form_Running_Rut.Show
            'On Error GoTo 0
                        'ThisWorkbook.Sheets(ActivSheet).Select
                        'Debug.Print "================== <<< Call_Rut_Recalcular_Tabla_JIs_303"
            'Exit Sub
'ManejoError:
            '        Unload Form_Running_Rut
            '        Call Rut_ConfigExcel_RESTABLECER
                'Static Intentos As Integer
                'If Err.Number = -2147417848 And Intentos < 5 Then
                    'Intentos = Intentos + 1
                    'Application.Wait Now + TimeValue("0:00:02") ' Espera 2 segundos
                    'Resume ' Reintenta la línea que falló
                'Else
                    'MsgBox "Error: " & Err.Description & vbCrLf & "Intentos: " & Intentos, vbCritical
                    'Intentos = 0
                'End If
                'MsgBox "<<< Err_Rut Form_Running_Rut >>>"
            'End Sub

' ==================================================================================================================================
'Sub Rut_Recalcular_Tabla_JIs_303()         '- EJECUTAR SÓLO --------------------
' ==================================================================================================================================
'Debug.Print "Rut_Recalcular_Tabla_JIs_303"
    
    '- Setting ListObjects ------------------------------------
'    Dim Lo_AdmP               As ListObject:      Set Lo_AdmP = Sht__BD.ListObjects(1)
    'Dim Lo_AdmP             As ListObject:      Set Lo_AdmP = Sht__BD_AdmP.ListObjects(1)
    'Dim Lo_JIs303           As ListObject:      Set Lo_JIs303 = Sht__BD_JIs_303.ListObjects(1)
    
    'Dim ContEmit        As Long:            ContEmit = 0
    'Dim ContEjeA        As Long:            ContEjeA = 0
    'Dim ContAnja        As Long:            ContAnja = 0
    'Dim ContSinTip      As Long:            ContSinTip = 0
    'Dim ContErrF        As Long:            ContErrF = 0
    'Dim TotReg          As Long:            TotReg = Lo_AdmP.ListRows.Count
    'Dim Lin             As Long
    'Dim L               As Integer
    'Dim ImpEmi          As Currency
    'Dim ImpCob          As Currency
    'Dim RngTEmi         As Range
    'Dim RngTEmiCob      As Range
    'Dim RngTEmiPdte     As Range
    'Dim RngTAant        As Range
    'Dim RngTAantCob     As Range
    'Dim RngTAantPdte    As Range
    'Dim RngTAnj         As Range
    'Dim RngTAnjCob      As Range
    'Dim RngTAnjPdte     As Range
    'Dim TxT_Progreso    As String
    'Dim TxT_Progreso2   As String
    'Dim AnoCont         As String:          AnoCont = Prog__APP.Range("APP_AnoCont")
    'Dim AnoEmi          As String
    'Dim AnoCob          As String
    'Dim TipoReg         As Range

    
    'Call Rut_Lo_Sort(Lo_AdmP, BD_ACont_Emi, xlAscending, True)
    
    'Lo_JIs303.HeaderRowRange.RowHeight = 30
    'With Lo_JIs303.DataBodyRange
        '.ClearContents
        '.Rows.RowHeight = 20
        'L = 1
        'Set RngTEmi = .Cells(L, JIs_Adm_Emi)
        'Set RngTEmiCob = .Cells(L, JIs_Adm_Cob)
        'Set RngTEmiPdte = .Cells(L, JIs_Adm_Pdte)
        '.Cells(L, JIs_TipRec) = "Emitido"
        '.Cells(L, JIs_ConcptEco) = "1303.00"
        '.Cells(L, JIs_ConcptNom) = "Servicios Administrativos"
        '.Cells(L, JIs_Descrip_JI) = "Liq.PPub_303.00__Emitido_" & Prog__APP.Range("APP_AnoCont") & "__Servicios Administrativos. "
        
        'L = 2
        'Set RngTAant = .Cells(L, JIs_Adm_Emi)
        'Set RngTAantCob = .Cells(L, JIs_Adm_Cob)
        'Set RngTAantPdte = .Cells(L, JIs_Adm_Pdte)
        '.Cells(L, JIs_TipRec) = "EjeAnt"
        '.Cells(L, JIs_ConcptEco) = "1303.00"
        '.Cells(L, JIs_ConcptNom) = "Servicios Administrativos"
        '.Cells(L, JIs_Descrip_JI) = "Liq.PPub_303.00__EjeAnt_" & Prog__APP.Range("APP_AnoCont") & "__Servicios Administrativos. "
        'L = 3
        'Set RngTAnj = .Cells(L, JIs_Adm_Emi)
        'Set RngTAnjCob = .Cells(L, JIs_Adm_Cob)
        'Set RngTAnjPdte = .Cells(L, JIs_Adm_Pdte)
        '.Cells(L, JIs_TipRec) = "Añejo"
        '.Cells(L, JIs_ConcptEco) = "1303.00"
        '.Cells(L, JIs_ConcptNom) = "Servicios Administrativos"
        '.Cells(L, JIs_Descrip_JI) = "Liq.PPub_303.00__Añejo_" & Prog__APP.Range("APP_AnoCont") & "__Servicios Administrativos. "
    'End With

    'With Lo_AdmP.DataBodyRange
        'For Lin = 1 To TotReg
            'AnoEmi = .Cells(Lin, BD_ACont_Emi)
            'AnoCob = .Cells(Lin, BD_ACont_Cob)
            'ImpEmi = .Cells(Lin, BD_ImpRec)
            'ImpCob = .Cells(Lin, BD_ImpCob)
            'Set TipoReg = .Cells(Lin, BD_Tipo_Rec)
            
            'If AnoCob = "" Then
                                            'TipoReg = "Emitido"
                                            'ContEmit = ContEmit + 1
                                            'RngTEmi = RngTEmi + ImpEmi
                                            'RngTEmiCob = RngTEmiCob + ImpCob
            'Else
                'Select Case AnoEmi
                    'Case Is > AnoCob
                                            'TipoReg = "ERR_Fech"
                                            'ContErrF = ContErrF + 1
                    'Case Is = AnoCont
                                            'TipoReg = "Emitido"
                                            'ContEmit = ContEmit + 1
                                            'RngTEmi = RngTEmi + ImpEmi
                                            'RngTEmiCob = RngTEmiCob + ImpCob
                    'Case Is = AnoCont - 1
                                            'TipoReg = "EjeAnt"
                                            'ContEjeA = ContEjeA + 1
                                            'RngTAant = RngTAant + ImpEmi
                                            'RngTAantCob = RngTAantCob + ImpCob
                    'Case Is < AnoCont - 1
                                            'TipoReg = "Añejo"
                                            'ContAnja = ContAnja + 1
                                            'RngTAnj = RngTAnj + ImpEmi
                                            'RngTAnjCob = RngTAnjCob + ImpCob
                    'Case Else
                                            'TipoReg = "¡SIN TIPO!"
                                            'ContSinTip = ContSinTip + 1
                'End Select
            'End If
               
        'Next Lin
        'RngTEmiPdte = RngTEmi - RngTEmiCob
        'RngTAantPdte = RngTAant - RngTAantCob
        'RngTAnjPdte = RngTAnj - RngTAnjCob
    'End With    '- Lo_AdmP.DataBodyRange

    'Sht__BD_JIs_303.Calculate
    'Sht__BD_JIs_303.Range("d2") = "Último Cálculo: " & Format(Now, "dd mmmm yyyy - hh:mm")
    'Sht__BD.Range("b3") = Sht__BD.Range("b3") & " Imp.Adm del Curso " & Prog__APP.Range("APP_C_Acad_Ant") & "."

    'Prog__APP.Range("APP_Last_Calc_JIs_303") = Format(Now(), "dd-mmm-yy hh:mm")
    
'End Sub


