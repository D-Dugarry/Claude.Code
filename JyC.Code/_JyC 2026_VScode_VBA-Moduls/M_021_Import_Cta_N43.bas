Attribute VB_Name = "M_021_Import_Cta_N43"
'- M_021_Import_Cta_Norma43
Option Explicit

' ==================================================================================================================================
            Sub Call_Rut_Import_Cta_N43()
                Prog__APP.Range("APP_Task_Rut") = "Rut_Import_Cta_N43"
                Form_Running_Rut.Show
            End Sub
' ==================================================================================================================================
Sub Rut_Import_Cta_N43()    '- Importa fich.N43 a la hoja N43_TxT y de esta lo pasa a la N43_CTA decodificando los campos y formateándolos
' ==================================================================================================================================
    Dim ID_Reg          As String           ' 1º y2º carácter, para identificar el tipo de línea de datos
    Dim Indice          As String           ' 3º y 4º carácter, para identificar de las líneas nº23, el orden de los Reg. Complementarios del Movimiento
    Dim LinTxtN43       As String           ' Toda la línea de datos
    Dim AñoCont         As String:      AñoCont = CStr(Right(Prog__APP.Range("App_AñoCont"), 2))    ' Para Controlar el cambio de años el en número de orden que genero
    
    Dim TotLin_TxT_N43  As Integer
    Dim Ld              As Integer                  ' Línea de Detalle
    Dim L_Cta           As Integer:     L_Cta = 0   ' Línea de Resultante
    Const Cr            As Integer = 1              ' Primera Columna de Línea Resultante con datos
    
    Dim Saldo_Ini       As Single
    Dim Saldo_Final     As Single
    Dim Nom_Cta         As String
    Dim Ordinal         As Long:        Ordinal = 0
    Dim SkipReg         As Boolean:     SkipReg = False
    Dim Sw_Canceled     As Boolean
    Dim NewRow          As ListRow
    
    Dim ActivForm    As Object      '- Identificamos qué Formulario está Activo.  ------------------------------
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)
    
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    Rut_Off_Functions
    Prog_N43_TxT.Visible = xlSheetVisible
    Prog_N43_CTA.Visible = xlSheetVisible
    
    ' ---------------        Importar Fichero de Texto Norma43     -------------------------------'
    Call Rut_Importar_Fichero_Norma43(Sw_Canceled)    '- Si Sw_Canceled=True el proceso ha sido Cancelado
    If Sw_Canceled Then
        With ActivForm.Controls("TBx_Informe")
            .Value = MsgBx_Msg
            ActivForm.Repaint
        End With
        GoTo Restablecer_Valores
    End If
    Prog_CTA_Tb.Visible = xlSheetVisible
    Prog_CTA_Tb.Unprotect
    Prog_CTA_Tb.Select
    
    Call Rut_LstObj_WrkSht_Preparar(Prog_CTA_Tb)
    Dim Lo_Cta          As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    Dim Last_Ordinal    As Long:        Last_Ordinal = Right(Lo_Cta.DataBodyRange.Cells(Lo_Cta.ListRows.Count, C_Cta_Ordinal), 6)
    Dim Saldo_Ultimo    As Single:      Saldo_Ultimo = Lo_Cta.DataBodyRange.Cells(Lo_Cta.ListRows.Count, C_Cta_Saldo)
    
    Lo_Cta.ShowTotals = False
              
        LinTxtN43 = Prog_N43_TxT.Cells(1, 1)
        If Mid(LinTxtN43, 3, 18) <> "008131910001030312" Then
            MsgBx_Msg = "No es la cuenta de JyC: 008131910001030312" & vbLf & _
                        "La Cta del fichero es la: " & Mid(LinTxtN43, 3, 18)
            MsgBx_Title = "Proceso: Importar Extracto-Norma43 de la Cuenta de JyC"
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(, True, , , "Stop"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
            With ActivForm.Controls("TBx_Informe")
                .Value = MsgBx_Msg
                ActivForm.Repaint
            End With
            GoTo Restablecer_Valores
        End If
              
              TotLin_TxT_N43 = Prog_N43_TxT.Cells(Rows.Count, 1).End(xlUp).Row
    For Ld = 1 To TotLin_TxT_N43 ' >>>>>>>>>>>>>>
        LinTxtN43 = Prog_N43_TxT.Cells(Ld, 1)
        ID_Reg = Left(LinTxtN43, 2)
        Select Case ID_Reg
        Case 22     '- 1ª línea de Detalle ----------------------------------------------------------------
            If AñoCont <> Mid(LinTxtN43, 17, 2) Then SkipReg = True: GoTo Siguiente_Lin '''' Ordinal = 0: AñoCont = Mid(LinTxtN43, 17, 2)
            SkipReg = False
            Ordinal = Ordinal + 1
            If Ordinal <= Last_Ordinal Then SkipReg = True: GoTo Siguiente_Lin
            L_Cta = L_Cta + 1
            Set NewRow = Lo_Cta.ListRows.Add
            With NewRow
                .Range(C_Cta_Ordinal) = "20" & Mid(LinTxtN43, 17, 2) & Right("00000" & Ordinal, 6)
                .Range(C_Cta_Bco) = Mid(LinTxtN43, 7, 4)
                .Range(C_Cta_F_OPE) = DateSerial(Mid(LinTxtN43, 11, 2), Mid(LinTxtN43, 13, 2), Mid(LinTxtN43, 15, 2))        'F.Op.
                .Range(C_Cta_F_VAL) = DateSerial(Mid(LinTxtN43, 17, 2), Mid(LinTxtN43, 19, 2), Mid(LinTxtN43, 21, 2))        'F.Valor
                .Range(C_Cta_N_Liq) = "x"
                .Range(C_Cta_Siglas) = "_Desconocido"
                If Mid(LinTxtN43, 28, 1) = 1 Then
                    .Range(C_Cta_Imp) = Val(Mid(LinTxtN43, 29, 14)) * -0.01
                Else
                    .Range(C_Cta_Imp) = Val(Mid(LinTxtN43, 29, 14)) * 0.01
                End If
                .Range(C_Cta_Imp).NumberFormat = "#,##0.00;[Red]-#,##0.00;0"
'                If L_Cta = 1 Then
'                    .Range(C_Cta_Saldo) = Saldo_Ultimo + .Range(C_Cta_Imp)
'                Else
'                    .Range(C_Cta_Saldo) = Lo_Cta.DataBodyRange.Cells(.Range(C_Cta_Saldo).Row - 1, C_Cta_Saldo) + .Range(C_Cta_Imp)
'                End If
                Saldo_Ultimo = Saldo_Ultimo + .Range(C_Cta_Imp)
                .Range(C_Cta_Saldo) = Saldo_Ultimo
                .Range(C_Cta_Saldo).NumberFormat = "#,##0.00;[Red]-#,##0.00;0"
                .Range(C_Cta_Doc) = Mid(LinTxtN43, 43, 10)
                .Range(C_Cta_Ref1) = Mid(LinTxtN43, 53, 50)
            End With
        Case 23     '- 2ª,3ª,4ª y 5ª Línea de Detalle (¡No tiene porque existir todas!) ------------------
            If SkipReg Then GoTo Siguiente_Lin
            Indice = Mid(LinTxtN43, 3, 2)
            NewRow.Range(C_Cta_Ref1 + Indice) = Mid(LinTxtN43, 5, 75)      'Reg. Complementarios de Mov.
        Case 11     '- Cabecera --------------------------------------------------------------------------
            
            Prog__APP.Range("APP_Task_Inf") = "Bco.:  " & Mid(LinTxtN43, 3, 4) & ",     Entidad:  " & Mid(LinTxtN43, 7, 4) & vbCrLf & _
                            "Cta.:  " & Mid(LinTxtN43, 11, 10) & vbCrLf & _
                            "Cta.:  " & Mid(LinTxtN43, 52, 50) & vbCrLf & _
                            "F.Inicio:  " & DateSerial(Mid(LinTxtN43, 21, 2), Mid(LinTxtN43, 23, 2), Mid(LinTxtN43, 25, 2)) & vbCrLf & _
                            "F.Final:   " & DateSerial(Mid(LinTxtN43, 27, 2), Mid(LinTxtN43, 29, 2), Mid(LinTxtN43, 31, 2))
            If Mid(LinTxtN43, 33, 1) = 1 Then
                Saldo_Ini = Val(Mid(LinTxtN43, 34, 14)) * -0.01 'Saldo Inic.
            Else
                Saldo_Ini = Val(Mid(LinTxtN43, 34, 14)) * 0.01    'Saldo Inic.
            End If
            With ActivForm.Controls("TBx_Informe")
                .Value = Prog__APP.Range("APP_Task_Inf")
                ActivForm.Repaint
            End With
        Case 33     '- Pié de informe --------------------------------------------------------------------
            If Mid(LinTxtN43, 59, 1) = 1 Then
                Saldo_Final = Val(Mid(LinTxtN43, 60, 14)) * -0.01 'Saldo Final
            Else
                Saldo_Final = Val(Mid(LinTxtN43, 60, 14)) * 0.01  'Saldo Final
            End If
            Prog__APP.Range("APP_Saldo_Cta_Fin") = Saldo_Final
            With H_INICI.Shapes.Range(Array("Saldo_Cta_JyC"))
                .TextFrame.Characters.Text = "Saldo Cta. = " & Format(Saldo_Final, "#,##0.00")
                .Fill.ForeColor.RGB = RGB(180, 200, 200)
            End With
            
            Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbCrLf & _
                            "Apuntes  Ant.: " & Format(Last_Ordinal, "#,##0") & vbCrLf & _
                            "Apuntes Total: " & Format(Val(Mid(LinTxtN43, 21, 5)) + Val(Mid(LinTxtN43, 40, 5)), "#,##0") & vbCrLf & _
                            "Apuntes  Debe: " & Format(Val(Mid(LinTxtN43, 21, 5)), "#,##0") & ", tot.: " & Format(Val(Mid(LinTxtN43, 26, 14)) * -0.01, "#,##0.00") & vbCrLf & _
                            "Apuntes Haber: " & Format(Val(Mid(LinTxtN43, 40, 5)), "#,##0") & ", tot.: " & Format(Val(Mid(LinTxtN43, 45, 14)) * 0.01, "#,##0.00") & vbCrLf & _
                            "Saldo  Inic.: " & Format(Val(Saldo_Ini), "#,##0.00") & vbCrLf & _
                            "Saldo Final.: " & Format(Val(Saldo_Final), "#,##0.00") & vbCrLf & vbCrLf & "Realizado: " & Now
            With ActivForm.Controls("TBx_Informe")
                .Value = Prog__APP.Range("APP_Task_Inf")
                ActivForm.Repaint
            End With
       End Select  ' ID_Reg = Left(LinTxtN43, 2)
Siguiente_Lin:
    Next Ld ' Ld = 1 To TotLin_TxT_N43     <<<<<<<<<<<<
    
    
    Lo_Cta.ListColumns(C_N43_Reg_Mov1).DataBodyRange.Resize(, 5).Select    '- Resize(,5) es para ampliar a 5 columnas desde "C_N43_Reg_Mov1"
    With Selection
        .Replace What:="ORDENANTE DE LA TRANSFERENCIA :", Replacement:="Ordenante:", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="NOMBRE DEL ORDENANTE", Replacement:="Ordenante: ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="BENEFICIARIO DE LA TRANSFERENCIA :", Replacement:="Beneficiario: ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="BENEFICIARIO ", Replacement:="Beneficiario: ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="OBSERVACIONES :", Replacement:="Observaciones: ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="CONCEPTO/OBSERVACIONES ", Replacement:="Concepto/Obs.: ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="PAIS DEL BANCO ORDENANTE", Replacement:="País Bco. ordenante", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="BANCO ORDENANTE", Replacement:="Bco. ordenante", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        Dim celda As Range      ' Quita los espacios en blanco repetidos >>>>>>>>>>>>>>>>>>>>>>>
            For Each celda In Selection
                celda.Value = WorksheetFunction.Trim(celda.Value)
            Next        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    End With
    
    For Ld = 1 To Lo_Cta.ListRows.Count
        Set NewRow = Lo_Cta.ListRows(Ld)
        NewRow.Range(C_Cta_Fusión) = NewRow.Range(C_Cta_Reg_Mov1) & " # " & NewRow.Range(C_Cta_Reg_Mov2) & " # " & NewRow.Range(C_Cta_Reg_Mov3) & " # " & _
                                     NewRow.Range(C_Cta_Reg_Mov4) & " # " & NewRow.Range(C_Cta_Reg_Mov5) & " # " & NewRow.Range(C_Cta_Inscrito)
    Next
    
    Lo_Cta.ListColumns(C_Cta_Fusión).DataBodyRange.Select
    With Selection
        .Replace What:="Ordenante: ", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="Beneficiario: ", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="Observaciones: ", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="Concepto/Obs.: ", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="País Bco. ordenante ", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="Bco. ordenante ", Replacement:="Bco. ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        ' Quita los espacios en blanco repetidos >>>>>>>>>>>>>>>>>>>>>>>
        For Each celda In Selection
            celda.Value = WorksheetFunction.Trim(celda.Value)
        Next
    End With
    
    Lo_Cta.ShowTotals = True
    
    If Ordinal <= Last_Ordinal Then
        MsgBx_Msg = "No hay movimientos nuevos en el extracto" & vbCrLf & vbCrLf & Prog__APP.Range("APP_Task_Inf")
        With ActivForm.Controls("TBx_Informe")
            .Value = .Value & vbLf & vbLf & "No hay movimientos nuevos en el extracto"
            ActivForm.Repaint
        End With
    End If
    
        '- Visualizo el progreso ---------------------------------------------------------------------------------------
        With ActivForm.Controls("TBx_Informe")
            .Value = .Value & vbLf & vbLf & "¡ Proceso concluido ! día: " & Format(Now(), "dd-mmm-yyyy ""a las"" hh:mm") & " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg."
            .SelStart = Len(.Text)
            .SetFocus
            ActivForm.Repaint
            Prog__APP.Range("APP_Task_Inf") = .Value
            Prog__APP.Range("APP_Task_Inf") = .Value
            MsgBx_Msg = Prog__APP.Range("APP_Task_Inf")
        End With

    Prog__APP.Range("APP_Date_Imp_Cta") = Date
   'Prog__APP.Range("APP_Date_Imp_Cta") = Format(Now(), "dd-mmm-yy")
   
'''    Call RefreshRibbon
    
Restablecer_Valores:
Rut_On_Functions
    Prog_N43_TxT.Visible = xlSheetVeryHidden
    Prog_N43_CTA.Visible = xlSheetVeryHidden
'    Prog_CTA_Tb.Visible = xlSheetVeryHidden
End Sub     '- Rut_Import_Cta_N43

' ==================================================================================================================================
Sub Rut_Importar_Fichero_Norma43(ByRef Sw_Canceled As Boolean)  '- Importa Fichero.N43 a la hoja N43_TxT ----------------------
' ==================================================================================================================================
Dim Fichero_Seleccionado    As String
Dim Fichero_TxT             As String
Dim Linea_de_Texto          As String
Dim TempFileNum             As Integer
Dim Cont_Linea              As Long
Dim AñoCont         As String:      AñoCont = CStr(Right(Prog__APP.Range("App_AñoCont"), 2))    ' Para Controlar el cambio de años el en número de orden que genero

    Application.DefaultFilePath = Fnc_NEXE_RutaAPP

    Fichero_Seleccionado = Application.GetOpenFilename(FileFilter:="Fich.Bancario Norma43 (*.n43), *.n43,Fich.Bancario formato Txt, *.txt", _
                                                       Title:="Buscar y Seleccionar Archivo Norma43", MultiSelect:=False)
    If Fichero_Seleccionado = "Falso" Then
        Sw_Canceled = True
        MsgBox "Proceso Cancelado", vbExclamation, "Buscar y Seleccionar Archivo Norma43"
        End       ' Si no hay ningún fichero seleccionado
    Else
        Sw_Canceled = False
    End If
    Prog_N43_TxT.Select
    Rut_WrkSheet_Vaciar (Prog_N43_TxT.Name)
                                              
    Range("b1") = Fichero_Seleccionado
    Fichero_TxT = Fichero_Seleccionado        ' Es el nombre del fichero con su ruta
    Columns("A:A").Select
    Selection.NumberFormat = "@"
    
    TempFileNum = FreeFile    'Store the first file number in TempFileNum
    Cont_Linea = 1
    Open Fichero_TxT For Input As #TempFileNum
        '- Extracto la 1ª línea y compruebo la fecha de inicio del extracto, que debe ser el 1 de enero del Año Contable actual. --------------------------------------------
        Line Input #TempFileNum, Linea_de_Texto           'Read data from each line of text file and store it in variable Linea_de_Texto
        Range("A" & Cont_Linea).Value = Linea_de_Texto     'Storing the text file values in column A
        Cont_Linea = Cont_Linea + 1
            '- Compruebo que el extracto empieza el 1 de enero. --------------------------------------------------
            If "0101" <> Mid(Range("a1"), 23, 4) Then    'F.Ini.
                MsgBx_Msg = "El extracto no empieza el uno de enero, no se puede operar así."
                MsgBx_Title = "Proceso: Importar Extracto en formato Norma43."
                Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Exclam"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
                Sw_Canceled = True
                Exit Sub
            End If
            '- Compruebo que el Año del extracto que debe coincidir con el AñoCont. ------------------------------
            If AñoCont <> Mid(Range("a1"), 21, 2) Then   'F.Ini.
                MsgBx_Msg = "El extracto no corresponde al año contable en cuestión: " & Prog__APP.Range("App_AñoCont")
                MsgBx_Title = "Proceso: Importar Extracto en formato Norma43."
                Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Exclam"): Form_MsgBox.Show  '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
                Sw_Canceled = True
                Exit Sub
            End If
    '- Termino de copiar el resto del extracto. -------------------------------------------------------------------
    Do While Not EOF(TempFileNum)
        Line Input #TempFileNum, Linea_de_Texto           'Read data from each line of text file and store it in variable Linea_de_Texto
        Range("A" & Cont_Linea).Value = Linea_de_Texto     'Storing the text file values in column A
        Cont_Linea = Cont_Linea + 1
    Loop
    Close #TempFileNum
    Columns("A:B").Select
    Selection.EntireColumn.AutoFit
    Range("b1").Select
    
End Sub     ' Rut_Importar_Fichero_Norma43
'---------------------------------------------------------------------------------------------------------------------------


' ==================================================================================================================================
Sub kk()
' ==================================================================================================================================
Dim ID_Reg          As String           ' 1º y2º carácter, para identificar el tipo de línea de datos
Dim Indice          As String           ' 3º y 4º carácter, para identificar de las líneas nº23, el orden de los Reg. Complementarios del Movimiento
Dim LinTxtN43       As String           ' Toda la línea de datos
Dim AñoCont         As String:      AñoCont = CStr(Right(Prog__APP.Range("App_AñoCont"), 2))    ' Para Controlar el cambio de años el en número de orden que genero

Dim TotLin_TxT_N43  As Integer
Dim Ld              As Integer                  ' Línea de Detalle
Dim L_Cta           As Integer:     L_Cta = 0   ' Línea de Resultante
Const Cr            As Integer = 1              ' Primera Columna de Línea Resultante con datos

Dim Saldo_Ini       As Single
Dim Saldo_Final     As Single
Dim Nom_Cta         As String
Dim Ordinal         As Long:        Ordinal = 0
Dim SkipReg         As Boolean:     SkipReg = False
Dim Sw_Canceled     As Boolean
Dim NewRow          As ListRow
    Rut_Off_Functions
    Prog_CTA_Tb.Visible = xlSheetVisible
    Prog_CTA_Tb.Unprotect
    Prog_CTA_Tb.Select
    Call Rut_LstObj_WrkSht_Preparar(Prog_CTA_Tb)
    Dim Lo_Cta          As ListObject:      Set Lo_Cta = Prog_CTA_Tb.ListObjects(1)
    
    For Ld = 1 To Lo_Cta.ListRows.Count
            Set NewRow = Lo_Cta.ListRows(Ld)
            NewRow.Range(C_Cta_Fusión) = NewRow.Range(C_Cta_Reg_Mov1) & " # " & NewRow.Range(C_Cta_Reg_Mov2) & " # " & NewRow.Range(C_Cta_Reg_Mov3) & " # " & NewRow.Range(C_Cta_Reg_Mov4) & " # " & NewRow.Range(C_Cta_Reg_Mov5) & " # " & NewRow.Range(C_Cta_Inscrito)
    Next
    
    Lo_Cta.ListColumns(C_Cta_Fusión).DataBodyRange.Select
    With Selection
        .Replace What:="ORDENANTE DE LA TRANSFERENCIA :", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="NOMBRE DEL ORDENANTE", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="BENEFICIARIO DE LA TRANSFERENCIA :", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="BENEFICIARIO ", Replacement:=" ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="OBSERVACIONES :", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="CONCEPTO/OBSERVACIONES ", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="PAIS DEL BANCO ORDENANTE ", Replacement:="", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        .Replace What:="BANCO ORDENANTE ", Replacement:="Bco. ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
'        .Replace What:=" #  # ", Replacement:=" # ", _
                      LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                      SearchFormat:=False, ReplaceFormat:=False
        ' Quita los espacios en blanco repetidos >>>>>>>>>>>>>>>>>>>>>>>
        For Each celda In Selection
            celda.Value = WorksheetFunction.Trim(celda.Value)
            celda.Value = Mid(celda.Value, 3)
        Next        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    End With
    
    Lo_Cta.ShowTotals = True
    
    If Ordinal <= Last_Ordinal Then
        MsgBx_Msg = "No hay movimientos nuevos en el extracto" & vbCrLf & vbCrLf & Prog__APP.Range("APP_Task_Inf")
        MsgBx_Title = "Proceso: Importar Extracto-Norma43 de la Cuenta de JyC"
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "stop"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
    Else
        MsgBx_Msg = Prog__APP.Range("APP_Task_Inf")
        MsgBx_Title = "Proceso: Importar Extracto-Norma43 de la Cuenta de JyC"
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(, False, , , "Ask"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
    End If
    
Restablecer_Valores:
Rut_On_Functions
    Prog_N43_TxT.Visible = xlSheetVeryHidden
    Prog_N43_CTA.Visible = xlSheetVeryHidden
    Prog_CTA_Tb.Visible = xlSheetVeryHidden
End Sub     '-


