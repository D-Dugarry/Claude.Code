Attribute VB_Name = "M09_Importar_Sol_Liq"
' Last Rev. 2026-09-14 12:25
'- M04_Importar_Sol_Liq
Option Explicit

'- TPV_Pagos ----------------------------------------------------
Public Const CSol_Plan             As Integer = 1       ' col: a
Public Const CSol_NomPlanVal       As Integer = 2       ' col: b
Public Const CSol_NomPlan          As Integer = 3       ' col: c
Public Const CSol_C_Acad           As Integer = 4       ' col: d
Public Const CSol_Nom              As Integer = 5       ' col: e
Public Const CSol_Expedte          As Integer = 6       ' col: f
Public Const CSol_DNI              As Integer = 7       ' col: g
Public Const CSol_Matricula        As Integer = 8       ' col: h
Public Const CSol_Anul             As Integer = 9       ' col: i
Public Const CSol_Ref              As Integer = 10       ' col: j
Public Const CSol_NumRec           As Integer = 11       ' col: k
Public Const CSol_ActivEco         As Integer = 12       ' col: l
Public Const CSol_FEmi             As Integer = 13       ' col: m
Public Const CSol_Fvnto            As Integer = 14       ' col: n
Public Const CSol_Fcob             As Integer = 15       ' col: o
Public Const CSol_ImpRec           As Integer = 16       ' col: p
Public Const CSol_ImpCob           As Integer = 17       ' col: q
Public Const CSol_PagTrjta         As Integer = 18       ' col: r
Public Const CSol_Comision         As Integer = 19       ' col: s
Public Const CSol_FormPag          As Integer = 20       ' col: t
Public Const CSol_RecMov           As Integer = 21       ' col: u
Public Const CSol_Grupo            As Integer = 22       ' col: v
Public Const CSol_Flist            As Integer = 23       ' col: w

Dim Finalizar_Proceso           As Boolean
Dim AñoCont                     As String           ' Para Controlar el cambio de años el en número de orden que genero

' ==================================================================================================================================
Sub Import_Sol_Liquid()
' ----------------------------------------------------------------------------------------------------------------------------------

Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '   Averigua el Número de Liquidación   ----------------------------------------------------------------------------------------
    Dim Pos_Ini     As Long
    Dim Pos_Fin     As Long
    Dim NumeLiquida      As String
    NumeLiquida = Form_Menu.TBx_Descripción
    Pos_Ini = InStr(NumeLiquida, "[[")
    Pos_Fin = InStr(NumeLiquida, "]]")
    If Pos_Ini * Pos_Fin = 0 Then   '---Controla que existe marca de inicio y fin y que hay algun dato entre marcas
        Form_Menu.TB_Informe = "Error: No hay un número de Liquidación, o no empieza por [[, o no acaba por ]]." & vbCrLf & Now
        Exit Sub
    End If
    NumeLiquida = Mid(NumeLiquida, Pos_Ini + 2, Pos_Fin - Pos_Ini - 2)    '---extraigo la NumeLiquida
    
    '   Seleccionar fichero     ---------------------------------------------------------------------------------------------------
    Dim Arch_Select         As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = Application.Workbooks(ThisWorkbook.Name).Path & "\"
        .Title = "Seleccionar el Fichero Excel de la Solicitud de Liquidación del Curso: "
        .ButtonName = "Aceptar"
        .AllowMultiSelect = False
        .Filters.Add "Sólo Ficheros Excel", "*.xls?", 1
        If .Show <> -1 Then
            MsgBx_Msg = "Cancelado"
            MsgBx_Title = "Proceso: Importar Solicitud de Liquidación de Curso"
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show '- (Font-Size, Red-Border, Buttons, Default-Button, Image)
'            APP_MnAux_Msg = "Proceso Cancelado: " & Now()
            GoTo Restablecer_Valores
        Else
            Arch_Select = .SelectedItems(1)
            'Nom_NewArch = Dir(Arch__EP_New)
        End If
    End With
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            Form_Menu.TB_Informe = "Importando Excel de Solicitud: " & Arch_Select & vbCrLf & Format(Now, "hh:mm:ss")
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
    '   Borrar el contenido de la hoja Prog_Sol_Liq    ------------------------------------------------------------
    Prog_Sol_Liq.Visible = xlSheetVisible
    Call Rut_WrkSheet_Vaciar(Prog_Sol_Liq.Name)
    '   Copio el excel    ------------------------------------------------------------------------------------------
    Dim ClsBk          As Variant
    Set ClsBk = Workbooks.Open(Arch_Select)
    ClsBk.Sheets(1).ListObjects(1).Range.Copy ThisWorkbook.Sheets(Prog_Sol_Liq.Name).Range("A1")
    Application.CutCopyMode = False
    ClsBk.Close SaveChanges:=False
    Set ClsBk = Nothing
    '   Si no viene con Tabla la Creo       ------------------------------------------------------------------------
    If Prog_Sol_Liq.ListObjects.Count = 0 Then
       Prog_Sol_Liq.ListObjects.Add(xlSrcRange, Sheets(Prog_Sol_Liq.Name).UsedRange, , xlYes).Name = "Tb_Sol_Liq"
    End If
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & "Importado Excel:       " & Format(Now, "hh:mm:ss")
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
    '----------------------------------------------------------
    '- Setting ListObjects ------------------------------------
    Dim Lo_Liq   As ListObject:     Set Lo_Liq = Wk_TitP_Liquid.ListObjects(1)
    Dim Lo_Sol   As ListObject:     Set Lo_Sol = Prog_Sol_Liq.ListObjects(1)
    '- Setting ListObjects ------------------------------------
    '----------------------------------------------------------

'    Prog_Sol_Liq.Select
    Prog_Sol_Liq.Visible = True
    Prog_Sol_Liq.Select
    Prog_Sol_Liq.Unprotect
    Call Rut_Lo_Filtros_Quitar(Lo_Sol)
    '- Formatear Columnas ------------------------------------------------------------------------------------------
    With Lo_Sol
'        .ShowTotals = False
        '- -----------------------------------------------------------------------------
'        .DataBodyRange.Columns(CSol_F_Emi).Select     '- Datos - Texto en Columnas PARA Números --------
'        Selection.TextToColumns Destination:=.DataBodyRange.Columns(CSol_F_Emi), DataType:=xlDelimited, _
'            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
'            :=Array(1, 4), TrailingMinusNumbers:=True
        .DataBodyRange.Columns(CSol_Ref).Select     '- Datos - Texto en Columnas - Finalizar - PARA NÚMEROS --------
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
            Selection.NumberFormat = "0000 000000000"
            Selection.Value = Selection.Value
        .DataBodyRange.Columns(CSol_ImpCob).Select
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
    End With
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & "Formateado Excel:      " & Format(Now, "hh:mm:ss")
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
    
    Call Rut_Lo_Sort(Lo_Sol, CSol_Ref, xlAscending, True)
    Call Rut_Lo_WrkSht_Preparar(Wk_TitP_Liquid)
    Call Rut_Lo_Sort(Lo_Liq, CLiq_Ref, xlAscending, True)

    Dim NotFound    As Long:        NotFound = 0
    Dim Encontrado  As Long:        Encontrado = 0
    Dim Cobrado     As Currency:    Cobrado = 0
    Dim LinSol      As Long
    Dim LinLiq      As Long:        LinLiq = 1
    Dim T_LinLiq    As Long:        T_LinLiq = Lo_Liq.ListRows.Count
    Wk_TitP_Liquid.Visible = True
    Wk_TitP_Liquid.Select
    Wk_TitP_Liquid.Unprotect
    Call Rut_Lo_WrkSht_Preparar(Wk_TitP_Liquid)
    Call Rut_Lo_Sort(Lo_Liq, CLiq_Ref, xlAscending, True)
    With Lo_Liq.DataBodyRange
        For LinSol = 1 To Lo_Sol.ListRows.Count
            If Lo_Sol.DataBodyRange.Cells(LinSol, CSol_Fcob) = "" Then GoTo Next_Lin
            Debug.Print Lo_Sol.DataBodyRange.Cells(LinSol, CSol_Ref), .Cells(LinLiq, CLiq_Ref)
            Select Case Lo_Sol.DataBodyRange.Cells(LinSol, CSol_Ref)
                '- OLD reg -----------------------------------------------------
                Case Is < .Cells(LinLiq, CLiq_Ref)
'                        Debug.Print "¡Error, No Existe la Tasa en la Tabla del Curso!    " & Lo_Sol.DataBodyRange.Cells(LinSol, CSol_Ref)
                    NotFound = NotFound + 1
                    GoTo Next_Lin
                '- Equal reg -----------------------------------------------------
                Case Is = .Cells(LinLiq, CLiq_Ref)
                    .Cells(LinLiq, CLiq_NumLiquid) = .Cells(LinLiq, CLiq_NumLiquid) & CStr(NumeLiquida)
                    Cobrado = Cobrado + .Cells(LinLiq, CLiq_Imp_Cob)
                    Encontrado = Encontrado + 1
                    If LinLiq < T_LinLiq Then LinLiq = LinLiq + 1
                '- NEW reg -----------------------------------------------------
                Case Is > .Cells(LinLiq, CLiq_Ref)
                    If LinLiq < T_LinLiq Then
                        LinLiq = LinLiq + 1
                        LinSol = LinSol - 1
                    Else
'                            Debug.Print "¡Error, He llegado al final de la Tabla de Liquidación del Curso y No Existe la Tasa !    " & Lo_Sol.DataBodyRange.Cells(LinSol, CSol_Ref)
                        NotFound = NotFound + 1
                        Exit For
                    End If
            End Select
Next_Lin:
        Next LinSol
    End With    ' Lo_Liq.DataBodyRange
    'Filtrar Liquidación -----------------------------------------------------------------------------------------------
    Call Rut_Lo_Sort(Lo_Liq, CLiq_Nombre, xlAscending, True)
    Lo_Liq.Range.AutoFilter Field:=CLiq_NumLiquid, Criteria1:="=" & NumeLiquida
    Lo_Liq.Range.Cells(1, 1).Select
    ActiveCell.Offset(1, 1).Select
            '- Visualizo el progreso ---------------------------------------------------------------------------------------
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & "Filtrado Excel:        " & Format(Now, "hh:mm:ss")
            Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False

'- Visualizo el progreso ---------------------------------------------------------------------------------------
Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!! día: " & Now() & " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & vbCrLf & _
        "Del Excel:  " & Mid(Arch_Select, InStrRev(Arch_Select, "\") + 1) & vbCrLf & vbCrLf & _
        "Ruta:  " & Left(Arch_Select, InStrRev(Arch_Select, "\")) & vbCrLf & vbCrLf & _
        "En la Solicitud del Curso:  " & Lo_Sol.DataBodyRange.Cells(1, CSol_Plan) & vbCrLf & vbCrLf & _
        Lo_Sol.DataBodyRange.Cells(1, CSol_NomPlan) & vbCrLf & vbCrLf & _
        Right("__________" & NotFound, 8) & "  Reg. no encontrados." & vbCrLf & _
        Right("__________" & Encontrado, 8) & "  Reg. encontrados." & vbCrLf & _
        Right("__________" & Lo_Sol.ListRows.Count - NotFound - Encontrado, 8) & "  Reg. Sin cobro." & vbCrLf & _
        Right("__________" & Lo_Sol.ListRows.Count, 8) & "  Reg. En la Solicitud." & vbCrLf & vbCrLf & _
        Right("__________" & Cobrado, 8) & "  Importe Cobrado en esta Liquidación." & vbCrLf
        
Restablecer_Valores:

Rut_On_Functions
'    Prog_N43_TxT.Visible = xlSheetVeryHidden
'    Prog_N43_CTA.Visible = xlSheetVeryHidden
End Sub     '- Extraer_Norma43
' ==================================================================================================================================




