Attribute VB_Name = "M_415_Copy_AE4_a_BD"
'2026-01-01
'-M_415_Copy_AE4_a_BD
Option Explicit


            Sub Rut_Copy_AE4x4_en_BDatos_ByHand()
                
                Call Rut_Copy_AE4x4_en_BDatos
                MsgBox "FIN"
'                Application.Speech.Speak "Proceso completado."
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- M_415_Copy_AE4_a_BD, Copiar los recibos AE4x4 a BDatos_Ant -----------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Copy_AE4x4_en_BDatos()
Debug.Print ">>> Rut_Copy_AE4x4_en_BDatos"

    Dim CantImpINSS      As Long:        CantImpINSS = 0
    Dim ImpTINSS         As Currency:      ImpTINSS = 0
    Dim F_BD            As Long
    Dim F_BDINSS        As Long
    Dim Found           As Long:        Found = 0
    Dim NotFound        As Long:        NotFound = 0
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim TimeLapSub      As Single:      TimeLapSub = LastTimeLap
    Dim TxT_Progreso    As String
    Dim TxT_ProgIni     As String
    Dim rowfind         As Variant
    
    Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_AE4          As ListObject:      Set Lo_AE4 = Sht__BD_AE4x4.ListObjects(1)
    Dim TF_BD           As Long:            TF_BD = Lo_BD.ListRows.Count
    Dim TF_AE4          As Long:            TF_AE4 = Lo_AE4.ListRows.Count
    
    Sht__BD.Visible = xlSheetVisible:           Sht__BD.Unprotect:          Lo_BD.ShowTotals = False
    Sht__BD_AE4x4.Visible = xlSheetVisible:     Sht__BD_AE4x4.Unprotect ':    Lo_AE4.ShowTotals = False
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    
    ' Optimizar entorno
    Dim SwScrUp     As Boolean:             SwScrUp = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Dim Sw_Calculation      As Boolean:     Sw_Calculation = Application.Calculation:   Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Rut_Copy_AE4x4_en_BDatos:
    '-            1º Borrar en Bdatos, todos los AE4 de los recibos de C_Acad_Ant y C_Acad_Pos
    '                   OJO sólo los de C_Acad Ant y Pos porque ocurrió que habían Rec. de Cursos más antiguos !!!
    '-            2º Copiar Lo_AE4 en BDatos.
    '- --------------------------------------------------------------------------------------------------------------
    
        '- Visualizo el progreso  <<<<>>>>
'        TxT_ProgIni = ActivForm.Controls("TBx_Informe")
'        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Procedimiento: Incorporar en BDatos, AE4: EFP y CFCyAFC de los Cursos " & C_Acad_Ant & " y " & C_Acad_Pos, 0)
        TxT_Progreso = ActivForm.Controls("TBx_Informe") & vbLf

    '- --------------------------------------------------------------------------------------------------------------
    '- Borrar en Bdatos, todos los AE4 ------------------------------------
        '- Visualizo el progreso  <<<<>>>>
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, "_") & "Borrando de BDatos Rec. AE4_" & C_Acad_Ant & " y AE4_" & C_Acad_Pos & " , wait please!", 0, , , , , , 3)
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Call Rut_Lo_WrkSht_Preparar(Sht__BD_AE4x4)
    rowfind = Lo_BD.ListRows.Count
    Call Rut_Lo_Sort(Lo_BD, BD_ActivEco, xlAscending, True)     '- Ordenar primero accelera un montón el borrado
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, False)      '- Ordenar primero accelera un montón el borrado
    Lo_BD.Range.AutoFilter Field:=BD_ActivEco, Criteria1:="=4"
    Lo_BD.Range.AutoFilter Field:=BD_C_Acad, Criteria1:="=" & C_Acad_Ant, Operator:=xlOr, Criteria2:="=" & C_Acad_Pos   '- Filtro los Recibos del Curso-Acad-Ant/Pos
    rowfind = Lo_BD.Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1   '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
    If rowfind > 0 Then
        Dim RngDel      As Range
        Set RngDel = Lo_BD.DataBodyRange.SpecialCells(xlCellTypeVisible)
        RngDel.EntireRow.Delete xlShiftUp
'        Lo_BD.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        '- Visualizo el progreso  <<<<>>>>
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del en BDatos Rec. AE4_" & C_Acad_Ant & " y AE4_" & C_Acad_Pos, 0, _
                                                        Format(rowfind, " #,##0") & " reg", _
                                                        "de " & Format(Lo_BD.ListRows.Count, "#,##0") & " reg", TxT_Progreso)
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay en BDatos Rec. AE4_" & C_Acad_Ant & " o AE4_" & C_Acad_Pos, 0, _
                                                        , "tiene " & Format(Lo_BD.ListRows.Count, "#,##0") & " reg", TxT_Progreso)
    End If
    
    '- Copiar Lo_AE4 en BDatos. ---------------------
        '- Visualizo el progreso  <<<<>>>>
        TxT_Progreso = ActivForm.Controls("TBx_Informe")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(10, "_") & "Copiando en BDatos Rec. AE4 y C_Acad " & C_Acad_Ant & " y " & C_Acad_Pos & " , wait please!", 0, , , , , , 3)
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    Call Rut_Lo_Filtros_Quitar(Lo_AE4)
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_AE4, Lo_BD)
        '- Visualizo el progreso  <<<<>>>>
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Copy en BDatos, Rec. AE4_" & C_Acad_Ant & " y AE4_" & C_Acad_Pos, 0, _
                                                        Format(Lo_AE4.ListRows.Count, " #,##0") & " reg", _
                                                        "BD= " & Format(Lo_BD.ListRows.Count, "#,##0") & " reg", TxT_Progreso)
    
Lo_BD.ShowTotals = True
Lo_AE4.ShowTotals = True

Restablecer_Valores:
    ' Restaurar entorno
    Application.ScreenUpdating = SwScrUp
    Application.Calculation = Sw_Calculation
    Application.EnableEvents = True
    
    Application.Speech.Speak "Proceso completado."
End Sub     ' Rut_Copy_AE4x4_en_BDatos     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================



