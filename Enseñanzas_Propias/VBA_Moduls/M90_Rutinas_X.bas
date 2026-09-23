Attribute VB_Name = "M90_Rutinas_X"
' Last Rev. 2026-09-23 01:20
Option Explicit

'' =================================================================================================
'            '######################################################################################
'                    Sub Rut_WrkSheet_ReducirPeso_xx()
''                        Application.Workbooks(ThisWorkbook.Name).Sheets(ActiveSheet.Name).ListObjects(1).DataBodyRange.Delete
'                        Call Rut_WrkSheet_ReducirPeso(ActiveSheet.Name)
'                    End Sub
''##################################################################################################
'Sub Rut_WrkSheet_ReducirPeso(ByVal WrkSht As String)  '--- Borra TODO a la Derecha y Abajo de .ListObjects(1)
'' =================================================================================================
'        Application.Calculation = xlManual
'        Application.ScreenUpdating = False
'    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
'            .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
'            .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
'            If .FilterMode Then .ShowAllData            ' Deshacer Filtros
'        With .ListObjects(1).Range                      ' Borra la filas de abajo y columnas de la derecha del la Tabla .ListObjects(1)
'            Range(.Cells(.Rows.Count, .Columns.Count).Address).Select   ' Selecciona la última celda de la tabla
'            ActiveCell.Offset(1, 1).Select
'            Range(ActiveCell.Address & ":" & Cells(Rows.Count, 1).Address).EntireRow.Delete
'            Range(ActiveCell.Address & ":" & Cells(1, Columns.Count).Address).EntireColumn.Delete
'            ActiveSheet.UsedRange                       ' Para restablecer el rango de celdas en uso
'        End With
'    End With
'        Application.ScreenUpdating = True
'        Application.Calculation = xlAutomatic
'End Sub
'' -------------------------------------------------------------------------------------------------
' ==================================================================================================
Sub Rut_x_Filtro_LoTb(ByRef LoTb As ListObject, columna As Integer, Criterio As String, Optional SW_Clear As Boolean = False)
' --------------------------------------------------------------------------------------------------
    If SW_Clear And Not LoTb.AutoFilter Is Nothing Then LoTb.AutoFilter.ShowAllData
    LoTb.Range.AutoFilter Field:=columna, Criteria1:=Criterio
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_x_Filtros_Quitar_ActivSheet_LstObj()      ' Muestra Todas las Solicitudes y Activar Filtros
' --------------------------------------------------------------------------------------------------
    With ActiveSheet.ListObjects(1)
        If .ShowAutoFilter Then
            With .AutoFilter
                 If .FilterMode Then .ShowAllData
            End With
        Else
            .ShowAutoFilter = True
        End If
    End With
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_x_Ocultar_Cuadro_Instrucciones()
' --------------------------------------------------------------------------------------------------
    ActiveSheet.Shapes("Cuadro_Instrucciones").Visible = False
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================


' ==================================================================================================
Sub Rut_Columnas_ShowHide_Row_1()   ' Sirve para cualquier Hoja y para la Tabla nº.1
Dim Cont_Col    As Integer
Rut_Off_Functions
If ActiveSheet.ProtectContents = True Then ActiveSheet.Protect UserInterfaceOnly:=True
With ActiveSheet.Shapes.Range(Array("Botón_Col_ShowHide_row1"))
    If .TextFrame.Characters.Text = "Show Col's" Then                    ' Mostrar Columnas
        .Fill.ForeColor.RGB = RGB(200, 200, 250)
        .TextFrame.Characters.Text = "Hide Col's"
        Columns.EntireColumn.Hidden = False
    Else                                        ' Ocultar Columnas
        .Fill.ForeColor.RGB = RGB(150, 200, 100)
        .TextFrame.Characters.Text = "Show Col's"
        For Cont_Col = 1 To ActiveSheet.ListObjects(1).Range.Columns.Count
            If IsEmpty(Cells(1, Cont_Col)) And Not Columns(Cont_Col).Hidden Then Columns(Cont_Col).Hidden = True
        Next Cont_Col
    End If
End With
Rut_On_Functions
End Sub     ' Rut_Columnas_ShowHide     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================

' ==================================================================================================
Sub Rut_Columnas_Show_Hide()
Dim Cont    As Integer
Dim ContC   As Integer
Dim fila    As Integer
Dim Clmn    As Integer
Rut_Off_Functions
ActiveSheet.Unprotect
        fila = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Row - 1
        Clmn = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Column
        If fila < 2 Then MsgBox "Error, la Tabla debe empezar al menos en la fila 3", vbExclamation + vbOKOnly: Exit Sub
        
    For Cont = Clmn To ActiveSheet.ListObjects(1).Range.Columns.Count + Clmn
        If Columns(Cont).Hidden Then ContC = ContC + 1
    Next Cont


    If ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Mostrar Col's" Then                 ' Mostrar Columnas
        SW_Col_Ocultas = False
        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(200, 200, 250)
        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Ocultar Col's"
        Columns.EntireColumn.Hidden = False
    Else                                        ' Ocultar Columnas
        SW_Col_Ocultas = True
        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(150, 200, 100)
        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Mostrar Col's"
        
        Columns.EntireColumn.Hidden = False
        For Cont = Clmn To ActiveSheet.ListObjects(1).Range.Columns.Count + Clmn
            If IsEmpty(Cells(fila, Cont)) And Not Columns(Cont).Hidden Then Columns(Cont).Hidden = True
        Next Cont
    End If
Rut_On_Functions
ActiveSheet.Protect allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True      '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA
End Sub     ' Rut_Columnas_Show_Hide     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================


' ==================================================================================================
Sub Rut_Quita_Ascii_160(ByRef Lo_Tb As ListObject, columna As Integer)
' ==================================================================================================
    Lo_Tb.ListColumns(6).DataBodyRange.Select
    Selection.Replace What:=Chr(160), Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'===================================================================================================
'Sub Rut_Columnas_Ajustar_Ancho()      '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'' =================================================================================================
'    Dim Cont_Col As Integer
'    Dim Ancho   As Integer
'    For Cont_Col = 1 To LastCol_Tb_Solicitudes
'        If Not Columns(Cont_Col).Hidden And Lo_Prog_Colns.DataBodyRange.Cells(7, Cont_Col) <> "Ocultar" Then
'            Ancho = Lo_Prog_Colns.DataBodyRange.Cells(4, Cont_Col)      '.Value2
'            Columns(Cont_Col).ColumnWidth = Ancho
'        End If
'    Next Cont_Col
'End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_Filas_Mostrar()      ' Muestra Todas las Solicitudes  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Rows.EntireRow.Hidden = False
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_Filas_Mostrar_WrkSht(WrkSht As String)      ' Muestra Todas las Solicitudes  >>>>>>>>>>>>>>>
    Sheets(WrkSht).Rows.EntireRow.Hidden = False
    Sheets(WrkSht).Rows.EntireRow.AutoFit
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_Columnas_Mostrar_WrkSht(WrkSht As String)     ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================
    Sheets(WrkSht).Columns.EntireColumn.Hidden = False
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
'Sub Rut_Columnas_Mostrar()      ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'' =================================================================================================
'Dim Cont_Col As Integer
'Application.ScreenUpdating = False
'        For Cont_Col = 1 To LastCol_Tb_Solicitudes
'            If Lo_Prog_Colns.DataBodyRange.Cells(7, Cont_Col) <> "Ocultar" Then Columns(Cont_Col).Hidden = False
'        Next Cont_Col
'Application.ScreenUpdating = True
'Form_Menu.TB_Informe = "Columnas Ocultas Visibles"
'End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' --------------------------------------------------------------------------------------------------
' ==================================================================================================
Sub Rut_Select_File(Título As String, ByRef NomFich As String, TipoFich As String)  '>>>>>>>>>>>>>>>
' ==================================================================================================
Dim fDialog As Office.FileDialog
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = Título
        .InitialFileName = ThisWorkbook.Path & "\"
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        .ButtonName = "Seleccionar"
        .Filters.Clear
        .Filters.Add "Sólo Ficheros Excel", TipoFich, 1
        If .Show = False Then
            MsgBox "Ha pulsado el botón <Cancelar>.", vbOKOnly, "Proceso: Seleccionar Archivo"
            NomFich = "Cancel"
        Else
            NomFich = .SelectedItems(1)
            'Nom_NewArch = Dir(Arch__EP_New)
        End If
    End With
Set fDialog = Nothing
' --------------------------------------------------------------------------------------------------
End Sub     '  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' ==================================================================================================
'Sub Rut_Visible_Hidde_Tablas_Prog()
'    Dim WkS As Worksheet
'    For Each WkS In Worksheets
'        If Left(WkS.Name, 5) = "Prog_" Then
'            Select Case Visibilidad_Hoja
'                Case Is = "Visible"
'                WkS.Visible = xlSheetVisible
'                Case Is = "Hidden"
'                WkS.Visible = xlSheetHidden
'                Case Is = "VeryHidden"
'                WkS.Visible = xlSheetVeryHidden
'            End Select  ' Case wks.Visible
'        End If
'    Next
'End Sub     '      <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================

' ==================================================================================================
Sub Rut_Quitar_Sombreados()

    ActiveSheet.ListObjects(1).DataBodyRange.Interior.Color = -1

End Sub     ' RuT_Quitar_Sombreados_DR     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================

' ==================================================================================================
'Sub Rut_Ocultar_Col_SN()
'Dim Cont    As Integer
'Dim Fila    As Integer
'Dim Clmn    As Integer
'Rut_Off_Functions
'    If SW_Col_Ocultas Then                    ' Mostrar Columnas
'        SW_Col_Ocultas = False
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(200, 200, 250)
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Ocultar Col's"
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN-2")).Visible = True
'        Columns.EntireColumn.Hidden = False
'    Else                                        ' Ocultar Columnas
'        SW_Col_Ocultas = True
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).Fill.ForeColor.RGB = RGB(150, 200, 100)
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN")).TextFrame.Characters.Text = "Mostrar Col's"
'        ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN-2")).Visible = False
'            Fila = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Row - 1
'            Clmn = ActiveSheet.ListObjects(1).Range.Cells(1, 1).Column
'        For Cont = Clmn To ActiveSheet.ListObjects(1).Range.Columns.Count + Clmn
'            If IsEmpty(Cells(Fila, Cont)) And Not Columns(Cont).Hidden Then Columns(Cont).Hidden = True
'            If IsEmpty(Prog_DefCol_BD.Cells(2, Cont)) And Not Columns(Cont).Hidden Then Columns(Cont).Hidden = True
'        Next Cont
'    End If
'Rut_On_Functions
'End Sub     ' RuT_Ocultar_Col_SN     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'' =================================================================================================

' ==================================================================================================
Sub Rut_Unhide_All_Sheets()     ' Para uso interno, hacer visible todas las hojas
    Dim WkS As Worksheet
 
    For Each WkS In ThisWorkbook.Worksheets
        WkS.Visible = xlSheetVisible
    Next WkS
End Sub
' ==================================================================================================
    

' ==================================================================================================
Sub Rut_Limpiar_BD_Deleted_Curso_Nuevo()   '- USO PUNTUAL: deja Prog_BD_Deleted vacia y sana
' ==================================================================================================
'-  Motivo (2026-09-23): la hoja quedo "viciada" -- el ListObject Tb_Deleted tenia sus registros,
'-  pero A CONTINUACION, FUERA de la tabla, habia miles de filas huerfanas (6.274 copias de una
'-  misma referencia) por un bug en la escritura de las altas de M07, ya corregido.
'-  El usuario decide empezar de cero esta hoja para el curso 2026-27.
'-
'-  RECORDATORIO de la operativa correcta (verificada 2026-09-23): a Prog_BD_Deleted van los
'-  recibos dados de baja SIN JI (filtro "=*Deleted *", con espacio). Los que tienen JI se quedan
'-  en Prog_BD marcados como "DeletedConJI", porque ya estan contabilizados y no se pueden quitar.
'- ---------------------------------------------------------------------------------------------
    Dim Lo_Del      As ListObject
    Dim ws          As Worksheet
    Dim FilasDentro As Long
    Dim UltFilaHoja As Long
    Dim FilPrimLibre As Long
    Dim Respuesta   As VbMsgBoxResult
    Dim RutaCopia   As String
    Dim Dummy       As String

    Set ws = Prog_BD_Deleted
    Set Lo_Del = ws.ListObjects(1)

    '- Diagnostico ANTES de tocar nada -----------------------------------------------------------
    If Lo_Del.DataBodyRange Is Nothing Then
        FilasDentro = 0
    Else
        FilasDentro = Lo_Del.ListRows.Count
    End If
    FilPrimLibre = Lo_Del.Range.Row + Lo_Del.Range.Rows.Count
    UltFilaHoja = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Respuesta = MsgBox("Se va a VACIAR por completo la hoja BD_Deleted:" & vbLf & vbLf & _
                       "   Registros DENTRO de la tabla : " & Format(FilasDentro, "#,##0") & vbLf & _
                       "   Ultima fila con datos        : " & Format(UltFilaHoja, "#,##0") & vbLf & _
                       "   1a fila libre tras la tabla  : " & Format(FilPrimLibre, "#,##0") & vbLf & vbLf & _
                       "Se hara una copia de seguridad del libro antes de continuar." & vbLf & vbLf & _
                       "Continuar?", vbYesNo + vbExclamation, "Limpiar BD_Deleted")
    If Respuesta <> vbYes Then Exit Sub

    Call Rut_Off_Functions
    On Error GoTo Gestion_Error

    '- 1) Copia de seguridad previa --------------------------------------------------------------
    RutaCopia = Fnc_CopSeg_Previa_Importacion("LimpiarBDDeleted")

    ws.Visible = xlSheetVisible
    ws.Unprotect
    Lo_Del.ShowTotals = False
    Call Rut_Lo_Filtros_Quitar(Lo_Del)

    '- 2) Vaciar la tabla, dejando cabecera + estructura -----------------------------------------
    If Not Lo_Del.DataBodyRange Is Nothing Then Lo_Del.DataBodyRange.Delete

    '- 3) Borrar TODO lo que quede fuera del ListObject ------------------------------------------
    '-    Se recalcula la geometria DESPUES de vaciar: la tabla ya es cabecera + 1 fila vacia.
    FilPrimLibre = Lo_Del.Range.Row + Lo_Del.Range.Rows.Count
    UltFilaHoja = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If UltFilaHoja >= FilPrimLibre Then
        ws.Range(ws.Rows(FilPrimLibre), ws.Rows(ws.Rows.Count)).Delete
    End If

    '- 4) Compactar el rango usado ---------------------------------------------------------------
    Dummy = ws.UsedRange.Address        '- hay que LEER UsedRange para que se recalcule

    Lo_Del.ShowTotals = True
    ws.Visible = xlSheetVeryHidden

    Call Rut_On_Functions
    MsgBox "BD_Deleted limpia." & vbLf & vbLf & _
           "Se han eliminado " & Format(FilasDentro, "#,##0") & " registros de la tabla" & vbLf & _
           "y todo el contenido huerfano que habia fuera de ella." & vbLf & vbLf & _
           IIf(Len(RutaCopia) > 0, "Copia de seguridad:" & vbLf & RutaCopia, _
                                   "OJO: no se pudo hacer la copia de seguridad."), _
           vbOKOnly + vbInformation, "Limpiar BD_Deleted"
    Exit Sub

Gestion_Error:
    Dim ErrN As Long, ErrD As String
    ErrN = Err.Number: ErrD = Err.Description
    On Error Resume Next
    Call Rut_Reset_NestLevel
    On Error GoTo 0
    MsgBox "Error " & ErrN & " al limpiar BD_Deleted:" & vbLf & vbLf & ErrD, _
           vbOKOnly + vbCritical, "Limpiar BD_Deleted"
End Sub     ' Rut_Limpiar_BD_Deleted_Curso_Nuevo
' --------------------------------------------------------------------------------------------------
