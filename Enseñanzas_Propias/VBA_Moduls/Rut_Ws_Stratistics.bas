Attribute VB_Name = "Rut_Ws_Stratistics"
'2026-02-05
Option Explicit

Sub Rut_Ws_All_Stratistics_01()
    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim wsRep       As Worksheet
    Dim Lo          As ListObject
    Dim rngDest     As Range
    Dim ur          As Range
    Dim cel         As Range
    
    Dim nFilas      As Long
    Dim nCols       As Long
    Dim nCeldas     As Double
    Dim nFormulas   As Double
    Dim nConstantes As Double
    Dim nBlancas    As Double
    Dim nErrores    As Double
    
    Dim Sw_EnableEvents     As Boolean:     Sw_EnableEvents = Application.EnableEvents: Application.EnableEvents = False
    '=== Configuración básica ===
    Set wb = ThisWorkbook  ' o ActiveWorkbook, según prefieras
    
    'Hoja donde se crea el ListObject de métricas
    On Error Resume Next
    Set wsRep = wb.Worksheets("Sheets_Data_Stratistics")
    On Error GoTo 0
    If wsRep Is Nothing Then
        Set wsRep = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsRep.Name = "Sheets_Data_Stratistics"
    End If
    
    'Limpiar (opcional) pero sin borrar la tabla si existe
    wsRep.Cells.Clear
    
    Set rngDest = wsRep.Range("A1")
    
    '=== Encabezados ===
    rngDest.Offset(0, 0).Value = "CodeName"
    rngDest.Offset(0, 1).Value = "NombreHoja"
    rngDest.Offset(0, 2).Value = "Visible"
    rngDest.Offset(0, 3).Value = "TipoHoja"
    rngDest.Offset(0, 4).Value = "UR_FilaIni"
    rngDest.Offset(0, 5).Value = "UR_ColIni"
    rngDest.Offset(0, 6).Value = "UR_NumFilas"
    rngDest.Offset(0, 7).Value = "UR_NumCols"
    rngDest.Offset(0, 8).Value = "UR_NumCeldas"
    rngDest.Offset(0, 9).Value = "Num_Formulas"
    rngDest.Offset(0, 10).Value = "Num_Constantes"
    rngDest.Offset(0, 11).Value = "Num_Blancas"
    rngDest.Offset(0, 12).Value = "Num_Errores"
    
    Dim fila As Long
    fila = 2
    
    '=== Recorrer hojas ===
    For Each ws In wb.Worksheets
        
        'Evitar contarnos a nosotros mismos si no quieres incluir la hoja de métricas
        If ws.Name = wsRep.Name Then GoTo SiguienteWs
        
        'UsedRange puede estar "sucio"; forzamos un pequeño truco
        'para actualizarlo (opcional, pero a veces ayuda)
        Dim tmp As String
        tmp = ws.UsedRange.Address    'toca ligeramente el UsedRange
        
        If ws.UsedRange.Cells.Count = 1 And IsEmpty(ws.UsedRange) Then
            'Hoja vacía
            nFilas = 0
            nCols = 0
            nCeldas = 0
            nFormulas = 0
            nConstantes = 0
            nBlancas = 0
            nErrores = 0
            Set ur = Nothing
        Else
            Set ur = ws.UsedRange
            
            nFilas = ur.Rows.Count                    'filas del UsedRange [web:16]
            nCols = ur.Columns.Count                  'columnas del UsedRange [web:16]
            nCeldas = CDbl(ur.Cells.Count)
            
            'Celdas con fórmula
            On Error Resume Next
            nFormulas = ur.SpecialCells(xlCellTypeFormulas).Count   'usa SpecialCells para fórmulas [web:17]
            If Err.Number <> 0 Then nFormulas = 0
            Err.Clear
            
            'Celdas con constantes (valores)
            nConstantes = 0
            If Not ur Is Nothing Then
                For Each cel In ur
                    If Not cel.HasFormula Then
                        If Not IsEmpty(cel.Value) Then
                            nConstantes = nConstantes + 1
                        End If
                    End If
                Next cel
            End If
            
            'Celdas con error
            nErrores = 0
            If Not ur Is Nothing Then
                For Each cel In ur
                    If cel.HasFormula Or Not IsEmpty(cel.Value) Then
                        If IsError(cel.Value) Then nErrores = nErrores + 1
                    End If
                Next cel
            End If
            
            'Blancas = total UsedRange - fórmulas - constantes
            nBlancas = nCeldas - nFormulas - nConstantes
            If nBlancas < 0 Then nBlancas = 0
        End If
        
        '=== Volcar datos ===
        With wsRep
            .Cells(fila, 1).Value = ws.CodeName
            .Cells(fila, 2).Value = ws.Name
            .Cells(fila, 3).Value = IIf(ws.Visible = xlSheetVisible, "Visible", _
                                        IIf(ws.Visible = xlSheetHidden, "Oculta", "Muy oculta"))
            .Cells(fila, 4).Value = TipoDeHoja(ws)
            
            If Not ur Is Nothing Then
                .Cells(fila, 5).Value = ur.Row
                .Cells(fila, 6).Value = ur.Column
                .Cells(fila, 7).Value = nFilas
                .Cells(fila, 8).Value = nCols
                .Cells(fila, 9).Value = nCeldas
            Else
                .Cells(fila, 5).Value = ""
                .Cells(fila, 6).Value = ""
                .Cells(fila, 7).Value = 0
                .Cells(fila, 8).Value = 0
                .Cells(fila, 9).Value = 0
            End If
            
            .Cells(fila, 10).Value = nFormulas
            .Cells(fila, 11).Value = nConstantes
            .Cells(fila, 12).Value = nBlancas
            .Cells(fila, 13).Value = nErrores
        End With
        
        fila = fila + 1
SiguienteWs:
    Next ws
    
    '=== Crear / actualizar la ListObject ===
    Dim lastCol As Long
    Dim lastRow As Long
    lastCol = wsRep.Cells(1, wsRep.Columns.Count).End(xlToLeft).Column
    lastRow = wsRep.Cells(wsRep.Rows.Count, 1).End(xlUp).Row
    
    On Error Resume Next
    Set Lo = wsRep.ListObjects("Lo_Stratistics")
    On Error GoTo 0
    
    If Lo Is Nothing Then
        Set Lo = wsRep.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=wsRep.Range(wsRep.Cells(1, 1), wsRep.Cells(lastRow, lastCol)), _
                    XlListObjectHasHeaders:=xlYes)  'crea la ListObject [web:14]
        Lo.Name = "Lo_Stratistics"
    Else
        Lo.Resize wsRep.Range(wsRep.Cells(1, 1), wsRep.Cells(lastRow, lastCol))
    End If
    
    MsgBox "Tabla de métricas creada/actualizada.", vbInformation
    Application.EnableEvents = Sw_EnableEvents
End Sub

Private Function TipoDeHoja(ByVal ws As Worksheet) As String
    'Permite distinguir hoja normal, gráfica, etc. (por si algún día añades ChartSheets)
    On Error Resume Next
    If TypeName(ws) = "Worksheet" Then
        TipoDeHoja = "Worksheet"
    Else
        TipoDeHoja = TypeName(ws)
    End If
End Function


