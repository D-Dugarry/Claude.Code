Attribute VB_Name = "Rut_Ws_Shapes_Stratistics"
Option Explicit

Sub Rut_Wb_Shapes_Statistics_List()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsOut As Worksheet
    Dim sh As Shape
    Dim fila As Long
    
    Set wb = ThisWorkbook   ' o ActiveWorkbook, según prefieras
    
    ' Crear / reutilizar hoja de salida
    On Error Resume Next
    Set wsOut = wb.Worksheets("Wb_Shapes_Stratistics")
    On Error GoTo 0
    
    If wsOut Is Nothing Then
        Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsOut.Name = "Wb_Shapes_Stratistics"
    Else
        wsOut.Cells.Clear
    End If
    
    ' Encabezados
    fila = 1
    With wsOut
        .Cells(fila, 1).Value = "Hoja"
        .Cells(fila, 2).Value = "ShapeName"
        .Cells(fila, 3).Value = "Tipo"
        .Cells(fila, 4).Value = "Macro (OnAction)"
        .Cells(fila, 5).Value = "Visible"
        .Cells(fila, 6).Value = "Left"
        .Cells(fila, 7).Value = "Top"
        .Cells(fila, 8).Value = "Width"
        .Cells(fila, 9).Value = "Height"
        .Cells(fila, 10).Value = "Texto"
    End With
    
    ' Recorrer todas las hojas y sus shapes
    For Each ws In wb.Worksheets
        For Each sh In ws.Shapes
            fila = fila + 1
            With wsOut
                .Cells(fila, 1).Value = ws.Name
                .Cells(fila, 2).Value = sh.Name
                .Cells(fila, 3).Value = TipoShapeTexto(sh)
                
                ' OnAction puede dar error si no aplica
                On Error Resume Next
                .Cells(fila, 4).Value = sh.OnAction
                On Error GoTo 0
                
                On Error Resume Next
                .Cells(fila, 5).Value = IIf(sh.Visible, "Sí", "No")
                .Cells(fila, 6).Value = sh.Left
                .Cells(fila, 7).Value = sh.Top
                .Cells(fila, 8).Value = sh.Width
                .Cells(fila, 9).Value = sh.Height
                On Error GoTo 0
                ' Texto, si aplica
                On Error Resume Next
                .Cells(fila, 10).Value = sh.TextFrame2.TextRange.Text
                On Error GoTo 0
            End With
        Next sh
    Next ws
    
    wsOut.Columns.AutoFit
    wsOut.Activate
End Sub

Private Function TipoShapeTexto(ByVal sh As Shape) As String
    ' Devuelve un texto legible para el tipo de shape
    ' Mapea algunos MsoShapeType comunes; si no, devuelve el número.
    
    Select Case sh.Type   ' MsoShapeType [web:6][web:7]
        Case msoAutoShape:        TipoShapeTexto = "AutoShape"
        Case msoCallout:          TipoShapeTexto = "Callout"
        Case msoChart:            TipoShapeTexto = "Chart"
        Case msoComment:          TipoShapeTexto = "Comment"
        Case msoFreeform:         TipoShapeTexto = "Freeform"
        Case msoGroup:            TipoShapeTexto = "Group"
        Case msoEmbeddedOLEObject: TipoShapeTexto = "OLE Object"
        Case msoFormControl:      TipoShapeTexto = "FormControl"
        Case msoLine:             TipoShapeTexto = "Line"
        Case msoPicture:          TipoShapeTexto = "Picture"
        Case msoPlaceholder:      TipoShapeTexto = "Placeholder"
        Case msoTextBox:          TipoShapeTexto = "TextBox"
        Case msoTable:            TipoShapeTexto = "Table"
        Case Else:                TipoShapeTexto = "Tipo " & CStr(sh.Type)
    End Select
End Function

