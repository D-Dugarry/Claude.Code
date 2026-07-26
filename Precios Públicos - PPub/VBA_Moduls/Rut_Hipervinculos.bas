Attribute VB_Name = "Rut_Hipervinculos"
Option Explicit

' ==================================================================================================================================
Sub Rut_ListarHipervinculos()
    Dim Ws As Worksheet
    Dim Celda As Range
    Dim NewLine As Long
    Dim hojaResultado As Worksheet

    ' Crear una nueva hoja para mostrar los resultados
    If Fnc_WrkSheet_Exist("Hipervínculos") Then
        Call Rut_WrkSheet_Vaciar(Sheets("Hipervínculos"))
        Set hojaResultado = ThisWorkbook.Sheets("Hipervínculos")
    Else
        Set hojaResultado = ThisWorkbook.Sheets.Add
        hojaResultado.Name = "Hipervínculos"
    End If
    hojaResultado.Cells(1, 1).Value = "Hoja"
    hojaResultado.Cells(1, 2).Value = "Celda"
    hojaResultado.Cells(1, 3).Value = "Dirección"
    hojaResultado.Cells(1, 4).Value = "Texto de la celda"
    NewLine = 2

    ' Recorrer todas las hojas
    For Each Ws In ThisWorkbook.Worksheets
        Debug.Print "Sheet:  " & Ws.Name
        For Each Celda In Ws.UsedRange
            If Celda.Hyperlinks.Count > 0 Then
                hojaResultado.Cells(NewLine, 1).Value = Ws.Name
                hojaResultado.Cells(NewLine, 2).Value = Celda.Address
                hojaResultado.Cells(NewLine, 3).Value = Celda.Hyperlinks(1).Address
                hojaResultado.Cells(NewLine, 4).Value = Celda.Value
                NewLine = NewLine + 1
            End If
        Next Celda
    Next Ws
    
    If NewLine = 2 Then
        hojaResultado.Cells(NewLine + 1, 2).Value = "¡ Sin Hipervínculos !"
        MsgBox "Proceso completado.     ¡ Sin Hipervínculos !"
    Else
        MsgBox "Proceso completado. Los hipervínculos están listados en la hoja 'Hipervínculos'." & vbLf & vbLf & _
                "¡ Hemos encontrado " & NewLine - 2 & " hipervínculos !"
    End If

End Sub
' ==================================================================================================================================

