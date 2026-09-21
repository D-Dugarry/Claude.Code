Attribute VB_Name = "Rut_Hipervinculos"
' Last Rev. 2026-09-21 12:12
Option Explicit

' ==================================================================================================
Sub ListarHipervinculos()
    Dim ws As Worksheet
    Dim Celda As Range
    Dim contador As Long
    Dim hojaResultado As Worksheet

    ' Crear una nueva hoja para mostrar los resultados
    If Fnc_WrkSheet_Exist("Hipervínculos") Then
        Set hojaResultado = ThisWorkbook.Sheets("Hipervínculos")
        Call Rut_WrkSheet_Vaciar(hojaResultado)
    Else
        Set hojaResultado = ThisWorkbook.Sheets.Add
        hojaResultado.Name = "Hipervínculos"
    End If
    hojaResultado.Cells(1, 1).Value = "Hoja"
    hojaResultado.Cells(1, 2).Value = "Celda"
    hojaResultado.Cells(1, 3).Value = "Dirección"
    hojaResultado.Cells(1, 4).Value = "Texto de la celda"
    contador = 2

    ' Recorrer todas las hojas
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "Sheet:  " & ws.Name
        For Each Celda In ws.UsedRange
            If Celda.Hyperlinks.Count > 0 Then
                hojaResultado.Cells(contador, 1).Value = ws.Name
                hojaResultado.Cells(contador, 2).Value = Celda.Address
                hojaResultado.Cells(contador, 3).Value = Celda.Hyperlinks(1).Address
                hojaResultado.Cells(contador, 4).Value = Celda.Value
                contador = contador + 1
            End If
        Next Celda
    Next ws

    MsgBox "Proceso completado. Los hipervínculos están listados en la hoja 'Hipervínculos'."
End Sub
' ==================================================================================================

