Attribute VB_Name = "Rut_Ws_Count_Cells"
Option Explicit

Sub Rut_WrkSht_Cells_Count()
    Dim Ws              As Worksheet
    Dim WsResumen       As Worksheet
    Dim LstObj          As ListObject
    Dim Rng             As Range
    Dim Fila            As Long
    Dim TotalCeldas     As Double
    
    ' Verificar si existe la hoja "Cuenta_Celdas", si no, crearla
    On Error Resume Next
    Set WsResumen = ThisWorkbook.Sheets("Cuenta_Celdas")
    On Error GoTo 0
    If WsResumen Is Nothing Then
        Set WsResumen = ThisWorkbook.Sheets.Add
        WsResumen.Name = "Cuenta_Celdas"
    Else
        ' Limpiar contenido previo en la hoja
        WsResumen.Activate
        WsResumen.Cells.Clear
    End If
    Application.ScreenUpdating = False
    
    ' Escribir encabezados en la hoja "Cuenta_Celdas"
    WsResumen.Cells(1, 1).Value = "Nombre de la Hoja"
    WsResumen.Cells(1, 2).Value = "Número de Celdas"
    
    ' Inicializar contador de filas
    Fila = 2
    
    ' Recorrer todas las hojas del libro
    For Each Ws In ThisWorkbook.Worksheets
        If Ws.Name <> WsResumen.Name Then ' Evitar contar la hoja "Cuenta_Celdas"
'            totalCeldas = ws.UsedRange.Cells.Count ' Contar todas las celdas de la hoja
            TotalCeldas = Application.CountA(Ws.Cells)
            ' Escribir el nombre de la hoja y el conteo en la hoja "Cuenta_Celdas"
            WsResumen.Cells(Fila, 1).Value = Ws.Name
            WsResumen.Cells(Fila, 2).Value = TotalCeldas
            
            Fila = Fila + 1
        End If
    Next Ws
    
    ' Convertir el rango en una tabla (ListObject)
    Set Rng = WsResumen.Range("A1:B" & Fila - 1)
    Set LstObj = WsResumen.ListObjects.Add(SourceType:=xlSrcRange, Source:=Rng, XlListObjectHasHeaders:=xlYes)
    LstObj.Name = "Lo_Num_Cells"
    
    ' Ordenar la tabla por la columna de conteo de celdas en orden descendente
    LstObj.Sort.SortFields.Clear
    LstObj.Sort.SortFields.Add Key:=LstObj.ListColumns("Número de Celdas").DataBodyRange, _
                            Order:=xlDescending
    LstObj.Sort.Apply
    
    ' Ajustar el ancho de las columnas para mejor visualización
    WsResumen.Columns("A:B").AutoFit
    WsResumen.Range("a1").Select
    
    Application.ScreenUpdating = True
    MsgBox "Proceso completado. Los resultados se encuentran en la hoja 'Cuenta_Celdas'.", vbInformation
'    Dim Ws As Worksheet
'    Dim ResumenHoja As Worksheet
'    Dim LstObj As ListObject
'    Dim Rng As Range
'    Dim Fila As Long
'    Dim Contador As Long
'
'    ' Verificar si existe la hoja "Cuenta_Celdas", si no, crearla
'    On Error Resume Next
'    Set ResumenHoja = ThisWorkbook.Sheets("Cuenta_Celdas")
'    If ResumenHoja Is Nothing Then
'        Set ResumenHoja = ThisWorkbook.Sheets.Add
'        ResumenHoja.Name = "Cuenta_Celdas"
'    Else
'        ' Limpiar contenido existente en la hoja "Cuenta_Celdas"
'        ResumenHoja.Cells.Clear
'    End If
'    On Error GoTo 0
'
'    ' Escribir encabezados en la hoja "Cuenta_Celdas"
'    ResumenHoja.Cells(1, 1).Value = "Nombre de la Hoja"
'    ResumenHoja.Cells(1, 2).Value = "Número de Celdas Utilizadas"
'
'    ' Inicializar la fila para escribir los resultados
'    Fila = 2
'
'    ' Recorrer cada hoja del libro
'    For Each Ws In ThisWorkbook.Sheets
'        ' Evitar procesar la hoja "Cuenta_Celdas" en el conteo
'        If Ws.Name <> "Cuenta_Celdas" Then
'            ' Contar el número de celdas no vacías en la hoja
'            Contador = Application.CountA(Ws.Cells)
'
'            ' Escribir el nombre de la hoja y el conteo en la hoja "Cuenta_Celdas"
'            ResumenHoja.Cells(Fila, 1).Value = Ws.Name
'            ResumenHoja.Cells(Fila, 2).Value = Contador
'
'            ' Incrementar la fila para el siguiente resultado
'            Fila = Fila + 1
'        End If
'    Next Ws
'
'    ' Ajustar el ancho de las columnas en la hoja "Cuenta_Celdas"
'    ResumenHoja.Columns("A:B").AutoFit
'
'    ' Mensaje de finalización
'    MsgBox "El conteo de celdas por hoja se ha completado. Los resultados están en la hoja 'Cuenta_Celdas'.", vbInformation
End Sub
