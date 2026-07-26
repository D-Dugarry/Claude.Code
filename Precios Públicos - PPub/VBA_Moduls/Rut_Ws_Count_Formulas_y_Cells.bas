Attribute VB_Name = "Rut_Ws_Count_Formulas_y_Cells"
Option Explicit

Sub RevisarFormulasEnHojas()
    Dim Ws              As Worksheet
    Dim WsResumen       As Worksheet
    Dim UltimaFila      As Long
    Dim tieneFormula    As Boolean
    Dim Celda           As Range
    Dim Rng             As Range
    Dim libro           As Workbook
    Dim LstObj          As ListObject
    Dim numFormulas     As Long
    Dim ArchSelect          As Variant
    Dim BuscarEnActual      As VbMsgBoxResult
    Dim HojaSelect          As String
    Dim AnalizarTodo        As VbMsgBoxResult
    Dim WsSelect            As Worksheet
    Dim TiempoInicio    As Single
    Dim TiempoTotal     As Single
    
    ' Iniciar el temporizador
    TiempoInicio = Timer
    
    ' Preguntar si se quiere analizar el libro actual
    BuscarEnActual = MsgBox("¿Quieres revisar este libro (Workbook actual)?", vbYesNo + vbQuestion, "Seleccionar Libro")
    
    If BuscarEnActual = vbYes Then
        Set libro = ThisWorkbook
    Else
        ' Seleccionar el libro de Excel (incluye archivos con macros habilitadas .xlsm)
        ArchSelect = Application.GetOpenFilename(FileFilter:="Archivos de Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", Title:="Selecciona un libro")
        
        ' Si se cancela la selección, salir del procedimiento
        If ArchSelect = False Then
            MsgBox "Proceso cancelado.", vbExclamation
            Exit Sub
        End If
        
        ' Abrir el libro seleccionado
        Set libro = Application.Workbooks.Open(ArchSelect)
    End If
    
    ' Preguntar si se desea analizar todas las hojas o una específica
    AnalizarTodo = MsgBox("¿Quieres analizar todas las hojas?", vbYesNo + vbQuestion, "Seleccionar Hojas")
    
    If AnalizarTodo = vbNo Then
        ' Solicitar al usuario que introduzca el nombre de la hoja
        HojaSelect = InputBox("Introduce el nombre de la hoja que quieres revisar:", "Seleccionar Hoja")
        
        ' Verificar si la hoja existe
        On Error Resume Next
            Set WsSelect = libro.Sheets(HojaSelect)
        On Error GoTo 0
        
        If WsSelect Is Nothing Then
            MsgBox "La hoja '" & HojaSelect & "' no existe.", vbExclamation
            Exit Sub
        Else
            ' Buscar si la hoja seleccionada tiene fórmulas y cuántas
            tieneFormula = False
            numFormulas = 0
            For Each Celda In WsSelect.UsedRange
                If Celda.HasFormula Then
                    tieneFormula = True
                    numFormulas = numFormulas + 1
                End If
            Next Celda
            
            ' Mostrar el resultado en un mensaje
            If tieneFormula Then
                MsgBox "La hoja '" & HojaSelect & "' tiene " & numFormulas & " fórmulas.", vbInformation
            Else
                MsgBox "La hoja '" & HojaSelect & "' no contiene fórmulas.", vbInformation
            End If
            
            ' Mostrar tiempo transcurrido
            TiempoTotal = Timer - TiempoInicio
            MsgBox "El proceso tomó " & Format(TiempoTotal, "0.00") & " segundos en completarse.", vbInformation
            Exit Sub
        End If
    End If
    
    ' Si se selecciona revisar todas las hojas, revisar si existe la hoja de WsResumen
    ' Verificar si existe la hoja "Cuenta_Fórmulas", si no, crearla
    On Error Resume Next
        Set WsResumen = libro.Sheets("Cuenta_Fórmulas")
    On Error GoTo 0
    If WsResumen Is Nothing Then
        Set WsResumen = ThisWorkbook.Sheets.Add
        WsResumen.Name = "Cuenta_Fórmulas"
    Else
        ' Limpiar contenido previo en la hoja
        WsResumen.Activate
        WsResumen.Cells.Clear
    End If
    
    ' Encabezados de la tabla
    WsResumen.Cells(1, 1).Value = "Nombre de la Hoja"
    WsResumen.Cells(1, 2).Value = "Contiene Fórmulas"
    WsResumen.Cells(1, 3).Value = "Celdas con Fórmulas"
    WsResumen.Cells(1, 4).Value = "Celdas NO vacias"
    
    Application.ScreenUpdating = False
    
    ' Iniciar el conteo de filas para la tabla
    UltimaFila = 2
    
    ' Recorrer todas las hojas del libro seleccionado
    For Each Ws In libro.Worksheets
        tieneFormula = False
        numFormulas = 0
        
        ' Buscar si la hoja tiene alguna fórmula y contar cuántas celdas con fórmulas hay
        For Each Celda In Ws.UsedRange
            If Celda.HasFormula Then
                tieneFormula = True
                numFormulas = numFormulas + 1
            End If
        Next Celda
        
        ' Escribir los resultados en la hoja de WsResumen
        WsResumen.Cells(UltimaFila, 1).Value = Ws.Name
        WsResumen.Cells(UltimaFila, 2).Value = IIf(tieneFormula, "Sí", "No")
        WsResumen.Cells(UltimaFila, 3).Value = numFormulas
        WsResumen.Cells(UltimaFila, 4).Value = Application.CountA(Ws.Cells)
        
        ' Avanzar a la siguiente fila
        UltimaFila = UltimaFila + 1
    Next Ws
    
    ' Convertir el rango en una tabla (ListObject)
    Set Rng = WsResumen.Range("a1:d" & UltimaFila - 1)
    Set LstObj = WsResumen.ListObjects.Add(SourceType:=xlSrcRange, Source:=Rng, XlListObjectHasHeaders:=xlYes)
    LstObj.Name = "Lo_Num_Formulas"
    
    ' Ordenar la tabla por la columna de conteo de celdas en orden descendente
    LstObj.Sort.SortFields.Clear
    LstObj.Sort.SortFields.Add Key:=LstObj.ListColumns("Celdas con Fórmulas").DataBodyRange, _
                            Order:=xlDescending
    LstObj.Sort.Apply
    
'    LstObj.ListColumns(2).DataBodyRange.HorizontalAlignment = xlCenter
'    LstObj.ListColumns(3).DataBodyRange.HorizontalAlignment = xlCenter
    LstObj.ListColumns(2).DataBodyRange.Resize(, 3).HorizontalAlignment = xlCenter
    
    ' Ajustar el tamaño de las columnas para que se vean bien
    WsResumen.Columns("A:D").AutoFit
    WsResumen.Range("a1").Select
    
    ' Calcular el tiempo total transcurrido
    TiempoTotal = Timer - TiempoInicio
    
    Application.ScreenUpdating = True
    ' Mensaje final con el tiempo de ejecución
    MsgBox "Revisión completada en la hoja 'Cuenta de Fórmulas'. El proceso tomó " & Format(TiempoTotal, "0.00") & " segundos en completarse.", vbInformation
End Sub

