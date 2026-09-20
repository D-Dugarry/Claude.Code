Attribute VB_Name = "M0999_Modif_Cols_BDatos"
Option Explicit
' Last Rev. 2026-09-20 23:20
Sub Reorganizar_Lo_BDatos_Cambio_Ene_26()
    Dim Lo As ListObject
    Set Lo = ActiveWorkbook.Worksheets("BDatos").ListObjects(1)
    '--- 1) Insertar 1 columnas VACÍAS a la derecha de "Col_46"
    Lo.ListColumns.Add Position:=46
    '--- 1) Insertar 1 columnas VACÍAS a la derecha de "Col_45"
    Lo.ListColumns.Add Position:=45
End Sub


Sub Reorganizar_Lo_BDatos_Cambio_Nov_2025()
    Dim Lo As ListObject
    Set Lo = ActiveWorkbook.Worksheets("BD_Hist").ListObjects(1)
    '--- 1) Insertar 2 columnas VACÍAS a la izquierda de "Col_33"
    Lo.ListColumns.Add Position:=32
    Lo.ListColumns.Add Position:=32
    Range("AD1:AG1").Copy Destination:=Range("AD4")
    '--- 2) Mover col 35 ANTES de col 34
    Lo.ListColumns.Add Position:=34
    Lo.ListColumns(35 + 1).Range.Copy Destination:=Lo.ListColumns(34).Range
    Lo.ListColumns(35 + 1).Delete
    Range("Ah1:Ai1").Copy Destination:=Range("Ah4")
    '--- 3) Mover col 45 ANTES de col 36
    Lo.ListColumns.Add Position:=36
    Lo.ListColumns(45 + 1).Range.Copy Destination:=Lo.ListColumns(36).Range
    Lo.ListColumns(45 + 1).Delete
    '--- 4) Mover col 46-47 ANTES de col 36
    Lo.ListColumns.Add Position:=36
    Lo.ListColumns.Add Position:=36
    Lo.ListColumns(46 + 2).Range.Copy Destination:=Lo.ListColumns(36).Range
    Lo.ListColumns(46 + 2).Delete
    Lo.ListColumns(46 + 2).Range.Copy Destination:=Lo.ListColumns(37).Range
    Lo.ListColumns(46 + 2).Delete
    Range("Aj1:Au1").Copy Destination:=Range("Aj4")
    '--- 5) Mover col 50-51 ANTES de col 48
    Lo.ListColumns.Add Position:=48
    Lo.ListColumns.Add Position:=48
    Lo.ListColumns(50 + 2).Range.Resize(, 2).Copy Destination:=Lo.ListColumns(48).Range
    Lo.ListColumns(50 + 2).Delete
    Lo.ListColumns(50 + 2).Delete
    Range("Av1:Ay1").Copy Destination:=Range("Av4")
End Sub



Sub Reorganizar_ListObject_Solo()
    Dim Lo As ListObject
    Dim iCol33 As Long, iCol36 As Long, iCol40 As Long
    Dim iCol45 As Long, iCol47 As Long, iCol50 As Long
    Dim iCol55 As Long, iCol57 As Long
    
    Set Lo = Worksheets("Hoja1").ListObjects(1)
    
    '--- Obtener índices iniciales
    iCol33 = Lo.ListColumns("Col_33").Index
    iCol36 = Lo.ListColumns("Col_36").Index
    iCol40 = Lo.ListColumns("Col_40").Index
    iCol45 = Lo.ListColumns("Col_45").Index
    iCol47 = Lo.ListColumns("Col_47").Index
    iCol50 = Lo.ListColumns("Col_50").Index
    iCol55 = Lo.ListColumns("Col_55").Index
    iCol57 = Lo.ListColumns("Col_57").Index
    
    '--- 1) Insertar 2 columnas VACÍAS a la izquierda de "Col_33"
    Lo.ListColumns.Add(Position:=iCol33).Name = "Nueva_Col1"
    Lo.ListColumns.Add(Position:=iCol33).Name = "Nueva_Col2"
    
    ' Recalcular índices (se movieron todos los de la derecha)
    iCol36 = Lo.ListColumns("Col_36").Index + 2
    iCol40 = Lo.ListColumns("Col_40").Index + 2
    iCol45 = Lo.ListColumns("Col_45").Index + 2
    iCol47 = Lo.ListColumns("Col_47").Index + 2
    iCol50 = Lo.ListColumns("Col_50").Index + 2
    iCol55 = Lo.ListColumns("Col_55").Index + 2
    iCol57 = Lo.ListColumns("Col_57").Index + 2
    
    '--- 2) Mover cols 45-47 ANTES de col 36 (copiar/pegar datos)
    Dim rngMover As Range, rngDestino As Range
    Set rngMover = Lo.ListColumns("Col_45").DataBodyRange.Resize(, 3)
    Lo.ListColumns("Col_45").Delete     ' Eliminar originales
    Lo.ListColumns("Col_46").Delete
    Lo.ListColumns("Col_47").Delete
    Lo.ListColumns.Add(Position:=iCol36 - 2).Name = "Col_45"
    Lo.ListColumns.Add(Position:=iCol36 - 1).Name = "Col_46"
    Lo.ListColumns.Add(Position:=iCol36).Name = "Col_47"
    ' Copiar datos de backup o regenerarlos
    
    '--- 3) Mover col 50 ANTES de col 40
    Lo.ListColumns("Col_50").DataBodyRange.Copy
    Lo.ListColumns("Col_50").Delete
    Lo.ListColumns.Add(Position:=iCol40 - 1).Name = "Col_50"
    Lo.ListColumns("Col_50").DataBodyRange.PasteSpecial xlPasteValues
    
    '--- 4) Mover cols 55-57 AL FINAL (copiar y pegar cada columna de una en una,
    '       igual que el patron del paso 3, para no perder el portapapeles previo)
    Lo.ListColumns("Col_55").DataBodyRange.Copy
    Lo.ListColumns("Col_55").Delete
    Lo.ListColumns.Add.Name = "Col_55"
    Lo.ListColumns("Col_55").DataBodyRange.PasteSpecial xlPasteValues
    Lo.ListColumns("Col_56").DataBodyRange.Copy
    Lo.ListColumns("Col_56").Delete
    Lo.ListColumns.Add.Name = "Col_56"
    Lo.ListColumns("Col_56").DataBodyRange.PasteSpecial xlPasteValues
    Lo.ListColumns("Col_57").DataBodyRange.Copy
    Lo.ListColumns("Col_57").Delete
    Lo.ListColumns.Add.Name = "Col_57"
    Lo.ListColumns("Col_57").DataBodyRange.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub

