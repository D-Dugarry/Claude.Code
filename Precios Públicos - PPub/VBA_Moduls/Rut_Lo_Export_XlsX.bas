Attribute VB_Name = "Rut_Lo_Export_XlsX"
' Last Rev. 2026-09-30 18:25
'2025-12-20
Option Explicit


'- ----------------------------------------------------------------------------------------------------------------------------
'- Exportar ListObject a New_WB
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Export_to_New_WB(ByVal Lo_Data As ListObject, _
                            Optional FichName As String)
    Dim WbNew       As Workbook
    Dim WsNew       As Worksheet
    Dim SaveNomArch       As Variant
    Dim RutaINI   As String
    Dim NomArch   As String
    Dim ArchFullName   As String
    Dim calcMode    As XlCalculation
    
    ' Proponer ruta = del libro actual
    If Len(ThisWorkbook.Path) > 0 Then
        RutaINI = ThisWorkbook.Path & Application.PathSeparator
    Else
        RutaINI = CurDir$ & Application.PathSeparator
    End If
    
    ' Proponer nombre = nombre de la tabla + .xlsx
    NomArch = Lo_Data.Name & ".xlsx"
    
    ' Cuadro Guardar como, con ruta y nombre sugeridos
    SaveNomArch = Application.GetSaveAsFilename( _
                    InitialFileName:=RutaINI & NomArch, _
                    FileFilter:="Excel (*.xlsx), *.xlsx")
    If SaveNomArch = False Then Exit Sub          ' Usuario cancela
    
    ArchFullName = CStr(SaveNomArch)
    
    ' Optimizar
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        calcMode = .Calculation
        .Calculation = xlCalculationManual
    End With
    
    On Error GoTo ErrHandler
    
    ' Crear nuevo libro y copiar solo la tabla
    Set WbNew = Workbooks.Add(xlWBATWorksheet)   ' Libro con 1 hoja
    Set WsNew = WbNew.Worksheets(1)
    
    ' Copiar rango completo de la tabla (incluye cabecera y datos)
    Lo_Data.Range.Copy Destination:=WsNew.Range("A1")
    
    ' Opcional: ajustar ancho de columnas
    WsNew.Columns.AutoFit
    
    ' Guardar como .xlsx
    WbNew.SaveAs Filename:=ArchFullName, FileFormat:=xlOpenXMLWorkbook
    Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(ArchFullName, FichName)
    
Fin:
    On Error Resume Next
    ' Cerrar libro nuevo (si quieres dejarlo abierto, comenta la siguiente línea)
    WbNew.Close SaveChanges:=False
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calcMode
    End With
    Exit Sub

ErrHandler:
    Resume Fin
End Sub





