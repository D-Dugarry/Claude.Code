Attribute VB_Name = "RuT_Wb_Sheets_List"
Option Explicit


' ==================================================================================================================================
    Sub RuT_WrkBook_Sheets_List()    ' Manejo interno, Hace la lista de todas las hojas de este libro       =============================
' ==================================================================================================================================
Dim WrkSht      As Worksheet
Dim NomHoja     As String
Dim NewRow      As ListRow
Application.ScreenUpdating = False
With Prog_HojasName
    .Unprotect
        ' --------------------------------=============  Preparar Tabla de DR_Unificada ==================
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        If .FilterMode Then .ShowAllData            ' Deshacer Filtros
        If Not .ListObjects(1).DataBodyRange Is Nothing Then .ListObjects(1).DataBodyRange.Delete
    With .ListObjects(1)
        If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
        For Each WrkSht In Worksheets
            Set NewRow = .ListRows.Add
            With NewRow
                .Range(1) = WrkSht.CodeName
                .Range(2) = WrkSht.Name
                .Range(3) = WrkSht.Visible
                Select Case .Range(3)
                    Case -1
                        .Range(4) = "Visible"
                    Case 0
                        .Range(4) = "Hidden"
                    Case 2
                        .Range(4) = "VeryHidden"
                End Select
            End With ' NewRow
        Next
        .Range.Sort key1:=.ListColumns(1), order1:=xlAscending, _
            Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    End With    ' .ListObjects(1)
'    Call Rut_Lo_DataBodyRange_Copy(Prog_HojasName.ListObjects(1), Prog_HojasName.ListObjects(1), True)
    .Protect
End With ' Prog_HojasName
Application.ScreenUpdating = True
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<



