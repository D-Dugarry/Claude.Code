Attribute VB_Name = "Rut_Ws"
' Last Rev. 2026-10-02 13:55
Option Explicit

'###################################################################################################################################
Sub Rut_WrkSheet_LstObj_LiberarEspacio(WrkSht As Worksheet)   '--- Borra TODO a la Derecha y Abajo de .ListObjects(1) ----------------------
' ==================================================================================================================================
    With WrkSht
            .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
            .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
            If .FilterMode Then .ShowAllData            ' Deshacer Filtros
'            If .ListObjects(1).FilterMode Then .ListObjects(1).ShowAllData            ' Deshacer Filtros
'            Rut_Lo_Filtros_Quitar (WrkSht.ListObjects(1))
        Dim UltCelda As Range
        With .ListObjects(1).Range                      ' Borra la filas de abajo y columnas de la derecha del la Tabla .ListObjects(1)
            Set UltCelda = .Cells(.Rows.Count, .Columns.Count)   ' Última celda de la tabla (sin Select, sin depender de la hoja activa)
        End With
            WrkSht.Range(WrkSht.Cells(UltCelda.Row + 1, UltCelda.Column + 1), WrkSht.Cells(WrkSht.Rows.Count, 1)).EntireRow.Delete
            WrkSht.Range(WrkSht.Cells(UltCelda.Row + 1, UltCelda.Column + 1), WrkSht.Cells(1, WrkSht.Columns.Count)).EntireColumn.Delete
            Dim Dummy As String: Dummy = WrkSht.UsedRange.Address   ' Para restablecer el rango de celdas en uso (hay que LEER la propiedad)
    End With
'    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
'            .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
'            .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
'            If .FilterMode Then .ShowAllData            ' Deshacer Filtros
'        With .ListObjects(1).Range                      ' Borra la filas de abajo y columnas de la derecha del la Tabla .ListObjects(1)
'            Range(.Cells(.Rows.Count, .Columns.Count).Address).Select   ' Selecciona la última celda de la tabla
'        End With
'            ActiveCell.Offset(1, 1).Select
'            Sheets(WrkSht).Range(ActiveCell.Address & ":" & Cells(Rows.Count, 1).Address).EntireRow.Delete
'            Sheets(WrkSht).Range(ActiveCell.Address & ":" & Cells(1, Columns.Count).Address).EntireColumn.Delete
'            ActiveSheet.UsedRange                       ' Para restablecer el rango de celdas en uso
'    End With
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Sub Rut_WrkSheet_Vaciar(WrkSht As Worksheet)      '--- Borra Toda la Hoja incluso los objetos (Shapes)  -------------------------------
' ==================================================================================================================================
Debug.Print ">>> Rut_WrkSheet_Vaciar, WrkSht=" & WrkSht.Name

    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht.Name)
            Dim Visual_Status   As Variant:  Visual_Status = .Visible   '--- para dejar la hoja en el mismo estado de Visibilidad ---
            Dim Prot_Estado     As T_Prot_Estado   '--- para dejar la hoja en el mismo estado de protección, con los mismos permisos (Rut_Ws_Protect_Status) ---
        .Visible = xlHidden
        Call Rut_Prot_Save(WrkSht, Prot_Estado)             '- anota sus permisos y la desprotege
            .Columns.Delete     ' --- con esto se borran hasta los "Shapes"
            Dim Dummy As String: Dummy = .UsedRange.Address   ' Restablece el rango de celdas en uso (hay que LEER la propiedad)
        .Visible = Visual_Status
        Call Rut_Prot_Restore(WrkSht, Prot_Estado, True)    '- la reprotege con los mismos permisos; True = UserInterfaceOnly
        .Select
    End With
Debug.Print ">>> Rut_WrkSheet_Vaciar, WrkSht=" & WrkSht.Name, "Protect_Status = " & Prot_Estado.Protegida
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Function Fnc_WrkSheet_Exist(SheetName As String) As Boolean
    Dim Ws As Worksheet
    On Error Resume Next
    Set Ws = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    Fnc_WrkSheet_Exist = Not Ws Is Nothing
End Function
' ==================================================================================================================================
Function Fnc_Range_Exist(RngName As String) As Boolean
    Dim Rng As Range
    On Error Resume Next
    Set Rng = Range(RngName)
    On Error GoTo 0
    Fnc_Range_Exist = Not Rng Is Nothing
End Function
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
' ==================================================================================================================================
' =====================     Rut_Columnas_ShowHide_RowX1     ================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
Sub Rut_WrkSheet_Col_ShowHide_RowX1(Show As Boolean)    ' Sirve para cualquier Hoja y para la Tabla nº.1
Dim Cont_Col    As Integer
    If Not Show Then                    ' Mostrar Columnas
        Columns.EntireColumn.Hidden = False
    Else                                        ' Ocultar Columnas
        For Cont_Col = 1 To ActiveSheet.ListObjects(1).Range.Columns.Count
            If IsEmpty(Cells(1, Cont_Col)) And Not Columns(Cont_Col).Hidden Then Columns(Cont_Col).Hidden = True
        Next Cont_Col
    End If
End Sub     ' Rut_Columnas_ShowHide_RowX1     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================



