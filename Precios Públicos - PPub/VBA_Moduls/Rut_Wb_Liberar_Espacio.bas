Attribute VB_Name = "Rut_Wb_Liberar_Espacio"
' Last Rev. 2026-10-10 08:11
Option Explicit
' Dependencias: Rut_Wb_State_Manager (Rut_Off_Functions / Rut_On_Functions) y Rut_Ws_Protect_Status (Rut_Prot_Save / Rut_Prot_Restore /
' Fnc_Prot_Hoja_Protegida, tipo T_Prot_Estado).

'##################################################################################################################################
' ==================================================================================================================================
Sub Rut_WrkSheet_LiberarEspacio(WrkSht As Worksheet, Optional ByVal Password As String = "", Optional ByRef Resultado As Long)   '--- Restablece el rango de celdas en uso ---
' ==================================================================================================================================
' Valido para CUALQUIER hoja (con o sin ListObjects). Borra las filas y columnas vacias que quedan por debajo y a la derecha del
' ULTIMO CONTENIDO REAL, para que no inflen el tamano del fichero, y lee .UsedRange para que Excel restablezca el rango en uso.
' Ultimo contenido = lo mas lejano entre: celdas con valor o formula (celdas combinadas completas), notas, tablas (ListObjects),
' tablas dinamicas y objetos (Shapes, graficos, botones). NO toca filtros ni filas/columnas ocultas. Deja la proteccion como estaba.
' OJO: lo que haya SOLO con formato (bordes, relleno) mas alla de ese ultimo contenido se pierde, porque se borran filas/columnas enteras.
' El tamano del fichero solo baja al GUARDAR el libro.
' Resultado (opcional, por referencia): 0 = bien, 1 = error en la hoja, 2 = omitida (protegida con contrasena).
' ==================================================================================================================================
    Dim Prot_Estado As T_Prot_Estado        '--- para dejar la hoja con los mismos permisos de proteccion (Rut_Ws_Protect_Status) ---
    Dim UltFila As Long, UltCol As Long, UsedFila As Long, UsedCol As Long
    Dim Rng As Range, Lo As ListObject, Pt As PivotTable, Shp As Shape, Cmt As Comment
    Dim UsedAntes As String, UsedDespues As String, Paso As String

    Resultado = 0
    Call Rut_Off_Functions
    Call Rut_Prot_Save(WrkSht, Prot_Estado, Password)       '- anota sus permisos y la desprotege
    On Error GoTo Fin

    If Fnc_Prot_Hoja_Protegida(WrkSht) Then             ' no se ha podido desproteger (contrasena): se deja tal cual
        Debug.Print ">>> Rut_WrkSheet_LiberarEspacio, WrkSht=" & WrkSht.Name & " -> PROTEGIDA con contrasena, se omite"
        Resultado = 2
        GoTo Fin
    End If

    Paso = "UsedRange inicial"
    UsedAntes = WrkSht.UsedRange.Address(False, False)
    UltFila = 1: UltCol = 1
    With WrkSht
        Paso = "Find por filas"
        '--- 1) Celdas con valor o formula: ultima fila y ultima columna con contenido ---
        Set Rng = .Cells.Find(What:="*", After:=.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not Rng Is Nothing Then Call Rut_Ws_Abarcar_Rango(Rng, UltFila, UltCol)
        Paso = "Find por columnas"
        Set Rng = .Cells.Find(What:="*", After:=.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
        If Not Rng Is Nothing Then Call Rut_Ws_Abarcar_Rango(Rng, UltFila, UltCol)
        '--- 2) Tablas (aunque tengan celdas vacias), tablas dinamicas, notas y objetos ---
        Paso = "Tablas"
        For Each Lo In .ListObjects
            Call Rut_Ws_Abarcar_Rango(Lo.Range, UltFila, UltCol)
        Next Lo
        Paso = "Tablas dinamicas"
        For Each Pt In .PivotTables
            Call Rut_Ws_Abarcar_Rango(Pt.TableRange2, UltFila, UltCol)
        Next Pt
        Paso = "Notas"
        For Each Cmt In .Comments
            Call Rut_Ws_Abarcar_Rango(Cmt.Parent, UltFila, UltCol)
        Next Cmt
        Paso = "Objetos (Shapes)"
        For Each Shp In .Shapes
            Set Rng = Nothing
            On Error Resume Next
            Set Rng = Shp.BottomRightCell
            On Error GoTo Fin
            Err.Clear                                       ' un BottomRightCell fallido no debe contar como error de la hoja
            If Not Rng Is Nothing Then Call Rut_Ws_Abarcar_Rango(Rng, UltFila, UltCol)
        Next Shp
        '--- 3) Borra lo de abajo y lo de la derecha ---
        '    Solo se borra si el rango en uso llega mas lejos que el contenido real (si ya coincide, no hay nada que liberar)
        With .UsedRange
            UsedFila = .Row + .Rows.Count - 1: UsedCol = .Column + .Columns.Count - 1
        End With
        If UsedFila > UltFila Then Paso = "Borrar filas " & (UltFila + 1) & ":" & UsedFila: .Range(.Rows(UltFila + 1), .Rows(UsedFila)).Delete
        If UsedCol > UltCol Then Paso = "Borrar columnas " & (UltCol + 1) & ":" & UsedCol: .Range(.Columns(UltCol + 1), .Columns(UsedCol)).Delete
        Paso = "UsedRange final"
        UsedDespues = .UsedRange.Address(False, False)  ' Restablece el rango de celdas en uso (hay que LEER la propiedad)
    End With
    Debug.Print ">>> Rut_WrkSheet_LiberarEspacio, WrkSht=" & WrkSht.Name & ", UsedRange: " & UsedAntes & " -> " & UsedDespues

Fin:
    If Err.Number <> 0 Then Resultado = 1: Debug.Print ">>> Rut_WrkSheet_LiberarEspacio, WrkSht=" & WrkSht.Name & " -> ERROR " & Err.Number & " en el paso '" & Paso & "' (UltFila=" & UltFila & ", UltCol=" & UltCol & "): " & Err.Description
    On Error Resume Next
    Call Rut_Prot_Restore(WrkSht, Prot_Estado, True)        '- la reprotege con los mismos permisos; True = UserInterfaceOnly
    On Error GoTo 0
    Call Rut_On_Functions
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Private Sub Rut_Ws_Abarcar_Rango(ByVal Rng As Range, ByRef UltFila As Long, ByRef UltCol As Long)   '--- Amplia UltFila/UltCol para incluir Rng ---
' ==================================================================================================================================
    Dim Fila As Long, Col As Long
    Dim Area As Range
    For Each Area In Rng.Areas
        Fila = Area.Row + Area.Rows.Count - 1
        Col = Area.Column + Area.Columns.Count - 1
        With Area.Cells(Area.Rows.Count, Area.Columns.Count).MergeArea   ' MergeArea solo es fiable con UNA celda (en un rango de varias da 1004)
            If .Row + .Rows.Count - 1 > Fila Then Fila = .Row + .Rows.Count - 1
            If .Column + .Columns.Count - 1 > Col Then Col = .Column + .Columns.Count - 1
        End With
        If Fila > UltFila Then UltFila = Fila
        If Col > UltCol Then UltCol = Col
    Next Area
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Sub Rut_WrkBook_LiberarEspacio(Optional ByVal Password As String = "")   '--- Restablece el rango de celdas de TODAS las hojas de ThisWorkbook ---
' ==================================================================================================================================
' Recorre todas las hojas de calculo de ThisWorkbook (tambien las ocultas) y llama a Rut_WrkSheet_LiberarEspacio. Guardar el libro
' despues para que el tamano del fichero baje. Cada hoja deja su traza (UsedRange antes -> despues) en la ventana Inmediato.
' ==================================================================================================================================
    Dim Ws As Worksheet
    Dim N As Long, Res As Long, NErr As Long, NOmit As Long
    Call Rut_Off_Functions
    For Each Ws In ThisWorkbook.Worksheets
        Call Rut_WrkSheet_LiberarEspacio(Ws, Password, Res)
        N = N + 1
        If Res = 1 Then NErr = NErr + 1
        If Res = 2 Then NOmit = NOmit + 1
    Next Ws
    Call Rut_On_Functions
    Debug.Print ">>> Rut_WrkBook_LiberarEspacio, hojas revisadas = " & N & ", con ERROR = " & NErr & ", omitidas (protegidas con contrasena) = " & NOmit & IIf(NErr > 0, "  <<< REVISAR las trazas de arriba", "") & " (guardar el libro para que baje el tamano)"
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
