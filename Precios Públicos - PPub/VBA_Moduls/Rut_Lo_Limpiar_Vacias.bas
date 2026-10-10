Attribute VB_Name = "Rut_Lo_Limpiar_Vacias"
' Last Rev. 2026-10-10 23:19
Option Explicit
' Dependencias: Rut_Wb_State_Manager (Rut_Off_Functions / Rut_On_Functions), Rut_Ws_Protect_Status (Rut_Prot_Save / Rut_Prot_Restore,
' tipo T_Prot_Estado) y la constante DefC_FormatCol (M_000_Ini_Var_APP).

'###################################################################################################################################
' ==================================================================================================================================
Sub Rut_Lo_Limpiar_Formato_Vacias(Lo As ListObject, Optional LoDefCol As ListObject = Nothing)   '--- Quita el formato a las celdas VACIAS ---
' ==================================================================================================================================
' Una celda vacia SIN formato propio no se guarda en el fichero; con formato (aunque sea General) Excel la guarda una a una:
' en BDatos eran ~5,4 millones de celdas vacias. Esta rutina deja sin formato SOLO las celdas vacias (los datos no se tocan).
' Si se pasa LoDefCol, solo se limpian las Col. con FormatCol = FALSO (las VERDADERO conservan su formato por si se rellenan luego).
' Los tramos de celdas vacias seguidas se limpian juntos (hasta 200 tramos por llamada). Guardar el libro para que baje el tamano.
' ==================================================================================================================================
    Dim Ws As Worksheet:        Set Ws = Lo.Parent
    Dim Prot_Estado As T_Prot_Estado
    Dim C As Long, R As Long, RIni As Long, N As Long
    Dim V As Variant
    Dim Col As Range, Rng As Range, Tmp As Range
    Dim NAreas As Long, NVaciasCol As Long, NVacias As Long, NCols As Long

    If Lo.DataBodyRange Is Nothing Then Exit Sub
    Call Rut_Off_Functions
    Call Rut_Prot_Save(Ws, Prot_Estado)             '- anota sus permisos y la desprotege
    On Error GoTo Fin
    N = Lo.DataBodyRange.Rows.Count
    For C = 1 To Lo.ListColumns.Count
        If Not LoDefCol Is Nothing Then
            If C > LoDefCol.ListRows.Count Then GoTo NextCol
            If LoDefCol.DataBodyRange.Cells(C, DefC_FormatCol).Value Then GoTo NextCol      '- VERDADERO: Col. con formato, no se toca
        End If
        Set Col = Lo.ListColumns(C).DataBodyRange
        If N = 1 Then
            ReDim V(1 To 1, 1 To 1): V(1, 1) = Col.Value2
        Else
            V = Col.Value2
        End If
        Set Rng = Nothing: NAreas = 0: NVaciasCol = 0
        R = 1
        Do While R <= N
            If IsEmpty(V(R, 1)) Then
                RIni = R
                Do While R < N
                    If Not IsEmpty(V(R + 1, 1)) Then Exit Do
                    R = R + 1
                Loop
                Set Tmp = Col.Cells(RIni, 1).Resize(R - RIni + 1, 1)
                NVaciasCol = NVaciasCol + (R - RIni + 1)
                If Rng Is Nothing Then Set Rng = Tmp Else Set Rng = Union(Rng, Tmp)
                NAreas = NAreas + 1
                If NAreas >= 200 Then Rng.ClearFormats: Set Rng = Nothing: NAreas = 0
            End If
            R = R + 1
        Loop
        If Not Rng Is Nothing Then Rng.ClearFormats
        If NVaciasCol > 0 Then
            NCols = NCols + 1: NVacias = NVacias + NVaciasCol
            Debug.Print ">>> Rut_Lo_Limpiar_Formato_Vacias, Col. " & C & " (" & Lo.ListColumns(C).Name & "): " & Format(NVaciasCol, "#,##0") & " celdas vacias"
        End If
NextCol:
    Next C
    Debug.Print ">>> Rut_Lo_Limpiar_Formato_Vacias, " & Ws.Name & ": " & NCols & " Col. con celdas vacias, " & Format(NVacias, "#,##0") & " celdas limpiadas (guardar el libro para que baje el tamano)"

Fin:
    If Err.Number <> 0 Then Debug.Print ">>> Rut_Lo_Limpiar_Formato_Vacias, " & Ws.Name & " -> ERROR " & Err.Number & " en la Col. " & C & ": " & Err.Description
    On Error Resume Next
    Call Rut_Prot_Restore(Ws, Prot_Estado, True)    '- la reprotege con los mismos permisos; True = UserInterfaceOnly
    On Error GoTo 0
    Call Rut_On_Functions
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Sub Rut_BDatos_Limpiar_Formato_Vacias()     '--- BDatos: quita el formato a las celdas vacias de las Col. con FormatCol = FALSO ---
' ==================================================================================================================================
    Call Rut_Lo_Limpiar_Formato_Vacias(Sht__BD.ListObjects(1), Prog_DefCol_BD.ListObjects(1))
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Sub Rut_Ws_Quitar_Estilo_Columna(Ws As Worksheet, Lo As ListObject)   '--- Quita el estilo (fuente) de las COLUMNAS de la hoja ---
' ==================================================================================================================================
' Las columnas de la hoja pueden llevar una fuente propia (en BDatos, Courier New): como las celdas vacias SIN formato propio son
' distintas de la columna, Excel las guarda una a una (4,2 millones de celdas en BDatos). Con la fuente por defecto del libro en las
' columnas, esas celdas ya no se guardan. Prueba del 2026-10-10 sobre una copia: 56.000 KB -> 41.300 KB, mismos valores y formatos.
' Pone la fuente por defecto (la del estilo Normal) en TODA la hoja y despues devuelve su nombre y tamano a las celdas que NO estan
' en el cuerpo de la tabla (cabeceras, filas de encima, fila de totales...), que si tienen fuentes propias. Conserva negrita, cursiva
' y color. Si esas celdas son mas de 50.000, no hace nada. Se ejecuta UNA vez: el estilo de columna queda guardado en el libro.
' ==================================================================================================================================
    Const MAX_CELDAS As Long = 50000
    Dim Prot_Estado As T_Prot_Estado
    Dim Normal As Style:            Set Normal = Ws.Parent.Styles("Normal")
    Dim Ur As Range, Body As Range, Bloque As Variant, Cel As Range
    Dim R1 As Long, R2 As Long, C1 As Long, C2 As Long, B1 As Long, B2 As Long, BC1 As Long, BC2 As Long
    Dim Bloques As New Collection, Guardadas As New Collection
    Dim Vistas As Long, Aplicado As Boolean, V As Variant

    If Lo.DataBodyRange Is Nothing Then Exit Sub
    Call Rut_Off_Functions
    Call Rut_Prot_Save(Ws, Prot_Estado)             '- anota sus permisos y la desprotege
    On Error GoTo Fin

    '--- 1) Celdas fuera del cuerpo de la tabla: encima, debajo, a la izquierda y a la derecha ---
    Set Ur = Ws.UsedRange:          Set Body = Lo.DataBodyRange
    R1 = Ur.Row:                    R2 = Ur.Row + Ur.Rows.Count - 1
    C1 = Ur.Column:                 C2 = Ur.Column + Ur.Columns.Count - 1
    B1 = Body.Row:                  B2 = Body.Row + Body.Rows.Count - 1
    BC1 = Body.Column:              BC2 = Body.Column + Body.Columns.Count - 1
    If B1 > R1 Then Bloques.Add Ws.Range(Ws.Cells(R1, C1), Ws.Cells(B1 - 1, C2))
    If R2 > B2 Then Bloques.Add Ws.Range(Ws.Cells(B2 + 1, C1), Ws.Cells(R2, C2))
    If BC1 > C1 Then Bloques.Add Ws.Range(Ws.Cells(B1, C1), Ws.Cells(B2, BC1 - 1))
    If C2 > BC2 Then Bloques.Add Ws.Range(Ws.Cells(B1, BC2 + 1), Ws.Cells(B2, C2))
    For Each Bloque In Bloques
        Vistas = Vistas + Bloque.Cells.Count
    Next Bloque
    If Vistas > MAX_CELDAS Then
        Debug.Print ">>> Rut_Ws_Quitar_Estilo_Columna, " & Ws.Name & ": " & Format(Vistas, "#,##0") & " celdas fuera de la tabla (maximo " & MAX_CELDAS & "), no se hace nada"
        GoTo Fin
    End If

    '--- 2) Anota la fuente de las que no usan la de por defecto ---
    For Each Bloque In Bloques
        For Each Cel In Bloque.Cells
            If Cel.Font.Name <> Normal.Font.Name Or Cel.Font.Size <> Normal.Font.Size Then
                Guardadas.Add Array(Cel.Address(False, False), Cel.Font.Name, Cel.Font.Size)
            End If
        Next Cel
    Next Bloque

    '--- 3) Fuente por defecto en toda la hoja (cambia el estilo de columna) y se devuelve la suya a las celdas anotadas ---
    Aplicado = True
    Ws.Cells.Font.ThemeFont = xlThemeFontMinor          ' la fuente del tema (Calibri), igual que la del estilo Normal
    Ws.Cells.Font.Size = Normal.Font.Size
    For Each V In Guardadas
        Ws.Range(V(0)).Font.Name = V(1)
        Ws.Range(V(0)).Font.Size = V(2)
    Next V
    Debug.Print ">>> Rut_Ws_Quitar_Estilo_Columna, " & Ws.Name & ": fuente por defecto en la hoja; devuelta su fuente a " & Guardadas.Count & _
                " celdas de " & Format(Vistas, "#,##0") & " fuera de la tabla (guardar el libro para que baje el tamano)"

Fin:
    If Err.Number <> 0 Then
        Debug.Print ">>> Rut_Ws_Quitar_Estilo_Columna, " & Ws.Name & " -> ERROR " & Err.Number & ": " & Err.Description
        If Aplicado Then                                '- fallo a medias: se intenta devolver su fuente a las celdas anotadas
            Err.Clear
            On Error Resume Next
            For Each V In Guardadas
                Ws.Range(V(0)).Font.Name = V(1)
                Ws.Range(V(0)).Font.Size = V(2)
            Next V
        End If
    End If
    On Error Resume Next
    Call Rut_Prot_Restore(Ws, Prot_Estado, True)        '- la reprotege con los mismos permisos; True = UserInterfaceOnly
    On Error GoTo 0
    Call Rut_On_Functions
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Sub Rut_BDatos_Quitar_Estilo_Columna()      '--- BDatos: quita el estilo (fuente) de las columnas, para que no se guarden las celdas vacias ---
' ==================================================================================================================================
    Call Rut_Ws_Quitar_Estilo_Columna(Sht__BD, Sht__BD.ListObjects(1))
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
