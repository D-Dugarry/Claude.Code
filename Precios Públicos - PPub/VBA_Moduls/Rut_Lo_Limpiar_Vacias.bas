Attribute VB_Name = "Rut_Lo_Limpiar_Vacias"
' Last Rev. 2026-10-12 09:46
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
Sub Rut_Ws_Quitar_Estilo_Columna(Ws As Worksheet)   '--- Devuelve el estilo Normal a las COLUMNAS de la hoja, sin cambiar sus celdas ---
' ==================================================================================================================================
' Una celda vacia se guarda en el fichero si su formato es distinto del de su COLUMNA. Las columnas de BDatos llevaban un estilo
' propio (Courier New hasta el 2026-10-10; despues, la fuente del tema, que a la vista es igual que Normal pero para Excel es otro
' estilo), y ClearFormats (lo hacen la importacion y el formateo en cada proceso) deja las celdas en Normal: al ser distintas de su
' columna, Excel guardaba una a una ~5 millones de celdas vacias. Por eso la version anterior de esta rutina (poner la fuente del
' tema en toda la hoja) bajo el libro a 41 MB y los procesos de despues lo volvieron a subir (66 MB el 2026-10-12).
' Ahora: copia el formato de las celdas en uso a una hoja temporal (por tramos de filas), quita el formato de TODA la hoja (asi las
' columnas quedan en Normal) y devuelve a las celdas su formato. Probado el 2026-10-12 sobre una copia del libro: 66,5 -> 52,8 MB,
' y las 160.543 x 69 celdas de BDatos con el mismo valor y el mismo formato (numero, fuente, relleno, bordes, alineacion y
' proteccion), los mismos anchos y altos y el mismo formato condicional. Se ejecuta UNA vez por hoja: el estilo de columna queda
' guardado en el libro. Si falla a medias, NO borra la hoja temporal (tiene el formato de las celdas): cerrar el libro SIN guardar.
' ==================================================================================================================================
    Const FILAS_TRAMO As Long = 20000
    Dim Prot_Estado As T_Prot_Estado
    Dim Wb As Workbook:             Set Wb = Ws.Parent
    Dim Hoja_Activa As Object:      Set Hoja_Activa = Wb.ActiveSheet
    Dim Alertas As Boolean:         Alertas = Application.DisplayAlerts
    Dim Tmp As Worksheet, Ur As Range
    Dim NFil As Long, NCol As Long, R As Long, R2 As Long
    Dim Sel As String, Scr_Fila As Long, Scr_Col As Long, Vista As Boolean
    Dim Paso As String, Aplicado As Boolean

    Call Rut_Off_Functions
    Call Rut_Prot_Save(Ws, Prot_Estado)             '- anota sus permisos y la desprotege
    On Error GoTo Fin

    Paso = "medir la hoja"
    Set Ur = Ws.UsedRange                           '- se trabaja desde A1 hasta la ultima celda en uso
    NFil = Ur.Row + Ur.Rows.Count - 1
    NCol = Ur.Column + Ur.Columns.Count - 1
    If Ws.Visible = xlSheetVisible Then             '- PasteSpecial cambia la seleccion de la hoja: se anota para dejarla igual
        Ws.Activate
        If TypeName(Selection) = "Range" Then Sel = Selection.Address
        Scr_Fila = ActiveWindow.ScrollRow:  Scr_Col = ActiveWindow.ScrollColumn
        Vista = True
    End If

    Paso = "copiar el formato a la hoja temporal"
    Set Tmp = Wb.Worksheets.Add(After:=Wb.Sheets(Wb.Sheets.Count))
    For R = 1 To NFil Step FILAS_TRAMO
        R2 = Application.Min(R + FILAS_TRAMO - 1, NFil)
        Ws.Range(Ws.Cells(R, 1), Ws.Cells(R2, NCol)).Copy
        Tmp.Cells(R, 1).PasteSpecial xlPasteFormats
    Next R
    Application.CutCopyMode = False

    Paso = "quitar el formato de la hoja"
    Aplicado = True
    Ws.Cells.ClearFormats                           '- columnas y filas en Normal, celdas sin formato

    Paso = "devolver el formato a las celdas"
    For R = 1 To NFil Step FILAS_TRAMO
        R2 = Application.Min(R + FILAS_TRAMO - 1, NFil)
        Tmp.Range(Tmp.Cells(R, 1), Tmp.Cells(R2, NCol)).Copy
        Ws.Cells(R, 1).PasteSpecial xlPasteFormats
    Next R
    Application.CutCopyMode = False

    Paso = "borrar la hoja temporal"
    Application.DisplayAlerts = False
    Tmp.Delete
    Application.DisplayAlerts = Alertas
    Aplicado = False
    Debug.Print ">>> Rut_Ws_Quitar_Estilo_Columna, " & Ws.Name & ": estilo Normal en las columnas; formato devuelto a " & _
                Format(NFil, "#,##0") & " filas x " & NCol & " col. (guardar el libro para que baje el tamano)"

Fin:
    If Err.Number <> 0 Then
        Debug.Print ">>> Rut_Ws_Quitar_Estilo_Columna, " & Ws.Name & " -> ERROR " & Err.Number & " al " & Paso & ": " & Err.Description
        Application.CutCopyMode = False
        Application.DisplayAlerts = Alertas
        If Aplicado Then
            MsgBox "Error al " & Paso & ": " & Err.Description & vbLf & vbLf & _
                   "La hoja " & Ws.Name & " puede haberse quedado sin formato: cierra el libro SIN GUARDAR." & vbLf & _
                   "(El formato de sus celdas esta en la hoja " & Tmp.Name & ".)", vbCritical, "Quitar el estilo de columna"
        Else
            MsgBox "Error al " & Paso & ": " & Err.Description & vbLf & vbLf & _
                   "La hoja " & Ws.Name & " no se ha cambiado.", vbExclamation, "Quitar el estilo de columna"
        End If
    End If
    On Error GoTo -1                                '- sale del manejador de errores, para que el Resume Next de abajo valga
    On Error Resume Next
    If Vista And Not Aplicado Then                  '- la seleccion y el desplazamiento de antes
        Ws.Activate
        If Sel <> "" Then Ws.Range(Sel).Select
        ActiveWindow.ScrollRow = Scr_Fila:  ActiveWindow.ScrollColumn = Scr_Col
    End If
    If Not Hoja_Activa Is Nothing Then Hoja_Activa.Activate
    Call Rut_Prot_Restore(Ws, Prot_Estado, True)    '- la reprotege con los mismos permisos; True = UserInterfaceOnly
    On Error GoTo 0
    Call Rut_On_Functions
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
Sub Rut_BDatos_Quitar_Estilo_Columna()      '--- BDatos: estilo Normal en las columnas, para que no se guarden las celdas vacias ---
' ==================================================================================================================================
    Call Rut_Ws_Quitar_Estilo_Columna(Sht__BD)
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
