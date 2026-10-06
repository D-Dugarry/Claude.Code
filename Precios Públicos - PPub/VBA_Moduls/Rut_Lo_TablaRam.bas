Attribute VB_Name = "Rut_Lo_TablaRam"
' Last Rev. 2026-10-06 11:50
Option Explicit

'===================================================================================================
' Rut_Lo_TablaRam
'
' Copia en memoria (RAM) de las columnas de un ListObject, para recorrerlas en VBA sin hacer una
' llamada a Excel por cada celda. Leer o escribir celda a celda (.Cells(i, j), ListRows(i).Range)
' cuesta una llamada a Excel cada vez: con ~175.000 filas son decenas de segundos. Con el array
' son décimas.
'
' Uso:
'   Dim T As T_TablaRam
'   Call Rut_TablaRam_Cargar(T, Lo, Array(BD_Ref, BD_Plan))   ' sin lista = todas las columnas
'   ... leer y escribir T.Datos(Fila, BD_xxx) ...               ' mismos índices que la tabla (BD_*)
'   T.Modificada(BD_Plan) = True                                ' marcar lo que hay que devolver
'   Call Rut_TablaRam_Volcar(T, Lo)                             ' escribe SOLO las columnas marcadas
'
' REGLAS:
'   - Entre Cargar y Volcar NO se puede ordenar la tabla ni añadir/borrar filas: T.Datos va por
'     posición de fila. Volcar comprueba al menos que el nº de filas no ha cambiado.
'   - Volcar reescribe la columna ENTERA, también las filas que no se tocaron. Un texto que parezca
'     número o fecha ("00123", "1/2") Excel lo convertiría al escribirlo: marcar como Modificada
'     solo columnas de importes o de textos que no puedan confundirse con números o fechas. En BDatos,
'     DNI, Form. Pago y CTA. CCC son textos con cifras: para vaciar celdas de esas columnas, usar
'     Rut_TablaRam_Vaciar_Celdas.
'   - .Value (por defecto) devuelve las fechas como Date y los importes con formato moneda como
'     Currency (redondeado a 4 decimales), igual que leer .Cells(i, j): úsalo cuando se monten
'     textos o comparaciones con esos valores. .Value2 (LeerValue2:=True) los deja como Double sin
'     redondear: úsalo en las columnas que se vayan a Volcar, para que las filas que no se tocan
'     queden idénticas.
'   - Las columnas que no se cargan valen Empty en T.Datos. Se pueden rellenar y Volcar igualmente.
'
' Para filtrar la copia en RAM igual que los filtros de Excel: módulo Rut_Lo_CriT_Ram.
'
' Subs públicas:
'   Rut_TablaRam_Cargar(T, Lo, [Columnas], [LeerValue2])    - Lee de la hoja las columnas pedidas
'   Rut_TablaRam_Volcar(T, Lo)                              - Escribe en la hoja las columnas Modificadas
'                                                             (las seguidas, de una sola vez)
'   Rut_TablaRam_Vaciar_Celdas(Lo, Filas, Columnas)         - Vacía en la hoja unas celdas sueltas
'   Rut_TablaRam_Ordenar_Tandas(Lo, Tandas, [Primera])      - Una ordenación con el resultado de varias
'   Rut_TablaRam_Anotar_Orden(Tandas, Cols)                 - Anota una ordenación para Ordenar_Tandas
'   Rut_TablaRam_Marcar(T, Filas, Col, Valor)               - Pone un valor en las filas marcadas
'===================================================================================================

' Copia en RAM de un ListObject. Pública porque viaja como parámetro entre módulos.
Public Type T_TablaRam
    NumFilas            As Long
    NumCols             As Long
    Titulos()           As Variant      ' (1 To NumCols): cabecera de cada columna
    Datos()             As Variant      ' (1 To NumFilas, 1 To NumCols): mismos índices que la tabla
    Cargada()           As Boolean      ' (1 To NumCols): columna leída de la hoja
    Modificada()        As Boolean      ' (1 To NumCols): columna a devolver a la hoja en Volcar
End Type

'===================================================================================================
'- Carga en T las Columnas pedidas de Lo (todas, si no se pasa la lista) ---------------------------
Sub Rut_TablaRam_Cargar(T As T_TablaRam, _
                        Lo As ListObject, _
                        Optional Columnas As Variant, _
                        Optional ByVal LeerValue2 As Boolean = False)
    Dim Cabecera    As Variant
    Dim ColData     As Variant
    Dim Col         As Variant
    Dim C           As Long
    Dim Fila        As Long

    T.NumCols = Lo.ListColumns.Count
    ReDim T.Titulos(1 To T.NumCols)
    ReDim T.Cargada(1 To T.NumCols)
    ReDim T.Modificada(1 To T.NumCols)
    Cabecera = Fnc_TablaRam_Leer(Lo.HeaderRowRange, False)
    For C = 1 To T.NumCols
        T.Titulos(C) = Cabecera(1, C)
    Next C

    If Lo.DataBodyRange Is Nothing Then
        T.NumFilas = 0
        Erase T.Datos
        Exit Sub
    End If
    T.NumFilas = Lo.ListRows.Count

    If IsMissing(Columnas) Then                     '- Toda la tabla de una sola lectura -----------
        T.Datos = Fnc_TablaRam_Leer(Lo.DataBodyRange, LeerValue2)
        For C = 1 To T.NumCols
            T.Cargada(C) = True
        Next C
    Else                                            '- Solo las columnas pedidas -------------------
        ReDim T.Datos(1 To T.NumFilas, 1 To T.NumCols)
        For Each Col In Columnas
            C = Col
            If C < 1 Or C > T.NumCols Then
                Err.Raise vbObjectError + 513, "Rut_TablaRam_Cargar", _
                          "La columna " & C & " no existe en la tabla " & Lo.Name & " (" & T.NumCols & " columnas)."
            End If
            If Not T.Cargada(C) Then
                ColData = Fnc_TablaRam_Leer(Lo.ListColumns(C).DataBodyRange, LeerValue2)
                For Fila = 1 To T.NumFilas
                    T.Datos(Fila, C) = ColData(Fila, 1)
                Next Fila
                T.Cargada(C) = True
            End If
        Next Col
    End If
End Sub     ' Rut_TablaRam_Cargar
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Escribe en Lo las columnas de T marcadas como Modificadas. Las que van seguidas, en una sola ----
'- escritura (2026-10-05; antes, una por columna)
Sub Rut_TablaRam_Volcar(T As T_TablaRam, Lo As ListObject)
    Dim Bloque()    As Variant
    Dim C           As Long
    Dim C2          As Long
    Dim K           As Long
    Dim Fila        As Long

    If T.NumFilas = 0 Then Exit Sub
    If Lo.ListRows.Count <> T.NumFilas Then
        Err.Raise vbObjectError + 514, "Rut_TablaRam_Volcar", _
                  "La tabla " & Lo.Name & " tiene " & Lo.ListRows.Count & " filas y su copia en RAM " & _
                  T.NumFilas & ": no se vuelca para no descuadrar los datos."
    End If
    C = 1
    Do While C <= T.NumCols
        If T.Modificada(C) Then
            C2 = C                                  '- Hasta dónde llegan las Modificadas seguidas
            Do While C2 < T.NumCols
                If Not T.Modificada(C2 + 1) Then Exit Do
                C2 = C2 + 1
            Loop
            ReDim Bloque(1 To T.NumFilas, 1 To C2 - C + 1)
            For K = C To C2
                For Fila = 1 To T.NumFilas
                    Bloque(Fila, K - C + 1) = T.Datos(Fila, K)
                Next Fila
                T.Modificada(K) = False
            Next K
            Lo.ListColumns(C).DataBodyRange.Resize(, C2 - C + 1).Value = Bloque
            C = C2 + 1
        Else
            C = C + 1
        End If
    Loop
End Sub     ' Rut_TablaRam_Volcar
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Vacía (ClearContents) en la hoja las celdas de las Columnas en las filas True de Filas ----------
'-   Para columnas que Volcar no puede reescribir enteras (textos con cifras, como CTA. CCC). Las filas
'-   seguidas van en un solo rango, y los rangos de cada columna se juntan hasta 250 caracteres de
'-   dirección por llamada. Filas va por posición, como T.Datos: llamarla antes de ordenar o borrar filas.
Sub Rut_TablaRam_Vaciar_Celdas(Lo As ListObject, Filas() As Boolean, Columnas As Variant)
    Dim Ws          As Worksheet:   Set Ws = Lo.Parent
    Dim Fila1       As Long:        Fila1 = Lo.DataBodyRange.Row
    Dim Col         As Variant
    Dim Letra       As String
    Dim Direcc      As String
    Dim Tramo       As String
    Dim Fila        As Long
    Dim Ini         As Long

    For Each Col In Columnas
        Letra = Split(Lo.ListColumns(Col).DataBodyRange.Cells(1, 1).Address(True, False), "$")(0)
        Direcc = ""
        Fila = 1
        Do While Fila <= UBound(Filas)
            If Filas(Fila) Then
                Ini = Fila                          '- Hasta dónde llega el tramo de filas seguidas
                Do While Fila < UBound(Filas)
                    If Not Filas(Fila + 1) Then Exit Do
                    Fila = Fila + 1
                Loop
                Tramo = Letra & (Fila1 + Ini - 1) & ":" & Letra & (Fila1 + Fila - 1)
                If Len(Direcc) + Len(Tramo) + 1 > 250 Then
                    Ws.Range(Direcc).ClearContents
                    Direcc = ""
                End If
                If Direcc = "" Then Direcc = Tramo Else Direcc = Direcc & "," & Tramo
            End If
            Fila = Fila + 1
        Loop
        If Direcc <> "" Then Ws.Range(Direcc).ClearContents
    Next Col
End Sub     ' Rut_TablaRam_Vaciar_Celdas
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Ordena Lo UNA vez, con el resultado de las ordenaciones ascendentes de Tandas hechas una tras ---
'- otra.
'-   Tandas = Array(Array(cols de la 1ª), Array(cols de la 2ª), ...), en el orden en que se hacían.
'-   Como la ordenación de Excel es estable, equivale a ordenar por las columnas de la última tanda,
'-   luego por las de la penúltima, y así hasta la primera (una columna repetida solo cuenta la 1ª vez).
'-   Primera (opcional): columna que va delante de todas (en M_111, la auxiliar de las filas a borrar,
'-   que así quedan juntas al final sin cambiar el orden de las demás).
'-   Quita antes los filtros: con filas ocultas, Excel solo ordenaría las visibles.
'-   Sirve para que el paso a RAM deje las filas en el mismo orden que el código anterior: M_112 se
'-   queda con el último duplicado de cada Ref y M_115 da el importe al primer recibo de cada matrícula.
Sub Rut_TablaRam_Ordenar_Tandas(Lo As ListObject, Tandas As Variant, Optional ByVal Primera As Long = 0)
    Dim Claves()    As Integer
    Dim N           As Long
    Dim I           As Long
    Dim K           As Long
    Dim Col         As Variant
    Dim Vistas      As String:      Vistas = ","

    If Lo.DataBodyRange Is Nothing Then Exit Sub
    ReDim Claves(1 To 64)
    If Primera > 0 Then
        N = 1
        Claves(1) = Primera
        Vistas = Vistas & Primera & ","
    End If
    For I = UBound(Tandas) To LBound(Tandas) Step -1
        For Each Col In Tandas(I)
            If InStr(Vistas, "," & Col & ",") = 0 Then
                N = N + 1
                Claves(N) = Col
                Vistas = Vistas & Col & ","
            End If
        Next Col
    Next I
    If N = 0 Then Exit Sub
    Call Rut_Lo_Filtros_Quitar(Lo)
    For K = 1 To N
        Call Rut_Lo_Sort(Lo, Claves(K), xlAscending, (K = 1), Aplicar:=(K = N))
    Next K
End Sub     ' Rut_TablaRam_Ordenar_Tandas
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Anota en Tandas una de las ordenaciones que hacía el código anterior, para Rut_TablaRam_Ordenar_Tandas.
'-   Cols = Array(columnas de esa ordenación, en su orden). Tandas empieza sin dimensionar.
'-   (Era Rut_Anotar_Orden de M_111; desde el 2026-10-06 la usa también M_211)
Sub Rut_TablaRam_Anotar_Orden(Tandas() As Variant, ByVal Cols As Variant)
    Dim N       As Long
    On Error Resume Next
    N = UBound(Tandas)                          '- Error si todavía está vacío: N se queda en 0
    On Error GoTo 0
    ReDim Preserve Tandas(1 To N + 1)
    Tandas(N + 1) = Cols
End Sub     ' Rut_TablaRam_Anotar_Orden
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Pone Valor en la columna Col de T, en las filas que valen True en Filas, y la marca como Modificada
'-   (Era Rut_Marcar_Obs de M_111, solo para Obs_Conta; desde el 2026-10-06 la usa también M_211)
Sub Rut_TablaRam_Marcar(T As T_TablaRam, Filas() As Boolean, ByVal Col As Long, ByVal Valor As Variant)
    Dim Fila    As Long
    For Fila = 1 To T.NumFilas
        If Filas(Fila) Then T.Datos(Fila, Col) = Valor
    Next Fila
    T.Modificada(Col) = True
End Sub     ' Rut_TablaRam_Marcar
'---------------------------------------------------------------------------------------------------

'- Devuelve SIEMPRE un array 2D (1 To filas, 1 To columnas): .Value de una sola celda no es array --
Private Function Fnc_TablaRam_Leer(Rng As Range, ByVal LeerValue2 As Boolean) As Variant
    Dim Arr     As Variant
    If Rng.Cells.CountLarge = 1 Then
        ReDim Arr(1 To 1, 1 To 1)
        If LeerValue2 Then Arr(1, 1) = Rng.Value2 Else Arr(1, 1) = Rng.Value
        Fnc_TablaRam_Leer = Arr
    ElseIf LeerValue2 Then
        Fnc_TablaRam_Leer = Rng.Value2
    Else
        Fnc_TablaRam_Leer = Rng.Value
    End If
End Function
'---------------------------------------------------------------------------------------------------
