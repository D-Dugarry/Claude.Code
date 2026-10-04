Attribute VB_Name = "Rut_Lo_TablaRam"
' Last Rev. 2026-10-02 20:15
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
'     solo columnas de importes o de textos que no puedan confundirse con números o fechas.
'   - .Value (por defecto) devuelve las fechas como Date y los importes con formato moneda como
'     Currency (redondeado a 4 decimales), igual que leer .Cells(i, j): úsalo cuando se monten
'     textos o comparaciones con esos valores. .Value2 (LeerValue2:=True) los deja como Double sin
'     redondear: úsalo en las columnas que se vayan a Volcar, para que las filas que no se tocan
'     queden idénticas.
'   - Las columnas que no se cargan valen Empty en T.Datos.
'
' Subs públicas:
'   Rut_TablaRam_Cargar(T, Lo, [Columnas], [LeerValue2])    - Lee de la hoja las columnas pedidas
'   Rut_TablaRam_Volcar(T, Lo)                              - Escribe en la hoja las columnas Modificadas
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
'- Escribe en Lo las columnas de T marcadas como Modificadas, una escritura por columna ------------
Sub Rut_TablaRam_Volcar(T As T_TablaRam, Lo As ListObject)
    Dim ColData()   As Variant
    Dim C           As Long
    Dim Fila        As Long

    If T.NumFilas = 0 Then Exit Sub
    If Lo.ListRows.Count <> T.NumFilas Then
        Err.Raise vbObjectError + 514, "Rut_TablaRam_Volcar", _
                  "La tabla " & Lo.Name & " tiene " & Lo.ListRows.Count & " filas y su copia en RAM " & _
                  T.NumFilas & ": no se vuelca para no descuadrar los datos."
    End If
    ReDim ColData(1 To T.NumFilas, 1 To 1)
    For C = 1 To T.NumCols
        If T.Modificada(C) Then
            For Fila = 1 To T.NumFilas
                ColData(Fila, 1) = T.Datos(Fila, C)
            Next Fila
            Lo.ListColumns(C).DataBodyRange.Value = ColData
            T.Modificada(C) = False
        End If
    Next C
End Sub     ' Rut_TablaRam_Volcar
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
