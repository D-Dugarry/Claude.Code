Attribute VB_Name = "Z_Verificar_CriT"
' Last Rev. 2026-10-05 23:11
Option Explicit

'===================================================================================================
' Z_Verificar_CriT   (módulo de un solo uso: se quita del libro cuando la fase 3 esté comprobada)
'
' Comprueba que los filtros en RAM (Rut_Lo_CriT_Ram) dejan las mismas filas que los filtros reales de
' Excel. Para cada rango Tb_CriT_* y para cada AutoFilter que usan M_111, M_113 y M_114, filtra la
' tabla de verdad, mira qué filas quedan visibles y lo compara con lo que dice la función en RAM.
'
'   Z_Verificar_CriT_BDatos  - Sobre la tabla real de BDatos. No cambia ningún dato: solo filtra (y
'                              deja la hoja como la deja cualquier proceso: columnas a la vista y sin
'                              filtros). Tarda 1 o 2 minutos.
'   Z_Verificar_CriT_Casos   - Sobre una tabla de casos raros (celdas vacías y con "", textos con
'                              cifras, fechas de fin de año, "M013B", "C404X"...) que monta en un libro
'                              nuevo y cierra sin guardar. Además prueba qué pasa con las filas ocultas
'                              por un AdvancedFilter: si Rut_Lo_Filtros_Quitar las vuelve a mostrar y si
'                              la ordenación de después las mueve. De eso depende el orden de las filas
'                              que reciben M_112 (qué duplicado se queda) y M_115 (el primer recibo de
'                              cada matrícula).
'
' El resultado sale en la ventana Inmediato y se añade al fichero Verificar_CriT.txt de la carpeta del
' libro.
'===================================================================================================

Private Const Z_Max_Ejemplos    As Long = 5
Private Const Z_Filas_Azar      As Long = 4000

Private Z_Informe               As String
Private Z_Errores               As Long
Private Z_Fila_Crit             As Long         '- Siguiente fila libre de la hoja de criterios copiados

'===================================================================================================
Sub Z_Verificar_CriT_BDatos()
    Dim Lo      As ListObject:  Set Lo = Sht__BD.ListObjects(1)

    Z_Informe = ""
    Z_Errores = 0
    Call Rut_Off_Functions
    On Error GoTo Fallo
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Call Rut_Z_Linea(String(100, "="))
    Call Rut_Z_Linea("Z_Verificar_CriT_BDatos  " & Format(Now, "dd/mm/yyyy hh:mm") & "  tabla " & Lo.Name & ", " & _
                     Format(Lo.ListRows.Count, "#,##0") & " filas")
    Call Rut_Z_Verificar_Lo(Lo, Nothing)
    Call Rut_Z_Linea(IIf(Z_Errores = 0, "TODO IGUAL", Z_Errores & " PRUEBAS CON DIFERENCIAS"))

Salir:
    On Error Resume Next
    Call Rut_Z_Quitar_Filtros(Lo)
    Call Rut_Lo_Totales_Mostrar
    Call Rut_On_Functions
    Call Rut_Z_Guardar_Informe
    Exit Sub
Fallo:
    Call Rut_Z_Linea("ERROR " & Err.Number & " (" & Err.Source & "): " & Err.Description)
    Resume Salir
End Sub
'---------------------------------------------------------------------------------------------------

'===================================================================================================
Sub Z_Verificar_CriT_Casos()
    Dim Wb      As Workbook
    Dim WsD     As Worksheet
    Dim WsC     As Worksheet
    Dim Lo      As ListObject

    Z_Informe = ""
    Z_Errores = 0
    Z_Fila_Crit = 1
    Call Rut_Off_Functions
    On Error GoTo Fallo
    Set Wb = Workbooks.Add(xlWBATWorksheet)
    Set WsD = Wb.Worksheets(1)
    Set WsC = Wb.Worksheets.Add(After:=WsD)
    Set Lo = Fnc_Z_Casos_Tabla(WsD, Sht__BD.ListObjects(1))
    Call Rut_Z_Linea(String(100, "="))
    Call Rut_Z_Linea("Z_Verificar_CriT_Casos  " & Format(Now, "dd/mm/yyyy hh:mm") & "  tabla de casos, " & _
                     Format(Lo.ListRows.Count, "#,##0") & " filas al azar")
    Call Rut_Z_Verificar_Lo(Lo, WsC)
    Call Rut_Z_Linea(IIf(Z_Errores = 0, "TODO IGUAL", Z_Errores & " PRUEBAS CON DIFERENCIAS"))
    Call Rut_Z_Orden(Lo, WsC)

Salir:
    On Error Resume Next
    If Not Wb Is Nothing Then Wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Call Rut_On_Functions
    Call Rut_Z_Guardar_Informe
    Exit Sub
Fallo:
    Call Rut_Z_Linea("ERROR " & Err.Number & " (" & Err.Source & "): " & Err.Description)
    Resume Salir
End Sub
'---------------------------------------------------------------------------------------------------

'- Compara, en la tabla Lo, cada filtro de Excel con su versión en RAM -----------------------------
'-   WsC: Nothing = los rangos de criterios del libro; una hoja = se copian (valores) a esa hoja,
'-   porque la tabla de casos está en otro libro.
Private Sub Rut_Z_Verificar_Lo(Lo As ListObject, WsC As Worksheet)
    Dim T           As T_TablaRam
    Dim Nombre      As Variant
    Dim Prueba      As Variant
    Dim Excel_()    As Boolean
    Dim Ram()       As Boolean
    Dim AnoCont     As Long:    AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim Cierre      As Date:    Cierre = CDate(Prog__APP.Range("APP_FechCierreCont"))
    Dim Desde       As Double
    Dim Hasta       As Double

    Call Rut_Z_Quitar_Filtros(Lo)
    Call Rut_TablaRam_Cargar(T, Lo, , True)                     '- Toda la tabla, con .Value2

    '- 1) Rangos de criterios, con AdvancedFilter ------------------------------------------------
    Call Rut_Z_Linea("-- AdvancedFilter con los rangos Tb_CriT_*")
    For Each Nombre In Fnc_Z_Nombres_CriT()
        Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit(CStr(Nombre), WsC)
        Excel_ = Fnc_Z_Visibles(Lo)
        Ram = Fnc_CriT_Filas(T, CStr(Nombre))
        Call Rut_Z_Comparar(CStr(Nombre), Excel_, Ram, T, Fnc_Z_Cols_CriT(T, CStr(Nombre)))
        Call Rut_Z_Quitar_Filtros(Lo)
    Next Nombre

    '- 2) Los AutoFilter de M_111, M_113 y M_114 (criterio de Excel, y su equivalente en RAM) ------
    Call Rut_Z_Linea("-- AutoFilter de M_111, M_113 y M_114")
    For Each Prueba In Array( _
            Array(BD_ActivEco, "=4", "=", 4), _
            Array(BD_ImpRec, "<0", "<", 0), _
            Array(BD_ImpRec, "=0", "=", 0), _
            Array(BD_DNI, "=1", "=", "1"), _
            Array(BD_ActivEco, "=300", "=", 300), _
            Array(BD_Anul, "=S", "=", "S"), _
            Array(BD_Matricula, "=N", "=", "N"), _
            Array(BD_Hinvalid, "=S", "=", "S"), _
            Array(BD_FCob, ">" & Format(Cierre, "mm\/dd\/yyyy"), ">", CDbl(Cierre)), _
            Array(BD_FEmi, ">" & Format(Cierre, "mm\/dd\/yyyy"), ">", CDbl(Cierre)), _
            Array(BD_FVto, "<01/01/" & AnoCont - 1, "<", CDbl(DateSerial(AnoCont - 1, 1, 1))), _
            Array(BD_FVto, ">=01/01/" & AnoCont + 1, ">=", CDbl(DateSerial(AnoCont + 1, 1, 1))), _
            Array(BD_ActivEco, 5, "=", 5), _
            Array(BD_ActivEco, 2, "=", 2), _
            Array(BD_ActivEco, 80, "=", 80), _
            Array(BD_ACont_Cob, "<" & AnoCont, "<", AnoCont), _
            Array(BD_TIO_EP, "=", "=", ""), _
            Array(BD_Tipo_Rec, "=", "=", ""))
        Lo.Range.AutoFilter Field:=Prueba(0), Criteria1:=Prueba(1)
        Excel_ = Fnc_Z_Visibles(Lo)
        Ram = Fnc_Filtro_Filas(T, Prueba(0), Prueba(2), Prueba(3))
        Call Rut_Z_Comparar(T.Titulos(Prueba(0)) & " " & Prueba(1), Excel_, Ram, T, Array(Prueba(0)))
        Call Rut_Z_Quitar_Filtros(Lo)
    Next Prueba
    '- Los dos de M_113 con dos condiciones sobre F_Vto
    For Each Prueba In Array(Array(AnoCont, AnoCont + 1), Array(AnoCont - 1, AnoCont))
        Desde = CDbl(DateSerial(Prueba(0), 1, 1))
        Hasta = CDbl(DateSerial(Prueba(1), 1, 1))
        Lo.Range.AutoFilter Field:=BD_FVto, Criteria1:=">=01/01/" & Prueba(0), Operator:=xlAnd, Criteria2:="<01/01/" & Prueba(1)
        Excel_ = Fnc_Z_Visibles(Lo)
        Ram = Fnc_Filtro_Filas(T, BD_FVto, "<", Hasta, Fnc_Filtro_Filas(T, BD_FVto, ">=", Desde))
        Call Rut_Z_Comparar(T.Titulos(BD_FVto) & " >=01/01/" & Prueba(0) & " y <01/01/" & Prueba(1), Excel_, Ram, T, Array(BD_FVto))
        Call Rut_Z_Quitar_Filtros(Lo)
    Next Prueba

    '- 3) Un filtro encima de otro sin quitar el primero (M_114 lo hace con Tb_CriT_Reg_Anul) -----
    '-    Si Excel solo mirase las filas que dejó el primero, aquí saldrían diferencias.
    Call Rut_Z_Linea("-- Un filtro aplicado sobre otro, sin quitar el primero")
    Lo.Range.AutoFilter Field:=BD_ImpRec, Criteria1:="<0"
    Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit("Tb_CriT_Reg_Anul", WsC)
    Excel_ = Fnc_Z_Visibles(Lo)
    Ram = Fnc_CriT_Filas(T, "Tb_CriT_Reg_Anul")
    Call Rut_Z_Comparar("AutoFilter Imp_Rec<0 y luego Tb_CriT_Reg_Anul", Excel_, Ram, T, Fnc_Z_Cols_CriT(T, "Tb_CriT_Reg_Anul"))
    Call Rut_Z_Quitar_Filtros(Lo)
    Lo.Range.AutoFilter Field:=BD_ImpRec, Criteria1:="<0"
    Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit("Tb_CriT_Emitido", WsC)
    Excel_ = Fnc_Z_Visibles(Lo)
    Ram = Fnc_CriT_Filas(T, "Tb_CriT_Emitido")
    Call Rut_Z_Comparar("AutoFilter Imp_Rec<0 y luego Tb_CriT_Emitido", Excel_, Ram, T, Fnc_Z_Cols_CriT(T, "Tb_CriT_Emitido"))
    Call Rut_Z_Quitar_Filtros(Lo)
    Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit("Tb_CriT_Emitido", WsC)
    Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit("Tb_CriT_Aplazado", WsC)
    Excel_ = Fnc_Z_Visibles(Lo)
    Ram = Fnc_CriT_Filas(T, "Tb_CriT_Aplazado")
    Call Rut_Z_Comparar("Tb_CriT_Emitido y luego Tb_CriT_Aplazado", Excel_, Ram, T, Fnc_Z_Cols_CriT(T, "Tb_CriT_Aplazado"))
    Call Rut_Z_Quitar_Filtros(Lo)
    Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit("Tb_CriT_Emitido", WsC)
    Call Rut_Lo_Filtros_Quitar(Lo)                              '- Lo que hace M_111 antes de su AutoFilter Imp_Rec=0
    Lo.Range.AutoFilter Field:=BD_ImpRec, Criteria1:="=0"
    Excel_ = Fnc_Z_Visibles(Lo)
    Ram = Fnc_Filtro_Filas(T, BD_ImpRec, "=", 0)
    Call Rut_Z_Comparar("Tb_CriT_Emitido, Rut_Lo_Filtros_Quitar y luego AutoFilter Imp_Rec=0", Excel_, Ram, T, Array(BD_ImpRec))
    Call Rut_Z_Quitar_Filtros(Lo)
End Sub     ' Rut_Z_Verificar_Lo
'---------------------------------------------------------------------------------------------------

'- Qué pasa con las filas ocultas por un AdvancedFilter (en la tabla de casos, que se puede estropear) -
'-   Reproduce las tres secuencias del código actual que ordenan justo después de un AdvancedFilter:
'-   M_111 (Tb_CriT_ImpMatCero, borrar visibles, Rut_Lo_Filtros_Quitar, ordenar por Imp_Rec),
'-   M_113 (Tb_CriT_RecAdm, Rut_Lo_Filtros_Quitar, ordenar por Plan) y
'-   M_114 (Tb_CriT_Reg_Err, .ShowAutoFilter = True, ordenar por Tipo_Rec).
Private Sub Rut_Z_Orden(Lo As ListObject, WsC As Worksheet)
    Dim Vis()   As Boolean

    Call Rut_Z_Linea("-- Filas ocultas por un AdvancedFilter y ordenación de después")

    Call Rut_Z_Quitar_Filtros(Lo)
    Call Rut_Lo_Sort(Lo, BD_Ref, xlAscending, True)             '- Orden de partida: Ref = nº de fila
    Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit("Tb_CriT_RecAdm", WsC)
    Call Rut_Z_Estado(Lo, "M_113: tras AdvancedFilter Tb_CriT_RecAdm")
    Call Rut_Lo_Filtros_Quitar(Lo)
    Call Rut_Z_Estado(Lo, "M_113: tras Rut_Lo_Filtros_Quitar")
    Call Rut_Lo_Sort(Lo, BD_Plan, xlAscending, True)
    Call Rut_Z_Estado_Orden(Lo, BD_Plan, "M_113: tras ordenar por Plan")

    Call Rut_Z_Quitar_Filtros(Lo)
    Call Rut_Lo_Sort(Lo, BD_Ref, xlAscending, True)
    Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit("Tb_CriT_Reg_Err", WsC)
    Call Rut_Z_Estado(Lo, "M_114: tras AdvancedFilter Tb_CriT_Reg_Err")
    Lo.ShowAutoFilter = True
    Call Rut_Z_Estado(Lo, "M_114: tras .ShowAutoFilter = True")
    Call Rut_Lo_Sort(Lo, BD_Tipo_Rec, xlAscending, True)
    Call Rut_Z_Estado_Orden(Lo, BD_Tipo_Rec, "M_114: tras ordenar por Tipo_Rec")

    Call Rut_Z_Quitar_Filtros(Lo)                               '- La última, porque borra filas
    Call Rut_Lo_Sort(Lo, BD_Ref, xlAscending, True)
    Lo.Range.AdvancedFilter xlFilterInPlace, Fnc_Z_Rango_Crit("Tb_CriT_ImpMatCero", WsC)
    Call Rut_Z_Estado(Lo, "M_111: tras AdvancedFilter Tb_CriT_ImpMatCero")
    Vis = Fnc_Z_Visibles(Lo)
    If Fnc_Filas_Contar(Vis) > 0 Then Lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
    Call Rut_Z_Estado(Lo, "M_111: tras borrar las visibles")
    Call Rut_Lo_Filtros_Quitar(Lo)
    Call Rut_Z_Estado(Lo, "M_111: tras Rut_Lo_Filtros_Quitar")
    Call Rut_Lo_Sort(Lo, BD_ImpRec, xlAscending, True)
    Call Rut_Z_Estado_Orden(Lo, BD_ImpRec, "M_111: tras ordenar por Imp_Rec")
End Sub
'---------------------------------------------------------------------------------------------------

'- Filas visibles y estado de los filtros ----------------------------------------------------------
Private Sub Rut_Z_Estado(Lo As ListObject, ByVal Momento As String)
    Dim Vis()   As Boolean:     Vis = Fnc_Z_Visibles(Lo)
    Call Rut_Z_Linea("  " & Momento & ": visibles " & Format(Fnc_Filas_Contar(Vis), "#,##0") & " de " & _
                     Format(Lo.ListRows.Count, "#,##0") & ", ShowAutoFilter=" & Lo.ShowAutoFilter & _
                     ", hoja FilterMode=" & Lo.Parent.FilterMode)
End Sub
'---------------------------------------------------------------------------------------------------

'- ¿Ha ordenado Excel todas las filas, solo las visibles o ninguna? --------------------------------
'-   Mira si la columna queda en orden (números, luego textos, luego vacías; Ref creciente en los empates,
'-   porque la ordenación es estable), en toda la tabla y solo entre las filas visibles.
Private Sub Rut_Z_Estado_Orden(Lo As ListObject, ByVal Col As Long, ByVal Momento As String)
    Dim T           As T_TablaRam
    Dim Vis()       As Boolean
    Dim Fila        As Long
    Dim Ant         As Long
    Dim Todas       As Boolean:     Todas = True
    Dim Visibles    As Boolean:     Visibles = True
    Dim Ocultas     As Long

    Vis = Fnc_Z_Visibles(Lo)
    Call Rut_TablaRam_Cargar(T, Lo, Array(Col, BD_Ref), True)
    For Fila = 2 To T.NumFilas
        If Fnc_Z_Antes(T.Datos(Fila, Col), T.Datos(Fila, BD_Ref), T.Datos(Fila - 1, Col), T.Datos(Fila - 1, BD_Ref)) Then Todas = False
    Next Fila
    Ant = 0
    For Fila = 1 To T.NumFilas
        If Vis(Fila) Then
            If Ant > 0 Then
                If Fnc_Z_Antes(T.Datos(Fila, Col), T.Datos(Fila, BD_Ref), T.Datos(Ant, Col), T.Datos(Ant, BD_Ref)) Then Visibles = False
            End If
            Ant = Fila
        Else
            Ocultas = Ocultas + 1
        End If
    Next Fila
    Call Rut_Z_Linea("  " & Momento & ": toda la tabla en orden=" & Todas & ", las visibles en orden=" & Visibles & _
                     " (ocultas " & Format(Ocultas, "#,##0") & ")")
End Sub
'---------------------------------------------------------------------------------------------------

'- ¿Va (V1, Ref1) antes que (V2, Ref2) en el orden ascendente de Excel? ----------------------------
Private Function Fnc_Z_Antes(V1 As Variant, R1 As Variant, V2 As Variant, R2 As Variant) As Boolean
    Dim C1      As Long:    C1 = Fnc_Z_Clase(V1)
    Dim C2      As Long:    C2 = Fnc_Z_Clase(V2)
    Dim Rel     As Long
    If C1 <> C2 Then
        Fnc_Z_Antes = (C1 < C2)
        Exit Function
    End If
    If C1 = 0 Then
        If V1 < V2 Then
            Rel = -1
        ElseIf V1 > V2 Then
            Rel = 1
        End If
    ElseIf C1 = 1 Then
        Rel = StrComp(V1, V2, vbTextCompare)
    End If
    If Rel = 0 Then Fnc_Z_Antes = (R1 < R2) Else Fnc_Z_Antes = (Rel < 0)
End Function

Private Function Fnc_Z_Clase(V As Variant) As Long          '- 0 número, 1 texto, 2 vacía
    If IsEmpty(V) Then
        Fnc_Z_Clase = 2
    ElseIf VarType(V) = vbString Then
        If Len(V) = 0 Then Fnc_Z_Clase = 2 Else Fnc_Z_Clase = 1
    Else
        Fnc_Z_Clase = 0
    End If
End Function
'---------------------------------------------------------------------------------------------------

'- Compara las filas de Excel con las de RAM y lo apunta en el informe -----------------------------
Private Sub Rut_Z_Comparar(ByVal Prueba As String, Excel_() As Boolean, Ram() As Boolean, _
                           T As T_TablaRam, Cols As Variant)
    Dim Fila        As Long
    Dim NExcel      As Long
    Dim NRam        As Long
    Dim NDif        As Long
    Dim Ejemplos    As String
    Dim C           As Variant

    For Fila = 1 To T.NumFilas
        If Excel_(Fila) Then NExcel = NExcel + 1
        If Ram(Fila) Then NRam = NRam + 1
        If Excel_(Fila) <> Ram(Fila) Then
            NDif = NDif + 1
            If NDif <= Z_Max_Ejemplos Then
                Ejemplos = Ejemplos & vbCrLf & "         fila " & Fila & ": Excel " & IIf(Excel_(Fila), "SÍ", "no") & _
                           ", RAM " & IIf(Ram(Fila), "SÍ", "no") & " |"
                For Each C In Cols
                    Ejemplos = Ejemplos & " " & T.Titulos(C) & "=" & Fnc_Z_Mostrar(T.Datos(Fila, C))
                Next C
            End If
        End If
    Next Fila
    If NDif = 0 Then
        Call Rut_Z_Linea("  OK   " & Prueba & ": " & Format(NExcel, "#,##0") & " filas")
    Else
        Z_Errores = Z_Errores + 1
        Call Rut_Z_Linea("  DIF  " & Prueba & ": Excel " & Format(NExcel, "#,##0") & ", RAM " & Format(NRam, "#,##0") & _
                         ", distintas " & Format(NDif, "#,##0") & Ejemplos)
    End If
End Sub
'---------------------------------------------------------------------------------------------------

'- Valor de una celda para el informe, con su tipo a la vista --------------------------------------
Private Function Fnc_Z_Mostrar(V As Variant) As String
    If IsEmpty(V) Then
        Fnc_Z_Mostrar = "(vacía)"
    ElseIf VarType(V) = vbString Then
        Fnc_Z_Mostrar = "'" & V & "'"
    ElseIf IsError(V) Then
        Fnc_Z_Mostrar = "(error)"
    Else
        Fnc_Z_Mostrar = CStr(V)
    End If
End Function
'---------------------------------------------------------------------------------------------------

'- Filas visibles del cuerpo de la tabla (por áreas, sin mirar fila a fila) ------------------------
Private Function Fnc_Z_Visibles(Lo As ListObject) As Boolean()
    Dim Vis()   As Boolean
    Dim Rng     As Range
    Dim Area    As Range
    Dim Fila0   As Long
    Dim Fila    As Long

    ReDim Vis(1 To Lo.ListRows.Count)
    Fila0 = Lo.DataBodyRange.Row - 1
    On Error Resume Next
    Set Rng = Lo.DataBodyRange.Columns(1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not Rng Is Nothing Then
        For Each Area In Rng.Areas
            For Fila = Area.Row - Fila0 To Area.Row - Fila0 + Area.Rows.Count - 1
                Vis(Fila) = True
            Next Fila
        Next Area
    End If
    Fnc_Z_Visibles = Vis
End Function
'---------------------------------------------------------------------------------------------------

'- Quita cualquier filtro (también el de un AdvancedFilter) y deja puestos los botones del AutoFilter -
Private Sub Rut_Z_Quitar_Filtros(Lo As ListObject)
    On Error Resume Next
    If Lo.Parent.FilterMode Then Lo.Parent.ShowAllData
    If Not Lo.ShowAutoFilter Then Lo.ShowAutoFilter = True
    If Lo.AutoFilter.FilterMode Then Lo.AutoFilter.ShowAllData
    Lo.Parent.Rows.Hidden = False
End Sub
'---------------------------------------------------------------------------------------------------

'- Rangos de criterios que se comprueban: todos menos Tb_CriT_Dto y Tb_CriT_DEV, que usan la columna -
'- Tipo_Dto, que BDatos no tiene (ni el código los usa)
Private Function Fnc_Z_Nombres_CriT() As Variant
    Fnc_Z_Nombres_CriT = Array( _
        "Tb_CriT_ImpMatCero", "Tb_CriT_Reg_Err", "Tb_CriT_Reg_AE6", "Tb_CriT_EFP", "Tb_CriT_CFC", "Tb_CriT_AFC", _
        "Tb_CriT_TNCT_M013", "Tb_CriT_TUP", "Tb_CriT_TNCT_PNB1", "Tb_CriT_RecAdm", "Tb_CriT_Reg_EURLE", _
        "Tb_CriT_Emitido", "Tb_CriT_EjeAnt", "Tb_CriT_Aneja", "Tb_CriT_Aplazado", "Tb_CriT_ADxAplz", "Tb_CriT_Reg_Anul", _
        "Tb_CriT_FVto_Act", "Tb_CriT_FVto_Ant", "Tb_CriT_FEmiCierreCont", "Tb_CriT_FCobCierreCont", _
        "Tb_CriT_AE4_Acont", "Tb_CriT_AE4_NO_Acont")
End Function
'---------------------------------------------------------------------------------------------------

'- Rango de criterios: el del libro, o una copia (valores) en WsC si la tabla está en otro libro ---
Private Function Fnc_Z_Rango_Crit(ByVal Nombre As String, WsC As Worksheet) As Range
    Dim V       As Variant
    Dim F       As Long
    Dim C       As Long
    Set Fnc_Z_Rango_Crit = ThisWorkbook.Names(Nombre).RefersToRange
    If WsC Is Nothing Then Exit Function
    V = Fnc_Z_Rango_Crit.Value
    For F = 1 To UBound(V, 1)
        For C = 1 To UBound(V, 2)
            If VarType(V(F, C)) = vbString Then V(F, C) = "'" & V(F, C)    '- Que "=" o ">0" entren como texto
        Next C
    Next F
    Set Fnc_Z_Rango_Crit = WsC.Cells(Z_Fila_Crit, 1).Resize(UBound(V, 1), UBound(V, 2))
    Fnc_Z_Rango_Crit.Value = V
    Z_Fila_Crit = Z_Fila_Crit + UBound(V, 1) + 1
End Function
'---------------------------------------------------------------------------------------------------

'- Columnas de la tabla que usa un rango de criterios (para enseñar sus valores en el informe) -----
Private Function Fnc_Z_Cols_CriT(T As T_TablaRam, ByVal Nombre As String) As Variant
    Dim Titulos As Variant
    Dim Tit     As Variant
    Dim C       As Long
    Dim Lista   As String
    Titulos = ThisWorkbook.Names(Nombre).RefersToRange.Rows(1).Value
    If Not IsArray(Titulos) Then Titulos = Array(Titulos)       '- Rango de una sola columna
    For Each Tit In Titulos
        For C = 1 To T.NumCols
            If StrComp(CStr(T.Titulos(C)), CStr(Tit), vbTextCompare) = 0 Then
                If InStr("," & Lista & ",", "," & C & ",") = 0 Then Lista = Lista & IIf(Lista = "", "", ",") & C
            End If
        Next C
    Next Tit
    Fnc_Z_Cols_CriT = Split(Lista, ",")
End Function
'---------------------------------------------------------------------------------------------------

'- Monta en WsD la tabla de casos: los títulos de BDatos y filas al azar con valores raros ---------
Private Function Fnc_Z_Casos_Tabla(WsD As Worksheet, Lo_BD As ListObject) As ListObject
    Dim A           As Long:        A = Prog__APP.Range("APP_AnoCont")
    Dim NCols       As Long:        NCols = Lo_BD.ListColumns.Count
    Dim Valores()   As Variant
    Dim Fechas      As Variant
    Dim AnoCont     As Variant
    Dim Lista       As Variant
    Dim Datos()     As Variant
    Dim F           As Long
    Dim C           As Long

    ReDim Valores(1 To NCols)
    AnoCont = Array(A - 2, A - 1, A, A + 1, Empty, "'" & A)
    Fechas = Array(DateSerial(A - 1, 12, 31), DateSerial(A, 1, 1), DateSerial(A, 12, 31), DateSerial(A, 12, 31) + 0.5, _
                   DateSerial(A + 1, 1, 15), DateSerial(A - 2, 6, 1), Empty, "abc", DateSerial(A + 1, 1, 1))
    Valores(BD_ACont_Emi) = AnoCont
    Valores(BD_ACont_Cob) = AnoCont
    Valores(BD_ACont_Vto) = AnoCont
    Valores(BD_Plan) = Array("C404", "C404X", "c404", "C40", Empty, "M013", "M013B", "PNB1", "Z1")
    Valores(BD_TipoCurso) = Array("TEP1", "tep1", "TEP12", "TNCT", "TUP", "TUPX", Empty, "TPE9", "TEP7")
    Valores(BD_DNI) = Array("'1", 1, "'01", "' 1", "1A", Empty)
    Valores(BD_Matricula) = Array("S", "N", "n", "No", Empty, "Si")
    Valores(BD_Anul) = Array("S", "s", "Si", "N", Empty)
    Valores(BD_ActivEco) = Array(2, 4, 5, 6, 7, 21, 80, 300, 710, "'4", Empty, 6.5)
    Valores(BD_FEmi) = Fechas
    Valores(BD_FVto) = Fechas
    Valores(BD_FCob) = Fechas
    Valores(BD_ImpRec) = Array(0, -1, 1, Empty, "'0", 0.0001, 100)
    Valores(BD_Hinvalid) = Array("S", "N", Empty, "s")
    Valores(BD_ImpDto) = Array(0, Empty, 5, "'0")
    Valores(BD_Concepto) = Array(Empty, "x", "=""""", 1310)     '- ="" deja una celda con el texto vacío
    Valores(BD_Tipo_Rec) = Array(Empty, "Emitido", "Aplazado")      '- Sin "": se ordena por ella
    Valores(BD_TIO_EP) = Array(Empty, "x", "=""""")

    ReDim Datos(1 To Z_Filas_Azar + 1, 1 To NCols)
    For C = 1 To NCols
        Datos(1, C) = Lo_BD.HeaderRowRange.Cells(1, C).Value
    Next C
    Rnd -1
    Randomize 20261005                                          '- Siempre las mismas filas
    For F = 2 To Z_Filas_Azar + 1
        For C = 1 To NCols
            If Not IsEmpty(Valores(C)) Then
                Lista = Valores(C)
                Datos(F, C) = Lista(Int(Rnd * (UBound(Lista) + 1)))
            End If
        Next C
        Datos(F, BD_Ref) = F - 1                                '- Ref = nº de fila de partida, para ver el orden
    Next F
    WsD.Range("A1").Resize(Z_Filas_Azar + 1, NCols).Value = Datos
    Set Fnc_Z_Casos_Tabla = WsD.ListObjects.Add(xlSrcRange, WsD.Range("A1").Resize(Z_Filas_Azar + 1, NCols), , xlYes)
End Function
'---------------------------------------------------------------------------------------------------

'- Informe -----------------------------------------------------------------------------------------
Private Sub Rut_Z_Linea(ByVal S As String)
    Debug.Print S
    Z_Informe = Z_Informe & S & vbCrLf
End Sub

Private Sub Rut_Z_Guardar_Informe()
    Dim Ruta    As String
    Dim F       As Integer
    Ruta = ThisWorkbook.Path
    If Ruta = "" Or InStr(Ruta, "://") > 0 Then Ruta = Environ$("TEMP")
    Ruta = Ruta & "\Verificar_CriT.txt"
    F = FreeFile
    Open Ruta For Append As #F
    Print #F, Z_Informe;
    Close #F
    Debug.Print "Informe añadido a " & Ruta
End Sub
'---------------------------------------------------------------------------------------------------
