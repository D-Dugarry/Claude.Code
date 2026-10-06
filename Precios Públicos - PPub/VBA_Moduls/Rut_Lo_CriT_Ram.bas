Attribute VB_Name = "Rut_Lo_CriT_Ram"
' Last Rev. 2026-10-06 11:50
Option Explicit

'===================================================================================================
' Rut_Lo_CriT_Ram
'
' Filtros de Excel hechos sobre la copia en RAM de una tabla (T_TablaRam, módulo Rut_Lo_TablaRam):
' dicen qué filas dejaría visibles el filtro, sin filtrar la hoja. Es la fase 3 del paso a RAM: que
' M_111, M_113 y M_114 recorran la tabla en memoria en lugar de filtrar la hoja, contar con
' SpecialCells y escribir en las celdas visibles en cada paso.
'
'   Fnc_CriT_Filas(T, "Tb_CriT_X", [Vivas])         - las filas que deja .AdvancedFilter xlFilterInPlace
'                                                      con el rango de criterios Tb_CriT_X
'   Fnc_Filtro_Filas(T, Col, Op, Operando, [Vivas])  - las que deja .AutoFilter Field:=Col,
'                                                      Criteria1:=Op & Operando
'   Fnc_Filas_Contar(Filas)                          - cuántas valen True
'   Fnc_Filas_Borrar(Vivas, Cumple, NVivas)          - borrado lógico: las que cumplen dejan de estar vivas
'
' Las dos primeras devuelven un Boolean(1 To T.NumFilas) con True en las filas que cumplen (con la
' tabla vacía, un (0 To 0) con False). Con Vivas, un array igual, solo se miran las filas que valen
' True en él y las demás salen False: sirve para el borrado lógico (filas ya borradas en RAM, que no
' deben contar) y para encadenar condiciones (Y). Las columnas que use el criterio tienen que estar
' cargadas en T; si no, salta un error que lo dice.
'
' El rango de criterios se lee de la hoja en cada llamada: siguen mandando las hojas Filtros_Tipo_Rec
' y Filtros_Conceptos, igual que con el AdvancedFilter.
'
' REGLAS que se aplican: las de los filtros de Excel, comparadas con el filtro real el 2026-10-05 con
' Z_Verificar_CriT (BDatos y 4.000 casos raros). Las marcadas con (*) se descubrieron así:
'   - Rango de criterios: la 1ª fila son los títulos de columna (sin distinguir mayúsculas). Cada fila
'     de debajo es una alternativa (O) y sus celdas se tienen que cumplir todas (Y). Celda vacía = sin
'     condición; una fila de criterios entera vacía deja pasar todas las filas.
'   - Celda de criterio:
'       número o fecha          -> igual a ese número
'       texto sin operador      -> EMPIEZA por ese texto: "M013" deja pasar "M013B"
'       "=" o "<>" solos        -> celda vacía / no vacía
'       =x  <>x                 -> igual / distinto; en textos, el texto entero, con comodines * ? ~
'       <x  >x  <=x  >=x        -> números con números y textos con textos (orden alfabético)
'     Los textos se comparan sin distinguir mayúsculas.
'   - El operando x es número si lo parece (2026, -1, 0,5), fecha si es mm/dd/aaaa válida, y texto en lo
'     demás. (*) Excel lee las fechas de los criterios a la AMERICANA aunque el libro esté en español:
'     "<=12/31/2026" deja pasar el 31/12/2026. Solo comprobado con 01/01 y 12/31; una fecha que se pueda
'     leer de las dos maneras (05/03) no se ha probado.
'   - Un número nunca es igual, mayor ni menor que un texto (sí distinto). Un texto frente a un operando
'     número: (*) con un criterio sin operador (la celda 2026) cuenta como igual si es ese texto ("2026"
'     sí, "02026" no); con operador ("<>2026") nunca es igual.
'   - (*) Celda vacía y texto vacío "" (una fórmula ="") son distintos: "=" solo deja pasar únicamente las
'     vacías. La vacía solo cumple "=" solo y "<>x"; el texto vacío se trata como un texto.
'   - AutoFilter (Fnc_Filtro_Filas): igual salvo que (*) "=" solo también deja pasar el texto vacío "", y
'     (*) "=1" deja pasar los textos "1" y " 1" (sin contar los espacios), pero no "01".
'===================================================================================================

'- Una condición: una celda del rango de criterios, o el criterio de un AutoFilter -----------------
Private Type T_Cond
    Col         As Long         ' columna de la tabla (mismos índices que T.Datos)
    Op          As String       ' "=", "<>", "<", ">", "<=", ">=", o "Empieza" (texto sin operador)
    Vacio       As Boolean      ' operando vacío: "=" o "<>" solos
    EsNum       As Boolean      ' el operando es un número (o una fecha)
    Num         As Double       ' el operando, si es número
    Txt         As String       ' el operando tal como se escribió
    Patron      As String       ' el operando en minúsculas, como patrón de Like (comodines traducidos)
    Auto        As Boolean      ' reglas del AutoFilter (si no, las del AdvancedFilter)
    Desnudo     As Boolean      ' criterio sin operador (la celda 2026, o el texto "M013")
End Type

'===================================================================================================
'- Filas que deja visibles .AdvancedFilter xlFilterInPlace con el rango de criterios NombreCriT ----
Public Function Fnc_CriT_Filas(T As T_TablaRam, ByVal NombreCriT As String, _
                               Optional Vivas As Variant) As Boolean()
    Dim Criterios   As Variant
    Dim ColTabla()  As Long
    Dim Conds()     As T_Cond
    Dim NumConds    As Long
    Dim Viva()      As Boolean
    Dim Cumple()    As Boolean
    Dim Cand()      As Boolean
    Dim FilaCrit    As Long
    Dim ColCrit     As Long
    Dim K           As Long
    Dim C           As Long
    Dim Fila        As Long

    Criterios = ThisWorkbook.Names(NombreCriT).RefersToRange.Value     '- Al menos 2 filas: siempre es un array 2D
    ReDim ColTabla(1 To UBound(Criterios, 2))
    For ColCrit = 1 To UBound(Criterios, 2)
        ColTabla(ColCrit) = Fnc_CriT_Columna(T, Criterios(1, ColCrit), NombreCriT)
    Next ColCrit

    Viva = Fnc_Filas_Vivas(T, Vivas)
    ReDim Cumple(LBound(Viva) To UBound(Viva))
    ReDim Cand(LBound(Viva) To UBound(Viva))
    ReDim Conds(1 To UBound(Criterios, 2))
    For FilaCrit = 2 To UBound(Criterios, 1)                        '- Cada fila de criterios: una alternativa (O)
        NumConds = 0
        For ColCrit = 1 To UBound(Criterios, 2)
            If Fnc_Cond_Leer(Criterios(FilaCrit, ColCrit), ColTabla(ColCrit), Conds(NumConds + 1), NombreCriT) Then
                NumConds = NumConds + 1
            End If
        Next ColCrit
        For Fila = 1 To T.NumFilas                                  '- Candidatas: vivas que aún no cumplen
            Cand(Fila) = Viva(Fila) And Not Cumple(Fila)
        Next Fila
        For K = 1 To NumConds                                       '- Sus condiciones, todas a la vez (Y)
            C = Conds(K).Col
            For Fila = 1 To T.NumFilas
                If Cand(Fila) Then Cand(Fila) = Fnc_Cond_Cumple(T.Datos(Fila, C), Conds(K))
            Next Fila
        Next K
        For Fila = 1 To T.NumFilas
            If Cand(Fila) Then Cumple(Fila) = True
        Next Fila
    Next FilaCrit
    Fnc_CriT_Filas = Cumple
End Function    ' Fnc_CriT_Filas
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Filas que deja visibles .AutoFilter Field:=Col, Criteria1:=Op & Operando ------------------------
'-   Op: "=", "<>", "<", ">", "<=" o ">=". Operando: número, fecha o texto ("" con "=" o "<>": vacía /
'-   no vacía). Las fechas, mejor como fecha o número que como texto: así no depende de cómo se escriban.
Public Function Fnc_Filtro_Filas(T As T_TablaRam, ByVal Col As Long, ByVal Op As String, _
                                 ByVal Operando As Variant, Optional Vivas As Variant) As Boolean()
    Dim Cond        As T_Cond
    Dim Cumple()    As Boolean
    Dim Fila        As Long

    If Col < 1 Or Col > T.NumCols Then
        Err.Raise vbObjectError + 520, "Fnc_Filtro_Filas", "La columna " & Col & " no existe en la tabla (" & T.NumCols & " columnas)."
    End If
    If Not T.Cargada(Col) Then
        Err.Raise vbObjectError + 521, "Fnc_Filtro_Filas", "La columna " & Col & " (" & T.Titulos(Col) & ") no se ha cargado en RAM."
    End If
    Select Case Op
        Case "=", "<>", "<", ">", "<=", ">="
        Case Else
            Err.Raise vbObjectError + 524, "Fnc_Filtro_Filas", "Operador no válido: '" & Op & "'."
    End Select
    Call Rut_Cond_Montar(Cond, Col, Op, Operando, True)
    Cumple = Fnc_Filas_Vivas(T, Vivas)
    For Fila = 1 To T.NumFilas
        If Cumple(Fila) Then Cumple(Fila) = Fnc_Cond_Cumple(T.Datos(Fila, Col), Cond)
    Next Fila
    Fnc_Filtro_Filas = Cumple
End Function    ' Fnc_Filtro_Filas
'---------------------------------------------------------------------------------------------------

'- Cuántas filas valen True ------------------------------------------------------------------------
Public Function Fnc_Filas_Contar(Filas() As Boolean) As Long
    Dim Fila    As Long
    For Fila = 1 To UBound(Filas)
        If Filas(Fila) Then Fnc_Filas_Contar = Fnc_Filas_Contar + 1
    Next Fila
End Function
'---------------------------------------------------------------------------------------------------

'- Borrado lógico: pone a False en Vivas las filas que cumplen y aún vivían; devuelve cuántas y ----
'- las descuenta de NVivas. (Era Fnc_Borrar_Filas de M_111; desde el 2026-10-06 la usa también M_211)
Public Function Fnc_Filas_Borrar(Vivas() As Boolean, Cumple() As Boolean, NVivas As Long) As Long
    Dim Fila    As Long
    For Fila = 1 To UBound(Vivas)
        If Cumple(Fila) And Vivas(Fila) Then
            Vivas(Fila) = False
            Fnc_Filas_Borrar = Fnc_Filas_Borrar + 1
        End If
    Next Fila
    NVivas = NVivas - Fnc_Filas_Borrar
End Function
'---------------------------------------------------------------------------------------------------

'- Vivas como array (todas True si no se pasa) -----------------------------------------------------
Private Function Fnc_Filas_Vivas(T As T_TablaRam, Vivas As Variant) As Boolean()
    Dim Viva()  As Boolean
    Dim Fila    As Long
    If IsMissing(Vivas) Then
        If T.NumFilas = 0 Then ReDim Viva(0 To 0) Else ReDim Viva(1 To T.NumFilas)
        For Fila = 1 To T.NumFilas
            Viva(Fila) = True
        Next Fila
    Else
        Viva = Vivas
        If UBound(Viva) <> T.NumFilas Then
            Err.Raise vbObjectError + 525, "Rut_Lo_CriT_Ram", "Vivas tiene " & UBound(Viva) & " filas y la tabla " & T.NumFilas & "."
        End If
    End If
    Fnc_Filas_Vivas = Viva
End Function
'---------------------------------------------------------------------------------------------------

'- Columna de la tabla con ese título (sin distinguir mayúsculas), que tiene que estar cargada -----
Private Function Fnc_CriT_Columna(T As T_TablaRam, ByVal Titulo As Variant, ByVal NombreCriT As String) As Long
    Dim C       As Long
    For C = 1 To T.NumCols
        If StrComp(CStr(T.Titulos(C)), CStr(Titulo), vbTextCompare) = 0 Then
            If Not T.Cargada(C) Then
                Err.Raise vbObjectError + 521, "Fnc_CriT_Filas", "El criterio " & NombreCriT & " usa la columna " & _
                          Titulo & " (" & C & "), que no se ha cargado en RAM."
            End If
            Fnc_CriT_Columna = C
            Exit Function
        End If
    Next C
    Err.Raise vbObjectError + 522, "Fnc_CriT_Filas", "El criterio " & NombreCriT & " usa la columna '" & Titulo & _
              "', que no está en la tabla."
End Function
'---------------------------------------------------------------------------------------------------

'- Lee una celda del rango de criterios. False si está vacía (sin condición) -----------------------
Private Function Fnc_Cond_Leer(ByVal Valor As Variant, ByVal Col As Long, Cond As T_Cond, _
                               ByVal NombreCriT As String) As Boolean
    Dim S       As String
    Dim Op      As String

    If IsEmpty(Valor) Then Exit Function
    If IsError(Valor) Or VarType(Valor) = vbBoolean Then
        Err.Raise vbObjectError + 523, "Fnc_CriT_Filas", "El criterio " & NombreCriT & " tiene una celda con un error " & _
                  "o un valor lógico, que no se sabe tratar."
    End If
    If VarType(Valor) <> vbString Then                              '- Número o fecha: igual a ese valor
        Call Rut_Cond_Montar(Cond, Col, "=", Valor, False)
        Cond.Desnudo = True
        Fnc_Cond_Leer = True
        Exit Function
    End If
    S = Valor
    If Len(S) = 0 Then Exit Function                                '- Texto vacío: sin condición
    Op = Fnc_Operador(S)
    If Op = "" Then                                                 '- Sin operador: el texto, empieza por;
        Call Rut_Cond_Montar(Cond, Col, "=", S, False)              '-   un número escrito como texto, igual
        If Not Cond.EsNum Then Cond.Op = "Empieza"
        Cond.Desnudo = True
    Else
        Call Rut_Cond_Montar(Cond, Col, Op, Mid$(S, Len(Op) + 1), False)
    End If
    Fnc_Cond_Leer = True
End Function
'---------------------------------------------------------------------------------------------------

'- Operador con el que empieza el texto ("" si no lleva) -------------------------------------------
Private Function Fnc_Operador(ByVal S As String) As String
    Dim Op      As Variant
    For Each Op In Array("<=", ">=", "<>", "=", "<", ">")
        If Left$(S, Len(Op)) = Op Then
            Fnc_Operador = Op
            Exit Function
        End If
    Next Op
End Function
'---------------------------------------------------------------------------------------------------

'- Rellena la condición con su operador y su operando (Auto: con las reglas del AutoFilter) --------
Private Sub Rut_Cond_Montar(Cond As T_Cond, ByVal Col As Long, ByVal Op As String, ByVal Operando As Variant, _
                            ByVal Auto As Boolean)
    Dim D       As Double
    Cond.Col = Col
    Cond.Op = Op
    Cond.Auto = Auto
    Cond.Desnudo = False
    Cond.Vacio = False
    Cond.EsNum = False
    Cond.Num = 0
    Cond.Patron = ""
    If VarType(Operando) = vbString Then
        Cond.Txt = Operando
        If Len(Cond.Txt) = 0 Then
            Cond.Vacio = True
        ElseIf Fnc_Texto_Numero(Cond.Txt, D) Then
            Cond.EsNum = True
            Cond.Num = D
        ElseIf Fnc_Texto_Fecha(Cond.Txt, D) Then
            Cond.EsNum = True
            Cond.Num = D
        Else
            Cond.Patron = Fnc_Comodines_Like(LCase$(Cond.Txt))
        End If
    Else
        Cond.EsNum = True
        Cond.Num = CDbl(Operando)
        Cond.Txt = CStr(Operando)
    End If
End Sub
'---------------------------------------------------------------------------------------------------

'- ¿Cumple el valor V de la celda la condición? ----------------------------------------------------
Private Function Fnc_Cond_Cumple(V As Variant, Cond As T_Cond) As Boolean
    Dim Tipo    As Long         '- 0 vacía, 1 número, 2 texto, 3 otra cosa (error, valor lógico)
    Dim X       As Double
    Dim Rel     As Long         '- -1, 0 o 1: la celda es menor, igual o mayor que el operando

    Select Case VarType(V)
        Case vbEmpty
            Tipo = 0
        Case vbString                                               '- El texto vacío "": vacía para el
            If Len(V) = 0 And Cond.Auto Then Tipo = 0 Else Tipo = 2 '-   AutoFilter, texto para el AdvancedFilter
        Case vbDouble, vbCurrency, vbDate, vbLong, vbInteger, vbSingle, vbDecimal, vbByte
            Tipo = 1
        Case Else
            Tipo = 3
    End Select

    If Cond.Vacio Then                                              '- "=" / "<>" solos: vacía / no vacía
        If Cond.Op = "=" Then
            Fnc_Cond_Cumple = (Tipo = 0)
        ElseIf Cond.Op = "<>" Then
            Fnc_Cond_Cumple = (Tipo <> 0)
        End If
        Exit Function
    End If
    If Tipo = 0 Or Tipo = 3 Then                                    '- Vacía (u otra cosa): solo "distinto de x"
        Fnc_Cond_Cumple = (Cond.Op = "<>")
        Exit Function
    End If

    If Cond.EsNum Then
        If Tipo = 2 Then                                            '- Texto frente a número
            If Cond.Auto Then                                       '- AutoFilter: "=1" deja pasar "1" y " 1"
                Rel = StrComp(Trim$(V), Cond.Txt, vbTextCompare)
                If Cond.Op = "=" Then
                    Fnc_Cond_Cumple = (Rel = 0)
                ElseIf Cond.Op = "<>" Then
                    Fnc_Cond_Cumple = (Rel <> 0)
                End If
            ElseIf Cond.Desnudo Then                                '- Criterio sin operador: el texto "2026"
                Fnc_Cond_Cumple = (StrComp(V, Cond.Txt, vbTextCompare) = 0)
            Else                                                    '- Con operador: nunca igual
                Fnc_Cond_Cumple = (Cond.Op = "<>")
            End If
            Exit Function
        End If
        X = V
        If X < Cond.Num Then
            Rel = -1
        ElseIf X > Cond.Num Then
            Rel = 1
        End If
    Else
        If Tipo = 1 Then                                            '- Número frente a texto: solo "distinto"
            Fnc_Cond_Cumple = (Cond.Op = "<>")
            Exit Function
        End If
        Select Case Cond.Op
            Case "Empieza"
                Fnc_Cond_Cumple = (LCase$(V) Like Cond.Patron & "*")
                Exit Function
            Case "="
                Fnc_Cond_Cumple = (LCase$(V) Like Cond.Patron)
                Exit Function
            Case "<>"
                Fnc_Cond_Cumple = Not (LCase$(V) Like Cond.Patron)
                Exit Function
        End Select
        Rel = StrComp(V, Cond.Txt, vbTextCompare)
    End If

    Select Case Cond.Op
        Case "=":   Fnc_Cond_Cumple = (Rel = 0)
        Case "<>":  Fnc_Cond_Cumple = (Rel <> 0)
        Case "<":   Fnc_Cond_Cumple = (Rel < 0)
        Case ">":   Fnc_Cond_Cumple = (Rel > 0)
        Case "<=":  Fnc_Cond_Cumple = (Rel <= 0)
        Case ">=":  Fnc_Cond_Cumple = (Rel >= 0)
    End Select
End Function    ' Fnc_Cond_Cumple
'---------------------------------------------------------------------------------------------------

'- Entero o decimal con coma y signo opcional (2026, -1, 0,5), como lo lee Excel en español --------
Private Function Fnc_Texto_Numero(ByVal S As String, D As Double) As Boolean
    Dim Cuerpo  As String
    Dim P       As Long
    Cuerpo = S
    If Left$(Cuerpo, 1) = "-" Or Left$(Cuerpo, 1) = "+" Then Cuerpo = Mid$(Cuerpo, 2)
    If Len(Cuerpo) = 0 Then Exit Function
    P = InStr(Cuerpo, ",")
    If P = 0 Then
        If Cuerpo Like "*[!0-9]*" Then Exit Function
    Else
        If P = 1 Or P = Len(Cuerpo) Then Exit Function
        If Left$(Cuerpo, P - 1) Like "*[!0-9]*" Then Exit Function
        If Mid$(Cuerpo, P + 1) Like "*[!0-9]*" Then Exit Function
    End If
    D = Val(Replace(S, ",", "."))                                   '- Val usa siempre el punto decimal
    Fnc_Texto_Numero = True
End Function
'---------------------------------------------------------------------------------------------------

'- Fecha mm/dd/aaaa (o m-d-aa): como lee Excel las fechas de los criterios, aunque el libro esté en español
Private Function Fnc_Texto_Fecha(ByVal S As String, D As Double) As Boolean
    Dim Partes  As Variant
    Dim Dia     As Long
    Dim Mes     As Long
    Dim Ano     As Long
    Partes = Split(Replace(S, "-", "/"), "/")
    If UBound(Partes) <> 2 Then Exit Function
    If Len(Partes(0)) < 1 Or Len(Partes(0)) > 2 Or Partes(0) Like "*[!0-9]*" Then Exit Function
    If Len(Partes(1)) < 1 Or Len(Partes(1)) > 2 Or Partes(1) Like "*[!0-9]*" Then Exit Function
    If (Len(Partes(2)) <> 2 And Len(Partes(2)) <> 4) Or Partes(2) Like "*[!0-9]*" Then Exit Function
    Mes = CLng(Partes(0))
    Dia = CLng(Partes(1))
    Ano = CLng(Partes(2))
    If Len(Partes(2)) = 2 Then
        If Ano < 30 Then Ano = Ano + 2000 Else Ano = Ano + 1900
    End If
    If Mes < 1 Or Mes > 12 Then Exit Function
    If Dia < 1 Or Dia > Day(DateSerial(Ano, Mes + 1, 0)) Then Exit Function
    D = CDbl(DateSerial(Ano, Mes, Dia))
    Fnc_Texto_Fecha = True
End Function
'---------------------------------------------------------------------------------------------------

'- Comodines de Excel (* ? y ~ para escaparlos) a patrón de Like, donde # y [ también son comodines --
Private Function Fnc_Comodines_Like(ByVal S As String) As String
    Dim I       As Long
    Dim Ch      As String
    Dim R       As String
    I = 1
    Do While I <= Len(S)
        Ch = Mid$(S, I, 1)
        Select Case Ch
            Case "~"
                If I < Len(S) And InStr("*?~", Mid$(S, I + 1, 1)) > 0 Then
                    I = I + 1
                    If Mid$(S, I, 1) = "~" Then R = R & "~" Else R = R & "[" & Mid$(S, I, 1) & "]"
                Else
                    R = R & "~"
                End If
            Case "[", "#"
                R = R & "[" & Ch & "]"
            Case Else
                R = R & Ch
        End Select
        I = I + 1
    Loop
    Fnc_Comodines_Like = R
End Function
'---------------------------------------------------------------------------------------------------
