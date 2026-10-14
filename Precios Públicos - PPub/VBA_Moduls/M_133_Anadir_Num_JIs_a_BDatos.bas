Attribute VB_Name = "M_133_Anadir_Num_JIs_a_BDatos"
' Last Rev. 2026-10-14 09:39
'2026-02-14
Option Explicit

'- Añadir los Números de JI's a BDatos, tomándolos de la Tabla Informe_Recibos. ------------------------
'- 2026-10-13: con control de errores (Paso, Rut_Off_Functions/Rut_On_Functions) y BDatos protegida al acabar como estaba
'  (Rut_Prot_Save/Rut_Prot_Restore). Se miran las 6 columnas de JI's (antes solo las 4 primeras: una fila con solo JI 4430 se
'  saltaba), y JI_Sin_Num ("·-·", recibos sin JI) cuenta como vacío: ya no se copia a BDatos como si fuera un JI. El resultado
'  queda en App_Task_Inf. Los JI's de BDatos al informe los pone ya M_130 al calcularlo (antes, M_132).
'- 2026-10-14: motor común de los dos informes (Docs/Plan_Inf_Conta_TIO.md). Rut_Inf_JIs_a_BDatos(Def) copia los JI's del
'  informe que describe Def (T_Inf_Def, M_130): Inf_Recibos_TIO (aquí) o Inf_Conta_TIO (M_141). Los dos, a las mismas 6
'  columnas de BDatos (BD_JI_Emi_Adm..BD_JI_443_Acad, AR:AW).

' ==================================================================================================================================
Sub Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos()    '- botón «JI's del Inf. a BD» en Inf_Recibos_TIO
' ==================================================================================================================================
    Dim Def         As T_Inf_Def
    Call Rut_Inf_Def_Recibos(Def)                   '- M_130
    Call Rut_Inf_JIs_a_BDatos(Def)
End Sub     '- Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos
' ==================================================================================================================================
'- Copia a BDatos los JI's del informe que describe Def (hasta el 2026-10-14, el cuerpo de Rut_Anadir_a_BDatos_JIs_de_Inf_Recibos)
Sub Rut_Inf_JIs_a_BDatos(Def As T_Inf_Def)
' ==================================================================================================================================
Debug.Print "Rut_Inf_JIs_a_BDatos: " & Def.Ws.Name

        Dim Sht_Inf         As Worksheet:       Set Sht_Inf = Def.Ws
        Dim Col_TipRec      As Integer:         Col_TipRec = Def.Col(InfRec_TipRec)         '- columnas de este informe
        Dim Col_ConcptEco   As Integer:         Col_ConcptEco = Def.Col(InfRec_ConcptEco)
        Dim Col_JI          As Integer:         Col_JI = Def.Col(InfRec_JI_Emi_Adm)         '- la primera de las 6 de JI's
        Dim Lo_Inf          As ListObject:      Set Lo_Inf = Sht_Inf.ListObjects(1)
        Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)

Dim Paso            As String:          Paso = "preparar las hojas"     '- para el aviso si algo falla
Dim Msg_Error       As String
Dim Off_Hecho       As Boolean
Dim Prot_BD         As T_Prot_Estado                                    '- protección de BDatos: Save al empezar, Restore en Fin
Dim Prot_Hecho      As Boolean
Dim T_Ini           As Double:          T_Ini = Timer
Dim Tp_Rec          As String
Dim Concept         As String
Dim F_Inf           As Integer
Dim K               As Integer
Dim rowfind         As Variant
Dim RwJI            As ListRow
Dim Valores         As Variant                                          '- (1, 1 To JI_N_Col): los JI's de la fila del informe
Dim Hay_JI          As Boolean
Dim N_Filas         As Long                                             '- filas del informe copiadas a BDatos
Dim N_Recibos       As Long                                             '- recibos de BDatos que las reciben
    On Error GoTo Fin

    Call Rut_Off_Functions:     Off_Hecho = True    '- cálculo manual, pantalla y eventos; los restaura Rut_On_Functions en Fin
    Application.DisplayAlerts = False

    Call Rut_Prot_Save(Sht__BD, Prot_BD):     Prot_Hecho = True     '- la desprotege; en Fin, protegida como estaba (antes se quedaba sin proteger)
    Sht_Inf.Visible = xlSheetVisible
    Sht_Inf.Select

    '------------ Preparo Sht__BD y Ordeno por Tipo_Tasa y Concepto_Económico -----------------------------------
    '- Para filtrar no hace falta, pero deja juntos los recibos de cada fila del informe y se escriben de una vez.
    '- Una sola ordenación con las dos columnas (Aplicar:=False en la primera): el mismo orden que las dos de antes.
    Call Rut_Lo_WrkSht_Preparar(Sht_Inf)
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Paso = "ordenar BDatos y el informe"
    Call Rut_Lo_Sort(Lo_BD, BD_Tipo_Rec, xlAscending, True, False)     '- Emitida, Aplazado, EjeAnt, ADxAplz, Añeja...
    Call Rut_Lo_Sort(Lo_BD, BD_Concepto, xlAscending, False)           '- 1303.00 1310.00 1311.00 Etc.
    Sht_Inf.Select
    Call Rut_Lo_Sort(Lo_Inf, Col_TipRec, xlAscending, True, False)  '- Emitida, Aplazado, EjeAnt, ADxAplz, Añeja...
    Call Rut_Lo_Sort(Lo_Inf, Col_ConcptEco, xlAscending, False)     '- 1303.00 1310.00 1311.00 Etc.

    '-------------- Leo la Tabla de JI's de Tasas por Tipo y por Concepto ------------------------------------
    '- Para cada Tipo de Recibos; Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. -----------------------------------
    For F_Inf = 1 To Lo_Inf.ListRows.Count
        Set RwJI = Lo_Inf.ListRows(F_Inf)
        Tp_Rec = RwJI.Range(Col_TipRec)
        Concept = "1" & RwJI.Range(Col_ConcptEco)
        Paso = "copiar los JI's de la fila " & F_Inf & " del informe (" & Tp_Rec & ", " & Concept & ")"
        If Left(Concept, 4) = "1311" Then GoTo Siguiente_Concepto                                       '- ¡¡ No podemos hacer de Enseñanzas Propias !!
        '- Los 6 JI's de la fila; los que no son dato (vacíos, blancos, errores y JI_Sin_Num) se escriben vacíos ---------------
        Valores = RwJI.Range(Col_JI).Resize(1, JI_N_Col).Value
        Hay_JI = False
        For K = 1 To JI_N_Col
            If Fnc_JI_Con_Dato(Valores(1, K)) Then Hay_JI = True Else Valores(1, K) = Empty
        Next K
        If Not Hay_JI Then GoTo Siguiente_Concepto      '- No hay JI's ¡¡ En ninguna columna !!
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '- Filtrar Tipo_Rec y Concepto
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="=" & Tp_Rec
        Lo_BD.Range.AutoFilter Field:=BD_Concepto, Criteria1:="=" & Val(Concept)
        rowfind = Fnc_Lo_Contar_Visibles(Lo_BD, BD_Ref)   '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            For K = 1 To JI_N_Col                         '- BD_JI_Emi_Adm..BD_JI_443_Acad, en el mismo orden que en el informe
                Lo_BD.ListColumns(BD_JI_Emi_Adm + K - 1).DataBodyRange.SpecialCells(xlCellTypeVisible).Cells.Value = Valores(1, K)
            Next K
            N_Filas = N_Filas + 1
            N_Recibos = N_Recibos + rowfind
        End If
Siguiente_Concepto:
    Next F_Inf

    Paso = "quitar los filtros de BDatos"
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    Lo_BD.ShowTotals = True
    Lo_Inf.ShowTotals = True

Fin:
    If Err.Number <> 0 Then
        Msg_Error = "Error " & Err.Number & " al " & Paso & ":" & vbLf & Err.Description
        Debug.Print ">>> Rut_Inf_JIs_a_BDatos -> ERROR " & Err.Number & " al " & Paso & ": " & Err.Description
    End If
    On Error GoTo -1                                '- sale del manejador de errores, para que el Resume Next de abajo valga
    On Error Resume Next
    If Msg_Error <> "" Then                         '- salida por error: BDatos sin filtros y las filas de totales visibles
        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        Call Rut_Lo_Totales_Mostrar
    End If
    If Prot_Hecho Then Call Rut_Prot_Restore(Sht__BD, Prot_BD, True)    '- después de Rut_Lo_Totales_Mostrar, que necesita BDatos sin proteger
    Application.DisplayAlerts = True
    If Msg_Error <> "" Then
        Prog__APP.Range("App_Task_Inf") = Format(Now, "dd/mm/yyyy hh:mm") & "  JI's de " & Sht_Inf.Name & " a BDatos: " & Replace(Msg_Error, vbLf, " ")
    Else
        Prog__APP.Range("App_Task_Inf") = Format(Now, "dd/mm/yyyy hh:mm") & "  JI's de " & Sht_Inf.Name & " a BDatos: " & N_Filas & _
                                          " filas del informe, " & Format(N_Recibos, "#,##0") & " recibos (" & Format(Timer - T_Ini, "0.0") & " s)"
    End If
    On Error GoTo 0
    If Off_Hecho Then Call Rut_On_Functions
    Debug.Print "   M_133: " & N_Filas & " filas del informe, " & N_Recibos & " recibos, " & Format(Timer - T_Ini, "0.00") & " s"
    If Msg_Error <> "" Then                         '- el aviso, ya con la pantalla activa
        MsgBox Msg_Error & vbLf & vbLf & "Puede que solo se hayan copiado a BDatos los JI's de algunas filas: vuelve a pasarlos.", _
               vbExclamation, "JI's de " & Sht_Inf.Name & " a BDatos"
    Else
        Application.Speech.Speak "Proceso completado puede verificar el resultado.", True
    End If
End Sub     '- Rut_Inf_JIs_a_BDatos
'--------------------------------------------------------------------------------------------------------------------------------------
