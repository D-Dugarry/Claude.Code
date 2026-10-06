Attribute VB_Name = "M_211_Del_Null_Reg_CAcadAnt"
' Last Rev. 2026-10-06 11:50
'2026-01-23
'- M_211_Remove_Null_Reg_CAcad
Option Explicit

'    - M_211, Filtrar y Borrar Registros NO deseados:
'        - Borrar Recibos AE4 Enseñanzas Propias
'        - Borrar Recibos de Movimiento Menos AE=80 'Pruebas Acceso UA'
'        - Borrar Recibos con Imp.Rec. < 0
'        - Borrar Recibos de Matrículas de coste CERO
'        - Borrar Borrar Recibos - Subvencionado-, Imp_Rec =0 porque Imp_Dto >0
'        - Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"
'        - Borrar Recibos BD_C_Acad <> C_Acad_Ant
'        - Borrar Recibos ANULADOS
'        - Borrar Recibos NO Martrícula
'        - Borrar Recibos INVALIDADOS

            Sub RuT_Remove_Null_Reg_ByHand()
                
                Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
                Dim Lo_BD_Ant           As ListObject:      Set Lo_BD_Ant = Sht__BD_Ant.ListObjects(1)
                Dim Lo_BD_Dpl           As ListObject:      Set Lo_BD_Dpl = Sht__BD_Dupl.ListObjects(1)
                Dim Lo_BD_ErrDate           As ListObject:      Set Lo_BD_ErrDate = Sht__BD_ErrDate.ListObjects(1)
                Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
                
                Call RuT_Remove_Null_Reg_CAcad(ActiveSheet.ListObjects(1))
                
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Filtrar y Borrar Registros NO deseados -------------------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Remove_Null_Reg_CAcad(Lo_G04 As ListObject)
'- Desde el 2026-10-06 trabaja en RAM, como M_111 (paso a RAM): cada paso filtra la copia en memoria de la tabla
'- (Rut_Lo_CriT_Ram, con las mismas reglas que los filtros de Excel) y borra solo en RAM (borrado lógico: Viva). Después, en
'- la hoja: la marca de Obs_Conta, UNA ordenación (con el mismo resultado que las que hacía cada paso antes de borrar: M_212
'- se queda con el último duplicado de cada Ref, así que el orden importa) que deja juntas al final las filas a borrar, y un
'- solo borrado. Antes se filtraba, contaba y borraba en la hoja en cada paso.
Debug.Print ">>> RuT_Remove_Null_Reg_CAcad"
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim T               As T_TablaRam
    Dim Viva()          As Boolean          '- La fila sigue en la tabla (las demás ya están borradas en RAM)
    Dim Cumple()        As Boolean
    Dim Tandas()        As Variant          '- Ordenaciones que hacía cada paso, en su orden
    Dim NVivas          As Long
    Dim NBorrar         As Long
    Dim rowfind         As Long
    Dim Fila            As Long
    Dim Sin_Registros   As Boolean          '- A la tabla no le quedan recibos

    Call Rut_Lo_Filtros_Quitar(Lo_G04)                '- Quitar filtros
    Lo_G04.ShowTotals = False
    If Lo_G04.DataBodyRange Is Nothing Then GoTo Terminar
    Call Rut_TablaRam_Cargar(T, Lo_G04, Array(BD_ActivEco, BD_RegMov, BD_ImpRec, BD_ImpDto, BD_DNI, BD_C_Acad, BD_Anul, _
                             BD_Matricula, BD_Hinvalid, BD_Obs_Conta), True)
    NVivas = T.NumFilas
    ReDim Viva(1 To T.NumFilas)
    For Fila = 1 To T.NumFilas
        Viva(Fila) = True
    Next Fila

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos AE4 Enseñanzas Propias -----------------------------------------------------------------------------------
    If Prog__APP_Switch.Range("Sw_DelRegAE4") Then
        Cumple = Fnc_Filtro_Filas(T, BD_ActivEco, "=", 4, Viva)
        rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
        Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_ActivEco))
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. AE4 ", 0, _
                                                            Format(rowfind, " #,##0") & " reg. ", _
                                                            " de " & Format(NVivas, "#,##0") & " reg.")
            If NVivas = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. AE4 ", 0)
        End If
    Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Se ha decidido mantener los Rec. AE4", 0)
    End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos de Movimiento Menos AE=80 'Pruebas Acceso UA' -----------------------------------------------------
    Cumple = Fnc_Filtro_Filas(T, BD_RegMov, "=", "S", Viva)
    Cumple = Fnc_Filtro_Filas(T, BD_ActivEco, "<>", 80, Cumple)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_RegMov, BD_ActivEco))
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec.Mov. Menos AE=80 'Pruebas Acceso UA'", 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        If NVivas = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de Movimiento. ", 0)
    End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos con Imp.Rec. < 0 -----------------------------------------------------------------------------------
    Cumple = Fnc_Filtro_Filas(T, BD_ImpRec, "<", 0, Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_ImpRec))
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. con Imp.Rec. < 0. ", 0, _
                                                        Format(rowfind, " #,##0") & " reg. ", _
                                                        " de " & Format(NVivas, "#,##0") & " reg.")
        If NVivas = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Negativos. ", 0)
    End If

'- ---------------------------------------------------------------------------------------------------------------
'- Borrar Recibos de Matrículas de coste CERO --------------------------------------------------------------------
'--------- El importe del recibo es Cero, el importe Académico es Cero, y el importe Administrativo es Cero. -----
'--------- Es decir que la Matrícula del estudio es gratuita, no tiene coste alguno. -----------------------------
'- ---------------------------------------------------------------------------------------------------------------
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_ImpRec, BD_ImpDto))
    Cumple = Fnc_CriT_Filas(T, "Tb_CriT_ImpMatCero", Viva)
    rowfind = Fnc_Filas_Contar(Cumple)
    If rowfind > 0 Then
        Call Fnc_Filas_Borrar(Viva, Cumple, NVivas)
        '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec. de Matrícula_Cero ", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de Matrícula_Cero", 0)
    End If

'- ---------------------------------------------------------------------------------------------------------------
'- Borrar Borrar Recibos - Subvencionado-, Imp_Rec =0 porque Imp_Dto >0 ------------------------------
'- ---------------------------------------------------------------------------------------------------------------
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_ImpRec))
    Cumple = Fnc_Filtro_Filas(T, BD_ImpRec, "=", 0, Viva)
    rowfind = Fnc_Filas_Contar(Cumple)
    If rowfind > 0 Then
        If Prog__APP_Switch.Range("Sw_DelRegMatriculaCero") Then
            Call Fnc_Filas_Borrar(Viva, Cumple, NVivas)
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec. de Matrícula_Cero Subvencionada", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        Else
            Call Rut_TablaRam_Marcar(T, Cumple, BD_Obs_Conta, "ImpMatCero")
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Marcados Rec. de Matrícula_Cero Subvencionada", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        End If
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de Matrícula_Cero Subvencionada", 0)
    End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"  --------------------------------------
    Cumple = Fnc_Filtro_Filas(T, BD_DNI, "=", "1", Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_DNI))
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Ficticios, DNI=1", 0, _
                                                        Format(rowfind, " #,##0") & " reg. ", _
                                                        " de " & Format(NVivas, "#,##0") & " reg.")
        If NVivas = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Ficticios, DNI=1", 0)
    End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos BD_C_Acad <> C_Acad_Ant ---------------------------------------------------------------------------
    Cumple = Fnc_Filtro_Filas(T, BD_C_Acad, "<>", C_Acad_Ant, Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_C_Acad))
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. de C_Acad. <> " & C_Acad_Ant, 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        If NVivas = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de C_Acad. <> " & C_Acad_Ant, 0)
    End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos ANULADOS ---------------------------------------------------------------------------------------
    Cumple = Fnc_Filtro_Filas(T, BD_Anul, "=", "S", Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_Anul))
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Anulados ", 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        If NVivas = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Anulados ", 0)
    End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos NO Martrícula ----------------------------------------------------------------------------------
    Cumple = Fnc_Filtro_Filas(T, BD_Matricula, "=", "N", Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_Matricula))
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. NO Matrícula ", 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        If NVivas = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. NO Matrícula ", 0)
    End If

'- ------------------------------------------------------------------------------------------------------------------
'- Borrar Recibos INVALIDADOS ------------------------------------------------------------------------------------
    Cumple = Fnc_Filtro_Filas(T, BD_Hinvalid, "=", "S", Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_Hinvalid))
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Invalidados ", 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        If NVivas = 0 Then GoTo Proceso_Finalizado_por_quedarse_sin_Registros
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Invalidados ", 0)
    End If
    GoTo Aplicar_en_la_Hoja

Proceso_Finalizado_por_quedarse_sin_Registros:
    Sin_Registros = True

'- -------------------------------------------------------------------------------------------------
'- En la hoja: la marca de Obs_Conta, una ordenación y un solo borrado -----------------------------
'- -------------------------------------------------------------------------------------------------
Aplicar_en_la_Hoja:
    NBorrar = T.NumFilas - NVivas
    If NVivas = 0 Then                          '- No queda ninguno: se borran todos, sin ordenar
        Lo_G04.DataBodyRange.Delete
    Else
        If NBorrar > 0 Then                     '- Columna auxiliar Tipo_Rec (vacía tras importar): 1 = a borrar
            For Fila = 1 To T.NumFilas
                If Viva(Fila) Then T.Datos(Fila, BD_Tipo_Rec) = 0 Else T.Datos(Fila, BD_Tipo_Rec) = 1
            Next Fila
            T.Modificada(BD_Tipo_Rec) = True
        End If
        Call Rut_TablaRam_Volcar(T, Lo_G04)     '- Obs_Conta (si hay marcas) y la columna auxiliar
        Erase T.Datos
        If NBorrar > 0 Then                     '- Las filas a borrar quedan juntas al final
            Call Rut_TablaRam_Ordenar_Tandas(Lo_G04, Tandas, BD_Tipo_Rec)
            Lo_G04.ListColumns(BD_Tipo_Rec).DataBodyRange.ClearContents
            Lo_G04.DataBodyRange.Rows(NVivas + 1).Resize(NBorrar).Delete Shift:=xlUp
        Else
            Call Rut_TablaRam_Ordenar_Tandas(Lo_G04, Tandas)
        End If
    End If

    If Sin_Registros Then
        MsgBx_Title = "Proceso: Importar LsGES04 C_Acad_Ant para ImpAdm"
        MsgBx_Msg = "¡ A la tabla Lo_G04 no le quedan Recibos procesables !"
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", MsgBx_Msg & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm"), 0)
        Prog__APP.Range("APP_Task_Inf") = ActivForm.Controls("TBx_Informe") & vbLf & vbLf & MsgBx_Msg & Now()
        Exit Sub
    End If

Terminar:
    Call Rut_Lo_Filtros_Quitar(Lo_G04)
Debug.Print "<<< RuT_Remove_Null_Reg_CAcad"

End Sub     ' RuT_Remove_Null_Reg_CAcad



