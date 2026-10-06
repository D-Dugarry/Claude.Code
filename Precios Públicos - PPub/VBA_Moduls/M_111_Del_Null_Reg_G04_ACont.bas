Attribute VB_Name = "M_111_Del_Null_Reg_G04_ACont"
' Last Rev. 2026-10-06 11:50
'Rev.: 2026-01-22
Option Explicit

'    - M_111, Filtrar y Borrar Registros NO deseados:
'        - Borrar Recibos AE4 Enseñanzas Propias
'        - Borrar Recibos con Imp.Rec. < 0
'        - Borrar Recibos de Matrículas de coste CERO ==> Imp_Rec = Imp_Dto = 0 ó Vacío (Tb_CriT_ImpMatCero)
'        - Borrar Recibos Importe CERO - Subvencionado- Imp_Rec =0 porque Imp_Dto >0
'        - Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"
'        - Borrar Recibos con AE=300 (M013 mal matriculado en Gestión Académica) ==>> Seminario orientación para preparación pruebas para mayores de 25 años.
'        - Borrar Recibos ANULADOS
'        - Borrar Recibos NO Martrícula
'        - Borrar Recibos INVALIDADOS
'        - Filtra Recibos con F_Cob > APP_FechCierreCont y Limpia-Clear las Columnas BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag y BD_HTipCob
'        - Borrar Recibos con F_Emi > APP_FechCierreCont "FUERA DEL PERÍODO CONTABLE"
'        - Borrar Incongruencias de Fechas, Recibos Cobrados en Años Anteriores o Posteriores a AnoCont

            Sub RuT_Remove_Null_Reg_ByHand()
                
                Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
                Dim Lo_BD_Ant           As ListObject:      Set Lo_BD_Ant = Sht__BD_Ant.ListObjects(1)
                Dim Lo_BD_Dpl           As ListObject:      Set Lo_BD_Dpl = Sht__BD_Dupl.ListObjects(1)
                Dim Lo_BD_ErrDate       As ListObject:      Set Lo_BD_ErrDate = Sht__BD_ErrDate.ListObjects(1)
                Dim Lo_BD_RegAnul       As ListObject:      Set Lo_BD_RegAnul = Sht__BD_RegAnul.ListObjects(1)
                Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
                
                Call RuT_Remove_Reg_No_Valid(Lo_BD, Lo_BD_ErrDate)
                
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Filtrar y Borrar Registros NO deseados: ------------------------------------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Remove_Reg_No_Valid(Lo_Data As ListObject, _
                        Optional Lo_BD_ErrDate As ListObject)
'- Desde el 2026-10-05 trabaja en RAM (fase 3 del paso a RAM): cada paso filtra la copia en memoria de la tabla
'- (Rut_Lo_CriT_Ram, con las mismas reglas que los filtros de Excel) y borra solo en RAM (borrado lógico: Viva).
'- Después, en la hoja: las marcas de Obs_Conta, los datos de cobro vaciados, UNA ordenación (con el mismo resultado
'- que las de antes, que hacía cada paso antes de borrar: M_112 se queda con el último duplicado de cada Ref, así
'- que el orden importa) que deja juntas al final las filas a borrar, la copia a BD_ErrDate y un solo borrado.
'- Antes se filtraba, contaba y borraba en la hoja en cada paso (13 s con 172.000 reg.).
Debug.Print ">>> RuT_Remove_Reg_No_Valid"
    Dim FechCierreCont  As String:      FechCierreCont = Prog__APP.Range("APP_FechCierreCont")
    Dim Cierre          As Double:      Cierre = CDbl(CDate(FechCierreCont))
    Dim Sh_Data         As Worksheet:   Set Sh_Data = Lo_Data.Parent
    Dim T               As T_TablaRam
    Dim Viva()          As Boolean          '- La fila sigue en la tabla (las demás ya están borradas en RAM)
    Dim Cumple()        As Boolean
    Dim Limpiar()       As Boolean          '- F_Cob posterior al cierre: se vacían los datos del cobro
    Dim ErrDate()       As Boolean          '- Fechas incongruentes: se copian a BD_ErrDate
    Dim Tandas()        As Variant          '- Ordenaciones que hacía cada paso, en su orden
    Dim NVivas          As Long
    Dim NBorrar         As Long
    Dim rowfind         As Long
    Dim Fila            As Long

    Lo_Data.ShowTotals = False
    If Lo_Data.DataBodyRange Is Nothing Then GoTo Terminar
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_TablaRam_Cargar(T, Lo_Data, Array(BD_ActivEco, BD_ImpRec, BD_ImpDto, BD_DNI, BD_Anul, BD_Matricula, _
                             BD_Hinvalid, BD_FCob, BD_FEmi, BD_ACont_Emi, BD_ACont_Cob, BD_Obs_Conta), True)
    NVivas = T.NumFilas
    ReDim Viva(1 To T.NumFilas)
    ReDim Limpiar(1 To T.NumFilas)
    ReDim ErrDate(1 To T.NumFilas)
    For Fila = 1 To T.NumFilas
        Viva(Fila) = True
    Next Fila

'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos AE4 Enseñanzas Propias -----------------------------------------------------------
    If Prog__APP_Switch.Range("Sw_DelRegAE4") Then
        Cumple = Fnc_Filtro_Filas(T, BD_ActivEco, "=", 4, Viva)
        rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
        Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_ActivEco))
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. AE4 ", 0, _
                                                            Format(rowfind, " #,##0") & " reg. ", _
                                                            " de " & Format(NVivas, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. AE4 ", 0)
        End If
    Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Se ha decidido mantener los Rec. AE4", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos con Imp.Rec. < 0 -----------------------------------------------------------------
    If Prog__APP_Switch.Range("Sw_DelRegNeg") Then
        Cumple = Fnc_Filtro_Filas(T, BD_ImpRec, "<", 0, Viva)
        rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
        Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_ImpRec))
        If rowfind > 0 Then
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Negativos. ", 0, _
                                                            Format(rowfind, " #,##0") & " reg. ", _
                                                            " de " & Format(NVivas, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Negativos. ", 0)
        End If
    Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Se ha decidido mantener los Rec. Negativos.", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos de Matrículas de Coste CERO ==> Imp_Rec = Imp_Dto = 0 ó Vacío (Tb_CriT_ImpMatCero) -------------
'--------- El importe del recibo es Cero, el importe Académico NO es Cero, o el importe Administrativo NO es Cero. -----
'--------- Es decir que la Matrícula del estudio es gratuita, aunque tiene coste. ------------------
'- -------------------------------------------------------------------------------------------------
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_ImpRec, BD_ImpDto))
    Cumple = Fnc_CriT_Filas(T, "Tb_CriT_ImpMatCero", Viva)
    rowfind = Fnc_Filas_Contar(Cumple)
    If rowfind > 0 Then
        If Prog__APP_Switch.Range("Sw_DelRegMatriculaCero") Then
            Call Fnc_Filas_Borrar(Viva, Cumple, NVivas)
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec. de Matrícula_Cero A Coste Cero", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        Else
            Call Rut_TablaRam_Marcar(T, Cumple, BD_Obs_Conta, "ImpMatCero")
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Marcados Rec. de Matrícula_Cero A Coste Cero", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        End If
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. de Matrícula_Cero A Coste Cero", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- Borrar Borrar Recibos - Subvencionado-, Imp_Rec =0 porque Imp_Dto >0 ----------------------------
'- -------------------------------------------------------------------------------------------------
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

'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"  ---------------------
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_DNI))
    Cumple = Fnc_Filtro_Filas(T, BD_DNI, "=", "1", Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Ficticio, DNI=1", 0, _
                                                        Format(rowfind, " #,##0") & " reg. ", _
                                                        " de " & Format(NVivas, "#,##0") & " reg.")
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Ficticio, DNI=1", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos con AE=300 (M013 mal matriculado en Gestión Académica) ==>> "SEMINARI D'ORIENTACIÓ PER A PREPARACIÓ DE PROVES PER A MAJORS DE 25 ANYS"  --------------------------------------
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_ActivEco))
    Cumple = Fnc_Filtro_Filas(T, BD_ActivEco, "=", 300, Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. AE=300 que debería ser AE=4 y Plan=M013.", 0, _
                                                        Format(rowfind, " #,##0") & " reg. ", _
                                                        " de " & Format(NVivas, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " - Son Rec. EFP, Plan=M013 'Seminario orientación pruebas > 65', mal matriculado x Secretaría de Acceso", 0)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", " - Mal matrículados con Rec.Mov. con AE=300, los borro porque vendrán en el AE4x4 ¡¡rectificados a mano!!", 0)
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Ficticio, DNI=1", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos ANULADOS -------------------------------------------------------------------------
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_Anul))
    Cumple = Fnc_Filtro_Filas(T, BD_Anul, "=", "S", Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Anulados ", 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Anulados ", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos NO Martrícula --------------------------------------------------------------------
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_Matricula))
    Cumple = Fnc_Filtro_Filas(T, BD_Matricula, "=", "N", Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. NO Matrícula ", 0, _
                                                        Format(rowfind, " #,##0") & " reg. ", _
                                                        " de " & Format(NVivas, "#,##0") & " reg.")
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. NO Matrícula ", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos INVALIDADOS ----------------------------------------------------------------------
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_Hinvalid))
    Cumple = Fnc_Filtro_Filas(T, BD_Hinvalid, "=", "S", Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. Invalidados ", 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. Invalidados ", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'-Filtra Recibos con F_Cob > APP_FechCierreCont y Limpia-Clear las Columnas BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag y BD_HTipCob -------
'-     (en la hoja se vacían al final, celda a celda agrupadas: CTA. CCC y Form. Pago son textos con cifras y Volcar los convertiría)
    Cumple = Fnc_Filtro_Filas(T, BD_FCob, ">", Cierre, Viva)
    rowfind = Fnc_Filas_Contar(Cumple)
    If rowfind > 0 Then
        For Fila = 1 To T.NumFilas
            If Cumple(Fila) Then
                Limpiar(Fila) = True
                T.Datos(Fila, BD_FCob) = Empty                      '- Lo mira Tb_CriT_Reg_Err, más abajo
            End If
        Next Fila
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clear Data Rec. F_Cob > " & Prog__APP.Range("APP_FechCierreCont"), 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. F_Cob > " & Prog__APP.Range("APP_FechCierreCont"), 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- ----------- Borrar Recibos con fechas FUERA DEL PERÍODO CONTABLE --------------------------------
'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos Emitidos en Años Posteriores a AnoCont -------------------------------------------
'-Filtra Recibos con F_Emi > APP_FechCierreCont y los Borra  ---------------------------------------
'-     (2026-10-05: antes, si había alguno, una línea de más borraba la tabla entera; ahora se borran solo esos)
    Call Rut_TablaRam_Anotar_Orden(Tandas, Array(BD_FEmi))
    Cumple = Fnc_Filtro_Filas(T, BD_FEmi, ">", Cierre, Viva)
    rowfind = Fnc_Filas_Borrar(Viva, Cumple, NVivas)
    If rowfind > 0 Then
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Borrados Rec. F_Emi > " & FechCierreCont, 0, _
             Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. F_Emi > " & FechCierreCont, 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- Borrar Incongruencias de Fechas -----------------------------------------------------------------
'- -------------------------------------------------------------------------------------------------
'- Borrar Recibos Cobrados en Años Anteriores o Posteriores a AnoCont ------------------------------
'- -------------------------------------------------------------------------------------------------
    Cumple = Fnc_CriT_Filas(T, "Tb_CriT_Reg_Err", Viva)
    rowfind = Fnc_Filas_Contar(Cumple)
    If rowfind > 0 Then
        Call Rut_TablaRam_Marcar(T, Cumple, BD_Obs_Conta, "Tb_CriT_Err_Date")
        ErrDate = Cumple                                            '- Se copian a BD_ErrDate, se borren o no
        If Prog__APP_Switch.Range("Sw_DelRegErrDate") Then
            Call Fnc_Filas_Borrar(Viva, Cumple, NVivas)
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Del Rec. con Fechas Incongruentes. ", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        Else
            '- Visualizo el progreso --------
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Marcados Rec. con Fechas Incongruentes. ", 0, _
                 Format(rowfind, " #,##0") & " reg. ", " de " & Format(NVivas, "#,##0") & " reg.")
        End If
    Else
        '- Visualizo el progreso --------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "No hay Rec. con Fechas Incongruentes.", 0)
    End If

'- -------------------------------------------------------------------------------------------------
'- En la hoja: marcas, datos de cobro vaciados, una ordenación, copia a BD_ErrDate y un solo borrado -----------------
'- -------------------------------------------------------------------------------------------------
    NBorrar = T.NumFilas - NVivas
    If Fnc_Filas_Contar(Limpiar) > 0 Then
        Call Rut_TablaRam_Vaciar_Celdas(Lo_Data, Limpiar, Array(BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag, BD_HTipCob))
    End If
    If NBorrar > 0 Then                         '- Columna auxiliar Tipo_Rec (vacía hasta M_114): 1 = a borrar
        For Fila = 1 To T.NumFilas
            If Viva(Fila) Then T.Datos(Fila, BD_Tipo_Rec) = 0 Else T.Datos(Fila, BD_Tipo_Rec) = 1
        Next Fila
        T.Modificada(BD_Tipo_Rec) = True
    End If
    Call Rut_TablaRam_Volcar(T, Lo_Data)        '- Obs_Conta (si hay marcas) y la columna auxiliar
    Erase T.Datos
    If NBorrar > 0 Then
        Call Rut_TablaRam_Ordenar_Tandas(Lo_Data, Tandas, BD_Tipo_Rec)
        Lo_Data.ListColumns(BD_Tipo_Rec).DataBodyRange.ClearContents
    Else
        Call Rut_TablaRam_Ordenar_Tandas(Lo_Data, Tandas)
    End If
    If Fnc_Filas_Contar(ErrDate) > 0 Then       '- Las de fechas incongruentes llevan la marca en Obs_Conta
        Lo_Data.Range.AutoFilter Field:=BD_Obs_Conta, Criteria1:="=Tb_CriT_Err_Date"
        Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_Data, Lo_BD_ErrDate, True)
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
    End If
    If NBorrar = T.NumFilas Then                '- Las filas a borrar han quedado juntas al final
        Lo_Data.DataBodyRange.Delete
    ElseIf NBorrar > 0 Then
        Lo_Data.DataBodyRange.Rows(NVivas + 1).Resize(NBorrar).Delete Shift:=xlUp
    End If

Terminar:
Debug.Print "<<< RuT_Remove_Null_Reg"
    Lo_Data.ShowTotals = True
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sh_Data)
End Sub     ' RuT_Remove_Reg_No_Valid
'- -------------------------------------------------------------------------------------------------
