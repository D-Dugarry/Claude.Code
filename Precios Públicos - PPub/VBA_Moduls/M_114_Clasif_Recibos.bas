Attribute VB_Name = "M_114_Clasif_Recibos"
' Last Rev. 2026-10-05 23:11
'Rev.: 2026-01-22
Option Explicit

            Sub RuT_Clasif_Recibos_ByHand()
                Sht__BD.Select
                Sht__BD.Unprotect
                Sht__BD.Columns(Sht__BD.ListObjects(1).Range.Columns(BD_Ref).Column).Hidden = False
                Call RuT_Clasif_Recibos
                Sht__BD.Unprotect
                MsgBox "FIN"
            End Sub
'- ----------------------------------------------------------------------------------------------------------------------------
'- Clasificar Recibos en Emitidos, Remesados, EjeAnt, ADxAplz, Añejas ---------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub RuT_Clasif_Recibos()
'- Desde el 2026-10-05 trabaja en RAM (fase 3 del paso a RAM): los filtros se evalúan sobre la copia en memoria de la
'- tabla (Rut_Lo_CriT_Ram, con las mismas reglas que los filtros de Excel) y las marcas se ponen en RAM. Después, en la
'- hoja: se devuelven Tipo_Rec, Incidencias y las columnas de las marcas, y UNA ordenación con el mismo resultado que las
'- dos de antes (M_115 da el importe al primer recibo de cada matrícula, así que el orden importa).
Debug.Print ">>> RuT_Clasif_Recibos"
    Dim RegsEmitido         As Long
    Dim RegsEjeAnt          As Long
    Dim RegsAnejo           As Long
    Dim RegsAplazado        As Long
    Dim RegsADxAplz         As Long
    Dim RegsSinTipo         As Long
    Dim RegsErrDate         As Long
    Dim RegsAnulado         As Long
    Dim RegsContabAnt       As Long
    Dim RegsClasifs         As Long
    Dim RegsDevolucion      As Long
    Dim RegsCanTot          As Long
    Dim RegsNOCUADRA        As Long

    Dim rowfind             As Long
    Dim APP_AnoCont         As String:      APP_AnoCont = Prog__APP.Range("APP_AnoCont")
    Dim TimeLapSub          As Single:      TimeLapSub = LastTimeLap
    Dim T                   As T_TablaRam
    Dim Cumple()            As Boolean
    Dim Fila                As Long

    Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)

    Sht__BD.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Prog__APP_Switch.Range("Sw_Col_Hide_Sht__BD") = False
    Sht__BD.Unprotect

        '- Visualizo el progreso
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificación de Recibos, Estadística:", 0, , , , , , 2)
    If Lo_BD.DataBodyRange Is Nothing Then GoTo Terminar
    With Lo_BD

        Call Rut_Lo_Filtros_Quitar(Lo_BD)
        .DataBodyRange.Columns(BD_Tipo_Rec).ClearContents
        .DataBodyRange.Columns(BD_CriT_Emi).Resize(, 56).ClearContents
        '.DataBodyRange.Columns(BD_CriT_Emi).Resize(, BD_CriT_ErrDate - BD_CriT_Emi + 1).ClearContents
        '.DataBodyRange.Columns(BD_CriT_Emi).Resize(, BD_CriT_ErrDate - BD_CriT_Emi + 2).ClearContents

        '- Tipo_Rec y las columnas de las marcas (BD_CriT_*) se acaban de vaciar: no hace falta leerlas, valen Empty en T
        Call Rut_TablaRam_Cargar(T, Lo_BD, Array(BD_ACont_Emi, BD_ACont_Vto, BD_ACont_Cob, BD_ImpRec, BD_ActivEco, _
                                 BD_Anul, BD_Matricula, BD_Hinvalid, BD_FEmi, BD_FCob, BD_Incidencias), True)

    '-Rec. Emitidos -------------------------------------------------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_Emitido")
        RegsEmitido = Fnc_Marcar(T, Cumple, BD_CriT_Emi, "Emitido")

    '-Rec. EjeAnt -------------------------------------------------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_EjeAnt")
        RegsEjeAnt = Fnc_Marcar(T, Cumple, BD_CriT_EjeAnt, "EjeAnt")

    '-Rec. Añejos -------------------------------------------------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_Aneja")
        RegsAnejo = Fnc_Marcar(T, Cumple, BD_CriT_Anejo, "Añejo")

    '-Rec. Aplazado -------------------------------------------------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_Aplazado")
        RegsAplazado = Fnc_Marcar(T, Cumple, BD_CriT_Aplazado, "Aplazado")

    '-Rec. ADxAplz -------------------------------------------------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_ADxAplz")
        RegsADxAplz = Fnc_Marcar(T, Cumple, BD_CriT_ADxAplz, "ADxAplz")

    '-----------------------------------------------------------------------------------------------
    '------------ A partir de ahora, las marcas pueden sobreescribir algún valor anterior ----------
    '-----------------------------------------------------------------------------------------------
    '-_Dev_EP_--------------------------------------------------------------------------------------
        '-Filtra Cobradas en Años anteriores al de Emisión -------------------------------------
        Cumple = Fnc_Filtro_Filas(T, BD_ImpRec, "<", 0)
        RegsDevolucion = Fnc_Marcar(T, Cumple, BD_CriT_DevEP, "_Dev_EP_")

    '-Filtra Recibos Anulados, NO Matrícula o Invalidados -------------------------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_Reg_Anul")
        RegsAnulado = Fnc_Marcar(T, Cumple, 0, "_Reg_Anul_")

    '-_Contab_Ant_----------------------------------------------------------------------------------
        '-Filtra Cobradas en Años anteriores al de Emisión -------------------------------------
        Cumple = Fnc_Filtro_Filas(T, BD_ACont_Cob, "<", APP_AnoCont)
        RegsContabAnt = Fnc_Marcar(T, Cumple, BD_CriT_ContabAnt, "_Contab_Ant_")

        rowfind = Fnc_Contar_Tipo(T, "Emitido")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. Emididos: ", 0, Format(rowfind, "#,##0") & " reg.")

        rowfind = Fnc_Contar_Tipo(T, "EjeAnt")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. EjeAnt: ", 0, Format(rowfind, "#,##0") & " reg.")

        rowfind = Fnc_Contar_Tipo(T, "Añejo")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. Añejos: ", 0, Format(rowfind, "#,##0") & " reg.")

        rowfind = Fnc_Contar_Tipo(T, "Aplazado")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. Aplazados: ", 0, Format(rowfind, "#,##0") & " reg.")

        rowfind = Fnc_Contar_Tipo(T, "ADxAplz")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. ADxAplz: ", 0, Format(rowfind, "#,##0") & " reg.")

        rowfind = Fnc_Contar_Tipo(T, "_Dev_EP_")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. de EP=AE4x4 _Devolución_EP_: ", 0, Format(rowfind, "#,##0") & " reg.")

        rowfind = Fnc_Contar_Tipo(T, "_Reg_Anul_")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. _Reg_Anul_: ", 0, Format(rowfind, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Registros: Anulados, NO Matrícula o Invalidados: ", 0)

        rowfind = Fnc_Contar_Tipo(T, "_Contab_Ant_")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. _Contab_Ant_: ", 0, Format(rowfind, "#,##0") & " reg.")
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "     Registros, Cobrados anteriormente y por lo tanto, ya Contabilizados. ", 0)

        '-------------------------------------------------------------------------------------------
        RegsNOCUADRA = Fnc_Contar_Tipo(T, "")
        RegsClasifs = T.NumFilas - RegsNOCUADRA
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Sumatorio Recibos clasificados (de " & Format(T.NumFilas, "#,##0") & " reg.)", 0, _
                                        Format(RegsClasifs, "#,##0") & " reg.", "faltan = " & Format(RegsNOCUADRA, "#,##0") & " reg.", , , , 2, 2)

        '-------------------------------------------------------------------------------------------
        '-Filtra Recibos ErrDate - Cobradas en Años anteriores al de Emisión -----------------------
        Cumple = Fnc_CriT_Filas(T, "Tb_CriT_Reg_Err")
        RegsErrDate = 0
        For Fila = 1 To T.NumFilas
            If Cumple(Fila) Then
                T.Datos(Fila, BD_Incidencias) = "_ERR_Date_"
                RegsErrDate = RegsErrDate + 1
            End If
        Next Fila
        T.Modificada(BD_Incidencias) = True
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificados Reg. _ERR_Date_: ", 0, _
                                        Format(RegsErrDate, "#,##0") & " reg.")

        '-------------------------------------------------------------------------------------------
        '-Filtra Recibos Sin Tipo ------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Total Recibos BDatos: ", 0, _
                                        Format(T.NumFilas, "#,##0") & " reg.", , , , , 2)
        RegsSinTipo = RegsNOCUADRA
Debug.Print "RegsSinTipo", RegsSinTipo
        RegsCanTot = T.NumFilas
Debug.Print "RegsCanTot", , Format(RegsCanTot, "#,##0")
        RegsNOCUADRA = RegsCanTot - (RegsEmitido + RegsEjeAnt + RegsAnejo + RegsAplazado + RegsADxAplz + RegsSinTipo)
Debug.Print "RegsNOCUADRA", RegsNOCUADRA

        '- En la hoja: las marcas y una ordenación con el resultado de las dos de antes: ACont_Emi, ACont_Vto y
        '- ACont_Cob (al empezar), y Tipo_Rec (al terminar).
        Call Rut_TablaRam_Volcar(T, Lo_BD)                              '- Tipo_Rec, Incidencias y BD_CriT_*
        Erase T.Datos
        Call Rut_TablaRam_Ordenar_Tandas(Lo_BD, Array(Array(BD_ACont_Emi, BD_ACont_Vto, BD_ACont_Cob), Array(BD_Tipo_Rec)))

    Call Rut_Lo_Filtros_Quitar(Lo_BD)

    End With        '- Lo_BD

Terminar:
    '- Visualizo el progreso -
    'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificación de registros finalizada.", TimeLapSub)

    Lo_BD.ShowTotals = True

Debug.Print "<<< RuT_Clasif_Recibos"

End Sub
'- -------------------------------------------------------------------------------------------------

'- Marca Tipo_Rec (y la columna ColMarca, si no es 0) de las filas que cumplen; devuelve cuántas son ---
Private Function Fnc_Marcar(T As T_TablaRam, Cumple() As Boolean, ByVal ColMarca As Long, ByVal Marca As String) As Long
    Dim Fila    As Long
    For Fila = 1 To T.NumFilas
        If Cumple(Fila) Then
            T.Datos(Fila, BD_Tipo_Rec) = Marca
            If ColMarca > 0 Then T.Datos(Fila, ColMarca) = Marca
            Fnc_Marcar = Fnc_Marcar + 1
        End If
    Next Fila
    T.Modificada(BD_Tipo_Rec) = True
    If ColMarca > 0 Then T.Modificada(ColMarca) = True
End Function

'- Cuántas filas tienen ese Tipo_Rec (sin distinguir mayúsculas, como CountIfs "=Emitido"); "" = vacío ---
Private Function Fnc_Contar_Tipo(T As T_TablaRam, ByVal Tipo As String) As Long
    Dim Fila    As Long
    For Fila = 1 To T.NumFilas
        If Tipo = "" Then
            If IsEmpty(T.Datos(Fila, BD_Tipo_Rec)) Then Fnc_Contar_Tipo = Fnc_Contar_Tipo + 1
        ElseIf Not IsEmpty(T.Datos(Fila, BD_Tipo_Rec)) Then
            If StrComp(T.Datos(Fila, BD_Tipo_Rec), Tipo, vbTextCompare) = 0 Then Fnc_Contar_Tipo = Fnc_Contar_Tipo + 1
        End If
    Next Fila
End Function
'- -------------------------------------------------------------------------------------------------
