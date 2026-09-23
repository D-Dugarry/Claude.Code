Attribute VB_Name = "M05_Asign_Tipo_Recibo"
' Last Rev. 2026-09-23 18:56
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M05_Asign_Tipo_Recibo - Tipificacion contable del recibo por ejercicio
' =================================================================================================
'
' PROPOSITO
'  Determina el TIPO de cada recibo segun la relacion entre sus anos de
'  emision, vencimiento y cobro frente al ano contable en curso (APP_AnoCont):
'  Emitido, EjeAnt, Anejo, Aplazado, ADxAplz, mas marcas especiales de
'  devolucion, cobro ya contabilizado y ajuste de matricula.
'  Es la clasificacion que decide en que ejercicio se imputa cada importe.
'
' INDICE DE RUTINAS Y FUNCIONES
'  RuT_Determinar_Tipo_Recibo_ByHand ... Lanzadera manual (prepara la hoja).
'  RuT_Determinar_Tipo_Recibo .......... Rutina principal. Sin argumentos:
'                                        trabaja siempre sobre Prog_LsGes04.
'
' TRAMOS DE PROGRAMACION
'    0. Ordena por ACont_Emi, ACont_Vto y ACont_Cob, y limpia tanto BD_Tipo_Rec
'       como el bloque de columnas de marca G04_Flag_Primera..G04_Flag_ADxAplz
'       (de una vez, con Resize sobre G04_Flag_Cuantas).
'
'    TIPIFICACION PRINCIPAL - una pasada por tipo, con AdvancedFilter sobre un
'    rango de criterios propio; cada una escribe BD_Tipo_Rec Y su columna de
'    marca G04_Flag_*, y suma al contador de tipificados:
'       Tb_CriT_Reg_Err  -> '_ERR_Date_' (fechas incoherentes; tambien en
'                           BD_Incidencias). No cuenta como tipificado.
'       Tb_CriT_Emitido  -> 'Emitido'   del ejercicio corriente.
'       Tb_CriT_EjeAnt   -> 'EjeAnt'    del ejercicio anterior.
'       Tb_CriT_Aneja    -> 'Anejo'     anteriores al ejercicio anterior.
'       Tb_CriT_Aplazado -> 'Aplazado'  anulados por aplazamiento el ano pasado.
'       Tb_CriT_ADxAplz  -> 'ADxAplz'   anulados por aplazamiento este ano, a
'                           cobrar el que viene.
'
'    MARCAS QUE SOBRESCRIBEN (se aplican DESPUES, con AutoFilter simple):
'       BD_ImpRec  < 0            -> '_Devol_'        (devolucion).
'       BD_ACont_Cob < AnoCont    -> '_AnoCont'<aa>_' (ya contabilizado).
'       BD_ImpAdm  < 0            -> '_Ajust_Matric_' (ajuste de matricula).
'       Tb_CriT_Reg_Anul          -> marca '_Reg_Anul_' en BD_Incidencias
'                                    (NO toca BD_Tipo_Rec).
'
'    CONTROL: cuenta los recibos sin tipificar. Si sale > 0 hay que revisar los
'    rangos de criterios: o dejan huecos, o se solapan. El MsgBox de aviso esta
'    comentado; el dato queda en el informe.
'
'    ULTIMO TRAMO: oculta varios bloques de columnas de Prog_LsGes04 para poder
'    revisar el resultado a ojo. Es ayuda de depuracion, no logica de negocio.
'
' NOTAS
'  Las columnas de marca G04_Flag_* permiten ver a posteriori por que filtros
'  paso un recibo, aunque BD_Tipo_Rec haya sido sobrescrito luego por _Devol_,
'  _AnoCont_ o _Ajust_Matric_. Por eso conviven las dos cosas.
'
'  El informe detallado solo llega a Form_Menu si Sw_Boss esta activo.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2026-01-06
Option Explicit

            Sub RuT_Determinar_Tipo_Recibo_ByHand()
                Prog_LsGes04.Unprotect
                Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)          '- Quita filtros, filas y columnas ocultas
                Call RuT_Determinar_Tipo_Recibo
            End Sub
'- -------------------------------------------------------------------------------------------------
'- Clasificar Recibos en Emitidos, Remesados, EjeAnt, ADxAplz, Añejas ------------------------------
'- -------------------------------------------------------------------------------------------------
Sub RuT_Determinar_Tipo_Recibo()
Debug.Print ">>> RuT_Determinar_Tipo_Recibo"
    Dim rowfind         As Variant
    Dim RegsCanTot      As Long
    Dim RegsSinTipo     As Long
    Dim APP_AñoCont     As String:          APP_AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim TxtProgreso     As String:          TxtProgreso = Form_Menu.TB_Informe
    Dim Sw_Boss         As Boolean:         Sw_Boss = Prog__APP_Switch.Range("Sw_Boss")
    
    Dim Lo_G04          As ListObject:      Set Lo_G04 = Prog_LsGes04.ListObjects(1)
    Application.ScreenUpdating = False
    Prog_LsGes04.Select
    Lo_G04.ShowTotals = False
    Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)
    
        '- Visualizo el progreso -------------------------------------------------------------------
        'Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Clasificación de Recibos, Estadística:", 0, , , , , , 2)
    With Lo_G04
    
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, G04_ACont_Emi, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        Call Rut_Lo_Sort(Lo_G04, G04_ACont_Vto, xlAscending, False)
        Call Rut_Lo_Sort(Lo_G04, G04_ACont_Cob, xlAscending, False)
            
        .DataBodyRange.Columns(G04_Tipo_Rec).ClearContents
        .DataBodyRange.Columns(G04_Flag_Primera).Resize(, G04_Flag_Cuantas).ClearContents

        TxtProgreso = TxtProgreso & vbCrLf & String(8, " ") & " Tipificado de Recibos. " & String(8, "_")
        
        '-Filtra Recibos ErrDate - -----------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, G04_FEmi, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_Err")
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Incidencias).SpecialCells(xlCellTypeVisible).Cells.Value = "_ERR_Date_"
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_ERR_Date_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros con errores de fechas."
    
        '-Filtra Recibos Emitidos ------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .AutoFilter.ShowAllData            ' Elimina los filtros
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Emitido")
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
            .DataBodyRange.Columns(G04_Flag_Emitido).SpecialCells(xlCellTypeVisible).Cells.Value = "Emitido"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros Emitidos."

        '-Filtra Recibos EjeAnt --------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_EjeAnt")
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "EjeAnt"
            .DataBodyRange.Columns(G04_Flag_EjeAnt).SpecialCells(xlCellTypeVisible).Cells.Value = "EjeAnt"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " Registros del Ejercicio Anterior. " & APP_AñoCont - 1
    
        '-Filtra Recibos Añejos --------------------------------------------------------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Añeja")
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Añejo"
            .DataBodyRange.Columns(G04_Flag_Añejos).SpecialCells(xlCellTypeVisible).Cells.Value = "Añejo"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " Registros Añejos, anteriores a " & APP_AñoCont - 1
    
        '-Filtra Recibos Aplazado ------------------------------------------------------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Aplazado")
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "Aplazado"
            .DataBodyRange.Columns(G04_Flag_Aplazado).SpecialCells(xlCellTypeVisible).Cells.Value = "Aplazado"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " Registros Aplazados, Anulados por Aplazamiento en " & APP_AñoCont - 1
    
        '-Filtra Recibos ADxAplz -------------------------------------------------------------------
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_ADxAplz")
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz"
            .DataBodyRange.Columns(G04_Flag_ADxAplz).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz"
            RegsCanTot = RegsCanTot + rowfind
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & _
                    " Registros ADxAplz, Anulados por Aplazamiento en " & APP_AñoCont & ", a cobrar en " & APP_AñoCont + 1
        
'        '-Filtra Recibos ADxAplz_SinCob -----------------------------------------------------------
'        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_ADxAplz_SinCob")
'        RowFind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
'        If RowFind > 0 Then
'            .DataBodyRange.Columns(BD_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz_SinCob"
'            .DataBodyRange.Columns(G04_Flag_ADxAplz).SpecialCells(xlCellTypeVisible).Cells.Value = "ADxAplz_SinCob"
'            RegsCanTot = RegsCanTot + RowFind
'        End If
'        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(RowFind, "#,##0"), 8) & _
'                    " Registros ADxAplz_SinCob, Anulados por Aplazamiento en " & APP_AñoCont - 1 & "  y  ¡¡¡ SIN Cobrar en " & APP_AñoCont & " !!!"
        
        
        '-------------------------------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------
        '-Filtra Recibos Devolución - --------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, G04_ImpRec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=G04_ImpRec, Criteria1:="<0"
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_Devol_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros de devolución."
        
        '-------------------------------------------------------------------------------------------
        '-_Contab_Ant_------------------------------------------------------------------------------
        '------------------ Filtra Cobradas en Años anteriores al de Emisión -----------------------
        '-------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .Range.AutoFilter Field:=G04_ACont_Cob, Criteria1:="<" & APP_AñoCont
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_AñoCont'" & Right(APP_AñoCont - 1, 2) & "_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros _Contab_Ant_, Cobrados anteriormente y por lo tanto, ya Contabilizado."
        
        '-------------------------------------------------------------------------------------------
        '-_Ajust_Matríc_----------------------------------------------------------------------------
        '------------------ Filtra Ajustes de Matrícula (ImpAdm negativo) --------------------------
        '-------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        .Range.AutoFilter Field:=G04_ImpAdm, Criteria1:="<0"
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Tipo_Rec).SpecialCells(xlCellTypeVisible).Cells.Value = "_Ajust_Matríc_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros _Ajust_Matríc_, Ajustes de Matrícula (Imp. Admin. negativo)."
        
        '-------------------------------------------------------------------------------------------
        '-Filtra Recibos Sin Tipo ------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        RegsSinTipo = Application.WorksheetFunction.CountIf(.DataBodyRange.Columns(G04_Tipo_Rec), "")

        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(RegsSinTipo, "#,##0"), 8) & " Registros SIN Tipificar, de" & Lo_G04.ListRows.Count & "reg."
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "=") & " " & Format(Lo_G04.ListRows.Count, "#,##0"), 8) & " Total Registros, Tipificados: " & RegsCanTot & "reg." & vbLf


        '-------------------------------------------------------------------------------------------
        '-Filtra Recibos Anul - -------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        Call Rut_Lo_Sort(Lo_G04, G04_Anul, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_Anul")
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then
            .DataBodyRange.Columns(G04_Incidencias).SpecialCells(xlCellTypeVisible).Cells.Value = "_Reg_Anul_"
        End If
        TxtProgreso = TxtProgreso & vbCrLf & Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " Registros Anulados."
    
        '-------------------------------------------------------------------------------------------
        '-Filtra Recibos Sin Tipo ------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_G04)
        rowfind = Application.WorksheetFunction.CountIf(.DataBodyRange.Columns(G04_Tipo_Rec), "")
        If rowfind > 0 Then
'''            MsgBox "¡¡¡ Recibos SIN identificar su Tipo_Recibos !!! " & RowFind & "reg.", vbOKOnly + vbExclamation
        End If
        
        TxtProgreso = TxtProgreso & vbLf & _
            Right(String(8, "_") & Format(Application.WorksheetFunction.CountIf(.DataBodyRange.Columns(G04_Tipo_Rec), "="), "#,##0"), 8) & " 'Sin Identificar' Rec. SIN Determinar su Tipo (Si > 0, HAY QUE VERIFICAR FILTROS)" & vbCrLf & _
                  String(8, " ") & " Son Reg. que se han quedado fuera de todos los filtros, o filtros que duplican tipo." & vbCrLf & _
            Right(String(8, "=") & " " & Format(.ListRows.Count, "#,##0"), 8) & " Reg. en BDatos." & vbCrLf
       
        '- Oculta columnas para poder verificar los resultados, SÓLO CUANDO SE ESTÁ COMPROBANDO LA PROGRAMACIÓN..............
        .DataBodyRange.Columns(3).Resize(, 3).Hidden = True
        .DataBodyRange.Columns(9).Resize(, 5).Hidden = True
        .DataBodyRange.Columns(18).Resize(, 7).Hidden = True
        .DataBodyRange.Columns(26).Resize(, 8).Hidden = True
        .DataBodyRange.Columns(38).Resize(, 10).Hidden = True
        .DataBodyRange.Columns(49).Resize(, 3).Hidden = True
        
    End With    '- Lo_G04
    
    If Sw_Boss Then Form_Menu.TB_Informe = TxtProgreso
    Call Rut_Lo_Filtros_Quitar(Lo_G04)
    Lo_G04.ShowTotals = True
Application.ScreenUpdating = True
'MsgBox "Finn"
Debug.Print "<<< RuT_Determinar_Tipo_Recibo"
End Sub


