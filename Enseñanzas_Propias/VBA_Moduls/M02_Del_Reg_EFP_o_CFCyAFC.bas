Attribute VB_Name = "M02_Del_Reg_EFP_o_CFCyAFC"
' Last Rev. 2026-09-27 21:30
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M02_Del_Reg_EFP_o_CFCyAFC - Separar EFP de CFCyAFC en la importacion
' =================================================================================================
'
' PROPOSITO
'  Cada copia del libro gestiona UN tipo de ensenanza. Esta rutina borra de la
'  importacion los recibos que no corresponden al tipo configurado en
'  APP_EFP_o_CFC (celda de Prog__APP): o los EFP, o todo lo demas.
'  La llama M01 sobre la copia en RAM, antes de volcar nada al libro.
'
' INDICE DE RUTINAS Y FUNCIONES
'  Rut_Borrar_Rec_EFP_o_CFCyAFC(Lo_Data) ... Unica rutina del modulo.
'
' TRAMOS DE PROGRAMACION
'    1. Quita filtros y ordena por BD_TipoCurso (ordenar antes acelera mucho
'       el borrado por rangos visibles).
'    2. Verifica que las cabeceras BD_TipoCurso y BD_ActivEco sean
'       'TipoCurso'/'Activ_Eco' (lo que ya define Tb_DefCols_Bdatos): el
'       AdvancedFilter exige que la cabecera coincida con la del rango de
'       criterios. Si no coincide, aborta con Err.Raise (M01 lo captura en
'       Gestion_Error y limpia).
'    3. Verifica que Tb_CriT_EFP (hoja Prog_Filtros_Concepto) exista y tenga al
'       menos una fila de criterios; si no, aborta igual que en el paso 2.
'    4. AdvancedFilter con Tb_CriT_EFP -> deja visibles los recibos de EFP.
'    5. Segun el tipo configurado:
'         - EFP     : anade una columna auxiliar, marca los visibles como 'EFP',
'                     invierte el filtro (<>EFP) y borra el resto. Borra la
'                     columna auxiliar justo antes de cada salida a FinRut.
'         - CFCyAFC : borra directamente los visibles (los EFP); no usa columna
'                     auxiliar, no la necesita para invertir nada.
'       Si no hay ningun recibo EFP, solo informa.
'       Si tras borrar la tabla se queda sin filas, sale a FinRut sin seguir.
'    6. FinRut: quita filtros. Es solo el punto de reunion; cada rama que usa
'       la columna auxiliar la borra ella misma antes de llegar aqui.
'
'  El recuento se hace SIEMPRE sobre BD_Ref con SpecialCells(xlCellTypeVisible)
'  menos 1 (la cabecera). Por eso la columna BD_Ref debe estar visible: si se
'  oculta, el conteo sale mal y se borra de menos o de mas.
'
' NOTAS
'  El nombre del End Sub ('Rut_Incorporar_Concept_Eco_y_Tipo_Ensenanza') es un
'  comentario heredado de un copiar-pegar: no corresponde a esta rutina.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'20265-01-11
Option Explicit

' ==================================================================================================
'- -------------------------------------------------------------------------------------------------
'- Borrar Tipo de Enseñanza EFP o CFCyAFC (EFP, CFC, AFC, TNCT) No Deseados ------------------------
'- -------------------------------------------------------------------------------------------------
Sub Rut_Borrar_Rec_EFP_o_CFCyAFC(Lo_Data As ListObject)
Debug.Print ">>> Rut_Borrar_Rec_EFP_o_CFCyAFC"
    
    Dim Curso_Acad      As String:      Curso_Acad = Prog__APP.Range("APP_CursAcad")
    Dim TipoCurso       As String:      TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")
    Dim rowfind             As Long
    Dim TxtMsg1  As String, TxtMsg2  As String, TxtMsg3  As String

    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_Lo_Sort(Lo_Data, G04_TipoCurso, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----

    With Lo_Data
        .ShowTotals = False
        '- El AdvancedFilter exige que la cabecera de criterios (Tb_CriT_EFP) case con la
        '- cabecera real de la columna. Ya coincide por Tb_DefCols_Bdatos: se verifica en vez
        '- de sobrescribirla (lo anterior era un no-op que escondia el supuesto).
        Dim Cab_TipoCurso As String, Cab_ActivEco As String
        Cab_TipoCurso = .Range(G04_TipoCurso).Value
        Cab_ActivEco = .Range(G04_ActivEco).Value
        If Cab_TipoCurso <> "TipoCurso" Or Cab_ActivEco <> "Activ_Eco" Then
            Err.Raise vbObjectError + 1, "Rut_Borrar_Rec_EFP_o_CFCyAFC", _
                "Cabeceras G04_TipoCurso/G04_ActivEco distintas de lo esperado por Tb_CriT_EFP ('" & Cab_TipoCurso & "' / '" & Cab_ActivEco & "'). Revisa Tb_DefCols_Bdatos."
        End If

        '- Verifica que el rango de criterios exista y tenga al menos 1 fila de criterios
        '- (la cabecera no cuenta): sin esto, un AdvancedFilter con Tb_CriT_EFP vacia o
        '- ausente no filtra nada y el bug pasa desapercibido (mensaje "NO hay Rec. de EFP").
        Dim Rg_CriT_EFP As Range
        On Error Resume Next
        Set Rg_CriT_EFP = Prog_Filtros_Concepto.Range("Tb_CriT_EFP")
        On Error GoTo 0
        If Rg_CriT_EFP Is Nothing Then
            Err.Raise vbObjectError + 2, "Rut_Borrar_Rec_EFP_o_CFCyAFC", _
                "No existe el rango con nombre Tb_CriT_EFP en Prog_Filtros_Concepto."
        ElseIf Rg_CriT_EFP.Rows.Count < 2 Then
            Err.Raise vbObjectError + 3, "Rut_Borrar_Rec_EFP_o_CFCyAFC", _
                "Tb_CriT_EFP no tiene ninguna fila de criterios (solo cabecera)."
        End If

        '-Filtra Recibos - EFP - Estudios de Formación Permanente: Máster, Especialista, Experto. --
        .Range.AdvancedFilter xlFilterInPlace, Rg_CriT_EFP
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then     '- hay rec. de EFP
            If TipoCurso = "EFP" Then
                .ListColumns.Add
                .DataBodyRange.Columns(.ListColumns.Count).SpecialCells(xlCellTypeVisible).Cells.Value = "EFP"
                '- Borra los que NO son "Tít. Propios" EFP ----
                Call Rut_Lo_Filtros_Quitar(Lo_Data)
                .Range.AutoFilter Field:=.ListColumns.Count, Criteria1:="<>EFP"     '- Selecciono los que NO mson EFP
                rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
                If rowfind > 0 Then
                    .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                        TxtMsg1 = "Borrados Recibos de estudios NO EFP_" & Curso_Acad
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
                    If .ListRows.Count = 0 Then
                        .ListColumns(.ListColumns.Count).Delete
                        GoTo FinRut        ' NO QUEDAN REGISTROS
                    End If
                End If
                .ListColumns(.ListColumns.Count).Delete
                GoTo FinRut
            Else
                .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                        TxtMsg1 = "Borrados Recibos de estudios EFP_" & Curso_Acad
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
                If .ListRows.Count = 0 Then GoTo FinRut        ' NO QUEDAN REGISTROS
            End If
        Else    '- No hay Rec. EFP
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe("NO hay Rec. de estudios de EFP_", , " quedan: " & Format(.ListRows.Count, "#,##0") & "reg.")
        End If
FinRut:
    End With    '-  Lo_Data
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
End Sub     '- Rut_Incorporar_Concept_Eco_y_Tipo_Enseñanza -----------------------------------------




