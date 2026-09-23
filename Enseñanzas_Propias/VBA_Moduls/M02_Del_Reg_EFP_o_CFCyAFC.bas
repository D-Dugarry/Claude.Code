Attribute VB_Name = "M02_Del_Reg_EFP_o_CFCyAFC"
' Last Rev. 2026-09-23 18:56
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
'    2. Anade una columna auxiliar al final de la tabla y rotula las cabeceras
'       BD_TipoCurso y BD_ActivEco ('TipoCurso' / 'Activ_Eco'): el AdvancedFilter
'       exige que la cabecera coincida con la del rango de criterios.
'    3. AdvancedFilter con Tb_CriT_EFP (hoja Prog_Filtros_Concepto) -> deja
'       visibles los recibos de EFP.
'    4. Segun el tipo configurado:
'         - EFP     : marca los visibles como 'EFP' en la columna auxiliar,
'                     invierte el filtro (<>EFP) y borra el resto.
'         - CFCyAFC : borra directamente los visibles (los EFP).
'       Si no hay ningun recibo EFP, solo informa.
'    5. FinRut: elimina la columna auxiliar y quita filtros.
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
    Dim rowfind             As Variant
    Dim TxT_Progreso        As String
    Dim TxtMsg1  As String, TxtMsg2  As String, TxtMsg3  As String
    Dim Cont_Fail           As Long

    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Call Rut_Lo_Sort(Lo_Data, G04_TipoCurso, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----

    Application.DisplayAlerts = False
    With Lo_Data
        .ListColumns.Add
        .ShowTotals = False
        .Range(G04_TipoCurso) = "TipoCurso"  ' para q funcionen los criterios de filtro deben tener ese nombre en la cabecera de la Col.
        .Range(G04_ActivEco) = "Activ_Eco"  ' para q funcionen los criterios de filtro deben tener ese nombre en la cabecera de la Col.
        '-Filtra Recibos - EFP - Estudios de Formación Permanente: Máster, Especialista, Experto. --
        .Range.AdvancedFilter xlFilterInPlace, Prog_Filtros_Concepto.Range("Tb_CriT_EFP")
        rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
        If rowfind > 0 Then     '- hay rec. de EFP
            If TipoCurso = "EFP" Then
                .DataBodyRange.Columns(.ListColumns.Count).SpecialCells(xlCellTypeVisible).Cells.Value = "EFP"
                '- Borra los que NO son "Tít. Propios" EFP ----
                Call Rut_Lo_Filtros_Quitar(Lo_Data)
                .Range.AutoFilter Field:=.ListColumns.Count, Criteria1:="<>EFP"     '- Selecciono los que NO mson EFP
                rowfind = .Range.Columns(G04_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA G04_Ref
                If rowfind > .ListColumns.Count Then
                    .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                        TxtMsg1 = "Borrados Recibos de estudios NO EFP_" & Curso_Acad
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
                End If
                GoTo FinRut
            Else
                .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                        TxtMsg1 = "Borrados Recibos de estudios EFP_" & Curso_Acad
                        TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                        TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
                    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            End If
        Else    '- No hay Rec. EFP
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe("NO hay Rec. de estudios de EFP_", , " quedan: " & Format(.ListRows.Count, "#,##0") & "reg.")
        End If
FinRut:
        .ListColumns(.ListColumns.Count).Delete
    End With    '-  Lo_Data
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Application.DisplayAlerts = True
End Sub     '- Rut_Incorporar_Concept_Eco_y_Tipo_Enseñanza -----------------------------------------




