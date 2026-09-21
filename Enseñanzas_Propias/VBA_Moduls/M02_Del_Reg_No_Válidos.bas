Attribute VB_Name = "M02_Del_Reg_No_Válidos"
' Last Rev. 2026-09-21 12:12
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M02_Del_Reg_No_Validos - Filtrado de recibos no procesables
' =================================================================================================
'
' PROPOSITO
'  Segunda criba de la importacion: elimina de la copia en RAM todos los
'  recibos que no deben entrar en la liquidacion, en cinco pasadas
'  independientes. La llama M01 justo despues de separar EFP/CFCyAFC.
'
' INDICE DE RUTINAS Y FUNCIONES
'  RuT_Del_Reg_NO_Validos(Lo_Data) ... Unica rutina del modulo.
'
' TRAMOS DE PROGRAMACION
'  Las cinco pasadas, todas con el mismo patron
'  (quitar filtros -> ordenar -> filtrar -> contar sobre BD_Ref -> borrar
'  visibles -> informar), y todas con salida anticipada a Restablecer_Valores
'  si la tabla se queda sin registros:
'
'    1. Otro curso academico: BD_C_Acad <> APP_CursAcad.
'    2. No matriculas: BD_Matricula = 'N'.
'    3. Fuera de Ensenanzas Propias: BD_ActivEco <> 4. Antes de filtrar hace
'       TextToColumns sobre la columna para convertir el texto a numero (si no,
'       el criterio numerico no casa).
'    4. Matricula de actividad academica a coste CERO (ImpRec = ImpDto = 0),
'       via AdvancedFilter con Tb_CriT_ImpMatCeroLsGes04.
'    5. Subvencionados al 100%: BD_ImpRec = 0 porque el descuento cubre el
'       importe academico.
'
'  Restablecer_Valores: reactiva totales, quita filtros y alertas.
'
' NOTAS
'  Los simbolos de los mensajes se escriben con ChrW (<> = &H2260,
'  >= = &H2265, => = &H21D2) para no depender de la codificacion del fichero.
'
'  Igual que en el resto del pipeline: el conteo exige BD_Ref VISIBLE.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2026-01-11
Option Explicit

'        - Borrar Recibos NO Pertinentes:  otros C_Acad, Matrícula=N, AE<>4
'                - Borrar Recibos de otro Curso_Acad
'                - Borrar Recibos que BD_Matricula = "N"
'                - Borrar Recibos que BD_ActivEco <> 4 (Enseñanzas Propias)
'                - Borrar Recibos de Matrículas de coste CERO - ImpRec=ImpDto=0 - Recibos Matrícula de Actividad Académica a Coste CERO
'                - Borrar Recibos Importe CERO - Subvencionado- Imp_Rec =0 porque Imp_Dto >0

'- -------------------------------------------------------------------------------------------------
'- Borrar Rec. NO Válidos: otros C_Acad, Matrícula=N, AE<>4 --------
'- -------------------------------------------------------------------------------------------------
Sub RuT_Del_Reg_NO_Válidos(Lo_Data As ListObject)
Debug.Print ">>> RuT_Del_Reg_NO_Válidos"
    Dim rowfind         As Variant
    Dim Cont_Fail       As Long
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim Curso_Acad      As String:      Curso_Acad = Prog__APP.Range("APP_CursAcad")
    Dim TxtMsg1  As String, TxtMsg2  As String, TxtMsg3  As String
    Lo_Data.ShowTotals = False
    Application.DisplayAlerts = False
    With Lo_Data
        .ShowTotals = False
        
        '- Borra los que son de otro Curso_Acad ----------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_C_Acad, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_C_Acad, Criteria1:="<>" & Curso_Acad   '- Elimino los que son de otro curso
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                TxtMsg1 = "Borrados Recibos de Curso-Acad " & ChrW(&H2260) & " " & Curso_Acad
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            If .ListRows.Count = 0 Then GoTo Restablecer_Valores        ' NO QUEDAN REGISTROS
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & "No hay Recibos de Curso-Acad " & ChrW(&H2260) & " " & Curso_Acad
        End If
        
        '- Borra los que BD_Matricula = "N" --------------------------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_Matricula, xlAscending, True)    '- Ordenar primero accelera un montón el borrado --
        .Range.AutoFilter Field:=BD_Matricula, Criteria1:="=N"                          '- Elimino las NO Martrículas
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                TxtMsg1 = "Borrados Recibos con Matrícula = N "
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            If .ListRows.Count = 0 Then GoTo Restablecer_Valores        ' NO QUEDAN REGISTROS
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & "No hay Recibos con Matrícula = N. "
        End If
        
        '- Borra los que BD_ActivEco <> 4 (Enseñanzas Propias) -------------------------------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        .DataBodyRange.Columns(BD_ActivEco).Select     '- Datos - Texto en Columnas - Finalizar - PARA NÚMEROS
        Selection.TextToColumns DataType:=xlDelimited, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
        Call Rut_Lo_Sort(Lo_Data, BD_ActivEco, xlAscending, True)    '- Ordenar primero accelera un montón el borrado ---
        '.Range.AutoFilter Field:=BD_ActivEco, Criteria1:="<>4", Operator:=xlAnd, Criteria2:="<>300"        '- Elimino las NO Títulos Propios
        .Range.AutoFilter Field:=BD_ActivEco, Criteria1:="<>4"                                              '- Elimino las NO Títulos Propios
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                TxtMsg1 = "Borrados Recibos con AE " & ChrW(&H2260) & " 4 "
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            If .ListRows.Count = 0 Then GoTo Restablecer_Valores        ' NO QUEDAN REGISTROS
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "No hay Recibos con AE " & ChrW(&H2260) & " 4."
        End If
        
        '-Borrar Recibos de Matrículas de coste CERO - ImpRec=ImpDto=0 - Recibos Matrícula de Actividad Académica a Coste CERO.
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_ImpRec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado --
        Call Rut_Lo_Sort(Lo_Data, BD_ImpDto, xlAscending, False)    '- Ordenar primero accelera un montón el borrado --
        .Range.AdvancedFilter xlFilterInPlace, Prog_Filtros_TipRec.Range("Tb_CriT_ImpMatCeroLsGes04")
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1   '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
'            On Error Resume Next
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
'            On Error GoTo 0
                TxtMsg1 = "Borrados Recibos Matrícula de Actividad Acad. a Coste CERO"
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            If .ListRows.Count = 0 Then GoTo Restablecer_Valores        ' NO QUEDAN REGISTROS
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "No hay Recibos de Matrícula de Actividad Académica a Coste CERO."
        End If
        
        '-Borrar Recibos Importe CERO - Subvencionado- Imp_Rec =0 porque Imp_Dto >0 ----------------
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        Call Rut_Lo_Sort(Lo_Data, BD_ImpRec, xlAscending, True)    '- Ordenar primero accelera un montón el borrado -----
        .Range.AutoFilter Field:=BD_ImpRec, Criteria1:="=0"
        rowfind = .Range.Columns(BD_Ref).SpecialCells(xlCellTypeVisible).Cells.Count - 1  '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA BD_Ref
        If rowfind > 0 Then
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
                TxtMsg1 = "Borrados Rec Subvencionados 100% (ImpDto " & ChrW(&H2265) & " ImpAcad " & ChrW(&H21D2) & " Rec=0€)"
                TxtMsg2 = Format(rowfind, "#,##0") & "reg."
                TxtMsg3 = " quedan " & Format(.ListRows.Count, "#,##0") & "reg."
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Func_Informe(TxtMsg1, TxtMsg2, TxtMsg3)
            If .ListRows.Count = 0 Then GoTo Restablecer_Valores        ' NO QUEDAN REGISTROS
        Else
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & "No hay Rec. Subvencionados ImpDto " & ChrW(&H2265) & " ImpAcad " & ChrW(&H21D2) & " ImpRec = 0"
        End If
        
    End With    ' Lo_Data
        
Restablecer_Valores:
    Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False     '- Visualizo el progreso
    Lo_Data.ShowTotals = True
    Call Rut_Lo_Filtros_Quitar(Lo_Data)
    Application.DisplayAlerts = True
End Sub
'- -------------------------------------------------------------------------------------------------




