Attribute VB_Name = "M08_Actualizar_Tb_Coef_VRI"
' Last Rev. 2026-09-21 18:03
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M08_Actualizar_Tb_Coef_VRI - Mantener la tabla de coeficientes de retencion del VRI
' =================================================================================================
'
' PROPOSITO
'  Mantiene al dia Prog_Coef_Ret_VRI, el catalogo de planes con su coeficiente
'  de retencion para el Vicerrectorado de Investigacion. Recorre Prog_BD y, por
'  cada plan que encuentra, actualiza su ficha o la crea si es un plan nuevo.
'  Ultimo paso del pipeline de importacion (lo llama M01 tras M07).
'
' INDICE DE RUTINAS Y FUNCIONES
'  RuT_Lo_Coef_VRI_Actualizar ... Unica rutina del modulo.
'
' TRAMOS DE PROGRAMACION
'    1. Prepara ambas hojas: ordena Prog_BD por BD_Plan y la tabla de
'       coeficientes por Orden y Plan.
'
'    2. Recorre Prog_BD detectando los CAMBIOS DE PLAN (por eso hace falta el
'       orden previo: cada plan se procesa una sola vez, en su primera fila).
'       Para cada plan nuevo busca su fila en Prog_Coef_Ret_VRI con Match:
'
'       - PLAN ENCONTRADO: actualiza concepto economico y tipo de ensenanza, y
'         ACUMULA la organica anotando el curso academico entre parentesis
'         ('<organica>_(<curso>)'), separando con salto de linea si ya habia
'         otras. No machaca el historico de organicas del plan.
'
'       - PLAN NO ENCONTRADO: anade la fila con plan, coeficiente, organica,
'         concepto y tipo, y la marca con 'x' en la columna Orden. Esa 'x' es
'         la senal de 'alta automatica pendiente de revisar'.
'
'    3. Solo en las filas marcadas con 'x' ajusta el coeficiente: se queda con
'       el MAXIMO entre el que ya tenia y el del registro (Application.Max), y
'       si queda en 0 aplica el coeficiente por defecto del libro, que es la
'       fila 1 de la tabla si APP_EFP_o_CFC = 'EFP' y la fila 2 en otro caso.
'       Las filas sin 'x' (revisadas a mano) quedan intactas.
'
'    4. Reordena la tabla, informa y, en Restablecer_Valores, reprotege Prog_BD
'       con UserInterfaceOnly:=True.
'
' NOTAS
'  Las dos primeras filas de Prog_Coef_Ret_VRI son los coeficientes POR
'  DEFECTO (1 = EFP, 2 = CFCyAFC): no son planes, y por eso M07 las usa con el
'  mismo criterio al dar de alta recibos.
'
'  El bloque de variables de importes declarado al principio (RegsEmis,
'  Imp_Cobr, Regs_Dev...) no llega a usarse: es residuo de una version previa.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2026-02-06
'- M31_Listar_PLANES
Option Explicit
'- Actualiza la Tabla de Referencia de los Coeficientes de Retención para el VRI
'===================================================================================================
Sub RuT_Lo_Coef_VRI_Actualizar()
'===================================================================================================
    Dim Cont                As Long
    Dim ContIni             As Long:        ContIni = 1
    Dim Coef_VRI            As Integer
    Dim Cod_Plan            As String
    Dim Orgánica            As String
    Dim Txt_Cabecera        As String
    Dim Cont_Plan           As Integer:     Cont_Plan = 1
    Dim PlanesSinCob        As Integer:     PlanesSinCob = 0
    Dim Regs_SinPlan        As Long:        Regs_SinPlan = 0    '- Anomalia: NO debe existir ningun registro sin Plan
    
    Dim RegsEmis            As Long     ' regs Emitidos
    Dim Imp_Emis            As Currency
    Dim RegsCobr            As Long     ' regs Cobrados
    Dim Imp_Cobr            As Currency
    Dim RegsPdts            As Long     ' regs Pendiente de pago
    Dim Imp_Pdte            As Currency
    
    Dim Regs_Dev            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula
    Dim Imp_Devo            As Currency
    Dim RDev_Pag            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, Pagados
    Dim IDev_Pag            As Currency
    Dim RDevPdte            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, No Pagados, pero Ajuste Matrícula NO deben ser pagados
    Dim IDevPdte            As Currency
    
    Dim Lo_BD       As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
    Dim Lo_RetVRI   As ListObject:      Set Lo_RetVRI = Prog_Coef_Ret_VRI.ListObjects(1)
    Dim RowNew      As ListRow
    Dim rowfind     As Variant
    
    Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso

    '- IMPRESCINDIBLE: M07 deja Lo_BD.ShowTotals = True. Si no se apaga aqui, la fila de
    '-  TOTALES entra en el recorrido y RowNew se queda Nothing -> Error 424 (bug historico).
    Dim Sw_TotBD        As Boolean:     Sw_TotBD = Lo_BD.ShowTotals
    Dim Sw_TotVRI       As Boolean:     Sw_TotVRI = Lo_RetVRI.ShowTotals
    Lo_BD.ShowTotals = False
    Lo_RetVRI.ShowTotals = False
    
    Dim T_ImpEmi        As Currency
    Dim T_ImpCob        As Currency
    Dim T_ImpPdt        As Currency

    ' =============  Preparar Hojas de Tablas ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, True)
    Call Rut_Lo_WrkSht_Preparar(Prog_Coef_Ret_VRI)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)


With Lo_BD.DataBodyRange
    ' Recorro toda la Tabla ------------------------------------------------------------------------
    For Cont = 1 To Lo_BD.ListRows.Count
        If .Cells(Cont, BD_Plan) <> Cod_Plan Then       '- <<<<<  Cambio PLAN BDatos NUEVO  >>>>
            Cod_Plan = .Cells(Cont, BD_Plan)
            Coef_VRI = .Cells(Cont, BD_Coef_VRI)
            Orgánica = .Cells(Cont, BD_Orgánica)
            Cont_Plan = Cont_Plan + 1
            '   Buscar Plan en Lo_RetVRI
            rowfind = Application.Match(Cod_Plan, Lo_RetVRI.DataBodyRange.Columns(CoefVRI_Plan), 0)
            If Not IsError(rowfind) Then    ' Plan Encontrado en Lo_RetVRI ==>>
                Set RowNew = Lo_RetVRI.ListRows(rowfind)
                RowNew.Range(CoefVRI_Concepto) = .Cells(Cont, BD_Concepto)
                RowNew.Range(CoefVRI_TipCurs) = .Cells(Cont, BD_Tipo_EP)
                If InStr(RowNew.Range(CoefVRI_Orgánica), Orgánica) = 0 And Orgánica <> "" Then
                    If RowNew.Range(CoefVRI_Orgánica) = "" Then
                        RowNew.Range(CoefVRI_Orgánica) = Orgánica & "_(" & .Cells(Cont, BD_C_Acad) & ")"
                    Else
                        RowNew.Range(CoefVRI_Orgánica) = RowNew.Range(CoefVRI_Orgánica) & vbLf & Orgánica & "_(" & .Cells(Cont, BD_C_Acad) & ")"
                    End If
                End If
            Else                            ' Plan No Existe, lo CREO en Lo_RetVRI==>>
                Set RowNew = Lo_RetVRI.ListRows.Add
                RowNew.Range(CoefVRI_Plan) = Cod_Plan
                RowNew.Range(CoefVRI_CoefVRI) = Coef_VRI
                If Orgánica <> "" Then RowNew.Range(CoefVRI_Orgánica) = Orgánica & "_(" & .Cells(Cont, BD_C_Acad) & ")"
                RowNew.Range(CoefVRI_Concepto) = .Cells(Cont, BD_Concepto)
                RowNew.Range(CoefVRI_TipCurs) = .Cells(Cont, BD_Tipo_EP)
                RowNew.Range(CoefVRI_ORden) = "x"
            End If
            
        End If
        '- Registro SIN PLAN: anomalia de datos (no debe existir ninguno). Lo contamos y avisamos
        '-  al final, en vez de abortar el proceso con Error 424 como hacia antes.
        If Cod_Plan = "" Then Regs_SinPlan = Regs_SinPlan + 1

        If Not RowNew Is Nothing Then
        If RowNew.Range(CoefVRI_ORden) = "x" Then
               RowNew.Range(CoefVRI_CoefVRI) = Application.Max(RowNew.Range(CoefVRI_CoefVRI), Coef_VRI)
            If RowNew.Range(CoefVRI_CoefVRI) = 0 Then
                If Prog__APP.Range("APP_EFP_o_CFC") = "EFP" Then    '- Cualificado: dependia de la hoja ACTIVA
                    RowNew.Range(CoefVRI_CoefVRI) = Lo_RetVRI.DataBodyRange.Cells(1, CoefVRI_CoefVRI)
                Else
                    RowNew.Range(CoefVRI_CoefVRI) = Lo_RetVRI.DataBodyRange.Cells(2, CoefVRI_CoefVRI)
                End If
            End If
        End If
        End If     '- Not RowNew Is Nothing
    Next
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)
    
    '- Visualizo el progreso
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                        " Actualizada la Tabla de Referencia de los Coef_VRI." & vbCrLf
End With    '- Lo_BD.DataBodyRange

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '- Anomalia de datos: NO debe existir ningun registro sin Plan ---------------------------
    If Regs_SinPlan > 0 Then
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & _
                Right(String(8, "_") & Format(Regs_SinPlan, "#,##0"), 8) & _
                " iOjo! Reg. de BDatos SIN PLAN (no deberia haber ninguno)." & vbCrLf
        MsgBox "Hay " & Regs_SinPlan & " registro(s) en BDatos SIN PLAN." & vbLf & vbLf & _
               "No deberia existir ninguno: revisa esos registros en Prog_BD.", _
               vbOKOnly + vbExclamation, "Actualizar Coef_VRI"
    End If
    '- Restaurar la fila de totales tal y como estaba al entrar ------------------------------
    Lo_BD.ShowTotals = Sw_TotBD
    Lo_RetVRI.ShowTotals = Sw_TotVRI
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA
'Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Lo_Coef_VRI_Actualizar
'===================================================================================================





