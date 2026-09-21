Attribute VB_Name = "M39_Resumen_Planes_C_Acad"
' Last Rev. 2026-09-21 12:12
'2026-02-05
'- M31_Cierre_Contable_PLANES
Option Explicit

'===================================================================================================
Sub RuT_Resumen_Planes_C_Acad()
'===================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim CursoAcad       As String:      CursoAcad = Prog__APP.Range("APP_CursAcad")
    Dim CursoAcadAnt    As String:      CursoAcadAnt = Prog__APP.Range("APP_C_Acad_Ant")
    Dim TipoCurso       As String:      TipoCurso = Prog__APP.Range("APP_EFP_o_CFC")

    Dim RngVisible  As Range
    Dim Lo_BD           As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
        Prog_BD.Unprotect:     Lo_BD.ShowTotals = False
        Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Dim Lo_BD_Filtrada  As ListObject
                
    Dim WsBuffer As Worksheet: Set WsBuffer = ThisWorkbook.Worksheets("Sheet_Buffer")  'hoja fija/oculta
    ' 1)Limpiar anterior tabla en Buffer
    If WsBuffer.ListObjects.Count > 0 Then WsBuffer.ListObjects(1).Delete
    ' 2) Ordenaciones
    Call Rut_Lo_Sort(Lo_BD, BD_C_Acad, xlAscending, True)
    Call Rut_Lo_Sort(Lo_BD, BD_ACont_Emi, xlAscending, False)
    Call Rut_Lo_Sort(Lo_BD, BD_Plan, xlAscending, False)
    ' 3) Filtros
    Lo_BD.Range.AutoFilter Field:=BD_Tipo_Rec, Criteria1:="<>_ERR_Date_"
    Lo_BD.Range.AutoFilter Field:=BD_ImpAdm, Criteria1:=">0"
    ' 4) Rango visible filtrado (incluye cabeceras)
    On Error Resume Next
        Set RngVisible = Lo_BD.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo SalirLimpiando
        If RngVisible Is Nothing Then GoTo SalirLimpiando
    On Error GoTo 0
    ' Limpia cualquier tabla previa en Buffer
    If WsBuffer.ListObjects.Count > 0 Then
        WsBuffer.ListObjects(1).Delete  '  .Unlist o .Delete si prefieres quitar formato tabla
    End If
    WsBuffer.Cells.Clear
    ' 6) Copiar solo lo visible
    RngVisible.Copy Destination:=WsBuffer.Range("A1")
    ' 7) Crear nuevo ListObject con los datos filtrados
    Set Lo_BD_Filtrada = WsBuffer.ListObjects.Add( _
                            SourceType:=xlSrcRange, _
                            Source:=WsBuffer.Range("A1").CurrentRegion, _
                            XlListObjectHasHeaders:=xlYes)
    ' 8) Llamar a tu rutina que trabaja con ListObject
        Call RuT_Resumen_Planes(Lo_BD_Filtrada, CursoAcad, TipoCurso, Sht__Inf_Rsm_Planes)
    GoTo Salir

    
SalirLimpiando:
    ' Opcional: quitar tabla en Buffer para dejar solo datos o la hoja limpia
    On Error Resume Next
    If Not WsBuffer Is Nothing Then
        If WsBuffer.ListObjects.Count > 0 Then WsBuffer.ListObjects(1).Unlist
        'Si quieres dejar la hoja totalmente vacía:
        WsBuffer.Cells.Clear
    End If
    On Error GoTo 0
    
Salir:
    ' Restaurar estado de Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Call Rut_EnableEvents_Status_Reset
    Application.Speech.Speak "Proceso Terminado"
End Sub '-------------------------------------------------------------------------------------------


'===================================================================================================
Sub RuT_Resumen_Planes(Lo_BD As ListObject, _
                                    CursoAcad As String, _
                                    TipoCurso As String, _
                                    Ws_Lista As Worksheet)
'===================================================================================================
    Dim AñoCont         As Integer:     AñoCont = Prog__APP.Range("APP_AñoCont")    '- Año Contable
    Dim ACont           As Integer:     ACont = Right(AñoCont, 2)
    Dim AñoContAnt      As Integer:     AñoContAnt = Left(CursoAcad, 4)             '- El 1º Año de Curso
    Dim AContAnt        As Integer:     AContAnt = Right(AñoContAnt, 2)             '- El 1º Año de Curso Corto
    Dim AñoContPos      As Integer:     AñoContPos = "20" & Right(CursoAcad, 2)     '- El 2º Año de Curso
    Dim AContPos        As Integer:     AContPos = Right(AñoContPos, 2)             '- El 2º Año de Curso Corto
    Dim Cont            As Long
    Dim ContIni         As Long:        ContIni = 1
    Dim Txt_Cabecera    As String
    Dim Cont_Plan       As Integer
    Dim PlanesSinCob    As Integer:     PlanesSinCob = 0
    
'    Dim RegsEmis            As Long     ' regs Emitidos
'    Dim Imp_Emis            As Currency
'    Dim Imp_EmisAdm         As Currency
'    Dim TImpEmisAdm         As Currency
'    Dim Imp_EmisAcad        As Currency
'    Dim TImpEmisAcad        As Currency
'
'    Dim Imp_EmisAnt         As Currency
'    Dim Imp_EmisAntAdm      As Currency
'    Dim TImpEmisAntAdm      As Currency
'    Dim Imp_EmisAntAcad     As Currency
'    Dim TImpEmisAntAcad     As Currency
'
'    Dim Imp_EmisPos         As Currency
'    Dim Imp_EmisPosAdm      As Currency
'    Dim TImpEmisPosAdm      As Currency
'    Dim Imp_EmisPosAcad     As Currency
'    Dim TImpEmisPosAcad     As Currency
'
'    Dim RegsAnul            As Long     ' regs Emitidos
'    Dim Imp_Anul            As Currency
'    Dim Imp_AnulAdm         As Currency
'    Dim TImpAnulAdm         As Currency
'    Dim Imp_AnulAcad        As Currency
'    Dim TImpAnulAcad        As Currency
'
'    Dim Imp_Cobr            As Currency
'    Dim Imp_CobrAdm         As Currency
'    Dim TImpCobrAdm         As Currency
'    Dim Imp_CobrAcad        As Currency
'    Dim TImpCobrAcad        As Currency
'
'    Dim Imp_CobrAnt         As Currency
'    Dim Imp_CobrAntAdm      As Currency
'    Dim TImpCobrAntAdm      As Currency
'    Dim Imp_CobrAntAcad     As Currency
'    Dim TImpCobrAntAcad     As Currency
'
'    Dim Imp_CobrPos         As Currency
'    Dim Imp_CobrPosAdm      As Currency
'    Dim TImpCobrPosAdm      As Currency
'    Dim Imp_CobrPosAcad     As Currency
'    Dim TImpCobrPosAcad     As Currency
'
'    Dim Imp__RDTpdt            As Currency
'    Dim Imp__RDTpdtAdm         As Currency
'    Dim TImpRDTpdtAdm         As Currency
'    Dim Imp__RDTpdtAcad        As Currency
'    Dim TImpRDTpdtAcad        As Currency
'
'    Dim Imp_SRDTpdt            As Currency
'    Dim Imp_SRDTpdtAdm         As Currency
'    Dim TImpSRDTpdtAdm         As Currency
'    Dim Imp_SRDTpdtAcad        As Currency
'    Dim TImpSRDTpdtAcad        As Currency
'
'    Dim Imp__ADx            As Currency
'    Dim Imp__ADxAdm         As Currency
'    Dim TImpADxAdm         As Currency
'    Dim Imp__ADxAcad        As Currency
'    Dim TImpADxAcad        As Currency
'
'    Dim Imp_Aplz            As Currency
'    Dim Imp_AplzAdm         As Currency
'    Dim TImpAplzAdm         As Currency
'    Dim Imp_AplzAcad        As Currency
'    Dim TImpAplzAcad        As Currency
'
'    Dim Imp_EjeAnt         As Currency
'    Dim Imp_EjeAntAdm      As Currency
'    Dim TImpEjeAntAdm      As Currency
'    Dim Imp_EjeAntAcad     As Currency
'    Dim TImpEjeAntAcad     As Currency
'
'    Dim Imp_PdtCob          As Currency
'    Dim Imp_PdtCobAdm       As Currency
'    Dim TimpPdtCobAdm       As Currency
'    Dim Imp_PdtCobAcad      As Currency
'    Dim TimpPdtCobAcad      As Currency
'
'    Dim Imp_PdtCobAnt       As Currency
'    Dim Imp_PdtCobAntAdm    As Currency
'    Dim TimpPdtCobAntAdm    As Currency
'    Dim Imp_PdtCobAntAcad   As Currency
'    Dim TimpPdtCobAntAcad   As Currency
'
'    Dim Imp_PdtCobPos       As Currency
'    Dim Imp_PdtCobPosAdm    As Currency
'    Dim TimpPdtCobPosAdm    As Currency
'    Dim Imp_PdtCobPosAcad   As Currency
'    Dim TimpPdtCobPosAcad   As Currency

    Dim Regs_Dev            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula
    Dim Imp_Devo            As Currency
    Dim RDev_Pag            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, Pagados
    Dim IDev_Pag            As Currency
    Dim RDevPdte            As Long     ' regs de Devolución (a pagar) / Ajuste Matrícula, No Pagados, pero Ajuste Matrícula NO deben ser pagados
    Dim IDevPdte            As Currency
    
'''    Dim Ws_Lista    As Worksheet:       Set Ws_Lista = Sheets("Cierre_Planes_CAcad_Ant_" & TipoCurso)
    Dim Lo_Rsm      As ListObject:      Set Lo_Rsm = Ws_Lista.ListObjects(1)
    Dim RowNew      As ListRow
    Dim Row_BD      As ListRow
    Dim rowfind     As Variant
    
    Dim Lo_RetVRI   As ListObject:      Set Lo_RetVRI = Prog_Coef_Ret_VRI.ListObjects(1)
    Call Rut_Lo_WrkSht_Preparar(Prog_Coef_Ret_VRI)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_ORden, xlAscending, True)
    Call Rut_Lo_Sort(Lo_RetVRI, CoefVRI_Plan, xlAscending, False)
    
    Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
'    On Error GoTo 0
    
    If Not Lo_Rsm.DataBodyRange Is Nothing Then Lo_Rsm.DataBodyRange.Delete

    '- Identificar Todos los Planes, Crea Dictionary para valores únicos (eficiente para grandes datos)
    Dim Collection_Planes      As New Collection
    Dim RngPlanes       As Range
    Set RngPlanes = Lo_BD.ListColumns(BD_Plan).DataBodyRange.SpecialCells(xlCellTypeVisible)
    Dim Celda As Range
    Dim CodPlan As String
    On Error Resume Next
    For Each Celda In RngPlanes.Cells
        If Not IsEmpty(Celda.Value) Then
            CodPlan = CStr(Celda.Value)
            Collection_Planes.Add Celda.Value, CodPlan   ' Solo agrega si no existe la CodPlan
        End If
    Next Celda
    On Error GoTo 0
    
    Call Rut_Lo_Sort(Lo_BD, BD_Tipo_Rec, xlAscending, False)
    ' Inicio Listado en Tabla excel
    Txt_Cabecera = "Resumen del Año Contable " & AñoCont & ", de los Planes de " & TipoCurso & "_" & CursoAcad & String(10, " ") & Now
    Ws_Lista.Range("k3") = Txt_Cabecera
    Ws_Lista.Range("h3") = Now
    Ws_Lista.Range("b3") = "Año Contable " & AñoCont & ", " & TipoCurso & "_" & CursoAcad
    
    Dim L0      As Integer:     L0 = Lo_Rsm.Range.Row - 2
    Dim L1      As Integer:     L1 = L0 + 1
    Dim C0      As Integer:     C0 = Lo_Rsm.Range.Column - 1
    
    '- Cabecera Tabla ------------------------------------------------------------------------------
    Cells(L0, C0 + Rsm_Imp_Emis) = "Emitido C_Acad " & CursoAcad
        Cells(L1, C0 + Rsm_Imp_Emis) = "Conta'" & ACont & vbLf & "Emitido" & vbLf & CursoAcad
        Cells(L1, C0 + Rsm_Imp_EmisAdm) = "Emitido" & vbLf & CursoAcad & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_EmisAcad) = "Emitido" & vbLf & CursoAcad & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_EmisAnt) = "Emitido " & AñoContAnt
        Cells(L1, C0 + Rsm_Imp_EmisAnt) = "Conta'" & ACont & vbLf & "Emitido" & vbLf & AñoContAnt
        Cells(L1, C0 + Rsm_Imp_EmisAntAdm) = "Emitido" & vbLf & AñoContAnt & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_EmisAntAcad) = "Emitido" & vbLf & AñoContAnt & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_EmisPos) = "Emitido " & AñoContPos
        Cells(L1, C0 + Rsm_Imp_EmisPos) = "Conta'" & ACont & vbLf & "Emitido" & vbLf & AñoContPos
        Cells(L1, C0 + Rsm_Imp_EmisPosAdm) = "Emitido" & vbLf & AñoContPos & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_EmisPosAcad) = "Emitido" & vbLf & AñoContPos & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_Anul) = "Anulado " & CursoAcad
        Cells(L1, C0 + Rsm_Imp_Anul) = "Importe" & vbLf & "Anulado" & vbLf & CursoAcad
        Cells(L1, C0 + Rsm_Imp_AnulAdm) = "Anulado" & vbLf & CursoAcad & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_AnulAcad) = "Anulado" & vbLf & CursoAcad & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_Cobr) = "Cobrado " & CursoAcad
        Cells(L1, C0 + Rsm_Imp_Cobr) = "Conta'" & ACont & vbLf & "Cobrado" & vbLf & CursoAcad
        Cells(L1, C0 + Rsm_Imp_CobrAdm) = "Cobrado" & vbLf & CursoAcad & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_CobrAcad) = "Cobrado" & vbLf & CursoAcad & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_CobrAnt) = "Cobrado " & AñoContAnt
        Cells(L1, C0 + Rsm_Imp_CobrAnt) = "Conta'" & ACont & vbLf & "Cobrado" & vbLf & AñoContAnt
        Cells(L1, C0 + Rsm_Imp_CobrAntAdm) = "Cobrado" & vbLf & AñoContAnt & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_CobrAntAcad) = "Cobrado" & vbLf & AñoContAnt & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_CobrPos) = "Cobrado " & AñoContPos
        Cells(L1, C0 + Rsm_Imp_CobrPos) = "Conta'" & ACont & vbLf & "Cobrado" & vbLf & AñoContPos
        Cells(L1, C0 + Rsm_Imp_CobrPosAdm) = "Cobrado" & vbLf & AñoContPos & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_CobrPosAcad) = "Cobrado" & vbLf & AñoContPos & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_RDTpdt) = "RDTpdt " & CursoAcad
        Cells(L1, C0 + Rsm_Imp_RDTpdt) = "Importe" & vbLf & "RDTpdt" & vbLf & CursoAcad
        Cells(L1, C0 + Rsm_Imp_RDTpdtAdm) = "RDTpdt" & vbLf & CursoAcad & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_RDTpdtAcad) = "RDTpdt" & vbLf & CursoAcad & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_RDTpdt) = "Pendiente de RDTpdt " & CursoAcad
        Cells(L1, C0 + Rsm_Imp_RDTpdt) = "Importe" & vbLf & "Pdte. RDTpdt" & vbLf & CursoAcad
        Cells(L1, C0 + Rsm_Imp_RDTpdtAdm) = "Pdte. RDTpdt" & vbLf & CursoAcad & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_RDTpdtAcad) = "Pdte. RDTpdt" & vbLf & CursoAcad & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_ADx) = "ADxAplz'" & AContAnt & " " & ChrW(&H21D2) & " Emi'" & AContAnt & " NoCob'" & AContAnt & " y Vto'" & AContPos
        Cells(L1, C0 + Rsm_Imp_ADx) = "Importe" & vbLf & "Emitido'" & AContAnt & vbLf & "ADxAplz'" & AContAnt
        Cells(L1, C0 + Rsm_Imp_ADxAdm) = "ADxAplz" & vbLf & AñoContAnt & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_ADxAcad) = "ADxAplz" & vbLf & AñoContAnt & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_Aplz) = "Aplazado'" & AContPos & " " & ChrW(&H21D2) & " Cob'" & AContPos & " y ADxAplz'" & AContAnt
        Cells(L1, C0 + Rsm_Imp_Aplz) = "Aplazado'" & AContPos & vbLf & "Cobrado'" & AContPos & vbLf & "ADxAplz'" & AContAnt
        Cells(L1, C0 + Rsm_Imp_AplzAdm) = "Aplazado" & vbLf & AñoContPos & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_AplzAcad) = "Aplazado" & vbLf & AñoContPos & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_EjeAnt) = "EjeAnt " & ChrW(&H21D2) & " Emi'" & AContAnt & " NoCob.'" & AContAnt & " y Cob.'" & AContPos
        Cells(L1, C0 + Rsm_Imp_EjeAnt) = "Rec. EjeAnt" & vbLf & "Emitido'" & AContAnt & vbLf & "No Cob.'" & AContAnt & vbLf & "y Cob.'" & AContPos
        Cells(L1, C0 + Rsm_Imp_EjeAntAdm) = "EjeAnt" & vbLf & AñoContPos & vbLf & "Imp. Adm."
        Cells(L1, C0 + Rsm_Imp_EjeAntAcad) = "EjeAnt" & vbLf & AñoContPos & vbLf & "Imp. Acad."
        
    Cells(L0, C0 + Rsm_Imp_PdtCob) = "Pendiente de Cobro " & CursoAcad
        Cells(L1, C0 + Rsm_Imp_PdtCob) = "Pendiente" & vbLf & "Cobro " & vbLf & CursoAcad
        Cells(L1, C0 + Rsm_Imp_PdtCobAdm) = "Pdte.Cob. " & vbLf & CursoAcad & vbLf & "Adm."
        Cells(L1, C0 + Rsm_Imp_PdtCobAcad) = "Pdte.Cob. " & vbLf & CursoAcad & vbLf & "Acad."
        
    Cells(L0, C0 + Rsm_Imp_PdtCobAnt) = "Pendiente de Cobro de " & AñoContAnt
        Cells(L1, C0 + Rsm_Imp_PdtCobAnt) = "Pendiente" & vbLf & "Cobro " & vbLf & AñoContAnt
        Cells(L1, C0 + Rsm_Imp_PdtCobAntAdm) = "Pdte.Cob. " & vbLf & AñoContAnt & vbLf & "Adm."
        Cells(L1, C0 + Rsm_Imp_PdtCobAntAcad) = "Pdte.Cob. " & vbLf & AñoContAnt & vbLf & "Acad."
        
    Cells(L0, C0 + Rsm_Imp_PdtCobPos) = "Pendiente de Cobro de " & AñoContPos
        Cells(L1, C0 + Rsm_Imp_PdtCobPos) = "Pendiente" & vbLf & "Cobro " & vbLf & AñoContPos
        Cells(L1, C0 + Rsm_Imp_PdtCobPosAdm) = "Pdte.Cob. " & vbLf & AñoContPos & vbLf & "Adm."
        Cells(L1, C0 + Rsm_Imp_PdtCobPosAcad) = "Pdte.Cob. " & vbLf & AñoContPos & vbLf & "Acad."
        
    Cells(L0, C0 + Rsm_Imp_PdtCobADx) = "Pendiente de Cobro de ADxAplz'" & AContPos
        Cells(L1, C0 + Rsm_Imp_PdtCobADx) = "Pendiente" & vbLf & "Cobro ADxAplz" & vbLf & AñoContPos
        Cells(L1, C0 + Rsm_Imp_PdtCobADxAdm) = "Pdte.Cob.ADx " & vbLf & AñoContPos & vbLf & "Adm."
        Cells(L1, C0 + Rsm_Imp_PdtCobADxAcad) = "Pdte.Cob.ADx " & vbLf & AñoContPos & vbLf & "Acad."
        
        Range(Cells(L1, C0 + 4), Cells(L1 + 1, C0 + Lo_Rsm.Range.Columns.Count - 1)).Select
        Selection.Interior.Color = RGB(180, 198, 231)    '-Azul Claro
        Selection.Font.ColorIndex = xlAutomatic
        Union(Cells(L1, C0 + Rsm_Imp_Emis), Cells(L1, C0 + Rsm_Imp_EmisAnt), Cells(L1, C0 + Rsm_Imp_EmisPos), Cells(L1, C0 + Rsm_Imp_Anul), _
              Cells(L1, C0 + Rsm_Imp_Cobr), Cells(L1, C0 + Rsm_Imp_CobrAnt), Cells(L1, C0 + Rsm_Imp_CobrPos), _
              Cells(L1, C0 + Rsm_Imp_RDTpdt), Cells(L1, C0 + Rsm_Imp_RDTpdt), Cells(L1, C0 + Rsm_Imp_ADx), Cells(L1, C0 + Rsm_Imp_Aplz), _
              Cells(L1, C0 + Rsm_Imp_EjeAnt), Cells(L1, C0 + Rsm_Imp_PdtCob), Cells(L1, C0 + Rsm_Imp_PdtCobAnt), _
              Cells(L1, C0 + Rsm_Imp_PdtCobPos), Cells(L1, C0 + Rsm_Imp_PdtCobADx)).Select
    With Lo_Rsm.HeaderRowRange
        Union(Selection, .Cells(Rsm_Imp_Emis), .Cells(Rsm_Imp_EmisAnt), .Cells(Rsm_Imp_EmisPos), .Cells(Rsm_Imp_Anul), _
              .Cells(Rsm_Imp_Cobr), .Cells(Rsm_Imp_CobrAnt), .Cells(Rsm_Imp_CobrPos), _
              .Cells(Rsm_Imp_RDTpdt), .Cells(Rsm_Imp_RDTpdt), .Cells(Rsm_Imp_ADx), .Cells(Rsm_Imp_Aplz), _
              .Cells(Rsm_Imp_EjeAnt), .Cells(Rsm_Imp_PdtCob), .Cells(Rsm_Imp_PdtCobAnt), _
              .Cells(Rsm_Imp_PdtCobPos), .Cells(Rsm_Imp_PdtCobADx)).Select
        With Selection
            .Interior.Color = RGB(31, 78, 120)          '- Azul oscuro
            .Font.Color = RGB(255, 255, 0)     '-Amarillo
            .Font.Name = "Arial Black"
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .WrapText = True
            .InsertIndent 1
        End With
        Union(.Cells(Rsm_Imp_AnulAdm), .Cells(Rsm_Imp_AnulAcad), .Cells(Rsm_Imp_PdtCobAdm), .Cells(Rsm_Imp_PdtCobAcad)).Select
            Selection.Font.Color = RGB(192, 0, 0)    '-Rojo oscuro
    End With

'    '- Visualizo el progreso ----------------------------------------------------------------------
'    Form_Menu.Lb_Tit_Informe.Caption = "Progreso de la Tarea."
'    Form_Menu.TB_Informe = Txt_Cabecera & vbCrLf & vbCrLf
'    Txt_Cabecera = String(16, " ") & AñoCont - 1 & String(11, " ") & AñoCont & String(11, " ") & AñoCont & "         Cob." & AñoCont & "       Cob." & AñoCont & vbLf & _
'                   "    Plan       Imp_Emi        Imp_Emi        ADxAplz       Aplazado        Eje_Ant"
'    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Txt_Cabecera & vbLf
    
    
    '- Recorro la Collection con todos los Planes
         Dim Cod_Plan As Variant
    With Lo_BD.DataBodyRange
    For Each Cod_Plan In Collection_Planes
        
        Cont_Plan = Cont_Plan + 1
        Set RowNew = Lo_Rsm.ListRows.Add
        
        RowNew.Range(Rsm_Orden) = Cont_Plan
        RowNew.Range(Rsm_Plan) = Cod_Plan
        '- Localizo el Coef_VRI
        rowfind = Application.Match(Cod_Plan, Lo_RetVRI.DataBodyRange.Columns(CoefVRI_Plan), 0)
        If Not IsError(rowfind) Then    ' Plan Encontrado en Lo_RetVRI ==>>
            RowNew.Range(Rsm_Coel_VRI) = Lo_RetVRI.DataBodyRange.Cells(rowfind, CoefVRI_CoefVRI)
        Else                            ' Plan No Existe, pongo valores estándares ==>>
            If RowNew.Range(CoefVRI_CoefVRI) = 0 Then
                If TipoCurso = "EFP" Then
                    RowNew.Range(Rsm_Coel_VRI) = Lo_RetVRI.DataBodyRange.Cells(1, CoefVRI_CoefVRI)
                Else
                    RowNew.Range(Rsm_Coel_VRI) = Lo_RetVRI.DataBodyRange.Cells(2, CoefVRI_CoefVRI)
                End If
            End If
        End If
            
'- Importe Emis C_Acad -----------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_Emis) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan)
            RowNew.Range(Rsm_RegsEmis) = Application.CountIfs( _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan)
            '- Importe-EmisAdm ---------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_EmisAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan)

            '- ¡¡ NO se puede hacer un SumIfs porque falla debido a la manipulación del Imp.Acad que utilizan con Imp. Negativos para ajustes de Matrículas !! ----
            '- Importe EmisAcad  -------------------------------------------------------------------
            If RowNew.Range(Rsm_Imp_Emis) = 0 Then GoTo Siguiente_Plan    '- Si todo está cobrado en año anterior
            RowNew.Range(Rsm_Imp_EmisAcad) = RowNew.Range(Rsm_Imp_Emis) - RowNew.Range(Rsm_Imp_EmisAdm)
            
'- Importe Emis_AñoContAnt -------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_EmisAnt) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Emi), AñoContAnt)
            '- Importe-EmisAntAdm ------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_EmisAntAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                    .Columns(BD_Rec_Imp_Adm), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Emi), AñoContAnt)
            '- Importe EmisAntAcad  ----------------------------------------------------------------
            RowNew.Range(Rsm_Imp_EmisAntAcad) = RowNew.Range(Rsm_Imp_EmisAnt) - RowNew.Range(Rsm_Imp_EmisAntAdm)

'- Importe Emis_AñoContPos -------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_EmisPos) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Emi), AñoContPos)
            '- Importe-EmisPosAdm ------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_EmisPosAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                    .Columns(BD_Rec_Imp_Adm), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Emi), AñoContPos)
            '- Importe EmisPosAcad  ----------------------------------------------------------------
            RowNew.Range(Rsm_Imp_EmisPosAcad) = RowNew.Range(Rsm_Imp_EmisPos) - RowNew.Range(Rsm_Imp_EmisPosAdm)

'- Importe Anul C_Acad -----------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_Anul) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                            .Columns(BD_ImpRec), "<0", _
                                                            .Columns(BD_Plan), Cod_Plan)
            If RowNew.Range(Rsm_Imp_Anul) < 0 Then
                RowNew.Range(Rsm_Imp_AnulAdm) = "?"
                RowNew.Range(Rsm_Imp_AnulAcad) = "?"
            End If

'- Importe Cobr C_Acad -----------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_Cobr) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                            .Columns(BD_ImpCob), ">0", _
                                                            .Columns(BD_Plan), Cod_Plan)
            '- Importe Adm Recibos Cob_AñoCont -----------------------------------------------------
            RowNew.Range(Rsm_Imp_CobrAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                .Columns(BD_ImpCob), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan)
            '- Importe Acad Recibos Cob_AñoCont ----------------------------------------------------
            RowNew.Range(Rsm_Imp_CobrAcad) = RowNew.Range(Rsm_Imp_Cobr) - RowNew.Range(Rsm_Imp_CobrAdm)

'- Importe Cobr AñoContAnt -------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_CobrAnt) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                                .Columns(BD_ImpCob), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Cob), AñoContAnt)
            '- Importe Adm Recibos Cob_AñoContAnt --------------------------------------------------
            RowNew.Range(Rsm_Imp_CobrAntAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                    .Columns(BD_ImpCob), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Cob), AñoContAnt)
            '- Importe Acad Recibos Cob_AñoContAnt -------------------------------------------------
            RowNew.Range(Rsm_Imp_CobrAntAcad) = RowNew.Range(Rsm_Imp_CobrAnt) - RowNew.Range(Rsm_Imp_CobrAntAdm)

'- Importe Cobr AñoContPos -------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_CobrPos) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                                .Columns(BD_ImpCob), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Cob), AñoContPos)
            '- Importe Adm Recibos Cob_AñoContPos --------------------------------------------------
            RowNew.Range(Rsm_Imp_CobrPosAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                    .Columns(BD_ImpCob), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Cob), AñoContPos)
            '- Importe Acad Recibos Cob_AñoContPos -------------------------------------------------
            RowNew.Range(Rsm_Imp_CobrPosAcad) = RowNew.Range(Rsm_Imp_CobrPos) - RowNew.Range(Rsm_Imp_CobrPosAdm)

'- Importe Recibos RDT -----------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_RDT) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                            .Columns(BD_ImpCob), ">0", _
                                                            .Columns(BD_Plan), Cod_Plan, _
                                                            .Columns(BD_RDT), "<>")
            '- Importe Adm Recibos Cob_AñoCont -----------------------------------------------------
            RowNew.Range(Rsm_Imp_RDTAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                    .Columns(BD_ImpCob), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_RDT), "<>")
            '- Importe Acad Recibos Cob_AñoCont ----------------------------------------------------
            RowNew.Range(Rsm_Imp_RDTAcad) = RowNew.Range(Rsm_Imp_RDT) - RowNew.Range(Rsm_Imp_RDTAdm)

'- Importe Sin_RDTpdt ------------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_RDTpdt) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                            .Columns(BD_ImpCob), ">0", _
                                                            .Columns(BD_Plan), Cod_Plan, _
                                                            .Columns(BD_RDT), "=")
            '- Importe Adm Recibos Cob_AñoCont -----------------------------------------------------
            RowNew.Range(Rsm_Imp_RDTpdtAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                            .Columns(BD_ImpCob), ">0", _
                                                            .Columns(BD_Plan), Cod_Plan, _
                                                            .Columns(BD_RDT), "=")
            '- Importe Acad Recibos Cob_AñoCont ----------------------------------------------------
            RowNew.Range(Rsm_Imp_RDTpdtAcad) = RowNew.Range(Rsm_Imp_RDTpdt) - RowNew.Range(Rsm_Imp_RDTpdtAdm)

'- Importe ADxAplz ---------------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_ADx) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                            .Columns(BD_ImpRec), ">0", _
                                                            .Columns(BD_Plan), Cod_Plan, _
                                                            .Columns(BD_ACont_Vto), AñoCont, _
                                                            .Columns(BD_ACont_Cob), "<>" & AñoCont - 1, _
                                                            .Columns(BD_ACont_Emi), AñoCont - 1)
            '- Importe-Adm ADxAplz -----------------------------------------------------------------
            RowNew.Range(Rsm_Imp_ADxAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Vto), AñoCont, _
                                                                .Columns(BD_ACont_Cob), "<>" & AñoCont - 1, _
                                                                .Columns(BD_ACont_Emi), AñoCont - 1)
            '- Importe Acad ADxAplz ----------------------------------------------------------------
            RowNew.Range(Rsm_Imp_ADxAcad) = RowNew.Range(Rsm_Imp_ADx) - RowNew.Range(Rsm_Imp_ADxAdm)

'- Importe Aplazado (ADxAplz el AñoCont -1 ---------------------------------------------------------
            RowNew.Range(Rsm_Imp_Aplz) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                            .Columns(BD_ImpRec), ">0", _
                                                            .Columns(BD_Plan), Cod_Plan, _
                                                            .Columns(BD_ACont_Vto), AñoCont, _
                                                            .Columns(BD_ACont_Cob), "<>" & AñoCont - 1, _
                                                            .Columns(BD_ACont_Emi), AñoCont - 1)
            '- Importe-Adm ADxAplz -----------------------------------------------------------------
            RowNew.Range(Rsm_Imp_AplzAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                            .Columns(BD_ImpRec), ">0", _
                                                            .Columns(BD_Plan), Cod_Plan, _
                                                            .Columns(BD_ACont_Vto), AñoCont, _
                                                            .Columns(BD_ACont_Cob), "<>" & AñoCont - 1, _
                                                            .Columns(BD_ACont_Emi), AñoCont - 1)
            '- Importe Acad ADxAplz ----------------------------------------------------------------
            RowNew.Range(Rsm_Imp_AplzAcad) = RowNew.Range(Rsm_Imp_Aplz) - RowNew.Range(Rsm_Imp_AplzAdm)

'- Importe EjeAnt ( Emitido el AñoCont-1, Vto AñoCont-1 y Cob Añocont ------------------------------
            RowNew.Range(Rsm_Imp_EjeAnt) = Application.SumIfs(.Columns(BD_ImpCob), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Emi), AñoContAnt, _
                                                                .Columns(BD_ACont_Vto), AñoContAnt, _
                                                                .Columns(BD_ACont_Cob), "=" & AñoContPos)
            '- Importe Adm EjeAnt ( Emitido el AñoCont-1, Vto AñoCont-1 y Cob Añocont --------------
            RowNew.Range(Rsm_Imp_EjeAntAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Emi), AñoContAnt, _
                                                                .Columns(BD_ACont_Vto), AñoContAnt, _
                                                                .Columns(BD_ACont_Cob), "=" & AñoContPos)
            '- Importe Acad EjeAnt ( Emitido el AñoCont-1, Vto AñoCont-1 y Cob Añocont -------------
            RowNew.Range(Rsm_Imp_EjeAntAcad) = RowNew.Range(Rsm_Imp_EjeAnt) - RowNew.Range(Rsm_Imp_EjeAntAdm)

'- Importe Pdte_Cob --------------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCob) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Cob), "=")
            '- Importe Adm Pdte_Cob ----------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                    .Columns(BD_ImpRec), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Cob), "=")
            '- Importe Acad Pdte_Cob ---------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobAcad) = RowNew.Range(Rsm_Imp_PdtCob) - RowNew.Range(Rsm_Imp_PdtCobAdm)
    
'- Importe Pdte_CobAnt -----------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobAnt) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                                    .Columns(BD_ImpRec), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Emi), AñoContAnt, _
                                                                    .Columns(BD_ACont_Vto), AñoContAnt, _
                                                                    .Columns(BD_ACont_Cob), "=")
            '- Importe Adm Pdte_CobAnt -------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobAntAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                    .Columns(BD_ImpRec), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Emi), AñoContAnt, _
                                                                    .Columns(BD_ACont_Vto), AñoContAnt, _
                                                                    .Columns(BD_ACont_Cob), "=")
            '- Importe Acad Pdte_CobAnt ------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobAntAcad) = RowNew.Range(Rsm_Imp_PdtCobAnt) - RowNew.Range(Rsm_Imp_PdtCobAntAdm)
    
'- Importe Pdte_CobPos -----------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobPos) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                                    .Columns(BD_ImpRec), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Emi), AñoContPos, _
                                                                    .Columns(BD_ACont_Cob), "=")
            '- Importe Adm Pdte_CobPos -------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobPosAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                    .Columns(BD_ImpRec), ">0", _
                                                                    .Columns(BD_Plan), Cod_Plan, _
                                                                    .Columns(BD_ACont_Emi), AñoContPos, _
                                                                    .Columns(BD_ACont_Cob), "=")
            '- Importe Acad Pdte_CobPos ------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobPosAcad) = RowNew.Range(Rsm_Imp_PdtCobPos) - RowNew.Range(Rsm_Imp_PdtCobPosAdm)
    
'- Importe Pdte_CobADx -----------------------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobADx) = Application.SumIfs(.Columns(BD_ImpRec), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Emi), AñoContAnt, _
                                                                .Columns(BD_ACont_Vto), AñoContPos, _
                                                                .Columns(BD_ACont_Cob), "=")
            '- Importe Adm Pdte_CobADx -------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobADxAdm) = Application.SumIfs(.Columns(BD_Rec_Imp_Adm), _
                                                                .Columns(BD_ImpRec), ">0", _
                                                                .Columns(BD_Plan), Cod_Plan, _
                                                                .Columns(BD_ACont_Emi), AñoContAnt, _
                                                                .Columns(BD_ACont_Vto), AñoContPos, _
                                                                .Columns(BD_ACont_Cob), "=")
            '- Importe Acad Pdte_CobADx ------------------------------------------------------------
            RowNew.Range(Rsm_Imp_PdtCobADxAcad) = RowNew.Range(Rsm_Imp_PdtCobADx) - RowNew.Range(Rsm_Imp_PdtCobADxAdm)
    
    
''- Informe ----------------------------------------------------------------------------------------
'            Form_Menu.TB_Informe = Form_Menu.TB_Informe & Format(Cont_Plan, "00") & "º " & Cod_Plan & " " & _
'                Right(String(12, " ") & Format(Imp_EmisAnt, "#,##0.00"), 14) & _
'                IIf(Imp_Cobr > 0, " " & Right(String(12, "·") & Format(Imp_EmisPos, "#,##0.00"), 14), "  No hay Cobros") & _
'                IIf(Imp__RDTpdt > 0, " " & Right(String(12, "·") & Format(Imp__ADx, "#,##0.00"), 14), " " & String(14, "·")) & _
'                IIf(Imp_SRDTpdt > 0, " " & Right(String(12, "·") & Format(Imp_Aplz, "#,##0.00"), 14), " " & String(14, "·")) & _
'                IIf(Imp_PdtCob > 0, " " & Right(String(12, "·") & Format(Imp_EjeAnt, "#,##0.00"), 14), " " & String(14, "·")) & vbLf
'
'            ' Intercalo Cabecera Columnas
'            If Cont_Plan Mod 20 = 0 Then Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Txt_Cabecera & vbLf
    
'            If Imp_Cobr = 0 Then PlanesSinCob = PlanesSinCob + 1
Siguiente_Plan:
    Next Cod_Plan
    End With    ' Lo_BD.DataBodyRange
    
    Lo_Rsm.DataBodyRange.Select
        With Selection
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.ColorIndex = xlAutomatic
        .Font.ThemeFont = xlThemeFontNone
        .Font.Bold = False
        .NumberFormat = "#,##0.00;[Red]-#,##0.00;"
            
'            .HorizontalAlignment = xlGeneral
'            .VerticalAlignment = xlCenter
        End With
        Range("Tb_Rsm_Planes[[Orden]:[Regs Emis]]").Select
        Selection.NumberFormat = "General"

'            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & Txt_Cabecera
'            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & " Totales "
'            Form_Menu.TB_Informe = Form_Menu.TB_Informe & _
'                Right(String(12, " ") & Format(TImpEmisAnt, "#,##0.00"), 14) & _
'                Right(String(12, " ") & Format(TImpEmisPos, "#,##0.00"), 15) & _
'                Right(String(12, " ") & Format(TImpADx, "#,##0.00"), 15) & _
'                Right(String(12, " ") & Format(TImpAplz, "#,##0.00"), 15) & _
'                Right(String(12, " ") & Format(TImpEjeAnt, "#,##0.00"), 15) & vbLf
'
'
'            If PlanesSinCob > 1 Then
'                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡¡ " & PlanesSinCob & " Planes sin Cobros !!"
'            ElseIf PlanesSinCob = 1 Then
'                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡¡ 1 Plan sin Cobros !!"
'            Else
'                Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡ No hay planes sin Cobros !"
'            End If
'
''- Visualizo el progreso --------------------------------------------------------------------------
'Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & _
'                                                "¡¡¡ Proceso concluido !!! día: " & Now() & _
'                                                " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg."

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA
'''Prog_BD.Visible = xlSheetVeryHidden
Rut_On_Functions
End Sub     ' RuT_Listar_Planes   ------------------------------------------------------------------
'===================================================================================================




