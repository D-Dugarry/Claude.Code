Attribute VB_Name = "M_900_____"
Option Explicit

Sub Rut_JyCList_LstObj_Refresh()

    Dim Color       As Integer:     Color = 19
    Dim LinX      As Long
    Dim RefAnt      As String:      RefAnt = ""
    Dim SiglasAnt   As String:      SiglasAnt = ""
    Dim RowTPV          As ListRow
    Dim RowCta          As ListRow
    Dim RowLst          As ListRow
    Dim Lin_Lst     As Variant
    Dim Lo_TPV      As ListObject:      Set Lo_TPV = Prog_TPV_Tb.ListObjects(1)
    Dim Lo_Lst      As ListObject:      Set Lo_Lst = Prog_JyC_List.ListObjects(1)

    Rut_Off_Functions

    H_INICI.Select
    H_INICI.Unprotect
'-- TPV ----------------------------------------------------------------------------------------------------- C_TPV_Liq_Ref
    Prog_TPV_Tb.Visible = xlSheetVisible
    Call Rut_LstObj_Filtros_Quitar(Lo_TPV)
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_N_Liq, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_TPV, C_TPV_Siglas, xlAscending, True)
    Call Rut_LstObj_Filtros_Quitar(Lo_Lst)
For LinX = 1 To Lo_TPV.ListRows.Count
    Set RowTPV = Lo_TPV.ListRows(LinX)
    If RowTPV.Range(C_TPV_Siglas) = "" Then GoTo SigLinX
    If Right(RowTPV.Range(C_TPV_F_Pago), 4) <> Prog__APP.Range("App_AñoCont") Then GoTo SigLinX
    If RefAnt <> RowTPV.Range(C_TPV_Siglas) Then
        RefAnt = RowTPV.Range(C_TPV_Siglas)
        
        Lin_Lst = Application.Match(RowTPV.Range(C_TPV_Siglas), Lo_Lst.DataBodyRange.Columns(1), 0)
        
        Set RowLst = Lo_Lst.ListRows(Lin_Lst)
        
        If RowLst.Range(C_JCL_Tipo) = "" Then RowLst.Range(C_JCL_Tipo) = "TPV"
        If InStr(RowLst.Range(C_JCL_Tipo), "TPV") = 0 Then RowLst.Range(C_JCL_Tipo) = RowLst.Range(C_JCL_Tipo) & ",TPV"
        
        RowLst.Range(TpvRef01) = RowTPV.Range(C_TPV_Liq_Ref)
        
        
        
        
        
        
        
    Else
        '- Acumular Importes ----
        RowLst.Range(C_JCL_TPV_RecEmi) = RowLst.Range(C_JCL_TPV_RecEmi) + RowTPV.Range(C_TPV_Imp)
        If RowTPV.Range(C_TPV_Neto) > 0 And RowTPV.Range(C_TPV_Imp) = "" Then
            RowLst.Range(C_JCL_TPV_ImpRec) = RowLst.Range(C_JCL_TPV_ImpRec) + RowTPV.Range(C_TPV_Neto)
            RowLst.Range(C_JCL_TPV_RecEmi) = RowLst.Range(C_JCL_TPV_RecEmi) + RowTPV.Range(C_TPV_Neto)
        Else
            RowLst.Range(C_JCL_TPV_ImpRec) = RowLst.Range(C_JCL_TPV_ImpRec) + RowTPV.Range(C_TPV_Imp)
        End If
        RowLst.Range(C_JCL_TPV_ComBco) = RowLst.Range(C_JCL_TPV_ComBco) + RowTPV.Range(C_TPV_ComBco)
        RowLst.Range(C_JCL_TPV_Neto) = RowLst.Range(C_JCL_TPV_Neto) + RowTPV.Range(C_TPV_Neto)
        If RowTPV.Range(C_TPV_DI) <> "" Then RowLst.Range(C_JCL_TPV_Dev) = RowLst.Range(C_JCL_TPV_Dev) + RowTPV.Range(C_TPV_DI_Imp)
        '- Acumular Observaciones ----
        If RowTPV.Range(C_TPV_Obs) <> "" Then
            If RowLst.Range(C_JCL_Obs) <> "" Then
                RowLst.Range(C_JCL_Obs) = RowLst.Range(C_JCL_Obs) & vbLf & "-.-" & vbLf & RowTPV.Range(C_TPV_Obs)
            Else
                RowLst.Range(C_JCL_Obs) = RowTPV.Range(C_TPV_Obs)
            End If
        End If
        '- Acumular Núm de Liq por si hay DEV de RDT ----
        If InStr(RowLst.Range(C_JCL_N_Liq), RowTPV.Range(C_TPV_N_Liq)) = 0 And RowLst.Range(C_JCL_N_Liq) <> "" Then RowLst.Range(C_JCL_N_Liq) = RowLst.Range(C_JCL_N_Liq) & vbLf & "-.-" & vbLf & RowTPV.Range(C_TPV_N_Liq)
        '- Acumular Núm de Ref-Sol ----
        If InStr(RowLst.Range(C_JCL_Ref_Sol), LCase(RowTPV.Range(C_TPV_Ref_Sol))) = 0 And RowLst.Range(C_JCL_Ref_Sol) <> "" Then RowLst.Range(C_JCL_Ref_Sol) = RowLst.Range(C_JCL_Ref_Sol) & vbLf & "-.-" & vbLf & LCase(RowTPV.Range(C_TPV_Ref_Sol))
        '- Acumular Orgánica ----
        If InStr(RowLst.Range(C_JCL_Org), UCase(RowTPV.Range(C_TPV_Org))) = 0 And RowLst.Range(C_JCL_Org) <> "" Then RowLst.Range(C_JCL_Org) = RowLst.Range(C_JCL_Org) & vbLf & "-.-" & vbLf & RowTPV.Range(C_TPV_Org)
        '- Acumular JI ----
        If InStr(RowLst.Range(C_JCL_JI), RowTPV.Range(C_TPV_JI)) = 0 And RowLst.Range(C_JCL_JI) <> "" Then RowLst.Range(C_JCL_JI) = RowLst.Range(C_JCL_JI) & vbLf & "-.-" & vbLf & RowTPV.Range(C_TPV_JI)
        '- Acumular ExpAdm ----
        If InStr(RowLst.Range(C_JCL_ExpAdm), RowTPV.Range(C_TPV_ExpAdm)) = 0 And RowLst.Range(C_JCL_ExpAdm) <> "" Then RowLst.Range(C_JCL_ExpAdm) = RowLst.Range(C_JCL_ExpAdm) & vbLf & "-.-" & vbLf & RowTPV.Range(C_TPV_ExpAdm)
        '- Acumular RDT ----
        If InStr(RowLst.Range(C_JCL_N_RDT), UCase(RowTPV.Range(C_TPV_RDT))) = 0 And RowLst.Range(C_JCL_N_RDT) <> "" Then RowLst.Range(C_JCL_N_RDT) = RowLst.Range(C_JCL_N_RDT) & vbLf & "-.-" & vbLf & RowTPV.Range(C_TPV_RDT)
        '- Acumular DI ----
        If InStr(RowLst.Range(C_JCL_DI), UCase(RowTPV.Range(C_TPV_DI))) = 0 And RowLst.Range(C_JCL_DI) <> "" Then RowLst.Range(C_JCL_DI) = RowLst.Range(C_JCL_DI) & vbLf & "-.-" & vbLf & RowTPV.Range(C_TPV_DI)
        '- Acumular PMP ----
        If InStr(RowLst.Range(C_JCL_PMP), UCase(RowTPV.Range(C_TPV_PMP))) = 0 And RowLst.Range(C_JCL_PMP) <> "" Then RowLst.Range(C_JCL_PMP) = RowLst.Range(C_JCL_PMP) & vbLf & "-.-" & vbLf & RowTPV.Range(C_TPV_PMP)
    End If

SigLinX:
Next LinX



    Call Rut_LstObj_Sort(Lo_Lst, C_JCL_N_Liq, xlAscending, True)
    Call Rut_LstObj_Sort(Lo_Lst, C_JCL_Siglas, xlAscending, True)
    '    Lo_Lst.DataBodyRange.Rows.AutoFit
    Lo_Lst.DataBodyRange.EntireRow.AutoFit
    '- Asigna color de fonda a cada grupo de JyC -----
    For LinX = 1 To Lo_Lst.ListRows.Count
        Set RowLst = Lo_Lst.ListRows(LinX)
        If RowLst.Range(C_JCL_N_Liq) = "x" Then RowLst.Range(C_JCL_Cta_Imp) = "Pdte. Liq." '.ClearContents
        If SiglasAnt <> RowLst.Range(C_JCL_Siglas) Then
            SiglasAnt = RowLst.Range(C_JCL_Siglas)
            If Color = 19 Then
                Color = 20
            Else
                Color = 19
            End If
        End If
        RowLst.Range.Interior.ColorIndex = Color
    Next LinX


Restablecer_Valores:
    Rut_On_Functions
    MsgBx_Msg = "¡ Proceso completado !"
    MsgBx_Title = "Proceso: Actualizar Tabla Resumen de Liquidaciones de JyC."
    Load Form_MsgBox: Call Form_MsgBox.SetParameter(, , , , "Exclam"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
    
'    Application.Speech.Speak "Proceso completado puede verificar el resultado."
End Sub

