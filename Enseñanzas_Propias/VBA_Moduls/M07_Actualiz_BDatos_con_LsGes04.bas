Attribute VB_Name = "M07_Actualiz_BDatos_con_LsGes04"
' Last Rev. 2026-09-21 12:12
'2026-01-09
Option Explicit

'            - Actualizar BDatos con Lo_Ges04
'            - Recorro toda LsGes04 para actualizar BDatos
'            --- Referencias IGUALES                            <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios
'            --- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO     <<<<  AÑADO UN NUEVO REGISTRO a BDatos
'            --- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO   <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
'            - Copiar Todos los Registros "Deleted" en Lo_Deleted y Borrarlos de BDatos

'- -------------------------------------------------------------------------------------------------
'- Actualizar BDatos con Lo_Ges04 ------------------------------------------------------------------
'- -------------------------------------------------------------------------------------------------
Sub RuT_Actualizar_BDatos_con_LsGes04()
Debug.Print ">>> RuT_Actualizar_BDatos_con_LsGes04"
    Dim F_Actualiz          As String:  F_Actualiz = Now()
    Dim Incidencia          As String
    Dim TxT_Progreso        As String
    Dim Ref_Ant             As String:      Ref_Ant = ""
    Dim Cont                As Integer
    Dim Cont_Fail           As Long:    Cont_Fail = 0
    Dim Cont_Repes          As Long:    Cont_Repes = 0
    Dim Cont_Nuevo          As Long:    Cont_Nuevo = 0
    Dim Cont_Modif          As Long:    Cont_Modif = 0
    Dim Cont_Deleted        As Long:    Cont_Deleted = 0
    Dim Cont_DeletedconJI   As Long:    Cont_DeletedconJI = 0
    Dim Cont_Mat_Anul       As Long:    Cont_Mat_Anul = 0
    Dim Chg_ImpRec          As Long:    Chg_ImpRec = 0
    Dim Chg_ImpCob          As Long:    Chg_ImpCob = 0
    Dim Chg_ImpAdm          As Long:    Chg_ImpAdm = 0
    Dim F_BD                As Long:    F_BD = 1
    Dim F_G4                As Long:    F_G4 = 1
    Dim RwG4            As ListRow
    Dim RwDB            As ListRow
    Dim RngG4           As Range
    Dim RngPH           As Range
    Dim rowfind             As Variant
        
    Dim Lo_BD               As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
    Dim Lo_Ges04            As ListObject:      Set Lo_Ges04 = Prog_LsGes04.ListObjects(1)
    Dim Lo_Tb_Ret_VRI       As ListObject:      Set Lo_Tb_Ret_VRI = Prog_Coef_Ret_VRI.ListObjects(1)
    Dim TRows_BD         As Long:        TRows_BD = Lo_BD.ListRows.Count
    Dim TRows_Ges04         As Long:        TRows_Ges04 = Lo_Ges04.ListRows.Count
    Prog_BD.Unprotect
    Prog_LsGes04.Unprotect
    Lo_BD.ShowTotals = False
    Lo_Ges04.ShowTotals = False
    
Call Rut_Off_Functions
    ' =============  Preparar Tabla de TitPH ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)    ' Ordenar por una Columna
    ' =============  Preparar Tabla de TitPH ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)
    Call Rut_Lo_Sort(Lo_Ges04, BD_Ref, xlAscending, True)    ' Ordenar por una Columna
            '- Visualizo el progreso
            TxT_Progreso = Form_Menu.TB_Informe & vbCrLf
            Form_Menu.TB_Informe = TxT_Progreso & "Incorporando LSGES04:  " & " 0 de " & Format(TRows_Ges04, "#,##0")
            Form_Menu.TB_Informe.SelStart = Len(Form_Menu.TB_Informe)
            Form_Menu.TB_Informe.SetFocus
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & vbLf
    '- ---------------------------------------------------------------------------------------------
    '- Actualizar BDatos con LsGes04 ---------------------------------------------------------------
    '        - Recorro toda LsGes04 para actualizar BDatos
    '        --- Referencias IGUALES                            <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios
    '        --- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO     <<<<  AÑADO UN NUEVO REGISTRO a BDatos
    '        --- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO   <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
    '- ---------------------------------------------------------------------------------------------
        
        If Not Lo_BD.DataBodyRange Is Nothing Then
            Lo_BD.ListColumns(BD_Incidencias).DataBodyRange.ClearContents    '- ClearContents -----
            Lo_BD.ListColumns(BD_EP_GestReg).DataBodyRange.ClearContents          '- ClearContents -----
        End If
    
    Do While F_G4 <= TRows_Ges04
    
    If F_G4 = TRows_Ges04 Then
        Debug.Print "F_G4 = TRows_Ges04"
    End If
        Set RwG4 = Lo_Ges04.ListRows(F_G4)
            '- Identificar Referencias Duplicadas, la SALTO
            If RwG4.Range(BD_Ref) = Ref_Ant Then
                RwG4.Range(BD_EP_Ctrl) = RwG4.Range(BD_EP_Ctrl) & "_Duplicaty"
                Lo_Ges04.ListRows(F_G4 - 1).Range(BD_EP_Ctrl) = Lo_Ges04.ListRows(F_G4 - 1).Range(BD_EP_Ctrl) & "_Duplicaty"
                Cont_Repes = Cont_Repes + 1
                F_G4 = F_G4 + 1
                GoTo Siguiente_Reg
            End If
        Set RwDB = Lo_BD.ListRows(F_BD)
        
    '--- Referencias IGUALES <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios  <
    '--- Referencias IGUALES <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios  <
    '--- Referencias IGUALES <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios  <
    '--- Referencias IGUALES <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios  <
        If Val(RwDB.Range(BD_Ref)) = Val(RwG4.Range(BD_Ref)) Then   '- YA EXISTE, LO ACTUALIZO -----
            '- Compruebo posibles INCIDENCIAS ------------------------------------------------------
            If RwDB.Range(BD_ImpRec) <> RwG4.Range(BD_ImpRec) * 1 Then                '- Cambio en el Imp. Recibo
                Incidencia = "Chg:PH_ImpRec=[" & RwDB.Range(BD_ImpRec) & "]_#_"
                RwDB.Range(BD_Incidencias) = RwDB.Range(BD_Incidencias) & Incidencia
                RwDB.Range(BD_H_Incidencias) = RwDB.Range(BD_H_Incidencias) & Incidencia
                Chg_ImpRec = Chg_ImpRec + 1
                End If
            If RwDB.Range(BD_ImpCob) > 0 And RwDB.Range(BD_ImpCob) <> RwG4.Range(BD_ImpCob) * 1 Then     '- Cambio en el Imp. Cobrado
                Incidencia = "Chg:PH_ImpCob=[" & RwDB.Range(BD_ImpCob) & "]_#_"
                RwDB.Range(BD_Incidencias) = RwDB.Range(BD_Incidencias) & Incidencia
                RwDB.Range(BD_H_Incidencias) = RwDB.Range(BD_H_Incidencias) & Incidencia
                Chg_ImpCob = Chg_ImpCob + 1
                End If
            If RwDB.Range(BD_ImpAdm) > 0 And RwDB.Range(BD_ImpAdm) <> RwG4.Range(BD_ImpAdm) * 1 Then     '- Cambio en el Imp. Adm.
                Incidencia = "Chg:PH_ImpAdm=[" & RwDB.Range(BD_ImpAdm) & "]_#_"
                RwDB.Range(BD_Incidencias) = RwDB.Range(BD_Incidencias) & Incidencia
                RwDB.Range(BD_H_Incidencias) = RwDB.Range(BD_H_Incidencias) & Incidencia
                Chg_ImpAdm = Chg_ImpAdm + 1
                End If
            If RwG4.Range(BD_Anul) = "S" Then     '- Tasa Anulada ----------------
                RwDB.Range(BD_Obs_Conta) = "Mat.Anulada_"
                Cont_Mat_Anul = Cont_Mat_Anul + 1
                End If
            
            '- Actualizo Todos los datos Nuevos  de Ges04 a BDatos ---------------------------------
            Set RngG4 = RwG4.Range.Cells(1, 1).Resize(1, BD_InfRegulariz)
            Set RngPH = RwDB.Range.Cells(1, 1).Resize(1, BD_InfRegulariz)
            RngPH.Value = RngG4.Value  ' para cambiar parte de una fila en una sólo sentencia
            '- Actualizo Todos los datos Asignados  en Ges04 a BDatos ------------------------------
            Set RngG4 = RwG4.Range.Cells(1, BD_ACont_Vto).Resize(1, BD_Cta_Ing - BD_Concepto + 1)
            Set RngPH = RwDB.Range.Cells(1, BD_ACont_Vto).Resize(1, BD_Cta_Ing - BD_Concepto + 1)
            RngPH.Value = RngG4.Value  ' para cambiar parte de una fila en una sólo sentencia
            Set RngG4 = RwG4.Range.Cells(1, BD_Rec_Imp_Acad)
            Set RngPH = RwDB.Range.Cells(1, BD_Rec_Imp_Acad)
            RngPH.Value = RngG4.Value  ' para cambiar parte de una fila en una sólo sentencia
            Set RngG4 = RwG4.Range.Cells(1, BD_Rec_Imp_Dto)
            Set RngPH = RwDB.Range.Cells(1, BD_Rec_Imp_Dto)
            RngPH.Value = RngG4.Value  ' para cambiar parte de una fila en una sólo sentencia
            
            Set RngG4 = RwG4.Range.Cells(1, BD_Rec_Imp_Adm)
            Set RngPH = RwDB.Range.Cells(1, BD_Rec_Imp_Adm)
            RngPH.Value = RngG4.Value  ' para cambiar parte de una fila en una sólo sentencia

            '- Preparo Salto de registro ------------------------------------
            RwDB.Range(BD_EP_GestReg) = RwDB.Range(BD_EP_GestReg) & "- Actualizado " & F_Actualiz & " - "
            RwG4.Range(BD_EP_GestReg) = "Actualizado BD, " & F_Actualiz
            Cont_Modif = Cont_Modif + 1
            If F_BD < TRows_BD Then F_BD = F_BD + 1
            F_G4 = F_G4 + 1
            Ref_Ant = RwG4.Range(BD_Ref)
            
        '--- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO  <<<<  AÑADO UN NUEVO REGISTRO a BDatos  <<<
        '--- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO  <<<<  AÑADO UN NUEVO REGISTRO a BDatos  <<<
        '--- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO  <<<<  AÑADO UN NUEVO REGISTRO a BDatos  <<<
        '--- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO  <<<<  AÑADO UN NUEVO REGISTRO a BDatos  <<<
        ElseIf Val(RwDB.Range(BD_Ref)) > Val(RwG4.Range(BD_Ref)) Or F_BD >= TRows_BD Then        '- NO EXISTE, AÑADO REGISTRO
            '--- Añado Registro -------------------------------------------------------------------
            Set RwDB = Lo_BD.ListRows.Add
            '- Actualizo Todos los datos Nuevos y Añadidos en Ges04 a BDatos -----------------------
            Set RngG4 = RwG4.Range.Cells(1, 1).Resize(1, BD_EP_GestReg)
            Set RngPH = RwDB.Range.Cells(1, 1).Resize(1, BD_EP_GestReg)
            RngPH.Value = RngG4.Value  ' para cambiar parte de una fila en una sólo sentencia
            ' -----------------=============  Buscar Tipo Plan  ==================------------------
            rowfind = Application.Match(RwDB.Range(BD_Plan), Lo_Tb_Ret_VRI.DataBodyRange.Columns(1), 0)
            If Not IsError(rowfind) Then    ' Plan Encontrado ==>> Tendrá características ESPECIALES
                RwDB.Range(BD_Coef_VRI) = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(rowfind)
            Else                            ' NO ENCONTRADO   ==>> Pongo el coeficiente establecido 15% ó 20%
                'If IsNumeric(Left(RwDB.Range(BD_Plan), 1)) Then
                If Range("APP_EFP_o_CFC") = "EFP" Then
                    RwDB.Range(BD_Coef_VRI) = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(1)
                Else
                    RwDB.Range(BD_Coef_VRI) = Lo_Tb_Ret_VRI.ListColumns("Coef_VRI").DataBodyRange(2)
                End If
            End If
           
            '- Preparo Salto de registro ------------------------------------
            RwDB.Range(BD_EP_GestReg) = "- Nuevo " & F_Actualiz & " - "
            RwG4.Range(BD_EP_GestReg) = "- Nuevo en BD, " & F_Actualiz
'''            RwDB.Range(BD_ACont_Vto) = Format(RwDB.Range(BD_FVto), "yyyy")
            Cont_Nuevo = Cont_Nuevo + 1
            F_G4 = F_G4 + 1
            Ref_Ant = RwG4.Range(BD_Ref)
            
        '--- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
        '--- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
        '--- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
        '--- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
        Else
            If RwDB.Range(BD_JI_Emi_Acad) = "" Then
                If InStr(RwDB.Range(BD_EP_GestReg), "Deleted") = 0 Then
                    RwDB.Range(BD_EP_GestReg) = RwDB.Range(BD_EP_GestReg) & "- Deleted " & F_Actualiz & " - "
                    RwDB.Range(BD_Tipo_Rec) = "Deleted"
                    Cont_Deleted = Cont_Deleted + 1
                End If
            Else
                If InStr(RwDB.Range(BD_EP_GestReg), "Deleted-conJI") = 0 Then
                    RwDB.Range(BD_EP_GestReg) = "- Deleted-conJI " & F_Actualiz & " - "
                    RwDB.Range(BD_Tipo_Rec) = "DeletedConJI"
                    Cont_DeletedconJI = Cont_DeletedconJI + 1
                End If
            End If
            If F_BD < TRows_BD Then F_BD = F_BD + 1
        End If
                '- Visualizo el progreso -----------------------------------------------------------
                If F_G4 Mod 500 = 0 Then
                    Form_Menu.TB_Informe = TxT_Progreso & "Incorporando LSGES04:  " & Format(F_G4, "#,##0") & " de " & Format(TRows_Ges04, "#,##0") & " reg."
'                    Application.ScreenUpdating = True:     DoEvents:         Application.ScreenUpdating = False
                End If
Siguiente_Reg:
    Loop
    '- Si quedan recibos en BDatos, NO EXISTE EN LsGes04 y quiere decir que <<<<< ES UN REGISTRO ELIMINADO
    Do While F_BD <= TRows_BD
        Set RwDB = Lo_BD.ListRows(F_BD)
        If RwDB.Range(BD_JI_Emi_Acad) = "" Then
            If InStr(RwDB.Range(BD_EP_GestReg), "Deleted") = 0 Then
                RwDB.Range(BD_EP_GestReg) = RwDB.Range(BD_EP_GestReg) & "- Deleted " & F_Actualiz & " - "
                RwDB.Range(BD_Tipo_Rec) = "Deleted"
                Cont_Deleted = Cont_Deleted + 1
            End If
        Else
            If InStr(RwDB.Range(BD_EP_GestReg), "Deleted-conJI") = 0 Then
                RwDB.Range(BD_EP_GestReg) = "- Deleted-conJI " & F_Actualiz & " - "
                RwDB.Range(BD_Tipo_Rec) = "DeletedConJI"
                Cont_DeletedconJI = Cont_DeletedconJI + 1
            End If
        End If
        If F_BD <= TRows_BD Then F_BD = F_BD + 1
    Loop
    
    '-Filtra Recibos Actualizar de BDatos ----------------------------------------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    rowfind = Application.WorksheetFunction.CountIf(Lo_BD.DataBodyRange.Columns(BD_EP_GestReg), "")
    If rowfind > 0 Then
        MsgBox "¡¡¡ Recibos SIN Actualizar !!!" & vbLf & vbLf & "¡¡¡ " & rowfind & "reg SIN ACTUALIZAR !!!", vbOKOnly + vbExclamation
    End If
    
    '  & Right(String(8, "_") & Format(.ListRows.Count, "#,##0"), 8) &
    
    
    
    Form_Menu.TB_Informe = TxT_Progreso & _
                 Right(String(8, "_") & Format(Cont_Nuevo, "#,##0"), 8) & " Incorporado Nuevos Recibos de LSGES04." & vbLf & _
                 Right(String(8, "_") & Format(Cont_Modif, "#,##0"), 8) & " Actualizados Recibos de BDatos con LSGES04." & vbLf & _
                 Right(String(8, "_") & Format(Cont_Deleted, "#,##0"), 8) & " Recibos de BDatos Deleted." & vbLf & _
                 Right(String(8, "_") & Format(Cont_DeletedconJI, "#,##0"), 8) & " ¡Ojo! Recibos de BDatos Borrados con JI." & vbLf & _
        Right("__________" & Chg_ImpRec + Chg_ImpCob + Chg_ImpAdm, 8) & "  Reg. Actualizados con incidencias (Cambios destacables)." & vbCrLf & _
        Right("__________" & Chg_ImpRec, 8) & "  Reg. Cambio en Importe de Recibo." & vbCrLf & _
        Right("__________" & Chg_ImpCob, 8) & "  Reg. Cambio en Importe Cobrado." & vbCrLf & _
        Right("__________" & Chg_ImpAdm, 8) & "  Reg. Cambio en Importe Administrativo." & vbCrLf & vbCrLf & _
        "Hay activos ahora un Total de:  " & Format(Lo_BD.ListRows.Count, "#,##0") & "  Reg." & _
                 Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " ¡Ojo! Recibos de BDatos Sin Actualizar."
    TxT_Progreso = Form_Menu.TB_Informe & vbCrLf
    
    
    If Cont_DeletedconJI > 0 Then
        MsgBox "¡¡¡ Recibos Borrados con JI's !!!" & Cont_DeletedconJI & " Recibos de BDatos Borrados con JI !!!", _
                vbOKOnly + vbExclamation, "Proceso: Actualizar BDatos con LsGes04"
    End If
    
    '- Copiar Todos los Registros "Deleted" en Lo_Deleted y Borrarlos de BDatos --------------------
    Dim Lo_Deleted      As ListObject:      Set Lo_Deleted = Prog_BD_Deleted.ListObjects(1)
    Call Rut_Lo_Sort(Lo_BD, BD_EP_GestReg, xlAscending, True)
    Lo_BD.Range.AutoFilter Field:=BD_EP_GestReg, Criteria1:="=*Deleted *"
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_BD, Lo_Deleted, False, True)
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    
    '- Visualizo el progreso
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                                " Actualizada BDatos con: " & Format(Lo_BD.ListRows.Count, "#,##0") & "reg." & vbCrLf

Call Rut_Lo_Filtros_Quitar(Lo_BD)
Call Rut_WrkSheet_ReducirPeso(Prog_LsGes04.Name, True)
Lo_BD.ShowTotals = True
Lo_Deleted.ShowTotals = True
Lo_Ges04.ShowTotals = True
       
Debug.Print "<<< RuT_Actualizar_BDatos_con_LsGes04"
End Sub




