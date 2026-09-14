Attribute VB_Name = "M71_Restituir_BDatos_EP_Work"
'2026-02-16
'- M71_Restituir_BDatos_EP_Work -----------------------------------------------------------------------------------------------------

Option Explicit

'==================================================================================================================================
Sub Mod_Restituir_Tabla_Prog_BD()
'==================================================================================================================================
Rut_Off_Functions
    Dim Respuesta   As Integer
    Dim TipoEP      As String
    Dim FichNom     As String
    Dim Proceso     As String
'''    Dim RowFind     As Variant
    Dim Lo_BD       As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
    
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    Application.EnableEvents = False
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)

    Form_Menu.TB_Informe = "Proceso para Restituir BDatos de otro Excel AE4." & vbLf
    
    '- Seleccionar fichero     -------------------------------------------------------------------------------------------------------
    Dim Arch__EP_New        As String
    Dim Nom_NewArch         As String
    Dim Path_NewArch        As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = Application.Workbooks(ThisWorkbook.Name).Path & "\"
        .Title = "Seleccionar el Fichero Excel de las Tasas de Enseñanzas Propias a Restituir."
        .ButtonName = "Aceptar"
        .AllowMultiSelect = False
        .Filters.Add "Sólo Ficheros Excel", "*.xls?", 1
        If .Show <> -1 Then
            MsgBox "Cancelado", , "Rutinas"
            Form_Menu.TB_Informe = "Proceso Cancelado: " & Now() & vbLf
            GoTo Restablecer_Valores
        Else
            Arch__EP_New = .SelectedItems(1)
'            Nom_NewArch = Dir(Arch__EP_New) '- Falla con NEXE
            Nom_NewArch = Right(Arch__EP_New, Len(Arch__EP_New) - InStrRev(Arch__EP_New, "/"))
            Path_NewArch = Left(Arch__EP_New, InStrRev(Arch__EP_New, "\"))
        End If
    End With
    '- Importar en ClsBook (RAM)   --------------------------------------------------------------
    Dim ClsBk       As Workbook:        Set ClsBk = Workbooks.Open(Arch__EP_New, ReadOnly:=True)
    Dim Ws_ClsBk    As Worksheet:       Set Ws_ClsBk = ClsBk.Sheets("BDatos")
    Ws_ClsBk.Unprotect
    Ws_ClsBk.ListObjects(1).Name = "TbSource"   '¡¡ Tengo que cambiar el nombre pq sino coinciden entre los dos ficheros !!
    Dim Lo_ClsBk    As ListObject:      Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
    '''RowFind = Lo_ClsBk.ListRows.Count
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & Nom_NewArch & vbLf & Now & _
                            " Importado Excel AE4: " & Format(Lo_ClsBk.ListRows.Count, "#,##0") & " reg." & vbLf

    '- Importar parámetros -------------------------------------------
    Prog__APP.Range("APP_AñoCont") = ClsBk.Sheets(Prog__APP.Name).Range("APP_AñoCont")
    Prog__APP.Range("APP_CursAcad") = ClsBk.Sheets(Prog__APP.Name).Range("APP_CursAcad")
    Prog__APP.Range("APP_C_Acad_Ant") = ClsBk.Sheets(Prog__APP.Name).Range("APP_C_Acad_Ant")
    Prog__APP.Range("APP_C_Acad_Pos") = ClsBk.Sheets(Prog__APP.Name).Range("APP_C_Acad_Pos")
    Prog__APP.Range("APP_EFP_o_CFC") = ClsBk.Sheets(Prog__APP.Name).Range("APP_EFP_o_CFC")
    Prog__APP.Range("APP_CopSeg_HD_Date") = ClsBk.Sheets(Prog__APP.Name).Range("APP_CopSeg_HD_Date")
    Prog__APP.Range("APP_BorrarPLAN") = ClsBk.Sheets(Prog__APP.Name).Range("APP_BorrarPLAN")
    Prog__APP.Range("Sw_VerRecNeg") = ClsBk.Sheets(Prog__APP.Name).Range("Sw_VerRecNeg")
    Prog__APP.Range("Sw_VerRecNoCob") = ClsBk.Sheets(Prog__APP.Name).Range("Sw_VerRecNoCob")
    Prog__APP.Range("APP_PlanMicroCred") = ClsBk.Sheets(Prog__APP.Name).Range("APP_PlanMicroCred")
    Dim ColPlan     As String
    ColPlan = UCase(ClsBk.Sheets(Wk_TitP_Liquid.Name).Range("Liquid_Plan"))

    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & _
                String(8, "_") & " Tipo de EP:   " & Prog__APP.Range("APP_EFP_o_CFC") & vbLf & _
                String(8, "_") & " Año Contable: " & Prog__APP.Range("APP_AñoCont") & vbLf & _
                String(8, "_") & " Curso Acad.:  " & Prog__APP.Range("APP_CursAcad") & vbLf & _
                String(8, "_") & " Plan activo:  " & ColPlan & vbLf
        
    '- Actualizar parámetros en el Formulario de Menú -----------------
    Form_Menu.Tbx_AñoContable = Prog__APP.Range("APP_AñoCont")
    Form_Menu.OpBtn_CAcadAnt.Caption = Prog__APP.Range("APP_C_Acad_Ant")
    Form_Menu.OpBtn_CAcadPos.Caption = Prog__APP.Range("APP_C_Acad_Pos")
    
    If Range("APP_CursAcad") = Prog__APP.Range("APP_C_Acad_Ant") Then
        Form_Menu.OpBtn_CAcadAnt.Value = True
    Else
        Form_Menu.OpBtn_CAcadPos.Value = True
    End If

    If Range("APP_EFP_o_CFC") = "EFP" Then
        Form_Menu.OpBt_TP.Value = True
    Else
        Form_Menu.OpBt_CR.Value = True
    End If

    '''    '- Para ajustar combios de estructuras de la Tabla
    '''    '--- 1) Insertar 1 columnas VACÍAS a la derecha de "Col_46"
    '''    Lo_ClsBk.ListColumns.Add Position:=46
    '''    '--- 1) Insertar 1 columnas VACÍAS a la derecha de "Col_45"
    '''    Lo_ClsBk.ListColumns.Add Position:=45
    
    '- --------------------------------------------------------------------------------------------------------------
    '- Copy ClsBk:
    '-            1º vaciar BDatos
    '-            2º Copiar Lo_ClsBk en Lo_BD.
    '- --------------------------------------------------------------------------------------------------------------
    
    '- Borrar Lo_BD y Copiar Lo_ClsBk en Lo_BD. ---------------------
    Prog_BD.Visible = xlSheetVisible
    Prog_BD.Unprotect
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_BD, True)
    Set Ws_ClsBk = Nothing
    Set Lo_ClsBk = Nothing
    
    Dim Ws_Target   As Worksheet
    Dim Lo_Target As ListObject
    '- Borrar Lo_BDeleted de ClsBook y Copiar Lo_ClsBk_Deleted en Lo_BDeleted. ---------------------
    Set Ws_ClsBk = ClsBk.Sheets("BD_Deleted")                               '- ¡poner la WorkSheet que corresponda del CloseBook!
    Set Ws_Target = Prog_BD_Deleted     '- ¡poner la WorkSheet que corresponda del WorkSheet.CodeName!
    Set Lo_Target = Ws_Target.ListObjects(1)
    Ws_ClsBk.Unprotect
    Ws_ClsBk.ListObjects(1).Name = "TbSource1"   '¡¡ Tengo que cambiar el nombre pq sino coinciden entre los dos ficheros !!
    Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
    Ws_Target.Visible = xlSheetVisible
    Ws_Target.Unprotect
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
    Call Rut_Lo_Filtros_Quitar(Lo_Target)
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_Target, True, False)
    Set Ws_ClsBk = Nothing
    Set Lo_ClsBk = Nothing
    Set Lo_Target = Nothing
    
    '- Borrar Lo_BD_Duplic de ClsBook y Copiar Lo_ClsBk_BD_Duplic en Lo_BD_Duplic. ---------------------
    Set Ws_ClsBk = ClsBk.Sheets("BD_Duplic")                               '- ¡poner la WorkSheet que corresponda del CloseBook!
    Set Ws_Target = Prog_BD_Dupl       '- ¡poner la WorkSheet que corresponda del WorkSheet.CodeName!
    Set Lo_Target = Ws_Target.ListObjects(1)
    Ws_ClsBk.Unprotect
    Ws_ClsBk.ListObjects(1).Name = "TbSource2"   '¡¡ Tengo que cambiar el nombre pq sino coinciden entre los dos ficheros !!
    Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
    Ws_Target.Visible = xlSheetVisible
    Ws_Target.Unprotect
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
    Call Rut_Lo_Filtros_Quitar(Lo_Target)
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_Target, True, False)
    Set Ws_ClsBk = Nothing
    Set Lo_ClsBk = Nothing
    Set Lo_Target = Nothing
    
    '- Cerrar CloseBook
    ClsBk.Close SaveChanges:=False
    
    Range("APP_Tb_Restituida_Date") = "Tabla Restituída el:   " & Now()
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                            " Copiado Excel AE4 a Lo_BD: " & Format(Lo_BD.ListRows.Count, "#,##0") & " reg." & vbLf
    Proceso = Form_Menu.TB_Informe
    Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & vbCrLf & Format(Now, "hh:mm:ss") & _
                            " Actualizando Tabla Liquidación de Plan, espere por favor."
    Wk_TitP_Liquid.Select
    Application.EnableEvents = True
    Wk_TitP_Liquid.Range("Liquid_Plan") = ColPlan
    Application.EnableEvents = False
    Form_Menu.TB_Informe = Proceso & vbCrLf & Format(Now, "hh:mm:ss") & " Actualizada la Tabla de Liquidación del Plan activo:  " & ColPlan & vbLf
   
    Prog_BD.Select
    Prog_BD.Unprotect
    Call Rut_WrkSheet_ReducirPeso(Prog_BD.Name)
    '- Preguntamos si Borramos Recibos No Válidos
    Respuesta = MsgBox("¿ Borramos Posible Recibos NO válidos ?" & vbCrLf & vbCrLf & _
                       "Necesario para los Excels antiguos que no los eliminaba" & vbCrLf & vbCrLf & _
                       "- C_Acad NO válido" & vbCrLf & _
                       "- AE<>4" & vbCrLf & _
                       "- Rec. Matrícula coste CERO" & vbLf & _
                       "Recibos Importe CERO (Subvencionado)" & vbLf & _
                       "Etc.", vbExclamation + vbYesNo, "Proceso: Importar BDatos de EP (AE4)")

    If Respuesta = vbYes Then
                Call RuT_Del_Reg_NO_Válidos_Bdatos_ByHand
    Else
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                            " No se han Borrado registros." & vbLf
    End If
    '- ¡¡ Si sólo se restituye el excel para actualizar la programación !!
    If Range("APP_EFP_o_CFC") = "EFP" Then TipoEP = "EFP_" Else TipoEP = "CFCyAFC_"
    FichNom = TipoEP & Range("APP_CursAcad") & "_BaseDatos_Liq_V2.3.xlsm"
    Respuesta = MsgBox("¿ Creamos ya en Nuevo Excel ?" & vbCrLf & vbCrLf & FichNom & vbCrLf & vbCrLf & _
                       "Adecuado Si sólo se restituye el excel para actualizar la programación", _
                       vbExclamation + vbYesNo, "Proceso: Importar BDatos de EP (AE4)")
    If Respuesta = vbYes Then
        Call Rut_Crear_WB_EFP_o_CFCyAFC_de_CAcad_Ant_o_Pos
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                            " Creado el nuevo: " & FichNom & vbLf
    Else
        Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                            " No se ha Creado: " & FichNom & vbLf
    End If
    
Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & "¡¡¡ Proceso concluido con éxito !!! día: " & Now() & " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
        "Restituido el Fichero Excel de trabajo de las Tasas de " & TipoEP & vbCrLf & _
        "Se han importado: " & Format(Lo_BD.ListRows.Count, "#,##0") & " reg." & vbCrLf & _
        "Del Archivo: " & Nom_NewArch & vbLf & _
        "de: " & Path_NewArch & vbLf
        
Prog_BD.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'''    Application.EnableEvents = False
    Debug.Print "SW_EnableEvents = " & Range("SW_EnableEvents")
    Call Rut_EnableEvents_Status_Reset
    Wk_TitP_Liquid.Select

Rut_Off_Functions
End Sub     ' Mod_Restituir_Tabla_Prog_BD   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'===================================================================================================================================


'    '- Copiar Todos los Registros Deleted en Lo_Deleted y los BORRA de Lo_Data --------------------------------------------------------------
'    Dim Lo_Deleted      As ListObject:      Set Lo_Deleted = Prog_BD_Deleted.ListObjects(1)
'    Call Rut_Lo_Sort(Lo_BD, BD_EP_GestReg, xlAscending, True)
'    Lo_BD.Range.AutoFilter Field:=BD_EP_GestReg, Criteria1:="=*Deleted *"
'    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_BD, Lo_Deleted, False, True)
'    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    





