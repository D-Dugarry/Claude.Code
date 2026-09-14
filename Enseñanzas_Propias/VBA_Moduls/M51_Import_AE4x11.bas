Attribute VB_Name = "M51_Import_AE4x11"
'2026-01-25
Option Explicit

'- ----------------------------------------------------------------------------------------------------------------------------
'- Seleccionar Excel pero no lo importa, sólo lo abre -------------------------------------------------------------------------
'- ----------------------------------------------------------------------------------------------------------------------------
Sub Rut_Lo_Import_AE4x1(Ws_AE4x1 As Worksheet, _
                                  Arch_New_Name As String, _
                                  SheetNom As String, _
                                  Optional NameFileAE4 As String, _
                                  Optional PathFileAE4 As String, _
                                  Optional SW_Inicilizar_Ws As Boolean = False)
                             
Debug.Print ">>> Rut_Lo_Import_LoData_LoDefCol_AE4x1"
    Dim Ccol            As Integer
    Dim RowsFind        As Variant
    Dim ArchRequest     As String:      ArchRequest = Arch_New_Name
    Dim AñoCont         As String:      AñoCont = Prog__APP.Range("APP_AñoCont")
    Dim C_Acad_Pos      As String:      C_Acad_Pos = Prog__APP.Range("APP_C_Acad_Pos")
    Dim C_Acad_Ant      As String:      C_Acad_Ant = Prog__APP.Range("APP_C_Acad_Ant")
    Dim FechCierreCont  As String:      FechCierreCont = Prog__APP.Range("APP_FechCierreCont")
    Dim Lo_AE4x1        As ListObject

    H_Inicio = Timer                '- Para saber el tiempo de proceso
        
    '- Select File -------------------------------------------------------------------------------------
'    Call Rut_File_Select_V2("Seleccionar el Fichero Excel " & Arch_New_Name & ": ", Arch_New_Name, "Excel", "*.xlsm")
    
    '- Seleccionar fichero Excel LSGES04 e importar en ClsBook (RAM) ------------------------------------------------------------------------
        Dim Arch__EP_New        As String
        Dim Nom_NewArch         As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = Application.Workbooks(ThisWorkbook.Name).Path & "\" & Arch_New_Name & "*"
        .Title = "Seleccionar el Fichero Excel de AE4. "
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        .ButtonName = "Seleccionar"
        .Filters.Clear
        .Filters.Add "Sólo Ficheros Excel", "*.xls?", 1
        If .Show <> -1 Then
            MsgBox "Cancelado", , "Rutinas"
            Form_Menu.TB_Informe = "Proceso Cancelado: " & Now()
            GoTo Cancel_Rut
        Else
            Arch_New_Name = .SelectedItems(1)
        End If
    End With
    If Arch_New_Name = "Cancel" Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ Cancelado a petición del Usuario !    "
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        Arch_New_Name = "Cancel"
        Exit Sub
    End If
    '- Comprueba que se ha seleccionado el nombre adecuado de Excel. --------------------------
    Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Arch_New_Name, NameFileAE4, PathFileAE4)
    If Left(NameFileAE4, Len(ArchRequest)) <> ArchRequest Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ Cancelado ¡" & vbLf & "El fichero debe ser:" & vbLf & ArchRequest & vbLf & vbLf & _
                    "  y se ha seleccionado:" & vbLf & NameFileAE4
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & vbLf & Now()
        Arch_New_Name = "Cancel"
        Exit Sub
    End If
    
    '- --------------------------------------------------------------------------------------------------------------
    Prog__APP.Range("SW_WB_Deactivate") = False     '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ---->>>
    Dim ClosedBook      As Workbook:        Set ClosedBook = Workbooks.Open(Arch_New_Name, ReadOnly:=True)
    Dim Ws_ClsBk        As Worksheet:       Set Ws_ClsBk = ClosedBook.Sheets(SheetNom)
    Dim Lo_ClsBk        As ListObject:      Set Lo_ClsBk = Ws_ClsBk.ListObjects(1)
    
    Ws_ClsBk.Unprotect
    Lo_ClsBk.ShowTotals = False
    
    '- Comprobar que la Tabla Lo_ClsBk No está vacía ---------------------
    If Lo_ClsBk.ListRows.Count = 0 Then
        MsgBx_Title = "Proceso: Importar " & ArchRequest
        MsgBx_Msg = "¡ La tabla no contiene datos !    "
        MsgBox MsgBx_Msg, vbExclamation, MsgBx_Title
        Prog__APP.Range("APP_Task_Inf") = Prog__APP.Range("APP_Task_Inf") & vbLf & vbLf & MsgBx_Msg & Now()
        GoTo Cancel_Rut
    End If

    '- Copy ClosedBook:
    Call Rut_Lo_Filtros_Quitar(Lo_ClsBk)
    Call Rut_WrkSheet_Preparar(Ws_AE4x1)
    '-          Si SW_Del_LoData=true Borrar Lo_AE4x1
    If SW_Inicilizar_Ws Then       '- Es el 1º, sólo copiar Lo_ClsBk en Ws_AE4x1
        Call Rut_WrkSheet_Vaciar(Ws_AE4x1.Name)
        Lo_ClsBk.Range.Copy Destination:=Ws_AE4x1.Range("A1")
        Set Lo_AE4x1 = Ws_AE4x1.ListObjects(1)
    Else
        Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_ClsBk, Lo_AE4x1, False)
    
    End If
    ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing
    Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<

    ' Restaurar entorno
Exit Sub

Cancel_Rut:
    Arch_New_Name = "Cancel"
    ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing
    Prog__APP.Range("SW_WB_Deactivate") = True      '- Esto parece que evita un ERROR al abrir el ClsBk que cierra el programa ----<<<
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------


