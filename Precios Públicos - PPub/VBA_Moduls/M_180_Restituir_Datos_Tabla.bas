Attribute VB_Name = "M_180_Restituir_Datos_Tabla"
'- M_180_Restituir_Datos_Tabla -----------------------------------------------------------------------------------------------------

Option Explicit

'==================================================================================================================================
Sub RuT_LstObj_Restore_BD()   ' Restituye la Tabla de otra versión del Excel --------
'==================================================================================================================================
Rut_Off_Functions
    H_Inicio = Timer                ' Para Saber el tiempo de proceso

'   Seleccionar fichero     -------------------------------------------------------------------------------------------------------
    Dim DatosFichDestino            As Variant
    Dim Arch__Tit_Prop_New         As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = Application.Workbooks(ThisWorkbook.Name).Path & "\" & "PPub_BDatos_Prog-RibbonX V-*"
        .Title = "Seleccionar el Fichero Excel de las Tasas de Títulos Propios y Cursos < 200h: "
        .ButtonName = "Aceptar"
        .AllowMultiSelect = False
        .Filters.Add "Sólo Ficheros Excel", "*.xls?", 1
        If .Show <> -1 Then
'            MsgBox "Cancelado", , "Rutinas"
            MsgBx_Msg = "Ha pulsado el botón <Cancelar>."
            MsgBx_Title = "Proceso: Restituir Tabla Tasas EP."
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", 1, "Ask"): Form_MsgBox.Show     '- ([Font-Size]=16, [Red-Border]=False, [Buttons]="Ok", [Default-Button]=1, [Image]="Msg")
            Prog__APP.Range("APP_Task_Inf") = "Proceso Cancelado: " & Now()
            GoTo Restablecer_Valores
        Else
            Arch__Tit_Prop_New = .SelectedItems(1)
        End If
    End With

    Sht__BD.Visible = xlSheetVisible
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Sht__BD.ListObjects(1).ShowTotals = False
    
'   Borrar el contenido de la Tabla     --------------------------------------------------------------------------------------------
    If Not Sht__BD.ListObjects(1).DataBodyRange Is Nothing Then Sht__BD.ListObjects(1).DataBodyRange.Delete
    
'   Copy Sheet from Another Workbook Without Opening it       ----------------------------------------------------------------------
    Dim ClosedBook          As Variant
    Set ClosedBook = Workbooks.Open(Arch__Tit_Prop_New)
    ClosedBook.Sheets(Sht__BD.Name).ListObjects(1).DataBodyRange.Copy
    Sht__BD.ListObjects(1).Range.Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    ClosedBook.Close SaveChanges:=False
    Set ClosedBook = Nothing
    Sht__BD.Select
    Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD)
    Range("d1").Select
    Range("d1") = "Tabla Restituída el:   " & Now()
    
Sht__BD.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'Sht__BD.Visible = xlSheetVeryHidden

Prog__APP.Range("APP_Task_Inf") = "¡¡¡ Proceso concluido con éxito !!! día: " & Now() & " - Tiempo: " & Round(Timer - H_Inicio, 2) & " seg." & vbCrLf & _
        "Restituido el Fichero Excel de trabajo de las Tasas de Títulos Propios y Cursos < 200h.  " & vbCrLf & _
        "Se han importado: " & Format(Sht__BD.ListObjects(1).ListRows.Count, "#,##0") & " reg." & vbCrLf & _
        "Del Archivo: " & Arch__Tit_Prop_New
Prog__APP.Range("APP_Last_BD_Restore") = Format(Now(), "dd-mmm-yy hh:mm")

Restablecer_Valores:    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Rut_Off_Functions
End Sub     ' RuT_Restituir_Tabla_Sht__BD   <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'===================================================================================================================================






