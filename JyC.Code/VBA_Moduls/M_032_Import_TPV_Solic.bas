Attribute VB_Name = "M_032_Import_TPV_Solic"
'- M_032_Import_TPV_Solic
' Generador de informe en: https://cvnet.cpd.ua.es/uaGenInf/Home/Consulta/38051
Option Explicit

Dim Finalizar_Proceso           As Boolean
Dim AñoCont                     As String           ' Para Controlar el cambio de años el en número de orden que genero

' ==================================================================================================================================
Sub Import_TPV_Solic()
' ----------------------------------------------------------------------------------------------------------------------------------
Rut_Off_Functions
    '   Seleccionar fichero     ---------------------------------------------------------------------------------------------------
    Dim Arch__Solic_TPV         As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = Application.Workbooks(ThisWorkbook.Name).Path & "\"
        .Title = "Seleccionar el Fichero Excel de la Solicitud de Liquidación de los pagos por TPV / Bizum del Usuario: "
        .ButtonName = "Aceptar"
        .AllowMultiSelect = False
        .Filters.Add "Sólo Ficheros Excel", "*.xls?", 1
        If .Show <> -1 Then
            MsgBx_Msg = "Cancelado"
            MsgBx_Title = "Proceso: Importar LSGES04_GE"
            Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
            Prog__APP.Range("APP_Task_Inf") = "Proceso Cancelado: " & vbCrLf & Now()
            GoTo Restablecer_Valores
        Else
            Arch__Solic_TPV = .SelectedItems(1)
        End If
    End With
    '   Borrar el contenido de la hoja Prog_TPV_Solic    -----------------------------------------------------------
    Prog_TPV_Solic.Visible = xlSheetVisible
    Call Rut_WrkSheet_Vaciar(Prog_TPV_Solic.Name)
    '   Copio el excel    ------------------------------------------------------------------------------------------
    Dim closedBook          As Workbook
    Set closedBook = Workbooks.Open(Arch__Solic_TPV)
    Application.DisplayAlerts = False
    '    closedBook.Sheets(1).Cells.Copy ThisWorkbook.Sheets(Prog_TPV_Solic.Name).Range("A1")
        closedBook.Sheets(1).UsedRange.Copy
        ThisWorkbook.Sheets(Prog_TPV_Solic.Name).Range("A1").PasteSpecial xlPasteAll
        ThisWorkbook.Sheets(Prog_TPV_Solic.Name).Range("A1").PasteSpecial xlPasteColumnWidths
        ThisWorkbook.Sheets(Prog_TPV_Solic.Name).UsedRange.Rows.AutoFit
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    closedBook.Close SaveChanges:=False
    Set closedBook = Nothing
    
    Prog_TPV_Solic.Select
    Prog_TPV_Solic.Visible = True
    Rows(1).RowHeight = 55
    Range("a1").Select
    
    If Range("c11") <> "Referencia TPV" Then
        MsgBx_Msg = "Algo pasa con el Excel, no tiene el formato correcto."
        MsgBx_Title = "Proceso: Importar Solicitud de Liquidación de TPV"
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK", , "Ask"): Form_MsgBox.Show '- ([Font-Size]16, [Red-Border]False, [Buttons]"Ok", [Default-Button]1, [Image]"Msg")
        Prog__APP.Range("APP_Task_Inf") = "Proceso Cancelado: " & vbCrLf & Now()
        GoTo Restablecer_Valores
    End If
    MsgBox "Excel importado: " & Arch__Solic_TPV
Restablecer_Valores:
Rut_On_Functions
    H_Liq_TPV.Select
    Prog_TPV_Solic.Visible = xlSheetVeryHidden
End Sub     '- Import_TPV_Solic
' ==================================================================================================================================

Sub kk()
Range(Cells(8900, 1), Cells(8900, 10)).Copy
Cells(8914, 1).Select
'Range(Cells(8914, 1), Cells(8914, 1)).Select
'Range("a8914").Select
'Selection.Paste
ActiveSheet.Paste
End Sub



