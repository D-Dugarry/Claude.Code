VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sht__Inf_Inf_Rsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Option Explicit

Private Sub Worksheet_Activate()
Debug.Print "Worksheet_Activate - " & ActiveSheet.Name
    Call Rut_Filtrar_Tareas
        Call Rut_Context_Buttons_Hide   '- Oculta las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
    Call RefreshRibbon
End Sub

Private Sub Worksheet_Deactivate()
Debug.Print "Worksheet_Deactivate - " & ActiveSheet.Name
    Call Rut_Context_Buttons_Restore    '- Restaura las opciones genéricas del Context-Menú Right-ClicK para dejar sólo visible las opciones Custom
End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub Worksheet_Change(ByVal Target As Range) '>>>>>>>>>>>> SI CAMBIA EL CONTENIDO DE LA CELDA SE ACTIVA (aunque acabe teniendo el mismo dato)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


If Not Prog__APP.Range("SW_Events") Then Exit Sub

    ' ============ Cambiar Plan  ==========================================================================
    If Not Application.Intersect(Target, Range("TP_Cod_Plan")) Is Nothing Then
    Dim Criterio As String
        Application.ScreenUpdating = False
        ActiveSheet.Unprotect
        With ActiveSheet.ListObjects(1)
            .ShowAutoFilter = True
            .AutoFilter.ShowAllData
            If Range("TP_Cod_Plan") = "" Then
                .AutoFilter.ShowAllData
                Exit Sub
            End If
            Criterio = Range("TP_Cod_Plan")
            .Range.AutoFilter Field:=1, Criteria1:=Criterio
            If .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count < 2 Then
                MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "No hay Datos de este Plan:    """ & Criterio & """", _
                vbOKOnly, "Proceso: Filtrado de Planes de Títulos Propios" & vbCrLf & "Quito Filtros."
                .AutoFilter.ShowAllData
            End If
        End With
        Range("f2").Select
        Application.ScreenUpdating = True
        ActiveSheet.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
        Exit Sub
    End If  ' -----------------------------------------------------------------------------------------------

End Sub





