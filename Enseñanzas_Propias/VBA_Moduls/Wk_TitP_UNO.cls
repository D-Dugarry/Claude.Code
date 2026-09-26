VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wk_TitP_UNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Option Explicit
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Sub Worksheet_Change(ByVal Target As Range) '>>>>>>>>>>>> SI CAMBIA EL CONTENIDO DE LA CELDA SE ACTIVA (aunque acabe teniendo el mismo dato)
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'If Not Prog__APP_Switch.Range("Sw_EnableEvents") Then Exit Sub
'
'    ' ============ Cambiar Plan  ==================================================================
'    If Not Application.Intersect(Target, Range("TP_Cod_Plan")) Is Nothing Then
'
'
'        Rut_Resumen_Tab_TitPropios_UNO (Range("TP_Cod_Plan"))
'
''    Dim Criterio As String
''        Application.ScreenUpdating = False
''        ActiveSheet.Unprotect
''        With ActiveSheet.ListObjects(1)
''            .ShowAutoFilter = True
''            .AutoFilter.ShowAllData
''            If Range("TP_Cod_Plan") = "" Then
''                .AutoFilter.ShowAllData
''                Exit Sub
''            End If
''            Criterio = Range("TP_Cod_Plan")
''            If Range("TP_Curso_Acad") <> "" Then
''                Criterio = Criterio & "_" & Range("TP_Curso_Acad")
''                .Range.AutoFilter Field:=CtTP_Plan_Curss, Criteria1:=Criterio
''                If .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count < 2 Then
''                    MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "No hay Datos de este Plan:    """ & Criterio & """", _
''                    vbOKOnly, "Proceso: Filtrado de Planes de Títulos Propios" & vbCrLf & "Quito Filtros."
''                    .AutoFilter.ShowAllData
''                End If
''            Else
''                .Range.AutoFilter Field:=CtTP_Cod_Plan, Criteria1:=Criterio
''                If .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count < 2 Then
''                    MsgBox "¡¡¡ Ups, Sorry !!!" & vbCrLf & vbCrLf & "No hay Datos de este Plan:    """ & Criterio & """", _
''                    vbOKOnly, "Proceso: Filtrado de Planes de Títulos Propios" & vbCrLf & "Quito Filtros."
''                    .AutoFilter.ShowAllData
''                End If
''            End If
''        End With
''        Range("f2").Select
''        Application.ScreenUpdating = True
''        ActiveSheet.Protect
'        Exit Sub
'    End If  ' -------------------------------------------------------------------------------------
'End Sub
'
'
'

