Attribute VB_Name = "Rut_Ws"
Option Explicit

Sub Rut_WrkSheet_Active_Protect()
    
    ActiveSheet.Protect allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
    
'    ActiveSheet.Protect AllowFiltering:=True, _        Permite Filtrar
'                        AllowSorting:=True, _          Permite Ordenar
'                        DrawingObjects:=True, _        Impide que se seleccionen las Shapes
'                        UserInterfaceOnly:=True        Permite que las Macros puedan modificar la Sheet

End Sub
            '###################################################################################################################################
                    Sub Rut_WrkSheet_LstObj_LiberarEspacio_ByHand()
                        Call Rut_WrkSheet_LstObj_LiberarEspacio(ActiveSheet)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Prog_DR)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Prog_DrWrk)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_IAdm_CAcadAnt)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_AdmP)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_Ant)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_Dupl)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_ErrDate)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_M013)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_PNB1)
'                        Call Rut_WrkSheet_LstObj_LiberarEspacio(Sht__BD_Ajust)
                         ActiveSheet.Range("a1").Select
                    End Sub
'###################################################################################################################################
Sub Rut_WrkSheet_LstObj_LiberarEspacio(WrkSht As Worksheet)   '--- Borra TODO a la Derecha y Abajo de .ListObjects(1) ----------------------
' ==================================================================================================================================
    With WrkSht
            .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
            .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
            If .FilterMode Then .ShowAllData            ' Deshacer Filtros
'            If .ListObjects(1).FilterMode Then .ListObjects(1).ShowAllData            ' Deshacer Filtros
'            Rut_Lo_Filtros_Quitar (WrkSht.ListObjects(1))
        Dim UltCelda As Range
        With .ListObjects(1).Range                      ' Borra la filas de abajo y columnas de la derecha del la Tabla .ListObjects(1)
            Set UltCelda = .Cells(.Rows.Count, .Columns.Count)   ' Última celda de la tabla (sin Select, sin depender de la hoja activa)
        End With
            WrkSht.Range(WrkSht.Cells(UltCelda.Row + 1, UltCelda.Column + 1), WrkSht.Cells(WrkSht.Rows.Count, 1)).EntireRow.Delete
            WrkSht.Range(WrkSht.Cells(UltCelda.Row + 1, UltCelda.Column + 1), WrkSht.Cells(1, WrkSht.Columns.Count)).EntireColumn.Delete
            WrkSht.UsedRange                            ' Para restablecer el rango de celdas en uso
    End With
'    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht)
'            .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
'            .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
'            If .FilterMode Then .ShowAllData            ' Deshacer Filtros
'        With .ListObjects(1).Range                      ' Borra la filas de abajo y columnas de la derecha del la Tabla .ListObjects(1)
'            Range(.Cells(.Rows.Count, .Columns.Count).Address).Select   ' Selecciona la última celda de la tabla
'        End With
'            ActiveCell.Offset(1, 1).Select
'            Sheets(WrkSht).Range(ActiveCell.Address & ":" & Cells(Rows.Count, 1).Address).EntireRow.Delete
'            Sheets(WrkSht).Range(ActiveCell.Address & ":" & Cells(1, Columns.Count).Address).EntireColumn.Delete
'            ActiveSheet.UsedRange                       ' Para restablecer el rango de celdas en uso
'    End With
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

            '###################################################################################################################################
                    Sub Rut_WrkSheet_Vaciar_ByHand()
                        Call Rut_WrkSheet_Vaciar(ActiveSheet)
                    End Sub
' ==================================================================================================================================
Sub Rut_WrkSheet_Vaciar(WrkSht As Worksheet)      '--- Borra Toda la Hoja incluso los objetos (Shapes)  -------------------------------
' ==================================================================================================================================
Debug.Print ">>> Rut_WrkSheet_Vaciar, WrkSht=" & WrkSht.Name

    With Application.Workbooks(ThisWorkbook.Name).Sheets(WrkSht.Name)
            Dim Visual_Status   As Variant:  Visual_Status = .Visible   '--- para dejar la hoja en el mismo estado de Visibilidad ---
            Dim Protect_Status   As Boolean:  Protect_Status = .ProtectContents   '--- para dejar la hoja en el mismo estado de protección ---
        .Visible = xlHidden
        .Unprotect
            .Columns.Delete     ' --- con esto se borran hasta los "Shapes"
            .UsedRange
        .Visible = Visual_Status
        If Protect_Status Then .Protect
        .Select
    End With
Debug.Print ">>> Rut_WrkSheet_Vaciar, WrkSht=" & WrkSht.Name, "Protect_Status = " & Protect_Status
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
    Sub RuT_AllSheets_Visible_OrNot()    ' Estable la visibilidad establecida de cada Sheet en la Tabla(2) de Prog_Hojas_Names
' ==================================================================================================================================
Dim Cont        As Integer
Dim NomHoja     As String
On Error Resume Next    ' por si ha desaparecido una hoja ----
With Prog_HojasName.ListObjects("Tb_Sheets_Config").DataBodyRange
    For Cont = 1 To .Rows.Count
        NomHoja = .Cells(Cont, 2)
        Sheets(NomHoja).Visible = Val(.Cells(Cont, 3))
    Next
End With
On Error GoTo 0
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<

            ' ======================================================================================================================
            Sub RuT_Sort_Sheets_ByHand()
                RuT_Sort_Sheets
            End Sub
' ==================================================================================================================================
Sub RuT_Sort_Sheets(Optional Sort_Ascending As Boolean = True)
' ==================================================================================================================================
Dim Pos_Outer  As Integer
Dim Pos_Inner  As Integer
Dim Cant_Sheets     As Integer:   Cant_Sheets = Sheets.Count

    For Pos_Outer = 1 To Cant_Sheets
        For Pos_Inner = 1 To Pos_Outer
            If Sort_Ascending Then
                If UCase(Sheets(Pos_Outer).Name) < UCase(Sheets(Pos_Inner).Name) Then
                        Sheets(Pos_Outer).Move Before:=Sheets(Pos_Inner)
                End If
            Else
                If UCase(Sheets(Pos_Outer).Name) > UCase(Sheets(Pos_Inner).Name) Then
                        Sheets(Pos_Outer).Move Before:=Sheets(Pos_Inner)
                End If
            End If
        Next Pos_Inner
    Next Pos_Outer
    
End Sub
' ==================================================================================================================================
Function Fnc_WrkSheet_Exist(SheetName As String) As Boolean
    Dim Ws As Worksheet
    On Error Resume Next
    Set Ws = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    Fnc_WrkSheet_Exist = Not Ws Is Nothing
End Function
' ==================================================================================================================================
Function Fnc_Range_Exist(RngName As String) As Boolean
    Dim Rng As Range
    On Error Resume Next
    Set Rng = Range(RngName)
    On Error GoTo 0
    Fnc_Range_Exist = Not Rng Is Nothing
End Function
' -------------------------------------------------------------------------------------------------------------------------------<<<

' ==================================================================================================================================
' ==================================================================================================================================
' =====================     Rut_Columnas_ShowHide_RowX1     ================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
Sub Rut_WrkSheet_Col_ShowHide_RowX1(Show As Boolean)    ' Sirve para cualquier Hoja y para la Tabla nº.1
Dim Cont_Col    As Integer
    If Not Show Then                    ' Mostrar Columnas
        Columns.EntireColumn.Hidden = False
    Else                                        ' Ocultar Columnas
        For Cont_Col = 1 To ActiveSheet.ListObjects(1).Range.Columns.Count
            If IsEmpty(Cells(1, Cont_Col)) And Not Columns(Cont_Col).Hidden Then Columns(Cont_Col).Hidden = True
        Next Cont_Col
    End If
End Sub     ' Rut_Columnas_ShowHide_RowX1     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================

' ==================================================================================================================================
' ==================================================================================================================================
' =====================     Rut_Columnas_ShowHide_RowX1  - ¡¡¡ CON BOTONES (Shapes) !!!   ==========================================
' ==================================================================================================================================
' ==================================================================================================================================
Sub Rut_Columnas_ShowHide_RowX1()   ' Sirve para cualquier Hoja y para la Tabla nº.1
Dim Cont_Col    As Integer
If ActiveSheet.ProtectContents = True Then ActiveSheet.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
With ActiveSheet.Shapes.Range(Array("Botón_Ver_Col_SN"))
    If .TextFrame.Characters.Text = "Show Col's" Then                    ' Mostrar Columnas
        .Fill.ForeColor.RGB = RGB(200, 200, 250)
        .TextFrame.Characters.Text = "Hide Col's"
        Columns.EntireColumn.Hidden = False
    Else                                        ' Ocultar Columnas
        .Fill.ForeColor.RGB = RGB(150, 200, 100)
        .TextFrame.Characters.Text = "Show Col's"
        For Cont_Col = 1 To ActiveSheet.ListObjects(1).Range.Columns.Count
            If IsEmpty(Cells(1, Cont_Col)) And Not Columns(Cont_Col).Hidden Then Columns(Cont_Col).Hidden = True
        Next Cont_Col
    End If
End With
End Sub     ' Rut_Columnas_ShowHide_RowX1     <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================================================



