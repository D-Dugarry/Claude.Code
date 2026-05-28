VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Prog_TPV_Solic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


' ==================================================================================================================================
' =====================     RuT_UTF8_ANSI_Col3     ======================================================================================
' ==================================================================================================================================
Sub RuT_UTF8_ANSI_TPV_Solic()
Dim CaracterUTF8, Ini, Largo_Txt As Integer
Dim Cambios As Boolean
Const ColC As Integer = 4
Dim T_TPV As Object
Set T_TPV = ThisWorkbook.Sheets(TPV_Solic.Name).ListObjects("TPV_Solic").DataBodyRange
Dim T_UTF8 As Object
Set T_UTF8 = ThisWorkbook.Sheets(Prog_UTF8.Name).ListObjects("Tb_UTF8").DataBodyRange

'On Error Resume Next
    Hora_Inicio = Timer                ' Para Saber el tiempo de proceso

With T_TPV
    
    For Cont_Fila = 1 To .Rows.Count
        Do
            Cambios = False
            CaracterUTF8 = 1
            Do Until CaracterUTF8 > 96
                Ini = InStr(.Cells(Cont_Fila, ColC), T_UTF8.Cells(CaracterUTF8, 1))
                If Ini > 0 Then
                    .Cells(Cont_Fila, ColC) = Left(.Cells(Cont_Fila, ColC), Ini - 1) & T_UTF8.Cells(CaracterUTF8, 2) & _
                                            Mid(.Cells(Cont_Fila, ColC), Ini + T_UTF8.Cells(CaracterUTF8, 5))
                    Debug.Print "Fila " & Cont_Fila & " Pos: " & Ini; "__" & T_UTF8.Cells(CaracterUTF8, 1) & "__ " & _
                        .Cells(Cont_Fila, 2) & " - Nombre: " & .Cells(Cont_Fila, ColC)
                    Cambios = True
                    Exit Do
                End If
                CaracterUTF8 = CaracterUTF8 + 1
            Loop ' CaracterUTF8 <=96
        Loop While Cambios
        
    Next Cont_Fila
End With ' TPV
            Debug.Print "Fin"

    MsgBox "He tardado: " & Timer - Hora_Inicio & "  segundos"
End Sub     ' RuT_UTF8_ANSI_Col3
' ==================================================================================================================================



