Attribute VB_Name = "Rz_Rut_Aux"
Option Explicit

'###################################################################################################################################
    ' Call Rut_LstObj_Buscar(ListObject, String, Columna , ByRef FilaFound As Long)
Sub Rut_LstObj_Buscar(ByRef Lo_Tb As ListObject, Dato As String, Columna As Integer, ByRef FilaFound As Long)

    Dim RowFind            As Variant                                                   '-- Si FilaFound=0 --> Dato NO Encontrado ------
        ' -----------------=============  Buscar Dato en la Columna  ==================--------------------------
        RowFind = Application.Match(Dato, Lo_Tb.DataBodyRange.Columns(Columna), 0)
        If Not IsError(RowFind) Then    ' DATO Encontrado ------------------------
            FilaFound = RowFind
        Else                            ' NO ENCONTRADO   ------------------------
            FilaFound = 0
        End If
    End Sub
    
'###################################################################################################################################
' Call Rut_LstObj_Buscar(ListObject, String, Columna , ByRef FilaFound As Long)
Sub Rut_LstObj_Buscar_NunLong(ByRef Lo_Tb As ListObject, Dato As Long, Columna As Integer, ByRef FilaFound As Long)

    Dim RowFind            As Variant                                                   '-- Si FilaFound=0 --> Dato NO Encontrado ------
        ' -----------------=============  Buscar Dato en la Columna  ==================--------------------------
'        RowFind = Application.Match(Dato, Lo_Tb.DataBodyRange.Columns(Columna), 0)
        RowFind = Application.Match(Dato, Lo_Tb.ListColumns(Columna).Range, 0)
        If Not IsError(RowFind) Then    ' DATO Encontrado ------------------------
            FilaFound = RowFind
        Else                            ' NO ENCONTRADO   ------------------------
            FilaFound = 0
        End If
    End Sub
