Attribute VB_Name = "Módulo_Filtro_Avanzado_Prueba"
' Last Rev. 2026-09-21 12:12
Option Explicit


'- -------------------------------------------------------------------------------------------------
'- Asignar Año de Vencimiento en ACont_Vto ---------------------------------------------------------
'- -------------------------------------------------------------------------------------------------
Sub RuT_filtro_Avanzado_Prueba()

Debug.Print ">>> RuT_Determinar_Año_Vto_Rec"
    Dim rowfind             As Variant
    Dim Lo_Data     As ListObject:  Set Lo_Data = Prog_LsGes04.ListObjects(1)
    Lo_Data.ShowTotals = False
        
    '- ---------------------------------------------------------------------------------------------
    '- Determinar Año Contable de Vencimiento del Recibo -------------------------------------------
    '- ---------------------------------------------------------------------------------------------
    With Lo_Data
        Call Rut_Lo_Filtros_Quitar(Lo_Data)
        
        .Range.AdvancedFilter xlFilterInPlace, Range("Tb_CriT_Reg_Err")
        rowfind = .Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1    '- OJO, TIENE QUE ESTAR VISIBLE LA COLUMNA Col_Ref
        Debug.Print rowfind
    
    End With    '-  Lo_Data
    Lo_Data.ShowAutoFilter = True
    Lo_Data.ShowTotals = True

End Sub

Sub kk()
    Dim Lo_Data     As ListObject:  Set Lo_Data = Prog_LsGes04.ListObjects(1)
        Call Rut_Lo_Filtros_Quitar(Lo_Data)

End Sub



