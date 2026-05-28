Attribute VB_Name = "Módulo5"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("C12").Select
    Selection.NumberFormat = "0.00"
    ActiveCell.Replace What:="'", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Cells.Find(What:="'", After:=ActiveCell, LookIn:=xlFormulas2, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    Range("C12").Select
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
End Sub
