Attribute VB_Name = "Módulo1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("Tb_Liquid_TitPropios[[Fecha Emisión]:[Fecha Cobro]]").Select
    ActiveSheet.Unprotect
    Selection.NumberFormat = "m/d/yyyy"
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("Tb_Liquid_TitPropios[[Fecha Emisión]:[Fecha Cobro]]").Select
    Selection.NumberFormat = "d-m-yyyy"
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("C6").Select
    With Selection.Font
        .Name = "Arial Narrow"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 6
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 6
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
