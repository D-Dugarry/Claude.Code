Attribute VB_Name = "Módulo1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("Tb_INSS21[ACont_Cob]").Select
    Selection.TextToColumns Destination:=Range("B8"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
End Sub

Sub gh()
    ActiveSheet.ListObjects(1).Range.AutoFilter Field:=1, Criteria1:="=2024"
    ActiveSheet.ListObjects(1).Range.AutoFilter Field:=2, Criteria1:="<>2025"
End Sub



'- Bloque comentado: Dim/Const sueltos tras End Sub (invalidos en VBA salvo comentarios; 
'- 'Solo se permiten comentarios despues de End Sub'). Son un borrador huerfano de variables
'- de M_600_Informe_EP_RSm.bas, no se usan en ningun Sub de este modulo.
'    Dim RegsEmis        As Long     ' regs Emitidos
'    Dim Imp_Emis        As Currency
'    Dim TimpEmis        As Currency
    
'    Dim RegsEmisAnt     As Long     ' regs Emitidos
'    Dim Imp_EmisAnt     As Currency
'    Dim TImpEmisAnt     As Currency
    
'    Dim RegsEmisPos     As Long     ' regs Emitidos
'    Dim Imp_EmisPos     As Currency
'    Dim TImpEmisPos     As Currency
    
'    Dim RegsEmisPosAdm     As Long     ' regs Emitidos
'    Dim Imp_EmisPosAdm     As Currency
'    Dim TimpEmisPosAdm     As Currency

'    Dim RegsEmisPosAcad    As Long     ' regs Emitidos
'    Dim Imp_EmisPosAcad    As Currency
'    Dim TimpEmisPosAcad    As Currency

'    Dim RegsCobr        As Long     ' regs Cobrados
'    Dim Imp_Cobr        As Currency
'    Dim TImpCobr        As Currency
    
'    Dim RegsCobrAnt     As Long     ' regs Cobrados
'    Dim Imp_CobrAnt     As Currency
'    Dim TImpCobrAnt     As Currency
    
'    Dim RegsCobrPos     As Long     ' regs Cobrados
'    Dim Imp_CobrPos     As Currency
'    Dim TImpCobrPos     As Currency
    
'    Dim Regs_RDT        As Long     ' regs Cobrados
'    Dim Imp__RDT        As Currency
'    Dim TImp_RDT        As Currency
    
'    Dim RegsSRDT        As Long     ' regs Pendiente de pago
'    Dim Imp_SRDT        As Currency
'    Dim TImpSRDT        As Currency
    
'    Dim Regs_ADx        As Long     ' regs Pendiente de pago
'    Dim Imp__ADx        As Currency
'    Dim TImp_ADx        As Currency

'    Dim Imp__ADxAdm     As Currency
'    Dim TImp_ADxAdm     As Currency
'    Dim Imp__ADxAcad    As Currency
'    Dim TImp_ADxAcad    As Currency

'    Dim RegsAplz        As Long     ' regs Pendiente de pago
'    Dim Imp_Aplz        As Currency
'    Dim TImpAplz        As Currency

'    Dim RegsPdtCob        As Long     ' regs Pendiente de pago
'    Dim Imp_PdtCob        As Currency
'    Dim TimpPdtCob        As Currency

'    Dim R_EjeAnt        As Long     ' regs Pendiente de pago
'    Dim I_EjeAnt        As Currency
'    Dim T_EjeAnt        As Currency

'    Const Lst_Orden         As Integer = 1
'    Const Lst_Plan          As Integer = 2
'    Const Lst_RegsEmis      As Integer = 3
'    Const Lst_Imp_Emis      As Integer = 4
'    Const Lst_Imp_EmisAnt   As Integer = 5
'    Const Lst_Imp_ADx       As Integer = 6
'    Const Lst_Imp_ADxAdm    As Integer = 7
'    Const Lst_Imp_ADxAcad   As Integer = 8
'    Const Lst_Imp_Aplz      As Integer = 9
'    Const Lst_ImpEjeAnt     As Integer = 10
'    Const Lst_Imp_EmisPos   As Integer = 11
'    Const Lst_Imp_EmisPosAdm    As Integer = 12
'    Const Lst_Imp_EmisPosAcad   As Integer = 13
'    Const Lst_Imp_Cobr      As Integer = 14
'    Const Lst_Imp_CobrAdm   As Integer = 15
'    Const Lst_Imp_CobrAcad  As Integer = 16
'    Const Lst_Imp__RDT      As Integer = 17
'    Const Lst_Imp_SRDT      As Integer = 18
'    Const Lst_Imp_PdtCob    As Integer = 19

