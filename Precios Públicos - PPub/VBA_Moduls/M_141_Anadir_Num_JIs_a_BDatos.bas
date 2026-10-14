Attribute VB_Name = "M_141_Anadir_Num_JIs_a_BDatos"
' Last Rev. 2026-10-14 09:39
Option Explicit

'- JI's del Informe Conta (hoja Inf_Conta_TIO) a BDatos (2026-10-14; Docs/Plan_Inf_Conta_TIO.md). Lo hace el motor de M_133
'  (Rut_Inf_JIs_a_BDatos), con las mismas reglas y en las mismas 6 columnas de BDatos (AR:AW) que el Informe de Recibos.
'  Botón «JI's del Inf. a BD» del grupo Calc_<hoja> (OnAct_JIs_Inf_a_BD, M___RibbonUI).

' ==================================================================================================================================
Sub Rut_Anadir_a_BDatos_JIs_de_Inf_Conta()      '- botón «JI's del Inf. a BD» en Inf_Conta_TIO
' ==================================================================================================================================
    Dim Def         As T_Inf_Def
    Call Rut_Inf_Def_Conta(Def)                     '- M_140
    Call Rut_Inf_JIs_a_BDatos(Def)                  '- M_133
End Sub     '- Rut_Anadir_a_BDatos_JIs_de_Inf_Conta
' ----------------------------------------------------------------------------------------------------------------------------------
