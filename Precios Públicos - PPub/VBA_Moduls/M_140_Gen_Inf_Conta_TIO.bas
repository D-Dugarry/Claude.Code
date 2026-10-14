Attribute VB_Name = "M_140_Gen_Inf_Conta_TIO"
' Last Rev. 2026-10-14 09:39
Option Explicit

'- Informe Conta (hoja Inf_Conta_TIO, 2026-10-14; Docs/Plan_Inf_Conta_TIO.md): el Informe de Recibos sin el INSS. No tiene las
'  columnas INSS Emi./Cob./Pdte., y en él Emitido, Cobrado y Pendiente son Adm. + Acad. Lo calcula el motor de M_130
'  (Rut_Inf_Calcular), con las mismas sumas en memoria y los mismos JI's de BDatos; aquí solo está lo propio de este informe.
'  Botón: el ábaco del grupo Calc_<hoja> (OnAct_ReCalculate_ListObj_Sht, M___RibbonUI), como en Inf_Recibos_TIO.

' ==================================================================================================================================
Sub Rut_Recalcular_Tabla_Inf_Conta()            '- botón del ábaco en Inf_Conta_TIO
' ==================================================================================================================================
    Dim Def         As T_Inf_Def
    Call Rut_Inf_Def_Conta(Def)
    Call Rut_Inf_Calcular(Def)                      '- M_130
End Sub     '- Rut_Recalcular_Tabla_Inf_Conta
' ----------------------------------------------------------------------------------------------------------------------------------
'- Descripción de Inf_Conta_TIO (T_Inf_Def, M_130) para los motores de M_130 (cálculo), M_133 (JI's a BDatos) y M_135 (correo).
'  Sus columnas son las de las constantes InfCon_* (M_000_Ini_Var_APP); las 3 del INSS no existen (0).
Public Sub Rut_Inf_Def_Conta(Def As T_Inf_Def)
    Set Def.Ws = Sht__Inf_Conta_TIO
    Def.Nom = "Informe Conta"
    Def.Titulo = "Informe Conta"
    Def.Con_INSS = False
    Def.Col(InfRec_TipRec) = InfCon_TipRec
    Def.Col(InfRec_Ensenanza) = InfCon_Ensenanza
    Def.Col(InfRec_ConcptNom) = InfCon_ConcptNom
    Def.Col(InfRec_Tot_Emi) = InfCon_Tot_Emi
    Def.Col(InfRec_Tot_Cob) = InfCon_Tot_Cob
    Def.Col(InfRec_Tot_Pdte) = InfCon_Tot_Pdte
    Def.Col(InfRec_Adm_INSS_Emi) = 0                         '- no la tiene
    Def.Col(InfRec_Adm_INSS_Cob) = 0                         '- no la tiene
    Def.Col(InfRec_Adm_INSS_Pdte) = 0                         '- no la tiene
    Def.Col(InfRec_Cta_Adm) = InfCon_Cta_Adm
    Def.Col(InfRec_Adm_Emi) = InfCon_Adm_Emi
    Def.Col(InfRec_Adm_Cob) = InfCon_Adm_Cob
    Def.Col(InfRec_Adm_Pdte) = InfCon_Adm_Pdte
    Def.Col(InfRec_ConcptEco) = InfCon_ConcptEco
    Def.Col(InfRec_Acad_Emi) = InfCon_Acad_Emi
    Def.Col(InfRec_Acad_Cob) = InfCon_Acad_Cob
    Def.Col(InfRec_Acad_Pdte) = InfCon_Acad_Pdte
    Def.Col(InfRec_JI_Emi_Adm) = InfCon_JI_Emi_Adm
    Def.Col(InfRec_JI_Emi_Acad) = InfCon_JI_Emi_Acad
    Def.Col(InfRec_AD_Emi_Adm) = InfCon_AD_Emi_Adm
    Def.Col(InfRec_AD_Emi_Acad) = InfCon_AD_Emi_Acad
    Def.Col(InfRec_JI_443_Adm) = InfCon_JI_443_Adm
    Def.Col(InfRec_JI_443_Acad) = InfCon_JI_443_Acad
    Def.Col(InfRec_Descrip_JI) = InfCon_Descrip_JI
    Def.Col(InfRec_Descrip_Contab) = InfCon_Descrip_Contab
    Def.Col(InfRec_ConcptEco2) = InfCon_ConcptEco2
    Def.Col(InfRec_ImpAcad_Crs_Ant) = InfCon_ImpAcad_Crs_Ant
    Def.Col(InfRec_ImpAcad_Crs_Pos) = InfCon_ImpAcad_Crs_Pos
    Def.Col(InfRec_ImpAcad_EmiAnt) = InfCon_ImpAcad_EmiAnt
    Def.Col(InfRec_ImpAcad_Emi) = InfCon_ImpAcad_Emi
    Def.Col(InfRec_ImpAcad_Cob_AcadAnt) = InfCon_ImpAcad_Cob_AcadAnt
    Def.Col(InfRec_ImpAcad_Cob_AcadPos) = InfCon_ImpAcad_Cob_AcadPos
    Def.Col(InfRec_ImpAcad_Pdte) = InfCon_ImpAcad_Pdte
    Def.Col(InfRec_ADxAplz) = InfCon_ADxAplz
    Def.Col(InfRec_Aplazado) = InfCon_Aplazado
    Def.Nom_Calc = "App_Last_Calc_Inf_Conta_TIO"
    Def.Nom_Mail = "App_Mail_Inf_Conta_TIO"
    Def.Nom_Ruta = "App_RutaRed_Inf_Conta_TIO"
    Def.Nom_Xlsx = "Inf_Conta_TIO"
    Def.Desglose = "desglosados en administrativos y académicos, <b>sin el seguro obligatorio (INSS)</b>"
End Sub     '- Rut_Inf_Def_Conta
' ----------------------------------------------------------------------------------------------------------------------------------
