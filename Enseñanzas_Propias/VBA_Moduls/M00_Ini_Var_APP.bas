Attribute VB_Name = "M00_Ini_Var_APP"
' Last Rev. 2026-09-19 20:52
'- M00_Ini_Var_APP -----------------------------------------------------------------------------------------------------
Option Explicit

'--- Tabla Prog_BD  EFP y CFCyAFC -----------------------------------------------
Public Const BD_ACont_Emi          As Integer = 1    ' col: a
Public Const BD_ACont_Cob          As Integer = 2    ' col: b
Public Const BD_Plan               As Integer = 3    ' col: c
Public Const BD_NomPlan            As Integer = 4    ' col: d
Public Const BD_TipoCurso          As Integer = 5    ' col: e
Public Const BD_C_Acad             As Integer = 6    ' col: f
Public Const BD_Nom                As Integer = 7    ' col: g
Public Const BD_DNI                As Integer = 8    ' col: h
Public Const BD_Matricula          As Integer = 9    ' col: i
Public Const BD_Anul               As Integer = 10   ' col: j
Public Const BD_Ref                As Integer = 11   ' col: k
Public Const BD_NumRec             As Integer = 12   ' col: l
Public Const BD_ActivEco           As Integer = 13   ' col: m
Public Const BD_FEmi               As Integer = 14   ' col: n
Public Const BD_FVto               As Integer = 15   ' col: o
Public Const BD_FCob               As Integer = 16   ' col: p
Public Const BD_ImpRec             As Integer = 17   ' col: q
Public Const BD_ImpCob             As Integer = 18   ' col: r
Public Const BD_FormPag            As Integer = 19   ' col: s
Public Const BD_CtaPag             As Integer = 20   ' col: t
Public Const BD_RegMov             As Integer = 21   ' col: u
Public Const BD_Grupo              As Integer = 22   ' col: v
Public Const BD_ImpAcad            As Integer = 23   ' col: w
Public Const BD_ImpAdm             As Integer = 24   ' col: x
Public Const BD_Expdte             As Integer = 25   ' col: y
Public Const BD_HTipCob            As Integer = 26   ' col: z
Public Const BD_Hinvalid           As Integer = 27   ' col: aa
Public Const BD_ImpDto             As Integer = 28   ' col: ab
Public Const BD_InfRegulariz       As Integer = 29   ' col: ac
Public Const BD_Rec_Imp_Acad       As Integer = 30   ' col: ad
Public Const BD_Rec_Imp_Adm        As Integer = 31   ' col: ae
Public Const BD_Rec_Imp_Dto        As Integer = 32   ' col: af
Public Const BD_Rec_Imp_INSS       As Integer = 33   ' col: ag
Public Const BD_ACont_Vto          As Integer = 34   ' col: ah
Public Const BD_Concepto           As Integer = 35   ' col: ai
Public Const BD_Tipo_Rec           As Integer = 36   ' col: aj
Public Const BD_Tipo_EP            As Integer = 37   ' col: ak
Public Const BD_Cta_Ing            As Integer = 38   ' col: al
Public Const BD_Coef_VRI           As Integer = 39   ' col: am
Public Const BD_Orgánica           As Integer = 40   ' col: an
Public Const BD_ExpAdm             As Integer = 41   ' col: ao
Public Const BD_Liquidado          As Integer = 42   ' col: ap
Public Const BD_RDT                As Integer = 43   ' col: aq
Public Const BD_JI_Emi_Adm         As Integer = 44   ' col: ar
Public Const BD_JI_Emi_Acad        As Integer = 45   ' col: as
Public Const BD_AD_Emi_Adm         As Integer = 46   ' col: at
Public Const BD_AD_Emi_Acad        As Integer = 47   ' col: au
Public Const BD_JI_443_Adm         As Integer = 48   ' col: av
Public Const BD_JI_443_Acad        As Integer = 49   ' col: aw
Public Const BD_Obs_Conta          As Integer = 50   ' col: ax
Public Const BD_Incidencias        As Integer = 51   ' col: ay
Public Const BD_H_Incidencias      As Integer = 52   ' col: bz
Public Const BD_EP_Ctrl            As Integer = 53   ' col: ba
Public Const BD_EP_GestReg         As Integer = 54   ' col: bb

'--- Flags internos de Tipo_Recibo en Prog_LsGes04 (M05_Asign_Tipo_Recibo) ------
'    Columnas de marca por tipo; el valor definitivo va en BD_Tipo_Rec.
Public Const G04_Flag_Emitido      As Integer = 52   ' col: az
Public Const G04_Flag_EjeAnt       As Integer = 53   ' col: ba
Public Const G04_Flag_Anejo        As Integer = 54   ' col: bb
Public Const G04_Flag_Aplazado     As Integer = 55   ' col: bc
Public Const G04_Flag_ADxAplz      As Integer = 56   ' col: bd
Public Const G04_Flag_Primera      As Integer = G04_Flag_Emitido
Public Const G04_Flag_Cuantas      As Integer = G04_Flag_ADxAplz - G04_Flag_Emitido + 1

'--- Tabla de Liquidación Plán ---------------------------------------------------
Public Const CLiq_DNI             As Integer = 1       ' col: a
Public Const CLiq_Nombre          As Integer = 2       ' col: b
Public Const CLiq_NumLiquid       As Integer = 3       ' col: c
Public Const CLiq_Núm_Rec         As Integer = 4       ' col: d
Public Const CLiq_AñoVto          As Integer = 5       ' col: e
Public Const CLiq_F_Emi           As Integer = 6       ' col: f
Public Const CLiq_F_Cobro         As Integer = 7       ' col: g
Public Const CLiq_Imp_Rec         As Integer = 8       ' col: h
Public Const CLiq_Imp_Cob         As Integer = 9       ' col: i
Public Const CLiq_Imp_Adm         As Integer = 10       ' col: j
Public Const CLiq_Ajst_Tadm       As Integer = 11       ' col: k
Public Const CLiq_JI_Emi          As Integer = 12       ' col: l
Public Const CLiq_AD_0010         As Integer = 13       ' col: m
Public Const CLiq_JI_443          As Integer = 14       ' col: n
Public Const CLiq_ExpAdm          As Integer = 15       ' col: o
Public Const CLiq_RDT             As Integer = 16       ' col: p
Public Const CLiq_Coef_VRI        As Integer = 17       ' col: q
Public Const CLiq_Orgánica        As Integer = 18       ' col: r
Public Const CLiq_Ref             As Integer = 19       ' col: s
Public Const CLiq_Obs             As Integer = 20       ' col: t
Public Const CLiq_ImpAcad         As Integer = 21       ' col: u
Public Const CLiq_ImpDto          As Integer = 22       ' col: v
Public Const CLiq_Rec_Emi         As Integer = 23       ' col: w
Public Const CLiq_Pag_X_Alu       As Integer = 24       ' col: x
Public Const CLiq_Diff            As Integer = 25       ' col: y
Public Const CLiq_T_Adm_Neg       As Integer = 26       ' col: z

'--- Tabla DefCol ---------------------------------------------------
Public Const DefC_Nom             As Integer = 1     ' col: a
Public Const DefC_TitColLstObj    As Integer = 2     ' col: b
Public Const DefC_TitColGenInf    As Integer = 3     ' col: c
Public Const DefC_TipVar          As Integer = 4     ' col: d
Public Const DefC_Format          As Integer = 5     ' col: e
Public Const DefC_Widht           As Integer = 6     ' col: f
Public Const DefC_CopyFormt       As Integer = 7     ' col: g
Public Const DefC_Align           As Integer = 8     ' col: h
Public Const DefC_Compare         As Integer = 9     ' col: i
Public Const DefC_WrapTxt         As Integer = 10    ' col: j
Public Const DefC_FormatCol       As Integer = 11    ' col: k
Public Const DefC_HiddenCol       As Integer = 12    ' col: l

' Tabla Sht__Inf_Contab_Rec.ListObjects(1) --------------------------------
Public Const InfRec_TipRec            As Integer = 1    ' col: a
Public Const InfRec_Enseñanza         As Integer = 2    ' col: b
Public Const InfRec_ConcptNom         As Integer = 3    ' col: c

Public Const InfRec_Tot_Emi           As Integer = 4    ' col: d
Public Const InfRec_Tot_Cob           As Integer = 5    ' col: e
Public Const InfRec_Tot_Pdte          As Integer = 6    ' col: f

Public Const InfRec_Adm_INSS_Emi      As Integer = 7    ' col: g
Public Const InfRec_Adm_INSS_Cob      As Integer = 8    ' col: h
Public Const InfRec_Adm_INSS_Pdte     As Integer = 9    ' col: i

Public Const InfRec_Cta_Adm           As Integer = 10   ' col: j
Public Const InfRec_Adm_Emi           As Integer = 11   ' col: k
Public Const InfRec_Adm_Cob           As Integer = 12   ' col: l
Public Const InfRec_Adm_Pdte          As Integer = 13   ' col: m

Public Const InfRec_ConcptEco         As Integer = 14   ' col: n
Public Const InfRec_Acad_Emi          As Integer = 15   ' col: o
Public Const InfRec_Acad_Cob          As Integer = 16   ' col: p
Public Const InfRec_Acad_Pdte         As Integer = 17   ' col: q
        '- Cols JI's
Public Const InfRec_JI_Emi_Adm           As Integer = 18   ' col: r
Public Const InfRec_JI_Emi_Acad          As Integer = 19   ' col: s
Public Const InfRec_AD_Emi_Adm           As Integer = 20   ' col: t
Public Const InfRec_AD_Emi_Acad          As Integer = 21   ' col: u
Public Const InfRec_JI_443_Adm           As Integer = 22   ' col: v
Public Const InfRec_JI_443_Acad          As Integer = 23   ' col: w

Public Const InfRec_Descrip_JI           As Integer = 24   ' col: x
Public Const InfRec_Descrip_Contab       As Integer = 25   ' col: y
        '- Cols Peter
Public Const InfRec_ConcptEco2           As Integer = 26   ' col: z
Public Const InfRec_ImpAcad_Crs_Ant      As Integer = 27   ' col: aa
Public Const InfRec_ImpAcad_Crs_Pos      As Integer = 28   ' col: ab
Public Const InfRec_ImpAcad_EmiAnt       As Integer = 29   ' col: ac
Public Const InfRec_ImpAcad_Emi          As Integer = 30   ' col: ad
Public Const InfRec_ImpAcad_Cob_AcadAnt  As Integer = 31   ' col: ae
Public Const InfRec_ImpAcad_Cob_AcadPos  As Integer = 32   ' col: af
Public Const InfRec_ImpAcad_Pdte         As Integer = 33   ' col: ag
Public Const InfRec_ADxAplz              As Integer = 34   ' col: ah
Public Const InfRec_Aplazado             As Integer = 35   ' col: ai

' Tabla Prog_Coef_Ret_VRI.ListObjects(1) --------------------------------
Public Const CoefVRI_Plan          As Integer = 1
Public Const CoefVRI_CoefVRI       As Integer = 2
Public Const CoefVRI_Concepto      As Integer = 3
Public Const CoefVRI_TipCurs       As Integer = 4
Public Const CoefVRI_Orgánica      As Integer = 5
Public Const CoefVRI_Obs           As Integer = 6
Public Const CoefVRI_ORden         As Integer = 7

' Tabla Prog_DefCol_Plan_Rsm.ListObjects(1) --------------------------------
Public Const Rsm_Orden                As Integer = 1    ' col: a
Public Const Rsm_Plan                 As Integer = 2    ' col: b
Public Const Rsm_Coel_VRI             As Integer = 3    ' col: c
Public Const Rsm_RegsEmis             As Integer = 4    ' col: d
Public Const Rsm_Imp_Emis             As Integer = 5    ' col: e
Public Const Rsm_Imp_EmisAdm          As Integer = 6    ' col: f
Public Const Rsm_Imp_EmisAcad         As Integer = 7    ' col: g
Public Const Rsm_Imp_EmisAnt          As Integer = 8    ' col: h
Public Const Rsm_Imp_EmisAntAdm       As Integer = 9    ' col: i
Public Const Rsm_Imp_EmisAntAcad      As Integer = 10   ' col: j
Public Const Rsm_Imp_EmisPos          As Integer = 11   ' col: k
Public Const Rsm_Imp_EmisPosAdm       As Integer = 12   ' col: l
Public Const Rsm_Imp_EmisPosAcad      As Integer = 13   ' col: m
Public Const Rsm_Imp_Anul             As Integer = 14   ' col: n
Public Const Rsm_Imp_AnulAdm          As Integer = 15   ' col: o
Public Const Rsm_Imp_AnulAcad         As Integer = 16   ' col: p
Public Const Rsm_Imp_Cobr             As Integer = 17   ' col: q
Public Const Rsm_Imp_CobrAdm          As Integer = 18   ' col: r
Public Const Rsm_Imp_CobrAcad         As Integer = 19   ' col: s
Public Const Rsm_Imp_CobrAnt          As Integer = 20   ' col: t
Public Const Rsm_Imp_CobrAntAdm       As Integer = 21   ' col: u
Public Const Rsm_Imp_CobrAntAcad      As Integer = 22   ' col: v
Public Const Rsm_Imp_CobrPos          As Integer = 23   ' col: w
Public Const Rsm_Imp_CobrPosAdm       As Integer = 24   ' col: x
Public Const Rsm_Imp_CobrPosAcad      As Integer = 25   ' col: y
Public Const Rsm_Imp_RDT              As Integer = 26   ' col: z
Public Const Rsm_Imp_RDTAdm           As Integer = 27   ' col: aa
Public Const Rsm_Imp_RDTAcad          As Integer = 28   ' col: ab
Public Const Rsm_Imp_RDTpdt           As Integer = 29   ' col: ac
Public Const Rsm_Imp_RDTpdtAdm        As Integer = 30   ' col: ad
Public Const Rsm_Imp_RDTpdtAcad       As Integer = 31   ' col: ae
Public Const Rsm_Imp_ADx              As Integer = 32   ' col: af
Public Const Rsm_Imp_ADxAdm           As Integer = 33   ' col: ag
Public Const Rsm_Imp_ADxAcad          As Integer = 34   ' col: ah
Public Const Rsm_Imp_Aplz             As Integer = 35   ' col: ai
Public Const Rsm_Imp_AplzAdm          As Integer = 36   ' col: aj
Public Const Rsm_Imp_AplzAcad         As Integer = 37   ' col: ak
Public Const Rsm_Imp_EjeAnt           As Integer = 38   ' col: al
Public Const Rsm_Imp_EjeAntAdm        As Integer = 39   ' col: am
Public Const Rsm_Imp_EjeAntAcad       As Integer = 40   ' col: an
Public Const Rsm_Imp_PdtCob           As Integer = 41   ' col: ao
Public Const Rsm_Imp_PdtCobAdm        As Integer = 42   ' col: ap
Public Const Rsm_Imp_PdtCobAcad       As Integer = 43   ' col: aq
Public Const Rsm_Imp_PdtCobAnt        As Integer = 44   ' col: ar
Public Const Rsm_Imp_PdtCobAntAdm     As Integer = 45   ' col: as
Public Const Rsm_Imp_PdtCobAntAcad    As Integer = 46   ' col: at
Public Const Rsm_Imp_PdtCobPos        As Integer = 47   ' col: au
Public Const Rsm_Imp_PdtCobPosAdm     As Integer = 48   ' col: av
Public Const Rsm_Imp_PdtCobPosAcad    As Integer = 49   ' col: aw
Public Const Rsm_Imp_PdtCobADx        As Integer = 50   ' col: ax
Public Const Rsm_Imp_PdtCobADxAdm     As Integer = 51   ' col: ay
Public Const Rsm_Imp_PdtCobADxAcad    As Integer = 52   ' col: az
Public Const Rsm_Obs                  As Integer = 53   ' col: a{

' Tabla Prog__MnAux.ListObjects(1) --------------------------------
Public Const Task_Tarea              As Integer = 1       ' col: a
Public Const Task_Usuario            As Integer = 2       ' col: b
Public Const Task_Nombre_Rut         As Integer = 3       ' col: c
Public Const Task_Descripción        As Integer = 4       ' col: d
Public Const Task_Rut_Informe        As Integer = 5       ' col: e
Public Const Task_Imagen             As Integer = 6       ' col: f
Public Const Task_Uribbon_Tags       As Integer = 7       ' col: g
Public Const Task_Visible            As Integer = 8       ' col: h
Public Const Task_Emails             As Integer = 9       ' col: i
Public Const Task_SheetsButton       As Integer = 10      ' col: j


'    Public SW_Probando                  As Boolean
    Public SW_Col_Ocultas               As Boolean

    Public Index_Tarea                  As Integer
    Public Index_Usuario                As Variant
    Public H_Inicio                     As Double
    Public Concepto                     As Single
    Public Coef_VRI                     As Integer

    Public MsgBx_Title          As String   '- Variable para el Formulario del MsgBox
    Public MsgBx_TitleBar       As Boolean  '- Variable para el Formulario del MsgBox
    Public MsgBx_Msg            As String   '- Variable para el Formulario del MsgBox
    Public MsgBx_Answer         As Integer  '- Variable para el Formulario del MsgBox, Botón pulsado


    Public ActivForm           As Object      '- Identificamos que Formulario esta Activo. ----------
    Public Lo_Tareas           As ListObject
    
Public CantChanges   As Integer
Public CantSelectionChange   As Integer

