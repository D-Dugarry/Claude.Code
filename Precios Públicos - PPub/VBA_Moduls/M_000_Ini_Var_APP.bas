Attribute VB_Name = "M_000_Ini_Var_APP"
'- M_00_Ini_Var_APP -----------------------------------------------------------------------------------------------------
Option Explicit

' Lo_Sht__BD  Base de Datos de Recibos de Precios Públicos  ----------------------
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
Public Const BD_TIO_EP             As Integer = 37   ' col: ak
Public Const BD_Cta_Ing            As Integer = 38   ' col: al
Public Const BD_Coef_VRI           As Integer = 39   ' col: am
Public Const BD_Orgánica           As Integer = 40   ' col: an
Public Const BD_ExpAdm             As Integer = 41   ' col: ao
Public Const BD_Liquidado          As Integer = 42   ' col: ap
Public Const BD_RDT                As Integer = 43   ' col: aq

Public Const BD_JI_Emi_Adm            As Integer = 44   ' col: ar
Public Const BD_JI_Emi_Acad           As Integer = 45   ' col: as
Public Const BD_AD_Emi_Adm            As Integer = 46   ' col: at
Public Const BD_AD_Emi_Acad           As Integer = 47   ' col: au
Public Const BD_JI_443_Adm            As Integer = 48   ' col: av
Public Const BD_JI_443_Acad           As Integer = 49   ' col: aw

Public Const BD_Obs_Conta             As Integer = 50   ' col: ax
Public Const BD_Incidencias           As Integer = 51   ' col: ay
Public Const BD_H_Incidencias         As Integer = 52   ' col: bz
Public Const BD_EP_Ctrl               As Integer = 53   ' col: ba
Public Const BD_EP_GestReg            As Integer = 54   ' col: bb

' - Columnas usadas como 'scratch' de marcado temporal por RuT_Clasif_Recibos (M_114) --------------
Public Const BD_CriT_Emi               As Integer = 52   ' col: bz  (mismo indice fisico que BD_H_Incidencias, reusado como scratch)
Public Const BD_CriT_EjeAnt            As Integer = 53   ' col: ba  (mismo indice fisico que BD_EP_Ctrl, reusado como scratch)
Public Const BD_CriT_Añejo             As Integer = 54   ' col: bb  (mismo indice fisico que BD_EP_GestReg, reusado como scratch)
Public Const BD_CriT_Aplazado          As Integer = 55   ' col: bc
Public Const BD_CriT_ADxAplz           As Integer = 56   ' col: bd
Public Const BD_CriT_ContabAnt         As Integer = 57   ' col: be
Public Const BD_CriT_DevEP             As Integer = 58   ' col: bf

' Lo_Prog_AE4.ListObjects(1) -----------------------------------
Public Const AE4_TipRec               As Integer = 1    ' col: a
Public Const AE4_Plan                 As Integer = 2    ' col: b
Public Const AE4_PlanNom              As Integer = 3    ' col: c
Public Const AE4_C_Acad               As Integer = 4    ' col: d
Public Const AE4_ConcptEco            As Integer = 5    ' col: e
Public Const AE4_Tot_Emi              As Integer = 6    ' col: f
Public Const AE4_Tot_Cob              As Integer = 7    ' col: g
Public Const AE4_Tot_Pdte             As Integer = 8    ' col: h
Public Const AE4_Adm_Emi              As Integer = 9    ' col: i
Public Const AE4_Adm_Cob              As Integer = 10   ' col: j
Public Const AE4_Adm_Pdte             As Integer = 11   ' col: k
Public Const AE4_Acad_Emi             As Integer = 12   ' col: l
Public Const AE4_Acad_Cob             As Integer = 13   ' col: m
Public Const AE4_Acad_Pdte            As Integer = 14   ' col: n
Public Const AE4_JI_Ant               As Integer = 15   ' col: o
Public Const AE4_AD_Ant               As Integer = 16   ' col: p
Public Const AE4_JI_Actual            As Integer = 17   ' col: q
Public Const AE4_AD_Actual            As Integer = 18   ' col: r
Public Const AE4_Descrip_JI           As Integer = 19   ' col: s

' Lo_LSace06.ListObjects(1) ------------------------------------
Public Const LS06_ACont_Emi           As Integer = 1    ' col: a
Public Const LS06_ACont_Cob           As Integer = 2    ' col: b
Public Const LS06_Ref                 As Integer = 3    ' col: c
Public Const LS06_C_Acad              As Integer = 4    ' col: d
Public Const LS06_FEmi                As Integer = 5    ' col: e
Public Const LS06_Expdte              As Integer = 6    ' col: f
Public Const LS06_DNI                 As Integer = 7    ' col: g
Public Const LS06_Nom                 As Integer = 8    ' col: h
Public Const LS06_Plan                As Integer = 9    ' col: i
Public Const LS06_Concept_Cod         As Integer = 10   ' col: j
Public Const LS06_Concept_Nom         As Integer = 11   ' col: k
Public Const LS06_Concept_Imp         As Integer = 12   ' col: l
Public Const LS06_Concept_Cant        As Integer = 13   ' col: m
Public Const LS06_Concept_TImp        As Integer = 14   ' col: n
Public Const LS06_Dto_Cod             As Integer = 15   ' col: o
Public Const LS06_Dto_Nom             As Integer = 16   ' col: p
Public Const LS06_Dto_Tot             As Integer = 17   ' col: q
Public Const LS06_F_Vto               As Integer = 18   ' col: r
Public Const LS06_ImpRec              As Integer = 19   ' col: s
Public Const LS06_ActivEco            As Integer = 20   ' col: t
Public Const LS06_NumRec              As Integer = 21   ' col: u
Public Const LS06_FormPag             As Integer = 22   ' col: v
Public Const LS06_CtaPag              As Integer = 23   ' col: w
Public Const LS06_FCob                As Integer = 24   ' col: x
Public Const LS06_ImpCob              As Integer = 25   ' col: y
Public Const LS06_InfRegulariz        As Integer = 26   ' col: z
Public Const LS06_Concept_TipoAcAd    As Integer = 27   ' col: aa
Public Const LS06_RecINSS             As Integer = 28   ' col: ab
Public Const LS06_RecFound            As Integer = 29   ' col: ac

' Tabla Prog_EPplazos.ListObjects(1) --------------------------------
Public Const IRs_Cod_Plan         As Integer = 1       ' col: a
Public Const IRs_Curso_Acad       As Integer = 2       ' col: b
Public Const IRs_Año_Emi          As Integer = 3       ' col: c
Public Const IRs_Plan_Curso       As Integer = 4       ' col: d
Public Const IRs_NomPlan          As Integer = 5       ' col: e
Public Const IRs_Orgánica         As Integer = 6       ' col: f
Public Const IRs_Ref_JI           As Integer = 7       ' col: g
Public Const IRs_Concepto         As Integer = 8       ' col: h
Public Const IRs_Incidencia       As Integer = 9       ' col: i
Public Const IRs_Cant_Reg         As Integer = 10       ' col: j
Public Const IRs_Ret_VRI          As Integer = 11       ' col: k

Public Const IRs_Emi_Tot          As Integer = 12       ' col: l
Public Const IRs_Emi_Org          As Integer = 13       ' col: m
Public Const IRs_Emi_VRI          As Integer = 14       ' col: n
Public Const IRs_Emi_Adm          As Integer = 15       ' col: o

Public Const IRs_Cob_Tot          As Integer = 16       ' col: p
Public Const IRs_Cob_Org          As Integer = 17       ' col: q
Public Const IRs_Cob_VRI          As Integer = 18       ' col: r
Public Const IRs_Cob_Adm          As Integer = 19       ' col: s

Public Const IRs_RDT_Tot          As Integer = 20       ' col: t
Public Const IRs_RDT_Org          As Integer = 21       ' col: u
Public Const IRs_RDT_VRI          As Integer = 22       ' col: v
Public Const IRs_RDT_Adm          As Integer = 23       ' col: w

Public Const IRs_RDT_Pte_Tot      As Integer = 24       ' col: x
Public Const IRs_RDT_Pte_Org      As Integer = 25       ' col: y
Public Const IRs_RDT_Pte_VRI      As Integer = 26       ' col: z
Public Const IRs_RDT_Pte_Adm      As Integer = 27       ' col: aa

Public Const IRs_AD_Tot           As Integer = 28       ' col: ab
Public Const IRs_AD_Org           As Integer = 29       ' col: ac
Public Const IRs_AD_VRI           As Integer = 30       ' col: ad
Public Const IRs_AD_Adm           As Integer = 31       ' col: ae

Public Const IRs_Descripc         As Integer = 32       ' col: af

' Tabla Prog_EPplazos.ListObjects(1) --------------------------------
Public Const Eplz_C_Acad              As Integer = 1    ' col: a
Public Const Eplz_Plan                As Integer = 2    ' col: b
Public Const Eplz_NomPlan             As Integer = 3    ' col: c
Public Const Eplz_Fech_P1             As Integer = 4    ' col: d
Public Const Eplz_Fech_P2             As Integer = 5    ' col: e
Public Const Eplz_Fech_P3             As Integer = 6    ' col: f
Public Const Eplz_Fech_P4             As Integer = 7    ' col: g
Public Const Eplz_Año_P1              As Integer = 8    ' col: h
Public Const Eplz_Año_P2              As Integer = 9    ' col: i
Public Const Eplz_Año_P3              As Integer = 10   ' col: j
Public Const Eplz_Año_P4              As Integer = 11   ' col: k
Public Const Eplz_Plazos              As Integer = 12   ' col: l
Public Const Eplz_PlazoAD             As Integer = 13   ' col: m

' Tabla Prog_JIs.ListObjects(1) --------------------------------
Public Const JIs_TipRec               As Integer = 1    ' col: a
Public Const JIs_Enseñanza            As Integer = 2    ' col: b
Public Const JIs_ConcptNom            As Integer = 3    ' col: c
Public Const JIs_Tot_Emi              As Integer = 4    ' col: d
Public Const JIs_Tot_Cob              As Integer = 5    ' col: e
Public Const JIs_Tot_Pdte             As Integer = 6    ' col: f
Public Const JIs_Cta_Adm              As Integer = 7    ' col: g
Public Const JIs_Adm_Emi              As Integer = 8    ' col: h
Public Const JIs_Adm_Cob              As Integer = 9    ' col: i
Public Const JIs_Adm_Pdte             As Integer = 10   ' col: j
Public Const JIs_ConcptEco            As Integer = 11   ' col: k
Public Const JIs_Acad_Emi             As Integer = 12   ' col: l
Public Const JIs_Acad_Cob             As Integer = 13   ' col: m
Public Const JIs_Acad_Pdte            As Integer = 14   ' col: n
Public Const JIs_JI_Ant               As Integer = 15   ' col: o
Public Const JIs_AD_Ant               As Integer = 16   ' col: p
Public Const JIs_JI_Actual            As Integer = 17   ' col: q
Public Const JIs_AD_Actual            As Integer = 18   ' col: r
Public Const JIs_Descrip_JI           As Integer = 19   ' col: s

' Tabla Prog_Inf_Recibos.ListObjects(1) --------------------------------
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
        '- Cols Inf Peter
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

' Tabla Prog__Menú_Aux.ListObjects(1) --------------------------------
Public Const DR_Ref                   As Integer = 1    ' col: a
Public Const DR_C_Acad                As Integer = 2    ' col: b
Public Const DR_Cod_Activ             As Integer = 3    ' col: c
Public Const DR_Imp_Rec               As Integer = 4    ' col: d
Public Const DR_F_Emi                 As Integer = 5    ' col: e
Public Const DR_F_Vto                 As Integer = 6    ' col: f
Public Const DR_PLAN                  As Integer = 7    ' col: g
Public Const DR_Exp                   As Integer = 8    ' col: h
Public Const DR_DNI                   As Integer = 9    ' col: i
Public Const DR_Cod_F_Pag             As Integer = 10   ' col: j
Public Const DR_Proforma              As Integer = 11   ' col: k
Public Const DR_N_Recibo              As Integer = 12   ' col: l
Public Const DR_H_Tip_Cob             As Integer = 13   ' col: m
Public Const DR_H_Cta                 As Integer = 14   ' col: n
Public Const DR_H_F_Emi               As Integer = 15   ' col: o
Public Const DR_H_F_Cob               As Integer = 16   ' col: p
Public Const DR_H_Imp_Cob             As Integer = 17   ' col: q
Public Const DR_H_Imp_Adm             As Integer = 18   ' col: r
Public Const DR_H_Cod_F_Pag           As Integer = 19   ' col: s
Public Const DR_H_Invalid             As Integer = 20   ' col: t
Public Const DR_H_AñoMes_Emi          As Integer = 21   ' col: u
Public Const DR_H_AñoMes_Rem          As Integer = 22   ' col: v
Public Const DR_H_AñoMes_Cob          As Integer = 23   ' col: w
Public Const DR_Ape_1                 As Integer = 24   ' col: x
Public Const DR_Ape_2                 As Integer = 25   ' col: y
Public Const DR_Nombre                As Integer = 26   ' col: z
Public Const DR_Tipo_Dto              As Integer = 27   ' col: aa
Public Const DR_Dto                   As Integer = 28   ' col: ab
Public Const DR_R_Regularizado        As Integer = 29   ' col: ac
Public Const DR_Tasa_Adm              As Integer = 30   ' col: ad
Public Const DR_Concepto              As Integer = 31   ' col: ae
Public Const DR_ACont_Vto             As Integer = 32   ' col: af
Public Const DR_JI_Emi                As Integer = 33   ' col: ag
Public Const DR_AD_0010               As Integer = 34   ' col: ah
Public Const DR_JI_443                As Integer = 35   ' col: ai
Public Const DR_AD_0001               As Integer = 36   ' col: aj
Public Const DR_Cta_Ing               As Integer = 37   ' col: ak
Public Const DR_Aux                   As Integer = 38   ' col: al
Public Const DR_Tipo_Tasa             As Integer = 39   ' col: am
Public Const DR_TIO_EP                As Integer = 40   ' col: an

'--- Tabla Lo_DefCol ---------------------------------------------------
Public Const DefC_Nom             As Integer = 1       ' col: a
Public Const DefC_TitColLstObj    As Integer = 2       ' col: b
Public Const DefC_TitColGenInf    As Integer = 3       ' col: c
Public Const DefC_TipVar          As Integer = 4       ' col: d
Public Const DefC_Format          As Integer = 5       ' col: e
Public Const DefC_Widht           As Integer = 6       ' col: f
Public Const DefC_CopyFormt       As Integer = 7       ' col: g
Public Const DefC_Align           As Integer = 8       ' col: h
Public Const DefC_Compare         As Integer = 9       ' col: i
Public Const DefC_WrapTxt         As Integer = 10      ' col: j
Public Const DefC_FormatCol       As Integer = 11      ' col: k
Public Const DefC_HiddenCol       As Integer = 12      ' col: l
Public Const DefC_ProtectData     As Integer = 13      ' col: m
Public Const DefC_Sort            As Integer = 14      ' col: n
Public Const DefC_LoTAdmAcad      As Integer = 15      ' col: o
Public Const DefC_RefreshBdAnt    As Integer = 16      ' col: p

' Tabla Prog__Menú_Aux.ListObjects(1) --------------------------------
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

'    Public Usuario_ID               As String
'    Public Usuario_Name             As String
'    Public Usuario_Ext              As String
'    Public App_MailUsu              As String
'    Public App_LetraUnidRed         As String
'    Public APP_letra_RutaAPP         As String
'    Public APP_SubName              As String   '- Nombre de la Rutina a ejecutar del MenúAux
'    Public APP_RutInform            As String   '- Informe de la Rutina ejecutada del MenúAux
    
'    Public SW_Boss                      As Boolean
    Public SW_Probando                  As Boolean
    Public SW_ShowHide_Col              As Boolean
'    Public SW_Events                    As Boolean
    Public SW_Cancelado                 As Boolean
    Public SW_Col_Ocultas               As Boolean
    Public SW_C_Acad_Ant                As Boolean
    Public SW_C_Acad_Pos                As Boolean
'    Public SW_Right_Click               As Boolean      '- Permite visualizar el Context-Menú Right-ClicK
    Public SW_Sheet_TitProp_LIQ         As Boolean
'    Public Sw_Cmb                       As Boolean

Public App_RutaAPP          As String   '- Ruta del WorkBook

'    Public Index_Tarea              As Integer
'    Public Index_RutAux              As Integer
    Public Index_Usuario            As Variant
    Public H_Inicio                 As Single
    Public LastTimeLap              As Single
    Public Concepto                 As Single
    Public Coef_VRI                 As Integer
    Public Index_Mail               As Integer  '--- Usado para seleccionar entre diff emails ---
    Public PlanAnt                  As String

    Public MsgBx_Title          As String   '- Variable para el Formulario del MsgBox
    Public MsgBx_TitleBar       As Boolean  '- Variable para el Formulario del MsgBox
    Public MsgBx_Msg            As String   '- Variable para el Formulario del MsgBox
    Public MsgBx_Answer         As Integer  '- Variable para el Formulario del MsgBox, Botón pulsado


    Public ActivForm        As Object                           '- Identificamos qué Formulario está Activo.  ----------
    Public Lo_Tareas        As ListObject
    
    Public Task_Inf         As String       '- Para el Informe de las actuaciones de las Rutinas.

