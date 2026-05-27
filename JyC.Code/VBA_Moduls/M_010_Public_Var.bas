Attribute VB_Name = "M_010_Public_Var"

' Lo_ N43_Cta
Public Const C_N43_Ordinal        As Integer = 1       ' col: a
Public Const C_N43_Bco            As Integer = 2       ' col: b
Public Const C_N43_F_OPE          As Integer = 3       ' col: c
Public Const C_N43_F_VAL          As Integer = 4       ' col: d
Public Const C_N43_Imp            As Integer = 5       ' col: e
Public Const C_N43_Saldo          As Integer = 6       ' col: f
Public Const C_N43_Doc            As Integer = 7       ' col: g
Public Const C_N43_Ref1           As Integer = 8       ' col: h
Public Const C_N43_Reg_Mov1       As Integer = 9       ' col: i
Public Const C_N43_Reg_Mov2       As Integer = 10       ' col: j
Public Const C_N43_Reg_Mov3       As Integer = 11       ' col: k
Public Const C_N43_Reg_Mov4       As Integer = 12       ' col: l
Public Const C_N43_Reg_Mov5       As Integer = 13       ' col: m

' Lo_ JyC_Lst
Public Const C_JCL_Siglas        As Integer = 1       ' col: a
Public Const C_JCL_Tipo          As Integer = 2       ' col: b
Public Const C_JCL_Nom           As Integer = 3       ' col: c
Public Const C_JCL_F_Ini         As Integer = 4       ' col: d
Public Const C_JCL_F_Fin         As Integer = 5       ' col: e
Public Const C_JCL_Organiza      As Integer = 6       ' col: f
Public Const C_JCL_Orgánica      As Integer = 7       ' col: g
Public Const C_JCL_Email         As Integer = 8       ' col: h
Public Const C_JCL_Contact       As Integer = 9       ' col: i
Public Const C_JCL_Tlno          As Integer = 10      ' col: j
Public Const C_JCL_Obs           As Integer = 11      ' col: k
Public Const C_JCL_TpvRefs       As Integer = 12      ' col: l
Public Const C_JCL_ExpAdm        As Integer = 13      ' col: m

' Lo_TPV_Lst
Public Const C_TPL_EventoRef        As Integer = 1       ' col: a
Public Const C_TPL_EventoNom        As Integer = 2       ' col: b
Public Const C_TPL_F_Vto            As Integer = 3       ' col: c
Public Const C_TPL_ImpRec           As Integer = 4       ' col: d
Public Const C_TPL_ComBco           As Integer = 5       ' col: e
Public Const C_TPL_ImpNeto          As Integer = 6       ' col: f
Public Const C_TPL_Pdte_ComBco      As Integer = 7       ' col: g
Public Const C_TPL_Clasif           As Integer = 8       ' col: h
Public Const C_TPL_Siglas           As Integer = 9       ' col: i
Public Const C_TPL_Obs              As Integer = 10      ' col: j

'- Lo_ JyC INICI ----------------------------------------------------
Public Const C_INI_Siglas          As Integer = 1      ' col: a
Public Const C_INI_N_Liq           As Integer = 2      ' col: b
Public Const C_INI_CtaTPV          As Integer = 3      ' col: c
Public Const C_INI_Ref_Sol         As Integer = 4      ' col: d
Public Const C_INI_TPV_ImpRec      As Integer = 5      ' col: e
Public Const C_INI_TPV_ComBco      As Integer = 6      ' col: f
Public Const C_INI_TPV_Neto        As Integer = 7      ' col: g
Public Const C_INI_TPV_Dev         As Integer = 8      ' col: h
Public Const C_INI_Cta_ImpRec      As Integer = 9      ' col: i
Public Const C_INI_Cta_Imp         As Integer = 10     ' col: j
Public Const C_INI_Cta_Dev         As Integer = 11     ' col: k
Public Const C_INI_Cta_Transf      As Integer = 12     ' col: l
Public Const C_INI_Org             As Integer = 13     ' col: m
Public Const C_INI_JI              As Integer = 14     ' col: n
Public Const C_INI_ExpAdm          As Integer = 15     ' col: o
Public Const C_INI_N_RDT           As Integer = 16     ' col: p
Public Const C_INI_DI              As Integer = 17     ' col: q
Public Const C_INI_ImpDI           As Integer = 18     ' col: r
Public Const C_INI_PMP             As Integer = 19     ' col: s
Public Const C_INI_F_Fin           As Integer = 20     ' col: t
Public Const C_INI_PdteLiq         As Integer = 21     ' col: u
Public Const C_INI_Obs             As Integer = 22     ' col: v

' Lo_Cta
Public Const C_Cta_Ordinal       As Integer = 1       ' col: a
Public Const C_Cta_Bco           As Integer = 2       ' col: b
Public Const C_Cta_F_OPE         As Integer = 3       ' col: c
Public Const C_Cta_F_VAL         As Integer = 4       ' col: d
Public Const C_Cta_Imp           As Integer = 5       ' col: e
Public Const C_Cta_Saldo         As Integer = 6       ' col: f
Public Const C_Cta_Doc           As Integer = 7       ' col: g
Public Const C_Cta_Ref1          As Integer = 8       ' col: h
Public Const C_Cta_Reg_Mov1      As Integer = 9       ' col: i
Public Const C_Cta_Reg_Mov2      As Integer = 10       ' col: j
Public Const C_Cta_Reg_Mov3      As Integer = 11       ' col: k
Public Const C_Cta_Reg_Mov4      As Integer = 12       ' col: l
Public Const C_Cta_Reg_Mov5      As Integer = 13       ' col: m
Public Const C_Cta_Inscrito      As Integer = 14       ' col: n
Public Const C_Cta_N_Liq         As Integer = 15       ' col: o
Public Const C_Cta_Siglas        As Integer = 16       ' col: p
Public Const C_Cta_Ref_Sol       As Integer = 17       ' col: q
Public Const C_Cta_Obs           As Integer = 18       ' col: r
Public Const C_Cta_Ope           As Integer = 19       ' col: s
Public Const C_Cta_Org           As Integer = 20       ' col: t
Public Const C_Cta_JI            As Integer = 21       ' col: u
Public Const C_Cta_ExpAdm        As Integer = 22       ' col: v
Public Const C_Cta_RDT           As Integer = 23       ' col: w
Public Const C_Cta_DI            As Integer = 24       ' col: x
Public Const C_Cta_DI_Imp        As Integer = 25       ' col: y
Public Const C_Cta_PMP           As Integer = 26       ' col: z
Public Const C_Cta_N_OT          As Integer = 27       ' col: aa
Public Const C_Cta_Fusión        As Integer = 28       ' col: ab


' Lo_CTA_Liq
Public Const C_Cta_Liq_Ordinal    As Integer = 1    ' col: a
Public Const C_Cta_Liq_F_OPE      As Integer = 2    ' col: b
Public Const C_Cta_Liq_Imp        As Integer = 3    ' col: c
Public Const C_Cta_Liq_Reg_Mov1   As Integer = 4    ' col: d
Public Const C_Cta_Liq_Reg_Mov3   As Integer = 5    ' col: e
Public Const C_Cta_Liq_Reg_Mov4   As Integer = 6    ' col: f
Public Const C_Cta_Liq_Inscrito   As Integer = 7    ' col: g
Public Const C_Cta_Liq_Obs        As Integer = 8    ' col: h
Public Const C_Cta_Liq_JI         As Integer = 9    ' col: i
Public Const C_Cta_Liq_ExpAdm     As Integer = 10   ' col: j
Public Const C_Cta_Liq_N_RDT      As Integer = 11   ' col: k
Public Const C_Cta_Liq_DI         As Integer = 12   ' col: l
Public Const C_Cta_Liq_DI_Imp     As Integer = 13   ' col: m
Public Const C_Cta_Liq_PMP        As Integer = 14   ' col: n
Public Const C_Cta_Liq_N_OT       As Integer = 15   ' col: o
Public Const C_Cta_Liq_Siglas     As Integer = 16   ' col: p
Public Const C_Cta_Liq_Ref_Sol    As Integer = 17   ' col: q
Public Const C_Cta_Liq_N_Liq      As Integer = 18   ' col: r
Public Const C_Cta_Liq_Org        As Integer = 19   ' col: s
Public Const C_Cta_Liq_Ope        As Integer = 20   ' col: t

' Lo_TPV_Pagos del generador de informes --------------------------
Public Const C_Pag_EventoRef       As Integer = 1      ' col: a
Public Const C_Pag_EventoNom       As Integer = 2      ' col: b
Public Const C_Pag_F_Ini           As Integer = 3      ' col: c
Public Const C_Pag_F_Fin           As Integer = 4      ' col: d
Public Const C_Pag_F_Emi           As Integer = 5      ' col: e
Public Const C_Pag_F_Vto           As Integer = 6      ' col: f
Public Const C_Pag_F_Pago          As Integer = 7      ' col: g
Public Const C_Pag_Ref             As Integer = 8      ' col: h
Public Const C_Pag_DNI             As Integer = 9      ' col: i
Public Const C_Pag_Nom             As Integer = 10     ' col: j
Public Const C_Pag_Ape1            As Integer = 11     ' col: k
Public Const C_Pag_Ape2            As Integer = 12     ' col: l
Public Const C_Pag_NomApe          As Integer = 13     ' col: m
Public Const C_Pag_Tipo_Tasa       As Integer = 14     ' col: n
Public Const C_Pag_Tasa_Imp        As Integer = 15     ' col: o
Public Const C_Pag_Tipo_Dto        As Integer = 16     ' col: p
Public Const C_Pag_Porc_Dto        As Integer = 17     ' col: q
Public Const C_Pag_Tipo_Pago       As Integer = 18     ' col: r
Public Const C_Pag_Imp             As Integer = 19     ' col: s
Public Const C_Pag_ComBco          As Integer = 20     ' col: t
Public Const C_Pag_Neto            As Integer = 21     ' col: u
Public Const C_Pag_Clave_PPTO      As Integer = 22     ' col: v
Public Const C_Pag_CtaCCC          As Integer = 23     ' col: w

' Lo_TPV_Pagos (Request) -----------------------------------------------
Public Const C_Sol_F_Pago           As Integer = 1     ' col: a
Public Const C_Sol_NomApe           As Integer = 2     ' col: b
Public Const C_Sol_Ref              As Integer = 3     ' col: c
Public Const C_Sol_Inscrit          As Integer = 4     ' col: d
Public Const C_Sol_Imp              As Integer = 5     ' col: e
Public Const C_Sol_ComBco           As Integer = 6     ' col: f
Public Const C_Sol_Neto             As Integer = 7     ' col: g

' Lo_ TPV  and  TPV_Liqdatos
Public Const C_TPV_EventoRef       As Integer = 1      ' col: a
Public Const C_TPV_EventoNom       As Integer = 2      ' col: b
Public Const C_TPV_F_Ini           As Integer = 3      ' col: c
Public Const C_TPV_F_Fin           As Integer = 4      ' col: d
Public Const C_TPV_F_Emi           As Integer = 5      ' col: e
Public Const C_TPV_F_Vto           As Integer = 6      ' col: f
Public Const C_TPV_F_Pago          As Integer = 7      ' col: g
Public Const C_TPV_Ref             As Integer = 8      ' col: h
Public Const C_TPV_DNI             As Integer = 9      ' col: i
Public Const C_TPV_Nom             As Integer = 10     ' col: j
Public Const C_TPV_Ape1            As Integer = 11     ' col: k
Public Const C_TPV_Ape2            As Integer = 12     ' col: l
Public Const C_TPV_NomApe          As Integer = 13     ' col: m
Public Const C_TPV_Tipo_Tasa       As Integer = 14     ' col: n
Public Const C_TPV_Tasa_Imp        As Integer = 15     ' col: o
Public Const C_TPV_Tipo_Dto        As Integer = 16     ' col: p
Public Const C_TPV_Porc_Dto        As Integer = 17     ' col: q
Public Const C_TPV_Tipo_Pago       As Integer = 18     ' col: r
Public Const C_TPV_Imp             As Integer = 19     ' col: s
Public Const C_TPV_ComBco          As Integer = 20     ' col: t
Public Const C_TPV_Neto            As Integer = 21     ' col: u
Public Const C_TPV_Clave_PPTO      As Integer = 22     ' col: v
Public Const C_TPV_CtaCCC          As Integer = 23     ' col: w
Public Const C_TPV_Dias            As Integer = 24     ' col: x
Public Const C_TPV_N_Liq           As Integer = 25     ' col: y
Public Const C_TPV_Siglas          As Integer = 26     ' col: z
Public Const C_TPV_Ref_Sol         As Integer = 27     ' col: aa
Public Const C_TPV_Obs             As Integer = 28     ' col: ab
Public Const C_TPV_Inscrito        As Integer = 29     ' col: ac
Public Const C_TPV_Org             As Integer = 30     ' col: ad
Public Const C_TPV_JI              As Integer = 31     ' col: ae
Public Const C_TPV_ExpAdm          As Integer = 32     ' col: af
Public Const C_TPV_RDT             As Integer = 33     ' col: ag
Public Const C_TPV_RDT_inv         As Integer = 34     ' col: ah
Public Const C_TPV_DI              As Integer = 35     ' col: ai
Public Const C_TPV_DI_Imp          As Integer = 36     ' col: aj
Public Const C_TPV_PMP             As Integer = 37     ' col: ak
Public Const C_TPV_Ope             As Integer = 38     ' col: al
Public Const C_TPV_Ctrl            As Integer = 39     ' col: am

' Lo_TPV_Liq
Public Const C_TPV_Liq_Ref        As Integer = 1      ' col: a
Public Const C_TPV_Liq_F_Pag      As Integer = 2      ' col: b
Public Const C_TPV_Liq_Form_Pag   As Integer = 3      ' col: c
Public Const C_TPV_Liq_Ordenante  As Integer = 4      ' col: d
Public Const C_TPV_Liq_Inscrito   As Integer = 5      ' col: e
Public Const C_TPV_Liq_Imp        As Integer = 6      ' col: f
Public Const C_TPV_Liq_ComBco     As Integer = 7      ' col: g
Public Const C_TPV_Liq_Neto       As Integer = 8      ' col: h
Public Const C_TPV_Liq_Obs        As Integer = 9      ' col: i
Public Const C_TPV_Liq_JI         As Integer = 10     ' col: j
Public Const C_TPV_Liq_ExpAdm     As Integer = 11     ' col: k
Public Const C_TPV_Liq_N_RDT      As Integer = 12     ' col: l
Public Const C_TPV_Liq_N_RDT_inv  As Integer = 13     ' col: m
Public Const C_TPV_Liq_DI         As Integer = 14     ' col: n
Public Const C_TPV_Liq_DI_Imp     As Integer = 15     ' col: o
Public Const C_TPV_Liq_PMP        As Integer = 16     ' col: p
Public Const C_TPV_Liq_N_OT       As Integer = 17     ' col: q
Public Const C_TPV_Liq_Siglas     As Integer = 18     ' col: r
Public Const C_TPV_Liq_Ref_Sol    As Integer = 19     ' col: s
Public Const C_TPV_Liq_N_Liq      As Integer = 20     ' col: t
Public Const C_TPV_Liq_Org        As Integer = 21     ' col: u
Public Const C_TPV_Liq_Ope        As Integer = 22     ' col: v

' Lo_Prog__MnAux
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

'--- Tabla Lo_DefCol ---------------------------------------------------
Public Const DefC_Nom                As Integer = 1       ' col: a
Public Const DefC_TitColLstObj       As Integer = 2       ' col: b
Public Const DefC_TitColGenInf       As Integer = 3       ' col: c
Public Const DefC_TipVar             As Integer = 4       ' col: d
Public Const DefC_Format             As Integer = 5       ' col: e
Public Const DefC_Weight             As Integer = 6       ' col: f
Public Const DefC_CopyFormt          As Integer = 7       ' col: g
Public Const DefC_Align              As Integer = 8       ' col: h
Public Const DefC_Compare            As Integer = 9       ' col: i
Public Const DefC_WrapTxt            As Integer = 10      ' col: j
Public Const DefC_FormatCol          As Integer = 11      ' col: k
Public Const DefC_HiddenCol          As Integer = 12      ' col: l
Public Const DefC_LoTAdmAcad         As Integer = 13      ' col: m

Public BooL_Col_Ocultas            As Boolean
Public Mostrar_2023                As Boolean
Public Mostrar_2024                As Boolean
Public Mostrar_Pagos               As Boolean
Public Mostrar_Cobros              As Boolean
Public Liquidado_T_S               As Boolean
Public Liquidado_T_Bco             As Boolean
Public Liquidado_T_N               As Boolean
Public Ocultar_Superfluos          As Boolean
Public Hora_Inicio                 As Single             ' Para Saber el tiempo de proceso

Public Const Lcab                   As Integer = 5      ' Línea de Cabecera de las Línea Resultante con datos EN LA HOJA JyC
Public Const C_Lin_Inicio           As Integer = 12     ' fila de inicio de datos en la tabla
Public Ultima_Col                   As Integer
Public Ultima_Fila                  As Integer
Public Cont_Col                     As Long
Public Cont_Fila                    As Long

Public Index_Usuario        As Variant
Public TaskIndice         As Integer
Public APP_MnAux_Msg        As String   '- Informe de la Rutina ejecutada del MenúAux
Public App_RutaAPP          As String   '- Ruta del WorkBook
'Public APP_Rut_Inform        As String
'Public APP_User_Mail          As String   '- Emial del Usuario
'Public APP_User_UnidRed          As String   '- Letra de la Unidad de Nexe del Usuario
    
Public MsgBx_Title       As String
Public MsgBx_TitleBar    As Boolean
Public MsgBx_Msg         As String
Public MsgBx_Rut         As String
Public MsgBx_Answer      As Integer

Public Lo_Lst           As ListObject
Public Lo_Cta           As ListObject
Public Lo_TPV           As ListObject
Public Lo_Liq           As ListObject

Public H_Inicio         As Single
