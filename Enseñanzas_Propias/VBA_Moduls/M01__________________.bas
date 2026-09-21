Attribute VB_Name = "M01__________________"
' Last Rev. 2026-09-21 12:12
'2026-01-14
Option Explicit

'    RuT_Importar_LSGES04_GE()   '- Importar Última Consulta de LSGES04_GE, para Actualizar registros existentes y Añadir Nuevos en BDatos.
'
'    - Seleccionar fichero Excel LSGES04 e importar en ClsBook (RAM)
'        - Si no viene Tabla la creo
'        - M02_Del_Reg_EFP_o_CFCyAFC
'        - M02_Del_Reg_No_Válidos
'                - Borrar Recibos de otro Curso_Acad
'                - Borrar Recibos que BD_Matricula = "N"
'                - Borrar Recibos que BD_ActivEco <> 4 (Enseñanzas Propias)
'                - Borrar Recibos de Matrículas de coste CERO - ImpRec=ImpDto=0 - Recibos Matrícula de Actividad Académica a Coste CERO
'                - Borrar Recibos Importe CERO - Subvencionado- Imp_Rec =0 porque Imp_Dto >0
'        - Copy ClsBk:
'                - Vaciar Lo_LSGes04 y
'                - Copiar Lo_ClsBk en Lo_Ges04
'                - Cerrar RAM
'        - Proceso Lo_LSGes04:
'            - Formatear Lo_Ges04
'            - M02_Manage_Duplicates
'            - Asignar Año de Vencimiento Rec. en ACont_Vto
'            - Asignar Col Cta_Ingreso con nº Cta. correspondiente
'            - Asignar Código Concepto-Eco y Tipo_Ensañanza: 1310.00, 1311.03... EFP, CFC, TNCT, UPUA...
'            - Asignar Tipo de Recibo: Emitido, EjeAnt, Añejo, ADxAplz o Aplazado
'            -
'        - Proceso Lo_BDatos:
'            - M07_Actualiz_BDatos_con_LsGes04, Actualizar BDatos con Lo_Ges04
'            - Recorro toda LsGes04 para actualizar BDatos
'            --- Referencias IGUALES                            <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios
'            --- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO     <<<<  AÑADO UN NUEVO REGISTRO a BDatos
'            --- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO   <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
'            - Copiar Todos los Registros "Deleted" en Lo_Deleted y Borrarlos de BDatos
'
'        - M08_Actualizar_Tb_Coef_VRI
'            - Actualiza la Tabla de Referencia de los Coeficientes de Retención para el VRI

'    - Identificar Rec. Imp_Adm e Acad



'###################################################################################################
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
'###################################################################################################

'    LSGES04_GE_SinDtos_Curso_2024-25_Plan_AE4_(2026-01-06).xlsx
'    09/01/2026 21:48:27 Importado Excel: 18.311reg.
'
'    Borrados Rec. de estudios EFP_2024-25.....891reg., quedan: 17.420reg.
'
'    No hay Recibos de Curso-Acad ? 2024-25
'    No hay Recibos con Matrícula = N.
'    No hay Recibos con AE ? 4.
'    Borrados Rec. Matrícula de Actividad Acad. a Coste CERO ...8.118reg., quedan: 9.302reg.
'    Borrados Recibos Subvencionados 100% ImpDto ? ImpAcad ? ImpRec = 0 ...2.113reg., quedan: 7.189reg.
'
'    21:48:31 Copiado Excel a Lo_Ges04: 7.189reg.
'    21:48:32 Formateado Lo_Ges04.
'    21:48:32 Añadido ACont_Vto.
'
'    _____305 Cta_CCC Imp_Rec < 0
'    _____295 Cta_CCC No Cobrado
'    ______36 Cta_CCC FLY WIRE
'    ______33 Cta_CCC FLY Regularizado
'    ______41 Cta_CCC Regularizado G.Acad
'    _______9 Cta_CCC Regularizado G.Acad ???
'    _______6 Cta_CCC Regularizado S.Inf.
'    _______1 Cta_CCC Regularizado S.Inf.
'    _____227 Rec. SIN Cta-CCC de Ingreso Asignados.
'    ___6.347 Rec. CON Cta-CCC de Ingreso Asignados por el sistema.
'    ___7.189 Rec. en BDatos21:48:36 Añadidas Cta. de ingreso.
'
'    ___4.881 '1311.03'    Reg. CFC: Cursos de Formación Contínua
'    ___1.879 '1312.02'    Reg. CFC_UPUA (TUP): Programa Univ. para Mayores de la UA. (Univ. Permanente)
'    _____175 '1311.03'    Reg. AFC: Actividades de Formación Complementaria
'    ______92 '1311.03'    Reg. TNCT_M013: Seminario Orientación Pruebas > 25 años (Secretaría de Acceso)
'    _____162 '1303.01'    Reg. TNCT_PNB1: Prueba de competencias idioma extranjero. (Centro Sup. Idiomas)
'    _______0 'Not Found'  Reg. SIN Tipo TIO-EP o Concepto Económico.
'    ___7.189 Reg. en BDatos.
'    21:48:36 Añadido Concepto Económico y Tipo de Enseñanza.
'
'             Tipificado de Recibos. ________
'    ___6.808 Registros Emitidos.
'    _______2 Registros del Ejercicio Anterior. 2024
'    _______0 Registros Añejos, anteriores a 2024
'    ______74 Registros Aplazados, Anulados por Aplazamiento en 2024
'    _______0 Registros ADxAplz, Anulados por Aplazamiento en 2025, a cobrar en 2026
'    _____305 Registros de devolución.
'    _______0 Registros SIN Tipificar.7189reg.
'
'    ___7.189 Total Registros, Tipificados: 6884reg.
'
'    _____517 Registros Anulados.
'    _______6 Registros con errores de fechas.
'0            'Sin Identificar' Rec. SIN Determinar su Tipo (Si > 0, HAY QUE VERIFICAR FILTROS)
'             Son Reg. que se han quedado fuera de todos los filtros, o filtros que duplican tipo.
'       7.189 Reg. en BDatos.
'
'    _______1 Incorporado Nuevos Recibos de LSGES04.
'    ___7.163 Actualizados Recibos de BDatos con LSGES04.
'    __10.030 Recibos de BDatos Deleted.
'    ______12 ¡Ojo! Recibos de BDatos Borrados con JI.
'    _______0 ¡Ojo! Recibos de BDatos Sin Actualizar.
'    21:49:35 Actualizada BDatos con: 7.176reg.
'

