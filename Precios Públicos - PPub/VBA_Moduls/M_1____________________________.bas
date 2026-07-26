Attribute VB_Name = "M_1____________________________"
'    Importar Última Consulta de LSGES04_GE, Actualizar registros existentes y Añadir Nuevos.
'
'    - Rut_Lo_Import_LoData_LoDefCol, Import LSGES04 del Año_Contable en Sht_BD
'    - Rut_Lo_ListColumns_ClearContents_DefC_ProtectData, Borrar por protección de Datos, Información sensible y no necesarias, según DefCol
'    - Rut_Lo_Format_LoData_LoDefColData, Format Sht__BD
'    - M_111, Filtrar y Borrar Registros NO deseados:
'        - Borrar Recibos AE4 Enseñanzas Propias
'        - Borrar Recibos de Matrículas de coste CERO
'        - Borrar Recibos con Imp.Rec. < 0
'        - Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"
'        - Borrar Recibos ANULADOS
'        - Borrar Recibos NO Martrícula
'        - Borrar Recibos INVALIDADOS
'        - Borrar Recibos con fechas FUERA DEL PERÍODO CONTABLE, ¡¡ o Borrar Datos del cobro !!
'        - Borrar Recibos Emitidos en Años Posteriores a AñoCont
'        - Borrar Incongruencias de Fechas
'    - M_112, Gestionar Duplicados
'    - M_113_Assign_Concept_Eco:
'        - Determinar Fecha de Vencimiento
'        - Determinar Cta-CCC Ingreso
'        - Determinar Concepto Económico
'        - Determinar Tipo de Enseñanza TIO-EP
'    - M_114, Clasificar Recibos en Emitidos, Aplazaados, EjeAnt, ADxAplz, Añejos


'    - M_115, Identificar y Asignar al primer registro de la matrícula el Importe Académico y el Administrativo
'    - M_215_Copy_IAdmCAcad_Ant_a_BD, Trasladar el ImpAdm, ImpAcad y ImpDto del C_Acad_Ant a BDatos

'    - M_120_Import_EFP_y_CFC, Importar Datos de Recibos de Enseñanzas Propias EP a BDatos
'
'    - Select File
'    - Con el ClsBk: (RAM)
'    - Compruebo ClsBk:
'    -            1º Si he dado una SheetNom si Existe
'    -            2º Si no Existe LisObject la creo
'    -            3º Comprobar que la Tabla Lo_ClsBk tiene datos
'    -            4º Comprobar que la cabecera de la Tabla Lo_ClsBk corresponde con la establecida en la LoDefCol
'    - Proceso:   1º Formateo
'    -            2º Borrar de Lo_ClsBk Recibos de LSave06 con C_Acad <> C_Acad_Ant y C_Acad_Pos
'    -            3º Borrar de Lo_ClsBk Recibos "<>INSS" en Nom_Concepto, para aligerar el peso de la Tabla.
'    -            4º M_314_Find_Rec_INSS, Identificar de un C_Acad, los 1º Rec. con Seguro obligatorio INSS, borrar los que no son.
'    -            5º Borro los datos de las columnas NO necesarias.
'    -            6º Determinar si en el LSace06 Hay UNO o DOS Cursos Académicos, Crea Dictionary para valores únicos (eficiente para grandes datos)
'    -            7º Borrar los Recibos de Lo_Data (Tabla DB_INSS) con el/los C_Acad del Nuevo LSace06
'                       Puesto que voy a copiar los nuevos registros, tengo que eliminar los viejos.
'                       Dependiendo de si en Lo_Source hay 1 ó 2 Cursos Académicos, filtro por 1 o 2 Cursos.
'    -
'    -



'    - Import BDatos de una versión anterior y Copiar en BDatos_Ant
'        - Actualizar BDatos con BDatos_Ant



'################################################################################################################
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
'################################################################################################################


'    Proceso: Importar LSGES04 para extraer Recibos Académicos de TIO y EP. Imp.Adm del Curso 2025-26.
'    Seleccionar el fichero y la ruta, para importar: LSGES04_GE_SinDtos_Año_2025
'    19:04:09 Lap:   5,04 seg. Excel Seleccionado: LSGES04_GE_SinDtos_Año_2025_BDatos (2026-01-22).xlsx
'Importado:     LSGES04_GE_SinDtos_Año_2025_BDatos(2026 - 1 - 22).xlsx
'    19:04:36 Lap:  25,03 seg. Importados nuevos datos: ................................ ........210.362 reg.
'    19:04:36 Lap:   0,01 seg. Formateando el Excel.
'    19:04:55 Lap:   0,53 seg. Formateadas: ................................17 de 58 col.
'    19:04:55 Lap:  18,68 seg. Formateado el Excel.
'
'    Borrados Rec. AE4 ..................................................... 20.797 reg. .... de 189.565 reg.
'    Borrados Rec. Negativos. ............................................... 7.144 reg. .... de 182.421 reg.
'    Del Rec. de Matrícula_Cero A Coste Cero................................. 1.741 reg. .... de 180.680 reg.
'    Del Rec. de Matrícula_Cero Subvencionada................................ 2.803 reg. .... de 177.877 reg.
'    Borrados Rec. Ficticio, DNI=1........................................... 3.541 reg. .... de 174.336 reg.
'    Borrados Rec. AE=300 que debería ser AE=4 y Plan=M013...................... 94 reg. .... de 174.242 reg.
'     - Son Rec. EFP, Plan=M013 'Seminario orientación pruebas > 65', mal matriculado x Secretaría de Acceso
'     - Mal matrículados con Rec.Mov. con AE=300, los borro porque vendrán en el AE4x4 ¡¡rectificados a mano!!
'    Borrados Rec. Anulados ................................................. 1.808 reg. .... de 172.434 reg.
'    No hay Rec. NO Matrícula
'    No hay Rec. Invalidados
'    Clear Data Rec. F_Cob > 31/12/2025..................................... 10.453 reg. .... de 172.434 reg.
'    No hay Rec. F_Emi > 31/12/2025
'    Marcados Rec. con Fechas Incongruentes. .................................... 3 reg. .... de 172.434 reg.
'    19:05:09 Lap:  14,34 seg. Recibos no requeridos para el procedimiento.
'
'    19:05:26 Lap:  17,05 seg. Referencias Repes-2 veces:.................... 5.008 reg. .... de 172.434 reg.
'    19:05:26 Lap:   0,05 seg. Referencias Repes-3 veces:...................... 383 reg. .... de 172.434 reg.
'    19:05:26 Lap:   0,05 seg. Referencias Repes-4 veces:....................... 20 reg. .... de 172.434 reg.
'    19:05:26 Lap:  17,21 seg. Find Ref. con Repeticiones: 5411 reg.  y Máx nº Repeticiones, : 4 veces).
'    19:05:27 Lap:   0,20 seg. Del en BDatos Rec. Repes NO finalistas........ 5.834 reg.  quedan 166.600 reg.
'    19:05:27 Lap:   0,74 seg. Copiados Repes finalistas a Bd_Duplic......... 5.411 reg. .. tiene 19.982 reg.
'    19:05:29 Lap:   1,59 seg. Del Repes-X de repetidos RpIdem............... 5.371 reg. . quedan 14.611 reg.
'    19:05:29 Lap:   0,14 seg. Del Repes_Cambios(en 0 Cols)..................... 40 reg. . quedan 14.571 reg.
'
'    No hay en BDatos Rec. AE4_2024-25 o AE4_2025-26...tiene 166.600 reg
'    Copy en BDatos, Rec. AE4_2024-25 y AE4_2025-26........................... 10.282 reg.....BD= 176.882 reg
'
'    Proceso: Asignación Fecha de Vto.
'     Año_Vto = 2025 ( F_Vto < 2024)....................................................0
'     Año_Vto = 2025 ( F_Vto = 2025)..............................................130.427
'     Año_Vto = 2024 ( F_Vto = 2024)...................................................72
'     Año_Vto = 2026 ( F_Vto = 2026)...............................................46.383
'     Errores de ACont_Vto < ACont_Emi, Cambiado Fecha: ACont_Vto = ACont_Emi...................8
'         Identificado Rec. AñoVto 2024................................................64
'         Identificado Rec. AñoVto 2025...........................................130.435
'         Identificado Rec. AñoVto 2026............................................46.383
'     Rec. SIN F_Vto. Asignada..........................................................0
'     Rec. en BDatos..............................................................176.882
'
'    Proceso: Asignación Concepto Económico y Tipo Ensañanza.
'    _134.623 '1310.00'    Reg. AE6 Grado
'    __13.259 '1310.01'    Reg. AE5 Master
'    ___3.783 '1310.02'    Reg. AE2 Doctorado
'    ___1.351 '1311.00'    Reg. AE4 EFP: Estudios de Formación Permanente: Master, Especialista y Experto.
'    ___4.561 '1311.03'    Reg. AE4 CFC: Cursos de Formación Contínua
'    _____429 '1311.03'    Reg. AE4 AFC: Actividades de Formación Complementaria
'    ______88 '1312.00'    Reg. AE300 TNCT_M013: Seminario Orientación Pruebas > 25 años
'                               (Secretaría de Acceso AE4/300/710)
'    ___3.723 '1312.02'    Reg. AE4 CFC_UPUA (TUP): Programa Univ. para Mayores UA. (Univ. Permanente)
'    ___5.000 '1315.00'    Reg. AE80 Pruebas Acceso Univ.
'    _____130 '1303.01'    Reg. AE4 TNCT_PNB1: Prueba de competencias idioma extranjero.
'                               (Centro Sup. Idiomas AE4/21)
'    ___9.405 '1303.00'    Reg. Rec_Adm: Recibos de una actividad púramente Administrativa.
'    _____530 'EURLElda'   Reg. AE6-Plan_C404 EURLElda: Esc. Univ. Relaciones Laborales Elda
'    _______0 'No Contab.' Reg. ImpMatCero: Matrícula de Actividad Académica a Coste CERO.
'    _______0 'No Contab.' Reg. Recibos SIN identificar Concepto-Eco o Tipo.
'    =176.882 Reg. en BDatos
'    19:05:58 Lap:  18,66 seg. Proceso Finalizado: Asignación Concepto Económico y Tipo Ensañanza.
'
'    Clasificación de Recibos, Estadística:
'    Clasificados Reg. Emididos: .............................................85.900 reg.
'    Clasificados Reg. EjeAnt: ...................................................65 reg.
'    Clasificados Reg. Añejos: ....................................................2 reg.
'    Clasificados Reg. Aplazados: ............................................44.707 reg.
'    Clasificados Reg. ADxAplz: ..............................................43.853 reg.
'    Clasificados Reg. de EP=AE4x4 _Devolución_EP_: ...........................1.889 reg.
'    Clasificados Reg. _Reg_Anul_: ..............................................456 reg.
'         Registros: Anulados, NO Matrícula o Invalidados:
'    Clasificados Reg. _Contab_Ant_: .............................................10 reg.
'         Registros, Cobrados anteriormente y por lo tanto, ya Contabilizados.
'
'    Sumatorio Recibos clasificados (de 176.882 reg.)........................176.882 reg......faltan = 0 reg.
'
'    Clasificados Reg. _ERR_Date_: ...............................................10 reg.
'
'    Total Recibos BDatos: ..................................................176.882 reg.
'
'    No se tendrán en cuenta los Rec. Anulados, con Imp.Adm. <0 ..................14 reg..........-18.521,86€
'
'    19:07:13 Lap:  64,11 seg. Resultado de Identificar Importes del AñoCont_2025, en: 176.882reg.
'    - Tasas Acad. de Matrícula por un importe de:    .....................28.975.834,53€......en 34.501 reg,
'    - Descuentos de Matrícula por un importe de:    ......................-8.520.272,96€......en 12.973 reg,
'    - Tasas Adm. de Matrícula por un importe de:    .........................856.570,75€......en 34.501 reg,
'
'    19:07:13 Lap:  98,24 seg. ----------------------------------------------------------------------------------------------------
'    Proceso Finalizado. 23-ene-26 19:07

