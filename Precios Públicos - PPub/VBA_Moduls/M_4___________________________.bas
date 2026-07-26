Attribute VB_Name = "M_4___________________________"
'   EN ESTE MÓDULO PROCESO TODO LO QUE PUEDO EN EL ClsBk PARA TRABAJAR AL MÁXIMO EN LA RAM

' M_410_Update_Lo_AE4
'
'    Importar las 4 ListObjects de los WB_AE4: EFP y CFCyAFC de Curso_Acad_Ant y Curso_Acad_Pos.
'    Rut_Lo_Import_LoData_LoDefCol_AE4x1, 4 veces
'
'    M_411_Import_AE4:
'    - Rut_Lo_Import_LoData_LoDefCol_AE4x1
'         - Rut_File_Select_V2 (4 veces; una por cada Fich.)
'         - Con el ClsBk: (RAM)
'         - Compruebo ClsBk:
'                 - Comprobar que la Hoja Existe, SI hemos solicitado una hoja concreta para copiar
'                 - Crear, si NO Existe ListObject Lo_ClsBk
'                 - Comprobar que la Tabla Lo_ClsBk No está vacía
'                 - Comprobar que la cabecera de la Tabla Lo_ClsBk corresponde con la establecida en la LoDefCol
'         - Proceso ClsBk: (para transferir sólo los recibos del AñoCont y no los del C_Acad)
'                 - Borrar Recibos ACont_Emi = AñoCont-1 y Acont_Cob <> AñoCont
'                 - Borrar Recibos ACont_Emi > AñoCont
'                 - Borrar Datos de Rec. con F_Cob > APP_FechCierreCont: Vaciar/Clear las Columnas BD_FCob, BD_ImpCob, BD_FormPag, BD_CtaPag y BD_HTipCob
'         - Copy ClsBk:
'                - Si SW_Del_LoData=true Borrar Lo_AE4x1 para iniciarlo
'                - Añadir Lo_ClsBk al final de Lo_AE4x1.
'
'    M_415_Copy_AE4_a_BD
'    - Rut_Copy_AE4x4_en_BDatos:
'         -            1º Borrar en Bdatos, todos los AE4 de los recibos de C_Acad_Ant y C_Acad_Pos
'                            OJO sólo los de C_Acad Ant y Pos porque ocurrió que habían Rec. de Cursos más antiguos !!!
'         -            2º Copiar Lo_AE4 en BDatos.



'################################################################################################################
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
'################################################################################################################
'    Proceso: Importar Rec. AE4 - EFPyAFC Del Curso_Acad 2024-25 y 2025-26


'    Proceso: Importar Rec. AE4 - EFPyAFC Del Curso_Acad 2024-25 y 2025-26
'    - Sólo los Recibos del Año Contable: 2025 (Del Acont_Cob = AñoCont-1), y añadir Rec. a tabla BDatos.
'
'    Importarmos 1º: EFP_2024-25
'    Excel Seleccionado: EFP_2024-25_BaseDatos_Liq_V2.4.xlsm.............883 reg
'              Borrados Rec. ACont_Emi = 2024 y ACont_Cob ?  2025.............. 500 reg......quedan 383 reg
'              No hay Rec. ACont_Emi > 2025
'              No hay Rec. con F_Cob > 31/12/2025
'    10:41:01 Lap:   0,14 seg. Copiados los nuevos recibos:
'                              en BD_AE4x4 Rec. de EFP_2024-25_BaseDatos........383 reg......Total: 383 reg
'
'    Importarmos 2º: EFP_2025-26
'    Excel Seleccionado: EFP_2025-26_BaseDatos_Liq_V2.4.xlsm...........1.018 reg
'              No hay Rec. ACont_Emi = 2024 y ACont_Cob ?  2025
'              Borrados Rec. ACont_Emi > 2025................................... 50 reg......quedan 968 reg
'              Clear Data en Rec. con F_Cob > 31/12/2025...................en  77 reg.
'    10:41:22 Lap:   1,50 seg. Copiados los nuevos recibos:
'                              en BD_AE4x4 Rec. de EFP_2025-26_BaseDatos........968 reg....Total: 1.351 reg
'
'    Importarmos 3º: CFCyAFC_2024-25
'    Excel Seleccionado: CFCyAFC_2024-25_BaseDatos_Liq_V2.4.xlsm...........7.176 reg
'              Borrados Rec. ACont_Emi = 2024 y ACont_Cob ?  2025............ 4.170 reg....quedan 3.006 reg
'              No hay Rec. ACont_Emi > 2025
'              Clear Data en Rec. con F_Cob > 31/12/2025....................en  1 reg.
'              Borrados Rec. ¡¡ M013 !!.......................................... 2 reg....quedan 3.004 reg
'    10:41:41 Lap:   4,71 seg. Copiados los nuevos recibos:
'                              en BD_AE4x4 Rec. de CFCyAFC_2024-25_BaseD......3.004 reg....Total: 4.355 reg
'
'    Importarmos 4º: CFCyAFC_2025-26
'    Excel Seleccionado: CFCyAFC_2025-26_BaseDatos_Liq_V2.4.xlsm...........6.223 reg
'              No hay Rec. ACont_Emi = 2024 y ACont_Cob ?  2025
'              Borrados Rec. ACont_Emi > 2025.................................. 199 reg....quedan 6.024 reg
'              Clear Data en Rec. con F_Cob > 31/12/2025...................en  21 reg.
'              Borrados Rec. ¡¡ M013 !!......................................... 88 reg....quedan 5.936 reg
'    10:42:06 Lap:   8,88 seg. Copiados los nuevos recibos:
'                              en BD_AE4x4 Rec. de CFCyAFC_2025-26_BaseD......5.936 reg...Total: 10.291 reg
'    No hay en BDatos Rec. AE4_2024-25 o AE4_2025-26...tiene 170.454 reg
'    Copy en BDatos, Rec. AE4_2024-25 y AE4_2025-26......................... 10.291 reg.....BD= 180.745 reg
'
'    10:42:20 Lap:   6,31 seg.
'    ----------------------------------------------------------------------------------------------------
'    Proceso Finalizado. 20-ene-26 10:42

