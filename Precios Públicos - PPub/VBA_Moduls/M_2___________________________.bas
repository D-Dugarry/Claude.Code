Attribute VB_Name = "M_2___________________________"
'    '- Importar LSGES04_GE por Curso_Acad_Ant, para hallar Imp.Acad. Imp.TAdm. e Imp.Dto del CAcad.Ant pagadas este AñoCont.
'
'    - Rut_Lo_Import_LoData_LoDefCol, Import LSGES04 por Curso_Acad_Ant
'    - Rut_Lo_ListColumns_ClearContents_DefC_ProtectData, Borrar por protección de Datos, Información sensible y no necesarias, según DefCol
'    - Rut_Lo_Format_LoData_LoDefColData, Format Sht__BD_IAdm_CAcadAnt
'    - M_211, Filtrar y Borrar Registros NO deseados:
'        - Borrar Recibos AE4 Enseñanzas Propias
'        - Borrar Recibos de Movimiento Menos AE=80 'Pruebas Acceso UA'
'        - Borrar Recibos con Imp.Rec. < 0
'        - Borrar Recibos de Matrículas de coste CERO
'        - Borrar Borrar Recibos - Subvencionado-, Imp_Rec =0 porque Imp_Dto >0
'        - Borrar Recibos con DNI=1 ==>> "NO BORRAR NO BORRAR, FICTICIO PARA RECIBOS"
'        - Borrar Recibos BD_C_Acad <> C_Acad_Ant
'        - Borrar Recibos ANULADOS
'        - Borrar Recibos NO Martrícula
'        - Borrar Recibos INVALIDADOS
'    - M_212, Gestionar Duplicados
'    - M_214_Find_IAdm_CAcad_Ant, Identificar Reg. de ImpAcad. ImpAdm. e ImpDto - Y - Borrar Reg SIN esos Datos

'################################################################################################################
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
'################################################################################################################

'    Proceso: Importar LSGES04 Del C_Acad_Ant 2024-25
'     Para extraer Recibos Académicos de TIO y EP. con Imp.Adm del Curso 2024-25
'     Y añadir el Imp.Acad. e Imp.Adm. a los Rec. de BDatos.
'    Seleccionar el fichero y la ruta, para importar: LSGES04_GE_SinDtos_Curso_2024-25
'    14:55:46 Lap:   5,80 seg. Excel Seleccionado: LSGES04_GE_SinDtos_Curso_2024-25_BDatos (2025-10-29).xlsx
'    Importado: LSGES04_GE_SinDtos_Curso_2024-25_BDatos (2025-10-29).xlsx
'    14:56:07 Lap:  20,16 seg. Importados nuevos datos: .............................. ........161.179 reg.
'    14:56:07 Lap:   0,00 seg. ¡ Información sensible SIN Eliminar !
'    14:56:20 Lap:  12,82 seg. Formateado el Excel, con  161.179 reg.
'    14:56:20 Lap:   0,42 seg. No hay Rec. de C_Acad. <> 2024-25
'    14:56:21 Lap:   0,53 seg. Borrados Rec. de Movimiento. .............. 13.650 reg. .... de 147.529 reg.
'    14:56:21 Lap:   0,44 seg. Borrados Rec. AE4 ......................... 18.417 reg. .... de 129.112 reg.
'    14:56:22 Lap:   0,91 seg. Borrados Rec. Negativos. ................... 3.738 reg. .... de 125.374 reg.
'    14:56:23 Lap:   0,84 seg. No hay Rec. Ficticio, DNI=1
'    14:56:23 Lap:   0,30 seg. No hay Rec. NO Matrícula
'    14:56:24 Lap:   0,95 seg. Borrados Rec. Anulados ..................... 3.053 reg. .... de 122.321 reg.
'    14:56:24 Lap:   0,28 seg. No hay Rec. Invalidados
'    14:56:26 Lap:   1,36 seg. Del Rec. de Matrícula_Cero ..................... 4 reg. .... de 122.317 reg.
'    14:56:29 Lap:   3,34 seg. Find Ref. con Repeticiones: 7218 y Máx nº Repeticiones, 4 veces.
'                              Borrados __________________________8007 reg., Quedan 114.310 reg.
'
'    14:56:33 Lap:   3,45 seg. Resultado de Identificar Importes del Curso: 2024-25, en: 114.310reg.
'
'         Borrados Rec. SIN Imp.Adm.Acad.Dto  86.903 reg.,  quedan 27.407 reg.
'         Identificas, Tasas Acad. de Matrícula por un importe de:     26.232.173,87€
'         Identificas, Tasas Adm.  de Matrícula por un importe de:        577.811,35€
'         Identificas, Descuentos  de Matrícula por un importe de:       -618.607,00€
'
'         Borrados Rec. Cobrados en Año_Cont_Ant 2024,  25.753 reg.,  quedan 1.654 reg.
'         Identificas, Tasas Acad. de Matrícula por un importe de:        801.621,51€
'         Identificas, Tasas Adm.  de Matrícula por un importe de:         51.232,57€
'         Identificas, Descuentos  de Matrícula por un importe de:        -12.432,00€
'
'    Procedimiento: Incorporar Imp.Acad./Adm./Dto., C_Acad_Ant_2024-25, a BDatos de Año_Cont_2025
'    14:56:51 Lap: 810,65 seg. Incorporados datos de: 1.006 reg. de un total de 1.655reg.
'         Tasas Acad. de Matrícula por un importe de:        475.839,47€.......en 1.006 reg,
'         Tasas Adm.  de Matrícula por un importe de:         29.940,89€.........en 998 reg,
'         Descuentos  de Matrícula por un importe de:        -11.484,00€..........en 35 reg,
'
'    14:56:53 Lap:  73,66 seg. ----------------------------------------------------------------------------------------------------
'    Proceso Finalizado. 14-ene-26 14:56

