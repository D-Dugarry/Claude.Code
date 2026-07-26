Attribute VB_Name = "M_3___________________________"
'2025-12-27
' Actualizo la sheet BD_INSS que contiene los Rec_INSS de 2 C_Acad, el Ant y el Pos
'   el C_Acad_Ant, no debería de cambiar a partir de Julio del 2º año del curso, puesto que se ha acabado el curso.
'   el C_Acad_Pos, habrá que actualizarlo en la misma fecha que el LSace04 para el cierre del año contable, y cada vez que se requiera generar los JI's

' Procedimiento:
'
' Rut_Lo_Import_LoData_LoDefCol_LSace06
'
'    Importar LSsace06 de Curso_Acad_Ant o Curso_Acad_Pos
'           Para Identificar Rec. Matrículas con seguro obligatorio INSS en Tasa Adm.
'           Tengo que tener los dos Cursos Acad actualizados.
'
'    - Select File
'    - Con el ClsBk: (RAM)
'    - Compruebo ClsBk:
'    _            1º Si he dado una SheetNom, ver si Existe
'    -            2º Si no Existe LisObject la creo
'    -            3º Comprobar que la Tabla Lo_ClsBk tiene datos
'    -            4º Comprobar que la cabecera de la Tabla Lo_ClsBk corresponde con la establecida en la LoDefCol
'    - --------------------------------------------------------------------------------------------------------------
'    - Proceso ClsBk:
'    -            Formateo.
'    -            Determinar si en el LSace06 Hay UNO o DOS Cursos Académicos, Crea Dictionary para valores únicos (eficiente para grandes datos)
'    -            Borrar de Lo_ClsBk Recibos de LSave06 con C_Acad <> C_Acad_Ant y C_Acad_Pos, para aligerar el peso de la Tabla.
'    -            Borrar de Lo_ClsBk Recibos "<>INSS" en Nom_Concepto, para aligerar el peso de la Tabla.
'    -            Borrar Recibos con Imp.Rec. < 0.
'    -            M_314_Find_Rec_INSS, Identificar de un C_Acad, los 1º Rec. con Seguro obligatorio INSS, borrar los que no son.
'    -  ?????          Borrar Recibos de C_Acad_Ant y Cobrados en Año_Cont_Ant.
'    -            Borro los datos de las columnas NO necesarias.
'    - --------------------------------------------------------------------------------------------------------------
'    - Copy ClsBk:
'    -            1º Borrar los Recibos de Lo_Data (Tabla DB_INSS) con el/los C_Acad del Nuevo LSace06
'                       Puesto que voy a copiar los nuevos registros, tengo que eliminar los viejos.
'                       Dependiendo de si en Lo_Source hay 1 ó 2 Cursos Académicos, filtro por 1 o 2 Cursos.
'    -            2º Copiar Lo_ClsBk en Lo_Data.
'    -            3º Oculto columnas en Lo_Data de BD_INSS
'    - --------------------------------------------------------------------------------------------------------------

' Repetir el procedimiento para cada C_Acad si se requiere.


'################################################################################################################
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
' - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME - INFORME -
'################################################################################################################

















