Attribute VB_Name = "M07_Actualiz_BDatos_con_LsGes04"
' Last Rev. 2026-09-23 18:56
' >>> DOC-MOD (generado) >>>
' =================================================================================================
' M07_Actualiz_BDatos_con_LsGes04 - Fusion de la importacion con la base de datos
' =================================================================================================
'
' PROPOSITO
'  El corazon del pipeline: vuelca Prog_LsGes04 (ya depurado y clasificado)
'  sobre Prog_BD, que es la base de datos historica y contiene ademas los
'  datos de gestion propios (JI, AD, RDT, organica, liquidado...) que NO
'  vienen de LSGES04 y hay que preservar.
'  Recorre las dos tablas EN PARALELO, ambas ordenadas por BD_Ref.
'
' INDICE DE RUTINAS Y FUNCIONES
'  RuT_Actualizar_BDatos_con_LsGes04 ... Unica rutina del modulo.
'
' TRAMOS DE PROGRAMACION
'    0. PREPARACION: desprotege ambas hojas, quita filtros y ordena las dos
'       tablas por BD_Ref. Limpia en Prog_BD las columnas BD_Incidencias y
'       BD_EP_GestReg (marcas de la pasada anterior).
'
'    BUCLE PRINCIPAL - avanza por LsGes04 comparando Val(BD_Ref) de cada tabla.
'    Tres casos, que es la clasica fusion de dos listas ordenadas:
'
'    a) REFERENCIAS IGUALES -> el recibo YA EXISTE: actualizar.
'       - Antes de sobrescribir, anota incidencias si cambian importes:
'         ImpRec, ImpCob (solo si ya habia cobro) o ImpAdm; guarda el valor
'         ANTERIOR en BD_Incidencias y en el historico BD_H_Incidencias.
'       - Si en LsGes04 viene BD_Anul = 'S', marca 'Mat.Anulada_' en observaciones.
'       - Copia por bloques contiguos (Resize, una sentencia por bloque):
'           columnas 1..BD_InfRegulariz      (datos originales del recibo)
'           BD_ACont_Vto..BD_Cta_Ing         (lo calculado en M01/M03/M04/M05)
'           BD_Rec_Imp_Acad, _Dto, _Adm      (lo imputado en M06)
'         Lo que NO entra en esos bloques (JI, AD, RDT, organica...) se conserva.
'       - Sella '- Actualizado <fecha> -' y avanza en las dos tablas.
'
'    b) REF. NUEVA (la de BD es mayor, o BD se acabo) -> ALTA.
'       Anade fila a Prog_BD, copia 1..BD_EP_GestReg y le asigna el coeficiente
'       de retencion del VRI: lo busca por plan en Prog_Coef_Ret_VRI y, si el
'       plan no existe, aplica el coeficiente por defecto segun APP_EFP_o_CFC.
'       Sella '- Nuevo <fecha> -'.
'
'    c) REF. ANTIGUA que ya no viene en LsGes04 -> BAJA.
'       Distingue dos situaciones, y es la distincion importante:
'         - Sin JI emitido        -> BD_Tipo_Rec = 'Deleted'.
'         - CON JI (BD_JI_Emi_Acad) -> 'DeletedConJI': el recibo ya genero un
'           documento contable, asi que su desaparicion es una ANOMALIA.
'       Solo avanza en Prog_BD (LsGes04 se queda donde esta).
'
'    Ademas: si dos filas seguidas de LsGes04 comparten referencia, marca ambas
'    '_Duplicaty' en BD_EP_Ctrl y salta la segunda.
'
'    BUCLE DE COLA: al agotarse LsGes04, todo lo que quede por recorrer en
'    Prog_BD son bajas, y se marcan con el mismo criterio Deleted/DeletedConJI.
'
'    CIERRE:
'     - Avisa por MsgBox si quedan recibos sin sellar en BD_EP_GestReg.
'     - Avisa por MsgBox si hubo bajas CON JI (requiere revision manual).
'     - Resumen al informe: altas, actualizaciones, bajas, bajas con JI y
'       desglose de cambios de importe.
'     - Filtra los '*Deleted *' y los traslada a Prog_BD_Deleted (el cuarto
'       argumento True de Rut_Lo_DataBodyRange_Filtered_Copy borra el origen).
'     - Rut_WrkSheet_ReducirPeso sobre Prog_LsGes04 y restaura totales.
'
' NOTAS
'  La fusion depende POR COMPLETO de que ambas tablas esten ordenadas por
'  BD_Ref y de que la comparacion sea numerica (Val). Cualquier cambio en la
'  ordenacion previa rompe el emparejamiento en silencio.
'
'  'DeletedConJI' no es un caso de borde: es la alarma de que se ha perdido un
'  recibo ya contabilizado.
' =================================================================================================
' <<< DOC-MOD (generado) <<<

'2026-01-09
Option Explicit

'            - Actualizar BDatos con Lo_Ges04
'            - Recorro toda LsGes04 para actualizar BDatos
'            --- Referencias IGUALES                            <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios
'            --- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO     <<<<  AÑADO UN NUEVO REGISTRO a BDatos
'            --- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO   <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
'            - Copiar Todos los Registros "Deleted" en Lo_Deleted y Borrarlos de BDatos

'- -------------------------------------------------------------------------------------------------
'- Actualizar BDatos con Lo_Ges04 ------------------------------------------------------------------
'- -------------------------------------------------------------------------------------------------
Sub RuT_Actualizar_BDatos_con_LsGes04()
Debug.Print ">>> RuT_Actualizar_BDatos_con_LsGes04"
    Dim F_Actualiz          As String:  F_Actualiz = Now()
    Dim Incidencia          As String
    Dim TxT_Progreso        As String
    Dim Ref_Ant             As String:      Ref_Ant = ""
    Dim Cont                As Integer
    Dim Cont_Fail           As Long:    Cont_Fail = 0
    Dim Cont_Repes          As Long:    Cont_Repes = 0
    Dim Cont_Nuevo          As Long:    Cont_Nuevo = 0
    Dim Cont_Modif          As Long:    Cont_Modif = 0
    Dim Cont_Deleted        As Long:    Cont_Deleted = 0
    Dim Cont_DeletedconJI   As Long:    Cont_DeletedconJI = 0
    Dim Cont_Mat_Anul       As Long:    Cont_Mat_Anul = 0
    Dim Chg_ImpRec          As Long:    Chg_ImpRec = 0
    Dim Chg_ImpCob          As Long:    Chg_ImpCob = 0
    Dim Chg_ImpAdm          As Long:    Chg_ImpAdm = 0
    Dim F_BD                As Long:    F_BD = 1
    Dim F_G4                As Long:    F_G4 = 1
    Dim rowfind             As Variant
        
    Dim Lo_BD               As ListObject:      Set Lo_BD = Prog_BD.ListObjects(1)
    Dim Lo_Ges04            As ListObject:      Set Lo_Ges04 = Prog_LsGes04.ListObjects(1)
    Dim Lo_Tb_Ret_VRI       As ListObject:      Set Lo_Tb_Ret_VRI = Prog_Coef_Ret_VRI.ListObjects(1)
    Dim TRows_BD         As Long:        TRows_BD = Lo_BD.ListRows.Count
    Dim TRows_Ges04         As Long:        TRows_Ges04 = Lo_Ges04.ListRows.Count
    '- Fase 4: el merge copia por posicion las col. 1..39 de Lo_Ges04 a Lo_BD. Si alguien
    '-  desalinea las constantes G04_*/BD_* en M00_Ini_Var_APP, paramos antes de escribir nada.
    If Not Fnc_G04_Alineada_con_BD() Then Err.Raise vbObjectError + 504, "M07", _
        "Las col. 1..39 de G04_* y BD_* no coinciden (revisar M00_Ini_Var_APP)."
    Prog_BD.Unprotect
    Prog_LsGes04.Unprotect
    Lo_BD.ShowTotals = False
    Lo_Ges04.ShowTotals = False
    
Call Rut_Off_Functions
    ' =============  Preparar Tabla de TitPH ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_BD)
    Call Rut_Lo_Sort(Lo_BD, BD_Ref, xlAscending, True)    ' Ordenar por una Columna
    ' =============  Preparar Tabla de TitPH ==================
    Call Rut_Lo_WrkSht_Preparar(Prog_LsGes04)
    Call Rut_Lo_Sort(Lo_Ges04, G04_Ref, xlAscending, True)    ' Ordenar por una Columna
            '- Visualizo el progreso
            TxT_Progreso = Form_Menu.TB_Informe & vbCrLf
            Form_Menu.TB_Informe = TxT_Progreso & "Incorporando LSGES04:  " & " 0 de " & Format(TRows_Ges04, "#,##0")
            Form_Menu.TB_Informe.SelStart = Len(Form_Menu.TB_Informe)
            Form_Menu.TB_Informe.SetFocus
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbLf & vbLf & vbLf
    '- ---------------------------------------------------------------------------------------------
    '- Actualizar BDatos con LsGes04 ---------------------------------------------------------------
    '        - Recorro toda LsGes04 para actualizar BDatos
    '        --- Referencias IGUALES                            <<<<  Ya existe en BDatos y hay que ver de Actualizar si hay Cambios
    '        --- Ref. NUEVA NO EXISTE, es un REGISTRO NUEVO     <<<<  AÑADO UN NUEVO REGISTRO a BDatos
    '        --- Ref. ANTIGUA NO EXISTE, es un REG. ELIMINADO   <<<< Lo marcamos y luego los copiamos en Lo_Deleted y Borramos de BDatos
    '- ---------------------------------------------------------------------------------------------
        
        If Not Lo_BD.DataBodyRange Is Nothing Then
            Lo_BD.ListColumns(BD_Incidencias).DataBodyRange.ClearContents    '- ClearContents -----
            Lo_BD.ListColumns(BD_EP_GestReg).DataBodyRange.ClearContents          '- ClearContents -----
        End If
    
    '- ===========================================================================================
    '-  FASE 2 (velocidad): este bucle era el 70-88% del tiempo de toda la importacion.
    '-  Antes hacia ~30 accesos COM por fila (RwDB.Range(...) / RwG4.Range(...)) sobre 8.400
    '-  filas = ~250.000 cruces VBA<->Excel. Ahora las dos tablas se vuelcan a arrays, TODO el
    '-  merge se resuelve en memoria y se escribe de una vez al final.
    '-  La logica de negocio es la MISMA: merge de dos tablas ordenadas por BD_Ref con la misma
    '-  comparacion Val(), las mismas 3 ramas y los mismos contadores.
    '- ===========================================================================================
    Dim aG4()       As Variant      '- Copia en RAM de Lo_Ges04
    Dim aBD()       As Variant      '- Copia en RAM de Lo_BD (se modifica y se vuelve a escribir)
    Dim aVRI()      As Variant      '- Tabla de Coef. VRI (para el Coef. de las altas)
    Dim aAltas()    As Variant      '- Registros NUEVOS, se anaden a Lo_BD de una sola vez
    Dim NumAltas    As Long:        NumAltas = 0
    Dim ColsBD      As Long:        ColsBD = Lo_BD.ListColumns.Count
    Dim c           As Long

    If TRows_Ges04 > 0 Then aG4 = Lo_Ges04.DataBodyRange.Value
    If TRows_BD > 0 Then aBD = Lo_BD.DataBodyRange.Value
    If Not Lo_Tb_Ret_VRI.DataBodyRange Is Nothing Then aVRI = Lo_Tb_Ret_VRI.DataBodyRange.Value
    ReDim aAltas(1 To TRows_Ges04 + 1, 1 To ColsBD)     '- Cota superior: como mucho, todas altas

    Do While F_G4 <= TRows_Ges04
        '- Guard de duplicados: CONSERVADO A PROPOSITO, pero hoy es DEFENSIVO ------------------
        '-  M02 (RuT_Duplicates_Search) ya borra los duplicados de Lo_Ges04 antes de llegar aqui
        '-  (en el log: "Del en Lo_Ges04 Rec. Repes NO finalistas"), asi que Cont_Repes no llega
        '-  a incrementarse nunca y no aparece en el informe. Se mantiene porque si algun dia M02
        '-  falla o se reordena el pipeline, sin este guard el merge emparejaria mal EN SILENCIO.
        If F_G4 > 1 Then
            If aG4(F_G4, G04_Ref) = Ref_Ant Then
                aG4(F_G4, G04_EP_Ctrl) = aG4(F_G4, G04_EP_Ctrl) & "_Duplicaty"
                aG4(F_G4 - 1, G04_EP_Ctrl) = aG4(F_G4 - 1, G04_EP_Ctrl) & "_Duplicaty"
                Cont_Repes = Cont_Repes + 1
                F_G4 = F_G4 + 1
                GoTo Siguiente_Reg
            End If
        End If

        '--- Referencias IGUALES <<<< Ya existe en BDatos: hay que ver si hay Cambios -----------
        If TRows_BD > 0 And Val(aBD(F_BD, BD_Ref)) = Val(aG4(F_G4, G04_Ref)) Then
            '- Compruebo posibles INCIDENCIAS --------------------------------------------------
            If aBD(F_BD, BD_ImpRec) <> aG4(F_G4, G04_ImpRec) * 1 Then          '- Cambio Imp. Recibo
                Incidencia = "Chg:PH_ImpRec=[" & aBD(F_BD, BD_ImpRec) & "]_#_"
                aBD(F_BD, BD_Incidencias) = aBD(F_BD, BD_Incidencias) & Incidencia
                aBD(F_BD, BD_H_Incidencias) = aBD(F_BD, BD_H_Incidencias) & Incidencia
                Chg_ImpRec = Chg_ImpRec + 1
            End If
            If aBD(F_BD, BD_ImpCob) > 0 And aBD(F_BD, BD_ImpCob) <> aG4(F_G4, G04_ImpCob) * 1 Then
                Incidencia = "Chg:PH_ImpCob=[" & aBD(F_BD, BD_ImpCob) & "]_#_"
                aBD(F_BD, BD_Incidencias) = aBD(F_BD, BD_Incidencias) & Incidencia
                aBD(F_BD, BD_H_Incidencias) = aBD(F_BD, BD_H_Incidencias) & Incidencia
                Chg_ImpCob = Chg_ImpCob + 1
            End If
            If aBD(F_BD, BD_ImpAdm) > 0 And aBD(F_BD, BD_ImpAdm) <> aG4(F_G4, G04_ImpAdm) * 1 Then
                Incidencia = "Chg:PH_ImpAdm=[" & aBD(F_BD, BD_ImpAdm) & "]_#_"
                aBD(F_BD, BD_Incidencias) = aBD(F_BD, BD_Incidencias) & Incidencia
                aBD(F_BD, BD_H_Incidencias) = aBD(F_BD, BD_H_Incidencias) & Incidencia
                Chg_ImpAdm = Chg_ImpAdm + 1
            End If
            If aG4(F_G4, G04_Anul) = "S" Then          '- Tasa Anulada
                aBD(F_BD, BD_Obs_Conta) = "Mat.Anulada_"
                Cont_Mat_Anul = Cont_Mat_Anul + 1
            End If

            '- Datos que vienen del LSGES04: columnas 1..BD_InfRegulariz ----------------------
            For c = 1 To BD_InfRegulariz
                aBD(F_BD, c) = aG4(F_G4, c)
            Next
            '- Datos calculados por el pipeline: BD_ACont_Vto..BD_Cta_Ing ---------------------
            '-  OJO: el codigo original hacia Resize(1, BD_Cta_Ing - BD_Concepto + 1) = 4 col.,
            '-  empezando en BD_ACont_Vto(34), con lo que copiaba 34..37 y DEJABA FUERA
            '-  BD_Cta_Ing(38), que M03 acaba de calcular. Confirmado con el usuario que SI debe
            '-  copiarse, asi que ahora el rango es BD_ACont_Vto..BD_Cta_Ing (34..38).
            For c = BD_ACont_Vto To BD_Cta_Ing
                aBD(F_BD, c) = aG4(F_G4, c)
            Next
            '- Importes de matricula que asigna M06 (col. contiguas 30..32) -------------------
            aBD(F_BD, BD_Rec_Imp_Acad) = aG4(F_G4, G04_Rec_Imp_Acad)
            aBD(F_BD, BD_Rec_Imp_Dto) = aG4(F_G4, G04_Rec_Imp_Dto)
            aBD(F_BD, BD_Rec_Imp_Adm) = aG4(F_G4, G04_Rec_Imp_Adm)

            '- Preparo Salto de registro ------------------------------------------------------
            aBD(F_BD, BD_EP_GestReg) = aBD(F_BD, BD_EP_GestReg) & "- Actualizado " & F_Actualiz & " - "
            aG4(F_G4, G04_EP_GestReg) = "Actualizado BD, " & F_Actualiz
            Cont_Modif = Cont_Modif + 1
            If F_BD < TRows_BD Then F_BD = F_BD + 1
            F_G4 = F_G4 + 1
            Ref_Ant = aG4(F_G4 - 1, G04_Ref)

        '--- Ref. NUEVA NO EXISTE <<<< ANADO UN NUEVO REGISTRO a BDatos -----------------------
        ElseIf TRows_BD = 0 Or Val(aBD(F_BD, BD_Ref)) > Val(aG4(F_G4, G04_Ref)) Or F_BD >= TRows_BD Then
            NumAltas = NumAltas + 1
            '- Todos los datos nuevos y anadidos, de Ges04 al registro de alta
            '- Fase 4: solo coinciden las col. 1..G04_Coef_VRI; las 4 de trabajo se mapean una a una
            '-  y los flags G04_Flag_* NO pasan a Prog_BD (decidido con el usuario el 2026-09-23).
            '-  Antes se copiaba 1..BD_EP_GestReg por posicion, y los flags Emitido/EjeAnt de
            '-  Lo_Ges04 acababan en BD_H_Incidencias/BD_EP_Ctrl del alta.
            For c = 1 To G04_Coef_VRI
                aAltas(NumAltas, c) = aG4(F_G4, c)
            Next
            aAltas(NumAltas, BD_Incidencias) = aG4(F_G4, G04_Incidencias)
            aAltas(NumAltas, BD_H_Incidencias) = aG4(F_G4, G04_H_Incidencias)
            aAltas(NumAltas, BD_EP_Ctrl) = aG4(F_G4, G04_EP_Ctrl)
            aAltas(NumAltas, BD_EP_GestReg) = aG4(F_G4, G04_EP_GestReg)
            ' -----------------=============  Buscar Tipo Plan  ==================-------------
            rowfind = Fnc_Buscar_Fila_VRI(aVRI, aG4(F_G4, G04_Plan))
            If rowfind > 0 Then         '- Plan Encontrado ==>> Tendra caracteristicas ESPECIALES
                aAltas(NumAltas, BD_Coef_VRI) = aVRI(rowfind, CoefVRI_CoefVRI)
            Else                        '- NO ENCONTRADO ==>> coeficiente establecido 15% o 20%
                If Prog__APP.Range("APP_EFP_o_CFC") = "EFP" Then
                    aAltas(NumAltas, BD_Coef_VRI) = aVRI(1, CoefVRI_CoefVRI)
                Else
                    aAltas(NumAltas, BD_Coef_VRI) = aVRI(2, CoefVRI_CoefVRI)
                End If
            End If
            '- Preparo Salto de registro ------------------------------------------------------
            aAltas(NumAltas, BD_EP_GestReg) = "- Nuevo " & F_Actualiz & " - "
            aG4(F_G4, G04_EP_GestReg) = "- Nuevo en BD, " & F_Actualiz
            Cont_Nuevo = Cont_Nuevo + 1
            F_G4 = F_G4 + 1
            Ref_Ant = aG4(F_G4 - 1, G04_Ref)

        '--- Ref. ANTIGUA NO EXISTE <<<< es un REG. ELIMINADO ---------------------------------
        Else
            If aBD(F_BD, BD_JI_Emi_Acad) = "" Then
                If InStr(aBD(F_BD, BD_EP_GestReg), "Deleted") = 0 Then
                    aBD(F_BD, BD_EP_GestReg) = aBD(F_BD, BD_EP_GestReg) & "- Deleted " & F_Actualiz & " - "
                    aBD(F_BD, BD_Tipo_Rec) = "Deleted"
                    Cont_Deleted = Cont_Deleted + 1
                End If
            Else
                If InStr(aBD(F_BD, BD_EP_GestReg), "Deleted-conJI") = 0 Then
                    aBD(F_BD, BD_EP_GestReg) = "- Deleted-conJI " & F_Actualiz & " - "
                    aBD(F_BD, BD_Tipo_Rec) = "DeletedConJI"
                    Cont_DeletedconJI = Cont_DeletedconJI + 1
                End If
            End If
            If F_BD < TRows_BD Then F_BD = F_BD + 1
        End If
                '- Visualizo el progreso -----------------------------------------------------
                If F_G4 Mod 2000 = 0 Then
                    Form_Menu.TB_Informe = TxT_Progreso & "Incorporando LSGES04:  " & Format(F_G4, "#,##0") & " de " & Format(TRows_Ges04, "#,##0") & " reg."
                End If
Siguiente_Reg:
    Loop
    '- Si quedan recibos en BDatos, NO EXISTEN en LsGes04 <<<<< SON REGISTROS ELIMINADOS ------
    Do While F_BD <= TRows_BD
        If aBD(F_BD, BD_JI_Emi_Acad) = "" Then
            If InStr(aBD(F_BD, BD_EP_GestReg), "Deleted") = 0 Then
                aBD(F_BD, BD_EP_GestReg) = aBD(F_BD, BD_EP_GestReg) & "- Deleted " & F_Actualiz & " - "
                aBD(F_BD, BD_Tipo_Rec) = "Deleted"
                Cont_Deleted = Cont_Deleted + 1
            End If
        Else
            If InStr(aBD(F_BD, BD_EP_GestReg), "Deleted-conJI") = 0 Then
                aBD(F_BD, BD_EP_GestReg) = "- Deleted-conJI " & F_Actualiz & " - "
                aBD(F_BD, BD_Tipo_Rec) = "DeletedConJI"
                Cont_DeletedconJI = Cont_DeletedconJI + 1
            End If
        End If
        F_BD = F_BD + 1
    Loop

    '- ===========================================================================================
    '-  VOLCADO A LAS HOJAS: 2 escrituras de bloque + 1 para las altas
    '- ===========================================================================================
    If TRows_BD > 0 Then Lo_BD.DataBodyRange.Value = aBD
    If TRows_Ges04 > 0 Then Lo_Ges04.DataBodyRange.Value = aG4

    '- Las ALTAS se anaden de UNA vez (antes era un Lo_BD.ListRows.Add por registro, que
    '-  reconstruye la tabla entera en cada llamada). Se amplia la tabla con Resize y se
    '-  escribe el bloque completo.  OJO: el 2o argumento de Resize es el ANCHO, no la col. final.
    If NumAltas > 0 Then
        '- OJO (bug corregido 2026-09-23): NO usar Lo_BD.DataBodyRange.Cells(TRows_BD + 1, 1)
        '-  como destino. TRows_BD es el tamano que tenia la tabla AL EMPEZAR, pero justo antes
        '-  se ha hecho 'Lo_BD.DataBodyRange.Value = aBD' y se va a redimensionar: el
        '-  DataBodyRange ya NO es el mismo rango y el bloque cae descolocado, replicando una
        '-  fila miles de veces fuera de la tabla. Se usan COORDENADAS ABSOLUTAS DE HOJA,
        '-  calculadas ANTES de tocar el tamano de la tabla.
        Dim aBloque()   As Variant
        Dim f2          As Long
        Dim wsBD        As Worksheet:   Set wsBD = Lo_BD.Parent
        Dim FilIniTb    As Long, ColIniTb As Long, ColFinTb As Long
        Dim FilPrimAlta As Long, FilFinTb  As Long

        ReDim aBloque(1 To NumAltas, 1 To ColsBD)
        For f2 = 1 To NumAltas
            For c = 1 To ColsBD
                aBloque(f2, c) = aAltas(f2, c)
            Next
        Next

        '- Geometria de la tabla ANTES de ampliarla -----------------------------------------
        FilIniTb = Lo_BD.Range.Row                          '- fila de la cabecera
        ColIniTb = Lo_BD.Range.Column
        ColFinTb = ColIniTb + Lo_BD.Range.Columns.Count - 1
        FilPrimAlta = FilIniTb + Lo_BD.Range.Rows.Count     '- 1a fila LIBRE tras la tabla
        FilFinTb = FilPrimAlta + NumAltas - 1               '- ultima fila tras ampliar

        '- 1o escribir las altas en la hoja, 2o ampliar la tabla para que las absorba -------
        wsBD.Range(wsBD.Cells(FilPrimAlta, ColIniTb), _
                   wsBD.Cells(FilFinTb, ColIniTb + ColsBD - 1)).Value = aBloque
        Lo_BD.Resize wsBD.Range(wsBD.Cells(FilIniTb, ColIniTb), wsBD.Cells(FilFinTb, ColFinTb))
        Erase aBloque
    End If

    Erase aAltas
    If TRows_Ges04 > 0 Then Erase aG4
    If TRows_BD > 0 Then Erase aBD
    
    '-Filtra Recibos Actualizar de BDatos ----------------------------------------------------------
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    rowfind = Application.WorksheetFunction.CountIf(Lo_BD.DataBodyRange.Columns(BD_EP_GestReg), "")
    If rowfind > 0 Then
        MsgBox "¡¡¡ Recibos SIN Actualizar !!!" & vbLf & vbLf & "¡¡¡ " & rowfind & "reg SIN ACTUALIZAR !!!", vbOKOnly + vbExclamation
    End If
    
    '  & Right(String(8, "_") & Format(.ListRows.Count, "#,##0"), 8) &
    
    
    
    Form_Menu.TB_Informe = TxT_Progreso & _
                 Right(String(8, "_") & Format(Cont_Nuevo, "#,##0"), 8) & " Incorporado Nuevos Recibos de LSGES04." & vbLf & _
                 Right(String(8, "_") & Format(Cont_Modif, "#,##0"), 8) & " Actualizados Recibos de BDatos con LSGES04." & vbLf & _
                 Right(String(8, "_") & Format(Cont_Deleted, "#,##0"), 8) & " Recibos de BDatos Deleted." & vbLf & _
                 Right(String(8, "_") & Format(Cont_DeletedconJI, "#,##0"), 8) & " ¡Ojo! Recibos de BDatos Borrados con JI." & vbLf & _
        Right("__________" & Chg_ImpRec + Chg_ImpCob + Chg_ImpAdm, 8) & "  Reg. Actualizados con incidencias (Cambios destacables)." & vbCrLf & _
        Right("__________" & Chg_ImpRec, 8) & "  Reg. Cambio en Importe de Recibo." & vbCrLf & _
        Right("__________" & Chg_ImpCob, 8) & "  Reg. Cambio en Importe Cobrado." & vbCrLf & _
        Right("__________" & Chg_ImpAdm, 8) & "  Reg. Cambio en Importe Administrativo." & vbCrLf & vbCrLf & _
        "Hay activos ahora un Total de:  " & Format(Lo_BD.ListRows.Count, "#,##0") & "  Reg." & _
                 Right(String(8, "_") & Format(rowfind, "#,##0"), 8) & " ¡Ojo! Recibos de BDatos Sin Actualizar."
    TxT_Progreso = Form_Menu.TB_Informe & vbCrLf
    
    
    If Cont_DeletedconJI > 0 Then
        MsgBox "¡¡¡ Recibos Borrados con JI's !!!" & Cont_DeletedconJI & " Recibos de BDatos Borrados con JI !!!", _
                vbOKOnly + vbExclamation, "Proceso: Actualizar BDatos con LsGes04"
    End If
    
    '- Copiar Todos los Registros "Deleted" en Lo_Deleted y Borrarlos de BDatos --------------------
    Dim Lo_Deleted      As ListObject:      Set Lo_Deleted = Prog_BD_Deleted.ListObjects(1)
    Call Rut_Lo_Sort(Lo_BD, BD_EP_GestReg, xlAscending, True)
    Lo_BD.Range.AutoFilter Field:=BD_EP_GestReg, Criteria1:="=*Deleted *"
    Call Rut_Lo_DataBodyRange_Filtered_Copy(Lo_BD, Lo_Deleted, False, True)
    Call Rut_Lo_Filtros_Quitar(Lo_BD)
    
    '- Visualizo el progreso
            Form_Menu.TB_Informe = Form_Menu.TB_Informe & vbCrLf & Format(Now, "hh:mm:ss") & _
                                " Actualizada BDatos con: " & Format(Lo_BD.ListRows.Count, "#,##0") & "reg." & vbCrLf

Call Rut_Lo_Filtros_Quitar(Lo_BD)
Call Rut_WrkSheet_ReducirPeso(Prog_LsGes04.Name, True)
       
Lo_BD.ShowTotals = True
Lo_Deleted.ShowTotals = True
Lo_Ges04.ShowTotals = True
Debug.Print "<<< RuT_Actualizar_BDatos_con_LsGes04"
End Sub

' ==================================================================================================
Function Fnc_Buscar_Fila_VRI(ByRef aVRI As Variant, ByVal Cod_Plan As Variant) As Long
' ==================================================================================================
'-  Equivalente en memoria de Application.Match(Plan, Lo_Tb_Ret_VRI...Columns(1), 0).
'-  Se usa dentro del bucle de altas de M07: llamar a Application.Match ahi dentro volveria a
'-  cruzar la frontera VBA<->Excel en cada alta, que es justo lo que la Fase 2 elimina.
'-  Devuelve la fila (1..n) o 0 si no lo encuentra.
    Dim f       As Long
    Fnc_Buscar_Fila_VRI = 0
    If IsEmpty(aVRI) Then Exit Function
    For f = LBound(aVRI, 1) To UBound(aVRI, 1)
        If aVRI(f, CoefVRI_Plan) = Cod_Plan Then
            Fnc_Buscar_Fila_VRI = f
            Exit Function
        End If
    Next
End Function    ' Fnc_Buscar_Fila_VRI
' --------------------------------------------------------------------------------------------------

'===================================================================================================
Function Fnc_G04_Alineada_con_BD() As Boolean   '- Las col. 1..39 de Tb_LsGes04 y Prog_BD coinciden?
'===================================================================================================
'-  Fase 4 (2026-09-23): M07 (merge y altas) y Rut_Lo_G04_Filtered_Copy_a_BD copian POR POSICION
'-  las col. 1..G04_Coef_VRI. Esta funcion es la red de seguridad de esa suposicion.
    Dim Ok      As Boolean:     Ok = True     '- Una comparacion por linea (VBA admite solo 24 ' _')
    Ok = Ok And (G04_ACont_Emi = BD_ACont_Emi)
    Ok = Ok And (G04_ACont_Cob = BD_ACont_Cob)
    Ok = Ok And (G04_Plan = BD_Plan)
    Ok = Ok And (G04_NomPlan = BD_NomPlan)
    Ok = Ok And (G04_TipoCurso = BD_TipoCurso)
    Ok = Ok And (G04_C_Acad = BD_C_Acad)
    Ok = Ok And (G04_Nom = BD_Nom)
    Ok = Ok And (G04_DNI = BD_DNI)
    Ok = Ok And (G04_Matricula = BD_Matricula)
    Ok = Ok And (G04_Anul = BD_Anul)
    Ok = Ok And (G04_Ref = BD_Ref)
    Ok = Ok And (G04_NumRec = BD_NumRec)
    Ok = Ok And (G04_ActivEco = BD_ActivEco)
    Ok = Ok And (G04_FEmi = BD_FEmi)
    Ok = Ok And (G04_FVto = BD_FVto)
    Ok = Ok And (G04_FCob = BD_FCob)
    Ok = Ok And (G04_ImpRec = BD_ImpRec)
    Ok = Ok And (G04_ImpCob = BD_ImpCob)
    Ok = Ok And (G04_FormPag = BD_FormPag)
    Ok = Ok And (G04_CtaPag = BD_CtaPag)
    Ok = Ok And (G04_RegMov = BD_RegMov)
    Ok = Ok And (G04_Grupo = BD_Grupo)
    Ok = Ok And (G04_ImpAcad = BD_ImpAcad)
    Ok = Ok And (G04_ImpAdm = BD_ImpAdm)
    Ok = Ok And (G04_Expdte = BD_Expdte)
    Ok = Ok And (G04_HTipCob = BD_HTipCob)
    Ok = Ok And (G04_Hinvalid = BD_Hinvalid)
    Ok = Ok And (G04_ImpDto = BD_ImpDto)
    Ok = Ok And (G04_InfRegulariz = BD_InfRegulariz)
    Ok = Ok And (G04_Rec_Imp_Acad = BD_Rec_Imp_Acad)
    Ok = Ok And (G04_Rec_Imp_Adm = BD_Rec_Imp_Adm)
    Ok = Ok And (G04_Rec_Imp_Dto = BD_Rec_Imp_Dto)
    Ok = Ok And (G04_Rec_Imp_INSS = BD_Rec_Imp_INSS)
    Ok = Ok And (G04_ACont_Vto = BD_ACont_Vto)
    Ok = Ok And (G04_Concepto = BD_Concepto)
    Ok = Ok And (G04_Tipo_Rec = BD_Tipo_Rec)
    Ok = Ok And (G04_Tipo_EP = BD_Tipo_EP)
    Ok = Ok And (G04_Cta_Ing = BD_Cta_Ing)
    Ok = Ok And (G04_Coef_VRI = BD_Coef_VRI)
    Fnc_G04_Alineada_con_BD = Ok
End Function    ' Fnc_G04_Alineada_con_BD   --------------------------------------------------------
'===================================================================================================

'===================================================================================================
Sub Rut_Lo_G04_Filtered_Copy_a_BD(Lo_G04 As ListObject, Lo_Tgt As ListObject)   '- Filas visibles de G04 a una tabla BD
'===================================================================================================
'-  Copia las filas VISIBLES de Lo_G04 (Tb_LsGes04) al final de una tabla con la estructura de
'-  Prog_BD (Tb_Duplic, Tb_BDatos). Fase 4 (2026-09-23): las dos tablas ya NO tienen la misma
'-  numeracion, asi que no vale Rut_Lo_DataBodyRange_Filtered_Copy (copia por posicion):
'-      G04 1..G04_Coef_VRI                  ->  BD 1..BD_Coef_VRI       (coinciden)
'-      G04_Incidencias..G04_EP_GestReg      ->  BD_Incidencias..BD_EP_GestReg
'-      G04_Flag_*                           ->  NO se copian
'-  Se usa Range.Copy (no .Value) para conservar los colores con que M02 marca los duplicados.
'-  Geometria con coordenadas NUMERICAS, calculadas antes de redimensionar la tabla destino.
    Dim Rg_Vis          As Range
    Dim Rg_Bloque1      As Range
    Dim Rg_Bloque2      As Range
    Dim ws              As Worksheet:   Set ws = Lo_Tgt.Parent
    Dim Sw_ShowTotals   As Boolean:     Sw_ShowTotals = Lo_Tgt.ShowTotals
    Dim Sw_DispAlerts   As Boolean:     Sw_DispAlerts = Application.DisplayAlerts
    Dim FilCabecera     As Long, ColIniTgt As Long, ColFinTgt As Long
    Dim StartRowAdd     As Long, NumFilas  As Long

    If Not Fnc_G04_Alineada_con_BD() Then Err.Raise vbObjectError + 504, "Rut_Lo_G04_Filtered_Copy_a_BD", _
        "Las col. 1..39 de G04_* y BD_* no coinciden (revisar M00_Ini_Var_APP)."
    If Lo_G04.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    Set Rg_Vis = Lo_G04.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Rg_Vis Is Nothing Then Exit Sub

    Lo_Tgt.ShowTotals = False
    Application.DisplayAlerts = False
    With Lo_G04.DataBodyRange
        Set Rg_Bloque1 = Intersect(Rg_Vis, .Columns(1).Resize(, G04_Coef_VRI))
        Set Rg_Bloque2 = Intersect(Rg_Vis, .Columns(G04_Incidencias).Resize(, G04_EP_GestReg - G04_Incidencias + 1))
        NumFilas = Intersect(Rg_Vis, .Columns(1)).Cells.Count    '- Cuenta TODAS las areas visibles
    End With

    FilCabecera = Lo_Tgt.HeaderRowRange.Row
    ColIniTgt = Lo_Tgt.Range.Column
    ColFinTgt = ColIniTgt + Lo_Tgt.Range.Columns.Count - 1
    If Lo_Tgt.DataBodyRange Is Nothing Then
        StartRowAdd = FilCabecera + 1
    Else
        StartRowAdd = FilCabecera + Lo_Tgt.ListRows.Count + 1
    End If

    Rg_Bloque1.Copy Destination:=ws.Cells(StartRowAdd, ColIniTgt + BD_ACont_Emi - 1)
    Rg_Bloque2.Copy Destination:=ws.Cells(StartRowAdd, ColIniTgt + BD_Incidencias - 1)
    Application.CutCopyMode = False

    '- Como se copia un Rango, Lo_Tgt NO se expande sola: se redimensiona para absorber las filas
    Lo_Tgt.Resize ws.Range(ws.Cells(FilCabecera, ColIniTgt), ws.Cells(StartRowAdd + NumFilas - 1, ColFinTgt))
    Lo_Tgt.ShowTotals = Sw_ShowTotals
    Application.DisplayAlerts = Sw_DispAlerts
End Sub     ' Rut_Lo_G04_Filtered_Copy_a_BD   ------------------------------------------------------
'===================================================================================================
