Attribute VB_Name = "M_130_Gen_Inf_Recibos"
' Last Rev. 2026-10-14 13:48
'Rev.: 2026-02-15
Option Explicit

'- Genera la Tabla Informe_Recibos
    '- Borrado Tabla explicativa de Tipos de Recibos: Emitido, ADxAplz, Aplazado, EjeAnt y Añejo.
    '- Vaciar Lo_Inf, Añadir Leyenda con AnoCont y Rellenar Lo_inf ----------------------------------------------
    '- Copio la Tabla explicativa de los Tipos de Recibos: Prog_TipoRec.ListObjects(1) ---
    '- Filas EURLE (Esc.Univ. Rel.Lab. Elda), las filtro, las copio al final del Informe y las borro de Lo_Inf.
    '- Las sumas de BDatos se hacen en memoria (Rut_Inf_Recibos_Sumas), sin ordenar BDatos (2026-10-12).
    '- Los JI's de BDatos van al informe en esa misma pasada (2026-10-13; antes los ponía M_132, con la forma
    '  «Añadir JI's de BD al Inf.»). BDatos queda protegida como estaba (Rut_Prot_Save/Rut_Prot_Restore).
    '- Motor común de los dos informes (2026-10-14; Docs/Plan_Inf_Conta_TIO.md): Rut_Inf_Calcular(Def) calcula el informe que
    '  describe Def (T_Inf_Def): Inf_Recibos_TIO (Rut_Inf_Def_Recibos, aquí) o Inf_Conta_TIO (Rut_Inf_Def_Conta, M_140), que no
    '  tiene las columnas del INSS y en el que Emitido, Cobrado y Pendiente son Adm. + Acad.
    '- Inf_Conta_TIO lleva debajo una tabla con el total del INSS (2026-10-14, Rut_Inf_Tabla_INSS).

'- Índices del tercer índice de Sumas(tipo de recibo, concepto, Sum_*), que rellena Rut_Inf_Recibos_Sumas ------------------------
Private Const Sum_Tot_Emi       As Long = 1     '- Imp.Rec. de los emitidos
Private Const Sum_Tot_Cob       As Long = 2     '- Imp.Cob. de los cobrados
Private Const Sum_INSS_Emi      As Long = 3     '- Rec_Imp_INSS de los emitidos
Private Const Sum_INSS_Cob      As Long = 4     '- Rec_Imp_INSS de los cobrados
Private Const Sum_Adm_Emi       As Long = 5     '- Rec_Imp_Adm de los emitidos
Private Const Sum_Adm_Cob       As Long = 6     '- Rec_Imp_Adm de los cobrados
Private Const Sum_Emi_Ant       As Long = 7     '- Imp.Rec. de los emitidos, curso Ant (distinto del Pos)
Private Const Sum_Emi_Pos       As Long = 8     '- Imp.Rec. de los emitidos, curso Pos
Private Const Sum_Adm_Ant       As Long = 9     '- Rec_Imp_Adm de los emitidos, curso Ant
Private Const Sum_Adm_Pos       As Long = 10    '- Rec_Imp_Adm de los emitidos, curso Pos
Private Const Sum_Cob_Ant       As Long = 11    '- Imp.Cob. cobrado en el año contable, curso Ant
Private Const Sum_Cob_Pos       As Long = 12    '- Imp.Cob. cobrado en el año contable, curso Pos
Private Const Sum_CobAdm_Ant    As Long = 13    '- Rec_Imp_Adm cobrado en el año contable, curso Ant
Private Const Sum_CobAdm_Pos    As Long = 14    '- Rec_Imp_Adm cobrado en el año contable, curso Pos
Private Const Sum_N             As Long = 14

'- JI's y AD's (2026-10-13): las 6 columnas van seguidas y en el mismo orden en BDatos (BD_JI_Emi_Adm..BD_JI_443_Acad) y en
'  el informe (InfRec_JI_Emi_Adm..InfRec_JI_443_Acad). JI_Sin_Num: hay recibos de ese tipo y concepto, pero ninguno tiene JI.
Public Const JI_N_Col          As Long = 6
Public Const JI_Sin_Num        As String = "·-·"

'- Descripción de un informe (2026-10-14). La rellenan Rut_Inf_Def_Recibos (aquí), Rut_Inf_Def_Conta (M_140) y
'  Fnc_Inf_Def_Hoja (la de la hoja activa, para los botones del Ribbon). La usan los motores de M_130 (cálculo), M_133 (JI's a
'  BDatos) y M_135 (correo). Col(InfRec_X) es la columna del dato InfRec_X en la tabla de ese informe (0 = no la tiene).
Public Type T_Inf_Def
    Ws                      As Worksheet    '- hoja del informe
    Nom                     As String       '- "Informe de Recibos" / "Informe Conta": para los textos de las preguntas y el correo
    Titulo                  As String       '- título de B2, delante del año contable
    Con_INSS                As Boolean      '- True: columnas del INSS y Emitido = Imp.Rec.; False: sin ellas y Emitido = Adm. + Acad.
    Col(1 To InfRec_N_Col)  As Integer      '- columna de cada dato en la tabla, por su constante InfRec_*
    Nom_Calc                As String       '- nombre de Excel con la fecha del último cálculo (App_Last_Calc_<hoja>)
    Nom_Mail                As String       '- nombre de Excel con los destinatarios del correo
    Nom_Ruta                As String       '- nombre de Excel con la carpeta de red de la copia
    Nom_Xlsx                As String       '- nombre del adjunto, delante de " (datos del AAAA-MM-DD).xlsx"
    Desglose                As String       '- cuerpo del correo: cómo van desglosados los importes (HTML)
End Type

' ==================================================================================================================================
Sub Rut_Recalcular_Tabla_Inf_Recibos()          '- botón del ábaco en Inf_Recibos_TIO
' ==================================================================================================================================
    Dim Def         As T_Inf_Def
    Call Rut_Inf_Def_Recibos(Def)
    Call Rut_Inf_Calcular(Def)
End Sub     '- Rut_Recalcular_Tabla_Inf_Recibos
' ----------------------------------------------------------------------------------------------------------------------------------
'- Descripción de Inf_Recibos_TIO (T_Inf_Def): sus columnas son las de las constantes InfRec_*.
Public Sub Rut_Inf_Def_Recibos(Def As T_Inf_Def)
    Dim K           As Integer
    Set Def.Ws = Sht__Inf_Recibos_TIO
    Def.Nom = "Informe de Recibos"
    Def.Titulo = "Informe Recibos"
    Def.Con_INSS = True
    For K = 1 To InfRec_N_Col
        Def.Col(K) = K
    Next K
    Def.Nom_Calc = "App_Last_Calc_Inf_Recibos_TIO"
    Def.Nom_Mail = "App_Mail_Inf_Recibos_TIO"
    Def.Nom_Ruta = "App_RutaRed_Inf_Recibos_TIO"
    Def.Nom_Xlsx = "Inf_Recibos_TIO"
    Def.Desglose = "desglosados en seguro obligatorio (INSS), administrativos y académicos"
End Sub     '- Rut_Inf_Def_Recibos
' ----------------------------------------------------------------------------------------------------------------------------------
'- Descripción del informe de la hoja Ws (la activa, desde los botones del Ribbon). False si no es Inf_Recibos_TIO ni
'  Inf_Conta_TIO de este libro.
Public Function Fnc_Inf_Def_Hoja(ByVal Ws As Object, Def As T_Inf_Def) As Boolean
    Dim Code        As String
    On Error Resume Next
    If Ws.Parent Is ThisWorkbook Then Code = Ws.CodeName
    On Error GoTo 0
    If Code = "" Then Exit Function
    Select Case Code
        Case Sht__Inf_Recibos_TIO.CodeName:     Call Rut_Inf_Def_Recibos(Def)
        Case Sht__Inf_Conta_TIO.CodeName:       Call Rut_Inf_Def_Conta(Def)         '- M_140
        Case Else:                              Exit Function
    End Select
    Fnc_Inf_Def_Hoja = True
End Function    '- Fnc_Inf_Def_Hoja
' ==================================================================================================================================
'- Calcula el informe que describe Def (hasta el 2026-10-14, el cuerpo de Rut_Recalcular_Tabla_Inf_Recibos). Cada columna, por
'  el mapa de Def: C(InfRec_X).
Sub Rut_Inf_Calcular(Def As T_Inf_Def)
' ==================================================================================================================================
Debug.Print "Rut_Inf_Calcular: " & Def.Ws.Name

    Dim Paso                    As String:      Paso = "leer la configuración y las tablas"   '- para el aviso si algo falla
    Dim Msg_Error               As String
    Dim Off_Hecho               As Boolean
    Dim Prot_BD                 As T_Prot_Estado    '- protección de BDatos: Rut_Prot_Save al empezar, Rut_Prot_Restore en Fin
    Dim Prot_Hecho              As Boolean
    Dim C(1 To InfRec_N_Col)    As Integer      '- columna de cada dato en este informe: C(InfRec_X) (2026-10-14)
    Dim K_Col                   As Integer
    Dim INSS_Emi                As Double       '- el INSS de la fila, en variables: Inf_Conta_TIO no tiene sus columnas
    Dim INSS_Cob                As Double
    Dim Fila_Nueva              As Boolean      '- la fila del informe es nueva (un concepto repetido reescribe la misma)
    Dim Tot_INSS_Emi            As Double       '- Inf_Conta_TIO: total del INSS para la tabla de debajo (Rut_Inf_Tabla_INSS)
    Dim Tot_INSS_Cob            As Double
    Dim Tot_INSS_Pdte           As Double
    For K_Col = 1 To InfRec_N_Col
        C(K_Col) = Def.Col(K_Col)
    Next K_Col
    On Error GoTo Fin

    Dim ImpEmiCursAnt           As Currency
    Dim ImpEmiCursPos           As Currency
    Dim ImpINSSCursAnt          As Currency
    Dim ImpINSSCursPos          As Currency
    Dim ImpAdmCursAnt           As Currency
    Dim ImpAdmCursPos           As Currency
    Dim ImpAcadCursAnt          As Currency
    Dim ImpAcadCursPos          As Currency
    Dim ImpCobCursAnt           As Currency
    Dim ImpCobCursPos           As Currency
    Dim ImpCobCursINSSAnt       As Currency
    Dim ImpCobCursINSSPos       As Currency
    Dim ImpCobCursAdmAnt        As Currency
    Dim ImpCobCursAdmPos        As Currency
    Dim ImpCobCursAcadAnt       As Currency
    Dim ImpCobCursAcadPos       As Currency

Dim Concept         As Variant
Dim Concept2        As String
Dim Tp_Rec          As String
Dim TipRec_Cncpt    As String:          TipRec_Cncpt = ""
Dim DescripcionA    As String
Dim DescripcionB    As String
Dim MenuAux_Msg     As String
Dim APP_AnoCont     As String:          APP_AnoCont = Prog__APP.Range("App_AnoCont")
Dim APP_ACont       As String:          APP_ACont = Right(Prog__APP.Range("App_AnoCont"), 2)
Dim APP_AcadAnt     As String:          APP_AcadAnt = Prog__APP.Range("App_C_Acad_Ant")
Dim APP_AcadPos     As String:          APP_AcadPos = Prog__APP.Range("App_C_Acad_Pos")

'- Día siguiente al cierre, como número de serie: cuentan los recibos con fecha anterior. Hasta el 2026-10-12 iba como texto en el
'  criterio de SUMIFS, que VBA pasa en formato americano: un cierre del 30/06 daba "<01/07/aaaa" = 7 de enero.
Dim Lim_Cierre      As Long:            Lim_Cierre = CLng(Int(DateAdd("d", 1, Prog__APP.Range("App_FechCierreCont"))))
Dim Sumas()         As Double                                   '- (tipo de recibo, concepto, Sum_*): Rut_Inf_Recibos_Sumas
Dim JIs()           As String                                   '- (tipo de recibo, concepto, 1 To JI_N_Col): Rut_Inf_Recibos_Sumas
Dim K_JI            As Long
Dim F_Concept       As Integer
Dim F_Tp_Rec        As Integer

Dim Lo_BD           As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
Dim Lo_Concept      As ListObject:      Set Lo_Concept = Prog_Concept.ListObjects(1)       '- "Lo_Conceptos" (con "Tb_Conceptos", error 9 desde que se renombró la tabla)
Dim Lo_Tipo_Rec     As ListObject:      Set Lo_Tipo_Rec = Prog_TipoRec.ListObjects(1)   '- "Lo_TipoRec"

Dim Sht_Inf         As Worksheet:       Set Sht_Inf = Def.Ws
Dim Lo_Inf          As ListObject:      Set Lo_Inf = Sht_Inf.ListObjects(1)

Dim RwJI            As ListRow

'- 2 filas por debajo de la de totales, se vea o no (si está oculta, de donde estaría): Docs/Plan_ShowTotals.md, fase 2.
Dim UltFila     As Integer:     UltFila = Lo_Inf.Range.Row + Lo_Inf.Range.Rows.Count + IIf(Lo_Inf.ShowTotals, 1, 2)
Dim TRows_LoTipoRec      As Integer:     TRows_LoTipoRec = Lo_Tipo_Rec.ListRows.Count + 2

    Paso = "preparar las hojas"
    Call Rut_Off_Functions:     Off_Hecho = True    '- cálculo manual, pantalla y eventos; los restaura Rut_On_Functions en Fin
    Application.DisplayAlerts = False

    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    
    Sht__BD.Visible = xlSheetVisible
    Call Rut_Prot_Save(Sht__BD, Prot_BD):     Prot_Hecho = True     '- la desprotege; en Fin, protegida como estaba (antes se quedaba sin proteger)
    Prog_Concept.Visible = xlSheetHidden
    Prog_TipoRec.Visible = xlSheetHidden
    Sht_Inf.Visible = xlSheetVisible
    Sht_Inf.Select
    '- Borrado Tabla explicativa de Tipos de Recibos: Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. --------------------------
    Sht_Inf.Range(Cells(UltFila, 1), Cells(UltFila + TRows_LoTipoRec + 20, 1)).EntireRow.Delete
    
    '------------ Preparo las hojas y ordeno Lo_Conceptos -----------------------------------------------------------
'    Call Rut_Lo_Filtros_Quitar(Lo_Inf)
    Call Rut_Lo_WrkSht_Preparar(Sht_Inf)
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Call Rut_Lo_Sort(Lo_Concept, 1, xlAscending, True)          '- Ordeno por el número de Aplicación-Concepto 1303, 1310, 1311...
    '- BDatos ya no se ordena por Tipo_Rec y Concepto (2026-10-12; unos 7 s): las sumas no dependen del orden, y los módulos que
    '  necesitan un orden lo ponen ellos (limpiando antes los criterios). El orden del informe lo da la ordenación de Lo_Inf.
    Paso = "sumar los importes y juntar los JI's de BDatos"
    Call Rut_Inf_Recibos_Sumas(Lo_BD, Lo_Tipo_Rec, Lo_Concept, Lim_Cierre, APP_AcadPos, APP_AnoCont, Sumas, JIs)
    
    Sht_Inf.Select
    
    '- Vaciar Lo_Inf, Añadir Leyenda con AnoCont y Rellenar Lo_inf ----------------------------------------------
    Paso = "vaciar el informe y escribir las cabeceras"
    If Not Lo_Inf.DataBodyRange Is Nothing Then Lo_Inf.DataBodyRange.Delete

    '-------------- Relleno Tabla para JI's de Tasas por Tipo y por Concepto ------------------------------------
    With Lo_Inf.HeaderRowRange
    
    
        .Cells(C(InfRec_JI_Emi_Adm)) = "JI-" & APP_AnoCont - 1 & " o JI-" & APP_AnoCont & vbLf & "1303.00" & vbLf & "Adm."
        .Cells(C(InfRec_JI_Emi_Acad)) = "JI-" & APP_AnoCont - 1 & " o JI-" & APP_AnoCont & vbLf & "1310.xx 00/01/02" & vbLf & "Acad."
        .Cells(C(InfRec_AD_Emi_Adm)) = "AD-" & APP_AnoCont & vbLf & "1303.00" & vbLf & "Adm."
        .Cells(C(InfRec_AD_Emi_Acad)) = "AD-" & APP_AnoCont & vbLf & "1310.xx 00/01/02" & vbLf & "Acad."
        .Cells(C(InfRec_JI_443_Adm)) = "JI-" & APP_AnoCont & vbLf & "1303.00" & vbLf & "Adm."
        .Cells(C(InfRec_JI_443_Acad)) = "JI-" & APP_AnoCont & vbLf & "1310.xx 00/01/02" & vbLf & "Acad."
    
        .Cells(C(InfRec_ConcptEco2)).Offset(-2, 0) = "Año Contable " & APP_AnoCont & ", Cursos Académicos " & APP_AcadAnt & " y " & APP_AcadPos & ", (importes académicos)"
        .Cells(C(InfRec_ImpAcad_Crs_Ant)) = "Emi.'" & APP_ACont & vbLf & "Curso" & vbLf & APP_AcadAnt
        .Cells(C(InfRec_ImpAcad_Crs_Pos)) = "Emi.'" & APP_ACont & vbLf & "Curso" & vbLf & APP_AcadPos
        
        .Cells(C(InfRec_ImpAcad_EmiAnt)) = "Emi.'" & APP_ACont - 1 & vbLf & "No Cob.'" & APP_ACont - 1 & vbLf & "JI-" & APP_AnoCont - 1
        .Cells(C(InfRec_ImpAcad_Emi)) = "Emi.'" & APP_ACont & vbLf & "No Cob.'" & APP_ACont & vbLf & "Vto.'" & APP_ACont + 1
        .Cells(C(InfRec_ImpAcad_Cob_AcadAnt)) = "Cob.'" & APP_ACont & vbLf & "Curso" & vbLf & APP_AcadAnt
        .Cells(C(InfRec_ImpAcad_Cob_AcadPos)) = "Cob.'" & APP_ACont & vbLf & "Curso" & vbLf & APP_AcadPos
        .Cells(C(InfRec_ImpAcad_Pdte)) = vbLf & "Pdte. Cob."
        .Cells(C(InfRec_ADxAplz)) = "Emi. y No Cob.'" & APP_ACont & vbLf & "Vto.'" & APP_ACont + 1 & vbLf & "AD-" & APP_AnoCont
        .Cells(C(InfRec_Aplazado)) = "AD'" & APP_ACont - 1 & " Cob.'" & APP_ACont & vbLf & "JI-" & APP_AnoCont & vbLf & "A la 4430"
    End With
    
    '- Para cada Tipo de Recibos; Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. -----------------------------------
    Paso = "rellenar la tabla"
    For F_Tp_Rec = 1 To Lo_Tipo_Rec.ListRows.Count
        
        Tp_Rec = Lo_Tipo_Rec.DataBodyRange.Cells(F_Tp_Rec, 1)
        DescripcionA = Lo_Tipo_Rec.DataBodyRange.Cells(F_Tp_Rec, 2)
        
        '- Para cada Concepto Económico= Tipo de Enseñanza; 1310.00, 1310.01, 1311.00, 1303.00 .. (TIO, EFP, CFC/AF y ADM.) -----------------------------
        For F_Concept = 1 To Lo_Concept.ListRows.Count
            
            If Lo_Concept.DataBodyRange.Cells(F_Concept, 14) = 4 Then GoTo Sigiente_Concepto
            Concept = Lo_Concept.DataBodyRange.Cells(F_Concept, 1)
            Paso = "rellenar la tabla (" & Tp_Rec & ", " & Concept & ")"
            Fila_Nueva = (TipRec_Cncpt <> (Tp_Rec & Concept))     '- un concepto repetido reescribe la fila del anterior
            If Fila_Nueva Then
                Set RwJI = Lo_Inf.ListRows.Add
                TipRec_Cncpt = Tp_Rec & Concept
            End If
            Concept2 = Lo_Concept.DataBodyRange.Cells(F_Concept, 5)     '- 303.00, 303.01, 310.00, 310.01, 310.02, 311.00, 311.03, 312.00, 312.02, 315.00
            RwJI.Range(C(InfRec_TipRec)) = Tp_Rec                             '- ADxAplz, Añejo, EjeAnt, Emitido, Aplazado
            RwJI.Range(C(InfRec_Ensenanza)) = Lo_Concept.DataBodyRange.Cells(F_Concept, 4)    '- TIO, EFP, TNCT, CFC, ADM
            RwJI.Range(C(InfRec_ConcptEco)) = Lo_Concept.DataBodyRange.Cells(F_Concept, 5)
            RwJI.Range(C(InfRec_ConcptEco2)) = RwJI.Range(C(InfRec_ConcptEco))
            RwJI.Range(C(InfRec_Cta_Adm)) = Lo_Concept.DataBodyRange.Cells(F_Concept, 6)      '- Cta. Adm
            DescripcionB = Lo_Concept.DataBodyRange.Cells(F_Concept, 2)
            RwJI.Range(C(InfRec_ConcptNom)) = DescripcionB
            
            '-----------------------------------------------------------------------------------------------------------------------
            '----- Primeras 12 columnas de la Tabla; Imp_Rec, Imp_INSS, Imp_Adm e Imp_Acad -----------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- de cada Tipo de Recibos; Emitido, ADxAplz, Aplazado, EjeAnt y Añejo. ------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Recibos Emitidos: Acad. + Adm. ----------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe Recibos Emitidos --------------------------------------------------------------------------------------------
            RwJI.Range(C(InfRec_Tot_Emi)) = Sumas(F_Tp_Rec, F_Concept, Sum_Tot_Emi)
            '- Importe Recibos Emitidos Cobrado ------------------------------------------------------------------------------------
            RwJI.Range(C(InfRec_Tot_Cob)) = Sumas(F_Tp_Rec, F_Concept, Sum_Tot_Cob)
            '- Saldo Recibos Emitidos Pendiente ----------------------------------------------------------------------------------
            RwJI.Range(C(InfRec_Tot_Pdte)) = RwJI.Range(C(InfRec_Tot_Emi)) - RwJI.Range(C(InfRec_Tot_Cob))

            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Adm.: Seg.Obl. INSS ---------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            INSS_Emi = Sumas(F_Tp_Rec, F_Concept, Sum_INSS_Emi)            '- en variables (2026-10-14): Inf_Conta_TIO no tiene
            INSS_Cob = Sumas(F_Tp_Rec, F_Concept, Sum_INSS_Cob)            '  estas columnas, y Adm. y Acad. los restan igual
            If Def.Con_INSS Then
                '- Importe INSS -----------------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Adm_INSS_Emi)) = INSS_Emi
                '- Importe INSS. Cobrado --------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Adm_INSS_Cob)) = INSS_Cob
                '- Saldo INSS. Pendiente --------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Adm_INSS_Pdte)) = INSS_Emi - INSS_Cob
            End If
            If Not Def.Con_INSS And Fila_Nueva And Not UCase$(DescripcionB) Like "EURLE*" Then
                '- Inf_Conta_TIO (2026-10-14): total del INSS para la tabla de debajo, sin EURLE (va aparte) y con cada fila
                '  redondeada como en Inf_Recibos_TIO: es lo que le falta a la fila de totales para dar la de Inf_Recibos_TIO.
                Tot_INSS_Emi = Tot_INSS_Emi + Round(INSS_Emi, 2)
                Tot_INSS_Cob = Tot_INSS_Cob + Round(INSS_Cob, 2)
                Tot_INSS_Pdte = Tot_INSS_Pdte + Round(INSS_Emi - INSS_Cob, 2)
            End If

            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Adm. SIN Imp.INSS -----------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe Adm. SIN INSS -----------------------------------------------------------------------------------------------
            RwJI.Range(C(InfRec_Adm_Emi)) = Sumas(F_Tp_Rec, F_Concept, Sum_Adm_Emi)
            RwJI.Range(C(InfRec_Adm_Emi)) = RwJI.Range(C(InfRec_Adm_Emi)) - INSS_Emi
            '- Importe Adm. Cobrado ------------------------------------------------------------------------------------------------
            RwJI.Range(C(InfRec_Adm_Cob)) = Sumas(F_Tp_Rec, F_Concept, Sum_Adm_Cob)
            RwJI.Range(C(InfRec_Adm_Cob)) = RwJI.Range(C(InfRec_Adm_Cob)) - INSS_Cob
            '- Saldo Imp. Adm. SIN INSS --------------------------------------------------------------------------------------------
            RwJI.Range(C(InfRec_Adm_Pdte)) = RwJI.Range(C(InfRec_Adm_Emi)) - RwJI.Range(C(InfRec_Adm_Cob))
           
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Académicos SIN Adm. y SIN INSS ----------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            If RwJI.Range(C(InfRec_Ensenanza)) <> "ADM" And RwJI.Range(C(InfRec_Ensenanza)) <> "TNCT" Then
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Acad_Emi)) = RwJI.Range(C(InfRec_Tot_Emi)) - RwJI.Range(C(InfRec_Adm_Emi)) - INSS_Emi
                '- Imp. Acad. Cobrado ----------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Acad_Cob)) = RwJI.Range(C(InfRec_Tot_Cob)) - RwJI.Range(C(InfRec_Adm_Cob)) - INSS_Cob
                '- Saldo Imp. Acad. ------------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Acad_Pdte)) = RwJI.Range(C(InfRec_Acad_Emi)) - RwJI.Range(C(InfRec_Acad_Cob))
            Else
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Adm_Emi)) = RwJI.Range(C(InfRec_Tot_Emi))
                '- Imp. Acad. Cobrado ----------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Adm_Cob)) = RwJI.Range(C(InfRec_Tot_Cob))
                '- Saldo Imp. Acad. ------------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_Adm_Pdte)) = RwJI.Range(C(InfRec_Tot_Pdte))
                
'''                GoTo Sigiente_Concepto
                
            End If
            
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Académicos por Curso_Acad ---------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '- Importe Recibos Emitidos x Curso_Acad Ant / Pos ----------------------------------------------------------------------
            ImpEmiCursAnt = Sumas(F_Tp_Rec, F_Concept, Sum_Emi_Ant)
            ImpEmiCursPos = Sumas(F_Tp_Rec, F_Concept, Sum_Emi_Pos)
            '- Importe Adm. x Curso_Acad Ant / Pos -------------------------------------------------------------------------
            ImpAdmCursAnt = Sumas(F_Tp_Rec, F_Concept, Sum_Adm_Ant)
            ImpAdmCursPos = Sumas(F_Tp_Rec, F_Concept, Sum_Adm_Pos)
            ImpAcadCursAnt = ImpEmiCursAnt - ImpAdmCursAnt
            ImpAcadCursPos = ImpEmiCursPos - ImpAdmCursPos
            
            '- Importe Recibos Cobrados x Curso_Acad Ant / Pos ----------------------------------------------------------------------
            ImpCobCursAnt = Sumas(F_Tp_Rec, F_Concept, Sum_Cob_Ant)
            ImpCobCursPos = Sumas(F_Tp_Rec, F_Concept, Sum_Cob_Pos)
            '- Importe Adm. Cobrados  x Curso_Acad Ant / Pos -------------------------------------------------------------------------
            ImpCobCursAdmAnt = Sumas(F_Tp_Rec, F_Concept, Sum_CobAdm_Ant)
            ImpCobCursAdmPos = Sumas(F_Tp_Rec, F_Concept, Sum_CobAdm_Pos)
            ImpCobCursAcadAnt = ImpCobCursAnt - ImpCobCursAdmAnt
            ImpCobCursAcadPos = ImpCobCursPos - ImpCobCursAdmPos
            '-----------------------------------------------------------------------------------------------------------------------
            '-------------------------------- Importes Académicos SIN Adm. y SIN INSS ----------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            If RwJI.Range(C(InfRec_Ensenanza)) <> "ADM" And RwJI.Range(C(InfRec_Ensenanza)) <> "TNCT" Then
                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
                RwJI.Range(C(InfRec_ImpAcad_Crs_Ant)) = ImpAcadCursAnt
                RwJI.Range(C(InfRec_ImpAcad_Crs_Pos)) = ImpAcadCursPos
                RwJI.Range(C(InfRec_ImpAcad_Cob_AcadAnt)) = ImpCobCursAcadAnt
                RwJI.Range(C(InfRec_ImpAcad_Cob_AcadPos)) = ImpCobCursAcadPos
                RwJI.Range(C(InfRec_ImpAcad_Pdte)) = RwJI.Range(C(InfRec_Acad_Emi)) - ImpCobCursAcadAnt - ImpCobCursAcadPos
            Else
'                '- Imp. Acad. ------------------------------------------------------------------------------------------------------
'                RwJI.Range(InfRec_ImpAcad_Crs_Ant) = ImpEmiCursAnt
'                RwJI.Range(InfRec_ImpAcad_Crs_Pos) = ImpEmiCursPos
                
'''                GoTo Sigiente_Concepto
            
            End If
                
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '----- Resto columnas de la Tabla (Exclusivas de Inf_Peter) ------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            '-----------------------------------------------------------------------------------------------------------------------
            
            '- Importe Recibos Emitido ---------------------------------------------------------------------------------------------
            If RwJI.Range(C(InfRec_TipRec)) = "Emitido" Then

                '- Opto por copiar de las Columnas anteriores para reducir el tiempo de ejecución -----------------------------
                RwJI.Range(C(InfRec_ImpAcad_Emi)) = RwJI.Range(C(InfRec_Acad_Emi))
'                RwJI.Range(InfRec_ImpAcad_Cob_AcadAnt) = RwJI.Range(InfRec_Acad_Cob)
'                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            
            End If

            '- Importe Recibos EjeAnt --------------------------------------------------------------------------------------------------------------
            If RwJI.Range(C(InfRec_TipRec)) = "EjeAnt" Then
                RwJI.Range(C(InfRec_ImpAcad_EmiAnt)) = RwJI.Range(C(InfRec_Acad_Emi))
'                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            End If
            '- Importe Recibos Añejo --------------------------------------------------------------------------------------------------------------
            If RwJI.Range(C(InfRec_TipRec)) = "Añejo" Then
                RwJI.Range(C(InfRec_ImpAcad_Emi)) = RwJI.Range(C(InfRec_Acad_Emi))
'                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
            End If
            '- Importe Recibos Aplazado --------------------------------------------------------------------------------------------------------------
            If RwJI.Range(C(InfRec_TipRec)) = "Aplazado" Then
                RwJI.Range(C(InfRec_ImpAcad_Emi)) = RwJI.Range(C(InfRec_Acad_Emi))
'                RwJI.Range(InfRec_ImpAcad_Pdte) = RwJI.Range(InfRec_Acad_Emi) - ImpCobCursAnt - ImpCobCursPos
                RwJI.Range(C(InfRec_Aplazado)) = RwJI.Range(C(InfRec_Acad_Cob))
            End If
            '- Importe Recibos ADxAplz --------------------------------------------------------------------------------------------------------------
            If RwJI.Range(C(InfRec_TipRec)) = "ADxAplz" Then
                RwJI.Range(C(InfRec_ImpAcad_Emi)) = RwJI.Range(C(InfRec_Acad_Emi))
                RwJI.Range(C(InfRec_ADxAplz)) = RwJI.Range(C(InfRec_Acad_Emi))
            End If
                    

            

            Select Case RwJI.Range(C(InfRec_TipRec))
            Case "EjeAnt"
                RwJI.Range(C(InfRec_Descrip_Contab)) = "Añadir línea de Cobro en JI-" & APP_AnoCont - 1 & " del Importe Cobrado. Concepto Eco. " _
                                             & Concept2 & "__" & DescripcionB & ".  "
            Case "Emitido", "Añejo"
                RwJI.Range(C(InfRec_Descrip_Contab)) = "JI-" & APP_AnoCont & " del importe Emitido y línea de cobro del Importe Cobrado. Concepto Eco. " _
                                             & Concept2 & "__" & DescripcionB & ".  "
            
            Case "ADxAplz"
                RwJI.Range(C(InfRec_Descrip_Contab)) = "JI-" & APP_AnoCont & " del importe Emitido y AD-" & APP_AnoCont & ", del mismo Importe. Concepto Eco. " _
                                             & Concept2 & "__" & DescripcionB & ".  "
            
            Case "Aplazado"
                RwJI.Range(C(InfRec_Descrip_Contab)) = "JI-" & APP_AnoCont & " a la Cta. 4430, del importe Cobrado y línea de cobro del mismo Importe. Concepto Eco. " _
                                             & Concept2 & "__" & DescripcionB & ".  "
            End Select
            
            If RwJI.Range(C(InfRec_TipRec)) = "EjeAnt" Then
                RwJI.Range(C(InfRec_Descrip_JI)) = "Añadir línea de Cobro en JI-" & APP_AnoCont - 1 & _
                                             " del Importe Cobrado y de Concepto Eco. " & Concept2 & _
                                             "__" & RwJI.Range(C(InfRec_Ensenanza)) & " - " & DescripcionB & ".  "
            Else
                RwJI.Range(C(InfRec_Descrip_JI)) = "Liq.PPub_" & Concept2 & "__" & Tp_Rec & "_" & APP_AnoCont & _
                                             "__" & RwJI.Range(C(InfRec_Ensenanza)) & " - " & DescripcionB & ".  "
            End If
            
            
            '- JI's y AD's de los recibos de BDatos (2026-10-13; antes los ponía M_132): los distintos, separados por ", ";
            '  JI_Sin_Num si hay recibos y ninguno tiene; vacía si no hay recibos de ese tipo y concepto.
            For K_JI = 1 To JI_N_Col
                If Len(JIs(F_Tp_Rec, F_Concept, K_JI)) > 0 Then RwJI.Range(C(InfRec_JI_Emi_Adm) + K_JI - 1) = JIs(F_Tp_Rec, F_Concept, K_JI)
            Next K_JI
            
            If WorksheetFunction.IsEven(F_Tp_Rec) Then             '- IsEven = Es Par
                RwJI.Range.Interior.ColorIndex = 19
            Else
                RwJI.Range.Interior.ColorIndex = 20
            End If
'            RwJI.Range.RowHeight = 30
            RwJI.Range.EntireRow.AutoFit
            RwJI.Range.VerticalAlignment = xlCenter
            
Sigiente_Concepto:      '- AE 4: sin fila (2026-10-12; antes saltaba a las descripciones y repintaba la fila anterior)
        Next F_Concept
    Next F_Tp_Rec
    
    Paso = "redondear los importes"
    '- Redondear todos los valores a 2 decimales ----------------------
    Dim Lin         As Integer, Col     As Integer
    Dim RngCol      As Range, Celda     As Range
    ' Recorrer columnas de la tabla
    For Col = C(InfRec_Tot_Emi) To C(InfRec_Acad_Pdte)
        If Col = C(InfRec_Cta_Adm) Or Col = C(InfRec_ConcptEco) Then GoTo SiguienteColumna
        Set RngCol = Lo_Inf.ListColumns(Col).DataBodyRange
        For Each Celda In RngCol
            If IsNumeric(Celda.Value) Then
                Celda.Value = Round(Celda.Value, 2)
            End If
        Next Celda
SiguienteColumna:
    Next Col
    If Not Def.Con_INSS Then
        '- Inf_Conta_TIO (2026-10-14): Emitido, Cobrado y Pendiente = Adm. + Acad. de la fila, con los importes ya redondeados,
        '  para que cuadren al céntimo con la suma de las columnas. Antes de borrar las filas sin Emitido: una fila que solo
        '  tuviera INSS se queda a 0 y se borra.
        Paso = "sumar Adm. + Acad. en Emitido, Cobrado y Pendiente"
        For Each RwJI In Lo_Inf.ListRows
            RwJI.Range(C(InfRec_Tot_Emi)) = Round(RwJI.Range(C(InfRec_Adm_Emi)).Value + RwJI.Range(C(InfRec_Acad_Emi)).Value, 2)
            RwJI.Range(C(InfRec_Tot_Cob)) = Round(RwJI.Range(C(InfRec_Adm_Cob)).Value + RwJI.Range(C(InfRec_Acad_Cob)).Value, 2)
            RwJI.Range(C(InfRec_Tot_Pdte)) = Round(RwJI.Range(C(InfRec_Adm_Pdte)).Value + RwJI.Range(C(InfRec_Acad_Pdte)).Value, 2)
        Next RwJI
    End If
    
    Paso = "dar formato al informe y quitar las filas sin importe"
    '- Indento la columna de Descripción de JIs -------------
    Lo_Inf.ListColumns(C(InfRec_Descrip_JI)).DataBodyRange.Select
    Selection.InsertIndent 1
    Lo_Inf.ListColumns(C(InfRec_Descrip_Contab)).DataBodyRange.Select
    Selection.InsertIndent 1
    Lo_Inf.ListColumns(C(InfRec_ConcptNom)).DataBodyRange.Select
    Selection.InsertIndent 1
    '- Filtro para mostrar la filas con datos ---------------
'    Lo_Inf.Range.AutoFilter Field:=4, Criteria1:="<>0"
    Call Rut_Lo_DataBodyRange_Filter_y_DEL(Lo_Inf, C(InfRec_Tot_Emi), "-")
    Call Rut_Lo_Sort(Lo_Inf, C(InfRec_TipRec), xlAscending, True)
    Call Rut_Lo_Sort(Lo_Inf, C(InfRec_ConcptEco), xlAscending, False)
    Lo_Inf.ShowTotals = True
    


'''''            '-----------------------------------------------------------------------------------------------------------------------
'''''            '- JIs del AñoCont_Ant
'''''            If .Range(BD_JI_Emi_Acad) <> "" Then
'''''                If InStr(RwJI.Range(InfRec_Adm_Pdte), .Columns(BD_JI_Emi_Acad)) Then
'''''                    RwJI.Range(InfRec_Adm_Pdte) = RwJI.Range(InfRec_Adm_Pdte) & "-" & .Columns(BD_JI_Emi_Acad)
'''''                End If
'''''            End If
            
            

'GoTo Finalizar
    
    '- ------------------------------------------------------
    '- Copio la Tabla explicativa de los Tipos de Recibos: Prog_TipoRec.ListObjects(1) ---
    Paso = "copiar la leyenda de los tipos de recibo"
    Dim ColIni      As Integer:     ColIni = C(InfRec_Descrip_JI)
    Dim ColFin      As Integer:     ColFin = ColIni + 1
    Dim Fila_INSS   As Long                     '- Inf_Conta_TIO: primera fila de la tabla del total del INSS
    UltFila = Lo_Inf.TotalsRowRange.Row + 3
    TRows_LoTipoRec = Prog_TipoRec.ListObjects(1).ListRows.Count
'    Prog_TipoRec.ListObjects(1).ListColumns(1).Range.Resize(, 2).Copy Destination:=Cells(UltFila , 3) ' Resize to include the second column
    Prog_TipoRec.ListObjects(1).ListColumns(1).Range.Resize(TRows_LoTipoRec + 1, 2).Copy    '- cabecera + datos, 2 Col. (sin la fila de totales aunque se vea)
    Cells(UltFila, ColIni).PasteSpecial Paste:=xlPasteFormats
    Cells(UltFila, ColIni).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    '- Añado Datos de AnoCont y Cursos Acad -----------------
    Cells(UltFila, ColIni + 1) = "Descripción para Recibos del Año Contable " & APP_AnoCont & vbLf & _
                            " Rec. que abarcan el Curso Académico " & APP_AcadAnt & " y " & APP_AcadPos & vbLf & _
                            " Y Son Rec. Emitidos en " & APP_AnoCont & " (Cobrados o NO)," & vbLf & _
                            " O Rec. Cobrados en " & APP_AnoCont & " (Emitidos en cualquier Año =< " & APP_AnoCont & ")"
    '- Cambio el alto de las filas y centro verticalmente ---
    Range(Cells(UltFila, ColIni), Cells(UltFila + TRows_LoTipoRec, ColIni + 1)).Select
            Selection.RowHeight = 80
            Selection.VerticalAlignment = xlCenter
    '- Combino Celdas para la Descripción -------------------
    Range(Cells(UltFila, ColIni + 1), Cells(UltFila + TRows_LoTipoRec, ColFin)).Select
    Selection.Merge True
    '- Bordes a toda la Tabla -------------------------------
    Range(Cells(UltFila, ColIni), Cells(UltFila + TRows_LoTipoRec, ColFin)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Union(Cells(UltFila, ColIni + 1), Range(Cells(UltFila + 1, ColIni), Cells(UltFila + TRows_LoTipoRec, ColFin))).Select
    With Selection.Font
        .Name = "Arial"
        .Size = 12
    End With
    
    '- Filas EURLE (Esc.Univ. Rel.Lab. Elda), las filtro, las copio al final del Informe y las borro de Lo_Inf.
    Paso = "separar las filas de EURLE"
    Dim RowsFind    As Variant
    Dim Rng         As Range
    With Lo_Inf
        .Range.AutoFilter Field:=C(InfRec_ConcptNom), Criteria1:="EURLE*"       '- Filtrar
        Sht_Inf.Calculate
        RowsFind = Fnc_Lo_Contar_Visibles(Lo_Inf, C(InfRec_ConcptNom))
        If RowsFind > 0 Then
            Set Rng = Range(Cells(2, Lo_Inf.Range.Column + C(InfRec_Tot_Emi) - 1), Cells(2, Lo_Inf.Range.Column + C(InfRec_Acad_Pdte) - 1))  '- E2:R2 / E2:O2
            Rng.Copy Destination:=Cells(UltFila + TRows_LoTipoRec + 3, 5)
            Set Rng = .Range.Resize(.Range.Rows.Count, C(InfRec_Acad_Pdte)).SpecialCells(xlCellTypeVisible)
            Rng.Copy Destination:=Cells(UltFila + TRows_LoTipoRec + 4, 2)
            .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        End If
    End With
    
    '- Inf_Conta_TIO (2026-10-14): tabla con el total del INSS, bajo Concepto NOM, Emitido, Cobrado y Pendiente, 2 filas por debajo
    '  de lo último que haya bajo la tabla: el bloque de EURLE (columna B) o, si no hay, la leyenda (columna ColIni). Como la
    '  leyenda, la dibuja cada cálculo.
    If Not Def.Con_INSS Then
        Paso = "escribir la tabla del total del INSS"
        Fila_INSS = WorksheetFunction.Max(Sht_Inf.Cells(Sht_Inf.Rows.Count, Lo_Inf.Range.Column).End(xlUp).Row, _
                                          Sht_Inf.Cells(Sht_Inf.Rows.Count, ColIni).End(xlUp).Row) + 2
        Call Rut_Inf_Tabla_INSS(Def, Lo_Inf, Fila_INSS, APP_AnoCont, Tot_INSS_Emi, Tot_INSS_Cob, Tot_INSS_Pdte)
    End If
    
Finalizar:
    Paso = "anotar la fecha del cálculo"
    Call Rut_Lo_Filtros_Quitar(Lo_Inf)
    Sht_Inf.Calculate
    Sht_Inf.Range("d2") = "Último Cálculo: " & Format(Now, "dd mmmm yyyy - hh:mm") & vbLf & "Fecha de cierre contable: " & Prog__APP.Range("App_FechCierreCont") & _
                          vbLf & "Descarga de los Recibos del " & Prog__APP.Range("App_Last_Import_Robot_PPub")   '- fecha de la carpeta del Robot (M_110), 2026-10-13
    Sht_Inf.Range("b2") = Def.Titulo & " " & APP_AnoCont       '- titulo de B2:C2 (combinadas: se escribe en la ancla, B2)
    Sht_Inf.Range("d2").Select
    Prog__APP.Range(Def.Nom_Calc) = Format(Now(), "dd-mmm-yy hh:mm")
    
''    '- Visualizo el progreso ---------------------------------------------------------------------------------------
''    MenuAux_Msg = Format(Now, "hh:mm:ss") & "  Tabla generada." & vbCrLf & _
''        vbCrLf & Format(Now, "hh:mm:ss") & "  Realizado el: " & Date & "  " & "-   Tiempo transcurrido: " & Round(Timer - H_Inicio, 2) & " seg."
''    MsgBox MenuAux_Msg
    Lo_BD.ShowTotals = True

Fin:
    If Err.Number <> 0 Then
        Msg_Error = "Error " & Err.Number & " al " & Paso & ":" & vbLf & Err.Description
        Debug.Print ">>> Rut_Inf_Calcular -> ERROR " & Err.Number & " al " & Paso & ": " & Err.Description
    End If
    On Error GoTo -1                                '- sale del manejador de errores, para que el Resume Next de abajo valga
    On Error Resume Next
    If Msg_Error <> "" Then                         '- salida por error: el informe sin filtros y las filas de totales visibles
        Application.CutCopyMode = False
        Call Rut_Lo_Filtros_Quitar(Lo_Inf)
        Call Rut_Lo_Totales_Mostrar                 '- las celdas de encima de BDatos leen su fila de totales
    End If
    If Prot_Hecho Then Call Rut_Prot_Restore(Sht__BD, Prot_BD, True)    '- después de Rut_Lo_Totales_Mostrar, que necesita BDatos sin proteger
    Application.DisplayAlerts = True
    On Error GoTo 0
    If Off_Hecho Then Call Rut_On_Functions         '- antes, Application.Calculation = Sw_Calculation (Boolean): se ignoraba y quedaba en Manual
    If Msg_Error <> "" Then                         '- el aviso, ya con la pantalla activa
        MsgBox Msg_Error & vbLf & vbLf & "El informe puede haber quedado incompleto: vuelve a calcularlo.", _
               vbExclamation, "Calcular " & Def.Nom
    Else
        Application.Speech.Speak "Proceso completado", True
    End If
End Sub     '- Rut_Inf_Calcular
' ==================================================================================================================================
'- Inf_Conta_TIO (2026-10-14): tabla con el total del INSS (emitido, cobrado y pendiente), que este informe no tiene en sus columnas.
'  Dos filas desde Fila (cabeceras e importes), bajo las columnas Concepto NOM, Emitido, Cobrado y Pendiente de la tabla, con el
'  formato de la que dibujó el usuario a mano: Arial, oro (Énfasis 4) aclarado un 60 % y un 80 % y bordes finos. Los importes,
'  con el formato de número de la fila de totales.
' ==================================================================================================================================
Private Sub Rut_Inf_Tabla_INSS(Def As T_Inf_Def, Lo_Inf As ListObject, ByVal Fila As Long, ByVal AnoCont As String, _
                               ByVal Emi As Double, ByVal Cob As Double, ByVal Pdte As Double)
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim Ws          As Worksheet:   Set Ws = Lo_Inf.Parent
    Dim Col0        As Long:        Col0 = Lo_Inf.Range.Column - 1      '- columna de la hoja = Col0 + columna de la tabla
    Dim C_Nom       As Long:        C_Nom = Col0 + Def.Col(InfRec_ConcptNom)
    Dim C_Emi       As Long:        C_Emi = Col0 + Def.Col(InfRec_Tot_Emi)
    Dim C_Cob       As Long:        C_Cob = Col0 + Def.Col(InfRec_Tot_Cob)
    Dim C_Pdte      As Long:        C_Pdte = Col0 + Def.Col(InfRec_Tot_Pdte)
    Dim Tabla       As Range
    Dim Cabeceras   As Range                    '- INSS Emi. e INSS Pdte (la de INSS Cob. va más grande y más clara)
    Dim Importes    As Range                    '- los de Emi. y Pdte (el de Cob. va más claro)

    Set Tabla = Union(Ws.Cells(Fila, C_Nom).Resize(2), Ws.Cells(Fila, C_Emi).Resize(2), _
                      Ws.Cells(Fila, C_Cob).Resize(2), Ws.Cells(Fila, C_Pdte).Resize(2))
    Set Cabeceras = Union(Ws.Cells(Fila, C_Emi), Ws.Cells(Fila, C_Pdte))
    Set Importes = Union(Ws.Cells(Fila + 1, C_Emi), Ws.Cells(Fila + 1, C_Pdte))

    Ws.Cells(Fila, C_Nom) = "Totales"
    Ws.Cells(Fila, C_Emi) = "INSS Emi."
    Ws.Cells(Fila, C_Cob) = "INSS Cob." & vbLf & AnoCont
    Ws.Cells(Fila, C_Pdte) = "INSS Pdte"
    Ws.Cells(Fila + 1, C_Nom) = "Imp.Adm. Seguro Obl. INSS"
    Ws.Cells(Fila + 1, C_Emi) = Round(Emi, 2)
    Ws.Cells(Fila + 1, C_Cob) = Round(Cob, 2)
    Ws.Cells(Fila + 1, C_Pdte) = Round(Pdte, 2)

    '- Toda la tabla: Arial negra (Texto 1), centrada en vertical, rellena de oro aclarado un 80 % y con bordes finos ---------------
    With Tabla
        .Font.Name = "Arial"
        .Font.Bold = False
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .IndentLevel = 0
        .Interior.Pattern = xlSolid
        .Interior.ThemeColor = xlThemeColorAccent4
        .Interior.TintAndShade = 0.8
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = xlAutomatic
    End With
    Union(Cabeceras, Importes).Interior.TintAndShade = 0.6     '- las de Emi. y Pdte, oro aclarado un 60 %
    With Ws.Cells(Fila, C_Nom)                                  '- "Totales": negrita y sin relleno
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .Interior.Pattern = xlNone
    End With
    With Union(Cabeceras, Ws.Cells(Fila, C_Cob))               '- cabeceras del INSS
        .Font.Bold = True
        .Font.TintAndShade = 0.05
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    Cabeceras.Font.Size = 12
    Ws.Cells(Fila, C_Cob).Font.Size = 14
    With Ws.Cells(Fila + 1, C_Nom)                              '- "Imp.Adm. Seguro Obl. INSS"
        .Font.Size = 18
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
    End With
    With Union(Importes, Ws.Cells(Fila + 1, C_Cob))            '- los totales
        .Font.Size = 16
        .NumberFormat = Lo_Inf.TotalsRowRange.Cells(1, Def.Col(InfRec_Tot_Emi)).NumberFormat
    End With
    Ws.Rows(Fila).Resize(2).AutoFit                             '- alto automático (36 y 22,5 el 2026-10-14)
End Sub     '- Rut_Inf_Tabla_INSS
' ==================================================================================================================================
'- Sumas de BDatos para el informe, en una sola pasada en memoria (2026-10-12). Antes eran 14 SUMIFS por fila del informe (420 en
'  total, unos 20 s); dan los mismos importes (comprobado con los 420). Sumas(tipo de recibo, concepto, Sum_*), con los criterios
'  de aquellos SUMIFS:
'   - Tipo_Rec: texto que CONTIENE el tipo de Lo_TipoRec (el "*Emitido*" de antes), sin distinguir mayúsculas.
'   - Concepto: igual al de Lo_Conceptos. Si este es un número escrito con punto ("1310.01"), igual a ese número: así leía el
'     criterio SUMIFS llamado desde VBA.
'   - Emitidos: Imp.Rec. > 0 y F_Emi anterior a Lim. Cobrados: Imp.Cob. > 0 y F_Cob anterior a Lim. Lim = día siguiente al cierre.
'   - Por curso: C_Acad igual al curso Pos, o distinto (vacías incluidas) para el Ant. Cobrado en el año contable: ACont_Cob igual
'     al año y F_Emi anterior a Lim (sin mirar F_Cob ni que Imp.Cob. sea > 0, como antes).
'   - Solo suman los números; textos, vacías y errores no cuentan, como en SUMIFS.
'  En la misma pasada, JIs(tipo de recibo, concepto, 1 To JI_N_Col) junta los JI's de BDatos (2026-10-13; antes M_132 filtraba
'  BDatos una vez por fila del informe, con «=» en Tipo_Rec y en Concepto), con sus mismas reglas:
'   - Tipo_Rec IGUAL al tipo (sin distinguir mayúsculas), no que lo contenga; sin mirar fechas ni importes.
'   - Los valores distintos de cada columna (.Value, como Celda.Value), en el orden de BDatos, separados por ", ".
'   - JI_Sin_Num si hay recibos y ninguno tiene dato; vacío si no hay recibos. No cuentan como dato los vacíos, los
'     blancos, los errores ni JI_Sin_Num (Fnc_JI_Con_Dato).
' ==================================================================================================================================
Private Sub Rut_Inf_Recibos_Sumas(Lo_BD As ListObject, Lo_Tipo_Rec As ListObject, Lo_Concept As ListObject, _
                                  ByVal Lim As Double, ByVal C_Acad_Pos As String, ByVal AnoCont As String, _
                                  Sumas() As Double, JIs() As String)
' ----------------------------------------------------------------------------------------------------------------------------------
    Dim T           As T_TablaRam
    Dim N_Tp        As Long:        N_Tp = Lo_Tipo_Rec.ListRows.Count
    Dim N_Cp        As Long:        N_Cp = Lo_Concept.ListRows.Count
    Dim Tipos()     As String
    Dim Alias_Cp()  As Long                         '- un concepto repetido en Lo_Conceptos suma lo mismo que el primero
    Dim Dic_Num     As Object:      Set Dic_Num = CreateObject("Scripting.Dictionary")
    Dim Dic_Txt     As Object:      Set Dic_Txt = CreateObject("Scripting.Dictionary")
    Dim Pos_Min     As String:      Pos_Min = LCase$(C_Acad_Pos)
    Dim ACont_Num   As Double:      ACont_Num = Val(AnoCont)
    Dim Fila        As Long, Tp As Long, Cp As Long, K As Long
    Dim V           As Variant, Tip As Variant
    Dim FEmi_OK     As Boolean, Emi_OK As Boolean, Cob_OK As Boolean, ACont_OK As Boolean
    Dim Es_Pos      As Boolean, Es_Ant As Boolean
    Dim Imp_Rec     As Double, Imp_Cob As Double, Imp_INSS As Double, Imp_Adm As Double
    Dim Datos_JI    As Variant                      '- (fila, 1 To JI_N_Col): las columnas de JI's de BDatos
    Dim Dic_JI()    As Object                       '- (tipo, concepto, columna): JI's distintos
    Dim N_Rec()     As Long                         '- (tipo, concepto): recibos con ese tipo IGUAL y ese concepto
    Dim V_JI        As Variant
    Dim Dic         As Object

    If N_Tp = 0 Or N_Cp = 0 Then Exit Sub
    ReDim Sumas(1 To N_Tp, 1 To N_Cp, 1 To Sum_N)
    ReDim JIs(1 To N_Tp, 1 To N_Cp, 1 To JI_N_Col)
    ReDim Dic_JI(1 To N_Tp, 1 To N_Cp, 1 To JI_N_Col)
    ReDim N_Rec(1 To N_Tp, 1 To N_Cp)
    ReDim Tipos(1 To N_Tp)
    ReDim Alias_Cp(1 To N_Cp)
    For Tp = 1 To N_Tp
        Tipos(Tp) = CStr(Lo_Tipo_Rec.DataBodyRange.Cells(Tp, 1).Value2)
    Next Tp
    For Cp = 1 To N_Cp                              '- concepto -> nº de fila en Lo_Conceptos
        Alias_Cp(Cp) = Cp
        V = Lo_Concept.DataBodyRange.Cells(Cp, 1).Value2
        If VarType(V) = vbString Then
            If Fnc_Es_Numero_Punto(CStr(V)) Then V = Val(CStr(V))
        End If
        If VarType(V) = vbDouble Then
            If Dic_Num.Exists(V) Then Alias_Cp(Cp) = Dic_Num(V) Else Dic_Num.Add V, Cp
        Else
            V = LCase$(CStr(V))
            If Dic_Txt.Exists(V) Then Alias_Cp(Cp) = Dic_Txt(V) Else Dic_Txt.Add V, Cp
        End If
    Next Cp

    Call Rut_TablaRam_Cargar(T, Lo_BD, Array(BD_ACont_Cob, BD_C_Acad, BD_FEmi, BD_FCob, BD_ImpRec, BD_ImpCob, _
                                             BD_Rec_Imp_Adm, BD_Rec_Imp_INSS, BD_Concepto, BD_Tipo_Rec), True)
    If T.NumFilas > 0 Then Datos_JI = Lo_BD.ListColumns(BD_JI_Emi_Adm).DataBodyRange.Resize(, JI_N_Col).Value
    For Fila = 1 To T.NumFilas
        Cp = 0
        V = T.Datos(Fila, BD_Concepto)
        If VarType(V) = vbDouble Then
            If Dic_Num.Exists(V) Then Cp = Dic_Num(V)
        ElseIf VarType(V) = vbString Then
            V = LCase$(V)
            If Dic_Txt.Exists(V) Then Cp = Dic_Txt(V)
        End If
        Tip = T.Datos(Fila, BD_Tipo_Rec)
        If Cp > 0 And VarType(Tip) = vbString Then
            V = T.Datos(Fila, BD_ImpRec):           Imp_Rec = 0:    If VarType(V) = vbDouble Then Imp_Rec = V
            V = T.Datos(Fila, BD_ImpCob):           Imp_Cob = 0:    If VarType(V) = vbDouble Then Imp_Cob = V
            V = T.Datos(Fila, BD_Rec_Imp_INSS):     Imp_INSS = 0:   If VarType(V) = vbDouble Then Imp_INSS = V
            V = T.Datos(Fila, BD_Rec_Imp_Adm):      Imp_Adm = 0:    If VarType(V) = vbDouble Then Imp_Adm = V
            V = T.Datos(Fila, BD_FEmi):             FEmi_OK = False: If VarType(V) = vbDouble Then FEmi_OK = (V < Lim)
            Emi_OK = FEmi_OK And (Imp_Rec > 0)
            V = T.Datos(Fila, BD_FCob):             Cob_OK = False: If VarType(V) = vbDouble Then Cob_OK = (V < Lim)
            Cob_OK = Cob_OK And (Imp_Cob > 0)
            V = T.Datos(Fila, BD_ACont_Cob)
            Select Case VarType(V)
                Case vbDouble:  ACont_OK = (V = ACont_Num)
                Case vbString:  ACont_OK = (Trim$(V) = AnoCont)
                Case Else:      ACont_OK = False
            End Select
            V = T.Datos(Fila, BD_C_Acad)
            Select Case VarType(V)
                Case vbString:  Es_Pos = (LCase$(V) = Pos_Min):     Es_Ant = Not Es_Pos
                Case vbError:   Es_Pos = False:                     Es_Ant = False
                Case Else:      Es_Pos = False:                     Es_Ant = True       '- vacías y números: distintos del Pos
            End Select
            For Tp = 1 To N_Tp
                If InStr(1, Tip, Tipos(Tp), vbTextCompare) > 0 Then
                    If Emi_OK Then
                        Sumas(Tp, Cp, Sum_Tot_Emi) = Sumas(Tp, Cp, Sum_Tot_Emi) + Imp_Rec
                        Sumas(Tp, Cp, Sum_INSS_Emi) = Sumas(Tp, Cp, Sum_INSS_Emi) + Imp_INSS
                        Sumas(Tp, Cp, Sum_Adm_Emi) = Sumas(Tp, Cp, Sum_Adm_Emi) + Imp_Adm
                        If Es_Pos Then
                            Sumas(Tp, Cp, Sum_Emi_Pos) = Sumas(Tp, Cp, Sum_Emi_Pos) + Imp_Rec
                            Sumas(Tp, Cp, Sum_Adm_Pos) = Sumas(Tp, Cp, Sum_Adm_Pos) + Imp_Adm
                        ElseIf Es_Ant Then
                            Sumas(Tp, Cp, Sum_Emi_Ant) = Sumas(Tp, Cp, Sum_Emi_Ant) + Imp_Rec
                            Sumas(Tp, Cp, Sum_Adm_Ant) = Sumas(Tp, Cp, Sum_Adm_Ant) + Imp_Adm
                        End If
                    End If
                    If Cob_OK Then
                        Sumas(Tp, Cp, Sum_Tot_Cob) = Sumas(Tp, Cp, Sum_Tot_Cob) + Imp_Cob
                        Sumas(Tp, Cp, Sum_INSS_Cob) = Sumas(Tp, Cp, Sum_INSS_Cob) + Imp_INSS
                        Sumas(Tp, Cp, Sum_Adm_Cob) = Sumas(Tp, Cp, Sum_Adm_Cob) + Imp_Adm
                    End If
                    If ACont_OK And FEmi_OK Then
                        If Es_Pos Then
                            Sumas(Tp, Cp, Sum_Cob_Pos) = Sumas(Tp, Cp, Sum_Cob_Pos) + Imp_Cob
                            Sumas(Tp, Cp, Sum_CobAdm_Pos) = Sumas(Tp, Cp, Sum_CobAdm_Pos) + Imp_Adm
                        ElseIf Es_Ant Then
                            Sumas(Tp, Cp, Sum_Cob_Ant) = Sumas(Tp, Cp, Sum_Cob_Ant) + Imp_Cob
                            Sumas(Tp, Cp, Sum_CobAdm_Ant) = Sumas(Tp, Cp, Sum_CobAdm_Ant) + Imp_Adm
                        End If
                    End If
                End If
                If StrComp(Tip, Tipos(Tp), vbTextCompare) = 0 Then      '- JI's: tipo IGUAL, sin mirar fechas ni importes
                    N_Rec(Tp, Cp) = N_Rec(Tp, Cp) + 1
                    For K = 1 To JI_N_Col
                        V_JI = Datos_JI(Fila, K)
                        If Fnc_JI_Con_Dato(V_JI) Then
                            If Dic_JI(Tp, Cp, K) Is Nothing Then Set Dic_JI(Tp, Cp, K) = CreateObject("Scripting.Dictionary")
                            Set Dic = Dic_JI(Tp, Cp, K)
                            Dic(V_JI) = True                         '- valores distintos, en el orden de BDatos
                        End If
                    Next K
                End If
            Next Tp
        End If
    Next Fila

    For Tp = 1 To N_Tp                              '- JI's: los distintos, separados por ", "; JI_Sin_Num si hay recibos sin ninguno
        For Cp = 1 To N_Cp
            For K = 1 To JI_N_Col
                If Not Dic_JI(Tp, Cp, K) Is Nothing Then
                    JIs(Tp, Cp, K) = Join(Dic_JI(Tp, Cp, K).Keys, ", ")
                ElseIf N_Rec(Tp, Cp) > 0 Then
                    JIs(Tp, Cp, K) = JI_Sin_Num
                End If
            Next K
        Next Cp
    Next Tp

    For Cp = 1 To N_Cp                              '- conceptos repetidos: las sumas y los JI's del primero
        If Alias_Cp(Cp) <> Cp Then
            For Tp = 1 To N_Tp
                For K = 1 To Sum_N
                    Sumas(Tp, Cp, K) = Sumas(Tp, Alias_Cp(Cp), K)
                Next K
                For K = 1 To JI_N_Col
                    JIs(Tp, Cp, K) = JIs(Tp, Alias_Cp(Cp), K)
                Next K
            Next Tp
        End If
    Next Cp
End Sub     '- Rut_Inf_Recibos_Sumas
' ----------------------------------------------------------------------------------------------------------------------------------
'- Un valor de una columna de JI's cuenta como dato si no está vacío ni en blanco, no es un error y no es JI_Sin_Num
'  (2026-10-13; lo usan Rut_Inf_Recibos_Sumas y M_133).
Public Function Fnc_JI_Con_Dato(ByVal V As Variant) As Boolean
    If IsEmpty(V) Or IsError(V) Then Exit Function
    If VarType(V) = vbString Then
        If Trim$(V) = "" Or V = JI_Sin_Num Then Exit Function
    End If
    Fnc_JI_Con_Dato = True
End Function
' ----------------------------------------------------------------------------------------------------------------------------------
'- "1310.01" o "1303": solo cifras y como mucho un punto (un número como lo lee SUMIFS llamado desde VBA)
Private Function Fnc_Es_Numero_Punto(ByVal Txt As String) As Boolean
    Txt = Trim$(Txt)
    If Txt = "" Or Txt = "." Then Exit Function
    If Txt Like "*[!0-9.]*" Then Exit Function
    If Len(Txt) - Len(Replace(Txt, ".", "")) > 1 Then Exit Function
    Fnc_Es_Numero_Punto = True
End Function
' ==================================================================================================================================
'Sub Ejemplo_Selcción_Múltiple()
'    Union(Cells(44 + 4, 4), Range(Cells(44 + 5, 3), Cells(44 + 3 + 5, 11))).Select
'End Sub
