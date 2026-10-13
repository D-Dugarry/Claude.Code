Attribute VB_Name = "M_135_Mail_Inf_Recibos"
' Last Rev. 2026-10-13 21:03
Option Explicit

'===================================================================================================
' M_135_Mail_Inf_Recibos
'
' Envía por correo el Informe de Recibos (hoja Inf_Recibos_TIO, CodeName Sht__Inf_Recibos_TIO) como
' un .xlsx con sus columnas B:Z. Plan y decisiones: Docs/Plan_Mail_Inf_Recibos.md.
'
' El adjunto se hace COPIANDO LA HOJA ENTERA a un libro nuevo y recortándola, no pegando un rango
' (el PasteSpecial de JyC M_825): así viajan la tabla con su fila de totales, los altos de fila, las
' celdas combinadas y las columnas que estén ocultas en la hoja. La hoja origen no se toca.
'
' Subs y funciones públicas:
'   Fnc_Inf_Recibos_A_Xlsx([ErrMsg])      - Genera el .xlsx en %TEMP% y devuelve su ruta ("" si falla)
'   Rut_Mail_Inf_Recibos                  - Botón «Enviar por correo»: pregunta, genera el .xlsx, lo
'                                           copia en la carpeta de red (App_Ruta_Inf_Recibos\<año>) si se
'                                           confirma, y lo envía (lo lanza OnAct_Mail_Inf_Recibos, M___RibbonUI)
'   Rut_Inf_Recibos_Probar_Xlsx           - Prueba: genera el .xlsx sin enviarlo y lo señala en el
'                                           Explorador (Inmediato: Call Rut_Inf_Recibos_Probar_Xlsx)
'===================================================================================================

'===================================================================================================
'- Genera el .xlsx del Informe de Recibos (columnas B:Z) en %TEMP% y devuelve su ruta --------------
'-   Nombre: "Inf_Recibos_TIO (datos del AAAA-MM-DD).xlsx", con la fecha de la carpeta de descarga del
'-   Robot (App_Last_Import_Robot_PPub, la anota M_110 al importar); sin esa fecha, "Inf_Recibos_TIO.xlsx".
'-   Con la hora detrás si hay un libro abierto con ese nombre o si el fichero anterior no se puede borrar.
'-   Si algo falla, devuelve "" y la causa en ErrMsg. No muestra ningún MsgBox: se lanzará con
'-   Form_Running_Rut abierto, y un MsgBox se abriría detrás del formulario.
'-   En la copia: todas las filas y columnas visibles y sin filtros; sin las columnas a la derecha de
'-   InfRec_Descrip_Contab (Z); sin las "x" de la fila 1 (las del botón Hide Cols_Row1) y con esa fila de
'-   alto normal; con los altos de las demás filas de la hoja; al 100 % de zoom; sin las formas con macro,
'-   y sin vínculos ni nombres que no sean de la copia.
'-   Se guarda con el cálculo en Automático: un libro guarda
'-   el modo de cálculo de Excel, y si el destinatario lo abriera el primero, su Excel quedaría en Manual.
Function Fnc_Inf_Recibos_A_Xlsx(Optional ByRef ErrMsg As String = "") As String
    Dim Ws          As Worksheet
    Dim Lo          As ListObject
    Dim WsAct       As Object               ' hoja activa al empezar, para volver a ella
    Dim VisAnt      As XlSheetVisibility
    Dim SwDeactAnt  As Variant
    Dim CalcAnt     As XlCalculation
    Dim WbN         As Workbook             ' el libro nuevo, con la copia
    Dim WsN         As Worksheet
    Dim LoN         As ListObject
    Dim Fila1       As Range
    Dim Celda       As Range
    Dim Vinculos    As Variant
    Dim i           As Long
    Dim Ruta        As String
    Dim Nombre      As String
    Dim Paso        As String
    Dim Off_Hecho   As Boolean
    Dim Sw_Hecho    As Boolean
    Dim Vis_Hecho   As Boolean
    Dim Calc_Hecho  As Boolean

    ErrMsg = ""
    On Error GoTo Err_Xlsx

    Paso = "comprobar el informe"
    Set Ws = Sht__Inf_Recibos_TIO
    Set Lo = Ws.ListObjects(1)
    If Lo.DataBodyRange Is Nothing Then
        ErrMsg = "El Informe de Recibos está vacío: hay que calcularlo antes de enviarlo."
        GoTo Salir
    End If
    Nombre = Fnc_Inf_Recibos_Nombre()
    Ruta = Fnc_Inf_Recibos_Ruta_Libre(Environ$("TEMP") & "\", Nombre)

    Paso = "copiar la hoja a un libro nuevo"
    If ActiveWorkbook Is ThisWorkbook Then Set WsAct = ActiveSheet
    '- Eventos apagados: la copia lleva el código de la hoja, que llama a rutinas de PPub (en el libro nuevo
    '- no existen), y el Workbook_Deactivate de PPub restablecería Excel a mitad (Sw_WB_Deactivate, por si acaso).
    Call Rut_Off_Functions:     Off_Hecho = True
    SwDeactAnt = Prog__APP_Switch.Range("Sw_WB_Deactivate").Value
    Prog__APP_Switch.Range("Sw_WB_Deactivate").Value = False:   Sw_Hecho = True
    VisAnt = Ws.Visible                     '- una hoja oculta no se puede copiar sola a un libro nuevo
    If VisAnt <> xlSheetVisible Then Ws.Visible = xlSheetVisible:   Vis_Hecho = True
    Ws.Copy
    Set WbN = ActiveWorkbook
    Set WsN = WbN.Worksheets(1)
    Set LoN = WsN.ListObjects(1)

    Paso = "recortar la copia"
    If LoN.ShowAutoFilter Then             '- el filtro de la tabla, desde la tabla (como Rut_Lo_Filtros_Quitar):
        If LoN.AutoFilter.FilterMode Then LoN.AutoFilter.ShowAllData    '- con WsN.ShowAllData da error 1004
    End If
    If WsN.FilterMode Then WsN.ShowAllData  '- si aún queda alguno (un filtro avanzado)
    WsN.Cells.EntireRow.Hidden = False
    WsN.Cells.EntireColumn.Hidden = False
    If Not LoN.ShowTotals Then LoN.ShowTotals = True
    WsN.Range(WsN.Cells(1, LoN.Range.Column + InfRec_Descrip_Contab), _
              WsN.Cells(1, WsN.Columns.Count)).EntireColumn.Delete
    Set Fila1 = Intersect(WsN.Rows(1), WsN.UsedRange)
    If Not Fila1 Is Nothing Then
        For Each Celda In Fila1.Cells
            If Not IsError(Celda.Value) Then
                If LCase$(Trim$(CStr(Celda.Value))) = "x" Then Celda.ClearContents
            End If
        Next Celda
    End If

    Paso = "copiar los altos de fila"
    '- Casi todas las filas tienen alto automático (M_130 les hace AutoFit), y en el libro nuevo Excel los vuelve a
    '- calcular a su manera (filas de 15,5 pasaban a 30). Se fijan los de la hoja, que son los que se ven en PPub.
    '- Una fila oculta en la hoja (por un filtro) da alto 0: se queda con el suyo.
    For i = 1 To Ws.UsedRange.Row + Ws.UsedRange.Rows.Count - 1
        If Not Ws.Rows(i).Hidden Then WsN.Rows(i).RowHeight = Ws.Rows(i).RowHeight
    Next i
    WsN.Rows(1).RowHeight = WsN.StandardHeight  '- la fila 1 queda vacía (sin las "x"): alto normal (decisión del usuario)

    Paso = "quitar las formas con macro, los vínculos y los nombres que sobran"
    For i = WsN.Shapes.Count To 1 Step -1
        If Fnc_Inf_Recibos_Forma_Con_Macro(WsN.Shapes(i)) Then WsN.Shapes(i).Delete
    Next i
    Vinculos = WbN.LinkSources(xlExcelLinks)
    If Not IsEmpty(Vinculos) Then
        For i = LBound(Vinculos) To UBound(Vinculos)
            WbN.BreakLink Vinculos(i), xlLinkTypeExcelLinks
        Next i
    End If
    For i = WbN.Names.Count To 1 Step -1
        Call Rut_Inf_Recibos_Nombre_Quitar_Si_Sobra(WbN.Names(i), WbN)
    Next i

    Paso = "dejar la vista al principio"
    WsN.Activate
    ActiveWindow.Zoom = 100                 '- la hoja está al 55 %; el adjunto, al 100 % (decisión del usuario)
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    WsN.Range("A1").Select

    Paso = "guardar el .xlsx"
    CalcAnt = Application.Calculation
    If CalcAnt <> xlCalculationAutomatic Then
        Application.Calculation = xlCalculationAutomatic:   Calc_Hecho = True
    End If
    Application.DisplayAlerts = False       '- descarta sin preguntar el código de la hoja (un .xlsx no lleva macros)
    WbN.SaveAs Filename:=Ruta, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    WbN.Close SaveChanges:=False
    Set WbN = Nothing
    Fnc_Inf_Recibos_A_Xlsx = Ruta

Salir:
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not WbN Is Nothing Then WbN.Close SaveChanges:=False
    If Calc_Hecho Then Application.Calculation = CalcAnt
    If Vis_Hecho Then Ws.Visible = VisAnt
    ThisWorkbook.Activate
    If Not WsAct Is Nothing Then
        If WsAct.Visible = xlSheetVisible Then WsAct.Activate
    End If
    If Sw_Hecho Then Prog__APP_Switch.Range("Sw_WB_Deactivate").Value = SwDeactAnt
    On Error GoTo 0
    If Off_Hecho Then Call Rut_On_Functions
    Exit Function

Err_Xlsx:
    ErrMsg = "Error " & Err.Number & " al " & Paso & ": " & Err.Description
    Resume Salir
End Function    ' Fnc_Inf_Recibos_A_Xlsx
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Botón «Enviar por correo» del grupo Informe de Recibos (OnAct_Mail_Inf_Recibos) -----------------
'- El callback abre y cierra Form_Running_Rut (Rut_Progreso_Abrir / Rut_Progreso_Cerrar): aquí solo
'- se escribe en su informe. No recalcula el informe: manda lo que hay en la hoja, y la pregunta
'- enseña la fecha del último cálculo. Destinatarios: App_Mail_Inf_Recibos (hoja ConfigAPP). Con
'- Sw_Test, el correo va a App_User_Mail y el .xlsx se queda en %TEMP% para revisarlo.
'- Preguntas y avisos, por delante del formulario (vbMsgBoxSetForeground + vbSystemModal).
'- Copia en la carpeta de red (2026-10-13): con su propia pregunta, después de la del envío; el .xlsx
'- se copia en App_Ruta_Inf_Recibos (Fnc_Inf_Recibos_Carpeta_Red) antes de enviarlo. Si la carpeta no
'- se encuentra, pregunta si se envía el correo de todos modos. Si la copia falla, avisa y envía igual.
'- Desde el 2026-10-13 (noche), dentro de la subcarpeta del año contable (App_AnoCont):
'- ...\DOCUMENTOS DUGARRY\2026\. Si no existe, la pregunta lo dice y se crea al copiar.
Sub Rut_Mail_Inf_Recibos()
    Dim Dest        As String
    Dim DestReal    As String
    Dim ErrMsg      As String
    Dim Ruta        As String
    Dim Asunto      As String
    Dim Cuerpo      As String
    Dim AnoCont     As String
    Dim Cierre      As String
    Dim Ult_Calc    As String
    Dim Descarga    As String
    Dim SwTest      As Boolean
    Dim Pregunta    As String
    Dim Estilo      As VbMsgBoxStyle
    Dim Respuesta   As VbMsgBoxResult
    Dim Carpeta_Red As String               '- carpeta de red para la copia ("" si no se encuentra)
    Dim Destino_Red As String               '- el .xlsx en la carpeta de red
    Dim Copiar_Red  As Boolean
    Dim Crear_Carpeta As Boolean            '- la subcarpeta del año contable no existe: se crea al copiar
    Dim Err_Red     As String
    Dim N_Err       As Long

    Set ActivForm = Form_Running_Rut
    H_Inicio = Timer
    LastTimeLap = Timer
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Proceso: enviar por correo el Informe de Recibos (hoja Inf_Recibos_TIO).", 0)

    '- Lo que dicen la pregunta, el asunto y el cuerpo
    SwTest = CBool(Prog__APP_Switch.Range("Sw_Test").Value)
    AnoCont = CStr(Prog__APP.Range("App_AnoCont").Value)
    Cierre = Format$(Prog__APP.Range("App_FechCierreCont").Value, "dd/mm/yyyy")
    Ult_Calc = Format$(Prog__APP.Range("App_Last_Calc_Inf_Recibos").Value, "dd/mm/yyyy hh:nn")
    Descarga = Fnc_Inf_Recibos_Fecha_Robot()                 '- AAAA-MM-DD, o vacía
    If Descarga <> "" Then Descarga = Format$(CDate(Descarga), "dd/mm/yyyy")

    Dest = Trim$(CStr(Prog__APP.Range("App_Mail_Inf_Recibos").Value))
    If Not Fnc_Valid_Email_Lista(Dest, ErrMsg) Then
        Call Rut_Mail_Inf_Recibos_Aviso("No se ha enviado nada: los destinatarios (App_Mail_Inf_Recibos, hoja ConfigAPP) " & _
                                        "no valen." & vbLf & ErrMsg)
        Exit Sub
    End If
    '- La pregunta: un envío real se ve como tal y lleva "No" por defecto (2026-10-13: con el botón
    '- "Prueba ON" del Ribbon a la vista, que quería decir "pulsa para activarla", se creyó en modo
    '- prueba y el informe llegó a los destinatarios reales).
    If SwTest Then
        Pregunta = "MODO PRUEBA (Sw_Test): irá solo a " & Prog__APP.Range("App_User_Mail").Value & vbLf & _
                   "(no a " & Replace(Dest, ";", "; ") & ")" & vbLf & vbLf
        Estilo = vbYesNo + vbQuestion
    Else
        Pregunta = "¡ ENVÍO REAL !   Va a:  " & Replace(Dest, ";", "; ") & vbLf & vbLf
        Estilo = vbYesNo + vbExclamation + vbDefaultButton2     '- un Intro sin mirar no envía nada
    End If
    Pregunta = Pregunta & "¿ Enviar el Informe de Recibos por correo ?" & vbLf & vbLf & _
               "Último cálculo: " & Ult_Calc & vbLf & _
               "Cierre contable: " & Cierre & _
               IIf(Descarga <> "", vbLf & "Descarga de los recibos: " & Descarga, "")
    Respuesta = MsgBox(Pregunta, Estilo + vbMsgBoxSetForeground + vbSystemModal, "Enviar el Informe de Recibos")
    If Respuesta <> vbYes Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Cancelado a petición del usuario: no se ha enviado nada.", 0)
        Exit Sub
    End If
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", IIf(SwTest, "Modo prueba (Sw_Test): irá solo a " & _
                         Prog__APP.Range("App_User_Mail").Value & ".", "Envío real a " & Replace(Dest, ";", "; ") & "."), 0)

    '- La copia en la carpeta de red (App_Ruta_Inf_Recibos), con su propia pregunta (2026-10-13)
    Carpeta_Red = Fnc_Inf_Recibos_Carpeta_Red(Err_Red)
    If Carpeta_Red <> "" Then
        If Not Fnc_Inf_Recibos_Existe(Carpeta_Red, True) Then
            Err_Red = "No se encuentra la carpeta " & Carpeta_Red & vbLf & "(¿ está conectada la unidad de red ?)"
            Carpeta_Red = ""
        End If
    End If
    If Carpeta_Red = "" Then
        Respuesta = MsgBox("No se puede copiar el informe en la carpeta de red:" & vbLf & Err_Red & vbLf & vbLf & _
                           "¿ Enviar el correo de todos modos ?", _
                           vbYesNo + vbExclamation + vbMsgBoxSetForeground + vbSystemModal, "Enviar el Informe de Recibos")
        If Respuesta <> vbYes Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Cancelado: no se puede copiar en la carpeta de red y no se ha " & _
                                 "enviado nada." & vbLf & Err_Red, 0)
            Exit Sub
        End If
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Sin copia en la carpeta de red: " & Replace(Err_Red, vbLf, " "), 0)
    Else
        '- En la subcarpeta del año contable (petición del usuario, 2026-10-13); si no existe, se crea al copiar
        If AnoCont <> "" Then Carpeta_Red = Carpeta_Red & "\" & AnoCont
        Crear_Carpeta = Not Fnc_Inf_Recibos_Existe(Carpeta_Red, True)
        Destino_Red = Carpeta_Red & "\" & Fnc_Inf_Recibos_Nombre() & ".xlsx"
        Pregunta = IIf(SwTest, "MODO PRUEBA (Sw_Test): la carpeta de red es la de verdad." & vbLf & vbLf, "") & _
                   "¿ Copiar también el informe en la carpeta de red ?" & vbLf & vbLf & _
                   Carpeta_Red & vbLf & Mid$(Destino_Red, Len(Carpeta_Red) + 2)
        If Crear_Carpeta Then
            Pregunta = Pregunta & vbLf & vbLf & "La carpeta " & AnoCont & " no existe: se creará."
        ElseIf Fnc_Inf_Recibos_Existe(Destino_Red, False) Then
            Pregunta = Pregunta & vbLf & vbLf & "Ya hay un fichero con ese nombre: se sustituirá."
        End If
        Estilo = vbYesNo + vbQuestion
        If SwTest Then Estilo = Estilo + vbDefaultButton2       '- en modo prueba, un Intro sin mirar no copia
        Respuesta = MsgBox(Pregunta, Estilo + vbMsgBoxSetForeground + vbSystemModal, "Copiar el Informe de Recibos")
        Copiar_Red = (Respuesta = vbYes)
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", IIf(Copiar_Red, "Se copiará en " & Destino_Red, _
                             "Sin copia en la carpeta de red (respuesta del usuario)."), 0)
    End If

    '- El adjunto
    Ruta = Fnc_Inf_Recibos_A_Xlsx(ErrMsg)
    If Ruta = "" Then
        Call Rut_Mail_Inf_Recibos_Aviso("No se ha enviado nada: no se ha podido generar el .xlsx." & vbLf & ErrMsg)
        Exit Sub
    End If
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Adjunto: " & Mid$(Ruta, InStrRev(Ruta, "\") + 1), LastTimeLap)
    LastTimeLap = Timer

    '- La copia en la carpeta de red, antes del envío: si falla, se avisa y el correo se envía igual
    If Copiar_Red Then
        On Error Resume Next
        If Crear_Carpeta Then MkDir Carpeta_Red
        If Err.Number = 0 Then FileCopy Ruta, Destino_Red
        N_Err = Err.Number:     Err_Red = Err.Description
        On Error GoTo 0
        If N_Err = 0 Then
            Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", IIf(Crear_Carpeta, "Creada la carpeta " & Carpeta_Red & vbLf, "") & _
                                 "Copiado en la carpeta de red: " & Destino_Red, LastTimeLap)
        Else
            Call Rut_Mail_Inf_Recibos_Aviso("No se ha podido copiar el informe en la carpeta de red (el correo se envía igual):" & _
                                            vbLf & Destino_Red & vbLf & "Error " & N_Err & ": " & Err_Red & _
                                            IIf(N_Err = 70, vbLf & "(¿ lo tiene alguien abierto ?)", ""))
        End If
        LastTimeLap = Timer
    End If

    '- Asunto y cuerpo (sin los totales: decisión del usuario)
    Asunto = "PPub - Informe de Recibos, año contable " & AnoCont & ", cierre " & Cierre
    Cuerpo = "<div style='line-height:1.8; font-family:Arial, sans-serif; font-size:16px;'>" & _
             "<b>Hola</b><br><br>" & _
             "Adjunto el <b>Informe de Recibos</b> de Precios Públicos del año contable " & AnoCont & _
             " (cursos " & Fnc_HTML_Texto(CStr(Prog__APP.Range("App_C_Acad_Ant").Value)) & " y " & _
             Fnc_HTML_Texto(CStr(Prog__APP.Range("App_C_Acad_Pos").Value)) & "): importes emitidos, cobrados " & _
             "y pendientes por tipo de recibo y concepto, desglosados en seguro obligatorio (INSS), " & _
             "administrativos y académicos, con los justificantes de ingreso (JI) y la descripción de " & _
             "cada movimiento contable. Las filas de EURLE van aparte, debajo de la tabla.<br><br>" & _
             IIf(Descarga <> "", "Recibos descargados el " & Descarga & ". ", "") & _
             "Fecha de cierre contable: " & Cierre & ". Último cálculo del informe: " & Ult_Calc & ".<br><br>" & _
             "</div>"

    '- El envío
    If Fnc_Email_Send(Dest, Asunto, Ruta, Cuerpo, ErrMsg, DestReal) Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Enviado a " & Replace(DestReal, ";", "; ") & _
                             IIf(SwTest, " (modo prueba, Sw_Test).", "."), LastTimeLap)
    Else
        Call Rut_Mail_Inf_Recibos_Aviso("No se ha enviado el correo." & vbLf & ErrMsg)
    End If
    If SwTest Then
        Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", "Modo prueba: el adjunto se queda para revisarlo en " & Ruta, 0)
    Else
        On Error Resume Next
        Kill Ruta
        On Error GoTo 0
    End If
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", String(100, "-") & vbLf & "Proceso Finalizado. " & _
                         Format(Now(), "dd-mmm-yy hh:mm"), H_Inicio, , , , , , 2)
End Sub     ' Rut_Mail_Inf_Recibos
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Aviso de un fallo del envío: al informe del formulario y en un MsgBox por delante de él ---------
Private Sub Rut_Mail_Inf_Recibos_Aviso(ByVal Texto As String)
    Call Rut_TimeLap_Inf(ActivForm, "TBx_Informe", Texto, 0)
    MsgBox Texto, vbExclamation + vbMsgBoxSetForeground + vbSystemModal, "Enviar el Informe de Recibos"
End Sub     ' Rut_Mail_Inf_Recibos_Aviso
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Prueba de la fase 2, desde Inmediato: Call Rut_Inf_Recibos_Probar_Xlsx --------------------------
'- Genera el .xlsx sin enviarlo y lo deja en %TEMP%, seleccionado en el Explorador para abrirlo.
Sub Rut_Inf_Recibos_Probar_Xlsx()
    Dim Ruta        As String
    Dim ErrMsg      As String
    Dim T_Ini       As Double

    T_Ini = Timer
    Ruta = Fnc_Inf_Recibos_A_Xlsx(ErrMsg)
    If Ruta = "" Then
        Debug.Print "Fnc_Inf_Recibos_A_Xlsx: no se ha generado. " & ErrMsg
        MsgBox "No se ha generado el .xlsx:" & vbLf & vbLf & ErrMsg, _
               vbExclamation + vbMsgBoxSetForeground, "Prueba del Informe de Recibos en .xlsx"
    Else
        Debug.Print "Fnc_Inf_Recibos_A_Xlsx: " & Format$(Timer - T_Ini, "0.00") & " s, " & Ruta
        Shell "explorer.exe /select,""" & Ruta & """", vbNormalFocus
        MsgBox "Generado en " & Format$(Timer - T_Ini, "0.0") & " s:" & vbLf & Ruta & vbLf & vbLf & _
               "Está seleccionado en el Explorador para abrirlo y revisarlo.", _
               vbInformation + vbMsgBoxSetForeground, "Prueba del Informe de Recibos en .xlsx"
    End If
End Sub     ' Rut_Inf_Recibos_Probar_Xlsx
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Ruta del .xlsx: Carpeta & Nombre & ".xlsx", o con la hora detrás si esa no está libre -----------
'- No está libre si hay un libro abierto con ese nombre (Excel no deja abrir dos libros con el mismo
'- nombre, y el SaveAs fallaría) o si el fichero ya existe y no se puede borrar (abierto en otro Excel).
Private Function Fnc_Inf_Recibos_Ruta_Libre(ByVal Carpeta As String, ByVal Nombre As String) As String
    Dim Wb          As Workbook
    Dim Ruta        As String
    Dim Libre       As Boolean

    Ruta = Carpeta & Nombre & ".xlsx"
    Libre = True
    For Each Wb In Application.Workbooks
        If StrComp(Wb.Name, Nombre & ".xlsx", vbTextCompare) = 0 Then Libre = False
    Next Wb
    If Libre Then
        On Error Resume Next
        If Len(Dir$(Ruta)) > 0 Then
            Kill Ruta
            Libre = (Err.Number = 0)
        End If
        On Error GoTo 0
    End If
    If Not Libre Then Ruta = Carpeta & Nombre & "_" & Format$(Now, "hhnnss") & ".xlsx"
    Fnc_Inf_Recibos_Ruta_Libre = Ruta
End Function    ' Fnc_Inf_Recibos_Ruta_Libre
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Si una forma tiene una macro asignada (algunos tipos de forma dan error al leer OnAction) -------
Private Function Fnc_Inf_Recibos_Forma_Con_Macro(ByVal Forma As Shape) As Boolean
    On Error Resume Next
    Fnc_Inf_Recibos_Forma_Con_Macro = (Len(Forma.OnAction) > 0)
End Function    ' Fnc_Inf_Recibos_Forma_Con_Macro
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Borra un nombre de la copia si no apunta a un rango de la propia copia (Wb) ---------------------
'- Así se van los rotos (#REF!), los que apuntan a otro libro y las constantes de PPub que Excel copia
'- con la hoja (RibbonPtr, el puntero del Ribbon). Los internos de Excel (_FilterDatabase) se quedan.
Private Sub Rut_Inf_Recibos_Nombre_Quitar_Si_Sobra(ByVal Nm As Name, ByVal Wb As Workbook)
    Dim Rng         As Range

    On Error Resume Next
    Set Rng = Nm.RefersToRange
    If Rng Is Nothing Then
        Nm.Delete
    ElseIf Not Rng.Worksheet.Parent Is Wb Then
        Nm.Delete
    End If
End Sub     ' Rut_Inf_Recibos_Nombre_Quitar_Si_Sobra
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Fecha de la carpeta de descarga del Robot (App_Last_Import_Robot_PPub) como "AAAA-MM-DD" --------
'- "" si la celda está vacía, no es una fecha o el nombre no existe: el adjunto va entonces sin fecha.
Private Function Fnc_Inf_Recibos_Fecha_Robot() As String
    Dim V           As Variant

    On Error Resume Next
    V = Prog__APP.Range("App_Last_Import_Robot_PPub").Value
    If IsDate(V) Then
        Fnc_Inf_Recibos_Fecha_Robot = Format$(CDate(V), "yyyy-mm-dd")
    ElseIf Not IsEmpty(V) And IsNumeric(V) Then                 '- un número de serie (celda sin formato de fecha)
        If CDbl(V) > 0 Then Fnc_Inf_Recibos_Fecha_Robot = Format$(CDate(CDbl(V)), "yyyy-mm-dd")
    End If
End Function    ' Fnc_Inf_Recibos_Fecha_Robot
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Nombre del .xlsx del informe, sin la extensión: "Inf_Recibos_TIO (datos del AAAA-MM-DD)" --------
'- ("Inf_Recibos_TIO" sin la fecha del Robot). Lo usan Fnc_Inf_Recibos_A_Xlsx y la copia en la red.
Private Function Fnc_Inf_Recibos_Nombre() As String
    Dim Fecha       As String

    Fnc_Inf_Recibos_Nombre = "Inf_Recibos_TIO"
    Fecha = Fnc_Inf_Recibos_Fecha_Robot()
    If Fecha <> "" Then Fnc_Inf_Recibos_Nombre = Fnc_Inf_Recibos_Nombre & " (datos del " & Fecha & ")"
End Function    ' Fnc_Inf_Recibos_Nombre
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Carpeta de red para la copia del informe (App_Ruta_Inf_Recibos, hoja ConfigAPP), sin "\" final --
'- La celda guarda la ruta sin la unidad ("\Espai_ServConta\..."), como App_RutaRed_Robot_PPub: se le
'- pone delante la unidad de red del usuario (App_User_Unid_Red, que Form_Usuario cambia con el
'- teletrabajo). Una ruta completa ("Y:\...", "\\servidor\...") se usa tal cual, y una de la Red Nexe
'- (https://nexe.ua.es/...) se traduce con Fnc_Format_Ruta. "" y la causa en ErrMsg si falta algo.
Private Function Fnc_Inf_Recibos_Carpeta_Red(Optional ByRef ErrMsg As String = "") As String
    Dim Ruta        As String
    Dim Unidad      As String

    ErrMsg = ""
    On Error Resume Next
    Ruta = Trim$(CStr(Prog__APP.Range("App_Ruta_Inf_Recibos").Value))
    Unidad = Trim$(CStr(Prog__APP.Range("App_User_Unid_Red").Value))
    On Error GoTo 0
    If Ruta = "" Then
        ErrMsg = "Falta la ruta de la carpeta (App_Ruta_Inf_Recibos, hoja ConfigAPP)."
        Exit Function
    End If
    If LCase$(Left$(Ruta, 8)) = "https://" Then             '- Red Nexe: a la unidad local
        Ruta = Fnc_URL_Decodificar(Fnc_Format_Ruta(Ruta))
    ElseIf Left$(Ruta, 2) = "\\" Or Mid$(Ruta, 2, 1) = ":" Then
                                                            '- ruta completa: tal cual
    Else
        If Unidad = "" Then
            ErrMsg = "Falta la unidad de red del usuario (App_User_Unid_Red, hoja ConfigAPP)."
            Exit Function
        End If
        If Left$(Ruta, 1) <> "\" Then Ruta = "\" & Ruta
        Ruta = Left$(Unidad, 1) & ":" & Ruta
    End If
    Do While Right$(Ruta, 1) = "\" And Len(Ruta) > 3
        Ruta = Left$(Ruta, Len(Ruta) - 1)
    Loop
    Fnc_Inf_Recibos_Carpeta_Red = Ruta
End Function    ' Fnc_Inf_Recibos_Carpeta_Red
'---------------------------------------------------------------------------------------------------

'===================================================================================================
'- Si existe la carpeta (Carpeta:=True) o el fichero; False también si la unidad no responde -------
Private Function Fnc_Inf_Recibos_Existe(ByVal Ruta As String, ByVal Carpeta As Boolean) As Boolean
    Dim Atrib       As Long

    On Error GoTo No_Existe
    Atrib = GetAttr(Ruta)
    Fnc_Inf_Recibos_Existe = (((Atrib And vbDirectory) = vbDirectory) = Carpeta)
No_Existe:
End Function    ' Fnc_Inf_Recibos_Existe
'---------------------------------------------------------------------------------------------------
