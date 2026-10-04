Attribute VB_Name = "Rut_Wb_CopSegTimed_USB_HD"
' Last Rev. 2026-10-04 17:08
'='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='=
' Rut_Wb_CopSegTimed_USB_HD  -  Copias de seguridad del libro con marca de tiempo
'
' Genera copias de seguridad en dos modalidades: USB con ruta configurable (con diálogo
' y actualización automática de path) o directorio local CopiaSeguridad/ sin diálogo.
' La copia HD se lanza desde Workbook_Open cuando han pasado APP_CopSeg_HD_Frecuency días
' (o más) desde la última (Fnc_CopSeg_HD_Toca); al hacerse, purga las copias del libro
' con más de APP_CopSeg_HD_FrecPurga meses (los backups manuales se conservan).
' Toda la configuración vive en rangos con nombre de ámbito Libro en la hoja ConfigCopSeg;
' Rut_CopSeg_Asegurar_Config crea hoja y rangos si faltan, así que el módulo es autoinstalable
' en cualquier libro. Los rangos se leen vía Fnc_CopSeg_Rng (por nombre, sin CodeName de hoja).
' Hacer las copias y gestionarlas son procesos independientes: el gestor (listar / borrar)
' vive en el módulo Rut_Wb_CopSegTimed_Gestor.
'
' Este módulo NO requiere ningún otro módulo del libro destino para compilar: es 100%
' autocontenido. Dos hooks opcionales (vía Application.Run, en tiempo de ejecución, sin
' dependencia de compilación) permiten integrarlo con un libro que ya tenga:
'   - Fnc_Format_Ruta(Ruta As String) As String   - adapta rutas de red (p.ej. NEXE/UA)
'   - Rut_CopSeg_Feedback_Host(Msg As String)      - recibe el feedback de cada copia/error
' Si el libro destino no tiene ninguna de las dos, el módulo funciona igual (fallback
' identidad en la ruta; feedback descartado en silencio).
'
' Configuración (hoja ConfigCopSeg, autocreada):
'   APP_CopSeg_Usb_Path, APP_CopSeg_Usb_Date, APP_CopSeg_HD_Date,
'   APP_CopSeg_HD_Frecuency, APP_CopSeg_HD_FrecPurga
'
' Índice de Subs/Functions:
'   Rut_WrkBook_CopSegTimed_USB               - Guarda copia del libro con marca de tiempo; ruta USB con diálogo
'   Rut_WrkBook_CopSegTimed_WB_HD             - Guarda copia del libro en CopiaSeguridad/ sin diálogo y purga las antiguas
'   Fnc_CopSeg_HD_Toca                        - True si toca copia HD según APP_CopSeg_HD_Frecuency (llamar desde Workbook_Open)
'   Fnc_CopSeg_Rng                            - Range de un nombre de ámbito Libro, viva donde viva su celda
'   Fnc_CopSeg_Range_Exist (Private)          - Comprueba si existe un nombre definido a nivel de libro
'   Fnc_CopSeg_Cfg_Long (Private)             - Lee un parámetro numérico de la config, con valor por defecto
'   Fnc_CopSeg_HD_Purgar (Private)            - Elimina de CopiaSeguridad/ las copias con más de APP_CopSeg_HD_FrecPurga meses
'   Rut_CopSeg_Asegurar_Config                - Setup idempotente: hoja ConfigCopSeg y sus rangos de configuración
'   Fnc_CopSeg_Ruta (Private)                 - Hook opcional: adapta la ruta vía Fnc_Format_Ruta del anfitrión si existe
'   Rut_CopSeg_Cfg_Rango (Private)            - Crea un rango de configuración (nombre + etiqueta + defecto) si falta
'   Fnc_CopSeg_Log (Private)                  - Hook opcional: feedback vía Rut_CopSeg_Feedback_Host del anfitrión si existe
'='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='=

Option Explicit

Private Const C_CopSeg_HojaCfg          As String = "ConfigCopSeg"                       ' Hoja de configuración del módulo
Private Const C_CopSeg_Usb_Path_Def     As String = "F:\__CopSeg Versiones Programas\"   ' Ruta USB por defecto (APP_CopSeg_Usb_Path)
Private Const C_CopSeg_HD_Frec_Def      As Long = 1                                      ' Días entre copias HD por defecto (APP_CopSeg_HD_Frecuency)
Private Const C_CopSeg_HD_FrecPurga_Def As Long = 3                                      ' Meses de retención por defecto (APP_CopSeg_HD_FrecPurga)

'===============================================================================
Sub Rut_WrkBook_CopSegTimed_USB(Optional Tipo As String = "")  '- Guarda Copia de Este Excel con marca de tiempo en el nombre del archivo. (Tipo="Data"/"VBA", etc.) ----------
    Call Rut_CopSeg_Asegurar_Config
    Dim Answer          As VbMsgBoxResult
    Dim FichNom         As String
        FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    Dim FichExt         As String
        FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
    Dim fichPath        As String
        fichPath = Fnc_CopSeg_Rng("APP_CopSeg_Usb_Path").Value & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
    Dim FichSelect      As Variant
        FichSelect = Application.GetSaveAsFilename(fichPath, "Excel Files (*" & FichExt & "), *" & FichExt)
        If VarType(FichSelect) = vbBoolean Then Exit Sub    ' El usuario cancela el diálogo: ni copia, ni sello de fecha, ni cambio de ruta
        On Error GoTo Finalizar
        Application.DisplayAlerts = False
        ThisWorkbook.SaveCopyAs Filename:=FichSelect
        Application.DisplayAlerts = True
        On Error GoTo 0
    If Fnc_CopSeg_Rng("APP_CopSeg_Usb_Path").Value <> Left(FichSelect, InStrRev(FichSelect, "\")) Then
        ' Pedir confirmación de cambio de ruta
        Dim Mensage     As String
        Mensage = "¿ Cambiamos esta ruta: " & Fnc_CopSeg_Rng("APP_CopSeg_Usb_Path").Value & vbLf & _
                  " por esta ? " & Left(FichSelect, InStrRev(FichSelect, "\")) & vbCrLf & vbCrLf
        Answer = MsgBox(Mensage, vbExclamation + vbYesNo + vbDefaultButton2, "Proceso: Cambio de Ruta para las Copias de Seguridad.")
        If Answer = vbYes Then
            Fnc_CopSeg_Rng("APP_CopSeg_Usb_Path").Value = Left(FichSelect, InStrRev(FichSelect, "\"))
        End If
    End If
    Fnc_CopSeg_Rng("APP_CopSeg_Usb_Date").Value = Now
    Call Fnc_CopSeg_Log("Copia USB realizada: " & FichSelect & "  -  " & Now())
Finalizar:
    Application.DisplayAlerts = True     ' Por si SaveCopyAs falla con los avisos desactivados
End Sub
' ------------------------------------------------------------------------------

' ==============================================================================
Sub Rut_WrkBook_CopSegTimed_WB_HD()   '- Guarda copia del libro en CopiaSeguridad/ sin diálogo; purga las copias con más de APP_CopSeg_HD_FrecPurga meses
' ==============================================================================
    '- Guardar estado de Excel (por si es manual o eventos desactivados) ----------
    Dim PrevScreen  As Boolean:    PrevScreen = Application.ScreenUpdating
    Dim PrevEvents  As Boolean:    PrevEvents = Application.EnableEvents
    Dim PrevCalc    As XlCalculation:  PrevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Call Rut_CopSeg_Asegurar_Config
    Dim T_Inicio                            As Single
                T_Inicio = Timer            ' Para saber el tiempo de proceso
    Dim FichNom                             As String
                FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    Dim FichExt                             As String
                FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
    Dim FPath                               As String
                FPath = Fnc_CopSeg_Ruta(ThisWorkbook.Path & "/CopiaSeguridad/")
    '- Fecha de los DATOS: el último guardado del fichero en disco, no la de hoy ---
    '  (la copia se lanza al abrir, pero su contenido es el del último día de trabajo;
    '   si no hubo guardados nuevos, el nombre se repite y la copia solo se sobreescribe)
    Dim FechaDatos                          As Date
                On Error Resume Next
                FechaDatos = FileDateTime(Fnc_CopSeg_Ruta(ThisWorkbook.FullName))
                On Error GoTo 0
                If FechaDatos = 0 Then FechaDatos = Now     ' Fallback: no se pudo leer el fichero en disco
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName As String
    IntialName = FPath & FichNom & " - " & Format(FechaDatos, "(yymmdd_hhmm)") & FichExt
            On Error GoTo GestError
            Application.DisplayAlerts = False
            ThisWorkbook.SaveCopyAs Filename:=IntialName
            Application.DisplayAlerts = True
            On Error GoTo 0
    Fnc_CopSeg_Rng("APP_CopSeg_HD_Date").Value = Now
    '- Purgar las copias con más de APP_CopSeg_HD_FrecPurga meses ---------------
    Dim Purgadas    As Long
    Purgadas = Fnc_CopSeg_HD_Purgar(FPath, FichNom, FichExt)

    Call Fnc_CopSeg_Log( _
        "Copia de Seguridad realizada en " & Round(Timer - T_Inicio, 2) & " seg.: " & IntialName & _
        IIf(Purgadas > 0, "  |  Eliminadas " & Purgadas & " copias antiguas", "") & "  -  " & Now())
    GoTo Salir_Sub
GestError:
    Application.DisplayAlerts = True
    Debug.Print "Error Rut_WrkBook_CopSegTimed_WB_HD", Err.Number, Err.Description, Err.Source
    MsgBox "Error: Rut_WrkBook_CopSegTimed_WB_HD" & vbLf & vbLf & "Err.Number:" & Err.Number & vbLf & vbLf & "Err.Description:" & Err.Description & _
            vbLf & vbLf & "Filename:=" & vbCrLf & vbCrLf & IntialName, vbExclamation + vbOKOnly, "Proceso: Copia de Seguridad."
    Call Fnc_CopSeg_Log("ERROR en la Copia de Seguridad: " & IntialName)
Salir_Sub:
    '- Restaurar estado de Excel -----------------------------------------------
    Application.ScreenUpdating = PrevScreen
    Application.EnableEvents = PrevEvents
    Application.Calculation = PrevCalc
End Sub     ' Rut_WrkBook_CopSegTimed_WB_HD
' ------------------------------------------------------------------------------

' ==============================================================================
Function Fnc_CopSeg_HD_Toca() As Boolean   '- True si toca copia HD: han pasado APP_CopSeg_HD_Frecuency días (o más) desde la última
' ==============================================================================
' Con Frecuency = 1 replica el comportamiento clásico: copia diaria si la última no es de hoy.
' Si el sello está vacío o no es fecha (1ª instalación), devuelve True.
    Call Rut_CopSeg_Asegurar_Config
    Dim UltimaCopia     As Date
    On Error Resume Next
    UltimaCopia = CDate(Fnc_CopSeg_Rng("APP_CopSeg_HD_Date").Value)
    On Error GoTo 0
    Fnc_CopSeg_HD_Toca = (Int(CDbl(UltimaCopia)) + Fnc_CopSeg_Cfg_Long("APP_CopSeg_HD_Frecuency", C_CopSeg_HD_Frec_Def) <= Date)
End Function     ' Fnc_CopSeg_HD_Toca
' ------------------------------------------------------------------------------

' ==============================================================================
Function Fnc_CopSeg_Rng(Nombre As String) As Range   '- Range de un nombre de ámbito Libro, viva donde viva su celda
' ==============================================================================
' Evita el error 1004 de Hoja.Range("Nombre") cuando el nombre apunta a una celda de otra
' hoja, y desacopla el módulo del CodeName de la hoja de configuración (portabilidad).
    Set Fnc_CopSeg_Rng = ThisWorkbook.Names(Nombre).RefersToRange
End Function     ' Fnc_CopSeg_Rng
' ------------------------------------------------------------------------------

' ==============================================================================
Private Function Fnc_CopSeg_Cfg_Long(Nombre As String, Defecto As Long) As Long   '- Lee un parámetro numérico de la config; Defecto si falta, no es numérico o es < 1
' ==============================================================================
    Fnc_CopSeg_Cfg_Long = Defecto
    Dim Valor           As Variant
    On Error Resume Next
    Valor = Fnc_CopSeg_Rng(Nombre).Value
    On Error GoTo 0
    If IsNumeric(Valor) Then
        If CLng(Valor) >= 1 Then Fnc_CopSeg_Cfg_Long = CLng(Valor)
    End If
End Function     ' Fnc_CopSeg_Cfg_Long
' ------------------------------------------------------------------------------

' ==============================================================================
Private Function Fnc_CopSeg_HD_Purgar(FPath As String, FichNom As String, FichExt As String) As Long   '- Elimina las copias del libro con más de APP_CopSeg_HD_FrecPurga meses; devuelve cuántas borró
' ==============================================================================
' Solo toca ficheros con el patrón exacto de Rut_WrkBook_CopSegTimed_WB_HD:
'   "<FichNom> - (yymmdd_hhmm)<FichExt>"
' Cualquier otro fichero de la carpeta (backups manuales, otros libros...) se conserva.
    Dim FSO             As Object:      Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim Fichero         As Object
    Dim FechaLimite     As Date:        FechaLimite = DateAdd("m", -Fnc_CopSeg_Cfg_Long("APP_CopSeg_HD_FrecPurga", C_CopSeg_HD_FrecPurga_Def), Date)
    Dim Prefijo         As String:      Prefijo = FichNom & " - ("
    Dim ABorrar         As Collection:  Set ABorrar = New Collection
    Dim RutaFich        As Variant
    If Not FSO.FolderExists(FPath) Then Exit Function
    '- 1ª pasada: recopilar (no borrar mientras se itera la colección Files) ------
    For Each Fichero In FSO.GetFolder(FPath).Files
        If Left(Fichero.Name, Len(Prefijo)) = Prefijo And LCase(Right(Fichero.Name, Len(FichExt))) = LCase(FichExt) Then
            If Fichero.DateLastModified < FechaLimite Then ABorrar.Add Fichero.Path
        End If
    Next Fichero
    '- 2ª pasada: borrar ---------------------------------------------------------
    For Each RutaFich In ABorrar
        On Error Resume Next
        FSO.DeleteFile RutaFich, True
        If Err.Number = 0 Then Fnc_CopSeg_HD_Purgar = Fnc_CopSeg_HD_Purgar + 1
        Err.Clear
        On Error GoTo 0
    Next RutaFich
    If Fnc_CopSeg_HD_Purgar > 0 Then Debug.Print "Fnc_CopSeg_HD_Purgar: eliminadas " & Fnc_CopSeg_HD_Purgar & " copias anteriores a " & FechaLimite
End Function     ' Fnc_CopSeg_HD_Purgar
' ------------------------------------------------------------------------------

' ==============================================================================
Sub Rut_CopSeg_Asegurar_Config()   '- Setup idempotente: hoja ConfigCopSeg y sus rangos de configuración
' ==============================================================================
' Se puede llamar siempre: crea la hoja y los rangos que falten y elimina nombres APP_CopSeg_*
' rotos (#REF!, p.ej. tras mover la config desde otra hoja). Lo que ya existe queda intacto.
' El setup de la hoja del gestor de copias es aparte: Rut_CopSeg_Gestor_Asegurar_Config
    '- 1) Hoja de configuración (crearla si falta) ------------------------------
    Dim Ws              As Worksheet
    Dim HojaNueva       As Boolean
    On Error Resume Next
    Set Ws = ThisWorkbook.Worksheets(C_CopSeg_HojaCfg)
    On Error GoTo 0
    If Ws Is Nothing Then
        Dim EventosActivos  As Boolean
        EventosActivos = Application.EnableEvents
        Application.EnableEvents = False            ' El Add dispara eventos de hoja/activación
        Set Ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
        Ws.Name = C_CopSeg_HojaCfg
        Application.EnableEvents = EventosActivos
        HojaNueva = True
    End If
    If Ws.Range("$A$1").Value = "" Then             ' Cabeceras (solo la 1ª vez)
        Ws.Range("$A$1").Value = "Nombre Rango"
        Ws.Range("$B$1").Value = "Rango Variable"
        Ws.Range("$C$1").Value = "Descripción"
        Ws.Range("$A$1:$C$1").Font.Bold = True
    End If
    '- 2) Nombres APP_CopSeg_* rotos (#REF!): eliminarlos para poder recrearlos --
    Dim Nm              As Name
    Dim ABorrar         As Collection:  Set ABorrar = New Collection
    Dim NomRoto         As Variant
    For Each Nm In ThisWorkbook.Names
        If Left(Nm.Name, Len("APP_CopSeg_")) = "APP_CopSeg_" And InStr(1, Nm.RefersTo, "#REF!") > 0 Then ABorrar.Add Nm.Name
    Next Nm
    For Each NomRoto In ABorrar
        ThisWorkbook.Names(NomRoto).Delete
    Next NomRoto
    '- 3) Rangos de configuración (uno por fila; solo se crean si el nombre no existe) --
    Call Rut_CopSeg_Cfg_Rango(Ws, 2, "APP_CopSeg_Usb_Path", C_CopSeg_Usb_Path_Def, "Ruta para la CopSeg USB")
    Call Rut_CopSeg_Cfg_Rango(Ws, 3, "APP_CopSeg_Usb_Date", Empty, "Fecha de la última CopSeg USB")
    Call Rut_CopSeg_Cfg_Rango(Ws, 4, "APP_CopSeg_HD_Date", Empty, "Fecha de la última CopSeg HD")
    Call Rut_CopSeg_Cfg_Rango(Ws, 5, "APP_CopSeg_HD_Frecuency", C_CopSeg_HD_Frec_Def, "Frecuencia de días para las copias")
    Call Rut_CopSeg_Cfg_Rango(Ws, 6, "APP_CopSeg_HD_FrecPurga", C_CopSeg_HD_FrecPurga_Def, "Frecuencia de meses para las purgas")
    If HojaNueva Then Ws.Columns("A:C").AutoFit
End Sub     ' Rut_CopSeg_Asegurar_Config
' ------------------------------------------------------------------------------

' ==============================================================================
Private Function Fnc_CopSeg_Ruta(Ruta As String) As String   '- Adapta ruta NEXE si el libro anfitrion tiene Fnc_Format_Ruta; fallback identidad
' ==============================================================================
' Intenta llamar a Fnc_Format_Ruta via Application.Run (hook opcional). Si no existe
' o falla, devuelve la ruta sin cambios. En libros normales (sin rutas NEXE),
' ThisWorkbook.Path ya es local, así que funciona sin adaptación.
' Diseñada para portabilidad: el módulo CopSeg funciona en cualquier .xlsm.
    Dim Resultado As Variant
    On Error Resume Next
    Resultado = Application.Run("'" & ThisWorkbook.Name & "'!Fnc_Format_Ruta", Ruta)
    On Error GoTo 0
    If VarType(Resultado) = vbString And Resultado <> "" Then
        Fnc_CopSeg_Ruta = Resultado
    Else
        Fnc_CopSeg_Ruta = Ruta    ' Fallback: devuelve la ruta original si el adaptador falta/falla
    End If
End Function     ' Fnc_CopSeg_Ruta
' ------------------------------------------------------------------------------

' ==============================================================================
Private Sub Fnc_CopSeg_Log(Mensaje As String)   '- Feedback opcional: llama a Rut_CopSeg_Feedback_Host en el anfitrion si existe
' ==============================================================================
' Hook para integrarse con el log del libro anfitrion (p.ej. una tabla de tareas).
' Si el anfitrion no tiene esta rutina, el mensaje simplemente se descarta.
' Para recibir feedback, crear en el libro destino:
'   Public Sub Rut_CopSeg_Feedback_Host(Msg As String)
'       ' ... escribir Msg donde convenga (log, celda, Tb_Tareas...) ...
'   End Sub
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!Rut_CopSeg_Feedback_Host", Mensaje
    On Error GoTo 0
End Sub     ' Fnc_CopSeg_Log
' ------------------------------------------------------------------------------

' ==============================================================================
Private Function Fnc_CopSeg_Range_Exist(RngName As String) As Boolean   '- Comprueba si existe un nombre definido a nivel de libro
' ==============================================================================
    Dim Nm As Name
    On Error Resume Next
    Set Nm = ThisWorkbook.Names(RngName)
    On Error GoTo 0
    Fnc_CopSeg_Range_Exist = Not Nm Is Nothing
End Function     ' Fnc_CopSeg_Range_Exist
' ------------------------------------------------------------------------------

' ==============================================================================
Private Sub Rut_CopSeg_Cfg_Rango(Ws As Worksheet, Fila As Long, Nombre As String, Defecto As Variant, Descripc As String)   '- Crea un rango de configuración (nombre + etiqueta + defecto) si falta
' ==============================================================================
' Nombre de ámbito Libro apuntando a B<Fila> de la hoja de config; escribe la etiqueta en A,
' el valor por defecto en B (solo si la celda está vacía) y la descripción en C.
    If Fnc_CopSeg_Range_Exist(Nombre) Then Exit Sub        ' Ya existe (apunte donde apunte): no tocar
    ThisWorkbook.Names.Add Name:=Nombre, RefersTo:="='" & Ws.Name & "'!$B$" & Fila
    Ws.Range("$A$" & Fila).Value = Nombre
    If Ws.Range("$B$" & Fila).Value = "" And Not IsEmpty(Defecto) Then Ws.Range("$B$" & Fila).Value = Defecto
    If Ws.Range("$C$" & Fila).Value = "" Then Ws.Range("$C$" & Fila).Value = Descripc
    If Right(Nombre, 5) = "_Date" Then Ws.Range("$B$" & Fila).NumberFormat = "dd/mm/yyyy hh:mm"
End Sub     ' Rut_CopSeg_Cfg_Rango
' ------------------------------------------------------------------------------
