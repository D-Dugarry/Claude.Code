Attribute VB_Name = "M_990_Import_Moduls"
'='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='=
' M_9_Import_Moduls  -  Importa o reemplaza modulos VBA en el libro activo desde disco
'
' Flujo:
'   1. Ofrece carpeta por defecto (carpeta del libro / VBA_Moduls si existe)
'   2. Pregunta si importar TODOS los .bas/.cls/.frm de la carpeta,
'      o seleccionar cuales con el dialogo de ficheros (Ctrl+clic = varios)
'   3. Para cada modulo seleccionado:
'        - Si NO existe en el proyecto -> importar como nuevo
'        - Si existe y es tipo normal (modulo/clase/formulario) -> quitar y reimportar
'        - Si existe y es tipo Document (hoja/libro) -> sustituir codigo sin quitar
'
' REQUISITO: activar "Confiar en el acceso al modelo de objetos de proyectos de VBA"
'   Archivo -> Opciones -> Centro de confianza -> Configuracion del centro de confianza
'   -> Configuracion de macros -> marcar la casilla correspondiente
'
' NOTA FORMULARIOS (.frm):
'   El fichero .frx binario companion debe estar en la misma carpeta que el .frm.
'   El selector filtra solo .bas/.cls/.frm; los .frx se usan automaticamente.
'
' Dependencias:
'   Ninguna (modulo autonomo, usa late binding para VBProject)
'
' Subs/Functions publicas:
'   Rut_Import_Moduls_All  - Punto de entrada: seleccion + importacion de modulos
'='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='='=

Option Explicit

Private Const COMP_DOCUMENT  As Integer = 100              ' vbext_ct_Document (hoja/ThisWorkbook)
Private Const MY_MODULE_NAME As String = "M_9_Import_Moduls" ' nombre de este modulo: nunca eliminarlo mientras corre
Private Const RES_NEW        As String = "NUEVO      "
Private Const RES_REPLACED   As String = "REEMPLAZADO"
Private Const RES_UPDATED    As String = "ACTUALIZADO"
Private Const RES_SKIPPED    As String = "OMITIDO    "
Private Const RES_ERROR      As String = "ERROR      "


'###################################################################################################################################
Public Sub Rut_Import_Moduls_All()
' ==================================================================================================================================

    ' --- Verificar acceso al VBProject (requiere opcion de confianza activada) ---
    Dim vbp As Object
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    On Error GoTo 0
    If vbp Is Nothing Then
        MsgBox "Sin acceso al VBProject." & vbCrLf & vbCrLf & _
               "Activa: Archivo > Opciones > Centro de confianza >" & vbCrLf & _
               "Configuracion de macros >" & vbCrLf & _
               "Confiar en el acceso al modelo de objetos de proyectos de VBA", _
               vbCritical, "Importar Modulos VBA"
        Exit Sub
    End If

    ' --- Ruta por defecto: carpeta del libro; si existe subcarpeta VBA_Moduls, usarla ---
    Dim defPath As String
    defPath = ThisWorkbook.Path
    If Dir(defPath & "\VBA_Moduls", vbDirectory) <> "" Then defPath = defPath & "\VBA_Moduls"
    defPath = defPath & "\"

    ' --- Preguntar modo de seleccion ---
    Dim mode As Integer
    mode = MsgBox("Importar TODOS los modulos de una carpeta?" & vbCrLf & vbCrLf & _
                  "   SI        -> seleccionar carpeta e importar todos los .bas/.cls/.frm" & vbCrLf & _
                  "   NO        -> seleccionar ficheros individuales (Ctrl+clic = varios)" & vbCrLf & _
                  "   CANCELAR  -> salir", _
                  vbYesNoCancel + vbQuestion, "Importar Modulos VBA")
    If mode = vbCancel Then Exit Sub

    ' --- Recoger lista de ficheros a importar ---
    Dim selFiles() As String
    Dim selCount   As Long
    selCount = 0

    If mode = vbYes Then
        ' Picker de carpeta -> enumerar .bas/.cls/.frm
        Dim fdFldr As FileDialog
        Set fdFldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fdFldr
            .Title = "Seleccionar carpeta con los modulos VBA"
            .InitialFileName = defPath
            If .Show = False Then Exit Sub
            Dim folder As String
            folder = .SelectedItems(1)
        End With
        If Right(folder, 1) <> "\" Then folder = folder & "\"

        Dim ext As Variant, fn As String
        For Each ext In Array("*.bas", "*.cls", "*.frm")
            fn = Dir(folder & ext)
            Do While fn <> ""
                ReDim Preserve selFiles(selCount)
                selFiles(selCount) = folder & fn
                selCount = selCount + 1
                fn = Dir()
            Loop
        Next ext

        If selCount = 0 Then
            MsgBox "No se encontraron ficheros .bas/.cls/.frm en la carpeta seleccionada.", _
                   vbInformation, "Importar Modulos VBA"
            Exit Sub
        End If

    Else
        ' Picker de ficheros con multi-seleccion
        Dim fdFiles As FileDialog
        Set fdFiles = Application.FileDialog(msoFileDialogFilePicker)
        With fdFiles
            .Title = "Seleccionar modulos a importar  (Ctrl+clic para varios)"
            .InitialFileName = defPath
            .AllowMultiSelect = True
            .Filters.Clear
            .Filters.Add "Modulos VBA", "*.bas; *.cls; *.frm"
            .Filters.Add "Todos los ficheros", "*.*"
            If .Show = False Then Exit Sub
            selCount = .SelectedItems.Count
            If selCount = 0 Then Exit Sub
            ReDim selFiles(selCount - 1)
            Dim i As Long
            For i = 1 To selCount
                selFiles(i - 1) = .SelectedItems(i)
            Next i
        End With
    End If

    ' --- Importar / reemplazar cada fichero ---
    Application.ScreenUpdating = False

    Dim nNew As Long, nReplaced As Long, nUpdated As Long, nSkipped As Long, nErr As Long
    Dim logText As String
    nNew = 0: nReplaced = 0: nUpdated = 0: nSkipped = 0: nErr = 0: logText = ""

    Dim j As Long, res As String, label As String
    For j = 0 To selCount - 1
        res = Fnc_Import_One(vbp, selFiles(j), label)
        Select Case res
            Case RES_NEW:      nNew = nNew + 1
            Case RES_REPLACED: nReplaced = nReplaced + 1
            Case RES_UPDATED:  nUpdated = nUpdated + 1
            Case RES_SKIPPED:  nSkipped = nSkipped + 1
            Case Else:         nErr = nErr + 1
        End Select
        logText = logText & res & "  " & label & vbCrLf
    Next j

    Application.ScreenUpdating = True

    ' --- Resumen final ---
    Dim warnSelf As String
    If nSkipped > 0 Then
        warnSelf = vbCrLf & "AVISO: " & MY_MODULE_NAME & " fue omitido (no puede reemplazarse a si mismo)." & vbCrLf & _
                            "Para actualizarlo: Alt+F11 -> clic derecho -> Quitar modulo -> Importar fichero." & vbCrLf
    End If

    MsgBox "Importacion completada:" & vbCrLf & vbCrLf & _
           "  Nuevos importados:         " & nNew & vbCrLf & _
           "  Reemplazados:              " & nReplaced & vbCrLf & _
           "  Actualizados (hoja/libro): " & nUpdated & vbCrLf & _
           "  Omitidos (modulo activo):  " & nSkipped & vbCrLf & _
           "  Errores:                   " & nErr & vbCrLf & _
           warnSelf & vbCrLf & logText, _
           IIf(nErr > 0, vbExclamation, vbInformation), "Importar Modulos VBA"

End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<


'###################################################################################################################################
Private Function Fnc_Import_One(ByVal vbp As Object, _
                                ByVal sPath As String, _
                                ByRef outLabel As String) As String
' ==================================================================================================================================
' Importa o reemplaza un unico fichero de modulo en el VBProject dado.
' Retorna una de las constantes RES_*; outLabel recibe el texto descriptivo del resultado.

    Dim vbName    As String
    Dim fn        As String
    Dim existing  As Object
    Dim eNum      As Long
    Dim eDesc     As String

    ' Verificar existencia del fichero
    If Dir(sPath) = "" Then
        outLabel = "No encontrado: " & sPath
        Fnc_Import_One = RES_ERROR
        Exit Function
    End If

    ' Obtener VB_Name del fichero (primeras 15 lineas); fallback al nombre del fichero
    vbName = Fnc_GetVBName(sPath)
    If vbName = "" Then
        fn = Mid(sPath, InStrRev(sPath, "\") + 1)
        vbName = Left(fn, InStrRev(fn, ".") - 1)
    End If

    ' Guardia: nunca eliminar el modulo que esta ejecutandose en este momento
    If vbName = MY_MODULE_NAME Then
        outLabel = vbName & "  (modulo activo - omitido para evitar crash)"
        Fnc_Import_One = RES_SKIPPED
        Exit Function
    End If

    ' Buscar componente existente en el proyecto
    Set existing = Nothing
    On Error Resume Next
    Set existing = vbp.VBComponents(vbName)
    On Error GoTo 0

    If existing Is Nothing Then
        ' --- Modulo nuevo: importar directamente ---
        On Error Resume Next
        vbp.VBComponents.Import sPath
        eNum = Err.Number
        eDesc = Err.Description
        On Error GoTo 0
        If eNum <> 0 Then
            outLabel = vbName & "  [Error " & eNum & ": " & eDesc & "]"
            Fnc_Import_One = RES_ERROR
        Else
            outLabel = vbName
            Fnc_Import_One = RES_NEW
        End If

    ElseIf existing.Type = COMP_DOCUMENT Then
        ' --- Modulo de hoja/libro: no se puede quitar, sustituir su codigo ---
        If Fnc_Update_Doc_Module(existing, sPath) Then
            outLabel = vbName & "  (tipo Document)"
            Fnc_Import_One = RES_UPDATED
        Else
            outLabel = vbName & "  [Fallo al actualizar modulo Document]"
            Fnc_Import_One = RES_ERROR
        End If

    Else
        ' --- Modulo normal (estandar/clase/formulario): quitar y reimportar ---
        On Error Resume Next
        vbp.VBComponents.Remove existing
        vbp.VBComponents.Import sPath
        eNum = Err.Number
        eDesc = Err.Description
        On Error GoTo 0
        If eNum <> 0 Then
            outLabel = vbName & "  [Error " & eNum & ": " & eDesc & "]"
            Fnc_Import_One = RES_ERROR
        Else
            outLabel = vbName
            Fnc_Import_One = RES_REPLACED
        End If
    End If

End Function
' -------------------------------------------------------------------------------------------------------------------------------<<<


'###################################################################################################################################
Private Function Fnc_GetVBName(ByVal sPath As String) As String
' ==================================================================================================================================
' Lee las primeras 15 lineas del fichero y extrae el valor de "Attribute VB_Name = ..."
' Devuelve cadena vacia si no se encuentra o hay error.

    Dim f     As Integer
    Dim sLine As String
    Dim p1    As Integer, p2 As Integer

    On Error GoTo ErrExit
    f = FreeFile
    Open sPath For Input As #f
    Dim i As Integer
    For i = 1 To 15
        If EOF(f) Then Exit For
        Line Input #f, sLine
        If InStr(1, sLine, "Attribute VB_Name") > 0 Then
            p1 = InStr(sLine, Chr(34))
            p2 = InStrRev(sLine, Chr(34))
            If p1 > 0 And p2 > p1 Then Fnc_GetVBName = Mid(sLine, p1 + 1, p2 - p1 - 1)
            Exit For
        End If
    Next i
    Close #f
    Exit Function
ErrExit:
    On Error Resume Next
    Close #f
    On Error GoTo 0
End Function
' -------------------------------------------------------------------------------------------------------------------------------<<<


'###################################################################################################################################
Private Function Fnc_Update_Doc_Module(ByVal comp As Object, ByVal sPath As String) As Boolean
' ==================================================================================================================================
' Para modulos tipo Document (hoja/libro), imposible Remove+Import.
' Lee el fichero, omite la cabecera VBA (VERSION/BEGIN..END/Attribute) y reemplaza
' el codigo del CodeModule con el contenido restante.

    Dim f        As Integer
    Dim sLine    As String
    Dim code     As String
    Dim inHeader As Boolean
    Dim inBegin  As Boolean
    inHeader = True
    inBegin  = False

    On Error GoTo ErrExit
    f = FreeFile
    Open sPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, sLine
        If inHeader Then
            If Left(sLine, 5) = "BEGIN" Then
                inBegin = True                          ' Inicio del bloque BEGIN...END
            ElseIf inBegin Then
                If Trim(sLine) = "END" Then inBegin = False   ' Fin del bloque
                ' Resto de lineas dentro de BEGIN...END se descartan
            Else
                Select Case True
                    Case Left(sLine, 9) = "Attribute": ' lineas Attribute VB_* -> descartar
                    Case Left(sLine, 7) = "VERSION":   ' linea VERSION -> descartar
                    Case Trim(sLine) = "":             ' lineas en blanco de cabecera -> descartar
                    Case Else: inHeader = False        ' primera linea de codigo real
                End Select
            End If
        End If
        If Not inHeader Then code = code & sLine & vbCrLf
    Loop
    Close #f

    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(code) > 0 Then .InsertLines 1, code
    End With

    Fnc_Update_Doc_Module = True
    Exit Function
ErrExit:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    Fnc_Update_Doc_Module = False
End Function
' -------------------------------------------------------------------------------------------------------------------------------<<<
