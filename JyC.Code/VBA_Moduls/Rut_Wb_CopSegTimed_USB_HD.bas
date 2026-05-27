Attribute VB_Name = "Rut_Wb_CopSegTimed_USB_HD"
Option Explicit

'==================================================================================================
' Rut_Wb_CopSegTimed_USB_HD
'
' Gestión de copias de seguridad del libro Excel con marca de tiempo en el nombre.
'
' PROCEDIMIENTOS PRINCIPALES:
'   Rut_WrkBook_CopSegTimed_USB       - Copia manual con diálogo SaveAs ? destino USB/red
'   Rut_WrkBooK_CopSegTimed_WB_HD     - Copia automática en subcarpeta CopiaSeguridad (sin diálogo)
'
' GESTOR DE COPIAS (listado y borrado):
'   Rut_WrkBook_CopSegTimed_USB_Organize  - Selecciona carpeta y lista las copias existentes
'   Rut_WrkBook_Folder_File_List           - Rellena Tb_WB_List y Tb_WB_Names con los ficheros
'   Rut_WrkBook_CopSegTimed_Selected_Del   - Borra los ficheros seleccionados en Tb_WB_List
'
' FORMATO DEL NOMBRE DE COPIA:
'   NombreOriginal (yymmdd_hhmm)[Tipo].xlsm
'   El "Nombre Corto" se extrae quitando el sufijo " (yymmdd_hhmm)" para agrupar versiones.
'
' DEPENDENCIAS (otros módulos):
'   Fnc_Range_Exist(nombre)   - Verifica si existe un rango con nombre
'   Fnc_Format_Ruta(ruta)     - Normaliza separadores de ruta
'   Rut_Off_Functions         - Desactiva eventos/cálculo/actualización
'   Rut_On_Functions          - Restaura eventos/cálculo/actualización
'   H_Inicio                  - Variable global Long/Single para medición de tiempo
'
' REFERENCIAS VBA REQUERIDAS:
'   Microsoft Scripting Runtime   (para FileSystemObject en Rut_WrkBook_Folder_File_List)
'   Ver Nota #01 al final del módulo si aparece error "No se puede definir el tipo"
'
' RANGOS CON NOMBRE REQUERIDOS:
'   APP_CopSeg_Usb_Path  (Prog__APP)    - Ruta por defecto para copias USB/red (con "\" final)
'   APP_CopSeg_Date      (Prog__APP)    - Fecha de la última copia automática en HD
'   APP_Task_Inf         (Prog__APP)    - Log de la última operación realizada
'   Ruta_CopSeg          (Prog_WB_List) - Ruta de la carpeta analizada por el gestor de copias
'
' HOJAS / OBJETOS REQUERIDOS:
'   Prog__APP     - Hoja de configuración de la aplicación
'   Prog_WB_List  - Hoja con las tablas Tb_WB_List y Tb_WB_Names
'   Form_Menu     - Formulario con campo TB_Informe (solo Rut_WrkBooK_CopSegTimed_WB_HD)
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' Rut_WrkBook_CopSegTimed_USB_Organize
'   Abre un diálogo para seleccionar la carpeta donde están las copias de seguridad,
'   guarda la ruta en Ruta_CopSeg y llama a Rut_WrkBook_Folder_File_List para listarlas.
'--------------------------------------------------------------------------------------------------
Sub Rut_WrkBook_CopSegTimed_USB_Organize()

    Dim FichDialog As FileDialog
    Dim FichRuta As String

    Set FichDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With FichDialog
        .Title = "Choose Folder where you have excel files"
        .ButtonName = "Choose"
        If .Show = True Then
            If .SelectedItems.Count > 0 Then FichRuta = .SelectedItems(1)
        End If
    End With

    If FichRuta = "" Then GoTo Final Else Range("Ruta_CopSeg") = FichRuta
    Call Rut_WrkBook_Folder_File_List(FichRuta)

Final:
End Sub


'--------------------------------------------------------------------------------------------------
' Rut_WrkBook_Folder_File_List
'   Lista todos los ficheros de la carpeta FichRuta en la hoja Prog_WB_List:
'     Tb_WB_List  : nombre, fecha modificación, tamaño (KB), nombre corto (sin timestamp)
'                   Ordenada por nombre corto ASC, fecha DESC (la copia más reciente arriba)
'     Tb_WB_Names : nombre corto único + número de copias existentes de ese fichero
'
'   REQUIERE referencia: Microsoft Scripting Runtime (ver Nota #01)
'--------------------------------------------------------------------------------------------------
Sub Rut_WrkBook_Folder_File_List(FichRuta As String)

    Dim pos             As Integer
    Dim ShortName       As String:          ShortName = ""
    Dim FichSistOjct    As New FileSystemObject     ' Ver Nota #01
    Dim Fichero         As File
    Dim Lo_WB_List      As ListObject:      Set Lo_WB_List  = Prog_WB_List.ListObjects("Tb_WB_List")
    Dim Lo_WB_Names     As ListObject:      Set Lo_WB_Names = Prog_WB_List.ListObjects("Tb_WB_Names")
    Dim NewRow          As ListRow
    Dim NewRowLo2       As ListRow

    Application.ScreenUpdating = False

    With Prog_WB_List
        .Unprotect
        .Columns.EntireColumn.Hidden = False
        .Rows.EntireRow.Hidden = False
        If .FilterMode Then .ShowAllData
        If Not Lo_WB_List.DataBodyRange  Is Nothing Then Lo_WB_List.DataBodyRange.Delete
        If Not Lo_WB_Names.DataBodyRange Is Nothing Then Lo_WB_Names.DataBodyRange.Delete

        With Lo_WB_List
            ' --- Rellenar Tb_WB_List ---
            For Each Fichero In FichSistOjct.GetFolder(FichRuta).Files
                Set NewRow = .ListRows.Add
                NewRow.Range(1) = Fichero.Name
                NewRow.Range(2) = Fichero.DateLastModified
                NewRow.Range(3) = Round(Fichero.Size / 1024, 2)     ' Tamaño en KB
                ' Nombre corto: quitar el sufijo " (yymmdd_hhmm)" si existe
                pos = InStr(Fichero.Name, " (") - 1
                If pos > 0 Then
                    NewRow.Range(4) = Left(Fichero.Name, pos)
                Else
                    NewRow.Range(4) = Left(Fichero.Name, Len(Fichero.Name) - 5)
                End If
            Next Fichero

            ' --- Ordenar: nombre corto ASC, fecha DESC ---
            With .Sort
                .SortFields.Clear
                .SortFields.Add Key:=Lo_WB_List.ListColumns(4).Range, SortOn:=xlSortOnValues, Order:=xlAscending,  DataOption:=xlSortNormal
                .SortFields.Add Key:=Lo_WB_List.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With

            ' --- Generar Tb_WB_Names: nombre corto único + nº de copias ---
            Dim i As Integer
            For i = 1 To Lo_WB_List.ListRows.Count
                Set NewRow = .ListRows(i)
                If ShortName <> NewRow.Range(4) Then
                    ShortName = NewRow.Range(4)
                    Set NewRowLo2 = Lo_WB_Names.ListRows.Add
                    NewRowLo2.Range(1) = ShortName
                    NewRowLo2.Range(2) = 1
                Else
                    NewRowLo2.Range(2) = NewRowLo2.Range(2) + 1
                End If
            Next i
        End With
    End With

    Range("b1").Select
    Application.ScreenUpdating = True
    MsgBox "All Files Listed successfully!", vbInformation

End Sub


'--------------------------------------------------------------------------------------------------
' Rut_WrkBook_CopSegTimed_Selected_Del
'   Borra del disco los ficheros correspondientes a las filas seleccionadas en Tb_WB_List.
'   La carpeta origen se lee de Ruta_CopSeg. Pide confirmación antes de borrar.
'   Tras borrar, recarga la lista con Rut_WrkBook_Folder_File_List.
'--------------------------------------------------------------------------------------------------
Sub Rut_WrkBook_CopSegTimed_Selected_Del()

    Dim Ws              As Worksheet:   Set Ws = Prog_WB_List
    Dim Lo_WB_List      As ListObject:  Set Lo_WB_List = Ws.ListObjects("Tb_WB_List")
    Dim FolderSelected  As String
    Dim FileToDel       As String
    Dim FSO             As Object
    Dim FilesDeleted    As Integer
    Dim Respuesta       As VbMsgBoxResult
    Dim FilaIndex       As Long
    Dim RangoFila       As Range
    Dim filasSeleccionadas As Range

    FolderSelected = Ws.Range("Ruta_CopSeg").Value
    If FolderSelected = "" Then
        MsgBox "No se encontró información de la carpeta en rango 'Ruta_CopSeg'.", vbExclamation
        Exit Sub
    End If

    If Lo_WB_List.DataBodyRange Is Nothing Then
        MsgBox "No hay datos en la tabla.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set filasSeleccionadas = Application.Intersect(Lo_WB_List.DataBodyRange, Selection)
    On Error GoTo 0

    If filasSeleccionadas Is Nothing Then
        MsgBox "Por favor, seleccione al menos una fila en la tabla de archivos.", vbExclamation
        Exit Sub
    End If

    Respuesta = MsgBox("¿Está seguro de que desea borrar " & filasSeleccionadas.Rows.Count & _
                       " archivo(s) de la carpeta?" & vbCrLf & vbCrLf & _
                       "Carpeta: " & FolderSelected & vbCrLf & vbCrLf & _
                       "Esta acción no se puede deshacer.", _
                       vbExclamation + vbYesNo + vbDefaultButton2, "Confirmar borrado")
    If Respuesta <> vbYes Then Exit Sub

    Set FSO = CreateObject("Scripting.FileSystemObject")
    FilesDeleted = 0

    For Each RangoFila In filasSeleccionadas.Rows
        FilaIndex = RangoFila.Row - Lo_WB_List.HeaderRowRange.Row
        If FilaIndex >= 1 And FilaIndex <= Lo_WB_List.ListRows.Count Then
            FileToDel = FolderSelected & "\" & Lo_WB_List.DataBodyRange(FilaIndex, 1)
            If FSO.FileExists(FileToDel) Then
                On Error Resume Next
                FSO.DeleteFile FileToDel, True
                On Error GoTo 0
                If Not FSO.FileExists(FileToDel) Then
                    FilesDeleted = FilesDeleted + 1
                    Lo_WB_List.DataBodyRange(FilaIndex, 1) = "[BORRADO] " & Lo_WB_List.DataBodyRange(FilaIndex, 1)
                Else
                    MsgBox "No se pudo borrar: " & FileToDel, vbExclamation
                End If
            Else
                MsgBox "Archivo no encontrado: " & FileToDel, vbExclamation
            End If
        End If
    Next RangoFila

    MsgBox "Operación completada." & vbCrLf & "Archivos borrados: " & FilesDeleted, vbInformation
    Call Rut_WrkBook_Folder_File_List(FolderSelected)
    Set FSO = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Rut_WrkBook_CopSegTimed_USB
'   Guarda una copia del libro con marca de tiempo mediante diálogo SaveAs.
'   La ruta por defecto se lee de APP_CopSeg_Usb_Path; si el usuario elige otra carpeta,
'   ofrece actualizar ese rango.
'
'   Tipo  : sufijo opcional al nombre (p.ej. "Data", "VBA"). Por defecto "".
'   Nombre generado: NombreOriginal (yymmdd_hhmm)[Tipo].xlsm
'--------------------------------------------------------------------------------------------------
Sub Rut_WrkBook_CopSegTimed_USB(Optional Tipo As String = "")

    Debug.Print "Rut_WrkBook_CopSegTimed_USB,   Tipo: " & Tipo

    Dim Answer      As VbMsgBoxResult
    Dim FichNom     As String:  FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    Dim FichExt     As String:  FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
    Dim fichPath    As String
    Dim FichSelect  As Variant

    If Not Fnc_Range_Exist("APP_CopSeg_Usb_Path") Then
        MsgBox "¡¡¡ Falta crear el Range('APP_CopSeg_Usb_Path') !!!", vbExclamation, "Copia de Seguridad"
        fichPath = "F:\__CopSeg Versiones Programas\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
    Else
        fichPath = Range("APP_CopSeg_Usb_Path") & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
    End If

    FichSelect = Application.GetSaveAsFilename(fichPath, "Excel Files (*" & FichExt & "), *" & FichExt)

    If FichSelect <> False Then
        On Error GoTo Finalizar
        Application.DisplayAlerts = False
        ThisWorkbook.SaveCopyAs Filename:=FichSelect
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If

    ' Ofrecer actualizar la ruta si el usuario guardó en una carpeta diferente
    If Fnc_Range_Exist("APP_CopSeg_Usb_Path") Then
        If Range("APP_CopSeg_Usb_Path") <> Left(FichSelect, InStrRev(FichSelect, "\")) Then
            Dim Mensage As String
            Mensage = "¿ Cambiamos esta ruta: " & Range("APP_CopSeg_Usb_Path") & vbLf & _
                      " por esta ? " & Left(FichSelect, InStrRev(FichSelect, "\")) & vbCrLf & vbCrLf
            Answer = MsgBox(Mensage, vbExclamation + vbYesNo + vbDefaultButton2, "Cambio de Ruta para Copias de Seguridad")
            If Answer = vbYes Then
                Range("APP_CopSeg_Usb_Path") = Left(FichSelect, InStrRev(FichSelect, "\"))
            End If
        End If
    End If

    Debug.Print Left(FichSelect, InStrRev(FichSelect, "\"))
    Prog__APP.Range("APP_Task_Inf") = "Copia Realizada en la Carpeta del USB: " & FichSelect & vbCrLf & _
                                      String(100, "-") & vbLf & _
                                      "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm")
Finalizar:
End Sub


'--------------------------------------------------------------------------------------------------
' Rut_WrkBooK_CopSegTimed_WB_HD
'   Guarda automáticamente una copia del libro en la subcarpeta "CopiaSeguridad" ubicada
'   junto al propio libro, sin mostrar ningún diálogo.
'   Actualiza APP_CopSeg_Date con la fecha de la copia y muestra el resultado en Form_Menu.
'
'   Nombre generado: NombreOriginal - (yymmdd_hhmm).xlsm
'--------------------------------------------------------------------------------------------------
Sub Rut_WrkBooK_CopSegTimed_WB_HD()

    Rut_Off_Functions

    Dim FichNom     As String:  FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    Dim FichExt     As String:  FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
    Dim FPath       As String:  FPath = Fnc_Format_Ruta(ThisWorkbook.Path & "/CopiaSeguridad/")
    Dim IntialName  As String:  IntialName = FPath & FichNom & " - " & Format(Now, "(yymmdd_hhmm)") & FichExt

    H_Inicio = Timer

    On Error GoTo GestError
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs Filename:=IntialName
    Application.DisplayAlerts = True
    On Error GoTo 0

    Range("APP_CopSeg_Date") = Date

    Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!  -.-  " & Now() & vbCrLf & _
                           "He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & vbLf & vbLf & _
                           "Copia de Seguridad: " & vbCrLf & IntialName
    GoTo Salir_Sub

GestError:
    Debug.Print "Error Rut_WrkBooK_CopSegTimed_WB_HD", Err.Number, Err.Description, Err.Source
    MsgBox "Error: Rut_WrkBooK_CopSegTimed_WB_HD" & vbLf & _
           "Err.Number: " & Err.Number & vbLf & _
           "Err.Description: " & Err.Description & vbLf & _
           "Filename: " & vbCrLf & IntialName, vbExclamation + vbOKOnly, "Copia de Seguridad"
    Form_Menu.TB_Informe = "Error Rut_WrkBooK_CopSegTimed_WB_HD" & vbCrLf & "Filename: " & IntialName

Salir_Sub:
    Rut_On_Functions

End Sub


' ==================================================================================================
' Nota #01 - FileSystemObject: referencia requerida
'
'   Si aparece el error "No se puede definir el tipo" en la declaración de FileSystemObject,
'   la librería Microsoft Scripting Runtime no está referenciada en el proyecto.
'
'   Solución recomendada:
'     Editor VBA ? Herramientas ? Referencias ? marcar "Microsoft Scripting Runtime" ? Aceptar
'
'   Alternativa sin referencia (late binding, sin IntelliSense):
'     Dim FichSistOjct As Object
'     Set FichSistOjct = CreateObject("Scripting.FileSystemObject")
' ==================================================================================================