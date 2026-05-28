Attribute VB_Name = "Rut_Wb_CopSegTimed_USB_HD"
Option Explicit

' ==================================================================================================================================
' ==================================================================================================================================
' ==================================================================================================================================
Sub Rut_WrkBook_CopSegTimed_USB_Organize()   '- Importa todos los ficheros de un directorio -------------------------
' ==================================================================================================================================
Dim FichDialog As FileDialog
Dim FichRuta As String
    '- Seleccionar Carpeta ---------------------------
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
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
Sub Rut_WrkBook_Folder_File_List(FichRuta As String)   '- Importa todos los ficheros de un directorio -------------------------
' ==================================================================================================================================
Dim pos             As Integer
Dim ShortName       As String:      ShortName = ""
Dim FichSistOjct    As New FileSystemObject    '- Ver Nota #01 -----
Dim Fichero         As File
Dim Lo_WB_List      As ListObject:      Set Lo_WB_List = Prog_WB_List.ListObjects("Tb_WB_List")
Dim Lo_WB_Names     As ListObject:      Set Lo_WB_Names = Prog_WB_List.ListObjects("Tb_WB_Names")
Dim NewRow          As ListRow
Dim NewRowLo2       As ListRow
    Application.ScreenUpdating = False
    With Prog_WB_List
        .Unprotect
        '- Preparar Tablas
        .Columns.EntireColumn.Hidden = False        ' Mostrar todas las Columnas
        .Rows.EntireRow.Hidden = False              ' Mostrar todas las Filas
        If .FilterMode Then .ShowAllData            ' Deshacer Filtros
        If Not Lo_WB_List.DataBodyRange Is Nothing Then Lo_WB_List.DataBodyRange.Delete
        If Not Lo_WB_Names.DataBodyRange Is Nothing Then Lo_WB_Names.DataBodyRange.Delete
        '- Generar Lista
        With Lo_WB_List
            For Each Fichero In FichSistOjct.GetFolder(FichRuta).Files
                Set NewRow = .ListRows.Add
                    NewRow.Range(1) = Fichero.Name
                    NewRow.Range(2) = Fichero.DateLastModified
                    NewRow.Range(3) = Round(Fichero.Size / 1024, 2) ' Tamaño en KB
                    pos = InStr(Fichero.Name, " (") - 1
                    If pos > 0 Then
                        NewRow.Range(4) = Left(Fichero.Name, pos)
                    Else
                        NewRow.Range(4) = Left(Fichero.Name, Len(Fichero.Name) - 5)
                    End If
            Next Fichero
            '- Ordeno lista por Nombre Corto y Fecha
            With .Sort
                .sortFields.Clear
                .sortFields.Add Key:=Lo_WB_List.ListColumns(4).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .sortFields.Add Key:=Lo_WB_List.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            '- Genero la 2ª Tabla que lista los Archivos con el número de copias que tiene
            Dim i   As Integer
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
End Sub     '- Rut_WrkBook_Folder_File_List
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
' Procedimiento para borrar archivos seleccionados
Sub Rut_WrkBook_CopSegTimed_Selected_Del()
' ==================================================================================================================================
    Dim Ws                  As Worksheet:           Set Ws = Prog_WB_List
    Dim Lo_WB_List          As ListObject:          Set Lo_WB_List = Ws.ListObjects("Tb_WB_List")
    Dim FolderSelected      As String
    Dim Fila                As ListRow
    Dim FileToDel           As String
    Dim FSO                 As Object       '- FSO = File Sistem Object
    Dim FilesDeleted        As Integer
    Dim Respuesta           As VbMsgBoxResult
    ' Obtener la carpeta seleccionada (guardada en rango "Ruta_CopSeg")
    FolderSelected = Ws.Range("Ruta_CopSeg").Value
        If FolderSelected = "" Then
            MsgBox "No se encontró información de la carpeta en rango 'Ruta_CopSeg'.", vbExclamation
            Exit Sub
        End If
    
    ' Verificar si hay Ficheros en la Tabla
    If Lo_WB_List.DataBodyRange Is Nothing Then
        MsgBox "No hay datos en la tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Contar filas seleccionadas
    Dim filasSeleccionadas As Range
    On Error Resume Next
    Set filasSeleccionadas = Application.Intersect(Lo_WB_List.DataBodyRange, Selection)
    On Error GoTo 0
    
    ' Verificar si hay filas seleccionadas de la Tabla
    If filasSeleccionadas Is Nothing Then
        MsgBox "Por favor, seleccione al menos una fila en la tabla de archivos.", vbExclamation
        Exit Sub
    End If
    
    ' Pedir confirmación
    Respuesta = MsgBox("¿Está seguro de que desea borrar " & filasSeleccionadas.Rows.Count & _
                       " archivo(s) de la carpeta?" & vbCrLf & vbCrLf & _
                       "Carpeta: " & FolderSelected & vbCrLf & vbCrLf & _
                       "Esta acción no se puede deshacer.", _
                       vbExclamation + vbYesNo + vbDefaultButton2, "Confirmar borrado")
    
    If Respuesta <> vbYes Then Exit Sub
    
    ' Crear objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FilesDeleted = 0
    
    ' Procesar cada fila seleccionada -----------------------------------------------------------------
    Dim FilaIndex As Long
    Dim RangoFila As Range
    
    For Each RangoFila In filasSeleccionadas.Rows
        FilaIndex = RangoFila.Row - Lo_WB_List.HeaderRowRange.Row
        
        If FilaIndex >= 1 And FilaIndex <= Lo_WB_List.ListRows.Count Then
            ' Obtener el nombre del archivo
            FileToDel = FolderSelected & "\" & Lo_WB_List.DataBodyRange(FilaIndex, 1)
            
            ' Verificar si el archivo existe y borrarlo
            If FSO.FileExists(FileToDel) Then
                On Error Resume Next
                FSO.DeleteFile FileToDel, True
                On Error GoTo 0
                
                If Not FSO.FileExists(FileToDel) Then
                    FilesDeleted = FilesDeleted + 1
                    ' Marcar la fila para eliminación (pero no eliminarla todavía)
                    Lo_WB_List.DataBodyRange(FilaIndex, 1) = "[BORRADO] " & Lo_WB_List.DataBodyRange(FilaIndex, 1)
                Else
                    MsgBox "No se pudo borrar: " & FileToDel, vbExclamation
                End If
            Else
                MsgBox "Archivo no encontrado: " & FileToDel, vbExclamation
            End If
        End If
    Next RangoFila
    
    MsgBox "Operación completada." & vbCrLf & _
           "Archivos borrados: " & FilesDeleted, vbInformation
    '- Actualizar la tabla
    Call Rut_WrkBook_Folder_File_List(FolderSelected)
    
    Set FSO = Nothing
    
End Sub     '- Rut_WrkBook_CopSegTimed_Selected_Del
' ----------------------------------------------------------------------------------------------------------------------------------

'===================================================================================================================================
Sub Rut_WrkBook_CopSegTimed_USB(Optional Tipo As String = "")  '- Guarda Copia de Este Excel con marca de tiempo en el nombre del archivo. (Tipo="Data"/"VBA", etc.) ----------
Debug.Print "Rut_WrkBook_CopSegTimed_USB,   Tipo: " & Tipo
    Dim Answer          As VbMsgBoxResult
    Dim FichNom         As String
        FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    Dim FichExt         As String
        FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
    Dim fichPath        As String
        'fichPath = ThisWorkbook.Path & "\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
        If Not Fnc_Range_Exist("APP_CopSeg_Usb_Path") Then
            MsgBox "¡¡¡ Falta crear el Range('APP_CopSeg_Usb_Path') !!!", vbExclamation, "Procedimiento: Copia de Seguridad"
            fichPath = "F:\__CopSeg Versiones Programas\" & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
        Else
            fichPath = Range("APP_CopSeg_Usb_Path") & FichNom & " " & Format(Now, "(yymmdd_hhmm)") & Tipo & FichExt
        End If
    Dim FichSelect      As Variant
        FichSelect = Application.GetSaveAsFilename(fichPath, "Excel Files (*" & FichExt & "), *" & FichExt)
        
        If FichSelect <> False Then
            On Error GoTo Finalizar
            Application.DisplayAlerts = False
            ThisWorkbook.SaveCopyAs Filename:=FichSelect 'ConflictResolution:=True     '??? no se lo que hace, habría que investigar
            Application.DisplayAlerts = True
            On Error GoTo 0
        End If
    If Fnc_Range_Exist("APP_CopSeg_Usb_Path") Then
        If Range("APP_CopSeg_Usb_Path") <> Left(FichSelect, InStrRev(FichSelect, "\")) Then
            ' Pedir confirmación
            Dim Mensage     As String
            Mensage = "¿ Cambiamos esta ruta: " & Range("APP_CopSeg_Usb_Path") & vbLf & _
                      " por esta ? " & Left(FichSelect, InStrRev(FichSelect, "\")) & vbCrLf & vbCrLf
            Answer = MsgBox(Mensage, vbExclamation + vbYesNo + vbDefaultButton2, "Proceso: Cambio de Ruta para las Copias de Seguridad.")
            If Answer = vbYes Then
                Range("APP_CopSeg_Usb_Path") = Left(FichSelect, InStrRev(FichSelect, "\"))
            End If
        End If
    End If
    Debug.Print Left(FichSelect, InStrRev(FichSelect, "\"))
    Prog__APP.Range("APP_Task_Inf") = "Copia Realizada en la Carpeta del USB: " & FichSelect & vbCrLf & String(100, "-") & vbLf & "Proceso Finalizado. " & Format(Now(), "dd-mmm-yy hh:mm")

'    ThisWorkbook.Close savechanges:=True
Finalizar:
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------

' ==================================================================================================================================
Sub Rut_WrkBooK_CopSegTimed_WB_HD()   '- Copia una Sheet concreta
' ==================================================================================================================================
Rut_Off_Functions
    Dim FichNom                             As String
                FichNom = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    Dim FichExt                             As String
                FichExt = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") + 1)
    Dim FPath                               As String
                FPath = Fnc_Format_Ruta(ThisWorkbook.Path & "/CopiaSeguridad/")
    H_Inicio = Timer                ' Para Saber el tiempo de proceso
    '- Indicar Nombre del Archivo y Ruta para almacenar --------------
    Dim IntialName As String
    Dim sFileSaveName As Variant
    IntialName = FPath & FichNom & " - " & Format(Now, "(yymmdd_hhmm)") & FichExt
            On Error GoTo GestError
            Application.DisplayAlerts = False
            ThisWorkbook.SaveCopyAs Filename:=IntialName    ', ConflictResolution:=True   '''', FileFormat:=FichExt
            Application.DisplayAlerts = True
            On Error GoTo 0
    Range("APP_CopSeg_Date") = Date
        
'    MsgBox "¡¡¡ Archivo guardado !!!", vbOKOnly, "Proceso: Copia de Seguridad"
    Form_Menu.TB_Informe = "¡¡¡ Proceso concluido con éxito !!!" & "  -.-  " & Now() & vbCrLf & _
                           "He tardado: " & Round(Timer - H_Inicio, 2) & " seg." & vbLf & vbLf & _
                           "Copia de Seguridad de Filename:= " & vbCrLf & IntialName
    GoTo Salir_Sub
GestError:
    Debug.Print "Error Rut_Exportar_La_Liquidación ", Err.Number, Err.Description, Err.Source
    Debug.Print sFileSaveName
    MsgBox "Error: Rut_Exportar_La_Liquidación " & vbLf & "Err.Number:" & Err.Number & vbLf & "Err.Description:" & Err.Description & _
            vbLf & "Filename:=" & vbCrLf & IntialName, vbExclamation + vbOKOnly, "Proceso: Copia de Seguridad."
    Form_Menu.TB_Informe = "Rut_Exportar_La_Liquidación " & vbCrLf & "Filename:=" & IntialName
Salir_Sub:
Rut_On_Functions
End Sub     ' Rut_WrkBooK_CopSegTimed_WB_HD
' ----------------------------------------------------------------------------------------------------------------------------------




''' Nota #01
'''
'''    El error "No se puede definir el tipo" en VBA ocurre porque la librería `FileSystemObject` no está referenciada en tu proyecto. Aquí te muestro cómo solucionarlo:
'''
'''    ## Solución 1: Referenciar la librería (Recomendado)
'''
'''    1. En el editor de VBA, ve a **Herramientas** ? **Referencias**
'''    2. Busca y marca la casilla: **"Microsoft Scripting Runtime"**
'''    3. Haz clic en **Aceptar**
'''
'''    ## Solución 2: Usar CreateObject (Alternativa)
'''
'''    Si prefieres no agregar la referencia, puedes usar:
'''
'''    ```vba
'''    Dim FichSistOjct As Object
'''    Set FichSistOjct = CreateObject("Scripting.FileSystemObject")
'''    ```
'''
'''    ## Ejemplo completo de uso:
'''
'''    ```vba
'''    Sub EjemploFileSystemObject()
'''        ' Método con referencia
'''        Dim FSO As New FileSystemObject
'''        Dim archivo As TextStream
'''
'''        ' Crear un archivo de texto
'''        Set archivo = FSO.CreateTextFile("C:\temp\ejemplo.txt", True)
'''        archivo.WriteLine "Hola mundo"
'''        archivo.Close
'''
'''        ' Método alternativo sin referencia
'''        Dim FSO2 As Object
'''        Set FSO2 = CreateObject("Scripting.FileSystemObject")
'''
'''        ' Verificar si existe un archivo
'''        If FSO2.FileExists("C:\temp\ejemplo.txt") Then
'''            MsgBox "El archivo existe"
'''        End If
'''    End Sub
'''    ```
'''
'''    ## Funciones comunes de FileSystemObject:
'''
'''    ```vba
'''    Sub UsosComunesFSO()
'''        Dim FSO As New FileSystemObject
'''
'''        ' Verificar existencia
'''        If FSO.FileExists("C:\ruta\archivo.txt") Then
'''            MsgBox "El archivo existe"
'''        End If
'''
'''        ' Copiar archivo
'''        FSO.CopyFile "origen.txt", "destino.txt"
'''
'''        ' Crear carpeta
'''        If Not FSO.FolderExists("C:\nueva_carpeta") Then
'''            FSO.CreateFolder "C:\nueva_carpeta"
'''        End If
'''
'''        ' Obtener información de archivo
'''        Dim archivo As File
'''        Set archivo = FSO.GetFile("archivo.txt")
'''        MsgBox "Tamaño: " & archivo.Size & " bytes"
'''    End Sub
'''    ```
'''
'''    **Recomendación**: Usa la **Solución 1** (agregar la referencia) ya que te dará acceso al IntelliSense y a todas las propiedades/métodos del objeto.
'''
'''
'''     EJEMPLO
'''     EJEMPLO
'''     EJEMPLO
'''     EJEMPLO
'''     EJEMPLO
'''
''' ==================================================================================================================================
'''     EJEMPLO: Procedimiento para actualizar la tabla (por si necesitas refrescar después de borrar)
Sub EJEMPLO_ActualizarTabla()
''' ==================================================================================================================================
    Dim Ws As Worksheet
    Dim tblArchivos As ListObject
    Dim carpetaSeleccionada As String
    Dim FSO As Object
    Dim Folder As Object
    Dim objFile As Object
    Dim i As Long
    
    Set Ws = ActiveSheet
    Set tblArchivos = Ws.ListObjects("TablaArchivos")
    carpetaSeleccionada = Ws.Range("Z1").Value
    
    If carpetaSeleccionada = "" Then
        MsgBox "No se encontró información de la carpeta.", vbExclamation
        Exit Sub
    End If
    
    ' Limpiar tabla existente (excepto encabezados)
    If tblArchivos.ListRows.Count > 0 Then
        tblArchivos.DataBodyRange.Delete
    End If
    
    ' Volver a llenar la tabla
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(carpetaSeleccionada)
    
    i = 1
    For Each objFile In Folder.Files
        tblArchivos.ListRows.Add
        tblArchivos.DataBodyRange(i, 1) = objFile.Name
        tblArchivos.DataBodyRange(i, 2) = objFile.DateLastModified
        tblArchivos.DataBodyRange(i, 3) = Round(objFile.Size / 1024, 2)
        i = i + 1
    Next objFile
    
    ' Reordenar la tabla
    With tblArchivos.Sort
        .sortFields.Clear
        .sortFields.Add Key:=tblArchivos.ListColumns("Nombre Archivo").DataBodyRange, _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .sortFields.Add Key:=tblArchivos.ListColumns("Fecha Modificación").DataBodyRange, _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    MsgBox "Tabla actualizada correctamente.", vbInformation
    Set FSO = Nothing
End Sub
''' ----------------------------------------------------------------------------------------------------------------------------------






