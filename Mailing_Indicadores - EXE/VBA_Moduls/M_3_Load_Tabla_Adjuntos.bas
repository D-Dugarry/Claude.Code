Attribute VB_Name = "M_3_Load_Tabla_Adjuntos"
' ==============================================================================
' Módulo     : M_3_Load_Tabla_Adjuntos
' Proyecto   : Mailing Indicadores UA
' Autor      : Dugarry
' Descripción: Carga de ficheros adjuntos en Tb_Datos extrayendo el email
'              del nombre del fichero.
'
' Subrutinas públicas:
'   · Rut_Load_Tabla_Adjuntos() — Rellena Tb_Datos con los ficheros de una
'                                  carpeta seleccionada por el usuario
' ==============================================================================
Option Explicit

' ==============================================================================
' Rut_Load_Tabla_Adjuntos
' ------------------------------------------------------------------------------
' Rellena Tb_Datos a partir de los ficheros de una carpeta seleccionada:
'   · Col 1 (email)   ? extrae el texto entre () del nombre y añade "@ua.es"
'   · Col 2 (fichero) ? nombre del fichero original sin modificar
'
' Convención de nombres esperada:
'   (usuario.dep) Indicadores Nombre Departamento 2023.xlsx
'    +--------+
'    Col 1: usuario.dep@ua.es
'    Col 2: (usuario.dep) Indicadores Nombre Departamento 2023.xlsx
'
' Si el fichero no contiene (), Col 1 queda vacío y Col 2 toma el nombre completo.
' La extensión a filtrar se pregunta al usuario (por defecto .xlsx, vacío = todos).
' ==============================================================================
Sub Rut_Load_Tabla_Adjuntos()

    Dim ExtFiltro           As String
    Dim Ruta                As String
    Dim Ob_Fichero          As Object
    Dim Ob_FileSistObjct    As Object
    Dim Ob_Ruta             As Object
    Dim NewRow              As ListRow
    Dim Tb_Tabla            As ListObject
    Dim sNombre             As String
    Dim sEmail              As String
    Dim posIni              As Integer
    Dim posFin              As Integer

    Application.ScreenUpdating = False

    ' --- Preguntar extensión a filtrar (por defecto .xlsx) ---
    ExtFiltro = Trim(InputBox("Extensión de ficheros a cargar:" & vbCrLf & vbCrLf & _
                              "Deja vacío para cargar todos los ficheros.", _
                              "Filtro de extensión", ".xlsx"))

    Set Tb_Tabla = DatosCorreo.ListObjects("Tb_Datos")
    Set Ob_FileSistObjct = CreateObject("Scripting.FileSystemObject")

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar Carpeta donde están los archivos"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        .Show
        If .SelectedItems.Count = 0 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        Ruta = .SelectedItems(1) & "\"
    End With

    Set Ob_Ruta = Ob_FileSistObjct.GetFolder(Ruta)

    For Each Ob_Fichero In Ob_Ruta.Files
        sNombre = Ob_Fichero.Name

        ' --- Filtro por extensión ---
        If ExtFiltro = "" Or LCase(Right(sNombre, Len(ExtFiltro))) = LCase(ExtFiltro) Then

            ' --- Extraer email de la parte entre () ---
            posIni = InStr(sNombre, "(")
            posFin = InStr(sNombre, ")")

            If posIni > 0 And posFin > posIni Then
                sEmail = Mid(sNombre, posIni + 1, posFin - posIni - 1) & "ua.es"
            Else
                sEmail = ""
            End If

            Set NewRow = Tb_Tabla.ListRows.Add
            NewRow.Range(1) = sEmail
            NewRow.Range(2) = sNombre      ' Nombre original sin modificar

        End If
    Next Ob_Fichero

    Application.ScreenUpdating = True

End Sub     ' Rut_Load_Tabla_Adjuntos
' ==============================================================================