Attribute VB_Name = "M_3_Load_Tabla_Adjuntos"
' ==============================================================================
' Módulo     : M_3_Load_Tabla_Adjuntos
' Proyecto   : Mailing Indicadores UA
' Autor      : Dugarry
' Descripción: Carga de ficheros adjuntos en Tb_Datos.
'
' Subrutinas públicas:
'   · Rut_Load_Tabla_Adjuntos() — Rellena Col2 de Tb_Datos con los ficheros
'                                  de una carpeta seleccionada por el usuario
' ==============================================================================
Option Explicit

' ==============================================================================
' Rut_Load_Tabla_Adjuntos
' ------------------------------------------------------------------------------
' Rellena la Col 2 de Tb_Datos con los nombres de los ficheros de una carpeta
' seleccionada por el usuario. Útil para preparar el envío: primero cargar los
' ficheros de Tablas_Indicadores/, luego pegar los emails en Col 1.
'
' Nota: EXT_FILTRO limita a ficheros .xlsx. Cambiar a "" para incluir todos.
' ==============================================================================
Sub Rut_Load_Tabla_Adjuntos()
    Const EXT_FILTRO        As String = ".xlsx"     ' "" para añadir todos los ficheros

    Dim Ruta                As String
    Dim Ob_Fichero          As Object
    Dim Ob_FileSistObjct    As Object
    Dim Ob_Ruta             As Object
    Dim NewRow              As ListRow
    Dim Tb_Tabla            As ListObject

    Application.ScreenUpdating = False

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
        If EXT_FILTRO = "" Or LCase(Right(Ob_Fichero.Name, Len(EXT_FILTRO))) = LCase(EXT_FILTRO) Then
            Set NewRow = Tb_Tabla.ListRows.Add
            NewRow.Range(2) = Ob_Fichero.Name
        End If
    Next Ob_Fichero

    Application.ScreenUpdating = True

End Sub     ' Rut_Load_Tabla_Adjuntos
' ==============================================================================
