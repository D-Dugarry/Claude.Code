Attribute VB_Name = "Rut_VBA_Moduls_Export"
Option Explicit

'==================================================================================================
' Rut_VBA_Moduls_Export
'
' Exporta todos los módulos VBA del libro activo a una carpeta del disco.
'
' PROCEDIMIENTOS:
'   Rut_VBA_Export_Moduls          - Exporta todos los componentes VBA (público)
'   Rut_Borrar_Contenido_Carpeta   - Borra todos los ficheros de una carpeta (privado)
'
' COMPONENTES EXPORTADOS:
'   Tipo 1   vbext_ct_StdModule   -> .bas  (módulos estándar)
'   Tipo 2   vbext_ct_ClassModule -> .cls  (módulos de clase)
'   Tipo 3   vbext_ct_MSForm      -> .frm  (formularios)
'   Tipo 100 vbext_ct_Document    -> .cls  (hojas y ThisWorkbook)
'
' CARPETA PROPUESTA:
'   [Ruta del libro]\[NombreLibro]_VBA-Moduls\
'   El usuario puede aceptarla, cambiarla manualmente o cancelar.
'   Si la carpeta ya existe y tiene ficheros, se ofrece borrar el contenido antes de exportar.
'
' USO:
'   Ejecutar Rut_VBA_Export_Moduls desde el editor VBA (F5) o desde botón en la hoja.
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' Rut_VBA_Export_Moduls
'   Exporta todos los componentes VBA del libro activo a una carpeta seleccionada por el usuario.
'   Propone como destino: [RutaLibro]\[NombreLibro]_VBA-Moduls\
'   Si la carpeta no existe la crea. Si ya tiene ficheros, ofrece borrarlos antes de exportar.
'--------------------------------------------------------------------------------------------------
Public Sub Rut_VBA_Export_Moduls()

    Dim VBComp      As Object
    Dim RutaDest    As String
    Dim RutaProp    As String
    Dim Extension   As String
    Dim Cont        As Integer
    Dim NomFich     As String
    Dim NomLibro    As String
    Dim Respuesta   As Integer

    ' --- Carpeta propuesta: misma ruta que el libro, sufijo "_VBA-Moduls" ---
    NomLibro = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    RutaProp = ThisWorkbook.Path & "\" & NomLibro & "_VBA-Moduls"

    ' --- Ofrecer carpeta propuesta o personalizada ---
    Respuesta = MsgBox("Carpeta de exportación propuesta:" & vbNewLine & vbNewLine & _
                       RutaProp & vbNewLine & vbNewLine & _
                       "¿Usar esta carpeta?" & vbNewLine & _
                       "   Sí     = Usar la carpeta propuesta" & vbNewLine & _
                       "   No     = Introducir una carpeta manualmente" & vbNewLine & _
                       "   Cancelar = Salir", _
                       vbQuestion + vbYesNoCancel, "Exportar módulos VBA")

    Select Case Respuesta
        Case vbYes
            RutaDest = RutaProp
        Case vbNo
            RutaDest = InputBox("Introduce la ruta de la carpeta de destino:", _
                                "Exportar módulos VBA", RutaProp)
            If RutaDest = "" Then
                MsgBox "Operación anulada.", vbInformation
                Exit Sub
            End If
        Case vbCancel
            Exit Sub
    End Select

    ' Asegurar que la ruta termina en "\"
    If Right(RutaDest, 1) <> "\" Then RutaDest = RutaDest & "\"

    ' --- Verificar / crear carpeta ---
    If Dir(RutaDest, vbDirectory) = "" Then
        MkDir RutaDest
    Else
        If Dir(RutaDest & "*.*") <> "" Then
            Respuesta = MsgBox("La carpeta ya existe y contiene archivos:" & vbNewLine & vbNewLine & _
                               RutaDest & vbNewLine & vbNewLine & _
                               "¿Borrar el contenido antes de exportar?", _
                               vbExclamation + vbYesNo, "Carpeta no vacía")
            If Respuesta = vbYes Then
                Call Rut_Borrar_Contenido_Carpeta(RutaDest)
            End If
            ' vbNo: continuar sin borrar (los ficheros existentes se sobreescriben)
        End If
    End If

    ' --- Exportar todos los componentes VBA ---
    Cont = 0

    For Each VBComp In ThisWorkbook.VBProject.VBComponents

        Select Case VBComp.Type
            Case 1      ' vbext_ct_StdModule    -> módulos estándar
                Extension = ".bas"
            Case 2      ' vbext_ct_ClassModule  -> módulos de clase
                Extension = ".cls"
            Case 3      ' vbext_ct_MSForm       -> formularios
                Extension = ".frm"
            Case 100    ' vbext_ct_Document     -> hojas y ThisWorkbook
                Extension = ".cls"
            Case Else
                Extension = ""
        End Select

        If Extension <> "" Then
            NomFich = RutaDest & VBComp.Name & Extension
            VBComp.Export NomFich
            Cont = Cont + 1
        End If

    Next VBComp

    MsgBox "Exportación completada." & vbNewLine & vbNewLine & _
           Cont & " módulos exportados en:" & vbNewLine & RutaDest, _
           vbInformation, "Exportar módulos VBA"

End Sub


'--------------------------------------------------------------------------------------------------
' Rut_Borrar_Contenido_Carpeta  (privada)
'   Borra todos los ficheros de la carpeta Ruta (no borra subcarpetas).
'   Usada por Rut_VBA_Export_Moduls antes de exportar si el usuario lo confirma.
'--------------------------------------------------------------------------------------------------
Private Sub Rut_Borrar_Contenido_Carpeta(ByVal Ruta As String)

    Dim Archivo As String

    Archivo = Dir(Ruta & "*.*")
    Do While Archivo <> ""
        Kill Ruta & Archivo
        Archivo = Dir
    Loop

End Sub