Attribute VB_Name = "Rut_VBA_Moduls_Export"
Option Explicit

'==================================================================================================
' M_000_Export_VBA
'   Exporta todos los modulos VBA del libro activo a una carpeta
'
' USO:
'   Ejecutar Rut_VBA_Export_Moduls desde el editor VBA (F5) o desde la hoja
'==================================================================================================

Public Sub Rut_VBA_Export_Moduls()

    Dim VBComp      As Object
    Dim RutaDest    As String
    Dim RutaProp    As String
    Dim Extension   As String
    Dim Cont        As Integer
    Dim NomFich     As String
    Dim NomLibro    As String
    Dim Respuesta   As Integer

    ' --- Construir nombre de carpeta propuesto ---
    NomLibro = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    RutaProp = ThisWorkbook.Path & "\" & NomLibro & "_VBA-Moduls"

    ' --- Ofrecer nombre propuesto o personalizado ---
    Respuesta = MsgBox("Carpeta de exportacion propuesta:" & vbNewLine & vbNewLine & _
                       RutaProp & vbNewLine & vbNewLine & _
                       "¿Usar esta carpeta?" & vbNewLine & _
                       "   Si  = Usar la carpeta propuesta" & vbNewLine & _
                       "   No  = Introducir una carpeta manualmente" & vbNewLine & _
                       "   Cancelar = Salir", _
                       vbQuestion + vbYesNoCancel, "Exportar modulos VBA")

    Select Case Respuesta
        Case vbYes
            RutaDest = RutaProp
        Case vbNo
            RutaDest = InputBox("Introduce la ruta de la carpeta de destino:", _
                                "Exportar modulos VBA", RutaProp)
            If RutaDest = "" Then
                MsgBox "Operacion anulada.", vbInformation
                Exit Sub
            End If
        Case vbCancel
            Exit Sub
    End Select

    ' Asegurar que la ruta termina en "\"
    If Right(RutaDest, 1) <> "\" Then RutaDest = RutaDest & "\"

    ' --- Verificar si la carpeta existe ---
    If Dir(RutaDest, vbDirectory) = "" Then
        ' No existe: crear
        MkDir RutaDest
    Else
        ' Existe: comprobar si esta vacia
        If Dir(RutaDest & "*.*") <> "" Then
            ' No esta vacia: preguntar si borrar contenido
            Respuesta = MsgBox("La carpeta ya existe y contiene archivos:" & vbNewLine & vbNewLine & _
                               RutaDest & vbNewLine & vbNewLine & _
                               "¿Borrar el contenido antes de exportar?", _
                               vbExclamation + vbYesNo, "Carpeta no vacia")
            If Respuesta = vbYes Then
                Call Rut_Borrar_Contenido_Carpeta(RutaDest)
            ElseIf Respuesta = vbNo Then
                ' Continuar sin borrar (los archivos existentes se sobreescribiran)
            End If
        End If
    End If

    ' --- Exportar todos los componentes VBA ---
    Cont = 0

    For Each VBComp In ThisWorkbook.VBProject.VBComponents

        Select Case VBComp.Type
            Case 1      ' vbext_ct_StdModule    -> .bas
                Extension = ".bas"
            Case 2      ' vbext_ct_ClassModule  -> .cls
                Extension = ".cls"
            Case 3      ' vbext_ct_MSForm       -> .frm
                Extension = ".frm"
            Case 100    ' vbext_ct_Document     -> hojas y ThisWorkbook -> .cls
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

    MsgBox "Exportacion completada." & vbNewLine & vbNewLine & _
           Cont & " modulos exportados en:" & vbNewLine & RutaDest, _
           vbInformation, "Exportar modulos VBA"

End Sub


'--------------------------------------------------------------------------------------------------
' Rut_Borrar_Contenido_Carpeta  (privada)
'   Borra todos los archivos de una carpeta (no borra subcarpetas)
'--------------------------------------------------------------------------------------------------
Private Sub Rut_Borrar_Contenido_Carpeta(ByVal Ruta As String)

    Dim Archivo As String

    Archivo = Dir(Ruta & "*.*")
    Do While Archivo <> ""
        Kill Ruta & Archivo
        Archivo = Dir
    Loop

End Sub
