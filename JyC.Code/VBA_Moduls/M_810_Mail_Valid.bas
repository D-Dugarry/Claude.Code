Attribute VB_Name = "M_810_Mail_Valid"
Option Explicit

'==================================================================================================
' M_810_Mail_Valid  -  Validación de direcciones de correo electrónico
'
' USO:
'   Fnc_Valid_Email_Multi(Destinatario)
'       -> Valida una o varias direcciones separadas por ";"
'       -> Si alguna no es válida, pregunta si cancelar o quitar esa dirección y seguir
'       -> Devuelve TRUE si se debe continuar el envío (con la lista ya limpia)
'       -> Devuelve FALSE si el usuario cancela, o no quedan destinatarios válidos
'       -> IMPORTANTE: Destinatario es ByRef, puede quedar modificado (sin las dir. erróneas)
'
'   Fnc_Valid_Email(Email)
'       -> Valida una sola dirección (sin MsgBox)
'       -> Útil para validaciones silenciosas
'
' EJEMPLOS:
'   If Not Fnc_Valid_Email_Multi(Destinatario) Then Exit Sub
'   ' Tras la llamada, Destinatario solo contiene las direcciones válidas
'
'   If Fnc_Valid_Email("user@domain.com") Then ...
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' Fnc_Valid_Email_Multi
'   Valida una o varias direcciones separadas por ";"
'   Si alguna no es válida, pregunta si cancelar o quitar esa dirección y seguir
'   Destinatario es ByRef: puede quedar modificado sin las direcciones erróneas
'   Devuelve TRUE si hay al menos una dirección válida y se debe continuar
'   Devuelve FALSE si el usuario cancela o no quedan destinatarios
'--------------------------------------------------------------------------------------------------
Public Function Fnc_Valid_Email_Multi(ByRef Destinatario As String) As Boolean

    Dim arrDest()       As String
    Dim i               As Integer
    Dim MailItem        As String
    Dim DestValidos     As String
    Dim Respuesta       As Integer

    Fnc_Valid_Email_Multi = False

    If Trim(Destinatario) = "" Then
        MsgBox "El destinatario está vacío.", vbExclamation, "Error Email"
        Exit Function
    End If

    arrDest = Split(Destinatario, ";")
    DestValidos = ""

    For i = 0 To UBound(arrDest)
        MailItem = Trim(arrDest(i))
        If MailItem <> "" Then
            If Fnc_Valid_Email(MailItem) Then
                ' Dirección correcta: la añadimos a la lista válida
                If DestValidos <> "" Then DestValidos = DestValidos & ";"
                DestValidos = DestValidos & MailItem
            Else
                ' Dirección errónea: preguntar qué hacer
                Respuesta = MsgBox( _
                    "Dirección de correo no válida:" & vbNewLine & vbNewLine & _
                    "   """ & MailItem & """" & vbNewLine & vbNewLine & _
                    "¿Qué deseas hacer?", _
                    vbExclamation + vbYesNo, _
                    "Dirección errónea — Sí=Quitar y seguir  /  No=Cancelar envío")

                If Respuesta = vbNo Then
                    ' Cancelar envío completo
                    Exit Function
                End If
                ' vbYes: quitar esta dirección y continuar con las demás
            End If
        End If
    Next i

    ' Comprobar que queda al menos un destinatario válido
    If DestValidos = "" Then
        MsgBox "No quedan destinatarios válidos. Envío cancelado.", vbCritical, "Error Email"
        Exit Function
    End If

    ' Devolver la lista limpia al módulo que llama
    Destinatario = DestValidos
    Fnc_Valid_Email_Multi = True

End Function


'--------------------------------------------------------------------------------------------------
' Fnc_Valid_Email
'   Valida una sola dirección de correo (sin MsgBox)
'   Devuelve TRUE si el formato es correcto
'--------------------------------------------------------------------------------------------------
Public Function Fnc_Valid_Email(ByVal Email As String) As Boolean

    Dim MyRegExp As Object

    Set MyRegExp = CreateObject("VBScript.RegExp")
    MyRegExp.IgnoreCase = True
    MyRegExp.Pattern = "^[a-z0-9_.+\-]+@[a-z0-9.\-]{2,}\.[a-z]{2,6}$"
    Fnc_Valid_Email = MyRegExp.Test(Trim(Email))
    Set MyRegExp = Nothing

End Function
