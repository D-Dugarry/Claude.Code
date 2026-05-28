Attribute VB_Name = "M_1_Gestion_Correos"
' ==============================================================================
' Módulo    : Mandar_Correos
' Proyecto  : Mailing Indicadores UA
' Autor     : Dugarry
' Descripción: Envío masivo de correos personalizados con adjuntos a departamentos
'              de la Universidad de Alicante.
'              Motor de envío: M_800_Mail_Send_New (CDO + Gmail SMTP).
'
' Subrutinas públicas:
'   · Enviar_Emails() — Rutina principal de envío masivo
'
' Dependencias:
'   · M_800_Mail_Send_New -> Rut_Email_Send  (envío CDO, credenciales, errores)
'   · M_810_Mail_Valid    -> Fnc_Valid_Email  (validación de dirección por fila)
'
' IMPORTANTE: M_800 requiere hoja con código VBA "Prog__APP" y los rangos nombrados:
'   APP_MailCta    -> cuenta Gmail SMTP (ej. ingresos@gcloud.ua.es)
'   APP_MailFrom   -> remitente visible (ej. ingresos@ua.es)
'   APP_MailClau   -> App Password de Google (16 chars, sin espacios)
'   APP_MailFirm   -> plantilla HTML de firma (puede estar vacía)
'   APP_User_Rubrica -> rúbrica del remitente (puede estar vacía)
'   SW_Test        -> TRUE/FALSE (modo prueba; TRUE redirige a DIR_PRUEBA)
'
' Celdas de configuración (hoja DatosCorreo):
'   B3 -> asunto del correo
'   B4, B5, B6 -> líneas del cuerpo del mensaje
' ==============================================================================
Option Explicit

' ==============================================================================
' Enviar_Emails
' ------------------------------------------------------------------------------
' Itera todas las filas de Tb_Datos y envía un correo por cada destinatario,
' delegando el envío CDO y la configuración SMTP a Rut_Email_Send (M_800).
'
' Columnas de Tb_Datos:
'   Col 1 -> email destinatario
'   Col 2 -> nombre del fichero adjunto (sin ruta)
'   Col 3 -> estado resultado: "Enviado" / "Sin Destinatario" /
'            "Email no válido" / "El Fichero NO Existe"
' ==============================================================================
Sub Enviar_Emails()
    Dim Tb_Tabla        As ListObject
    Dim FicheroAdjunto  As String
    Dim Ultima_Fila     As Integer
    Dim Fila            As Integer
    Dim Cont_Emails     As Integer
    Dim Pregunta        As VbMsgBoxResult
    Dim Ruta            As String
    Dim Hora_Inicio     As Long
    Dim Destinatario    As String
    Dim CuerpoHTML      As String

    Hora_Inicio = Timer
    Set Tb_Tabla = DatosCorreo.ListObjects("Tb_Datos")

    ' --- Doble confirmación antes de enviar ---
    Pregunta = MsgBox("¿Seguro que quieres mandar todos los E-mails?", vbOKCancel, "Dugarry's Botones")
    If Pregunta = vbCancel Then Exit Sub
    Pregunta = MsgBox("¿Seguro que deseas continuar?", vbYesNo + vbQuestion, "Dugarry's Botones")
    If Pregunta = vbNo Then Exit Sub

    ' --- Selección de carpeta con los ficheros adjuntos ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar Carpeta donde están los archivos"
        .InitialFileName = ThisWorkbook.Path
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        Ruta = .SelectedItems(1) & "\"
    End With

    ' --- Preparar tabla: limpiar Col3 y ordenar ---
    With Tb_Tabla
        .ListColumns(3).DataBodyRange.Clear
        .Sort.SortFields.Clear
        .Range.Sort key1:=.ListColumns(1), order1:=xlAscending, _
                    key2:=.ListColumns(2), order2:=xlAscending, Header:=xlYes
        Ultima_Fila = .DataBodyRange.Rows.Count
    End With

    ' --- Cuerpo HTML construido una sola vez desde las celdas de configuración ---
    CuerpoHTML = "<html><body>" & _
                 "<p>" & DatosCorreo.Range("B4").Value & "</p>" & _
                 "<p>" & DatosCorreo.Range("B5").Value & "</p>" & _
                 "<p>" & DatosCorreo.Range("B6").Value & "</p>" & _
                 "</body></html>"

    Cont_Emails = 0

    ' ==============================================================================
    ' Bucle principal de envío
    ' ==============================================================================
    For Fila = 1 To Ultima_Fila
        Destinatario = Trim(CStr(Tb_Tabla.DataBodyRange.Cells(Fila, 1)))

        ' --- Validaciones previas al envío ---
        If Destinatario = "" Then
            Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "Sin Destinatario"
            GoTo SiguienteFila
        End If
        If Not Fnc_Valid_Email(Destinatario) Then
            Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "Email no válido"
            GoTo SiguienteFila
        End If
        FicheroAdjunto = Ruta & CStr(Tb_Tabla.DataBodyRange.Cells(Fila, 2))
        If Dir(FicheroAdjunto) = "" Then
            Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "El Fichero NO Existe"
            GoTo SiguienteFila
        End If

        Cont_Emails = Cont_Emails + 1
        Application.StatusBar = ">>>>>>>>>>>>>>>>>     Enviando Email nº: " & Cont_Emails & _
                                " a: " & Destinatario

        ' --- Enviar mediante M_800 (CDO + Gmail SMTP) ---
        Rut_Email_Send Destinatario, DatosCorreo.Range("B3").Value, FicheroAdjunto, CuerpoHTML
        Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "Enviado"

SiguienteFila:
    Next Fila

    ' --- Resumen final ---
    Application.StatusBar = True
    Application.ScreenUpdating = True
    MsgBox "He tardado: " & Timer - Hora_Inicio & "  segundos" & vbCrLf & vbCrLf & _
           "He mandado: " & Cont_Emails & "  Emails.", , "Dugarry's Botones"

End Sub     ' Enviar_Emails
' ==============================================================================
