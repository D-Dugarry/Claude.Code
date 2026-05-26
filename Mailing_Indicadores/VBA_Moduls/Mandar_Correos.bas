Attribute VB_Name = "Mandar_Correos"
'   Nota Importante: Tenemos que Activar en este libro de Excel en Herramientas->Referencias->Microsoft CDO for Windows 2000 library
Option Explicit

'=================================================================================================================
'=================================================================================================================
Sub Enviar_Emails()

Dim FicheroAdjunto          As String
Dim FicheroAdjuntoExiste    As String
Dim Password_Correo         As String

Dim Tb_Tabla As Object
Set Tb_Tabla = ThisWorkbook.Sheets(DatosCorreo.Name).ListObjects("Tb_Datos")

Dim Correo                  As CDO.Message
Dim Configuracion_Correo    As CDO.Configuration
Dim Campos                  As Variant
Dim msConfigURL             As String
Dim Ultima_Fila             As Integer
Dim Fila                    As Integer
Dim Cont_Emails             As Integer
Dim Pregunta                As Integer
Dim Ruta                    As String
Dim Hora_Inicio                     As Long             ' Para Saber el tiempo de proceso

Hora_Inicio = Timer                ' Para Saber el tiempo de proceso

Pregunta = MsgBox("¿Seguro que quieres mandar todos los E-mails?", vbOKCancel, "Dugarry's Botones")
If Pregunta = 2 Then MsgBox "Elegiste Cancelar", , "Dugarry's Botones": Exit Sub
Pregunta = MsgBox("¿Seguro que deseas continuar?", vbYesNo + vbQuestion, "Dugarry's Botones")
If Pregunta = vbNo Then Exit Sub

Ruta = ActiveWorkbook.Path   ' Averigua la carpeta de la hoja de cálculo actual

With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Seleccionar Carpeta donde están los archivos"
    .InitialFileName = Ruta  ' = "C:\" o = "C:\Ejercicios"... - en caso de conocer la ruta
    .Show
    If .SelectedItems.Count = 0 Then Exit Sub
    Ruta = .SelectedItems(1) & "\"
End With

' -------------------------------------------------------
' Ordenar x Email(6) en Ascendente
' -------------------------------------------------------
With Tb_Tabla
    Tb_Tabla.ListColumns(3).DataBodyRange.Clear
    .Sort.SortFields.Clear
    .Range.Sort key1:=.ListColumns(1), order1:=xlAscending, _
                key2:=.ListColumns(2), order2:=xlAscending, Header:=xlYes      ' Ordenar por Columna Núm_Artículos
    Application.ScreenUpdating = True                          'Evita que la pantalla esté constantemente actualizándose

Password_Correo = InputBox("Introduce la contraseña del Correo: ", "Envío masivo de correos con adjuntos")
    ' -------------------------------------------------------
    ' Enviar los Emails ---------------------------------------
    ' -------------------------------------------------------
    Cont_Emails = 0
    Ultima_Fila = .DataBodyRange.Rows.Count            ' Última Filas con datos
End With    ' Tb_Tabla

For Fila = 1 To Ultima_Fila

    Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "Enviado"

    If Tb_Tabla.DataBodyRange.Cells(Fila, 1) = "" Then Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "Sin Destinatario": GoTo SiguienteFila
    FicheroAdjunto = Ruta & Tb_Tabla.DataBodyRange.Cells(Fila, 2)
    FicheroAdjuntoExiste = Dir(FicheroAdjunto)
    If FicheroAdjuntoExiste = "" Then Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "El Fichero NO Existe": GoTo SiguienteFila
    
    Cont_Emails = Cont_Emails + 1
    Application.StatusBar = ">>>>>>>>>>>>>>>>>     Enviando Email nº: " & Cont_Emails & " a: " & Cells(Fila, 1)
    
    Set Correo = New CDO.Message
    Set Configuracion_Correo = New CDO.Configuration
    
    Configuracion_Correo.Load -1
    
    Set Campos = Configuracion_Correo.Fields
    
    With Correo
        .From = Range("B2")             ' Emisor
        .Subject = Range("B3")          ' Asunto del correo
        .To = Tb_Tabla.DataBodyRange.Cells(Fila, 1)                    ' Destinatario
        .CC = ""
        .BCC = ""
        .TextBody = Range("B4") & vbCrLf & vbCrLf & vbCrLf & _
                    Range("B5") & vbCrLf & vbCrLf & _
                    Range("B6")
        .AddAttachment FicheroAdjunto
        ' ----------- Para mandar varios ficheros en el mismo correo ---------------------------
        While Tb_Tabla.DataBodyRange.Cells(Fila + 1, 1) = Tb_Tabla.DataBodyRange.Cells(Fila, 1)
            Fila = Fila + 1
            FicheroAdjunto = Ruta & Tb_Tabla.DataBodyRange.Cells(Fila, 2)
            FicheroAdjuntoExiste = Dir(FicheroAdjunto)
            If FicheroAdjuntoExiste = "" Or Tb_Tabla.DataBodyRange.Cells(Fila, 2) = "" Then
                Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "El Fichero NO Existe"
            Else
                .AddAttachment Ruta & Tb_Tabla.DataBodyRange.Cells(Fila, 2)
                Tb_Tabla.DataBodyRange.Cells(Fila, 3) = "Enviado"
            End If
        Wend
        '----------------------------------------------------------------------------------------
        
    End With    ' Correo
    
    msConfigURL = "http://schemas.microsoft.com/cdo/configuration"
    
    With Campos
        .Item(msConfigURL & "/smtpusessl") = True
        .Item(msConfigURL & "/smtpauthenticate") = 1
        
        .Item(msConfigURL & "/smtpserver") = "smtp.gmail.com"
        .Item(msConfigURL & "/smtpserverport") = 465
        .Item(msConfigURL & "/sendusing") = 2
        
        .Item(msConfigURL & "/sendusername") = Range("B2")
        .Item(msConfigURL & "/sendpassword") = Password_Correo
        .Update
    
    End With    ' Campos
    
    Correo.Configuration = Configuracion_Correo
    Correo.Send
    
SiguienteFila:
Next Fila

exit_line:
'// Release object memory
Set Correo = Nothing
Set Configuracion_Correo = Nothing

Application.StatusBar = True
Application.ScreenUpdating = True                          'Evita que la pantalla esté constantemente actualizándose

MsgBox "He tardado: " & Timer - Hora_Inicio & "  segundos" & vbCrLf & vbCrLf & vbCrLf & _
        "He mandado los : " & Cont_Emails & "  Emails. ", , "Dugarry's Botones"

End Sub     ' Enviar_Emails
'=================================================================================================================

'=================================================================================================================
'=================================================================================================================
Sub Borrar_Datos_Tabla()
    Dim Pregunta    As Integer
    Dim Tb_Tabla    As Object
    Dim Tb_Rango    As Range
    Set Tb_Tabla = ThisWorkbook.Sheets(DatosCorreo.Name).ListObjects("Tb_Datos")
    Set Tb_Rango = Tb_Tabla.DataBodyRange
    If Tb_Rango Is Nothing Then MsgBox "No hay Datos en la Tabla para Borrar", , "Dugarry's Botones": Exit Sub
    ActiveWorkbook.Sheets(DatosCorreo.Name).ListObjects("Tb_Datos").DataBodyRange.Select
    Pregunta = MsgBox("¿Seguro que deseas borrar todas estas filas?", vbYesNo + vbQuestion, "Dugarry's Botones")
    If Pregunta = vbYes Then ThisWorkbook.Sheets(DatosCorreo.Name).ListObjects("Tb_Datos").DataBodyRange.Delete
    ThisWorkbook.Sheets(DatosCorreo.Name).ListObjects("Tb_Datos").HeaderRowRange.Select
Application.ScreenUpdating = True                          'Evita que la pantalla esté constantemente actualizándose
End Sub      'Borrar_Datos_Tabla
'=================================================================================================================

'=================================================================================================================
'=================================================================================================================
Sub Añadir_Lista_Ficheros()
 
Dim Ruta                    As String
Dim Ob_Fichero          As Object

Dim Tb_Tabla As Object
Set Tb_Tabla = ThisWorkbook.Sheets(DatosCorreo.Name).ListObjects("Tb_Datos")

Dim Ob_FileSistObjct    As Object
Set Ob_FileSistObjct = CreateObject("Scripting.FileSystemObject")

    Ruta = ActiveWorkbook.Path   ' Averigua la carpeta de la hoja de cálculo actual
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar Carpeta donde están los archivos"
        .AllowMultiSelect = False
        .InitialFileName = Ruta  ' = "C:\" o = "C:\Ejercicios"... - en caso de conocer la ruta
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        Ruta = .SelectedItems(1) & "\"
    End With

Dim Ob_Ruta             As Object
Set Ob_Ruta = Ob_FileSistObjct.GetFolder(Ruta)

Dim NewRow As ListRow
 
For Each Ob_Fichero In Ob_Ruta.Files
    Set NewRow = Tb_Tabla.ListRows.Add
    NewRow.Range(2) = Ob_Fichero.Name
Next Ob_Fichero
Application.ScreenUpdating = True                          'Evita que la pantalla esté constantemente actualizándose
 
End Sub     ' Añadir_Lista_Ficheros
'=================================================================================================================


Sub Rut_Iniciar()
    Application.DisplayFullScreen = True                        'Ves pantalla completa
    Application.DisplayFormulaBar = False                       'Oculta la barra de formulas
    If ActiveWindow.DisplayGridlines Then ActiveWindow.DisplayGridlines = False         'Oculta las lineas de la cuadricula
    If ActiveSheet.DisplayPageBreaks Then ActiveSheet.DisplayPageBreaks = False         'Oculta las líneas de salto de página
    If CommandBars("Ribbon").Controls(1).Height > 100 Then CommandBars.ExecuteMso ("MinimizeRibbon")
    Application.ScreenUpdating = True                          'Evita que la pantalla esté constantemente actualizándose
    Application.EnableEvents = True
End Sub


