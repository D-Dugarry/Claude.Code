Attribute VB_Name = "M_820_Range_TO_HTML"
Option Explicit

'==================================================================================================
' M_820_Range_TO_HTML
'
' Convierte un rango de Excel a HTML para incluir en el cuerpo de un correo.
'
' FUNCIONAMIENTO:
'   1. Copia el rango a un libro temporal
'   2. Publica ese libro como fichero HTML en la carpeta Temp del sistema
'   3. Auto-detecta el charset del HTML generado (UTF-8 o Windows-1252)
'   4. Lee el HTML con ADODB.Stream usando el charset detectado
'   5. Limpia: cierra el libro temporal y borra el fichero HTML
'
' CODIFICACION:
'   Excel PublishObjects puede generar el fichero HTML en UTF-8 o Windows-1252 segun
'   la version/build de Excel y la configuracion regional del sistema.
'   Ej: Excel 365 reciente genera UTF-8; versiones anteriores generan Windows-1252.
'   SOLUCION: se lee primero con Windows-1252 (charset seguro para el meta tag ASCII),
'   se detecta el charset real del meta tag, y si es UTF-8 se relee correctamente.
'   Esto hace el modulo compatible con cualquier entorno sin cambiar codigo.
'
' NOTA SW_WB_Deactivate:
'   Se pone a FALSE antes del PasteSpecial para evitar que Workbook_Deactivate
'   llame a Rut_ConfigExcel_Restablecer, que interfiere con la operacion interna.
'   Se restaura a TRUE en el bloque Cleanup (siempre, con o sin error).
'==================================================================================================


'--------------------------------------------------------------------------------------------------
' Fnc_RangeToHTML
'   Convierte un rango de Excel a HTML estatico
'   Auto-detecta el encoding (UTF-8 o Windows-1252) segun lo que genere Excel en este entorno
'   Devuelve el HTML como String, o "" si ocurre algun error
'--------------------------------------------------------------------------------------------------
Public Function Fnc_RangeToHTML(ByVal Rango As Range) As String

    Dim ADO         As Object
    Dim TempFile    As String
    Dim TempWB      As Workbook

    Fnc_RangeToHTML = ""

    On Error GoTo Err_RangeToHTML

    ' Deshabilitar Workbook_Deactivate para evitar interferencia con PasteSpecial
    Application.EnableEvents = False
    Prog__APP.Range("SW_WB_Deactivate") = False

    ' --- Fichero HTML temporal ---
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".html"

    ' --- Copiar rango a libro temporal ---
    Rango.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial xlPasteColumnWidths
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo Err_RangeToHTML
    End With

    ' --- Publicar como HTML estatico ---
    With TempWB.PublishObjects.Add( _
            SourceType:=xlSourceRange, _
            Filename:=TempFile, _
            Sheet:=TempWB.Sheets(1).Name, _
            Source:=TempWB.Sheets(1).UsedRange.Address, _
            HtmlType:=xlHtmlStatic)
        .Publish True
    End With

    ' --- Auto-detectar encoding: leer con Windows-1252 (seguro para el meta tag ASCII) ---
    Dim sRaw        As String
    Dim sCharset    As String

    Set ADO = CreateObject("ADODB.Stream")
    With ADO
        .Type = 2               ' adTypeText
        .Charset = "windows-1252"
        .Open
        .LoadFromFile TempFile
        sRaw = .ReadText
        .Close
    End With
    Set ADO = Nothing

    ' El meta tag <meta ... charset=xxx> es ASCII puro -> detectable con cualquier encoding
    If InStr(1, LCase(sRaw), "charset=utf-8") > 0 Then
        sCharset = "utf-8"
    Else
        sCharset = "windows-1252"
    End If

    ' --- Leer con el charset correcto ---
    If sCharset = "utf-8" Then
        Set ADO = CreateObject("ADODB.Stream")
        With ADO
            .Type = 2
            .Charset = "utf-8"
            .Open
            .LoadFromFile TempFile
            Fnc_RangeToHTML = .ReadText
            .Close
        End With
        Set ADO = Nothing
    Else
        Fnc_RangeToHTML = sRaw
    End If

    ' --- Ajustes sobre el HTML generado ---
    Fnc_RangeToHTML = Replace(Fnc_RangeToHTML, "align=center x:publishsource=", "align=left x:publishsource=")
    Fnc_RangeToHTML = Replace(Fnc_RangeToHTML, "display:none", "")

Cleanup:
    On Error Resume Next
    If Not ADO Is Nothing Then ADO.Close
    If Not TempWB Is Nothing Then TempWB.Close SaveChanges:=False
    If Dir(TempFile) <> "" Then Kill TempFile
    Set ADO = Nothing
    Set TempWB = Nothing
    Prog__APP.Range("SW_WB_Deactivate") = True
    Application.EnableEvents = True
    On Error GoTo 0
    Exit Function

Err_RangeToHTML:
    MsgBox "Error en Fnc_RangeToHTML:" & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Fnc_RangeToHTML"
    Fnc_RangeToHTML = ""
    Resume Cleanup

End Function
