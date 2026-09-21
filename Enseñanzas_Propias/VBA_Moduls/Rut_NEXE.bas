Attribute VB_Name = "Rut_NEXE"
' Last Rev. 2026-09-21 12:12
Option Explicit

' ==================================================================================================
Function Fnc_Format_Ruta(Ruta_Red As String) As String       ' Cambia la Ruta de la Red Nexe a Unidades de disco Local
' ==================================================================================================

    If Left(ActiveWorkbook.Path, 18) = "https://nexe.ua.es" Then
        Debug.Print "Red"
    Else
        Debug.Print "Local"
    End If
    Fnc_Format_Ruta = Replace(Ruta_Red, "/", "\")                                           '- Cambio / por \
'    Fnc_Format_Ruta = Mid(Fnc_Format_Ruta, InStr(1, Fnc_Format_Ruta, "\Espai_ServConta\"))  '- Quito la parte de Https\\...
    If InStr(1, Fnc_Format_Ruta, "\Espai_ServConta\") > 0 Then
        Fnc_Format_Ruta = Mid(Fnc_Format_Ruta, InStr(1, Fnc_Format_Ruta, "\Espai_ServConta\"))   '- Quito la parte de Https\\...
'       Fnc_Format_Ruta = Mid(Fnc_Format_Ruta, InStr(1, Fnc_Format_Ruta, Range("App_MailUsu")) + Len(Range("App_MailUsu")))
        Fnc_Format_Ruta = Range("App_LetraUnidRed") & ":" & Fnc_Format_Ruta                     '- Añado la letra de la Unidad
    End If
Debug.Print Fnc_Format_Ruta
Debug.Print Fnc_Format_Ruta
End Function        ' Fnc_Format_Referencia
'---------------------------------------------------------------------------------------------------

'###################################################################################################
        'call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Fich_SelectedItem, [NomFich], [PathFich])
Sub Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Fich_SelectedItem As String, Optional NomFich As String = "", Optional RutaFich As String = "")
' ==================================================================================================
    Debug.Print "Rut_ArchFullName_SeparaEn_NameFile_y_PathFile: " & Fich_SelectedItem
    If InStr(Fich_SelectedItem, "/") > 0 Then
        NomFich = Right(Fich_SelectedItem, Len(Fich_SelectedItem) - InStrRev(Fich_SelectedItem, "/"))   '- ¡¡ Funciona con NEXE !!
        RutaFich = Left(Fich_SelectedItem, InStrRev(Fich_SelectedItem, "/"))                            ' Extrae solo la ruta del directorio
    Else
        NomFich = Right(Fich_SelectedItem, Len(Fich_SelectedItem) - InStrRev(Fich_SelectedItem, "\"))   '- ¡¡ Funciona con NEXE !!
        RutaFich = Left(Fich_SelectedItem, InStrRev(Fich_SelectedItem, "\"))                            ' Extrae solo la ruta del directorio
    End If
End Sub             ' Fnc_Format_Referencia

