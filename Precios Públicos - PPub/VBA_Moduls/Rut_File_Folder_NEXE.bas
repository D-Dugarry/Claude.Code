Attribute VB_Name = "Rut_File_Folder_NEXE"
' Last Rev. 2026-10-07 13:48
Option Explicit

'###################################################################################################################################
Function Fnc_NEXE_RutaAPP() As String  ' Formatea la Ruta de RED del Excel Actual, dependiendo de la Red Nexe o del disco Local -------
' ==================================================================================================================================
    Fnc_NEXE_RutaAPP = Application.Workbooks(ThisWorkbook.Name).Path
    Debug.Print "Fnc_NEXE_RutaAPP: " & Fnc_NEXE_RutaAPP
    If Left(Fnc_NEXE_RutaAPP, 18) = "https://nexe.ua.es" Then
        Fnc_NEXE_RutaAPP = Replace(Fnc_NEXE_RutaAPP, "/", "\")                                        '- Cambio / por \
        Fnc_NEXE_RutaAPP = Mid(Fnc_NEXE_RutaAPP, InStr(1, Fnc_NEXE_RutaAPP, Prog__APP.Range("APP_User_Mail")) + Len(Prog__APP.Range("APP_User_Mail")))
        Fnc_NEXE_RutaAPP = Prog__APP.Range("App_User_Unid_Red") & ":" & Fnc_NEXE_RutaAPP                     '- Añado la letra de la Unidad
    End If
End Function        ' Fnc_Format_Referencia
'###################################################################################################################################
Function Fnc_Format_Ruta(Ruta_Red As String) As String  ' Cambia una ruta de la Red Nexe a la unidad local
' =================================================================================================
'- El mismo cambio que Fnc_NEXE_RutaAPP, pero para cualquier ruta. La llama por su nombre, con
'  Application.Run, el módulo de copias Rut_Wb_CopSegTimed_USB_HD (skill excel-copseg-backup) antes de
'  guardar en CopiaSeguridad\: si se renombra, deja de usarse sin avisar. Una ruta que no es de la Red
'  Nexe se devuelve tal cual.
    Dim Mail        As String
    Dim PosMail     As Long
    Fnc_Format_Ruta = Ruta_Red
    If Left(Ruta_Red, 18) <> "https://nexe.ua.es" Then Exit Function
    Mail = Prog__APP.Range("APP_User_Mail")
    If Mail = "" Then Exit Function
    PosMail = InStr(1, Ruta_Red, Mail)
    If PosMail = 0 Then Exit Function
    Fnc_Format_Ruta = Prog__APP.Range("App_User_Unid_Red") & ":" & Replace(Mid(Ruta_Red, PosMail + Len(Mail)), "/", "\")
End Function        ' Fnc_Format_Ruta
'##################################################################################################
        'call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Fich_SelectedItem, [NomFich], [PathFich])
Sub Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(Fich_SelectedItem As String, Optional NomFich As String = "", Optional RutaFich As String = "")
' ==================================================================================================================================
    Debug.Print "Rut_ArchFullName_SeparaEn_NameFile_y_PathFile: " & Fich_SelectedItem
    If InStr(Fich_SelectedItem, "/") > 0 Then
        NomFich = Right(Fich_SelectedItem, Len(Fich_SelectedItem) - InStrRev(Fich_SelectedItem, "/"))   '- ¡¡ Funciona con NEXE !!
        RutaFich = Left(Fich_SelectedItem, InStrRev(Fich_SelectedItem, "/"))                            ' Extrae solo la ruta del directorio
    Else
        NomFich = Right(Fich_SelectedItem, Len(Fich_SelectedItem) - InStrRev(Fich_SelectedItem, "\"))   '- ¡¡ Funciona con NEXE !!
        RutaFich = Left(Fich_SelectedItem, InStrRev(Fich_SelectedItem, "\"))                            ' Extrae solo la ruta del directorio
    End If
    If InStr(Fich_SelectedItem, "/") > 0 Then NomFich = Fnc_URL_Decodificar(NomFich)             '- Los nombres de una URL vienen con %C3%B1 etc.
End Sub             ' Fnc_Format_Referencia
Function Fnc_URL_Decodificar(ByVal Txt As String) As String  ' Decodifica %C3%B1 (UTF-8 en porcentaje), como los nombres de una URL de NEXE
' ==================================================================================================================================
    Dim i As Long, B As Long, B2 As Long, B3 As Long
    Dim Res As String
    i = 1
    Do While i <= Len(Txt)
        If Mid$(Txt, i, 1) = "%" And i + 2 <= Len(Txt) Then
            If Mid$(Txt, i + 1, 2) Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
                B = CLng("&H" & Mid$(Txt, i + 1, 2))
                If B >= &HC2 And B <= &HDF And Mid$(Txt, i + 3, 1) = "%" And i + 5 <= Len(Txt) Then          '- 2 bytes (tildes, enye)
                    B2 = CLng("&H" & Mid$(Txt, i + 4, 2))
                    Res = Res & ChrW$(((B And &H1F) * 64) Or (B2 And &H3F)): i = i + 6
                ElseIf B >= &HE0 And B <= &HEF And Mid$(Txt, i + 3, 1) = "%" And Mid$(Txt, i + 6, 1) = "%" And i + 8 <= Len(Txt) Then   '- 3 bytes
                    B2 = CLng("&H" & Mid$(Txt, i + 4, 2)): B3 = CLng("&H" & Mid$(Txt, i + 7, 2))
                    Res = Res & ChrW$(((B And &HF) * 4096) Or ((B2 And &H3F) * 64) Or (B3 And &H3F)): i = i + 9
                Else
                    Res = Res & Chr$(B): i = i + 3                          '- 1 byte (ASCII, p.ej. %20)
                End If
            Else
                Res = Res & "%": i = i + 1
            End If
        Else
            Res = Res & Mid$(Txt, i, 1): i = i + 1
        End If
    Loop
    Fnc_URL_Decodificar = Res
End Function        ' Fnc_URL_Decodificar
'###################################################################################################################################
   'call Rut_File_Select ("Título...", NomFich, ["Excel"], ["*.xls?"])      '- NomFich = "Cancel"
Sub Rut_File_Select(Titulo As String, _
                    NomFich As String, _
                    Optional TipoFich_Txt As String = "Cualquier Fichero", _
                    Optional TipoFich As String = "*.*")
' ==================================================================================================================================
Debug.Print ">>> Rut_File_Select"
    With Application.FileDialog(msoFileDialogFilePicker)
            .Title = Titulo
            .InitialFileName = Fnc_NEXE_RutaAPP & "\" & NomFich & "*"
            .InitialView = msoFileDialogViewDetails
            .ButtonName = "Seleccionar" ' o "Aceptar" o ...
            .Filters.Clear
            .Filters.Add TipoFich_Txt, TipoFich, 1
            .AllowMultiSelect = False
        If .Show = True Then
            Call Rut_ArchFullName_SeparaEn_NameFile_y_PathFile(.SelectedItems(1), NomFich)
        Else
'            MsgBox "Ha pulsado el botón <Cancelar>.", vbOKOnly, "Proceso: " & Título
            NomFich = "Cancel"
        End If
    End With
Debug.Print "<<< Rut_File_Select"
End Sub
'###################################################################################################################################
   'call Rut_File_Select ("Título...", NameFile, ["Excel"], ["*.xls?"])      '- NameFile = "Cancel"
Sub Rut_File_Select_V2(Titulo As String, _
                        NameFile As String, _
                        Optional TipoFich_Txt As String = "Cualquier Fichero", _
                        Optional TipoFich As String = "*.*", _
                        Optional PathFile As String = "")
' ==================================================================================================================================
Debug.Print ">>> Rut_File_Select_V2"
    With Application.FileDialog(msoFileDialogFilePicker)
            .Title = Titulo
            .InitialFileName = Fnc_NEXE_RutaAPP & "\" & NameFile & "*"
            .InitialView = msoFileDialogViewDetails
            .ButtonName = "Seleccionar" ' o "Aceptar" o ...
            .Filters.Clear
            .Filters.Add TipoFich_Txt, TipoFich, 1
            .AllowMultiSelect = False
        If .Show = True Then
            NameFile = .SelectedItems(1)
        Else
'            MsgBox "Ha pulsado el botón <Cancelar>.", vbOKOnly, "Proceso: " & Título
            NameFile = "Cancel"
        End If
    End With
Debug.Print "<<< Rut_File_Select_V2"
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------
Sub Rut_TimeLap_Inf(ActivForm As Object, _
                    Kontrol As String, _
                    MsgTxt_A As String, _
                    TimeIni As Single, _
                    Optional MsgTxt_B As String = "", _
                    Optional MsgTxt_C As String = "", _
                    Optional TxT_Progreso As String = "", _
                    Optional SW_ReseTxT As Boolean = False, _
                    Optional IniLap As Boolean = False, _
                    Optional RowsJumpBefore As Integer = 1, _
                    Optional RowsJumpAfter As Integer = 1)
                    
    Dim Cont        As Integer
    Dim MsgTxt      As String
    Dim TxTAnt      As String
    Const Tab99     As Integer = 104    '- Ancho máx línea
    Const Tab02     As Integer = 18
    Const Tab03     As Integer = 20
    Dim Tab01       As Integer:       Tab01 = Tab99 - Tab02 - Tab03
    
    If TimeIni = 0 Then                                     '- Línea SIN Hora NI Lapso
        MsgTxt = MsgTxt_A
    Else                                                    '- Línea Con Hora Y Lapso
        MsgTxt = Format(Now, "hh:mm:ss") & " Lap: " & Right("    " & Format(Round(Timer - TimeIni, 2), "#0.00"), 6) & " seg. " & MsgTxt_A
    End If
    If MsgTxt = "" Then MsgTxt = String(Tab01, " ")
    If MsgTxt_B <> "" Then                                  '- .......... Txt_B
        MsgTxt = Left(MsgTxt & String(Tab01, "."), Tab01)
        MsgTxt = MsgTxt & Right(String(Tab02, ".") & MsgTxt_B, Tab02)
    End If                                                  '- .......... Txt_C
    If MsgTxt_C <> "" Then
        MsgTxt = MsgTxt & Right(String(Tab03, ".") & MsgTxt_C, Tab03)
    End If
    
    If ActivForm Is Nothing Then
        TxTAnt = Task_Inf
    Else
        TxTAnt = ActivForm.Controls(Kontrol).Value
        Task_Inf = ""
    End If
    
    If SW_ReseTxT Then TxTAnt = ""                      '- Reinicio el Contenido a Nada
    If TxT_Progreso <> "" Then TxTAnt = TxT_Progreso    '- Reinicio el Contenido a TxT_Progreso
    For Cont = 1 To RowsJumpBefore - 1                      '- Añado líneas en Blanco antes del Último Mensaje
        TxTAnt = TxTAnt & vbLf
    Next
    For Cont = 1 To RowsJumpAfter                           '- Añado líneas en Blanco después del Último Mensaje
        MsgTxt = MsgTxt & vbLf
    Next
    
    TxTAnt = TxTAnt & MsgTxt
    
    If ActivForm Is Nothing Then
        Task_Inf = TxTAnt
        Debug.Print Task_Inf
    Else
        With ActivForm.Controls(Kontrol)
            .Value = TxTAnt
            .SelStart = Len(.Text)      '- Mueve el cursor a la posición X, en este caso el final del Text.
            .SelLength = 0
            .SetFocus
        End With
        ActivForm.Repaint
    End If
    
'    With ActivForm.Controls(Kontrol)
'        If SW_ReseTxT Then .Value = ""                      '- Reinicio el Contenido a Nada
'        If TxT_Progreso <> "" Then .Value = TxT_Progreso    '- Reinicio el Contenido a TxT_Progreso
'        For Cont = 1 To RowsJumpBefore - 1                          '- Añado líneas en Blanco antes del Último Mensaje
'            .Value = .Value & vbLf
'        Next
'        .Value = .Value & MsgTxT & vbLf
'        .SelStart = Len(.Text)      '- Mueve el cursor a la posición X, en este caso el final del Text.
'        .SetFocus
'    End With
'                ActivForm.Repaint
                LastTimeLap = Timer
End Sub
' -------------------------------------------------------------------------------------------------------------------------------<<<
''-----------------------------------------------------------------------------------------------------------------------------------
'Sub Rut_Form_Show_Task_Inf(Kontrol As String, MsgTxt As String)     '- PARECE QUE NO FUNCIONA
'    Dim ActivForm        As Object                          '- Identificamos qué Formulario está Activo.  ----------
'    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
'    With ActivForm.Controls(Kontrol)
'        .Value = MsgTxt
''        .SelStart = Len(.Text)      '- Mueve el cursor a la posición X, en este caso el final del Text.
'        .SetFocus
'    End With
''        ActivForm.Repaint
'    Set ActivForm = Nothing
'End Sub
'

'- Eliminadas 2 Subs ajenas al proyecto (ImportMultipleFiles, RuT_WrkBook_Sheets_Select_Inport):
'- referencian objetos de un TPV (H_Resumen, Tb_TPV, columnas Importe/Comision/Cobrado) que no existen
'- en este libro, y ImportMultipleFiles usaba la variable Lr sin declarar (error de compilacion con
'- Option Explicit). Ninguna de las dos tenia ningun llamador en el proyecto.
