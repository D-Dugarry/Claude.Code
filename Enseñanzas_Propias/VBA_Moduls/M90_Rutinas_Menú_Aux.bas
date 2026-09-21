Attribute VB_Name = "M90_Rutinas_Menú_Aux"
' Last Rev. 2026-09-21 12:12
Option Explicit


' ==================================================================================================
Sub Rut_Activar_Programación()
'    RuT_Al_Abrir_WorkBook
    Call Rut_Sheets_Show
    Form_Menu.TB_Informe = "Estado de Programación Activado"
    Unload Form_Menu
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_Reset_App()
    Call RuT_Al_Abrir_WorkBook
    Form_Menu.TB_Informe = "App Reset"
    Unload Form_Menu
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_Btn_Menú_Aux()
        Form_Menu.Show
        Application.ScreenUpdating = True
        DoEvents
End Sub
'===================================================================================================
Sub Rut_Cerrar_Menú()
    Unload Form_Menu
End Sub
' ==================================================================================================
Sub Rut_Chg_Usuario()
        Form_Usuario.Show
        Application.ScreenUpdating = True
        DoEvents
End Sub
' ==================================================================================================
Sub Rut_OnOff_SW_VerRecNeg()
    Prog__APP_Switch.Range("Sw_VerRecNeg") = Not Prog__APP_Switch.Range("Sw_VerRecNeg")
    'If Prog__APP_Switch.Range("Sw_VerRecNeg") Then Range("Liquid_Sw_VerRecNeg") = "Hide Rec.Neg." Else Range("Liquid_Sw_VerRecNeg") = "Ver Rec.Neg."
    If Prog__APP_Switch.Range("Sw_VerRecNeg") Then
        Range("Liquid_Sw_VerRecNeg").Interior.ColorIndex = 6
        Range("Liquid_Sw_VerRecNeg").Font.ColorIndex = 13
        Range("Liquid_Sw_VerRecNeg") = "Hide Rec.Neg."
    Else
        Range("Liquid_Sw_VerRecNeg").Interior.ColorIndex = 13
        Range("Liquid_Sw_VerRecNeg").Font.ColorIndex = 6
        Range("Liquid_Sw_VerRecNeg") = "Ver Rec.Neg."
    End If
    Call Rut_00_Liquid_TitProp(UCase(Range("Liquid_Plan")), Range("Liquid_Curso_Acad"))
    Form_Menu.TB_Informe = "Visualizar registros negativos: Sw_VerRecNeg = " & Prog__APP_Switch.Range("Sw_VerRecNeg")
End Sub
' ==================================================================================================
Sub Rut_OnOff_SW_VerRecNoCob()
    Prog__APP_Switch.Range("Sw_VerRecNoCob") = Not Prog__APP_Switch.Range("Sw_VerRecNoCob")
    If Prog__APP_Switch.Range("Sw_VerRecNoCob") Then
        Range("Liquid_Sw_VerRecNoCob").Interior.ColorIndex = 6
        Range("Liquid_Sw_VerRecNoCob").Font.ColorIndex = 13
        Range("Liquid_Sw_VerRecNoCob") = "Hide No.Cob."
    Else
        Range("Liquid_Sw_VerRecNoCob").Interior.ColorIndex = 13
        Range("Liquid_Sw_VerRecNoCob").Font.ColorIndex = 6
        Range("Liquid_Sw_VerRecNoCob") = "Ver No.Cob."
    End If
    Call Rut_00_Liquid_TitProp(UCase(Range("Liquid_Plan")), Range("Liquid_Curso_Acad"))
    Form_Menu.TB_Informe = "Visualizar registros NO cobrados: Sw_VerRecNoCob = " & Prog__APP_Switch.Range("Sw_VerRecNoCob")
End Sub
'Sub Rut_OnOff_SW_Probando()
'    If Prog__APP_Switch.Range("Sw_Probando") Then
'        Prog__APP_Switch.Range("Sw_Probando") = False
'        Form_Menu.Lb_SW_Probando.Visible = False
'        Form_Menu.TB_Informe = "SW_Probando DesActivado"
'    Else
'        Prog__APP_Switch.Range("Sw_Probando") = True
'        Form_Menu.Lb_SW_Probando.Visible = True
'        Form_Menu.TB_Informe = "SW_Probando Activado"
'    End If
'End Sub
' ==================================================================================================
Sub Rut_Show_All_Columns()
        ActiveSheet.Unprotect
    Columns.EntireColumn.Hidden = False
'    IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ------------------------
    ActiveSheet.Protect , allowFiltering:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA
    Form_Menu.TB_Informe = "All columns are visible"
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
'Sub Rut_Mostrar_Col_Ocultas()       ' CuadroTextoVistaColmns    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Dim Cont_Col As Integer
''Solicitudes.Unprotect
'Application.ScreenUpdating = False
'        For Cont_Col = 1 To LastCol_Tb_Solicitudes
'            If Lo_Prog_Colns.DataBodyRange.Cells(7, Cont_Col) = "Ocultar" Then Columns(Cont_Col).Hidden = False
'        Next Cont_Col
'Application.ScreenUpdating = True
'Solicitudes.Protect
'Form_Menu.TB_Informe = "Columnas Ocultas al Usuario Visibles"
'End Sub     ' RuT_Mostrar_Ocultar_Col     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' ==================================================================================================
Sub Rut_ProtectUnProtect_ActivSheet()
    If ActiveSheet.ProtectContents Then
        ActiveSheet.Unprotect
        Form_Menu.TB_Informe = "ActiveSheet.UnProtect"
    Else
        ActiveSheet.Protect
        Form_Menu.TB_Informe = "ActiveSheet.Protect"
    End If
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_Sheets_Show()
Dim WrkSht          As Worksheet
    For Each WrkSht In Worksheets
                WrkSht.Visible = xlSheetVisible
    Next
    Call Rut_Reset_ToolsBar
    Form_Menu.TB_Informe = "All sheets Visible"
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================
Sub Rut_Sheets_Hide()
Dim WrkSht          As Worksheet

    Wk_TitP_Liquid.Visible = xlSheetVisible
    Wk_TitP_Liquid.Select

    For Each WrkSht In Worksheets
        If Left(WrkSht.CodeName, 5) = "Prog_" Then WrkSht.Visible = xlSheetVeryHidden   'xlSheetVisible   '
    Next
    
    Form_Menu.TB_Informe = "All sheets Hide"
End Sub     ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' ==================================================================================================

'===================================================================================================
Sub Rut_Tiempo_Transcurrido()
    Dim Hora_Inicio                  As Single
    Hora_Inicio = Now()                ' Para Saber el tiempo de proceso
    Debug.Print Format(Now() - Hora_Inicio, "hh.mm.ss")
End Sub






'    ' ============================  Para mostrar u ocultar Columnas según Lista [[lista]]   =======
'    ' =============================================================================================
'    Sub Rut_Mostrar_Columnas_Lista()       ' CuadroTextoVistaColmns    <<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Dim Pos_Ini     As Integer
'    Dim Pos_Fin     As Integer
'    Dim Cont            As Integer
'    Dim Pos_Caract      As Integer
'    Dim Letra_Col   As String
'    Dim Cadena      As String
'    Dim Caract      As String
'
'        Cadena = Form_Menu.TBx_Descripción
'        Pos_Ini = InStr(Cadena, "[[")
'        Pos_Fin = InStr(Cadena, "]]")
'        If Pos_Ini * Pos_Fin = 0 Then   '---Controla que existe marca de inicio y fin y que hay algun dato entre marcas
'            Form_Menu.TB_Informe = "Error: No hay lista de Columnas a visualizar, o no empieza por [[, o no acaba por ]]." & vbCrLf & Now
'            Exit Sub
'        End If
'
'        Cadena = Mid(Cadena, Pos_Ini + 2, Pos_Fin - Pos_Ini - 2)    '---extraigo la Cadena
'        Cadena = Func_Normalizar_Lista_Columnas(Cadena)                           '---si hay espacios, los quito
'        If Left(Cadena, 5) = "Error" Then
'            Form_Menu.TB_Informe = Cadena & vbCrLf & Now
'            Exit Sub
'        End If
'
'        Form_Menu.TB_Informe = "Lista de Columnas a visualizar: [[" & Cadena & "]]" & vbCrLf & Now
'        Form_Menu.TBx_Descripción = Left(Form_Menu.TBx_Descripción, Pos_Ini + 1) & Cadena & Mid(Form_Menu.TBx_Descripción, Pos_Fin)
'
'        Range(Columns(2), Columns(LastCol_Tb_Solicitudes)).Hidden = True    '---Oculta de la Columna 2 a la última
'
'        Pos_Ini = 1
'        Do
'            If InStr(",:", Mid(Cadena, Pos_Ini + 1, 1)) > 0 Then        ' Controlar si columna de 1 o 2 letras
'                Letra_Col = Mid(Cadena, Pos_Ini, 1)
'                Pos_Ini = Pos_Ini + 1
'            Else
'                Letra_Col = Mid(Cadena, Pos_Ini, 2)
'                Pos_Ini = Pos_Ini + 2
'            End If
'            If Mid(Cadena, Pos_Ini, 1) = ":" Then                       ' Controlo si ":"
'                Letra_Col = Letra_Col & ":"
'                Pos_Ini = Pos_Ini + 1
'
'                If InStr(",:", Mid(Cadena, Pos_Ini + 1, 1)) > 0 Then    ' Controlar si columna de 1 o 2 letras
'                    Letra_Col = Letra_Col & Mid(Cadena, Pos_Ini, 1)
'                    Pos_Ini = Pos_Ini + 1
'                Else
'                    Letra_Col = Letra_Col & Mid(Cadena, Pos_Ini, 2)
'                    Pos_Ini = Pos_Ini + 2
'                End If
'            Else
'            End If
'            Pos_Ini = Pos_Ini + 1
'            Columns(Letra_Col).Hidden = False
'        Loop While Pos_Ini < Len(Cadena)
'
'        LO_Tb_LS_CAS.DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1).Select        '  Me sitúo en la segunda celda del rango DataBodyRange de las celdas visibles que será la segunda celda de las primera fila visible.
'
'    End Sub     ' Rut_Mostrar_Columnas_Lista     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    '==============================================================================================
'

