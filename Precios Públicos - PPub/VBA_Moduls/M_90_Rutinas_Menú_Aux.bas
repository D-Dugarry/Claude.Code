Attribute VB_Name = "M_90_Rutinas_Menú_Aux"
Option Explicit

'==================================================================================================================================
'===================================================================================================================================
Sub RuT_Ejecutar_Rut(TaskRut As String, Optional Show_Msg As Boolean = True)
Debug.Print "RuT_Ejecutar_Rut"
    Dim TaskIndice      As Variant
    TaskIndice = Application.Match(TaskRut, Prog__Menú_Aux.ListObjects(1).DataBodyRange.Columns(Task_Nombre_Rut), 0)
    If IsError(TaskIndice) Then     ' ¡¡¡ NO Existe la Rutina !!! ------------------------
        MsgBx_Msg = "¡ No Existe la Tarea o su nombre ha cambiado !"
    Else                            ' ¡¡¡ Existe la Rutina !!! ------------------------
        Prog__APP.Range("APP_Task_Rut") = TaskRut
        Prog__APP.Range("APP_Task_Index") = TaskIndice
        '--------------
        Application.Run Prog__APP.Range("APP_Task_Rut").Value
        '--------------
        MsgBx_Msg = Prog__APP.Range("APP_Task_Inf")
        Prog__Menú_Aux.ListObjects(1).DataBodyRange.Cells(TaskIndice, Task_Rut_Informe) = Prog__APP.Range("APP_Task_Inf")
    End If
    If Show_Msg Then
        MsgBx_Title = "Informe del Proceso de Ejecutar la Tarea:  " & TaskRut
        Load Form_MsgBox: Call Form_MsgBox.SetParameter(16, True, "OK"): Form_MsgBox.Show  '- ([Font-Size]=16, [Red-Border]=False, [Buttons]="Ok", [Default-Button]=1, [Image]="Msg")
    End If
End Sub
'==================================================================================================================================
'===================================================================================================================================
Sub RuT_Load_Task_Data(TaskRut As String)
Debug.Print "RuT_Load_Task_Data"
    Dim Rutinas_Name    As String
    Dim TaskIndice      As Variant
    TaskIndice = Application.Match(TaskRut, Prog__Menú_Aux.ListObjects(1).DataBodyRange.Columns(Task_Nombre_Rut), 0)
        Prog__APP.Range("APP_Task_Rut") = TaskRut
        Prog__APP.Range("APP_Task_Index") = TaskIndice
        
        Prog__Menú_Aux.ListObjects(1).DataBodyRange.Cells(TaskIndice, Task_Rut_Informe) = Prog__APP.Range("APP_Task_Inf")
End Sub
'==================================================================================================================================
'===================================================================================================================================
Sub RuT_Save_Task_Data(TaskRut As String, Optional Task_Inf As String = "")
Debug.Print "RuT_Save_Task_Data"
    Dim Rutinas_Name    As String
    Dim TaskIndice      As Variant
    TaskIndice = Application.Match(TaskRut, Prog__Menú_Aux.ListObjects(1).DataBodyRange.Columns(Task_Nombre_Rut), 0)
    If IsError(TaskIndice) Then
        Debug.Print "RuT_Save_Task_Data - ERROR - Task NOT FOUND -------------------<<<"
    Else
        Prog__APP.Range("APP_Task_Rut") = TaskRut
        Prog__APP.Range("APP_Task_Index") = TaskIndice
        If Task_Inf = "" Then
            Prog__Menú_Aux.ListObjects(1).DataBodyRange.Cells(TaskIndice, Task_Rut_Informe) = Prog__APP.Range("APP_Task_Inf")
        Else
            Debug.Print Task_Inf
        Dim Rng     As Range:     Set Rng = Prog__Menú_Aux.ListObjects(1).DataBodyRange.Cells(TaskIndice, Task_Rut_Informe)
            Rng.Value = Task_Inf
        End If
    End If
End Sub
'==================================================================================================================================
'===================================================================================================================================
Sub Rut_Chg_Usuario()
Debug.Print "Rut_Chg_Usuario"
    Form_Usuario.Show
    Call RuT_Load_Task_Data("Rut_Chg_Usuario")
    MyRibbon.InvalidateControl "Btn_ChangeUser"     '- 1º el Botón, Actualiza solo este Control_ID
    MyRibbon.InvalidateControl "GroupChangeUser"    '- 2º el Grupo, Actualiza solo este Control_ID
End Sub
'===================================================================================================================================
Sub Rut_Activar_Programación()
Debug.Print "Rut_Activar_Programación"
    Call RuT_Al_Abrir_WorkBook
    Call Rut_Sheets_ShowAll
    Call Rut_ConfigExcel_RESTABLECER
    Prog__APP.Range("APP_Task_Inf") = "Estado de Programación Activado  -  " & Now
End Sub
'===================================================================================================================================
Sub Rut_Lo_Export_WorkSheet()
Debug.Print "Rut_Lo_Export_WorkSheet"
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
    Call Rut_Lo_Export_WS_Xlsx(Sheets(ActiveSheet.Name), ActiveSheet.Name & "__" & _
                        Format(Prog__APP.Range("APP_FechCierreCont"), "yyyy-mmm-dd") & "__" & Format(Now(), "(dd-mm-yy hh.mm)"), True, False)
    Prog__APP.Range("APP_Last_Exp_JIsPPub") = Format(Now(), "dd-mmm-yy hh:mm")
End Sub
'===================================================================================================================================
Sub Rut_Lo_Export_Hist_Bdatos()
Debug.Print "Rut_Lo_Export_Hist_Bdatos"
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
'    Call Rut_Lo_Export(Sht__BD, "Histórico_BDatos")
    Call Rut_Lo_Export_WS_Xlsx(Sht__BD, "Histórico_BDatos__a_" & _
                        Format(Prog__APP.Range("APP_FechCierreCont"), "yyyy-mmm") & "__" & Format(Now(), "(yyyy-mm-dd hhmm)"), False)
    Prog__APP.Range("APP_Last_BD_Export") = Format(Now(), "dd-mmm-yy hh:mm")
    MyRibbon.InvalidateControl "BtnExportBDatos"     '- 1º el Botón, Actualiza solo este Control_ID
End Sub
'===================================================================================================================================
Sub Rut_Lo_Import_Hist_Bdatos()
Debug.Print "Rut_Lo_Export_Bdatos"
    Set ActivForm = VBA.UserForms(VBA.UserForms.Count - 1)  '- Identificamos qué Formulario está Activo.  ----------
    Application.ScreenUpdating = False
    Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
    Rut_Off_Functions
    
    Sht__BD.Visible = xlSheetVisible
    Sht__BD.Select
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Sht__BD.Unprotect
    Lo_DefCol_BD.TotalsRowRange(DefC_HiddenCol) = False
    Prog__APP.Range("SW_Col_Hide_Sht__BD") = Not Prog__APP.Range("SW_Col_Hide_Sht__BD")
    Lo_BD.ShowTotals = False
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    
    Dim Arch_New_Name         As String:    Arch_New_Name = "Histórico_BDatos_" & Prog__APP.Range("APP_AñoCont") & "-"
    Call Rut_Lo_Import_LoData_LoDefCol(Lo_BD, Lo_DefCol_BD, DefC_TitColLstObj, Arch_New_Name)
    Prog__APP.Range("APP_Last_BD_Import") = Format(Now(), "dd-mmm-yy hh:mm")
    
    Sht__BD.Calculate
    Lo_DefCol_BD.ShowTotals = True
    Application.ScreenUpdating = True
    Rut_On_Functions
    MyRibbon.InvalidateControl "BtnImportBDatos"     '- 1º el Botón, Actualiza solo este Control_ID
End Sub
'===================================================================================================================================
Sub Rut_Reset_App()
    Call RuT_Al_Abrir_WorkBook
    Prog__APP.Range("APP_Task_Inf") = "App Reset" & vbCrLf & Now
    Call RuT_Load_Task_Data("Rut_Reset_App")
End Sub
' ==================================================================================================================================
Sub Rut_RibbonRefresh()
    Call Rut_Filtrar_Tareas
    Call RefreshRibbon
    Prog__APP.Range("APP_Task_Inf") = "RibbonX Refreshed " & Now
End Sub
'===================================================================================================================================
Sub Rut_Btn_Menú_Aux()
    Form_Menu.Show
End Sub
'===================================================================================================================================
Sub Rut_Columns_Show_All()
    Call Rut_Protect_Status_Save(ThisWorkbook.ActiveSheet)
    ActiveSheet.Unprotect
    Columns.EntireColumn.Hidden = False
    Call Rut_Protect_Status_Restore(ThisWorkbook.ActiveSheet)
    Prog__APP.Range("APP_Task_Inf") = "All columns are visible" & vbCrLf & Now
End Sub
'===================================================================================================================================
Sub Rut_RibbonX_ShowAll()
    Call Rut_Menú_ShowAll
    Prog__APP.Range("APP_Task_Inf") = "RibbonX Showed" & vbCrLf & Now
End Sub
'===================================================================================================================================
Sub Rut_OnOff_SW_Probando()
    If Prog__APP.Range("SW_Test") Then
        Prog__APP.Range("SW_Test") = False
        Form_Menu.Lb_SW_Test.Visible = False
        Prog__APP.Range("APP_Task_Inf") = "SW-B DesActivado" & vbCrLf & Now
    Else
        Prog__APP.Range("SW_Test") = True
        Form_Menu.Lb_SW_Test.Visible = True
        Prog__APP.Range("APP_Task_Inf") = "SW-B Activado" & vbCrLf & Now
    End If
End Sub
'===================================================================================================================================
Sub Rut_OnOff_SW_DelRegNeg()
    If Prog__APP.Range("SW_DelRegNeg") Then
        Prog__APP.Range("SW_DelRegNeg") = False
        Form_Menu.Lb_SW_DelRegNeg.Visible = False
        Prog__APP.Range("APP_Task_Inf") = "SW-B DesActivado" & vbCrLf & Now
    Else
        Prog__APP.Range("SW_DelRegNeg") = True
        Form_Menu.Lb_SW_DelRegNeg.Visible = True
        Prog__APP.Range("APP_Task_Inf") = "SW-B Activado" & vbCrLf & Now
    End If
End Sub
'===================================================================================================================================
Sub Rut_Sheets_ShowAll()
Dim WrkSht          As Worksheet
    For Each WrkSht In Worksheets
                WrkSht.Visible = xlSheetVisible
    Next
    ActiveWindow.DisplayWorkbookTabs = True
    Prog__APP.Range("APP_Task_Inf") = "All sheets Visible" & vbCrLf & Now
End Sub
'===================================================================================================================================
Sub Rut_Sheets_HideAll()     '- Hide All excep ActiveSheet  ---
Debug.Print "Rut_Sheets_HideAll", ActiveSheet.Name
Dim WrkSht          As Worksheet
    For Each WrkSht In Worksheets
        If WrkSht.Name <> ActiveSheet.Name Then WrkSht.Visible = xlSheetVeryHidden   'xlSheetVisible   '
    Next
    ActiveWindow.DisplayWorkbookTabs = False
    Prog__APP.Range("APP_Task_Inf") = "All sheets Hide" & vbCrLf & Now
End Sub
'===================================================================================================================================
Sub Rut_ProtectUnProtect_ActivSheet()
    If ActiveSheet.ProtectContents Then
        ActiveSheet.Unprotect
        Prog__APP.Range("APP_Task_Inf") = "ActiveSheet.UnProtect" & vbCrLf & Now
    Else
        ActiveSheet.Protect
        Prog__APP.Range("APP_Task_Inf") = "ActiveSheet.Protect" & vbCrLf & Now
    End If
End Sub
'===================================================================================================================================
Sub Rut_Events_Status_Change()        ' Para permitir las rutinas que se activan cuando ocurre un evento
    If Application.EnableEvents Then
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
        Prog__APP.Range("SW_Events") = False
        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is OFF" & vbCrLf & Now
    Else
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
        Prog__APP.Range("SW_Events") = True
        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is On" & vbCrLf & Now
    End If
End Sub
'===================================================================================================================================
    'Call Rut_Events_Status_Choose(Optional Choose As String = "CHANGE")    '- EnableEvents: ["CHANGE"], "ON", "OFF"
Sub Rut_Events_Status_Choose(Optional Choose As String = "CHANGE")      ' Para permitir las rutinas que se activan cuando ocurre un evento
    Select Case UCase(Choose)
        Case "CHANGE"
                        If Application.EnableEvents Then
                            Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
                            Prog__APP.Range("SW_Events") = False
                            Prog__APP.Range("APP_Task_Inf") = "Events Status Change is OFF" & vbCrLf & Now
                        Else
                            Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
                            Prog__APP.Range("SW_Events") = True
                            Prog__APP.Range("APP_Task_Inf") = "Events Status Change is On" & vbCrLf & Now
                        End If
        Case "ON"
                        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
                        Prog__APP.Range("SW_Events") = True
                        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is On" & vbCrLf & Now
        Case "OFF"
                        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
                        Prog__APP.Range("SW_Events") = False
                        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is OFF" & vbCrLf & Now
    End Select
End Sub
'===================================================================================================================================
'###################################################################################################################################
Sub RuT_Quitar_Sombreados()
    ActiveSheet.ListObjects(1).DataBodyRange.Interior.Color = -1
End Sub

'===================================================================================================================================
Sub Rut_x_Help_ShowHide()
'===================================================================================================================================
Debug.Print "Rut_x_Help_ShowHide   -------------- Por ahora está inabiltado"
MsgBox "Rut_x_Help_ShowHide   -------------- Por ahora está inabiltado"
'Call RuT_Load_Task_Data("Rut_x_Help_ShowHide")
'    Dim F_Help  As Integer
'    Dim Nombre  As String
'    Dim rng     As Range
'    Dim Lo_TPLqHelp        As ListObject
'    Set Lo_TPLqHelp = Prog__TitP_Liq_Help.ListObjects(1)
'
'    Wk_TitP_Liquid.Select
'    Wk_TitP_Liquid.Unprotect
'
'    If ActiveSheet.Comments.Count = 0 Then
'        With Lo_TPLqHelp.DataBodyRange      '-Asigno los comentarios
'            For F_Help = 1 To .Rows.Count
'                Nombre = .Cells(F_Help, 1)
'                Set rng = Range(Nombre)
'                Wk_TitP_Liquid.Range(rng.Address).AddComment .Cells(F_Help, 2).Value2
'                Wk_TitP_Liquid.Range(rng.Address).Comment.Shape.TextFrame.AutoSize = True
'            Next F_Help
'        End With
'        ActiveSheet.Shapes("Cuadro_Instrucciones").Visible = True
'        Prog__APP.Range("APP_Help_State") = "Activados"
'        Prog__APP.Range("APP_Task_Inf") =   "Estado de Comentarios de Celdas Activado  -  " & Now
'    Else
'        ActiveSheet.UsedRange.ClearComments
'        Prog__APP.Range("APP_Help_State") = "Desactivados"
'        Prog__APP.Range("APP_Task_Inf") =   "Estado de Comentarios de Celdas Desactivado  -  " & Now
'    End If
'        'MyRibbon.InvalidateControl "BtnHelpComments"    '- Actualiza solo este botón
'        'MyRibbon.InvalidateControl "GroupHelp"    '- Actualiza solo este botón
'    Wk_TitP_Liquid.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, userinterfaceonly:=True   '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
End Sub
'=======================================================================================================





'===================================================================================================================================
'Sub Rut_OnOff_SW_VerRecNeg()
'    Prog__APP.Range("APP_VerRecNeg") = Not Prog__APP.Range("APP_VerRecNeg")
'    If Prog__APP.Range("APP_VerRecNeg") Then Range("Liquid_APP_VerRecNeg") = "Hide Rec.Neg." Else Range("Liquid_APP_VerRecNeg") = "Ver Rec.Neg."
'    Call Rut_Liquid_EP_Load(UCase(Range("Liquid_Plan")), Prog__APP.Range("APP_CursAcad"))
'    Prog__APP.Range("APP_Task_Inf") =   "Visualizar registros negativos: " & Prog__APP.Range("APP_VerRecNeg") & vbCrLf & Now
'End Sub
'===================================================================================================================================
'===================================================================================================================================
'Sub Rut_Mostrar_Col_Ocultas()       ' CuadroTextoVistaColmns    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Dim Cont_Col As Integer
'Solicitudes.Unprotect
'Application.ScreenUpdating = False
'        For Cont_Col = 1 To LastCol_Tb_Solicitudes
'            If Lo_Prog_Colns.DataBodyRange.Cells(7, Cont_Col) = "Ocultar" Then Columns(Cont_Col).Hidden = False
'        Next Cont_Col
'Application.ScreenUpdating = True
'Solicitudes.Protect
'Prog__APP.Range("APP_Task_Inf") =   "Columnas Ocultas al Usuario Visibles" & vbCrLf & Now
'End Sub
'==================================================================================================================================


'    '=============================  Para mostrar u ocultar Columnas según Lista [[lista]]   ============================================
'    '===================================================================================================================================
'    Sub Rut_Mostrar_Columnas_Lista()       ' CuadroTextoVistaColmns    <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
'            Prog__APP.Range("APP_Task_Inf") =   "Error: No hay lista de Columnas a visualizar, o no empieza por [[, o no acaba por ]]." & vbCrLf & Now
'            Exit Sub
'        End If
'
'        Cadena = Mid(Cadena, Pos_Ini + 2, Pos_Fin - Pos_Ini - 2)    '---extraigo la Cadena
'        Cadena = Func_Normalizar_Lista_Columnas(Cadena)                           '---si hay espacios, los quito
'        If Left(Cadena, 5) = "Error" Then
'            Prog__APP.Range("APP_Task_Inf") =   Cadena & vbCrLf & Now
'            Exit Sub
'        End If
'
'        Prog__APP.Range("APP_Task_Inf") =   "Lista de Columnas a visualizar: [[" & Cadena & "]]" & vbCrLf & Now
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
'    End Sub     ' Rut_Mostrar_Columnas_Lista     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    '==================================================================================================================================
'

