Attribute VB_Name = "M_90_Rutinas_Menu_Aux"
' Last Rev. 2026-10-07 10:49
Option Explicit

'- Estado de protección de las hojas desprotegidas con el botón Protect/UnProtect
'  del Ribbon (una entrada por hoja, por CodeName), para devolverlas con los mismos permisos.
'  Ver Rut_Ws_Protect_Status.
Private Prot_Hojas(1 To 60)    As String
Private Prot_Estados(1 To 60)  As T_Prot_Estado
Private Progreso_Hoja           As String     '- hoja activa al abrir el progreso (Rut_Progreso_Cerrar con Volver_Hoja)

'==================================================================================================================================
'===================================================================================================================================
'- Formulario de progreso de los botones del Ribbon (2026-10-04, patron de Jornadas y Congresos).
'  El boton abre Form_Running_Rut, llama DIRECTAMENTE a su rutina y lo cierra:
'
'      Call Rut_Progreso_Abrir("Titulo de la tarea")
'      Call RuT_Update_LSGES04_ACont               '- escribe su progreso en ActivForm (= Form_Running_Rut)
'      Call Rut_Progreso_Cerrar(control.Tag)       '- informe a APP_Task_Inf y a Lo_RibbonUI, fondo verde
'
'  Sustituye a RuT_Ejecutar_Rut, RuT_Load_Task_Data y RuT_Save_Task_Data, y al camino viejo: el
'  boton dejaba el nombre de la rutina en APP_Task_Rut y Form_Running_Rut la buscaba en Tb_Tareas
'  (_Menu_Aux) y la lanzaba con Application.Run. Como el cierre lo hace el boton, cualquier salida
'  de la rutina (cancelar un dialogo incluido) deja visible el boton de salida del formulario.
Public Sub Rut_Progreso_Abrir(ByVal Titulo As String)
    Progreso_Hoja = ThisWorkbook.ActiveSheet.Name
    Application.ScreenUpdating = False          '- ANTES del Show: si se apaga justo despues (lo hacen las
                                                '  rutinas al empezar) el formulario no llega a pintarse
    Load Form_Running_Rut
    Set ActivForm = Form_Running_Rut
    Form_Running_Rut.Lb_Tit_Informe.Caption = "Progreso de la Tarea: " & Titulo
    Form_Running_Rut.TBx_Informe = ""
    Form_Running_Rut.Show vbModeless            '- vbModeless: el Show devuelve el control al boton
    DoEvents                                    '- que se pinte antes del proceso largo
End Sub
'==================================================================================================
Public Sub Rut_Progreso_Cerrar(Optional ByVal Tag As String = "", Optional ByVal Volver_Hoja As Boolean = False)
    On Error Resume Next                        '- el cierre no debe fallar aunque la rutina haya dejado algo a medias
    Prog__APP.Range("APP_Task_Inf") = Form_Running_Rut.TBx_Informe.Text   '- .Text: con el control, 1004 en logs largos
    If Len(Tag) > 0 Then Call Rut_RibbonUI_Guardar_Informe(Tag)
    If Volver_Hoja Then ThisWorkbook.Sheets(Progreso_Hoja).Select
    If Fnc_Get_NestLevel() > 0 Then Call Rut_Reset_State   '- red de seguridad: la rutina se salto algun Rut_On_Functions
    Call Rut_Lo_Totales_Mostrar                 '- BDatos, Inf_Recibos_TIO y las DefCol, con la fila de totales (Docs/Plan_ShowTotals.md, fase 4)
    Application.ScreenUpdating = True
    Form_Running_Rut.Rut_Finalizada             '- fondo verde y boton de salida
    Call RefreshRibbon                          '- rotulos (fecha de la ultima importacion...) y supertips al dia
    On Error GoTo 0
End Sub
'==================================================================================================
'- Feedback del modulo de copias Rut_Wb_CopSegTimed_USB_HD (skill excel-copseg-backup), que la llama
'  por su nombre con Application.Run tras cada copia o error: lo deja en APP_Task_Inf, como hacia la
'  copia USB antigua. Si se renombra, el modulo de copias deja de llamarla sin avisar.
Public Sub Rut_CopSeg_Feedback_Host(Msg As String)
    Prog__APP.Range("APP_Task_Inf") = Msg
End Sub
'==================================================================================================================================
'===================================================================================================================================
Sub Rut_Activar_Programacion()
Debug.Print "Rut_Activar_Programacion"
    Call RuT_Al_Abrir_WorkBook
    Call Rut_Sheets_ShowAll
    Call Rut_ConfigExcel_RESTABLECER
    Prog__APP.Range("APP_Task_Inf") = "Estado de Programación Activado  -  " & Now
End Sub
'===================================================================================================================================
Sub Rut_Lo_Export_WorkSheet()
Debug.Print "Rut_Lo_Export_WorkSheet"
    Set ActivForm = Form_Running_Rut                      '- lo abre el boton del Ribbon (Rut_Progreso_Abrir)
    Call Rut_Lo_Export_WS_Xlsx(Sheets(ActiveSheet.Name), ActiveSheet.Name & "__" & _
                        Format(Prog__APP.Range("APP_FechCierreCont"), "yyyy-mmm-dd") & "__" & Format(Now(), "(dd-mm-yy hh.mm)"), True, False)
    Prog__APP.Range("APP_Last_Exp_JIsPPub") = Format(Now(), "dd-mmm-yy hh:mm")
End Sub
'===================================================================================================================================
Sub Rut_Lo_Export_Hist_Bdatos()
Debug.Print "Rut_Lo_Export_Hist_Bdatos"
    Set ActivForm = Form_Running_Rut                      '- lo abre el boton del Ribbon (Rut_Progreso_Abrir)
'    Call Rut_Lo_Export(Sht__BD, "Histórico_BDatos")
    Call Rut_Lo_Export_WS_Xlsx(Sht__BD, "Histórico_BDatos__a_" & _
                        Format(Prog__APP.Range("APP_FechCierreCont"), "yyyy-mmm") & "__" & Format(Now(), "(yyyy-mm-dd hhmm)"), False)
    Prog__APP.Range("APP_Last_BD_Export") = Format(Now(), "dd-mmm-yy hh:mm")
    MyRibbon.InvalidateControl "BtnExportBDatos"     '- 1º el Botón, Actualiza solo este Control_ID
End Sub
'===================================================================================================================================
Sub Rut_Lo_Import_Hist_Bdatos()
Debug.Print "Rut_Lo_Export_Bdatos"
    Set ActivForm = Form_Running_Rut                      '- lo abre el boton del Ribbon (Rut_Progreso_Abrir)
    Application.ScreenUpdating = False
    Dim Lo_BD               As ListObject:      Set Lo_BD = Sht__BD.ListObjects(1)
    Dim Lo_DefCol_BD        As ListObject:      Set Lo_DefCol_BD = Prog_DefCol_BD.ListObjects(1)
    Rut_Off_Functions
    
    Sht__BD.Visible = xlSheetVisible
    Sht__BD.Select
    Call Rut_Lo_WrkSht_Preparar(Sht__BD)
    Sht__BD.Unprotect
    Call Rut_ColHide_Set("Sht__BD", Not Fnc_ColHide_Get("Sht__BD"))
    
    H_Inicio = Timer                '- Para saber el tiempo de proceso
    LastTimeLap = Timer             '- Para saber tiempos intermedios
    
    Dim Arch_New_Name         As String:    Arch_New_Name = "Histórico_BDatos_" & Prog__APP.Range("APP_AnoCont") & "-"
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
    Call Rut_Reset_State            '- Deja a 0 el contador de Rut_Off/On_Functions
    Call RuT_Al_Abrir_WorkBook
    Prog__APP.Range("APP_Task_Inf") = "App Reset" & vbCrLf & Now
    Call Rut_RibbonUI_Guardar_Informe("Reset")     '- informe en Lo_RibbonUI (supertip de Boss)
End Sub
' ==================================================================================================================================
Sub Rut_RibbonRefresh()
    Call RefreshRibbon
    Prog__APP.Range("APP_Task_Inf") = "RibbonX Refreshed " & Now
End Sub
'===================================================================================================================================
Sub Rut_Btn_Menu_Aux()
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
Sub Rut_OnOff_SW_DelRegNeg()
    If Prog__APP_Switch.Range("Sw_DelRegNeg") Then
        Prog__APP_Switch.Range("Sw_DelRegNeg") = False
        Form_Menu.Lb_SW_DelRegNeg.Visible = False
        Prog__APP.Range("APP_Task_Inf") = "SW-B DesActivado" & vbCrLf & Now
    Else
        Prog__APP_Switch.Range("Sw_DelRegNeg") = True
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
'- Alterna la protección de una hoja SIN perder sus permisos (filtrar, ordenar...): al desprotegerla
'  los anota (Rut_Prot_Save) y al volver a protegerla le devuelve esos mismos (Rut_Prot_Restore).
'  El estado es POR HOJA (Prot_Hojas/Prot_Estados): desproteger una, cambiar a otra y proteger no
'  mezcla sus permisos. Si no hay nada anotado de esa hoja (nunca se desprotegió por aquí, o un
'  error + "Fin" vació las variables), aplica el esquema por defecto de la app.
'  Devuelve True si la hoja queda protegida. La usa el botón Protect/UnProtect
'  del Ribbon (OnAct_SheetProtect, M___RibbonUI).
Function Fnc_ProtectUnProtect_Hoja(ByVal Hoja As Worksheet) As Boolean
    Dim Slot    As Long:    Slot = Fnc_Prot_Slot(Hoja.CodeName)
    If Fnc_Prot_Hoja_Protegida(Hoja) Then                   '- no basta ProtectContents (ver Rut_Ws_Protect_Status)
        If Slot > 0 Then
            Call Rut_Prot_Save(Hoja, Prot_Estados(Slot))    '- anota los permisos de ESTA hoja y la desprotege
        Else
            Hoja.Unprotect
        End If
    Else
        If Slot > 0 Then Call Rut_Prot_Restore(Hoja, Prot_Estados(Slot), True)    '- True = UserInterfaceOnly
        If Not Fnc_Prot_Hoja_Protegida(Hoja) Then           '- nada anotado de esta hoja: esquema por defecto
            Hoja.Protect AllowFiltering:=True, AllowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True
        End If
    End If
    Fnc_ProtectUnProtect_Hoja = Fnc_Prot_Hoja_Protegida(Hoja)
End Function
'- Casilla de Prot_Estados de una hoja: la suya si ya tiene, si no la primera libre. 0 = tabla llena.
Private Function Fnc_Prot_Slot(ByVal CodeName As String) As Long
    Dim i As Long, Libre As Long
    For i = LBound(Prot_Hojas) To UBound(Prot_Hojas)
        If Prot_Hojas(i) = CodeName Then Fnc_Prot_Slot = i: Exit Function
        If Libre = 0 And Len(Prot_Hojas(i)) = 0 Then Libre = i
    Next i
    If Libre > 0 Then Prot_Hojas(Libre) = CodeName
    Fnc_Prot_Slot = Libre
End Function
'===================================================================================================================================
Sub Rut_Events_Status_Change()        ' Para permitir las rutinas que se activan cuando ocurre un evento
    If Application.EnableEvents Then
        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
        Prog__APP_Switch.Range("Sw_Events") = False
        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is OFF" & vbCrLf & Now
    Else
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
        Prog__APP_Switch.Range("Sw_Events") = True
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
                            Prog__APP_Switch.Range("Sw_Events") = False
                            Prog__APP.Range("APP_Task_Inf") = "Events Status Change is OFF" & vbCrLf & Now
                        Else
                            Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
                            Prog__APP_Switch.Range("Sw_Events") = True
                            Prog__APP.Range("APP_Task_Inf") = "Events Status Change is On" & vbCrLf & Now
                        End If
        Case "ON"
                        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
                        Prog__APP_Switch.Range("Sw_Events") = True
                        Prog__APP.Range("APP_Task_Inf") = "Events Status Change is On" & vbCrLf & Now
        Case "OFF"
                        Application.EnableEvents = False                       ' INHABILITA LOS EVENTOS
                        Prog__APP_Switch.Range("Sw_Events") = False
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
'        Cadena = Form_Menu.Tbx_Descripcion
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
'        Form_Menu.Tbx_Descripcion = Left(Form_Menu.Tbx_Descripcion, Pos_Ini + 1) & Cadena & Mid(Form_Menu.Tbx_Descripcion, Pos_Fin)
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
