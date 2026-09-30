Attribute VB_Name = "Rut_Wb_State_Manager"
' Last Rev. 2026-09-30 15:05
'===================================================================================================
' Rut_Wb_State_Manager  (skill excel-state-manager, adaptado a PPub)
'
' Gestión centralizada del estado de Excel (cálculo, pantalla, eventos) con contador de
' anidamiento: una rutina anidada que llama a Rut_On_Functions ya NO reactiva la pantalla ni
' el recálculo antes de que termine la rutina exterior. Solo la salida del nivel más
' externo restaura el estado.
'
' Adaptación a PPub respecto al skill original:
'   - Al restaurar NO se vuelve al 'estado previo': se aplica la política de la app (cálculo
'     según SW_App_Calculation, eventos activos y SW_Events sincronizado), igual que hacía
'     el antiguo Rut_On_Functions de M_000_Ini_APP (de donde se han movido ambas rutinas).
'   - Rut_Off_Functions aplica el apagado SIEMPRE, también en niveles anidados (como antes).
'   - No toca DisplayStatusBar: PPub la mantiene oculta (ver Rut_ConfigExcel_Establecer).
'   - Un Rut_On_Functions 'suelto' (nivel 0) restaura igualmente, como antes.
'   - Rut_Reset_State: red de seguridad si algún camino se salta su Rut_On_Functions
'     (la llaman RuT_Ejecutar_Rut al terminar cada tarea y Rut_Reset_App).
'
' REGLA: cada Rut_Off_Functions necesita su Rut_On_Functions en TODOS los caminos de salida
'        (Exit Sub incluidos). Depurar con  ? Fnc_Get_NestLevel()  en Inmediato (0 = libre).
'===================================================================================================
Option Explicit

Private m_NestLevel     As Long         ' Contador de anidamiento (0 = libre)

'---------------------------------------------------------------------------------------------------
' Rut_Off_Functions: pausa cálculo, pantalla y eventos, y sube un nivel de anidamiento.
'---------------------------------------------------------------------------------------------------
Public Sub Rut_Off_Functions()
If m_NestLevel = 0 Then Debug.Print "Rut_Off_Functions"
    m_NestLevel = m_NestLevel + 1
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False:     Prog__APP.Range("SW_Events") = False               ' DesHABILITA LOS EVENTOS (SIEMPRE, sin depender del switch)
End Sub

'---------------------------------------------------------------------------------------------------
' Rut_On_Functions: baja un nivel; solo en la salida más externa restaura el estado.
'---------------------------------------------------------------------------------------------------
Public Sub Rut_On_Functions()
    If m_NestLevel > 1 Then
        m_NestLevel = m_NestLevel - 1       ' Nivel anidado: restaurará la rutina exterior
        Exit Sub
    End If
    m_NestLevel = 0
    Call Rut_Restaurar_Estado
End Sub

'---------------------------------------------------------------------------------------------------
' Rut_Reset_State: fuerza el contador a 0 y restaura (red de seguridad).
'---------------------------------------------------------------------------------------------------
Public Sub Rut_Reset_State()
    If m_NestLevel > 0 Then Debug.Print "!!! Rut_Reset_State: Off/On desbalanceado (nivel " & m_NestLevel & "), se fuerza a 0"
    m_NestLevel = 0
    Call Rut_Restaurar_Estado
End Sub

'---------------------------------------------------------------------------------------------------
' Fnc_Get_NestLevel: nivel actual de anidamiento (utilidad de depuración).
'---------------------------------------------------------------------------------------------------
Public Function Fnc_Get_NestLevel() As Long
    Fnc_Get_NestLevel = m_NestLevel
End Function

'---------------------------------------------------------------------------------------------------
' Rut_Restaurar_Estado: política de restauración de PPub (cuerpo del antiguo Rut_On_Functions).
'---------------------------------------------------------------------------------------------------
Private Sub Rut_Restaurar_Estado()
    If Prog__APP.Range("SW_App_Calculation") Then
        Application.Calculation = xlCalculationAutomatic
    Else
         Application.Calculation = xlCalculationManual
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If Prog__APP.Range("SW_Events") Then
        Application.EnableEvents = True
    Else
        Application.EnableEvents = True                       ' HABILITA LOS EVENTOS
        Prog__APP.Range("SW_Events") = True
    End If
Debug.Print "Rut_On_Functions"
End Sub
