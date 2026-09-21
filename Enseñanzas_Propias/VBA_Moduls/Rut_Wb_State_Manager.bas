Attribute VB_Name = "Rut_Wb_State_Manager"
' Last Rev. 2026-09-21 12:12
'===================================================================================================
' Rut_Wb_State_Manager
'
' Gestión centralizada del estado de Excel (pantalla, cálculo, eventos) con soporte
' de reentrancia mediante contador de anidamiento.
'
' PROBLEMA QUE RESUELVE:
'   Si una rutina A llama a Rut_Off_Functions y luego B (llamada por A) llama a
'   Rut_On_Functions, el estado se restaura antes de que A termine. Con este módulo,
'   solo se restaura cuando se sale del nivel más externo.
'
' NOTA (Enseñanzas_Propias): además del estado real de Application, este módulo
' mantiene sincronizado Prog__APP_Switch.Range("Sw_EnableEvents") como efecto
' colateral, porque varios Worksheet_Change/SelectionChange (Wk_TitP_Liquid,
' Sht__Inf_EPs_UXXI, Sht__Inf_EP_Rsm_1) y el checkbox de Form_Menu lo consultan
' directamente como guard de re-entrada. La FUENTE DE VERDAD para la restauración
' sigue siendo el estado real capturado por Rut_Off_Functions (m_PrevEvents), no el
' switch de hoja.
'
' SUBS/FUNCTIONS PÚBLICAS:
'   Rut_Off_Functions   -> Desactiva pantalla/cálculo/eventos (en el nivel 0)
'   Rut_On_Functions    -> Reactiva (solo en el nivel 0)
'
'===================================================================================================
Option Explicit

Private m_NestLevel     As Long         ' Contador de anidamiento (0 = libre)
Private m_PrevCalc      As XlCalculation
Private m_PrevScreen    As Boolean
Private m_PrevEvents    As Boolean

'---------------------------------------------------------------------------------------------------
' Rut_Off_Functions
'   Desactiva: recálculo (manual), pantalla, eventos.
'   Si es la primera llamada (nivel 0), guarda el estado previo.
'   Las llamadas sucesivas (nivel > 0) solo incrementan el contador.
'---------------------------------------------------------------------------------------------------
Public Sub Rut_Off_Functions()
    If m_NestLevel = 0 Then
        m_PrevCalc = Application.Calculation
        m_PrevScreen = Application.ScreenUpdating
        m_PrevEvents = Application.EnableEvents

        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = False        ' Bonus: congelar la barra de estado
        Prog__APP_Switch.Range("Sw_EnableEvents") = False    ' Mantiene sincronizados los guards de hoja
    End If
    m_NestLevel = m_NestLevel + 1
End Sub

'---------------------------------------------------------------------------------------------------
' Rut_On_Functions
'   Solo en el nivel 0 (salida más externa), restaura el estado guardado.
'   Las llamadas internas simplemente decrementan el contador.
'---------------------------------------------------------------------------------------------------
Public Sub Rut_On_Functions()
    If m_NestLevel > 0 Then m_NestLevel = m_NestLevel - 1

    If m_NestLevel = 0 Then
        Application.Calculation = m_PrevCalc
        Application.ScreenUpdating = m_PrevScreen
        Application.EnableEvents = m_PrevEvents
        Application.DisplayStatusBar = True
        Application.DisplayAlerts = True
        Prog__APP_Switch.Range("Sw_EnableEvents") = m_PrevEvents    ' Mantiene sincronizados los guards de hoja
    End If
End Sub

'---------------------------------------------------------------------------------------------------
' Fnc_Get_NestLevel (utilidad de depuración)
'   Devuelve el nivel actual de anidamiento.
'---------------------------------------------------------------------------------------------------
Public Function Fnc_Get_NestLevel() As Long
    Fnc_Get_NestLevel = m_NestLevel
End Function
