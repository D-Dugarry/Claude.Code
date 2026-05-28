Attribute VB_Name = "M_0_Ini_APP"
' ==============================================================================
' Módulo     : M_0_Ini_APP
' Proyecto   : Mailing Indicadores UA
' Autor      : Dugarry
' Descripción: Inicialización de la interfaz de la aplicación al abrir el libro.
'
' Subrutinas públicas:
'   · Rut_Iniciar_APP() — Configura la UI de Excel al arrancar
' ==============================================================================
Option Explicit

' ==============================================================================
' Rut_Iniciar_APP
' ------------------------------------------------------------------------------
' Configura la interfaz de Excel al abrir el libro: pantalla completa,
' oculta barra de fórmulas, cuadrícula, saltos de página y ribbon.
' Llamada desde ThisWorkbook.Workbook_Open().
' ==============================================================================
Sub Rut_Iniciar_APP()
    Application.EnableEvents   = True
    Application.ScreenUpdating = False          ' Evita parpadeo durante la inicialización

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False

    ThisWorkbook.Windows(1).DisplayGridlines = False

    ' Ocultar saltos de página en todas las hojas del libro
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.DisplayPageBreaks = False
    Next ws

    ' Minimizar ribbon si está expandido (altura > 1 px indica ribbon expandido)
    On Error Resume Next
    If CommandBars("Ribbon").Controls(1).Height > 1 Then
        CommandBars.ExecuteMso "MinimizeRibbon"
    End If
    On Error GoTo 0

    Application.ScreenUpdating = True
End Sub     ' Rut_Iniciar_APP
' ==============================================================================
