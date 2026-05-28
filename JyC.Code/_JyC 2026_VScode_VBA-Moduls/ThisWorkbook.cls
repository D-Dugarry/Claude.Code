VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
' ThisWorkbook
'
' Gestiona los eventos principales del libro: apertura, cierre, desactivacion y menu contextual.
'
' EVENTOS:
'   Workbook_Open              -> Inicializa la aplicacion y la deja en primer plano
'   Workbook_Activate          -> Refresca el ribbon al recuperar el foco
'   Workbook_Deactivate        -> Restaura config de Excel al perder el foco
'   Workbook_BeforeClose       -> Restaura interfaz de Excel al estado estandar
'   Workbook_SheetBeforeRightClick -> Controla el menu contextual
'
' PROPIEDADES:
'   ribbonUI   -> Referencia al objeto IRibbonUI (ribbon personalizado via RibbonX)
'                 Se asigna en OnLoad_MyRibbon (M_000_RibbonUI) y se usa en RefreshRibbon
'   xl_edit1   -> Valor del campo EditBox del ribbon (persiste entre invalidaciones)
'
' SWITCHES EN Prog__APP (Config_APP):
'   SW_Events        -> TRUE = eventos de hoja habilitados
'   SW_Test          -> TRUE = modo prueba activo (correos redirigidos a DIR_PRUEBA)
'   SW_RightClick    -> TRUE = menu contextual habilitado en todas las hojas
'   SW_WB_Deactivate -> TRUE = Workbook_Deactivate llama a Rut_ConfigExcel_Restablecer
'                       FALSE = desactivado temporalmente (ej: durante Fnc_RangeToHTML)
'
' NOTA RIBBON (RibbonX):
'   RefreshRibbon se llama con Application.OnTime (+1 seg) para evitar el error
'   "Ribbon UI Refresh Failed" que ocurre cuando Invalidate se ejecuta mientras
'   Application.Visible = False durante la secuencia de apertura.
'==================================================================================================

Private myRibUI     As IRibbonUI
Private v_edit_1    As String

'--------------------------------------------------------------------------------------------------
' ribbonUI  (Property Let/Get)
'   Almacena la referencia al objeto IRibbonUI del ribbon personalizado.
'   Se asigna desde OnLoad_MyRibbon en M_000_RibbonUI cuando el ribbon carga.
'--------------------------------------------------------------------------------------------------
Public Property Let ribbonUI(irib As IRibbonUI)
    Set myRibUI = irib
End Property
Public Property Get ribbonUI() As IRibbonUI
    Set ribbonUI = myRibUI
End Property

'--------------------------------------------------------------------------------------------------
' xl_edit1  (Property Let/Get)
'   Almacena el valor del campo EditBox del ribbon.
'   Persiste entre llamadas a Invalidate (que borrararia el valor si no se guardara aqui).
'--------------------------------------------------------------------------------------------------
Public Property Let xl_edit1(xval As String)
    v_edit_1 = xval
End Property
Public Property Get xl_edit1() As String
    xl_edit1 = v_edit_1
End Property

'==================================================================================================
' Workbook_Open
'   Se ejecuta automaticamente al abrir el libro.
'   Objetivo: inicializar la aplicacion y dejarla en primer plano sobre cualquier otro Excel.
'
'   SECUENCIA:
'   1. Minimiza otros libros Excel
'   2. Oculta la app temporalmente (evita parpadeos)
'   3. Inicializa switches de control
'   4. Configura menu contextual y oculta ribbon nativo
'   5. Oculta todas las hojas excepto H_INICI
'   6. Llama a RuT_Al_Abrir_WorkBook (inicializacion principal)
'   7. Restaura visibilidad y activa pantalla completa
'   8. Programa RefreshRibbon con OnTime (+1 seg) para evitar error RibbonX
'==================================================================================================
Private Sub Workbook_Open()
Debug.Print "Workbook_Open >>> ThisWorkbook"

    Dim WrkSht As Worksheet

    On Error GoTo Err_Open

    ' --- 1. Minimizar otros libros Excel para quedar en primer plano ---
    Call Rut_WrkBook_MinimizeAllExcelExceptThisWB

    ' --- 2. Suprimir pantalla durante la inicializacion ---
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    Application.Visible = False      ' Ocultar temporalmente para evitar parpadeos

    ' --- 3. Inicializar switches de control en Prog__APP ---
    Prog__APP.Range("SW_Events") = True
    Prog__APP.Range("SW_Test") = False
    Prog__APP.Range("SW_RightClick") = True
    Prog__APP.Range("SW_WB_Deactivate") = True

    ' --- 4. Configurar menu contextual y ribbon ---
    Application.CommandBars("Ply").Enabled = True
    Call Rut_Context_Buttons_Hide                                      ' Oculta opciones genericas del menu contextual
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"    ' startFromScratch=true en el XML ya elimina tabs nativos; True = ribbon expandido

    ' --- 5. Mostrar solo la hoja de inicio ---
    H_INICI.Visible = xlSheetVisible
    For Each WrkSht In Worksheets
        If WrkSht.CodeName <> "H_INICI" Then WrkSht.Visible = xlSheetVeryHidden
    Next WrkSht

    ' --- 6. Inicializacion principal de la aplicacion ---
    Call RuT_Al_Abrir_WorkBook

    ' --- 7. Restaurar visibilidad y activar pantalla completa ---
    Application.WindowState = xlMaximized   ' DisplayFullScreen ocultaria el ribbon
    Application.Visible = True
    Application.ScreenUpdating = True

    ' --- 8. Refrescar ribbon personalizado (diferido 1 seg para evitar error RibbonX) ---
    '        Si se llama directamente con Application.Visible=False, MyRibbon.Invalidate falla.
    Call Rut_Menú_ShowAll   ' Configura estados del semaforo y visibilidad del ribbon

    Application.OnTime Now + TimeValue("00:00:01"), "RefreshRibbon"

Debug.Print "Workbook_Open <<< ThisWorkbook"
    Exit Sub

Err_Open:
    ' Garantizar que la aplicacion vuelve a ser visible aunque falle la inicializacion
    Application.Visible = True
    Application.ScreenUpdating = True
    MsgBox "Error en Workbook_Open:" & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error al abrir"

End Sub

'==================================================================================================
' Workbook_Activate
'   Se ejecuta cuando el libro recupera el foco.
'   Refresca el ribbon por si ha quedado desactualizado al volver desde otra aplicacion.
'   On Error Resume Next: si el ribbon no esta listo, sale silenciosamente.
'==================================================================================================
Private Sub Workbook_Activate()
Debug.Print "Workbook_Activate"
    On Error Resume Next
    Call RefreshRibbon
    On Error GoTo 0
End Sub

'==================================================================================================
' Workbook_Deactivate
'   Se ejecuta cuando el libro pierde el foco.
'   Restaura la configuracion de Excel si SW_WB_Deactivate = True.
'
'   SW_WB_Deactivate se pone a FALSE temporalmente en Fnc_RangeToHTML (M_820_Range_TO_HTML)
'   para evitar que Rut_ConfigExcel_Restablecer interfiera con el PasteSpecial interno
'   de esa funcion. Se restaura a TRUE al terminar.
'==================================================================================================
Private Sub Workbook_Deactivate()
    If Prog__APP.Range("SW_WB_Deactivate") Then
Debug.Print "Workbook_Deactivate"
        Call Rut_ConfigExcel_Restablecer
    End If
End Sub

'==================================================================================================
' Workbook_BeforeClose
'   Se ejecuta antes de cerrar el libro.
'   Restaura todos los elementos de interfaz de Excel al estado estandar
'   para no afectar a otros libros que puedan estar abiertos.
'==================================================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Debug.Print "Workbook_BeforeClose <<< ThisWorkbook"

    ' --- Restaurar menus contextuales ---
    Application.CommandBars("Ply").Enabled = True
    Application.CommandBars("Cell").Reset
    Application.CommandBars("List Range Popup").Reset

    ' --- Restaurar eventos y actualizacion de pantalla ---
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' --- Restaurar elementos de interfaz de la aplicacion ---
    Application.DisplayStatusBar = True
    Application.DisplayFormulaBar = True
    Application.WindowState = xlNormal      ' Restaurar tamano de ventana

    ' --- Restaurar elementos de la ventana activa ---
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True

    ' --- Restaurar ribbon nativo de Excel ---
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"

    ' --- Dejar la hoja de inicio visible y seleccionada ---
    H_INICI.Visible = xlSheetVisible
    H_INICI.Select

End Sub

'==================================================================================================
' Workbook_SheetBeforeRightClick
'   Se ejecuta al hacer clic derecho en cualquier hoja del libro.
'   Controla la disponibilidad del menu contextual segun SW_RightClick:
'     TRUE  -> menu contextual habilitado (Cancel = False)
'     FALSE -> menu contextual deshabilitado (Cancel = True)
'
'   NOTA: Para controlar el Right-Click en una hoja concreta individualmente,
'   anadir este codigo en el modulo de esa hoja:
'       Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'           Cancel = Not Prog__APP.Range("SW_RightClick")
'       End Sub
'==================================================================================================
Private Sub Workbook_SheetBeforeRightClick(ByVal SheetX As Object, ByVal Target As Range, Cancel As Boolean)
Debug.Print "Workbook_SheetBeforeRightClick"
    Cancel = Not Prog__APP.Range("SW_RightClick")
End Sub


