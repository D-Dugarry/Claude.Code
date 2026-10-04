Attribute VB_Name = "M___RibbonUI_Rules"
' Last Rev. 2026-10-04 13:09
'==================================================================================================
' M___RibbonUI_Rules
'
' Motor de visibilidad y supertips del Ribbon, portado de M_003_RibbonUI_Rules de Jornadas y
' Congresos (2026-10-04). Lo usan GetVsbl_CtrlTab y getStip_CtrlTab (M___RibbonUI).
'
' FUENTE: la tabla Lo_RibbonUI de la hoja Prog__RibbonUI (pestana "RibbonUI"), una fila por Tag
' de control: Uribbon-Tags, Usuario, SheetsNames, Nom_Rut, Descripion, Informe_Rut, Group-Tag
' (constantes Rib_* en M_000_Ini_Var_APP). Hasta el 2026-10-04 lo decidia la columna Visible de
' Tb_Tareas (_Menu_Aux), que Rut_Filtrar_Tareas reescribia en cada activacion de hoja, y
' Func_CtrlTab_View la leia casando el Tag POR PREFIJO (Like Tag & "*", sensible a mayusculas).
' Ahora el Tag se compara ENTERO y sin distinguir mayusculas: cada control necesita su fila.
'
' Listas Usuario y SheetsNames: elementos separados por "," o ";", sin distinguir mayusculas;
' lista vacia = sin restriccion (Fnc_Lista_Contiene). Un elemento que no es ninguna hoja ("no")
' deja el control oculto en todas.
'
' PARA CAMBIAR LA VISIBILIDAD O EL SUPERTIP DE UN BOTON se edita su fila en Lo_RibbonUI; se aplica
' al pulsar Ribbon Refresh o al cambiar de hoja (RefreshRibbon llama a Rules_Reset). Un Tag sin
' fila queda OCULTO. Si la tabla no se puede leer se ve TODO: es deliberado, un ribbon vacio deja
' al usuario sin acceso a nada, ni siquiera a lo que lo arreglaria.
'
' RUTINAS:
'   Ribbon_IsVisible(Tag)      - True si el Tag tiene fila y el usuario y la hoja activos casan.
'   Fnc_SuperTip_Data(Tag)     - array 0..2 (Descripion, Nom_Rut, Informe_Rut) o Empty sin fila.
'   Fnc_RibbonUI_Rut(Tag)      - Nom_Rut del Tag ("" sin fila). Lo usa el boton RunRutPrueba.
'   Rut_RibbonUI_Guardar_Informe(Tag, [Texto]) - guarda el informe de la ultima ejecucion en la
'                    columna Informe_Rut del Tag (Texto o, si se omite, APP_Task_Inf); lo muestra
'                    el supertip de Boss. Sin fila, no hace nada.
'   Rules_Reset                - vacia la cache para que la proxima consulta relea la tabla.
'   Fnc_Lista_Contiene(Lista, Valor) - True si Valor es un elemento COMPLETO de Lista.
'
' REQUIERE: Prog__RibbonUI (Lo_RibbonUI), Prog__APP (APP_User_ID, APP_Task_Inf),
'           M_000_Ini_Var_APP (constantes Rib_*).
'==================================================================================================
Option Explicit

Private Dic_Usuario     As Object      '- Tag -> lista de Usuarios permitidos (col Rib_User)
Private Dic_Hojas       As Object      '- Tag -> lista de Hojas permitidas (col Rib_SheetsNames)
Private Dic_Supertip    As Object      '- Tag -> array(Rib_Rut_Descrip, Rib_Rut, Rib_Rut_Informe)
Private Rules_OK        As Boolean     '- True si la ultima carga trajo al menos una fila

'==================================================================================================
Public Sub Rules_Reset()               '- la proxima consulta volvera a leer Lo_RibbonUI
    Set Dic_Usuario = Nothing
    Set Dic_Hojas = Nothing
    Set Dic_Supertip = Nothing
    Rules_OK = False
End Sub
'==================================================================================================
Private Sub Rules_Load()
    Dim Datos       As Variant
    Dim Fila        As Long
    Dim Tag         As String

    Set Dic_Usuario = CreateObject("Scripting.Dictionary")
    Set Dic_Hojas = CreateObject("Scripting.Dictionary")
    Set Dic_Supertip = CreateObject("Scripting.Dictionary")
    Dic_Usuario.CompareMode = vbTextCompare
    Dic_Hojas.CompareMode = vbTextCompare
    Dic_Supertip.CompareMode = vbTextCompare
    Rules_OK = False

    On Error GoTo Fallo                '- sin tabla legible se sale con Rules_OK = False
    Datos = Prog__RibbonUI.ListObjects(1).DataBodyRange.Value
    For Fila = 1 To UBound(Datos, 1)
        Tag = Trim$(CStr(Datos(Fila, Rib_Tag)))
        If Len(Tag) > 0 Then
            If Not Dic_Usuario.Exists(Tag) Then          '- si un Tag estuviera repetido, manda la 1a fila
                Dic_Usuario.Add Tag, Trim$(CStr(Datos(Fila, Rib_User)))
                Dic_Hojas.Add Tag, Trim$(CStr(Datos(Fila, Rib_SheetsNames)))
                Dic_Supertip.Add Tag, Array(CStr(Datos(Fila, Rib_Rut_Descrip)), _
                                            CStr(Datos(Fila, Rib_Rut)), _
                                            CStr(Datos(Fila, Rib_Rut_Informe)))
            End If
        End If
    Next Fila
    Rules_OK = (Dic_Usuario.Count > 0)
Fallo:
End Sub
'==================================================================================================
Public Function Ribbon_IsVisible(Tag As String) As Boolean
    On Error GoTo Fallo
    If Dic_Usuario Is Nothing Then Rules_Load
    If Not Rules_OK Then                                   '- sin tabla legible no se oculta nada:
        Ribbon_IsVisible = True                            '-   ver el porque en la cabecera
        Exit Function
    End If
    If Not Dic_Usuario.Exists(Tag) Then Exit Function      '- un Tag sin fila queda oculto

    Ribbon_IsVisible = Fnc_Lista_Contiene(CStr(Dic_Usuario(Tag)), CStr(Prog__APP.Range("APP_User_ID").Value)) _
                   And Fnc_Lista_Contiene(CStr(Dic_Hojas(Tag)), ActiveSheet.Name)
    Exit Function
Fallo:
    Ribbon_IsVisible = False      '- un fallo no debe tumbar el pintado del ribbon
End Function
'==================================================================================================
Public Function Fnc_SuperTip_Data(Tag As String) As Variant
    On Error GoTo Fallo
    If Dic_Usuario Is Nothing Then Rules_Load
    If Not Rules_OK Then GoTo Fallo
    If Not Dic_Supertip.Exists(Tag) Then GoTo Fallo
    Fnc_SuperTip_Data = Dic_Supertip(Tag)
    Exit Function
Fallo:
    Fnc_SuperTip_Data = Empty      '- Tag sin fila o tabla no legible: el llamador decide el mensaje
End Function
'==================================================================================================
Public Function Fnc_RibbonUI_Rut(Tag As String) As String
    Dim Datos       As Variant
    Datos = Fnc_SuperTip_Data(Tag)
    If Not IsEmpty(Datos) Then Fnc_RibbonUI_Rut = Trim$(CStr(Datos(1)))
End Function
'==================================================================================================
'- Guarda en Informe_Rut el informe de la ultima ejecucion del boton. Busca la fila en vivo (no en
'  la cache) por si la tabla se ha ordenado. Se escribe como CADENA: pasando el Range o el control,
'  Excel da 1004 con textos largos (log de M_110, 2026-10-04).
Public Sub Rut_RibbonUI_Guardar_Informe(ByVal Tag As String, Optional ByVal Texto As Variant)
    Dim Pos         As Variant
    Dim Informe     As String
    On Error GoTo Fallo
    Pos = Application.Match(Tag, Prog__RibbonUI.ListObjects(1).ListColumns(Rib_Tag).DataBodyRange, 0)
    If IsError(Pos) Then Exit Sub                          '- boton sin fila: no se guarda nada
    If IsMissing(Texto) Then Texto = Prog__APP.Range("APP_Task_Inf").Value
    Informe = Left$(CStr(Texto), 32000)                    '- una celda admite 32.767 caracteres
    If Informe Like "[=+@-]*" Then Informe = "'" & Informe  '- que Excel no lo tome por formula
    Prog__RibbonUI.ListObjects(1).DataBodyRange.Cells(Pos, Rib_Rut_Informe).Value = Informe
    Call Rules_Reset                                       '- que el supertip lea el informe nuevo
    Exit Sub
Fallo:
    Debug.Print "Rut_RibbonUI_Guardar_Informe: no se pudo guardar el informe de " & Tag & ": " & Err.Description
End Sub
'==================================================================================================
'- Comprueba si Valor figura en Lista (elementos separados por "," o ";"), sin distinguir
'  mayusculas y comparando elementos COMPLETOS (con InStr, "Rafa" casaria con "RafaG").
'  Lista vacia = sin restriccion, devuelve True. Valor vacio = False.
Public Function Fnc_Lista_Contiene(Lista As String, Valor As String) As Boolean
    Dim Elementos() As String
    Dim Cont_Elem   As Long
    If Trim$(Lista) = "" Then Fnc_Lista_Contiene = True: Exit Function
    If Trim$(Valor) = "" Then Exit Function
    Elementos = Split(Replace(Lista, ";", ","), ",")
    For Cont_Elem = LBound(Elementos) To UBound(Elementos)
        If StrComp(Trim$(Elementos(Cont_Elem)), Trim$(Valor), vbTextCompare) = 0 Then
            Fnc_Lista_Contiene = True
            Exit Function
        End If
    Next Cont_Elem
End Function
