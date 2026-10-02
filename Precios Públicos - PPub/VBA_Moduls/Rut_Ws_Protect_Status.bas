Attribute VB_Name = "Rut_Ws_Protect_Status"
'Last Rev. 2026-10-02 11:54
Option Explicit

'===================================================================================================
' Rut_Ws_Protect_Status
'
' Guarda y restaura el estado de protección de una hoja, conservando la contraseña real y los
' permisos originales (filtrado, ordenación, formato...) en vez de forzar un esquema fijo al
' desproteger y volver a proteger.
'
' Skill "vba-ws-protect-status": es la única fuente de este módulo. Lo usa, entre otros,
' Rut_Wb_Restituir_Datos (skill "vba-restituir-datos", que lo declara como dependencia), y sirve
' a cualquier rutina que necesite escribir en una hoja protegida sin cambiarle los permisos.
'
' REENTRANTE: el estado guardado viaja en una variable del llamador (T_Prot_Estado), no en
' variables a nivel de módulo. Por eso se pueden anidar varios Save/Restore sobre hojas distintas
' sin que uno pise el estado del otro:
'
'   Dim Est As T_Prot_Estado
'   Call Rut_Prot_Save(Ws, Est, "miClave")   ' guarda el estado y DEJA LA HOJA DESPROTEGIDA
'   ' ... escribir en la hoja ...
'   Call Rut_Prot_Restore(Ws, Est)           ' vuelve a dejarla exactamente como estaba
'
' ESTADO POR HOJA: si desproteger y reproteger ocurren en llamadas distintas (p.ej. un botón que
' alterna Protect/Unprotect), el llamador debe conservar un T_Prot_Estado POR HOJA. Uno solo
' compartido aplicaría los permisos de una hoja a otra si se cambia de hoja entre las dos
' pulsaciones. Ejemplo resuelto en el SKILL.md del skill.
'
' La contraseña NO se hardcodea: se pasa como parámetro (el llamador la lee de una celda con
' nombre, del hook del libro, etc.). Cadena vacía = hoja sin contraseña.
'
' LIMITACIONES conocidas (de Excel, no de este módulo):
'   - UserInterfaceOnly no es una propiedad legible: si la hoja estaba protegida con
'     UserInterfaceOnly:=True, al restaurar se perdería ese flag. Por eso Rut_Prot_Restore admite
'     el parámetro opcional UserInterfaceOnly, que el llamador puede poner a True si sabe que la
'     hoja lo usaba.
'   - AllowSelectingLockedCells / AllowSelectingUnlockedCells tampoco son legibles desde
'     Worksheet.Protection; no se guardan ni se restauran.
'   - "Protegida" NO es ProtectContents: una hoja protegida sin la casilla "Proteger hoja y
'     contenido de celdas bloqueadas" (Contents:=False) tiene ProtectContents = False, pero
'     Excel la da por protegida, y un Protect sobre ella NO cambia nada (sin dar error). Por
'     eso se mira con Fnc_Prot_Hoja_Protegida y se guarda/restaura también Contents.
'
' Subs/Functions públicas:
'   Rut_Prot_Save(Ws, Estado, [Password])                   - Guarda el estado y desprotege la hoja
'   Rut_Prot_Restore(Ws, Estado, [UserInterfaceOnly])       - Restaura el estado guardado
'   Fnc_Prot_Estaba_Protegida(Estado)                       - True si la hoja estaba protegida al guardar
'   Fnc_Prot_Hoja_Protegida(Ws)                             - True si la hoja tiene CUALQUIER protección
'===================================================================================================

' Estado de protección de UNA hoja. Público porque viaja como parámetro entre módulos.
Public Type T_Prot_Estado
    Guardado                    As Boolean   ' True tras un Rut_Prot_Save (evita restaurar sin haber guardado)
    Protegida                   As Boolean   ' la hoja estaba protegida
    Contents                    As Boolean   ' celdas bloqueadas protegidas (False es posible con la hoja protegida)
    Password                    As String
    DrawingObjects              As Boolean
    Scenarios                   As Boolean
    AllowFormattingCells        As Boolean
    AllowFormattingColumns      As Boolean
    AllowFormattingRows         As Boolean
    AllowInsertingColumns       As Boolean
    AllowInsertingRows          As Boolean
    AllowInsertingHyperlinks    As Boolean
    AllowDeletingColumns        As Boolean
    AllowDeletingRows           As Boolean
    AllowSorting                As Boolean
    AllowFiltering              As Boolean
    AllowUsingPivotTables       As Boolean
End Type

'---------------------------------------------------------------------------------------------------
' Rut_Prot_Save
'   Guarda en Estado la protección actual de la hoja y la DEJA DESPROTEGIDA, lista para escribir.
'   Si la hoja no estaba protegida no hace nada más que anotarlo (el Restore la dejará igual).
'---------------------------------------------------------------------------------------------------
Public Sub Rut_Prot_Save(ByVal Ws As Worksheet, ByRef Estado As T_Prot_Estado, Optional ByVal Password As String = "")

    Dim Vacio As T_Prot_Estado

    Estado = Vacio                          ' limpiar restos de un uso anterior de la misma variable
    Estado.Guardado = True
    If Ws Is Nothing Then Exit Sub

    On Error Resume Next                    ' alguna propiedad puede no estar disponible según el tipo de hoja
    Estado.Protegida = Fnc_Prot_Hoja_Protegida(Ws)
    If Estado.Protegida Then
        Estado.Password = Password
        Estado.Contents = Ws.ProtectContents
        Estado.DrawingObjects = Ws.ProtectDrawingObjects
        Estado.Scenarios = Ws.ProtectScenarios
        Estado.AllowFormattingCells = Ws.Protection.AllowFormattingCells
        Estado.AllowFormattingColumns = Ws.Protection.AllowFormattingColumns
        Estado.AllowFormattingRows = Ws.Protection.AllowFormattingRows
        Estado.AllowInsertingColumns = Ws.Protection.AllowInsertingColumns
        Estado.AllowInsertingRows = Ws.Protection.AllowInsertingRows
        Estado.AllowInsertingHyperlinks = Ws.Protection.AllowInsertingHyperlinks
        Estado.AllowDeletingColumns = Ws.Protection.AllowDeletingColumns
        Estado.AllowDeletingRows = Ws.Protection.AllowDeletingRows
        Estado.AllowSorting = Ws.Protection.AllowSorting
        Estado.AllowFiltering = Ws.Protection.AllowFiltering
        Estado.AllowUsingPivotTables = Ws.Protection.AllowUsingPivotTables

'       Desproteger con la contraseña si la hay: Unprotect sin argumento sobre una hoja con
'       contraseña abre el cuadro de Excel pidiéndola, en medio de un proceso automático.
        If Len(Password) > 0 Then
            Ws.Unprotect Password:=Password
        Else
            Ws.Unprotect
        End If
    End If
    On Error GoTo 0
End Sub     ' Rut_Prot_Save
'---------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------
' Rut_Prot_Restore
'   Vuelve a proteger la hoja con los mismos permisos y contraseña que tenía. Si no estaba
'   protegida, la deja desprotegida. UserInterfaceOnly:=True protege la hoja de cara al usuario
'   pero permite que el código VBA siga escribiendo en ella (Excel no lo conserva al guardar el
'   libro: hay que volver a aplicarlo en cada sesión).
'---------------------------------------------------------------------------------------------------
Public Sub Rut_Prot_Restore(ByVal Ws As Worksheet, ByRef Estado As T_Prot_Estado, _
                            Optional ByVal UserInterfaceOnly As Boolean = False)

    If Ws Is Nothing Then Exit Sub
    If Not Estado.Guardado Then Exit Sub    ' nunca se guardó: no tocar la protección de la hoja
    If Not Estado.Protegida Then
        Estado.Guardado = False
        Exit Sub                            ' no estaba protegida: se queda como está
    End If

    On Error Resume Next
    Ws.Protect Password:=Estado.Password, _
               DrawingObjects:=Estado.DrawingObjects, _
               Contents:=Estado.Contents, _
               Scenarios:=Estado.Scenarios, _
               UserInterfaceOnly:=UserInterfaceOnly, _
               AllowFormattingCells:=Estado.AllowFormattingCells, _
               AllowFormattingColumns:=Estado.AllowFormattingColumns, _
               AllowFormattingRows:=Estado.AllowFormattingRows, _
               AllowInsertingColumns:=Estado.AllowInsertingColumns, _
               AllowInsertingRows:=Estado.AllowInsertingRows, _
               AllowInsertingHyperlinks:=Estado.AllowInsertingHyperlinks, _
               AllowDeletingColumns:=Estado.AllowDeletingColumns, _
               AllowDeletingRows:=Estado.AllowDeletingRows, _
               AllowSorting:=Estado.AllowSorting, _
               AllowFiltering:=Estado.AllowFiltering, _
               AllowUsingPivotTables:=Estado.AllowUsingPivotTables
    On Error GoTo 0

    Estado.Guardado = False                 ' consumido: un segundo Restore seguido no vuelve a proteger
End Sub     ' Rut_Prot_Restore
'---------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------
' Fnc_Prot_Estaba_Protegida
'   True si la hoja estaba protegida en el momento de guardar el estado. Útil para decidir si hay
'   que avisar al usuario de algo que solo aplica a hojas protegidas.
'---------------------------------------------------------------------------------------------------
Public Function Fnc_Prot_Estaba_Protegida(ByRef Estado As T_Prot_Estado) As Boolean
    Fnc_Prot_Estaba_Protegida = (Estado.Guardado And Estado.Protegida)
End Function     ' Fnc_Prot_Estaba_Protegida
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
' Fnc_Prot_Hoja_Protegida
'   True si la hoja tiene CUALQUIER protección activa. No basta con ProtectContents: protegida sin
'   la casilla de celdas bloqueadas (Contents:=False) da ProtectContents = False, pero Excel la da
'   por protegida (en Revisar sale "Desproteger hoja") y un Protect sobre ella no cambia nada.
'---------------------------------------------------------------------------------------------------
Public Function Fnc_Prot_Hoja_Protegida(ByVal Ws As Worksheet) As Boolean
    If Ws Is Nothing Then Exit Function
    Fnc_Prot_Hoja_Protegida = Ws.ProtectContents Or Ws.ProtectDrawingObjects Or Ws.ProtectScenarios
End Function     ' Fnc_Prot_Hoja_Protegida
'---------------------------------------------------------------------------------------------------
