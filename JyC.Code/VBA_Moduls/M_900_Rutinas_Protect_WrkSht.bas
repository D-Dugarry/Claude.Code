Attribute VB_Name = "M_900_Rutinas_Protect_WrkSht"
Option Explicit

' Variables globales para guardar el estado de protección
Dim wasProtected As Boolean
Dim protectionPassword As String

' Permisos de protección específicos
Dim allowFormattingCells As Boolean
Dim allowFormattingColumns As Boolean
Dim allowFormattingRows As Boolean
Dim allowInsertingColumns As Boolean
Dim allowInsertingRows As Boolean
Dim allowInsertingHyperlinks As Boolean
Dim allowDeletingColumns As Boolean
Dim allowDeletingRows As Boolean
Dim allowSorting As Boolean
Dim allowFiltering As Boolean
Dim allowUsingPivotTables As Boolean
Dim allowSelectingLockedCells As Boolean
Dim allowSelectingUnlockedCells As Boolean
Dim allowDrawingObjects As Boolean

Sub Rut_Protect_Status_Display_ByHand()
    
    H_INICI.Protect , allowFiltering:=True, allowSorting:=True, DrawingObjects:=True, UserInterfaceOnly:=True       '=== IMPORTANTE, Mantiene protegida la hoja pero permite modificar con VBA  ================
'    H_INICI.Unprotect
'        Debug.Print "allowSorting", H_INICI.Protection.allowSorting
'        Debug.Print "allowFiltering", H_INICI.Protection.allowFiltering
    Call Rut_Protect_Status_Display(ActiveSheet)
End Sub

Sub Rut_Protect_Status_Display(WrkSht As Worksheet)
    On Error Resume Next  ' Ignora errores si alguna propiedad no está disponible
    ' Guarda el estado de protección de la hoja
    wasProtected = WrkSht.ProtectContents
    protectionPassword = "" ' Reiniciar contraseña por seguridad

    ' Si la hoja está protegida, guardamos la contraseña y detalles de protección
    If wasProtected Then
        protectionPassword = "Comptabilitat2020" ' Cambia esto por la contraseña actual
        ' Guarda los permisos de protección
        Debug.Print "allowFormattingCells", WrkSht.Protection.allowFormattingCells
        Debug.Print "allowFormattingColumns", WrkSht.Protection.allowFormattingColumns
        Debug.Print "allowFormattingRows", WrkSht.Protection.allowFormattingRows
        Debug.Print "allowInsertingColumns", WrkSht.Protection.allowInsertingColumns
        Debug.Print "allowInsertingRows", WrkSht.Protection.allowInsertingRows
        Debug.Print "allowInsertingHyperlinks", WrkSht.Protection.allowInsertingHyperlinks
        Debug.Print "allowDeletingColumns", WrkSht.Protection.allowDeletingColumns
        Debug.Print "allowDeletingRows", WrkSht.Protection.allowDeletingRows
        Debug.Print "allowSorting", WrkSht.Protection.allowSorting
        Debug.Print "allowFiltering", WrkSht.Protection.allowFiltering
        Debug.Print "allowUsingPivotTables", WrkSht.Protection.allowUsingPivotTables
        Debug.Print "allowDrawingObjects", WrkSht.Protection.DrawingObjects
'        allowSelectingLockedCells = WrkSht.Protection.allowSelectingLockedCells
'        allowSelectingUnlockedCells = WrkSht.Protection.allowSelectingUnlockedCells
    End If
        Debug.Print "allowFormattingCells", WrkSht.Protection.allowFormattingCells
        Debug.Print "allowFormattingColumns", WrkSht.Protection.allowFormattingColumns
        Debug.Print "allowFormattingRows", WrkSht.Protection.allowFormattingRows
        Debug.Print "allowInsertingColumns", WrkSht.Protection.allowInsertingColumns
        Debug.Print "allowInsertingRows", WrkSht.Protection.allowInsertingRows
        Debug.Print "allowInsertingHyperlinks", WrkSht.Protection.allowInsertingHyperlinks
        Debug.Print "allowDeletingColumns", WrkSht.Protection.allowDeletingColumns
        Debug.Print "allowDeletingRows", WrkSht.Protection.allowDeletingRows
        Debug.Print "allowSorting", WrkSht.Protection.allowSorting
        Debug.Print "allowFiltering", WrkSht.Protection.allowFiltering
        Debug.Print "allowUsingPivotTables", WrkSht.Protection.allowUsingPivotTables
        Debug.Print "allowDrawingObjects", WrkSht.Protection.DrawingObjects
    On Error GoTo 0  ' Reestablece manejo de errores
End Sub

Sub Rut_Protect_Status_Save(WrkSht As Worksheet)
    On Error Resume Next  ' Ignora errores si alguna propiedad no está disponible
    ' Guarda el estado de protección de la hoja
    wasProtected = WrkSht.ProtectContents
    protectionPassword = "" ' Reiniciar contraseña por seguridad

    ' Si la hoja está protegida, guardamos la contraseña y detalles de protección
    If wasProtected Then
        protectionPassword = "Comptabilitat2020" ' Cambia esto por la contraseña actual
        ' Guarda los permisos de protección
        allowFormattingCells = WrkSht.Protection.allowFormattingCells
        allowFormattingColumns = WrkSht.Protection.allowFormattingColumns
        allowFormattingRows = WrkSht.Protection.allowFormattingRows
        allowInsertingColumns = WrkSht.Protection.allowInsertingColumns
        allowInsertingRows = WrkSht.Protection.allowInsertingRows
        allowInsertingHyperlinks = WrkSht.Protection.allowInsertingHyperlinks
        allowDeletingColumns = WrkSht.Protection.allowDeletingColumns
        allowDeletingRows = WrkSht.Protection.allowDeletingRows
        allowSorting = WrkSht.Protection.allowSorting
        allowFiltering = WrkSht.Protection.allowFiltering
        allowUsingPivotTables = WrkSht.Protection.allowUsingPivotTables
'        allowSelectingLockedCells = WrkSht.Protection.allowSelectingLockedCells
'        allowSelectingUnlockedCells = WrkSht.Protection.allowSelectingUnlockedCells
    End If
    On Error GoTo 0  ' Reestablece manejo de errores
End Sub

Sub Rut_Protect_Status_Restore(WrkSht As Worksheet)
    On Error Resume Next  ' Ignora errores si alguna propiedad no está disponible
    ' Restablece el estado de protección con los detalles originales
    If wasProtected Then
        ' Si estaba protegida, aplicar nuevamente la protección con los mismos permisos
        WrkSht.Protect Password:=protectionPassword, _
                        allowFormattingCells:=allowFormattingCells, _
                        allowFormattingColumns:=allowFormattingColumns, _
                        allowFormattingRows:=allowFormattingRows, _
                        allowInsertingColumns:=allowInsertingColumns, _
                        allowInsertingRows:=allowInsertingRows, _
                        allowInsertingHyperlinks:=allowInsertingHyperlinks, _
                        allowDeletingColumns:=allowDeletingColumns, _
                        allowDeletingRows:=allowDeletingRows, _
                        allowSorting:=allowSorting, _
                        allowFiltering:=allowFiltering, _
                        allowUsingPivotTables:=allowUsingPivotTables ', _
                        allowSelectingLockedCells:=allowSelectingLockedCells, _
                        allowSelectingUnlockedCells:=allowSelectingUnlockedCells
    Else
        ' Si no estaba protegida, simplemente desproteger
        WrkSht.Unprotect
    End If
    On Error GoTo 0  ' Reestablece manejo de errores
End Sub

