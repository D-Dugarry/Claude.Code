Attribute VB_Name = "M_2_Inicializar_Tabla"
' ==============================================================================
' Módulo     : M_2_Inicializar_Tabla
' Proyecto   : Mailing Indicadores UA
' Autor      : Dugarry
' Descripción: Operaciones de inicialización y limpieza sobre Tb_Datos.
'
' Subrutinas públicas:
'   · Rut_Inicializar_Tabla() — Borra los datos de Tb_Datos con confirmación
' ==============================================================================
Option Explicit

' ==============================================================================
' Rut_Inicializar_Tabla
' ------------------------------------------------------------------------------
' Elimina todas las filas de datos de Tb_Datos previa confirmación del usuario.
' Si la tabla está vacía, avisa y sale sin hacer nada.
' ==============================================================================
Sub Rut_Inicializar_Tabla()
    Dim Pregunta    As VbMsgBoxResult
    Dim Tb_Tabla    As ListObject
    Dim Tb_Rango    As Range

    Set Tb_Tabla = ThisWorkbook.Sheets(DatosCorreo.Name).ListObjects(1)
    Set Tb_Rango = Tb_Tabla.DataBodyRange

    ' Comprobar tabla vacía ANTES de desactivar ScreenUpdating
    If Tb_Rango Is Nothing Then
        MsgBox "No hay Datos en la Tabla para Borrar", , "Dugarry's Botones"
        Exit Sub
    End If

    Pregunta = MsgBox("¿Seguro que deseas borrar todas estas filas?", vbYesNo + vbQuestion, "Dugarry's Botones")
    If Pregunta = vbYes Then
        Application.ScreenUpdating = False
        Tb_Rango.Delete                     ' No hace falta .Select previo
        Tb_Tabla.Range.Cells(1).Select
        Application.ScreenUpdating = True
    End If

End Sub     ' Rut_Inicializar_Tabla
' ==============================================================================
