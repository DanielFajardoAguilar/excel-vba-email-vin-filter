Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub CorreosPorNumeroVINs()

    Dim celdaInicioBase As Range
    Dim celdaInicioResultado As Range
    
    Dim wsBaseDatos As Worksheet
    Dim wsResultado As Worksheet
    
    Dim colCorreo As Long
    Dim colVIN As Long
    
    Dim filaInicioDatos As Long
    Dim filaUltimoDato As Long
    
    Dim numeroVINsBuscados As Long
    
    ' =========================
    ' 1) Seleccionar inicio de la base (EMAIL)
    ' =========================
    On Error Resume Next
    Set celdaInicioBase = Application.InputBox( _
        Prompt:="Selecciona la celda donde INICIA la base (primer CORREO)." & vbCrLf & _
                "El VIN debe estar en la columna inmediata a la derecha.", _
        Title:="Inicio de base", Type:=8)
    On Error GoTo 0
    If celdaInicioBase Is Nothing Then Exit Sub
    
    Set wsBaseDatos = celdaInicioBase.Worksheet
    colCorreo = celdaInicioBase.Column
    colVIN = colCorreo + 1
    filaInicioDatos = celdaInicioBase.Row
    
    ' =========================
    ' 2) Preguntar cuántos VINs se buscan
    ' =========================
    numeroVINsBuscados = CLng(Application.InputBox( _
        Prompt:="¿Cuántos VINs EXACTOS por correo deseas filtrar?", _
        Title:="Filtro por número de VINs", Default:=2, Type:=1))
    
    If numeroVINsBuscados <= 0 Then
        MsgBox "El número de VINs debe ser mayor que cero.", vbExclamation
        Exit Sub
    End If
    
    ' =========================
    ' 3) Seleccionar inicio de salida
    ' =========================
    On Error Resume Next
    Set celdaInicioResultado = Application.InputBox( _
        Prompt:="Selecciona la celda donde comenzará la HOJA RESULTADO.", _
        Title:="Ubicación del resultado", Type:=8)
    On Error GoTo 0
    If celdaInicioResultado Is Nothing Then Exit Sub
    
    Set wsResultado = celdaInicioResultado.Worksheet
    
    ' =========================
    ' 4) Detectar última fila de datos
    ' =========================
    filaUltimoDato = wsBaseDatos.Cells(wsBaseDatos.Rows.Count, colCorreo).End(xlUp).Row
    If filaUltimoDato < filaInicioDatos Then
        MsgBox "No se encontraron datos debajo de la celda seleccionada.", vbExclamation
        Exit Sub
    End If
    
    ' =========================
    ' 5) Agrupar VINs por correo
    ' =========================
    Dim dictVINsPorCorreo As Object
    Set dictVINsPorCorreo = CreateObject("Scripting.Dictionary")
    
    Dim fila As Long
    Dim correoActual As String
    Dim vinActual As String
    
    For fila = filaInicioDatos To filaUltimoDato
        correoActual = Trim$(CStr(wsBaseDatos.Cells(fila, colCorreo).Value))
        vinActual = Trim$(CStr(wsBaseDatos.Cells(fila, colVIN).Value))
        
        If Len(correoActual) > 0 And Len(vinActual) > 0 Then
            If Not dictVINsPorCorreo.Exists(correoActual) Then
                dictVINsPorCorreo.Add correoActual, CreateObject("System.Collections.ArrayList")
            End If
            
            If Not ExisteVIN(dictVINsPorCorreo(correoActual), vinActual) Then
                dictVINsPorCorreo(correoActual).Add vinActual
            End If
        End If
    Next fila
    
    ' =========================
    ' 6) Limpiar área de salida
    ' =========================
    Dim filaInicioResultado As Long
    Dim colInicioResultado As Long
    
    filaInicioResultado = celdaInicioResultado.Row
    colInicioResultado = celdaInicioResultado.Column
    
    wsResultado.Range( _
        wsResultado.Cells(filaInicioResultado, colInicioResultado), _
        wsResultado.Cells(filaInicioResultado + dictVINsPorCorreo.Count + 5, _
                          colInicioResultado + numeroVINsBuscados) _
    ).Clear
    
    ' =========================
    ' 7) Encabezados
    ' =========================
    wsResultado.Cells(filaInicioResultado, colInicioResultado).Value = "CORREO"
    
    Dim i As Long
    For i = 1 To numeroVINsBuscados
        wsResultado.Cells(filaInicioResultado, colInicioResultado + i).Value = "VIN_" & i
    Next i
    
    ' =========================
    ' 8) Escribir resultados
    ' =========================
    Dim claveCorreo As Variant
    Dim filaResultadoActual As Long
    Dim cantidadVINs As Long
    
    filaResultadoActual = filaInicioResultado + 1
    
    For Each claveCorreo In dictVINsPorCorreo.Keys
        cantidadVINs = dictVINsPorCorreo(claveCorreo).Count
        
        If cantidadVINs = numeroVINsBuscados Then
            wsResultado.Cells(filaResultadoActual, colInicioResultado).Value = CStr(claveCorreo)
            
            For i = 1 To numeroVINsBuscados
                wsResultado.Cells(filaResultadoActual, colInicioResultado + i).Value = _
                    dictVINsPorCorreo(claveCorreo)(i - 1)
            Next i
            
            filaResultadoActual = filaResultadoActual + 1
        End If
    Next claveCorreo
    
    ' =========================
    ' 9) Formato final
    ' =========================
    wsResultado.Rows(filaInicioResultado).Font.Bold = True
    wsResultado.Columns.AutoFit
    
    MsgBox "Proceso terminado. Correos con EXACTAMENTE " & _
           numeroVINsBuscados & " VIN(s) exportados.", vbInformation

End Sub

' =========================
' Función auxiliar: validar duplicados de VIN
' =========================
Private Function ExisteVIN(ByVal listaVINs As Object, ByVal vinBuscado As String) As Boolean
    Dim idx As Long
    For idx = 0 To listaVINs.Count - 1
        If CStr(listaVINs(idx)) = vinBuscado Then
            ExisteVIN = True
            Exit Function
        End If
    Next idx
    ExisteVIN = False
End Function




