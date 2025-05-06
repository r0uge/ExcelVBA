'Escenario:
'Dado una tabla en excel en la hoja "Data" se desea filtrar segun cada valor de una columna en la hoja "Sheet2"
'Luego del filtrado, copiar el resultado y pegarlo en cada hoja que su nombre coincide con cada valor de la columna de la hoja "Sheet2" (mismo valor usado para fitrar)
'Repetir la operacion hasta terminar con todos los valores de la columna de la hoja "Sheet2"

Sub FiltrarYCopiarDatos()

    Dim wsOrigen As Worksheet
    Dim wsValores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaOrigen As Long
    Dim ultimaFilaValores As Long
    Dim celdaValor As Range
    Dim rngDatos As Range
    Dim valorFiltro As String
    Dim columnaFiltro As String
    
    ' CONFIGURACIÓN
    Set wsOrigen = ThisWorkbook.Sheets("Data")
    Set wsValores = ThisWorkbook.Sheets("Sheet2")
    columnaFiltro = "C"  ' Columna a filtrar en Hoja1 (ajústalo si necesario)

    ' Encontrar última fila en Hoja1 y Hoja2
    ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, columnaFiltro).End(xlUp).Row
    ultimaFilaValores = wsValores.Cells(wsValores.Rows.Count, "A").End(xlUp).Row
    
    ' Rango total de datos (suponiendo que los datos tienen encabezados en la fila 1)
    Set rngDatos = wsOrigen.Range("A3").CurrentRegion

    Application.ScreenUpdating = False
    
    For Each celdaValor In wsValores.Range("A1:A" & ultimaFilaValores) ' Suponemos que hay encabezado en A1
        valorFiltro = Trim(celdaValor.Value)
        
        If Len(valorFiltro) > 0 Then
            ' Aplicar filtro
            rngDatos.AutoFilter Field:=3, Criteria1:=valorFiltro
            
            On Error Resume Next
            Set wsDestino = ThisWorkbook.Sheets(valorFiltro)
            On Error GoTo 0
            
            If Not wsDestino Is Nothing Then
                ' Limpiar destino
                wsDestino.Cells.ClearContents
                
                ' Copiar encabezado + datos visibles
                rngDatos.SpecialCells(xlCellTypeVisible).Copy
                wsDestino.Range("A1").PasteSpecial Paste:=xlPasteValues
            End If
            
            Set wsDestino = Nothing
        End If
    Next celdaValor

    ' Quitar filtro
    wsOrigen.AutoFilterMode = False

    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    MsgBox "Proceso completado."

End Sub
