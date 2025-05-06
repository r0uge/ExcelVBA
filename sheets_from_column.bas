'Genero la cantidad de hojas que necesito a partir de una serie de datos que se encuentran en columna
'Al nombre de cada hojas le aplica el valor de cada celda de la columna (no verifica si tiene caracteres validos ni longitud)

Sub CrearHojasDesdeLista()
    Dim ws As Worksheet
    Dim nombreHoja As String
    Dim celda As Range
    Dim existe As Boolean
    
    ' Asume que los nombres están en la hoja activa, columna A, desde la fila 1 hacia abajo
    Set ws = ActiveSheet
    
    For Each celda In ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        nombreHoja = Trim(celda.Value)
        
        If nombreHoja <> "" Then
            ' Verifica si ya existe una hoja con ese nombre
            existe = False
            For Each sh In ThisWorkbook.Sheets
                If sh.Name = nombreHoja Then
                    existe = True
                    Exit For
                End If
            Next sh
            
            ' Si no existe, crea la hoja
            If Not existe Then
                Sheets.Add(After:=Sheets(Sheets.Count)).Name = nombreHoja
            End If
        End If
    Next celda
    
    MsgBox "Hojas creadas desde la lista.", vbInformation
End Sub
