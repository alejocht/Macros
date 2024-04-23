Attribute VB_Name = "Módulo1"
Sub GenerarReceta()
    Dim mueble As String
    Dim cantidad As Double
    Dim CantidadDeFilas As Long
    Dim HojaFormulas As Worksheet
    Dim HojaResultados As Worksheet
    Dim HojaFormulario As Worksheet
    Dim i As Long
    Dim j As Long
    Dim FilaInicial As Long
    Dim CantidadDeEntradas As Long
    Dim VariableDeControlId As Long
    Dim FilaInicialErrores As Long
    Set HojaErrores = ThisWorkbook.Sheets("Errores")
    Set HojaFormulas = ThisWorkbook.Sheets("Recetas")
    Set HojaResultados = ThisWorkbook.Sheets("Resultados")
    Set HojaFormulario = ThisWorkbook.Sheets("Formulario")
    'Limpiar residuos de calculos anteriores
    HojaResultados.Range("A2:E1000").Clear
    HojaErrores.Range("A2:E1000").Clear
    'Obtener cantidad de entradas
    CantidadDeEntradas = HojaFormulario.Cells(HojaFormulario.Rows.Count, "B").End(xlUp).Row
    'CantidadDeEntradas = CantidadDeEntradas - 1
    'Contar filas de la hoja de formulas
    CantidadDeFilas = HojaFormulas.Cells(HojaFormulas.Rows.Count, "A").End(xlUp).Row
    'Fila inicial para imprimir resultados
    FilaInicial = 2
    FilaInicialErrores = 2
    VariableDeControlId = 1
    
    For j = 2 To CantidadDeEntradas
        mueble = HojaFormulario.Cells(j, "E").Value
        cantidad = HojaFormulario.Cells(j, "D").Value
        
        'Bucle de busqueda + calculo de formulas
        For i = 2 To CantidadDeFilas
            If HojaFormulas.Cells(i, "A").Value = mueble And HojaFormulas.Cells(i, "C") = 1 Then
                'Escritura en Resultados
                
                'IdFila
                HojaResultados.Cells(FilaInicial, "A").Value = VariableDeControlId
                'Mueble
                 HojaResultados.Cells(FilaInicial, "B").Value = HojaFormulario.Cells(j, "E").Value
                 'Producto
                 HojaResultados.Cells(FilaInicial, "C").Value = HojaFormulas.Cells(i, "D").Value
                 'Cantidad multiplicada
                 HojaResultados.Cells(FilaInicial, "D").Value = HojaFormulas.Cells(i, "E").Value * cantidad
                 
                 FilaInicial = FilaInicial + 1
            Else
                'agregar el producto a la hoja errores
                
                'IdFila
                HojaErrores.Cells(FilaInicialErrores, "A").Value = VariableDeControlId
                'Mueble
                HojaErrores.Cells(FilaInicialErrores, "B").Value = HojaFormulario.Cells(j, "E").Value
                'Error generico
                HojaErrores.Cells(FilaInicialErrores, "C").Value = "El mueble no pudo ser reconocido. Use un mueble disponible en HojaRecetas"
                
                FilaInicialErrores = FilaInicialErrores + 1
            
         End If
        Next i
        VariableDeControlId = VariableDeControlId + 1
    Next j
        
        
            
        
    
    
      
    'Mostrar mensaje de finalización
    MsgBox "Receta generada correctamente.", vbInformation
End Sub
