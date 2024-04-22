Attribute VB_Name = "Módulo1"
Sub GenerarReceta()
    Dim mueble As String
    Dim cantidad As Double
    Dim CantidadDeFilas As Long
    Dim HojaFormulas As Worksheet
    Dim HojaResultados As Worksheet
    Dim HojaFormulario As Worksheet
    Dim i As Long
    Dim FilaInicial As Long
    Set HojaFormulas = ThisWorkbook.Sheets("Recetas")
    Set HojaResultados = ThisWorkbook.Sheets("Resultados")
    Set HojaFormulario = ThisWorkbook.Sheets("Formulario")
    
    
    'Obtener el producto seleccionado
    mueble = HojaFormulario.Cells(2, "B").Value
    
    'Obtener la cantidad deseada
    cantidad = HojaFormulario.Cells(2, "C").Value
    
    'Contar filas de la hoja de formulas
    CantidadDeFilas = HojaFormulas.Cells(HojaFormulas.Rows.Count, "A").End(xlUp).Row
    FilaInicial = 2
    
    'Bucle de busqueda + calculo de formulas
    For i = 2 To CantidadDeFilas
        If HojaFormulas.Cells(i, "A").Value = mueble And HojaFormulas.Cells(i, "C") = 1 Then
            HojaResultados.Cells(FilaInicial, "A").Value = HojaFormulas.Cells(i, "D").Value
            HojaResultados.Cells(FilaInicial, "B").Value = HojaFormulas.Cells(i, "E").Value * cantidad
            FilaInicial = FilaInicial + 1
        End If
        Next i
        
        
            
        
    
    
      
    'Mostrar mensaje de finalización
    MsgBox "Receta generada correctamente.", vbInformation
End Sub
