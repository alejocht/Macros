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
    Set HojaFormulas = ThisWorkbook.Sheets("Recetas")
    Set HojaResultados = ThisWorkbook.Sheets("Resultados")
    Set HojaFormulario = ThisWorkbook.Sheets("Formulario")
    'Limpiar residuos de calculos anteriores
    HojaResultados.Range("A2:E1000").Clear
    'Obtener cantidad de entradas
    CantidadDeEntradas = HojaFormulario.Cells(HojaFormulario.Rows.Count, "B").End(xlUp).Row
    'CantidadDeEntradas = CantidadDeEntradas - 1
    'Contar filas de la hoja de formulas
    CantidadDeFilas = HojaFormulas.Cells(HojaFormulas.Rows.Count, "A").End(xlUp).Row
    'Fila inicial para imprimir resultados
    FilaInicial = 2
    
    For j = 2 To CantidadDeEntradas
        mueble = HojaFormulario.Cells(j, "E").Value
        cantidad = HojaFormulario.Cells(j, "D").Value
        
        'Bucle de busqueda + calculo de formulas
        For i = 2 To CantidadDeFilas
         If HojaFormulas.Cells(i, "A").Value = mueble And HojaFormulas.Cells(i, "C") = 1 Then
             HojaResultados.Cells(FilaInicial, "A").Value = HojaFormulas.Cells(i, "D").Value
              HojaResultados.Cells(FilaInicial, "B").Value = HojaFormulas.Cells(i, "E").Value * cantidad
               FilaInicial = FilaInicial + 1
         End If
        Next i
    Next j
        
        
            
        
    
    
      
    'Mostrar mensaje de finalización
    MsgBox "Receta generada correctamente.", vbInformation
End Sub
