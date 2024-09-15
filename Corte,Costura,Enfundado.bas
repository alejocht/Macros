Attribute VB_Name = "Módulo1"
Sub PasarDeCorteACostura()
    Set HojaCorte = Sheets("CORTE")
    Set HojaCostura = Sheets("COSTURA")
    Set hojaRespaldo = Sheets("RESPALDO")
    
    Dim i As Integer
    Dim fechahora As Date
    fechahora = Now
    
    ultimaFilaCorte = HojaCorte.Cells(HojaCorte.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaCostura = HojaCostura.Cells(HojaCostura.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaRespaldo = hojaRespaldo.Cells(hojaRespaldo.Rows.Count, "D").End(xlUp).Row + 1
    
    For i = 2 To ultimaFilaCorte Step 1
        If HojaCorte.Cells(i, "L").Value Then
            columna = 1
            For columna = 1 To 10
                HojaCostura.Cells(ultimaFilaCostura, columna).Value = HojaCorte.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, columna).Value = HojaCorte.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, 11).Value = fechahora
                hojaRespaldo.Cells(ultimaFilaRespaldo, 12).Value = "Se movio de corte a costura"
            Next columna
            ultimaFilaCostura = ultimaFilaCostura + 1
            ultimaFilaRespaldo = ultimaFilaRespaldo + 1
            HojaCorte.Rows(i).Delete
            ultimaFilaCorte = ultimaFilaCorte - 1
            i = i - 1
        End If
    Next i
        
End Sub

Sub PasarDeCosturaACorte()
    Set HojaCorte = Sheets("CORTE")
    Set HojaCostura = Sheets("COSTURA")
    Set hojaRespaldo = Sheets("RESPALDO")
    
    Dim i As Integer
    Dim fechahora As Date
    fechahora = Now
    
    ultimaFilaCorte = HojaCorte.Cells(HojaCorte.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaCostura = HojaCostura.Cells(HojaCostura.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaRespaldo = hojaRespaldo.Cells(hojaRespaldo.Rows.Count, "D").End(xlUp).Row + 1
    
    For i = 2 To ultimaFilaCostura Step 1
        If HojaCostura.Cells(i, "L").Value Then
            columna = 1
            For columna = 1 To 10
                HojaCorte.Cells(ultimaFilaCorte, columna).Value = HojaCostura.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, columna).Value = HojaCostura.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, 11).Value = fechahora
                hojaRespaldo.Cells(ultimaFilaRespaldo, 12).Value = "Se movio de costura a corte"
            Next columna
            ultimaFilaCostura = ultimaFilaCostura - 1
            ultimaFilaRespaldo = ultimaFilaRespaldo + 1
            ultimaFilaCorte = ultimaFilaCorte + 1
            HojaCostura.Rows(i).Delete
            i = i - 1
        End If
    Next i
        
End Sub

Sub PasarDeCosturaAEnfundado()
    Set HojaCostura = Sheets("COSTURA")
    Set HojaEnfundado = Sheets("ENFUNDADO")
    Set hojaRespaldo = Sheets("RESPALDO")
    
    Dim i As Integer
    Dim fechahora As Date
    fechahora = Now
    
    ultimaFilaCostura = HojaCostura.Cells(HojaCostura.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaEnfundado = HojaEnfundado.Cells(HojaEnfundado.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaRespaldo = hojaRespaldo.Cells(hojaRespaldo.Rows.Count, "D").End(xlUp).Row + 1
    
    For i = 2 To ultimaFilaCostura
        If HojaCostura.Cells(i, "L").Value Then
            columna = 1
            For columna = 1 To 10
                HojaEnfundado.Cells(ultimaFilaEnfundado, columna).Value = HojaCostura.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, columna).Value = HojaCostura.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, 11).Value = fechahora
                hojaRespaldo.Cells(ultimaFilaRespaldo, 12).Value = "Se movio de costura a enfundado"
            Next columna
            ultimaFilaCostura = ultimaFilaCostura - 1
            ultimaFilaRespaldo = ultimaFilaRespaldo + 1
            ultimaFilaEnfundado = ultimaFilaEnfundado + 1
            HojaCostura.Rows(i).Delete
            i = i - 1
        End If
    Next i
        
End Sub

Sub PasarDeEnfundadoAListo()
    Set HojaEnfundado = Sheets("ENFUNDADO")
    Set HojaListo = Sheets("LISTOS")
    Set hojaRespaldo = Sheets("RESPALDO")
    
    Dim i As Integer
    Dim fechahora As Date
    fechahora = Now
    
    ultimaFilaEnfundado = HojaEnfundado.Cells(HojaEnfundado.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaListo = HojaListo.Cells(HojaListo.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaRespaldo = hojaRespaldo.Cells(hojaRespaldo.Rows.Count, "D").End(xlUp).Row + 1
    
    For i = 2 To ultimaFilaEnfundado
        If HojaEnfundado.Cells(i, "L").Value Then
            columna = 1
            For columna = 1 To 10
                HojaListo.Cells(ultimaFilaListo, columna).Value = HojaEnfundado.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, columna).Value = HojaEnfundado.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, 11).Value = fechahora
                hojaRespaldo.Cells(ultimaFilaRespaldo, 12).Value = "Se movio de enfundado a listo"
            Next columna
            ultimaFilaEnfundado = ultimaFilaEnfundado - 1
            ultimaFilaRespaldo = ultimaFilaRespaldo + 1
            ultimaFilaListo = ultimaFilaListo + 1
            HojaEnfundado.Rows(i).Delete
            i = i - 1
        End If
    Next i
        
End Sub

Sub PasarDeEnfundadoACostura()
    Set HojaEnfundado = Sheets("ENFUNDADO")
    Set HojaCostura = Sheets("COSTURA")
    Set hojaRespaldo = Sheets("RESPALDO")
    
    Dim i As Integer
    Dim fechahora As Date
    fechahora = Now
    
    ultimaFilaEnfundado = HojaEnfundado.Cells(HojaEnfundado.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaCostura = HojaCostura.Cells(HojaCostura.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaRespaldo = hojaRespaldo.Cells(hojaRespaldo.Rows.Count, "D").End(xlUp).Row + 1
    
    For i = 2 To ultimaFilaEnfundado
        If HojaEnfundado.Cells(i, "L").Value Then
            columna = 1
            For columna = 1 To 10
                HojaCostura.Cells(ultimaFilaCostura, columna).Value = HojaEnfundado.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, columna).Value = HojaEnfundado.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, 11).Value = fechahora
                hojaRespaldo.Cells(ultimaFilaRespaldo, 12).Value = "Se movio de enfundado a costura"
            Next columna
            ultimaFilaCostura = ultimaFilaCostura + 1
            ultimaFilaRespaldo = ultimaFilaRespaldo + 1
            ultimaFilaEnfundado = ultimaFilaEnfundado - 1
            HojaEnfundado.Rows(i).Delete
            i = i - 1
        End If
    Next i
        
End Sub

Sub PasarDeListoAEnfundado()
    Set HojaListo = Sheets("LISTOS")
    Set HojaEnfundado = Sheets("ENFUNDADO")
    Set hojaRespaldo = Sheets("RESPALDO")
    
    Dim i As Integer
    Dim fechahora As Date
    fechahora = Now
    
    ultimaFilaListo = HojaListo.Cells(HojaListo.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaEnfundado = HojaEnfundado.Cells(HojaEnfundado.Rows.Count, "D").End(xlUp).Row + 1
    ultimaFilaRespaldo = hojaRespaldo.Cells(hojaRespaldo.Rows.Count, "D").End(xlUp).Row + 1
    
    
    For i = 2 To ultimaFilaListo
        If HojaListo.Cells(i, "K").Value Then
            columna = 1
            For columna = 1 To 10
                HojaEnfundado.Cells(ultimaFilaEnfundado, columna).Value = HojaListo.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, columna).Value = HojaListo.Cells(i, columna).Value
                hojaRespaldo.Cells(ultimaFilaRespaldo, 11).Value = fechahora
                hojaRespaldo.Cells(ultimaFilaRespaldo, 12).Value = "Se movio de listo a enfundado"
            Next columna
            ultimaFilaEnfundado = ultimaFilaEnfundado + 1
            ultimaFilaRespaldo = ultimaFilaRespaldo + 1
            ultimaFilaListo = ultimaFilaListo - 1
            HojaListo.Rows(i).Delete
            i = i - 1
        End If
    Next i
        
End Sub
