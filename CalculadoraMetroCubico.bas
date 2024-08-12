Attribute VB_Name = "Módulo1"
Sub abrirForm()
    Formulario.Show
End Sub

Sub cerrarForm()
    Formulario.Hide
End Sub

Function AgregarRegistro(txtProducto As String, txtCantidad As String) As Boolean
    Dim HojaCalculo As Worksheet
    Dim Tabla As ListObject
    Dim nuevafila As ListRow
    
    If Not (Validar(txtProducto, txtCantidad)) Then
        AgregarRegistro = False
        Exit Function
    End If
    
    Set HojaCalculo = ThisWorkbook.Sheets("Hoja1")
    Set Tabla = HojaCalculo.ListObjects("TablaDeCalculo")
    Set nuevafila = Tabla.ListRows.Add
    nuevafila.Range(1, 1).Value = Tabla.ListRows.Count
    nuevafila.Range(1, 2).Value = txtProducto
    nuevafila.Range(1, 3).Value = txtCantidad
    
    AgregarRegistro = True
    
End Function

Sub LimpiarTabla()
    Dim HojaCalculo As Worksheet
    Dim Tabla As ListObject
    
    Set HojaCalculo = ThisWorkbook.Sheets("HojaCalculo")
    Set Tabla = HojaCalculo.ListObjects("TablaDeCalculo")
    
    For i = Tabla.ListRows.Count To 1 Step -1
        Tabla.ListRows(i).Delete
    Next i
End Sub

Function Validar(txtProducto As String, txtCantidad As String) As Boolean
    Validar = True
    If (txtProducto = "") Then
        MsgBox "No hay producto cargada"
        Validar = False
        Exit Function
    End If
    If (txtCantidad = "") Then
        MsgBox "No hay Cantidad cargada"
        Validar = False
        Exit Function
    End If
    If Not IsNumeric(txtCantidad) Then
        MsgBox "El campo cantidad no tiene un numero valido"
        Exit Function
    End If
End Function



