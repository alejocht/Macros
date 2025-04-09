Attribute VB_Name = "Módulo1"
Dim formularioWs As Worksheet
Dim corteWs As Worksheet
Dim soldadoWs As Worksheet
Dim cargaCalidadWs As Worksheet
Dim pintureriaWs As Worksheet
Dim finalWs As Worksheet

Dim fechaPedido As Range
Dim fechaProducir As Range
Dim numOrdenLote As Range
Dim modelo As Range
Dim cantidad As Range
Dim turno As Range
Dim observacion As Range
Dim estado As Range

Sub inicializar()
    Set formularioWs = ActiveWorkbook.Worksheets("PCP")
    Set corteWs = ActiveWorkbook.Worksheets("SEC. CORTE")
    Set soldadoWs = ActiveWorkbook.Worksheets("SEC. SOLDADO")
    Set cargaCalidadWs = ActiveWorkbook.Worksheets("CARGA Y CALIDAD")
    Set pintureriaWs = ActiveWorkbook.Worksheets("PINTURERIA")
    Set finalWs = ActiveWorkbook.Worksheets("FINAL")
    
    Set fechaPedido = formularioWs.Range("B3")
    Set fechaProducir = formularioWs.Range("B5")
    Set numOrdenLote = formularioWs.Range("B7")
    Set modelo = formularioWs.Range("B9")
    Set cantidad = formularioWs.Range("B11")
    Set turno = formularioWs.Range("B13")
    Set observacion = formularioWs.Range("B15")
    Set estado = formularioWs.Range("B17")
End Sub

Sub vaciarFormulario()
    fechaPedido.ClearContents
    fechaProducir.ClearContents
    numOrdenLote.ClearContents
    modelo.ClearContents
    cantidad.ClearContents
    turno.ClearContents
    observacion.ClearContents
    estado.ClearContents
End Sub

Sub agregarRegistroDesdeFormulario(tabla As ListObject, fechaPedido As Range, fechaProducir As Range, numOrdenLote As Range, modelo As Range, cantidad As Range, turno As Range, observacion As Range, estado As Range)
    'agregar fila
    Dim nuevaFila As ListRow
    Set nuevaFila = tabla.ListRows.Add
    'pasar datos
    nuevaFila.Range(1, 1).Value = fechaPedido.Value
    nuevaFila.Range(1, 2).Value = fechaProducir.Value
    nuevaFila.Range(1, 3).Value = numOrdenLote.Value
    nuevaFila.Range(1, 4).Value = modelo.Value
    nuevaFila.Range(1, 5).Value = cantidad.Value
    nuevaFila.Range(1, 6).Value = turno.Value
    nuevaFila.Range(1, 7).Value = observacion.Value
    nuevaFila.Range(1, 8).Value = estado.Value
End Sub

'De Hoja Formulario a Hoja X

Sub FormularioACorte()
    'inicializar variables
    inicializar
    'set tabla corte
    Dim corteTbl As ListObject
    corteWs.Select
    Set corteTbl = corteWs.ListObjects("tblSec.Corte")
    'agregar nuevo registro
    Call agregarRegistroDesdeFormulario(corteTbl, fechaPedido, fechaProducir, numOrdenLote, modelo, cantidad, turno, observacion, estado)
    'vaciar formulario
    vaciarFormulario
End Sub

Sub FormularioASoldado()
    'inicializar variables
    inicializar
    'set tabla corte
    Dim soldadoTbl As ListObject
    soldadoWs.Select
    Set soldadoTbl = soldadoWs.ListObjects("tblSec.Soldado")
    'agregar nuevo registro
    Call agregarRegistroDesdeFormulario(soldadoTbl, fechaPedido, fechaProducir, numOrdenLote, modelo, cantidad, turno, observacion, estado)
    'vaciar formulario
    vaciarFormulario
End Sub

Sub FormularioACargaCalidad()
    'inicializar variables
    inicializar
    'set tabla corte
    Dim cargaCalidadTbl As ListObject
    cargaCalidadWs.Select
    Set cargaCalidadTbl = cargaCalidadWs.ListObjects("tblCargaCalidad")
    'agregar nuevo registro
    Call agregarRegistroDesdeFormulario(cargaCalidadTbl, fechaPedido, fechaProducir, numOrdenLote, modelo, cantidad, turno, observacion, estado)
    'vaciar formulario
    vaciarFormulario
End Sub

Sub FormularioAFinal()
    'inicializar variables
    inicializar
    'set tabla corte
    Dim finalTbl As ListObject
    finalWs.Select
    Set finalTbl = finalWs.ListObjects("tblFinal")
    'agregar nuevo registro
    Call agregarRegistroDesdeFormulario(finalTbl, fechaPedido, fechaProducir, numOrdenLote, modelo, cantidad, turno, observacion, estado)
    'vaciar formulario
    vaciarFormulario
End Sub

'De Hoja A a Hoja B

Function listaDeFilasSeleccionadas(tabla As ListObject, columna As Integer) As Collection
    Dim i As Integer
    Dim filasSeleccionadas As New Collection
    
    i = 1
    For Each fila In tabla.ListRows
        If fila.Range.Cells(1, columna).Value Then
            filasSeleccionadas.Add i
        End If
        i = i + 1
    Next fila
    
    Set listaDeFilasSeleccionadas = filasSeleccionadas
End Function

Sub eliminarRegistrosDeTabla(tablaOrigen As ListObject, filasABorrar As Collection)
    Dim i As Long
    ' Recorremos de atrás hacia adelante para evitar conflictos al eliminar
    For i = filasABorrar.Count To 1 Step -1
        Dim index As Long
        index = filasABorrar(i)
        tablaOrigen.ListRows(index).Delete
    Next i
End Sub


Sub agregarRegistrosDeTabla(tablaOrigen As ListObject, tablaDestino As ListObject, filasAgregar As Collection)
    Dim nuevaFila As ListRow
    
    For Each fila In filasAgregar
        Set nuevaFila = tablaDestino.ListRows.Add
        nuevaFila.Range(1, 1).Value = tablaOrigen.DataBodyRange(fila, 1).Value
        nuevaFila.Range(1, 2).Value = tablaOrigen.DataBodyRange(fila, 2).Value
        nuevaFila.Range(1, 3).Value = tablaOrigen.DataBodyRange(fila, 3).Value
        nuevaFila.Range(1, 4).Value = tablaOrigen.DataBodyRange(fila, 4).Value
        nuevaFila.Range(1, 5).Value = tablaOrigen.DataBodyRange(fila, 5).Value
        nuevaFila.Range(1, 6).Value = tablaOrigen.DataBodyRange(fila, 6).Value
        nuevaFila.Range(1, 7).Value = tablaOrigen.DataBodyRange(fila, 7).Value
        nuevaFila.Range(1, 8).Value = tablaOrigen.DataBodyRange(fila, 8).Value
    Next fila
End Sub

Sub CorteASoldado()
    inicializar
    Dim corteTbl As ListObject
    Dim soldadoTbl As ListObject
    Set corteTbl = corteWs.ListObjects("tblSec.Corte")
    Set soldadoTbl = soldadoWs.ListObjects("tblSec.Soldado")
    'chequear los tildados en seleccion
    Dim filasSeleccionadas As Collection
    Set filasSeleccionadas = listaDeFilasSeleccionadas(corteTbl, 9)
    'agregarlos en la otra tabla
    agregarRegistrosDeTabla corteTbl, soldadoTbl, filasSeleccionadas
    'borrar fila de origen
    eliminarRegistrosDeTabla corteTbl, filasSeleccionadas
    
End Sub

Sub SoldadoACargaCalidad()
    inicializar
    Dim cargaCalidadTbl As ListObject
    Dim soldadoTbl As ListObject
    Set cargaCalidadTbl = cargaCalidadWs.ListObjects("tblCargaCalidad")
    Set soldadoTbl = soldadoWs.ListObjects("tblSec.Soldado")
    'chequear los tildados en seleccion
    Dim filasSeleccionadas As Collection
    Set filasSeleccionadas = listaDeFilasSeleccionadas(soldadoTbl, 9)
    'agregarlos en la otra tabla
    agregarRegistrosDeTabla soldadoTbl, cargaCalidadTbl, filasSeleccionadas
    'borrar fila de origen
    eliminarRegistrosDeTabla soldadoTbl, filasSeleccionadas
End Sub

Sub CargaCalidadAFin()
    inicializar
    Dim cargaCalidadTbl As ListObject
    Dim finTbl As ListObject
    Set cargaCalidadTbl = cargaCalidadWs.ListObjects("tblCargaCalidad")
    Set finTbl = finalWs.ListObjects("tblFinal")
    'chequear los tildados en seleccion
    Dim filasSeleccionadas As Collection
    Set filasSeleccionadas = listaDeFilasSeleccionadas(cargaCalidadTbl, 9)
    'agregarlos en la otra tabla
    agregarRegistrosDeTabla cargaCalidadTbl, finTbl, filasSeleccionadas
    'borrar fila de origen
    eliminarRegistrosDeTabla cargaCalidadTbl, filasSeleccionadas
End Sub

'preguntar si este orden esta bien
Sub PintureriaFinal()

End Sub

