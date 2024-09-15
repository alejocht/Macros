Attribute VB_Name = "Módulo1"

Sub DividirArchivoPorFilas()

    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim rowsPerFile As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim fileCounter As Integer
    Dim newFileName As String
    Dim filePath As String
    Dim savePath As String
    Dim nombreHoja As String
    
    nombreHoja = Cells(3, 2).Value
    ' Validar que el nombre de hoja no sea vacio
    If IsEmpty(nombreHoja) Then
        MsgBox "La celda B3 está vacía. Por favor, ingresa el nombre de la hoja.", vbExclamation
        Exit Sub
    End If
    ' Define la hoja de trabajo activa
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    ' Define la cantidad de filas por archivo
    rowsPerFile = Cells(3, 1).Value ' Cambia esto a la cantidad de filas que desees por archivo
    
    ' Validar que rowsPerFile sea un número positivo
    If rowsPerFile <= 0 Then
        MsgBox "La cantidad de filas por archivo debe ser un número positivo.", vbExclamation
        Exit Sub
    End If
    
    ' Encuentra la última fila de datos
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Solicitar al usuario la carpeta para guardar los archivos
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecciona la carpeta para guardar los archivos"
        If .Show = -1 Then
            savePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Inicializar contadores
    fileCounter = 1
    startRow = 2 ' Suponiendo que la fila 1 contiene los encabezados
    
    ' Dividir el archivo en partes
    Do While startRow <= lastRow
        ' Crear un nuevo libro de trabajo
        Set wsNew = Workbooks.Add(xlWorksheet).Sheets(1)
        
        ' Copiar el encabezado
        ws.Rows(1).Copy Destination:=wsNew.Rows(1)
        
        ' Determinar la última fila del nuevo archivo
        endRow = Application.WorksheetFunction.Min(startRow + rowsPerFile - 1, lastRow)
        
        ' Copiar los datos al nuevo archivo
        ws.Rows(startRow & ":" & endRow).Copy Destination:=wsNew.Rows(2)
        
        ' Guardar el nuevo archivo
        newFileName = savePath & "\Parte_" & fileCounter & ".xlsx"
        wsNew.Parent.SaveAs Filename:=newFileName, FileFormat:=xlOpenXMLWorkbook
        wsNew.Parent.Close SaveChanges:=False
        
        ' Actualizar contadores
        fileCounter = fileCounter + 1
        startRow = endRow + 1
    Loop
    
    ' Mensaje de confirmación
    MsgBox "División completa. Archivos guardados en: " & savePath

End Sub
