Attribute VB_Name = "Funciones"
Option Explicit

Public Type TablaDatos
    Encabezados() As String
    Valores As Variant
End Type

Private Const ERR_PREFIX As Long = &H500

Public Function ObtenerTabla(ByVal nombreTabla As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set ObtenerTabla = ws.ListObjects(nombreTabla)
        On Error GoTo 0
        If Not ObtenerTabla Is Nothing Then Exit Function
    Next ws
    Err.Raise ERR_PREFIX + 1, "Funciones.ObtenerTabla", _
              "No se encontró la tabla '" & nombreTabla & "' en el libro." & vbCrLf & _
              "Verifique el nombre y que la tabla exista." 
End Function

Public Function CargarTabla(ByVal nombreTabla As String) As TablaDatos
    Dim lo As ListObject
    Dim datos As TablaDatos
    Dim totalColumnas As Long
    Dim origen As Variant
    Dim i As Long

    Set lo = ObtenerTabla(nombreTabla)
    totalColumnas = lo.ListColumns.Count

    ReDim datos.Encabezados(1 To totalColumnas)
    For i = 1 To totalColumnas
        datos.Encabezados(i) = CStr(lo.ListColumns(i).Name)
    Next i

    If lo.DataBodyRange Is Nothing Then
        datos.Valores = Empty
    Else
        origen = lo.DataBodyRange.Value
        If Not IsArray(origen) Then
            ReDim origen(1 To 1, 1 To totalColumnas)
            For i = 1 To totalColumnas
                origen(1, i) = lo.DataBodyRange.Cells(1, i).Value
            Next i
        End If
        datos.Valores = origen
    End If

    CargarTabla = datos
End Function

Public Function BuscarFila(ByVal nombreTabla As String, _
                           ByVal indiceColumna As Long, _
                           ByVal valorClave As Variant) As Long
    Dim lo As ListObject
    Dim arr As Variant
    Dim vector() As Variant
    Dim posicion As Variant

    Set lo = ObtenerTabla(nombreTabla)
    If lo.DataBodyRange Is Nothing Then Exit Function

    arr = lo.DataBodyRange.Columns(indiceColumna).Value
    vector = ConvertirAColumna(arr)
    If (Not IsArray(vector)) Or LBound(vector) > UBound(vector) Then Exit Function

    posicion = Application.Match(valorClave, vector, 0)
    If Not IsError(posicion) Then
        BuscarFila = CLng(posicion)
    End If
End Function

Public Function ActualizarFilaTabla(ByVal nombreTabla As String, _
                                    ByVal indiceColumnaClave As Long, _
                                    ByVal valorClave As Variant, _
                                    ByVal nuevosValores As Variant) As Boolean
    Dim lo As ListObject
    Dim fila As Long

    Set lo = ObtenerTabla(nombreTabla)
    fila = BuscarFila(nombreTabla, indiceColumnaClave, valorClave)
    If fila = 0 Then Exit Function

    If UBound(nuevosValores, 1) <> 1 Then
        Err.Raise ERR_PREFIX + 2, "Funciones.ActualizarFilaTabla", _
                  "El parámetro nuevosValores debe ser un arreglo de una sola fila." 
    End If

    lo.DataBodyRange.Rows(fila).Value = nuevosValores
    ActualizarFilaTabla = True
End Function

Public Function ConvertirAColumna(ByVal matriz As Variant) As Variant
    Dim resultado() As Variant
    Dim totalFilas As Long
    Dim i As Long

    If Not IsArray(matriz) Then
        ReDim resultado(1 To 1)
        resultado(1) = matriz
        ConvertirAColumna = resultado
        Exit Function
    End If

    On Error GoTo limpiar
    totalFilas = UBound(matriz, 1)
    ReDim resultado(1 To totalFilas)
    For i = 1 To totalFilas
        resultado(i) = matriz(i, 1)
    Next i
    ConvertirAColumna = resultado
    Exit Function

limpiar:
    ConvertirAColumna = Empty
End Function

Public Function CrearVectorFila(ByVal numeroColumnas As Long) As Variant
    Dim valores() As Variant
    ReDim valores(1 To 1, 1 To numeroColumnas)
    CrearVectorFila = valores
End Function

Public Function EsValorVacio(ByVal valor As Variant) As Boolean
    If IsMissing(valor) Then
        EsValorVacio = True
    ElseIf IsObject(valor) Then
        EsValorVacio = valor Is Nothing
    ElseIf IsError(valor) Then
        EsValorVacio = True
    ElseIf VarType(valor) = vbString Then
        EsValorVacio = Len(Trim$(valor)) = 0
    Else
        EsValorVacio = IsEmpty(valor)
    End If
End Function

Public Sub ExportarRango(ByVal destino As String, _
                         ByVal datos As Variant, _
                         ByVal encabezados() As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim filas As Long
    Dim columnas As Long

    If Not EsMatrizValida(datos) Then Exit Sub

    filas = UBound(datos, 1)
    columnas = UBound(datos, 2)

    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)

    For i = LBound(encabezados) To UBound(encabezados)
        ws.Cells(1, i).Value = encabezados(i)
        ws.Cells(1, i).Font.Bold = True
    Next i

    If filas >= 1 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(filas + 1, columnas)).Value = datos
    End If

    ws.ListObjects.Add(xlSrcRange, _
                       ws.Range(ws.Cells(1, 1), ws.Cells(filas + 1, columnas)), _
                       , xlYes).Name = "TablaExportada"
    ws.Cells.EntireColumn.AutoFit

    wb.SaveAs destino, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
End Sub

Public Function EsMatrizValida(ByVal datos As Variant) As Boolean
    On Error GoTo salir
    If Not IsArray(datos) Then Exit Function
    If IsEmpty(datos) Then Exit Function
    If UBound(datos, 1) < LBound(datos, 1) Then Exit Function
    If UBound(datos, 2) < LBound(datos, 2) Then Exit Function
    EsMatrizValida = True
salir:
End Function
