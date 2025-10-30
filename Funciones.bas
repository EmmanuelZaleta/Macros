Attribute VB_Name = "Funciones"
Option Explicit

' =====================================================================
' Configuración LoadFactor (único archivo)
' =====================================================================
Private Const LOADFACTOR_FILENAME As String = "ENSAMBLE_LOADFACTOR.TXT"

' Pon True si quieres EXCLUIR renglones cuyo campo(2) contenga "Troquel"
Private Const EXCLUIR_TROQUEL As Boolean = False

' ===== Config por defecto (edítalas si quieres) =====
Public Const DEFAULT_UNC As String = "\\Yazaki.local\na\elcom\chihuahua\Area_General\Materiales\Archivos Macro PCD\EP1\Extractor\"
Public Const DEFAULT_FILENAME As String = "ENSAMBLE_ORDER_STAT_Query.TXT"

Private Function TieneTroquel(ByVal s As String) As Boolean
    TieneTroquel = (InStr(1, s, "Troquel", vbTextCompare) > 0)
End Function

' === Helpers numéricos / texto =======================================
Private Function Num(ByVal s As String) As Double
    ' Convierte a Double tolerando comas y texto extra
    Num = Val(Replace(Trim$(s), ",", "."))
End Function

Private Function EficienciaDeTexto(ByVal s As String) As Double
    ' Extrae el número de un texto tipo "Efficiency 85%" -> 0.85
    Dim i As Long, c As String, buf As String, n As Double
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If (c >= "0" And c <= "9") Or c = "." Or c = "," Then buf = buf & c
    Next i
    If buf <> "" Then n = Num(buf)
    If n = 0 Then
        EficienciaDeTexto = 1#
    ElseIf n > 1.001 Then
        EficienciaDeTexto = n / 100#
    Else
        EficienciaDeTexto = n
    End If
End Function

' =====================================================================
' ORDENES
' =====================================================================
Sub traeInformacionOrdenes(pPlan As String, txtFechaFin As String)

    Dim rutaArchivos As String, vArchivo As String, fullPath As String
    Dim binData As String, lineas() As String, campos() As String
    Dim fila As Long, i As Long, fechaETD As Long, fechaFin As Long
    Dim anio As Integer, mes As Integer, dia As Integer
    Dim fechaConvertida As Date
    Dim wsDestino As Worksheet

    ' Obtener ruta y archivo dinámico
    rutaArchivos = ThisWorkbook.Sheets("Macro").Range("B1").Value
    vArchivo = buscaArchivo("Ordenes")
    fullPath = rutaArchivos & vArchivo

    If Dir(fullPath) = "" Then
        MsgBox "No se encontró el archivo: " & fullPath, vbCritical
        Exit Sub
    End If

    ' Validar y convertir txtFechaFin (formato aaaammdd)
    If Len(txtFechaFin) <> 8 Or Not IsNumeric(txtFechaFin) Then GoTo formatoInvalido

    anio = CInt(Left(txtFechaFin, 4))
    mes = CInt(Mid(txtFechaFin, 5, 2))
    dia = CInt(Right(txtFechaFin, 2))

    If mes < 1 Or mes > 12 Or dia < 1 Or dia > 31 Then GoTo formatoInvalido

    On Error GoTo formatoInvalido
    fechaConvertida = DateSerial(anio, mes, dia)
    fechaFin = CLng(Format(fechaConvertida, "yyyymmdd"))
    On Error GoTo 0

    ' Preparar hoja destino
    Set wsDestino = Workbooks(pPlan).Sheets("Orderstats")
    wsDestino.Cells.ClearContents
    wsDestino.Range("A1:J1").Value = Array("PartNo", "Control", "Item", "ETD", "Qty", "St", "PO", "Fecha PO", "Linea", "Planta")

    ' Leer archivo completo como texto (soporta LF o CRLF)
    Open fullPath For Binary As #1
        binData = Space$(LOF(1))
        Get #1, , binData
    Close #1

    ' Normalizar saltos de línea: CRLF y CR -> LF
    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    ' Procesar líneas (empezamos en 1 para saltar encabezado)
    fila = 2
    For i = 1 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            campos = Split(lineas(i), "|")

            If UBound(campos) >= 9 Then
                If IsNumeric(campos(3)) And Len(Trim(campos(3))) = 8 Then
                    fechaETD = CLng(campos(3))

                    If fechaETD <= fechaFin Then
                        With wsDestino
                            .Cells(fila, "A").Value = campos(0)
                            .Cells(fila, "B").Value = campos(1)
                            .Cells(fila, "C").Value = QuitarCerosIzquierda(campos(2))
                            .Cells(fila, "D").Value = DateSerial(Left(campos(3), 4), Mid(campos(3), 5, 2), Right(campos(3), 2))
                            .Cells(fila, "E").Value = campos(4)
                            .Cells(fila, "F").Value = campos(5)
                            .Cells(fila, "G").Value = campos(6)
                            .Cells(fila, "H").Value = campos(7)
                            .Cells(fila, "I").Value = campos(8)
                            .Cells(fila, "J").Value = campos(9)
                        End With
                        fila = fila + 1
                    End If
                End If
            End If
        End If
    Next i

    ' Validación final
    If fila <= 2 Then
        MsgBox "No se encontraron registros válidos en el archivo para la fecha indicada.", vbExclamation
        Exit Sub
    End If

    ' Ordenar por columna D (ETD)
    With wsDestino.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsDestino.Range("D2:D" & fila - 1), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsDestino.Range("A1:J" & fila - 1)
        .Header = xlYes
        .Apply
    End With

    Exit Sub

formatoInvalido:
    MsgBox "La fecha ingresada (" & txtFechaFin & ") no es válida. Usa el formato aaaammdd, por ejemplo: 20250424", vbCritical

End Sub



' ---------------------------
' Wrappers con filtro de fecha
' ---------------------------

' Carga solo hasta End (<= End), formato yyyymmdd (ej. "20251022")
Public Sub CargarOrderStat_DesdeUNC_Hasta(pPlan As String, _
                                          endYYYYMMDD As String, _
                                          Optional carpeta As String = vbNullString, _
                                          Optional archivo As String = vbNullString)
    If Len(carpeta) = 0 Then carpeta = DEFAULT_UNC
    If Len(archivo) = 0 Then archivo = DEFAULT_FILENAME

    Dim fullPath As String
    If Right$(carpeta, 1) <> "\" And Right$(carpeta, 1) <> "/" Then carpeta = carpeta & "\"
    fullPath = carpeta & archivo

    CargarOrderStatDesdeArchivo pPlan, fullPath, vbNullString, endYYYYMMDD
End Sub

' Carga entre Start y End (ambos inclusivos), formato yyyymmdd
Public Sub CargarOrderStat_DesdeUNC_Rango(pPlan As String, _
                                          startYYYYMMDD As String, _
                                          endYYYYMMDD As String, _
                                          Optional carpeta As String = vbNullString, _
                                          Optional archivo As String = vbNullString)
    If Len(carpeta) = 0 Then carpeta = DEFAULT_UNC
    If Len(archivo) = 0 Then archivo = DEFAULT_FILENAME

    Dim fullPath As String
    If Right$(carpeta, 1) <> "\" And Right$(carpeta, 1) <> "/" Then carpeta = carpeta & "\"
    fullPath = carpeta & archivo

    CargarOrderStatDesdeArchivo pPlan, fullPath, startYYYYMMDD, endYYYYMMDD
End Sub
Public Sub CargarOrderStatDesdeArchivo(pPlan As String, fullPath As String, _
                                       Optional startYYYYMMDD As String = "", _
                                       Optional endYYYYMMDD As String = "")
    Dim binData As String, lineas() As String, campos() As String
    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, fila As Long, fnum As Integer

    If Dir(fullPath) = "" Then
        MsgBox "No se encontró el archivo:" & vbCrLf & fullPath, vbCritical
        Exit Sub
    End If

    ' === Parseo de fechas de filtro ===
    Dim hasStart As Boolean, hasEnd As Boolean
    Dim dStart As Date, dEnd As Date
    If Len(startYYYYMMDD) > 0 Then dStart = ParseYYYYMMDD(startYYYYMMDD): hasStart = True
    If Len(endYYYYMMDD) > 0 Then dEnd = ParseYYYYMMDD(endYYYYMMDD): hasEnd = True

    ' === Selección de hoja destino ===
    Set wb = Workbooks(pPlan)
    Set ws = EnsureSheet(wb, "OrderStats")

    ' === Encabezados ===
    ws.Cells.ClearContents
    ws.Range("A1:J1").Value = Array("CUST. CD.", "S/T", "PARTNO", "ETD", "ETA", _
                                    "QUANTITY", "SHIPPING QTY", "Remain", "CUST. PO", "ORDER FLG")

    ' === Lectura de archivo ===
    fnum = FreeFile
    Open fullPath For Binary As #fnum
        binData = Space$(LOF(fnum))
        Get #fnum, , binData
    Close #fnum

    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    fila = 2
    For i = LBound(lineas) To UBound(lineas)
        Dim ln As String: ln = Trim$(lineas(i))
        If ln <> "" Then
            If Left$(ln, 1) = "|" _
               And InStr(1, ln, "|Plnt|", vbTextCompare) = 0 _
               And InStr(1, ln, "---", vbTextCompare) = 0 _
               And InStr(1, ln, "Demanda de cliente", vbTextCompare) = 0 Then

                campos = Split(ln, "|")
                If UBound(campos) >= 7 Then
                    On Error GoTo SaltarFila

                    Dim vPlnt As String, vMat As String, vFecha As Date
                    Dim vSched As Double, vIssued As Double, vPO As String, vDeliv As Double
                    Dim vFecha2 As Date

                    vPlnt = "YZP" & Trim$(campos(1))
                    vMat = Trim$(campos(3))
                    vFecha = Trim$(campos(4))
                    vFecha2 = DateAdd("d", -Val(campos(8)), vFecha)
                    vSched = CleanNumber(campos(5))
                    vIssued = CleanNumber(campos(6))
                    vPO = Trim$(campos(7))
                    vDeliv = CleanNumber(campos(8))

                    ' Filtro por fechas (inclusivo)
                    If hasStart And vFecha < dStart Then GoTo SaltarFila
                    If hasEnd And vFecha > dEnd Then GoTo SaltarFila

                    ' === Escribir fila ===
                    With ws
                        .Cells(fila, 1).Value = vPlnt
                        .Cells(fila, 2).Value = vPlnt
                        .Cells(fila, 3).Value = "'" & vMat                   ' Mantiene como texto
                        .Cells(fila, 4).Value = vFecha2  ' ETD
                        .Cells(fila, 5).Value = vFecha  ' ETA
                        .Cells(fila, 6).Value = vSched
                        .Cells(fila, 7).Value = vIssued
                        .Cells(fila, 8).Value = vSched - vIssued             ' Remain = Qty - Shipped
                        .Cells(fila, 9).Value = vPO
                        .Cells(fila, 11).Value = vFecha2 ' ETD
                        Dim hoy As Date, limite As Date
                        hoy = Date
                        limite = hoy + 7   ' Ejemplo: 7 días adelante, reemplázalo por lo que uses como lv_datu2
                        
                        If vFecha >= hoy And vFecha <= limite Then
                            .Cells(fila, 10).Value = "O"  ' Orden actual (Open)
                        ElseIf vFecha > limite Then
                            .Cells(fila, 10).Value = "F"  ' Futura
                        ElseIf vFecha < hoy Then
                            .Cells(fila, 10).Value = "P"  ' Pasada
                        End If

                        ' Formatos específicos
                        .Cells(fila, 3).NumberFormat = "General"
                    End With
                    fila = fila + 1
                End If
            End If
        End If
SaltarFila:
        On Error GoTo 0
    Next i

    If fila <= 2 Then
        Dim msg As String: msg = "No hubo filas dentro del rango."
        If hasStart Then msg = msg & vbCrLf & "Desde: " & Format$(dStart, "yyyy-mm-dd")
        If hasEnd Then msg = msg & vbCrLf & "Hasta: " & Format$(dEnd, "yyyy-mm-dd")
        MsgBox msg, vbExclamation
        Exit Sub
    End If

    ' === Ajustes finales ===
    With ws
        .Columns("A:J").AutoFit
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A1:J" & fila - 1).AutoFilter
    End With

    MsgBox "Cargado Order Stat en 'OrderStatRaw' (" & (fila - 2) & " filas).", vbInformation
    
    ' ===== SEGUNDO ARCHIVO: ENSAMBLE_ORDER_STAT_Query2.TXT =====
Dim p As Long, carpeta2 As String, fullPath2 As String
Dim t As String
p = InStrRev(fullPath, "\"): If p = 0 Then p = InStrRev(fullPath, "/")
If p > 0 Then carpeta2 = Left$(fullPath, p) Else carpeta2 = ""
fullPath2 = carpeta2 & "ENSAMBLE_ORDER_STAT_Query2.TXT"

If Dir(fullPath2) <> "" Then
    fnum = FreeFile
    Open fullPath2 For Binary As #fnum
        binData = Space$(LOF(fnum))
        Get #fnum, , binData
    Close #fnum

    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    For i = LBound(lineas) To UBound(lineas)
        Dim ln2 As String: ln2 = Trim$(lineas(i))
        If ln2 <> "" Then
            If Left$(ln2, 1) = "|" _
               And InStr(1, ln2, "|Plnt|", vbTextCompare) = 0 _
               And InStr(1, ln2, "---", vbTextCompare) = 0 Then

                campos = Split(ln2, "|")
                If UBound(campos) >= 9 Then
                    On Error GoTo Saltar2

                    Dim vCust As String, vPlnt2 As String, vMat2 As String
                    Dim vGI As Date, vDeliv2 As Date
                    Dim vQty As Double, vShip As Double, vRemain As Double
                    Dim vDoc As String, vSold As String

                    vCust = Trim$(campos(1))
                    vPlnt2 = Trim$(campos(2))
                    vMat2 = Trim$(campos(3))

                    ' Excluir materiales que inician con "M"
                    If Len(vMat2) > 0 Then
                        If UCase$(Left$(vMat2, 1)) = "M" Then GoTo Saltar2
                    End If

                    ' Fechas: MM/DD/YYYY o YYYYMMDD
                    t = Trim$(campos(4))
                    If Len(t) = 10 And Mid$(t, 3, 1) = "/" And Mid$(t, 6, 1) = "/" Then
                        vGI = DateSerial(CInt(Right$(t, 4)), CInt(Left$(t, 2)), CInt(Mid$(t, 4, 2)))
                    ElseIf Len(t) = 8 And IsNumeric(t) Then
                        vGI = DateSerial(CInt(Left$(t, 4)), CInt(Mid$(t, 5, 2)), CInt(Right$(t, 2)))
                    Else
                        vGI = CDate(t)
                    End If

                    t = Trim$(campos(5))
                    If Len(t) = 10 And Mid$(t, 3, 1) = "/" And Mid$(t, 6, 1) = "/" Then
                        vDeliv2 = DateSerial(CInt(Right$(t, 4)), CInt(Left$(t, 2)), CInt(Mid$(t, 4, 2)))
                    ElseIf Len(t) = 8 And IsNumeric(t) Then
                        vDeliv2 = DateSerial(CInt(Left$(t, 4)), CInt(Mid$(t, 5, 2)), CInt(Right$(t, 2)))
                    Else
                        vDeliv2 = CDate(t)
                    End If

                    vQty = CleanNumber(campos(6))
                    If Trim$(campos(7)) = "" Then
                        vShip = 0
                    Else
                        vShip = CleanNumber(campos(7))
                    End If
                    vRemain = vQty - vShip

                    vDoc = Trim$(campos(8))
                    vSold = Trim$(campos(9))

                    With ws
                        .Cells(fila, 1).Value = vCust
                        .Cells(fila, 2).Value = vPlnt2
                        .Cells(fila, 3).Value = "'" & vMat2
                        .Cells(fila, 4).Value = vGI
                        .Cells(fila, 5).Value = vDeliv2
                        .Cells(fila, 6).Value = vQty
                        .Cells(fila, 7).Value = vShip
                        .Cells(fila, 8).Value = vRemain
                        .Cells(fila, 9).Value = vDoc
                        .Cells(fila, 11).Value = vGI
                        If vFecha >= hoy And vFecha <= limite Then
                            .Cells(fila, 10).Value = "O"  ' Orden actual (Open)
                        ElseIf vFecha > limite Then
                            .Cells(fila, 10).Value = "F"  ' Futura
                        ElseIf vFecha < hoy Then
                            .Cells(fila, 10).Value = "P"  ' Pasada
                        End If

                        .Cells(fila, 3).NumberFormat = "General"
                    End With
                    fila = fila + 1
                End If
            End If
        End If
Saltar2:
        On Error GoTo 0
    Next i
End If

End Sub


' ----------------- Helpers -----------------
Private Function EnsureSheet(wb As Workbook, nombre As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nombre, vbTextCompare) = 0 Then
            Set EnsureSheet = ws
            Exit Function
        End If
    Next ws
    Set EnsureSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    EnsureSheet.Name = nombre
End Function

Private Function CleanNumber(ByVal s As String) As Double
    s = Replace$(s, ",", "")
    s = Replace$(s, " ", "")
    s = Trim$(s)
    If Len(s) = 0 Or s = "-" Then
        CleanNumber = 0
    Else
        CleanNumber = CDbl(s)
    End If
End Function

Private Function ParseDateMMDDYYYY(ByVal s As String) As Date
    Dim mm As Integer, dd As Integer, yy As Integer
    s = Trim$(s)
    If Len(s) = 10 And Mid$(s, 3, 1) = "/" And Mid$(s, 6, 1) = "/" Then
        mm = CInt(Left$(s, 2))
        dd = CInt(Mid$(s, 4, 2))
        yy = CInt(Right$(s, 4))
        ParseDateMMDDYYYY = DateSerial(yy, mm, dd)
    Else
        ParseDateMMDDYYYY = CDate(s)
    End If
End Function

Private Function ParseYYYYMMDD(ByVal s As String) As Date
    Dim y As Integer, m As Integer, d As Integer
    s = Trim$(s)
    If Len(s) <> 8 Or Not IsNumeric(s) Then Err.Raise 5, , "Fecha inválida (yyyymmdd): " & s
    y = CInt(Left$(s, 4))
    m = CInt(Mid$(s, 5, 2))
    d = CInt(Right$(s, 2))
    ParseYYYYMMDD = DateSerial(y, m, d)
End Function
' =====================================================================
' LOAD FACTOR (Sólo ENSAMBLE_LOADFACTOR.TXT) con Capacidad
' =====================================================================
Sub TraeInformacionLoadFactor(pPlan As String)

    Dim rutaArchivos As String
    Dim vArchivoLF As String
    Dim wsDestino As Worksheet
    Dim fila As Long
    Dim dictUnicos As Object

    On Error GoTo manejaError

    Set dictUnicos = CreateObject("Scripting.Dictionary")

    ' Obtener ruta y hoja destino
    rutaArchivos = ThisWorkbook.Sheets("Macro").Range("B1").Value
    Set wsDestino = Workbooks(pPlan).Sheets("Load Factor")
    wsDestino.Cells.ClearContents
    fila = 2

    ' Encabezados (según tu hoja)
    wsDestino.Range("A1:O1").Value = Array( _
        "PartNo", "CONTROL", "DIE", "Dep", "Group Code", "Eng Lev", _
        "Std Cav", "Act Cav", "Cycle Time", "Piece Weight", "Shot Weight", _
        "Pcs/Hour", "Capacidad", "Ensamble", "")

    ' Único archivo
    vArchivoLF = LOADFACTOR_FILENAME

    ' Procesar archivo
    LlamaArchivo rutaArchivos & vArchivoLF, wsDestino, fila, dictUnicos
    ActualizarLoadFactorDesdeMDMQ0400_Fast pPlan, True
    Exit Sub

manejaError:
    MsgBox "Se produjo un error: " & Err.Description, vbCritical
    Close #1

End Sub

Sub LlamaArchivo(fullPath As String, wsDestino As Worksheet, ByRef fila As Long, dictUnicos As Object)

    Dim binData As String, lineas() As String, linea As String
    Dim campos() As String
    Dim i As Long
    Dim claveUnica As String
    Dim partNo As String, campoDie As String
    Dim basePH As Double, cap As Double
    Dim eff As Double, ct As Double, actCav As Double

    If Dir(fullPath) = "" Then
        MsgBox "No se encontró: " & fullPath, vbCritical
        Exit Sub
    End If

    ' Leer todo el archivo (soporta LF/CRLF/CR y UTF-8 BOM)
    Open fullPath For Binary As #1
        binData = Space$(LOF(1))
        Get #1, , binData
    Close #1

    ' Quitar BOM UTF-8 si existe
    If Len(binData) >= 3 Then
        If Left$(binData, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
            binData = Mid$(binData, 4)
        End If
    End If

    ' Normalizar saltos de línea
    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    ' Recorrer renglones (1..UBound) para saltar encabezado
    For i = 1 To UBound(lineas)
        linea = Trim$(lineas(i))
        If linea <> "" Then
            If InStr(linea, "|") > 0 Then
                campos = Split(linea, "|")

                ' Asegurar columnas suficientes para cálculo (>=20 incluye Working Rate)
                If UBound(campos) >= 20 Then

                    ' Filtro opcional de "Troquel" en campo(2)
                    If EXCLUIR_TROQUEL And TieneTroquel(campos(2)) Then GoTo siguiente

                    partNo = QuitarCerosIzquierda(campos(0))
                    campoDie = Trim$(campos(2))

                    ' Clave única por PartNo + DIE recortado a 5 (si aplica)
                    If Len(campoDie) > 5 Then
                        claveUnica = partNo & Left$(campoDie, 5)
                    Else
                        claveUnica = partNo & campoDie
                    End If

                    If Not dictUnicos.exists(claveUnica) Then
                        dictUnicos.Add claveUnica, vbNullString

                        ' ---- Cálculo de Capacidad ----
                        ' 1) Pcs/Hour si viene; si no, 3600/CycleTime * (ActCav si >0, si no 1)
                        If IsNumeric(campos(11)) Then
                            basePH = Num(campos(11))
                        Else
                            ct = Num(campos(9))                  ' Cycle Time
                            actCav = Num(campos(8))              ' Act Cav
                            If actCav <= 0 Then actCav = 1
                            If ct > 0 Then basePH = (3600# / ct) * actCav Else basePH = 0
                        End If
                        ' 2) Eficiencia desde "Working Rate" (col 20, ej. "Efficiency 85%")
                        eff = EficienciaDeTexto(campos(20))
                        ' 3) Capacidad final
                        cap = basePH * eff

                        ' --- Mapeo de columnas ---
                        With wsDestino
                            .Cells(fila, "A").Value = partNo                 ' PartNo (0)
                            .Cells(fila, "B").Value = campos(3)              ' Control
                            .Cells(fila, "C").Value = campoDie               ' Die
                            .Cells(fila, "D").Value = campos(4)              ' Dep
                            .Cells(fila, "E").Value = campos(5)              ' Group Code
                            .Cells(fila, "F").Value = campos(6)              ' Eng Lev
                            .Cells(fila, "G").Value = campos(7)              ' Std Cav
                            .Cells(fila, "H").Value = campos(8)              ' Act Cav
                            .Cells(fila, "I").Value = campos(9)              ' Cycle Time
                            .Cells(fila, "J").Value = campos(10)             ' Piece Weight
                            .Cells(fila, "K").Value = ""                     ' Shot Weight (si aplica)
                            .Cells(fila, "L").Value = basePH                 ' Pcs/Hour base
                            .Cells(fila, "M").Value = cap                    ' Capacidad
                            .Cells(fila, "N").Value = ""                     ' Ensamble (placeholder)
                        End With

                        fila = fila + 1
                    End If
                End If
            End If
        End If
siguiente:
    Next i

    ' Formato visual
    If fila > 2 Then
        wsDestino.Rows(1).Font.Bold = True
        wsDestino.Columns("A:O").AutoFit
    End If

End Sub

' =====================================================================
' ITEM MASTER
' =====================================================================
Sub traeInformacionItemMaster(pPlan As String)

    Dim rutaArchivo As String
    Dim archivo As String
    Dim binData As String
    Dim lineas() As String
    Dim datos() As String
    Dim fila As Long
    Dim wsDestino As Worksheet
    Dim valorModC As String, valorModD As String
    Dim i As Long

    On Error GoTo manejaError

    ' Obtener ruta y archivo
    rutaArchivo = ThisWorkbook.Sheets("Macro").Range("B1").Value
    archivo = rutaArchivo & "ENSAMBLE_ITEMMASTER.TXT"

    If Dir(archivo) = "" Then
        MsgBox "No se encontró el archivo: " & archivo, vbCritical
        Exit Sub
    End If

    ' Preparar hoja destino
    Set wsDestino = Workbooks(pPlan).Sheets("Item Master")
    ' wsDestino.Cells.ClearContents
    fila = 2

    ' Leer archivo como texto completo
    Open archivo For Binary As #1
        binData = Space$(LOF(1))
        Get #1, , binData
    Close #1

    ' Normalizar saltos de línea
    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    ' Procesar líneas (empezando en 1 para saltar encabezado)
    For i = 1 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            datos = Split(lineas(i), "|")
            If UBound(datos) >= 8 Then
                ' Lógica de modificación
                Select Case UCase(Left(Trim(datos(2)), 1))
                    Case "E": valorModC = "3"
                    Case "F": valorModC = "2"
                    Case Else: valorModC = "DEFAULT"
                End Select

                Select Case UCase(Left(Trim(datos(3)), 2))
                    Case "Z1": valorModD = "3"
                    Case "Z2": valorModD = "2"
                    Case Else: valorModD = "DEFAULT"
                End Select

                ' Llenar datos
                With wsDestino
                    .Cells(fila, "A").Value = QuitarCerosIzquierda(CStr(datos(0)))
                    .Cells(fila, "B").Value = datos(1)
                    .Cells(fila, "C").Value = datos(2)
                    .Cells(fila, "D").Value = datos(5)
                    .Cells(fila, "E").Value = datos(3)
                    .Cells(fila, "F").Value = valorModC
                    .Cells(fila, "G").Value = valorModD
                    .Cells(fila, "H").Value = datos(6)
                    .Cells(fila, "J").Value = datos(7)
                End With

                fila = fila + 1
            End If
        End If
    Next i

    Exit Sub

manejaError:
    MsgBox "Se produjo un error: " & Err.Description, vbCritical
    Close #1

End Sub

' =====================================================================
' HELPERS BÁSICOS
' =====================================================================
Function buscaArchivo(pNombre As String) As String
    Dim vLstRen As Long
    Dim h As Integer
    With ThisWorkbook.Sheets("Macro")
        vLstRen = .Range("A65536").End(xlUp).Row
        For h = 1 To vLstRen
            If .Range("A" & h).Value = pNombre Then
                buscaArchivo = .Range("B" & h).Value
                Exit Function
            End If
        Next h
    End With
End Function

Function QuitarCerosIzquierda(txt As String) As String
    If Len(txt) = 0 Then
        QuitarCerosIzquierda = ""
        Exit Function
    End If
    Do While Left(txt, 1) = "0" And Len(txt) > 1
        txt = Mid(txt, 2)
    Loop
    QuitarCerosIzquierda = txt
End Function

' =====================================================================
' FLEX PLAN (desde ENSAMBLE_ORDER_STAT.TXT a matriz por semanas)
' =====================================================================
Sub traeInformacionFlexPlan(pPlan As String)

    Dim wsDestino As Worksheet
    Dim fullPath As String
    Dim binData As String, lineas() As String, campos() As String
    Dim dictProyeccion As Object, dictPartes As Object
    Dim semanas() As Variant
    Dim fechaDato As Date
    Dim partNumber As Variant
    Dim cantidad As Double
    Dim fila As Long, colBase As Long, i As Long
    Dim claveSemana As String
    Dim minFecha As Date, maxFecha As Date
    Dim idx As Long

    ' Inicialización
    Set dictProyeccion = CreateObject("Scripting.Dictionary")
    Set dictPartes = CreateObject("Scripting.Dictionary")
    colBase = 3 ' Columna C en adelante son las semanas

    ' Ruta del archivo fuente
    fullPath = "\\Yazaki.local\na\elcom\chihuahua\Area_General\Materiales\Archivos Macro PCD\EP1\Extractor\ENSAMBLE_ORDER_STAT.TXT"

    ' Leer archivo completo como texto y normalizar saltos de línea
    Open fullPath For Binary As #1
        binData = Space$(LOF(1))
        Get #1, , binData
    Close #1

    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    ' Detectar fechas mínima y máxima
    minFecha = DateSerial(2099, 12, 31)
    maxFecha = DateSerial(2000, 1, 1)

    For i = 1 To UBound(lineas) ' Saltamos encabezado (línea 0)
        If Trim(lineas(i)) <> "" Then
            campos = Split(lineas(i), "|")
            If UBound(campos) >= 5 Then
                If IsNumeric(campos(3)) And Len(campos(3)) = 8 Then
                    fechaDato = DateSerial(Left(campos(3), 4), Mid(campos(3), 5, 2), Right(campos(3), 2))
                    If fechaDato < minFecha Then minFecha = fechaDato
                    If fechaDato > maxFecha Then maxFecha = fechaDato
                End If
            End If
        End If
    Next i

    ' Ajustar a domingo más cercano
    minFecha = minFecha - Weekday(minFecha, vbSunday) + 1
    maxFecha = maxFecha + (7 - Weekday(maxFecha, vbSunday))

    ' Construir arreglo de semanas
    ReDim semanas(0 To 0)
    i = 0
    Do While minFecha <= maxFecha
        ReDim Preserve semanas(i)
        semanas(i) = Array(minFecha, minFecha + 6)
        minFecha = minFecha + 7
        i = i + 1
    Loop

    ' Leer datos y llenar diccionarios
    For i = 1 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            campos = Split(lineas(i), "|")
            If UBound(campos) >= 5 Then
                partNumber = Trim(campos(2))
                If partNumber <> "" And IsNumeric(campos(5)) Then
                    cantidad = Val(campos(5))
                    If IsNumeric(campos(3)) And Len(campos(3)) = 8 Then
                        fechaDato = DateSerial(Left(campos(3), 4), Mid(campos(3), 5, 2), Right(campos(3), 2))
                        For idx = 0 To UBound(semanas)
                            If fechaDato >= semanas(idx)(0) And fechaDato <= semanas(idx)(1) Then
                                claveSemana = partNumber & "|" & idx
                                dictProyeccion(claveSemana) = dictProyeccion(claveSemana) + cantidad
                                Exit For
                            End If
                        Next idx
                        dictPartes(partNumber) = True
                    End If
                End If
            End If
        End If
    Next i

    ' Crear hoja destino
    Set wsDestino = Workbooks(pPlan).Sheets("Flex-plan")
    wsDestino.Cells.Clear

    ' Encabezados
    wsDestino.Range("A1").Value = "Common Name"
    wsDestino.Range("B1").Value = "Part NO"
    For i = 0 To UBound(semanas)
        wsDestino.Cells(1, colBase + i).Value = _
            Format(semanas(i)(0), "m/d/yyyy") & " - " & Format(semanas(i)(1), "m/d/yyyy")
    Next i

    ' Llenar datos
    fila = 2
    For Each partNumber In dictPartes.Keys
        wsDestino.Cells(fila, 1).Value = " "  ' Common Name fijo
        wsDestino.Cells(fila, 2).Value = partNumber
        For i = 0 To UBound(semanas)
            claveSemana = partNumber & "|" & i
            If dictProyeccion.exists(claveSemana) Then
                wsDestino.Cells(fila, colBase + i).Value = dictProyeccion(claveSemana)
            Else
                wsDestino.Cells(fila, colBase + i).Value = 0
            End If
        Next i
        fila = fila + 1
    Next

    ' Formato final
    wsDestino.Rows(1).Font.Bold = True
    wsDestino.Columns("A:Z").AutoFit

End Sub

' =====================================================================
' INVENTARIO FG
' =====================================================================
Sub traeInformacionInventarioFG(pPlan As String)

    Dim rutaArchivos As String
    Dim vArchivo As String
    Dim fullPath As String
    Dim binData As String
    Dim lineas() As String
    Dim campos() As String
    Dim dictSumatoria As Object
    Dim clave As Variant
    Dim fila As Long
    Dim wsDestino As Worksheet
    Dim partNumber As String, fechaInj As String, invLoc As String
    Dim boxUnit As Double
    Dim i As Long
    Dim totalFiltradosHOLD As Long

    ' Obtener ruta y archivo
    rutaArchivos = ThisWorkbook.Sheets("Macro").Range("B1").Value
    vArchivo = buscaArchivo("InvLocWIPFG")
    fullPath = rutaArchivos & vArchivo

    If Dir(fullPath) = "" Then
        MsgBox "No se encontró el archivo: " & fullPath, vbCritical
        Exit Sub
    End If

    ' Inicializar objetos
    Set dictSumatoria = CreateObject("Scripting.Dictionary")
    Set wsDestino = Workbooks(pPlan).Sheets("Inventario FG")
    wsDestino.Cells.ClearContents

    ' Leer archivo como texto normalizando saltos de línea
    Open fullPath For Binary As #1
        binData = Space$(LOF(1))
        Get #1, , binData
    Close #1

    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    ' Agrupar datos
    For i = 1 To UBound(lineas) ' Saltar encabezado (línea 0)
        If Trim(lineas(i)) <> "" Then
            campos = Split(lineas(i), "|")

            If UBound(campos) >= 5 Then
                ' Excluir si campos(1) inicia con "HOLD"
                If UCase(Left(Trim(campos(1)), 4)) = "HOLD" Then
                    totalFiltradosHOLD = totalFiltradosHOLD + 1
                    GoTo SaltarLinea
                End If

                partNumber = QuitarCerosIzquierda(campos(3))
                fechaInj = Trim(campos(4))
                boxUnit = Val(campos(2))
                invLoc = Trim(campos(1))

                If partNumber <> "" And fechaInj <> "" And invLoc <> "" Then
                    clave = partNumber & "|" & fechaInj & "|" & invLoc
                    If dictSumatoria.exists(clave) Then
                        dictSumatoria(clave) = dictSumatoria(clave) + boxUnit
                    Else
                        dictSumatoria.Add clave, boxUnit
                    End If
                End If
            End If
        End If
SaltarLinea:
    Next i

    ' Escribir resultados
    fila = 2
    wsDestino.Range("A1:H1").Value = Array("SEQN2", "PARTNO", "DIE NO.", "BOX UNIT", "INV LOCATION", "", "", "INJ. DATE")

    For Each clave In dictSumatoria.Keys
        Dim clavePartes() As String
        clavePartes = Split(clave, "|")
        partNumber = clavePartes(0)
        fechaInj = clavePartes(1)
        invLoc = clavePartes(2)

        With wsDestino
            .Cells(fila, "B").Value = partNumber
            .Cells(fila, "D").Value = dictSumatoria(clave)
            .Cells(fila, "E").Value = invLoc
            .Cells(fila, "H").Value = fechaInj
        End With

        fila = fila + 1
    Next
End Sub

' =====================================================================
' WIP (InvLocWIP + InvCompon)
' =====================================================================
Sub traeInformacionInvLocWIP(pPlan As String)

    Dim rutaArchivos As String
    Dim vArchivo As String, vArchivo2 As String
    Dim fullPath As String, fullPath2 As String
    Dim binData As String, binData2 As String
    Dim lineas() As String, lineas2() As String
    Dim campos() As String
    Dim fila As Long, i As Long
    Dim wsDestino As Worksheet
    Dim linea2 As String  ' Para mostrar en error

    ' Obtener ruta
    rutaArchivos = ThisWorkbook.Sheets("Macro").Range("B1").Value

    ' Archivos a procesar
    vArchivo = buscaArchivo("InvLocWIP")
    vArchivo2 = buscaArchivo("InvCompon")
    fullPath = rutaArchivos & vArchivo
    fullPath2 = rutaArchivos & vArchivo2

    ' Validar existencia de ambos archivos
    If Dir(fullPath) = "" Then
        MsgBox "No se encontró el archivo: " & fullPath, vbCritical
        Exit Sub
    End If

    If Dir(fullPath2) = "" Then
        MsgBox "No se encontró el archivo: " & fullPath2, vbCritical
        Exit Sub
    End If

    ' Preparar hoja destino
    Set wsDestino = Workbooks(pPlan).Sheets("WIP")
    wsDestino.Cells.ClearContents
    fila = 2

    ' === Leer InvLocWIP ===
    Open fullPath For Binary As #1
        binData = Space$(LOF(1))
        Get #1, , binData
    Close #1

    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    ' Procesar InvLocWIP
    For i = 1 To UBound(lineas) ' Saltamos encabezado
        If Trim(lineas(i)) <> "" Then
            campos = Split(lineas(i), "|")
            If UBound(campos) >= 5 Then
                On Error GoTo errorHandler
                With wsDestino
                    .Cells(fila, "A").Value = campos(1) ' INV Location
                    If IsNumeric(campos(2)) Then .Cells(fila, "B").Value = Val(campos(2))
                    .Cells(fila, "C").Value = QuitarCerosIzquierda(campos(3)) ' Part#
                    .Cells(fila, "D").Value = campos(4) ' Inj. Date
                    .Cells(fila, "E").Value = campos(5) ' Dept
                    .Cells(fila, "F").Value = "InvLocWIP" ' Origen
                End With
                fila = fila + 1
                On Error GoTo 0
            End If
        End If
    Next i

    ' === Leer InvCompon ===
    Open fullPath2 For Binary As #2
        binData2 = Space$(LOF(2))
        Get #2, , binData2
    Close #2

    binData2 = Replace(binData2, vbCrLf, vbLf)
    binData2 = Replace(binData2, vbCr, vbLf)
    lineas2 = Split(binData2, vbLf)

    ' Procesar InvCompon
    For i = 1 To UBound(lineas2) ' Saltamos encabezado
        linea2 = lineas2(i)
        If Trim(linea2) <> "" Then
            If InStr(linea2, "|") = 0 Then GoTo continuar2

            campos = Split(linea2, "|")
            If UBound(campos) >= 6 Then
                On Error GoTo errorHandler
                With wsDestino
                    .Cells(fila, "A").Value = campos(3)
                    If IsNumeric(campos(4)) Then .Cells(fila, "B").Value = Val(campos(4))
                    .Cells(fila, "C").Value = QuitarCerosIzquierda(campos(0))
                    .Cells(fila, "D").Value = campos(6)
                    .Cells(fila, "E").Value = campos(5)
                    .Cells(fila, "F").Value = "InvCompon"
                End With
                fila = fila + 1
                On Error GoTo 0
            End If
        End If
continuar2:
    Next i

   Exit Sub

' === MANEJO DE ERRORES GENERAL ===
errorHandler:
    MsgBox "ERROR - Fila: " & fila & vbCrLf & _
           "Línea: " & linea2 & vbCrLf & _
           "Descripción: " & Err.Description, vbCritical
    On Error Resume Next
    Close #1
    Close #2
    Exit Sub

End Sub

' =====================================================================
' DATASET LIMPIO LOADFACTOR (único archivo)
' =====================================================================
Sub ConstruirDatasetLimpioLoadFactor(ByRef dictDataset As Object)

    Dim rutaBase As String
    Dim fullPath As String
    Dim binData As String
    Dim lineas() As String
    Dim campos() As String
    Dim i As Long
    Dim campo0 As String, campo2 As String, campo11 As Double
    Dim claveUnica As String

    ' Inicializar
    Set dictDataset = CreateObject("Scripting.Dictionary")
    rutaBase = ThisWorkbook.Sheets("Macro").Range("B1").Value
    fullPath = rutaBase & LOADFACTOR_FILENAME

    If Dir(fullPath) = "" Then Exit Sub

    ' Leer archivo binario
    Open fullPath For Binary As #1
        binData = Space$(LOF(1))
        Get #1, , binData
    Close #1

    ' Normalizar saltos de línea
    binData = Replace(binData, vbCrLf, vbLf)
    binData = Replace(binData, vbCr, vbLf)
    lineas = Split(binData, vbLf)

    ' Procesar líneas (desde 1 para saltar encabezado)
    For i = 1 To UBound(lineas)
        If Trim(lineas(i)) <> "" Then
            campos = Split(lineas(i), "|")
            If UBound(campos) >= 11 Then
                ' Filtro opcional de "Troquel" en campo(2)
                If EXCLUIR_TROQUEL And TieneTroquel(campos(2)) Then
                    ' Saltar
                Else
                    campo0 = QuitarCerosIzquierda(Trim(campos(0))) ' PartNo
                    If Len(Trim(campos(2))) >= 5 Then
                        campo2 = Left(Trim(campos(2)), 5)          ' Die recortado
                    Else
                        campo2 = Trim(campos(2))
                    End If

                    ' Clave única = PartNo + Die(recortado)
                    claveUnica = campo0 & campo2

                    If Not dictDataset.exists(claveUnica) Then
                        If IsNumeric(campos(11)) Then
                            campo11 = CDbl(campos(11))              ' Pcs/hour
                        Else
                            campo11 = 0
                        End If
                        dictDataset.Add claveUnica, Array(campo0, campo2, campo11)
                    End If
                End If
            End If
        End If
    Next i

End Sub

' =====================================================================
' CAPACIDADES (usa sólo ENSAMBLE_LOADFACTOR.TXT + ItemMaster + BOM)
' =====================================================================
Sub traeInformacionCapacidades(pPlan As String)

    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim rutaBase As String, archivoLF As String, archivoItemMaster As String
    Dim binData As String, lineas() As String, campos() As String
    Dim partNo As String, child As String, claveUnica As String
    Dim dictCapEnsamble As Object, dictCapMoldeo As Object, dictUnicos As Object
    Dim dictBusquedaItemMaster As Object
    Dim dictExcluirPartNosZ2 As Object, dictExcluirChildsZ2F As Object
    Dim i As Long, fila As Long, destinoFila As Long
    Dim datos() As String
    Dim basePH As Double, capRow As Double, eff As Double, ct As Double, actCav As Double

    ' Crear diccionarios
    Set dictCapEnsamble = CreateObject("Scripting.Dictionary")
    Set dictCapMoldeo = CreateObject("Scripting.Dictionary")
    Set dictUnicos = CreateObject("Scripting.Dictionary")            ' evita duplicar por PartNo|Die
    Set dictBusquedaItemMaster = CreateObject("Scripting.Dictionary")
    Set dictExcluirPartNosZ2 = CreateObject("Scripting.Dictionary")
    Set dictExcluirChildsZ2F = CreateObject("Scripting.Dictionary")

    ' Definir rutas de archivos
    rutaBase = ThisWorkbook.Sheets("Macro").Range("B1").Value
    archivoLF = rutaBase & LOADFACTOR_FILENAME
    archivoItemMaster = rutaBase & "ENSAMBLE_ITEMMASTER.TXT"

    ' === 1. Leer ItemMaster y registrar exclusiones ===
    If Dir(archivoItemMaster) <> "" Then
        Open archivoItemMaster For Binary As #1
            binData = Space$(LOF(1))
            Get #1, , binData
        Close #1

        binData = Replace(binData, vbCrLf, vbLf)
        binData = Replace(binData, vbCr, vbLf)
        lineas = Split(binData, vbLf)

        For i = 1 To UBound(lineas)
            If Trim(lineas(i)) <> "" Then
                datos = Split(lineas(i), "|")
                If UBound(datos) >= 4 Then
                    Dim clavePartNo As String, claveChild As String
                    clavePartNo = QuitarCerosIzquierda(datos(0))
                    claveChild = QuitarCerosIzquierda(datos(0))

                    ' Excluir PartNos con status Z2
                    If UCase(Trim(datos(3))) = "Z2" Then
                        If Not dictExcluirPartNosZ2.exists(clavePartNo) Then
                            dictExcluirPartNosZ2.Add clavePartNo, True
                        End If
                    End If

                    ' Excluir Childs con tipo F y status Z2
                    If UCase(Trim(datos(2))) = "F" And UCase(Trim(datos(3))) = "Z2" Then
                        If Not dictExcluirChildsZ2F.exists(claveChild) Then
                            dictExcluirChildsZ2F.Add claveChild, True
                        End If
                    End If

                    If Trim(datos(4)) <> "" Then
                        dictBusquedaItemMaster(clavePartNo) = Trim(datos(4))
                    End If
                End If
            End If
        Next i
    Else
        MsgBox "No se encontró el archivo: " & archivoItemMaster, vbCritical
        Exit Sub
    End If

    ' === 2. Leer ÚNICO LoadFactor y calcular CAPACIDADES por PartNo ===
    If Dir(archivoLF) <> "" Then
        Open archivoLF For Binary As #1
            binData = Space$(LOF(1))
            Get #1, , binData
        Close #1

        binData = Replace(binData, vbCrLf, vbLf)
        binData = Replace(binData, vbCr, vbLf)
        lineas = Split(binData, vbLf)

        For i = 1 To UBound(lineas)
            If Trim(lineas(i)) <> "" Then
                campos = Split(lineas(i), "|")
                If UBound(campos) >= 20 Then
                    Dim np As String, die As String
                    np = QuitarCerosIzquierda(Trim(campos(0)))
                    die = QuitarCerosIzquierda(Trim(campos(2)))    ' Die completo
                    claveUnica = np & "|" & die

                    If Not dictUnicos.exists(claveUnica) Then
                        dictUnicos.Add claveUnica, ""

                        ' ---- Capacidad por renglón (igual que en Load factor) ----
                        If IsNumeric(campos(11)) Then
                            basePH = Num(campos(11))
                        Else
                            ct = Num(campos(9))         ' Cycle Time
                            actCav = Num(campos(8))     ' Act Cav
                            If actCav <= 0 Then actCav = 1
                            If ct > 0 Then basePH = (3600# / ct) * actCav Else basePH = 0
                        End If
                        eff = EficienciaDeTexto(campos(20))
                        capRow = basePH * eff

                        ' Guardar capacidad acumulada por PartNo (para ensamble y moldeo)
                        dictCapEnsamble(np) = dictCapEnsamble(np) + capRow
                        dictCapMoldeo(np) = dictCapMoldeo(np) + capRow
                    End If
                End If
            End If
        Next i
    Else
        MsgBox "No se encontró el archivo: " & archivoLF, vbCritical
        Exit Sub
    End If

    ' === 3. Procesar BOM ===
    Set wbOrigen = Workbooks.Open("\\Yazaki.local\na\elcom\chihuahua\Area_General\Materiales\Archivos Macro PCD\EP1\Extractor\BOM_SIN_EMPAQUE_FICR4700_1815.xlsx")
    Set wsOrigen = wbOrigen.Sheets(1)
    Set wsDestino = Workbooks(pPlan).Sheets("Capacidades")
    wsDestino.Cells.ClearContents

    wsDestino.Range("A1").Value = "#Parte"
    wsDestino.Range("B1").Value = "Child"
    wsDestino.Range("C1").Value = "Capacidad de Moldeo"
    wsDestino.Range("D1").Value = "Capacidad de Ensamble"

    destinoFila = 2

    For fila = 2 To wsOrigen.Cells(wsOrigen.Rows.Count, "B").End(xlUp).Row
        If wsOrigen.Cells(fila, "I").Value = 1 Then
            partNo = QuitarCerosIzquierda(wsOrigen.Cells(fila, "B").Value)
            child = QuitarCerosIzquierda(wsOrigen.Cells(fila, "C").Value)

            ' Exclusiones combinadas
            If dictExcluirPartNosZ2.exists(partNo) Then GoTo SaltarFila
            If dictExcluirChildsZ2F.exists(child) Then GoTo SaltarFila

            If Left(child, 3) <> "M51" And Left(child, 1) <> "Y" Then
                wsDestino.Cells(destinoFila, "A").Value = partNo
                wsDestino.Cells(destinoFila, "B").Value = child

                ' Moldeo: se consulta por CHILD
                If dictCapMoldeo.exists(child) Then
                    wsDestino.Cells(destinoFila, "C").Value = dictCapMoldeo(child)
                Else
                    wsDestino.Cells(destinoFila, "C").Value = "No encontrado: --"
                End If

                ' Ensamble: se consulta por PART NO (FG)
                If dictCapEnsamble.exists(partNo) Then
                    wsDestino.Cells(destinoFila, "D").Value = dictCapEnsamble(partNo)
                Else
                    wsDestino.Cells(destinoFila, "D").Value = "No encontrado: --"
                End If

                destinoFila = destinoFila + 1
            End If
        End If
SaltarFila:
    Next fila

    wbOrigen.Close SaveChanges:=False

End Sub

' =====================================================================
' FLEX PLAN: copiar desde Excel externo
' =====================================================================
Public Sub CargarArchivoFlexPlanDesdeRuta(rutaBase As String)

    Dim archivoNombre As String
    Dim archivoCompleto As String
    Dim wbOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim hojaDestino As Worksheet
    Dim seleccionManual As Variant

    archivoNombre = "YCC Flex Planning.xlsx"
    archivoCompleto = rutaBase & archivoNombre

    On Error Resume Next
    Set wbOrigen = Workbooks.Open(archivoCompleto, ReadOnly:=True)
    On Error GoTo 0

    ' Si no existe, permitir selección manual
    If wbOrigen Is Nothing Then
        seleccionManual = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona el archivo de Flex Plan")
        If seleccionManual = "False" Then
            MsgBox "No se seleccionó ningún archivo.", vbExclamation
            Exit Sub
        End If

        Set wbOrigen = Workbooks.Open(seleccionManual, ReadOnly:=True)
        If wbOrigen Is Nothing Then
            MsgBox "No se pudo abrir el archivo seleccionado.", vbCritical
            Exit Sub
        End If
    End If

    ' Buscar hoja "Flex-plan"
    On Error Resume Next
    Set hojaOrigen = wbOrigen.Sheets("Flex-plan")
    On Error GoTo 0

    If hojaOrigen Is Nothing Then
        MsgBox "No se encontró la hoja 'Flex-plan' en el archivo seleccionado.", vbExclamation
        wbOrigen.Close SaveChanges:=False
        Exit Sub
    End If

    ' Copiar a hoja local "FlexPlan"
    Set hojaDestino = ThisWorkbook.Sheets("FlexPlan")
    hojaDestino.Cells.ClearContents
    hojaOrigen.UsedRange.Copy Destination:=hojaDestino.Range("A1")

    ' Guardar referencia
    ThisWorkbook.Sheets("Macro").Range("B1").Value = wbOrigen.FullName

    wbOrigen.Close SaveChanges:=False

    MsgBox "Datos cargados correctamente desde 'Flex-plan'.", vbInformation
End Sub

Public Function WorksheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set sht = wb.Sheets(sheetName)
    WorksheetExists = Not sht Is Nothing
    Set sht = Nothing
    On Error GoTo 0
End Function


' === "Vacío" robusto: "", Empty, espacios, tabs, NBSP, CR/LF, "" de fórmula.
Private Function EsVacioRobusto(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then EsVacioRobusto = True: Exit Function
    Dim s As String
    s = CStr(v)
    ' limpiar caracteres molestos
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, Chr$(160), "") ' NBSP
    s = Trim$(s)
    EsVacioRobusto = (LenB(s) = 0)
End Function

' === Última fila real (no se confunde con filtros/formatos)
Private Function ultimaFila(ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                          MatchCase:=False)
    On Error GoTo 0
    ultimaFila = IIf(c Is Nothing, 1, c.Row)
End Function


' === ACTUALIZADOR PRINCIPAL ===
' Actualiza Load Factor (col C) con Short Text (E) de MDMQ0400.
' Coincidencia:  (MDMQ0400.A, MDMQ0400.F)  <->  (LoadFactor.A, LoadFactor.N)
' soloASVacio=True => solo filas donde AS se considere "vacío".
Public Sub ActualizarLoadFactorDesdeMDMQ0400_Fast(ByVal pPlan As String, Optional ByVal soloASVacio As Boolean = True)
    Const COL_MATERIAL& = 1   ' A
    Const COL_SHORTTXT& = 5   ' E
    Const COL_WORKCTR& = 6    ' F
    Const COL_AS& = 45        ' AS

    Dim ruta As String, fullPath As String
    Dim wbM As Workbook, wsM As Worksheet
    Dim wsLF As Worksheet
    Dim rFinM As Long, rFinLF As Long
    Dim arrA As Variant, arrE As Variant, arrF As Variant, arrAS As Variant
    Dim arrLFA As Variant, arrLFN As Variant, arrLFC As Variant
    Dim dict As Object, i As Long, k As String
    Dim actualizados As Long, total As Long, tomados As Long

    ' ---- aceleradores
    Dim prevCalc As XlCalculation, prevScreen As Boolean, prevEvents As Boolean, prevAlerts As Boolean
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo fallo

    ruta = ThisWorkbook.Sheets("Macro").Range("B1").Value
    fullPath = "\\Yazaki.local\na\elcom\chihuahua\Area_General\Materiales\Archivos Macro PCD\EP1\Extractor\MDMQ0400.XLS"
    If Dir(fullPath) = "" Then fullPath = ruta & "MDMQ0400.xlsx"
    If Dir(fullPath) = "" Then Err.Raise vbObjectError + 100, , "No se encontró MDMQ0400.xlsx (revisa Macro!B1 o la ruta fija)."

    ' --- abrir MDMQ0400
    Set wbM = Workbooks.Open(fullPath, ReadOnly:=True)
    Set wsM = wbM.Worksheets(1)

    ' --- hoja Load Factor
    On Error Resume Next
    Set wsLF = Workbooks(pPlan).Worksheets("Load Factor")
    On Error GoTo fallo
    If wsLF Is Nothing Then Err.Raise vbObjectError + 101, , "No existe la hoja 'Load Factor' en " & pPlan

    ' --- construir diccionario desde MDMQ0400 (leyendo columnas a arrays)
    rFinM = ultimaFila(wsM)
    If rFinM < 2 Then GoTo reporte

    arrA = wsM.Range(wsM.Cells(2, COL_MATERIAL), wsM.Cells(rFinM, COL_MATERIAL)).Value2
    arrE = wsM.Range(wsM.Cells(2, COL_SHORTTXT), wsM.Cells(rFinM, COL_SHORTTXT)).Value2
    arrF = wsM.Range(wsM.Cells(2, COL_WORKCTR), wsM.Cells(rFinM, COL_WORKCTR)).Value2
    arrAS = wsM.Range(wsM.Cells(2, COL_AS), wsM.Cells(rFinM, COL_AS)).Value2

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    For i = 1 To UBound(arrA, 1)
        total = total + 1
        If (Not soloASVacio) Or EsVacioRobusto(arrAS(i, 1)) Then
            Dim mat As String, wc As String, st As String
            mat = QuitarCerosIzquierda(CStr(arrA(i, 1)))
            wc = Trim$(CStr(arrF(i, 1)))
            If LenB(mat) > 0 And LenB(wc) > 0 Then
                ' <<< CAMBIO >>> solo primeros 5 caracteres del ShortText
                st = Left$(Trim$(CStr(arrE(i, 1))), 5)
                k = mat & "|" & wc
                dict(k) = st
                tomados = tomados + 1
            End If
        End If
    Next i

    ' --- leer Load Factor y preparar columna C en memoria
    rFinLF = ultimaFila(wsLF)
    If rFinLF < 1 Then GoTo reporte

    arrLFA = wsLF.Range("A1:A" & rFinLF).Value2
    arrLFN = wsLF.Range("N1:N" & rFinLF).Value2
    arrLFC = wsLF.Range("C1:C" & rFinLF).Value2

    For i = 1 To UBound(arrLFA, 1)
        Dim a As String, n As String
        a = Trim$(CStr(arrLFA(i, 1)))
        n = Trim$(CStr(arrLFN(i, 1)))
        If LenB(a) > 0 And LenB(n) > 0 Then
            k = QuitarCerosIzquierda(a) & "|" & n
            If dict.exists(k) Then
                If arrLFC(i, 1) <> dict(k) Then
                    arrLFC(i, 1) = dict(k)
                    actualizados = actualizados + 1
                End If
            End If
        End If
    Next i

    ' --- volcado único a la hoja
    wsLF.Range("C1:C" & rFinLF).Value2 = arrLFC

reporte:

salida:
    On Error Resume Next
    If Not wbM Is Nothing Then wbM.Close SaveChanges:=False
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    Exit Sub

fallo:
    MsgBox "Error al actualizar desde MDMQ0400 (FAST): " & Err.Description, vbCritical
    Resume salida
End Sub



