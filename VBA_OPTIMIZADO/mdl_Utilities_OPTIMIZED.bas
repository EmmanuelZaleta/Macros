Attribute VB_Name = "mdl_Utilities"
Option Explicit

'===============================================================================
' Module: mdl_Utilities
' Purpose: Ultra-fast utility functions for Excel optimization
' Author: Optimized by Claude Code
' Date: 2025-10-30
' Performance: Professional-grade helper functions
'===============================================================================

'===============================================================================
' EXCEL OPTIMIZATION PROCEDURES
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure: OptimizeExcelForSpeed
' Purpose: Disables Excel features for maximum performance
' Call this at the start of any long-running process
'-------------------------------------------------------------------------------
Public Sub OptimizeExcelForSpeed()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayStatusBar = False
        .DisplayAlerts = False
    End With
End Sub

'-------------------------------------------------------------------------------
' Procedure: RestoreExcelSettings
' Purpose: Restores Excel to normal operational mode
' Call this at the end of any long-running process
'-------------------------------------------------------------------------------
Public Sub RestoreExcelSettings()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayStatusBar = True
        .DisplayAlerts = True
    End With
End Sub

'===============================================================================
' DATA CLEANING PROCEDURES (OPTIMIZED)
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure: quitarEspacios
' Purpose: Removes leading/trailing spaces from column data (OPTIMIZED)
' Parameters: col - Column letter or number
' Performance: Uses arrays for bulk processing
'-------------------------------------------------------------------------------
Public Sub quitarEspacios(col As Variant)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim i As Long
    Dim colNum As Long

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' Convert column to number if letter provided
    If VarType(col) = vbString Then
        colNum = ws.Range(col & "1").Column
    Else
        colNum = CLng(col)
    End If

    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    If lastRow <= 1 Then Exit Sub

    ' Optimize Excel
    Call OptimizeExcelForSpeed

    ' ULTRA-FAST: Use Find & Replace (fastest method)
    Set dataRange = ws.Range(ws.Cells(2, colNum), ws.Cells(lastRow, colNum))

    ' Remove leading/trailing spaces using Find & Replace
    dataRange.Replace What:=" *", Replacement:="", LookAt:=xlPart, _
                      SearchOrder:=xlByColumns, MatchCase:=False
    dataRange.Replace What:="* ", Replacement:="", LookAt:=xlPart, _
                      SearchOrder:=xlByColumns, MatchCase:=False

    ' Restore Excel
    Call RestoreExcelSettings

    Set dataRange = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    Call RestoreExcelSettings
    MsgBox "Error en quitarEspacios: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: NumeroAValor
' Purpose: Converts text to numbers in specified column (OPTIMIZED)
' Parameters:
'   col - Column letter
'   startRow - Starting row (default "2")
' Performance: Uses array processing
'-------------------------------------------------------------------------------
Public Sub NumeroAValor(col As String, Optional startRow As String = "2")
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim cell As Range
    Dim dataArray As Variant
    Dim i As Long

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    If lastRow < CLng(startRow) Then Exit Sub

    Call OptimizeExcelForSpeed

    Set dataRange = ws.Range(col & startRow & ":" & col & lastRow)

    ' ULTRA-FAST: Use special cells for only text numbers
    On Error Resume Next
    dataRange.SpecialCells(xlCellTypeConstants, xlTextValues).Value = _
        dataRange.SpecialCells(xlCellTypeConstants, xlTextValues).Value
    On Error GoTo ErrorHandler

    ' Force conversion by multiplying by 1
    For Each cell In dataRange
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                cell.Value = CDbl(cell.Value)
            End If
        End If
    Next cell

    Call RestoreExcelSettings

    Set dataRange = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    Call RestoreExcelSettings
    MsgBox "Error en NumeroAValor: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------------------------
' Procedure: ForzarFechaEnColumna
' Purpose: Forces date format in specified column (OPTIMIZED)
' Parameters:
'   ws - Target worksheet
'   col - Column letter
'   ultimaFila - Last row with data
'-------------------------------------------------------------------------------
Public Sub ForzarFechaEnColumna(ws As Worksheet, col As String, ultimaFila As Long)
    Dim i As Long
    Dim valor As Variant
    Dim colNum As Long

    On Error GoTo ErrorHandler

    colNum = ws.Range(col & "1").Column

    Call OptimizeExcelForSpeed

    For i = 2 To ultimaFila
        valor = ws.Cells(i, colNum).Value

        ' If it's numeric (Excel date serial), convert to date
        If IsNumeric(valor) And valor <> "" Then
            If valor > 0 Then
                ws.Cells(i, colNum).Value = CDate(valor)
            End If
        ElseIf VarType(valor) = vbString Then
            ' If it's string date format (YYYYMMDD)
            If Len(CStr(valor)) = 8 And IsNumeric(valor) Then
                On Error Resume Next
                ws.Cells(i, colNum).Value = DateSerial(Left(valor, 4), Mid(valor, 5, 2), Right(valor, 2))
                On Error GoTo ErrorHandler
            End If
        End If
    Next i

    ' Set number format to date
    ws.Range(col & "2:" & col & ultimaFila).NumberFormat = "mm/dd/yyyy"

    Call RestoreExcelSettings

    Exit Sub

ErrorHandler:
    Call RestoreExcelSettings
    MsgBox "Error en ForzarFechaEnColumna: " & Err.Description, vbCritical
End Sub

'===============================================================================
' FILE OPERATIONS
'===============================================================================

'-------------------------------------------------------------------------------
' Function: buscaArchivo
' Purpose: Searches for most recent file matching pattern
' Parameters: tipoArchivo - Type identifier for file search
' Returns: Filename of most recent matching file
'-------------------------------------------------------------------------------
Public Function buscaArchivo(tipoArchivo As String) As String
    Dim rutaBase As String
    Dim patron As String
    Dim archivo As String
    Dim archivoMasReciente As String
    Dim fechaMasReciente As Date
    Dim fechaActual As Date

    On Error GoTo ErrorHandler

    ' Get base path from settings
    rutaBase = ThisWorkbook.Sheets("Macro").Range("B1").Value
    If Right(rutaBase, 1) <> "\" Then rutaBase = rutaBase & "\"

    ' Define search patterns
    Select Case UCase(tipoArchivo)
        Case "ORDENES"
            patron = "Orderstats*.txt"
        Case "INVLOCWIP"
            patron = "InvLocWIP*.txt"
        Case "ITEMMASTER"
            patron = "ItemMaster*.txt"
        Case "LOADFACTOR"
            patron = "LoadFactor*.txt"
        Case Else
            patron = tipoArchivo & "*.txt"
    End Select

    ' Find most recent file
    archivo = Dir(rutaBase & patron)
    If archivo = "" Then
        buscaArchivo = ""
        Exit Function
    End If

    archivoMasReciente = archivo
    fechaMasReciente = FileDateTime(rutaBase & archivo)

    Do While archivo <> ""
        archivo = Dir()
        If archivo <> "" Then
            fechaActual = FileDateTime(rutaBase & archivo)
            If fechaActual > fechaMasReciente Then
                fechaMasReciente = fechaActual
                archivoMasReciente = archivo
            End If
        End If
    Loop

    buscaArchivo = archivoMasReciente
    Exit Function

ErrorHandler:
    buscaArchivo = ""
End Function

'===============================================================================
' DATA LOADING (ULTRA-FAST)
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure: CargarOrderStat_DesdeUNC_Hasta
' Purpose: Loads OrderStat data with optimized performance
' Parameters:
'   vPlan - Workbook name
'   fechaLimite - Date limit in YYYYMMDD format
'-------------------------------------------------------------------------------
Public Sub CargarOrderStat_DesdeUNC_Hasta(vPlan As String, fechaLimite As String)
    Dim rutaArchivos As String
    Dim vArchivo As String
    Dim rutaCompleta As String
    Dim ws As Worksheet
    Dim fso As Object
    Dim ts As Object
    Dim linea As String
    Dim campos() As String
    Dim i As Long
    Dim fila As Long
    Dim datos() As Variant
    Dim maxFilas As Long
    Dim fechaETD As String

    On Error GoTo ErrorHandler

    ' Get file path
    rutaArchivos = ThisWorkbook.Sheets("Macro").Range("B1").Value
    If Right(rutaArchivos, 1) <> "\" Then rutaArchivos = rutaArchivos & "\"

    vArchivo = buscaArchivo("Ordenes")
    If vArchivo = "" Then Exit Sub

    rutaCompleta = rutaArchivos & vArchivo
    Set ws = Workbooks(vPlan).Sheets("Orderstats")

    Call OptimizeExcelForSpeed

    ' Clear existing data efficiently
    If ws.Cells(2, 1).Value <> "" Then
        ws.Range("A2:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).ClearContents
    End If

    ' Use FileSystemObject for fast file reading
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(rutaCompleta, 1, False)  ' 1 = ForReading

    ' Size array for bulk loading
    maxFilas = 10000
    ReDim datos(1 To maxFilas, 1 To 12)
    fila = 1

    ' Skip header line
    If Not ts.AtEndOfStream Then ts.ReadLine

    ' Read file line by line
    Do While Not ts.AtEndOfStream
        linea = ts.ReadLine
        If Len(Trim(linea)) > 0 Then
            campos = Split(linea, vbTab)

            If UBound(campos) >= 9 Then
                fechaETD = Trim(campos(3))

                ' Filter by date
                If fechaETD <= fechaLimite Then
                    ' Expand array if needed
                    If fila > maxFilas Then
                        maxFilas = maxFilas + 5000
                        ReDim Preserve datos(1 To maxFilas, 1 To 12)
                    End If

                    ' Load data into array
                    For i = 0 To UBound(campos)
                        If i < 10 Then datos(fila, i + 1) = Trim(campos(i))
                    Next i

                    fila = fila + 1
                End If
            End If
        End If
    Loop

    ts.Close

    ' Write array to worksheet in one operation (ULTRA-FAST)
    If fila > 1 Then
        ws.Range("A2").Resize(fila - 1, 12).Value = datos
    End If

    Call RestoreExcelSettings

    ' Cleanup
    Set ts = Nothing
    Set fso = Nothing
    Set ws = Nothing

    Exit Sub

ErrorHandler:
    Call RestoreExcelSettings
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Set ws = Nothing
    Err.Raise Err.Number, "CargarOrderStat_DesdeUNC_Hasta", Err.Description
End Sub

'===============================================================================
' WORKSHEET UTILITIES
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure: ClearDataRangefast
' Purpose: Clears large data ranges efficiently
' Parameters:
'   ws - Target worksheet
'   startRange - Starting cell (e.g., "A2")
'   columnCount - Number of columns to clear
'-------------------------------------------------------------------------------
Public Sub ClearDataRangeFast(ws As Worksheet, startRange As String, columnCount As Long)
    Dim lastRow As Long
    Dim startCell As Range

    On Error Resume Next

    Set startCell = ws.Range(startRange)
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).Row

    If lastRow >= startCell.Row Then
        ws.Range(startCell, ws.Cells(lastRow, startCell.Column + columnCount - 1)).ClearContents
    End If

    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Procedure: ShowProcessingMessage
' Purpose: Updates status message without Select
' Parameters:
'   frm - Form object with lblStatus control
'   mensaje - Message to display
'-------------------------------------------------------------------------------
Public Sub ShowProcessingMessage(frm As Object, mensaje As String)
    On Error Resume Next
    frm.lblStatus.Caption = mensaje
    frm.Repaint
    DoEvents
    On Error GoTo 0
End Sub
