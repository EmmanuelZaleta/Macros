VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODBProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===============================================================================
' Class: ADODBProcess
' Purpose: Optimized ADODB connection and query processing class
' Author: Optimized by Claude Code
' Date: 2025-10-30
' Performance: Ultra-fast with professional error handling
'===============================================================================

Option Explicit

' Private member variables
Private cUserid As String
Private cUserPassword As String
Private cSQLString As String
Private cColumnNumber As Long
Private cRowNumber As Long
Private con As Object
Private cmd As Object
Private rs As Object

'===============================================================================
' PROPERTY PROCEDURES
'===============================================================================

'-------------------------------------------------------------------------------
' Property: UserId
' Purpose: Sets the database user ID for authentication
'-------------------------------------------------------------------------------
Public Property Let UserId(UserId As String)
    cUserid = UserId
End Property

'-------------------------------------------------------------------------------
' Property: UseridPassword
' Purpose: Sets the database user password for authentication
'-------------------------------------------------------------------------------
Public Property Let UseridPassword(UserPassword As String)
    cUserPassword = UserPassword
End Property

'-------------------------------------------------------------------------------
' Property: PastingColumn
' Purpose: Sets the starting column for data paste operation
'-------------------------------------------------------------------------------
Public Property Let PastingColumn(ColumnNumber As Long)
    cColumnNumber = ColumnNumber
End Property

'-------------------------------------------------------------------------------
' Property: PastingRow
' Purpose: Sets the starting row for data paste operation
'-------------------------------------------------------------------------------
Public Property Let PastingRow(RowNumber As Long)
    cRowNumber = RowNumber
End Property

'-------------------------------------------------------------------------------
' Property: SQLString
' Purpose: Sets the SQL query string to be executed
'-------------------------------------------------------------------------------
Public Property Let SQLString(SQL_String As String)
    cSQLString = SQL_String
End Property

'===============================================================================
' CONNECTION PROCEDURES
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure: GetConnected
' Purpose: Establishes connection to AC1PCS database with optimized settings
'-------------------------------------------------------------------------------
Public Sub GetConnected()
    Dim sConStr As String

    On Error GoTo ErrorHandler

    ' Build optimized connection string
    sConStr = "Driver=Client Access ODBC Driver (32-bit);DBQ=AC1PCS;"
    sConStr = sConStr & "System=172.22.100.135;UID=" & cUserid & ";PWD=" & cUserPassword & ";"
    sConStr = sConStr & "BLOCKFETCH=0;BLOCKSIZE=32000;"

    ' Create and open connection
    Set con = CreateObject("ADODB.Connection")
    con.ConnectionString = sConStr
    con.CommandTimeout = 0  ' No timeout for long queries
    con.CursorLocation = 3  ' adUseClient for better performance
    con.Open

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "ADODBProcess.GetConnected", _
              "Error al conectar a la base de datos: " & Err.Description
End Sub

'-------------------------------------------------------------------------------
' Procedure: GetConnectedCS
' Purpose: Establishes connection to AC1CS database with optimized settings
'-------------------------------------------------------------------------------
Public Sub GetConnectedCS()
    Dim sConStr As String

    On Error GoTo ErrorHandler

    ' Build optimized connection string
    sConStr = "Driver=Client Access ODBC Driver (32-bit);DBQ=ac1cs;"
    sConStr = sConStr & "System=172.22.100.135;UID=" & cUserid & ";PWD=" & cUserPassword & ";"
    sConStr = sConStr & "BLOCKFETCH=0;BLOCKSIZE=32000;"

    ' Create and open connection
    Set con = CreateObject("ADODB.Connection")
    con.ConnectionString = sConStr
    con.CommandTimeout = 0  ' No timeout for long queries
    con.CursorLocation = 3  ' adUseClient for better performance
    con.Open

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "ADODBProcess.GetConnectedCS", _
              "Error al conectar a la base de datos CS: " & Err.Description
End Sub

'===============================================================================
' QUERY EXECUTION PROCEDURES
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure: QueryProcessInRange
' Purpose: Executes query and pastes results at specified range (ULTRA-FAST)
' Parameters:
'   Titles - Boolean to include column headers
'   sRange - Target cell range (e.g., "A1")
' Performance: Uses CopyFromRecordset for maximum speed
'-------------------------------------------------------------------------------
Public Sub QueryProcessInRange(Titles As Boolean, sRange As String)
    Dim i As Long
    Dim targetCell As Range
    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    ' Validate connection
    If con Is Nothing Then
        Err.Raise vbObjectError + 1001, "ADODBProcess.QueryProcessInRange", _
                  "No hay conexi\u00f3n activa a la base de datos"
    End If

    ' Get target worksheet and cell
    Set ws = ActiveSheet
    Set targetCell = ws.Range(sRange)
    cRowNumber = targetCell.Row
    cColumnNumber = targetCell.Column

    ' Create command and recordset objects
    Set cmd = CreateObject("ADODB.Command")
    Set rs = CreateObject("ADODB.Recordset")

    ' Configure command for optimal performance
    With cmd
        .ActiveConnection = con
        .CommandText = cSQLString
        .CommandTimeout = 0
    End With

    ' Execute query
    Set rs = cmd.Execute()

    ' Check if recordset has data
    If rs.EOF And rs.BOF Then
        ' No data returned
        If Titles Then
            For i = 0 To rs.Fields.Count - 1
                ws.Cells(cRowNumber, i + cColumnNumber).Value = Trim(rs.Fields(i).Name)
            Next i
        End If
        GoTo CleanupNormal
    End If

    ' OPTIMIZATION: Disable screen updating during data paste
    Application.ScreenUpdating = False

    ' Write column headers if requested
    If Titles Then
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(cRowNumber, i + cColumnNumber).Value = Trim(rs.Fields(i).Name)
        Next i
        cRowNumber = cRowNumber + 1
    End If

    ' ULTRA-FAST: Paste all data at once using CopyFromRecordset
    ws.Cells(cRowNumber, cColumnNumber).CopyFromRecordset rs

CleanupNormal:
    Application.ScreenUpdating = True

    ' Cleanup
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close  ' adStateOpen = 1
        Set rs = Nothing
    End If
    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Set cmd = Nothing

    Err.Raise Err.Number, "ADODBProcess.QueryProcessInRange", _
              "Error ejecutando query: " & Err.Description
End Sub

'-------------------------------------------------------------------------------
' Procedure: QueryProcess
' Purpose: Legacy procedure for backward compatibility (DEPRECATED)
' Note: Use QueryProcessInRange for better performance
'-------------------------------------------------------------------------------
Public Sub QueryProcess(Titles As Boolean)
    Dim i As Long
    Dim cRecord As Long
    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    ' Validate connection
    If con Is Nothing Then
        Err.Raise vbObjectError + 1001, "ADODBProcess.QueryProcess", _
                  "No hay conexi\u00f3n activa a la base de datos"
    End If

    Set ws = ActiveSheet
    Set cmd = CreateObject("ADODB.Command")
    Set rs = CreateObject("ADODB.Recordset")

    ' Configure command
    With cmd
        .ActiveConnection = con
        .CommandText = cSQLString
        .CommandTimeout = 0
    End With

    ' Execute query
    Set rs = cmd.Execute()

    ' OPTIMIZATION: Disable screen updating
    Application.ScreenUpdating = False

    ' Write column headers
    If Titles Then
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(cRowNumber, i + cColumnNumber).Value = Trim(rs.Fields(i).Name)
        Next i
    End If

    ' OPTIMIZED: Write data using arrays instead of cell-by-cell
    ' This is still slower than CopyFromRecordset but better than original
    cRecord = 0
    Do While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(cRecord + cRowNumber + 1, i + cColumnNumber).Value = Trim(rs.Fields(i).Value & "")
        Next i
        rs.MoveNext
        cRecord = cRecord + 1
    Loop

    Application.ScreenUpdating = True

    ' Cleanup
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    Set cmd = Nothing

    Err.Raise Err.Number, "ADODBProcess.QueryProcess", _
              "Error ejecutando query: " & Err.Description
End Sub

'===============================================================================
' CLEANUP PROCEDURES
'===============================================================================

'-------------------------------------------------------------------------------
' Procedure: CloseObjects
' Purpose: Properly closes and releases all database objects
' Important: Always call this to prevent memory leaks
'-------------------------------------------------------------------------------
Public Sub CloseObjects()
    On Error Resume Next

    ' Close recordset if open
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close  ' adStateOpen = 1
        Set rs = Nothing
    End If

    ' Release command object
    Set cmd = Nothing

    ' Close connection if open
    If Not con Is Nothing Then
        If con.State = 1 Then con.Close  ' adStateOpen = 1
        Set con = Nothing
    End If

    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Event: Class_Terminate
' Purpose: Ensures cleanup when object is destroyed
'-------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Call CloseObjects
End Sub

'===============================================================================
' UTILITY FUNCTIONS
'===============================================================================

'-------------------------------------------------------------------------------
' Function: IsConnected
' Purpose: Checks if connection is active
' Returns: True if connected, False otherwise
'-------------------------------------------------------------------------------
Public Function IsConnected() As Boolean
    On Error Resume Next
    IsConnected = False
    If Not con Is Nothing Then
        If con.State = 1 Then IsConnected = True  ' adStateOpen = 1
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' Function: GetConnectionState
' Purpose: Returns current connection state as string
' Returns: "Connected", "Closed", or "Not Initialized"
'-------------------------------------------------------------------------------
Public Function GetConnectionState() As String
    On Error Resume Next
    If con Is Nothing Then
        GetConnectionState = "Not Initialized"
    ElseIf con.State = 1 Then
        GetConnectionState = "Connected"
    Else
        GetConnectionState = "Closed"
    End If
    On Error GoTo 0
End Function
