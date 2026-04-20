Attribute VB_Name = "LibraryDBADO"
' 2025 08 16 - Code review by Copilot/ChatGPT-5 and Mike T

Option Explicit

' Connection string examples:
' "Initial Catalog=" & sDatabase & ";Data Source=" & sServer & ";Integrated Security=SSPI;"
' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=NewTest;Data Source=."

Public db_Records As ADODB.Recordset
Private FConn As ADODB.Connection
Private FConnected As Boolean
Private FServer As String
Private FDatabase As String
Private FConnectionString As String

' Connect using an ODBC DSN
Public Sub ConnectToODBC(ADSN As String, AUser As String, APassword As String)
    FConnectionString = "DSN=" & ADSN & ";UID=" & AUser & ";PWD=" & APassword
    ConnectToDB
End Sub

' Connect using SQLOLEDB provider
Public Sub ConnectToSQLOLEDB(AServer As String, ADatabase As String, sUser As String, sPassword As String)
    FServer = AServer
    FDatabase = ADatabase

    If sUser = "" Then
        FConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                            "Initial Catalog=" & FDatabase & ";Data Source=" & FServer
    Else
        FConnectionString = "Provider=SQLOLEDB.1;User ID='" & sUser & "';Password='" & sPassword & "';" & _
                            "Persist Security Info=False;Initial Catalog=" & FDatabase & ";Data Source=" & FServer
    End If

    ConnectToDB
End Sub

' Establish connection using the current connection string
Public Sub ConnectToDB()
    On Error GoTo ErrHandler

    If ConnectedToDb() Then CloseConnection

    Set FConn = New ADODB.Connection
    FConn.Open FConnectionString

    Set db_Records = Nothing
    FConnected = (FConn.State = adStateOpen)
    Exit Sub

ErrHandler:
    FConnected = False
    MsgBox "Failed to connect to database: " & Err.Description, vbCritical
End Sub

' Close the active connection
Public Sub CloseConnection()
    On Error GoTo ErrHandler

    If Not FConn Is Nothing Then
        If FConn.State = adStateOpen Then FConn.Close
    End If

    Set FConn = Nothing
    FConnected = False
    Exit Sub

ErrHandler:
    MsgBox "Error closing connection: " & Err.Description, vbExclamation
End Sub

' Check if connection is active
Public Function ConnectedToDb() As Boolean
    ConnectedToDb = (Not FConn Is Nothing) And (FConn.State = adStateOpen) And FConnected
End Function

' Execute a query and store results in db_Records
Public Sub RunQuery(sQuery As String)
    On Error GoTo ErrHandler

    If Not ConnectedToDb() Then
        MsgBox "No active database connection.", vbExclamation
        Exit Sub
    End If

    Set db_Records = FConn.Execute(sQuery)
    Exit Sub

ErrHandler:
    MsgBox "Query failed: " & Err.Description, vbCritical
    Set db_Records = Nothing
End Sub

' Clear the current recordset
Public Sub ClearQuery()
    If Not db_Records Is Nothing Then
        If db_Records.State = adStateOpen Then db_Records.Close
        Set db_Records = Nothing
    End If
End Sub

' Return a single value from a query
Public Function QuickValue(sQuery As String, sReturnField As String) As Variant
    On Error GoTo ErrHandler

    RunQuery sQuery

    If db_Records Is Nothing Or db_Records.EOF Then
        QuickValue = Null
    Else
        db_Records.MoveFirst
        QuickValue = db_Records.fields(sReturnField).Value
    End If
    Exit Function

ErrHandler:
    MsgBox "QuickValue failed: " & Err.Description, vbExclamation
    QuickValue = Null
End Function

' Return a value only if exactly one record is returned
Public Function QuickValueSingle(sQuery As String, sReturnField As String) As Variant
    On Error GoTo ErrHandler

    RunQuery sQuery

    If db_Records Is Nothing Or db_Records.EOF Then
        QuickValueSingle = Null
        Exit Function
    End If

    Dim iCount As Integer: iCount = 0
    db_Records.MoveFirst

    Do While Not db_Records.EOF
        iCount = iCount + 1
        db_Records.MoveNext
    Loop

    db_Records.MoveFirst
    If iCount = 1 Then
        QuickValueSingle = db_Records.fields(sReturnField).Value
    Else
        QuickValueSingle = "XXX" ' Consider replacing with Null or raising an error
    End If
    Exit Function

ErrHandler:
    MsgBox "QuickValueSingle failed: " & Err.Description, vbExclamation
    QuickValueSingle = Null
End Function
