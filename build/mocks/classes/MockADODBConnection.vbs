' Mock ADODB.Connection for VBS testing
Class MockADODBConnection
    Private connectionString_
    Private state_
    Private errors_

    Private Sub Class_Initialize()
        connectionString_ = ""
        state_ = 0 ' adStateClosed
        Set errors_ = CreateObject("Scripting.Dictionary")
    End Sub

    Public Property Get ConnectionString()
        ConnectionString = connectionString_
    End Property

    Public Property Let ConnectionString(ByVal value)
        connectionString_ = value
    End Property

    Public Property Get State()
        State = state_
    End Property

    Public Property Get Errors()
        Set Errors = errors_
    End Property

    Public Sub Open(ByVal connStr)
        If connStr <> "" Then connectionString_ = connStr
        state_ = 1 ' adStateOpen
        DebugPrint "[MockADODB] Connection.Open: " & connectionString_
    End Sub

    Public Sub Close()
        state_ = 0 ' adStateClosed
        DebugPrint "[MockADODB] Connection.Close"
    End Sub

    Public Function Execute(ByVal commandText)
        DebugPrint "[MockADODB] Connection.Execute: " & commandText
        Set Execute = New MockADODBRecordset
    End Function

    Public Function BeginTrans()
        DebugPrint "[MockADODB] BeginTrans"
        BeginTrans = 1
    End Function

    Public Sub CommitTrans()
        DebugPrint "[MockADODB] CommitTrans"
    End Sub

    Public Sub RollbackTrans()
        DebugPrint "[MockADODB] RollbackTrans"
    End Sub
End Class

Function CreateMockADODBConnection()
    Set CreateMockADODBConnection = New MockADODBConnection
End Function
