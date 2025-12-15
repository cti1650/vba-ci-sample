' Mock ADODB.Recordset for VBS testing
Class MockADODBRecordset
    Private fields_
    Private data_
    Private currentRow_
    Private eof_
    Private bof_

    Private Sub Class_Initialize()
        Set fields_ = CreateObject("Scripting.Dictionary")
        Set data_ = CreateObject("Scripting.Dictionary")
        currentRow_ = -1
        eof_ = True
        bof_ = True
    End Sub

    Public Property Get EOF()
        EOF = eof_
    End Property

    Public Property Get BOF()
        BOF = bof_
    End Property

    Public Property Get Fields()
        Set Fields = fields_
    End Property

    Public Property Get RecordCount()
        RecordCount = data_.Count
    End Property

    Public Sub Open(ByVal source, ByVal conn)
        DebugPrint "[MockADODB] Recordset.Open: " & source
        eof_ = True
        bof_ = True
    End Sub

    Public Sub Close()
        DebugPrint "[MockADODB] Recordset.Close"
    End Sub

    Public Sub MoveFirst()
        If data_.Count > 0 Then
            currentRow_ = 0
            eof_ = False
            bof_ = False
        End If
    End Sub

    Public Sub MoveNext()
        currentRow_ = currentRow_ + 1
        If currentRow_ >= data_.Count Then
            eof_ = True
        End If
    End Sub

    Public Sub MoveLast()
        If data_.Count > 0 Then
            currentRow_ = data_.Count - 1
            eof_ = False
            bof_ = False
        End If
    End Sub

    Public Sub AddNew()
        DebugPrint "[MockADODB] Recordset.AddNew"
    End Sub

    Public Sub Update()
        DebugPrint "[MockADODB] Recordset.Update"
    End Sub

    Public Sub Delete()
        DebugPrint "[MockADODB] Recordset.Delete"
    End Sub
End Class

Function CreateMockADODBRecordset()
    Set CreateMockADODBRecordset = New MockADODBRecordset
End Function
