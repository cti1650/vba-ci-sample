' CreateObject mock/wrapper for testing
' Allows intercepting and replacing COM object creation with mock objects

' Mock object registry
Dim mock_Objects
Set mock_Objects = CreateObject("Scripting.Dictionary")

' Register a mock object for a ProgID
Sub RegisterMockObject(ByVal progId, ByVal mockObj)
    mock_Objects(LCase(progId)) = mockObj
End Sub

' Unregister a mock object
Sub UnregisterMockObject(ByVal progId)
    If mock_Objects.Exists(LCase(progId)) Then
        mock_Objects.Remove LCase(progId)
    End If
End Sub

' Clear all registered mocks
Sub ClearAllMocks()
    mock_Objects.RemoveAll
End Sub

' CreateObject wrapper (mock-aware)
Function CreateObjectMock(ByVal progId)
    Dim lowerProgId
    lowerProgId = LCase(progId)

    ' Return registered mock if available
    If mock_Objects.Exists(lowerProgId) Then
        Set CreateObjectMock = mock_Objects(lowerProgId)
        Exit Function
    End If

    ' Otherwise create real object
    Set CreateObjectMock = CreateObject(progId)
End Function
