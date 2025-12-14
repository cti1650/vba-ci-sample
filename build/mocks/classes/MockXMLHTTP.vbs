' Mock XMLHTTP / ServerXMLHTTP for VBS testing
Class MockXMLHTTP
    Private status_
    Private statusText_
    Private responseText_
    Private responseXML_
    Private readyState_
    Private requestMethod_
    Private requestUrl_
    Private requestHeaders_

    Private Sub Class_Initialize()
        status_ = 200
        statusText_ = "OK"
        responseText_ = ""
        Set responseXML_ = Nothing
        readyState_ = 0
        requestMethod_ = ""
        requestUrl_ = ""
        Set requestHeaders_ = CreateObject("Scripting.Dictionary")
    End Sub

    Public Property Get Status()
        Status = status_
    End Property

    Public Property Get statusText()
        statusText = statusText_
    End Property

    Public Property Get responseText()
        responseText = responseText_
    End Property

    Public Property Get responseXML()
        Set responseXML = responseXML_
    End Property

    Public Property Get readyState()
        readyState = readyState_
    End Property

    Public Sub Open(ByVal method, ByVal url, ByVal async)
        requestMethod_ = method
        requestUrl_ = url
        readyState_ = 1
        DebugPrint "[MockXMLHTTP] Open: " & method & " " & url
    End Sub

    Public Sub setRequestHeader(ByVal header, ByVal value)
        requestHeaders_(header) = value
    End Sub

    Public Sub send(ByVal body)
        readyState_ = 4
        DebugPrint "[MockXMLHTTP] Send: " & requestUrl_
        If body <> "" Then
            DebugPrint "[MockXMLHTTP] Body: " & Left(body, 100)
        End If
        ' Default empty response
        responseText_ = "{}"
    End Sub

    Public Function getResponseHeader(ByVal header)
        getResponseHeader = ""
    End Function

    Public Function getAllResponseHeaders()
        getAllResponseHeaders = ""
    End Function

    ' Test helper: Set mock response
    Public Sub SetMockResponse(ByVal statusCode, ByVal text)
        status_ = statusCode
        responseText_ = text
        If statusCode >= 200 And statusCode < 300 Then
            statusText_ = "OK"
        Else
            statusText_ = "Error"
        End If
    End Sub
End Class

Function CreateMockXMLHTTP()
    Set CreateMockXMLHTTP = New MockXMLHTTP
End Function
