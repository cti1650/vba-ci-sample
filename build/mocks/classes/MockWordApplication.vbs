' Mock Word.Application for VBS testing
Class MockWordApplication
    Private documents_
    Private visible_
    Private displayAlerts_

    Private Sub Class_Initialize()
        Set documents_ = CreateObject("Scripting.Dictionary")
        visible_ = False
        displayAlerts_ = 0
    End Sub

    Public Property Get Visible()
        Visible = visible_
    End Property

    Public Property Let Visible(ByVal value)
        visible_ = value
    End Property

    Public Property Get DisplayAlerts()
        DisplayAlerts = displayAlerts_
    End Property

    Public Property Let DisplayAlerts(ByVal value)
        displayAlerts_ = value
    End Property

    Public Property Get Documents()
        Set Documents = documents_
    End Property

    Public Function Quit()
        DebugPrint "[MockWord] Application.Quit called"
    End Function

    Public Property Get Version()
        Version = "16.0"
    End Property

    Public Property Get Name()
        Name = "Microsoft Word"
    End Property
End Class

Function CreateMockWordApplication()
    Set CreateMockWordApplication = New MockWordApplication
End Function
