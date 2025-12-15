' Mock Excel.Application for VBS testing
Class MockExcelApplication
    Private worksheets_
    Private workbooks_
    Private visible_
    Private displayAlerts_
    Private screenUpdating_
    Private calculation_

    Private Sub Class_Initialize()
        Set worksheets_ = CreateObject("Scripting.Dictionary")
        Set workbooks_ = CreateObject("Scripting.Dictionary")
        visible_ = False
        displayAlerts_ = True
        screenUpdating_ = True
        calculation_ = -4105 ' xlCalculationAutomatic
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

    Public Property Get ScreenUpdating()
        ScreenUpdating = screenUpdating_
    End Property

    Public Property Let ScreenUpdating(ByVal value)
        screenUpdating_ = value
    End Property

    Public Property Get Calculation()
        Calculation = calculation_
    End Property

    Public Property Let Calculation(ByVal value)
        calculation_ = value
    End Property

    Public Property Get Workbooks()
        Set Workbooks = workbooks_
    End Property

    Public Function Quit()
        DebugPrint "[MockExcel] Application.Quit called"
    End Function

    Public Function Run(ByVal macroName)
        DebugPrint "[MockExcel] Application.Run: " & macroName
        Run = Empty
    End Function

    Public Property Get Version()
        Version = "16.0"
    End Property

    Public Property Get Name()
        Name = "Microsoft Excel"
    End Property
End Class

Function CreateMockExcelApplication()
    Set CreateMockExcelApplication = New MockExcelApplication
End Function
