Attribute VB_Name = "TestWebAPI"
Option Explicit

' WebAPIクラスのテスト
' 注意: 実際のHTTP通信はCI環境では行わず、ユーティリティ関数のみテスト

Public Sub Test_UrlEncode_AlphaNumeric()
    Dim api As WebAPI
    Set api = New WebAPI

    If api.UrlEncode("abc123") <> "abc123" Then
        Utils.Fail 3001, "UrlEncode should not change alphanumeric"
    End If
End Sub

Public Sub Test_UrlEncode_Space()
    Dim api As WebAPI
    Set api = New WebAPI

    If api.UrlEncode("hello world") <> "hello+world" Then
        Utils.Fail 3002, "UrlEncode should convert space to +"
    End If
End Sub

Public Sub Test_UrlEncode_SpecialChars()
    Dim api As WebAPI
    Set api = New WebAPI

    ' 予約されていない文字はそのまま
    If api.UrlEncode("test-value_1.0~") <> "test-value_1.0~" Then
        Utils.Fail 3003, "UrlEncode should not change unreserved chars"
    End If
End Sub

Public Sub Test_UrlEncode_Japanese()
    Dim api As WebAPI
    Set api = New WebAPI

    Dim encoded As String
    encoded = api.UrlEncode("&=")

    ' & と = はエンコードされるべき
    If InStr(encoded, "&") > 0 Or InStr(encoded, "=") > 0 Then
        Utils.Fail 3004, "UrlEncode should encode special characters"
    End If
End Sub

Public Sub Test_BuildQueryString_Simple()
    Dim api As WebAPI
    Set api = New WebAPI

    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "name", "john"
    params.Add "age", "30"

    Dim result As String
    result = api.BuildQueryString(params)

    ' nameとageの両方が含まれているか確認
    If InStr(result, "name=john") = 0 Then
        Utils.Fail 3005, "BuildQueryString should contain name=john"
    End If

    If InStr(result, "age=30") = 0 Then
        Utils.Fail 3006, "BuildQueryString should contain age=30"
    End If

    If InStr(result, "&") = 0 Then
        Utils.Fail 3007, "BuildQueryString should contain &"
    End If
End Sub

Public Sub Test_BuildQueryString_WithSpaces()
    Dim api As WebAPI
    Set api = New WebAPI

    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "query", "hello world"

    Dim result As String
    result = api.BuildQueryString(params)

    If InStr(result, "query=hello+world") = 0 Then
        Utils.Fail 3008, "BuildQueryString should encode spaces"
    End If
End Sub

Public Sub Test_DefaultTimeout()
    Dim api As WebAPI
    Set api = New WebAPI

    If api.Timeout <> 30000 Then
        Utils.Fail 3009, "Default timeout should be 30000ms"
    End If
End Sub

Public Sub Test_SetTimeout()
    Dim api As WebAPI
    Set api = New WebAPI

    api.Timeout = 60000

    If api.Timeout <> 60000 Then
        Utils.Fail 3010, "Timeout should be settable"
    End If
End Sub

Public Sub Test_InitialStatus()
    Dim api As WebAPI
    Set api = New WebAPI

    ' 初期状態ではLastStatusは0
    If api.LastStatus <> 0 Then
        Utils.Fail 3011, "Initial LastStatus should be 0"
    End If
End Sub

Public Sub Test_IsSuccess_NoRequest()
    Dim api As WebAPI
    Set api = New WebAPI

    ' リクエスト前はFalse
    If api.IsSuccess() Then
        Utils.Fail 3012, "IsSuccess should be False before any request"
    End If
End Sub
