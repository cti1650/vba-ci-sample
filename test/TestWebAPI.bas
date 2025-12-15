Attribute VB_Name = "TestWebAPI"
Option Explicit

' WebAPIクラスのテスト
' 注意: 実際のHTTP通信はCI環境では行わず、ユーティリティ関数のみテスト

Public Sub Test_UrlEncode_AlphaNumeric()
    Dim api As WebAPI
    Set api = New WebAPI

    expect(api.UrlEncode("abc123")).toBe "abc123"
    expect(api.UrlEncode("ABC")).toBe "ABC"
    expect(api.UrlEncode("test")).toBe "test"
End Sub

Public Sub Test_UrlEncode_Space()
    Dim api As WebAPI
    Set api = New WebAPI

    expect(api.UrlEncode("hello world")).toBe "hello+world"
    expect(api.UrlEncode(" ")).toBe "+"
End Sub

Public Sub Test_UrlEncode_SpecialChars()
    Dim api As WebAPI
    Set api = New WebAPI

    ' 予約されていない文字はそのまま
    expect(api.UrlEncode("test-value_1.0~")).toBe "test-value_1.0~"
    expect(api.UrlEncode("-_.~")).toBe "-_.~"
End Sub

Public Sub Test_UrlEncode_Reserved()
    Dim api As WebAPI
    Set api = New WebAPI

    ' 予約文字はエンコードされる
    Dim encoded As String
    encoded = api.UrlEncode("&=")

    expect(encoded).Not_.toContain "&"
    expect(encoded).Not_.toContain "="
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

    expect(result).toContain "name=john"
    expect(result).toContain "age=30"
    expect(result).toContain "&"
End Sub

Public Sub Test_BuildQueryString_WithSpaces()
    Dim api As WebAPI
    Set api = New WebAPI

    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "query", "hello world"

    Dim result As String
    result = api.BuildQueryString(params)

    expect(result).toContain "query=hello+world"
End Sub

Public Sub Test_DefaultTimeout()
    Dim api As WebAPI
    Set api = New WebAPI

    expect(api.Timeout).toBe 30000
End Sub

Public Sub Test_SetTimeout()
    Dim api As WebAPI
    Set api = New WebAPI

    api.Timeout = 60000
    expect(api.Timeout).toBe 60000

    api.Timeout = 5000
    expect(api.Timeout).toBe 5000
End Sub

Public Sub Test_InitialStatus()
    Dim api As WebAPI
    Set api = New WebAPI

    expect(api.LastStatus).toBe 0
End Sub

Public Sub Test_IsSuccess_NoRequest()
    Dim api As WebAPI
    Set api = New WebAPI

    expect(api.IsSuccess()).toBeFalsy
End Sub
