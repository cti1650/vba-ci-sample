VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' WebAPI - HTTP通信を行うクラス
' GET/POST/PUT/DELETE メソッドをサポート

Private http_ As Object
Private lastStatus_ As Long
Private lastStatusText_ As String
Private lastResponseText_ As String
Private timeout_ As Long

Private Sub Class_Initialize()
    timeout_ = 30000 ' デフォルト30秒
End Sub

Private Sub Class_Terminate()
    Set http_ = Nothing
End Sub

' ===========================================
' プロパティ
' ===========================================

Public Property Get LastStatus() As Long
    LastStatus = lastStatus_
End Property

Public Property Get LastStatusText() As String
    LastStatusText = lastStatusText_
End Property

Public Property Get LastResponseText() As String
    LastResponseText = lastResponseText_
End Property

Public Property Get Timeout() As Long
    Timeout = timeout_
End Property

Public Property Let Timeout(ByVal value As Long)
    timeout_ = value
End Property

' ===========================================
' HTTP メソッド
' ===========================================

Public Function Fetch(ByVal URL As String, Optional ByVal headers As Object = Nothing) As String
    Fetch = Request("GET", URL, "", headers)
End Function

Public Function Send(ByVal URL As String, ByVal body As String, Optional ByVal headers As Object = Nothing) As String
    Send = Request("POST", URL, body, headers)
End Function

Public Function Update(ByVal URL As String, ByVal body As String, Optional ByVal headers As Object = Nothing) As String
    Update = Request("PUT", URL, body, headers)
End Function

Public Function Remove(ByVal URL As String, Optional ByVal headers As Object = Nothing) As String
    Remove = Request("DELETE", URL, "", headers)
End Function

' ===========================================
' メインリクエスト処理
' ===========================================

Public Function Request(ByVal method As String, ByVal URL As String, Optional ByVal body As String = "", Optional ByVal headers As Object = Nothing) As String
    On Error Resume Next

    ' HTTPオブジェクト作成
    Set http_ = CreateHttpRequest()
    If http_ Is Nothing Then
        lastStatus_ = 0
        lastStatusText_ = "Failed to create HTTP object"
        lastResponseText_ = ""
        Request = ""
        Exit Function
    End If

    ' リクエスト設定
    http_.Open method, URL, False
    http_.setRequestHeader "User-Agent", "VBA-WebAPI/1.0"

    ' カスタムヘッダー設定
    If Not headers Is Nothing Then
        Dim key As Variant
        For Each key In headers.Keys
            http_.setRequestHeader CStr(key), CStr(headers(key))
        Next
    End If

    ' Content-Type設定（POSTの場合）
    If method = "POST" Or method = "PUT" Then
        If headers Is Nothing Then
            http_.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
        ElseIf Not headers.Exists("Content-Type") Then
            http_.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
        End If
    End If

    ' リクエスト送信
    If body <> "" Then
        http_.send body
    Else
        http_.send
    End If

    ' エラーチェック
    If Err.Number <> 0 Then
        lastStatus_ = 0
        lastStatusText_ = "Error: " & Err.Description
        lastResponseText_ = ""
        Request = ""
        Set http_ = Nothing
        On Error GoTo 0
        Exit Function
    End If

    ' レスポンス取得
    lastStatus_ = http_.Status
    lastStatusText_ = http_.statusText
    lastResponseText_ = http_.responseText
    Request = lastResponseText_

    Set http_ = Nothing
    On Error GoTo 0
End Function

' ===========================================
' JSON ヘルパー
' ===========================================

Public Function PostJson(ByVal URL As String, ByVal jsonBody As String) As String
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "Content-Type", "application/json; charset=UTF-8"
    headers.Add "Accept", "application/json"
    PostJson = Request("POST", URL, jsonBody, headers)
End Function

Public Function GetJson(ByVal URL As String) As String
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers.Add "Accept", "application/json"
    GetJson = Request("GET", URL, "", headers)
End Function

' ===========================================
' ユーティリティ
' ===========================================

Public Function IsSuccess() As Boolean
    IsSuccess = (lastStatus_ >= 200 And lastStatus_ < 300)
End Function

Public Function UrlEncode(ByVal text As String) As String
    Dim i As Long
    Dim c As String
    Dim result As String

    result = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        Select Case c
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", ".", "~"
                result = result & c
            Case " "
                result = result & "+"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next
    UrlEncode = result
End Function

Public Function BuildQueryString(ByVal params As Object) As String
    Dim key As Variant
    Dim parts() As String
    Dim i As Long

    ReDim parts(params.Count - 1)
    i = 0
    For Each key In params.Keys
        parts(i) = UrlEncode(CStr(key)) & "=" & UrlEncode(CStr(params(key)))
        i = i + 1
    Next
    BuildQueryString = Join(parts, "&")
End Function

' ===========================================
' 内部関数
' ===========================================

Private Function CreateHttpRequest() As Object
    Dim progIDs As Variant
    Dim ret As Object
    Dim i As Long

    Set ret = Nothing
    progIDs = Array("MSXML2.ServerXMLHTTP.6.0", _
                    "MSXML2.ServerXMLHTTP", _
                    "Msxml2.XMLHTTP.6.0", _
                    "Msxml2.XMLHTTP", _
                    "Microsoft.XMLHTTP")

    On Error Resume Next
    For i = LBound(progIDs) To UBound(progIDs)
        Set ret = CreateObject(progIDs(i))
        If Not ret Is Nothing Then Exit For
    Next
    On Error GoTo 0

    Set CreateHttpRequest = ret
End Function
