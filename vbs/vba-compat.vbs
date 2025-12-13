' vba-compat.vbs - VBA固有の機能をVBSでモックする互換レイヤー
Option Explicit

' ============================================
' Debug オブジェクトのモック
' ============================================
' VBSでは Print が予約語のため、Debug.Print は直接モックできない
' 代わりに変換スクリプトで Debug.Print を DebugPrint に変換する
Sub DebugPrint(ByVal msg)
    WScript.Echo msg
End Sub

' ============================================
' VBA関数のモック
' ============================================

' FreeFile - VBSでは使えないが、モックとして0を返す
Function FreeFile()
    FreeFile = 0
End Function

' DoEvents - VBSでは何もしない
Sub DoEvents()
    ' No-op
End Sub

' ============================================
' Collection クラス (VBSにはないのでモック)
' ============================================
Class CollectionClass
    Private items_

    Private Sub Class_Initialize()
        Set items_ = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub Add(ByVal item, ByVal key)
        If key = "" Then
            key = "item_" & (items_.Count + 1)
        End If
        items_.Add key, item
    End Sub

    Public Sub Remove(ByVal index)
        If IsNumeric(index) Then
            Dim i, k
            i = 0
            For Each k In items_.Keys
                i = i + 1
                If i = index Then
                    items_.Remove k
                    Exit Sub
                End If
            Next
        Else
            items_.Remove index
        End If
    End Sub

    Public Property Get Item(ByVal index)
        If IsNumeric(index) Then
            Dim i, k
            i = 0
            For Each k In items_.Keys
                i = i + 1
                If i = index Then
                    If IsObject(items_(k)) Then
                        Set Item = items_(k)
                    Else
                        Item = items_(k)
                    End If
                    Exit Property
                End If
            Next
        Else
            If IsObject(items_(index)) Then
                Set Item = items_(index)
            Else
                Item = items_(index)
            End If
        End If
    End Property

    Public Property Get Count()
        Count = items_.Count
    End Property

    Public Sub Clear()
        items_.RemoveAll
    End Sub
End Class

' Collection のファクトリ関数
Function CreateCollection()
    Set CreateCollection = New CollectionClass
End Function

' ============================================
' FileSystemObject ヘルパー (よく使う関数)
' ============================================
Dim fso_
Set fso_ = CreateObject("Scripting.FileSystemObject")

' VBA互換のファイル書き込み
Sub WriteTextFile(ByVal path, ByVal text)
    Dim f
    Set f = fso_.CreateTextFile(path, True)
    f.WriteLine text
    f.Close
End Sub

' VBA互換のファイル読み込み
Function ReadTextFile(ByVal path)
    If fso_.FileExists(path) Then
        Dim f
        Set f = fso_.OpenTextFile(path, 1)
        ReadTextFile = f.ReadAll
        f.Close
    Else
        ReadTextFile = ""
    End If
End Function

' ============================================
' 型変換関数 (VBSにもあるが念のため)
' ============================================
' CLng, CStr, CInt, CDbl, CBool, CDate は VBS標準で使える

' CLngPtr - VBAのポインタ型、VBSではCLngにフォールバック
Function CLngPtr(ByVal value)
    CLngPtr = CLng(value)
End Function

' ============================================
' テスト用アサーション関数
' ============================================
Sub Assert(ByVal condition, ByVal message)
    If Not condition Then
        Err.Raise 9999, "Assert", "Assertion failed: " & message
    End If
End Sub

Sub AssertEqual(ByVal expected, ByVal actual, ByVal message)
    If expected <> actual Then
        Err.Raise 9999, "AssertEqual", message & " - Expected: " & CStr(expected) & ", Actual: " & CStr(actual)
    End If
End Sub

Sub AssertTrue(ByVal condition, ByVal message)
    Assert condition, message
End Sub

Sub AssertFalse(ByVal condition, ByVal message)
    Assert Not condition, message
End Sub
