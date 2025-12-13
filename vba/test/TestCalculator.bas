Attribute VB_Name = "TestCalculator"
Option Explicit

Public Sub Test_Add()
  Dim c As Calculator
  Set c = New Calculator

  If c.Add(1, 2) <> 3 Then
    Utils.Fail 1001, "Add(1,2) expected 3"
  End If
End Sub

Public Sub Test_Subtract()
  Dim c As Calculator
  Set c = New Calculator

  If c.Subtract(10, 4) <> 6 Then
    Utils.Fail 1002, "Subtract(10,4) expected 6"
  End If
End Sub

' 構文エラーテスト用（CI動作確認用）
Public Sub Test_SyntaxError(
  ' ← 括弧が閉じていない構文エラー
End Sub
