Attribute VB_Name = "TestCalculator"
Option Explicit

Public Sub Test_Add()
    Dim c As Calculator
    Set c = New Calculator

    expect(c.Add(1, 2)).toBe 3
    expect(c.Add(0, 0)).toBe 0
    expect(c.Add(-1, 1)).toBe 0
End Sub

Public Sub Test_Subtract()
    Dim c As Calculator
    Set c = New Calculator

    expect(c.Subtract(10, 4)).toBe 6
    expect(c.Subtract(5, 5)).toBe 0
    expect(c.Subtract(0, 10)).toBe -10
End Sub

' === 意図的に失敗するテスト（検証用） ===
Public Sub Test_ShouldFail_ToBe()
    Dim c As Calculator
    Set c = New Calculator

    ' 1 + 2 = 3 だが、わざと 5 を期待して失敗させる
    expect(c.Add(1, 2)).toBe 5
End Sub

Public Sub Test_ShouldFail_ToContain()
    Dim text As String
    text = "hello world"

    ' "foo" は含まれていないので失敗
    expect(text).toContain "foo"
End Sub

Public Sub Test_ShouldFail_ToBeGreaterThan()
    ' 3 > 10 は偽なので失敗
    expect(3).toBeGreaterThan 10
End Sub

Public Sub Test_ShouldFail_NotToBe()
    ' 5 は 5 なので、Not_.toBe 5 は失敗
    expect(5).Not_.toBe 5
End Sub
