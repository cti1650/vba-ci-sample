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
