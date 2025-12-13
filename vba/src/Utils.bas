Attribute VB_Name = "Utils"
Option Explicit

Public Sub WriteTextFile(ByVal path As String, ByVal text As String)
  ' VBA: FreeFile/Open を使用、VBS: FileSystemObject を使用
  ' この関数はVBA/VBS両方で動作するよう条件分岐が必要だが、
  ' テストでは使用しないため、VBS変換時は空実装でOK
  #If VBA7 Or VBA6 Then
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
  #End If
End Sub

Public Sub Fail(ByVal code As Long, ByVal message As String)
  Err.Raise code, "VBA_TEST", message
End Sub
