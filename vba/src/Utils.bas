Attribute VB_Name = "Utils"
Option Explicit

Public Sub WriteTextFile(ByVal path As String, ByVal text As String)
  Dim f As Integer
  f = FreeFile
  Open path For Output As #f
  Print #f, text
  Close #f
End Sub

Public Sub Fail(ByVal code As Long, ByVal message As String)
  Err.Raise code, "VBA_TEST", message
End Sub
