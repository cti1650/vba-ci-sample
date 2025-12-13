Attribute VB_Name = "RunTests"
Option Explicit

Private Const SUCCESS_PATH As String = "C:\temp\success.txt"
Private Const ERROR_PATH As String = "C:\temp\error.txt"

Public Sub RunAll()
  On Error GoTo ErrHandler

  ' === Register tests here ===
  Call TestCalculator.Test_Add
  Call TestCalculator.Test_Subtract

  ' Mark success
  Call Utils.WriteTextFile(SUCCESS_PATH, "OK")
  Exit Sub

ErrHandler:
  Dim msg As String
  msg = CStr(Err.Number) & ":" & Err.Description
  Call Utils.WriteTextFile(ERROR_PATH, msg)
End Sub
