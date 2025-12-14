' run-tests.vbs - テスト実行エントリーポイント
' VBAから変換されたVBSファイルを読み込み、Test_* 関数を自動検出して実行する
'
' 重要: すべてのコードを1つの文字列に結合してからExecuteGlobalを呼ぶ
' これにより、すべてのコードが同じグローバルスコープに存在し、相互参照が可能になる
Option Explicit

Dim fso, scriptDir, genDir, compatPath
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
genDir = fso.BuildPath(scriptDir, "generated")
compatPath = fso.BuildPath(scriptDir, "vba-compat.vbs")

If Not fso.FolderExists(genDir) Then
    WScript.Echo "ERROR: generated folder not found: " & genDir
    WScript.Quit 1
End If

' ============================================
' Step 1: すべてのコードを1つの文字列に結合
' ============================================
Dim allCode, file, files, fileContent, enumsPath

' GetScriptDir関数を最初に定義（scriptDirの値を埋め込む）
allCode = "' === GetScriptDir ===" & vbCrLf & _
          "Function GetScriptDir()" & vbCrLf & _
          "    GetScriptDir = """ & scriptDir & """" & vbCrLf & _
          "End Function" & vbCrLf & vbCrLf

' VBA互換レイヤーを追加
If fso.FileExists(compatPath) Then
    Dim compatFile
    Set compatFile = fso.GetFile(compatPath)
    If compatFile.Size > 0 Then
        allCode = allCode & "' === vba-compat.vbs ===" & vbCrLf & _
                  fso.OpenTextFile(compatPath).ReadAll & vbCrLf & vbCrLf
    End If
End If

' _enums.vbs を追加（Enum定数の定義）
enumsPath = fso.BuildPath(genDir, "_enums.vbs")
If fso.FileExists(enumsPath) Then
    Dim enumsFile
    Set enumsFile = fso.GetFile(enumsPath)
    If enumsFile.Size > 0 Then
        allCode = allCode & "' === _enums.vbs ===" & vbCrLf & _
                  fso.OpenTextFile(enumsPath).ReadAll & vbCrLf & vbCrLf
    End If
End If

' 各生成ファイルを追加
Set files = fso.GetFolder(genDir).Files
For Each file In files
    If LCase(fso.GetExtensionName(file.Name)) = "vbs" Then
        If LCase(file.Name) <> "_enums.vbs" Then
            ' 空ファイルをスキップ（ReadAllで"Input past end of file"エラーを防ぐ）
            If file.Size > 0 Then
                allCode = allCode & "' === " & file.Name & " ===" & vbCrLf & _
                          fso.OpenTextFile(file.Path).ReadAll & vbCrLf & vbCrLf
            End If
        End If
    End If
Next

' ============================================
' Step 2: すべてのコードを一度にExecuteGlobal
' ============================================
On Error Resume Next
ExecuteGlobal allCode
If Err.Number <> 0 Then
    WScript.Echo "ERROR loading code: " & Err.Description
    WScript.Echo "Error source: " & Err.Source
    WScript.Quit 1
End If
On Error GoTo 0

' ============================================
' Step 3: Test_で始まる関数を検出
' ============================================
Dim passCount, failCount, testNames
passCount = 0
failCount = 0

Set testNames = CreateObject("Scripting.Dictionary")
Dim regex, matches, match
Set regex = New RegExp
regex.Global = True
regex.IgnoreCase = True
regex.Pattern = "(?:Sub|Function)\s+(Test_\w+)\s*\("

Set matches = regex.Execute(allCode)
For Each match In matches
    testNames(match.SubMatches(0)) = True
Next

WScript.Echo "========================================="
WScript.Echo "VBS Test Runner"
WScript.Echo "========================================="
WScript.Echo ""

' ============================================
' Step 4: 各テストを実行
' ============================================
Dim testName
For Each testName In testNames.Keys
    On Error Resume Next
    ExecuteGlobal "Call " & testName & "()"
    If Err.Number <> 0 Then
        WScript.Echo "[FAIL] " & testName & ": " & Err.Description
        failCount = failCount + 1
        Err.Clear
    Else
        WScript.Echo "[PASS] " & testName
        passCount = passCount + 1
    End If
    On Error GoTo 0
Next

' ============================================
' Step 5: 結果サマリー
' ============================================
WScript.Echo ""
WScript.Echo "========================================="
WScript.Echo "Results: " & passCount & " passed, " & failCount & " failed"
WScript.Echo "========================================="

If passCount + failCount = 0 Then
    WScript.Echo "ERROR: No tests found!"
    WScript.Quit 1
End If

If failCount > 0 Then
    WScript.Quit 1
Else
    WScript.Quit 0
End If
